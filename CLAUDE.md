# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

WMI Explorer (`WMIExp`): a native Windows GUI (C++20, WTL/ATL, Win32) for browsing WMI namespaces, classes, properties, methods and instances. x64 only. Work in progress; there are no tests or lint tooling.

## Build

Visual Studio 2026 (toolset `v145`). Solution: `WMIExplorer.sln`, which contains `WMIExp` and the `WTLHelper` static library (a project reference).

```
nuget restore WMIExplorer.sln
msbuild WMIExplorer.sln /p:Configuration=Debug /p:Platform=x64 /m
```

- Configurations: `Debug`, `Release`, `ReleaseSigned` (all `x64`).
- NuGet packages restore into `packages/`. `WMIExp` uses `wtl`; `WTLHelper` also uses `Microsoft.Windows.ImplementationLibrary` (WIL, `<wil\com.h>` is used throughout) and `Detours`.
- `msbuild` is not on PATH by default. Run from a VS Developer prompt or use `"C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe"`.
- Output goes to `x64\<Config>\`.

## WTLHelper dependency

`wtlhelper/` (https://github.com/zodiacon/wtlhelper) is listed in `.gitmodules`, but its files are tracked as ordinary files in this repo (path `WTLHelper/` in git), not as a gitlink. The paths are case-insensitive on Windows, so `..\WTLHelper\WTLHelper` (include dir) and `..\wtlhelper\WTLHelper\WTLHelper.vcxproj` (project reference) point to the same place. It provides the reusable UI building blocks the main frame mixes in: `CVirtualListView`, `CTreeViewHelper` (`GetFullItemPath`), `CCustomSplitterWindow`, and the dark mode support (`WTLHelper.h`, `DarkMode/`). Look there before writing new generic UI helpers. Treat it as a shared library: don't make WMIExp-specific changes in it.

## Architecture

The app is small and almost all logic lives in `CMainFrame` (`WMIExp/MainFrm.*`):

- **Startup** (`WMIExp.cpp`): STA COM init, then `CoInitializeSecurity` with impersonation set to `RPC_C_IMP_LEVEL_DELEGATE`. Without this, WMI returns access-denied errors. Next, settings are loaded and `WTLHelper::InitDarkMode` runs, both before any window is created. Then the WTL message loop.
- **WMI access** (`WMIHelper.*`): static wrappers over `IWbemServices`/`IWbemClassObject` that enumerate namespaces, classes, instances, properties and methods, plus value formatting (`FormatValue`, `CompareValues`, `CimTypeToString`), parsing text into values (`ParseValue`), and descriptions (`Get*Description`, which need an object fetched with `WBEM_FLAG_USE_AMENDED_QUALIFIERS`). `GetErrorText` turns an HRESULT into readable text.
- **Connections**: `WMIConnection` holds the computer and credentials. `WMIHelper::CurrentConnection()` is the one the main window uses. Connect with `WMIHelper::Connect`, and open namespaces with `WMIHelper::OpenNamespace`, never `IWbemServices::OpenNamespace` directly. Every proxy (services or enumerator) needs `WMIHelper::SetSecurity`, which sets encryption and the credentials on remote proxies; the `WMIHelper` enumerations already do this. Connections are never freed, because proxies keep pointing to their `COAUTHIDENTITY`. File → Connect (`CConnectDlg`) replaces the connection and resets the tree, search, query and event dialogs.
- **Background queries**: `EnumInstancesAsync` and `ExecQueryAsync` run a semi-synchronous query on a worker thread (the services proxy is marshaled to it), not an async sink: WMI calling back a sink fails with remote computers in many setups. When done, the worker posts a heap `std::shared_ptr<WMIQueryJob>*`, which the receiver takes with `WMIHelper::TakeJob` (that frees it) and compares with its current job. Cancelling sets `Cancelled`: the worker stops within 500 ms and posts nothing.
- **Tree (left pane)**: namespaces and classes, lazily populated. Every namespace gets a placeholder child (`NodeType::HasChildren`) without being opened. `OnTreeItemExpanding` replaces it by opening the namespace and calling `BuildTree`, and removes the expand button if the namespace turns out to be empty or can't be opened. `BuildTree` sorts items itself and adds them at the end (no `TVI_SORT`). The `WMIHelper` enumerations read results in batches (`ReadAll`), not one object at a time. The `NodeType` of each item is stored in its item data. The namespace path is derived from the tree path (`GetFullItemPath`) of the namespace item (`GetNamespaceItem`), relative to the `ROOT` node. With View → Class Hierarchy, classes are nested under their superclasses (a class whose superclass is hidden stays at the top), so a class's parent can be a class: use `GetNamespaceItem` and `FindClassItem`, not the parent item. Tree tooltips show class descriptions (`OnTreeGetInfoTip`, cached in `m_ClassDescriptions`).
- **Selection**: `TVN_SELCHANGED` starts a 200 ms timer (id 2), which calls `TreeItemSelected`. That cancels any instance enumeration in progress (`CancelInstanceEnum`), sets `m_spCurrentNamespace` and `m_spCurrentClass`, starts a new enumeration and calls `UpdateList`. If the namespace can't be opened, `m_spCurrentNamespace` becomes null (it is never left pointing to the previous namespace) and the error goes to status pane 2. The frame keeps the active job in `m_EnumJob`, and `OnAddInstances` drops any other job's results.
- **Lists (right side, horizontal splitter)**: `m_List` shows class members (`m_Items`: properties and methods) or namespace contents. `m_InstanceList` shows instances (`m_Objects`). Its columns are rebuilt for each class by `BuildInstanceColumns`: key properties first, then `__CLASS` if the instances are of mixed classes, then the other properties. Cells read values on demand with `IWbemClassObject::Get`, and all value text goes through `FormatValue`. `m_List` has a Description column: properties and methods get theirs in `UpdateList`, and classes in a namespace's list get theirs when first painted (`WmiItem::DescriptionLoaded`). Double-clicking a method runs it (`ExecuteMethod`, `CExecMethodDlg`): a static method runs on the class, and any other on the selected instance. Double-clicking an instance shows its MOF (`CTextDlg`). Both are virtual list views driven by `GetColumnText`, `GetRowImage`, `DoSort` and `OnStateChanged` through `CVirtualListView`.
- **Settings** (`Settings.*`, `AppSettings.h`): registry-backed singleton (`AppSettings::Get()`) at `HKCU\Software\ScorpioSoftware\WmiExp`, declared with the `SETTING` and `DEF_SETTING` macros. To add a setting, add both a `SETTING(...)` entry and a `DEF_SETTING(...)` accessor. Settings are loaded in `wWinMain`, not in `OnCreate`.
- **UI resources**: commands and IDs are in `resource.h` / `WMIExp.rc`. Menu icons are attached as bitmaps by `InitMenu` (through `WTLHelper::InitMenu`), and toolbar icons by `InitToolBar`.

## Conventions

- Tabs for indentation; the opening brace goes on the same line.
- COM smart pointers: both `CComPtr` and `wil::com_ptr` are used. Strings: `CString`, `CComBSTR`, `std::wstring`. Use `std::format` for formatted UI text.
- `pch.h` holds all the ATL/WTL/WMI/STL includes; every `.cpp` includes it first.

## Dark mode

Dark mode follows the AstroStudio project and uses WTLHelper's dark mode library. `WTLHelper::InitDarkMode` installs a thread `WH_CALLWNDPROCRET` hook that themes every window as it is created, and Detours hooks on `GetSysColor`/`GetSysColorBrush`. Menus are plain (not owner-drawn): the dark mode library paints the menu bar. Don't bring back `COwnerDrawnMenu` or `ThemeHelper`, because they conflict with it.
- The `DarkMode` setting holds 1 (dark), 0 (light) or -1 (follow the system; this is the default until the user toggles it).
- **Options → Dark Mode** (`OnToggleDarkMode`) calls `WTLHelper::SwitchToMode`, then calls `InitMenu` again because menu icon bitmaps are pre-filled with the background color.
- Custom-drawn controls can react to `WTLHelper::ThemeChangedMessage`, which is sent to all descendants when the mode switches.
- Before showing a common dialog that doesn't support dark mode, wrap it in `WTLHelper::SuspendHook`/`ResumeHook` (see `WTLHelper::InvokeFontDialog`).

## Editing gotcha

The sources use CRLF line endings, and `WMIExp.rc` is ANSI (ISO-8859), not UTF-8. Git Bash `sed -i` drops the CRs, so use the Edit tool, or `perl -i -pe` for `.rc` changes.

## Search

`CSearchDlg` (`SearchDlg.*`, Edit → Find, Ctrl+F) is a modeless, resizable dialog (`CDialogResize`). It is hidden on Close, not destroyed, so its results are kept. `CMainFrame::PreTranslateMessage` sends it its keyboard messages before the main window's accelerators.
- The search runs on a `std::jthread` (MTA) with its own `ROOT` connection, because COM pointers from the UI thread's STA can't be used there. It goes through every namespace and posts `SearchProgress` objects (`WM_SEARCH_PROGRESS`, deleted by the receiver). A generation number makes the dialog ignore posts from a replaced search.
- Properties are matched with `WBEM_FLAG_LOCAL_ONLY`, so an inherited property is listed once, under the class that defines it. Whatever the View options hide (system classes and properties) is not searched.
- Double-clicking a result calls `ISearchNavigator::NavigateTo`, which `CMainFrame` implements. It walks the tree path and loads each namespace with `LoadNamespaceChildren` (the same code expanding a namespace uses), then selects the item. For a property, it also selects the row in the list.

## Tools dialogs

- **Query** (`CQueryDlg`, Tools → Query, Ctrl+Q) and **Event Viewer** (`CEventsDlg`, Ctrl+E) are modeless and hidden on Close, like Search. `PreTranslateMessage` routes keyboard messages to all three. Query runs `ExecQueryAsync`, with columns from the results' properties. Event Viewer listens on its own `std::jthread` (`ExecNotificationQuery`, `Next` with a 500 ms timeout so it notices a stop request) and posts `EventBatch` objects with a generation number.
- **Show MOF** (`CTextDlg`, Ctrl+M) shows `GetObjectText` for the selected instance or class, with Copy and Save As (UTF-16 with a BOM, which mofcomp accepts). The save dialog is wrapped in `SuspendResumeHook`, because it doesn't support dark mode.
- To test the app from another process (UI Automation sees only panes here), send window messages. A cross-process `SendMessage` is an input-synchronous call: COM calls made while handling it fail with `RPC_E_CANTCALLOUT_ININPUTSYNCCALL`. So post anything that makes the app call WMI (tree expand and select, commands, button clicks).
