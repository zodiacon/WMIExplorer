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
- **WMI access** (`WMIHelper.*`): static wrappers over `IWbemServices`/`IWbemClassObject` that enumerate namespaces, classes, instances, properties and methods. `EnumInstancesAsync` uses a `CComObject` sink (`IWbemObjectSink`) that collects results and posts a message (`WM_INSTANCES`) to the frame with an `IObjectsCallback*` in `lParam`. The posted message holds its own reference, so the receiver must always call `Release()`, even for results it ignores. `GetStatus()` gives the final HRESULT, and `GetErrorText` turns it into readable text.
- **Tree (left pane)**: namespaces and classes, lazily populated. Every namespace gets a placeholder child (`NodeType::HasChildren`) without being opened. `OnTreeItemExpanding` replaces it by opening the namespace and calling `BuildTree`, and removes the expand button if the namespace turns out to be empty or can't be opened. `BuildTree` sorts items itself and adds them at the end (no `TVI_SORT`). The `WMIHelper` enumerations read results in batches (`ReadAll`), not one object at a time. The `NodeType` of each item is stored in its item data. The namespace path is derived from the tree path (`GetFullItemPath`) relative to the `ROOT` node, and opened via `m_spWmi->OpenNamespace`.
- **Selection**: `TVN_SELCHANGED` starts a 200 ms timer (id 2), which calls `TreeItemSelected`. That cancels any instance enumeration in progress (`CancelInstanceEnum`, which calls `CancelAsyncCall`), sets `m_spCurrentNamespace` and `m_spCurrentClass`, starts a new enumeration and calls `UpdateList`. The frame keeps the active sink in `m_spEnumSink`, and `OnAddInstances` drops any result whose sink is not that one.
- **Lists (right side, horizontal splitter)**: `m_List` shows class members (`m_Items`: properties and methods) or namespace contents. `m_InstanceList` shows instances (`m_Objects`). Its columns are rebuilt for each class by `BuildInstanceColumns`: key properties first, then `__CLASS` if the instances are of mixed classes, then the other properties. Cells read values on demand with `IWbemClassObject::Get`, and all value text goes through `FormatValue`. Both are virtual list views driven by `GetColumnText`, `GetRowImage`, `DoSort` and `OnStateChanged` through `CVirtualListView`.
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
