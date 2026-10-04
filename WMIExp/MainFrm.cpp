// MainFrm.cpp : implmentation of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#include "pch.h"
#include "resource.h"
#include "AboutDlg.h"
#include "MainFrm.h"
#include "SecurityHelper.h"
#include "AppSettings.h"
#include "IconHelper.h"
#include <WTLHelper.h>
#include <SortHelper.h>
#include <ListViewhelper.h>
#include <ClipboardHelper.h>
#include "TextDlg.h"
#include "ConnectDlg.h"
#include "ExecMethodDlg.h"

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg) {
	// the (modeless) dialogs get their keyboard handling, and not the main window's accelerators
	for (HWND hDlg : { m_SearchDlg.m_hWnd, m_QueryDlg.m_hWnd, m_EventsDlg.m_hWnd }) {
		if (hDlg && (pMsg->hwnd == hDlg || ::IsChild(hDlg, pMsg->hwnd)))
			return ::IsDialogMessage(hDlg, pMsg);
	}

	return CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle() {
	UIUpdateToolBar();
	return FALSE;
}

CString CMainFrame::GetColumnText(HWND h, int row, int col) const {
	if (h == m_InstanceList) {
		auto index = GetColumnManager(h)->GetColumnTag<int>(col);
		if (row >= (int)m_Objects.size() || index >= (int)m_InstanceColumns.size())
			return L"";

		// values are read as they are shown: an instance is a local copy, so this is cheap
		auto& column = m_InstanceColumns[index];
		CComVariant value;
		if (FAILED(m_Objects[row].Object->Get(column.Name, 0, &value, nullptr, nullptr)))
			return L"";
		return WMIHelper::FormatValue(value, column.Type);
	}

	auto column = GetColumnManager(h)->GetColumnTag<ColumnType>(col);
	if (h == m_List) {
		auto& item = m_Items[row];
		switch (column) {
			case ColumnType::Name: return item.Name.c_str();
			case ColumnType::Type: return NodeTypeToText(item.Type);
			case ColumnType::CimType:
				if (item.Type == NodeType::Property) {
					auto text = WMIHelper::CimTypeToString(item.CimType);
					return text;
				}
				break;

			case ColumnType::Value: return GetObjectValue(item);
			case ColumnType::Details: return GetObjectDetails(item);
			case ColumnType::Description:
				if (!item.DescriptionLoaded && item.Type == NodeType::Class && m_spCurrentNamespace) {
					// one more call to WMI per class, so only for the classes shown
					CComPtr<IWbemClassObject> spClass;
					auto hr = m_spCurrentNamespace->GetObject(CComBSTR(item.Name.c_str()), WBEM_FLAG_USE_AMENDED_QUALIFIERS, nullptr, &spClass, nullptr);
					// (asked by another process, such as an accessibility tool, COM cannot call out: tried again when painted)
					item.DescriptionLoaded = hr != RPC_E_CANTCALLOUT_ININPUTSYNCCALL;
					if (SUCCEEDED(hr))
						item.Description = FlattenText(WMIHelper::GetClassDescription(spClass));
				}
				return item.Description;
		}
	}
	return L"";
}

CString CMainFrame::FlattenText(CString text) {
	// one line, for a list view
	text.Replace(L"\r\n", L" ");
	text.Replace(L'\n', L' ');
	text.Replace(L'\t', L' ');
	return text.Trim();
}

bool CMainFrame::OnDoubleClickList(HWND h, int row, int, POINT const&) {
	if (h == m_InstanceList) {
		if (row < 0 || row >= (int)m_Objects.size())
			return false;
		CTextDlg::ShowObject(m_hWnd, m_Objects[row].Object.get());
		return true;
	}
	if (h == m_List && row >= 0 && row < (int)m_Items.size()) {
		auto& item = m_Items[row];
		switch (item.Type) {
			case NodeType::Method:
				ExecuteMethod(item);
				return true;

			case NodeType::Class:
			case NodeType::Namespace:
			{
				// into it, in the tree
				auto hParent = m_Tree.GetSelectedItem();
				if (hParent == nullptr)
					return false;
				LoadNamespaceChildren(hParent);
				auto hItem = item.Type == NodeType::Class ? FindClassItem(hParent, item.Name.c_str()) : FindChildItem(hParent, item.Name.c_str(), NodeType::Namespace);
				if (hItem) {
					m_Tree.EnsureVisible(hItem);
					m_Tree.SelectItem(hItem);
				}
				return true;
			}
		}
	}
	return false;
}

int CMainFrame::GetRowImage(HWND h, int row, int) const {
	if (h == m_List) {
		switch (m_Items[row].Type) {
			case NodeType::Namespace: return 0;
			case NodeType::Class: return 1;
			case NodeType::Instance: return 5;
			case NodeType::Property:
				return m_Items[row].Name.substr(0, 2) == L"__" ? 4 : 2;

			case NodeType::Method: return 3;
		}
	}
	else {
		return 5;
	}
	return -1;
}

void CMainFrame::OnStateChanged(HWND h, int from, int to, UINT oldState, UINT newState) {
	if (h != m_InstanceList)
		return;

	int index = m_InstanceList.GetSelectedIndex();
	if (index >= 0 && index < (int)m_Objects.size())
		m_ObjPropValues = WMIHelper::EnumProperties(m_Objects[index].Object.get());
	else
		m_ObjPropValues.clear();
	m_List.RedrawItems(m_List.GetTopIndex(), m_List.GetTopIndex() + m_List.GetCountPerPage());
}

void CMainFrame::ClearInstanceColumns() {
	ClearSort(m_InstanceList);
	GetColumnManager(m_InstanceList)->Clear();
	m_InstanceColumns.clear();
}

void CMainFrame::BuildInstanceColumns() {
	ClearInstanceColumns();
	if (m_spCurrentClass == nullptr)
		return;

	auto add = [&](CString const& name, CIMTYPE type, PCWSTR header = nullptr) {
		if (std::ranges::any_of(m_InstanceColumns, [&](auto const& c) { return c.Name.CompareNoCase(name) == 0; }))
			return;

		int width = 120;
		int format = LVCFMT_LEFT;
		if (type & CIM_FLAG_ARRAY)
			width = 180;
		else {
			switch (type) {
				case CIM_STRING: case CIM_REFERENCE: width = 180; break;
				case CIM_DATETIME: width = 140; break;
				case CIM_BOOLEAN: width = 70; break;
				case CIM_SINT8: case CIM_UINT8: case CIM_SINT16: case CIM_UINT16: case CIM_SINT32: case CIM_UINT32:
				case CIM_SINT64: case CIM_UINT64: case CIM_REAL32: case CIM_REAL64:
					width = 90;
					format = LVCFMT_RIGHT;
					break;
			}
		}
		if (header == nullptr)
			header = name;
		width = std::max(width, (int)wcslen(header) * 7 + 24);
		GetColumnManager(m_InstanceList)->AddColumn(header, format, width, (int)m_InstanceColumns.size());
		m_InstanceColumns.push_back({ name, type });
	};

	// key properties first, as they identify the instance
	for (auto& prop : WMIHelper::EnumProperties(m_spCurrentClass, WBEM_FLAG_KEYS_ONLY))
		add(CString(prop.Name), prop.Type);

	// the class of each instance, if not all are of the selected class
	if (!m_Objects.empty()) {
		auto className = WMIHelper::GetStringProperty(m_spCurrentClass, L"__CLASS");
		if (std::ranges::any_of(m_Objects, [&](auto const& obj) { return WMIHelper::GetStringProperty(obj.Object.get(), L"__CLASS").CompareNoCase(className) != 0; }))
			add(L"__CLASS", CIM_STRING, L"Class");
	}

	for (auto& prop : WMIHelper::EnumProperties(m_spCurrentClass, WBEM_FLAG_NONSYSTEM_ONLY))
		add(CString(prop.Name), prop.Type);

	if (AppSettings::Get().ViewSystemProperties() || m_InstanceColumns.empty()) {
		for (auto& prop : WMIHelper::EnumProperties(m_spCurrentClass, WBEM_FLAG_SYSTEM_ONLY))
			add(CString(prop.Name), prop.Type);
	}
}

void CMainFrame::SortInstances(const SortInfo* si) {
	auto index = GetColumnManager(m_InstanceList)->GetColumnTag<int>(si->SortColumn);
	if (index >= (int)m_InstanceColumns.size())
		return;

	auto& column = m_InstanceColumns[index];
	auto count = m_Objects.size();
	std::vector<CComVariant> values(count);
	for (size_t i = 0; i < count; i++)
		m_Objects[i].Object->Get(column.Name, 0, &values[i], nullptr, nullptr);

	std::vector<size_t> order(count);
	for (size_t i = 0; i < count; i++)
		order[i] = i;
	std::ranges::stable_sort(order, [&](size_t i1, size_t i2) {
		auto result = WMIHelper::CompareValues(values[i1], values[i2], column.Type);
		return si->SortAscending ? result < 0 : result > 0;
		});

	// keep the selected instance selected (the list view keeps the index, not the instance)
	int selected = m_InstanceList.GetSelectedIndex();
	int newSelected = -1;
	std::vector<WmiItem> sorted;
	sorted.reserve(count);
	for (size_t i = 0; i < count; i++) {
		if ((int)order[i] == selected)
			newSelected = (int)i;
		sorted.push_back(std::move(m_Objects[order[i]]));
	}
	m_Objects = std::move(sorted);

	if (newSelected >= 0 && newSelected != selected) {
		m_InstanceList.SetItemState(newSelected, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
		m_InstanceList.EnsureVisible(newSelected, FALSE);
	}
}

LRESULT CMainFrame::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/) {
	auto& settings = AppSettings::Get();

	m_hSingleInstMutex = ::CreateMutex(nullptr, FALSE, L"WmiExpSingleInstanceMutex");
	if (settings.SingleInstance() && m_hSingleInstMutex) {
		if (::GetLastError() == ERROR_ALREADY_EXISTS) {
			//
			// not first instance
			//
			auto hMainWnd = ::FindWindow(GetWndClassName(), nullptr);
			if (hMainWnd) {
				::SetActiveWindow(hMainWnd);
				::SetForegroundWindow(hMainWnd);
				return -1;
			}
		}
	}

	CMenuHandle menu = GetMenu();
	if (SecurityHelper::IsRunningElevated()) {
		auto fileMenu = menu.GetSubMenu(0);
		fileMenu.DeleteMenu(0, MF_BYPOSITION);
		fileMenu.DeleteMenu(0, MF_BYPOSITION);
	}

	InitMenu(menu);
	UIAddMenu(menu);

	CToolBarCtrl tb;
	tb.Create(m_hWnd, nullptr, nullptr, ATL_SIMPLE_TOOLBAR_PANE_STYLE, 0, ATL_IDW_TOOLBAR);
	InitToolBar(tb, 24);
	UIAddToolBar(tb);

	CreateSimpleReBar(ATL_SIMPLE_REBAR_NOBORDER_STYLE);
	AddSimpleReBarBand(tb, nullptr, TRUE);

	CReBarCtrl rb(m_hWndToolBar);
	rb.LockBands(true);

	CreateSimpleStatusBar(ATL_IDS_IDLEMESSAGE, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | SBARS_SIZEGRIP | SBT_TOOLTIPS);
	m_StatusBar.SubclassWindow(m_hWndStatusBar);
	int panes[] = { 100, 300, 700 };
	m_StatusBar.SetParts(_countof(panes), panes);

	m_hWndClient = m_Splitter.Create(m_hWnd, rcDefault, nullptr,
		WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_CLIENTEDGE);

	m_Tree.Create(m_Splitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
		TVS_HASBUTTONS | TVS_LINESATROOT | TVS_HASLINES | TVS_SHOWSELALWAYS | TVS_INFOTIP, 0, TreeId);
	m_Tree.SetExtendedStyle(TVS_EX_DOUBLEBUFFER | TVS_EX_RICHTOOLTIP, 0);
	// (class descriptions can be long: wrapped)
	CToolTipCtrl(m_Tree.GetToolTips()).SetMaxTipWidth(500);

	m_DetailSplitter.Create(m_Splitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
	m_DetailSplitter.SetSplitterPosPct(55);

	CImageList images;
	images.Create(16, 16, ILC_COLOR32 | ILC_COLOR | ILC_MASK, 10, 4);
	UINT icons[] = {
		IDI_NAMESPACE, IDI_CLASS, IDI_PROPERTY, IDI_OBJECT, IDI_PROPERTY2,
		IDI_METHOD
	};
	for (auto icon : icons)
		images.AddIcon(AtlLoadIconImage(icon, 0, 16, 16));
	m_Tree.SetImageList(images, TVSIL_NORMAL);
	//::SetWindowTheme(m_Tree, L"Explorer", nullptr);

	m_List.Create(m_DetailSplitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
		| LVS_OWNERDATA | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SHAREIMAGELISTS, 0);
	m_List.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP);
	m_List.SetImageList(images, LVSIL_SMALL);

	auto cm = GetColumnManager(m_List);
	cm->AddColumn(L"Name", LVCFMT_LEFT, 220, ColumnType::Name);
	cm->AddColumn(L"Type", LVCFMT_LEFT, 110, ColumnType::Type);
	cm->AddColumn(L"CIM Type", LVCFMT_LEFT, 120, ColumnType::CimType);
	cm->AddColumn(L"Value", LVCFMT_LEFT, 250, ColumnType::Value);
	cm->AddColumn(L"Property Value", LVCFMT_LEFT, 400, ColumnType::Details);
	cm->AddColumn(L"Description", LVCFMT_LEFT, 500, ColumnType::Description);

	m_InstanceList.Create(m_DetailSplitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
		| LVS_OWNERDATA | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SHAREIMAGELISTS | LVS_SINGLESEL, 0);
	m_InstanceList.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);
	m_InstanceList.SetImageList(images, LVSIL_SMALL);
	// columns are the properties of the selected class (see BuildInstanceColumns)

	m_Splitter.SetSplitterPanes(m_Tree, m_DetailSplitter);
	m_DetailSplitter.SetSplitterPanes(m_List, m_InstanceList);

	m_Splitter.SetSplitterPosPct(20);

	auto pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	//
	// update UI based on settings
	//
	UISetCheck(ID_OPTIONS_ALWAYSONTOP, settings.AlwaysOnTop());
	UISetCheck(ID_OPTIONS_SINGLEINSTANCE, settings.SingleInstance());
	UISetCheck(ID_VIEW_SYSTEMPROPERTIES, settings.ViewSystemProperties());
	UISetCheck(ID_VIEW_SYSTEMCLASSES, settings.ViewSystemClasses());
	UISetCheck(ID_VIEW_NAMESPACESINLIST, settings.ShowNamespacesInList());
	UISetCheck(ID_VIEW_DERIVEDINSTANCES, settings.DerivedInstances());
	UISetCheck(ID_VIEW_CLASSHIERARCHY, settings.ClassHierarchy());
	UISetCheck(ID_OPTIONS_DARKMODE, WTLHelper::IsDarkMode());

	if (settings.AlwaysOnTop())
		SetWindowPos(HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

	UpdateLayout();

	UpdateTitle();
	if (auto hr = WMIHelper::Connect(WMIHelper::CurrentConnection(), m_RootName, &m_spWmi); FAILED(hr))
		m_StatusBar.SetText(2, L"Error connecting to WMI: " + WMIHelper::GetErrorText(hr));
	InitTree();

	return 0;
}

LRESULT CMainFrame::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
	CancelInstanceEnum();
	if (m_SearchDlg.IsWindow())
		m_SearchDlg.DestroyWindow();
	if (m_QueryDlg.IsWindow())
		m_QueryDlg.DestroyWindow();
	if (m_EventsDlg.IsWindow())
		m_EventsDlg.DestroyWindow();
	AppSettings::Get().Save();

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);
	pLoop->RemoveIdleHandler(this);

	bHandled = FALSE;
	return 1;
}

LRESULT CMainFrame::OnTimer(UINT, WPARAM id, LPARAM, BOOL&) {
	if (id == 2) {
		KillTimer(id);
		TreeItemSelected(nullptr);
	}
	return 0;
}

LRESULT CMainFrame::OnAddInstances(UINT, WPARAM, LPARAM lp, BOOL& bHandled) {
	auto job = WMIHelper::TakeJob(lp);

	// results of an enumeration that was since cancelled or replaced are dropped
	if (job != m_EnumJob)
		return 0;

	m_EnumJob.reset();
	m_Objects.clear();
	m_Objects.reserve(job->Objects.size());
	for (auto& obj : job->Objects) {
		WmiItem item;
		item.Type = NodeType::Instance;
		item.Object = obj.p;
		m_Objects.push_back(std::move(item));
	}

	BuildInstanceColumns();
	m_InstanceList.SetItemCount((int)m_Objects.size());
	if (FAILED(job->Status))
		m_StatusBar.SetText(2, std::format(L"{} Objects (Error: {})", m_Objects.size(), (PCWSTR)WMIHelper::GetErrorText(job->Status)).c_str());
	else
		m_StatusBar.SetText(2, std::format(L"{} Objects", m_Objects.size()).c_str());
	return 0;
}

void CMainFrame::CancelInstanceEnum() {
	if (m_EnumJob) {
		// the worker stops soon, and does not post its results
		m_EnumJob->Cancelled = true;
		m_EnumJob.reset();
	}
}

LRESULT CMainFrame::OnFileExit(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {
	PostMessage(WM_CLOSE);
	return 0;
}

LRESULT CMainFrame::OnViewToolBar(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {
	static BOOL bVisible = TRUE;	// initially visible
	bVisible = !bVisible;
	CReBarCtrl rebar = m_hWndToolBar;
	int nBandIndex = rebar.IdToIndex(ATL_IDW_BAND_FIRST + 1);	// toolbar is 2nd added band
	rebar.ShowBand(nBandIndex, bVisible);
	UISetCheck(wID, bVisible);
	UpdateLayout();
	return 0;
}

LRESULT CMainFrame::OnViewStatusBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {
	BOOL bVisible = !::IsWindowVisible(m_hWndStatusBar);
	::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
	UpdateLayout();
	return 0;
}

LRESULT CMainFrame::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {
	CAboutDlg dlg;
	dlg.DoModal();
	return 0;
}

LRESULT CMainFrame::OnTreeItemExpanding(int, LPNMHDR hdr, BOOL&) {
	auto tv = reinterpret_cast<NMTREEVIEW*>(hdr);
	LoadNamespaceChildren(tv->itemNew.hItem);
	return 0;
}

void CMainFrame::LoadNamespaceChildren(HTREEITEM hItem) {
	auto hChild = m_Tree.GetChildItem(hItem);
	if (hChild == nullptr || GetTreeNodeType(hChild) != NodeType::HasChildren)
		return;		// loaded already

	auto path = GetFullItemPath(m_Tree, hItem);
	path = path.Mid(path.Find(L'\\') + 1);
	CComPtr<IWbemServices> spNamespace;
	CWaitCursor wait;
	m_Tree.DeleteItem(hChild);
	auto hr = WMIHelper::OpenNamespace(m_spWmi, path, &spNamespace);
	// (the current namespace is the selected one's: TreeItemSelected sets it)
	if (SUCCEEDED(hr))
		BuildTree(spNamespace, hItem);
	else {
		m_StatusBar.SetText(2, std::format(L"Error opening {}: {}", (PCWSTR)path, (PCWSTR)WMIHelper::GetErrorText(hr)).c_str());
	}

	if (m_Tree.GetChildItem(hItem) == nullptr) {
		// nothing in it (or it cannot be opened): no expand button
		TVITEM tvi{ TVIF_CHILDREN };
		tvi.hItem = hItem;
		tvi.cChildren = 0;
		m_Tree.SetItem(&tvi);
	}
}

HTREEITEM CMainFrame::FindChildItem(HTREEITEM hParent, CString const& name, NodeType type) {
	CString text;
	for (auto hItem = m_Tree.GetChildItem(hParent); hItem; hItem = m_Tree.GetNextSiblingItem(hItem)) {
		if (GetTreeNodeType(hItem) == type && m_Tree.GetItemText(hItem, text) && text.CompareNoCase(name) == 0)
			return hItem;
	}
	return nullptr;
}

//
// a class in a namespace: at the top, or (with the class hierarchy) under its superclass
//
HTREEITEM CMainFrame::FindClassItem(HTREEITEM hNamespace, CString const& name) {
	if (auto hItem = FindChildItem(hNamespace, name, NodeType::Class))
		return hItem;

	if (!AppSettings::Get().ClassHierarchy())
		return nullptr;

	std::function<HTREEITEM(HTREEITEM)> find = [&](HTREEITEM hParent) -> HTREEITEM {
		CString text;
		for (auto hItem = m_Tree.GetChildItem(hParent); hItem; hItem = m_Tree.GetNextSiblingItem(hItem)) {
			if (GetTreeNodeType(hItem) != NodeType::Class)
				continue;
			if (m_Tree.GetItemText(hItem, text) && text.CompareNoCase(name) == 0)
				return hItem;
			if (auto hFound = find(hItem))
				return hFound;
		}
		return nullptr;
	};
	return find(hNamespace);
}

HTREEITEM CMainFrame::GetNamespaceItem(HTREEITEM hItem) const {
	while (hItem && GetTreeNodeType(hItem) != NodeType::Namespace)
		hItem = m_Tree.GetParentItem(hItem);
	return hItem;
}

CString CMainFrame::GetClassDescription(CString const& nsPath, CString const& className) {
	auto key = nsPath + L":" + className;
	if (auto it = m_ClassDescriptions.find(key); it != m_ClassDescriptions.end())
		return it->second;

	// the namespace of the last tooltip is kept: the mouse usually moves between classes of one namespace
	if (m_TipNamespacePath != nsPath) {
		m_spTipNamespace.Release();
		m_TipNamespacePath = nsPath;
		if (nsPath.CompareNoCase(m_RootName) == 0)
			m_spTipNamespace = m_spWmi;
		else if (m_spWmi)
			WMIHelper::OpenNamespace(m_spWmi, nsPath.Mid(m_RootName.GetLength() + 1), &m_spTipNamespace);
	}

	CString description;
	CComPtr<IWbemClassObject> spClass;
	if (m_spTipNamespace && SUCCEEDED(m_spTipNamespace->GetObject(CComBSTR(className), WBEM_FLAG_USE_AMENDED_QUALIFIERS, nullptr, &spClass, nullptr)))
		description = WMIHelper::GetClassDescription(spClass);
	m_ClassDescriptions[key] = description;
	return description;
}

LRESULT CMainFrame::OnTreeGetInfoTip(int, LPNMHDR hdr, BOOL&) {
	auto tip = reinterpret_cast<NMTVGETINFOTIP*>(hdr);
	if (GetTreeNodeType(tip->hItem) != NodeType::Class)
		return 0;

	CString name;
	m_Tree.GetItemText(tip->hItem, name);
	auto description = GetClassDescription(GetFullItemPath(m_Tree, GetNamespaceItem(tip->hItem)), name);
	if (!description.IsEmpty())
		::StringCchCopy(tip->pszText, tip->cchTextMax, description);
	return 0;
}

bool CMainFrame::NavigateTo(SearchResult const& result) {
	auto path = result.Namespace;
	if (result.Type == SearchResult::Kind::Namespace)
		path += L"\\" + result.Name;

	// walk down the namespaces, loading each on the way (as expanding it would)
	int start = 0;
	auto segment = path.Tokenize(L"\\", start);
	if (segment.CompareNoCase(m_RootName) != 0)
		return false;

	auto hItem = m_hRoot;
	for (segment = path.Tokenize(L"\\", start); !segment.IsEmpty(); segment = path.Tokenize(L"\\", start)) {
		LoadNamespaceChildren(hItem);
		hItem = FindChildItem(hItem, segment, NodeType::Namespace);
		if (hItem == nullptr)
			return false;
	}

	if (result.Type != SearchResult::Kind::Namespace) {
		LoadNamespaceChildren(hItem);
		hItem = FindClassItem(hItem, result.Type == SearchResult::Kind::Class ? result.Name : result.Class);
		if (hItem == nullptr)
			return false;
	}

	m_Tree.EnsureVisible(hItem);
	m_Tree.SelectItem(hItem);
	// now, not when the selection timer fires, so the property can be selected in the list
	KillTimer(2);
	TreeItemSelected(hItem);

	if (result.Type == SearchResult::Kind::Property || result.Type == SearchResult::Kind::Method) {
		auto type = result.Type == SearchResult::Kind::Property ? NodeType::Property : NodeType::Method;
		auto it = std::ranges::find_if(m_Items, [&](auto const& item) {
			return item.Type == type && result.Name.CompareNoCase(item.Name.c_str()) == 0;
			});
		if (it == m_Items.end())
			return false;

		int index = static_cast<int>(it - m_Items.begin());
		m_List.SetItemState(-1, 0, LVIS_SELECTED);
		m_List.SetItemState(index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
		m_List.EnsureVisible(index, FALSE);
	}
	return true;
}

LRESULT CMainFrame::OnEditFind(WORD, WORD, HWND, BOOL&) {
	m_SearchDlg.Activate(m_hWnd);
	return 0;
}

LRESULT CMainFrame::OnTreeSelChanged(int, LPNMHDR, BOOL&) {
	m_InstanceList.SetItemCount(0);
	SetTimer(2, 200, nullptr);
	return 0;
}

LRESULT CMainFrame::OnViewSystemClasses(WORD, WORD id, HWND, BOOL&) {
	bool view;
	AppSettings::Get().ViewSystemClasses(view = !AppSettings::Get().ViewSystemClasses());
	UISetCheck(id, view);
	// recreate tree
	auto node = m_Tree.GetSelectedItem();
	CString path;
	if (node) {
		path = GetFullItemPath(m_Tree, node);
	}
	m_List.SetItemCount(0);
	m_InstanceList.SetItemCount(0);
	InitTree();
	if (!path.IsEmpty()) {
		node = FindItem(m_Tree, TVI_ROOT, path);
		if (node)
			m_Tree.SelectItem(node);
	}

	return 0;
}

LRESULT CMainFrame::OnViewSystemProperties(WORD, WORD id, HWND, BOOL&) {
	bool view;
	AppSettings::Get().ViewSystemProperties(view = !AppSettings::Get().ViewSystemProperties());
	UISetCheck(id, view);
	UpdateList();
	if (!m_Objects.empty()) {
		BuildInstanceColumns();
		m_InstanceList.Invalidate();
	}
	return 0;
}

LRESULT CMainFrame::OnViewNamespacesInList(WORD, WORD id, HWND, BOOL&) {
	bool view;
	AppSettings::Get().ShowNamespacesInList(view = !AppSettings::Get().ShowNamespacesInList());
	UISetCheck(id, view);
	UpdateList();
	return 0;
}

LRESULT CMainFrame::OnViewDerivedInstances(WORD, WORD id, HWND, BOOL&) {
	bool view;
	AppSettings::Get().DerivedInstances(view = !AppSettings::Get().DerivedInstances());
	UISetCheck(id, view);
	if (m_spCurrentClass)
		TreeItemSelected(nullptr);
	return 0;
}

LRESULT CMainFrame::OnViewRefresh(WORD, WORD, HWND, BOOL&) {
	TreeItemSelected(nullptr);
	return 0;
}

LRESULT CMainFrame::OnEditCopy(WORD, WORD, HWND, BOOL&) {
	auto hFocus = ::GetFocus();
	CString text;
	if (hFocus == m_List || hFocus == m_InstanceList) {
		text = ListViewHelper::GetSelectedRowsAsString(CListViewCtrl(hFocus));
	}
	else if (hFocus == m_Tree) {
		if (auto hItem = m_Tree.GetSelectedItem())
			m_Tree.GetItemText(hItem, text);
	}
	if (!text.IsEmpty())
		ClipboardHelper::CopyText(m_hWnd, text);
	return 0;
}

LRESULT CMainFrame::OnSingleInstance(WORD, WORD id, HWND, BOOL&) {
	bool single;
	AppSettings::Get().SingleInstance(single = !AppSettings::Get().SingleInstance());
	UISetCheck(id, single);
	return 0;
}

PCWSTR CMainFrame::NodeTypeToText(NodeType type) {
	switch (type) {
		case NodeType::Class: return L"Class";
		case NodeType::Namespace: return L"Namespace";
		case NodeType::Method: return L"Method";
		case NodeType::Property: return L"Property";
		case NodeType::Instance: return L"Object";
	}
	return L"";
}

void CMainFrame::InitMenu(HMENU menu) {
	MenuItemData commands[] = {
		{ ID_FILE_RUNASADMINISTRATOR, 0, IconHelper::GetShieldIcon() },
		{ ID_EDIT_COPY, IDI_COPY },
		{ ID_VIEW_REFRESH, IDI_REFRESH },
	};
	WTLHelper::InitMenu(menu, commands, _countof(commands));
}

void CMainFrame::InitToolBar(CToolBarCtrl& tb, int size) {
	CImageList tbImages;
	tbImages.Create(size, size, ILC_COLOR32, 8, 4);
	tb.SetImageList(tbImages);

	const struct {
		UINT id;
		int image;
		BYTE style = BTNS_BUTTON;
		PCWSTR text = nullptr;
	} buttons[] = {
		{ ID_VIEW_REFRESH, IDI_REFRESH },
		{ 0 },
		{ ID_EDIT_COPY, IDI_COPY },
	};
	for (auto& b : buttons) {
		if (b.id == 0)
			tb.AddSeparator(0);
		else {
			auto hIcon = AtlLoadIconImage(b.image, 0, size, size);
			ATLASSERT(hIcon);
			int image = tbImages.AddIcon(hIcon);
			tb.AddButton(b.id, b.style, TBSTATE_ENABLED, image, b.text, 0);
		}
	}
}

void CMainFrame::InitTree() {
	m_spCurrentNamespace = m_spWmi;
	m_NamespacePath = m_RootName;
	m_Tree.LockWindowUpdate();
	m_Tree.DeleteAllItems();
	m_hRoot = InsertTreeItem(m_RootName, 0, TVI_ROOT, NodeType::Namespace);
	if (m_spWmi) {
		BuildTree(m_spWmi, m_hRoot);
		m_Tree.Expand(m_hRoot, TVE_EXPAND);
	}
	m_Tree.LockWindowUpdate(FALSE);
	m_Tree.SelectItem(m_hRoot);
	m_Tree.SetFocus();
}

void CMainFrame::BuildTree(IWbemServices* pWmi, HTREEITEM hParent) {
	struct Node {
		CString Name;
		CString SuperClass;
		bool IsNamespace;
	};
	bool hierarchy = AppSettings::Get().ClassHierarchy();
	std::vector<Node> nodes;
	for (auto& spObj : WMIHelper::EnumClasses(pWmi, true, AppSettings::Get().ViewSystemClasses()))
		nodes.push_back({ WMIHelper::GetStringProperty(spObj, L"__CLASS"), hierarchy ? WMIHelper::GetStringProperty(spObj, L"__SUPERCLASS") : CString(), false });
	for (auto& spObj : WMIHelper::EnumNamespaces(pWmi))
		nodes.push_back({ WMIHelper::GetStringProperty(spObj, L"NAME"), L"", true });

	// sorting here and adding at the end is much faster than inserting sorted (TVI_SORT) one by one
	std::ranges::sort(nodes, [](auto const& n1, auto const& n2) { return ::lstrcmpiW(n1.Name, n2.Name) < 0; });

	// with the hierarchy: the subclasses of each class (sorted, as the nodes are)
	std::map<CString, std::vector<Node const*>> subclasses;
	std::set<CString> names;
	auto isSubclass = [&](Node const& node) {
		// (a superclass can be hidden, as a system class: then the class is at the top)
		return hierarchy && !node.IsNamespace && !node.SuperClass.IsEmpty() && names.contains(node.SuperClass);
	};
	if (hierarchy) {
		for (auto& node : nodes)
			if (!node.IsNamespace)
				names.insert(node.Name);
		for (auto& node : nodes)
			if (isSubclass(node))
				subclasses[node.SuperClass].push_back(&node);
	}

	m_Tree.SetRedraw(FALSE);
	std::function<void(Node const&, HTREEITEM)> insert = [&](Node const& node, HTREEITEM hParentItem) {
		auto hItem = InsertTreeItem(node.Name, node.IsNamespace ? 0 : 1, hParentItem, node.IsNamespace ? NodeType::Namespace : NodeType::Class);
		//
		// a namespace gets a placeholder child (so it can be expanded) without opening it;
		// it is opened when expanded, and loses its expand button then if it turns out empty
		//
		if (node.IsNamespace)
			InsertTreeItem(L"\\\\", 0, hItem, NodeType::HasChildren);
		else if (auto it = subclasses.find(node.Name); it != subclasses.end()) {
			for (auto child : it->second)
				insert(*child, hItem);
		}
	};
	for (auto& node : nodes) {
		if (!isSubclass(node))
			insert(node, hParent);
	}
	m_Tree.SetRedraw(TRUE);
	m_Tree.Invalidate();
}

void CMainFrame::UpdateList() {
	m_List.SetItemCount(0);
	m_Items.clear();

	auto& settings = AppSettings::Get();
	if (m_spCurrentClass) {
		// descriptions are in the class with its amended (localized) qualifiers
		CComPtr<IWbemClassObject> spAmended;
		m_spCurrentNamespace->GetObject(CComBSTR(WMIHelper::GetStringProperty(m_spCurrentClass, L"__CLASS")),
			WBEM_FLAG_USE_AMENDED_QUALIFIERS, nullptr, &spAmended, nullptr);

		for (auto& prop : WMIHelper::EnumProperties(m_spCurrentClass)) {
			if (!settings.ViewSystemProperties() && CString(prop.Name).Left(2) == L"__")
				continue;
			WmiItem item;
			item.Name = prop.Name;
			item.Type = NodeType::Property;
			item.CimType = prop.Type;
			item.Value = prop.Value;
			if (spAmended)
				item.Description = FlattenText(WMIHelper::GetPropertyDescription(spAmended, prop.Name));
			item.DescriptionLoaded = true;
			m_Items.push_back(std::move(item));
		}
		for (auto& method : WMIHelper::EnumMethods(m_spCurrentClass)) {
			WmiItem item;
			item.Name = method.Name;
			item.Type = NodeType::Method;
			item.Object = method.spInParams;
			item.Object2 = method.spOutParams;
			item.Value = method.ClassName.c_str();
			if (spAmended)
				item.Description = FlattenText(WMIHelper::GetMethodDescription(spAmended, method.Name.c_str()));
			item.DescriptionLoaded = true;
			m_Items.push_back(std::move(item));
		}
	}
	else {
		if (m_spCurrentNamespace == nullptr)
			return;

		if (settings.ShowNamespacesInList()) {
			for (auto& ns : WMIHelper::EnumNamespaces(m_spCurrentNamespace)) {
				WmiItem item;
				CComBSTR name;
				item.Name = WMIHelper::GetStringProperty(ns, L"NAME");
				item.Object = ns;
				item.Type = NodeType::Namespace;
				m_Items.push_back(std::move(item));
			}
		}
		for (auto& cls : WMIHelper::EnumClasses(m_spCurrentNamespace, true, settings.ViewSystemClasses())) {
			WmiItem item;
			CComBSTR name;
			item.Name = WMIHelper::GetStringProperty(cls, L"__CLASS");
			item.Object = cls;
			item.Type = NodeType::Class;
			m_Items.push_back(std::move(item));
		}
	}
	RefreshList();
}

void CMainFrame::DoSort(const SortInfo* si) {
	if (si->hWnd == m_InstanceList) {
		SortInstances(si);
		return;
	}

	auto column = GetColumnManager(si->hWnd)->GetColumnTag<ColumnType>(si->SortColumn);
	ATLASSERT(si->hWnd == m_List);
	if (column == ColumnType::Value || column == ColumnType::Details) {
		SortItemsByValue(si, column);
		return;
	}

	auto sort =[&](const auto& i1, const auto& i2) {
		switch (column) {
			case ColumnType::Name: return SortHelper::Sort(i1.Name, i2.Name, si->SortAscending);
			case ColumnType::Type: return SortHelper::Sort(i1.Type, i2.Type, si->SortAscending);
			case ColumnType::CimType: return SortHelper::Sort(i1.CimType, i2.CimType, si->SortAscending);
		}
		return false;
		};
	std::sort(m_Items.begin(), m_Items.end(), sort);
}

//
// the Value and Details columns: empty values first, then numbers (by value), then text.
// keeping the three groups apart makes the order consistent when items of different types are mixed
//
void CMainFrame::SortItemsByValue(const SortInfo* si, ColumnType column) {
	struct SortKey {
		int Group;		// 0: empty, 1: number, 2: text
		double Number;
		CString Text;
	};

	auto makeKey = [&](WmiItem const& item) {
		SortKey key{ 2, 0 };
		key.Text = column == ColumnType::Value ? GetObjectValue(item) : GetObjectDetails(item);
		if (key.Text.IsEmpty()) {
			key.Group = 0;
			return key;
		}
		if (item.Type != NodeType::Property)
			return key;

		// the value shown: the class's (Value) or the selected instance's (Details)
		CComVariant value;
		auto type = item.CimType;
		if (column == ColumnType::Value)
			value = item.Value;
		else {
			auto it = std::ranges::find_if(m_ObjPropValues, [&](auto const& p) { return p.Name == item.Name.c_str(); });
			if (it == m_ObjPropValues.end())
				return key;
			value = it->Value;
			type = it->Type;
		}
		switch (type) {
			case CIM_SINT8: case CIM_UINT8: case CIM_SINT16: case CIM_UINT16:
			case CIM_SINT32: case CIM_UINT32: case CIM_SINT64: case CIM_UINT64:
			case CIM_REAL32: case CIM_REAL64:
				// (64-bit integers come as strings)
				if (SUCCEEDED(value.ChangeType(VT_R8))) {
					key.Group = 1;
					key.Number = value.dblVal;
				}
				break;
		}
		return key;
	};

	std::vector<SortKey> keys;
	keys.reserve(m_Items.size());
	for (auto& item : m_Items)
		keys.push_back(makeKey(item));

	std::vector<size_t> order(m_Items.size());
	std::iota(order.begin(), order.end(), size_t(0));
	std::ranges::stable_sort(order, [&](size_t i1, size_t i2) {
		auto& k1 = keys[i1];
		auto& k2 = keys[i2];
		int result;
		if (k1.Group != k2.Group)
			result = k1.Group - k2.Group;
		else if (k1.Group == 1)
			result = k1.Number < k2.Number ? -1 : (k1.Number > k2.Number ? 1 : 0);
		else
			result = k1.Text.CompareNoCase(k2.Text);
		return si->SortAscending ? result < 0 : result > 0;
		});

	std::vector<WmiItem> items;
	items.reserve(m_Items.size());
	for (auto i : order)
		items.push_back(std::move(m_Items[i]));
	m_Items = std::move(items);
}

CString CMainFrame::GetObjectDetails(WmiItem const& item) const {
	switch (item.Type) {
		case NodeType::Property:
		{
			if (m_ObjPropValues.empty())
				return L"";

			auto it = std::find_if(m_ObjPropValues.begin(), m_ObjPropValues.end(), [&](auto const& p) { return p.Name == item.Name.c_str(); });
			if (it == m_ObjPropValues.end())
				return L"";

			return WMIHelper::FormatValue(it->Value, it->Type);
		}

		case NodeType::Method:
			if (item.Object) {	// in params
				auto props = WMIHelper::EnumProperties(item.Object.get());
				CString text;
				for (auto& p : props) {
					text += WMIHelper::CimTypeToString(p.Type) + L" " + p.Name + L", ";
				}
				if (!text.IsEmpty())
					text = text.Left(text.GetLength() - 2);
				return L"(" + text + L")";
			}
			break;
	}
	return L"";
}

CString CMainFrame::GetObjectValue(WmiItem const& item) const {
	switch (item.Type) {
		case NodeType::Property:
			return WMIHelper::FormatValue(item.Value, item.CimType);
	}
	return L"";
}

void CMainFrame::TreeItemSelected(HTREEITEM hItem) {
	CancelInstanceEnum();
	m_StatusBar.SetText(2, L"");
	m_InstanceList.SetItemCount(0);
	m_Objects.clear();
	m_ObjPropValues.clear();
	ClearInstanceColumns();

	if(hItem == nullptr)
		hItem = m_Tree.GetSelectedItem();
	if (hItem == nullptr) {
		m_spCurrentClass = nullptr;
		m_spCurrentNamespace = nullptr;
		return;
	}
	auto path = GetFullItemPath(m_Tree, hItem);
	auto type = GetTreeNodeType(hItem);
	CString name;
	m_Tree.GetItemText(hItem, name);
	switch (type) {
		case NodeType::Namespace:
		{
			if (hItem == m_hRoot) {
				m_NamespacePath = m_RootName;
				m_spCurrentNamespace = m_spWmi;
			}
			else {
				CComPtr<IWbemServices> spNamespace;
				path = path.Mid(path.Find(L'\\') + 1);
				auto hr = WMIHelper::OpenNamespace(m_spWmi, path, &spNamespace);
				if (FAILED(hr)) {
					// don't leave the previous namespace current: its contents would be shown as this one's
					m_spCurrentNamespace = nullptr;
					m_spCurrentClass = nullptr;
					m_NamespacePath.Empty();
					m_Items.clear();
					RefreshList();
					m_StatusBar.SetText(2, std::format(L"Error opening {}: {}", (PCWSTR)path, (PCWSTR)WMIHelper::GetErrorText(hr)).c_str());
					return;
				}
				m_spCurrentNamespace = spNamespace;
				m_NamespacePath = m_RootName + L"\\" + path;
			}
			m_spCurrentClass = nullptr;
			break;
		}
		case NodeType::Class:
		{
			// (with the class hierarchy, the parent can be a class)
			auto hNamespace = GetNamespaceItem(hItem);
			if (m_NamespacePath != GetFullItemPath(m_Tree, hNamespace))
				TreeItemSelected(hNamespace);
			m_spCurrentClass = nullptr;
			if (m_spCurrentNamespace == nullptr)
				return;		// the namespace could not be opened (the error is in the status bar)

			auto hr = m_spCurrentNamespace->GetObject(CComBSTR(name), 0, nullptr, &m_spCurrentClass, nullptr);
			if (m_spCurrentClass) {
				m_InstanceList.SetItemCount(0);
				m_Objects.clear();
				m_EnumJob = WMIHelper::EnumInstancesAsync(m_hWnd, WM_INSTANCES, name, m_spCurrentNamespace, AppSettings::Get().DerivedInstances());
				m_StatusBar.SetText(2, L"Enumerating Objects...");
			}
			else {
				m_Items.clear();
				RefreshList();
				m_StatusBar.SetText(2, std::format(L"Error getting class {}: {}", (PCWSTR)name, (PCWSTR)WMIHelper::GetErrorText(hr)).c_str());
				return;
			}
			break;
		}

		default:
			ATLASSERT(false);
			return;
	}

	UpdateList();
}

void CMainFrame::RefreshList() {
	m_List.SetItemCountEx(static_cast<int>(m_Items.size()), LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
	m_List.RedrawItems(m_List.GetTopIndex(), m_List.GetTopIndex() + m_List.GetCountPerPage());
	m_StatusBar.SetText(1, std::format(L"{} Items", m_Items.size()).c_str());
}

HTREEITEM CMainFrame::InsertTreeItem(PCWSTR text, int image, HTREEITEM hParent, NodeType type) {
	auto hItem = m_Tree.InsertItem(text, image, image, hParent, TVI_LAST);
	ATLASSERT(hItem);
	m_Tree.SetItemData(hItem, static_cast<ULONG_PTR>(type));
	return hItem;
}

CMainFrame::NodeType CMainFrame::GetTreeNodeType(HTREEITEM hItem) const {
	return static_cast<NodeType>(m_Tree.GetItemData(hItem));
}

LRESULT CMainFrame::OnShowWindow(UINT, WPARAM, LPARAM, BOOL&) {
	static bool shown = false;
	if (!shown) {
		shown = true;
		auto wp = AppSettings::Get().MainWindowPlacement();
		if (wp.showCmd)
			SetWindowPlacement(&wp);
		SetAlwaysOnTop(AppSettings::Get().AlwaysOnTop());
	}
	return 0;
}

void CMainFrame::SetAlwaysOnTop(bool onTop) {
	SetWindowPos(onTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	UISetCheck(ID_OPTIONS_ALWAYSONTOP, onTop);
}

LRESULT CMainFrame::OnRunAsAdmin(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {
	AppSettings::Get().Save();
	if (SecurityHelper::RunElevated())
		PostMessage(WM_CLOSE);
	return 0;
}

LRESULT CMainFrame::OnAlwaysOnTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {
	auto& settings = AppSettings::Get();
	settings.AlwaysOnTop(!settings.AlwaysOnTop());
	SetAlwaysOnTop(settings.AlwaysOnTop());

	return 0;
}

LRESULT CMainFrame::OnToggleDarkMode(WORD, WORD, HWND, BOOL&) {
	WTLHelper::SwitchToMode(WTLHelper::IsDarkMode() ? DarkModeKind::Classic : DarkModeKind::Dark, m_hWnd);
	AppSettings::Get().DarkMode(WTLHelper::IsDarkMode() ? 1 : 0);
	InitMenu(GetMenu());
	DrawMenuBar();
	UISetCheck(ID_OPTIONS_DARKMODE, WTLHelper::IsDarkMode());
	return 0;
}

void CMainFrame::UpdateTitle() {
	CString title;
	title.LoadString(IDR_MAINFRAME);
	auto conn = WMIHelper::CurrentConnection();
	if (!conn->Computer.IsEmpty()) {
		title += L" - \\\\" + conn->Computer;
		if (conn->HasCredentials())
			title += L" (" + conn->User + L")";
	}
	if (SecurityHelper::IsRunningElevated())
		title += L" (Administrator)";
	SetWindowText(title);
}

LRESULT CMainFrame::OnConnect(WORD, WORD, HWND, BOOL&) {
	CConnectDlg dlg;
	if (dlg.DoModal() != IDOK)
		return 0;

	// everything from the previous computer goes
	CancelInstanceEnum();
	m_SearchDlg.Reset();
	m_QueryDlg.Reset();
	m_EventsDlg.Reset();
	WMIHelper::SetCurrentConnection(dlg.GetConnection());
	m_spWmi = dlg.GetRoot();
	m_spCurrentNamespace = nullptr;
	m_spCurrentClass = nullptr;
	m_NamespacePath.Empty();
	m_ClassDescriptions.clear();
	m_spTipNamespace.Release();
	m_TipNamespacePath.Empty();
	m_Items.clear();
	RefreshList();
	m_Objects.clear();
	m_ObjPropValues.clear();
	m_InstanceList.SetItemCount(0);
	ClearInstanceColumns();
	m_StatusBar.SetText(2, L"");

	UpdateTitle();
	InitTree();
	return 0;
}

LRESULT CMainFrame::OnQuery(WORD, WORD, HWND, BOOL&) {
	m_QueryDlg.Activate(m_hWnd, m_NamespacePath.IsEmpty() ? CString(L"ROOT\\CIMV2") : m_NamespacePath,
		m_spCurrentClass ? WMIHelper::GetStringProperty(m_spCurrentClass, L"__CLASS") : CString());
	return 0;
}

LRESULT CMainFrame::OnEvents(WORD, WORD, HWND, BOOL&) {
	m_EventsDlg.Activate(m_hWnd);
	return 0;
}

LRESULT CMainFrame::OnExecuteMethod(WORD, WORD, HWND, BOOL&) {
	int index = m_List.GetSelectedIndex();
	if (index < 0 || index >= (int)m_Items.size() || m_Items[index].Type != NodeType::Method) {
		AtlMessageBox(m_hWnd, L"Select a method of the class in the list.", IDS_TITLE, MB_ICONINFORMATION);
		return 0;
	}
	ExecuteMethod(m_Items[index]);
	return 0;
}

//
// a static method runs on the class; any other on the instance selected in the instance list
//
void CMainFrame::ExecuteMethod(WmiItem const& method) {
	if (m_spCurrentClass == nullptr || m_spCurrentNamespace == nullptr)
		return;

	CString path;
	if (WMIHelper::IsStaticMethod(m_spCurrentClass, method.Name.c_str()))
		path = WMIHelper::GetStringProperty(m_spCurrentClass, L"__CLASS");
	else {
		int index = m_InstanceList.GetSelectedIndex();
		if (index < 0 || index >= (int)m_Objects.size()) {
			AtlMessageBox(m_hWnd, std::format(L"{} is not a static method: select the instance to run it on.", method.Name).c_str(),
				IDS_TITLE, MB_ICONINFORMATION);
			return;
		}
		path = WMIHelper::GetStringProperty(m_Objects[index].Object.get(), L"__RELPATH");
	}
	CExecMethodDlg dlg(m_spCurrentNamespace, path, method.Name.c_str(), method.Object.get());
	dlg.DoModal(m_hWnd);
}

LRESULT CMainFrame::OnShowMof(WORD, WORD, HWND, BOOL&) {
	// the selected instance or class in the focused list, otherwise the selected class
	auto hFocus = ::GetFocus();
	if (hFocus == m_InstanceList) {
		int index = m_InstanceList.GetSelectedIndex();
		if (index >= 0 && index < (int)m_Objects.size()) {
			CTextDlg::ShowObject(m_hWnd, m_Objects[index].Object.get());
			return 0;
		}
	}
	else if (hFocus == m_List) {
		int index = m_List.GetSelectedIndex();
		if (index >= 0 && index < (int)m_Items.size() && m_Items[index].Type == NodeType::Class && m_Items[index].Object) {
			CTextDlg::ShowObject(m_hWnd, m_Items[index].Object.get());
			return 0;
		}
	}
	if (m_spCurrentClass) {
		CTextDlg::ShowObject(m_hWnd, m_spCurrentClass);
		return 0;
	}
	AtlMessageBox(m_hWnd, L"Select a class or an instance.", IDS_TITLE, MB_ICONINFORMATION);
	return 0;
}

LRESULT CMainFrame::OnViewClassHierarchy(WORD, WORD id, HWND, BOOL&) {
	bool hierarchy;
	AppSettings::Get().ClassHierarchy(hierarchy = !AppSettings::Get().ClassHierarchy());
	UISetCheck(id, hierarchy);

	// the selected item is selected again in the rebuilt tree (where its path may differ)
	std::optional<SearchResult> selected;
	auto hItem = m_Tree.GetSelectedItem();
	if (hItem && hItem != m_hRoot) {
		CString name;
		m_Tree.GetItemText(hItem, name);
		if (GetTreeNodeType(hItem) == NodeType::Class)
			selected = SearchResult{ SearchResult::Kind::Class, name, GetFullItemPath(m_Tree, GetNamespaceItem(hItem)) };
		else
			selected = SearchResult{ SearchResult::Kind::Namespace, name, GetFullItemPath(m_Tree, m_Tree.GetParentItem(hItem)) };
	}
	InitTree();
	if (selected)
		NavigateTo(*selected);
	return 0;
}
