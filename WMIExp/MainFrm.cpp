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
#include "WMIHelper.h"

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg) {
	return CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle() {
	UIUpdateToolBar();
	return FALSE;
}

CString CMainFrame::GetColumnText(HWND h, int row, int col) const {
	if (h == m_List) {
		auto& item = m_Items[row];
		switch (GetColumnManager(h)->GetColumnTag<ColumnType>(col)) {
			case ColumnType::Name: return item.Name;
			case ColumnType::Type: return NodeTypeToText(item.Type);
			case ColumnType::CimType:
				if (item.Type == NodeType::Property) {
					auto text = CimTypeToString(item.CimType);
					return text;
				}
				break;

			case ColumnType::Value: return GetObjectValue(item);
			case ColumnType::Details: return GetObjectDetails(item);
		}
	}
	return L"";
}

int CMainFrame::GetRowImage(HWND h, int row) const {
	if (h == m_List) {
		switch (m_Items[row].Type) {
			case NodeType::Namespace: return 0;
			case NodeType::Class: return 1;
			case NodeType::Instance: return 5;
			case NodeType::Property:
				return m_Items[row].Name.Left(2) == L"__" ? 4 : 2;

			case NodeType::Method: return 3;
		}
	}
	return -1;
}

DWORD CMainFrame::OnPrePaint(int, LPNMCUSTOMDRAW cd) {
	if (cd->hdr.hwndFrom != m_List)
		return CDRF_DODEFAULT;

	return CDRF_NOTIFYITEMDRAW;
}

DWORD CMainFrame::OnItemPrePaint(int, LPNMCUSTOMDRAW cd) {
	int index = (int)cd->dwItemSpec;
	auto& item = m_Items[index];
	if (item.Type != NodeType::Instance)
		return CDRF_DODEFAULT;

	CDCHandle dc(cd->hdc);
	dc.SetBkMode(TRANSPARENT);
	bool selected = m_List.GetItemState(index, LVIS_SELECTED) > 0;
	auto& rc = cd->rc;
	dc.FillSolidRect(&rc, ::GetSysColor(selected ? COLOR_HIGHLIGHT : COLOR_WINDOW));
	dc.SetTextColor(::GetSysColor(selected ? COLOR_HIGHLIGHTTEXT : COLOR_WINDOWTEXT));
	CRect rcIcon(CPoint(rc.left + 2, rc.top + 2), CSize(16, 16));
	m_List.GetImageList(LVSIL_SMALL).DrawEx(5, dc, rcIcon, CLR_NONE, CLR_NONE, ILD_NORMAL);
	rc.left += 24;
	dc.DrawText(item.Name, -1, &rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

	return CDRF_SKIPDEFAULT;
}

LRESULT CMainFrame::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/) {
	auto& settings = AppSettings::Get();

	settings.Load(L"Software\\ScorpioSoftware\\WmiExp");
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
		CString text;
		GetWindowText(text);
		text += L" (Administrator)";
		SetWindowText(text);
	}

	HWND hWndCmdBar = m_CmdBar.Create(m_hWnd, rcDefault, nullptr, ATL_SIMPLE_CMDBAR_PANE_STYLE);

	m_CmdBar.SetAlphaImages(true);
	m_CmdBar.AttachMenu(menu);
	InitCommandBar();

	UIAddMenu(menu);
	SetMenu(nullptr);

	CToolBarCtrl tb;
	tb.Create(m_hWnd, nullptr, nullptr, ATL_SIMPLE_TOOLBAR_PANE_STYLE, 0, ATL_IDW_TOOLBAR);
	InitToolBar(tb, 24);
	UIAddToolBar(tb);

	CreateSimpleReBar(ATL_SIMPLE_REBAR_NOBORDER_STYLE);
	AddSimpleReBarBand(hWndCmdBar);
	AddSimpleReBarBand(tb, nullptr, TRUE);

	CReBarCtrl rb(m_hWndToolBar);
	rb.LockBands(true);

	CreateSimpleStatusBar(ATL_IDS_IDLEMESSAGE, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | SBARS_SIZEGRIP | SBT_TOOLTIPS);
	m_StatusBar.SubclassWindow(m_hWndStatusBar);
//	int panes[] = { 24, 1300 };
//	m_StatusBar.SetParts(_countof(panes), panes);

	m_hWndClient = m_Splitter.Create(m_hWnd, rcDefault, nullptr,
		WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_CLIENTEDGE);

	m_Tree.Create(m_Splitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
		TVS_HASBUTTONS | TVS_LINESATROOT | TVS_HASLINES | TVS_SHOWSELALWAYS, 0, TreeId);
	m_Tree.SetExtendedStyle(TVS_EX_DOUBLEBUFFER | TVS_EX_RICHTOOLTIP, 0);

	m_DetailSplitter.Create(m_Splitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
	m_DetailSplitter.SetSplitterPosPct(50);

	CImageList images;
	images.Create(16, 16, ILC_COLOR32 | ILC_COLOR | ILC_MASK, 10, 4);
	UINT icons[] = {
		IDI_NAMESPACE, IDI_CLASS, IDI_PROPERTY, IDI_METHOD, IDI_PROPERTY2,
		IDI_OBJECT
	};
	for (auto icon : icons)
		images.AddIcon(AtlLoadIconImage(icon, 0, 16, 16));
	m_Tree.SetImageList(images, TVSIL_NORMAL);
	//::SetWindowTheme(m_Tree, L"Explorer", nullptr);

	m_List.Create(m_DetailSplitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
		| LVS_OWNERDATA | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SHAREIMAGELISTS, 0);
	m_List.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP);
	m_List.SetImageList(images, LVSIL_SMALL);
	//::SetWindowTheme(m_List, L"Explorer", nullptr);

	auto cm = GetColumnManager(m_List);
	cm->AddColumn(L"Name", LVCFMT_LEFT, 220, ColumnType::Name);
	cm->AddColumn(L"Type", LVCFMT_LEFT, 110, ColumnType::Type);
	//cm->AddColumn(L"Size", LVCFMT_RIGHT, 70, ColumnType::Size);
	cm->AddColumn(L"Value", LVCFMT_LEFT, 250, ColumnType::Value);
	cm->AddColumn(L"CIM Type", LVCFMT_LEFT, 120, ColumnType::CimType);
	cm->AddColumn(L"Details", LVCFMT_LEFT, 250, ColumnType::Details);
	cm->UpdateColumns();

	m_InstanceList.Create(m_DetailSplitter, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
		| LVS_OWNERDATA | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SHAREIMAGELISTS, 0);
	m_InstanceList.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP);
	cm = GetColumnManager(m_InstanceList);
	cm->AddColumn(L"Property", LVCFMT_LEFT, 150, ColumnType::Name);
	cm->AddColumn(L"Value", LVCFMT_LEFT, 350, ColumnType::Value);

	m_Splitter.SetSplitterPanes(m_Tree, m_DetailSplitter);
	m_DetailSplitter.SetSplitterPanes(m_List, m_InstanceList);

	m_Splitter.SetSplitterPosPct(25);
	m_Splitter.UpdateSplitterLayout();

	auto pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	//
	// update UI based on settings
	//
	UISetCheck(ID_OPTIONS_ALWAYSONTOP, settings.AlwaysOnTop());
	UISetCheck(ID_OPTIONS_SINGLEINSTANCE, settings.SingleInstance());

	if (settings.AlwaysOnTop())
		SetWindowPos(HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

	UpdateLayout();
//	UpdateUI();
	InitTree();

	return 0;
}

LRESULT CMainFrame::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
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
	auto cb = reinterpret_cast<IObjectsCallback*>(lp);
	ATLASSERT(cb);

	auto count = cb->GetObjectCount();
	for(auto i = 0; i < count; i++) {
		WmiItem item;
		item.Type = NodeType::Instance;
		item.Object = cb->GetItem(i);
		CComBSTR text;
		item.Object->GetObjectText(0, &text);
		item.Name = text;
		m_Items.push_back(std::move(item));
	}
	cb->Release();
	RefreshList();

	return 0;
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
	auto hItem = tv->itemNew.hItem;
	CString text;
	if (m_Tree.GetItemText(m_Tree.GetChildItem(hItem), text) && text != L"\\\\")
		return 0;

	auto path = GetFullPath(hItem);
	CComPtr<IWbemServices> spNamespace;
	auto hr = m_spWmi->OpenNamespace(CComBSTR(path), 0, nullptr, &spNamespace, nullptr);
	if (SUCCEEDED(hr)) {
		CWaitCursor wait;
		m_NamespacePath = L"ROOT\\" + path;
		m_Tree.DeleteItem(m_Tree.GetChildItem(hItem));
		m_spCurrentNamespace = spNamespace;
		BuildTree(spNamespace, hItem);
	}
	return 0;
}

LRESULT CMainFrame::OnTreeSelChanged(int, LPNMHDR, BOOL&) {
	SetTimer(2, 200, nullptr);
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

CString CMainFrame::CimTypeToString(CIMTYPE type) {
	CString text;
	switch (type & 0xff) {
		case CIM_EMPTY: text = L"Empty"; break;
		case CIM_SINT8: text = L"Signed Byte (8 bit)"; break;
		case CIM_UINT8: text = L"Byte (8 bit)"; break;
		case CIM_SINT16: text = L"Signed Word (16 bit)"; break;
		case CIM_UINT16: text = L"Word (16 bit)"; break;
		case CIM_SINT32: text = L"Signed Int (32 bit)"; break;
		case CIM_UINT32: text = L"Int (32 bit)"; break;
		case CIM_SINT64: text = L"Signed QWord (64 bit)"; break;
		case CIM_UINT64: text = L"QWord (64 bit)"; break;
		case CIM_REAL32: text = L"Real (32 bit)"; break;
		case CIM_REAL64: text = L"Real (64 bit)"; break;
		case CIM_BOOLEAN: text = L"Boolean"; break;
		case CIM_STRING: text = L"String"; break;
		case CIM_DATETIME: text = L"Date Time"; break;
		case CIM_REFERENCE: text = L"Reference"; break;
		case CIM_CHAR16: text = L"Character"; break;
		case CIM_OBJECT: text = L"Object"; break;
	}
	if (type & CIM_FLAG_ARRAY)
		text += L" [Array]";
	return text;
}

CString CMainFrame::GetArrayValue(CComVariant& value, CIMTYPE type) {
	CString text;
	switch (type & 0xff) {
		case CIM_STRING:
		{
			// string array
			CComSafeArray<BSTR> arr(value.parray);
			auto count = arr.GetCount();
			for (ULONG i = 0; i < count; i++)
				text += CString(arr.GetAt(i).m_str) + L", ";
			if (!text.IsEmpty())
				text = text.Left(text.GetLength() - 2);
			break;
		}

		case CIM_SINT8:
		case CIM_UINT8:
		{
			BYTE* data;
			CComSafeArray<BYTE> arr(value.parray);
			auto count = std::min(arr.GetCount(), (ULONG)64);
			if (SUCCEEDED(::SafeArrayAccessData(value.parray, reinterpret_cast<void**>(&data)))) {
				for (ULONG i = 0; i < count; i++) {
					CString str;
					str.Format(L"%02X ", data[i]);
					text += str;
				}
				::SafeArrayUnaccessData(value.parray);
			}
			break;
		}
	}
	return text;
}

void CMainFrame::InitCommandBar() {
	struct {
		UINT id, icon;
		HICON hIcon = nullptr;
	} cmds[] = {
		{ ID_FILE_RUNASADMINISTRATOR, 0, IconHelper::GetShieldIcon() },
		{ ID_EDIT_COPY, IDI_COPY },
		//{ ID_VIEW_REFRESH, IDI_REFRESH },
		//{ ID_EDIT_CUT, IDI_CUT },
		//{ ID_EDIT_DELETE, IDI_DELETE },
		//{ ID_EDIT_FIND, IDI_FIND },
	};
	for (auto& cmd : cmds) {
		HICON hIcon = cmd.hIcon;
		if (!hIcon) {
			hIcon = AtlLoadIconImage(cmd.icon, 0, 16, 16);
			ATLASSERT(hIcon);
		}
		m_CmdBar.AddIcon(cmd.icon ? hIcon : cmd.hIcon, cmd.id);
	}
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
	WMIHelper::Init(nullptr, L"ROOT", &m_spWmi);
	m_spCurrentNamespace = m_spWmi;
	m_Tree.LockWindowUpdate();
	m_hRoot = InsertTreeItem(L"ROOT", 0, TVI_ROOT, NodeType::Namespace);
	if (m_spWmi) {
		BuildTree(m_spWmi, m_hRoot);
		m_Tree.Expand(m_hRoot, TVE_EXPAND);
	}
	m_Tree.LockWindowUpdate(FALSE);
	m_Tree.SelectItem(m_hRoot);
	m_Tree.SetFocus();
}

void CMainFrame::BuildTree(IWbemServices* pWmi, HTREEITEM hParent) {
	auto classes = WMIHelper::EnumClasses(pWmi, true);
	for (auto& spObj : classes) {
		auto name = WMIHelper::GetStringProperty(spObj, L"__CLASS");
		auto hItem = InsertTreeItem(name, 1, hParent, NodeType::Class);
	}

	auto ns = WMIHelper::EnumNamespaces(pWmi);
	for (auto& spObj : ns) {
		auto name = WMIHelper::GetStringProperty(spObj, L"NAME");
		ATLTRACE(L"Namespace: %s\n", (PCWSTR)name);
		auto hItem = InsertTreeItem(name, 0, hParent, NodeType::Namespace);
		CComPtr<IWbemServices> spNamespace;
		auto hr = pWmi->OpenNamespace(CComBSTR(name), 0, nullptr, &spNamespace, nullptr);
		if (FAILED(hr))
			continue;

		if (IsChildNamespaceOrClass(spNamespace))
			InsertTreeItem(L"\\\\", 0, hItem, NodeType::HasChildren);
	}
	m_Tree.SortChildren(hParent);
}

bool CMainFrame::IsChildNamespaceOrClass(IWbemServices* pWmi) const {
	CComPtr<IEnumWbemClassObject> spEnum;
	pWmi->CreateInstanceEnum(CComBSTR(L"__NAMESPACE"), 0, nullptr, &spEnum);
	return spEnum != nullptr;
}

void CMainFrame::UpdateList() {
	m_List.SetItemCount(0);
	m_Items.clear();

	if (m_spCurrentClass) {
		CString name;
		m_Tree.GetItemText(m_Tree.GetSelectedItem(), name);
		for (auto& prop : WMIHelper::EnumProperties(m_spCurrentClass)) {
			WmiItem item;
			item.Name = prop.Name;
			item.Type = NodeType::Property;
			item.CimType = prop.Type;
			item.Value = prop.Value;
			m_Items.push_back(std::move(item));
		}
		for (auto& method : WMIHelper::EnumMethods(m_spCurrentClass)) {
			WmiItem item;
			item.Name = method.Name;
			item.Type = NodeType::Method;
			item.Object = method.spInParams;
			item.Object2 = method.spOutParams;
			item.Value = method.ClassName;
			m_Items.push_back(std::move(item));
		}
		WMIHelper::EnumInstancesAsync(m_hWnd, WM_INSTANCES, name, m_spCurrentNamespace, false);
		//for (auto& inst : WMIHelper::EnumInstances(name, m_spCurrentNamespace, false)) {
		//	WmiItem item;
		//	item.Type = NodeType::Instance;
		//	item.Object = inst;
		//	CComBSTR text;
		//	inst->GetObjectText(0, &text);
		//	item.Name = text;
		//	m_Items.push_back(std::move(item));
		//}
	}
	else {
		if (m_spCurrentNamespace == nullptr)
			return;

		for (auto& ns : WMIHelper::EnumNamespaces(m_spCurrentNamespace)) {
			WmiItem item;
			CComBSTR name;
			item.Name = WMIHelper::GetStringProperty(ns, L"NAME");
			item.Object = ns;
			item.Type = NodeType::Namespace;
			m_Items.push_back(std::move(item));
		}

		for (auto& cls : WMIHelper::EnumClasses(m_spCurrentNamespace, true)) {
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

CString CMainFrame::GetObjectDetails(WmiItem const& item) const {
	switch (item.Type) {
		case NodeType::Method:
			return L"Class: " + CString(item.Value.bstrVal);
	}
	return L"";
}

CString CMainFrame::GetObjectValue(WmiItem const& item) const {
	switch (item.Type) {
		case NodeType::Property:
			CComVariant value(item.Value);
			if (value.vt == VT_NULL)
				return L"";

			if ((item.CimType & CIM_FLAG_ARRAY) && value.parray) {
				return GetArrayValue(value, item.CimType);
			}
			if (value.vt == VT_BOOL)
				return value.boolVal ? L"True" : L"False";
			if(SUCCEEDED(value.ChangeType(VT_BSTR)))
				return CString(value.bstrVal);
			break;
	}
	return L"";
}

void CMainFrame::TreeItemSelected(HTREEITEM /* hItem */) {
	auto hItem = m_Tree.GetSelectedItem();
	if (hItem == nullptr) {
		m_spCurrentClass = nullptr;
		m_spCurrentNamespace = nullptr;
		return;
	}
	auto path = GetFullPath(hItem);
	auto type = GetTreeNodeType(hItem);
	CString name;
	m_Tree.GetItemText(hItem, name);
	switch (type) {
		case NodeType::Namespace:
		{
			if (hItem == m_hRoot) {
				m_NamespacePath = L"ROOT";
				m_spCurrentNamespace = m_spWmi;
			}
			else {
				CComPtr<IWbemServices> spNamespace;
				m_spWmi->OpenNamespace(CComBSTR(path), 0, nullptr, &spNamespace, nullptr);
				if (spNamespace) {
					m_spCurrentNamespace = spNamespace;
					m_NamespacePath = L"ROOT\\" + path;
				}
			}
			m_spCurrentClass = nullptr;
			break;
		}
		case NodeType::Class:
			m_spCurrentClass = nullptr;
			m_spCurrentNamespace->GetObject(CComBSTR(name), 0, nullptr, &m_spCurrentClass, nullptr);
			break;

		default:
			ATLASSERT(false);
			return;
	}

	UpdateList();
}

void CMainFrame::RefreshList() {
	m_List.SetItemCountEx(static_cast<int>(m_Items.size()), LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
	m_List.RedrawItems(m_List.GetTopIndex(), m_List.GetTopIndex() + m_List.GetCountPerPage());
}

CString CMainFrame::GetFullPath(HTREEITEM hItem) const {
	CString path, text;
	while (hItem) {
		m_Tree.GetItemText(hItem, text);
		if (text == L"ROOT")
			break;
		path = text + L"\\" + path;
		hItem = m_Tree.GetParentItem(hItem);
	}
	return path.TrimRight(L"\\");
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
