// MainFrm.h : interface of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once

#include <VirtualListView.h>
#include "WMIHelper.h"
#include <CustomSplitterWindow.h>
#include <TreeViewHelper.h>
#include "SearchDlg.h"
#include "QueryDlg.h"
#include "EventsDlg.h"

class CMainFrame :
	public CFrameWindowImpl<CMainFrame>,
	public CAutoUpdateUI<CMainFrame>,
	public CVirtualListView<CMainFrame>,
	public CTreeViewHelper<CMainFrame>,
	public CMessageFilter,
	public CIdleHandler,
	public ISearchNavigator {
public:
	DECLARE_FRAME_WND_CLASS(L"WMIEXPWNDCLASS", IDR_MAINFRAME)

	// ISearchNavigator
	bool NavigateTo(SearchResult const& result) override;

	const UINT WM_INSTANCES = WM_APP + 6;

	enum { TreeId = 123, ListId };

	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnIdle();

	CString GetColumnText(HWND, int row, int col) const;
	int GetRowImage(HWND, int row, int) const;

	void OnStateChanged(HWND h, int from, int to, UINT oldState, UINT newState);
	void DoSort(const SortInfo* si);
	bool OnDoubleClickList(HWND h, int row, int col, POINT const& pt);

	BEGIN_MSG_MAP(CMainFrame)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)
		NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDING, OnTreeItemExpanding)
		NOTIFY_CODE_HANDLER(TVN_SELCHANGED, OnTreeSelChanged)
		NOTIFY_CODE_HANDLER(TVN_GETINFOTIP, OnTreeGetInfoTip)
		MESSAGE_HANDLER(WM_INSTANCES, OnAddInstances)
		COMMAND_ID_HANDLER(ID_VIEW_SYSTEMCLASSES, OnViewSystemClasses)
		COMMAND_ID_HANDLER(ID_VIEW_SYSTEMPROPERTIES, OnViewSystemProperties)
		COMMAND_ID_HANDLER(ID_VIEW_NAMESPACESINLIST, OnViewNamespacesInList)
		COMMAND_ID_HANDLER(ID_VIEW_DERIVEDINSTANCES, OnViewDerivedInstances)
		COMMAND_ID_HANDLER(ID_VIEW_CLASSHIERARCHY, OnViewClassHierarchy)
		COMMAND_ID_HANDLER(ID_FILE_CONNECT, OnConnect)
		COMMAND_ID_HANDLER(ID_TOOLS_QUERY, OnQuery)
		COMMAND_ID_HANDLER(ID_TOOLS_EVENTS, OnEvents)
		COMMAND_ID_HANDLER(ID_TOOLS_EXECUTEMETHOD, OnExecuteMethod)
		COMMAND_ID_HANDLER(ID_TOOLS_SHOWMOF, OnShowMof)
		COMMAND_ID_HANDLER(ID_VIEW_REFRESH, OnViewRefresh)
		COMMAND_ID_HANDLER(ID_EDIT_COPY, OnEditCopy)
		COMMAND_ID_HANDLER(ID_EDIT_FIND, OnEditFind)
		COMMAND_ID_HANDLER(ID_OPTIONS_SINGLEINSTANCE, OnSingleInstance)
		COMMAND_ID_HANDLER(ID_APP_EXIT, OnFileExit)
		COMMAND_ID_HANDLER(ID_VIEW_TOOLBAR, OnViewToolBar)
		COMMAND_ID_HANDLER(ID_OPTIONS_ALWAYSONTOP, OnAlwaysOnTop)
		COMMAND_ID_HANDLER(ID_OPTIONS_DARKMODE, OnToggleDarkMode)
		COMMAND_ID_HANDLER(ID_VIEW_STATUS_BAR, OnViewStatusBar)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		MESSAGE_HANDLER(WM_SHOWWINDOW, OnShowWindow)
		COMMAND_ID_HANDLER(ID_FILE_RUNASADMINISTRATOR, OnRunAsAdmin)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		CHAIN_MSG_MAP(CAutoUpdateUI<CMainFrame>)
		CHAIN_MSG_MAP(CVirtualListView<CMainFrame>)
		CHAIN_MSG_MAP(CFrameWindowImpl<CMainFrame>)
	END_MSG_MAP()

	// Handler prototypes (uncomment arguments if needed):
	//	LRESULT MessageHandler(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	//	LRESULT CommandHandler(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
	//	LRESULT NotifyHandler(int /*idCtrl*/, LPNMHDR /*pnmh*/, BOOL& /*bHandled*/)

private:
	enum class ColumnType {
		Name, Value, Type, Size, CimType, Details, Description
	};
	enum class NodeType {
		Computer, Namespace, Class, Property, Method, Instance, HasChildren = 0x80
	};
	struct WmiItem {
		std::wstring Name;
		wil::com_ptr<IWbemClassObject> Object, Object2;
		CIMTYPE CimType;
		NodeType Type;
		CComVariant Value;
		// loaded when first shown (for classes in a namespace's list: it needs another call to WMI)
		mutable CString Description;
		mutable bool DescriptionLoaded{ false };
	};

	struct WmiObject {
		wil::com_ptr<IWbemClassObject> Object;
	};

	// a column of the instance list: a property of the selected class
	struct InstanceColumn {
		CString Name;
		CIMTYPE Type;
	};

	static PCWSTR NodeTypeToText(NodeType type);
	static CString FlattenText(CString text);

	void InitMenu(HMENU menu);
	void InitToolBar(CToolBarCtrl& tb, int size = 24);
	void InitTree();
	void BuildTree(IWbemServices* pWmi, HTREEITEM hParent);
	void LoadNamespaceChildren(HTREEITEM hItem);
	HTREEITEM FindChildItem(HTREEITEM hParent, CString const& name, NodeType type);
	HTREEITEM FindClassItem(HTREEITEM hNamespace, CString const& name);
	HTREEITEM GetNamespaceItem(HTREEITEM hItem) const;
	CString GetClassDescription(CString const& nsPath, CString const& className);
	void ExecuteMethod(WmiItem const& method);
	void UpdateTitle();
	void UpdateList();
	void CancelInstanceEnum();
	void BuildInstanceColumns();
	void ClearInstanceColumns();
	void SortInstances(const SortInfo* si);
	void SortItemsByValue(const SortInfo* si, ColumnType column);
	CString GetObjectDetails(WmiItem const& item) const;
	CString GetObjectValue(WmiItem const& item) const;
	void TreeItemSelected(HTREEITEM hItem);
	void RefreshList();

	HTREEITEM InsertTreeItem(PCWSTR text, int image, HTREEITEM hParent, NodeType type);
	NodeType GetTreeNodeType(HTREEITEM hItem) const;
	void SetAlwaysOnTop(bool onTop);

	LRESULT OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAddInstances(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnFileExit(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewToolBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewStatusBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnTreeItemExpanding(int /*idCtrl*/, LPNMHDR /*pnmh*/, BOOL& /*bHandled*/);
	LRESULT OnTreeSelChanged(int /*idCtrl*/, LPNMHDR /*pnmh*/, BOOL& /*bHandled*/);
	LRESULT OnViewSystemClasses(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewSystemProperties(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewNamespacesInList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewDerivedInstances(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewRefresh(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEditCopy(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEditFind(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSingleInstance(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnShowWindow(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnRunAsAdmin(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlwaysOnTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnToggleDarkMode(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewClassHierarchy(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnConnect(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnQuery(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEvents(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnExecuteMethod(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnShowMof(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnTreeGetInfoTip(int /*idCtrl*/, LPNMHDR /*pnmh*/, BOOL& /*bHandled*/);

	CCustomSplitterWindow m_Splitter;
	CCustomHorSplitterWindow m_DetailSplitter;
	CTreeViewCtrlEx m_Tree;
	CPaneContainer m_TreePane;
	CListViewCtrl m_List;
	CListViewCtrl m_InstanceList;
	CMultiPaneStatusBarCtrl m_StatusBar;
	std::vector<WmiItem> m_Items;
	std::vector<WmiItem> m_Objects;
	std::vector<InstanceColumn> m_InstanceColumns;
	std::vector<WMIProperty> m_ObjPropValues;
	HANDLE m_hSingleInstMutex;
	HTREEITEM m_hRoot;
	CSearchDlg m_SearchDlg{ this };
	CQueryDlg m_QueryDlg;
	CEventsDlg m_EventsDlg;
	CString m_NamespacePath;
	CComPtr<IWbemServices> m_spWmi;
	CComPtr<IWbemServices> m_spCurrentNamespace;
	CComPtr<IWbemClassObject> m_spCurrentClass;
	// the instance enumeration in progress (if any)
	std::shared_ptr<WMIQueryJob> m_EnumJob;
	// class descriptions for the tree's tooltips (key: namespace path:class), and the namespace of the last one
	std::map<CString, CString> m_ClassDescriptions;
	CComPtr<IWbemServices> m_spTipNamespace;
	CString m_TipNamespacePath;
	const CString m_RootName{ L"ROOT" };
};
