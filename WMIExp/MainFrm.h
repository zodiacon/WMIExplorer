// MainFrm.h : interface of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once

#include <VirtualListView.h>
#include "WMIHelper.h"
#include <OwnerDrawnMenu.h>
#include <CustomSplitterWindow.h>

class CMainFrame :
	public CFrameWindowImpl<CMainFrame>,
	public CAutoUpdateUI<CMainFrame>,
	public CVirtualListView<CMainFrame>,
	public COwnerDrawnMenu<CMainFrame>,
	public CMessageFilter, 
	public CIdleHandler {
public:
	DECLARE_FRAME_WND_CLASS(L"WMIEXPWNDCLASS", IDR_MAINFRAME)

	const UINT WM_INSTANCES = WM_APP + 6;

	enum { TreeId = 123, ListId };

	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnIdle();

	CString GetColumnText(HWND, int row, int col) const;
	int GetRowImage(HWND, int row, int) const;

	void OnStateChanged(HWND h, int from, int to, UINT oldState, UINT newState);
	void DoSort(const SortInfo* si);

	BEGIN_MSG_MAP(CMainFrame)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)
		NOTIFY_CODE_HANDLER(TVN_ITEMEXPANDING, OnTreeItemExpanding)
		NOTIFY_CODE_HANDLER(TVN_SELCHANGED, OnTreeSelChanged)
		MESSAGE_HANDLER(WM_INSTANCES, OnAddInstances)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		COMMAND_ID_HANDLER(ID_VIEW_SYSTEMCLASSES, OnViewSystemClasses)
		COMMAND_ID_HANDLER(ID_VIEW_SYSTEMPROPERTIES, OnViewSystemProperties)
		COMMAND_ID_HANDLER(ID_VIEW_NAMESPACESINLIST, OnViewNamespacesInList)
		COMMAND_ID_HANDLER(ID_APP_EXIT, OnFileExit)
		COMMAND_ID_HANDLER(ID_VIEW_TOOLBAR, OnViewToolBar)
		COMMAND_ID_HANDLER(ID_VIEW_STATUS_BAR, OnViewStatusBar)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		CHAIN_MSG_MAP(CAutoUpdateUI<CMainFrame>)
		CHAIN_MSG_MAP(CVirtualListView<CMainFrame>)
		CHAIN_MSG_MAP(CFrameWindowImpl<CMainFrame>)
		CHAIN_MSG_MAP(COwnerDrawnMenu<CMainFrame>)
	END_MSG_MAP()

	// Handler prototypes (uncomment arguments if needed):
	//	LRESULT MessageHandler(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	//	LRESULT CommandHandler(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
	//	LRESULT NotifyHandler(int /*idCtrl*/, LPNMHDR /*pnmh*/, BOOL& /*bHandled*/)

private:
	enum class ColumnType {
		Name, Value, Type, Size, CimType, Details
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
	};

	struct WmiObject {
		wil::com_ptr<IWbemClassObject> Object;
	};

	static PCWSTR NodeTypeToText(NodeType type);
	static CString CimTypeToString(CIMTYPE type);
	static CString GetArrayValue(CComVariant& value, CIMTYPE type);

	void InitCommandBar();
	void InitToolBar(CToolBarCtrl& tb, int size = 24);
	void InitTree();
	void BuildTree(IWbemServices* pWmi, HTREEITEM hParent);
	bool IsChildNamespaceOrClass(IWbemServices* pWmi) const;
	void UpdateList();
	CString GetObjectDetails(WmiItem const& item) const;
	CString GetObjectValue(WmiItem const& item) const;
	void TreeItemSelected(HTREEITEM hItem);
	void RefreshList();

	CString GetFullPath(HTREEITEM hItem) const;

	HTREEITEM InsertTreeItem(PCWSTR text, int image, HTREEITEM hParent, NodeType type);
	NodeType GetTreeNodeType(HTREEITEM hItem) const;

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

	CCustomSplitterWindow m_Splitter;
	CCustomHorSplitterWindow m_DetailSplitter;
	CTreeViewCtrlEx m_Tree;
	CPaneContainer m_TreePane;
	CListViewCtrl m_List;
	CListViewCtrl m_InstanceList;
	CMultiPaneStatusBarCtrl m_StatusBar;
	std::vector<WmiItem> m_Items;
	std::vector<WmiItem> m_Objects;
	std::vector<WMIProperty> m_ObjPropValues;
	HANDLE m_hSingleInstMutex;
	HTREEITEM m_hRoot;
	CString m_NamespacePath;
	CComPtr<IWbemServices> m_spWmi;
	CComPtr<IWbemServices> m_spCurrentNamespace;
	CComPtr<IWbemClassObject> m_spCurrentClass;
};
