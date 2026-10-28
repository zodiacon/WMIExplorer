#pragma once

#include <VirtualListView.h>
#include "DialogHelper.h"
#include "resource.h"

//
// the properties of an object (an instance, a class or an embedded object), or the elements of an array value.
// double-clicking a reference opens the object it points to; an embedded object or an array opens in another one
//
class CObjectDlg :
	public CDialogImpl<CObjectDlg>,
	public CDialogResize<CObjectDlg>,
	public CDialogHelper<CObjectDlg>,
	public CVirtualListView<CObjectDlg> {
public:
	enum { IDD = IDD_OBJECT };

	// pSvc: the object's namespace (references can be relative to it)
	CObjectDlg(IWbemServices* pSvc, IWbemClassObject* pObj);
	CObjectDlg(IWbemServices* pSvc, CString const& name, CComVariant const& array, CIMTYPE type);

	// a reference (non-empty), an embedded object or an array: what ShowValue opens
	static bool CanShowValue(CComVariant const& value, CIMTYPE type);
	// opens a value in an object dialog (a reference: the object it points to). false if it can't be opened
	static bool ShowValue(HWND hParent, IWbemServices* pSvc, CString const& name, CComVariant const& value, CIMTYPE type);
	// the object of a path (an error message if it can't be read)
	static void ShowReference(HWND hParent, IWbemServices* pSvc, CString const& path);

	CString GetColumnText(HWND, int row, int col) const;
	int GetRowImage(HWND, int row, int) const;
	void DoSort(const SortInfo* si);
	bool OnDoubleClickList(HWND, int row, int col, POINT const& pt);

	BEGIN_MSG_MAP(CObjectDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDC_COPY, OnCopy)
		COMMAND_ID_HANDLER(IDC_MOF, OnShowMof)
		COMMAND_ID_HANDLER(IDCANCEL, OnCloseCmd)
		CHAIN_MSG_MAP(CVirtualListView<CObjectDlg>)
		CHAIN_MSG_MAP(CDialogResize<CObjectDlg>)
	END_MSG_MAP()

	BEGIN_DLGRESIZE_MAP(CObjectDlg)
		DLGRESIZE_CONTROL(IDC_RESULTS, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_COPY, DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDC_MOF, DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDC_STATUS, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

private:
	enum class ColumnType {
		Name, Type, Value
	};

	struct Row {
		CString Name;
		CComVariant Value;
		CIMTYPE Type;
		int Index;		// the position in the object or array (the order of "Name" for an array)
		bool System{ false };
	};

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnCopy(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnShowMof(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCloseCmd(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	CComPtr<IWbemServices> m_spSvc;
	CComPtr<IWbemClassObject> m_spObj;		// null: an array
	CString m_Title;
	std::vector<Row> m_Rows;
	CListViewCtrl m_List;
};
