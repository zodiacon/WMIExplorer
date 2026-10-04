#pragma once

#include <VirtualListView.h>
#include "DialogHelper.h"
#include "resource.h"

struct WMIQueryJob;

//
// modeless: runs WQL queries (on a worker thread); a column for each property of the results
//
class CQueryDlg :
	public CDialogImpl<CQueryDlg>,
	public CDialogResize<CQueryDlg>,
	public CDialogHelper<CQueryDlg>,
	public CVirtualListView<CQueryDlg> {
public:
	enum { IDD = IDD_QUERY };

	// the namespace and query are filled in (if the query box is empty)
	void Activate(HWND hOwner, CString const& ns, CString const& className);
	void Reset();

	CString GetColumnText(HWND, int row, int col) const;
	int GetRowImage(HWND, int row, int) const;
	void DoSort(const SortInfo* si);
	bool OnDoubleClickList(HWND, int row, int col, POINT const& pt);

	BEGIN_MSG_MAP(CQueryDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_QUERY_DONE, OnQueryDone)
		COMMAND_ID_HANDLER(IDC_RUN, OnRun)
		COMMAND_ID_HANDLER(IDC_STOP, OnStop)
		COMMAND_ID_HANDLER(IDCANCEL, OnCloseCmd)
		CHAIN_MSG_MAP(CVirtualListView<CQueryDlg>)
		CHAIN_MSG_MAP(CDialogResize<CQueryDlg>)
	END_MSG_MAP()

	BEGIN_DLGRESIZE_MAP(CQueryDlg)
		DLGRESIZE_CONTROL(IDC_NAMESPACE, DLSZ_SIZE_X)
		DLGRESIZE_CONTROL(IDC_QUERY, DLSZ_SIZE_X)
		DLGRESIZE_CONTROL(IDC_RUN, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_STOP, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_RESULTS, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_STATUS, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

private:
	static constexpr UINT WM_QUERY_DONE = WM_APP + 11;

	struct Column {
		CString Name;
		CIMTYPE Type;
	};

	void BuildColumns();
	void Cancel();
	void UpdateButtons();

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnQueryDone(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnRun(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnStop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCloseCmd(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	CListViewCtrl m_List;
	std::vector<CComPtr<IWbemClassObject>> m_Objects;
	std::vector<Column> m_Columns;
	std::shared_ptr<WMIQueryJob> m_Job;
	DWORD m_StartTime{ 0 };
};
