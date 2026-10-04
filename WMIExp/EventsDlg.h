#pragma once

#include <VirtualListView.h>
#include "DialogHelper.h"
#include "resource.h"
#include <thread>

struct WMIConnection;

//
// modeless: shows the events of an event query (ExecNotificationQuery) as they arrive
//
class CEventsDlg :
	public CDialogImpl<CEventsDlg>,
	public CDialogResize<CEventsDlg>,
	public CDialogHelper<CEventsDlg>,
	public CVirtualListView<CEventsDlg> {
public:
	enum { IDD = IDD_EVENTS };

	void Activate(HWND hOwner);
	void Reset();

	CString GetColumnText(HWND, int row, int col) const;
	int GetRowImage(HWND, int row, int) const;
	void DoSort(const SortInfo* si);
	bool OnDoubleClickList(HWND, int row, int col, POINT const& pt);

	BEGIN_MSG_MAP(CEventsDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_EVENTS, OnEvents)
		COMMAND_ID_HANDLER(IDC_START, OnStart)
		COMMAND_ID_HANDLER(IDC_STOP, OnStop)
		COMMAND_ID_HANDLER(IDC_CLEAR, OnClear)
		COMMAND_ID_HANDLER(IDCANCEL, OnCloseCmd)
		CHAIN_MSG_MAP(CVirtualListView<CEventsDlg>)
		CHAIN_MSG_MAP(CDialogResize<CEventsDlg>)
	END_MSG_MAP()

	BEGIN_DLGRESIZE_MAP(CEventsDlg)
		DLGRESIZE_CONTROL(IDC_NAMESPACE, DLSZ_SIZE_X)
		DLGRESIZE_CONTROL(IDC_QUERY, DLSZ_SIZE_X)
		DLGRESIZE_CONTROL(IDC_START, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_STOP, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_CLEAR, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_RESULTS, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_STATUS, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

private:
	static constexpr UINT WM_EVENTS = WM_APP + 12;

	enum class ColumnType {
		Time, Class, Summary
	};

	struct Event {
		SYSTEMTIME Time;
		CString Class, Summary;
		CComPtr<IWbemClassObject> Object;
	};

	// posted by the listening thread (LPARAM); the receiver deletes it
	struct EventBatch {
		std::vector<Event> Events;
		CString Error;
		int Generation;
		bool Done{ false };
	};

	void Stop();
	void UpdateButtons();
	void UpdateStatus();
	void ListenThread(std::stop_token st, WMIConnection const* conn, CString ns, CString query, int generation);
	static CString GetSummary(IWbemClassObject* pEvent);

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnEvents(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnStart(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnStop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnClear(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCloseCmd(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	CListViewCtrl m_List;
	std::vector<Event> m_Events;
	std::jthread m_Thread;
	int m_Generation{ 0 };
	bool m_Listening{ false };
	CString m_Error;
};
