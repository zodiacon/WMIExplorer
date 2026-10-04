#pragma once

#include <VirtualListView.h>
#include "DialogHelper.h"
#include "resource.h"
#include <thread>

struct SearchResult {
	enum class Kind {
		Namespace, Class, Property, Method
	};
	Kind Type;
	CString Name;
	CString Namespace;		// the namespace the item is in
	CString Class;			// for a property or method: the class that defines it
	CIMTYPE CimType{ 0 };	// for a property
};

struct WMIConnection;

struct ISearchNavigator {
	virtual bool NavigateTo(SearchResult const& result) = 0;
};

//
// modeless: stays open (and keeps its results) while the user navigates in the main window
//
class CSearchDlg :
	public CDialogImpl<CSearchDlg>,
	public CDialogResize<CSearchDlg>,
	public CDialogHelper<CSearchDlg>,
	public CVirtualListView<CSearchDlg> {
public:
	enum { IDD = IDD_SEARCH };

	explicit CSearchDlg(ISearchNavigator* navigator) : m_Navigator(navigator) {}

	void Activate(HWND hOwner);
	// (when connecting to another computer)
	void Reset();

	CString GetColumnText(HWND, int row, int col) const;
	int GetRowImage(HWND, int row, int) const;
	void DoSort(const SortInfo* si);
	bool OnDoubleClickList(HWND, int row, int col, POINT const& pt);

	BEGIN_MSG_MAP(CSearchDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SEARCH_PROGRESS, OnSearchProgress)
		COMMAND_ID_HANDLER(IDC_SEARCH, OnSearch)
		COMMAND_ID_HANDLER(IDC_STOP, OnStop)
		COMMAND_ID_HANDLER(IDCANCEL, OnCloseCmd)
		CHAIN_MSG_MAP(CVirtualListView<CSearchDlg>)
		CHAIN_MSG_MAP(CDialogResize<CSearchDlg>)
	END_MSG_MAP()

	BEGIN_DLGRESIZE_MAP(CSearchDlg)
		DLGRESIZE_CONTROL(IDC_TEXT, DLSZ_SIZE_X)
		DLGRESIZE_CONTROL(IDC_SEARCH, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_STOP, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_RESULTS, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_STATUS, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

private:
	static constexpr UINT WM_SEARCH_PROGRESS = WM_APP + 10;

	enum class ColumnType {
		Name, Type, Namespace, Class
	};

	struct SearchOptions {
		CString Text;
		bool Namespaces, Classes, Properties, Methods;
		bool MatchCase;
		bool SystemClasses, SystemProperties;
	};

	// posted by the search thread (LPARAM); the receiver deletes it
	struct SearchProgress {
		std::vector<SearchResult> Results;
		CString Namespace;		// being searched now
		CString Error;
		int Generation;			// of the search that posted it
		bool Done{ false };
	};

	void StartSearch();
	void StopSearch();
	void UpdateButtons();
	void SearchThread(std::stop_token st, WMIConnection const* conn, SearchOptions options, int generation);
	void SearchNamespace(std::stop_token const& st, SearchOptions const& options, IWbemServices* pSvc, CString const& path, int generation);
	bool Post(SearchProgress* progress);
	static bool Matches(SearchOptions const& options, PCWSTR text);
	static PCWSTR KindToText(SearchResult::Kind kind);

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnSearchProgress(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnSearch(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnStop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCloseCmd(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	ISearchNavigator* m_Navigator;
	CListViewCtrl m_List;
	std::vector<SearchResult> m_Results;
	std::jthread m_Thread;
	int m_Generation{ 0 };
	bool m_Searching{ false };
	bool m_StopRequested{ false };
};
