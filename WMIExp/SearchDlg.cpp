#include "pch.h"
#include "SearchDlg.h"
#include "WMIHelper.h"
#include "AppSettings.h"

void CSearchDlg::Reset() {
	if (!IsWindow())
		return;
	StopSearch();
	++m_Generation;
	m_Searching = false;
	m_Results.clear();
	m_List.SetItemCount(0);
	UpdateButtons();
	SetDlgItemText(IDC_STATUS, L"");
}

void CSearchDlg::Activate(HWND hOwner) {
	if (!IsWindow())
		Create(hOwner);
	ShowWindow(SW_SHOW);
	SetActiveWindow();
	GotoDlgCtrl(GetDlgItem(IDC_TEXT));
	SendDlgItemMessage(IDC_TEXT, EM_SETSEL, 0, -1);
}

CString CSearchDlg::GetColumnText(HWND, int row, int col) const {
	auto& item = m_Results[row];
	switch (GetColumnManager(m_List)->GetColumnTag<ColumnType>(col)) {
		case ColumnType::Name: return item.Name;
		case ColumnType::Type: return KindToText(item.Type);
		case ColumnType::Namespace: return item.Namespace;
		case ColumnType::Class: return item.Class;
	}
	return L"";
}

int CSearchDlg::GetRowImage(HWND, int row, int) const {
	return static_cast<int>(m_Results[row].Type);
}

void CSearchDlg::DoSort(const SortInfo* si) {
	auto column = GetColumnManager(m_List)->GetColumnTag<ColumnType>(si->SortColumn);
	auto compare = [&](SearchResult const& r1, SearchResult const& r2) {
		switch (column) {
			case ColumnType::Name: return r1.Name.CompareNoCase(r2.Name);
			case ColumnType::Type: return static_cast<int>(r1.Type) - static_cast<int>(r2.Type);
			case ColumnType::Namespace: return r1.Namespace.CompareNoCase(r2.Namespace);
			case ColumnType::Class: return r1.Class.CompareNoCase(r2.Class);
		}
		return 0;
	};
	std::ranges::stable_sort(m_Results, [&](auto const& r1, auto const& r2) {
		auto result = compare(r1, r2);
		return si->SortAscending ? result < 0 : result > 0;
		});
}

bool CSearchDlg::OnDoubleClickList(HWND, int row, int, POINT const&) {
	if (row < 0 || row >= (int)m_Results.size())
		return false;

	if (!m_Navigator->NavigateTo(m_Results[row]))
		SetDlgItemText(IDC_STATUS, L"Not found in the main view (it may be hidden by the View options, or the tree is out of date: press F5)");
	return true;
}

PCWSTR CSearchDlg::KindToText(SearchResult::Kind kind) {
	switch (kind) {
		case SearchResult::Kind::Namespace: return L"Namespace";
		case SearchResult::Kind::Class: return L"Class";
		case SearchResult::Kind::Property: return L"Property";
		case SearchResult::Kind::Method: return L"Method";
	}
	return L"";
}

bool CSearchDlg::Matches(SearchOptions const& options, PCWSTR text) {
	return options.MatchCase ? wcsstr(text, options.Text) != nullptr : ::StrStrIW(text, options.Text) != nullptr;
}

void CSearchDlg::StartSearch() {
	SearchOptions options;
	GetDlgItemText(IDC_TEXT, options.Text);
	options.Text.Trim();
	if (options.Text.IsEmpty())
		return;

	options.Namespaces = IsDlgButtonChecked(IDC_NAMESPACES) == BST_CHECKED;
	options.Classes = IsDlgButtonChecked(IDC_CLASSES) == BST_CHECKED;
	options.Properties = IsDlgButtonChecked(IDC_PROPERTIES) == BST_CHECKED;
	options.Methods = IsDlgButtonChecked(IDC_METHODS) == BST_CHECKED;
	options.MatchCase = IsDlgButtonChecked(IDC_MATCHCASE) == BST_CHECKED;
	if (!options.Namespaces && !options.Classes && !options.Properties && !options.Methods) {
		SetDlgItemText(IDC_STATUS, L"Select what to search for");
		return;
	}
	// what is hidden in the main view is not searched (it could not be shown there)
	options.SystemClasses = AppSettings::Get().ViewSystemClasses();
	options.SystemProperties = AppSettings::Get().ViewSystemProperties();

	StopSearch();
	m_Results.clear();
	m_List.SetItemCount(0);
	m_Searching = true;
	m_StopRequested = false;
	UpdateButtons();
	SetDlgItemText(IDC_STATUS, L"Searching...");

	auto generation = ++m_Generation;
	auto conn = WMIHelper::CurrentConnection();
	m_Thread = std::jthread([this, conn, options, generation](std::stop_token st) {
		SearchThread(st, conn, options, generation);
		});
}

void CSearchDlg::StopSearch() {
	if (m_Thread.joinable()) {
		m_Thread.request_stop();
		m_Thread.join();
	}
}

void CSearchDlg::UpdateButtons() {
	GetDlgItem(IDC_SEARCH).EnableWindow(!m_Searching);
	GetDlgItem(IDC_STOP).EnableWindow(m_Searching && !m_StopRequested);
}

bool CSearchDlg::Post(SearchProgress* progress) {
	if (!PostMessage(WM_SEARCH_PROGRESS, 0, reinterpret_cast<LPARAM>(progress))) {
		delete progress;
		return false;
	}
	return true;
}

//
// runs on its own (MTA) thread, with its own connection to WMI
//
void CSearchDlg::SearchThread(std::stop_token st, WMIConnection const* conn, SearchOptions options, int generation) {
	::CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	auto done = new SearchProgress;
	done->Generation = generation;
	done->Done = true;
	{
		CComPtr<IWbemServices> spRoot;
		auto hr = WMIHelper::Connect(conn, L"ROOT", &spRoot);
		if (SUCCEEDED(hr))
			SearchNamespace(st, options, spRoot, L"ROOT", generation);
		else
			done->Error = WMIHelper::GetErrorText(hr);
	}
	Post(done);
	::CoUninitialize();
}

void CSearchDlg::SearchNamespace(std::stop_token const& st, SearchOptions const& options, IWbemServices* pSvc, CString const& path, int generation) {
	if (st.stop_requested())
		return;

	auto starting = new SearchProgress;
	starting->Generation = generation;
	starting->Namespace = path;
	Post(starting);

	auto progress = new SearchProgress;
	progress->Generation = generation;
	if (options.Classes || options.Properties || options.Methods) {
		for (auto& spClass : WMIHelper::EnumClasses(pSvc, true, options.SystemClasses)) {
			if (st.stop_requested())
				break;

			auto name = WMIHelper::GetStringProperty(spClass, L"__CLASS");
			if (options.Classes && Matches(options, name))
				progress->Results.push_back({ SearchResult::Kind::Class, name, path });

			if (options.Properties) {
				// local properties only: an inherited property is found in the class that defines it
				spClass->BeginEnumeration(WBEM_FLAG_LOCAL_ONLY);
				BSTR prop;
				CIMTYPE type;
				while (S_OK == spClass->Next(0, &prop, nullptr, &type, nullptr)) {
					CComBSTR name2;
					name2.Attach(prop);
					if ((options.SystemProperties || wcsncmp(name2, L"__", 2) != 0) && Matches(options, name2))
						progress->Results.push_back({ SearchResult::Kind::Property, CString(name2), path, name, type });
				}
				spClass->EndEnumeration();
			}
			if (options.Methods) {
				for (auto& method : WMIHelper::EnumMethods(spClass, true)) {
					if (Matches(options, method.Name.c_str()))
						progress->Results.push_back({ SearchResult::Kind::Method, method.Name.c_str(), path, name });
				}
			}
		}
	}

	std::vector<CString> children;
	for (auto& spNamespace : WMIHelper::EnumNamespaces(pSvc)) {
		auto name = WMIHelper::GetStringProperty(spNamespace, L"NAME");
		if (options.Namespaces && Matches(options, name))
			progress->Results.push_back({ SearchResult::Kind::Namespace, name, path });
		children.push_back(name);
	}
	if (progress->Results.empty())
		delete progress;
	else
		Post(progress);

	for (auto& name : children) {
		if (st.stop_requested())
			return;

		CComPtr<IWbemServices> spChild;
		if (SUCCEEDED(WMIHelper::OpenNamespace(pSvc, name, &spChild)))
			SearchNamespace(st, options, spChild, path + L"\\" + name, generation);
	}
}

LRESULT CSearchDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	DlgResize_Init(true, true, WS_THICKFRAME | WS_CLIPCHILDREN);
	SetDialogIcon(IDR_MAINFRAME);

	m_List.Attach(GetDlgItem(IDC_RESULTS));
	m_List.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);

	CImageList images;
	images.Create(16, 16, ILC_COLOR32 | ILC_MASK, 4, 1);
	// in the order of SearchResult::Kind
	for (auto icon : { IDI_NAMESPACE, IDI_CLASS, IDI_PROPERTY, IDI_METHOD })
		images.AddIcon(AtlLoadIconImage(icon, 0, 16, 16));
	m_List.SetImageList(images, LVSIL_SMALL);

	auto cm = GetColumnManager(m_List);
	cm->AddColumn(L"Name", LVCFMT_LEFT, 220, ColumnType::Name);
	cm->AddColumn(L"Type", LVCFMT_LEFT, 80, ColumnType::Type);
	cm->AddColumn(L"Namespace", LVCFMT_LEFT, 200, ColumnType::Namespace);
	cm->AddColumn(L"Class", LVCFMT_LEFT, 200, ColumnType::Class);

	CheckDlgButton(IDC_NAMESPACES, BST_CHECKED);
	CheckDlgButton(IDC_CLASSES, BST_CHECKED);
	CheckDlgButton(IDC_PROPERTIES, BST_CHECKED);
	CheckDlgButton(IDC_METHODS, BST_CHECKED);
	UpdateButtons();

	return TRUE;
}

LRESULT CSearchDlg::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled) {
	StopSearch();
	// progress still in the queue would be lost with the window
	MSG msg;
	while (::PeekMessage(&msg, m_hWnd, WM_SEARCH_PROGRESS, WM_SEARCH_PROGRESS, PM_REMOVE))
		delete reinterpret_cast<SearchProgress*>(msg.lParam);

	bHandled = FALSE;
	return 0;
}

LRESULT CSearchDlg::OnSearchProgress(UINT, WPARAM, LPARAM lParam, BOOL&) {
	std::unique_ptr<SearchProgress> progress(reinterpret_cast<SearchProgress*>(lParam));
	if (progress->Generation != m_Generation)
		return 0;		// from a search that was since replaced

	if (!progress->Results.empty()) {
		m_Results.insert(m_Results.end(), std::make_move_iterator(progress->Results.begin()), std::make_move_iterator(progress->Results.end()));
		m_List.SetItemCountEx((int)m_Results.size(), LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
	}

	if (progress->Done) {
		m_Searching = false;
		UpdateButtons();
		Sort(m_List);
		CString status;
		if (!progress->Error.IsEmpty())
			status = L"Error: " + progress->Error;
		else
			status = std::format(L"{}{} item(s) found", m_StopRequested ? L"Stopped. " : L"", m_Results.size()).c_str();
		SetDlgItemText(IDC_STATUS, status);
	}
	else if (!progress->Namespace.IsEmpty()) {
		SetDlgItemText(IDC_STATUS, std::format(L"Searching {}... ({} found)", (PCWSTR)progress->Namespace, m_Results.size()).c_str());
	}
	return 0;
}

LRESULT CSearchDlg::OnSearch(WORD, WORD, HWND, BOOL&) {
	StartSearch();
	return 0;
}

LRESULT CSearchDlg::OnStop(WORD, WORD, HWND, BOOL&) {
	// the search thread notices soon, and reports when it is done
	m_Thread.request_stop();
	m_StopRequested = true;
	UpdateButtons();
	return 0;
}

LRESULT CSearchDlg::OnCloseCmd(WORD, WORD, HWND, BOOL&) {
	// hidden, not destroyed: the results are there when it is shown again
	ShowWindow(SW_HIDE);
	return 0;
}
