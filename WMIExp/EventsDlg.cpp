#include "pch.h"
#include "EventsDlg.h"
#include "WMIHelper.h"
#include "TextDlg.h"

void CEventsDlg::Activate(HWND hOwner) {
	if (!IsWindow())
		Create(hOwner);
	ShowWindow(SW_SHOW);
	SetActiveWindow();
	GotoDlgCtrl(GetDlgItem(IDC_QUERY));
}

void CEventsDlg::Reset() {
	if (!IsWindow())
		return;
	Stop();
	m_Events.clear();
	m_List.SetItemCount(0);
	UpdateStatus();
}

CString CEventsDlg::GetColumnText(HWND, int row, int col) const {
	auto& e = m_Events[row];
	switch (GetColumnManager(m_List)->GetColumnTag<ColumnType>(col)) {
		case ColumnType::Time:
			return std::format(L"{:02}:{:02}:{:02}.{:03}", e.Time.wHour, e.Time.wMinute, e.Time.wSecond, e.Time.wMilliseconds).c_str();
		case ColumnType::Class: return e.Class;
		case ColumnType::Summary: return e.Summary;
	}
	return L"";
}

int CEventsDlg::GetRowImage(HWND, int, int) const {
	return 0;
}

void CEventsDlg::DoSort(const SortInfo* si) {
	auto column = GetColumnManager(m_List)->GetColumnTag<ColumnType>(si->SortColumn);
	auto compare = [&](Event const& e1, Event const& e2) -> int {
		switch (column) {
			case ColumnType::Time:
			{
				FILETIME ft1, ft2;
				::SystemTimeToFileTime(&e1.Time, &ft1);
				::SystemTimeToFileTime(&e2.Time, &ft2);
				return ::CompareFileTime(&ft1, &ft2);
			}
			case ColumnType::Class: return e1.Class.CompareNoCase(e2.Class);
			case ColumnType::Summary: return e1.Summary.CompareNoCase(e2.Summary);
		}
		return 0;
	};
	std::ranges::stable_sort(m_Events, [&](auto const& e1, auto const& e2) {
		auto result = compare(e1, e2);
		return si->SortAscending ? result < 0 : result > 0;
		});
}

bool CEventsDlg::OnDoubleClickList(HWND, int row, int, POINT const&) {
	if (row < 0 || row >= (int)m_Events.size())
		return false;

	CTextDlg::ShowObject(m_hWnd, m_Events[row].Object);
	return true;
}

//
// for an instance event: the instance's class and keys; otherwise, the first few properties
//
CString CEventsDlg::GetSummary(IWbemClassObject* pEvent) {
	CComVariant target;
	if (SUCCEEDED(pEvent->Get(L"TargetInstance", 0, &target, nullptr, nullptr)) && target.vt == VT_UNKNOWN && target.punkVal) {
		CComQIPtr<IWbemClassObject> spTarget(target.punkVal);
		if (spTarget) {
			CString text = WMIHelper::GetStringProperty(spTarget, L"__CLASS");
			bool first = true;
			auto add = [&](PCWSTR name, CComVariant const& value, CIMTYPE type) {
				text += std::format(L"{}{}={}", first ? L" " : L", ", name, (PCWSTR)WMIHelper::FormatValue(value, type)).c_str();
				first = false;
			};
			for (auto& prop : WMIHelper::EnumProperties(spTarget, WBEM_FLAG_KEYS_ONLY))
				add(prop.Name, prop.Value, prop.Type);
			CComVariant name;
			if (SUCCEEDED(spTarget->Get(L"Name", 0, &name, nullptr, nullptr)) && name.vt == VT_BSTR && text.Find(L" Name=") < 0)
				add(L"Name", name, CIM_STRING);
			return text;
		}
	}

	CString text;
	int count = 0;
	for (auto& prop : WMIHelper::EnumProperties(pEvent, WBEM_FLAG_NONSYSTEM_ONLY)) {
		if (prop.Value.vt == VT_NULL || prop.Value.vt == VT_EMPTY || (prop.Type & ~CIM_FLAG_ARRAY) == CIM_OBJECT ||
			prop.Name == L"SECURITY_DESCRIPTOR" || prop.Name == L"TIME_CREATED")
			continue;
		if (count++)
			text += L", ";
		text += std::format(L"{}={}", (PCWSTR)prop.Name, (PCWSTR)WMIHelper::FormatValue(prop.Value, prop.Type)).c_str();
		if (count == 4)
			break;
	}
	return text;
}

//
// runs on its own (MTA) thread, with its own connection
//
void CEventsDlg::ListenThread(std::stop_token st, WMIConnection const* conn, CString ns, CString query, int generation) {
	::CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	auto post = [&](EventBatch* batch) {
		batch->Generation = generation;
		if (!PostMessage(WM_EVENTS, 0, reinterpret_cast<LPARAM>(batch)))
			delete batch;
	};

	HRESULT hr;
	{
		CComPtr<IWbemServices> spSvc;
		CComPtr<IEnumWbemClassObject> spEnum;
		hr = WMIHelper::Connect(conn, ns, &spSvc);
		if (SUCCEEDED(hr))
			hr = spSvc->ExecNotificationQuery(CComBSTR(L"WQL"), CComBSTR(query), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &spEnum);
		if (SUCCEEDED(hr)) {
			WMIHelper::SetSecurity(spEnum, conn);
			IWbemClassObject* objects[64];
			while (!st.stop_requested()) {
				ULONG count = 0;
				// a timeout, to notice a stop request
				hr = spEnum->Next(500, _countof(objects), objects, &count);
				if (count) {
					auto batch = new EventBatch;
					for (ULONG i = 0; i < count; i++) {
						Event e;
						::GetLocalTime(&e.Time);
						e.Object.Attach(objects[i]);
						e.Class = WMIHelper::GetStringProperty(e.Object, L"__CLASS");
						e.Summary = GetSummary(e.Object);
						batch->Events.push_back(std::move(e));
					}
					post(batch);
				}
				if (FAILED(hr))		// (WBEM_S_TIMEDOUT is a success code)
					break;
			}
		}
	}
	auto done = new EventBatch;
	done->Done = true;
	if (FAILED(hr))
		done->Error = WMIHelper::GetErrorText(hr);
	post(done);
	::CoUninitialize();
}

void CEventsDlg::Stop() {
	if (m_Thread.joinable()) {
		m_Thread.request_stop();
		m_Thread.join();
	}
	m_Listening = false;
	UpdateButtons();
}

void CEventsDlg::UpdateButtons() {
	GetDlgItem(IDC_START).EnableWindow(!m_Listening);
	GetDlgItem(IDC_STOP).EnableWindow(m_Listening);
}

void CEventsDlg::UpdateStatus() {
	CString text;
	if (!m_Error.IsEmpty())
		text = L"Error: " + m_Error;
	else
		text = std::format(L"{}{} event(s)", m_Listening ? L"Listening... " : L"", m_Events.size()).c_str();
	SetDlgItemText(IDC_STATUS, text);
}

LRESULT CEventsDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	DlgResize_Init(true, true, WS_THICKFRAME | WS_CLIPCHILDREN);
	SetDialogIcon(IDR_MAINFRAME);

	SetDlgItemText(IDC_NAMESPACE, L"ROOT\\CIMV2");
	SetDlgItemText(IDC_QUERY, L"SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'");

	m_List.Attach(GetDlgItem(IDC_RESULTS));
	m_List.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);
	CImageList images;
	images.Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);
	images.AddIcon(AtlLoadIconImage(IDI_OBJECT, 0, 16, 16));
	m_List.SetImageList(images, LVSIL_SMALL);

	auto cm = GetColumnManager(m_List);
	cm->AddColumn(L"Time", LVCFMT_LEFT, 90, ColumnType::Time);
	cm->AddColumn(L"Event Class", LVCFMT_LEFT, 200, ColumnType::Class);
	cm->AddColumn(L"Details", LVCFMT_LEFT, 500, ColumnType::Summary);

	UpdateButtons();
	return TRUE;
}

LRESULT CEventsDlg::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled) {
	Stop();
	MSG msg;
	while (::PeekMessage(&msg, m_hWnd, WM_EVENTS, WM_EVENTS, PM_REMOVE))
		delete reinterpret_cast<EventBatch*>(msg.lParam);
	bHandled = FALSE;
	return 0;
}

LRESULT CEventsDlg::OnEvents(UINT, WPARAM, LPARAM lParam, BOOL&) {
	std::unique_ptr<EventBatch> batch(reinterpret_cast<EventBatch*>(lParam));
	if (batch->Generation != m_Generation)
		return 0;

	if (!batch->Events.empty()) {
		m_Events.insert(m_Events.end(), std::make_move_iterator(batch->Events.begin()), std::make_move_iterator(batch->Events.end()));
		m_List.SetItemCountEx((int)m_Events.size(), LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
		m_List.EnsureVisible((int)m_Events.size() - 1, FALSE);
	}
	if (batch->Done) {
		m_Listening = false;
		m_Error = batch->Error;
		UpdateButtons();
	}
	UpdateStatus();
	return 0;
}

LRESULT CEventsDlg::OnStart(WORD, WORD, HWND, BOOL&) {
	CString ns, query;
	GetDlgItemText(IDC_NAMESPACE, ns);
	GetDlgItemText(IDC_QUERY, query);
	ns.Trim();
	query.Trim();
	if (ns.IsEmpty() || query.IsEmpty())
		return 0;

	Stop();
	m_Error.Empty();
	m_Listening = true;
	UpdateButtons();
	UpdateStatus();
	auto generation = ++m_Generation;
	auto conn = WMIHelper::CurrentConnection();
	m_Thread = std::jthread([=, this](std::stop_token st) {
		ListenThread(st, conn, ns, query, generation);
		});
	return 0;
}

LRESULT CEventsDlg::OnStop(WORD, WORD, HWND, BOOL&) {
	// (a few events may still arrive from the stopped query: they are dropped)
	++m_Generation;
	Stop();
	UpdateStatus();
	return 0;
}

LRESULT CEventsDlg::OnClear(WORD, WORD, HWND, BOOL&) {
	m_Events.clear();
	m_List.SetItemCount(0);
	UpdateStatus();
	return 0;
}

LRESULT CEventsDlg::OnCloseCmd(WORD, WORD, HWND, BOOL&) {
	// hidden, not destroyed: it keeps listening
	ShowWindow(SW_HIDE);
	return 0;
}
