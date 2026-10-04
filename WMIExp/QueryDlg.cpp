#include "pch.h"
#include "QueryDlg.h"
#include "WMIHelper.h"
#include "TextDlg.h"

void CQueryDlg::Activate(HWND hOwner, CString const& ns, CString const& className) {
	if (!IsWindow())
		Create(hOwner);

	if (GetDlgItem(IDC_QUERY).GetWindowTextLength() == 0) {
		SetDlgItemText(IDC_NAMESPACE, ns);
		if (!className.IsEmpty())
			SetDlgItemText(IDC_QUERY, L"SELECT * FROM " + className);
	}
	ShowWindow(SW_SHOW);
	SetActiveWindow();
	GotoDlgCtrl(GetDlgItem(IDC_QUERY));
}

void CQueryDlg::Reset() {
	if (!IsWindow())
		return;
	Cancel();
	m_Objects.clear();
	m_List.SetItemCount(0);
	SetDlgItemText(IDC_STATUS, L"");
}

CString CQueryDlg::GetColumnText(HWND, int row, int col) const {
	auto index = GetColumnManager(m_List)->GetColumnTag<int>(col);
	if (row >= (int)m_Objects.size() || index >= (int)m_Columns.size())
		return L"";

	CComVariant value;
	if (FAILED(m_Objects[row]->Get(m_Columns[index].Name, 0, &value, nullptr, nullptr)))
		return L"";
	return WMIHelper::FormatValue(value, m_Columns[index].Type);
}

int CQueryDlg::GetRowImage(HWND, int, int) const {
	return 0;
}

void CQueryDlg::DoSort(const SortInfo* si) {
	auto index = GetColumnManager(m_List)->GetColumnTag<int>(si->SortColumn);
	if (index >= (int)m_Columns.size())
		return;

	auto& column = m_Columns[index];
	auto count = m_Objects.size();
	std::vector<CComVariant> values(count);
	for (size_t i = 0; i < count; i++)
		m_Objects[i]->Get(column.Name, 0, &values[i], nullptr, nullptr);

	std::vector<size_t> order(count);
	std::iota(order.begin(), order.end(), size_t(0));
	std::ranges::stable_sort(order, [&](size_t i1, size_t i2) {
		auto result = WMIHelper::CompareValues(values[i1], values[i2], column.Type);
		return si->SortAscending ? result < 0 : result > 0;
		});

	std::vector<CComPtr<IWbemClassObject>> sorted;
	sorted.reserve(count);
	for (auto i : order)
		sorted.push_back(std::move(m_Objects[i]));
	m_Objects = std::move(sorted);
}

bool CQueryDlg::OnDoubleClickList(HWND, int row, int, POINT const&) {
	if (row < 0 || row >= (int)m_Objects.size())
		return false;

	CTextDlg::ShowObject(m_hWnd, m_Objects[row]);
	return true;
}

//
// the properties of the results: the classes of the results may differ (and so may their properties, with a projection)
//
void CQueryDlg::BuildColumns() {
	ClearSort(m_List);
	GetColumnManager(m_List)->Clear();
	m_Columns.clear();

	auto add = [&](CString const& name, CIMTYPE type, PCWSTR header = nullptr) {
		if (std::ranges::any_of(m_Columns, [&](auto const& c) { return c.Name.CompareNoCase(name) == 0; }))
			return;
		int format = LVCFMT_LEFT;
		switch (type) {
			case CIM_SINT8: case CIM_UINT8: case CIM_SINT16: case CIM_UINT16: case CIM_SINT32: case CIM_UINT32:
			case CIM_SINT64: case CIM_UINT64: case CIM_REAL32: case CIM_REAL64:
				format = LVCFMT_RIGHT;
				break;
		}
		GetColumnManager(m_List)->AddColumn(header ? header : (PCWSTR)name, format, 130, (int)m_Columns.size());
		m_Columns.push_back({ name, type });
	};

	std::vector<CString> classes;
	for (auto& obj : m_Objects) {
		auto className = WMIHelper::GetStringProperty(obj, L"__CLASS");
		if (std::ranges::find(classes, className) != classes.end())
			continue;
		classes.push_back(className);
	}
	if (classes.size() > 1)
		add(L"__CLASS", CIM_STRING, L"Class");

	// the properties of the first object of each class
	classes.clear();
	for (auto& obj : m_Objects) {
		auto className = WMIHelper::GetStringProperty(obj, L"__CLASS");
		if (std::ranges::find(classes, className) != classes.end())
			continue;
		classes.push_back(className);
		for (auto& prop : WMIHelper::EnumProperties(obj, WBEM_FLAG_NONSYSTEM_ONLY))
			add(CString(prop.Name), prop.Type);
	}
	// such as SELECT __PATH FROM ...
	if (m_Columns.empty() && !m_Objects.empty()) {
		for (auto& prop : WMIHelper::EnumProperties(m_Objects[0], WBEM_FLAG_SYSTEM_ONLY))
			add(CString(prop.Name), prop.Type);
	}
}

void CQueryDlg::Cancel() {
	if (m_Job) {
		m_Job->Cancelled = true;
		m_Job.reset();
	}
	UpdateButtons();
}

void CQueryDlg::UpdateButtons() {
	GetDlgItem(IDC_RUN).EnableWindow(m_Job == nullptr);
	GetDlgItem(IDC_STOP).EnableWindow(m_Job != nullptr);
}

LRESULT CQueryDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	DlgResize_Init(true, true, WS_THICKFRAME | WS_CLIPCHILDREN);
	SetDialogIcon(IDR_MAINFRAME);

	m_List.Attach(GetDlgItem(IDC_RESULTS));
	m_List.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);
	CImageList images;
	images.Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);
	images.AddIcon(AtlLoadIconImage(IDI_OBJECT, 0, 16, 16));
	m_List.SetImageList(images, LVSIL_SMALL);

	UpdateButtons();
	return TRUE;
}

LRESULT CQueryDlg::OnRun(WORD, WORD, HWND, BOOL&) {
	CString ns, query;
	GetDlgItemText(IDC_NAMESPACE, ns);
	GetDlgItemText(IDC_QUERY, query);
	ns.Trim();
	query.Trim();
	if (ns.IsEmpty())
		ns = L"ROOT\\CIMV2";
	if (query.IsEmpty())
		return 0;

	Cancel();
	m_Objects.clear();
	m_List.SetItemCount(0);

	CComPtr<IWbemServices> spSvc;
	HRESULT hr;
	{
		CWaitCursor wait;
		hr = WMIHelper::Connect(WMIHelper::CurrentConnection(), ns, &spSvc);
	}
	if (FAILED(hr)) {
		SetDlgItemText(IDC_STATUS, std::format(L"Error opening {}: {}", (PCWSTR)ns, (PCWSTR)WMIHelper::GetErrorText(hr)).c_str());
		return 0;
	}

	m_StartTime = ::GetTickCount();
	m_Job = WMIHelper::ExecQueryAsync(m_hWnd, WM_QUERY_DONE, query, spSvc);
	SetDlgItemText(IDC_STATUS, L"Running...");
	UpdateButtons();
	return 0;
}

LRESULT CQueryDlg::OnQueryDone(UINT, WPARAM, LPARAM lParam, BOOL&) {
	auto job = WMIHelper::TakeJob(lParam);
	if (job != m_Job)
		return 0;		// stopped, or replaced by another query

	m_Job.reset();
	UpdateButtons();
	m_Objects = std::move(job->Objects);
	BuildColumns();
	m_List.SetItemCount((int)m_Objects.size());

	auto elapsed = (::GetTickCount() - m_StartTime) / 1000.0;
	if (FAILED(job->Status) && m_Objects.empty())
		SetDlgItemText(IDC_STATUS, L"Error: " + WMIHelper::GetErrorText(job->Status));
	else if (FAILED(job->Status))
		SetDlgItemText(IDC_STATUS, std::format(L"{} object(s). Error: {}", m_Objects.size(), (PCWSTR)WMIHelper::GetErrorText(job->Status)).c_str());
	else
		SetDlgItemText(IDC_STATUS, std::format(L"{} object(s) ({:.2f} sec)", m_Objects.size(), elapsed).c_str());
	return 0;
}

LRESULT CQueryDlg::OnStop(WORD, WORD, HWND, BOOL&) {
	Cancel();
	SetDlgItemText(IDC_STATUS, L"Stopped.");
	return 0;
}

LRESULT CQueryDlg::OnCloseCmd(WORD, WORD, HWND, BOOL&) {
	// hidden, not destroyed: the query and results are there when it is shown again
	ShowWindow(SW_HIDE);
	return 0;
}
