#include "pch.h"
#include "ObjectDlg.h"
#include "WMIHelper.h"
#include "TextDlg.h"
#include "AppSettings.h"
#include <ListViewhelper.h>
#include <ClipboardHelper.h>

CObjectDlg::CObjectDlg(IWbemServices* pSvc, IWbemClassObject* pObj) : m_spSvc(pSvc), m_spObj(pObj) {
	// an instance's path, or the class name (an embedded object has no path)
	auto relPath = WMIHelper::GetStringProperty(pObj, L"__RELPATH");
	m_Title = relPath.IsEmpty() ? WMIHelper::GetStringProperty(pObj, L"__CLASS") : relPath;

	int index = 0;
	for (auto& prop : WMIHelper::EnumProperties(pObj, WBEM_FLAG_NONSYSTEM_ONLY))
		m_Rows.push_back({ CString(prop.Name), prop.Value, prop.Type, index++ });
	if (AppSettings::Get().ViewSystemProperties() || m_Rows.empty()) {
		for (auto& prop : WMIHelper::EnumProperties(pObj, WBEM_FLAG_SYSTEM_ONLY))
			m_Rows.push_back({ CString(prop.Name), prop.Value, prop.Type, index++, true });
	}
}

CObjectDlg::CObjectDlg(IWbemServices* pSvc, CString const& name, CComVariant const& array, CIMTYPE type) : m_spSvc(pSvc) {
	auto elementType = type & ~CIM_FLAG_ARRAY;
	auto sa = (array.vt & VT_ARRAY) ? array.parray : nullptr;
	LONG lower = 0, upper = -1;
	VARTYPE vt = VT_EMPTY;
	if (sa) {
		::SafeArrayGetLBound(sa, 1, &lower);
		::SafeArrayGetUBound(sa, 1, &upper);
		::SafeArrayGetVartype(sa, &vt);
	}
	m_Title = std::format(L"{} ({} elements)", (PCWSTR)name, upper - lower + 1).c_str();

	for (LONG i = lower; i <= upper; i++) {
		// the element is copied into the variant's data (the union is the same for all types)
		CComVariant element;
		if (vt == VT_VARIANT) {
			if (FAILED(::SafeArrayGetElement(sa, &i, &element)))
				continue;
		}
		else {
			if (FAILED(::SafeArrayGetElement(sa, &i, &element.llVal)))
				continue;
			element.vt = vt;
		}
		m_Rows.push_back({ std::format(L"[{}]", i - lower).c_str(), std::move(element), elementType, (int)(i - lower) });
	}
}

bool CObjectDlg::CanShowValue(CComVariant const& value, CIMTYPE type) {
	if (type & CIM_FLAG_ARRAY)
		return (value.vt & VT_ARRAY) && value.parray;
	switch (type) {
		case CIM_REFERENCE: return value.vt == VT_BSTR && ::SysStringLen(value.bstrVal) > 0;
		case CIM_OBJECT: return value.vt == VT_UNKNOWN && value.punkVal;
	}
	return false;
}

bool CObjectDlg::ShowValue(HWND hParent, IWbemServices* pSvc, CString const& name, CComVariant const& value, CIMTYPE type) {
	if (!CanShowValue(value, type))
		return false;

	if (type & CIM_FLAG_ARRAY) {
		CObjectDlg dlg(pSvc, name, value, type);
		dlg.DoModal(hParent);
		return true;
	}
	if (type == CIM_REFERENCE) {
		ShowReference(hParent, pSvc, value.bstrVal);
		return true;
	}

	CComQIPtr<IWbemClassObject> spObj(value.punkVal);
	if (spObj == nullptr)
		return false;
	CObjectDlg dlg(pSvc, spObj);
	dlg.DoModal(hParent);
	return true;
}

void CObjectDlg::ShowReference(HWND hParent, IWbemServices* pSvc, CString const& path) {
	CComPtr<IWbemClassObject> spObj;
	CComPtr<IWbemServices> spSvc;
	HRESULT hr;
	{
		CWaitCursor wait;
		hr = WMIHelper::GetObjectByPath(pSvc, path, &spObj, &spSvc);
	}
	if (FAILED(hr)) {
		AtlMessageBox(hParent, std::format(L"Failed to get {}: {}", (PCWSTR)path, (PCWSTR)WMIHelper::GetErrorText(hr)).c_str(),
			IDS_TITLE, MB_ICONERROR);
		return;
	}
	// (its references are relative to its own namespace)
	CObjectDlg dlg(spSvc, spObj);
	dlg.DoModal(hParent);
}

CString CObjectDlg::GetColumnText(HWND, int row, int col) const {
	auto& r = m_Rows[row];
	switch (GetColumnManager(m_List)->GetColumnTag<ColumnType>(col)) {
		case ColumnType::Name: return r.Name;
		case ColumnType::Type: return WMIHelper::CimTypeToString(r.Type);
		case ColumnType::Value: return WMIHelper::FormatValue(r.Value, r.Type);
	}
	return L"";
}

int CObjectDlg::GetRowImage(HWND, int row, int) const {
	auto& r = m_Rows[row];
	if (CanShowValue(r.Value, r.Type))
		return 2;
	return r.System ? 1 : 0;
}

void CObjectDlg::DoSort(const SortInfo* si) {
	auto column = GetColumnManager(m_List)->GetColumnTag<ColumnType>(si->SortColumn);
	std::ranges::stable_sort(m_Rows, [&](auto const& r1, auto const& r2) {
		int result = 0;
		switch (column) {
			case ColumnType::Name:
				// array elements by their position
				result = m_spObj ? r1.Name.CompareNoCase(r2.Name) : r1.Index - r2.Index;
				break;
			case ColumnType::Type: result = WMIHelper::CimTypeToString(r1.Type).CompareNoCase(WMIHelper::CimTypeToString(r2.Type)); break;
			case ColumnType::Value: result = WMIHelper::CompareValues(r1.Value, r2.Value, r1.Type); break;
		}
		return si->SortAscending ? result < 0 : result > 0;
		});
}

bool CObjectDlg::OnDoubleClickList(HWND, int row, int, POINT const&) {
	if (row < 0 || row >= (int)m_Rows.size())
		return false;
	auto& r = m_Rows[row];
	return ShowValue(m_hWnd, m_spSvc, r.Name, r.Value, r.Type);
}

LRESULT CObjectDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	DlgResize_Init(true, true, WS_THICKFRAME | WS_CLIPCHILDREN);
	SetDialogIcon(IDR_MAINFRAME);
	SetWindowText(m_Title);

	m_List.Attach(GetDlgItem(IDC_RESULTS));
	m_List.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP);
	CImageList images;
	images.Create(16, 16, ILC_COLOR32 | ILC_MASK, 3, 1);
	for (auto icon : { IDI_PROPERTY, IDI_PROPERTY2, IDI_OBJECT })
		images.AddIcon(AtlLoadIconImage(icon, 0, 16, 16));
	m_List.SetImageList(images, LVSIL_SMALL);

	auto cm = GetColumnManager(m_List);
	cm->AddColumn(m_spObj ? L"Property" : L"Element", LVCFMT_LEFT, 200, ColumnType::Name);
	cm->AddColumn(L"Type", LVCFMT_LEFT, 130, ColumnType::Type);
	cm->AddColumn(L"Value", LVCFMT_LEFT, 380, ColumnType::Value);
	m_List.SetItemCount((int)m_Rows.size());

	GetDlgItem(IDC_MOF).EnableWindow(m_spObj != nullptr);
	SetDlgItemText(IDC_STATUS, std::any_of(m_Rows.begin(), m_Rows.end(), [](auto const& r) { return CanShowValue(r.Value, r.Type); })
		? L"Double-click a reference, an embedded object or an array to open it." : L"");

	GotoDlgCtrl(m_List);
	if (!m_Rows.empty())
		m_List.SetItemState(0, LVIS_FOCUSED, LVIS_FOCUSED);
	return FALSE;
}

LRESULT CObjectDlg::OnCopy(WORD, WORD, HWND, BOOL&) {
	// the selected rows, or all of them
	CString text;
	if (m_List.GetSelectedCount() > 0)
		text = ListViewHelper::GetSelectedRowsAsString(m_List);
	else {
		for (int i = 0; i < m_List.GetItemCount(); i++)
			text += ListViewHelper::GetRowAsString(m_List, i) + L"\r\n";
	}
	ClipboardHelper::CopyText(m_hWnd, text);
	return 0;
}

LRESULT CObjectDlg::OnShowMof(WORD, WORD, HWND, BOOL&) {
	if (m_spObj)
		CTextDlg::ShowObject(m_hWnd, m_spObj);
	return 0;
}

LRESULT CObjectDlg::OnCloseCmd(WORD, WORD, HWND, BOOL&) {
	EndDialog(IDCANCEL);
	return 0;
}
