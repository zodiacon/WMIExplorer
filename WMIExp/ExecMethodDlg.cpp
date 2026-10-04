#include "pch.h"
#include "ExecMethodDlg.h"
#include "WMIHelper.h"

LRESULT CExecMethodDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	SetDialogIcon(IDR_MAINFRAME);
	SetWindowText(L"Execute " + m_Method);
	SetDlgItemText(IDC_TARGET, std::format(L"Method {} on {}", (PCWSTR)m_Method, (PCWSTR)m_Path).c_str());
	CreateParameterRows();

	if (!m_Parameters.empty() && m_Parameters[0].Edit.IsWindowEnabled()) {
		GotoDlgCtrl(m_Parameters[0].Edit);
		return FALSE;
	}
	return TRUE;
}

//
// a label and an edit box for each parameter, below the "Input parameters" label;
// what is below it is moved down, and the dialog made taller
//
void CExecMethodDlg::CreateParameterRows() {
	if (m_spInParams) {
		for (auto& prop : WMIHelper::EnumProperties(m_spInParams, WBEM_FLAG_NONSYSTEM_ONLY)) {
			Parameter p;
			p.Name = prop.Name;
			p.Type = prop.Type;
			p.Id = INT_MAX;
			// the ID qualifier is the position of the parameter in the method's signature
			CComPtr<IWbemQualifierSet> spSet;
			CComVariant id;
			if (SUCCEEDED(m_spInParams->GetPropertyQualifierSet(prop.Name, &spSet)) && SUCCEEDED(spSet->Get(L"ID", 0, &id, nullptr))
				&& SUCCEEDED(id.ChangeType(VT_I4)))
				p.Id = id.lVal;
			m_Parameters.push_back(std::move(p));
		}
		std::ranges::stable_sort(m_Parameters, {}, &Parameter::Id);
	}
	SetDlgItemText(IDC_PARAMS, m_Parameters.empty() ? L"No input parameters." : L"Input parameters:");
	if (m_Parameters.empty())
		return;

	auto dluX = [&](int x) { CRect rc(0, 0, x, 0); MapDialogRect(&rc); return rc.right; };
	auto dluY = [&](int y) { CRect rc(0, 0, 0, y); MapDialogRect(&rc); return rc.bottom; };

	CRect label;
	GetDlgItem(IDC_PARAMS).GetWindowRect(&label);
	ScreenToClient(&label);
	int rowHeight = dluY(18);
	int delta = rowHeight * (int)m_Parameters.size() + dluY(4);

	for (auto child = GetWindow(GW_CHILD); child; child = child.GetWindow(GW_HWNDNEXT)) {
		CRect rc;
		child.GetWindowRect(&rc);
		ScreenToClient(&rc);
		if (rc.top > label.top)
			child.SetWindowPos(nullptr, rc.left, rc.top + delta, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
	}
	CRect rcDlg;
	GetWindowRect(&rcDlg);
	SetWindowPos(nullptr, 0, 0, rcDlg.Width(), rcDlg.Height() + delta, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

	auto font = GetFont();
	int y = label.bottom + dluY(4);
	int left = dluX(7), labelWidth = dluX(120), editLeft = dluX(130), editWidth = dluX(183);
	// in the tab order after the label (new windows are added last)
	HWND hAfter = GetDlgItem(IDC_PARAMS);
	UINT id = 2000;
	for (auto& p : m_Parameters) {
		CRect rcText(left, y + dluY(3), left + labelWidth, y + dluY(11));
		CStatic text;
		text.Create(m_hWnd, rcText,
			p.Name + L" (" + WMIHelper::CimTypeToString(p.Type) + L")", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS | SS_NOPREFIX, 0, id++);
		text.SetFont(font);
		text.SetWindowPos(hAfter, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		hAfter = text;

		CRect rcEdit(editLeft, y, editLeft + editWidth, y + dluY(14));
		p.Edit.Create(m_hWnd, rcEdit, nullptr,
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, WS_EX_CLIENTEDGE, id++);
		p.Edit.SetFont(font);
		p.Edit.SetWindowPos(hAfter, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		hAfter = p.Edit;

		auto type = p.Type & ~CIM_FLAG_ARRAY;
		if (type == CIM_OBJECT) {
			p.Edit.SetWindowText(L"(objects are not supported)");
			p.Edit.EnableWindow(FALSE);
		}
		else if (p.Type & CIM_FLAG_ARRAY)
			p.Edit.SetCueBannerText(L"Comma separated values");
		else if (type == CIM_BOOLEAN)
			p.Edit.SetCueBannerText(L"True or False");
		else if (type == CIM_DATETIME)
			p.Edit.SetCueBannerText(L"yyyymmddHHMMSS.000000+000");
		y += rowHeight;
	}
}

LRESULT CExecMethodDlg::OnExecute(WORD, WORD, HWND, BOOL&) {
	CComPtr<IWbemClassObject> spIn;
	if (m_spInParams) {
		auto hr = m_spInParams->SpawnInstance(0, &spIn);
		if (FAILED(hr)) {
			SetDlgItemText(IDC_OUTPUT, L"Error: " + WMIHelper::GetErrorText(hr));
			return 0;
		}
		for (auto& p : m_Parameters) {
			CString text;
			if (!p.Edit.IsWindowEnabled() || p.Edit.GetWindowText(text) == 0)
				continue;		// left null

			CComVariant value;
			if (FAILED(WMIHelper::ParseValue(text, p.Type, value)) || FAILED(hr = spIn->Put(p.Name, 0, &value, 0))) {
				AtlMessageBox(m_hWnd, (PCWSTR)std::format(L"Invalid value for {} ({})", (PCWSTR)p.Name,
					(PCWSTR)WMIHelper::CimTypeToString(p.Type)).c_str(), IDS_TITLE, MB_ICONERROR);
				GotoDlgCtrl(p.Edit);
				return 0;
			}
		}
	}

	CComPtr<IWbemClassObject> spOut;
	HRESULT hr;
	{
		CWaitCursor wait;
		hr = m_spSvc->ExecMethod(CComBSTR(m_Path), CComBSTR(m_Method), 0, nullptr, spIn, &spOut, nullptr);
	}

	SYSTEMTIME now;
	::GetLocalTime(&now);
	auto output = std::format(L"{:02}:{:02}:{:02} ", now.wHour, now.wMinute, now.wSecond);
	if (FAILED(hr))
		output += L"Error: " + std::wstring(WMIHelper::GetErrorText(hr));
	else {
		output += L"Done.";
		if (spOut) {
			for (auto& prop : WMIHelper::EnumProperties(spOut, WBEM_FLAG_NONSYSTEM_ONLY))
				output += std::format(L"\r\n{} = {}", (PCWSTR)prop.Name, (PCWSTR)WMIHelper::FormatValue(prop.Value, prop.Type));
		}
	}
	SetDlgItemText(IDC_OUTPUT, output.c_str());
	return 0;
}

LRESULT CExecMethodDlg::OnCloseCmd(WORD, WORD, HWND, BOOL&) {
	EndDialog(IDCANCEL);
	return 0;
}
