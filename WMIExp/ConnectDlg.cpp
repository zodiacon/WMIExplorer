#include "pch.h"
#include "ConnectDlg.h"
#include "WMIHelper.h"

LRESULT CConnectDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	SetDialogIcon(IDR_MAINFRAME);
	auto current = WMIHelper::CurrentConnection();
	SetDlgItemText(IDC_COMPUTER, current->Computer);
	SetDlgItemText(IDC_USER, current->User);
	return TRUE;
}

LRESULT CConnectDlg::OnOK(WORD, WORD, HWND, BOOL&) {
	CString computer, user, password;
	GetDlgItemText(IDC_COMPUTER, computer);
	GetDlgItemText(IDC_USER, user);
	GetDlgItemText(IDC_PASSWORD, password);
	computer.Trim();
	computer.TrimLeft(L'\\');
	user.Trim();
	if (computer == L"." || computer.CompareNoCase(L"localhost") == 0)
		computer.Empty();

	auto conn = WMIHelper::NewConnection(computer, user, password);
	CComPtr<IWbemServices> spRoot;
	HRESULT hr;
	{
		// (an unreachable computer takes a while to fail)
		CWaitCursor wait;
		hr = WMIHelper::Connect(conn, L"ROOT", &spRoot);
	}
	if (FAILED(hr)) {
		AtlMessageBox(m_hWnd, (PCWSTR)(L"Failed to connect: " + WMIHelper::GetErrorText(hr)), IDS_TITLE, MB_ICONERROR);
		return 0;
	}

	m_Connection = conn;
	m_spRoot = spRoot;
	EndDialog(IDOK);
	return 0;
}

LRESULT CConnectDlg::OnCancel(WORD, WORD, HWND, BOOL&) {
	EndDialog(IDCANCEL);
	return 0;
}
