#pragma once

#include "DialogHelper.h"
#include "resource.h"

struct WMIConnection;

//
// connects to the ROOT namespace of a computer; closes only when connected (or cancelled)
//
class CConnectDlg :
	public CDialogImpl<CConnectDlg>,
	public CDialogHelper<CConnectDlg> {
public:
	enum { IDD = IDD_CONNECT };

	WMIConnection const* GetConnection() const {
		return m_Connection;
	}
	IWbemServices* GetRoot() const {
		return m_spRoot;
	}

	BEGIN_MSG_MAP(CConnectDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDOK, OnOK)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
	END_MSG_MAP()

private:
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnOK(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	WMIConnection const* m_Connection{ nullptr };
	CComPtr<IWbemServices> m_spRoot;
};
