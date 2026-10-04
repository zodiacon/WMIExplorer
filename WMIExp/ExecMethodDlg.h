#pragma once

#include "DialogHelper.h"
#include "resource.h"

//
// collects the input parameters of a method, and runs it (on a class for a static method, otherwise on an instance)
//
class CExecMethodDlg :
	public CDialogImpl<CExecMethodDlg>,
	public CDialogHelper<CExecMethodDlg> {
public:
	enum { IDD = IDD_EXECMETHOD };

	// path: the class name (static method) or the instance's relative path. pInParams: may be null (no parameters)
	CExecMethodDlg(IWbemServices* pSvc, CString const& path, CString const& method, IWbemClassObject* pInParams) :
		m_spSvc(pSvc), m_Path(path), m_Method(method), m_spInParams(pInParams) {}

	BEGIN_MSG_MAP(CExecMethodDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDC_RUN, OnExecute)
		COMMAND_ID_HANDLER(IDCANCEL, OnCloseCmd)
	END_MSG_MAP()

private:
	struct Parameter {
		CString Name;
		CIMTYPE Type;
		int Id;			// the parameter's position
		CEdit Edit;
	};

	void CreateParameterRows();

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnExecute(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCloseCmd(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	CComPtr<IWbemServices> m_spSvc;
	CString m_Path, m_Method;
	CComPtr<IWbemClassObject> m_spInParams;
	std::vector<Parameter> m_Parameters;
};
