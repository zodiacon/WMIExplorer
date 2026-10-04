#pragma once

#include "DialogHelper.h"
#include "resource.h"

//
// shows text (an object's MOF) that can be copied or saved to a file
//
class CTextDlg :
	public CDialogImpl<CTextDlg>,
	public CDialogResize<CTextDlg>,
	public CDialogHelper<CTextDlg> {
public:
	enum { IDD = IDD_TEXT };

	CTextDlg(CString const& title, CString const& text, CString const& fileName) : m_Title(title), m_Text(text), m_FileName(fileName) {}

	// the MOF text of an object (class or instance)
	static void ShowObject(HWND hParent, IWbemClassObject* pObj);

	BEGIN_MSG_MAP(CTextDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDC_COPY, OnCopy)
		COMMAND_ID_HANDLER(IDC_SAVE, OnSave)
		COMMAND_ID_HANDLER(IDCANCEL, OnCloseCmd)
		CHAIN_MSG_MAP(CDialogResize<CTextDlg>)
	END_MSG_MAP()

	BEGIN_DLGRESIZE_MAP(CTextDlg)
		DLGRESIZE_CONTROL(IDC_TEXT, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_COPY, DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDC_SAVE, DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

private:
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnCopy(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSave(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCloseCmd(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	CString m_Title, m_Text, m_FileName;
	CFont m_Font;
};
