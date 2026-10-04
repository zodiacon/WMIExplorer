#include "pch.h"
#include "TextDlg.h"
#include "WMIHelper.h"
#include <WTLHelper.h>
#include <ClipboardHelper.h>

void CTextDlg::ShowObject(HWND hParent, IWbemClassObject* pObj) {
	auto name = WMIHelper::GetStringProperty(pObj, L"__CLASS");
	// (an instance's path, or the class name)
	auto relPath = WMIHelper::GetStringProperty(pObj, L"__RELPATH");
	CTextDlg dlg(relPath.IsEmpty() ? name : relPath, WMIHelper::GetObjectText(pObj), name);
	dlg.DoModal(hParent);
}

LRESULT CTextDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
	DlgResize_Init(true, true, WS_THICKFRAME | WS_CLIPCHILDREN);
	SetDialogIcon(IDR_MAINFRAME);
	SetWindowText(m_Title);

	m_Font.CreatePointFont(100, L"Consolas");
	auto edit = GetDlgItem(IDC_TEXT);
	edit.SetFont(m_Font);
	edit.SendMessage(EM_SETLIMITTEXT, 0);	// (the default is 30000 characters)
	edit.SetWindowText(m_Text);

	// not the text: it would be all selected
	GotoDlgCtrl(GetDlgItem(IDCANCEL));
	return FALSE;
}

LRESULT CTextDlg::OnCopy(WORD, WORD, HWND, BOOL&) {
	ClipboardHelper::CopyText(m_hWnd, m_Text);
	return 0;
}

LRESULT CTextDlg::OnSave(WORD, WORD, HWND, BOOL&) {
	CFileDialog dlg(FALSE, L"mof", m_FileName, OFN_OVERWRITEPROMPT | OFN_EXPLORER | OFN_ENABLESIZING,
		L"MOF Files (*.mof)\0*.mof\0Text Files (*.txt)\0*.txt\0All Files\0*.*\0", m_hWnd);
	{
		// the file dialog does not support dark mode
		SuspendResumeHook suspend;
		if (dlg.DoModal(m_hWnd) != IDOK)
			return 0;
	}

	// UTF-16 with a BOM, which mofcomp accepts
	wil::unique_hfile hFile(::CreateFile(dlg.m_szFileName, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, 0, nullptr));
	if (!hFile) {
		AtlMessageBox(m_hWnd, L"Failed to create the file.", IDS_TITLE, MB_ICONERROR);
		return 0;
	}
	const WCHAR bom = 0xFEFF;
	DWORD written;
	if (!::WriteFile(hFile.get(), &bom, sizeof(bom), &written, nullptr) ||
		!::WriteFile(hFile.get(), (PCWSTR)m_Text, m_Text.GetLength() * sizeof(WCHAR), &written, nullptr))
		AtlMessageBox(m_hWnd, L"Failed to write the file.", IDS_TITLE, MB_ICONERROR);
	return 0;
}

LRESULT CTextDlg::OnCloseCmd(WORD, WORD, HWND, BOOL&) {
	EndDialog(IDCANCEL);
	return 0;
}
