// WMIExp.cpp : main source file for WMIExp.exe
//

#include "pch.h"
#include "resource.h"
#include "MainFrm.h"
#include <ThemeHelper.h>

CAppModule _Module;

int Run(LPTSTR /*lpstrCmdLine*/ = nullptr, int nCmdShow = SW_SHOWDEFAULT) {
	CMessageLoop theLoop;
	_Module.AddMessageLoop(&theLoop);

	CMainFrame wndMain;

	if (wndMain.CreateEx() == nullptr) 	{
		ATLTRACE(_T("Main window creation failed!\n"));
		return 0;
	}

	wndMain.ShowWindow(nCmdShow);

	int nRet = theLoop.Run();

	_Module.RemoveMessageLoop();
	return nRet;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR lpstrCmdLine, int nCmdShow) {
	auto hr = ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	ATLASSERT(SUCCEEDED(hr));

	//
	// needed to prevent access denied errors from WMI
	//
	hr = ::CoInitializeSecurity(
		nullptr,
		-1,								// COM authentication
		nullptr,                        // Authentication services
		nullptr,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_DELEGATE,
		nullptr,                        // Authentication info
		EOAC_NONE,						// Additional capabilities 
		nullptr);
	ATLASSERT(SUCCEEDED(hr));

	AtlInitCommonControls(ICC_COOL_CLASSES | ICC_BAR_CLASSES | ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES);

	ThemeHelper::Init();
	hr = _Module.Init(nullptr, hInstance);
	ATLASSERT(SUCCEEDED(hr));

	int nRet = Run(lpstrCmdLine, nCmdShow);

	_Module.Term();
	::CoUninitialize();

	return nRet;
}
