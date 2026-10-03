#pragma once

#include <wil\com.h>

struct WMIProperty {
	CComBSTR Name;
	CComVariant Value;
	CIMTYPE Type;
	long Flavor;
};

struct WMIMethod {
	std::wstring Name;
	wil::com_ptr<IWbemClassObject> spInParams, spOutParams;
	std::wstring ClassName;
};

//
// posted (as LPARAM) to the window when an async enumeration completes, fails or is cancelled.
// the receiver must call Release.
//
struct IObjectsCallback {
	virtual int GetObjectCount() const = 0;
	virtual CComPtr<IWbemClassObject> GetItem(int i) const = 0;
	virtual HRESULT GetStatus() const = 0;
	virtual IWbemObjectSink* GetSink() = 0;
	virtual ULONG Release() = 0;
};

struct WMIHelper abstract final {
	static HRESULT Init(PCWSTR computerName, PCWSTR ns, IWbemServices** ppWmi);
	static CString GetStringProperty(IWbemClassObject* pObj, PCWSTR name);
	static std::vector<CComPtr<IWbemClassObject>> EnumNamespaces(IWbemServices* pWmi);
	static std::vector<CComPtr<IWbemClassObject>> EnumClasses(IWbemServices* pSvc, bool deep, bool includeSystemClasses = false);
	static std::vector<CComPtr<IWbemClassObject>> EnumInstances(PCWSTR name, IWbemServices* pSvc, bool deep);
	static HRESULT EnumInstancesAsync(HWND hWnd, UINT msg, PCWSTR name, IWbemServices* pSvc, bool deep, IWbemObjectSink** ppSink);
	static CString GetErrorText(HRESULT hr);
	static std::vector<WMIProperty> EnumProperties(IWbemClassObject* pObj, long flags = 0);
	static std::vector<WMIMethod> EnumMethods(IWbemClassObject* pObj, bool localOnly = false, bool inheritedOnly = false);
	static std::vector<CComBSTR> GetNames(IWbemClassObject* pObj);
};
