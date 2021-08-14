#pragma once

struct WMIProperty {
	CComBSTR Name;
	CComVariant Value;
	CIMTYPE Type;
	long Flavor;
};

struct WMIMethod {
	CComBSTR Name;
	CComPtr<IWbemClassObject> spInParams, spOutParams;
	CComBSTR ClassName;
};

struct IObjectsCallback {
	virtual int GetObjectCount() const = 0;
	virtual CComPtr<IWbemClassObject> GetItem(int i) const = 0;
	virtual ULONG Release() = 0;
};

struct WMIHelper abstract final {
	static HRESULT Init(PCWSTR computerName, PCWSTR ns, IWbemServices** ppWmi);
	static CString GetStringProperty(IWbemClassObject* pObj, PCWSTR name);
	static std::vector<CComPtr<IWbemClassObject>> EnumNamespaces(IWbemServices* pWmi);
	static std::vector<CComPtr<IWbemClassObject>> EnumClasses(IWbemServices* pSvc, bool deep, bool includeSystemClasses = false);
	static std::vector<CComPtr<IWbemClassObject>> EnumInstances(PCWSTR name, IWbemServices* pSvc, bool deep);
	static bool EnumInstancesAsync(HWND hWnd, UINT msg, PCWSTR name, IWbemServices* pSvc, bool deep);
	static std::vector<WMIProperty> EnumProperties(IWbemClassObject* pObj);
	static std::vector<WMIMethod> EnumMethods(IWbemClassObject* pObj, bool localOnly = false, bool inheritedOnly = false);
	static std::vector<CComBSTR> GetNames(IWbemClassObject* pObj);
};
