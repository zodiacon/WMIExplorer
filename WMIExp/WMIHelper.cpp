#include "pch.h"
#include "WMIHelper.h"

class CObjectSink : 
	public IObjectsCallback,
	public CComObjectRoot, 
	public IWbemObjectSink {
public:
	BEGIN_COM_MAP(CObjectSink)
		COM_INTERFACE_ENTRY(IWbemObjectSink)
	END_COM_MAP()

	void Init(HWND hWnd, UINT msg) {
		m_hWnd = hWnd;
		m_Msg = msg;
	}

	int GetObjectCount() const override {
		return (int)m_Objects.size();
	}
	CComPtr<IWbemClassObject> GetItem(int i) const override {
		return m_Objects[i];
	}

private:
	// Inherited via IWbemObjectSink
	HRESULT __stdcall Indicate(long lObjectCount, IWbemClassObject** apObjArray) override {
		for (int i = 0; i < lObjectCount; i++)
			m_Objects.push_back(apObjArray[i]);
		return S_OK;
	}
	HRESULT __stdcall SetStatus(long lFlags, HRESULT hr, BSTR strParam, IWbemClassObject* pObjParam) override {
		if (hr == S_OK && lFlags == WBEM_STATUS_COMPLETE) {
			AddRef();
			::PostMessage(m_hWnd, m_Msg, 0, reinterpret_cast<LPARAM>(static_cast<IObjectsCallback*>(this)));
		}
		return S_OK;
	}

	HWND m_hWnd;
	UINT m_Msg;
	std::vector<CComPtr<IWbemClassObject>> m_Objects;
};

HRESULT WMIHelper::Init(PCWSTR computerName, PCWSTR ns, IWbemServices** ppWmi) {
	CComPtr<IWbemLocator> spLocator;
	auto hr = spLocator.CoCreateInstance(__uuidof(WbemLocator));
	if (FAILED(hr))
		return hr;

	return spLocator->ConnectServer(CComBSTR(ns),
		nullptr, nullptr, nullptr, WBEM_FLAG_CONNECT_USE_MAX_WAIT, nullptr, nullptr, ppWmi);
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumNamespaces(IWbemServices* pWmi) {
	std::vector<CComPtr<IWbemClassObject>> ns;
	CComPtr<IEnumWbemClassObject> spEnum;
	//auto hr = pWmi->ExecQuery(CComBSTR(L"WQL"), CComBSTR(L"SELECT * FROM __NAMESPACE"), 0, nullptr, &spEnum);
	auto hr = pWmi->CreateInstanceEnum(CComBSTR(L"__NAMESPACE"), 0, nullptr, &spEnum);
	if (FAILED(hr))
		return ns;

	CComPtr<IWbemClassObject> spObj;
	ULONG count;
	while (S_OK == spEnum->Next(WBEM_INFINITE, 1, &spObj, &count)) {
		ns.push_back(spObj);
		spObj.Release();
	}
	return ns;
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumClasses(IWbemServices* pSvc, bool deep, bool includeSystemClasses) {
	std::vector<CComPtr<IWbemClassObject>> classes;
	CComPtr<IEnumWbemClassObject> spEnum;
	auto hr = pSvc->CreateClassEnum(nullptr, (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW) | WBEM_FLAG_FORWARD_ONLY, nullptr, &spEnum);
	if (FAILED(hr))
		return classes;

	CComPtr<IWbemClassObject> spObj;
	ULONG count;
	while (S_OK == spEnum->Next(WBEM_INFINITE, 1, &spObj, &count)) {
		if (!includeSystemClasses) {
			auto dynasty = GetStringProperty(spObj, L"__DYNASTY");
			if (dynasty.CompareNoCase(L"__SystemClass") == 0) {
				spObj.Release();
				continue;
			}
		}
		classes.push_back(spObj);
		spObj.Release();
	}
	return classes;
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumInstances(PCWSTR name, IWbemServices* pSvc, bool deep) {
	std::vector<CComPtr<IWbemClassObject>> instances;
	CComPtr<IEnumWbemClassObject> spEnum;
	auto hr = pSvc->CreateInstanceEnum(CComBSTR(name), (deep ? WBEM_FLAG_DEEP : 0) | WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &spEnum);
	if (FAILED(hr))
		return instances;

	CComPtr<IWbemClassObject> spObj;
	ULONG count;
	while (S_OK == spEnum->Next(WBEM_INFINITE, 1, &spObj, &count)) {
		instances.push_back(spObj);
		spObj.Release();
	}
	return instances;
}

bool WMIHelper::EnumInstancesAsync(HWND hWnd, UINT msg, PCWSTR name, IWbemServices* pSvc, bool deep) {
	CComObject<CObjectSink>* pSink;
	pSink->CreateInstance(&pSink);
	pSink->Init(hWnd, msg);
	auto hr = pSvc->CreateInstanceEnumAsync(CComBSTR(name), (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW), nullptr, pSink);
	if (FAILED(hr))
		return false;

	return true;
}

std::vector<WMIProperty> WMIHelper::EnumProperties(IWbemClassObject* pObj) {
	std::vector<WMIProperty> props;
	pObj->BeginEnumeration(0);
	WMIProperty prop;
	while (S_OK == pObj->Next(0, &prop.Name, &prop.Value, &prop.Type, &prop.Flavor)) {
		props.push_back(std::move(prop));
		prop.Value.Clear();
	}
	pObj->EndEnumeration();
	return props;
}

std::vector<WMIMethod> WMIHelper::EnumMethods(IWbemClassObject* pObj, bool localOnly, bool inheritedOnly) {
	std::vector<WMIMethod> methods;
	pObj->BeginMethodEnumeration((localOnly ? WBEM_FLAG_LOCAL_ONLY : 0) | (inheritedOnly ? WBEM_FLAG_PROPAGATED_ONLY : 0));
	CComBSTR name;
	WMIMethod method;
	while (S_OK == pObj->NextMethod(0, &name, method.spInParams.addressof(), method.spOutParams.addressof())) {
		method.Name = name.m_str;
		pObj->GetMethodOrigin(method.Name.c_str(), &name);
		method.ClassName = name.m_str;
		methods.push_back(std::move(method));
	}
	pObj->EndMethodEnumeration();
	return methods;
}

std::vector<CComBSTR> WMIHelper::GetNames(IWbemClassObject* pObj) {
	std::vector<CComBSTR> names;
	SAFEARRAY* sa;
	pObj->GetNames(nullptr, 0, nullptr, &sa);
	auto count = sa->cbElements;
	for (ULONG i = 0; i < count; i++) {
		LONG index = i;
		BSTR name;
		::SafeArrayGetElement(sa, &index, &name);
		CComBSTR bname;
		bname.Attach(name);
		names.push_back(bname);
	}
	::SafeArrayDestroy(sa);
	return names;
}

CString WMIHelper::GetStringProperty(IWbemClassObject* pObj, PCWSTR name) {
	CComVariant value;
	if (FAILED(pObj->Get(name, 0, &value, nullptr, nullptr)))
		return L"";

	if (value.vt != VT_BSTR)
		return L"";
	return value.bstrVal;
}
