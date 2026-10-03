#include "pch.h"
#include "WMIHelper.h"
#include <functional>

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
	HRESULT GetStatus() const override {
		return m_Status;
	}
	IWbemObjectSink* GetSink() override {
		return this;
	}

private:
	// Inherited via IWbemObjectSink
	HRESULT __stdcall Indicate(long lObjectCount, IWbemClassObject** apObjArray) override {
		for (int i = 0; i < lObjectCount; i++)
			m_Objects.push_back(apObjArray[i]);
		return S_OK;
	}
	HRESULT __stdcall SetStatus(long lFlags, HRESULT hr, BSTR strParam, IWbemClassObject* pObjParam) override {
		if (lFlags == WBEM_STATUS_COMPLETE) {
			m_Status = hr;
			//
			// the posted message holds a reference, released by the receiver
			//
			AddRef();
			if (!::PostMessage(m_hWnd, m_Msg, 0, reinterpret_cast<LPARAM>(static_cast<IObjectsCallback*>(this))))
				static_cast<IWbemObjectSink*>(this)->Release();
		}
		return S_OK;
	}

	HWND m_hWnd;
	UINT m_Msg;
	HRESULT m_Status{ S_OK };
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

//
// reads all the objects of an enumerator, many in each call (each call is a round trip to the WMI service)
//
static std::vector<CComPtr<IWbemClassObject>> ReadAll(IEnumWbemClassObject* pEnum, std::function<bool(IWbemClassObject*)> const& filter = nullptr) {
	std::vector<CComPtr<IWbemClassObject>> objects;
	IWbemClassObject* batch[256];
	ULONG count;
	HRESULT hr;
	do {
		count = 0;
		hr = pEnum->Next(WBEM_INFINITE, _countof(batch), batch, &count);
		for (ULONG i = 0; i < count; i++) {
			CComPtr<IWbemClassObject> spObj;
			spObj.Attach(batch[i]);
			if (!filter || filter(spObj))
				objects.push_back(std::move(spObj));
		}
	} while (hr == WBEM_S_NO_ERROR);	// WBEM_S_FALSE: the last (partial) batch
	return objects;
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumNamespaces(IWbemServices* pWmi) {
	CComPtr<IEnumWbemClassObject> spEnum;
	auto hr = pWmi->CreateInstanceEnum(CComBSTR(L"__NAMESPACE"), WBEM_FLAG_SHALLOW | WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &spEnum);
	if (FAILED(hr))
		return {};

	return ReadAll(spEnum);
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumClasses(IWbemServices* pSvc, bool deep, bool includeSystemClasses) {
	CComPtr<IEnumWbemClassObject> spEnum;
	auto hr = pSvc->CreateClassEnum(nullptr, (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW) | WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &spEnum);
	if (FAILED(hr))
		return {};

	if (includeSystemClasses)
		return ReadAll(spEnum);

	return ReadAll(spEnum, [](auto pObj) {
		return GetStringProperty(pObj, L"__DYNASTY").CompareNoCase(L"__SystemClass") != 0;
		});
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumInstances(PCWSTR name, IWbemServices* pSvc, bool deep) {
	CComPtr<IEnumWbemClassObject> spEnum;
	auto hr = pSvc->CreateInstanceEnum(CComBSTR(name), (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW) | WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &spEnum);
	if (FAILED(hr))
		return {};

	return ReadAll(spEnum);
}

HRESULT WMIHelper::EnumInstancesAsync(HWND hWnd, UINT msg, PCWSTR name, IWbemServices* pSvc, bool deep, IWbemObjectSink** ppSink) {
	CComObject<CObjectSink>* pSink;
	auto hr = CComObject<CObjectSink>::CreateInstance(&pSink);
	if (FAILED(hr))
		return hr;

	CComPtr<IWbemObjectSink> spSink(pSink);
	pSink->Init(hWnd, msg);
	hr = pSvc->CreateInstanceEnumAsync(CComBSTR(name), (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW), nullptr, spSink);
	if (FAILED(hr))
		return hr;

	*ppSink = spSink.Detach();
	return S_OK;
}

CString WMIHelper::GetErrorText(HRESULT hr) {
	CComPtr<IWbemStatusCodeText> spText;
	if (SUCCEEDED(spText.CoCreateInstance(__uuidof(WbemStatusCodeText)))) {
		CComBSTR text;
		if (SUCCEEDED(spText->GetErrorCodeText(hr, 0, 0, &text)) && text.Length() > 0) {
			CString result(text);
			result.TrimRight(L"\r\n ");
			return result;
		}
	}
	return std::format(L"Error 0x{:08X}", static_cast<ULONG>(hr)).c_str();
}

std::vector<WMIProperty> WMIHelper::EnumProperties(IWbemClassObject* pObj, long flags) {
	std::vector<WMIProperty> props;
	pObj->BeginEnumeration(flags);
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
