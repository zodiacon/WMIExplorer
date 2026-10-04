#include "pch.h"
#include "WMIHelper.h"

namespace {
	// connections are never freed: proxies (in any thread) keep pointing to their identity
	std::vector<std::unique_ptr<WMIConnection>> s_Connections;
	std::atomic<WMIConnection const*> s_CurrentConnection;
}

WMIConnection const* WMIHelper::CurrentConnection() {
	if (s_CurrentConnection == nullptr)
		s_CurrentConnection = NewConnection(L"", L"", L"");
	return s_CurrentConnection;
}

void WMIHelper::SetCurrentConnection(WMIConnection const* conn) {
	s_CurrentConnection = conn;
}

WMIConnection const* WMIHelper::NewConnection(CString const& computer, CString const& user, CString const& password) {
	auto conn = std::make_unique<WMIConnection>();
	conn->Computer = computer;
	conn->User = user;
	conn->Password = password;
	if (auto slash = user.Find(L'\\'); slash >= 0) {
		conn->m_Domain = user.Left(slash);
		conn->m_UserName = user.Mid(slash + 1);
	}
	else {
		conn->m_UserName = user;		// (user@domain has no separate domain)
	}
	auto& id = conn->m_Identity;
	id.User = (USHORT*)(PCWSTR)conn->m_UserName;
	id.UserLength = conn->m_UserName.GetLength();
	id.Domain = conn->m_Domain.IsEmpty() ? nullptr : (USHORT*)(PCWSTR)conn->m_Domain;
	id.DomainLength = conn->m_Domain.GetLength();
	id.Password = (USHORT*)(PCWSTR)conn->Password;
	id.PasswordLength = conn->Password.GetLength();
	id.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

	s_Connections.push_back(std::move(conn));
	return s_Connections.back().get();
}

HRESULT WMIHelper::SetSecurity(IUnknown* pProxy, WMIConnection const* conn) {
	if (pProxy == nullptr || conn == nullptr || conn->Computer.IsEmpty())
		return S_OK;		// local: the process defaults (CoInitializeSecurity) are used

	// remote: encrypted (some namespaces require it), with the connection's credentials (if any)
	auto identity = conn->HasCredentials() ? const_cast<COAUTHIDENTITY*>(&conn->m_Identity) : nullptr;
	auto set = [&](IUnknown* p) {
		return ::CoSetProxyBlanket(p, RPC_C_AUTHN_DEFAULT, RPC_C_AUTHZ_DEFAULT, COLE_DEFAULT_PRINCIPAL,
			RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE, identity, EOAC_NONE);
	};
	auto hr = set(pProxy);
	if (FAILED(hr))
		return hr;

	// the proxy's IUnknown is separate (used for QueryInterface and Release), and needs the same
	CComPtr<IUnknown> spUnk;
	if (SUCCEEDED(pProxy->QueryInterface(&spUnk)) && spUnk != pProxy)
		set(spUnk);
	return S_OK;
}

HRESULT WMIHelper::Connect(WMIConnection const* conn, PCWSTR ns, IWbemServices** ppWmi) {
	CComPtr<IWbemLocator> spLocator;
	auto hr = spLocator.CoCreateInstance(__uuidof(WbemLocator));
	if (FAILED(hr))
		return hr;

	CString path(ns);
	if (!conn->Computer.IsEmpty())
		path = L"\\\\" + conn->Computer + L"\\" + path;
	CComBSTR user, password;
	if (conn->HasCredentials()) {
		user = conn->User;
		password = conn->Password;
	}
	CComPtr<IWbemServices> spWmi;
	hr = spLocator->ConnectServer(CComBSTR(path), user, password, nullptr, WBEM_FLAG_CONNECT_USE_MAX_WAIT, nullptr, nullptr, &spWmi);
	if (FAILED(hr))
		return hr;

	hr = SetSecurity(spWmi, conn);
	if (FAILED(hr))
		return hr;

	*ppWmi = spWmi.Detach();
	return S_OK;
}

HRESULT WMIHelper::OpenNamespace(IWbemServices* pParent, PCWSTR path, IWbemServices** ppWmi) {
	CComPtr<IWbemServices> spWmi;
	auto hr = pParent->OpenNamespace(CComBSTR(path), 0, nullptr, &spWmi, nullptr);
	if (FAILED(hr))
		return hr;

	SetSecurity(spWmi, CurrentConnection());
	*ppWmi = spWmi.Detach();
	return S_OK;
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

	SetSecurity(spEnum, CurrentConnection());
	return ReadAll(spEnum);
}

std::vector<CComPtr<IWbemClassObject>> WMIHelper::EnumClasses(IWbemServices* pSvc, bool deep, bool includeSystemClasses) {
	CComPtr<IEnumWbemClassObject> spEnum;
	auto hr = pSvc->CreateClassEnum(nullptr, (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW) | WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &spEnum);
	if (FAILED(hr))
		return {};

	SetSecurity(spEnum, CurrentConnection());
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

	SetSecurity(spEnum, CurrentConnection());
	return ReadAll(spEnum);
}

static void PostJob(HWND hWnd, UINT msg, std::shared_ptr<WMIQueryJob> const& job) {
	auto p = new std::shared_ptr<WMIQueryJob>(job);
	if (!::PostMessage(hWnd, msg, 0, reinterpret_cast<LPARAM>(p)))
		delete p;
}

std::shared_ptr<WMIQueryJob> WMIHelper::TakeJob(LPARAM lParam) {
	std::unique_ptr<std::shared_ptr<WMIQueryJob>> p(reinterpret_cast<std::shared_ptr<WMIQueryJob>*>(lParam));
	return *p;
}

//
// semi-synchronous on a worker thread, rather than asynchronous with a sink:
// a sink is called back by the WMI service, which fails with remote computers in many setups
//
std::shared_ptr<WMIQueryJob> WMIHelper::StartJob(HWND hWnd, UINT msg, IWbemServices* pSvc, std::function<HRESULT(IWbemServices*, IEnumWbemClassObject**)> start) {
	auto job = std::make_shared<WMIQueryJob>();
	// the proxy belongs to the caller's apartment: the worker gets its own
	IStream* pStream = nullptr;
	auto hr = ::CoMarshalInterThreadInterfaceInStream(__uuidof(IWbemServices), pSvc, &pStream);
	if (FAILED(hr)) {
		job->Status = hr;
		PostJob(hWnd, msg, job);
		return job;
	}

	auto conn = CurrentConnection();
	std::thread([=]() {
		::CoInitializeEx(nullptr, COINIT_MULTITHREADED);
		{
			CComPtr<IWbemServices> spSvc;
			auto hr = ::CoGetInterfaceAndReleaseStream(pStream, __uuidof(IWbemServices), reinterpret_cast<void**>(&spSvc));
			if (SUCCEEDED(hr)) {
				SetSecurity(spSvc, conn);
				CComPtr<IEnumWbemClassObject> spEnum;
				hr = start(spSvc, &spEnum);
				if (SUCCEEDED(hr)) {
					SetSecurity(spEnum, conn);
					IWbemClassObject* batch[256];
					while (!job->Cancelled) {
						ULONG count = 0;
						// a timeout, so a cancelled job does not wait for a slow query
						hr = spEnum->Next(500, _countof(batch), batch, &count);
						for (ULONG i = 0; i < count; i++) {
							CComPtr<IWbemClassObject> spObj;
							spObj.Attach(batch[i]);
							job->Objects.push_back(std::move(spObj));
						}
						if (hr != WBEM_S_NO_ERROR && hr != WBEM_S_TIMEDOUT)
							break;
					}
					if (hr == WBEM_S_FALSE || hr == WBEM_S_TIMEDOUT)
						hr = S_OK;
				}
			}
			job->Status = hr;
		}
		if (!job->Cancelled)
			PostJob(hWnd, msg, job);
		::CoUninitialize();
		}).detach();
	return job;
}

std::shared_ptr<WMIQueryJob> WMIHelper::EnumInstancesAsync(HWND hWnd, UINT msg, PCWSTR className, IWbemServices* pSvc, bool deep) {
	return StartJob(hWnd, msg, pSvc, [name = CComBSTR(className), deep](auto pSvc, auto ppEnum) {
		return pSvc->CreateInstanceEnum(name, (deep ? WBEM_FLAG_DEEP : WBEM_FLAG_SHALLOW) | WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, ppEnum);
		});
}

std::shared_ptr<WMIQueryJob> WMIHelper::ExecQueryAsync(HWND hWnd, UINT msg, PCWSTR query, IWbemServices* pSvc) {
	return StartJob(hWnd, msg, pSvc, [text = CComBSTR(query)](auto pSvc, auto ppEnum) {
		return pSvc->ExecQuery(CComBSTR(L"WQL"), text, WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, ppEnum);
		});
}

CString WMIHelper::GetQualifier(IWbemQualifierSet* pSet, PCWSTR name) {
	CComVariant value;
	if (pSet == nullptr || FAILED(pSet->Get(name, 0, &value, nullptr)) || value.vt != VT_BSTR)
		return L"";
	return value.bstrVal;
}

CString WMIHelper::GetClassDescription(IWbemClassObject* pClass) {
	CComPtr<IWbemQualifierSet> spSet;
	pClass->GetQualifierSet(&spSet);
	return GetQualifier(spSet, L"Description");
}

CString WMIHelper::GetPropertyDescription(IWbemClassObject* pClass, PCWSTR name) {
	CComPtr<IWbemQualifierSet> spSet;
	pClass->GetPropertyQualifierSet(name, &spSet);
	return GetQualifier(spSet, L"Description");
}

CString WMIHelper::GetMethodDescription(IWbemClassObject* pClass, PCWSTR name) {
	CComPtr<IWbemQualifierSet> spSet;
	pClass->GetMethodQualifierSet(name, &spSet);
	return GetQualifier(spSet, L"Description");
}

bool WMIHelper::IsStaticMethod(IWbemClassObject* pClass, PCWSTR name) {
	CComPtr<IWbemQualifierSet> spSet;
	CComVariant value;
	if (FAILED(pClass->GetMethodQualifierSet(name, &spSet)) || FAILED(spSet->Get(L"Static", 0, &value, nullptr)))
		return false;
	return value.vt == VT_BOOL && value.boolVal;
}

CString WMIHelper::GetObjectText(IWbemClassObject* pObj) {
	CComBSTR text;
	if (FAILED(pObj->GetObjectText(0, &text)))
		return L"";
	// for an edit control
	CString result(text);
	result.Replace(L"\r\n", L"\n");
	result.Replace(L"\n", L"\r\n");
	return result;
}

static VARTYPE CimTypeToVarType(CIMTYPE type) {
	switch (type) {
		case CIM_SINT8: case CIM_SINT16: case CIM_CHAR16: return VT_I2;
		case CIM_UINT8: return VT_UI1;
		case CIM_UINT16: case CIM_SINT32: case CIM_UINT32: return VT_I4;
		case CIM_REAL32: return VT_R4;
		case CIM_REAL64: return VT_R8;
		case CIM_BOOLEAN: return VT_BOOL;
	}
	return VT_BSTR;		// strings, dates, references, and 64-bit integers
}

HRESULT WMIHelper::ParseValue(CString const& text, CIMTYPE type, CComVariant& value) {
	value.Clear();
	if (type & CIM_FLAG_ARRAY) {
		auto elementType = type & ~CIM_FLAG_ARRAY;
		std::vector<CComVariant> elements;
		int start = 0;
		for (auto token = text.Tokenize(L",", start); start >= 0; token = text.Tokenize(L",", start)) {
			CComVariant element;
			auto hr = ParseValue(token.Trim(), elementType, element);
			if (FAILED(hr))
				return hr;
			elements.push_back(std::move(element));
		}
		auto vt = CimTypeToVarType(elementType);
		auto sa = ::SafeArrayCreateVector(vt, 0, (ULONG)elements.size());
		if (sa == nullptr)
			return E_OUTOFMEMORY;
		for (LONG i = 0; i < (LONG)elements.size(); i++) {
			auto& e = elements[i];
			// a BSTR is passed as is (and copied); other types by address
			::SafeArrayPutElement(sa, &i, vt == VT_BSTR ? static_cast<void*>(e.bstrVal) : static_cast<void*>(&e.llVal));
		}
		value.vt = VT_ARRAY | vt;
		value.parray = sa;
		return S_OK;
	}

	switch (type) {
		case CIM_STRING: case CIM_DATETIME: case CIM_REFERENCE:
			value = text;
			return S_OK;

		case CIM_CHAR16:
			if (text.GetLength() != 1)
				return E_INVALIDARG;
			value = (short)text[0];
			return S_OK;

		case CIM_BOOLEAN:
			if (text.CompareNoCase(L"true") == 0 || text == L"1")
				value = true;
			else if (text.CompareNoCase(L"false") == 0 || text == L"0")
				value = false;
			else
				return E_INVALIDARG;
			return S_OK;

		case CIM_REAL32: case CIM_REAL64:
		{
			PWSTR end;
			auto number = wcstod(text, &end);
			if (text.IsEmpty() || *end)
				return E_INVALIDARG;
			if (type == CIM_REAL32)
				value = (float)number;
			else
				value = number;
			return S_OK;
		}

		case CIM_SINT8: case CIM_UINT8: case CIM_SINT16: case CIM_UINT16:
		case CIM_SINT32: case CIM_UINT32: case CIM_SINT64: case CIM_UINT64:
		{
			PWSTR end;
			// decimal, or hex with 0x
			auto number = type == CIM_UINT64 ? (long long)_wcstoui64(text, &end, 0) : _wcstoi64(text, &end, 0);
			if (text.IsEmpty() || *end)
				return E_INVALIDARG;
			switch (CimTypeToVarType(type)) {
				case VT_I2: value = (short)number; break;
				case VT_UI1: value = (BYTE)number; break;
				case VT_I4: value = (long)number; break;
				default:
					// 64-bit integers are passed as strings
					value = type == CIM_UINT64 ? std::to_wstring((unsigned long long)number).c_str() : std::to_wstring(number).c_str();
					break;
			}
			return S_OK;
		}
	}
	return E_NOTIMPL;	// embedded objects
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

CString WMIHelper::CimTypeToString(CIMTYPE type) {
	CString text;
	switch (type & 0xff) {
		case CIM_EMPTY: text = L"Empty"; break;
		case CIM_SINT8: text = L"Signed Byte (8 bit)"; break;
		case CIM_UINT8: text = L"Byte (8 bit)"; break;
		case CIM_SINT16: text = L"Signed Word (16 bit)"; break;
		case CIM_UINT16: text = L"Word (16 bit)"; break;
		case CIM_SINT32: text = L"Signed Int (32 bit)"; break;
		case CIM_UINT32: text = L"Int (32 bit)"; break;
		case CIM_SINT64: text = L"Signed QWord (64 bit)"; break;
		case CIM_UINT64: text = L"QWord (64 bit)"; break;
		case CIM_REAL32: text = L"Real (32 bit)"; break;
		case CIM_REAL64: text = L"Real (64 bit)"; break;
		case CIM_BOOLEAN: text = L"Boolean"; break;
		case CIM_STRING: text = L"String"; break;
		case CIM_DATETIME: text = L"Date Time"; break;
		case CIM_REFERENCE: text = L"Reference"; break;
		case CIM_CHAR16: text = L"Character"; break;
		case CIM_OBJECT: text = L"Object"; break;
	}
	if (type & CIM_FLAG_ARRAY)
		text += L" [Array]";
	return text;
}

CString WMIHelper::GetArrayValue(CComVariant const& value, CIMTYPE type) {
	if ((value.vt & VT_ARRAY) == 0 || value.parray == nullptr)
		return L"";

	auto sa = value.parray;
	LONG lower = 0, upper = -1;
	::SafeArrayGetLBound(sa, 1, &lower);
	::SafeArrayGetUBound(sa, 1, &upper);

	CString text;
	if ((type & ~CIM_FLAG_ARRAY) == CIM_UINT8 || (type & ~CIM_FLAG_ARRAY) == CIM_SINT8) {
		// bytes are shown in hex (at most 64)
		BYTE* data;
		auto count = std::min(upper - lower + 1, 64L);
		if (SUCCEEDED(::SafeArrayAccessData(sa, reinterpret_cast<void**>(&data)))) {
			for (LONG i = 0; i < count; i++)
				text += std::format(L"{:02X} ", data[i]).c_str();
			::SafeArrayUnaccessData(sa);
		}
		return text.TrimRight();
	}

	VARTYPE vt;
	if (FAILED(::SafeArrayGetVartype(sa, &vt)))
		return L"";

	const LONG maxCount = 100;
	for (LONG i = lower; i <= upper && i < lower + maxCount; i++) {
		// the element is copied into the variant's data (the union is the same for all types)
		CComVariant element;
		if (vt == VT_VARIANT) {
			if (FAILED(::SafeArrayGetElement(sa, &i, &element)))
				continue;
		}
		else {
			if (FAILED(::SafeArrayGetElement(sa, &i, &element.llVal)))
				continue;
			element.vt = vt;
		}
		if (i > lower)
			text += L", ";
		text += FormatValue(element, type & ~CIM_FLAG_ARRAY);
	}
	if (upper - lower + 1 > maxCount)
		text += L", ...";
	return text;
}

CString WMIHelper::FormatDateTime(PCWSTR dmtf) {
	// date and time: yyyymmddHHMMSS.mmmmmmsUUU, interval: ddddddddHHMMSS.mmmmmm:000
	std::wstring_view s(dmtf);
	if (s.size() < 25 || !std::all_of(s.begin(), s.begin() + 14, [](wchar_t ch) { return ch >= L'0' && ch <= L'9'; }))
		return dmtf;		// (fields can be wildcards)

	if (s[21] == L':')
		return std::format(L"{} days, {}:{}:{}", _wtoi(std::wstring(s.substr(0, 8)).c_str()),
			s.substr(8, 2), s.substr(10, 2), s.substr(12, 2)).c_str();
	return std::format(L"{}-{}-{} {}:{}:{}", s.substr(0, 4), s.substr(4, 2), s.substr(6, 2),
		s.substr(8, 2), s.substr(10, 2), s.substr(12, 2)).c_str();
}

CString WMIHelper::FormatValue(CComVariant const& value, CIMTYPE type) {
	if (value.vt == VT_NULL || value.vt == VT_EMPTY)
		return L"";

	if (type & CIM_FLAG_ARRAY)
		return GetArrayValue(value, type);
	if (value.vt == VT_BOOL)
		return value.boolVal ? L"True" : L"False";
	if (type == CIM_DATETIME && value.vt == VT_BSTR)
		return FormatDateTime(value.bstrVal);
	if (type == CIM_OBJECT)
		return L"(Object)";

	CComVariant text;
	if (SUCCEEDED(text.ChangeType(VT_BSTR, &value)))
		return CString(text.bstrVal);
	return L"";
}

int WMIHelper::CompareValues(CComVariant const& v1, CComVariant const& v2, CIMTYPE type) {
	bool null1 = v1.vt == VT_NULL || v1.vt == VT_EMPTY;
	bool null2 = v2.vt == VT_NULL || v2.vt == VT_EMPTY;
	if (null1 || null2)
		return null1 == null2 ? 0 : (null1 ? -1 : 1);

	if ((type & CIM_FLAG_ARRAY) == 0 && v1.vt == v2.vt) {
		// 64-bit integers come as strings
		if (type == CIM_UINT64 && v1.vt == VT_BSTR) {
			auto n1 = _wcstoui64(v1.bstrVal, nullptr, 10), n2 = _wcstoui64(v2.bstrVal, nullptr, 10);
			return (n1 > n2) - (n1 < n2);
		}
		if (type == CIM_SINT64 && v1.vt == VT_BSTR) {
			auto n1 = _wcstoi64(v1.bstrVal, nullptr, 10), n2 = _wcstoi64(v2.bstrVal, nullptr, 10);
			return (n1 > n2) - (n1 < n2);
		}
		switch (::VarCmp(const_cast<VARIANT*>(static_cast<const VARIANT*>(&v1)), const_cast<VARIANT*>(static_cast<const VARIANT*>(&v2)),
			LOCALE_USER_DEFAULT, NORM_IGNORECASE)) {
			case VARCMP_LT: return -1;
			case VARCMP_EQ: return 0;
			case VARCMP_GT: return 1;
		}
	}
	return FormatValue(v1, type).CompareNoCase(FormatValue(v2, type));
}
