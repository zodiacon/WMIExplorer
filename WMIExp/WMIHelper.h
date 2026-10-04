#pragma once

#include <wil\com.h>
#include <atomic>

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
// the computer (and credentials) to connect to.
// a connection is never freed (see WMIHelper::NewConnection): proxies keep pointing to its Identity
//
struct WMIConnection {
	CString Computer;	// empty: the local computer
	CString User;		// empty: the current user. "domain\user" or "user@domain"
	CString Password;

	bool HasCredentials() const {
		return !User.IsEmpty();
	}

private:
	friend struct WMIHelper;
	CString m_Domain, m_UserName;
	COAUTHIDENTITY m_Identity{};
};

//
// a query (or instance enumeration) running on a worker thread.
// when it is done, a pointer to a std::shared_ptr<WMIQueryJob> is posted (as LPARAM) to the window;
// the receiver must delete it (WMIHelper::TakeJob does), even for results it ignores
//
struct WMIQueryJob {
	std::vector<CComPtr<IWbemClassObject>> Objects;
	HRESULT Status{ S_OK };
	std::atomic<bool> Cancelled{ false };
};

struct WMIHelper abstract final {
	// connects to a namespace of the computer (and with the credentials) of the connection
	static HRESULT Connect(WMIConnection const* conn, PCWSTR ns, IWbemServices** ppWmi);
	static HRESULT OpenNamespace(IWbemServices* pParent, PCWSTR path, IWbemServices** ppWmi);
	// the connection the main window uses (and new windows and threads start with)
	static WMIConnection const* CurrentConnection();
	static WMIConnection const* NewConnection(CString const& computer, CString const& user, CString const& password);
	static void SetCurrentConnection(WMIConnection const* conn);
	// a proxy (services or enumerator) must have the connection's credentials set
	static HRESULT SetSecurity(IUnknown* pProxy, WMIConnection const* conn);

	static CString GetStringProperty(IWbemClassObject* pObj, PCWSTR name);
	static std::vector<CComPtr<IWbemClassObject>> EnumNamespaces(IWbemServices* pWmi);
	static std::vector<CComPtr<IWbemClassObject>> EnumClasses(IWbemServices* pSvc, bool deep, bool includeSystemClasses = false);
	static std::vector<CComPtr<IWbemClassObject>> EnumInstances(PCWSTR name, IWbemServices* pSvc, bool deep);

	// on a worker thread; the result is posted to hWnd with msg
	static std::shared_ptr<WMIQueryJob> EnumInstancesAsync(HWND hWnd, UINT msg, PCWSTR className, IWbemServices* pSvc, bool deep);
	static std::shared_ptr<WMIQueryJob> ExecQueryAsync(HWND hWnd, UINT msg, PCWSTR query, IWbemServices* pSvc);
	static std::shared_ptr<WMIQueryJob> TakeJob(LPARAM lParam);

	static CString GetErrorText(HRESULT hr);
	static std::vector<WMIProperty> EnumProperties(IWbemClassObject* pObj, long flags = 0);
	static std::vector<WMIMethod> EnumMethods(IWbemClassObject* pObj, bool localOnly = false, bool inheritedOnly = false);
	static std::vector<CComBSTR> GetNames(IWbemClassObject* pObj);

	// descriptions (and other qualifiers) need an object fetched with WBEM_FLAG_USE_AMENDED_QUALIFIERS
	static CString GetClassDescription(IWbemClassObject* pClass);
	static CString GetPropertyDescription(IWbemClassObject* pClass, PCWSTR name);
	static CString GetMethodDescription(IWbemClassObject* pClass, PCWSTR name);
	static bool IsStaticMethod(IWbemClassObject* pClass, PCWSTR name);
	static CString GetObjectText(IWbemClassObject* pObj);

	// values as text
	static CString CimTypeToString(CIMTYPE type);
	static CString FormatValue(CComVariant const& value, CIMTYPE type);
	static CString FormatDateTime(PCWSTR dmtf);
	static int CompareValues(CComVariant const& v1, CComVariant const& v2, CIMTYPE type);
	// text to a value that can be Put into a property of the given type (arrays: comma separated)
	static HRESULT ParseValue(CString const& text, CIMTYPE type, CComVariant& value);

private:
	static CString GetArrayValue(CComVariant const& value, CIMTYPE type);
	static CString GetQualifier(IWbemQualifierSet* pSet, PCWSTR name);
	static std::shared_ptr<WMIQueryJob> StartJob(HWND hWnd, UINT msg, IWbemServices* pSvc, std::function<HRESULT(IWbemServices*, IEnumWbemClassObject**)> start);
};
