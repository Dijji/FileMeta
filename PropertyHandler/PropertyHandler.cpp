// Copyright (c) 2013, 2020 Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

#include <windows.h>
#include <shlwapi.h>
#include <propkey.h>
#include <propvarutil.h>
#include "dll.h"
#include "RegisterExtension.h"

#pragma region Logging support functions
#define LOGGING = 1
#ifdef LOGGING

#include <strsafe.h>
#include <iostream>
#include <string>
#include <fstream>
#include <cstdarg>
#include <vector>
#include <ctime>

std::wstring FormatGuid(GUID guid)
{
	WCHAR buffer[250];
	swprintf_s(buffer, 250, L"{%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
		guid.Data1, guid.Data2, guid.Data3,
		guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
		guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
	std::wstring result(buffer);
	return result;
}

std::wstring FormatVariant(REFPROPVARIANT propvar)
{
	WCHAR szBuffer[128];

	HRESULT hr = PropVariantToString(propvar, szBuffer, ARRAYSIZE(szBuffer));

	if (SUCCEEDED(hr) || hr == STRSAFE_E_INSUFFICIENT_BUFFER)
		return std::wstring(szBuffer);
	else
		return std::wstring(L"???");
}

std::wstring getCurrentDateTime(std::wstring s)
{
	time_t now = std::time(0);
	struct tm  tstruct;
	WCHAR  buf[80];
	localtime_s(&tstruct, &now);
	if (s == L"now")
		wcsftime(buf, sizeof(buf), L"%Y-%m-%d %X", &tstruct);
	else if (s == L"date")
		wcsftime(buf, sizeof(buf), L"%Y-%m-%d", &tstruct);
	return std::wstring(buf);
}

BOOL DirectoryExists(std::wstring dirName) {
	DWORD attribs = ::GetFileAttributes(dirName.c_str());
	if (attribs == INVALID_FILE_ATTRIBUTES) {
		return false;
	}
	return (attribs & FILE_ATTRIBUTE_DIRECTORY);
}

std::wstring PropName(REFPROPERTYKEY key)
{
	IPropertyDescription *pPropDesc;
	std::wstring result(L"Unknown");
	HRESULT hr = PSGetPropertyDescription(key, IID_IPropertyDescription, (void**)&pPropDesc);
	if (SUCCEEDED(hr))
	{
		LPWSTR pszName;
		hr = pPropDesc->GetCanonicalName(&pszName);
		if (SUCCEEDED(hr))
		{
			result = std::wstring(pszName);
			CoTaskMemFree(pszName);
		}
		SafeRelease(&pPropDesc);
	}
	return result;
}

void WriteLog(const WCHAR* fmt, ...)
{
	// Format the message
	WCHAR buffer[512];
	va_list vl;
	va_start(vl, fmt);
	_vsnwprintf_s(buffer, ARRAYSIZE(buffer), _TRUNCATE, fmt, vl);
	va_end(vl);

	// Write the message to the log
	std::wstring path = L"c:\\FilemetaLogs";
	if (!DirectoryExists(path) &&
		!CreateDirectory(path.c_str(), NULL))
		return; // Maybe throw an exception?
	std::wstring filePath = path + L"\\log_" + getCurrentDateTime(L"date") + L".log";
	std::wofstream ofs(filePath.c_str(), std::ios_base::out | std::ios_base::app);
	std::wstring now = getCurrentDateTime(L"now");
	ofs << now << L"  " << buffer << L'\n';
	ofs.flush();
	ofs.close();
}
#else
#define WriteLog(...) ((void)0)
#endif // Logging


#pragma endregion // logging functions 

extern PFN_STGOPENSTGEX v_pfnStgOpenStorageEx;

static const WCHAR* PropertyHandlerDescription = L"File Metadata Property Handler";

class CPropertyHandler : public IPropertyStore, public IInitializeWithFile
{
public:
    CPropertyHandler() : _cRef(1), _pStore(NULL), _bReadWrite(FALSE), _pChainedPropStore(NULL), _bHaveChainedPropCount(FALSE)
    {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void ** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CPropertyHandler, IPropertyStore),
			QITABENT(CPropertyHandler, IInitializeWithFile),
            {0, 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long cRef = InterlockedDecrement(&_cRef);
        if (cRef == 0)
        {
			WriteLog(L"Final release received for instance");
            delete this;
        }
        return cRef;
    }

    // IPropertyStore
    IFACEMETHODIMP GetCount(DWORD *pcProps);
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY *pkey);
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar);
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);
    IFACEMETHODIMP Commit();

	// IInitializeWithFile
	IFACEMETHODIMP Initialize(LPCWSTR pszFilePath,DWORD grfMode); 

private:
	
	~CPropertyHandler()
    {
        SafeRelease(&_pStore);
		SafeRelease(&_pChainedPropStore);
        DllRelease();
    }
	HRESULT OpenStore(BOOL bRead);
	DWORD ChainedPropCount();

    long _cRef;

	WCHAR					_pszFilePath[MAX_PATH];
	BOOL		            _bReadWrite;     // Whether storage set currently open read write
	IPropertyStore *		_pStore;		// Wrapper over set of storages.
    IPropertyStore *		_pChainedPropStore;	// Chained properties store
	BOOL					_bHaveChainedPropCount; // Whether we have read the count of chained properties
    DWORD					_cChainedPropCount;	// Count of properties in the chained properties store
};

HRESULT CPropertyHandler_CreateInstance(REFIID riid, void **ppv)
{
    HRESULT hr = E_OUTOFMEMORY;
    CPropertyHandler *pirm = new (std::nothrow) CPropertyHandler();
    if (pirm)
    {
        hr = pirm->QueryInterface(riid, ppv);
        pirm->Release();
    }
    return hr;
}

HRESULT CPropertyHandler::GetCount(DWORD *pcProps)
{
	try
	{
		WriteLog(L"GetCount called");
		HRESULT hr = OpenStore(FALSE);
		if (SUCCEEDED(hr))
		{
			hr = _pStore->GetCount(pcProps);
			if (SUCCEEDED(hr))
			{
				DWORD chained = ChainedPropCount();
				if (chained > 0)
					*pcProps += chained;
			}
		}
		else
			*pcProps = 0;
		WriteLog(L"GetCount returned 0x%08X with value %u", hr, *pcProps);
		return hr;
	}
	catch (...)
	{
		WriteLog(L"Unknown exception in GetCount");
		throw;
		//return E_FAIL;
	}
}

HRESULT CPropertyHandler::GetAt(DWORD iProp, PROPERTYKEY *pkey)
{
	try
	{
		WriteLog(L"GetAt called for index %d", iProp);
		//int * p = 0; *p = 1; // generate test access violation
		HRESULT hr = OpenStore(FALSE);
		// We take chained properties first because their number remains constant,
		// whereas the number of File Meta properties varies as they are set,
		// which would mean that if we took File Meta properties first,
		// repeated calls with the same index into the chained properties could return different key values
		if (SUCCEEDED(hr))
		{
			if (iProp < ChainedPropCount()) 
				hr = _pChainedPropStore->GetAt(iProp, pkey);
			else
				hr = _pStore->GetAt(iProp - ChainedPropCount(), pkey);
		}
		else
			*pkey = PKEY_Null;

		WriteLog(L"GetAt returned 0x%08X with value %s = %s %d", hr, PropName(*pkey).c_str(),
			FormatGuid(pkey->fmtid).c_str(), pkey->pid);
		return hr;
	}
	catch (...)
	{
		WriteLog(L"Unknown exception in GetAt");
		throw;
		//return E_FAIL;
	}
}

HRESULT CPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar)
{
	try
	{
		WriteLog(L"GetValue called for %s = %s %d pVar = 0x%X", PropName(key).c_str(), FormatGuid(key.fmtid).c_str(), key.pid, pPropVar);
		HRESULT hr = OpenStore(FALSE);
		
		if (SUCCEEDED(hr))
		{
			// Take the File Meta property value first, and if there isn't one,
			// see if this is a check for the software product name, which we use as a marker, or if not
			// try for a chained property value
			hr = _pStore->GetValue(key, pPropVar);
			if (SUCCEEDED(hr))
			{
				if (pPropVar->vt == VT_EMPTY)
				{
					if (key == PKEY_Software_ProductName)
						hr = InitPropVariantFromString(L"FileMetadata", pPropVar);
					else if (_pChainedPropStore != NULL)
					{
						WriteLog(L"Supplying value from the chained property store");
						hr = _pChainedPropStore->GetValue(key, pPropVar);
					}
				}
			}
		}
		else
			PropVariantClear(pPropVar);

		WriteLog(L"GetValue returned 0x%08X with value '%s' pVar = 0x%X", hr, FormatVariant(*pPropVar).c_str(), pPropVar);
		return hr;
	}
	catch (...)
	{
		WriteLog(L"Unknown exception in GetValue");
		throw;
		//return E_FAIL;
	}
}

// SetValue just updates the File Meta property store's value cache
HRESULT CPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
	try
	{
		WriteLog(L"SetValue called with value '%s' for %s = %s %d", FormatVariant(propVar).c_str(),
			PropName(key).c_str(), FormatGuid(key.fmtid).c_str(), key.pid);
		HRESULT hr = OpenStore(TRUE);
		if (SUCCEEDED(hr))
			hr = _pStore->SetValue(key, propVar);
		WriteLog(L"SetValue returned 0x%08X", hr);
		return hr;
	}
	catch (...)
	{
		WriteLog(L"Unknown exception in SetValue");
		throw;
		//return E_FAIL;
	}
}

// Commit writes updates out to thw alternate stream
HRESULT CPropertyHandler::Commit()
{
	try
	{
		WriteLog(L"Commit called");
		HRESULT hr = OpenStore(TRUE);
		if (SUCCEEDED(hr))
			hr = _pStore->Commit();
		WriteLog(L"Commit returned 0x%08X", hr);
	}
	catch (...)
	{
		WriteLog(L"Unknown exception in Commit");
		throw;
		//return E_FAIL;
	}
}

HRESULT CPropertyHandler::OpenStore(BOOL bReadWrite)
{
	if (_pStore)
		return S_OK;
	else
		return E_UNEXPECTED;
}

DWORD CPropertyHandler::ChainedPropCount()
{
	if (!_bHaveChainedPropCount)
	{
		_cChainedPropCount = 0;

		if (_pChainedPropStore != NULL)
		{
			WriteLog(L"Calling chained GetCount");
			_bHaveChainedPropCount = SUCCEEDED(_pChainedPropStore->GetCount(&_cChainedPropCount));
			WriteLog(L"Chained GetCount returned %u", _cChainedPropCount);
		}
		else
			_bHaveChainedPropCount = TRUE;
	}
	return _cChainedPropCount;
}

HRESULT CPropertyHandler::Initialize(LPCWSTR pszFilePath, DWORD grfMode)
{
	try
	{
		HRESULT hr = S_OK;
		WriteLog(L"Initialize for '%s' mode=0x%X", pszFilePath, grfMode);

		wcscpy_s(_pszFilePath, MAX_PATH, pszFilePath);

		if (!v_pfnStgOpenStorageEx)
		{
			BOOL fRunningOnNT = ((GetVersion() & 0x80000000) != 0x80000000);
			v_pfnStgOpenStorageEx = ((fRunningOnNT) ? (PFN_STGOPENSTGEX)GetProcAddress(GetModuleHandle(L"OLE32"), "StgOpenStorageEx") : NULL);
		}

		// In case of multiple calls to Initialize
		SafeRelease(&_pStore);
		SafeRelease(&_pChainedPropStore);

		if (v_pfnStgOpenStorageEx)
		{
			IPropertySetStorage* pPropSetStg = NULL;
			BOOL bReadWrite = (grfMode & STGM_READWRITE) == STGM_READWRITE;

			DWORD grfModeStorage = grfMode;
			if (grfMode == (STGM_READ | STGM_SHARE_DENY_NONE))
			{
				WriteLog(L"Upgrading storage from deny none to deny write so that the indexing service works");
				grfModeStorage = (STGM_READ | STGM_SHARE_DENY_WRITE);
			}

			WriteLog(L"Opening property store %s", bReadWrite ? L"Read/Write" : L"Read only");
			hr = (v_pfnStgOpenStorageEx)(_pszFilePath, grfModeStorage, STGFMT_FILE, 0, NULL, 0,
				IID_IPropertySetStorage, (void**)&pPropSetStg);

			if (SUCCEEDED(hr))
				// To make IPropertyStore work for Write, it is necessary to use STGM_READWRITE, which the MS documentation says
				// explicitly will not work.  The recommended STGM_READ fails with E_ACCESSDENIED on Write and Commit, which
				// is what you would expect. The only bug appears to be in the documentation.
				hr = PSCreatePropertyStoreFromPropertySetStorage(pPropSetStg, grfMode, IID_IPropertyStore, (void **)&_pStore);

			SafeRelease(&pPropSetStg);

			WriteLog(L"Open property store returned 0x%08X", hr);
			if (SUCCEEDED(hr))
				_bReadWrite = bReadWrite;
			else
				_pStore = NULL;
		}
		else
			hr = E_UNEXPECTED;

		// Check if a chained property handler is configured, and if there is one, load and initialize it too.
		// This is simplified by the fact that we only ever open a chained property handler read-only
		PCWSTR pszExt = PathFindExtension(_pszFilePath);
		if (SUCCEEDED(hr) && *pszExt)
		{
			// Chained property handlers are configured in a Chained property value added to the standard property system key
			HKEY propertyHandlerKey;
			WCHAR szKeyName[MAX_PATH] = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\";
			wcscat_s(szKeyName, MAX_PATH, pszExt);
			if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyName, 0, KEY_QUERY_VALUE, &propertyHandlerKey) == ERROR_SUCCESS)
			{
				WCHAR szBuf[64];
				DWORD nBufLen = sizeof(szBuf);
				DWORD dwType;

				if (RegGetValue(propertyHandlerKey, NULL, L"Chained", RRF_RT_REG_SZ, &dwType, szBuf, &nBufLen) == ERROR_SUCCESS)
				{
					CLSID clsid;
					if (SUCCEEDED(CLSIDFromString(szBuf, &clsid)))
					{
						WriteLog(L"Initializing read-only chained property handler: %s", FormatGuid(clsid).c_str());
						HRESULT hr2 = CoCreateInstance(clsid, NULL, CLSCTX_ALL, IID_IPropertyStore, (void **)&_pChainedPropStore);
						if (SUCCEEDED(hr2))
						{
							IInitializeWithFile *pChainedPropInit;
							hr2 = _pChainedPropStore->QueryInterface(IID_IInitializeWithFile, (void **)&pChainedPropInit);
							if (SUCCEEDED(hr2))
							{
								hr2 = pChainedPropInit->Initialize(_pszFilePath, STGM_READ);
								pChainedPropInit->Release();
							}
							else
							{
								IInitializeWithStream *pChainedPropInitWithStream;
								hr2 = _pChainedPropStore->QueryInterface(IID_IInitializeWithStream, (void **)&pChainedPropInitWithStream);
								if (SUCCEEDED(hr2))
								{
									IStream *pStream;
									hr2 = SHCreateStreamOnFileEx(_pszFilePath, STGM_READ, 0, FALSE, NULL, &pStream);
									if (SUCCEEDED(hr2))
									{
										hr2 = pChainedPropInitWithStream->Initialize(pStream, STGM_READ);
										pStream->Release();
									}
									pChainedPropInitWithStream->Release();
								}
							}
							if (FAILED(hr2))
							{
								_pChainedPropStore->Release();
								_pChainedPropStore = NULL;
							}
						}
						WriteLog(L"Chained property store Initialize returned 0x%08X", hr);
					}
				}

				RegCloseKey(propertyHandlerKey);
			}
		}

		return hr;
	}
	catch (...)
	{
		WriteLog(L"Unknown exception in Initialize");
		throw;
		//return E_FAIL;
	}
}

HRESULT RegisterPropertyHandler()
{
    // register the property handler COM object, and set the options it uses
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.RegisterInProcServer(PropertyHandlerDescription, L"Both");
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterInProcServerAttribute(L"DisableProcessIsolation", TRUE);
        if (SUCCEEDED(hr))
		{
			hr = re.RegisterInProcServerAttribute(L"ManualSafeSave", TRUE);
			if (SUCCEEDED(hr))
			{
				hr = re.RegisterInProcServerAttribute(L"EnableShareDenyWrite", TRUE);
				if (SUCCEEDED(hr))
				{
					hr = re.RegisterInProcServerAttribute(L"EnableShareDenyNone", TRUE);
				}
			}
		}
    }

    return hr;
}

HRESULT UnregisterPropertyHandler()
{
    // Unregister the property handler COM object.
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();

    return hr;
}
