// Copyright (c) 2013, 2020 Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

#include <windows.h>
#include <shlwapi.h>
#include <propkey.h>
#include <strsafe.h>
#include <propvarutil.h>
#include "dll.h"
#include "RegisterExtension.h"
#include <iostream>
#include <string>
#include <fstream>
#include <cstdarg>
#include <vector>
#include <ctime>

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
	HRESULT OpenStore(BOOL bReadWrite);
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

void format(std::string& a_string, const char* fmt, ...)
{
	va_list vl;
	va_start(vl, fmt);
	int size = _vscprintf(fmt, vl);
	if (size > 0)
	{
		a_string.resize(++size);
		vsnprintf_s((char*)a_string.data(), size, _TRUNCATE, fmt, vl);
	}
	va_end(vl);
}

std::string FormatGuid(GUID guid)
{
	std::string result;
	format (result, "{%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
		guid.Data1, guid.Data2, guid.Data3,
		guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
		guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
	return result;
}

std::wstring FormatVariant(REFPROPVARIANT propvar)
{
	WCHAR szBuffer[128];

	HRESULT hr = PropVariantToString(propvar, szBuffer, ARRAYSIZE(szBuffer));

	if (SUCCEEDED(hr) || hr == STRSAFE_E_INSUFFICIENT_BUFFER)
		return szBuffer;
	else
		return L"???";
}

std::string getCurrentDateTime(std::string s)
{
	time_t now = std::time(0);
	struct tm  tstruct;
	char  buf[80];
	localtime_s(&tstruct, &now);
	if (s == "now")
		strftime(buf, sizeof(buf), "%Y-%m-%d %X", &tstruct);
	else if (s == "date")
		strftime(buf, sizeof(buf), "%Y-%m-%d", &tstruct);
	return std::string(buf);
}

BOOL DirectoryExists(std::string dirName) {
	DWORD attribs = ::GetFileAttributesA(dirName.c_str());
	if (attribs == INVALID_FILE_ATTRIBUTES) {
		return false;
	}
	return (attribs & FILE_ATTRIBUTE_DIRECTORY);
}


void WriteLogMsg(const std::string& message)
{
	// Write the message to the log
	std::string path = "c:\\Filemeta logs";
	if (!DirectoryExists(path) &&
		!CreateDirectoryA(path.c_str(), NULL))
			return; // Maybe throw an exception?
	std::string filePath = path + "\\log_" + getCurrentDateTime("date") + ".txt";
	std::string now = getCurrentDateTime("now");
	std::ofstream ofs(filePath.c_str(), std::ios_base::out | std::ios_base::app);
	ofs << now << '\t' << message << '\n';
	ofs.close();
}
void WriteLog(const char* fmt, ...)
{
	// Format the message
	std::string message;
	va_list vl;
	va_start(vl, fmt);
	format(message, fmt, vl);
	va_end(vl);
	WriteLogMsg(message);
}

HRESULT CPropertyHandler::GetCount(DWORD *pcProps)
{
	try
	{
		WriteLogMsg("GetCount called");
		*pcProps = 0;
		HRESULT hr = OpenStore(FALSE);
		hr = SUCCEEDED(hr) ? _pStore->GetCount(pcProps) : hr;
		if (SUCCEEDED(hr))
			*pcProps += ChainedPropCount();
		WriteLog("GetCount returning 0x%X with value %u", hr, *pcProps);
		return hr;
	}
	catch (...)
	{
		WriteLogMsg("Unknown exception in GetCount");
		return E_FAIL;
	}
}

HRESULT CPropertyHandler::GetAt(DWORD iProp, PROPERTYKEY *pkey)
{
	try
	{
		WriteLog("GetAt called for index %d", iProp);
		*pkey = PKEY_Null;
		HRESULT hr = OpenStore(FALSE);
		// We take chained properties first because their number remains constant,
		// whereas the number of File Meta properties varies as they are set,
		// which would mean that if we took File Meta properties first,
		// repeated calls with the same index into the chained properties could return different key values
		hr = SUCCEEDED(hr) ?
			(iProp < ChainedPropCount()) ?
			_pChainedPropStore->GetAt(iProp, pkey) :
			_pStore->GetAt(iProp - ChainedPropCount(), pkey) :
			hr;
		WriteLog("GetAt returning 0x%X with value %s %d", hr, FormatGuid(pkey->fmtid), pkey->pid);
		return hr;
	}
	catch (...)
	{
		WriteLogMsg("Unknown exception in GetAt");
		return E_FAIL;
	}
}

HRESULT CPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar)
{
	try
	{
		WriteLog("GetValue called for %s %d", FormatGuid(key.fmtid), key.pid);
		PropVariantInit(pPropVar);
		HRESULT hr = OpenStore(FALSE);
		// Take the File Meta property value first, and if there isn't one,
		// see if this is a check for the software product name, which we use as a marker, or if not
		// try for a chained property value
		hr = SUCCEEDED(hr) ? _pStore->GetValue(key, pPropVar) : hr;
		if (SUCCEEDED(hr) && pPropVar->vt == VT_EMPTY && key == PKEY_Software_ProductName)
			hr = InitPropVariantFromString(L"FileMetadata", pPropVar);
		else if (_pChainedPropStore != NULL && SUCCEEDED(hr) && pPropVar->vt == VT_EMPTY)
		{
			WriteLogMsg("Value will be taken from chained property store");
			hr = _pChainedPropStore->GetValue(key, pPropVar);
		}
		WriteLog("GetValue returning 0x%X with value %S", hr, FormatVariant(*pPropVar));
		return hr;
	}
	catch (...)
	{
		WriteLogMsg("Unknown exception in GetValue");
		return E_FAIL;
	}
}

// SetValue just updates the File Meta property store's value cache
HRESULT CPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
	try
	{
		WriteLog("SetValue called for %s %d with value %S", FormatGuid(key.fmtid), key.pid, FormatVariant(propVar));
		HRESULT hr = OpenStore(TRUE);
		return SUCCEEDED(hr) ? _pStore->SetValue(key, propVar) : hr;
		WriteLog("SetValue returning 0x%X", hr);
	}
	catch (...)
	{
		WriteLogMsg("Unknown exception in SetValue");
		return E_FAIL;
	}
}

// Commit writes updates out to thw alternate stream
HRESULT CPropertyHandler::Commit()
{
	try
	{
		WriteLogMsg("Commit called");
		HRESULT hr = OpenStore(TRUE);
		if (SUCCEEDED(hr))
			hr = _pStore->Commit();
		WriteLog("Commit returning 0x%X", hr);
	}
	catch (...)
	{
		WriteLogMsg("Unknown exception in Commit");
		return E_FAIL;
	}
}

HRESULT CPropertyHandler::OpenStore(BOOL bReadWrite)
{
    HRESULT hr = E_UNEXPECTED;

	if (_pStore)
	{
		// If open and read/write, or only read wanted, we're good
		if (_bReadWrite || !bReadWrite)
			return S_OK;
		// Must be open read but read/write wanted - close ready to re-open
		else 
			SafeRelease(&_pStore);
	}

	if (v_pfnStgOpenStorageEx)
	{
		IPropertySetStorage* pPropSetStg = NULL;
		DWORD dwReadWrite = bReadWrite ? STGM_READWRITE : STGM_READ;

		WriteLog("Changing property store access to %s", bReadWrite ? "Read/Write" : "Read");
		hr = (v_pfnStgOpenStorageEx)(_pszFilePath, dwReadWrite | STGM_SHARE_EXCLUSIVE, STGFMT_FILE, 0, NULL, 0, 
				IID_IPropertySetStorage, (void**)&pPropSetStg);

		if (SUCCEEDED(hr))
			// To make IPropertyStore work for Write, it is necessary to use STGM_READWRITE, which the MS documentation says
			// explicitly will not work.  The recommended STGM_READ fails with E_ACCESSDENIED on Write and Commit, which
			// is what you would expect. The only bug appears to be in the documentation.
			hr = PSCreatePropertyStoreFromPropertySetStorage(pPropSetStg, dwReadWrite, IID_IPropertyStore, (void **)&_pStore);

		SafeRelease(&pPropSetStg);

		WriteLog("Property store access change returned 0x%X", hr);
		if (SUCCEEDED(hr))
			_bReadWrite = bReadWrite;
	}

	return hr;
}

DWORD CPropertyHandler::ChainedPropCount()
{
	if (!_bHaveChainedPropCount)
	{
		_cChainedPropCount = 0;

		if (_pChainedPropStore != NULL)
		{
			WriteLogMsg("Calling chained GetCount");
			_bHaveChainedPropCount = SUCCEEDED(_pChainedPropStore->GetCount(&_cChainedPropCount));
			WriteLog("Chains GetCount returned %u", _cChainedPropCount);
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
		WriteLog("Initialize for '%S' mode=s", pszFilePath);// , grfMode == 0 ? "Read" : "Read/Write");

		wcscpy_s(_pszFilePath, MAX_PATH, pszFilePath);

		if (!v_pfnStgOpenStorageEx)
		{
			BOOL fRunningOnNT = ((GetVersion() & 0x80000000) != 0x80000000);
			v_pfnStgOpenStorageEx = ((fRunningOnNT) ? (PFN_STGOPENSTGEX)GetProcAddress(GetModuleHandle(L"OLE32"), "StgOpenStorageEx") : NULL);
		}

		// Check if a chained property handler is configured, and if there is one, load and initialize it too.
		// This is simplified by the fact that we only ever open a chained property handler read-only
		SafeRelease(&_pChainedPropStore);
		PCWSTR pszExt = PathFindExtension(_pszFilePath);
		if (*pszExt)
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
						WriteLog("Attempting to initialize read-only chained property handler: %s", FormatGuid(clsid));
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
						WriteLog("Chained property store Initialize returned 0x%X", hr);
					}
				}

				RegCloseKey(propertyHandlerKey);
			}
		}

		return hr;
	}
	catch (...)
	{
		WriteLogMsg("Unknown exception in Initialize");
		return E_FAIL;
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
