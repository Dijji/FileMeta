// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

#include <windows.h>
#include <shlwapi.h>
#include <propkey.h>
#include "dll.h"
#include "RegisterExtension.h"

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

HRESULT CPropertyHandler::GetCount(DWORD *pcProps)
{
    *pcProps = 0;
	HRESULT hr = OpenStore(FALSE);
    hr = SUCCEEDED(hr) ? _pStore->GetCount(pcProps) : hr;
    if (SUCCEEDED(hr))
         *pcProps += ChainedPropCount();
    return hr;
}

HRESULT CPropertyHandler::GetAt(DWORD iProp, PROPERTYKEY *pkey)
{
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
    return hr;
}

HRESULT CPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar)
{
    PropVariantInit(pPropVar);
	HRESULT hr = OpenStore(FALSE);
	// Take the File Meta property value first, and if there isn't one,
	// try for a chained property value
    hr = SUCCEEDED(hr) ? _pStore->GetValue(key, pPropVar) : hr;
    if (_pChainedPropStore != NULL && SUCCEEDED(hr) && pPropVar->vt == VT_EMPTY)
        hr = _pChainedPropStore->GetValue(key, pPropVar);
    return hr;
}

// SetValue just updates the File Meta property store's value cache
HRESULT CPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
	HRESULT hr = OpenStore(TRUE);
    return SUCCEEDED(hr) ? _pStore->SetValue(key, propVar) : hr;
}

// Commit writes updates out to thw alternate stream
HRESULT CPropertyHandler::Commit()
{
	HRESULT hr = OpenStore(TRUE);
    return SUCCEEDED(hr) ? _pStore->Commit() : hr;
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

		hr = (v_pfnStgOpenStorageEx)(_pszFilePath, dwReadWrite | STGM_SHARE_EXCLUSIVE, STGFMT_FILE, 0, NULL, 0, 
				IID_IPropertySetStorage, (void**)&pPropSetStg);

		if (SUCCEEDED(hr))
			// To make IPropertyStore work for Write, it is necessary to use STGM_READWRITE, which the MS documentation says
			// explicitly will not work.  The recommended STGM_READ fails with E_ACCESSDENIED on Write and Commit, which
			// is what you would expect. The only bug appears to be in the documentation.
			hr = PSCreatePropertyStoreFromPropertySetStorage(pPropSetStg, dwReadWrite, IID_IPropertyStore, (void **)&_pStore);

		SafeRelease(&pPropSetStg);

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
			_bHaveChainedPropCount = SUCCEEDED(_pChainedPropStore->GetCount(&_cChainedPropCount));
		else
			_bHaveChainedPropCount = TRUE;
	}
	return _cChainedPropCount;
}

HRESULT CPropertyHandler::Initialize(LPCWSTR pszFilePath, DWORD grfMode)
{
    HRESULT hr = S_OK;

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
                }
            }

            RegCloseKey(propertyHandlerKey);
        }
 	}

    return hr;
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
