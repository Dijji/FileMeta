// Copyright (c) 2013, Dijii, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

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
    CPropertyHandler() : _cRef(1), _pStore(NULL), _fReadOnly(FALSE)
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
        DllRelease();
    }

    long _cRef;

	BOOL		            _fReadOnly;     // Whether storage set currently open read only
	IPropertyStore *		_pStore;		// Wrapper over set of storages
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
    return _pStore ? _pStore->GetCount(pcProps) : E_UNEXPECTED;
}

HRESULT CPropertyHandler::GetAt(DWORD iProp, PROPERTYKEY *pkey)
{
    *pkey = PKEY_Null;
    return _pStore ? _pStore->GetAt(iProp, pkey) : E_UNEXPECTED;
}

HRESULT CPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar)
{
    PropVariantInit(pPropVar);
    return _pStore ? _pStore->GetValue(key, pPropVar) : E_UNEXPECTED;
}

// SetValue just updates the property store's value cache
HRESULT CPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
    return _pStore ? _pStore->SetValue(key, propVar) : E_UNEXPECTED;
}

// Commit writes updates out to thw alternate stream
HRESULT CPropertyHandler::Commit()
{
    return _pStore ? _pStore->Commit() : E_UNEXPECTED;
}

HRESULT CPropertyHandler::Initialize(LPCWSTR pszFilePath, DWORD grfMode)
{
    HRESULT hr = E_UNEXPECTED;

	if (!v_pfnStgOpenStorageEx)
	{
		BOOL fRunningOnNT = ((GetVersion() & 0x80000000) != 0x80000000);
		v_pfnStgOpenStorageEx = ((fRunningOnNT) ? (PFN_STGOPENSTGEX)GetProcAddress(GetModuleHandle(L"OLE32"), "StgOpenStorageEx") : NULL);
	}

	if (v_pfnStgOpenStorageEx)
	{
		IPropertySetStorage* pPropSetStg = NULL;

		// On Win2K+ we can get the NTFS version of OLE properties (saved in alt stream)...
		hr = (v_pfnStgOpenStorageEx)(pszFilePath, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, STGFMT_FILE, 0, NULL, 0, 
				IID_IPropertySetStorage, (void**)&pPropSetStg);

		// If we failed to gain write access, try for just read access
		if ((hr == STG_E_ACCESSDENIED) && (!_fReadOnly))
		{
			_fReadOnly = TRUE;  // we don't do anything with this, but we could use it when writing, rather than just allowing failures
			hr = (v_pfnStgOpenStorageEx)(pszFilePath, (STGM_READ | STGM_SHARE_EXCLUSIVE), STGFMT_FILE,
				0, NULL, 0, IID_IPropertySetStorage, (void**)&pPropSetStg);
		}

		if (SUCCEEDED(hr))
			// To make IPropertyStore work for Write, it is necessary to use STGM_READWRITE, which the MS documentation says
			// explicitly will not work.  The recommended STGM_READ fails with E_ACCESSDENIED on Write and Commit, which
			// is what you would expect. The only bug appears to be in the documentation.
			hr = PSCreatePropertyStoreFromPropertySetStorage(pPropSetStg, STGM_READWRITE, IID_IPropertyStore, (void **)&_pStore);

		SafeRelease(&pPropSetStg);
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

    return hr;
}

HRESULT UnregisterPropertyHandler()
{
    // Unregister the property handler COM object.
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();

    return hr;
}
