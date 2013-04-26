// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved
// Copied from IdealPropertyHandler sample http://archive.msdn.microsoft.com/shellintegration and modified

#include "RegisterExtension.h"
#include <shlobj.h>
#include <shlwapi.h>
#include <strsafe.h>
#include <shobjidl.h>

#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "shlwapi.lib")     // link to this

__inline HRESULT ResultFromKnownLastError() { const DWORD err = GetLastError(); return err == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(err); }

// retrieve the HINSTANCE for the current DLL or EXE using this symbol that
// the linker provides for every module, avoids the need for a global HINSTANCE variable
// and provides access to this value for static libraries
EXTERN_C IMAGE_DOS_HEADER __ImageBase;
__inline HINSTANCE GetModuleHINSTANCE() { return (HINSTANCE)&__ImageBase; }

CRegisterExtension::CRegisterExtension(REFCLSID clsid /* = CLSID_NULL */, HKEY hkeyRoot /* = HKEY_CURRENT_USER */) : _hkeyRoot(hkeyRoot), _fAssocChanged(false)
{
    SetHandlerCLSID(clsid);
    GetModuleFileName(GetModuleHINSTANCE(), _szModule, ARRAYSIZE(_szModule));
}

CRegisterExtension::~CRegisterExtension()
{
    if (_fAssocChanged)
    {
        // inform Explorer, et al that file association data has changed
        SHChangeNotify(SHCNE_ASSOCCHANGED, 0, 0, 0);
    }
}

void CRegisterExtension::SetHandlerCLSID(REFCLSID clsid)
{
    _clsid = clsid;
    StringFromGUID2(_clsid, _szCLSID, ARRAYSIZE(_szCLSID));
}

HRESULT CRegisterExtension::_EnsureModule() const
{
    return _szModule[0] ? S_OK : E_FAIL;
}

HRESULT CRegisterExtension::RegisterInProcServer(PCWSTR pszFriendlyName, PCWSTR pszThreadingModel) const
{
    HRESULT hr = _EnsureModule();
    if (SUCCEEDED(hr))
    {
        hr = RegSetKeyValuePrintf(_hkeyRoot, L"Software\\Classes\\CLSID\\%s", L"", pszFriendlyName, _szCLSID);
        if (SUCCEEDED(hr))
        {
            hr = RegSetKeyValuePrintf(_hkeyRoot, L"Software\\Classes\\CLSID\\%s\\InProcServer32", L"", _szModule, _szCLSID);
            if (SUCCEEDED(hr))
            {
                hr = RegSetKeyValuePrintf(_hkeyRoot, L"Software\\Classes\\CLSID\\%s\\InProcServer32", L"ThreadingModel", pszThreadingModel, _szCLSID);
            }
        }
    }
    return hr;
}

// use for
// ManualSafeSave = REG_DWORD:<1>
// EnableShareDenyNone = REG_DWORD:<1>
// EnableShareDenyWrite = REG_DWORD:<1>

HRESULT CRegisterExtension::RegisterInProcServerAttribute(PCWSTR pszAttribute, DWORD dwValue) const
{
    return RegSetKeyValuePrintf(_hkeyRoot, L"Software\\Classes\\CLSID\\%s", pszAttribute, dwValue, _szCLSID);
}

HRESULT CRegisterExtension::UnRegisterObject() const
{
    // might have an AppID value, try that
    HRESULT hr = RegDeleteKeyPrintf(_hkeyRoot, L"Software\\Classes\\AppID\\%s", _szCLSID);
    if (SUCCEEDED(hr))
    {
        hr = RegDeleteKeyPrintf(_hkeyRoot, L"Software\\Classes\\CLSID\\%s", _szCLSID);
    }
    return hr;
}


// in process context menu handler

HRESULT CRegisterExtension::RegisterContextMenuHandler(PCWSTR pszProgID, PCWSTR pszDescription) const
{
    return RegSetKeyValuePrintf(_hkeyRoot, L"Software\\Classes\\%s\\shellex\\ContextMenuHandlers\\%s",
        L"", pszDescription, pszProgID, _szCLSID);
}

HRESULT CRegisterExtension::RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, PCWSTR pszValue, ...) const
{
    va_list argList;
    va_start(argList, pszValue);

    WCHAR szKeyName[512];
    HRESULT hr = StringCchVPrintf(szKeyName, ARRAYSIZE(szKeyName), pszKeyFormatString, argList);
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(RegSetKeyValueW(hkey, szKeyName, pszValueName, REG_SZ, pszValue,
            lstrlen(pszValue) * sizeof(*pszValue)));
    }

    va_end(argList);

    _UpdateAssocChanged(hr, pszKeyFormatString);
    return hr;
}

HRESULT CRegisterExtension::RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, DWORD dwValue, ...) const
{
    va_list argList;
    va_start(argList, dwValue);

    WCHAR szKeyName[512];
    HRESULT hr = StringCchVPrintf(szKeyName, ARRAYSIZE(szKeyName), pszKeyFormatString, argList);
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(RegSetKeyValueW(hkey, szKeyName, pszValueName, REG_DWORD, &dwValue, sizeof(dwValue)));
    }

    va_end(argList);

    _UpdateAssocChanged(hr, pszKeyFormatString);
    return hr;
}

__inline HRESULT MapNotFoundToSuccess(HRESULT hr)
{
    return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) == hr ? S_OK : hr;
}

HRESULT CRegisterExtension::RegDeleteKeyPrintf(HKEY hkey, PCWSTR pszKeyFormatString, ...) const
{
    va_list argList;
    va_start(argList, pszKeyFormatString);

    WCHAR szKeyName[512];
    HRESULT hr = StringCchVPrintf(szKeyName, ARRAYSIZE(szKeyName), pszKeyFormatString, argList);
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(RegDeleteTree(hkey, szKeyName));
    }

    va_end(argList);

    _UpdateAssocChanged(hr, pszKeyFormatString);
    return MapNotFoundToSuccess(hr);
}


void CRegisterExtension::_UpdateAssocChanged(HRESULT hr, PCWSTR pszKeyFormatString) const
{
    static const WCHAR c_szProgIDPrefix[] = L"Software\\Classes\\%s";
    if (SUCCEEDED(hr) && !_fAssocChanged &&
        (StrCmpNIC(pszKeyFormatString, c_szProgIDPrefix, ARRAYSIZE(c_szProgIDPrefix) - 1) == 0 ||
         StrStrI(pszKeyFormatString, L"PropertyHandlers") ||
         StrStrI(pszKeyFormatString, L"KindMap")))
    {
        const_cast<CRegisterExtension*>(this)->_fAssocChanged = true;
    }
}
