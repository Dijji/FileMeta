// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved
// Copied from IdealPropertyHandler sample http://archive.msdn.microsoft.com/shellintegration and modified

#pragma once

#include <windows.h>

// production code should use an installer technology like MSI to register its handlers
// rather than this class.
// this class is used for demontration purpouses, it encapsulate the different types
// of handler registrations, schematizes those by provding methods that have parameters
// that map to the supported extension schema and makes it easy to create self registering
// .exe and .dlls

class CRegisterExtension
{
public:
    CRegisterExtension(REFCLSID clsid = CLSID_NULL, HKEY hkeyRoot = HKEY_CURRENT_USER);
    ~CRegisterExtension();
    void SetHandlerCLSID(REFCLSID clsid);
   
    HRESULT RegisterInProcServer(PCWSTR pszFriendlyName, PCWSTR pszThreadingModel) const;
    HRESULT RegisterInProcServerAttribute(PCWSTR pszAttribute, DWORD dwValue) const;

    // remove a COM object registration
    HRESULT UnRegisterObject() const;

    HRESULT RegisterContextMenuHandler(PCWSTR pszProgID, PCWSTR pszDescription) const;

    // this should probably be private but they are useful
    HRESULT RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, PCWSTR pszValue, ...) const;
    HRESULT RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, DWORD dwValue, ...) const;
    
    HRESULT RegDeleteKeyPrintf(HKEY hkey, PCWSTR pszKeyFormatString, ...) const;
    
    PCWSTR GetCLSIDString() const { return _szCLSID; };

    bool HasClassID() const { return _clsid != CLSID_NULL ? true : false; };

private:

    HRESULT _EnsureModule() const;
    bool _IsBaseClassProgID(PCWSTR pszProgID)  const;
    HRESULT _EnsureBaseProgIDVerbIsNone(PCWSTR pszProgID) const;
    void _UpdateAssocChanged(HRESULT hr, PCWSTR pszKeyFormatString) const;

    CLSID _clsid;
    HKEY _hkeyRoot;
    WCHAR _szCLSID[39];
    WCHAR _szModule[MAX_PATH];
    bool _fAssocChanged;
};
