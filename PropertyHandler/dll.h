// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved
// Copied from IdealPropertyHandler sample http://archive.msdn.microsoft.com/shellintegration and modified

#pragma once
#include <windows.h>
#include <new> // std::nothrow
#include <propsys.h>

template <class T> void SafeRelease(T **ppT)
{
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = NULL;
    }
}

#ifdef _WIN64
class DECLSPEC_UUID("D06391EE-2FEB-419B-9667-AD160D0849F3") CPropertyHandler;
#else
class DECLSPEC_UUID("60211757-EF87-465e-B6C1-B37CF98295F9") CPropertyHandler;
#endif
HRESULT CPropertyHandler_CreateInstance(REFIID riid, void **ppv);

typedef HRESULT (CALLBACK* PFN_STGOPENSTGEX)(const WCHAR*, DWORD, DWORD, DWORD, void*, void*, REFIID riid, void **);

void DllAddRef();
void DllRelease();

HRESULT RegisterPropertyHandler();
HRESULT UnregisterPropertyHandler();
