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
class DECLSPEC_UUID("28D14D00-2D80-4956-9657-9D50C8BB47A5") CContextMenuHandler;
#else
class DECLSPEC_UUID("DA38301B-BE91-4397-B2C8-E27A0BD80CC5") CContextMenuHandler;
#endif
HRESULT CContextMenuHandler_CreateInstance(REFIID riid, void **ppv);

void DllAddRef();
void DllRelease();

HRESULT RegisterContextMenuHandler();
HRESULT UnregisterContextMenuHandler();
