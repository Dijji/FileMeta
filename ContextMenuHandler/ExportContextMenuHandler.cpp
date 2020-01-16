// Some parts copied from CppShellExtContextMenuHandler Copyright (c) Microsoft Corporation and subject to the Microsoft Public License: http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other code Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

#include <windows.h>
#include <WinUser.h>
#include <shlwapi.h>
#include <propkey.h>
#include <Propvarutil.h>
#include <shlobj.h>     // For IShellExtInit and IContextMenu
#include <atlbase.h>	// For CComPtr
#include <strsafe.h>
#include "..\CommandLine\XmlHelpers.h"
#include "resource.h"
#include "dll.h"
#include "RegisterExtension.h"
#include <iostream>
#include <fstream>
#include <sstream>
#include <algorithm>
using namespace std;
using namespace rapidxml;

#define IDM_EXPORT              0  // The commands' identifier offsets

static const WCHAR* PropertyHandlerDescription = L"File Metadata Export Context Menu Handler";
static const WCHAR* ExportVerb			= L"CMHExport";

class CExportContextMenuHandler : public IShellExtInit, public IContextMenu
{
public:
    CExportContextMenuHandler() : _cRef(1), m_pdtobj(NULL)
    {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void ** ppv)
    {
        static const QITAB qit[] =
        {
	        QITABENT(CExportContextMenuHandler, IContextMenu),
	        QITABENT(CExportContextMenuHandler, IShellExtInit), 
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

    // IShellExtInit
    IFACEMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID);

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCommand, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);


private:
	~CExportContextMenuHandler()
    {
		SafeRelease(&m_pdtobj);
        DllRelease();
    }

    long _cRef;
	CExtensionChecker m_checker;

	IDataObject *m_pdtobj;
	std::vector<std::wstring> m_files;
};


#pragma region IShellExtInit

// Initialize the context menu handler.
// If any value other than S_OK is returned from the method, the context 
// menu item is not displayed.
IFACEMETHODIMP CExportContextMenuHandler::Initialize(
    LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
	// This is important because Initialize can be called multiple times,
	// e.g. when more than 16 files are selected
	SafeRelease(&m_pdtobj);

    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }
	else
	{ 
        m_pdtobj = pDataObj; 
        m_pdtobj->AddRef(); 
    }

	return S_OK;
}

#pragma endregion

#pragma region IContextMenu

//
//   FUNCTION: CExportContextMenuHandler::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP CExportContextMenuHandler::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

	HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    // The pDataObj pointer contains the objects being acted upon
    if (NULL != m_pdtobj && SUCCEEDED(m_pdtobj->GetData(&fe, &stm)))
    {
        // Get an HDROP handle.
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {
            // Determine how many files are involved in this operation. 
            UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
	 		WCHAR buff[MAX_PATH];

			for (int i=0; i < nFiles; i++)
			{
		      DragQueryFile(hDrop, i, buff, sizeof(buff));
			  m_files.push_back(buff);
			}

            GlobalUnlock(stm.hGlobal);
        }

        ReleaseStgMedium(&stm);
    }

	WCHAR szXmlTarget[MAX_PATH];	
    hr = S_OK;

    if (m_files.size() == 1)
    {
		wcscpy_s(szXmlTarget, MAX_PATH, m_files[0].c_str());
		wcscat_s(szXmlTarget, MAX_PATH, MetadataFileSuffix);
    }
	else
	{
		AccessResourceString(IDS_XML_FILE, szXmlTarget, MAX_PATH);
	}
	
    // First, create and populate a submenu.
    HMENU hSubmenu = CreatePopupMenu();
    UINT uID = idCmdFirst;
	WCHAR buffer[2*MAX_PATH];

	// Export
    AccessResourceString(IDS_EXPORT, buffer, MAX_PATH);
	wcscat_s(buffer, 2*MAX_PATH, szXmlTarget);
	if (!InsertMenu ( hSubmenu, 0, MF_BYPOSITION, uID++, buffer) )
		return HRESULT_FROM_WIN32(GetLastError());

    // Insert the submenu into the ctx menu provided by Explorer.
	MENUITEMINFO mii = { sizeof(MENUITEMINFO) };

    mii.fMask = MIIM_SUBMENU | MIIM_STRING | MIIM_ID;
    mii.wID = uID++;
    mii.hSubMenu = hSubmenu;

    AccessResourceString(IDS_METADATA, buffer, MAX_PATH);
	mii.dwTypeData = (LPWSTR)buffer;

    if (!InsertMenuItem ( hMenu, indexMenu, TRUE, &mii ) )
		return HRESULT_FROM_WIN32(GetLastError());

    return MAKE_HRESULT ( SEVERITY_SUCCESS, FACILITY_NULL, uID - idCmdFirst );
}

//
//   FUNCTION: CExportContextMenuHandler::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP CExportContextMenuHandler::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	HRESULT hr = E_FAIL;

	// Commands can be indicated by offset id or command verb string
	// We support only the id form
	if (IS_INTRESOURCE(pici->lpVerb)) 
	{
		switch(LOWORD(pici->lpVerb))
		{
		case IDM_EXPORT:
			try
			{
				for (int i = 0; i < m_files.size(); i++)
				{
					// In the multi-file case, we know only that at least one of the files has our context menu,
					// so we need to check the extension of each file to see if a propertyhandler is configured for it
					if (m_files.size() > 1)
						if (0 == m_checker.HasPropertyHandler(m_files[i]))
							continue;

					// Build an XML document containing thw metadata
					xml_document<WCHAR> doc;
					ExportMetadata(&doc, m_files[i], true);

					// writing to a string rather than directly to the stream is odd, but writing directly does not compile
					// (trying to access a private constructor on traits - a typically arcane template issue)
					wstring s;
					print(std::back_inserter(s), doc, 0);

					wstring szXmlTarget = m_files[i];
					szXmlTarget += MetadataFileSuffix;

					// Now write from the XML string to a file stream
					// This used to be STL, but wofstream by default writes 8-bit encoded files, and changing that is complex
					FILE *pfile;
					errno_t err = _wfopen_s(&pfile, szXmlTarget.c_str(), L"w+, ccs=UTF-16LE");
					if (0 == err)
					{
						fwrite(s.c_str(), sizeof(WCHAR), s.length(), pfile);
						fclose(pfile);
					}
					else
						throw CPHException(err, E_FAIL, IDS_E_FILEOPEN_1, err);
				}

				hr = S_OK;
			}
			catch(CPHException& e)
			{
				WCHAR buffer[MAX_PATH];
				AccessResourceString(IDS_EXPORT_FAILED, buffer, MAX_PATH);
				MessageBox(NULL, e.GetMessage(), buffer, MB_OK);
				hr = e.GetHResult();
			}
			break;
		}
	}

    return hr;
}


//
//   FUNCTION: CCExportContextMenuHandler::GetCommandString
//
//   PURPOSE: This method can be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested.
//            We only implement support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP CExportContextMenuHandler::GetCommandString(UINT_PTR idCommand, 
    UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;

    switch (uFlags)
    {
    // GCS_VERBW is an optional feature that enables a caller to 
    // discover the canonical name for the verb passed in through 
    // idCommand.
    case GCS_VERBW:
		switch (idCommand)
		{
		case IDM_EXPORT:
        hr = StringCchCopyW(reinterpret_cast<PWSTR>(pszName), cchMax, ExportVerb);
        break;
		}
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}


#pragma region Registration and other COM mechanics

HRESULT CExportContextMenuHandler_CreateInstance(REFIID riid, void **ppv)
{
    HRESULT hr = E_OUTOFMEMORY;
    CExportContextMenuHandler *pirm = new (std::nothrow) CExportContextMenuHandler();
    if (pirm)
    {
        hr = pirm->QueryInterface(riid, ppv);
        pirm->Release();
    }
    return hr;
}

HRESULT RegisterExportContextMenuHandler()
{
    // register the context menu handler COM object
    CRegisterExtension re(__uuidof(CExportContextMenuHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.RegisterInProcServer(PropertyHandlerDescription, L"Both");

	return hr;
}

HRESULT UnregisterExportContextMenuHandler()
{
    // Unregister the context menu handler COM object.
    CRegisterExtension re(__uuidof(CExportContextMenuHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();

    return hr;
}
#pragma endregion