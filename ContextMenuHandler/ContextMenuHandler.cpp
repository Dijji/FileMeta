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
#define IDM_IMPORT              1
#define IDM_DELETE              2

static const WCHAR* PropertyHandlerDescription = L"File Metadata Context Menu Handler";
static const WCHAR* ExportVerb			= L"CMHExport";
static const WCHAR* ImportVerb			= L"CMHImport";

class CContextMenuHandler : public IShellExtInit, public IContextMenu
{
public:
    CContextMenuHandler() : _cRef(1), m_pdtobj(NULL)
    {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void ** ppv)
    {
        static const QITAB qit[] =
        {
	        QITABENT(CContextMenuHandler, IContextMenu),
	        QITABENT(CContextMenuHandler, IShellExtInit), 
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
	~CContextMenuHandler()
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
IFACEMETHODIMP CContextMenuHandler::Initialize(
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
//   FUNCTION: CContextMenuHandler::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP CContextMenuHandler::QueryContextMenu(
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

	// Import
    AccessResourceString(IDS_IMPORT, buffer, MAX_PATH);
	wcscat_s(buffer, 2*MAX_PATH, szXmlTarget);
	UINT uMenuFlags = MF_BYPOSITION;

	// Grey menu item if single file does not exist or cannot be opened
    if (m_files.size() == 1)
    {
		wifstream myfile;
		myfile.open (szXmlTarget, ios_base::in );
		if (!myfile.is_open())
			uMenuFlags |= MF_GRAYED;
		myfile.close();
	}

    if (!InsertMenu ( hSubmenu, 1, uMenuFlags, uID++, buffer) )
		return HRESULT_FROM_WIN32(GetLastError());

	// Delete
	AccessResourceString(IDS_DELETE, buffer, MAX_PATH);
	uMenuFlags = MF_BYPOSITION;

	// Grey menu item if single file and no metadata present
    if (m_files.size() == 1)
    {
		if (S_OK != MetadataPresent(m_files[0].c_str()))
			uMenuFlags |= MF_GRAYED;
	}

    if (!InsertMenu ( hSubmenu, 2, uMenuFlags, uID++, buffer) )
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
//   FUNCTION: CContextMenuHandler::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP CContextMenuHandler::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
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
					ExportMetadata(&doc, m_files[i]);

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

		case IDM_IMPORT:
			try
			{
				for (int i = 0; i < m_files.size(); i++)
				{
					// In the multi-file case, we know only that at least one of the files has our context menu,
					// so we need to check the extension of each file to see if our propertyhandler is configured for it
					if (m_files.size() > 1)
						if (1 != m_checker.HasPropertyHandler(m_files[i]))
							continue;

					wstring szXmlTarget = m_files[i];
					szXmlTarget += MetadataFileSuffix;

					// rapidxml parsing works only from a string, so read the whole file
					FILE *pfile;
					errno_t err = _wfopen_s(&pfile, szXmlTarget.c_str(), L"rb");
					if (0 == err)
					{
						fseek (pfile , 0 , SEEK_END);
						size_t lSize = ftell (pfile);  // size in bytes
						rewind (pfile);

						CHAR* buffer = new CHAR[lSize + sizeof(WCHAR)];
						WCHAR* wbuffer;
						size_t lwSize;
						size_t offset = 0;

						fread(buffer, sizeof(CHAR), lSize, pfile);
						fclose(pfile);

						// export files are now UTF-16 with BOM
						if (buffer[0] == (CHAR)0xFF && buffer[1] == (CHAR)0xFE)
						{
							wbuffer = (WCHAR*) buffer;
							lwSize = lSize / sizeof(WCHAR);
							wbuffer[lwSize] = L'\0';  // ensure termination
							offset++;  // skip BOM
						}
						// but also cope with ASCII files from our previous versions, or that have been hand-edited in ASCII
						else
						{
							wbuffer = new WCHAR[lSize + 1];
							lwSize = lSize;
							size_t conv;
							mbstowcs_s(&conv, wbuffer, lSize+1, buffer, lSize);
							delete [] buffer;
						}

						// parse the XML
						xml_document<WCHAR> doc;
						WCHAR * xml = doc.allocate_string(wbuffer+offset, lwSize+1-offset);  

						if (offset == 0) // ASCII file
							delete [] wbuffer;
						else
							delete [] buffer;
	
						try
						{
							doc.parse<0>(xml);
						}
						catch(parse_error& e)
						{
							size_t size = strlen(e.what()) + 1;
							WCHAR * error = new WCHAR[size];
							size_t convertedChars = 0;
							mbstowcs_s(&convertedChars, error, size, e.what(), _TRUNCATE);

	#define MAX_ERRLENGTH 20
							WCHAR content[MAX_ERRLENGTH + 1];
							size = wcslen(e.where<WCHAR>());
							if (size > MAX_ERRLENGTH)
								size = MAX_ERRLENGTH;
							wmemcpy(content, e.where<WCHAR>(), size);
							content[MAX_ERRLENGTH] = L'\0';  // ensure termination

							CPHException cphe = CPHException(ERROR_XML_PARSE_ERROR, E_FAIL, IDS_E_XML_PARSE_ERROR_3, error, content, szXmlTarget);
							delete [] error;
							throw cphe;
						}

						// apply it 
						ImportMetadata(&doc, m_files[i]);
					}

					// Tolerate file access problems in the multi-file case
					else if (m_files.size() == 1)
						throw CPHException(err, E_FAIL, IDS_E_FILEOPEN_1, err);
				}
			}
			catch(CPHException& e)
			{
				WCHAR buffer[MAX_PATH];
				AccessResourceString(IDS_IMPORT_FAILED, buffer, MAX_PATH);
				MessageBox(NULL, e.GetMessage(), buffer, MB_OK);

				hr = e.GetHResult();
			}
			break;

		case IDM_DELETE:
			try
			{
				for (int i = 0; i < m_files.size(); i++)
				{
					// In the multi-file case, we know only that at least one of the files has our context menu,
					// so we need to check the extension of each file to see if our propertyhandler is configured for it
					if (m_files.size() > 1)
					{
						if (1 != m_checker.HasPropertyHandler(m_files[i]))
							continue;

						// Also, we need to check if metadata is present, to avoid adding an empty alternate stream by opening
						// r/w when no metadata stream is present
						if (S_OK != MetadataPresent(m_files[i]))
							continue;
					}

					DeleteMetadata(m_files[i]);
				}
			}
			catch(CPHException& e)
			{
				WCHAR buffer[MAX_PATH];
				AccessResourceString(IDS_DELETE_FAILED, buffer, MAX_PATH);
				MessageBox(NULL, e.GetMessage(), buffer, MB_OK);

				hr = e.GetHResult();
			}
			break;
		}
	}

    return hr;
}


//
//   FUNCTION: CCContextMenuHandler::GetCommandString
//
//   PURPOSE: This method can be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested.
//            We only implement support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP CContextMenuHandler::GetCommandString(UINT_PTR idCommand, 
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
		case IDM_IMPORT:
        hr = StringCchCopyW(reinterpret_cast<PWSTR>(pszName), cchMax, ImportVerb);
        break;
		}
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}


#pragma region Registration and other COM mechanics

HRESULT CContextMenuHandler_CreateInstance(REFIID riid, void **ppv)
{
    HRESULT hr = E_OUTOFMEMORY;
    CContextMenuHandler *pirm = new (std::nothrow) CContextMenuHandler();
    if (pirm)
    {
        hr = pirm->QueryInterface(riid, ppv);
        pirm->Release();
    }
    return hr;
}

HRESULT RegisterContextMenuHandler()
{
    // register the context menu handler COM object
    CRegisterExtension re(__uuidof(CContextMenuHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.RegisterInProcServer(PropertyHandlerDescription, L"Both");

	return hr;
}

HRESULT UnregisterContextMenuHandler()
{
    // Unregister the context menu handler COM object.
    CRegisterExtension re(__uuidof(CContextMenuHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();

    return hr;
}
#pragma endregion