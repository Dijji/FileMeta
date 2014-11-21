// Some parts copied from Microsoft's EnumAll sample: http://msdn.microsoft.com/en-us/library/aa379016(v=vs.85).aspx (no visible license terms_
// Some parts copied from CppShellExtContextMenuHandler Copyright (c) Microsoft Corporation and subject to the Microsoft Public License: http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other code Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

#include "stdafx.h"
#include <string>
#include <iostream>
#include <algorithm>
#include <map>
#include <shlwapi.h>
#include <propsys.h>
#include <propkey.h>
#include <Propvarutil.h>
#include "tclap/CmdLine.h"
#undef RAPIDXML_NO_EXCEPTIONS
#include "rapidxml.hpp"
#include "rapidxml_print.hpp"
#include "resource.h"
#include <strsafe.h>
#include <direct.h>

using namespace rapidxml;
using namespace TCLAP;
using namespace std;

static const WCHAR* MetadataFileSuffix	= L".metadata.xml";
static const WCHAR* MetadataNodeName	= L"Metadata";
static const WCHAR* StorageNodeName		= L"Storage";
static const WCHAR* DescriptionAttrName = L"Description";
static const WCHAR* FormatIDAttrName	= L"FormatID";
static const WCHAR* PropertyNodeName	= L"Property";
static const WCHAR* PropertyIdAttrName	= L"Id";
static const WCHAR* NameAttrName		= L"Name";
static const WCHAR* TypeAttrName		= L"Type";
static const WCHAR* TypeIdAttrName	    = L"TypeId";
static const WCHAR* ValueNodeName		= L"Value";

static const WCHAR* OurPropertyHandlerGuid64 = L"{D06391EE-2FEB-419B-9667-AD160D0849F3}";
static const WCHAR* OurPropertyHandlerGuid32 = L"{60211757-EF87-465e-B6C1-B37CF98295F9}";

int AccessResourceString(UINT uId, LPWSTR lpBuffer, int nBufferMax)
{
	return LoadStringW(GetModuleHandle(NULL), uId, lpBuffer, nBufferMax);
}

class CPHException 
{
public:
	CPHException (int err, UINT uResourceId, ...)
	{
		WCHAR    lpszFormat[MAX_PATH];
		va_list  fmtList;

		_err = err;
		AccessResourceString(uResourceId, lpszFormat, MAX_PATH);
		va_start( fmtList, uResourceId );
		vswprintf_s( _pszError, MAX_PATH, lpszFormat, fmtList );
		va_end( fmtList );
	}

	int		_err;
	WCHAR   _pszError[MAX_PATH];
};

class CExtensionChecker
{
private:
	std::map<std::wstring, BOOL> m_extensions;		// map of extensions checked so far

public:
	// See if file is handled by our property handler
	BOOL HasOurPropertyHandler(wstring fileName)
	{
		WCHAR pszExt[_MAX_EXT];
		BOOL val = FALSE;

		// Try and get the extension
		if (0 == _wsplitpath_s(fileName.c_str(), NULL, 0, NULL, 0, NULL, 0, pszExt, _MAX_EXT))
		{
			CharLower(pszExt); // use lower case for the map

			// If we've already checked this extension, return the result
			if (m_extensions.find(pszExt) != m_extensions.end())
				return m_extensions[pszExt];

			wstring subKey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\";
			subKey.append(pszExt);

			WCHAR buffer[_MAX_EXT];
			DWORD size = _MAX_EXT;
			// Not finding the key results in FALSE
			if (ERROR_SUCCESS == RegGetValue(HKEY_LOCAL_MACHINE, subKey.c_str(), NULL, RRF_RT_REG_SZ, NULL, buffer, &size))
			{
				// If the key is found, TRUE iff our property handler
#ifdef _WIN64
				val = (0 == StrCmpI(buffer, OurPropertyHandlerGuid64));
#else
				val = (0 == StrCmpI(buffer, OurPropertyHandlerGuid32));
#endif	
			}

			m_extensions[pszExt] = val;
		}

		return val;
	}
};


void ExportMetadata (xml_document<WCHAR> *doc, wstring targetFile);
void ImportMetadata (xml_document<WCHAR> *doc, wstring targetFile);
void DeleteMetadata (wstring targetFile);

int wmain(int argc, WCHAR* argv[])
{
	int result = 0;
	bool prompt = false;
	CoInitialize(NULL);

	try
	{  
		CExtensionChecker checker;

		// Define the command line object.
		CmdLine cmd(L"Export, import or delete File Meta metadata properties", L'=', L"0.1");

		// Define function switches
		SwitchArg deleteSwitch(L"d",L"delete",L"Remove all metadata from file", false);
		SwitchArg importSwitch(L"i",L"import",L"Import metadata from XML", false);
		SwitchArg exportSwitch(L"e",L"export",L"Export metadata to XML", true);
		vector<Arg*> functions;
		functions.push_back(&deleteSwitch);
		functions.push_back(&importSwitch);
		functions.push_back(&exportSwitch);
		cmd.xorAdd(functions);

		// Define prompt switch
		SwitchArg promptSwitch(L"p",L"prompt",L"After execution, prompt to continue", false);
		cmd.add(promptSwitch);

		// Define XML file name override
		ValueArg<wstring> xmlFileArg(L"x",L"xml",L"Name of XML file (only valid if one target file)",false,L"",L"file name");
		cmd.add( xmlFileArg );

		// Define XML file directory override
		ValueArg<wstring> xmlDirArg(L"f",L"folder",L"Directory for XML files (default is same directory as target file)",false,L"",L"directory name");
		cmd.add( xmlDirArg );

		// Define XML console override
		SwitchArg xmlConsoleSwitch(L"c",L"console",L"Output XML to console instead of file (only valid for --export)",false);
		cmd.add( xmlConsoleSwitch );

		// Define target file
		UnlabeledMultiArg<wstring> fileArg(L"file",L"Names of target files", true,L"file name",false);
		cmd.add( fileArg );

		// Parse the args.
		cmd.parse( argc, argv );

		prompt = promptSwitch.getValue();
		vector<wstring> targetFiles(fileArg.getValue());

		if (xmlFileArg.isSet())
		{
			if (targetFiles.size() > 1)
				throw ArgException(L"-x cannot be used with multiple files", L"xml");
			else if (xmlDirArg.isSet())
				throw ArgException(L"-x and -f cannot be used together", L"xml");
			else if (xmlConsoleSwitch.isSet())
				throw ArgException(L"-x and -c cannot be used together", L"xml");
		}
		else if (xmlDirArg.isSet())
		{
			if (!PathIsDirectory(xmlDirArg.getValue().c_str()))
				throw ArgException(L"Cannot find directory", L"f");
			else if (xmlConsoleSwitch.isSet())
				throw ArgException(L"-f and -c cannot be used together", L"console");
		}
		else if (xmlConsoleSwitch.isSet())
		{
			if (!exportSwitch.isSet())
				throw ArgException(L"-c can only be used with -e", L"console");
		}

		for (auto pos = targetFiles.begin(); pos != targetFiles.end(); ++pos)
		{
			wstring targetFile(*pos);

			if (!PathFileExists(targetFile.c_str()))
			{
				wcerr << L"Cannot find file \"" << targetFile.c_str() << L"\"" << endl;
				result = ERROR_FILE_NOT_FOUND;
				break;
			}
			else if (!checker.HasOurPropertyHandler(targetFile))
			{
				// Skip files that do not have our property handler
				continue;
			}
			else if (deleteSwitch.isSet())
			{
				DeleteMetadata(targetFile);

				wcout << L"Removed all metadata from " << targetFile <<  endl;
			}
			else
			{
				wstring xmlFile;

				// Build XML file name 
				if (xmlFileArg.isSet())
				{
					// use specified file
					xmlFile = xmlFileArg.getValue();
				}
				else if (xmlDirArg.isSet())
				{
					// build from specified directory and target file stem
					WCHAR buf[MAX_PATH];
					wcscpy_s(buf, MAX_PATH, xmlDirArg.getValue().c_str());
					PathAppend(buf, PathFindFileName(targetFile.c_str()));
					xmlFile = buf;
					xmlFile += MetadataFileSuffix;
				}
				else
				{
					// build from full target file name
					xmlFile = targetFile + MetadataFileSuffix;
				}


				if (exportSwitch.isSet())
				{
					xml_document<WCHAR> doc;
					ExportMetadata(&doc, targetFile);

					// writing to a string rather than directly to the stream is odd, but writing directly does not compile
					// (trying to access a private constructor on traits - a typically arcane template issue)
					wstring s;
					print(std::back_inserter(s), doc, 0);

					if (xmlConsoleSwitch.isSet())
					{
						wcout << s << endl;
					}
					else
					{
						// Now write from the XML string to a file stream
						// This used to be STL, but wofstream by default writes 8-bit encoded files, and changing that is complex
						FILE *pfile;
						errno_t err = _wfopen_s(&pfile, xmlFile.c_str(), L"w+, ccs=UTF-16LE");
						if (0 == err)
						{
							fwrite(s.c_str(), sizeof(WCHAR), s.length(), pfile);
							fclose(pfile);
						}
						else
							throw CPHException(err, IDS_E_FILEOPEN_1, err);

						wcout << L"Exported metadata to " << xmlFile << endl;
					}
				}
				// Must be import, so XML file needs to exist
				else if (!PathFileExists(xmlFile.c_str()))
				{
					wcerr << L"Cannot find XML file \"" << xmlFile.c_str() << L"\"" << endl;
					result = ERROR_FILE_NOT_FOUND;
					break;
				}
				else
				{
					// rapidxml parsing works only from a string, so read the whole file
					FILE *pfile;
					errno_t err = _wfopen_s(&pfile, xmlFile.c_str(), L"rb");
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

							CPHException cphe = CPHException(ERROR_XML_PARSE_ERROR, IDS_E_XML_PARSE_ERROR_2, error, content);
							delete [] error;
							throw cphe;
						}

						// apply it 
						ImportMetadata(&doc, targetFile);
					}
					else
						throw CPHException(err, IDS_E_FILEOPEN_1, err);

					wcout << L"Imported metadata to " << targetFile << L" from " << xmlFile <<  endl;
				}
			}
		}
	}
	catch (ArgException &e)  // catch any exceptions
	{
		wcerr << L"error: " << e.error() << L" for arg " << e.argId() << endl; 
		result = ERROR_INVALID_PARAMETER;
	}
	catch (CPHException &e)
	{
		wcerr << e._pszError << endl;
		result = e._err;
	}

	if (prompt)
	{
		wcout << L"Hit any key to continue...";
		_getwch();
	}

	CoUninitialize();

	return result;
}

//-------------------------------------------
// Everything below is copied from ContextMenuHandler.cpp, with the following changes:
// - remove class from method names
// - replace 2x m_szSelectedFile references with references to new targetFile param
// - add forward references for import and export helpers
//
// todo make source common and remove copied material
//-------------------------------------------


#pragma region Converters and string helpers

std::vector<std::wstring> &wsplit(const std::wstring &s, WCHAR delim, std::vector<std::wstring> &elems) {
    std::wstringstream ss(s);
    std::wstring item;
    while (std::getline(ss, item, delim)) {
        elems.push_back(item);
    }
    return elems;
}


std::vector<std::wstring> wsplit(const std::wstring &s, WCHAR delim) {
    std::vector<std::wstring> elems;
    wsplit(s, delim, elems);
    return elems;
}

//+-------------------------------------------------------------------
//
//  ConvertVarTypeToString
//  
//  Generate a string for a given PROPVARIANT variable type (VT). 
//  For the given vt, write the string to pwszType, which is a buffer
//  of size cchType characters.
//
//+-------------------------------------------------------------------

void
ConvertVarTypeToString( VARTYPE vt, WCHAR *pwszType, size_t cchType )
{
    const WCHAR *pwszModifier;

    // Ensure that the output string is terminated
    // (wcsncpy does not guarantee termination)

    pwszType[ cchType-1 ] = L'\0';
    --cchType;

    // Create a string using the basic type.

    switch( vt & VT_TYPEMASK )
    {
    case VT_EMPTY:
        wcsncpy_s( pwszType, cchType, L"VT_EMPTY", cchType );
        break;
    case VT_NULL:
        wcsncpy_s( pwszType, cchType, L"VT_NULL", cchType );
        break;
    case VT_I2:
        wcsncpy_s( pwszType, cchType, L"VT_I2", cchType );
        break;
    case VT_I4:
        wcsncpy_s( pwszType, cchType, L"VT_I4", cchType );
        break;
    case VT_I8:
        wcsncpy_s( pwszType, cchType, L"VT_I8", cchType );
        break;
    case VT_UI2:
        wcsncpy_s( pwszType, cchType, L"VT_UI2", cchType );
        break;
    case VT_UI4:
        wcsncpy_s( pwszType, cchType, L"VT_UI4", cchType );
        break;
    case VT_UI8:
        wcsncpy_s( pwszType, cchType, L"VT_UI8", cchType );
        break;
    case VT_R4:
        wcsncpy_s( pwszType, cchType, L"VT_R4", cchType );
        break;
    case VT_R8:
        wcsncpy_s( pwszType, cchType, L"VT_R8", cchType );
        break;
    case VT_CY:
        wcsncpy_s( pwszType, cchType, L"VT_CY", cchType );
        break;
    case VT_DATE:
        wcsncpy_s( pwszType, cchType, L"VT_DATE", cchType );
        break;
    case VT_BSTR:
        wcsncpy_s( pwszType, cchType, L"VT_BSTR", cchType );
        break;
    case VT_ERROR:
        wcsncpy_s( pwszType, cchType, L"VT_ERROR", cchType );
        break;
    case VT_BOOL:
        wcsncpy_s( pwszType, cchType, L"VT_BOOL", cchType );
        break;
    case VT_VARIANT:
        wcsncpy_s( pwszType, cchType, L"VT_VARIANT", cchType );
        break;
    case VT_DECIMAL:
        wcsncpy_s( pwszType, cchType, L"VT_DECIMAL", cchType );
        break;
    case VT_I1:
        wcsncpy_s( pwszType, cchType, L"VT_I1", cchType );
        break;
    case VT_UI1:
        wcsncpy_s( pwszType, cchType, L"VT_UI1", cchType );
        break;
    case VT_INT:
        wcsncpy_s( pwszType, cchType, L"VT_INT", cchType );
        break;
    case VT_UINT:
        wcsncpy_s( pwszType, cchType, L"VT_UINT", cchType );
        break;
    case VT_VOID:
        wcsncpy_s( pwszType, cchType, L"VT_VOID", cchType );
        break;
    case VT_SAFEARRAY:
        wcsncpy_s( pwszType, cchType, L"VT_SAFEARRAY", cchType );
        break;
    case VT_USERDEFINED:
        wcsncpy_s( pwszType, cchType, L"VT_USERDEFINED", cchType );
        break;
    case VT_LPSTR:
        wcsncpy_s( pwszType, cchType, L"VT_LPSTR", cchType );
        break;
    case VT_LPWSTR:
        wcsncpy_s( pwszType, cchType, L"VT_LPWSTR", cchType );
        break;
    case VT_RECORD:
        wcsncpy_s( pwszType, cchType, L"VT_RECORD", cchType );
        break;
    case VT_FILETIME:
        wcsncpy_s( pwszType, cchType, L"VT_FILETIME", cchType );
        break;
    case VT_BLOB:
        wcsncpy_s( pwszType, cchType, L"VT_BLOB", cchType );
        break;
    case VT_STREAM:
        wcsncpy_s( pwszType, cchType, L"VT_STREAM", cchType );
        break;
    case VT_STORAGE:
        wcsncpy_s( pwszType, cchType, L"VT_STORAGE", cchType );
        break;
    case VT_STREAMED_OBJECT:
        wcsncpy_s( pwszType, cchType, L"VT_STREAMED_OBJECT", cchType );
        break;
    case VT_STORED_OBJECT:
        wcsncpy_s( pwszType, cchType, L"VT_BLOB_OBJECT", cchType );
        break;
    case VT_CF:
        wcsncpy_s( pwszType, cchType, L"VT_CF", cchType );
        break;
    case VT_CLSID:
        wcsncpy_s( pwszType, cchType, L"VT_CLSID", cchType );
        break;
    default:
        _snwprintf_s( pwszType, cchType, cchType, L"Unknown (%d)", 
                    vt & VT_TYPEMASK );
        break;
    }

    // Adjust cchType for the added characters.

    cchType -= wcslen(pwszType);

    // Add the type modifiers, if present.

    if( vt & VT_VECTOR )
    {
        pwszModifier = L" | VT_VECTOR";        
        wcsncat_s( pwszType, cchType, pwszModifier, cchType );
        cchType -= wcslen( pwszModifier );
    }

    if( vt & VT_ARRAY )
    {
        pwszModifier = L" | VT_ARRAY";        
        wcsncat_s( pwszType, cchType, pwszModifier, cchType );
        cchType -= wcslen( pwszModifier );
    }

    if( vt & VT_RESERVED )
    {
        pwszModifier = L" | VT_RESERVED";        
        wcsncat_s( pwszType, cchType, pwszModifier, cchType );
        cchType -= wcslen( pwszModifier );
    }

}
#pragma endregion

#pragma region Tracing
#if defined(_DEBUG) && defined(WIN32)
#define TRACEF OutputDebugStringFormat

void OutputDebugStringFormat( WCHAR* lpszFormat, ... )
{
	WCHAR    lpszBuffer[MAX_PATH];
	va_list  fmtList;

	va_start( fmtList, lpszFormat );
	vswprintf_s( lpszBuffer, MAX_PATH, lpszFormat, fmtList );
	va_end( fmtList );

   ::OutputDebugStringW( lpszBuffer );
}
#else
#define TRACEF	__noop
#endif
#pragma endregion

typedef HRESULT (CALLBACK* PFN_STGOPENSTGEX)(const WCHAR*, DWORD, DWORD, DWORD, void*, void*, REFIID riid, void **);

PFN_STGOPENSTGEX  v_pfnStgOpenStorageEx; // StgOpenStorageEx (Win2K/XP only)

BOOL GetStgOpenStorageEx()
{
	if (!v_pfnStgOpenStorageEx)
	{
		BOOL fRunningOnNT = ((GetVersion() & 0x80000000) != 0x80000000);
		v_pfnStgOpenStorageEx = ((fRunningOnNT) ? (PFN_STGOPENSTGEX)GetProcAddress(GetModuleHandle(L"OLE32"), "StgOpenStorageEx") : NULL);
	}

	return (v_pfnStgOpenStorageEx != NULL);
}


inline bool operator<(const PROPERTYKEY& a, const PROPERTYKEY& b)
{
	if (a.fmtid.Data1 != b.fmtid.Data1)
		return a.fmtid.Data1 < b.fmtid.Data1;
	else if (a.fmtid.Data2 != b.fmtid.Data2)
		return a.fmtid.Data2 < b.fmtid.Data2;
	else if (a.fmtid.Data3 != b.fmtid.Data3)
		return a.fmtid.Data3 < b.fmtid.Data3;
	// should be enough without Data4 checks (it is a char[8])
	else
		return a.pid < b.pid;
}

void ExportPropertySetData (xml_document<WCHAR> *doc, xml_node<WCHAR> *root, PROPERTYKEY* keys, DWORD& index, CComPtr<IPropertyStore> pStore);

void ExportMetadata (xml_document<WCHAR> *doc, wstring targetFile)
{
    HRESULT hr = E_UNEXPECTED;

	xml_node<WCHAR> *root = doc->allocate_node(node_element, MetadataNodeName);
	doc->append_node(root);

	if (GetStgOpenStorageEx())
	{
		CComPtr<IPropertySetStorage> pPropSetStg;
		CComPtr<IPropertyStore> pStore;
		PROPERTYKEY * keys = NULL;

		try
		{
			hr = (v_pfnStgOpenStorageEx)(targetFile.c_str(), STGM_READ | STGM_SHARE_EXCLUSIVE, STGFMT_FILE, 0, NULL, 0, 
					IID_IPropertySetStorage, (void**)&pPropSetStg);
			if( FAILED(hr) ) 
				throw CPHException(ERROR_OPEN_FAILED, IDS_E_IPSS_1, hr);

			hr = PSCreatePropertyStoreFromPropertySetStorage(pPropSetStg, STGM_READ, IID_IPropertyStore, (void **)&pStore);
			pPropSetStg.Release();
			if( FAILED(hr) ) 
				throw CPHException(ERROR_OPEN_FAILED, IDS_E_PSCREATE_1, hr);

			DWORD cProps;
			hr = pStore->GetCount(&cProps);
			if( FAILED(hr) ) 
				throw CPHException(ERROR_OPEN_FAILED, IDS_E_IPS_GETCOUNT_1, hr);

			keys = new PROPERTYKEY[cProps];

			for (DWORD i = 0; i < cProps; i++)
			{
				hr = pStore->GetAt(i, &keys[i]);
				if( FAILED(hr) ) 
					throw CPHException(ERROR_UNKNOWN_PROPERTY, IDS_E_IPS_GETAT_1, hr);
			}

			// Sort keys into their property sets
			// We used to use IPropertyStorage to get the grouping, but this worked badly with Unicode property value
			sort(keys, &keys[cProps]);

			// Loop through all the properties
			DWORD index = 0;

			while( index < cProps)
			{
				// Export the properties in the property set - throws exceptions on error
				ExportPropertySetData( doc, root, keys, index, pStore );
			}
		}
		catch (CPHException& e)
		{
			delete [] keys;
			throw e;
		}

		delete [] keys;
	}
}

// throws CPHException on error
void ExportPropertySetData (xml_document<WCHAR> *doc, xml_node<WCHAR> *root, PROPERTYKEY* keys, DWORD& index, CComPtr<IPropertyStore> pStore)
{
    HRESULT hr = E_UNEXPECTED;

    PROPVARIANT propvar;
	GUID currFmtid = keys[index].fmtid;

	PropVariantInit( &propvar );

	WCHAR * pGuid = doc->allocate_string(NULL, 64);
	StringFromGUID2( currFmtid, pGuid, 64);

	xml_node<WCHAR> *storage = doc->allocate_node(node_element, StorageNodeName);
	root->append_node(storage);

	xml_attribute<WCHAR> * attr = NULL;
    if( FMTID_SummaryInformation == currFmtid )
		attr = doc->allocate_attribute(DescriptionAttrName, L"SummaryInformation");
    else if( FMTID_DocSummaryInformation == currFmtid )
       attr = doc->allocate_attribute(DescriptionAttrName, L"DocumentSummaryInformation" );
    else if( FMTID_UserDefinedProperties == currFmtid )
       attr = doc->allocate_attribute(DescriptionAttrName, L"UserDefined" );
	if (attr != NULL)
		storage->append_attribute(attr);

	attr = doc->allocate_attribute(FormatIDAttrName, pGuid);
	storage->append_attribute(attr);

	try
	{
		// Loop through each property with the same FMTID

		for(; currFmtid == keys[index].fmtid; index++ )
		{
			// Read the property out of the property set
			PropVariantInit( &propvar );
			hr = pStore->GetValue(keys[index], &propvar);
			if( FAILED(hr) ) 
				throw CPHException(ERROR_UNKNOWN_PROPERTY, IDS_E_IPS_GETVALUE_3, hr, keys[index].pid, pGuid);

			// Export the property value, type, and so on.
			WCHAR* wszId = doc->allocate_string(NULL, 20);
			WCHAR* wszTypeId = doc->allocate_string(NULL, 20);
			WCHAR* wszType = doc->allocate_string(NULL, MAX_PATH + 1);
			WCHAR* wszValue = doc->allocate_string(NULL, MAX_PATH + 1);

			StringCbPrintf (wszId, 20, L"%d", keys[index].pid);
			StringCbPrintf (wszTypeId, 20, L"%d", propvar.vt);
			ConvertVarTypeToString( propvar.vt, wszType, MAX_PATH);

			xml_node<WCHAR> *prop = doc->allocate_node(node_element, PropertyNodeName);
			storage->append_node(prop);

			PWSTR pName = NULL;
			hr = PSGetNameFromPropertyKey(keys[index], &pName);

			// If we don't get a name, don't worry as it is for documentation only and not read on import
			if (SUCCEEDED(hr))
			{
				WCHAR* wszName = doc->allocate_string(pName, wcslen(pName)+1);
				CoTaskMemFree(pName);
				attr = doc->allocate_attribute(NameAttrName, wszName);
				prop->append_attribute(attr);
			}

			attr = doc->allocate_attribute(PropertyIdAttrName, wszId);
			prop->append_attribute(attr);

			attr = doc->allocate_attribute(TypeAttrName, wszType);
			prop->append_attribute(attr);

			attr = doc->allocate_attribute(TypeIdAttrName, wszTypeId);
			prop->append_attribute(attr);

			// PSFormatForDisplay would be natural here if we wanted max readability, as it formats nicely and respects locale,
			// but we use coercion because we're more concerned with round-tripping the value when we import it again.
			// The exception is the vector (array) types where we want the multi-value formatting, and coercion to a simple string fails anyway.
			// It does put a blank after each semicolon separator though, which we have to remove on import.
			if (propvar.vt & VT_VECTOR)
			{
				hr = PSFormatForDisplay(keys[index], propvar, PDFF_DEFAULT, wszValue, MAX_PATH);

				if (SUCCEEDED(hr)) 
				{
					xml_node<WCHAR> *node = doc->allocate_node(node_element, ValueNodeName, wszValue);
					prop->append_node(node);
				}
				else
					throw CPHException(ERROR_INVALID_FUNCTION, IDS_E_PSFORMAT_3, hr, keys[index].pid, pGuid);
			}
			else
			{
				PROPVARIANT propvarString = {0};
				hr = PropVariantChangeType(&propvarString, propvar, 0, VT_LPWSTR);

				if (SUCCEEDED(hr)) 
				{
					WCHAR* wszDisp = doc->allocate_string(propvarString.pwszVal, wcslen(propvarString.pwszVal)+1);

					xml_node<WCHAR> *node = doc->allocate_node(node_element, ValueNodeName, wszDisp);
					prop->append_node(node);
				}
				PropVariantClear(&propvarString);
				if (FAILED(hr))
					throw CPHException(ERROR_INVALID_FUNCTION, IDS_E_PSFORMAT_3, hr, keys[index].pid, pGuid);
			}
        }
    }
	catch (CPHException& e)
    {
		PropVariantClear( &propvar );
		throw e;
    }

    PropVariantClear( &propvar );
}

void ImportPropertySetData (xml_document<WCHAR> *doc, xml_node<WCHAR> *stor, FMTID fmtid, CComPtr<IPropertyStore> pStore);

// throws CPHException on error
void ImportMetadata (xml_document<WCHAR> *doc, wstring targetFile)
{
    HRESULT hr = E_UNEXPECTED;

	xml_node<WCHAR> *root = doc->allocate_node(node_element, MetadataNodeName);
	doc->append_node(root);

	if (GetStgOpenStorageEx())
	{
		CComPtr<IPropertySetStorage> pPropSetStg;
		CComPtr<IPropertyStore> pStore;	
		 
		hr = (v_pfnStgOpenStorageEx)(targetFile.c_str(), STGM_READWRITE | STGM_SHARE_EXCLUSIVE, STGFMT_FILE, 0, NULL, 0, 
				IID_IPropertySetStorage, (void**)&pPropSetStg);
		if( FAILED(hr) ) 
			throw CPHException(ERROR_OPEN_FAILED, IDS_E_IPSS_1, hr);

		// We use IPropertyStore for writing for simplicity
		hr = PSCreatePropertyStoreFromPropertySetStorage(pPropSetStg, STGM_READWRITE, IID_IPropertyStore, (void **)&pStore);
		pPropSetStg.Release();
		if( FAILED(hr) ) 
			throw CPHException(ERROR_OPEN_FAILED, IDS_E_PSCREATE_1, hr);

		xml_node<WCHAR>* root = doc->first_node();
		if (wcscmp(root->name(), MetadataNodeName) != 0)
			throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_ROOT_1, root->name());

		// iterate over the storages
		xml_node<WCHAR>* stor = root->first_node();
		while (stor)
		{
			if (wcscmp(stor->name(), StorageNodeName) != 0)
				throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_STORAGE_1, stor->name());

			xml_attribute<WCHAR>* id = stor->first_attribute(FormatIDAttrName);
			if (!id)
				throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_NOFORMATID);

			FMTID fmtid;
			hr = CLSIDFromString (id->value(), &fmtid);
			if (FAILED(hr))
				throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_BADFORMATID_1, id->value());

			ImportPropertySetData(doc, stor, fmtid, pStore);

			stor = stor->next_sibling();
		}

		pStore->Commit();
	}
}


// throws CPHException on error
void ImportPropertySetData (xml_document<WCHAR> *doc, xml_node<WCHAR> *stor, FMTID fmtid, CComPtr<IPropertyStore> pStore)
{
 	// iterate over the properties
	xml_node<WCHAR>* prop = stor->first_node();
	while (prop)
	{
		if (wcscmp(prop->name(), PropertyNodeName) != 0)
			throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_PROPERTY_1, prop->name());

		// OK if this is missing
		xml_attribute<WCHAR>* name = prop->first_attribute(NameAttrName);

		xml_attribute<WCHAR>* id = prop->first_attribute(PropertyIdAttrName);
		if (!id)
			throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_NOID);

		xml_attribute<WCHAR>* idType = prop->first_attribute(TypeIdAttrName);
		if (!idType)
			throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_NOTYPEID_1, name != NULL ? name->value(): id->value());

		xml_node<WCHAR>* val = prop->first_node(ValueNodeName);
		if (!val)
			throw CPHException(ERROR_XML_PARSE_ERROR, IDS_E_NOVALUE_1, name != NULL ? name->value(): id->value());

		WCHAR* stop;
		VARTYPE vt = (VARTYPE) wcstol(idType->value(), &stop, 10);
		PROPERTYKEY key;
		key.fmtid = fmtid;
		key.pid =  wcstol(id->value(), &stop, 10);

		PROPVARIANT propvarString = {0};
		PROPVARIANT propvarValue = {0};

		try
		{
			HRESULT hr;

			// Coercion does not handle array strings well, or other array types at all
			// We need to split the input into an array of string values that can be coerced
			if (vt & VT_VECTOR)
			{
				wstring s = val->value();
				std::vector<std::wstring> ss = wsplit(s, L';');
				
				PCWSTR * ps = NULL;
				ps = new PCWSTR[ss.size()];
				for (unsigned int i = 0; i < ss.size(); i++)
				{
					// Non-first elements begin with a blank after each ';' put there by formatting for display on export
					// If present, remove
					if (i > 0 && ss[i].size() > 0 && ss[i][0] == L' ')
						ps[i] = ss[i].c_str() + 1;
					else
						ps[i] = ss[i].c_str();
				}

				hr = InitPropVariantFromStringVector(ps, (ULONG)ss.size(), &propvarString);
				delete [] ps;
			}
			else
				hr = InitPropVariantFromString(val->value(), &propvarString);

			if (SUCCEEDED(hr))
			{
				hr = PropVariantChangeType(&propvarValue, propvarString, 0, vt);
				if (SUCCEEDED(hr))
				{
					hr = pStore->SetValue(key, propvarValue);
					if (FAILED(hr))
						throw CPHException(ERROR_UNKNOWN_PROPERTY, IDS_E_IPS_SETVALUE_2, hr, name != NULL ? name->value(): id->value());

					TRACEF(L"Set property with Name or Id %s to %s\n",  name != NULL ? name->value(): id->value(), val->value() );
	
					PropVariantClear(&propvarString);
					PropVariantClear(&propvarValue);
				}
				else
					throw CPHException(ERROR_INVALID_FUNCTION, IDS_E_VAR_COERCE_2, hr, name != NULL ? name->value(): id->value());
			}
			else
				throw CPHException(ERROR_INVALID_FUNCTION, IDS_E_VAR_INIT_2, hr, name != NULL ? name->value(): id->value());
		}
		catch (CPHException& e)
		{
			PropVariantClear(&propvarString);
			PropVariantClear(&propvarValue);
			throw e;
		}

		prop = prop->next_sibling();
	}
}

void DeleteMetadata (wstring targetFile)
{
    HRESULT hr = E_UNEXPECTED;

	if (GetStgOpenStorageEx())
	{
		CComPtr<IPropertySetStorage> pPropSetStg;
		CComPtr<IEnumSTATPROPSETSTG> penum;
		STATPROPSETSTG statpropsetstg;

		hr = (v_pfnStgOpenStorageEx)(targetFile.c_str(), STGM_READWRITE | STGM_SHARE_EXCLUSIVE, STGFMT_FILE, 0, NULL, 0, 
				IID_IPropertySetStorage, (void**)&pPropSetStg);
		if( FAILED(hr) ) 
			throw CPHException(ERROR_OPEN_FAILED, IDS_E_IPSS_1, hr);

		hr = pPropSetStg->Enum( &penum );
		if( FAILED(hr) ) 
			throw CPHException(ERROR_OPEN_FAILED, IDS_E_IPSS_ENUM_1, hr);

		memset( &statpropsetstg, 0, sizeof(statpropsetstg) );
	    hr = penum->Next( 1, &statpropsetstg, NULL );

	    // Delete all the property sets.
		while( S_OK == hr )
		{
			// Delete the property set.
			hr = pPropSetStg->Delete( statpropsetstg.fmtid);
			if( FAILED(hr) ) 
			{
				WCHAR pGuid[64];
				StringFromGUID2( statpropsetstg.fmtid, pGuid, 64);		
				throw CPHException(ERROR_UNKNOWN_PROPERTY, IDS_E_IPSS_DELETE_2, hr, pGuid);
			}

			// Get the next property set in the enumeration.
			hr = penum->Next( 1, &statpropsetstg, NULL );

		}
		if( FAILED(hr) )
			throw CPHException(ERROR_OPEN_FAILED, IDS_E_IPSS_ENUM_NEXT_1, hr);
	}
}


