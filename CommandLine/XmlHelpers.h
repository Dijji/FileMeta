// Copyright (c) 2014, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.#include "stdafx.h"
#include "stdafx.h"
#include <string>
#include <map>
#include <vector>
#include <shlwapi.h>
#include <propsys.h>
#include <propkey.h>
#include <Propvarutil.h>
#undef RAPIDXML_NO_EXCEPTIONS
#include "rapidxml.hpp"
#include "rapidxml_print.hpp"
#include "resource.h"

using namespace rapidxml;
using namespace std;

static const WCHAR* MetadataFileSuffix	= L".metadata.xml";

class CPHException 
{
public:
	CPHException (int err, HRESULT hr, UINT uResourceId, ...);

	int GetError();
	HRESULT GetHResult();
	WCHAR*  GetMessage();

private:
	int		_err;		// error code from winerror.h
	HRESULT _hr;		// COM error
	WCHAR   _pszError[MAX_PATH];
};

class CExtensionChecker
{
private:
	std::map<std::wstring, BOOL> m_extensions;		// map of extensions checked so far

public:
	// See if file is handled by our property handler
	BOOL HasOurPropertyHandler(wstring fileName);
};

int AccessResourceString(UINT uId, LPWSTR lpBuffer, int nBufferMax);

// Converters and string helpers
std::vector<std::wstring> &wsplit(const std::wstring &s, WCHAR delim, std::vector<std::wstring> &elems);
std::vector<std::wstring> wsplit(const std::wstring &s, WCHAR delim);

void ConvertVarTypeToString( VARTYPE vt, WCHAR *pwszType, size_t cchType );

// Tracing
#if defined(_DEBUG) && defined(WIN32)
#define TRACEF OutputDebugStringFormat
void OutputDebugStringFormat( WCHAR* lpszFormat, ... );
#else
#define TRACEF	__noop
#endif

typedef HRESULT (CALLBACK* PFN_STGOPENSTGEX)(const WCHAR*, DWORD, DWORD, DWORD, void*, void*, REFIID riid, void **);

HRESULT MetadataPresent(wstring targetFile);
void ExportMetadata (xml_document<WCHAR> *doc, wstring targetFile);
void ExportPropertySetData (xml_document<WCHAR> *doc, xml_node<WCHAR> *root, PROPERTYKEY* keys, DWORD& index, CComPtr<IPropertyStore> pStore);

void ImportMetadata (xml_document<WCHAR> *doc, wstring targetFile);
void ImportPropertySetData (xml_document<WCHAR> *doc, xml_node<WCHAR> *stor, FMTID fmtid, CComPtr<IPropertyStore> pStore);

void DeleteMetadata (wstring targetFile);



