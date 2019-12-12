// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

#include "stdafx.h"
#include "XmlHelpers.h"
#include "tclap/CmdLine.h"
#include "resource.h"
#include <iostream>
#include <algorithm>
#include <strsafe.h>
#include <direct.h>

using namespace rapidxml;
using namespace TCLAP;
using namespace std;

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

		// Define Explorer view of meta data switch
		SwitchArg explorerSwitch(L"v", L"explorer", L"Export metadata thisExplorer sees", false);
		cmd.add(explorerSwitch);

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
		else if (explorerSwitch.isSet())
		{
			if (!exportSwitch.isSet())
				throw ArgException(L"-v can only be used with -e", L"explorer");
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
			else if (0 == checker.HasPropertyHandler(targetFile) ||
					 (-1 == checker.HasPropertyHandler(targetFile) && !explorerSwitch.isSet()))
			{
				// Skip files that do not have our property handler,
				// unless we were asked for the Explorer view
				continue;
			}
			else if (deleteSwitch.isSet())
			{
				// We need to check if metadata is present, to avoid adding an empty alternate stream by opening
				// r/w when no metadata stream is present
				if (S_OK != MetadataPresent(targetFile))
					continue;

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
					ExportMetadata(&doc, targetFile, explorerSwitch.isSet());

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
							throw CPHException(err, E_FAIL, IDS_E_FILEOPEN_1, err);

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

							CPHException cphe = CPHException(ERROR_XML_PARSE_ERROR, E_FAIL, IDS_E_XML_PARSE_ERROR_3, error, content, xmlFile.c_str());
							delete [] error;
							throw cphe;
						}

						// apply it 
						ImportMetadata(&doc, targetFile);
					}
					else
						throw CPHException(err, E_FAIL, IDS_E_FILEOPEN_1, err);

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
		wcerr << e.GetMessage() << endl;
		result = e.GetError();
	}

	if (prompt)
	{
		wcout << L"Hit any key to continue...";
		_getwch();
	}

	CoUninitialize();

	return result;
}

// An implementation of this is required by XmlHelpers
int AccessResourceString(UINT uId, LPWSTR lpBuffer, int nBufferMax)
{
	return LoadStringW(GetModuleHandle(NULL), uId, lpBuffer, nBufferMax);
}

