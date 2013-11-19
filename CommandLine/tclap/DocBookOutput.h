// -*- Mode: c++; c-basic-offset: 4; tab-width: 4; -*-

/****************************************************************************** 
 * 
 *  file:  DocBookOutput.h
 * 
 *  Copyright (c) 2004, Michael E. Smoot
 *  All rights reverved.
 * 
 *  See the file COPYING in the top directory of this distribution for
 *  more information.
 *  
 *  THE SOFTWARE IS PROVIDED _AS IS_, WITHOUT WARRANTY OF ANY KIND, EXPRESS 
 *  OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
 *  THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
 *  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
 *  DEALINGS IN THE SOFTWARE.  
 *  
 *****************************************************************************/ 

#ifndef TCLAP_DOCBOOKOUTPUT_H
#define TCLAP_DOCBOOKOUTPUT_H

#include <string>
#include <vector>
#include <list>
#include <iostream>
#include <algorithm>

#include <CmdLineInterface.h>
#include <CmdLineOutput.h>
#include <XorHandler.h>
#include <Arg.h>

namespace TCLAP {

/**
 * A class that generates DocBook output for usage() method for the 
 * given CmdLine and its Args.
 */
class DocBookOutput : public CmdLineOutput
{

	public:

		/**
		 * Prints the usage to stdout.  Can be overridden to 
		 * produce alternative behavior.
		 * \param c - The CmdLine object the output is generated for. 
		 */
		virtual void usage(CmdLineInterface& c);

		/**
		 * Prints the version to stdout. Can be overridden 
		 * to produce alternative behavior.
		 * \param c - The CmdLine object the output is generated for. 
		 */
		virtual void version(CmdLineInterface& c);

		/**
		 * Prints (to stderr) an error message, short usage 
		 * Can be overridden to produce alternative behavior.
		 * \param c - The CmdLine object the output is generated for. 
		 * \param e - The ArgException that caused the failure. 
		 */
		virtual void failure(CmdLineInterface& c, 
						     ArgException& e );

	protected:

		/**
		 * Substitutes the char r for string x in string s.
		 * \param s - The string to operate on. 
		 * \param r - The char to replace. 
		 * \param x - What to replace r with. 
		 */
		void substituteSpecialChars( std::wstring& s, TCHAR r, std::wstring& x );
		void removeChar( std::wstring& s, TCHAR r);
		void basename( std::wstring& s );

		void printShortArg(Arg* it);
		void printLongArg(Arg* it);

		TCHAR theDelimiter;
};


inline void DocBookOutput::version(CmdLineInterface& _cmd) 
{ 
	std::wcout << _cmd.getVersion() << std::endl;
}

inline void DocBookOutput::usage(CmdLineInterface& _cmd ) 
{
	std::list<Arg*> argList = _cmd.getArgList();
	std::wstring progName = _cmd.getProgramName();
	std::wstring xversion = _cmd.getVersion();
	theDelimiter = _cmd.getDelimiter();
	XorHandler xorHandler = _cmd.getXorHandler();
	std::vector< std::vector<Arg*> > xorList = xorHandler.getXorList();
	basename(progName);

	std::wcout << "<?xml version='1.0'?>" << std::endl;
	std::wcout << "<!DOCTYPE refentry PUBLIC \"-//OASIS//DTD DocBook XML V4.2//EN\"" << std::endl;
	std::wcout << "\t\"http://www.oasis-open.org/docbook/xml/4.2/docbookx.dtd\">" << std::endl << std::endl;

	std::wcout << "<refentry>" << std::endl;

	std::wcout << "<refmeta>" << std::endl;
	std::wcout << "<refentrytitle>" << progName << "</refentrytitle>" << std::endl;
	std::wcout << "<manvolnum>1</manvolnum>" << std::endl;
	std::wcout << "</refmeta>" << std::endl;

	std::wcout << "<refnamediv>" << std::endl;
	std::wcout << "<refname>" << progName << "</refname>" << std::endl;
	std::wcout << "<refpurpose>" << _cmd.getMessage() << "</refpurpose>" << std::endl;
	std::wcout << "</refnamediv>" << std::endl;

	std::wcout << "<refsynopsisdiv>" << std::endl;
	std::wcout << "<cmdsynopsis>" << std::endl;

	std::wcout << "<command>" << progName << "</command>" << std::endl;

	// xor
	for ( int i = 0; (unsigned int)i < xorList.size(); i++ )
	{
		std::wcout << "<group choice='req'>" << std::endl;
		for ( ArgVectorIterator it = xorList[i].begin(); 
						it != xorList[i].end(); it++ )
			printShortArg((*it));

		std::wcout << "</group>" << std::endl;
	}
	
	// rest of args
	for (ArgListIterator it = argList.begin(); it != argList.end(); it++)
		if ( !xorHandler.contains( (*it) ) )
			printShortArg((*it));

 	std::wcout << "</cmdsynopsis>" << std::endl;
	std::wcout << "</refsynopsisdiv>" << std::endl;

	std::wcout << "<refsect1>" << std::endl;
	std::wcout << "<title>Description</title>" << std::endl;
	std::wcout << "<para>" << std::endl;
	std::wcout << _cmd.getMessage() << std::endl; 
	std::wcout << "</para>" << std::endl;
	std::wcout << "</refsect1>" << std::endl;

	std::wcout << "<refsect1>" << std::endl;
	std::wcout << "<title>Options</title>" << std::endl;

	std::wcout << "<variablelist>" << std::endl;
	
	for (ArgListIterator it = argList.begin(); it != argList.end(); it++)
		printLongArg((*it));

	std::wcout << "</variablelist>" << std::endl;
	std::wcout << "</refsect1>" << std::endl;

	std::wcout << "<refsect1>" << std::endl;
	std::wcout << "<title>Version</title>" << std::endl;
	std::wcout << "<para>" << std::endl;
	std::wcout << xversion << std::endl; 
	std::wcout << "</para>" << std::endl;
	std::wcout << "</refsect1>" << std::endl;
	
	std::wcout << "</refentry>" << std::endl;

}

inline void DocBookOutput::failure( CmdLineInterface& _cmd,
				    ArgException& e ) 
{ 
	static_cast<void>(_cmd); // unused
	std::wcout << e.what() << std::endl;
	throw ExitException(1);
}

inline void DocBookOutput::substituteSpecialChars( std::wstring& s,
				                                   TCHAR r,
												   std::wstring& x )
{
	size_t p;
	while ( (p = s.find_first_of(r)) != std::wstring::npos )
	{
		s.erase(p,1);
		s.insert(p,x);
	}
}

inline void DocBookOutput::removeChar( std::wstring& s, TCHAR r)
{
	size_t p;
	while ( (p = s.find_first_of(r)) != std::wstring::npos )
	{
		s.erase(p,1);
	}
}

inline void DocBookOutput::basename( std::wstring& s )
{
	size_t p = s.find_last_of('/');
	if ( p != std::wstring::npos )
	{
		s.erase(0, p + 1);
	}
}

inline void DocBookOutput::printShortArg(Arg* a)
{
	std::wstring lt = "&lt;"; 
	std::wstring gt = "&gt;"; 

	std::wstring id = a->shortID();
	substituteSpecialChars(id,'<',lt);
	substituteSpecialChars(id,'>',gt);
	removeChar(id,'[');
	removeChar(id,']');
	
	std::wstring choice = "opt";
	if ( a->isRequired() )
		choice = "plain";

	std::wcout << "<arg choice='" << choice << '\'';
	if ( a->acceptsMultipleValues() )
		std::wcout << " rep='repeat'";


	std::wcout << '>';
	if ( !a->getFlag().empty() )
		std::wcout << a->flagStartChar() << a->getFlag();
	else
		std::wcout << a->nameStartString() << a->getName();
	if ( a->isValueRequired() )
	{
		std::wstring arg = a->shortID();
		removeChar(arg,'[');
		removeChar(arg,']');
		removeChar(arg,'<');
		removeChar(arg,'>');
		arg.erase(0, arg.find_last_of(theDelimiter) + 1);
		std::wcout << theDelimiter;
		std::wcout << "<replaceable>" << arg << "</replaceable>";
	}
	std::wcout << "</arg>" << std::endl;

}

inline void DocBookOutput::printLongArg(Arg* a)
{
	std::wstring lt = "&lt;"; 
	std::wstring gt = "&gt;"; 

	std::wstring desc = a->getDescription();
	substituteSpecialChars(desc,'<',lt);
	substituteSpecialChars(desc,'>',gt);

	std::wcout << "<varlistentry>" << std::endl;

	if ( !a->getFlag().empty() )
	{
		std::wcout << "<term>" << std::endl;
		std::wcout << "<option>";
		std::wcout << a->flagStartChar() << a->getFlag();
		std::wcout << "</option>" << std::endl;
		std::wcout << "</term>" << std::endl;
	}

	std::wcout << "<term>" << std::endl;
	std::wcout << "<option>";
	std::wcout << a->nameStartString() << a->getName();
	if ( a->isValueRequired() )
	{
		std::wstring arg = a->shortID();
		removeChar(arg,'[');
		removeChar(arg,']');
		removeChar(arg,'<');
		removeChar(arg,'>');
		arg.erase(0, arg.find_last_of(theDelimiter) + 1);
		std::wcout << theDelimiter;
		std::wcout << "<replaceable>" << arg << "</replaceable>";
	}
	std::wcout << "</option>" << std::endl;
	std::wcout << "</term>" << std::endl;

	std::wcout << "<listitem>" << std::endl;
	std::wcout << "<para>" << std::endl;
	std::wcout << desc << std::endl;
	std::wcout << "</para>" << std::endl;
	std::wcout << "</listitem>" << std::endl;

	std::wcout << "</varlistentry>" << std::endl;
}

} //namespace TCLAP
#endif 
