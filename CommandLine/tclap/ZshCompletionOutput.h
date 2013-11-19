// -*- Mode: c++; c-basic-offset: 4; tab-width: 4; -*-

/****************************************************************************** 
 * 
 *  file:  ZshCompletionOutput.h
 * 
 *  Copyright (c) 2006, Oliver Kiddle
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

#ifndef TCLAP_ZSHCOMPLETIONOUTPUT_H
#define TCLAP_ZSHCOMPLETIONOUTPUT_H

#include <string>
#include <vector>
#include <list>
#include <iostream>
#include <map>

#include <CmdLineInterface.h>
#include <CmdLineOutput.h>
#include <XorHandler.h>
#include <Arg.h>

namespace TCLAP {

/**
 * A class that generates a Zsh completion function as output from the usage()
 * method for the given CmdLine and its Args.
 */
class ZshCompletionOutput : public CmdLineOutput
{

	public:

		ZshCompletionOutput();

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

		void basename( std::wstring& s );
		void quoteSpecialChars( std::wstring& s );

		std::wstring getMutexList( CmdLineInterface& _cmd, Arg* a );
		void printOption( Arg* it, std::wstring mutex );
		void printArg( Arg* it );

		std::map<std::wstring, std::wstring> common;
		TCHAR theDelimiter;
};

ZshCompletionOutput::ZshCompletionOutput()
: common(std::map<std::wstring, std::wstring>()),
  theDelimiter('=')
{
	common["host"] = "_hosts";
	common["hostname"] = "_hosts";
	common["file"] = "_files";
	common["filename"] = "_files";
	common["user"] = "_users";
	common["username"] = "_users";
	common["directory"] = "_directories";
	common["path"] = "_directories";
	common["url"] = "_urls";
}

inline void ZshCompletionOutput::version(CmdLineInterface& _cmd)
{
	std::wcout << _cmd.getVersion() << std::endl;
}

inline void ZshCompletionOutput::usage(CmdLineInterface& _cmd )
{
	std::list<Arg*> argList = _cmd.getArgList();
	std::wstring progName = _cmd.getProgramName();
	std::wstring xversion = _cmd.getVersion();
	theDelimiter = _cmd.getDelimiter();
	basename(progName);

	std::wcout << "#compdef " << progName << std::endl << std::endl <<
		"# " << progName << " version " << _cmd.getVersion() << std::endl << std::endl <<
		"_arguments -s -S";

	for (ArgListIterator it = argList.begin(); it != argList.end(); it++)
	{
		if ( (*it)->shortID().at(0) == '<' )
			printArg((*it));
		else if ( (*it)->getFlag() != "-" )
			printOption((*it), getMutexList(_cmd, *it));
	}

	std::wcout << std::endl;
}

inline void ZshCompletionOutput::failure( CmdLineInterface& _cmd,
				                ArgException& e )
{
	static_cast<void>(_cmd); // unused
	std::wcout << e.what() << std::endl;
}

inline void ZshCompletionOutput::quoteSpecialChars( std::wstring& s )
{
	size_t idx = s.find_last_of(':');
	while ( idx != std::wstring::npos )
	{
		s.insert(idx, 1, '\\');
		idx = s.find_last_of(':', idx);
	}
	idx = s.find_last_of('\'');
	while ( idx != std::wstring::npos )
	{
		s.insert(idx, "'\\'");
		if (idx == 0)
			idx = std::wstring::npos;
		else
			idx = s.find_last_of('\'', --idx);
	}
}

inline void ZshCompletionOutput::basename( std::wstring& s )
{
	size_t p = s.find_last_of('/');
	if ( p != std::wstring::npos )
	{
		s.erase(0, p + 1);
	}
}

inline void ZshCompletionOutput::printArg(Arg* a)
{
	static int count = 1;

	std::wcout << " \\" << std::endl << "  '";
	if ( a->acceptsMultipleValues() )
		std::wcout << '*';
	else
		std::wcout << count++;
	std::wcout << ':';
	if ( !a->isRequired() )
		std::wcout << ':';

	std::wcout << a->getName() << ':';
	std::map<std::wstring, std::wstring>::iterator compArg = common.find(a->getName());
	if ( compArg != common.end() )
	{
		std::wcout << compArg->second;
	}
	else
	{
		std::wcout << "_guard \"^-*\" " << a->getName();
	}
	std::wcout << '\'';
}

inline void ZshCompletionOutput::printOption(Arg* a, std::wstring mutex)
{
	std::wstring flag = a->flagStartChar() + a->getFlag();
	std::wstring name = a->nameStartString() + a->getName();
	std::wstring desc = a->getDescription();

	// remove full stop and capitalisation from description as
	// this is the convention for zsh function
	if (!desc.compare(0, 12, "(required)  "))
	{
		desc.erase(0, 12);
	}
	if (!desc.compare(0, 15, "(OR required)  "))
	{
		desc.erase(0, 15);
	}
	size_t len = desc.length();
	if (len && desc.at(--len) == '.')
	{
		desc.erase(len);
	}
	if (len)
	{
		desc.replace(0, 1, 1, tolower(desc.at(0)));
	}

	std::wcout << " \\" << std::endl << "  '" << mutex;

	if ( a->getFlag().empty() )
	{
		std::wcout << name;
	}
	else
	{
		std::wcout << "'{" << flag << ',' << name << "}'";
	}
	if ( theDelimiter == '=' && a->isValueRequired() )
		std::wcout << "=-";
	quoteSpecialChars(desc);
	std::wcout << '[' << desc << ']';

	if ( a->isValueRequired() )
	{
		std::wstring arg = a->shortID();
		arg.erase(0, arg.find_last_of(theDelimiter) + 1);
		if ( arg.at(arg.length()-1) == ']' )
			arg.erase(arg.length()-1);
		if ( arg.at(arg.length()-1) == ']' )
		{
			arg.erase(arg.length()-1);
		}
		if ( arg.at(0) == '<' )
		{
			arg.erase(arg.length()-1);
			arg.erase(0, 1);
		}
		size_t p = arg.find('|');
		if ( p != std::wstring::npos )
		{
			do
			{
				arg.replace(p, 1, 1, ' ');
			}
			while ( (p = arg.find_first_of('|', p)) != std::wstring::npos );
			quoteSpecialChars(arg);
			std::wcout << ": :(" << arg << ')';
		}
		else
		{
			std::wcout << ':' << arg;
			std::map<std::wstring, std::wstring>::iterator compArg = common.find(arg);
			if ( compArg != common.end() )
			{
				std::wcout << ':' << compArg->second;
			}
		}
	}

	std::wcout << '\'';
}

inline std::wstring ZshCompletionOutput::getMutexList( CmdLineInterface& _cmd, Arg* a)
{
	XorHandler xorHandler = _cmd.getXorHandler();
	std::vector< std::vector<Arg*> > xorList = xorHandler.getXorList();
	
	if (a->getName() == "help" || a->getName() == "version")
	{
		return "(-)";
	}

	std::ostringstream list;
	if ( a->acceptsMultipleValues() )
	{
		list << '*';
	}

	for ( int i = 0; static_cast<unsigned int>(i) < xorList.size(); i++ )
	{
		for ( ArgVectorIterator it = xorList[i].begin();
			it != xorList[i].end();
			it++)
		if ( a == (*it) )
		{
			list << '(';
			for ( ArgVectorIterator iu = xorList[i].begin();
				iu != xorList[i].end();
				iu++ )
			{
				bool notCur = (*iu) != a;
				bool hasFlag = !(*iu)->getFlag().empty();
				if ( iu != xorList[i].begin() && (notCur || hasFlag) )
					list << ' ';
				if (hasFlag)
					list << (*iu)->flagStartChar() << (*iu)->getFlag() << ' ';
				if ( notCur || hasFlag )
					list << (*iu)->nameStartString() << (*iu)->getName();
			}
			list << ')';
			return list.str();
		}
	}
	
	// wasn't found in xor list
	if (!a->getFlag().empty()) {
		list << "(" << a->flagStartChar() << a->getFlag() << ' ' <<
			a->nameStartString() << a->getName() << ')';
	}
	
	return list.str();
}

} //namespace TCLAP
#endif
