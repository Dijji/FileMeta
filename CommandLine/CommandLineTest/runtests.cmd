@ECHO off
SETLOCAL
:: Run this with CommandLineTest as the working folder
::
:: If FileMeta.exe is not on the path, tweak the following line to point to its location
SET _filemeta="FileMeta.exe"
::SET _filemeta="C:\Program Files (x86)\File Metadata\FileMeta.exe"

ECHO Setting up test files...
type NUL > allprops.txt
type NUL > fewprops.txt
CALL %_filemeta% -i allprops.txt > nul || goto error
CALL %_filemeta% -i fewprops.txt > nul || goto error

ECHO Test 1: Export 600 properties, import onto new file, export again and compare
SET Test=Test1
CALL %_filemeta% -e -x=props.xml allprops.txt > nul || goto error
type NUL > temp.txt
CALL %_filemeta% -i -x=props.xml temp.txt > nul || goto error
CALL %_filemeta% -e -x=props2.xml temp.txt > nul || goto error
fc /b props.xml props2.xml > nul || goto error
del temp.txt
del props.xml
del props2.xml
ECHO %Test% passed

ECHO Test2: Check that failures return error codes
SET Test=Test2
CALL %_filemeta% -e nofile.txt 2> nul && goto error
type NUL > temp.txt
CALL %_filemeta% -i -x=noxml.xml temp.txt 2> nul && goto error
CALL %_filemeta% -e -x=nodir\noxml.xml temp.txt 2> nul && goto error
CALL %_filemeta% -e -f=nodir temp.txt 2> nul && goto error
del temp.txt
ECHO %Test% passed

ECHO Test3: Check that illegal option combinations fail
SET Test=Test3
CALL %_filemeta% -e -x=props.xml -f=temp fewprops.txt 2> nul && goto error
CALL %_filemeta% -e -x=props.xml *.txt 2> nul && goto error
CALL %_filemeta% -e -x=props.xml -d=nodir fewprops.txt 2> nul && goto error
CALL %_filemeta% -ec -x=props.xml fewprops.txt 2> nul && goto error
CALL %_filemeta% -ec -f=. fewprops.txt 2> nul && goto error
type NUL > temp.txt
CALL %_filemeta% -ic temp.txt 2> nul && goto error
CALL %_filemeta% -dc temp.txt 2> nul && goto error
del temp.txt
ECHO %Test% passed

ECHO Test4: Export to and import from subfolder
SET Test=Test4
md temp
type NUL > temp.txt
CALL %_filemeta% -e -x=temp\props.xml fewprops.txt > nul || goto error
CALL %_filemeta% -i -x=temp\props.xml temp.txt > nul || goto error
CALL %_filemeta% -e -x=props2.xml temp.txt > nul || goto error
fc /b temp\props.xml props2.xml > nul || goto error
CALL %_filemeta% -e -f=temp fewprops.txt > nul || goto error
ren temp\fewprops.txt.metadata.xml temp.txt.metadata.xml > nul || goto error
CALL %_filemeta% -i -f=temp temp.txt > nul || goto error
fc /b temp\props.xml temp\temp.txt.metadata.xml > nul || goto error
del temp.txt
del props2.xml
del /q temp\*.*
rd temp
ECHO %Test% passed

ECHO Test5: Test wildcards against FOR loops
SET Test=Test5
md temp
md temp2
type NUL > temp.txt
CALL %_filemeta% -e -f=temp *.txt > nul || goto error
FOR /f %%G IN ('dir /b *.txt') DO (
  CALL %_filemeta% -e -f=temp2 %%G > nul || goto error )
FOR /f %%G IN ('dir /b temp\*.xml') DO (
  fc /b temp\%%G temp2\%%G > nul || goto error )
del /q temp\*.*
del /q temp2\*.*
del temp.txt
rd temp
rd temp2
ECHO %Test% passed

ECHO Test6: Test delete metadata
SET Test=Test6
copy fewprops.txt temp.txt > nul || goto error
CALL %_filemeta% -e -x=before.xml temp.txt > nul || goto error
CALL %_filemeta% -d temp.txt > nul || goto error
CALL %_filemeta% -e -x=after.xml temp.txt > nul || goto error
set size=0
call :filesize before.xml
IF %size% LEQ 32 goto error
call :filesize after.xml
IF %size% GTR 32 goto error
del temp.txt
del before.xml
del after.xml
ECHO %Test% passed

:passed
del allprops.txt
del fewprops.txt
ECHO All tests completed successfully
goto :eof

:: set filesize of 1st argument in %size% variable, and return
:filesize
  set size=%~z1
  exit /b 0

:error
 ECHO %Test% failed
