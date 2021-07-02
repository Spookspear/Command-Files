::--------------------------------------------------------
::      Author: G Bishop
::        Date: 27th August 2020
::        Name: %sLinkName%_ReadFileNew.cmd
:: Description:
::     History: See end of file
::--------------------------------------------------------
@echo off

setlocal
setlocal EnableExtensions EnableDelayedExpansion
set VERSION=01.01.01

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

set Q_ALLOW_PAUSE=Yes
set Q_ALLOW_LOGFILE=No

call:Main
goto:myExit

::--------------------------------------------------------
::-- Main
::--------------------------------------------------------
:Main
if /i "%DEBUG%" equ "Yes" (@echo Main - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call:preSetup
    call:doWork

if /i "%DEBUG%" equ "Yes" (@echo Main - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- doWork
::--------------------------------------------------------
:doWork
if /i "%DEBUG%" equ "Yes" (@echo doWork - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    for /f "tokens=* delims= " %%f in (%sCurrFile%.txt) do (
        call:fixFolder "%%f"
    )

if /i "%DEBUG%" equ "Yes" (@echo doWork - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- fixFolder
::--------------------------------------------------------
:fixFolder
if /i "%DEBUG%" equ "Yes" (@echo fixFolder - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFullPath=%~1
    if /i "%DEBUG%" equ "Yes" (pause)

    if exist "%sFullPath%" (

        call:Logging "%SEPERATOR1%"
        call:Logging "Folder: %sFullPath%"
        if /i "%DEBUG%" equ "Yes" (pause)

        call:Logging "Transferring: %sLinkName%"
        copy %sLinkName% "%sFullPath%" >nul
        if /i "%DEBUG%" equ "Yes" (pause)

        pushd "%sFullPath%"
        if /i "%DEBUG%" equ "Yes" (pause)

        %sLinkName% No
        if /i "%DEBUG%" equ "Yes" (pause)

        :: del %sLinkName% /q >nul
        if /i "%DEBUG%" equ "Yes" (pause)

        popd
        if /i "%DEBUG%" equ "Yes" (pause)

    ) else (
        call:Logging "%sFullPath% Not found"

    )

    call:Logging "%SEPERATOR1%"

if /i "%DEBUG%" equ "Yes" (@echo fixFolder - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- preSetup
::--------------------------------------------------------
:preSetup
if /i "%DEBUG%" equ "Yes" (@echo preSetup - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call:setVars
    call:setScreen
    call:echoVersion
    call:preChecks

if /i "%DEBUG%" equ "Yes" (@echo preSetup - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- preChecks
::--------------------------------------------------------
:preChecks
if /i "%DEBUG%" equ "Yes" (@echo preChecks - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not exist %sLinkName% (call:myExitError "Missing files ...")

if /i "%DEBUG%" equ "Yes" (@echo preChecks - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
if /i "%DEBUG%" equ "Yes" (@echo setVars - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not defined Q_ALLOW_PAUSE         (set Q_ALLOW_PAUSE=Yes)
    if not defined Q_ALLOW_LOGFILE       (set Q_ALLOW_LOGFILE=No)
    if not defined Q_ALLOW_SPECIAL_CHARS (set Q_ALLOW_SPECIAL_CHARS=Yes)

    set /a iTimeOut=3
    set sSeperatorChar1=-
    set sSeperatorChar2==

    set sLinkName=reNameFolders.cmd.lnk

    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar1=Ä)
    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar2=Í)
    call:getRoot myRoot

    if "%Q_ALLOW_LOGFILE%" equ "Yes" (
        set LOGFOLDER=C:\Bin\Logs
        call:setLogFile "%myRoot%" LOGFILE
    )

    call:setEchoData

if /i "%DEBUG%" equ "Yes" (@echo setVars - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::---------------------------------------------------------
::-- setLogFile
::---------------------------------------------------------
:setLogFile
if /i "%DEBUG%" equ "Yes" (@echo setLogFile - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not defined sCurrFile (call:getNameNoExt sCurrFile)

    for /F "tokens=1-4 delims=/- " %%A in ('date/T') do set myFileDate=%%A%%B
    for /F "tokens=1,2* delims=: " %%i in ('time/T') do set myFileTime=%%i%%j%%k
    set "%~2=%~1\%sCurrFile%-%myFileDate%-%MyFileTime%.Log"

if /i "%DEBUG%" equ "Yes" (@echo setLogFile - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setScreen
::--------------------------------------------------------
:setScreen
if /i "%DEBUG%" equ "Yes" (@echo setScreen - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call:calculateHeight %COMPUTERNAME% iWidth iHeight iColor

    if /i "%DEBUG%" equ "Yes" (@echo  iWidth %iWidth%)
    if /i "%DEBUG%" equ "Yes" (@echo iHeight %iHeight%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%DEBUG%" neq "Yes" (mode con: cols=%iWidth% lines=%iHeight%)
    :: mode con: cols=%iWidth% lines=%iHeight%
    set /a iWidth-=iSubtractWidth

    call:Replicate %iWidth% "%sSeperatorChar1%" SEPERATOR1
    call:Replicate %iWidth% "%sSeperatorChar2%" SEPERATOR2

    call:setColour %iColor%

if /i "%DEBUG%" equ "Yes" (@echo setScreen - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- calculateHeight
::--------------------------------------------------------
:calculateHeight
if /i "%DEBUG%" equ "Yes" (@echo calculateHeight - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not defined iW_Buff  (set /a iW_Buff=0)
    if not defined iH_Buff  (set /a iH_Buff=0)

    if /i "%DEBUG%" equ "Yes" (@echo iEchoWidth %iEchoWidth%)
    if /i "%DEBUG%" equ "Yes" (pause)

    set /a iW=%iEchoWidth%
    set /a iH=3333
    set iC=1b

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (set /a iH_Buff+=1)

    set /a %~2=%iW%+%iW_Buff%
    set /a %~3=%iH%+%iH_Buff%
    set %~4=%iC%

if /i "%DEBUG%" equ "Yes" (@echo calculateHeight - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
if /i "%DEBUG%" equ "Yes" (@echo echoVersion - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    Title %strEcho%
    call:Logging "%SEPERATOR2%"
    call:Logging "%strEcho%"
    call:Logging "%SEPERATOR1%"

if /i "%DEBUG%" equ "Yes" (@echo echoVersion - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setEchoData
::--------------------------------------------------------
:setEchoData
if /i "%DEBUG%" equ "Yes" (@echo setEchoData - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not defined sCurrFile (call:getNameNoExt sCurrFile)

    set strEcho=%sCurrFile% - Executing command file contents - will send file: %sLinkName% - %cd% - Version: %VERSION%
    set sSubtractMsg=%computername% %date% %time% :=

    set strEchoMeasure=%sSubtractMsg% %strEcho%

    call:calculateLength "%sSubtractMsg%"   iSubtractWidth
    call:calculateLength "%strEchoMeasure%" iEchoWidth

    set /a iH_Buff=3

    if /i "%DEBUG%" equ "Yes" (@echo iSubtractWidth %iSubtractWidth%)
    if /i "%DEBUG%" equ "Yes" (@echo iEchoWidth     %iEchoWidth%)
    if /i "%DEBUG%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo setEchoData - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setColour
::--------------------------------------------------------
:setColour
if /i "%DEBUG%" equ "Yes" (@echo setColour - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%~1" equ "Error" (color cf
    ) else (if /i "%~1" equ "Warning" (color 5e
    ) else (if /i "%~1" equ "Info"    (color 1f
    ) else (if /i "%~1" equ "Debug"   (color 3e
    ) else (if /i "%~1" equ "Good"    (color 2f
    ) else (if /i "%~1" equ "Present" (color 9f
    ) else (color %~1))))))

if /i "%DEBUG%" equ "Yes" (@echo setColour - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- Replicate
::--------------------------------------------------------
:Replicate
if /i "%DEBUG%" equ "Yes" (@echo Replicate - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set /a iLoop=%~1-1
    set repeatChar=%~2
    set returnStr=
    for /L %%G in (1,1,!iLoop!) do (set "returnStr=!returnStr!!repeatChar!")
    set %~3=%returnStr%

if /i "%DEBUG%" equ "Yes" (@echo Replicate - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- myExitError
::--------------------------------------------------------
:myExitError
if /i "%DEBUG%" equ "Yes" (@echo myExitError - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set Q_ALLOW_PAUSE=Yes
    call:setColour Error
    call:Logging "%~1"
    goto:myExit

if /i "%DEBUG%" equ "Yes" (@echo myExitError - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- getNameNoExt
::--------------------------------------------------------
:getNameNoExt
    set "%~1=%~n0"
goto:eof

::---------------------------------------------------------
::-- calculateLength
::---------------------------------------------------------
:calculateLength
    (echo "%~1" & echo.) | findstr /O . | more +1 | (set /p result= & call exit /b %%result%%)
    set /a %2=%errorlevel%-4
goto:eof

::---------------------------------------------------------
::-- deQuote
::---------------------------------------------------------
:deQuote
    for /f "delims=" %%a in ('echo %%%1%%') do set %2=%%~a
goto:eof

::--------------------------------------------------------
::-- getRoot or set lRoot=%~dp0
::--------------------------------------------------------
:getRoot
    set lRoot=%cd%
    call:removeSlash "%lRoot%" lRoot
    set "%~1=%lRoot%"
goto:eof

::--------------------------------------------------------
::-- removeSlash
::--------------------------------------------------------
:removeSlash
    set var=%~1
    if %var:~-1%==\ (set %2=%var:~0,-1%)
goto:eof

::--------------------------------------------------------
::-- getFirstXChars
::--------------------------------------------------------
:getFirstXChars
    set sValIn=%~1
    set /a iNo=%2
    set vWorkVal=%%sValIn:~0,%iNo%%%
    call:getFirstX %vWorkVal% vWorkVal2
    set %3=%vWorkVal2%

:getFirstX
    set %2=%~1
goto:eof

::--------------------------------------------------------
::-- Logging
::--------------------------------------------------------
:Logging
    @echo off
    if /i "%DEBUG%" equ "Yes" (set DEBUG_Remember=%DEBUG%)

    set sMsg=%~1
    set sMessage=%computername% %date% %time% := %sMsg%
    @echo %sMessage%
    if "%Q_ALLOW_LOGFILE%" equ "Yes" (@echo %sMessage% >> "%LOGFILE%")

    set DEBUG=%DEBUG_Remember%
    if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)
goto:eof

::--------------------------------------------------------
:: myExit
::--------------------------------------------------------
:myExit
if /i "%DEBUG%" equ "Yes" (@echo myExit - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call:Logging "%SEPERATOR2%"
    call:Logging "Complete ..."
    call:Logging "%SEPERATOR2%"

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        pause >nul
    ) else (
        timeout /t %iTimeOut% >nul
    )
    endlocal
    endlocal

if /i "%DEBUG%" equ "Yes" (@echo myExit - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
:: Revision History:
::-----------+------------+----------+--------------------
:: Modified  | Date       | Ver      | Reason
::-----------+------------+----------+--------------------
:: G Bishop  | 29-04-2020 | 01.01.01 | initial dev
:: G Bishop  | 14-08-2020 | 01.02.00 | Amended: ExcelFiles will now move TTD back
:: G Bishop  | 18-08-2020 | 00.00.00 | Amended: will place icon on root
::-----------+------------+----------+--------------------
:: To-Do: Get this checked into git
::--------------------------------------------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
