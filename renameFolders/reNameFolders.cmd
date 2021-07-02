::--------------------------------------------------------
::      Author: G Bishop
::        Date: 24-08-2020
::        Name: reNameFoldersNew.cmd
:: Description: renames eCas folders to have the vessel number and name
::     History: See end of file
::--------------------------------------------------------
@echo off

setlocal
setlocal EnableExtensions EnableDelayedExpansion

if [%~1] equ [] (set sAllowResize=Yes)    else (set sAllowResize=%~1)

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

set VERSION=01.02.00

set Q_ALLOW_PAUSE=No
set Q_ALLOW_LOGFILE=No

call:Main
goto:myExit

::--------------------------------------------------------
::-- Main
::--------------------------------------------------------
:Main
    call:preSetup
    call:doWork
goto:eof

::--------------------------------------------------------
::-- doWork
::--------------------------------------------------------
:doWork
    for /D %%G in (*) do (
        call:workerProcessCaller "%%G"
    )
goto:eof

::--------------------------------------------------------
::-- getVesselID
::--------------------------------------------------------
:getVesselID
    call:getCallingFolderName sCallingFolder
    call:getFirstXChars "%sCallingFolder%" 3 sVesselNo_A
    call:getFirstXChars "%sCallingFolder%" 4 sVesselNo_B

    set sVesselID=%sVesselNo_B%
    if /i %sVesselNo_A% equ %sVesselNo_B% (set sVesselID=%sVesselNo_A%)

    call:calculateLength "%sCallingFolder%" iCallingFolderLength
    set /a iCallingFolderLength-=7
    call:getLastXChars "%sCallingFolder%" %iCallingFolderLength% sCallingFolderDescOnly

    set %~1=%sVesselID% %sCallingFolderDescOnly%
goto:eof

::--------------------------------------------------------
::-- workerProcessCaller
::--------------------------------------------------------
:workerProcessCaller
    set sFolder=%~1

    call:checkOkToGo "%sFolder%" sIsOK

    if /i "%sIsOK%" equ "Yes" (
        call:workerProcess "!sFolder!"
    ) else (
        set sAnyErrors=Yes
        call:Logging "Problem with folder name: %sFolder%"
        if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)
    )
goto:eof

::--------------------------------------------------------
::-- workerProcess
::--------------------------------------------------------
:workerProcess
    :: get first 12 chars for incident name
    call:getFirstXChars "%sFolder%" 12 sIncidentNo

    :: so will need to measure string and use that number for the number of right chars to get
    call:calculateLength "%sFolder%" iFolderLength

    :: subtract incident number and ' - eCas - ' equals:23
    set /a iFolderLength-=23

    :: getting the last x chars leaving folder description
    call:getLastXChars "%sFolder%" %iFolderLength% sFolderDescOnly

    :: concatenate the incident, vessel and descripiton to return
    set sFolderNewName=%sIncidentNo% - %sVessel% - %sFolderDescOnly%

    if not exist "%sFolderNewName%" (
        call:Logging "ren '%sFolder%' '%sFolderNewName%'"
        if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)
        ren "%sFolder%" "%sFolderNewName%"
        if %errorlevel% neq 0 (set sAnyErrors=Yes)
    ) else (
        call:Logging "Folder:'%sFolderNewName%' Already exists"
        if %errorlevel% neq 0 (set sAnyErrors=Yes)
    )
goto:eof

::--------------------------------------------------------
::-- checkOkToGo
::--------------------------------------------------------
:checkOkToGo
    set sTemp=%~1
    set sRetVal=No
    call:getMidXChars "%sTemp%" 13 10 sTemp1
    if /i "%sTemp1%" equ " - eCas - " (set sRetVal=Yes)
    set %~2=%sRetVal%
goto:eof

::--------------------------------------------------------
::-- preSetup
::--------------------------------------------------------
:preSetup
    call:setVars
    call:setScreen
    call:echoVersion
goto:eof

::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
    if not defined Q_ALLOW_PAUSE         (set Q_ALLOW_PAUSE=Yes)
    if not defined Q_ALLOW_LOGFILE       (set Q_ALLOW_LOGFILE=No)
    if not defined Q_ALLOW_SPECIAL_CHARS (set Q_ALLOW_SPECIAL_CHARS=Yes)

    set /a iTimeOut=3
    set sSeperatorChar1=-
    set sSeperatorChar2==
    set sAnyErrors=No

    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar1=Ä)
    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar2=Í)
    call:getRoot myRoot

    call:getVesselID sVessel

    if "%Q_ALLOW_LOGFILE%" equ "Yes" (
        set LOGFOLDER=C:\Bin\Logs
        call:setLogFile "%myRoot%" LOGFILE
    )

    call:setEchoData

goto:eof

::---------------------------------------------------------
::-- setLogFile
::---------------------------------------------------------
:setLogFile
    if not defined sCurrFile (call:getNameNoExt sCurrFile)

    for /F "tokens=1-4 delims=/- " %%A in ('date/T') do set myFileDate=%%A%%B
    for /F "tokens=1,2* delims=: " %%i in ('time/T') do set myFileTime=%%i%%j%%k
    set "%~2=%~1\%sCurrFile%-%myFileDate%-%MyFileTime%.Log"
goto:eof

::--------------------------------------------------------
::-- setScreen
::--------------------------------------------------------
:setScreen
    call:calculateHeight %COMPUTERNAME% iWidth iHeight iColor

    if /i "%sAllowResize%" equ "Yes" (
        if /i "%DEBUG%" neq "Yes" (
            mode con: cols=%iWidth% lines=%iHeight%
        )
    )

    set /a iWidth-=iSubtractWidth
    call:Replicate %iWidth% "%sSeperatorChar1%" SEPERATOR1
    call:Replicate %iWidth% "%sSeperatorChar2%" SEPERATOR2
    call:setColour %iColor%
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
    if not defined strEcho (set strEcho=%sCurrFile% - %sVessel% - Version: %VERSION%)

    Title %strEcho%
    call:Logging "%SEPERATOR2%"
    call:Logging "%strEcho%"
    call:Logging "%SEPERATOR1%"
goto:eof

::--------------------------------------------------------
::-- setEchoData
::--------------------------------------------------------
:setEchoData
    if not defined sCurrFile (call:getNameNoExt sCurrFile)

    set strEcho=%sCurrFile% - %sVessel% - Version: %VERSION%
    set sSubtractMsg=%computername% %date% %time% :=

    set strEchoMeasure=%sSubtractMsg% %strEcho%

    call:calculateLength "%sVessel%"   iVesselWidth
    call:measureFoldersLoopCaller iW_Buff iH_Buff
    call:calculateLength "%sSubtractMsg%"   iSubtractWidth
    call:calculateLength "%strEchoMeasure%" iEchoWidth

    if %iEchoWidth% gtr iW_Buff (set iW_Buff=%iEchoWidth%)
goto:eof

::--------------------------------------------------------
::-- calculateHeight
::--------------------------------------------------------
:calculateHeight
    if not defined iW_Buff  (set /a iW_Buff=0)
    if not defined iH_Buff  (set /a iH_Buff=0)

    set /a iW=1
    set /a iH=7
    set iC=1b

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (set /a iH_Buff+=1)

    set /a %~2=%iW%+%iW_Buff%
    set /a %~3=%iH%+%iH_Buff%
    set %~4=%iC%
goto:eof

::--------------------------------------------------------
::-- setColour
::--------------------------------------------------------
:setColour
    if /i "%~1" equ "Error" (color cf
    ) else (if /i "%~1" equ "Warning" (color 5e
    ) else (if /i "%~1" equ "Info"    (color 1f
    ) else (if /i "%~1" equ "Debug"   (color 3e
    ) else (if /i "%~1" equ "Good"    (color 2f
    ) else (if /i "%~1" equ "Present" (color 9f
    ) else (color %~1))))))
goto:eof

::--------------------------------------------------------
::-- Replicate
::--------------------------------------------------------
:Replicate
    set /a iLoop=%~1-1
    set repeatChar=%~2
    set returnStr=
    for /L %%G in (1,1,!iLoop!) do (set "returnStr=!returnStr!!repeatChar!")
    set %~3=%returnStr%
goto:eof

::--------------------------------------------------------
::-- myExitError
::--------------------------------------------------------
:myExitError
    set Q_ALLOW_PAUSE=Yes
    call:setColour Error
    call:Logging "%~1"
    goto:myExit
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
    set vWorkVal="%%sValIn:~0,%iNo%%%"
    call:getFirstX %vWorkVal% vWorkVal2
    call:deQuote vWorkVal2 vWorkVal2
    set %3=%vWorkVal2%

:getFirstX
    set %~2=%~1
goto:eof

::--------------------------------------------------------
::-- getLastXChars
::--------------------------------------------------------
:getLastXChars
    set sValIn=%~1
    set /a iNo=%2
    set vWorkVal="%%sValIn:~-%iNo%%%"
    call:getLastX %vWorkVal% vWorkVal2
    call:deQuote vWorkVal2 vWorkVal2
    set %3=%vWorkVal2%

:getLastX
    set %~2=%1
goto:eof

::--------------------------------------------------------
::-- getMidXChars
::--------------------------------------------------------
:getMidXChars
    set sValIn=%~1
    set /a iStart=%~2
    set /a iLength=%~3
    call:calculateLength "%sValIn%" iWorkValLength
    set /a iWorkValLength-=%iStart%
    call:getLastXChars "%sValIn%" %iWorkValLength% vWorkVal
    call:getFirstXChars "%vWorkVal%" %iLength% vRetVal
    set %~4=%vRetVal%
goto:eof

::--------------------------------------------------------
::-- getCallingFolderName
::--------------------------------------------------------
:getCallingFolderName
    for %%I in (.) do set fldr=%%~nxI
    set "%~1=%fldr%"
goto:eof

::--------------------------------------------------------
::-- :measureFoldersLoopCaller - get the longest
::--------------------------------------------------------
:measureFoldersLoopCaller
    set /a iLongest=0
    set /a iLineLen=0
    set /a iHeightCount=0

    for /D %%i in (*) do (
        call:measureFoldersLoop "%%i" iLineLen
        if !iLineLen! gtr !iLongest! (set /a iLongest=!iLineLen!)
        set /a iHeightCount+=1
    )

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (set /a iHeightCount*=2)

    set /a %1=!iLongest!
    set /a %2=!iHeightCount!
goto:eof

::--------------------------------------------------------
::-- measureFoldersLoop - get the longest
::--------------------------------------------------------
:measureFoldersLoop
    set varPath=%1
    set /a iLocalLineLen=0
    call:deQuote varPath _varPath
    set sMeasureThis="%sSubtractMsg% ren '%_varPath%' '%_varPath%'"
    call:calculateLength %sMeasureThis% iLocalLineLen
    set /a iLocalLineLen-=4
    set /a iLocalLineLen+=iVesselWidth
    set /a iLocalLineLen-=2
    set /a %2=!iLocalLineLen!
goto:eof

::--------------------------------------------------------
::-- Logging
::--------------------------------------------------------
:Logging
    set sMsg=%~1
    set sMessage=%computername% %date% %time% := %sMsg%
    @echo %sMessage%
    if "%Q_ALLOW_LOGFILE%" equ "Yes" (@echo %sMessage% >> "%LOGFILE%")
goto:eof

::--------------------------------------------------------
:: myExit
::--------------------------------------------------------
:myExit
    call:Logging "%SEPERATOR2%"
    call:Logging "Complete - any errors: %sAnyErrors%"
    call:Logging "%SEPERATOR2%"

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        pause >nul
    ) else (
        timeout /t %iTimeOut% >nul
    )

    if /i "%sAnyErrors%" equ "No" (del reNameFolders.cmd.lnk)

    endlocal
    endlocal

goto:eof

::--------------------------------------------------------
::-- Revision History:
::-------------+------------+----------+------------------
:: Modified    | Date       | Ver      | Reason
::-------------+------------+----------+------------------
:: G Bishop    | 24-08-2020 | 01.00.00 | Created
:: G Bishop    | 27-08-2020 | 01.01.00 | Added: getMidXChars() and dequoted getFirstXChars() & :getLastXChars()
:: G Bishop    | 27-08-2020 | 01.02.00 | Amended: will pass in sAllowResize
::-------------+------------+----------+------------------
:: To-Do
::--------------------------------------------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
