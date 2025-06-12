::--------------------------------------------------------
::      Author: G Bishop
::        Date: 4th November 2024
::        Name: measureDuration.cmd
:: Description: Measures time taken to run a process in milliseconds
::     History: See end of file
::--------------------------------------------------------
@echo off

if [%1] equ [] (set Q_ARCHIVE=Archive) else (set Q_ARCHIVE=%~1)

set VERSION=01.03.00

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

setlocal
setlocal EnableExtensions EnableDelayedExpansion

set Q_ALLOW_PAUSE=Yes

call :setVars
call :setScreen
call :echoVersion

call :durationTiming Start

call :Logging "Running process, please wait ..."

if /i "%Q_ARCHIVE%" equ "Archive"    (call ArchiveFiles.vbs.lnk)
if /i "%Q_ARCHIVE%" equ "Unarchive"  (call ArchiveFilesDe.vbs.lnk)
if /i "%Q_ARCHIVE%" equ "Duplicates" (call ArchiveFilesDup.vbs.lnk)

call :durationTiming Stop

:: pause
goto:myExit


::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
    set SEPERATOR1=----------------------------------------
    set SEPERATOR2=========================================
goto:eof

::--------------------------------------------------------
::-- setScreen
::--------------------------------------------------------
:setScreen
    set /a iWidth=60
    set /a iHeight=10
    mode con: cols=%iWidth% lines=%iHeight%
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
    :: call :getFileNameThis sThisName
    :: set strEcho=Version ~ %VERSION% - %sThisName%
    set strEcho=Version ~ %VERSION% - Archiving is:
    set strEcho=!strEcho! %Q_ARCHIVE%

    Title %strEcho%
    call :Logging "%SEPERATOR2%"
    call :Logging "%strEcho%"
    call :Logging "%SEPERATOR1%"
goto:eof

::---------------------------------------------------------
::-- deQuote
::---------------------------------------------------------
:deQuote
    for /f "delims=" %%a in ('echo %%%1%%') do set %2=%%~a
goto:eof

::---------------------------------------------------------
::-- durationTiming
::---------------------------------------------------------
:durationTiming
    set sDoWhat=%1

    if /i "%sDoWhat%" equ "Start" (set iTimeStart=%time%)

    if /i "%sDoWhat%" equ "Stop" (
        set iTimeEnd=!time!
        call :getDuration iTimeStart iTimeEnd iDuration
        call :formatThis !iDuration! sDuration
        call :Logging "Execution took ~ !sDuration! milliseconds."
    )

goto:eof

::---------------------------------------------------------
::-- getDuration
::---------------------------------------------------------
:getDuration        ' Calculate difference
    call :getMSeconds %1 iTimeStartMs
    call :getMSeconds %2 iTimeEndMs
    set /a %3=%iTimeEndMs%-%iTimeStartMs%
goto:eof

::---------------------------------------------------------
::-- getMSeconds
::---------------------------------------------------------
:getMSeconds
    call :parseTime      %1              TimeAsArgs
    call :calcMSeconds   %TimeAsArgs%    %2
goto:eof

::---------------------------------------------------------
::-- calcMSeconds
::---------------------------------------------------------
:calcMSeconds
    set /a iHour=%1
    set /a iMin=%2
    set /a iSec=%3
    set /a iMSec=%4
    set /a iHour*=(3600*1000)
    set /a iMin*=(60*1000)
    set /a iSec*=(1000)
    set /a %5 = (%iHour%) + (%iMin%) + (%iSec%) + (%iMSec%)
goto:eof

::---------------------------------------------------------
::-- parseTimeSwapped
::---------------------------------------------------------
:parseTime      ' Mask time like " 0:23:29,12"
    set %2=!%1: 0=0!
    set %2=!%2::= !
    set %2=!%2:.= !
    set %2=!%2:,= !
    set %2=!%2: 0= !
goto:eof

::--------------------------------------------------------
::-- formatThis
::--------------------------------------------------------
:formatThis
    set "sIn1=%1"
    set "sTemp="
    set "sSign="
    set /a iNoLoops=7
    if "%sIn1:~0,1%" equ "-" set "sSign=-" & set "sIn1=%sIn1:~1%"
    for /L %%i in (1,1,%iNoLoops%) do if defined sIn1 (
       set "sTemp=,!sIn1:~-3!!sTemp!"
       set "sIn1=!sIn1:~0,-3!"
    )
    set "%2=%sSign%%sTemp:~1%
goto:eof

::--------------------------------------------------------
::-- getFileNameThis
::--------------------------------------------------------
:getFileNameThis
    set "%1=%~nx0"
goto:eof

::--------------------------------------------------------
::-- Logging
::--------------------------------------------------------
:Logging
    set sMsg=%~1
    set sMessage=%time% := %sMsg%
    @echo %sMessage%
goto:eof

::--------------------------------------------------------
::-- myExit
::--------------------------------------------------------
:myExit
if /i "%DEBUG%" equ "Yes" (@echo myExit - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call :Logging "%SEPERATOR1%"
    call :Logging "Complete ..."
    call :Logging "%SEPERATOR2%"
    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo myExit - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
:: Revision History:
::-----------+------------+----------+--------------------
:: Modified  | Date       | Ver      | Reason
::-----------+------------+----------+--------------------
:: G Bishop  | 04-11-2024 | 01.00.00 | Initial dev
:: G Bishop  | 05-11-2024 | 01.00.01 | Renamed: measureDuration.cmd
:: G Bishop  | 05-11-2024 | 01.02.00 | Added: Parameter: Q_ARCHIVE
:: G Bishop  | 06-12-2024 | 01.03.00 | Amended: Parameter: Q_ARCHIVE now does 3 things
::-----------+------------+----------+--------------------
:: ToDo: Add command line parameters
::-----------+------------+----------+--------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
