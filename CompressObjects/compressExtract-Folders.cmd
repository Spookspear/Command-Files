::--------------------------------------------------------
::        Author: G Bishop
::        Date: 10:53 13 December 2024
::        Name: compressExtract-Folders.cmd
:: Description: Performs specified actions on folders
::     History: See end of file
::--------------------------------------------------------
@echo off
if [%1] equ [] (set Q_COMMAND=Extract) else (set Q_COMMAND=%~1)

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

set VERSION=01.02.00

set Q_ALLOW_PAUSE=No
set Q_DELETE_EXTRACTER=No


:: Finer control
:: set Q_COMMAND=compressorDeploy
:: set Q_COMMAND=compressorExecute

:: set Q_COMMAND=extractorDeploy
:: set Q_COMMAND=extractorExecute

if /i "%Q_COMMAND%" equ "Extract" (set Q_COMMAND=extractorDeployAndExecute)
if /i "%Q_COMMAND%" equ "Compress" (set Q_COMMAND=compressorDeployAndExecute)

:: set Q_COMMAND=compressorDeployAndExecute
:: set Q_COMMAND=extractorDeployAndExecute

call :Main
goto :myExit

::--------------------------------------------------------
::-- Main
::--------------------------------------------------------
:Main
if /i "%DEBUG%" equ "Yes" (@echo Main - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call :setVars
    call :echoVersion

    for /f "delims=" %%i in ('dir /b /A:d') do (
        call :%Q_COMMAND% "%%i"
    )

if /i "%DEBUG%" equ "Yes" (@echo Main - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
if /i "%DEBUG%" equ "Yes" (@echo setVars - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set /a iTimeout=2
    set sLocationCompressor=%USERPROFILE%\Documents\Development\Command Files\CompressExtract\Compress\7Z-EmailSendSplit.cmd.lnk
    set sLocationExtractor=%USERPROFILE%\Documents\Development\Command Files\CompressExtract\Extract\extractCompressed.cmd.lnk

    set SEPERATOR1=ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ?
    set SEPERATOR2=ÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ?

if /i "%DEBUG%" equ "Yes" (@echo setVars - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- compressorDeployAndExecute
::--------------------------------------------------------
:compressorDeployAndExecute
if /i "%DEBUG%" equ "Yes" (@echo compressorDeployAndExecute - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFolder=%~1
    call :Logging "compressorDeployAndExecute - %sFolder%"

    if not exist "%sFolder%\7Z-EmailSendSplit.cmd.lnk" (
        call :compressorDeploy "%sFolder%"
    )

    call :compressorExecute "%sFolder%"
    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo compressorDeployAndExecute - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- compressorDeploy
::--------------------------------------------------------
:compressorDeploy
if /i "%DEBUG%" equ "Yes" (@echo compressorDeploy - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFolder=%~1

    call :Logging "%SEPERATOR2%"
    call :Logging "compressorDeploy - %sFolder%"
    call :Logging "%SEPERATOR1%"

    copy "%sLocationCompressor%"  "%sFolder%"

if /i "%DEBUG%" equ "Yes" (@echo compressorDeploy - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- compressorExecute
::--------------------------------------------------------
:compressorExecute
if /i "%DEBUG%" equ "Yes" (@echo compressorExecute - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFolder=%~1
    call :Logging "%SEPERATOR2%"
    call :Logging "compressorExecute - %sFolder%"
    call :Logging "%SEPERATOR1%"

    pushd "%sFolder%"
    7Z-EmailSendSplit.cmd.lnk
    popd

if /i "%DEBUG%" equ "Yes" (@echo compressorExecute - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- extractorDeployAndExecute
::--------------------------------------------------------
:extractorDeployAndExecute
if /i "%DEBUG%" equ "Yes" (@echo extractorDeployAndExecute - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFolder=%~1
    call :Logging "extractorDeployAndExecute - %sFolder%"

    if not exist "%sFolder%\extractCompressed.cmd.lnk" (
        call :extractorDeploy "%sFolder%"
    )

    call :extractorExecute "%sFolder%"
    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo extractorDeployAndExecute - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- extractorDeploy
::--------------------------------------------------------
:extractorDeploy
if /i "%DEBUG%" equ "Yes" (@echo extractorDeploy - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFolder=%~1

    call :Logging "%SEPERATOR2%"
    call :Logging "extractorDeploy - %sFolder%"
    call :Logging "%SEPERATOR1%"

    copy "%sLocationExtractor%"  "%sFolder%" >nul:

if /i "%DEBUG%" equ "Yes" (@echo extractorDeploy - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- extractorExecute
::--------------------------------------------------------
:extractorExecute
if /i "%DEBUG%" equ "Yes" (@echo extractorExecute - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFolder=%~1
    call :Logging "%SEPERATOR2%"
    call :Logging "extractorExecute - %sFolder%"
    call :Logging "%SEPERATOR1%"

    pushd "%sFolder%"
    extractCompressed.cmd.lnk
    popd

if /i "%DEBUG%" equ "Yes" (@echo extractorExecute - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
if /i "%DEBUG%" equ "Yes" (@echo echoVersion - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set strEcho=VERSION: %VERSION%
    Title %strEcho%

    call:Logging "%SEPERATOR2%"
    call:Logging "%strEcho%"
    call:Logging "Q_COMMAND is %Q_COMMAND%"
    call:Logging "%Q_COMMAND%"
    call:Logging "%SEPERATOR1%"

if /i "%DEBUG%" equ "Yes" (@echo echoVersion - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto :eof

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
    @echo %date% %time% := %~1
goto:eof

::--------------------------------------------------------
::-- myExit
::--------------------------------------------------------
:myExit
if /i "%DEBUG%" equ "Yes" (@echo myExit - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sMessage=goodbye ...

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        set sMessage=%sMessage% press any key to close window
    ) else (
        set sMessage=%sMessage% closed in %iTimeout%
    )

    if "%Q_DELETE_EXTRACTER%" equ "Yes" (
        set sMessage=%sMessage% Caller to be deleted
    )
    call:Logging "%sMessage% ..."

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        pause
    ) else (
        if "!Q_DELETE_EXTRACTER!" equ "Yes" (
            timeout /T !iTimeout!
        ) else (
            timeout /T !iTimeout! >nul
        )
    )

    if "%Q_DELETE_EXTRACTER%" equ "Yes" (
        call:getFileNameThis sThisName
        set sThisName=!sThisName!.lnk
        del !sThisName! /q
    )

    endlocal
    endlocal

    :: exit

if /i "%DEBUG%" equ "Yes" (@echo myExit - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
:: Revision History:
::-----------+------------+----------+--------------------
:: Modified  | Date       | Ver      | Reason
::-----------+------------+----------+--------------------
:: G Bishop  | 12-12-2024 | 01.00.00 | Initial Dev.
:: G Bishop  | 13-12-2024 | 01.01.00 | Added: more code
:: G Bishop  | 20-03-2025 | 01.02.00 | Amended: Pass in parameters of what to do
::-----------+------------+----------+--------------------
:: To-Do:
::--------------------------------------------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
