set DEBUG=Yes
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

@echo off

setlocal EnableDelayedExpansion

set VERSION=01.01.00

call :Main
goto:myExit

::--------------------------------------------------------
::-- Main
::--------------------------------------------------------
:Main
if /i "%DEBUG%" equ "Yes" (@echo Main - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call :getNameNoExt sCurrFile
    set sCurrFile=%sCurrFile%.txt
    
    @echo sCurrFile %sCurrFile%
    @echo sCurrFile %sCurrFile%
    @echo sCurrFile %sCurrFile%
    pause
    

    for /f "tokens=* delims= " %%f in (%sCurrFile%) do (
        call:autoMateWorker "%%f"
    )

if /i "%DEBUG%" equ "Yes" (@echo Main - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- autoMateWorker
::--------------------------------------------------------
:autoMateWorker
if /i "%DEBUG%" equ "Yes" (@echo autoMateWorker - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sTARGETFolder=%~1

    if /i "%DEBUG%" equ "Yes" (@echo    sTARGETFolder %sTARGETFolder%)
    if /i "%DEBUG%" equ "Yes" (pause)

    pushd "%sTARGETFolder%"
    if /i "%DEBUG%" equ "Yes" (pause)
    
    popd
    if /i "%DEBUG%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo autoMateWorker - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof


::--------------------------------------------------------
::-- getNameNoExt
::--------------------------------------------------------
:getNameNoExt
    set "%~1=%~n0"
goto:eof

::--------------------------------------------------------
::-- myExit
::--------------------------------------------------------
:myExit
if /i "%DEBUG%" equ "Yes" (@echo myExit - Start)
if /i "%DEBUG%" equ "Yes" (pause)
    @echo bye
    pause

    endlocal
    exit

if /i "%DEBUG%" equ "Yes" (@echo myExit - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof