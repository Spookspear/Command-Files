::--------------------------------------------------------
::      Author: G Bishop
::        Date: 26 September 2019
::        Name: cleanDownFolders.cmd
:: Description: Deletes files and folders in supplied text file
::     History: See end of file
::--------------------------------------------------------
@echo off

setlocal EnableDelayedExpansion

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

set VERSION=01.01.00

if /i "%cd%" equ "C:\Windows\System32" (
    cd "C:\Users\grant.bishop\OneDrive - Harbour Energy plc\Documents\Development\Command Files\cleanFolder\VMVersion"
)

call :getNameNoExt G_INFILE
set G_INFILE=!G_INFILE!
set G_INFILE=!G_INFILE!.txt

if /i "%DEBUG%" equ "Yes" (@echo G_INFILE %G_INFILE%)
if /i "%DEBUG%" equ "Yes" (pause)

set G_ALLOW_PAUSE=No
set G_TAKEOWNERSHIP=Yes
set G_DELETEFILES=Yes
:: set G_METHOD=Fast
set G_METHOD=Slow

:: Remove the top folder
set G_REMOVEROOT=No

call :Main
call :myExit Yes

::--------------------------------------------------------
::-- Main
::--------------------------------------------------------
:Main
if /i "%DEBUG%" equ "Yes" (@echo Main - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call :setScreen
    call :preSetup
    call :sayStuff

    if /i "%DEBUG%" equ "Yes" (@echo G_INFILE %G_INFILE%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i %G_TAKEOWNERSHIP% equ Yes (call :takeOwnershipCaller "%G_INFILE%")
    if /i %G_DELETEFILES% equ Yes (call :cleanDownFoldersCaller "%G_INFILE%")

if /i "%DEBUG%" equ "Yes" (@echo Main - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- sayStuff
::--------------------------------------------------------
:sayStuff
    :: Used for displaying
    call :deQuote G_INFILE _G_INFILE

    call :Logging "%SEPERATOR1%"
    call :Logging " Clearing folders defined in %_G_INFILE%"
    call :Logging "   Delete files: %G_DELETEFILES%"
    call :Logging " Take Ownership: %G_TAKEOWNERSHIP%"
    call :Logging "    Allow Pause: %G_ALLOW_PAUSE%"
    call :Logging "  Delete Method: %G_METHOD%"
    call :Logging "    Remove Root: %G_REMOVEROOT%"
    call :Logging "%SEPERATOR1%"
    pause
goto:eof

::--------------------------------------------------------
::-- preSetup
::--------------------------------------------------------
:preSetup
    set iTimeOut=10
goto:eof

::--------------------------------------------------------
::-- takeOwnershipCaller - not tested
::--------------------------------------------------------
:takeOwnershipCaller
if /i "%DEBUG%" equ "Yes" (@echo takeOwnershipCaller - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sCurrFile=%~1

    if /i "%DEBUG%" equ "Yes" (@echo sCurrFile %sCurrFile%)
    if /i "%DEBUG%" equ "Yes" (pause)

    for /f "tokens=* delims= " %%f in (%sCurrFile%) do (
        call :takeOwnership "%%f"
    )

if /i "%DEBUG%" equ "Yes" (@echo takeOwnershipCaller - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- take ownership
::--------------------------------------------------------
:takeOwnership
if /i "%DEBUG%" equ "Yes" (@echo takeOwnership - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sPath=%~1

    if /i "%DEBUG%" equ "Yes" (@echo sPath %sPath%)
    if /i "%DEBUG%" equ "Yes" (pause)

    for /d %%i in (%sPath%\*) do (Takeown /r /f "%%i")

if /i "%DEBUG%" equ "Yes" (@echo takeOwnership - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- cleanDownFoldersCaller
::--------------------------------------------------------
:cleanDownFoldersCaller
if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersCaller - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sCurrFile=%~1

    if /i "%DEBUG%" equ "Yes" (@echo sCurrFile %sCurrFile%)
    if /i "%DEBUG%" equ "Yes" (pause)

    for /f "tokens=* delims= " %%f in (%sCurrFile%) do (
        if /i %G_METHOD% equ Slow (
            call :cleanDownFoldersSlow "%%f"
        )

        if /i %G_METHOD% equ Fast (
            call :cleanDownFoldersFast "%%f"
        )
    )


    REM for /f "tokens=* delims= " %%f in (%~1) do (
        REM if /i %G_METHOD% equ Slow (
            REM call :cleanDownFoldersSlow "%%f"
        REM )

        REM if /i %G_METHOD% equ Fast (
            REM call :cleanDownFoldersFast "%%f"
        REM )
    REM )

if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersCaller - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- Clean down folders Slow
::--------------------------------------------------------
:cleanDownFoldersSlow
if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersSlow - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sPath=%~1
    call :Logging "%sPath%"
    if /i %G_ALLOW_PAUSE% equ Yes (pause)

    if exist "%sPath%" (
        call :Logging "Removing attributes"
        for /r "%sPath%\" %%g in (*.*) do (attrib -h -r -s "%%g")
        if /i "%DEBUG%" equ "Yes" (pause)

        call :Logging "Deleting files"
        for /r "%sPath%\" %%g in (*.*) do (del /f /q "%%g")
        if /i "%DEBUG%" equ "Yes" (pause)

        call :Logging "Removing Folders"
        for /d %%i in ("%sPath%\*") do (rd /s /q "%%i")
        if /i "%DEBUG%" equ "Yes" (pause)

        if /i %G_REMOVEROOT% equ Yes (
            call :Logging "Removing Top Folder"
            rd /s /q "%sPath%"
        )
        if /i "%DEBUG%" equ "Yes" (pause)
    )

if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersSlow - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- Clean down folders Fast
::--------------------------------------------------------
:cleanDownFoldersFast
if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersFast - Start)
if /i "%DEBUG%" equ "Yes" (pause)


    set sPath=%~1
    call :Logging "%sPath%"
    if /i %G_ALLOW_PAUSE% equ Yes (pause)

    if exist "%sPath%" (

        call :Logging "Removing attributes"
        attrib -h -r -s /s "%sPath%\*.*"

        call :Logging "Removing Files"
        del /f /q /s "%sPath%\*" >nul

        call :Logging "Removing folders"
        for /d %%i in ("%sPath%\*") do rd /s /q "%%i"

        del "%sPath%\*.*" /q
        call :Logging "Removing *.* from top folder"

        if /i %G_REMOVEROOT% equ Yes (
            call :Logging "Removing Top Folder"
            rd /s /q "%sPath%"
        )
    )


if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersFast - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setScreen
::--------------------------------------------------------
:setScreen
if /i "%DEBUG%" equ "Yes" (@echo setScreen - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sTitle=Handling Folders - VERSION %VERSION%
    Title %sTitle%

    call :calculateHeight %G_METHOD% iWidth iHeight iColor
    call :setColour %iColor%
    :: mode con: cols=%iWidth% lines=%iHeight%
    if /i "%DEBUG%" neq "Yes" (mode con: cols=%iWidth% lines=%iHeight%)

    call :replicate %iWidth% "-" SEPERATOR1
    call :replicate %iWidth% "=" SEPERATOR2

    call :Logging "%SEPERATOR2%"
    call :Logging "%sTitle%"
    call :Logging "%SEPERATOR2%"

if /i "%DEBUG%" equ "Yes" (@echo setScreen - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- calculateHeight default is for ukn6441
::--------------------------------------------------------
:calculateHeight
    set lPC=%1
    set /a iW=105
    set /a iH=60
    set /a iBuff=0
    set iC=Info

    if /i %lPC% equ Slow (
        set /a iW=232
        set /a iH=112
    )

    call :getDrive sDrive

    if /i %sDrive% equ G: (set /a iBuff=38)

    set /a %~2=%iW%+%iBuff%
    set /a %~3=%iH%
    set %~4=%iC%
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
::-- setColour
::--------------------------------------------------------
:setColour
    if /i %1 equ Error (color cf
    ) else (if /i %1 equ Warning (color 5e
    ) else (if /i %1 equ Info    (color 1f
    ) else (if /i %1 equ Debug   (color 3e
    ) else (color %1))))
goto:eof

::--------------------------------------------------------
::-- getDrive
::--------------------------------------------------------
:getDrive
    set %~1=%~d0
goto:eof

::---------------------------------------------------------
::-- deQuote
::---------------------------------------------------------
:deQuote
    for /f "delims=" %%a in ('echo %%%1%%') do set %2=%%~a
goto:eof

::--------------------------------------------------------
::-- getNameNoExt
::--------------------------------------------------------
:getNameNoExt
    set "%~1=%~n0"
goto:eof

::--------------------------------------------------------
::-- Logging
::--------------------------------------------------------
:Logging
    @echo %~1
goto:eof

::--------------------------------------------------------
::-- myExit
::--------------------------------------------------------
:myExit
if /i "%DEBUG%" equ "Yes" (@echo myExit - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if [%1]==[] (set sayByeBye=Yes) else (set sayByeBye=%1)
    if /i %sayByeBye%==Yes (
        call :Logging "%SEPERATOR2%"
        call :Logging "Done ..."
        call :Logging "%SEPERATOR2%"
    ) else (
        call :Logging "Errors?"
    )
    timeout /T %iTimeOut%
    if /i %G_ALLOW_PAUSE% equ Yes (pause)
    endlocal
    exit

if /i "%DEBUG%" equ "Yes" (@echo myExit - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
:: Revision History:
::-----------+------------+----------+--------------------
:: Modified  | Date       | Ver      | Reason
::-----------+------------+----------+--------------------
:: G Bishop  | 26/09/2019 | 01.00.00 | Created
:: G Bishop  | 07/10/2019 | 01.00.01 | renamed input file to: ListOfFolders.txt
:: G Bishop  | 11/10/2019 | 01.00.02 | Added in G_TAKEOWNERSHIP - call :takeOwnershipCaller(filename)
:: G Bishop  | 11/10/2019 | 01.00.03 | G_DELETEFILES  - call :cleanDownFoldersCaller(filename)
:: G Bishop  | 01/11/2019 | 01.00.04 | Continue if any one folder doesnt exist
:: G Bishop  | 29-08-2024 | 01.01.00 | Amended: Call convention, made similar to other scripts
::-----------+------------+----------+--------------------
:: To-Do
:: test takeOwnershipCaller
::--------------------------------------------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
