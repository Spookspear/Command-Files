::--------------------------------------------------------
::      Author: G Bishop
::        Date: 24th January 2012
::        Name: extractCompressed.cmd
:: Description: Extracts all compressed files
::     History: See end of file
::--------------------------------------------------------
@echo off

setlocal
setlocal EnableExtensions EnableDelayedExpansion

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

set VERSION=02.03.00
set Q_ALLOW_PAUSE=Yes

set Q_DEL_AFTER_EXTRACT=Yes
set Q_MOVETOFOLDER=No
::set Q_SILENT=Yes
set Q_SILENT=No

call:Main
goto:myExit

::++++++++++++++++++++++++++++++++++++++++++++++++++++++++
::++ Main
::++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:Main
    call:preSetup
    call:preChecks
    if /i "%Q_SILENT%" equ "No"         (call:updateScreen)
    if /i "%Q_ALLOW_EXTRACT%" equ "Y"   (call:unCompressFilesCaller)
goto:eof


::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
    set ZIPFOLDER=C:\Program Files\7-Zip

    set /a iTimeOut=15
    set sSeperatorChar1=Ä
    set sSeperatorChar2=Í
    set sIsSplit=No

    set Q_ALLOW_EXTRACT=Y
    if /i "%Q_SILENT%" equ "No" (set Q_ALLOW_EXTRACT=N)

    call:getRoot myRoot
    call:getCallingFolderName "%myRoot%" UTILITY

    set iNumFiles[0][0]=7z
    set iNumFiles[1][0]=Zip
    set iNumFiles[2][0]=001

    set /a intArrLen=0

    call:countFiles "%myRoot%" *.!iNumFiles[0][0]! iNumFiles[0]
    call:countFiles "%myRoot%" *.!iNumFiles[1][0]! iNumFiles[1]
    call:countFiles "%myRoot%" *.!iNumFiles[2][0]! iNumFiles[2]
    call:arrayLength %intArrLen% intArrLen

    set /a iNumFiles.Length=%intArrLen%

    call:sumUpArray intNumerOfFiles
    call:setEchoWidth
goto:eof

::--------------------------------------------------------
::-- sumUpArray
::--------------------------------------------------------
:sumUpArray
    set /a iRetVal=0
    For /L %%C in (0,1,%iNumFiles.Length%) do (
        if !iNumFiles[%%C]! gtr 0 (
            set /a iRetVal+=!iNumFiles[%%C]!
        )
    )
    set /a %~1=%iRetVal%
goto:eof

::--------------------------------------------------------
::-- unCompressFilesCaller
::--------------------------------------------------------
:unCompressFilesCaller
    For /L %%C in (0,1,%iNumFiles.Length%) do (
        set sExt=!iNumFiles[%%C][0]!
        if !iNumFiles[%%C]! gtr 0 (
            call:unCompressFilesLoop !sExt!
        )
    )
goto:eof

::--------------------------------------------------------
::-- unCompressFilesLoop
::--------------------------------------------------------
:unCompressFilesLoop
    set sExt=%~1
    for %%f in ("%myRoot%\*.%sExt%") do (
        call:unCompressFiles "%%~nf" %%~xf
    )
goto:eof

::--------------------------------------------------------
::-- unCompressFiles
::--------------------------------------------------------
:unCompressFiles
    set sName=%~1
    set sExt=%~2
    set sFullName=%sName%%sExt%

    if /i "%Q_MOVETOFOLDER%" equ "Yes" (
        md "%sName%"
        move "%sFullName%" "%sName%"
        cd "%sName%"
    )

    if /i "%Q_SILENT%" equ "No" (
        call:logging "Extracting %sFullName%"
        7z x -y "%sFullName%" >nul
    ) else (
        7z x -y "%sFullName%"
    )

    set anyError=%errorlevel%
    if %anyError% equ 0 (
        if /i "%Q_DEL_AFTER_EXTRACT%" equ "Yes" (
            if /i "%sExt%" equ ".001" (
                call:getFileNoExt "!sName!" sNewName
                call:getFileExt   "!sName!" sNewExt
                set sDelName=!sNewName!!sNewExt!.????
                del "!sDelName!" /f /q
            ) else (
                del "%sFullName%" /f /q
            )
        )
    ) else (
        call:Logging "One of more files had an error"
        pause
    )

    if /i "%Q_MOVETOFOLDER%" equ "Yes" (cd ..)

goto:eof

::--------------------------------------------------------
::-- preSetup
::--------------------------------------------------------
:preSetup
    call:setVars
    call:setScreen
    if /i "%Q_SILENT%" equ "No" (
        call:echoVersion
    )
goto:eof

::--------------------------------------------------------
::-- setScreen
::--------------------------------------------------------
:setScreen
    call:calculateHeight %COMPUTERNAME% iWidth iHeight iColor
    if /i "%DEBUG%" neq "Yes" (mode con: cols=%iWidth% lines=%iHeight%)
    call:setColour %iColor%
    set /a iWidth-=%iSubtractWidth%
    call:Replicate %iWidth% "%sSeperatorChar1%" SEPERATOR1
    call:Replicate %iWidth% "%sSeperatorChar2%" SEPERATOR2
goto:eof

::--------------------------------------------------------
::-- calculateHeight
::--------------------------------------------------------
:calculateHeight
    set /a iW=%iVersionEchoWidth%
    set /a iH=2

    if /i "%Q_SILENT%" equ "No" (set /a iH=12)

    set /a iW_Buff=0
    set /a iH_Buff=0
    set /a iH_Buff=intNumerOfFiles

    if /i "%Q_SILENT%" equ "No" (
        set /a iH_Buff+=1
    ) else (
        set /a iH_Buff*=22
    )

    set iC=1b

    if /i "%sIsSplit%" equ "Yes" (set /a iH_Buff+=9)
    if "%Q_ALLOW_PAUSE%" equ "Yes" (set /a iH_Buff+=1)

    set /a %~2=%iW%+%iW_Buff%
    set /a %~3=%iH%+%iH_Buff%
    set %~4=%iC%
goto:eof

::--------------------------------------------------------
::-- setEchoWidth
::--------------------------------------------------------
:setEchoWidth
    set sSubtractMsg=%date% %time% :=
    set strEchoMsg=Extract: %intNumerOfFiles% files from: %UTILITY% - Ver %VERSION%
    set strEchoMeasure=%sSubtractMsg% %strEchoMsg%
    call:calculateLength "%sSubtractMsg%" iSubtractWidth
    call:calculateLength "%strEchoMeasure%" iVersionEchoWidth
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
    set strEcho=%strEchoMsg%
    Title %strEcho%
    call:Logging "%SEPERATOR2%"
    call:Logging "%strEcho%"
    call:Logging "%SEPERATOR2%"
goto:eof

::--------------------------------------------------------
::-- updateScreen
::--------------------------------------------------------
:updateScreen
    call:Logging  "Remove source: %Q_DEL_AFTER_EXTRACT%"
    call:Logging  "Move to folder: %Q_MOVETOFOLDER%"

    call:Logging  "Ext   Files"
    call:Logging  " !iNumFiles[0][0]!:  !iNumFiles[0]!"
    call:Logging "!iNumFiles[1][0]!:  !iNumFiles[1]!"
    call:Logging "!iNumFiles[2][0]!:  !iNumFiles[2]!"
    call:Logging "Tot:  %intNumerOfFiles%"
    call:Logging "%SEPERATOR2%"
    call:askQuestions
goto:eof

::--------------------------------------------------------
::-- askQuestions
::--------------------------------------------------------
:askQuestions
    call:getAnswer "OK to extract?" %Q_ALLOW_EXTRACT% Q_ALLOW_EXTRACT
goto:eof

::--------------------------------------------------------
::-- getAnswer
::--------------------------------------------------------
:getAnswer
    @choice /C:YN /D:%2 /T:%iTimeOut% /M "%~1"
    if errorlevel 2 (set %~3=N)
    if errorlevel 1 (set %~3=Y)
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
::-- countFiles
::--------------------------------------------------------
:countFiles
    set sInFolder=%~1
    set sInMask=%2
    set /a iCount=0
    for %%x in ("%sInFolder%\*!sInMask!") do set /a iCount+=1
    set %~3=%iCount%
goto:eof

::--------------------------------------------------------
::-- preChecks
::--------------------------------------------------------
:preChecks
    if not exist "%ZIPFOLDER%" (call:fail_zipMajor)
    if /i "%UTILITY%" equ "CompressObjects" (call:fail_StartIn)
    path=%path%;%zipfolder%
goto:eof

::--------------------------------------------------------
::-- getCallingFolderName
::--------------------------------------------------------
:getCallingFolderName
    for %%I in (%1) do set fldr=%%~nxI
    set "%~2=%fldr%"
goto:eof

::---------------------------------------------------------
::-- calculateLength
::---------------------------------------------------------
:calculateLength
    (echo "%~1" & echo.) | findstr /O . | more +1 | (set /p result= & call exit /b %%result%%)
    set /a %2=%errorlevel%-4
goto:eof

::--------------------------------------------------------
::-- fail_StartIn
::--------------------------------------------------------
:fail_StartIn
    call:myExitError "Please remove 'Start In' from shortcut" Error
goto:eof

::--------------------------------------------------------
::-- fail_zipMajor
::--------------------------------------------------------
:fail_zipMajor
    call:myExitError "Unable to locate: 7Zip %ZIPFOLDER% on device %COMPUTERNAME%" Error
goto:eof

::--------------------------------------------------------
::-- fail_zipErrors
::--------------------------------------------------------
:fail_zipErrors
    call:myExitError "Issues with *.%sExt%" Error
goto:eof

::--------------------------------------------------------
::-- myExitError
::--------------------------------------------------------
:myExitError
    set Q_ALLOW_PAUSE=Yes
    set Q_ALLOW_EXTRACT=N
    call:setColour %2
    call:Logging "%~1"
    goto:myExit
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
::-- getFileNoExt
::--------------------------------------------------------
:getFileNoExt
    set %2=%~n1
goto:eof

 ::--------------------------------------------------------
 ::-- getFileExt
 ::--------------------------------------------------------
 :getFileExt
     set %2=%~x1
 goto:eof

::--------------------------------------------------------
::-- removeSlash - get the new one - 1gvb1
::--------------------------------------------------------
:removeSlash
    set var=%~1
    if %var:~-1%==\ (set %2=%var:~0,-1%)
goto:eof

::--------------------------------------------------------
::-- arrayLength this var must be a global
::--------------------------------------------------------
:arrayLength
    set iArrLen=%1
    call:measureArray
    set /a iArrLen-=1
    set /a %~2=%iArrLen%

:measureArray
    if defined iNumFiles[%iArrLen%] (
        set /a iArrLen+=1
        call:measureArray
    )
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
    call:Logging "goodbye ..."
    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)
    endlocal
    endlocal
    exit
goto:eof

::--------------------------------------------------------
:: Revision History:
::-----------+------------+----------+--------------------
:: Modified  | Date       | Ver      | Reason
::-----------+------------+----------+--------------------
:: G Bishop  | 24/01/2012 | 01.01.00 | Added in this header
:: G Bishop  | 24/01/2012 | 01.02.00 | Tailored to be run from central DFS location:
:: G Bishop  | 19/03/2016 | 01.03.00 | Amended: now uses *.*
:: G Bishop  | 24/03/2016 | 01.04.00 | Amended: no longer uses: pkzip25 -extract -directories *.zip
:: G Bishop  | 24/03/2016 | 01.05.00 | Amended: now uses 7Zip syntax
:: G Bishop  | 18-11-2018 | 01.05.01 | Amended: added -y to 7Z & *.7z
:: G Bishop  | 23/01/2020 | 01.05.02 |   Added: Q_DELETE_SOURCE flag
:: G Bishop  | 17/03/2020 | 02.00.00 | Amended: to new style
:: G Bishop  | 19/03/2020 | 02.00.01 | Bug fix: would have deleted all files
:: G Bishop  | 20/03/2020 | 02.01.01 | Amended: Handles split files
:: G Bishop  | 24-08-2020 | 02.01.02 | renamed: calculateLength()
:: G Bishop  | 21-04-2021 | 02.02.00 |   Added: Q_MOVETOFOLDER to move .zip file to its own folder
:: G Bishop  | 22-05-2021 | 02.02.01 | Removed: Q_DELETE_SOURCE as it doesnt do anything
:: G Bishop  | 22-05-2021 | 02.03.00 |   Added: Q_SILENT and setScreen() and echo echoVersion() turn off output?
:: G Bishop  | 22-05-2021 | 02.03.01 | Testing: if Q_SILENT = Yes
:: G Bishop  | 25-05-2021 | 00.00.00 | Amended: unCompressFiles will now delete split zip files
::-----------+------------+----------+--------------------
:: To-Do:
:: Remove debugging info
:: Put deleting folders in an extention
:: create a split routine!
::--------------------------------------------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
