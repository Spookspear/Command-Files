::--------------------------------------------------------
::      Author: Grant Bishop
::        Date: 17 August 2016
::        Name: 7Z-EmailSendSplit.cmd
:: Description: Reads the variable myRoot and compresses all sub folders into 7Z format
::     History: See end of file
::--------------------------------------------------------
set VERSION=06.05.01

:: set myRoot=C:\WorkOS\DM1542 - Application Domain Migration\Amulet

@echo off
setlocal
setlocal EnableExtensions EnableDelayedExpansion

if [%1] equ [] (set Q_COMPRESS_TYPE=Zip) else (set Q_COMPRESS_TYPE=%~1)

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

::++++++++++++++++++++++++++++
::+ User Defined Variables
::++++++++++++++++++++++++++++
set Q_ALLOW_PAUSE=No
set Q_SEND_AN_EMAIL=No
set Q_SEND_SOMEWHERE=No
set Q_SPLIT_ZIPFILE=No
set Q_REMOVE_COMPRESSED_FOLDERS=Yes
set Q_CREATE_EXTRACTER=No
set Q_LOG_RESULTS=No
set Q_LOG_RESULTS_VERBOSE=No
set Q_DEL_IN_EXTRACTOR=No
set Q_DEL_RESULTANT_ZIP=No
set Q_ALLOW_SPECIAL_CHARS=No
set Q_EXCLUDE_ARCHIVE=Yes
set Q_DEL_RESULTANT_FILE=No
set Q_REMOVE_CALLER=Yes

set Q_UPDATE_SCREEN=No
:: set Q_COMPRESS_TYPE=Zip
:: set Q_COMPRESS_TYPE=7z

:: set Q_FOLDER_OR_FILES=Files
set Q_FOLDER_OR_FILES=Folders

::++++++++++++++++++++++++++++++++++++++++++++++++++++++++
::--------------------------------------------------------
::++++++++++++++++++++++++++++++++++++++++++++++++++++++++

if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (set TARGETSHARE="\\In1graapp027\c$\Work\BO")

call :Main
goto:myExit

::++++++++++++++++++++++++++++++++++++++++++++++++++++++++
::++ Main
::++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:Main

    call:durationTiming Start

    call :preSetup
    call :updateScreen
    call :compressUsing7Zip
    call :Logging "%SEPERATOR2%"
goto:eof

::--------------------------------------------------------
::-- updateScreen
::--------------------------------------------------------
:updateScreen
    call :Logging "%SEPERATOR1%"
    call :Logging "      Send email: %Q_SEND_AN_EMAIL%"
    call :Logging "  Working Folder: %myRoot% [%UTILITY%]"
    call :Logging "     Log to file: %Q_LOG_RESULTS% = Verbose [%Q_LOG_RESULTS_VERBOSE%]"
    call :Logging "  Send somewhere: %Q_SEND_SOMEWHERE% = [%_TARGETSHARE%]"
    call :Logging "      Split file: %Q_SPLIT_ZIPFILE%"
    call :Logging " Folder or Files: %Q_FOLDER_OR_FILES%"
    call :Logging "Type of compress: %Q_COMPRESS_TYPE%"
    call :Logging "  Compress Level: %Q_COMPRESS_LEVEL%     *** New"
    call :Logging "  Remove folders: %Q_REMOVE_COMPRESSED_FOLDERS%"
    call :Logging "Create Extracter: %Q_CREATE_EXTRACTER%"
    call :Logging "Del In Extracter: %Q_DEL_IN_EXTRACTOR%"
    call :Logging "Remove Zip After: %Q_DEL_RESULTANT_ZIP%"
    call :Logging " Exclude Archive: %Q_EXCLUDE_ARCHIVE%"
    call :Logging "   Remove Caller: %Q_REMOVE_CALLER%"

    if /i "%Q_FOLDER_OR_FILES%" equ "Files" (
        call :Logging "     Remove File: %Q_DEL_RESULTANT_FILE%"
    )

    if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (call :Logging "Robocopy Switchs: %ROBOSWITCHES%")
    call :Logging "%SEPERATOR2%"
    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause)
    ::--------------------------------------------------------
goto:eof

::++++++++++++++++++++++++++++++++++++++++++++++++++++++++
::++ Function section starts below here
::++++++++++++++++++++++++++++++++++++++++++++++++++++++++

::--------------------------------------------------------
::-- preSetup
::--------------------------------------------------------
:preSetup
    call :setVars
    call :setScreen
    call :echoVersion
    call :preChecks
goto:eof

::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
    set /a iTimeout=2
    set DIDANYWORK=No
    if not defined Q_COMPRESS_LEVEL (set Q_COMPRESS_LEVEL=9)

    call :setRoot myRoot

    if /i "%Q_LOG_RESULTS%" equ "Yes" (
        call :setLogFile "%myRoot%" LOGFILE
    )

    call :getCallingFolderName "%myRoot%" UTILITY

    if /i "%UTILITY%" equ "7Z-Compress" (call :fail_StartIn)
    if /i "%UTILITY%" equ "Bin" (call :fail_StartIn)
    if /i "%Q_SEND_AN_EMAIL%" equ "Yes" (set MAILTO=grant_bishop@hotmail.com)

    set COMPRESS_UNIT=K
    set COMPRESS_SIZE=10000

    set sSeperatorChar1=-
    set sSeperatorChar2==
    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar1=Ä)
    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar2=Í)

    set ROBO_TIMEOUT=60
    set ROBO_RETRYS=5

    set EMAILER=C:\Bin\sendEmail.exe
    set ZIPFOLDER=C:\Program Files\7-Zip
    path=%path%;%ZIPFOLDER%

    set sExtractFileName=Extract.cmd
    call :setSwitches ROBOSWITCHES

    if defined TARGETSHARE (call :deQuote TARGETSHARE _TARGETSHARE)
    if not defined TARGETSHARE (set TARGETSHARE=%myRoot%)
    call :createSplitOptions SPLITVAR ZIP_MASK

    set pipeTo=nul
    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (set pipeTo="%LOGFILE%")

    call :setEchoData

goto:eof

::--------------------------------------------------------
::-- preChecks
::--------------------------------------------------------
:preChecks
    if not exist "%ZIPFOLDER%" (call :reportProblems)
    call :undoContradictions
    if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (if not exist "%_TARGETSHARE%"             (goto:fail_Destination "%_TARGETSHARE%"))
    if /i "%Q_SPLIT_ZIPFILE%" equ "Yes"  (if exist "%Q_COMPRESS_TYPE%%ZIP_MASK%"    (call :fail_NotImplemented))
    if /i "%Q_SEND_AN_EMAIL%" equ "Yes"  (if not exist "%EMAILER%"                  (call :fail_NoMailer))
goto:eof

::--------------------------------------------------------
::-- undoContradictions undo contradictions
::--------------------------------------------------------
:undoContradictions
    :: if /i "%Q_FOLDER_OR_FILES%" equ "Files"     (set Q_REMOVE_COMPRESSED_FOLDERS=No)
    if /i "%Q_CREATE_EXTRACTER%" equ "No"       (set Q_DEL_IN_EXTRACTOR=No)
    if /i "%Q_LOG_RESULTS%" equ "No"            (set Q_LOG_RESULTS_VERBOSE=No)
    if /i "%Q_SEND_SOMEWHERE%" equ "No"         (set Q_DEL_RESULTANT_ZIP=No)
    if /i "%Q_FOLDER_OR_FILES%" equ "Folders"   (set Q_DEL_RESULTANT_FILE=No)
goto:eof

::--------------------------------------------------------
::-- compressUsing7Zip
::--------------------------------------------------------
:compressUsing7Zip
    if /i "%Q_FOLDER_OR_FILES%" equ "Files"     (call :compressFilesCaller)
    if /i "%Q_FOLDER_OR_FILES%" equ "Folders"   (call :compressFoldersCaller)
    if /i "%Q_CREATE_EXTRACTER%" equ "Yes"      (call :createExtractor)
    if /i "%Q_SEND_AN_EMAIL%" equ "Yes"         (call :sendEmail)
goto:eof

::--------------------------------------------------------
::-- compressFilesCaller
::--------------------------------------------------------
:compressFilesCaller
    pushd "%myRoot%"
    call :Logging "Objects handled.: %Q_FOLDER_OR_FILES%"
    for %%f in ("%myRoot%\*.*") do (call :compressFiles "%%~nf" %%~xf)
goto:eof

::--------------------------------------------------------
::-- compressFiles / pass in a file name and handle each folder
::--------------------------------------------------------
:compressFiles
    set filename=%~1
    set ext=%2
    set varPath=%filename%%ext%
    set canProceed=Yes

    :: Do not allow known list of programs
    if /i "%varPath%" equ "7Z-EmailSendSplit.cmd"                       (set canProceed=No)
    if /i "%varPath%" equ "7Z-EmailSendSplit.cmd.lnk"                   (set canProceed=No)
    if /i "%varPath%" equ "7Z-EmailSendSplit.cmd - Shortcut.lnk"        (set canProceed=No)
    if /i "%varPath%" equ "7Z-Extracter-7Z.cmd"                         (set canProceed=No)
    if /i "%varPath%" equ "extractCompressed.cmd"                       (set canProceed=No)
    if /i "%varPath%" equ "extractCompressed.cmd.lnk"                   (set canProceed=No)
    if /i "%varPath%" equ "extractCompressed.cmd - Shortcut.lnk"        (set canProceed=No)
    if /i "%varPath%" equ "%sCurrFile%-%myFileDate%-%MyFileTime%.Log"   (set canProceed=No)

    if /i "%canProceed%" equ "Yes" (
        set sLocalMsg=    Compressing.: %varPath%
        set DIDANYWORK=Yes
        if not exist "%filename%" (md "%filename%")
        move "%varPath%" "%filename%" >nul
        7z a -t%Q_COMPRESS_TYPE% -mx%Q_COMPRESS_LEVEL% "%filename%".%Q_COMPRESS_TYPE% "%filename%" %SPLITVAR% >nul

        if %errorlevel% equ 1 (call :reportProblems)

        if /i "%Q_DEL_RESULTANT_FILE%" equ "Yes" (
            set sLocalMsg=!sLocalMsg! - Removing
            rd "%filename%" /s /q
        ) else (
            set sLocalMsg=!sLocalMsg! - Keeping Generated folder
        )

        call :Logging "!sLocalMsg!"

    )
goto:eof

::--------------------------------------------------------
::-- compressFoldersCaller
::--------------------------------------------------------
:compressFoldersCaller
    call :Logging "%SEPERATOR1%"
    call :Logging "Type of objects handled.: %Q_FOLDER_OR_FILES%"
    call :Logging "%SEPERATOR1%"

    set bAllowed=Yes

    for /D %%i in ("%myRoot%\*") do (

        if /i "!Q_EXCLUDE_ARCHIVE!" equ "Yes" (
            set bAllowed=Yes
            if "%%i" equ "%myRoot%\_Archive" (set bAllowed=No)
            if "%%i" equ "%myRoot%\Archive" (set bAllowed=No)
        )

        if /i "!bAllowed!" equ "Yes" (
            call :compressFolders "%%i"
        )
    )
goto:eof

::--------------------------------------------------------
::-- compressFolders - pass in a file name and handle each folder
::-- Note: cannot use logging on syntax below on files with ()
::--------------------------------------------------------
:compressFolders

    set varPath=%~1

    call :Logging "Folder.: !varPath!" Logfile

    @echo %time% 7z a -t%Q_COMPRESS_TYPE% -mx%Q_COMPRESS_LEVEL% "!varPath!.%Q_COMPRESS_TYPE%" "!varPath!" %SPLITVAR%

    if /i "%Q_LOG_RESULTS%" equ "Yes" (
        @echo %time% 7z a -t%Q_COMPRESS_TYPE% -mx%Q_COMPRESS_LEVEL% "!varPath!.%Q_COMPRESS_TYPE%" "!varPath!" %SPLITVAR% >> "%LOGFILE%"
    )

    7z a -t%Q_COMPRESS_TYPE% -mx%Q_COMPRESS_LEVEL% "!varPath!.%Q_COMPRESS_TYPE%" "!varPath!" %SPLITVAR% >> %pipeTo%

    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (call :Logging "Complete ...")
    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (call :Logging "%SEPERATOR2%")

    if %errorlevel% equ 0 (call :cleanDownFoldersCaller "%varPath%")
    if %errorlevel% equ 1 (call :reportProblems)
    if %errorlevel% equ 2 (call :fail_NotImplemented)

    set DIDANYWORK=Yes

    if /i "%DIDANYWORK%" equ "Yes" (
        if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (
            call :sendDestination "!varPath!%ZIP_MASK%"
            if /i "%Q_DEL_RESULTANT_ZIP%" equ "Yes" (call :removeCompressedFile "!varPath!%ZIP_MASK%")
        )
    )

goto:eof

::--------------------------------------------------------
::-- sendDestination
::--------------------------------------------------------
:sendDestination
    set sSource=%~1

    call :getNameNoExt "%sSource%"          sSourceName
    set sSourceName=%sSourceName%.????

    call :Logging "%SEPERATOR2%"
    call :Logging "RoboCopy From/To.: %myRoot%\%sSourceName% / %_TARGETSHARE%"

    call :Logging "Syntax:"
    call :Logging "robocopy %myRoot% %_TARGETSHARE% %sSourceName% %ROBOSWITCHES%"

    robocopy "%myRoot%" "%_TARGETSHARE%" "%sSourceName%" %ROBOSWITCHES%

    call :checkErrorLevel %errorlevel% COPYOK
    if !COPYOK! leq 1 (call :Logging "Robocopy complete and was OK")
goto:eof

::--------------------------------------------------------
::-- removeCompressedFile only after transfer
::--------------------------------------------------------
:removeCompressedFile
    set sSource=%~1

    call :getNameNoExt "%sSource%"          sSourceName
    set sSourceName=%sSourceName%.????

    :: Only allow to delete if previously copied somewhere
    if /i "%Q_DEL_RESULTANT_ZIP%" equ "Yes" (
        if "%Q_LOG_RESULTS%" equ "Yes" (
            del "%sSourceName%" /q
        ) else (
            del "%sSourceName%" /q > nul
        )
    )
goto:eof

::--------------------------------------------------------
::-- cleanDownFoldersCaller - needed to handle files with ()
::--------------------------------------------------------
:cleanDownFoldersCaller
    if /i "%Q_REMOVE_COMPRESSED_FOLDERS%" equ "Yes" (
        call :cleanDownFolders %1
        if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (
            call :Logging "Removing Folder.: Complete"
            call :Logging "%SEPERATOR1%"
        )
    )
goto:eof

::--------------------------------------------------------
::-- cleanDownFolders Clean down Split folders
::--------------------------------------------------------
:cleanDownFolders
    set varPath=%~1
    call :getCallingFolderName "!varPath!" sFolderName
    call :Logging "Cleaning up.: Folder !sFolderName!"
    attrib -h -r -s /s "%varPath%\*.*"                  >> %pipeTo%
    del /f /q /s "%varPath%\*"                          >> %pipeTo%
    for /d %%i in ("%varPath%\*") do (rd /s /q "%%i")   >> %pipeTo%
    rd /s /q "%varPath%"                                >> %pipeTo%
goto:eof

::------------------------------------------------------------------------------
::-- setSwitches - sets common switches to be used wuith all copy commands
::------------------------------------------------------------------------------
:: Current Switches - meaning:
::  /MIR                    :: MIRror a directory tree (equivalent to /E plus /PURGE)
::              /E          :: copy subdirectories, including Empty ones.
::              /PURGE      :: delete dest files/dirs that no longer exist in source.
::  /XX                     :: eXclude eXtra files and directories.
::  /XF *.db    /XF file    :: eXclude Files matching given names/paths/wildcards, all links
::  /W:n                    :: Wait time between retries: default is 30 seconds. This 30
::  /LOG+                   :: Log file usually combined with /TEE
::  /TEE                    :: output to console window, as well as the log file.
::  /E          /E          :: copy subdirectories, including Empty ones.
::  /R:0        /R:n        :: number of Retries on failed copies: default 1 million.
::  /W:5        /W:5        :: Wait time between retries: 5 seconds.
::  /TS         /TS         :: include source file Time Stamps in the output.
::  /NJH        /NJH        :: No Job Header.
::  /NJS        /NJS        :: No Job Summary.
::  /NDL        /NDL        :: No Directory List - don't log directory names.
::  /NP         /NP         :: No Progress - don't display percentage copied.
:: ------------------------------------------------------------------------------
:: If verbose = No turn off top and bottom
:: AKA: /NJH /NJS /NDL
:: ------------------------------------------------------------------------------
:setSwitches
    set sSw1=
    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (set sSw1=/LOG+:"%LOGFILE%" /TEE)
    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "No" (
        set sSw2=/NP
        set sSw3=/NJH /NJS /NDL
    )
    set sSw4=/XF *.db
    set sSw5=/R:%ROBO_RETRYS%
    set sSw6=/W:%ROBO_TIMEOUT%
    set retVal=%sSw1% %sSw2% %sSw3% %sSw4% %sSw5% %sSw6%
    set "%1=%retVal%"
goto:eof

::--------------------------------------------------------
::-- setScreen
::--------------------------------------------------------
:setScreen
    call :calculateHeight %COMPUTERNAME% iWidth iHeight iColor
    
    if /i "%DEBUG%" equ "Yes" (set Q_UPDATE_SCREEN=Yes)
    if /i "%Q_UPDATE_SCREEN%" equ "Yes" (mode con: cols=%iWidth% lines=%iHeight%)
    :: if /i "%DEBUG%" equ "No"  (mode con: cols=%iWidth% lines=%iHeight%)
    set /a iWidth-=iSubtractWidth
    call :Replicate %iWidth% "%sSeperatorChar1%" SEPERATOR1
    call :Replicate %iWidth% "%sSeperatorChar2%" SEPERATOR2
    call :setColour %iColor%
goto:eof

::--------------------------------------------------------
::-- setEchoData
::--------------------------------------------------------
:setEchoData
    set strEcho=Version: %VERSION% Compress folder using %Q_COMPRESS_TYPE% %myRoot% ...
    set sSubtractMsg=%time%
    set strEchoMeasure=%sSubtractMsg% %strEcho%
    call :calculateLength "%sSubtractMsg%"   iSubtractWidth
    call :calculateLength "%strEchoMeasure%" iEchoWidth
    call :measureLoopFolders iW_Buff iH_Buff
goto:eof

::--------------------------------------------------------
::-- calculateHeight
::--------------------------------------------------------
:calculateHeight
    set /a iW=%iEchoWidth%
    set /a iH=29
    set iC=1b

    set /a iW_Buff+=4
    set /a iH_Buff*=3

    if /i "%Q_FOLDER_OR_FILES%" equ "Files" (
        set /a iH+=5
    )

    if /i "%Q_SPLIT_ZIPFILE%" equ "Yes" (set /a iH_Buff*=2)

    if %iW_Buff% gtr %iW% (set /a iW=0)
    if %iW_Buff% lss %iW% (set /a iW_Buff=0)

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (set /a iH_Buff+=1)
    :: if /i "%Q_DEL_RESULTANT_ZIP%" equ "Yes" (set /a iH_Buff+=3)

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
::-- measureFoldersLoop - get the longest
::--------------------------------------------------------
:measureLoopFolders
    set /a iLongest=0
    set /a iLineLen=0
    set /a iHeight=0
    set /a iHeightCount=0
    set /a iLineLenRobo=0

    :: measure the z7 syntax vs the robocopy syntax and return the longest
    for /D %%i in ("%myRoot%\*") do (

        if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (
            call :measureFoldersRoboCopy "%%i" iLineLenRobo
        )

        call :measureFolders7Zip "%%i" iLineLen7Zip

        set iLineLen=!iLineLenRobo!
        if !iLineLen7Zip! gtr !iLineLenRobo! (set iLineLen=!iLineLen7Zip!)
        if !iLineLen! gtr !iLongest! (set /a iLongest=!iLineLen!)

        @echo Measuring folder lengths "%%i" !iLongest!

        if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (set /a iHeight+=7)
        if /i "%Q_SEND_SOMEWHERE%" equ "No" (set /a iHeight+=2)
        set /a iHeightCount+=1

    )

    set /a iHeight=(iHeight-iHeightCount)

    set /a %1=!iLongest!
    set /a %2=!iHeight!
goto:eof

::--------------------------------------------------------
::-- measureFoldersRoboCopy
::--------------------------------------------------------
:measureFoldersRoboCopy
    set varPath=%1
    set /a iLocalLineLen=0
    call :deQuote varPath _varPath
    set sMeasureThis="%time% robocopy %myRoot% %_TARGETSHARE% %_varPath%.???? %ROBOSWITCHES%"
    call :calculateLength %sMeasureThis% iLocalLineLen
    set /a %2=!iLocalLineLen!
goto:eof

::--------------------------------------------------------
::-- measureFolders7Zip - get the longest
::--------------------------------------------------------
:measureFolders7Zip
    set varPath=%1
    set /a iLocalLineLen=0
    call :deQuote varPath _varPath
    set sMeasureThis="%time% 7z a -t%Q_COMPRESS_TYPE% -mx%Q_COMPRESS_LEVEL% !_varPath!.%Q_COMPRESS_TYPE% !_varPath! %SPLITVAR%"
    call :calculateLength %sMeasureThis% iLocalLineLen
    if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (set /a iLocalLineLen-=16)
    set /a %2=!iLocalLineLen!
goto:eof

::---------------------------------------------------------
::-- calculateLength
::---------------------------------------------------------
:calculateLength
    (echo "%~1" & echo.) | findstr /O . | more +1 | (set /p result= & call exit /b %%result%%)
    set /a %2=%errorlevel%-4
goto:eof

::--------------------------------------------------------
::-- createExtractor
::--------------------------------------------------------
:createExtractor
    set sExtractFile="%myRoot%\%sExtractFileName%"

    if /i "%Q_SPLIT_ZIPFILE%" equ "Yes" (set sMask=.001)
    if /i "%Q_SPLIT_ZIPFILE%" equ "No" (set sMask=.%Q_COMPRESS_TYPE%)

    set myPercent=%%

    set sVar1a=path=%myPercent%pa
    set sVar1b=th%myPercent%;
    set sVar1c=C:\Program Files\7-Zip

    @echo @echo off> %sExtractFile%
    @echo pushd "%_myRoot%"> %sExtractFile%

    @echo %sVar1a%%sVar1b%%sVar1c%>> %sExtractFile%

    :: construct syntax
    set sVar1=for %myPercent%%myPercent%i in ("%myPercent%
    set sVar2=~dp0"*%sMask%) do (7z x "
    set sVar3=%myPercent%%myPercent%i" -y)

    @echo %sVar1%%sVar2%%sVar3%>> %sExtractFile%

    if /i "%Q_DEL_IN_EXTRACTOR%" equ "Yes" (
        @echo del *.%Q_COMPRESS_TYPE%.???? >> %sExtractFile%
        @echo del %sExtractFileName% >> %sExtractFile%
    )

    if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (
        copy %sExtractFile% %TARGETSHARE% >> %pipeTo%
    )

    if /i "%Q_LOG_RESULTS%" equ "Yes" (call :Logging "Extractor created and copied" Yes)
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
    Title %strEcho%
    call :Logging "%SEPERATOR2%"
    call :Logging "%strEcho%"
    call :Logging "%SEPERATOR2%"
goto:eof

::--------------------------------------------------------
::-- checkErrorLevel
::--------------------------------------------------------
:checkErrorLevel
    set var=%1
    call :Logging "%SEPERATOR1%"
    if %var% equ 16 set sMessage="***FATAL ERROR***"
    if %var% equ 15 set sMessage="OKCOPY + FAIL + MISMATCHES + XTRA"
    if %var% equ 14 set sMessage="FAIL + MISMATCHES + XTRA"
    if %var% equ 13 set sMessage="OKCOPY + FAIL + MISMATCHES"
    if %var% equ 12 set sMessage="FAIL + MISMATCHES"
    if %var% equ 11 set sMessage="OKCOPY + FAIL + XTRA"
    if %var% equ 10 set sMessage="FAIL + XTRA"
    if %var% equ 9 set sMessage="OKCOPY + FAIL"
    if %var% equ 8 set sMessage="FAIL"
    if %var% equ 7 set sMessage="OKCOPY + MISMATCHES + XTRA"
    if %var% equ 6 set sMessage="MISMATCHES + XTRA"
    if %var% equ 5 set sMessage="OKCOPY + MISMATCHES"
    if %var% equ 4 set sMessage="MISMATCHES"
    if %var% equ 3 set sMessage="OKCOPY + XTRA"
    if %var% equ 2 set sMessage="XTRA"
    if %var% equ 1 set sMessage="OKCOPY"
    if %var% equ 0 set sMessage="No Change"

    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (call :Logging "error level: %var%" Yes)
    if /i "%Q_LOG_RESULTS_VERBOSE%" equ "Yes" (call :Logging %sMessage% Yes)

    if %var% geq 8 (call :fail_copyMajor %var% %sMessage%)
    set "%~2=%var%"
goto:eof

::--------------------------------------------------------
::-- getMailServer
::--------------------------------------------------------
:getMailServer     - work out mail server
    set %~1=mailhost.in1.net
goto:eof

::--------------------------------------------------------
::-- createSplitOptions, sets file ext and split volumns
::--------------------------------------------------------
:createSplitOptions
    if /i "%Q_SPLIT_ZIPFILE%" equ "Yes" (
        set %1=-v%COMPRESS_SIZE%%COMPRESS_UNIT% -y
        set %2=.%Q_COMPRESS_TYPE%.????
    ) else (
        set %1=-y
        set %2=*.%Q_COMPRESS_TYPE%
    )
goto:eof

::--------------------------------------------------------
::-- myExitError
::--------------------------------------------------------
:myExitError
    set Q_ALLOW_PAUSE=Yes
    call :setColour Error
    call :Logging "%~1"
    goto:myExit
goto:eof

::--------------------------------------------------------
::-- reportProblems
::--------------------------------------------------------
:reportProblems
    set Q_DEL_RESULTANT_FILE=No
    call :myExitError "There were problems with the process on device %COMPUTERNAME% - (%ZIPFOLDER%) ?"
goto:eof

::--------------------------------------------------------
::-- fail_StartIn
::--------------------------------------------------------
:fail_StartIn
    set /a iTimeout+=1
    set Q_LOG_RESULTS=No
    call :myExitError "Please remove 'Start In' from shortcut"
goto:eof

::--------------------------------------------------------
::-- fail_Destination
::--------------------------------------------------------
:fail_Destination
    md %~1
    set vErr=%errorlevel%
    @echo vErr %errorlevel%
    pause
    if %vErr% neq 0 (
        call :myExitError "Destination folder cannot be accessed: %_TARGETSHARE%"
    )
goto:eof

::--------------------------------------------------------
::-- fail_NotImplemented
::--------------------------------------------------------
:fail_NotImplemented
    call :myExitError "Cannot update spit zip - please remove any %ZIP_MASK% that may remain files"
goto:eof

::--------------------------------------------------------
::-- fail_copyMajor
::--------------------------------------------------------
:fail_copyMajor
    call :myExitError "RoboCopy failed to complete transfer %3 %1 %2"
goto:eof

::--------------------------------------------------------
::-- fail_NoMailer
::--------------------------------------------------------
:fail_NoMailer
    call :myExitError "Send email is set to Yes but %EMAILER% cannot be found"
goto:eof

::--------------------------------------------------------
::-- getCallingFolderName
::--------------------------------------------------------
:getCallingFolderName
    for %%I in (%1) do set fldr=%%~nxI
    set "%~2=%fldr%"
goto:eof

::--------------------------------------------------------
::-- getFileName
::--------------------------------------------------------
:getFileName
    set "%2=%~nx1"
goto:eof

::--------------------------------------------------------
::-- getFileNameThis
::--------------------------------------------------------
:getFileNameThis
    set "%1=%~nx0"
goto:eof

::--------------------------------------------------------
::-- getNameNoExtThis
::--------------------------------------------------------
:getNameNoExtThis
    set "%~1=%~n0"
goto:eof

::--------------------------------------------------------
::-- getNameNoExt
::--------------------------------------------------------
:getNameNoExt
    set "%~2=%~n1"
goto:eof

::---------------------------------------------------------
::-- deQuote
::---------------------------------------------------------
:deQuote
    for /f "delims=" %%a in ('echo %%%1%%') do set %2=%%~a
goto:eof

::--------------------------------------------------------
::-- setRoot
::--------------------------------------------------------
:setRoot
    if not defined myRoot (call :getRoot myRoot)
    call :removeSlash "%myRoot%" myRoot
    call :deQuote myRoot  _myRoot
    if not exist "%_myRoot%" (call :myExitError "Folder does not exist _myRoot %_myRoot%")
    pushd "%_myRoot%"
    call :windowsFolderCheck
    set "%~1=%_myRoot%"
goto:eof

::--------------------------------------------------------
::-- getRoot or set lRoot=%~dp0
:: Todo - pass in one parameter that says whether to use start-in or current Folder
::--------------------------------------------------------
:getRoot
    set "%~1=%cd%"
goto:eof

::--------------------------------------------------------
::-- windowsFolderCheck
::--------------------------------------------------------
:windowsFolderCheck
    if /i "%windir%\system32" equ "%_myRoot%"   (call :myExitError "Check 1 - myRoot is in %windir%\system32 - DO NOT RUN AS ADMIN")
    if /i "%windir%\system32" equ  "%cd%"       (call :myExitError "Check 2 - cd is in %windir%\system32")
goto:eof

::--------------------------------------------------------
::-- removeSlashQ
::--------------------------------------------------------
:removeSlash
    set varIn="%~1"
    call :deQuote varIn _varIn
    if "%_varIn:~-1%" equ "\" (set varOut="%_varIn:~0,-1%") else (set varOut=%varIn%)
    set %2=%varOut%
goto:eof

::---------------------------------------------------------
::-- setLogFile
::---------------------------------------------------------
:setLogFile
    call :getNameNoExtThis sCurrFile
    for /F "tokens=1-4 delims=/- " %%A in ('date/T') do set myFileDate=%%A%%B
    for /F "tokens=1,2* delims=: " %%i in ('time/T') do set myFileTime=%%i%%j%%k
    set "%~2=%~1\%sCurrFile%-%myFileDate%-%MyFileTime%.Log"
goto:eof

::---------------------------------------------------------
::-- durationTiming
::---------------------------------------------------------
:durationTiming
    set sDoWhat=%1

    if /i "%sDoWhat%" equ "Start" (set iTimeStart=%time%)

    if /i "%sDoWhat%" equ "Stop" (
        set iTimeEnd=!time!
        call:getDuration iTimeStart iTimeEnd iDuration
        call:formatThis !iDuration! sDuration
        call:Logging "Execution took ~ !sDuration! milliseconds."
    )
goto:eof

::---------------------------------------------------------
::-- getDuration
::---------------------------------------------------------
:getDuration        ' Calculate difference
    call:getMSeconds %1 iTimeStartMs
    call:getMSeconds %2 iTimeEndMs
    set /a %3=%iTimeEndMs%-%iTimeStartMs%
goto:eof

::---------------------------------------------------------
::-- getMSeconds
::---------------------------------------------------------
:getMSeconds
    Call:parseTime      %1              TimeAsArgs
    Call:calcMSeconds   %TimeAsArgs%    %2
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
::-- sendEmail
::--------------------------------------------------------
:sendEmail
    call :getMailServer MAILSERVER
    set SUBJECT=File on Host Name: %COMPUTERNAME% - compressed
    if /i "%Q_SEND_SOMEWHERE%" equ "Yes" (set SUBJECT=%SUBJECT% and sent to %_TARGETSHARE%)

    if /i "%DIDANYWORK%" equ "Yes" (
        call :Logging "Done.:"

        if exist "%LOGFILE%" (
            if defined MAILCC (
                %EMAILER% -f noreply@ineos.com -t %MAILTO% -cc %MAILCC% -u "%SUBJECT%" -s %MAILSERVER% < "%LOGFILE%"
            ) else (
                %EMAILER% -f noreply@ineos.com -t %MAILTO% -u "%SUBJECT%" -s %MAILSERVER% < "%LOGFILE%"
            )
        ) else (
            if defined MAILCC (
                %EMAILER% -f noreply@ineos.com -t %MAILTO% -cc %MAILCC% -u "%SUBJECT%" -s %MAILSERVER% -m %SUBJECT%
            ) else (
                %EMAILER% -f noreply@ineos.com -t %MAILTO% --u "%SUBJECT%" -s %MAILSERVER% -m %SUBJECT%
            )
        )
    )
goto:eof

::--------------------------------------------------------
::-- Logging
::--------------------------------------------------------
:Logging
    set sMsg=%~1
    set sMessage=%time% %sMsg%
    @echo %sMessage%
    if /i "%Q_LOG_RESULTS%" equ "Yes" (
        @echo %sMessage% >> "%LOGFILE%"
    )
goto:eof

::--------------------------------------------------------
:: myExit
::--------------------------------------------------------
:myExit

    call:durationTiming Stop

    if /i "%DIDANYWORK%" equ "No" (
        set sMessage=No work done, Q_FOLDER_OR_FILES?
    ) else (
        set sMessage=Complete, good-bye.
    )

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        set sMessage=%sMessage% press any key to close window
    ) else (
        set sMessage=%sMessage% windows will close in %iTimeout% seconds
    )

    if "%Q_REMOVE_CALLER%" equ "Yes" (
        set sMessage=%sMessage% Caller will be deleted
    )

    call :Logging "%sMessage% ..."

    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        pause
    ) else (
        if "!Q_REMOVE_CALLER!" equ "Yes" (
            timeout /T !iTimeout!
        ) else (
            timeout /T !iTimeout! >nul
        )
    )

    :: if /i "%Q_SEND_AN_EMAIL%" equ "No" (if /i "%Q_ALLOW_PAUSE%" equ "Yes" (pause))

    if /i "%Q_REMOVE_CALLER%" equ "Yes" (
        call :getFileNameThis sThisName
        set sThisName=!sThisName!.lnk
        del !sThisName! /q
    )

    endlocal
    endlocal
    :: exit

goto:eof



::--------------------------------------------------------
::-- Revision History:
::-----------+------------+----------+--------------------
:: Modified  | Date       | Ver      | Reason
::-----------+------------+----------+--------------------
:: G Bishop  | 19-08-2016 | 01.00.00 | added a line to delete root files + defining compress level
:: G Bishop  | 22-08-2016 | 01.00.01 | added: set my root to current folder using switch: %~dp0
:: G Bishop  | 31-08-2016 | 01.01.00 | replaced: removefiles with cleanDownFolders which attempts to remove file attributes
:: G Bishop  | 01-09-2016 | 01.01.01 | removed: on 7z command line
:: G Bishop  | 22-09-2016 | 01.02.00 | Now sends email when finished ...
:: G Bishop  | 15-11-2016 | 01.02.01 | was calling routine: compressUsing7Zip and subsequently sending file twice
:: G Bishop  | 25-09-2017 | 01.03.00 | can now specify server destination share
:: G Bishop  | 27-09-2017 | 01.04.00 | optional to send email & what server and to split archive - all in one
:: G Bishop  | 03-10-2017 | 01.05.00 | will handle files or folders \ move or copy & resume copying
:: G Bishop  | 03-10-2017 | 01.05.01 | This had quite a serious bug in that it doesn't handle spaces in folder names - needs more testing
:: G Bishop  | 03-10-2017 | 02.00.00 | Will now convert to 7z OR zip with use of new var called: Q_COMPRESS_TYPE
:: G Bishop  | 19-12-2017 | 02.01.00 | don't allow deleting files if Q_FOLDER_OR_FILES=Files as it will delete the results
:: G Bishop  | 23-01-2018 | 03.00.00 | Took out picking up where we left off \ or replaced with a flag
:: G Bishop  | 10-09-2018 | 03.01.00 | Amended: sendDestination for larger extensions as was restricted to 3 .??? now 4 .????
:: G Bishop  | 18-09-2018 | 03.02.00 | Added: /i for if checks (cases sensitivity)
:: G Bishop  | 18-10-2018 | 03.03.00 | Added: new Flag for whether to log the output: Q_LOG_RESULTS
:: G Bishop  | 08-11-2018 | 03.03.01 | Added: new Flag for Compression Units = COMPRESS_UNIT, so far tested M & K
:: G Bishop  | 08-11-2018 | 03.04.00 | Added: Allows files to be deleted Q_REMOVE_COMPRESSED_FOLDERS
:: G Bishop  | 24-11-2018 | 03.04.01 | Amended: createSplitOptions to make cleared what is does and took out resume copying
:: G Bishop  | 25-12-2018 | 03.04.02 | Amended: minor changes
:: G Bishop  | 22-10-2019 | 03.05.00 | Log file is appended to perpetually
:: G Bishop  | 30-10-2019 | 04.01.00 | can now be run from the shortcut using an absolute path
:: G Bishop  | 11-12-2019 | 04.01.01 | Optimised and renamed variables
:: G Bishop  | 13-12-2019 | 04.01.02 | moved TARGET share higher up so it can be edited
:: G Bishop  | 13-12-2019 | 04.01.03 | Bug Fix
:: G Bishop  | 08-01-2020 | 05.00.00 | Major rewrite \ fixed many issues and probably introduced new ones
:: G Bishop  | 09-01-2020 | 05.01.00 | Added: call :calculateLength so screen is always correct length
:: G Bishop  | 15-01-2020 | 05.02.00 | Added: call :setLogFile & call :getNameNoExtThis
:: G Bishop  | 22-01-2020 | 05.03.00 | Added: Q_ALLOW_PAUSE
:: G Bishop  | 24-01-2020 | 05.03.01 | Bugfix: put quotes around UTILITY for compare
:: G Bishop  | 24-01-2020 | 05.03.02 | Bugfix: better exit for when there is an eror
:: G Bishop  | 24-01-2020 | 05.03.03 | Bugfix: TARGETSHARE destination was not checking Q_SEND_SOMEWHERE
:: G Bishop  | 06-03-2020 | 05.03.04 | Bugfix: used _TARGETSHARE, warning for current directory, moved setLogFile up and added quotes around LOGFILE
:: G Bishop  | 06-03-2020 | 05.04.01 | Amended: Will now handle spaces in paths correctly, double height if Q_SPLIT_ZIPFILE=Yes
:: G Bishop  | 19-03-2020 | 05.05.01 | Amended: Does each folder invidually
:: G Bishop  | 19-03-2020 | 05.05.02 | Amended: Correctly handles spaces
:: G Bishop  | 20-03-2020 | 05.06.00 | Added: fail_NoMailer and lots of effort
:: G Bishop  | 03-04-2020 | 05.07.00 | Amended: setColour()
:: G Bishop  | 22-05-2020 | 05.07.01 | Amended: Placed Q_ in front on questions
:: G Bishop  | 24-08-2020 | 05.07.02 | renamed: calculateLength()
:: G Bishop  | 14-09-2020 | 05.07.03 | Amended: fail_Destination() to create folder and not just fail
:: G Bishop  | 15-09-2020 | 05.07.04 | Added: Logging of Robocopy Syntax
:: G Bishop  | 30-09-2020 | 05.08.00 | Amended: measureLoopFolders measure the z7 syntax vs the robocopy syntax and return the longest
:: G Bishop  | 11-12-2020 | 05.09.00 | Amended: measureFoldersLoop() changed loop type to: 'for /D %%i in' from: 'for /F %%i in'
:: G Bishop  | 12-03-2021 | 05.10.00 | Amended: Some work on the files part, other cosmetic
:: G Bishop  | 12-03-2021 | 05.10.01 | Amended: Removed most debug data for folders and screen updating
:: G Bishop  | 17-03-2021 | 05.10.02 | Bugfix: measurement of lines when no robocopy, renamed: measureLoopFolders()
:: G Bishop  | 25-05-2021 | 05.10.03 | Amended: Fixed height calculations and Q_LOG_RESULTS=No now works, includes: Q_ALLOW_SPECIAL_CHARS
:: G Bishop  | 02-07-2021 | 05.10.04 | Amended: Can I see this in github?
:: G Bishop  | 06-05-2023 | 05.10.05 | Amended: Final message, when no work is done
:: G Bishop  | 21-07-2023 | 06.00.00 | Amended: get working with this RunAs issue\configuration
:: G Bishop  | 21-07-2023 | 06.00.01 | Added: Q_EXCLUDE_ARCHIVE exclude Archive folders, bugfix in getRoot(), changed final messages
:: G Bishop  | 16-08-2023 | 06.01.00 | Amended: Location and switches
:: G Bishop  | 31-08-2023 | 06.01.01 | Amended: Minor changes to switches
:: G Bishop  | 30-10-2023 | 06.01.02 | Amended: Removed redudant code set /a on numeric vars
:: G Bishop  | 27-01-2024 | 06.02.00 | Renamed: Q_DEL_RESULTANT_ZIP Added: Q_DEL_RESULTANT_FILE, amended: Screen size for both cases
:: G Bishop  | 12-02-2024 | 06.03.00 | Removed: vars passed in on cmd line and much debugging
:: G Bishop  | 12-02-2024 | 06.04.00 | Added: DurationTimine Start and Stop
:: G Bishop  | 18-03-2024 | 06.04.01 | Amended: declare and renamed: Q_REMOVE_CALLER
:: G Bishop  | 08-04-2024 | 06.04.02 | Amended: Timeouts and spelling, EM for running on server as Admin
:: G Bishop  | 13-12-2024 | 06.05.00 | Added: Q_COMPRESS_TYPE on command line, defaults to Zip if not supplied
:: G Bishop  | 13-12-2024 | 06.05.01 | Added: Q_UPDATE_SCREEN - for when run from a caller
::-----------+------------+----------+--------------------
:: To-do:      12-03-2021 - deleting with the zip_mask is still a but flakey
:: Todo:       28-02-2024 - pass in one parameter that says weather to use start-in or current Folder
::-----------+------------+----------+--------------------
:: 12-03-2021 - Shorten widths if no robocopy - DONE
:: 14-09-2020 - test fail_Destination in the field
:: 11-12-2020 - is 16 chars too wide
:: 12-02-2024 - Include: durationTiming() - DONE - not tested
:: 12-02-2024 - Restart versioning
::::------------+------------+----------+-------------------
:: Operator | Description
:: EQU      | equal to
:: NEQ      | not equal to
:: LSS      | less than
:: LEQ      | less than or equal to
:: GTR      | greater than
:: GEQ      | greater than or equal to
::--------------------------------------------------------
