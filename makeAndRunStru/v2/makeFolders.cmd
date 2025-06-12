::--------------------------------------------------------
::      Author: G Bishop
::        Date: 20th August 2018
::        Name: makeFolders.cmd
:: Description: sets up the current working item folder all in one
::     History: See end of file
::--------------------------------------------------------
@echo off
setlocal
setlocal EnableExtensions EnableDelayedExpansion

set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

set VERSION=04.05.00

set Q_ALLOW_PAUSE=Yes
set Q_FOLDER_ICON=Yes
set Q_DEVICE=HE-5H9YZHMREDIB
:: set Q_DEVICE=DESKTOP-OIFDHFS

:: Required, due to 'Run As' being C:\Windows

if /i "%DEBUG%" equ "Yes" (
    @echo cd    %cd%
    @echo cd    %~dp0
    pause
)

if /i "%COMPUTERNAME%" neq "%Q_DEVICE%" (
    set sRootDrv=K:
) else (
    set sRootDrv=C:
)

set G_Root=%sRootDrv%\Users\grant.bishop\OneDrive - Harbour Energy plc

if /i "%DEBUG%" equ "Yes" (
    @echo G_Root  %G_Root%
    @echo G_Root  %~dp0
    pause
)
:: pushd %cd%
pushd %~dp0

if /i "%DEBUG%" equ "Yes" (
    @echo cd    %cd%
    pause
)

Title %VERSION%

call :setFolderNamesCaller

call :preChecks
call :preSetup
call :Main
call :myExit Yes

::--------------------------------------------------------
::-- setFolderNamesCaller
::--------------------------------------------------------
:setFolderNamesCaller
if /i "%DEBUG%" equ "Yes" (@echo setFolderNamesCaller - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not defined iX (set /a iX=0)

    if /i "%DEBUG%" equ "Yes" (@echo sCurrFile: %sCurrFile%)
    if /i "%DEBUG%" equ "Yes" (pause)

    call :getNameNoExt sCurrFile
    set sCurrFile=%sCurrFile%.txt
    if not exist %sCurrFile% (call :myExitError "Unable to find input file")

    for /f "tokens=* delims= " %%f in (%sCurrFile%) do (
        call :autoMateWorker "%%f"
    )

    if %iX% neq 0 (
        call :Pad %iX%      strFoldesTotal
    )

if /i "%DEBUG%" equ "Yes" (@echo setFolderNamesCaller - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- autoMateWorker
::--------------------------------------------------------
:autoMateWorker
if /i "%DEBUG%" equ "Yes" (@echo autoMateWorker - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sTARGETFolder=%~1

    if /i "%DEBUG%" equ "Yes" (@echo sTARGETFolder: %sTARGETFolder%)
    if /i "%DEBUG%" equ "Yes" (pause)

    call :getFirstXChars "%sTARGETFolder%" 2 sSep

    if /i "%DEBUG%" equ "Yes" (@echo sSep: %sSep%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%sSep%" neq "::" (
        call :setFolderNames "%sTARGETFolder%"
    )

if /i "%DEBUG%" equ "Yes" (@echo autoMateWorker - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setFolderNames
::--------------------------------------------------------
:setFolderNames
if /i "%DEBUG%" equ "Yes" (@echo setFolderNames - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if /i [%1] equ [] (set sFldr="Problems") else (set sFldr="%~1")
    if /i [%2] equ [] (set sRITM="")         else (set sRITM=%~2)
    if /i [%3] equ [] (set sTASK="")         else (set sTASK=%~3)

    if not defined iX (set /a iX=0)
    set sFldr=%~1
    set sTARGETFilePath[!iX!]=%sFldr%

    if defined sRITM (
        set sTARGETFilePath[!iX!][0]=%sRITM%
    ) else (
        @echo was not def1
    )

    if defined sTASK (
        set sTARGETFilePath[!iX!][1]=%sTASK%
    ) else (
        @echo was not def2
    )

    set /a iX+=1

if /i "%DEBUG%" equ "Yes" (@echo setFolderNames - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- Main
::--------------------------------------------------------
:Main
if /i "%DEBUG%" equ "Yes" (@echo Main - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    For /L %%C in (0,1,%sTARGETFilePath.Length%) do (
        call :createFolderCaller "!sTARGETFilePath[%%C]!" "!sTARGETFilePath[%%C][0]!" "!sTARGETFilePath[%%C][1]!" %%C
    )

if /i "%DEBUG%" equ "Yes" (@echo Main - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- createFolderCaller
::--------------------------------------------------------
:createFolderCaller
if /i "%DEBUG%" equ "Yes" (@echo createFolderCaller - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if /i [%1] equ [] (set strFolder="Problems") else (set strFolder="%~1")
    if /i [%2] equ [] (set sRITM=)               else (set sRITM=%~2)
    if /i [%3] equ [] (set sTASK=)               else (set sTASK=%~3)
    if /i [%4] equ [] (set intNumber=)           else (set intNumber=%~4)

    set /a intNumber+=1
    set bSubFolder=False

    if [%sRITM%] equ [""] (set sRITM=RITM0000000)
    if [%sTASK%] equ [""] (set sTASK=SCTASK0000000)

    if /i "%DEBUG%" equ "Yes" (@echo sRITM: %sRITM%)
    if /i "%DEBUG%" equ "Yes" (@echo sTASK: %sTASK%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%DEBUG%" equ "Yes" (@echo before RTrim)
    if /i "%DEBUG%" equ "Yes" (@echo strFolder: %strFolder%)
    if /i "%DEBUG%" equ "Yes" (pause)

    call :RTrim %strFolder%  strFolder

    if /i "%DEBUG%" equ "Yes" (@echo After RTrim)
    if /i "%DEBUG%" equ "Yes" (@echo strFolder: %strFolder%)
    if /i "%DEBUG%" equ "Yes" (pause)

    call :deQuote strFolder  _strFolder

    if /i "%DEBUG%" equ "Yes" (@echo dequoted)
    if /i "%DEBUG%" equ "Yes" (@echo _strFolder: %_strFolder%)
    if /i "%DEBUG%" equ "Yes" (@echo pause1)
    if /i "%DEBUG%" equ "Yes" (pause)

    set /a iLengthTicket=10
    set /a iLengthPadding=3
    set /a iLengthTotal=iLengthTicket
    set /a iLengthTotal+=iLengthPadding

    call :calculateStringLength "%_strFolder%" intLengthFolder

    if /i "%DEBUG%" equ "Yes" (@echo pause2)
    if /i "%DEBUG%" equ "Yes" (pause)

    set /a iLengthSubtract+=iLengthTicket
    set /a iLengthSubtract+=iLengthPadding

    call :getFirstXChars "%_strFolder%" %iLengthTicket% sNumber
    call :getMidXChars   "%_strFolder%" %iLengthTotal% %intLengthFolder% sTitle
    call :getFirstXChars "%_strFolder%" 3 sINCorREQ

    if /i "%DEBUG%" equ "Yes" (@echo pause3)
    if /i "%DEBUG%" equ "Yes" (pause)

    set _strFolder=%sNumber% -%sTitle%

    call :workoutFolderCaller "%sINCorREQ%" myRoot
    call :setStandardFiles "%sINCorREQ%"

    if /i "%DEBUG%" equ "Yes" (@echo myRoot: %myRoot%)
    if /i "%DEBUG%" equ "Yes" (pause)

    pushd "%myRoot%"
    if /i "%DEBUG%" equ "Yes" (@echo pause5)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%DEBUG%" equ "Yes" (@echo _strFolder: %_strFolder%)
    if /i "%DEBUG%" equ "Yes" (@echo    sNumber: %sNumber%)
    if /i "%DEBUG%" equ "Yes" (@echo     sTitle: %sTitle%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%DEBUG%" equ "Yes" (@echo  sINCorREQ: %sINCorREQ%)
    if /i "%DEBUG%" equ "Yes" (@echo bSubFolder: %bSubFolder%)
    if /i "%DEBUG%" equ "Yes" (pause)

    :: Always filed by category/subfolder apart from problems

    :: set bSubFolder=True
    if /i "%sINCorREQ%" neq "PRB" (
        set bSubFolder=True
    )

    :: 28-01-2025 - Remove this after a couple of weeks
    :: if /i "%sINCorREQ%" equ "REQ" (
    ::     set bSubFolder=True
    :: )
    :: if /i "%sINCorREQ%" equ "ISB" (
    ::     set bSubFolder=True
    :: )
    :: if /i "%sINCorREQ%" equ "CHG" (
    ::     set bSubFolder=True
    :: )
    :: if /i "%sINCorREQ%" equ "INC" (
    ::     set bSubFolder=True
    :: )
    :: if /i "%sINCorREQ%" equ "PER" (
    ::     set bSubFolder=True
    :: )

    if /i "%DEBUG%" equ "Yes" (@echo bSubFolder: %bSubFolder%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%bSubFolder%" equ "True" (
        call :switchFolder "%sTitle%"
    )

    call :createFolder "%_strFolder%" "%sNumber%" "%sTitle%" !intNumber!

if /i "%DEBUG%" equ "Yes" (@echo createFolderCaller - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- switchFolder
::--------------------------------------------------------
:switchFolder
if /i "%DEBUG%" equ "Yes" (@echo switchFolder - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sTitle=%~1

    call :LTrim "%sTitle%" _sTitle

    if /i "%DEBUG%" equ "Yes" (
        @echo _sTitle %_sTitle%
        pause
    )

    call :inStr "%_sTitle%" "-" intPosition
    set /a intPosition-=2

    if /i "%DEBUG%" equ "Yes" (
        @echo intPosition %intPosition%
        pause
    )

    call :getFirstXChars "%_sTitle%" %intPosition% sCat

    if /i "%DEBUG%" equ "Yes" (
        @echo sCat#1 XX%sCat%YY
        pause
    )

    if /i "%DEBUG%" equ "Yes" (
        @echo sINCorREQ %sINCorREQ%
        @echo sINCorREQ %sINCorREQ%
        pause
    )

    if /i "%DEBUG%" equ "Yes" (
        @echo sCat#2 XX%sCat%YY
        pause
    )

    if not exist "%sCat%\Desktop.ini" (
        call :placeIconOnFolderBuiltIn "%cd%" "%sCat%" "%ICON_CATEGORY%"
    )

    cd "%sCat%"
    if /i "%DEBUG%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo switchFolder - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- preChecks
::--------------------------------------------------------
:preChecks
if /i "%DEBUG%" equ "Yes" (@echo preChecks - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if %iX% equ 0 (
        call :myExitError "folder count is zero"
    ) else (
        call :Logging "folder count is: %strFoldesTotal%"
        pause
    )

if /i "%DEBUG%" equ "Yes" (@echo preChecks - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- preSetup
::--------------------------------------------------------
:preSetup
if /i "%DEBUG%" equ "Yes" (@echo preSetup - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call :setVars
    call :setScreen
    call :echoVersion
    call :postChecks

if /i "%DEBUG%" equ "Yes" (@echo preSetup - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
:: XCOPY_SWITCHES:          :: meaning:
:: -------------------------------------------------------
:: /D                       :: Only newer files
:: /C                       :: Continues copying even if errors occur.
:: /R                       :: Overwrites read-only files.
:: /H                       :: Copies hidden and system files also.
:: /K                       :: Copies attributes. Normal XCopy will reset read-only attributes.
:: /Y                       :: Suppresses prompting to confirm you want to overwrite an existing destination file.
:: /Q                       :: Does not display file names while copying.
:: /E                       :: Copies directories and subdirectories, including empty ones.
:: /I                       :: If destination does not exist and copying more than one file, assumes that destination must be a directory.
:: /V                       :: Verifies the size of each new file.
:: /Z                       :: Copies networked files in restartable mode.
:: /F                       :: Displays full source and destination file names while copying.
::--------------------------------------------------------

::--------------------------------------------------------
::-- setVars
::--------------------------------------------------------
:setVars
if /i "%DEBUG%" equ "Yes" (@echo setVars - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    :: set G_Root=C:\Work
    :: set G_Root=%OneDrive%
    :: set G_Root=C:\Users\grant.bishop\OneDrive - Harbour Energy plc

    call :setStandardFilesGlobal
    call :setStandardIcons

    if not defined Q_ALLOW_SPECIAL_CHARS (set Q_ALLOW_SPECIAL_CHARS=No)
    set sSeperatorChar1=-
    set sSeperatorChar2==

    set /a iTimeOut=5
    set /a iCount=0
    set /a intArrLen=0

    set XCOPY_SWITCHES=/C /K /R /V /Y /Z /F /Q

    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar1=Ä)
    if /i "%Q_ALLOW_SPECIAL_CHARS%" equ "Yes" (set sSeperatorChar2=Í)

    call :arrayLength %intArrLen% intArrLen
    set sTARGETFilePath.Length=%intArrLen%

    set boolTurnOffErrors=False

if /i "%DEBUG%" equ "Yes" (@echo setVars - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setStandardFilesGlobal
::--------------------------------------------------------
:setStandardFilesGlobal

    if /i "%COMPUTERNAME%" equ "UKN6441" (
        set FOLDER_TTD=%G_Root%\Documents\Templates\ThingsToDo.{App}{Ref}.xlsx
        set FOLDER_ARCHIVE=%G_Root%\Documents\Development\VisualBasic.Script\File Shifter\ArchiveFiles.vbs.lnk
    ) else (

        if /i "%COMPUTERNAME%" equ "%Q_DEVICE%" (
            set FOLDER_TTD=%G_Root%\Documents\Templates\ThingsToDo.{App}{Ref}.xlsx
            set FOLDER_ARCHIVE=%G_Root%\Documents\Development\VisualBasic.Script\File Shifter\Shortcuts\ArchiveFiles.vbs.lnk
        )
    )

    if not defined FOLDER_TTD       (call :myExitError "TTD Path no defined")
    if not defined FOLDER_ARCHIVE   (call :myExitError "TTD Path no defined")

goto:eof

::--------------------------------------------------------
::-- setStandardFiles
::--------------------------------------------------------
:setStandardFiles
if /i "%DEBUG%" equ "Yes" (@echo setStandardFiles - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set FOLDER_TTD=%G_Root%\Documents\Templates\ThingsToDo.%sINCorREQ%.{App}{Ref}.xlsx

if /i "%DEBUG%" equ "Yes" (@echo setStandardFiles - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setStandardIcons
::--------------------------------------------------------
:setStandardIcons
    set ICON_SEARCH=%windir%\System32\SHELL32.dll,218
    set ICON_ARCHIVE=%windir%\System32\SHELL32.dll,45
    set ICON_SCREENSHOT=%windir%\System32\SHELL32.dll,127
    set ICON_CATEGORY=%windir%\System32\SHELL32.dll,80
    set ICON_PDF=%G_Root%\Pictures\Icons\PDFs.ico,0
goto:eof

::--------------------------------------------------------
::-- echoVersion
::--------------------------------------------------------
:echoVersion
    set strEcho=Creating folders - Version %VERSION%
    Title %strEcho%
    call :Logging "%SEPERATOR2%"
    call :Logging "%strEcho%"
goto:eof

::--------------------------------------------------------
::-- postChecks
::--------------------------------------------------------
:postChecks
if /i "%DEBUG%" equ "Yes" (@echo postChecks - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if not exist "%FOLDER_ARCHIVE%"     (call :myExitError "Cannot find %FOLDER_ARCHIVE%")
    if /i "%G_Root%" neq "%G_Root%"   (call :setColour Warning)

if /i "%DEBUG%" equ "Yes" (@echo postChecks - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- createFolder
::--------------------------------------------------------
:createFolder
if /i "%DEBUG%" equ "Yes" (@echo createFolder - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set strFolder=%~1
    set sNumber=%~2
    set sTitle=%~3
    set intNumber=%~4

    if /i "%DEBUG%" equ "Yes" (
        @echo strFolder     %strFolder%
        @echo sNumber       %sNumber%
        @echo sTitle        %sTitle%
        @echo intNumber     %intNumber%
    )

    call :Logging "%SEPERATOR1%"
    call :Logging "Creating: %strFolder%"
    call :Logging "%SEPERATOR1%"

    set boolTurnOffErrors=False

    if /i "%Q_FOLDER_ICON%" equ "No" (
        if not exist "%strFolder%" (md "%strFolder%")
    ) else (
        call :placeIconOnFolderBuiltIn "%cd%" "%strFolder%" "%ICON_SEARCH%"
    )

    call :makeSubFldrs "%strFolder%" %intNumber%
    call :getFilesCommon "%strFolder%"
    call :nameTTD "%strFolder%" "%sNumber%" "%sTitle%"

if /i "%DEBUG%" equ "Yes" (@echo createFolder - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- nameTTD
::--------------------------------------------------------
:nameTTD
if /i "%DEBUG%" equ "Yes" (@echo nameTTD - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set strFolder=%~1
    set sNumber=%~2
    set sTitle=%~3

    if /i "%DEBUG%" equ "Yes" (
        @echo strFolder      %strFolder%
        @echo sNumber        %sNumber%
        @echo sTitle         %sTitle%
        pause
    )

    call :LTrim "%sTitle%" _sTitle
    call :inStr "%sTitle%" "-" intPosition
    call :calculateStringLength "%sTitle%" intRemain
    set /a intPosition+=1
    call :getMidXChars "%_sTitle%" %intPosition% %intRemain% sNewName

    set sFile=ThingsToDo.%sNumber%.%sNewName%.xlsx
    if /i "%DEBUG%" equ "Yes" (pause)

    pushd %strFolder%
    if /i "%DEBUG%" equ "Yes" (pause)

    call :renameFile "ThingsToDo.%sINCorREQ%.{APP}{Ref}.xlsx" "%sFile%"
    if /i "%DEBUG%" equ "Yes" (pause)
    popd

if /i "%DEBUG%" equ "Yes" (@echo nameTTD - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- makeSubFldrs - Images is common to all
::--------------------------------------------------------
:makeSubFldrs
    set strFolder=%~1
    set intNumber=%~2

    call :Pad %intNumber%      sPad1
    call :Logging "  Making: Sub Folders - [%sPad1%/!strFoldesTotal!]"

    :: Common
    call :makeSubSubNew "%strFolder%" "_Archive" "%ICON_ARCHIVE%"
    call :makeSubSubNew "%strFolder%" "Images" "%ICON_SCREENSHOT%"
    :: call :makeSubSubNew "%strFolder%" "PDFs" "%ICON_PDF%"

goto:eof

::--------------------------------------------------------
::-- Pad
::--------------------------------------------------------
:Pad
    set /a iN=%1
    set sR=000
    if %iN% leq 9  (set sR=00)
    if %iN% gtr 9  (set sR=0)
    if %iN% gtr 99 (set sR=)
    set %~2=%sR%%iN%
goto:eof

::--------------------------------------------------------
::-- getFilesCommon
::--------------------------------------------------------
:getFilesCommon
    set strFolder=%~1
    call :Logging " Getting: Common files"

    call :xcopySourceTarget "%FOLDER_ARCHIVE%" "%strFolder%\_Archive"

    if not exist "%strFolder%\ThingsToDo*.xlsx" (
        call :xcopySourceTarget "%FOLDER_TTD%"  "%strFolder%"
    ) else (
        call :Logging " Message: ThingsToDo.%sINCorREQ%.{APP}{Ref}.xlsx already exists"
        set boolTurnOffErrors=True
    )
goto:eof

::--------------------------------------------------------
::-- renameFile
::--------------------------------------------------------
:renameFile
if /i "%DEBUG%" equ "Yes" (@echo renameFile - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sFileSource=%~1
    set sFileTarget=%~2

    call :Logging "  Naming: %sFileTarget%"

    if /i "%DEBUG%" equ "Yes" (@echo sFileTarget: "%sFileTarget%")
    if /i "%DEBUG%" equ "Yes" (@echo sFileSource: "%sFileSource%")
    if /i "%DEBUG%" equ "Yes" (pause)

    if not exist "%sFileTarget%" (
        if exist "%sFileSource%" (
            ren "%sFileSource%" "%sFileTarget%"
        ) else (
            if "%boolTurnOffErrors%" neq "True" (
                call :Logging "PROBLEM - Missing: %sFileSource%
                pause
           )
        )
    )

if /i "%DEBUG%" equ "Yes" (@echo renameFile - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- getCallingFolderName
::--------------------------------------------------------
:getCallingFolderName
    for %%I in (%1) do set fldr=%%~nxI
    set "%~2=%fldr%"
goto:eof

::--------------------------------------------------------
::-- getNameNoExt
::--------------------------------------------------------
:getNameNoExt
    set "%~1=%~n0"
goto:eof

::--------------------------------------------------------
::-- makeSubSub
::--------------------------------------------------------
:makeSubSub
    set strFolder=%~1
    set strFolderSub=%~2
    call :Logging "creating: %strFolder%\%strFolderSub%"
    if not exist "%strFolder%\%strFolderSub%" (md "%strFolder%\%strFolderSub%" >nul)
goto:eof

::--------------------------------------------------------
::-- makeSubSubNew - assign icon on the fly
::--------------------------------------------------------
:makeSubSubNew
    set strFolder=%~1
    set strFolderSub=%~2
    set sSourceIcon=%~3

    call :Logging "creating: %strFolder%\%strFolderSub%"
    call :placeIconOnFolderBuiltIn "%strFolder%" "%strFolderSub%" "%sSourceIcon%"

goto:eof

::--------------------------------------------------------
::-- placeIconOnFolderBuiltIn
::--------------------------------------------------------
:placeIconOnFolderBuiltIn
if /i "%DEBUG%" equ "Yes" (@echo placeIconOnFolderBuiltIn - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sCurrentFolder=%~1
    set sNewFolderName=%~2
    set sSourceIcon=%~3

    set bPopBack=No
     if /i "%DEBUG%" equ "Yes" (pause)

    if /i "%DEBUG%" equ "Yes" (@echo 1 - changed folder? - !cd!)
    if /i "%DEBUG%" equ "Yes" (pause)

    if "!cd!" neq "%sCurrentFolder%" (
        pushd "%sCurrentFolder%"
        set bPopBack=Yes
    )

    if /i "%DEBUG%" equ "Yes" (@echo 2- folder changed? - %cd%)
    if /i "%DEBUG%" equ "Yes" (pause)

    if not exist "%sNewFolderName%" (
        md "%sNewFolderName%"
    )
    if /i "%DEBUG%" equ "Yes" (pause)

    if exist "%sNewFolderName%" (

        set sDeskTopINI=%sNewFolderName%\Desktop.ini

        if not exist "!sDeskTopINI!" (
            attrib -r "%sNewFolderName%"
            if /i "%DEBUG%" equ "Yes" (pause)

            @echo [.ShellClassInfo]> "!sDeskTopINI!"
            @echo IconResource=%sSourceIcon%>> "!sDeskTopINI!"
            @echo IconIndex=0 >> "!sDeskTopINI!"
            @echo InfoTip=!sNewFolderName!>> "!sDeskTopINI!"
            @echo [ViewState]>> "!sDeskTopINI!"
            @echo Mode=>> "!sDeskTopINI!"
            @echo Vid=>> "!sDeskTopINI!"
            @echo FolderType=StorageProviderGeneric>> "!sDeskTopINI!"

            if /i "%DEBUG%" equ "Yes" (pause)

            attrib +h +r "!sDeskTopINI!" >nul:
            attrib +r "%sNewFolderName%" >nul:
        )

    )

    if /i "%DEBUG%" equ "Yes" (@echo 3 - popd folder? - !cd!)
    if /i "%DEBUG%" equ "Yes" (pause)

    if "!bPopBack!" equ "Yes" (
        if /i "%DEBUG%" equ "Yes" (@echo will pop back)
        popd
    )

    if /i "%DEBUG%" equ "Yes" (@echo 4 - popd folder after? - !cd!)
    if /i "%DEBUG%" equ "Yes" (pause)

if /i "%DEBUG%" equ "Yes" (@echo placeIconOnFolderBuiltIn - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- xcopySourceTarget
::--------------------------------------------------------
::-- /C  Continues copying even if errors occur.
::-- /K  Copies attributes. Normal Xcopy will reset read-only attributes.
::-- /R  Overwrites read-only files.
::-- /V  Verifies the size of each new file.
::-- /Y  Suppresses prompting to confirm you want to overwrite an existing destination file.
::-- /Z  Copies networked files in restartable mode.
::-- /F  Displays full source and destination file names while copying.
::-- /Q  Does not display file names while copying.
::--------------------------------------------------------
:xcopySourceTarget
    set SOURCE=%~1
    set TARGET=%~2

    call :getFileName sSourceName "%SOURCE%"
    call :Logging " Copying: %sSourceName%"

    if not exist "%SOURCE%\%sSourceName%" (
        xcopy "%SOURCE%" "%TARGET%" %XCOPY_SWITCHES% >Nul
        if %errorlevel% neq 0 (call :fail_copyMajor)
    ) else (
        call :Logging "File: %TARGET%\%sSourceName% = missing"
        pause
    )
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
::-- setScreen
::--------------------------------------------------------
:setScreen
    call :calculateHeight %iArrLen% iWidth iHeight iColor
    if /i "%DEBUG%" neq "Yes" (mode con: cols=%iWidth% lines=%iHeight%)
    call :setColour %iColor%
    call :Replicate %iWidth% "%sSeperatorChar1%" SEPERATOR1
    call :Replicate %iWidth% "%sSeperatorChar2%" SEPERATOR2
goto:eof

::--------------------------------------------------------
::-- calculateHeight
::--------------------------------------------------------
:calculateHeight
if /i "%DEBUG%" equ "Yes" (@echo calculateHeight - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set /a iNoCalls=%1
    set /a iW=2

    set /a iH=8
    set /a iH_Buff=0

    set /a iW_Buff=2
    set /a iW_Buff1=0
    set /a iW_Buff2=0
    set iC=Info

    call :setEchoWidth            iW_Buff1
    call :calculateGreatestLength iW_Buff2

    set /a iW_Buff2-=9

    set /a iW_Buff=iW_Buff1
    if %iW_Buff2% gtr %iW_Buff1%  (set /a iW_Buff=%iW_Buff2%)

    set /a iHeightEach=10
    set /a iH_Buff+=iHeightEach
    set /a iH_Buff*=iNoCalls

    set /a %~2=iW+iW_Buff
    set /a %~3=iH+iH_Buff
    set %~4=%iC%

if /i "%DEBUG%" equ "Yes" (@echo calculateHeight - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::---------------------------------------------------------
::-- calculateGreatestLength
::---------------------------------------------------------
:calculateGreatestLength
if /i "%DEBUG%" equ "Yes" (@echo calculateGreatestLength - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set /a iStrLen=0
    set /a iStrLenRetVal=0
    For /L %%C in (0,1,%sTARGETFilePath.Length%) do (
        call :calculateStringLength "Creating: 000000 - !sTARGETFilePath[%%C]!\Images" iStrLen
        if !iStrLen! gtr !iStrLenRetVal! (set iStrLenRetVal=!iStrLen!)
    )
    set %1=%iStrLenRetVal%

if /i "%DEBUG%" equ "Yes" (@echo calculateGreatestLength - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::---------------------------------------------------------
::-- calculateStringLength
::---------------------------------------------------------
:calculateStringLength
    (echo "%~1" & echo.) | findstr /O . | more +1 | (set /p result= & call exit /b %%result%%)
    set /a %2=%errorlevel%-4
goto:eof

::--------------------------------------------------------
::-- setEchoWidth
::--------------------------------------------------------
:setEchoWidth
if /i "%DEBUG%" equ "Yes" (@echo setEchoWidth - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set strEchoMsg=Creating folders - Version %VERSION%
    call :calculateStringLength "%strEchoMsg%" %1

if /i "%DEBUG%" equ "Yes" (@echo setEchoWidth - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- getFirstXChars
::--------------------------------------------------------
:getFirstXChars
    set sValIn=%~1
    set /a iNo=%2
    set vWorkVal="%%sValIn:~0,%iNo%%%"
    call :getFirstX %vWorkVal% vWorkVal2
    call :deQuote vWorkVal2 vWorkVal2
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
    call :getLastX %vWorkVal% vWorkVal2
    call :deQuote vWorkVal2 vWorkVal2
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
    call :calculateStringLength "%sValIn%" iWorkValLength
    set /a iWorkValLength-=%iStart%
    call :getLastXChars "%sValIn%" %iWorkValLength% vWorkVal
    call :getFirstXChars "%vWorkVal%" %iLength% vRetVal
    set %~4=%vRetVal%
goto:eof

::--------------------------------------------------------
::-- arrayLength this var must be a global
::--------------------------------------------------------
:arrayLength
    set iArrLen=%1
    call :measureArray
    set /a iArrLen-=1
    set /a %~2=%iArrLen%

:measureArray
    if defined sTARGETFilePath[%iArrLen%] (
        set /a iArrLen+=1
        call :measureArray
    )
goto:eof

::---------------------------------------------------------
::-- deQuote
::---------------------------------------------------------
:deQuote
    for /f "delims=" %%a in ('echo %%%1%%') do set %2=%%~a
goto:eof

::--------------------------------------------------------
::-- fail_copyMajor
::--------------------------------------------------------
:fail_copyMajor
    call :myExitError "xcopy failed - quitting"
goto:eof

::--------------------------------------------------------
::-- getFileName
::--------------------------------------------------------
:getFileName
    set %1=%~nx2
goto:eof

::--------------------------------------------------------
::-- inStr - Returns the position of the first occurrence of one string within another.
::--------------------------------------------------------
:inStr

    set sValue=%~1
    set sFind=%~2

    set /a intAllowed=1
    set /a intNumFound=0
    call :calculateStringLength "%sValue%" iLoop

    for /L %%G in (0,1,!iLoop!) do (
        set sScanVar=!sValue:~%%G,1!
        if /i !sScanVar! equ !sFind! (
            set /a intNumFound+=1
            if !intNumFound! leq !intAllowed! (
                set /a iResult=%%G
            )
        )
    )

    :: exit /b
    if /i "%DEBUG%" equ "Yes" (
        @echo intNumFound %intNumFound%
        @echo iResult     %iResult%
        pause
    )

    set /a %~3=%iResult%+1
goto:eof

::--------------------------------------------------------
::-- workoutFolder
::--------------------------------------------------------
:workoutFolder
    set sValueIn=%~1
    call :inStr "%sValueIn%" "-" intPosition
    call :calculateStringLength "%sValueIn%" intLengthAPP
    set /a intPosition+=2
    set /a intRemain=intLengthAPP
    set /a intRemain-=intPosition
    call :getMidXChars "%sValueIn%" %intPosition% %intRemain% sRetVal
    set %~2=%sRetVal%
goto:eof

::--------------------------------------------------------
::-- workoutFolderCaller
::--------------------------------------------------------
:workoutFolderCaller
if /i "%DEBUG%" equ "Yes" (@echo workoutFolderCaller - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set sINCorREQ=%~1

    set lVarLocal=%G_Root%

    if /i "%sINCorREQ%" equ "INC" (
        set sRetVal=%lVarLocal%\Documents\Service Now\Incidents
    )

    if /i "%sINCorREQ%" equ "REQ" (
        set sRetVal=%lVarLocal%\Documents\Service Now\Requests
    )

    if /i "%sINCorREQ%" equ "CHG" (
        set sRetVal=%lVarLocal%\Documents\Service Now\Changes
    )

    if /i "%sINCorREQ%" equ "ISB" (
        set sRetVal=%lVarLocal%\Documents\Work
    )

    if /i "%sINCorREQ%" equ "PER" (
        set sRetVal=%lVarLocal%\Documents\Service Now\Personal
    )

    if /i "%sINCorREQ%" equ "PRB" (
        set sRetVal=%lVarLocal%\Documents\Service Now\Problem
    )

    set %~2=%sRetVal%

if /i "%DEBUG%" equ "Yes" (@echo workoutFolderCaller - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- myExitError
::--------------------------------------------------------
:myExitError
if /i "%DEBUG%" equ "Yes" (@echo myExitError - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    if /i [%1] equ [] (set sLocalMsg="Problems") else (set sLocalMsg=%1)
    if /i [%2] equ [] (set sAllOK=No)            else (set sAllOK=%2)

    call :setColour Error
    call :Logging %sLocalMsg%
    call :myExit %sAllOK%

if /i "%DEBUG%" equ "Yes" (@echo myExitError - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- getRoot or set lRoot=%~dp0 - not used
::--------------------------------------------------------
:getRoot
    :: set lRoot=%cd%
    set lRoot=%G_Root%\Documents\Service Now\Requests
    call :removeSlash "%lRoot%" lRoot
    set %~1=%lRoot%
goto:eof

::--------------------------------------------------------
::-- removeSlash
::--------------------------------------------------------
:removeSlash
    set varIn="%~1"
    call :deQuote varIn _varIn
    if "%_varIn:~-1%" equ "\" (set varOut="%_varIn:~0,-1%") else (set varOut=%varIn%)
    set %2=%varOut%
goto:eof

::--------------------------------------------------------
::-- LTrim
::--------------------------------------------------------
:LTrim
    set varIn=%~1
    set vWorkVal=%varIn:~0,1%

    if "%vWorkVal%" neq " " (
        set varOut=%varIn%
    ) else (
        set varOut=%varIn:~1%
    )
    set %2=%varOut%
goto:eof

::--------------------------------------------------------
::-- RTrim
::--------------------------------------------------------
:RTrim
if /i "%DEBUG%" equ "Yes" (@echo RTrim - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set varIn=%~1

    if "%varIn:~-1%" equ " " (set varOut="%varIn:~0,-1%") else (set varOut=%varIn%)

    if /i "%DEBUG%" equ "Yes" (
        @echo varOut %varOut%
        pause
    )

    set %2=%varOut%

if /i "%DEBUG%" equ "Yes" (@echo RTrim - End)
if /i "%DEBUG%" equ "Yes" (pause)
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

    if /i [%1] equ [] (set sAllOK=No) else (set sAllOK=%1)

    if /i "%sAllOK%" equ "Yes" (
        call :Logging "%SEPERATOR2%"
        call :Logging "Done ..."
        call :Logging "%SEPERATOR2%"
    ) else (
        call :Logging "Errors?"
    )
    if /i "%Q_ALLOW_PAUSE%" equ "Yes" (
        pause
    ) else (
        timeout /T %iTimeOut%
    )

    popd
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
:: G Bishop  | 20-08-2018 | 01.01.00 | getFldrINC() had wrong length of 12 - whereas TASKS are 13 long
:: G Bishop  | 24-08-2018 | 01.02.00 | Moved: makeSubFldrs\Archive to common and screen shots to Citi
:: G Bishop  | 10-05-2019 | 01.03.00 | Added: BFC and redeveloped
:: G Bishop  | 23-05-2019 | 01.04.00 | Added: preSetup, and included APP in subject
:: G Bishop  | 14/06/2019 | 01.05.00 | Added: timout instead of pause
:: G Bishop  | 03/07/2019 | 01.06.00 | Added: screen dimentions
:: G Bishop  | 29/07/2019 | 01.07.00 | Amended: Tweaking of screen dimentions
:: G Bishop  | 29/07/2019 | 01.07.01 | Bugfix: renameFile() has %1 twice
:: G Bishop  | 19-08-2019 | 01.08.00 | Amended: for both CRM365 & 7CRM := CRM and fixed dynamic height
:: G Bishop  | 19-08-2019 | 01.09.00 | Added: optional exit message
:: G Bishop  | 09/09/2019 | 01.10.00 | Amended: QlikView folder layout
:: G Bishop  | 16/09/2019 | 01.11.00 | Added: debug info
:: G Bishop  | 17/09/2019 | 01.12.00 | Added: Replicate()
:: G Bishop  | 11/10/2019 | 01.13.00 | Added: QlikView & changed setColour
:: G Bishop  | 24-10-2019 | 01.14.00 | Added: new setColour
:: G Bishop  | 01/12/2019 | 01.15.00 | Amended: to create screenshots no matter the app
:: G Bishop  | 06/12/2019 | 01.16.00 | Amended: Will also copy: UserRightsNT.xlsx into \Archive & moved sections around
:: G Bishop  | 06/12/2019 | 02.01.00 | Amended: Split out TASK and Incidents
:: G Bishop  | 06/12/2019 | 02.02.00 | Amended: Height for QlikView
:: G Bishop  | 11/12/2019 | 02.03.00 | Amended: APP now allows spaces
:: G Bishop  | 15/02/2020 | 02.03.01 | Bugfix: Sorted location of ArchiveFiles.vbs.lnk
:: G Bishop  | 15/02/2020 | 02.03.02 | Amended: Sorted Versioning + new debug sections
:: G Bishop  | 04/03/2020 | 02.03.03 | Added: quotes where applicable
:: G Bishop  | 03/04/2020 | 02.04.00 | Amended: setColour()
:: G Bishop  | 24/04/2020 | 02.04.01 | Added: DEBUG to required routines, renamed Q_ALLOW_PAUSE
:: G Bishop  | 24/04/2020 | 02.05.00 | Added: Add icons onto folders dynamically
:: G Bishop  | 24-04-2020 | 02.05.01 | Bugfix: for some reason it was updating the files before getting them
:: G Bishop  | 24-04-2020 | 02.05.02 | Amended: Will apply icon if folder exists
:: G Bishop  | 27/04/2020 | 02.06.00 | Amended: Will apply icon to base folder
:: G Bishop  | 27/04/2020 | 02.06.01 | Amended: made adding CreateUserTemplate optional
:: G Bishop  | 27/04/2020 | 02.06.02 | Added: Q_FOLDER_ICON - made placing icon on root folder optional
:: G Bishop  | 11-06-2020 | 02.06.03 | Amended: for QlikView & tookout Maximo
:: G Bishop  | 09-07-2020 | 02.07.00 | Replaced: with getFirstXChars(); measureName(), getFldrNameTASK(), getFldrNameINC()
:: G Bishop  | 28-08-2020 | 02.07.01 | Removed: doActivitiesTask() - deleting: ThingsToDo.{App}{Ref}.xlsx
:: G Bishop  | 28-08-2020 | 02.07.02 | Amended: getFilesCommon() don't copy ThingsToDo.{App}{Ref}.xlsx if already exists
:: G Bishop  | 31-08-2020 | 02.07.03 | Amended: will only rename ThingsToDo.{App}{Ref}.xlsx if exists
:: G Bishop  | 31-08-2020 | 02.07.04 | Amended: Renamed: Screen shots
:: G Bishop  | 18-11-2020 | 02.08.00 | Amended: getFilesCommon() Changed location of ThingsToDo.{App}{Ref}.xlsx
:: G Bishop  | 10-05-2021 | 02.08.01 | Amended: Not implemented, measure path for screen width - later
:: G Bishop  | 07-07-2021 | 02.09.00 | Amended: Works out from calling folders name
:: G Bishop  | 16-12-2021 | 02.10.00 | Amended: Now uses _Archive
:: G Bishop  | 20-12-2021 | 02.11.00 | Amended: Works out the path it file is in
:: G Bishop  | 20-12-2021 | 02.11.01 | Amended: Now works from a central location
:: G Bishop  | 22-12-2021 | 02.11.02 | Amended: Uses placeIconOnFolderBuiltIn
:: G Bishop  | 22-12-2021 | 02.12.00 | Amended: Measure array for screen dimentions
:: G Bishop  | 09-04-2022 | 02.13.00 | Amended: Dont report missing if already accounted as absent [boolTurnOffErrors]
:: G Bishop  | 16-06-2022 | 02.13.01 | Amended: Renamed: 'Screen Shots' to: 'Images'
:: G Bishop  | 16-06-2022 | 02.13.02 | Amended: calculateHeight() corrected height caluclations, althogh buffer needs to be 8 to work correctly
:: G Bishop  | 27-08-2022 | 02.13.03 | Bugfix: Removed debugging message
:: G Bishop  | 31-08-2022 | 03.01.00 | Amended: Paths for Home and Harbour Energy
:: G Bishop  | 31-08-2022 | 03.02.00 | Amended: Handle both REQ and INC
:: G Bishop  | 15-02-2023 | 03.03.00 | Added: CHG - folder and title
:: G Bishop  | 15-02-2023 | 03.04.00 | Amended: nameTTD() Added: LTrim to extract the title part of the file name
:: G Bishop  | 20-02-2023 | 03.04.01 | Amended: Completed, instr() return only 1st occurrence
:: G Bishop  | 20-02-2023 | 03.04.02 | Amended: placeIconOnFolderBuiltIn()
:: G Bishop  | 22-02-2023 | 03.04.03 | Amended: setStandardFilesGlobal() corrected FOLDER_ARCHIVE for HE-5H9YZHMREDIB
:: G Bishop  | 27-02-2023 | 03.05.00 | Added: ISB for support work that isnt a call or project
:: G Bishop  | 28-02-2023 | 03.05.01 | Added: function :specialCheck() that puts hyphen back into the name
:: G Bishop  | 01-03-2023 | 03.05.02 | Amended: createFolderCaller() Check type of request
:: G Bishop  | 20-03-2023 | 03.05.03 | Amended: Removed special check for ISBAU12
:: G Bishop  | 04-04-2023 | 03.05.04 | Bugfix: Introduced RTrim for spaces causing error - needs more work
:: G Bishop  | 27-04-2023 | 03.05.05 | Bugfix: Handle if no folders are present
:: G Bishop  | 11-05-2023 | 03.05.06 | Amended: Moved referenced files under \Documents i.e. \Documents\Development\
:: G Bishop  | 17-05-2023 | 03.05.07 | Amended: better name for varible that's making and switching into a sub folder
:: G Bishop  | 04-07-2023 | 04.00.00 | Amended: Reads input file
:: G Bishop  | 10-08-2023 | 04.00.01 | Amended: Minor modifications, added current device to Global var Q_DEVICE
:: G Bishop  | 16-11-2023 | 04.01.00 | Amended: Templates location
:: G Bishop  | 25-11-2023 | 04.01.01 | Renamed: removeSlash
:: G Bishop  | 23-01-2024 | 04.02.00 | Amended: Now using OneDrive to pushd to
:: G Bishop  | 30-01-2024 | 04.02.01 | Bugfix: Get to run correctly when called as admin
:: G Bishop  | 13-05-2024 | 04.03.00 | Added: Problem numbering and filing to ..\Service Now\Problem
:: G Bishop  | 29-10-2024 | 04.03.01 | Amended: setStandardIcons(Use %windir%)
:: G Bishop  | 02-12-2024 | 04.04.00 | Removed: specialCheck()
:: G Bishop  | 28-01-2025 | 04.05.00 | Amended: createFolderCaller(Always filed by category/subfolder apart from PROBLEMS)
::-----------+------------+----------+--------------------
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
