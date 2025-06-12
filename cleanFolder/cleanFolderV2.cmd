set DEBUG=No
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

call :setUpScreen

call :cleanDownFoldersCaller "C:\Work\2022\Folder 01"
call :cleanDownFoldersCaller "C:\Work\2022\Folder 02"
call :cleanDownFoldersCaller "C:\Work\2022\Folder 03"
call :cleanDownFoldersCaller "C:\Work\2022\Folder 04"
call :cleanDownFoldersCaller "C:\Work\2022"


:: call :cleanDownFoldersCaller X:\Amulet\ScheduledTasks\C3\Schedules\C3 Auto Calculator\LogBackups\2021
if /i "%DEBUG%" equ "Yes" (pause)
:: call :cleanDownFoldersCaller X:\Amulet\ScheduledTasks\C3\Schedules\C3 Auto Calculator\LogBackups\2022
:: call :cleanDownFoldersCaller X:\Amulet\ScheduledTasks\C3\Schedules\C3 Auto Calculator\LogBackups\2023


call :myExit

::--------------------------------------------------------
::-- deletePatchCaller
::--------------------------------------------------------
:cleanDownFoldersCaller
if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersCaller - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set PATHNAME=%~1
    call :setDOSPointer "%PATHNAME%"

    for /F %%i in ('dir /b "%PATHNAME%\*.*"') do (
        call :deleteContentFnc "%PATHNAME%"
    )
    
    cd ..
    rd /s /q "%PATHNAME%"

if /i "%DEBUG%" equ "Yes" (@echo cleanDownFoldersCaller - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- deleteContentFnc
::--------------------------------------------------------
:deleteContentFnc
if /i "%DEBUG%" equ "Yes" (@echo deleteContentFnc - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    set PATHNAME=%~1

    call :setDOSPointer "%PATHNAME%"

    for /d %%i in ("%PATHNAME%\*") do (call :cleanDownFoldersNew "%%i")
    if /i "%DEBUG%" equ "Yes" (pause)


if /i "%DEBUG%" equ "Yes" (@echo deleteContentFnc - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- Clean down folders
::--------------------------------------------------------
:cleanDownFoldersNew
    set varpath=%~1

    @echo "%varpath%"
    if /i "%DEBUG%" equ "Yes" (pause)

    @echo Removing attributes on: %varpath%
    attrib -h -r -s /s "%varpath%\*.*"
    if /i "%DEBUG%" equ "Yes" (pause)
    del /f /q "%varpath%\*"
    if /i "%DEBUG%" equ "Yes" (pause)
    del "%varpath%\*.*" /q
    if /i "%DEBUG%" equ "Yes" (pause)
    rd /s /q "%varpath%"
    if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setDOSPointer - change drive and folder
::--------------------------------------------------------
:setDOSPointer
if /i "%DEBUG%" equ "Yes" (@echo setDOSPointer - Start)
if /i "%DEBUG%" equ "Yes" (pause)

    call :getDriveLtr "%~1" SOURCEDRIVE
    
    if /i "%DEBUG%" equ "Yes" (@echo SOURCEDRIVE %SOURCEDRIVE%)
    if /i "%DEBUG%" equ "Yes" (pause)

    %SOURCEDRIVE%
    cd "%~1"

if /i "%DEBUG%" equ "Yes" (@echo setDOSPointer - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- getDriveLtr
::--------------------------------------------------------
:getDriveLtr
if /i "%DEBUG%" equ "Yes" (@echo getDriveLtr - Start)
if /i "%DEBUG%" equ "Yes" (pause)

     set "%~2=%~d1"

if /i "%DEBUG%" equ "Yes" (@echo getDriveLtr - End)
if /i "%DEBUG%" equ "Yes" (pause)
goto:eof

::--------------------------------------------------------
::-- setUpScreen
::-- sets the screen up
::--------------------------------------------------------
:setUpScreen
    Title eCas upgrade to: %VERNO%
    color 1b
    mode con: cols=200 lines=333
goto:eof

::--------------------------------------------------------
::-- myExit
::--------------------------------------------------------
:myExit
    @echo "Done ..."
    pause
    exit
goto:eof

