set DEBUG=Yes
if /i "%DEBUG%" equ "Yes" (@echo on) else (@echo off)

call :setUpScreen


call :cleanDownFoldersCaller C:\found.000
if /i "%DEBUG%" equ "Yes" (pause)

call :MyExit

::--------------------------------------------------------
::-- deletePatchCaller
::--------------------------------------------------------
:cleanDownFoldersCaller
    set PATHNAME=%1
    call :setDOSPointer %PATHNAME%

    for /F %%i in ('dir /b "%PATHNAME%\*.*"') do (
        call :deleteContentFnc %PATHNAME%
    )

goto:eof

::--------------------------------------------------------
::-- deleteContentFnc
::--------------------------------------------------------
:deleteContentFnc
    set PATHNAME=%1

    call :setDOSPointer %PATHNAME%

    for /d %%i in (%PATHNAME%\*) do (call :cleanDownFoldersNew %%i)
    if /i "%DEBUG%" equ "Yes" (pause)

goto:eof

::--------------------------------------------------------
::-- Clean down folders
::--------------------------------------------------------
:cleanDownFoldersNew
    set varpath=%1

    @echo %varpath%
    if /i "%DEBUG%" equ "Yes" (pause)

    @echo Removing attributes on: %varpath%
    attrib -h -r -s /s %varpath%\*.*
    if /i "%DEBUG%" equ "Yes" (pause)
    del /f /q %varpath%\*
    if /i "%DEBUG%" equ "Yes" (pause)
    del %varpath%\*.* /q
    if /i "%DEBUG%" equ "Yes" (pause)
    rd /s /q %varpath%
    if /i "%DEBUG%" equ "Yes" (pause)
goto:eof


::--------------------------------------------------------
::-- setDOSPointer - change drive and folder
::--------------------------------------------------------
:setDOSPointer
    call :getDriveLtr SOURCEDRIVE %1
    %SOURCEDRIVE%
    cd %1
goto:eof

::--------------------------------------------------------
::-- getDriveLtr
::--------------------------------------------------------
:getDriveLtr
    set str=%2
    set str=%str:~0,2%
    set "%~1=%str%"
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
::-- MyExit
::--------------------------------------------------------
:MyExit
    @echo "Done ..."
    pause
exit

::--------------------------------------------------------
::-- Clean down folders
::--------------------------------------------------------
:cleanDownFolders
    set varpath=%1
    if exist %varpath% (
        attrib -h -r -s /s %varpath%\*.*
        for /d %%i in (%varpath%\*) do rd /s /q "%%i"
        del %varpath%\*.* /q
        rd /s /q %varpath%
    )
goto:eof
