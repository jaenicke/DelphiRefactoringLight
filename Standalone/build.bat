@echo off
rem Standalone Refactoring Light - msbuild driver.
rem
rem Override the config from the command line: build.bat Release
rem Default config is Debug to match the .dproj's $(Config) default.
call "C:\Program Files (x86)\Embarcadero\Studio\37.0\bin\rsvars.bat"
cd /d "%~dp0"
if not exist DCU mkdir DCU
if not exist Output mkdir Output
set CONFIG=%1
if "%CONFIG%"=="" set CONFIG=Debug
msbuild RefactoringLightStandalone.dproj /t:Build /p:Config=%CONFIG% /p:Platform=Win32 /p:DCC_ExeOutput=.\Output /p:DCC_DcuOutput=.\DCU /v:minimal
echo.
echo Exit Code: %ERRORLEVEL%
