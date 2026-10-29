@echo off
setlocal
rem Standalone Refactoring Light - msbuild driver.
rem
rem   build.bat [config] [delphi version] [installation directory]
rem
rem Default config is Debug to match the .dproj's $(Config) default, and the
rem default version is Delphi 13 (37.0) - pass 23.0 for Delphi 12 or 22.0 for
rem Delphi 11. The installation directory is read from that version's registry
rem branch; pass it as the third argument when the registry does not say.
set "CONFIG=%~1"
if not defined CONFIG set "CONFIG=Debug"
set "BDSVER=%~2"
if not defined BDSVER set "BDSVER=37.0"
set "RADSTUDIO=%~3"
if not defined RADSTUDIO for /f "tokens=2,*" %%A in ('reg query "HKCU\Software\Embarcadero\BDS\%BDSVER%" /v RootDir 2^>nul') do set "RADSTUDIO=%%B"
if not defined RADSTUDIO set "RADSTUDIO=%BDS%"
if not exist "%RADSTUDIO%\bin\rsvars.bat" (
  echo RAD Studio %BDSVER% not found in the registry. Pass its installation
  echo directory as the third argument, e.g. build.bat Debug %BDSVER% "D:\Studio"
  exit /b 1
)
call "%RADSTUDIO%\bin\rsvars.bat"
if errorlevel 1 exit /b 1
pushd "%~dp0"
if not exist DCU mkdir DCU
if not exist Output mkdir Output
msbuild RefactoringLightStandalone.dproj /t:Build /p:Config=%CONFIG% /p:Platform=Win32 /p:DCC_ExeOutput=.\Output /p:DCC_DcuOutput=.\DCU\Win32\%CONFIG% /v:minimal
set "BUILD_RESULT=%ERRORLEVEL%"
popd
echo.
echo Exit Code: %BUILD_RESULT%
exit /b %BUILD_RESULT%
