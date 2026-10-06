@echo off
setlocal

:: ============================================================================
:: Builds delinst.exe from the dih directory
:: Tries msbuild first, falls back to bds.exe if msbuild is not available
:: Output goes to the parent directory (the example/project root)
:: ============================================================================

set BDSVER=%~1
if "%BDSVER%"=="" goto :no_version

set SCRIPTDIR=%~dp0
set DPROJ=%SCRIPTDIR%delinst.dproj
set GROUPPROJ=%SCRIPTDIR%delinst.groupproj
:: Output delinst.exe to the parent directory
set EXEDIR=%SCRIPTDIR%..

:: Read BDS root directory from registry
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Embarcadero\BDS\%BDSVER%" /v RootDir 2^>nul') do set "BDSROOT=%%b"

if "%BDSROOT%"=="" goto :no_bds

set "RSVARS=%BDSROOT%bin\rsvars.bat"
set "BDSEXE=%BDSROOT%bin\bds.exe"

echo.
echo Building delinst.exe ...
echo.

:: Try msbuild first, fall back to bds.exe if it fails
call "%RSVARS%" 2>nul

echo Trying msbuild ...
set "MSBUILD_LOG=%TEMP%\dih_msbuild.log"
:: /tv:4.0: without it MSBuild takes the machine's default toolset, and a
:: default of 2.0 needs .NET 3.5 for its task assemblies (MSB4036 on the
:: forum machine of 2026-10-06). Delphi points at .NET 4.x itself.
msbuild "%DPROJ%" /t:Build /tv:4.0 /p:Platform=Win32 /p:Config=Release /p:DCC_ExeOutput="%EXEDIR%\." /v:m /nologo > "%MSBUILD_LOG%" 2>&1
set MSBUILD_ERR=%ERRORLEVEL%

:: Check for "does not support command line compiling" in output
findstr /i /c:"does not support command line" "%MSBUILD_LOG%" >nul 2>&1
if %ERRORLEVEL% EQU 0 goto :try_bds

:: msbuild ran - show output and check result
type "%MSBUILD_LOG%"
if %MSBUILD_ERR% EQU 0 goto :check_result_msbuild

:try_bds
:: msbuild failed or no cmd compiler - try bds.exe
if not exist "%BDSEXE%" goto :no_compiler

echo.
echo msbuild not available, falling back to bds.exe ...
echo.
set "DIH_ExeOutput=%EXEDIR%\."
:: MSBuild - and the IDE's own project builder, which is what bds.exe -b
:: uses - reads EVERY environment variable as a property. An inherited
:: Config or Platform therefore decides this build: a caller that does
:: "set CONFIG=<config file>" makes DCC_DcuOutput ".\Win32\<that path>",
:: which cannot be created, and a Visual Studio prompt exports
:: Platform=x64, which no Delphi project has. The msbuild branch above
:: passes both with /p: and cannot be fooled; here only the environment
:: can say it, so pin both to what we are building.
set "Config=Release"
set "Platform=Win32"

:: Start background watcher to auto-close bds.exe save dialogs
:: The watcher shares THIS console, so it has to be stopped when the
:: build is over - an orphan keeps the window open after the final
:: pause (user report 2026-10-05). It writes its pid for us and has a
:: timeout of its own as a second net.
set "DLGPID=%TEMP%\dih_closedialog_%RANDOM%.pid"
start /b powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTDIR%closedialog.ps1" -ProcessName bds -PidFile "%DLGPID%" 2>nul

"%BDSEXE%" -b -ns "%GROUPPROJ%"
:: FIRST the exit code, THEN the call - a call resets ERRORLEVEL.
set BDS_ERR=%ERRORLEVEL%
call :stop_watcher
if %BDS_ERR% NEQ 0 goto :build_failed

:: Signal to calling scripts: no cmd compiler, use bds.exe
echo.>"%EXEDIR%\.dih_usebds"
if not exist "%EXEDIR%\delinst.exe" goto :exe_missing
echo delinst.exe built successfully.
exit /b 0

:check_result_msbuild
:: Signal to calling scripts: msbuild works
if exist "%EXEDIR%\.dih_usebds" del "%EXEDIR%\.dih_usebds" >nul 2>&1
if not exist "%EXEDIR%\delinst.exe" goto :exe_missing
echo delinst.exe built successfully.
exit /b 0

:no_version
echo ERROR: BDS version parameter required, e.g. builddih.cmd 37.0
exit /b 1

:no_bds
echo ERROR: BDS %BDSVER% not found in registry.
exit /b 1

:no_compiler
echo ERROR: msbuild failed and bds.exe not found.
exit /b 1

:build_failed
echo.
echo ERROR: Failed to build delinst.exe
exit /b 1

:exe_missing
echo ERROR: delinst.exe was not created
exit /b 1

:stop_watcher
:: Best effort, and by PID only - never by image name: a developer has
:: other powershell processes, and this one may already have ended.
if not defined DLGPID exit /b 0
if exist "%DLGPID%" for /f "usebackq tokens=1" %%p in ("%DLGPID%") do taskkill /f /pid %%p >nul 2>&1
if exist "%DLGPID%" del /f /q "%DLGPID%" >nul 2>&1
exit /b 0
