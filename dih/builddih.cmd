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
msbuild "%DPROJ%" /t:Build /p:Platform=Win32 /p:Config=Release /p:DCC_ExeOutput="%EXEDIR%\." /v:m /nologo > "%MSBUILD_LOG%" 2>&1
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

:: Start background watcher to auto-close bds.exe save dialogs
start /b powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTDIR%closedialog.ps1" -ProcessName bds 2>nul

"%BDSEXE%" -b -ns "%GROUPPROJ%"
if %ERRORLEVEL% NEQ 0 goto :build_failed

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
