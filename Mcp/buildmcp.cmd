@echo off
setlocal

:: ============================================================================
:: Builds the MCP bridge RefactoringLightMcp.exe and installs it to
:: %LOCALAPPDATA%\DelphiRefactoringLight\mcp\ - a fixed path, so the Claude
:: Code registration never has to change. Called by install.cmd.
::
:: A running Claude Code session keeps the exe open. Windows lets a running
:: exe be RENAMED, though, so the old one is moved aside first and the new
:: one copied in.
::
:: THE ASIDE NAME IS UNIQUE (2026-10-04). It used to be one fixed
:: "<exe>.old", deleted before the move - and a bridge process started FROM
:: that .old holds it, so the delete failed, the move had nowhere to go and
:: the copy could not overwrite the live exe. Then NOTHING was installed
:: although the build succeeded, which is how the user's bridge stayed four
:: versions behind. With a unique name the move always has a free
:: destination, so an install works WHILE Claude Code sessions are running -
:: which is the normal case, because install.cmd is run from one.
:: ============================================================================

set BDSVER=%~1
if "%BDSVER%"=="" set BDSVER=37.0
set SCRIPTDIR=%~dp0
set DPROJ=%SCRIPTDIR%RefactoringLightMcp.dproj
set GROUPPROJ=%SCRIPTDIR%RefactoringLightMcp.groupproj
set CLOSEDLG=%SCRIPTDIR%..\dih\closedialog.ps1
set BUILT=%SCRIPTDIR%Bin\RefactoringLightMcp.exe
set TARGETDIR=%LOCALAPPDATA%\DelphiRefactoringLight\mcp
set TARGET=%TARGETDIR%\RefactoringLightMcp.exe

for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Embarcadero\BDS\%BDSVER%" /v RootDir 2^>nul') do set "BDSROOT=%%b"
if "%BDSROOT%"=="" (
    echo MCP bridge: BDS %BDSVER% not found in the registry - skipped.
    exit /b 1
)

echo.
echo Building the MCP bridge (RefactoringLightMcp.exe) ...
call "%BDSROOT%bin\rsvars.bat" 2>nul
set "MCP_LOG=%TEMP%\refactoringlight_mcp_build.log"
msbuild "%DPROJ%" /t:Build /p:Platform=Win32 /p:Config=Release /v:m /nologo > "%MCP_LOG%" 2>&1
set MSBUILD_ERR=%ERRORLEVEL%
:: No command-line compiler (Community Edition) - build with the IDE, the
:: same fallback dih\builddih.cmd has. Without it the bridge was simply
:: never built there and every MCP tool was missing, while install.cmd
:: told the user to open the project by hand (user report 2026-10-05).
findstr /i /c:"does not support command line" "%MCP_LOG%" >nul 2>&1
if %ERRORLEVEL% EQU 0 goto :try_bds
if %MSBUILD_ERR% NEQ 0 goto :msbuild_failed
if not exist "%BUILT%" goto :no_exe
goto :install

:msbuild_failed
type "%MCP_LOG%"
echo MCP bridge: build FAILED - see above.
exit /b 1

:no_exe
echo MCP bridge: build produced no exe.
exit /b 1

:try_bds
if not exist "%BDSROOT%bin\bds.exe" goto :no_compiler
echo MCP bridge: no command-line compiler, falling back to bds.exe ...
:: MSBuild - and the IDE's own project builder - reads EVERY environment
:: variable as a property, and nothing on the bds command line can correct
:: it, so an inherited Config / Platform would decide this build.
set "Config=Release"
set "Platform=Win32"
set "DLGPID=%TEMP%\mcp_closedialog_%RANDOM%.pid"
if exist "%CLOSEDLG%" start /b powershell -NoProfile -ExecutionPolicy Bypass -File "%CLOSEDLG%" -ProcessName bds -PidFile "%DLGPID%" 2>nul
"%BDSROOT%bin\bds.exe" -b -ns "%GROUPPROJ%"
:: FIRST the exit code, THEN the call - a call resets ERRORLEVEL.
set BDS_ERR=%ERRORLEVEL%
call :stop_watcher
if %BDS_ERR% NEQ 0 goto :bds_failed
if not exist "%BUILT%" goto :no_exe
goto :install

:bds_failed
echo MCP bridge: the IDE build FAILED.
if exist "%SCRIPTDIR%RefactoringLightMcp.err" type "%SCRIPTDIR%RefactoringLightMcp.err"
exit /b 1

:no_compiler
echo MCP bridge: neither a command-line compiler nor bds.exe was found.
exit /b 1

:install

if not exist "%TARGETDIR%" mkdir "%TARGETDIR%"
:: Best effort only: an .old a bridge still runs from cannot be deleted, and
:: that must never stop the install (it used to).
del /f /q "%TARGETDIR%\*.old" >nul 2>&1
:: One line on purpose: %RANDOM% is expanded where it stands, so no delayed
:: expansion is needed (plain "setlocal" above does not have it).
if exist "%TARGET%" move /y "%TARGET%" "%TARGET%.%RANDOM%%RANDOM%.old" >nul 2>&1
copy /y "%BUILT%" "%TARGET%" >nul
if %ERRORLEVEL% NEQ 0 goto :notinstalled

:: VERIFY - a copy that reports success is NOT proof (2026-09-30): the exe
:: installed here had been nine days old while every build succeeded, because
:: the rename-aside failed and the copy could not overwrite the live exe. So
:: ASK THE INSTALLED EXE what it is and compare it with the built one; only
:: that distinguishes "installed" from "still the old file".
set "BUILTVER="
set "TARGETVER="
for /f "tokens=2" %%v in ('"%BUILT%" --version 2^>nul') do set "BUILTVER=%%v"
for /f "tokens=2" %%v in ('"%TARGET%" --version 2^>nul') do set "TARGETVER=%%v"
:: Both empty compares EQUAL, so an exe that cannot start at all (a missing
:: dependency, a truncated copy) used to report "installed" (audit #36, L2f).
:: An answer from the BUILT exe is the precondition for comparing anything.
if not defined BUILTVER goto :notinstalled
if not "%BUILTVER%"=="%TARGETVER%" goto :notinstalled

echo MCP bridge installed: %TARGET% (%TARGETVER%)
exit /b 0

:notinstalled
echo.
echo ****************************************************************
echo  MCP bridge NOT updated - the old exe is still installed.
echo ****************************************************************
echo.
echo  built    : %BUILTVER%  %BUILT%
if defined TARGETVER echo  installed: %TARGETVER%  %TARGET%
if not defined TARGETVER echo  installed: the exe there did not answer --version
echo.
echo  The file could not be replaced - and a running Claude Code session is
echo  NOT a reason any more (the exe is moved aside under a unique name).
echo  Check that %TARGETDIR%
echo  is writable, then run install.cmd again.
exit /b 1

:stop_watcher
:: Best effort, and by PID only - never by image name: a developer has
:: other powershell processes, and this one may already have ended.
if not defined DLGPID exit /b 0
if exist "%DLGPID%" for /f "usebackq tokens=1" %%p in ("%DLGPID%") do taskkill /f /pid %%p >nul 2>&1
if exist "%DLGPID%" del /f /q "%DLGPID%" >nul 2>&1
exit /b 0
