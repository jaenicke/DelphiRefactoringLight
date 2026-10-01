@echo off
setlocal

:: ============================================================================
:: Builds the MCP bridge RefactoringLightMcp.exe and installs it to
:: %LOCALAPPDATA%\DelphiRefactoringLight\mcp\ - a fixed path, so the Claude
:: Code registration never has to change. Called by install.cmd.
::
:: A running Claude Code session keeps the exe open. Windows lets a running
:: exe be RENAMED, though, so the old one is moved aside first and the new
:: one copied in; the old file is removed by the next install.
:: ============================================================================

set BDSVER=%~1
if "%BDSVER%"=="" set BDSVER=37.0
set SCRIPTDIR=%~dp0
set DPROJ=%SCRIPTDIR%RefactoringLightMcp.dproj
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
if %ERRORLEVEL% NEQ 0 (
    findstr /i /c:"does not support command line" "%MCP_LOG%" >nul 2>&1
    if not errorlevel 1 (
        echo MCP bridge: this Delphi edition has no command-line compiler.
        echo Open Mcp\RefactoringLightMcp.dproj in the IDE and build it there,
        echo then run install.cmd again.
    ) else (
        type "%MCP_LOG%"
        echo MCP bridge: build FAILED - see above.
    )
    exit /b 1
)
if not exist "%BUILT%" (
    echo MCP bridge: build produced no exe.
    exit /b 1
)

if not exist "%TARGETDIR%" mkdir "%TARGETDIR%"
if exist "%TARGET%.old" del /f /q "%TARGET%.old" >nul 2>&1
if exist "%TARGET%" (
    move /y "%TARGET%" "%TARGET%.old" >nul 2>&1
)
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
echo  The file could not be replaced. Almost always a RUNNING Claude Code
echo  session (or a "claude" in a terminal) still holds it.
echo  Close every one of them, then run install.cmd again.
exit /b 1
