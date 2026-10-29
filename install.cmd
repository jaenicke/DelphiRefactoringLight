@echo off
setlocal

:: ============================================================================
:: Delphi Refactoring Light - Install Script
:: Builds delinst.exe (DIH), then uses it to build and install the package
:: ============================================================================

:: The Delphi version to install into, named as its registry branch is
:: (Delphi 13 = 37.0, Delphi 12 = 23.0, Delphi 11 = 22.0). The _D11 /
:: _D12 scripts beside this one pass theirs and change nothing else;
:: without an argument it stays Delphi 13.
set BDSVER=%~1
if not defined BDSVER set BDSVER=37.0
:: Which IDEs get a BPL. Delphi 12 was the first with a 64-bit IDE
:: (bin64ds.exe) and its own " x64" registry branch, so Delphi 11
:: is Win32 alone - its wrapper passes that.
set PLATFORMS=%~2
if not defined PLATFORMS set PLATFORMS=Win32,Win64
set DELINST=%~dp0delinst.exe
set DIHCONFIG=%~dp0DelphiRefactoringLight.xml
set USEBDS=

:: Build delinst.exe first (also detects if cmd compiler is available)
call "%~dp0dih\builddih.cmd" %BDSVER%
if %ERRORLEVEL% NEQ 0 goto :error

:: builddih.cmd creates .dih_usebds marker if bds.exe was used
if exist "%~dp0.dih_usebds" set USEBDS=-usebds

echo.
echo ============================================
echo  Delphi Refactoring Light - Install (BDS %BDSVER%)
echo ============================================
echo.

"%DELINST%" %BDSVER% -config "%DIHCONFIG%" -platforms %PLATFORMS% -configs Release -action install -verbose log %USEBDS%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Installation completed with errors. See DelphiRefactoringLight.log for details.
    echo The summary above says WHICH entry failed - the package, the MCP
    echo bridge for Claude Code, or both.
    goto :error
)

echo.
echo Installation completed successfully.
echo See DelphiRefactoringLight.log for details.

:: The bridge is an entry of the installer now (it builds it, installs it
:: next to the package and verifies that the installed exe answers), so its
:: result is a row of the summary above. What follows is only what the
:: installer cannot do: registering it in Claude Code, which is the user's
:: own configuration.
set MCPEXE=%LOCALAPPDATA%\DelphiRefactoringLight\mcp\RefactoringLightMcp.exe

echo.
echo ============================================
echo  Claude Code: MCP bridge
echo ============================================
echo.
echo  The bridge gives Claude Code the IDE's diagnostics and Refactoring
echo  Light's quick fixes - for every running RAD Studio instance.
echo  Claude Code starts the bridge itself; it only has to be registered once.
echo.
set MCPREG=
where claude >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto :mcp_steps
call claude mcp get delphi-refactoring-light >nul 2>&1
if %ERRORLEVEL% EQU 0 set MCPREG=1

:mcp_steps
if defined MCPREG goto :mcp_registered
echo  1. Register the bridge - EITHER once for every Claude Code session:
echo.
echo       claude mcp add --scope user delphi-refactoring-light -- "%MCPEXE%"
echo.
echo     OR only for one Delphi project (run it in that project folder; keeps
echo     other, non-Delphi sessions free of it):
echo.
echo       claude mcp add --scope local delphi-refactoring-light -- "%MCPEXE%"
echo.
echo     The name is up to you - it prefixes every tool name. While no IDE
echo     runs, the bridge offers only two tools; the rest appear with the IDE.
echo.
echo  2. Restart RAD Studio (the IDE side is part of the package).
echo  3. Start Claude Code in your project folder and type /mcp -
echo     "delphi-refactoring-light" should show as connected.
echo  4. Optional check - which IDEs the bridge sees:
echo.
echo       "%MCPEXE%" --list
goto :restart_hint

:mcp_registered
echo  The bridge is already registered in Claude Code (delphi-refactoring-light).
echo  Running Claude Code sessions still use the previous bridge until you
echo  reconnect it with /mcp or start a new session.
echo  Check which IDEs the bridge sees:
echo.
echo       "%MCPEXE%" --list
:: This block used to fall through into a warning block below, which is
:: why every user with a registered bridge got "MCP BRIDGE IS NOT
:: installed" on a SUCCESSFUL install from 1.15.2 until 1.16.30. There is
:: nothing below it any more - the bridge's own result is a row of the
:: installer's summary - but the goto stays: a label block that ends
:: without one does whatever is appended next.
goto :restart_hint

:restart_hint
echo.
echo IMPORTANT: Restart RAD Studio so the IDE picks up the new package version.
goto :end

:error
echo.
pause
exit /b 1

:end
pause
exit /b 0
