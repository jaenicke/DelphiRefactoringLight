@echo off
setlocal

:: ============================================================================
:: Delphi Refactoring Light - Install Script
:: Builds delinst.exe (DIH), then uses it to build and install the package
:: ============================================================================

set BDSVER=37.0
set DELINST=%~dp0delinst.exe
set CONFIG=%~dp0DelphiRefactoringLight.xml
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

"%DELINST%" %BDSVER% -config "%CONFIG%" -platforms Win32,Win64 -configs Release -action install -verbose log %USEBDS%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Installation completed with errors. See DelphiRefactoringLight.log for details.
    goto :error
)

echo.
echo Installation completed successfully.
echo See DelphiRefactoringLight.log for details.

:: MCP bridge for Claude Code (optional - a failure here does not fail the
:: package installation)
set MCPEXE=%LOCALAPPDATA%\DelphiRefactoringLight\mcp\RefactoringLightMcp.exe
call "%~dp0Mcp\buildmcp.cmd" %BDSVER%
if %ERRORLEVEL% NEQ 0 goto :mcp_failed

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
:: WITHOUT this goto the block FALLS THROUGH into :mcp_failed below.
:: That is what happened from 1.15.2 until 1.16.30: every user whose
:: bridge was already registered got the "MCP BRIDGE IS NOT installed"
:: warning on every SUCCESSFUL install.
goto :restart_hint

:mcp_failed
:: The package IS installed, so this must not read like a plain success - a
:: stale bridge exe is invisible from inside Claude Code (it only reports its
:: version to the MCP client, never to the IDE), and that cost two sessions of
:: chasing wrong timeouts on 2026-09-30.
echo.
echo ################################################################
echo  WARNING: the package is installed, the MCP BRIDGE IS NOT.
echo ################################################################
echo.
echo  Claude Code will keep talking to the OLD bridge exe, so its tools
echo  behave like the build that exe came from - wrong timeouts, missing
echo  tools - and nothing inside the IDE can tell you. The reason is in
echo  the lines above.
echo.
echo  A running Claude Code session is NOT the reason (since 1.22.1 the
echo  old exe is moved aside under a unique name), so read the lines above.
echo  To see what is installed at any time:
echo.
echo       "%MCPEXE%" --version
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
