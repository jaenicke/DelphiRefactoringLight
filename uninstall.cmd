@echo off
setlocal

:: ============================================================================
:: Delphi Refactoring Light - Uninstall Script
:: Builds delinst.exe (DIH), then uses it to uninstall the package
:: ============================================================================

:: The Delphi version to install into, named as its registry branch is
:: (Delphi 13 = 37.0, Delphi 12 = 23.0, Delphi 11 = 22.0). The _D11 /
:: _D12 scripts beside this one pass theirs and change nothing else;
:: without an argument it stays Delphi 13.
set BDSVER=%~1
if not defined BDSVER set BDSVER=37.0
:: Which IDEs get a BPL. A second argument pins it ("Win32"); with none
:: it is DERIVED from that installation below - see platforms.cmd, which
:: says why a version number cannot answer this.
set PLATFORMS=%~2
set DELINST=%~dp0delinst.exe
set DIHCONFIG=%~dp0DelphiRefactoringLight.xml
set USEBDS=

:: Build delinst.exe first (also detects if cmd compiler is available)
call "%~dp0dih\builddih.cmd" %BDSVER%
if %ERRORLEVEL% NEQ 0 goto :error

:: builddih.cmd creates .dih_usebds marker if bds.exe was used
if exist "%~dp0.dih_usebds" set USEBDS=-usebds

:: Now that the version is known to be installed, ask it which platforms
:: it can take. A pinned list is left alone.
if not defined PLATFORMS call "%~dp0platforms.cmd" %BDSVER%
if defined PLATFORMNOTE echo Note: %PLATFORMNOTE%

echo.
echo ============================================
echo  Delphi Refactoring Light - Uninstall (BDS %BDSVER%, %PLATFORMS%)
echo ============================================
echo.

"%DELINST%" %BDSVER% -config "%DIHCONFIG%" -platforms %PLATFORMS% -configs Release -action uninstall -verbose log %USEBDS%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Uninstall completed with errors. See DelphiRefactoringLight.log for details.
    goto :error
)

echo.
echo Uninstall completed successfully.
goto :end

:error
echo.
pause
exit /b 1

:end
pause
exit /b 0
