@echo off
setlocal

:: ============================================================================
:: Delphi Refactoring Light - Install Script
:: Builds delinst.exe (DIH), then uses it to build and install the package
:: ============================================================================

set BDSVER=37.0
set DELINST=%~dp0delinst.exe
set CONFIG=DelphiRefactoringLight.xml
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

"%DELINST%" %BDSVER% -config "%CONFIG%" -platforms Win32 -configs Release -action install -verbose log %USEBDS%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Installation completed with errors. See DelphiRefactoringLight.log for details.
    goto :error
)

echo.
echo Installation completed successfully.
echo See DelphiRefactoringLight.log for details.
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
