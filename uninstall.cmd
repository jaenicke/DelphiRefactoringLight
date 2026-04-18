@echo off
setlocal

:: ============================================================================
:: Delphi Refactoring Light - Uninstall Script
:: Builds delinst.exe (DIH), then uses it to uninstall the package
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
echo  Delphi Refactoring Light - Uninstall (BDS %BDSVER%)
echo ============================================
echo.

"%DELINST%" %BDSVER% -config "%CONFIG%" -platforms Win32 -configs Release -action uninstall -verbose log %USEBDS%

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
