@echo off
setlocal

:: ============================================================================
:: Delphi Refactoring Light - Build Only (no IDE registration)
:: Builds delinst.exe (DIH), then builds the package without modifying the IDE
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
echo  Delphi Refactoring Light - Build Only (BDS %BDSVER%)
echo ============================================
echo.

"%DELINST%" %BDSVER% -config "%CONFIG%" -platforms Win32 -configs Release -entries DelphiRefactoringLight -action build -verbose log %USEBDS%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Build completed with errors. See DelphiRefactoringLight.log for details.
    goto :error
)

echo.
echo Build completed successfully.
goto :end

:error
echo.
pause
exit /b 1

:end
pause
exit /b 0
