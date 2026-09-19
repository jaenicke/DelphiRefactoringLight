@echo off
rem Builds and runs the DUnitX tests (IDE-free, no DelphiLSP needed).
rem   run-tests.cmd            Win32 Debug
rem   run-tests.cmd Win64      Win64 Debug
setlocal
rem not PLATFORM: rsvars.bat clears that variable
set RL_PLATFORM=%1
if "%RL_PLATFORM%"=="" set RL_PLATFORM=Win32
call "C:\Program Files (x86)\Embarcadero\Studio\37.0\bin\rsvars.bat"
cd /d "%~dp0"
msbuild DelphiRefactoringLightTests.dproj /t:Build /p:Config=Debug /p:Platform=%RL_PLATFORM% /v:minimal
if errorlevel 1 exit /b 2
"%~dp0%RL_PLATFORM%\Debug\DelphiRefactoringLightTests.exe" --exitbehavior:Continue --xmlfile:"%~dp0%RL_PLATFORM%\Debug\results.xml"
exit /b %ERRORLEVEL%
