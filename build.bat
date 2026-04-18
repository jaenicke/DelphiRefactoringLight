@echo off
call "C:\Program Files (x86)\Embarcadero\Studio\37.0\bin\rsvars.bat"
cd /d "%~dp0"
if not exist DCU mkdir DCU
if not exist Output2 mkdir Output2
msbuild Packages\DelphiRefactoringLight.dproj /t:Build /p:Config=Debug /p:Platform=Win32 /p:DCC_BplOutput=.\Output2 /v:minimal
echo.
echo Exit Code: %ERRORLEVEL%
