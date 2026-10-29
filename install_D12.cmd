@echo off
:: Delphi 12 Athens is BDS 23.0 in the registry. Everything this does is
:: install.cmd's - the version is the only difference, so there is one place
:: where the install itself can change.
call "%~dp0install.cmd" 23.0
exit /b %ERRORLEVEL%
