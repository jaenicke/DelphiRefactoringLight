@echo off
:: Delphi 11 Alexandria is BDS 22.0 in the registry. Everything this does is
:: rebuild.cmd's - the version is the only difference, and which platforms
:: that installation can take is read from the installation itself.
call "%~dp0rebuild.cmd" 22.0
exit /b %ERRORLEVEL%
