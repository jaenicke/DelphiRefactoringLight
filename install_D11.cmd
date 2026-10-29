@echo off
:: Delphi 11 Alexandria is BDS 22.0 in the registry. Everything this does is
:: install.cmd's - this passes the version and, because Delphi 11 has no
:: 64-bit IDE, Win32 alone. There is one install to keep working.
call "%~dp0install.cmd" 22.0 Win32
exit /b %ERRORLEVEL%
