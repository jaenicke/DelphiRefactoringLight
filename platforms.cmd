@echo off
:: ============================================================================
:: Which platforms does this package get built and registered for?
:: ============================================================================
:: Called as:  call platforms.cmd <bds-version>
:: Answers:    PLATFORMS      the -platforms list for delinst.exe
::             PLATFORMNOTE   why Win64 was left out, when it was
::
:: A design-time package is built for the IDE it is loaded INTO: a Win32 BPL
:: for the 32-bit IDE, a Win64 BPL for the 64-bit one. So Win64 is only worth
:: a round when that installation HAS a 64-bit IDE, and it is only POSSIBLE
:: when it also ships the Win64 design-time .dcp our .dpk requires
:: ("requires rtl, vcl, designide") - without it the Win64 round fails at the
:: compiler, which is what a version without a 64-bit IDE produces.
::
:: READ FROM THE INSTALLATION, never from a version table: "Delphi 12 has a
:: 64-bit IDE" was exactly such a remembered claim, it was wrong, and it broke
:: the Delphi 12 install. The two probes below answer per installation, so a
:: 12.x that does ship both is served as well as 13 is.
set PLATFORMS=
set PLATFORMNOTE=
set "_PFVER=%~1"
if not defined _PFVER goto :pf_noversion

set "_PFROOT="
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Embarcadero\BDS\%_PFVER%" /v RootDir 2^>nul') do set "_PFROOT=%%b"
if not defined _PFROOT goto :pf_noroot

:: Win32 is in every Delphi; only the second half is a question.
set PLATFORMS=Win32
if not exist "%_PFROOT%bin64\bds.exe" (
    set "PLATFORMNOTE=this installation has no 64-bit IDE, so there is nothing a Win64 BPL could be loaded into"
    goto :pf_done
)
if not exist "%_PFROOT%lib\win64\release\designide.dcp" (
    set "PLATFORMNOTE=this installation ships no Win64 designide.dcp, so a Win64 design-time package cannot be compiled here"
    goto :pf_done
)
set PLATFORMS=Win32,Win64
goto :pf_done

:pf_noversion
:: No version to look up - leave the caller's own default alone.
set "PLATFORMNOTE=no version given, so the platforms were not derived"
goto :pf_done

:pf_noroot
:: builddih.cmd reports a version that is not installed, and it words it
:: better than this would; answering Win32 keeps this script out of that.
set PLATFORMS=Win32
set "PLATFORMNOTE=BDS %_PFVER% has no RootDir in the registry"
goto :pf_done

:pf_done
set "_PFVER="
set "_PFROOT="
goto :eof
