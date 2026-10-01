(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.RestartHint;

interface

type
  /// <summary>Shows a restart hint after the package is (re-)installed
  ///  inside a running IDE session. See Check for the detection logic.</summary>
  TRestartHint = class
  private
    class function GetMarkerFilePath: string; static;
    class function ProcessStamp: string; static;
    class procedure PruneDeadMarkers; static;
    class procedure ShowRestartHint; static;
  public
    /// <summary>Checks whether the package was (re-)installed inside a
    ///  running IDE session and, if so, shows a restart hint. On a
    ///  normal IDE start (fresh process) no hint is shown.</summary>
    class procedure Check; static;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils,
  Winapi.Windows, Vcl.Forms, Vcl.Dialogs, System.UITypes;

{ TRestartHint }

class function TRestartHint.GetMarkerFilePath: string;
begin
  // ONE MARKER PER PROCESS (audit #40, L5j): a single shared file meant
  // IDE B overwrote IDE A's pid, so reinstalling in A found B's number and
  // showed no hint - and a stale marker whose pid Windows later reused
  // produced a hint nobody earned. The file is named after the pid and
  // holds the process's CREATION TIME, which is what tells "the same IDE
  // again" from "a new process that happens to have that pid".
  Result := TPath.Combine(TPath.GetTempPath,
    Format('DelphiRefactoringLight.%d.pid', [GetCurrentProcessId]));
end;

/// <summary>The current process's creation time as a stable string - two
///  processes can share a pid over time, never a pid plus a start
///  time.</summary>
class function TRestartHint.ProcessStamp: string;
var
  Created, Exited, Kernel, User: TFileTime;
begin
  Result := '';
  if GetProcessTimes(GetCurrentProcess, Created, Exited, Kernel, User) then
    Result := Format('%d-%d', [Created.dwHighDateTime, Created.dwLowDateTime]);
end;

/// <summary>Removes the markers of processes that are gone, so %TEMP% does
///  not collect one file per IDE start for ever.</summary>
class procedure TRestartHint.PruneDeadMarkers;
var
  Files: TArray<string>;
begin
  try
    Files := TDirectory.GetFiles(TPath.GetTempPath,
      'DelphiRefactoringLight.*.pid');
  except
    Exit;
  end;
  for var F in Files do
  begin
    var Name := TPath.GetFileNameWithoutExtension(F);   // ...Light.<pid>
    var Dot := LastDelimiter('.', Name);
    if Dot <= 0 then Continue;
    var Pid := StrToUIntDef(Copy(Name, Dot + 1, MaxInt), 0);
    if (Pid = 0) or (Pid = GetCurrentProcessId) then Continue;
    // A handle we cannot open at all, or one whose process has exited,
    // means the marker is stale.
    var H := OpenProcess(SYNCHRONIZE, False, Pid);
    var Dead := H = 0;
    if H <> 0 then
    begin
      Dead := WaitForSingleObject(H, 0) = WAIT_OBJECT_0;
      CloseHandle(H);
    end;
    if Dead then
      try TFile.Delete(F); except end;
  end;
end;

class procedure TRestartHint.ShowRestartHint;
begin
  // Delay the dialog so the IDE message pump is fully up.
  TThread.ForceQueue(nil,
    procedure
    begin
      MessageDlg(
        'Delphi Refactoring Light was just installed or updated.' + sLineBreak +
        sLineBreak +
        'IMPORTANT: Please restart RAD Studio so the expert works ' +
        'correctly.' + sLineBreak +
        sLineBreak +
        'Without a restart, access violations or unexpected behavior may occur, ' +
        'because the IDE still holds references to the previous package version.',
        mtWarning, [mbOK], 0);
    end);
end;

class procedure TRestartHint.Check;
var
  Stored: string;
  MarkerFile: string;
  MarkerExists: Boolean;
  ShouldHint: Boolean;
begin
  PruneDeadMarkers;
  MarkerFile := GetMarkerFilePath;
  MarkerExists := FileExists(MarkerFile);

  Stored := '';
  if MarkerExists then
  begin
    try
      Stored := Trim(TFile.ReadAllText(MarkerFile));
    except
      Stored := '';
    end;
  end;

  // Only hint on re-install within the running IDE session:
  //   - Marker exists AND stored PID equals the current PID
  //     (= this package was already loaded once in this session and is
  //        being loaded again -> IDE restart required).
  //
  // Missing marker means first load (e.g. after external install.cmd
  // while the IDE was closed). No restart needed in that case - the
  // IDE is starting fresh anyway.
  // The marker is already named after this pid, so what has to match is
  // the process START: the same IDE loading the package a second time.
  ShouldHint := MarkerExists and (Stored <> '') and (Stored = ProcessStamp);

  // Always update the marker - also on normal IDE starts.
  try
    TFile.WriteAllText(MarkerFile, ProcessStamp);
  except
    // Tolerate write errors silently.
  end;

  if ShouldHint then
    ShowRestartHint;
end;

end.
