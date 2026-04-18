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
  Result := TPath.Combine(TPath.GetTempPath, 'DelphiRefactoringLight.pid');
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
  CurrentPID: DWORD;
  StoredPIDStr: string;
  StoredPID: DWORD;
  MarkerFile: string;
  MarkerExists: Boolean;
  ShouldHint: Boolean;
begin
  CurrentPID := GetCurrentProcessId;
  MarkerFile := GetMarkerFilePath;
  MarkerExists := FileExists(MarkerFile);

  StoredPID := 0;
  if MarkerExists then
  begin
    try
      StoredPIDStr := Trim(TFile.ReadAllText(MarkerFile));
      StoredPID := StrToUIntDef(StoredPIDStr, 0);
    except
      StoredPID := 0;
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
  ShouldHint := MarkerExists and (StoredPID = CurrentPID);

  // Always update the PID - also on normal IDE starts.
  try
    TFile.WriteAllText(MarkerFile, IntToStr(CurrentPID));
  except
    // Tolerate write errors silently.
  end;

  if ShouldHint then
    ShowRestartHint;
end;

end.
