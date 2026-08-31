(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DialogHelper;

{
  Shared helpers for the plugin's dialogs:

    - PrepareDialog: ties the dialog to the IDE main window (so it stays
      on the same monitor instead of jumping to a secondary one) and
      applies the active IDE VCL style (dark / light).

    - RegisterDialogClass: registers the form class with the IDE's
      theming services so newly created instances pick up the theme on
      construction.

  The theming uses IOTAIDEThemingServices (Delphi 10.4+); if the IDE
  doesn't expose it (older versions) or theming is disabled, the
  functions degrade gracefully.

  The monitor-jumping issue is fixed by setting PopupMode := pmExplicit
  and PopupParent := <main IDE window>. Without this, a non-owned form
  shown via TForm.Show may be placed on whichever monitor the OS
  currently considers "default", which on multi-monitor setups isn't
  always the one the IDE is running on.
}

interface

uses
  System.Classes, Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls;

type
  /// <summary>Class for the ad-hoc progress/tool windows built with
  ///  CreateNew. A DISTINCT class name matters: the IDE theming service
  ///  registers form CLASSES, and registering plain TForm would affect
  ///  every TForm instance in the process.</summary>
  TThemedToolForm = class(TForm);

  /// <summary>Shared progress window for the project-wide checks (DFM
  ///  events, circular refs, interface GUIDs) so they all look the
  ///  same: themed tool window, path label with ellipsis, progress bar.
  ///  Built + shown by CreateCheckProgress; drive it via Step.</summary>
  TCheckProgressWindow = class(TThemedToolForm)
  private
    FLbl: TLabel;
    FBar: TProgressBar;
  public
    /// <summary>Updates label + bar and pumps the message queue (the
    ///  checks run synchronously on the UI thread). ATotal > 0 renders
    ///  "Current / Total  -  Text" and moves the bar; otherwise AText
    ///  is shown as-is and the bar is left untouched.</summary>
    procedure Step(ACurrent, ATotal: Integer; const AText: string);
    property ProgressLabel: TLabel read FLbl;
    property Bar: TProgressBar read FBar;
  end;

/// <summary>Registers AClass with the IDE theming service so new
///  instances pick up the active theme. Safe to call multiple times;
///  the theming service deduplicates internally.</summary>
procedure RegisterDialogClass(AClass: TCustomFormClass);

/// <summary>Wires AForm to the IDE main window (PopupParent) and
///  applies the active IDE theme. Call from the dialog's constructor
///  AFTER all child controls have been created so the theme reaches
///  them as well.</summary>
procedure PrepareDialog(AForm: TForm; AOwner: TComponent);

/// <summary>Builds, themes (EnableThemes + PrepareDialog) and SHOWS a
///  TCheckProgressWindow with the reference layout used by the project
///  checks. Owner form -> centered on it, else screen center. The
///  caller frees the window.</summary>
function CreateCheckProgress(const ACaption: string; AOwner: TComponent;
  const AInitialText: string = 'Scanning...'): TCheckProgressWindow;

/// <summary>ShowMessage replacement that follows the IDE theme (the VCL
///  message box stays light in dark mode). Deliberately built on our own
///  registered form class - registering the VCL's TMessageForm with the
///  IDE theming service would restyle every plugin's ShowMessage.</summary>
procedure ShowThemedMessage(const AMsg: string);

implementation

uses
  System.SysUtils, Winapi.Windows, Vcl.Graphics, Expert.IdeThemes
  {$IFNDEF STANDALONE_BUILD}, ToolsAPI {$ENDIF};

{$IFNDEF STANDALONE_BUILD}
function ThemingServices: IOTAIDEThemingServices;
begin
  if not Supports(BorlandIDEServices, IOTAIDEThemingServices, Result) then
    Result := nil;
end;
{$ENDIF}

procedure RegisterDialogClass(AClass: TCustomFormClass);
{$IFNDEF STANDALONE_BUILD}
var
  TS: IOTAIDEThemingServices;
{$ENDIF}
begin
{$IFNDEF STANDALONE_BUILD}
  TS := ThemingServices;
  if (TS <> nil) and TS.IDEThemingEnabled then
  try
    TS.RegisterFormClass(AClass);
  except
    // Defensive: don't let a theming bug break dialog creation.
  end;
{$ENDIF}
end;

procedure PrepareDialog(AForm: TForm; AOwner: TComponent);
var
  Anchor: TCustomForm;
  {$IFNDEF STANDALONE_BUILD}
  TS: IOTAIDEThemingServices;
  {$ENDIF}
begin
  // Anchor the dialog to the IDE main window. This both fixes the
  // multi-monitor jump and makes Windows treat the dialog as a child
  // of the IDE for task switching / focus purposes.
  if AOwner is TCustomForm then
    Anchor := TCustomForm(AOwner)
  else
    Anchor := Application.MainForm;

  if Anchor <> nil then
  begin
    AForm.PopupMode := pmExplicit;
    AForm.PopupParent := Anchor;
  end;

{$IFNDEF STANDALONE_BUILD}
  // Apply the IDE theme to the form and all its child controls.
  TS := ThemingServices;
  if (TS <> nil) and TS.IDEThemingEnabled then
  try
    TS.ApplyTheme(AForm);
  except
    // See RegisterDialogClass.
  end;
{$ENDIF}
end;

procedure TCheckProgressWindow.Step(ACurrent, ATotal: Integer; const AText: string);
begin
  if ATotal > 0 then
  begin
    FLbl.Caption := Format('%d / %d  -  %s', [ACurrent, ATotal, AText]);
    FBar.Max := ATotal;
    if ACurrent > ATotal then FBar.Position := ATotal
    else FBar.Position := ACurrent;
  end
  else
    FLbl.Caption := AText;
  Application.ProcessMessages;
end;

function CreateCheckProgress(const ACaption: string; AOwner: TComponent;
  const AInitialText: string): TCheckProgressWindow;
begin
  Result := TCheckProgressWindow.CreateNew(AOwner);
  Result.Caption := ACaption;
  Result.BorderStyle := bsToolWindow;
  Result.FormStyle := fsStayOnTop;
  if AOwner is TCustomForm then
    Result.Position := poOwnerFormCenter
  else
    Result.Position := poScreenCenter;
  Result.ClientWidth := 460;
  Result.ClientHeight := 62;
  Result.FLbl := TLabel.Create(Result);
  Result.FLbl.Parent := Result;
  Result.FLbl.AlignWithMargins := True;
  Result.FLbl.Align := alTop;
  Result.FLbl.Margins.SetBounds(10, 8, 10, 2);
  Result.FLbl.EllipsisPosition := epPathEllipsis;
  Result.FLbl.Caption := AInitialText;
  Result.FBar := TProgressBar.Create(Result);
  Result.FBar.Parent := Result;
  Result.FBar.AlignWithMargins := True;
  Result.FBar.Align := alTop;
  Result.FBar.Margins.SetBounds(10, 4, 10, 6);
  Result.FBar.Height := 18;
  Result.FBar.Min := 0;
  Result.FBar.Max := 100;
  EnableThemes(Result);
  PrepareDialog(Result, AOwner);
  Result.Show;
end;

procedure ShowThemedMessage(const AMsg: string);
const
  MaxTextWidth = 480;
var
  Dlg: TThemedToolForm;
  Lbl: TLabel;
  Btn: TButton;
  Measure: Vcl.Graphics.TBitmap;
  R: TRect;
begin
  Dlg := TThemedToolForm.CreateNew(nil);
  try
    Dlg.BorderStyle := bsDialog;
    Dlg.Caption := 'Refactoring Light';
    Dlg.Position := poScreenCenter;

    // Measure the wrapped text with the form's font.
    Measure := Vcl.Graphics.TBitmap.Create;
    try
      Measure.Canvas.Font := Dlg.Font;
      R := Rect(0, 0, MaxTextWidth, 0);
      DrawText(Measure.Canvas.Handle, PChar(AMsg), Length(AMsg), R,
        DT_CALCRECT or DT_WORDBREAK or DT_NOPREFIX);
    finally
      Measure.Free;
    end;
    if R.Right < 220 then R.Right := 220;

    Lbl := TLabel.Create(Dlg);
    Lbl.Parent := Dlg;
    Lbl.AutoSize := False;
    Lbl.WordWrap := True;
    Lbl.ShowAccelChar := False;
    Lbl.SetBounds(16, 16, R.Right, R.Bottom);
    Lbl.Caption := AMsg;

    Btn := TButton.Create(Dlg);
    Btn.Parent := Dlg;
    Btn.Caption := 'OK';
    Btn.Default := True;
    Btn.Cancel := True;
    Btn.ModalResult := mrOk;
    Btn.SetBounds(16 + (R.Right - 90) div 2, R.Bottom + 28, 90, 26);

    Dlg.ClientWidth := R.Right + 32;
    Dlg.ClientHeight := Btn.Top + Btn.Height + 12;

    EnableThemes(Dlg);
    PrepareDialog(Dlg, nil);
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
