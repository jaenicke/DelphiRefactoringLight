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

/// <summary>Scales a dialog whose controls were placed with HARD-CODED
///  96-dpi coordinates (all of ours are built in code) up to the DPI the
///  form actually runs at. GEOMETRY ONLY - fonts are left alone on
///  purpose: at a scaled DPI the VCL already hands the form a larger
///  font, and scaling it again would overshoot. The failure mode we are
///  fixing is text that no longer fits its box, so the boxes are what
///  must grow. A no-op at 96 dpi.</summary>
procedure ScaleDialogFrom96(AForm: TForm; AAnchor: TCustomForm = nil);

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

/// <summary>Themed confirmation with an explicit DEFAULT of "no": used
///  before a fix DELETES code the user may still want (an unused private
///  method whose body is not empty). ACaption labels the confirming
///  button, so it says what will happen instead of a bare "OK".</summary>
function AskThemedConfirm(const AMsg, AConfirmCaption: string): Boolean;

/// <summary>Installs the confirmation the "remove unused private member"
///  fix asks before it deletes a NON-EMPTY body. Call once at startup -
///  without it that fix refuses to delete anything.</summary>
procedure InstallRemovePrivateConfirm;

implementation

uses
  System.SysUtils, Winapi.Windows, Vcl.Graphics, Expert.IdeThemes,
  Expert.AutoImport
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

procedure ScaleControlTree(AControl: TControl; ANum, ADen: Integer);
var
  I: Integer;
  W: TWinControl;
begin
  if AControl = nil then Exit;
  AControl.SetBounds(
    MulDiv(AControl.Left, ANum, ADen), MulDiv(AControl.Top, ANum, ADen),
    MulDiv(AControl.Width, ANum, ADen), MulDiv(AControl.Height, ANum, ADen));
  if AControl.Constraints <> nil then
  begin
    AControl.Constraints.MinWidth := MulDiv(AControl.Constraints.MinWidth, ANum, ADen);
    AControl.Constraints.MinHeight := MulDiv(AControl.Constraints.MinHeight, ANum, ADen);
    AControl.Constraints.MaxWidth := MulDiv(AControl.Constraints.MaxWidth, ANum, ADen);
    AControl.Constraints.MaxHeight := MulDiv(AControl.Constraints.MaxHeight, ANum, ADen);
  end;
  if AControl is TWinControl then
  begin
    W := TWinControl(AControl);
    for I := 0 to W.ControlCount - 1 do
      ScaleControlTree(W.Controls[I], ANum, ADen);
  end;
end;

procedure ScaleDialogFrom96(AForm: TForm; AAnchor: TCustomForm);
var
  Target, I: Integer;
begin
  if AForm = nil then Exit;
  // Which DPI to scale TO: the monitor the dialog will appear on. Note
  // that CurrentPPI is NOT a reliable answer here - whatever the VCL did
  // while constructing the form, every coordinate was overwritten
  // afterwards by our own hard-coded 96-dpi assignments, so the geometry
  // in front of us is raw 96 either way.
  // The dialog has no position yet, so ITS monitor is whatever happens to
  // contain (0,0) - ask the window we anchor to (the IDE) first.
  Target := 0;
  if (AAnchor <> nil) and (AAnchor.Monitor <> nil) then
    Target := AAnchor.Monitor.PixelsPerInch;
  if (Target <= 0) and (AForm.Monitor <> nil) then
    Target := AForm.Monitor.PixelsPerInch;
  if Target <= 0 then Target := AForm.CurrentPPI;
  if Target <= 0 then Target := Screen.PixelsPerInch;
  if (Target <= 96) or (Target > 1000) then Exit;   // 100% -> nothing to do

  AForm.DisableAlign;
  try
    for I := 0 to AForm.ControlCount - 1 do
      ScaleControlTree(AForm.Controls[I], Target, 96);
    AForm.ClientWidth := MulDiv(AForm.ClientWidth, Target, 96);
    AForm.ClientHeight := MulDiv(AForm.ClientHeight, Target, 96);
    if AForm.Constraints <> nil then
    begin
      AForm.Constraints.MinWidth := MulDiv(AForm.Constraints.MinWidth, Target, 96);
      AForm.Constraints.MinHeight := MulDiv(AForm.Constraints.MinHeight, Target, 96);
    end;
  finally
    AForm.EnableAlign;
  end;
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

  // Our dialogs are built in code at 96 dpi - grow them to the real DPI
  // before anything else looks at their geometry.
  ScaleDialogFrom96(AForm, Anchor);

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

procedure InstallRemovePrivateConfirm;
begin
  RemovePrivateConfirm :=
    function(const AInfo: TPrivateMember): Boolean
    var
      Msg: string;
    begin
      Msg := Format('The private method "%s.%s" is reported as never used.'#13#10#13#10 +
        'Its implementation is NOT empty - removing it deletes %d statement(s) ' +
        'together with the declaration.'#13#10#13#10 +
        'The compiler cannot see uses through RTTI, a DFM event binding or an ' +
        'interface, so please make sure the method is really dead.',
        [AInfo.TypeName, AInfo.Name, AInfo.BodyLines]);
      Result := AskThemedConfirm(Msg, 'Remove anyway');
    end;
end;

function AskThemedConfirm(const AMsg, AConfirmCaption: string): Boolean;
const
  MaxTextWidth = 520;
var
  Dlg: TThemedToolForm;
  Lbl: TLabel;
  BtnOk, BtnCancel: TButton;
  Measure: Vcl.Graphics.TBitmap;
  R: TRect;
  BtnW: Integer;
begin
  Dlg := TThemedToolForm.CreateNew(nil);
  try
    Dlg.BorderStyle := bsDialog;
    Dlg.Caption := 'Refactoring Light';
    Dlg.Position := poScreenCenter;

    Measure := Vcl.Graphics.TBitmap.Create;
    try
      Measure.Canvas.Font := Dlg.Font;
      R := Rect(0, 0, MaxTextWidth, 0);
      DrawText(Measure.Canvas.Handle, PChar(AMsg), Length(AMsg), R,
        DT_CALCRECT or DT_WORDBREAK or DT_NOPREFIX);
    finally
      Measure.Free;
    end;
    if R.Right < 320 then R.Right := 320;

    Lbl := TLabel.Create(Dlg);
    Lbl.Parent := Dlg;
    Lbl.AutoSize := False;
    Lbl.WordWrap := True;
    Lbl.ShowAccelChar := False;
    Lbl.SetBounds(16, 16, R.Right, R.Bottom);
    Lbl.Caption := AMsg;

    BtnW := 130;
    BtnCancel := TButton.Create(Dlg);
    BtnCancel.Parent := Dlg;
    BtnCancel.Caption := 'Cancel';
    // Cancel is the DEFAULT: deleting code must be a deliberate choice.
    BtnCancel.Default := True;
    BtnCancel.Cancel := True;
    BtnCancel.ModalResult := mrCancel;
    BtnCancel.SetBounds(16 + R.Right - 90, R.Bottom + 28, 90, 26);

    BtnOk := TButton.Create(Dlg);
    BtnOk.Parent := Dlg;
    BtnOk.Caption := AConfirmCaption;
    BtnOk.ModalResult := mrOk;
    BtnOk.SetBounds(BtnCancel.Left - BtnW - 8, R.Bottom + 28, BtnW, 26);

    Dlg.ClientWidth := R.Right + 32;
    Dlg.ClientHeight := BtnCancel.Top + BtnCancel.Height + 12;

    EnableThemes(Dlg);
    PrepareDialog(Dlg, nil);
    Result := Dlg.ShowModal = mrOk;
  finally
    Dlg.Free;
  end;
end;

end.
