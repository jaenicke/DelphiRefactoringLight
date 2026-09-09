(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.OptionsFrame;

{
  TFrame embedded into the Tools > Options dialog by Expert.OptionsPage.
  Shows one labelled shortcut field per refactoring feature. The fields
  capture the next pressed key combination (Ctrl/Alt/Shift + key);
  Backspace/Delete clears the shortcut.
}

interface

uses
  Winapi.Messages, System.SysUtils, System.Classes, Vcl.Controls, Vcl.Forms,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Menus,
  Expert.Shortcuts;

type
  TLspOptionsFrame = class(TFrame)
    grpShortcuts: TGroupBox;
    lblRename: TLabel;
    edtRename: TEdit;
    lblCompletion: TLabel;
    edtCompletion: TEdit;
    lblExtract: TLabel;
    edtExtract: TEdit;
    lblFindRef: TLabel;
    edtFindRef: TEdit;
    lblFindImp: TLabel;
    edtFindImp: TEdit;
    lblAlign: TLabel;
    edtAlign: TEdit;
    lblRemoveWith: TLabel;
    edtRemoveWith: TEdit;
    lblMoveToUnit: TLabel;
    edtMoveToUnit: TEdit;
    lblFindOriginal: TLabel;
    edtFindOriginal: TEdit;
    lblHint: TLabel;
    grpLsp: TGroupBox;
    cbxPrewarmLsp: TCheckBox;
    lblLspNote: TLabel;
    grpBlame: TGroupBox;
    cbxLiveBlame: TCheckBox;
    lblBlameInfo: TLabel;
    cbxBlameInfo: TComboBox;
    lblBlameWidth: TLabel;
    edtBlameWidth: TEdit;
    lblBlameOffset: TLabel;
    edtBlameOffset: TEdit;
    cbxTortoise: TCheckBox;
    lblBlameNote: TLabel;
    btnDefaults: TButton;
    procedure ShortcutEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ShortcutEditKeyPress(Sender: TObject; var Key: Char);
    procedure btnDefaultsClick(Sender: TObject);
  private
    function EditFor(Kind: TShortcutKind): TEdit;
    procedure ApplyToEdit(Kind: TShortcutKind);
    procedure AdjustLayout;
    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
  public
    /// <summary>Fills the edits from the current settings.</summary>
    procedure LoadFromSettings;
    /// <summary>Writes the edits back to the settings (no registry I/O).</summary>
    procedure StoreToSettings;
  end;

implementation

{$R *.dfm}

uses
  Winapi.Windows, System.Math, Vcl.Graphics, Expert.PluginSettings
  {$IFNDEF STANDALONE_BUILD}, Expert.BlameGutter{$ENDIF};

{$IFDEF STANDALONE_BUILD}
procedure ApplyBlameSettings;
begin
  // no editor gutter outside the IDE
end;
{$ENDIF}

{ TLspOptionsFrame }

function TLspOptionsFrame.EditFor(Kind: TShortcutKind): TEdit;
begin
  case Kind of
    skRename:     Result := edtRename;
    skCompletion: Result := edtCompletion;
    skExtract:    Result := edtExtract;
    skFindRef:    Result := edtFindRef;
    skFindImp:    Result := edtFindImp;
    skAlign:      Result := edtAlign;
    skRemoveWith: Result := edtRemoveWith;
    skMoveToUnit: Result := edtMoveToUnit;
    skFindOriginal: Result := edtFindOriginal;
  else
    Result := nil;
  end;
end;

procedure TLspOptionsFrame.ApplyToEdit(Kind: TShortcutKind);
var
  E: TEdit;
  SC: TShortCut;
begin
  E := EditFor(Kind);
  if E = nil then Exit;
  SC := TExpertsShortCut.Shortcuts[Kind];
  if SC = 0 then
    E.Text := '(none)'
  else
    E.Text := ShortCutToText(SC);
  E.Tag := Integer(Kind);
end;

// The options host restyles the page with its own (larger) font AFTER the
// frame is streamed - the designed 96-dpi edit column (Left = 160) then
// starts inside the longer labels ("Remove with (project-wide):" was cut
// off). Re-derive the column from the REAL label widths at runtime.
procedure TLspOptionsFrame.AdjustLayout;
var
  I, MaxRight, EditLeft, LineH, Gap, Row: Integer;
  C: TControl;

  // Right edge of the widest label in AParent, ignoring the note labels
  // (those are wrapped to the full width and would win every time).
  function LabelColumn(AParent: TWinControl;
    const AIgnore: array of TControl): Integer;
  var
    K, N: Integer;
    Ctl: TControl;
    Skip: Boolean;
  begin
    Result := 0;
    for K := 0 to AParent.ControlCount - 1 do
    begin
      Ctl := AParent.Controls[K];
      if not (Ctl is TLabel) then Continue;
      Skip := False;
      for N := Low(AIgnore) to High(AIgnore) do
        if Ctl = AIgnore[N] then Skip := True;
      if Skip then Continue;
      if Ctl.Left + Ctl.Width > Result then Result := Ctl.Left + Ctl.Width;
    end;
  end;

begin
  // The options host restyles the page with its own (larger) font AFTER
  // the frame is streamed, so nothing may rely on the designed 96-dpi
  // coordinates. Everything below is laid out as a VERTICAL FLOW: each
  // section is positioned after the previous one and sized from its own
  // children. That is also what a tester's screenshot forced - the
  // "Restore defaults" button was placed after grpLsp while grpBlame
  // still sat at its designed position, so the button landed INSIDE it.
  LineH := Abs(lblHint.Font.Height) + 4;
  Gap := 12;

  // ---- shortcuts: derive the edit column from the real label widths ----
  MaxRight := LabelColumn(grpShortcuts, [lblHint]);
  EditLeft := MaxRight + Gap;
  for I := 0 to grpShortcuts.ControlCount - 1 do
  begin
    C := grpShortcuts.Controls[I];
    if C is TEdit then
    begin
      C.Left := EditLeft;
      if EditLeft + C.Width > grpShortcuts.ClientWidth - Gap then
        C.Width := Max(80, grpShortcuts.ClientWidth - Gap - EditLeft);
    end;
  end;
  lblHint.AutoSize := False;
  lblHint.WordWrap := True;
  lblHint.Width := grpShortcuts.ClientWidth - lblHint.Left - Gap;
  lblHint.Height := 2 * LineH;
  grpShortcuts.Height := lblHint.Top + lblHint.Height + 14;

  // ---- LSP -------------------------------------------------------------
  grpLsp.Top := grpShortcuts.Top + grpShortcuts.Height + Gap;
  cbxPrewarmLsp.Width := grpLsp.ClientWidth - cbxPrewarmLsp.Left - Gap;
  lblLspNote.AutoSize := False;
  lblLspNote.WordWrap := True;
  lblLspNote.Width := grpLsp.ClientWidth - lblLspNote.Left - Gap;
  lblLspNote.Height := 2 * LineH;
  grpLsp.Height := lblLspNote.Top + lblLspNote.Height + Gap;

  // ---- live blame ------------------------------------------------------
  grpBlame.Top := grpLsp.Top + grpLsp.Height + Gap;

  Row := 22;
  cbxLiveBlame.SetBounds(16, Row, grpBlame.ClientWidth - 16 - Gap,
    cbxLiveBlame.Height);

  // One label column for "Show:" and "Column width (px):".
  Row := Row + cbxLiveBlame.Height + 10;
  // lblBlameOffset is EXCLUDED on purpose: it is placed to the RIGHT of
  // the width edit further down, so counting it here would push the
  // column right - and AdjustLayout runs more than once (LoadFromSettings
  // and CM_FONTCHANGED), so it would grow on every call. That feedback
  // was visible in the render test: the combo box wandered off the group.
  EditLeft := LabelColumn(grpBlame, [lblBlameNote, lblBlameOffset]) + Gap;
  // Never let the label column eat the whole group either.
  EditLeft := EnsureRange(EditLeft, 100, grpBlame.ClientWidth div 2);
  lblBlameInfo.Top := Row + 4;
  cbxBlameInfo.SetBounds(EditLeft, Row,
    Min(240, grpBlame.ClientWidth - EditLeft - Gap), cbxBlameInfo.Height);

  Row := Row + cbxBlameInfo.Height + 10;
  lblBlameWidth.Top := Row + 4;
  edtBlameWidth.SetBounds(EditLeft, Row, 70, edtBlameWidth.Height);
  lblBlameOffset.Left := edtBlameWidth.Left + edtBlameWidth.Width + 16;
  lblBlameOffset.Top := Row + 4;
  edtBlameOffset.SetBounds(
    lblBlameOffset.Left + lblBlameOffset.Width + 8, Row, 70,
    edtBlameOffset.Height);

  Row := Row + edtBlameWidth.Height + 10;
  cbxTortoise.SetBounds(16, Row, grpBlame.ClientWidth - 16 - Gap,
    cbxTortoise.Height);

  Row := Row + cbxTortoise.Height + 6;
  lblBlameNote.AutoSize := False;
  lblBlameNote.WordWrap := True;
  lblBlameNote.SetBounds(34, Row, grpBlame.ClientWidth - 34 - Gap, 2 * LineH);
  grpBlame.Height := lblBlameNote.Top + lblBlameNote.Height + Gap;

  // ---- and only THEN the button ---------------------------------------
  btnDefaults.Top := grpBlame.Top + grpBlame.Height + 10;

  // THE FRAME'S OWN HEIGHT must follow the content. The options host
  // scrolls its page by that height, so a frame that stays at its
  // designed size simply cuts off whatever the flow pushed below it -
  // which is what a tester saw: the last note line half visible with the
  // scrollbar already at the end. Constraints.MinHeight carries it even
  // when the host aligns the frame.
  var Bottom := btnDefaults.Top + btnDefaults.Height + 16;
  Constraints.MinHeight := Bottom;
  if Height < Bottom then Height := Bottom;
end;

procedure TLspOptionsFrame.CMFontChanged(var Message: TMessage);
begin
  inherited;
  // The host applies its font after streaming; the AutoSize labels have
  // grown by now - move the edit column out of their way.
  if not (csLoading in ComponentState) then
    AdjustLayout;
end;

procedure TLspOptionsFrame.LoadFromSettings;
var
  K: TShortcutKind;
begin
  AdjustLayout;
  for K := Low(TShortcutKind) to High(TShortcutKind) do
    ApplyToEdit(K);
  cbxPrewarmLsp.Checked := TPluginSettings.PrewarmLspOnProjectOpen;

  if cbxBlameInfo.Items.Count = 0 then
  begin
    cbxBlameInfo.Items.Add('Revision only');
    cbxBlameInfo.Items.Add('Revision and author');
    cbxBlameInfo.Items.Add('Revision, author and age');
  end;
  cbxLiveBlame.Checked := TPluginSettings.LiveBlame;
  cbxBlameInfo.ItemIndex :=
    EnsureRange(TPluginSettings.BlameInfo, 0, cbxBlameInfo.Items.Count - 1);
  edtBlameWidth.Text := IntToStr(TPluginSettings.BlameColumnWidth);
  edtBlameOffset.Text := IntToStr(TPluginSettings.BlameColumnOffset);
  cbxTortoise.Checked := TPluginSettings.BlameUseTortoise;
end;

procedure TLspOptionsFrame.StoreToSettings;
var
  K: TShortcutKind;
  E: TEdit;
  SC: TShortCut;
begin
  for K := Low(TShortcutKind) to High(TShortcutKind) do
  begin
    E := EditFor(K);
    if E = nil then Continue;
    SC := TextToShortCut(E.Text);
    TExpertsShortCut.Shortcuts[K] := SC;
  end;
  TPluginSettings.PrewarmLspOnProjectOpen := cbxPrewarmLsp.Checked;

  TPluginSettings.BlameInfo := Max(0, cbxBlameInfo.ItemIndex);
  // 0 is a legitimate value ("do not touch the gutter"); anything wider
  // than half the editor would just hide code.
  TPluginSettings.BlameColumnWidth :=
    EnsureRange(StrToIntDef(Trim(edtBlameWidth.Text),
      TPluginSettings.BlameColumnWidth), 0, 600);
  TPluginSettings.BlameColumnOffset :=
    EnsureRange(StrToIntDef(Trim(edtBlameOffset.Text),
      TPluginSettings.BlameColumnOffset), 0, 600);
  TPluginSettings.BlameUseTortoise := cbxTortoise.Checked;
  // The switch takes effect immediately - the gutter width is restored
  // when it goes off, so a stale wide gutter can never be left behind.
  TPluginSettings.LiveBlame := cbxLiveBlame.Checked;
  ApplyBlameSettings;
end;

procedure TLspOptionsFrame.ShortcutEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  E: TEdit;
  SC: TShortCut;
begin
  E := Sender as TEdit;

  // Allow tab to leave the field normally.
  if Key = VK_TAB then Exit;

  // Backspace / Delete clears the shortcut.
  if (Key = VK_BACK) or (Key = VK_DELETE) then
  begin
    E.Text := '(none)';
    Key := 0;
    Exit;
  end;

  // Ignore lone modifier keys - we only commit on a real key.
  case Key of
    VK_SHIFT, VK_CONTROL, VK_MENU,
    VK_LSHIFT, VK_RSHIFT, VK_LCONTROL, VK_RCONTROL, VK_LMENU, VK_RMENU,
    VK_LWIN, VK_RWIN, VK_CAPITAL, VK_NUMLOCK, VK_SCROLL:
      Exit;
  end;

  SC := ShortCut(Key, Shift);
  E.Text := ShortCutToText(SC);
  Key := 0;
end;

procedure TLspOptionsFrame.ShortcutEditKeyPress(Sender: TObject; var Key: Char);
begin
  // Suppress the character that would otherwise appear in the edit.
  Key := #0;
end;

procedure TLspOptionsFrame.btnDefaultsClick(Sender: TObject);
var
  K: TShortcutKind;
  E: TEdit;
begin
  for K := Low(TShortcutKind) to High(TShortcutKind) do
  begin
    E := EditFor(K);
    if E <> nil then
      E.Text := ShortCutToText(TExpertsShortCut.Default(K));
  end;
  cbxPrewarmLsp.Checked := TPluginSettings.DefaultPrewarm;
end;

end.
