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
  Winapi.Windows, System.Math, Vcl.Graphics, Expert.PluginSettings;

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
  I, MaxRight, EditLeft: Integer;
  C: TControl;
begin
  MaxRight := 0;
  for I := 0 to grpShortcuts.ControlCount - 1 do
  begin
    C := grpShortcuts.Controls[I];
    if (C is TLabel) and (C <> lblHint) then
      if C.Left + C.Width > MaxRight then
        MaxRight := C.Left + C.Width;
  end;
  EditLeft := MaxRight + 12;
  for I := 0 to grpShortcuts.ControlCount - 1 do
  begin
    C := grpShortcuts.Controls[I];
    if C is TEdit then
    begin
      C.Left := EditLeft;
      // Keep the edit inside the group when the column moved far right.
      if EditLeft + C.Width > grpShortcuts.ClientWidth - 12 then
        C.Width := Max(80, grpShortcuts.ClientWidth - 12 - EditLeft);
    end;
  end;
  // Long texts: wrap within the group width instead of clipping at the
  // designed 96-dpi widths, and grow the groups to fit the wrapped lines.
  var LineH := Abs(lblHint.Font.Height) + 4;
  lblHint.AutoSize := False;
  lblHint.WordWrap := True;
  lblHint.Width := grpShortcuts.ClientWidth - lblHint.Left - 12;
  lblHint.Height := 2 * LineH;
  grpShortcuts.Height := lblHint.Top + lblHint.Height + 14;

  grpLsp.Top := grpShortcuts.Top + grpShortcuts.Height + 12;
  cbxPrewarmLsp.Width := grpLsp.ClientWidth - cbxPrewarmLsp.Left - 12;
  lblLspNote.Width := grpLsp.ClientWidth - lblLspNote.Left - 12;
  lblLspNote.Height := 2 * LineH;
  grpLsp.Height := lblLspNote.Top + lblLspNote.Height + 12;

  btnDefaults.Top := grpLsp.Top + grpLsp.Height + 10;
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
