(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
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
  System.SysUtils, System.Classes, Vcl.Controls, Vcl.Forms,
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
    lblHint: TLabel;
    btnDefaults: TButton;
    procedure ShortcutEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ShortcutEditKeyPress(Sender: TObject; var Key: Char);
    procedure btnDefaultsClick(Sender: TObject);
  private
    function EditFor(Kind: TShortcutKind): TEdit;
    procedure ApplyToEdit(Kind: TShortcutKind);
  public
    /// <summary>Fills the edits from the current settings.</summary>
    procedure LoadFromSettings;
    /// <summary>Writes the edits back to the settings (no registry I/O).</summary>
    procedure StoreToSettings;
  end;

implementation

{$R *.dfm}

uses
  Winapi.Windows, Vcl.Graphics;

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

procedure TLspOptionsFrame.LoadFromSettings;
var
  K: TShortcutKind;
begin
  for K := Low(TShortcutKind) to High(TShortcutKind) do
    ApplyToEdit(K);
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
end;

end.
