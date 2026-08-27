(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UsesEditor;

// Small, reusable helper to add a unit to a Pascal file's uses clause,
// either in the interface or the implementation section. Comment-aware.
// Writes through IEditorHelper, so the change is undoable when the file
// is open in the IDE and a plain disk write otherwise.

interface

type
  TUsesSection = (usInterface, usImplementation);

/// <summary>True if AUnit already appears in any uses clause of AContent
///  (whole-token, case-insensitive; dotted names compared whole).</summary>
function UnitInUsesText(const AContent, AUnit: string): Boolean;

/// <summary>Adds AUnit to the requested section's uses clause of AFilePath
///  (creating the clause if the section has none). Returns True if the file
///  was changed; False if AUnit was already reachable or on error.</summary>
function AddUnitToUses(const AFilePath, AUnit: string;
  ASection: TUsesSection): Boolean;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils,
  Expert.EditorHelperIntf;

function StripLineComment(const L: string): string;
var
  I, N: Integer;
  InStr: Boolean;
begin
  N := Length(L);
  InStr := False;
  I := 1;
  while I <= N do
  begin
    if L[I] = '''' then
      InStr := not InStr
    else if (not InStr) and (L[I] = '/') and (I < N) and (L[I + 1] = '/') then
      Exit(TrimRight(Copy(L, 1, I - 1)));
    Inc(I);
  end;
  Result := L;
end;

function TokenEquals(const ALine, AUnit: string): Boolean;
// True when AUnit appears as a whole token in ALine (bounded by
// non-identifier, non-'.' characters).
var
  U, Needle: string;
  P, AfterIdx: Integer;
begin
  Result := False;
  U := UpperCase(ALine);
  Needle := UpperCase(AUnit);
  P := Pos(Needle, U);
  while P > 0 do
  begin
    var OkBefore := (P = 1) or
      not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_', '.']);
    AfterIdx := P + Length(Needle);
    var OkAfter := (AfterIdx > Length(U)) or
      not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_', '.']);
    if OkBefore and OkAfter then Exit(True);
    P := Pos(Needle, U, P + 1);
  end;
end;

function UnitInUsesText(const AContent, AUnit: string): Boolean;
var
  SL: TStringList;
  I: Integer;
  InUses: Boolean;
  L, Low: string;
begin
  Result := False;
  if AUnit = '' then Exit;
  SL := TStringList.Create;
  try
    SL.Text := AContent;
    InUses := False;
    for I := 0 to SL.Count - 1 do
    begin
      L := StripLineComment(SL[I]);
      Low := LowerCase(Trim(L));
      if not InUses then
      begin
        if (Low = 'uses') or Low.StartsWith('uses ') or Low.StartsWith('uses'#9) then
          InUses := True
        else
          Continue;
      end;
      if TokenEquals(L, AUnit) then Exit(True);
      if Pos(';', L) > 0 then InUses := False;
    end;
  finally
    SL.Free;
  end;
end;

function AddUnitToUses(const AFilePath, AUnit: string;
  ASection: TUsesSection): Boolean;
var
  Content: string;
  SL: TStringList;
  I, StartIdx, EndIdx, UsesIdx, P: Integer;
  Low: string;
begin
  Result := False;
  if (AUnit = '') or (AFilePath = '') then Exit;

  if (Editor = nil) or not Editor.ReadEditorContent(AFilePath, Content) then
  begin
    if not TFile.Exists(AFilePath) then Exit;
    try Content := TFile.ReadAllText(AFilePath); except Exit; end;
  end;
  if UnitInUsesText(Content, AUnit) then Exit;   // already reachable

  SL := TStringList.Create;
  try
    SL.Text := Content;

    // Delimit the section we are inserting into.
    StartIdx := -1; EndIdx := SL.Count;
    for I := 0 to SL.Count - 1 do
    begin
      Low := LowerCase(Trim(StripLineComment(SL[I])));
      if (ASection = usInterface) and (Low = 'interface') and (StartIdx < 0) then
        StartIdx := I
      else if Low = 'implementation' then
      begin
        if ASection = usImplementation then StartIdx := I
        else if (ASection = usInterface) and (StartIdx >= 0) then
        begin EndIdx := I; Break; end;
      end;
    end;
    if StartIdx < 0 then Exit;   // section not found

    // Find a uses clause inside the section.
    UsesIdx := -1;
    for I := StartIdx + 1 to EndIdx - 1 do
    begin
      Low := LowerCase(Trim(StripLineComment(SL[I])));
      if (Low = 'uses') or Low.StartsWith('uses ') or Low.StartsWith('uses'#9) then
      begin UsesIdx := I; Break; end;
    end;

    if UsesIdx >= 0 then
    begin
      // Insert ", AUnit" before the ';' that closes the clause.
      for I := UsesIdx to EndIdx - 1 do
      begin
        P := Pos(';', StripLineComment(SL[I]));
        if P > 0 then
        begin
          P := Pos(';', SL[I]);
          SL[I] := Copy(SL[I], 1, P - 1) + ', ' + AUnit + Copy(SL[I], P, MaxInt);
          Break;
        end;
      end;
    end
    else
      SL.Insert(StartIdx + 1, 'uses ' + AUnit + ';');

    Result := Editor.ReplaceFileContent(AFilePath, SL.Text);
  finally
    SL.Free;
  end;
end;

end.
