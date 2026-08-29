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

/// <summary>Cuts a trailing '//' comment off a source line
///  (string-literal aware). Shared helper for line-level Pascal parsing.</summary>
function StripLineComment(const L: string): string;

/// <summary>Adds AUnit to the requested section's uses clause of AFilePath
///  (creating the clause if the section has none). Returns True if the file
///  was changed; False if AUnit was already reachable or on error.
///
///  Section aware: a unit listed only under implementation does NOT satisfy
///  an interface-level need - in that case it is MOVED (removed from the
///  implementation clause, added to the interface clause; listing it in
///  both would not compile). The reverse direction is a no-op: a unit in
///  the interface uses is reachable from the implementation too.
///  Clauses containing { } or (* *) (e.g. IFDEFs) are never rewritten -
///  the move is refused (False) rather than risking a mangled clause.</summary>
function AddUnitToUses(const AFilePath, AUnit: string;
  ASection: TUsesSection): Boolean;

/// <summary>Removes AUnit from whichever uses clause of AFilePath lists it
///  (interface or implementation). Same layout handling and IFDEF safety
///  gate as the move logic in AddUnitToUses. False when the unit is not
///  listed or the clause cannot be rewritten safely.</summary>
function RemoveUnitFromUses(const AFilePath, AUnit: string): Boolean;

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

// Line index of the 'uses' keyword inside [AFrom..ATo), and the index of
// the line carrying the closing ';'. False when the range has no clause.
function FindUsesClause(SL: TStringList; AFrom, ATo: Integer;
  out AUsesIdx, ASemiIdx: Integer): Boolean;
var
  I: Integer;
  Low: string;
begin
  Result := False;
  AUsesIdx := -1; ASemiIdx := -1;
  for I := AFrom to ATo - 1 do
  begin
    Low := LowerCase(Trim(StripLineComment(SL[I])));
    if (Low = 'uses') or Low.StartsWith('uses ') or Low.StartsWith('uses'#9) then
    begin
      AUsesIdx := I;
      Break;
    end;
  end;
  if AUsesIdx < 0 then Exit;
  for I := AUsesIdx to ATo - 1 do
    if Pos(';', StripLineComment(SL[I])) > 0 then
    begin
      ASemiIdx := I;
      Exit(True);
    end;
end;

function ClauseContains(SL: TStringList; AUsesIdx, ASemiIdx: Integer;
  const AUnit: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := AUsesIdx to ASemiIdx do
    if TokenEquals(StripLineComment(SL[I]), AUnit) then Exit(True);
end;

// Removes AUnit from the clause [AUsesIdx..ASemiIdx]. Handles the common
// layouts (single line; one-unit-per-line with trailing commas); refuses
// clauses containing block comments / IFDEFs and exotic layouts (returns
// False, file untouched). Deletes the whole clause when AUnit was its only
// unit.
function RemoveFromClause(SL: TStringList; AUsesIdx, ASemiIdx: Integer;
  const AUnit: string): Boolean;
var
  I, LineIdx, A, B, P, UnitCount: Integer;
  Stripped, U, Needle, Rest: string;
begin
  Result := False;

  // Safety gate + unit count: work on the stripped clause text.
  UnitCount := 0;
  for I := AUsesIdx to ASemiIdx do
  begin
    Stripped := StripLineComment(SL[I]);
    if (Pos('{', Stripped) > 0) or (Pos('(*', Stripped) > 0) then Exit;
    // Count commas later; here just refuse braces.
  end;
  var ClauseText := '';
  for I := AUsesIdx to ASemiIdx do
    ClauseText := ClauseText + ' ' + StripLineComment(SL[I]);
  P := Pos(';', ClauseText);
  if P > 0 then ClauseText := Copy(ClauseText, 1, P - 1);
  P := Pos('uses', LowerCase(ClauseText));
  if P > 0 then ClauseText := Copy(ClauseText, P + 4, MaxInt);
  for var Part in ClauseText.Split([',']) do
    if Trim(Part) <> '' then Inc(UnitCount);

  // The only unit -> drop the entire clause.
  if UnitCount <= 1 then
  begin
    for I := ASemiIdx downto AUsesIdx do SL.Delete(I);
    Exit(True);
  end;

  // Locate the token line and its position (positions in the stripped
  // prefix are identical to the original line - StripLineComment only
  // cuts a // tail).
  LineIdx := -1;
  for I := AUsesIdx to ASemiIdx do
    if TokenEquals(StripLineComment(SL[I]), AUnit) then
    begin LineIdx := I; Break; end;
  if LineIdx < 0 then Exit;

  Stripped := StripLineComment(SL[LineIdx]);
  U := UpperCase(Stripped);
  Needle := UpperCase(AUnit);
  A := 0;
  P := Pos(Needle, U);
  while P > 0 do
  begin
    var OkBefore := (P = 1) or
      not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_', '.']);
    var AfterIdx := P + Length(Needle);
    var OkAfter := (AfterIdx > Length(U)) or
      not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_', '.']);
    if OkBefore and OkAfter then begin A := P; Break; end;
    P := Pos(Needle, U, P + 1);
  end;
  if A = 0 then Exit;
  B := A + Length(AUnit);   // first char AFTER the token (1-based)

  // (a) ", " after the token -> remove token + comma + spaces.
  P := B;
  while (P <= Length(Stripped)) and (Stripped[P] = ' ') do Inc(P);
  if (P <= Length(Stripped)) and (Stripped[P] = ',') then
  begin
    Inc(P);
    while (P <= Length(Stripped)) and (Stripped[P] = ' ') do Inc(P);
    SL[LineIdx] := Copy(SL[LineIdx], 1, A - 1) + Copy(SL[LineIdx], P, MaxInt);
    if Trim(SL[LineIdx]) = '' then SL.Delete(LineIdx);
    Exit(True);
  end;

  // (b) "," before the token -> remove comma + spaces + token.
  P := A - 1;
  while (P >= 1) and (Stripped[P] = ' ') do Dec(P);
  if (P >= 1) and (Stripped[P] = ',') then
  begin
    SL[LineIdx] := Copy(SL[LineIdx], 1, P - 1) + Copy(SL[LineIdx], B, MaxInt);
    if Trim(SL[LineIdx]) = '' then SL.Delete(LineIdx);
    Exit(True);
  end;

  // (c) Token alone on its line ("  B," was case (a); here "  B;" or "  B").
  Rest := Trim(Copy(Stripped, B, MaxInt));
  if Trim(Copy(Stripped, 1, A - 1)) = '' then
  begin
    if Rest = ';' then
    begin
      // Line "  B;" - the previous clause line must end with ',': turn
      // that into the closing ';' and drop this line.
      for I := LineIdx - 1 downto AUsesIdx do
      begin
        var PrevStripped := StripLineComment(SL[I]);
        var Q := Length(PrevStripped);
        while (Q >= 1) and (PrevStripped[Q] = ' ') do Dec(Q);
        if (Q >= 1) and (PrevStripped[Q] = ',') then
        begin
          SL[I] := Copy(SL[I], 1, Q - 1) + ';' + Copy(SL[I], Q + 1, MaxInt);
          SL.Delete(LineIdx);
          Exit(True);
        end;
        if Trim(PrevStripped) <> '' then Break;   // unexpected layout
      end;
    end
    else if Rest = '' then
    begin
      // Line "  B" with the ';' further down (unusual) - just drop the
      // line if the NEXT clause line starts the terminator.
      SL.Delete(LineIdx);
      Exit(True);
    end;
  end;

  // Exotic layout (comma-first style etc.) - leave the file untouched.
end;

function RemoveUnitFromUses(const AFilePath, AUnit: string): Boolean;
var
  Content, Low: string;
  SL: TStringList;
  I, IntfIdx, ImplIdx, UsesIdx, SemiIdx: Integer;
  Removed: Boolean;
begin
  Result := False;
  if (AUnit = '') or (AFilePath = '') then Exit;
  if (Editor = nil) or not Editor.ReadEditorContent(AFilePath, Content) then
  begin
    if not TFile.Exists(AFilePath) then Exit;
    try Content := TFile.ReadAllText(AFilePath); except Exit; end;
  end;

  SL := TStringList.Create;
  try
    SL.Text := Content;
    // Anchor at the section keywords (like AddUnitToUses) so text above
    // 'interface' - the unit header comment, say - can never be mistaken
    // for a uses clause. IntfIdx falls back to 0 for headerless files.
    IntfIdx := 0;
    ImplIdx := SL.Count;
    for I := 0 to SL.Count - 1 do
    begin
      Low := LowerCase(Trim(StripLineComment(SL[I])));
      if (Low = 'interface') and (IntfIdx = 0) then
        IntfIdx := I + 1
      else if Low = 'implementation' then
      begin
        ImplIdx := I;
        Break;
      end;
    end;

    Removed := False;
    // Interface clause first, then implementation clause.
    if FindUsesClause(SL, IntfIdx, ImplIdx, UsesIdx, SemiIdx)
      and ClauseContains(SL, UsesIdx, SemiIdx, AUnit) then
      Removed := RemoveFromClause(SL, UsesIdx, SemiIdx, AUnit)
    else if (ImplIdx < SL.Count)
      and FindUsesClause(SL, ImplIdx + 1, SL.Count, UsesIdx, SemiIdx)
      and ClauseContains(SL, UsesIdx, SemiIdx, AUnit) then
      Removed := RemoveFromClause(SL, UsesIdx, SemiIdx, AUnit);

    if Removed then
      Result := Editor.ReplaceFileContent(AFilePath, SL.Text);
  finally
    SL.Free;
  end;
end;

function AddUnitToUses(const AFilePath, AUnit: string;
  ASection: TUsesSection): Boolean;
var
  Content: string;
  SL: TStringList;
  I, IntfIdx, ImplIdx, StartIdx, EndIdx, UsesIdx, SemiIdx, P: Integer;
  Low: string;
  InIntf, InImpl: Boolean;
begin
  Result := False;
  if (AUnit = '') or (AFilePath = '') then Exit;

  if (Editor = nil) or not Editor.ReadEditorContent(AFilePath, Content) then
  begin
    if not TFile.Exists(AFilePath) then Exit;
    try Content := TFile.ReadAllText(AFilePath); except Exit; end;
  end;

  SL := TStringList.Create;
  try
    SL.Text := Content;

    // Section boundaries.
    IntfIdx := -1; ImplIdx := SL.Count;
    for I := 0 to SL.Count - 1 do
    begin
      Low := LowerCase(Trim(StripLineComment(SL[I])));
      if (Low = 'interface') and (IntfIdx < 0) then
        IntfIdx := I
      else if Low = 'implementation' then
      begin
        ImplIdx := I;
        Break;
      end;
    end;

    // Where is the unit already listed?
    InIntf := False; InImpl := False;
    if (IntfIdx >= 0)
      and FindUsesClause(SL, IntfIdx + 1, ImplIdx, UsesIdx, SemiIdx) then
      InIntf := ClauseContains(SL, UsesIdx, SemiIdx, AUnit);
    if (ImplIdx < SL.Count)
      and FindUsesClause(SL, ImplIdx + 1, SL.Count, UsesIdx, SemiIdx) then
      InImpl := ClauseContains(SL, UsesIdx, SemiIdx, AUnit);

    if ASection = usInterface then
    begin
      if InIntf then Exit;   // already reachable at interface level
      if InImpl then
      begin
        // MOVE: a unit under implementation does not satisfy an interface
        // need, and listing it in both clauses would not compile. Removing
        // implementation lines never shifts the interface indices (the
        // implementation section comes after).
        if not FindUsesClause(SL, ImplIdx + 1, SL.Count, UsesIdx, SemiIdx) then Exit;
        if not RemoveFromClause(SL, UsesIdx, SemiIdx, AUnit) then Exit;
      end;
      StartIdx := IntfIdx;
      EndIdx := ImplIdx;
    end
    else
    begin
      if InIntf or InImpl then Exit;   // reachable either way
      if ImplIdx >= SL.Count then Exit;
      StartIdx := ImplIdx;
      EndIdx := SL.Count;
    end;
    if StartIdx < 0 then Exit;   // section not found

    // Insert into the section's uses clause (create one if missing).
    if FindUsesClause(SL, StartIdx + 1, EndIdx, UsesIdx, SemiIdx) then
    begin
      P := Pos(';', StripLineComment(SL[SemiIdx]));
      if P <= 0 then Exit;
      P := Pos(';', SL[SemiIdx]);
      SL[SemiIdx] := Copy(SL[SemiIdx], 1, P - 1) + ', ' + AUnit
        + Copy(SL[SemiIdx], P, MaxInt);
    end
    else
      SL.Insert(StartIdx + 1, 'uses ' + AUnit + ';');

    Result := Editor.ReplaceFileContent(AFilePath, SL.Text);
  finally
    SL.Free;
  end;
end;

end.
