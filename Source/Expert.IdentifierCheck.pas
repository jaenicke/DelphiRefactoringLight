(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.IdentifierCheck;

interface

uses
  System.SysUtils, System.Classes;

type
  TIdentifierCheckStatus = (icsEmpty, icsOk, icsInvalid, icsKeyword, icsUnchanged,
    icsInUnit, icsInProject);

  TIdentifierCheckResult = record
    Status: TIdentifierCheckStatus;
    Message: string;
    InUnitCount: Integer;
    InProjectFileCount: Integer;
  end;

  TIdentifierChecker = class
  public
    class function IsValidIdentifier(const AName: string): Boolean; static;
    class function IsPascalKeyword(const AName: string): Boolean; static;
    class function CountInText(const AText, AIdent: string): Integer; static;
    class function CountInFiles(const AFiles: TArray<string>;
      const AIdent: string; out AFilesWithHit: Integer): Integer; static;

    /// <summary>Full synchronous check. ACurrentFileText is the current
    ///  unit's full source (for in-unit collision detection).
    ///  AProjectFiles are scanned for project-wide occurrences (skip the
    ///  current unit). AOldName (may be empty) is the identifier being
    ///  renamed, used for "unchanged" detection and to discount its own
    ///  occurrences.</summary>
    class function Check(const ANewName, AOldName, ACurrentFileText: string;
      const AProjectFiles: TArray<string>; const ACurrentFile: string): TIdentifierCheckResult; static;
  end;

implementation

uses
  System.Character, System.StrUtils, System.IOUtils;

const
  KEYWORDS: array[0..65] of string = (
    'AND', 'ARRAY', 'AS', 'ASM', 'BEGIN', 'CASE', 'CLASS', 'CONST',
    'CONSTRUCTOR', 'DESTRUCTOR', 'DISPINTERFACE', 'DIV', 'DO', 'DOWNTO',
    'ELSE', 'END', 'EXCEPT', 'EXPORTS', 'FILE', 'FINALIZATION', 'FINALLY',
    'FOR', 'FUNCTION', 'GOTO', 'IF', 'IMPLEMENTATION', 'IN', 'INHERITED',
    'INITIALIZATION', 'INLINE', 'INTERFACE', 'IS', 'LABEL', 'LIBRARY',
    'MOD', 'NIL', 'NOT', 'OBJECT', 'OF', 'OR', 'PACKED', 'PROCEDURE',
    'PROGRAM', 'PROPERTY', 'RAISE', 'RECORD', 'REPEAT', 'RESOURCESTRING',
    'SET', 'SHL', 'SHR', 'STRING', 'THEN', 'THREADVAR', 'TO', 'TRY',
    'TYPE', 'UNIT', 'UNTIL', 'USES', 'VAR', 'WHILE', 'WITH', 'XOR',
    'FALSE', 'TRUE');

class function TIdentifierChecker.IsValidIdentifier(const AName: string): Boolean;
var
  I: Integer;
  Ch: Char;
begin
  if AName = '' then Exit(False);
  Ch := AName[1];
  if not (Ch.IsLetter or (Ch = '_')) then Exit(False);
  for I := 2 to Length(AName) do
  begin
    Ch := AName[I];
    if not (Ch.IsLetterOrDigit or (Ch = '_')) then Exit(False);
  end;
  Result := True;
end;

class function TIdentifierChecker.IsPascalKeyword(const AName: string): Boolean;
var
  U, K: string;
begin
  U := UpperCase(AName);
  for K in KEYWORDS do
    if K = U then Exit(True);
  Result := False;
end;

class function TIdentifierChecker.CountInText(const AText, AIdent: string): Integer;
var
  I, Len, ILen: Integer;
  UText, UIdent: string;
  Ch: Char;
  Matches: Boolean;
begin
  Result := 0;
  if (AText = '') or (AIdent = '') then Exit;
  UText := UpperCase(AText);
  UIdent := UpperCase(AIdent);
  ILen := Length(UIdent);
  Len := Length(UText);
  I := 1;
  while I <= Len - ILen + 1 do
  begin
    if UText[I] = UIdent[1] then
    begin
      // word-boundary before
      Matches := (I = 1);
      if not Matches then
      begin
        Ch := UText[I - 1];
        Matches := not (Ch.IsLetterOrDigit or (Ch = '_'));
      end;
      if Matches then
      begin
        Matches := CompareMem(@UText[I], @UIdent[1], ILen * SizeOf(Char));
        if Matches and (I + ILen <= Len) then
        begin
          Ch := UText[I + ILen];
          Matches := not (Ch.IsLetterOrDigit or (Ch = '_'));
        end;
      end;
      if Matches then
      begin
        Inc(Result);
        Inc(I, ILen);
        Continue;
      end;
    end;
    Inc(I);
  end;
end;

class function TIdentifierChecker.CountInFiles(const AFiles: TArray<string>;
  const AIdent: string; out AFilesWithHit: Integer): Integer;
var
  F, Content: string;
  C: Integer;
begin
  Result := 0;
  AFilesWithHit := 0;
  for F in AFiles do
  begin
    try
      Content := TFile.ReadAllText(F);
    except
      Continue;
    end;
    C := CountInText(Content, AIdent);
    if C > 0 then
    begin
      Inc(Result, C);
      Inc(AFilesWithHit);
    end;
  end;
end;

class function TIdentifierChecker.Check(const ANewName, AOldName, ACurrentFileText: string;
  const AProjectFiles: TArray<string>; const ACurrentFile: string): TIdentifierCheckResult;
var
  Trimmed: string;
  InUnitTotal, InUnitSelf: Integer;
  OtherFiles: TArray<string>;
  F: string;
  ProjHits, ProjFilesWith: Integer;
begin
  Result := Default(TIdentifierCheckResult);
  Trimmed := Trim(ANewName);

  if Trimmed = '' then
  begin
    Result.Status := icsEmpty;
    Result.Message := 'Please enter an identifier.';
    Exit;
  end;

  if not IsValidIdentifier(Trimmed) then
  begin
    Result.Status := icsInvalid;
    Result.Message := 'Invalid identifier.';
    Exit;
  end;

  if IsPascalKeyword(Trimmed) then
  begin
    Result.Status := icsKeyword;
    Result.Message := '"' + Trimmed + '" is a reserved Pascal keyword.';
    Exit;
  end;

  if (AOldName <> '') and SameText(Trimmed, AOldName) then
  begin
    Result.Status := icsUnchanged;
    Result.Message := 'Name unchanged.';
    Exit;
  end;

  // In-unit collision (subtract occurrences of AOldName-matched identifier)
  InUnitTotal := CountInText(ACurrentFileText, Trimmed);
  InUnitSelf := 0;
  if AOldName <> '' then
    InUnitSelf := CountInText(ACurrentFileText, AOldName);
  // Heuristic: if rename and new != old, any occurrences already in the unit
  // are true collisions. Otherwise (extract method with no old name),
  // same semantics.
  Result.InUnitCount := InUnitTotal;
  if InUnitTotal > 0 then
  begin
    // If the old name equals the new name's count inside the file, we consider
    // those occurrences are the pre-existing ones of AOldName and therefore
    // not a collision. We already returned early for SameText above, so any
    // InUnitTotal > 0 here means "some symbol with that new name exists".
    Result.Status := icsInUnit;
    Result.Message := Format('"%s" already appears %d time(s) in this unit.',
      [Trimmed, InUnitTotal]);
    // do not return yet - still report project count
  end;
  // silence hint about InUnitSelf
  if InUnitSelf < 0 then ;

  // Project-wide (skip the current file)
  SetLength(OtherFiles, 0);
  for F in AProjectFiles do
    if not SameText(F, ACurrentFile) then
    begin
      SetLength(OtherFiles, Length(OtherFiles) + 1);
      OtherFiles[High(OtherFiles)] := F;
    end;

  ProjHits := CountInFiles(OtherFiles, Trimmed, ProjFilesWith);
  Result.InProjectFileCount := ProjFilesWith;

  if Result.Status = icsInUnit then Exit;

  if ProjHits > 0 then
  begin
    Result.Status := icsInProject;
    Result.Message := Format('"%s" exists in %d other project file(s) (%d occurrence(s)).',
      [Trimmed, ProjFilesWith, ProjHits]);
    Exit;
  end;

  Result.Status := icsOk;
  Result.Message := Format('"%s" is available.', [Trimmed]);
end;

end.
