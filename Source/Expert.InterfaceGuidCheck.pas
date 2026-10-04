(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.InterfaceGuidCheck;

// Scans project sources for interface declarations and their GUIDs.
// Duplicate GUIDs - the classic copy/paste accident - cause silently
// wrong behaviour in Supports / QueryInterface (the first interface
// with the GUID wins), so they are flagged for the dialog to show in
// red at the top of the list.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type
  TInterfaceGuidEntry = record
    InterfaceName: string;
    /// <summary>GUID including braces, original spelling.</summary>
    Guid: string;
    FileName: string;
    /// <summary>1-based line of the interface declaration.</summary>
    Line: Integer;
    IsDuplicate: Boolean;
    /// <summary>True for declarations without any GUID - not an error
    ///  per se (interfaces used without Supports/QueryInterface don't
    ///  need one) but worth seeing in the list.</summary>
    HasGuid: Boolean;
    /// <summary>True for "= dispinterface" declarations. Type-library
    ///  imports pair every dual interface with a dispinterface that
    ///  INTENTIONALLY shares its GUID - one interface plus one
    ///  dispinterface on the same GUID is therefore NOT flagged as a
    ///  duplicate.</summary>
    IsDispInterface: Boolean;
  end;

/// <summary>True when ALINE declares an interface: "IFoo = interface[(...)]"
///  or "= dispinterface", with ANAME the declared name. Both tests matter and
///  both were missing (found on a real project, 2026-09-30, where the check
///  listed two of them as interfaces without a GUID):
///  * the '=' must be a real one - "result := InterfaceArrayFind(...)"
///    contains the text "= INTERFACE" but is an assignment;
///  * the keyword must END at a word boundary - an alias
///    "TFoo = Interfaces.GTIDLL.TFoo;" is not an interface declaration.
///  A forward declaration ("= interface;") answers False: it carries no GUID,
///  the full declaration elsewhere does.</summary>
function IsInterfaceDeclLine(const ALine: string; out AName: string;
  out AIsDisp: Boolean): Boolean;

type
  /// <summary>What assigning a GUID to one interface declaration would
  ///  change. Ok = False means nothing is written and Problem says why -
  ///  the shapes a human has to decide (see PlanInterfaceGuidEdit).</summary>
  TGuidEditPlan = record
    Ok: Boolean;
    Problem: string;
    /// <summary>The whole text with the edit applied.</summary>
    Lines: TArray<string>;
    /// <summary>0-based line that was rewritten, or the inserted one.</summary>
    Line: Integer;
    /// <summary>True when the declaration had no GUID and a line was
    ///  INSERTED for it; False when an existing GUID was replaced.</summary>
    Inserted: Boolean;
    /// <summary>The GUID that stood there before ('' when there was none),
    ///  so a caller can report what it replaced.</summary>
    OldGuid: string;
  end;

/// <summary>True for the text of a GUID literal: "{8-4-4-4-12}" hex digits
///  with braces, which is what Delphi accepts behind an interface. Nothing
///  else may ever reach a source file through the fix.</summary>
function IsGuidText(const AGuid: string): Boolean;

/// <summary>The 0-based line carrying the GUID of the interface declared at
///  ADeclLine, or -1 when it has none. EXACTLY the search
///  TInterfaceGuidChecker.Scan does (the declaration line or the next three,
///  stopped by the next interface declaration - audit #37, M24), because a
///  fix that edited another line than the check reported would be worse than
///  no fix at all. AMASKED is MaskCommentsAndStrings(ALines).</summary>
function FindInterfaceGuidLine(const ALines, AMasked: TArray<string>;
  ADeclLine: Integer; out AGuid: string): Integer;

/// <summary>Every 0-based line of ALINES that declares an interface named
///  ANAME. The fix locates its target by NAME in the text as it is NOW
///  (buffer or disk) instead of trusting the line the check reported, which
///  any edit since can have moved. More than one answer means a conditional
///  declaration ("{$IF} IFoo = interface ... {$ELSE} IFoo = interface") -
///  there the caller must refuse rather than pick one.</summary>
function InterfaceDeclLines(const ALines: TArray<string>;
  const AName: string): TArray<Integer>;

/// <summary>Assigns ANEWGUID to the interface declared at ADECLLINE
///  (0-based): replaces the GUID it has, or inserts "['{...}']" on its own
///  line below the declaration. REFUSED (Ok = False) when ADeclLine does not
///  declare an interface any more (the file changed since the check) and when
///  the declaration is closed on its own line ("IFoo = interface end;") -
///  there the GUID's place is a question for a human, and a wrong guess
///  produces a unit that does not compile.</summary>
function PlanInterfaceGuidEdit(const ALines: TArray<string>;
  ADeclLine: Integer; const ANewGuid: string): TGuidEditPlan;

type
  TInterfaceGuidChecker = class
  public
    /// <summary>Scans AFiles (only .pas are considered) and returns
    ///  every interface declaration found. Entries with duplicate
    ///  GUIDs have IsDuplicate = True. AProgress (optional) is called
    ///  per file with (current, total, filename) so the caller can
    ///  show feedback during large scans.</summary>
    class function Scan(const AFiles: TArray<string>;
      const AProgress: TProc<Integer, Integer, string> = nil): TArray<TInterfaceGuidEntry>;

    /// <summary>Scans a single file. IsDuplicate is NOT set - callers
    ///  merging into an existing entry set run MarkDuplicates over the
    ///  combined array afterwards. Used by the dialog's live refresh
    ///  when the active unit changes.</summary>
    class function ScanSingleFile(const AFile: string): TArray<TInterfaceGuidEntry>;

    /// <summary>(Re)computes IsDuplicate across the whole array,
    ///  honouring the interface/dispinterface dual-pair rule.</summary>
    class procedure MarkDuplicates(var AEntries: TArray<TInterfaceGuidEntry>);
  end;

implementation

uses
  System.IOUtils, System.StrUtils,
  Expert.EditorHelperIntf, Delphi.FileEncoding, Expert.PascalScanner;

function SplitLines(const AContent: string): TArray<string>;
begin
  // Normalize CRLF / lone LF / lone CR before splitting - files coming
  // out of git with LF-only endings would otherwise parse as one line.
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10)
    .Split([#10], TStringSplitOptions.None);
end;

function ReadFileLines(const AFile: string): TArray<string>;
var
  Content: string;
begin
  Result := nil;
  if (Editor <> nil) and Editor.ReadEditorContent(AFile, Content) then
    Exit(SplitLines(Content));
  if not TFile.Exists(AFile) then Exit;
  try
    Content := TDelphiFileEncoding.ReadAll(AFile);
    Result := SplitLines(Content);
  except
    Result := nil;
  end;
end;

/// <summary>Extracts "['{...}']" from a line; returns '' if absent.
///  AMASKED is the same line from MaskCommentsAndStrings. The GUID IS a
///  string literal, so the mask blanks it: reading it from the masked line
///  found no GUID at all from 1.16.18 on - every interface was listed as
///  "(no GUID)" and no duplicate could be reported any more. The text comes
///  from ALINE; the mask only decides whether its '[' is code, so a GUID
///  inside a comment still does not count (audit #37, L3d).
///  Blanks inside the brackets are allowed: "[ '{...}' ]" compiles and IS
///  the interface's GUID, but the exact text "['{" was looked for, so such
///  an interface was listed as "(no GUID)".</summary>
function ExtractGuid(const ALine, AMasked: string): string;
var
  P, Q, R: Integer;
begin
  Result := '';
  // A '[' of the CODE: the mask keeps every position and blanks comments
  // and strings, so a hit in AMASKED is the same column in ALINE.
  P := Pos('[', AMasked);
  while P > 0 do
  begin
    Q := P + 1;
    while (Q <= Length(ALine)) and CharInSet(ALine[Q], [' ', #9]) do Inc(Q);
    if Copy(ALine, Q, 2) = '''{' then
    begin
      R := PosEx('}''', ALine, Q + 2);
      if R > 0 then
      begin
        var C := R + 2;
        while (C <= Length(ALine)) and CharInSet(ALine[C], [' ', #9]) do Inc(C);
        if (C <= Length(ALine)) and (ALine[C] = ']') then
          Exit(Copy(ALine, Q + 1, R - Q));  // {....}
      end;
    end;
    P := PosEx('[', AMasked, P + 1);
  end;
end;

function IsGuidText(const AGuid: string): Boolean;
const
  Groups: array[0..4] of Integer = (8, 4, 4, 4, 12);
var
  P: Integer;
begin
  Result := False;
  if Length(AGuid) <> 38 then Exit;          // {8-4-4-4-12} with braces
  if (AGuid[1] <> '{') or (AGuid[Length(AGuid)] <> '}') then Exit;
  P := 2;
  for var G := 0 to High(Groups) do
  begin
    if G > 0 then
    begin
      if AGuid[P] <> '-' then Exit;
      Inc(P);
    end;
    for var I := 1 to Groups[G] do
    begin
      if not CharInSet(AGuid[P], ['0'..'9', 'A'..'F', 'a'..'f']) then Exit;
      Inc(P);
    end;
  end;
  Result := P = Length(AGuid);
end;

function FindInterfaceGuidLine(const ALines, AMasked: TArray<string>;
  ADeclLine: Integer; out AGuid: string): Integer;

  function MaskedLine(AIndex: Integer): string;
  begin
    if AIndex <= High(AMasked) then Result := AMasked[AIndex]
    else Result := ALines[AIndex];
  end;

var
  Name: string;
  Disp: Boolean;
  J: Integer;
begin
  AGuid := '';
  Result := -1;
  if (ADeclLine < 0) or (ADeclLine > High(ALines)) then Exit;
  AGuid := ExtractGuid(ALines[ADeclLine], MaskedLine(ADeclLine));
  if AGuid <> '' then Exit(ADeclLine);
  J := ADeclLine;
  while (J < High(ALines)) and (J < ADeclLine + 3) do
  begin
    Inc(J);
    if IsInterfaceDeclLine(MaskedLine(J), Name, Disp) then Break;
    AGuid := ExtractGuid(ALines[J], MaskedLine(J));
    if AGuid <> '' then Exit(J);
  end;
end;

function InterfaceDeclLines(const ALines: TArray<string>;
  const AName: string): TArray<Integer>;
var
  Masked: TArray<string>;
  Name: string;
  Disp: Boolean;
begin
  Result := nil;
  if (Length(ALines) = 0) or (AName = '') then Exit;
  Masked := MaskCommentsAndStrings(ALines);
  for var I := 0 to High(ALines) do
  begin
    var L: string;
    if I <= High(Masked) then L := Masked[I] else L := ALines[I];
    if IsInterfaceDeclLine(L, Name, Disp) and SameText(Name, AName) then
      Result := Result + [I];
  end;
end;

function PlanInterfaceGuidEdit(const ALines: TArray<string>;
  ADeclLine: Integer; const ANewGuid: string): TGuidEditPlan;

  function Indentation(const ALine: string): string;
  begin
    Result := '';
    for var I := 1 to Length(ALine) do
      if CharInSet(ALine[I], [' ', #9]) then Result := Result + ALine[I]
      else Break;
  end;

  // 'end' as a WORD of the code (the line comes in masked, so a comment or
  // a string cannot match). Written here rather than pulled in from
  // Expert.AutoImport: this unit stays free of that dependency.
  function ClosesOnThisLine(const AMasked: string): Boolean;
  var
    U: string;
    P: Integer;
  begin
    U := UpperCase(AMasked);
    P := Pos('END', U);
    while P > 0 do
    begin
      if ((P = 1) or not IsIdentChar(U[P - 1])) and
         ((P + 3 > Length(U)) or not IsIdentChar(U[P + 3])) then Exit(True);
      P := Pos('END', U, P + 1);
    end;
    Result := False;
  end;

var
  Masked: TArray<string>;
  Name, Line, Indent: string;
  Disp: Boolean;
  GLine, P: Integer;
begin
  Result := Default(TGuidEditPlan);
  Result.Line := -1;
  if (ADeclLine < 0) or (ADeclLine > High(ALines)) then
  begin
    Result.Problem := 'line ' + IntToStr(ADeclLine + 1) + ' is outside the file';
    Exit;
  end;
  if not IsGuidText(ANewGuid) then
  begin
    Result.Problem := '"' + ANewGuid + '" is not a GUID';
    Exit;
  end;
  Masked := MaskCommentsAndStrings(ALines);
  if not IsInterfaceDeclLine(Masked[ADeclLine], Name, Disp) then
  begin
    Result.Problem := 'line ' + IntToStr(ADeclLine + 1) + ' does not declare ' +
      'an interface - the file has changed since the check, run it again';
    Exit;
  end;
  // "IFoo = interface end;" - where the GUID belongs in a declaration that is
  // opened and closed on one line is a question for a human.
  if ClosesOnThisLine(Masked[ADeclLine]) then
  begin
    Result.Problem := Name + ' is declared and closed on one line - add the ' +
      'GUID by hand';
    Exit;
  end;

  Result.Lines := Copy(ALines, 0, Length(ALines));
  GLine := FindInterfaceGuidLine(ALines, Masked, ADeclLine, Result.OldGuid);
  if GLine >= 0 then
  begin
    // Replace the GUID where it stands: brackets, quotes, blanks and the
    // indentation of that line stay exactly as the author wrote them.
    Line := Result.Lines[GLine];
    P := Pos(Result.OldGuid, Line);
    if P <= 0 then
    begin
      Result.Problem := 'the GUID of ' + Name + ' could not be located on ' +
        'line ' + IntToStr(GLine + 1);
      Result.Lines := nil;
      Exit;
    end;
    Result.Lines[GLine] := Copy(Line, 1, P - 1) + ANewGuid +
      Copy(Line, P + Length(Result.OldGuid), MaxInt);
    Result.Line := GLine;
  end
  else
  begin
    // No GUID: its own line below the declaration, which is also right for
    // "IFoo = interface(IBar)" - the parent list stays untouched.
    Indent := Indentation(ALines[ADeclLine]) + '  ';
    if ADeclLine < High(ALines) then
    begin
      var NextIndent := Indentation(ALines[ADeclLine + 1]);
      if (Trim(ALines[ADeclLine + 1]) <> '') and
         (Length(NextIndent) > Length(Indentation(ALines[ADeclLine]))) then
        Indent := NextIndent;
    end;
    Insert([Indent + '[''' + ANewGuid + ''']'], Result.Lines, ADeclLine + 1);
    Result.Line := ADeclLine + 1;
    Result.Inserted := True;
  end;
  Result.Ok := True;
end;

function IsInterfaceDeclLine(const ALine: string; out AName: string;
  out AIsDisp: Boolean): Boolean;
var
  U: string;
  P, Q, KwEnd: Integer;
begin
  Result := False;
  AName := '';
  AIsDisp := False;
  U := UpperCase(ALine);
  P := 1;
  while True do
  begin
    P := PosEx('=', U, P);
    if P = 0 then Exit;
    // ':=' is an assignment, '<=' / '>=' comparisons - none of them declare
    if (P > 1) and CharInSet(U[P - 1], [':', '<', '>']) then
    begin
      Inc(P);
      Continue;
    end;
    Q := P + 1;
    while (Q <= Length(U)) and CharInSet(U[Q], [' ', #9]) do Inc(Q);
    if Copy(U, Q, Length('DISPINTERFACE')) = 'DISPINTERFACE' then
    begin
      AIsDisp := True;
      KwEnd := Q + Length('DISPINTERFACE');
    end
    else if Copy(U, Q, Length('INTERFACE')) = 'INTERFACE' then
    begin
      AIsDisp := False;
      KwEnd := Q + Length('INTERFACE');
    end
    else
    begin
      Inc(P);
      Continue;
    end;
    // WORD BOUNDARY: "Interfaces.GTIDLL.TFoo" is an alias and
    // "InterfaceArrayFind(" a call - neither is the keyword
    if (KwEnd <= Length(U)) and CharInSet(U[KwEnd], ['A'..'Z', '0'..'9', '_']) then
    begin
      Inc(P);
      Continue;
    end;
    if Trim(Copy(U, KwEnd, MaxInt)).StartsWith(';') then Exit;   // forward decl
    AName := Trim(Copy(ALine, 1, P - 1));
    // "type IFoo = interface" names the INTERFACE, not "type IFoo"
    // (audit #37, L3d): a one-line declaration kept the keyword in the
    // name, so every report and every duplicate comparison used a name
    // that does not exist.
    if AName.ToUpper.StartsWith('TYPE ') then
      AName := Trim(Copy(AName, 6, MaxInt));
    if (AName = '') or not CharInSet(AName[1], ['A'..'Z', 'a'..'z', '_']) then Exit;
    // The name is ONE identifier, optionally with a generic parameter list
    // ("IList<T: IObject>" is a real shape in this project, so the check
    // stops at the '<' rather than rejecting it).
    for var CI := 2 to Length(AName) do
    begin
      if AName[CI] = '<' then Break;
      if not CharInSet(AName[CI], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then Exit;
    end;
    Exit(True);
  end;
end;

class function TInterfaceGuidChecker.ScanSingleFile(
  const AFile: string): TArray<TInterfaceGuidEntry>;
var
  Entries: TList<TInterfaceGuidEntry>;
  L, Name, Guid: string;
  Lines: TArray<string>;
  I: Integer;
  E: TInterfaceGuidEntry;
begin
  Result := nil;
  if not SameText(ExtractFileExt(AFile), '.pas') then Exit;
  Entries := TList<TInterfaceGuidEntry>.Create;
  try
    Lines := ReadFileLines(AFile);
    // MASKED, not just '//'-stripped (audit #37, L3d): an interface inside
    // { } or (* *) was reported, which also made the REAL one a duplicate
    // of itself. Masking keeps the line length, so every position below
    // still refers to the real line - which is where the GUID is READ from,
    // because the mask blanks string literals and the GUID is one.
    var Masked := MaskCommentsAndStrings(Lines);
    for I := 0 to High(Lines) do
    begin
      if I <= High(Masked) then L := Masked[I] else L := Lines[I];
      var IsDisp := False;
      if not IsInterfaceDeclLine(L, Name, IsDisp) then Continue;

      // GUID on the declaration line or within the next 3, stopped by the
      // next declaration (audit #37, M24). ONE implementation, shared with
      // the fix of 1.23.0: if the two searched differently, the fix would
      // rewrite another line than the one this check reported.
      FindInterfaceGuidLine(Lines, Masked, I, Guid);

      E := Default(TInterfaceGuidEntry);
      E.InterfaceName := Name;
      E.Guid := Guid;
      E.FileName := AFile;
      E.Line := I + 1;
      E.HasGuid := Guid <> '';
      E.IsDispInterface := IsDisp;
      Entries.Add(E);
    end;
    Result := Entries.ToArray;
  finally
    Entries.Free;
  end;
end;

class procedure TInterfaceGuidChecker.MarkDuplicates(
  var AEntries: TArray<TInterfaceGuidEntry>);
type
  TGuidUse = record
    IntfCount: Integer;   // "= interface" declarations with this GUID
    DispCount: Integer;   // "= dispinterface" declarations with this GUID
  end;
var
  GuidCount: TDictionary<string, TGuidUse>;
  I: Integer;
  Use: TGuidUse;
begin
  GuidCount := TDictionary<string, TGuidUse>.Create;
  try
    for I := 0 to High(AEntries) do
      if AEntries[I].HasGuid then
      begin
        var Key := UpperCase(AEntries[I].Guid);
        if not GuidCount.TryGetValue(Key, Use) then
          Use := Default(TGuidUse);
        if AEntries[I].IsDispInterface then Inc(Use.DispCount) else Inc(Use.IntfCount);
        GuidCount.AddOrSetValue(Key, Use);
      end;

    // One interface + one dispinterface on the same GUID is the
    // legitimate COM dual-interface pattern (type-library imports
    // generate exactly that) - only a second declaration OF THE SAME
    // KIND makes a GUID collision.
    for I := 0 to High(AEntries) do
    begin
      AEntries[I].IsDuplicate := False;
      if AEntries[I].HasGuid
         and GuidCount.TryGetValue(UpperCase(AEntries[I].Guid), Use)
         and ((Use.IntfCount > 1) or (Use.DispCount > 1)) then
        AEntries[I].IsDuplicate := True;
    end;
  finally
    GuidCount.Free;
  end;
end;

class function TInterfaceGuidChecker.Scan(const AFiles: TArray<string>;
  const AProgress: TProc<Integer, Integer, string>): TArray<TInterfaceGuidEntry>;
var
  All: TList<TInterfaceGuidEntry>;
  I: Integer;
begin
  All := TList<TInterfaceGuidEntry>.Create;
  try
    for I := 0 to High(AFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(I + 1, Length(AFiles), AFiles[I]);
      All.AddRange(ScanSingleFile(AFiles[I]));
    end;
    Result := All.ToArray;
  finally
    All.Free;
  end;
  MarkDuplicates(Result);
end;

end.
