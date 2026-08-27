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
  Expert.EditorHelperIntf, Delphi.FileEncoding;

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

/// <summary>Extracts "['{...}']" from a line; returns '' if absent.</summary>
function ExtractGuid(const ALine: string): string;
var
  P1, P2: Integer;
begin
  Result := '';
  P1 := Pos('[''{', ALine);
  if P1 = 0 then Exit;
  P2 := PosEx('}'']', ALine, P1);
  if P2 = 0 then Exit;
  Result := Copy(ALine, P1 + 2, P2 - P1 - 1);  // {....}
end;

function StripLineComment(const ALine: string): string;
var
  P: Integer;
begin
  Result := ALine;
  P := Pos('//', Result);
  if P > 0 then Result := Copy(Result, 1, P - 1);
end;

class function TInterfaceGuidChecker.ScanSingleFile(
  const AFile: string): TArray<TInterfaceGuidEntry>;
var
  Entries: TList<TInterfaceGuidEntry>;
  L, U, Name, Guid: string;
  Lines: TArray<string>;
  I, J, EqPos: Integer;
  E: TInterfaceGuidEntry;
begin
  Result := nil;
  if not SameText(ExtractFileExt(AFile), '.pas') then Exit;
  Entries := TList<TInterfaceGuidEntry>.Create;
  try
    Lines := ReadFileLines(AFile);
    for I := 0 to High(Lines) do
    begin
      L := StripLineComment(Lines[I]);
      U := UpperCase(L);
      // "IFoo = interface" or "= dispinterface"; reject forward
      // declarations ("= interface;") - they never carry the GUID,
      // the full declaration elsewhere does.
      EqPos := Pos('= INTERFACE', U);
      if EqPos = 0 then EqPos := Pos('= DISPINTERFACE', U);
      if EqPos = 0 then Continue;
      var AfterKw := Trim(Copy(U, EqPos + Length('= INTERFACE'), MaxInt));
      if U.Contains('= DISPINTERFACE') then
        AfterKw := Trim(Copy(U, Pos('= DISPINTERFACE', U) + Length('= DISPINTERFACE'), MaxInt));
      if AfterKw.StartsWith(';') then Continue; // forward decl

      Name := Trim(Copy(L, 1, EqPos - 1));
      if (Name = '') or not CharInSet(Name[1], ['A'..'Z', 'a'..'z', '_']) then Continue;

      // GUID on the same line or within the next 3 lines.
      Guid := ExtractGuid(L);
      J := I;
      while (Guid = '') and (J < High(Lines)) and (J < I + 3) do
      begin
        Inc(J);
        Guid := ExtractGuid(StripLineComment(Lines[J]));
        // Stop early if the next declaration starts.
        if Pos('= INTERFACE', UpperCase(Lines[J])) > 0 then Break;
      end;

      E := Default(TInterfaceGuidEntry);
      E.InterfaceName := Name;
      E.Guid := Guid;
      E.FileName := AFile;
      E.Line := I + 1;
      E.HasGuid := Guid <> '';
      E.IsDispInterface := Pos('= DISPINTERFACE', U) > 0;
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
