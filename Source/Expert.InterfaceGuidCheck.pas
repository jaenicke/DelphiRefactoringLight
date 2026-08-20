(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
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
  end;

  TInterfaceGuidChecker = class
  public
    /// <summary>Scans AFiles (only .pas are considered) and returns
    ///  every interface declaration found. Entries with duplicate
    ///  GUIDs have IsDuplicate = True.</summary>
    class function Scan(const AFiles: TArray<string>): TArray<TInterfaceGuidEntry>;
  end;

implementation

uses
  System.IOUtils, System.StrUtils,
  Expert.EditorHelperIntf, Delphi.FileEncoding;

function ReadFileLines(const AFile: string): TArray<string>;
var
  Content: string;
begin
  Result := nil;
  if (Editor <> nil) and Editor.ReadEditorContent(AFile, Content) then
    Exit(Content.Split([sLineBreak], TStringSplitOptions.None));
  if not TFile.Exists(AFile) then Exit;
  try
    Content := TDelphiFileEncoding.ReadAll(AFile);
    Result := Content.Split([sLineBreak], TStringSplitOptions.None);
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

class function TInterfaceGuidChecker.Scan(
  const AFiles: TArray<string>): TArray<TInterfaceGuidEntry>;
var
  Entries: TList<TInterfaceGuidEntry>;
  GuidCount: TDictionary<string, Integer>;
  F, L, U, Name, Guid: string;
  Lines: TArray<string>;
  I, J, EqPos: Integer;
  E: TInterfaceGuidEntry;
begin
  Entries := TList<TInterfaceGuidEntry>.Create;
  GuidCount := TDictionary<string, Integer>.Create;
  try
    for F in AFiles do
    begin
      if not SameText(ExtractFileExt(F), '.pas') then Continue;
      Lines := ReadFileLines(F);
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
        E.FileName := F;
        E.Line := I + 1;
        E.HasGuid := Guid <> '';
        Entries.Add(E);

        if Guid <> '' then
        begin
          var Key := UpperCase(Guid);
          var C: Integer;
          if GuidCount.TryGetValue(Key, C) then
            GuidCount[Key] := C + 1
          else
            GuidCount.Add(Key, 1);
        end;
      end;
    end;

    // Mark duplicates.
    for I := 0 to Entries.Count - 1 do
      if Entries[I].HasGuid then
      begin
        var C: Integer;
        if GuidCount.TryGetValue(UpperCase(Entries[I].Guid), C) and (C > 1) then
        begin
          E := Entries[I];
          E.IsDuplicate := True;
          Entries[I] := E;
        end;
      end;

    Result := Entries.ToArray;
  finally
    GuidCount.Free;
    Entries.Free;
  end;
end;

end.
