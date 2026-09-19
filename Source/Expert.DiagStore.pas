(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DiagStore;

{$OVERFLOWCHECKS OFF}
{$RANGECHECKS OFF}

// The RAW diagnostics each source last reported, per file and per source.
//
// The live checker keeps only the RESOLVED quick fixes, and only for the
// active buffer. The MCP bridge needs the diagnostics themselves, for any
// file, and must know whether they still describe the buffer: every entry
// carries the hash of the content it was computed for, and a consumer
// compares that with the current buffer - a mismatch means "stale", never
// "no errors".
//
// Thread-safe (the LSP worker writes from a background thread, the MCP
// handler reads from another). Bounded: at most MaxFiles files, the least
// recently written one is dropped first.

interface

uses
  System.SysUtils, Lsp.Protocol;

type
  TStoredDiags = record
    Source: string;          // SrcStructure / SrcCompiler / SrcLsp
    ContentHash: Cardinal;   // DiagContentHash of the analysed content
    Diags: TArray<TLspErrorDiag>;
    Stamp: TDateTime;
  end;

function DiagContentHash(const AContent: string): Cardinal;

procedure StoreDiagnostics(const AFile, ASource, AContent: string;
  const ADiags: TArray<TLspErrorDiag>);
procedure StoreDiagnosticsHashed(const AFile, ASource: string; AHash: Cardinal;
  const ADiags: TArray<TLspErrorDiag>);
function StoredDiagnostics(const AFile: string): TArray<TStoredDiags>;
procedure ClearStoredDiagnostics;

/// <summary>Merges the sources whose hash matches AHash. Duplicates (same
///  code and start position) are reported once; ASources[i] lists the
///  sources that reported Result[i] ("structure+lsp"). AUsed / AStale name
///  the sources that were (not) current for this content.</summary>
function MergeFreshDiagnostics(const ASets: TArray<TStoredDiags>; AHash: Cardinal;
  out ASources: TArray<string>; out AUsed, AStale: string): TArray<TLspErrorDiag>;

implementation

uses
  System.Generics.Collections, System.SyncObjs;

const
  MaxFiles = 64;

type
  TFileEntry = class
    Sets: TArray<TStoredDiags>;
    Touched: Int64;
  end;

var
  GLock: TCriticalSection;
  GFiles: TObjectDictionary<string, TFileEntry>;
  GClock: Int64;

function DiagContentHash(const AContent: string): Cardinal;
var
  I: Integer;
begin
  // FNV-1a over the UTF-16 code units - fast enough for a 10,000-line
  // unit on every Structure view notification (< 1 ms).
  Result := 2166136261;
  for I := 1 to Length(AContent) do
    Result := (Result xor Ord(AContent[I])) * 16777619;
end;

function Key(const AFile: string): string;
begin
  Result := UpperCase(ExpandFileName(AFile));
end;

procedure StoreDiagnosticsHashed(const AFile, ASource: string; AHash: Cardinal;
  const ADiags: TArray<TLspErrorDiag>);
var
  E: TFileEntry;
  S: TStoredDiags;
  I: Integer;
begin
  if (AFile = '') or (ASource = '') then Exit;
  S.Source := ASource;
  S.ContentHash := AHash;
  S.Diags := Copy(ADiags);
  S.Stamp := Now;
  GLock.Enter;
  try
    if not GFiles.TryGetValue(Key(AFile), E) then
    begin
      if GFiles.Count >= MaxFiles then
      begin
        var OldKey := '';
        var OldT := High(Int64);
        for var P in GFiles do
          if P.Value.Touched < OldT then
          begin
            OldT := P.Value.Touched;
            OldKey := P.Key;
          end;
        if OldKey <> '' then GFiles.Remove(OldKey);
      end;
      E := TFileEntry.Create;
      GFiles.Add(Key(AFile), E);
    end;
    Inc(GClock);
    E.Touched := GClock;
    for I := 0 to High(E.Sets) do
      if SameText(E.Sets[I].Source, ASource) then
      begin
        E.Sets[I] := S;
        Exit;
      end;
    E.Sets := E.Sets + [S];
  finally
    GLock.Leave;
  end;
end;

procedure StoreDiagnostics(const AFile, ASource, AContent: string;
  const ADiags: TArray<TLspErrorDiag>);
begin
  StoreDiagnosticsHashed(AFile, ASource, DiagContentHash(AContent), ADiags);
end;

function StoredDiagnostics(const AFile: string): TArray<TStoredDiags>;
var
  E: TFileEntry;
begin
  Result := nil;
  GLock.Enter;
  try
    if GFiles.TryGetValue(Key(AFile), E) then
      Result := Copy(E.Sets);
  finally
    GLock.Leave;
  end;
end;

procedure ClearStoredDiagnostics;
begin
  GLock.Enter;
  try
    GFiles.Clear;
  finally
    GLock.Leave;
  end;
end;

function MergeFreshDiagnostics(const ASets: TArray<TStoredDiags>; AHash: Cardinal;
  out ASources: TArray<string>; out AUsed, AStale: string): TArray<TLspErrorDiag>;

  procedure AddName(var AList: string; const AName: string);
  begin
    if AList <> '' then AList := AList + ', ';
    AList := AList + AName;
  end;

var
  Index: TDictionary<string, Integer>;
  ByLine: TDictionary<string, Integer>;   // code|line -> index, -1 = several
  K, KL: string;
  Idx: Integer;

  function ZeroLen(const R: TLspRange): Boolean;
  begin
    Result := (R.Start.Line = R.End_.Line) and (R.Start.Character = R.End_.Character);
  end;

begin
  Result := nil;
  ASources := nil;
  AUsed := '';
  AStale := '';
  Index := TDictionary<string, Integer>.Create;
  ByLine := TDictionary<string, Integer>.Create;
  try
    for var S in ASets do
    begin
      if S.ContentHash <> AHash then
      begin
        AddName(AStale, S.Source);
        Continue;
      end;
      AddName(AUsed, S.Source);
      for var D in S.Diags do
      begin
        K := UpperCase(D.Code) + '|' + IntToStr(D.Range.Start.Line) + '|' +
          IntToStr(D.Range.Start.Character);
        KL := UpperCase(D.Code) + '|' + IntToStr(D.Range.Start.Line);
        Idx := -1;
        if not Index.TryGetValue(K, Idx) then
        begin
          Idx := -1;
          // The Structure view reports a POSITION only (zero-length range)
          // and its column may differ from the LSP's range start: the same
          // error on the same line, when it is the only one of its code
          // there, is still one error.
          var Cand: Integer;
          if ByLine.TryGetValue(KL, Cand) and (Cand >= 0) and
             (ZeroLen(D.Range) or ZeroLen(Result[Cand].Range)) and
             (Pos(S.Source, ASources[Cand]) = 0) then
            Idx := Cand;
        end;
        if Idx >= 0 then
        begin
          if Pos(S.Source, ASources[Idx]) = 0 then
            ASources[Idx] := ASources[Idx] + '+' + S.Source;
          // a zero-length range gains the real extent
          if ZeroLen(Result[Idx].Range) and not ZeroLen(D.Range) then
            Result[Idx].Range := D.Range;
          Continue;
        end;
        Index.Add(K, Length(Result));
        if ByLine.ContainsKey(KL) then
          ByLine[KL] := -1
        else
          ByLine.Add(KL, Length(Result));
        Result := Result + [D];
        ASources := ASources + [S.Source];
      end;
    end;
  finally
    ByLine.Free;
    Index.Free;
  end;
end;

initialization
  GLock := TCriticalSection.Create;
  GFiles := TObjectDictionary<string, TFileEntry>.Create([doOwnsValues]);

finalization
  FreeAndNil(GFiles);
  FreeAndNil(GLock);

end.
