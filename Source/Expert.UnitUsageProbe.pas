(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnitUsageProbe;

{
  "Is this unit referenced anywhere?" - fast enough to answer BEFORE a
  dialog is shown (Save As of a unit).

  The expensive answer is the NEGATIVE one: proving that a name occurs
  nowhere means every file must be looked at, there is no early exit.
  Three things make that affordable:

  1. RAW BYTES. The file is searched as bytes, case-insensitively - no
     encoding detection, no UTF-16 conversion of the whole content, no
     UpperCase() copy of every file. A unit name is ASCII, so this is
     exact for ANSI and UTF-8 (the two encodings Delphi sources use).
     A UTF-16 file (BOM) is decoded on the slow path - rare.
  2. PARALLEL. The scan is I/O bound and the OS cache makes repeat runs
     nearly free.
  3. A TIME BUDGET. When the budget is spent the answer is uurUnknown
     and the CALLER keeps its old behaviour. Being slow must never turn
     into being wrong.

  Comments and strings are NOT excluded: a mention in a comment yields
  "used", which only means the dialog appears - the safe direction.
}

interface

uses
  System.SysUtils, System.Types;

type
  TUnitUsageResult = (
    uurUsed,      // at least one file references the name
    uurUnused,    // no file does - checked completely
    uurUnknown);  // budget exceeded / unreadable file: no answer

const
  /// <summary>Default budget. A "Save As" must not feel blocked; when
  ///  the project is too large to answer in time the dialog simply
  ///  appears as before.</summary>
  DefaultUnitUsageBudgetMs = 1200;

/// <summary>True when ABytes reference AUnitName as a WHOLE WORD
///  (case-insensitive, ASCII). Pure and exported for the tests.</summary>
function BytesReferenceUnit(const ABytes: TBytes; const AUnitName: string): Boolean;

/// <summary>Same for already-decoded text (open editor buffers).</summary>
function TextReferencesUnit(const AContent, AUnitName: string): Boolean;

/// <summary>Does any of AFiles reference AUnitName? Files open in the
///  editor are read from the BUFFER (unsaved changes) and skipped in the
///  disk pass. AFirstHit names the file that answered, AElapsedMs says
///  how long it took.</summary>
function ProbeUnitUsage(const AFiles: TArray<string>; const AUnitName: string;
  ABudgetMs: Integer; out AFirstHit: string; out AElapsedMs: Integer): TUnitUsageResult;

implementation

uses
  System.Classes, System.IOUtils, System.Threading, System.SyncObjs,
  System.Diagnostics, System.Generics.Collections, Expert.EditorHelperIntf;

function IsIdentByte(B: Byte): Boolean; inline;
begin
  Result := ((B >= Ord('A')) and (B <= Ord('Z')))
    or ((B >= Ord('a')) and (B <= Ord('z')))
    or ((B >= Ord('0')) and (B <= Ord('9')))
    or (B = Ord('_'));
end;

function LowerByte(B: Byte): Byte; inline;
begin
  if (B >= Ord('A')) and (B <= Ord('Z')) then
    Result := B + 32
  else
    Result := B;
end;

function BytesReferenceUnit(const ABytes: TBytes; const AUnitName: string): Boolean;
var
  Pat: TBytes;
  N, M, I, J: Integer;
begin
  Result := False;
  M := Length(AUnitName);
  N := Length(ABytes);
  if (M = 0) or (N < M) then Exit;

  // A UTF-16 file cannot be scanned byte-wise - decode it instead.
  if (N >= 2) and (((ABytes[0] = $FF) and (ABytes[1] = $FE))
    or ((ABytes[0] = $FE) and (ABytes[1] = $FF))) then
    Exit(TextReferencesUnit(TEncoding.Unicode.GetString(ABytes), AUnitName));

  SetLength(Pat, M);
  for I := 1 to M do
    Pat[I - 1] := LowerByte(Byte(Ord(AUnitName[I])));

  for I := 0 to N - M do
  begin
    if LowerByte(ABytes[I]) <> Pat[0] then Continue;
    J := 1;
    while (J < M) and (LowerByte(ABytes[I + J]) = Pat[J]) do Inc(J);
    if J < M then Continue;
    // Whole word: an identifier character on either side disqualifies.
    // A dot does NOT - 'Other.OldUnit' is a reference too.
    if (I > 0) and IsIdentByte(ABytes[I - 1]) then Continue;
    if (I + M < N) and IsIdentByte(ABytes[I + M]) then Continue;
    Exit(True);
  end;
end;

function TextReferencesUnit(const AContent, AUnitName: string): Boolean;
var
  N, M, I, J: Integer;

  function IsIdentChar(C: Char): Boolean;
  begin
    Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
  end;

begin
  Result := False;
  M := Length(AUnitName);
  N := Length(AContent);
  if (M = 0) or (N < M) then Exit;

  for I := 1 to N - M + 1 do
  begin
    J := 0;
    while (J < M) and (UpCase(AContent[I + J]) = UpCase(AUnitName[J + 1])) do
      Inc(J);
    if J < M then Continue;
    if (I > 1) and IsIdentChar(AContent[I - 1]) then Continue;
    if (I + M <= N) and IsIdentChar(AContent[I + M]) then Continue;
    Exit(True);
  end;
end;

// A buffer-read file must not be read from disk again - and the two
// lists come from different ToolsAPI calls, so compare them normalised.
function NormKey(const AFile: string): string;
begin
  try
    Result := LowerCase(ExpandFileName(AFile));
  except
    Result := LowerCase(AFile);
  end;
end;

function ProbeUnitUsage(const AFiles: TArray<string>; const AUnitName: string;
  ABudgetMs: Integer; out AFirstHit: string; out AElapsedMs: Integer): TUnitUsageResult;
var
  SW: TStopwatch;
  Found, Expired: Integer;
  HitFile: string;
  HitLock: TCriticalSection;
  Live, Arr: TArray<string>;
  DiskFiles: TList<string>;
  Content: string;
  Seen: TDictionary<string, Boolean>;
begin
  AFirstHit := '';
  AElapsedMs := 0;
  Result := uurUnknown;
  if (AUnitName = '') or (Length(AFiles) = 0) then Exit;

  SW := TStopwatch.StartNew;
  Found := 0;
  Expired := 0;
  HitFile := '';

  Seen := TDictionary<string, Boolean>.Create;
  DiskFiles := TList<string>.Create;
  HitLock := TCriticalSection.Create;
  try
    // ---- Open buffers first: they hold the UNSAVED text, and there are
    // only a handful of them. Must run on the calling thread (ToolsAPI).
    try
      Live := Editor.GetOpenSourceFiles;
    except
      Live := nil;
    end;
    for var LF in Live do
      if Editor.ReadEditorContent(LF, Content) then
      begin
        Seen.AddOrSetValue(NormKey(LF), True);
        if TextReferencesUnit(Content, AUnitName) then
        begin
          AFirstHit := LF;
          AElapsedMs := Integer(SW.ElapsedMilliseconds);
          Exit(uurUsed);
        end;
      end;

    // THE PROJECT FILE is the trap this whole probe stumbled over first:
    // on "Save as" the IDE rewrites the uses entry in its BUFFER, while
    // the .dpr on disk keeps the OLD unit name until the project is
    // saved - so a disk read reports "still used" forever. Its module is
    // named after the .dproj, so ask for the .dpr explicitly.
    var DprFile := '';
    try
      DprFile := Editor.GetCurrentProjectDproj;
      if DprFile <> '' then DprFile := ChangeFileExt(DprFile, '.dpr');
    except
      DprFile := '';
    end;
    if (DprFile <> '') and not Seen.ContainsKey(NormKey(DprFile))
      and Editor.ReadEditorContent(DprFile, Content) then
    begin
      Seen.AddOrSetValue(NormKey(DprFile), True);
      if TextReferencesUnit(Content, AUnitName) then
      begin
        AFirstHit := DprFile;
        AElapsedMs := Integer(SW.ElapsedMilliseconds);
        Exit(uurUsed);
      end;
    end;

    for var F in AFiles do
      if not Seen.ContainsKey(NormKey(F)) then
        DiskFiles.Add(F);

    if DiskFiles.Count = 0 then
    begin
      AElapsedMs := Integer(SW.ElapsedMilliseconds);
      Exit(uurUnused);
    end;

    Arr := DiskFiles.ToArray;
    TParallel.For(0, High(Arr),
      procedure(AIndex: Integer; ALoopState: TParallel.TLoopState)
      var
        Bytes: TBytes;
      begin
        if (Found <> 0) or (Expired <> 0) then
        begin
          ALoopState.Stop;
          Exit;
        end;
        if SW.ElapsedMilliseconds > ABudgetMs then
        begin
          TInterlocked.Exchange(Expired, 1);
          ALoopState.Stop;
          Exit;
        end;
        try
          Bytes := TFile.ReadAllBytes(Arr[AIndex]);
        except
          // Unreadable file: absence can no longer be PROVEN.
          TInterlocked.Exchange(Expired, 1);
          Exit;
        end;
        if BytesReferenceUnit(Bytes, AUnitName) then
        begin
          HitLock.Enter;
          try
            if HitFile = '' then HitFile := Arr[AIndex];
          finally
            HitLock.Leave;
          end;
          TInterlocked.Exchange(Found, 1);
          ALoopState.Stop;
        end;
      end);

    AElapsedMs := Integer(SW.ElapsedMilliseconds);
    if Found <> 0 then
    begin
      AFirstHit := HitFile;
      Result := uurUsed;
    end
    else if Expired <> 0 then
      Result := uurUnknown
    else
      Result := uurUnused;
  finally
    HitLock.Free;
    DiskFiles.Free;
    Seen.Free;
  end;
end;

end.
