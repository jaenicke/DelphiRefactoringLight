(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Delphi.FileEncoding;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils;

type
  /// <summary>Helpers for reading and writing Delphi source files with
  ///  automatic encoding detection. Grouped into a class to keep the unit's
  ///  public API free of global routines.</summary>
  TDelphiFileEncoding = class
  public
    /// <summary>Detects the encoding of a Delphi source file.
    ///  UTF-8 BOM -> UTF-8, UTF-16 LE/BE BOM -> UTF-16 LE/BE; without a BOM
    ///  the CONTENT decides: valid UTF-8 -> UTF-8 WITHOUT preamble (so a
    ///  write-back does not add a BOM), otherwise ANSI (system code page).
    ///  Pure ASCII counts as UTF-8 - both agree there. The returned
    ///  instance is shared and must NOT be freed.</summary>
    class function Detect(const AFilePath: string): TEncoding; static;
    /// <summary>The same decision on bytes already in memory.
    ///  APreambleLen receives the length of the BOM (0 without).</summary>
    class function DetectBytes(const ABytes: TBytes; ACount: Integer;
      out APreambleLen: Integer): TEncoding; static;
    /// <summary>Well-formed UTF-8 (no overlongs, no surrogates,
    ///  &lt;= U+10FFFF) in ABytes[AStart .. AStart + ALen - 1]?</summary>
    class function IsValidUtf8(const ABytes: TBytes; AStart, ALen: Integer): Boolean; static;
    /// <summary>The shared UTF-8 encoding without preamble.</summary>
    class function UTF8NoBOM: TEncoding; static;

    /// <summary>Reads a Delphi source file as a single string using the
    ///  auto-detected encoding.</summary>
    class function ReadAll(const AFilePath: string): string; static;

    /// <summary>Reads a Delphi source file line by line using the
    ///  auto-detected encoding (CRLF, LF and CR all end a line; a final
    ///  line break does not produce an empty last line).</summary>
    class function ReadLines(const AFilePath: string): TArray<string>; static;

    /// <summary>Writes a Delphi source file with the given encoding. The
    ///  caller typically passes back the encoding originally detected.
    ///  ATOMIC: the content goes to a temporary file beside the target,
    ///  which then replaces it - a failing write never leaves the unit
    ///  empty or half-written. An ANSI encoding that cannot represent the
    ///  content (characters outside the code page) is upgraded to UTF-8
    ///  with BOM instead of silently writing '?'.</summary>
    class procedure WriteAll(const AFilePath, AContent: string; AEncoding: TEncoding); static;
  end;

// Backwards-compat free-function wrappers. These forward to
// TDelphiFileEncoding and exist so existing callers in other units keep
// compiling during the gradual migration.
function DetectFileEncoding(const AFilePath: string): TEncoding;
function ReadDelphiFile(const AFilePath: string): string;
function ReadDelphiFileLines(const AFilePath: string): TArray<string>;
procedure WriteDelphiFile(const AFilePath, AContent: string; AEncoding: TEncoding);

implementation

{$IFDEF MSWINDOWS}
uses
  Winapi.Windows;
{$ENDIF}

type
  TUTF8NoBOMEncoding = class(TUTF8Encoding)
  public
    function GetPreamble: TBytes; override;
  end;

function TUTF8NoBOMEncoding.GetPreamble: TBytes;
begin
  Result := nil;
end;

var
  GUTF8NoBOM: TEncoding;

{ TDelphiFileEncoding }

class function TDelphiFileEncoding.UTF8NoBOM: TEncoding;
var
  E: TEncoding;
begin
  if GUTF8NoBOM = nil then
  begin
    E := TUTF8NoBOMEncoding.Create;
    if AtomicCmpExchange(Pointer(GUTF8NoBOM), Pointer(E), nil) <> nil then
      E.Free;
  end;
  Result := GUTF8NoBOM;
end;

class function TDelphiFileEncoding.IsValidUtf8(const ABytes: TBytes;
  AStart, ALen: Integer): Boolean;
var
  I, Stop, Need: Integer;
  C: Byte;
  CP: Cardinal;
begin
  I := AStart;
  Stop := AStart + ALen;
  while I < Stop do
  begin
    C := ABytes[I];
    if C < $80 then begin Inc(I); Continue; end;
    if (C and $E0) = $C0 then begin Need := 1; CP := C and $1F; end
    else if (C and $F0) = $E0 then begin Need := 2; CP := C and $0F; end
    else if (C and $F8) = $F0 then begin Need := 3; CP := C and $07; end
    else Exit(False);
    if I + Need >= Stop then Exit(False);
    for var K := 1 to Need do
    begin
      if (ABytes[I + K] and $C0) <> $80 then Exit(False);
      CP := (CP shl 6) or (ABytes[I + K] and $3F);
    end;
    case Need of
      1: if CP < $80 then Exit(False);
      2: if (CP < $800) or ((CP >= $D800) and (CP <= $DFFF)) then Exit(False);
      3: if (CP < $10000) or (CP > $10FFFF) then Exit(False);
    end;
    Inc(I, Need + 1);
  end;
  Result := True;
end;

class function TDelphiFileEncoding.DetectBytes(const ABytes: TBytes;
  ACount: Integer; out APreambleLen: Integer): TEncoding;
begin
  APreambleLen := 0;
  if (ACount >= 3) and (ABytes[0] = $EF) and (ABytes[1] = $BB) and (ABytes[2] = $BF) then
  begin
    APreambleLen := 3;
    Exit(TEncoding.UTF8);
  end;
  if (ACount >= 2) and (ABytes[0] = $FF) and (ABytes[1] = $FE) then
  begin
    APreambleLen := 2;
    Exit(TEncoding.Unicode);          // UTF-16 LE
  end;
  if (ACount >= 2) and (ABytes[0] = $FE) and (ABytes[1] = $FF) then
  begin
    APreambleLen := 2;
    Exit(TEncoding.BigEndianUnicode); // UTF-16 BE
  end;
  // No BOM: the content decides, as the RTL's one-argument
  // TFile.ReadAllText does. Before, this was ANSI unconditionally - a
  // BOM-less UTF-8 unit then lost every non-ASCII character to '?' on the
  // first write-back, irrecoverably.
  if IsValidUtf8(ABytes, 0, ACount) then
    Result := UTF8NoBOM
  else
    Result := TEncoding.Default;       // ANSI
end;

class function TDelphiFileEncoding.Detect(const AFilePath: string): TEncoding;
var
  Bytes: TBytes;
  Dummy: Integer;
begin
  if not FileExists(AFilePath) then
    Exit(TEncoding.Default);
  Bytes := TFile.ReadAllBytes(AFilePath);
  Result := DetectBytes(Bytes, Length(Bytes), Dummy);
end;

class function TDelphiFileEncoding.ReadAll(const AFilePath: string): string;
var
  Bytes: TBytes;
  Enc: TEncoding;
  Pre: Integer;
begin
  Bytes := TFile.ReadAllBytes(AFilePath);
  Enc := DetectBytes(Bytes, Length(Bytes), Pre);
  if Length(Bytes) - Pre <= 0 then Exit('');
  Result := Enc.GetString(Bytes, Pre, Length(Bytes) - Pre);
end;

class function TDelphiFileEncoding.ReadLines(const AFilePath: string): TArray<string>;
var
  S: string;
  N, I, Start, Count: Integer;
begin
  S := ReadAll(AFilePath);
  N := Length(S);
  Result := nil;
  if N = 0 then Exit;
  SetLength(Result, 64);
  Count := 0;
  Start := 1;
  I := 1;
  while I <= N do
  begin
    if (S[I] = #13) or (S[I] = #10) then
    begin
      if Count = Length(Result) then SetLength(Result, Count * 2);
      Result[Count] := Copy(S, Start, I - Start);
      Inc(Count);
      if (S[I] = #13) and (I < N) and (S[I + 1] = #10) then Inc(I);
      Start := I + 1;
    end;
    Inc(I);
  end;
  if Start <= N then
  begin
    if Count = Length(Result) then SetLength(Result, Count + 1);
    Result[Count] := Copy(S, Start, N - Start + 1);
    Inc(Count);
  end;
  SetLength(Result, Count);
end;

class procedure TDelphiFileEncoding.WriteAll(const AFilePath, AContent: string;
  AEncoding: TEncoding);
var
  Enc: TEncoding;
  Data, Pre: TBytes;
  Tmp: string;
  FS: TFileStream;
  Replaced: Boolean;
begin
  Enc := AEncoding;
  if Enc = nil then Enc := TEncoding.Default;
  // An ANSI code page cannot hold every character. Writing anyway would
  // substitute '?' - data loss nobody can undo. UTF-8 with BOM reads the
  // same in Delphi and keeps everything.
  if (Enc.CodePage <> 65001) and (Enc.CodePage <> 1200) and (Enc.CodePage <> 1201) and
     (Enc.GetString(Enc.GetBytes(AContent)) <> AContent) then
    Enc := TEncoding.UTF8;
  Pre := Enc.GetPreamble;
  Data := Enc.GetBytes(AContent);

  Tmp := AFilePath + '.rl~tmp';
  try
    FS := TFileStream.Create(Tmp, fmCreate);
    try
      if Length(Pre) > 0 then FS.WriteBuffer(Pre[0], Length(Pre));
      if Length(Data) > 0 then FS.WriteBuffer(Data[0], Length(Data));
      {$IFDEF MSWINDOWS}
      FlushFileBuffers(FS.Handle);
      {$ENDIF}
    finally
      FS.Free;
    end;
    {$IFDEF MSWINDOWS}
    Replaced := False;
    // ReplaceFile keeps the target's attributes, ACL and creation time;
    // it needs an existing target, MoveFileEx covers the rest.
    if FileExists(AFilePath) then
      Replaced := ReplaceFile(PChar(AFilePath), PChar(Tmp), nil,
        REPLACEFILE_IGNORE_MERGE_ERRORS, nil, nil);
    if not Replaced then
      if not MoveFileEx(PChar(Tmp), PChar(AFilePath),
           MOVEFILE_REPLACE_EXISTING or MOVEFILE_WRITE_THROUGH) then
        RaiseLastOSError;
    {$ELSE}
    TFile.Copy(Tmp, AFilePath, True);
    TFile.Delete(Tmp);
    {$ENDIF}
  except
    if FileExists(Tmp) then
      try System.SysUtils.DeleteFile(Tmp); except end;
    raise;
  end;
end;

{ Backwards-compat wrappers }

function DetectFileEncoding(const AFilePath: string): TEncoding;
begin
  Result := TDelphiFileEncoding.Detect(AFilePath);
end;

function ReadDelphiFile(const AFilePath: string): string;
begin
  Result := TDelphiFileEncoding.ReadAll(AFilePath);
end;

function ReadDelphiFileLines(const AFilePath: string): TArray<string>;
begin
  Result := TDelphiFileEncoding.ReadLines(AFilePath);
end;

procedure WriteDelphiFile(const AFilePath, AContent: string; AEncoding: TEncoding);
begin
  TDelphiFileEncoding.WriteAll(AFilePath, AContent, AEncoding);
end;

initialization

finalization
  FreeAndNil(GUTF8NoBOM);

end.
