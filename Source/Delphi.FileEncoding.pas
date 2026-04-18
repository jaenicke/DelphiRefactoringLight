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
  ///  automatic BOM-based encoding detection. Grouped into a class to
  ///  keep the unit's public API free of global routines.</summary>
  TDelphiFileEncoding = class
  public
    /// <summary>Detects the encoding of a Delphi source file by its BOM.
    ///  UTF-8 BOM -> UTF-8, UTF-16 LE BOM -> UTF-16 LE,
    ///  UTF-16 BE BOM -> UTF-16 BE, otherwise -> ANSI (system default).</summary>
    class function Detect(const AFilePath: string): TEncoding; static;

    /// <summary>Reads a Delphi source file as a single string using the
    ///  auto-detected encoding.</summary>
    class function ReadAll(const AFilePath: string): string; static;

    /// <summary>Reads a Delphi source file line by line using the
    ///  auto-detected encoding.</summary>
    class function ReadLines(const AFilePath: string): TArray<string>; static;

    /// <summary>Writes a Delphi source file with the given encoding. The
    ///  caller typically passes back the encoding originally detected.</summary>
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

{ TDelphiFileEncoding }

class function TDelphiFileEncoding.Detect(const AFilePath: string): TEncoding;
var
  Stream: TFileStream;
  BOM: array[0..2] of Byte;
  BytesRead: Integer;
begin
  Result := TEncoding.Default; // ANSI (Windows-1252) as fallback

  if not FileExists(AFilePath) then
    Exit;

  Stream := TFileStream.Create(AFilePath, fmOpenRead or fmShareDenyNone);
  try
    BytesRead := Stream.Read(BOM, 3);
    if (BytesRead >= 3) and (BOM[0] = $EF) and (BOM[1] = $BB) and (BOM[2] = $BF) then
      Result := TEncoding.UTF8
    else if (BytesRead >= 2) and (BOM[0] = $FF) and (BOM[1] = $FE) then
      Result := TEncoding.Unicode          // UTF-16 LE
    else if (BytesRead >= 2) and (BOM[0] = $FE) and (BOM[1] = $FF) then
      Result := TEncoding.BigEndianUnicode // UTF-16 BE
    else
      Result := TEncoding.Default;         // ANSI
  finally
    Stream.Free;
  end;
end;

class function TDelphiFileEncoding.ReadAll(const AFilePath: string): string;
begin
  Result := TFile.ReadAllText(AFilePath, Detect(AFilePath));
end;

class function TDelphiFileEncoding.ReadLines(const AFilePath: string): TArray<string>;
begin
  Result := TFile.ReadAllLines(AFilePath, Detect(AFilePath));
end;

class procedure TDelphiFileEncoding.WriteAll(const AFilePath, AContent: string;
  AEncoding: TEncoding);
begin
  TFile.WriteAllText(AFilePath, AContent, AEncoding);
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

end.
