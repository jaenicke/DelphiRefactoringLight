(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Lsp.Uri;

interface

uses
  System.SysUtils, System.NetEncoding;

type
  /// <summary>Conversion between Windows file paths and "file://" URIs
  ///  as required by the Language Server Protocol.</summary>
  TLspUri = class
  public
    /// <summary>Converts a Windows file path into a "file://" URI.</summary>
    class function PathToFileUri(const APath: string): string; static;

    /// <summary>Converts a "file://" URI back into a Windows file path.</summary>
    class function FileUriToPath(const AUri: string): string; static;
  end;

implementation

type
  /// <summary>File-private helpers for percent-encoding/decoding paths.</summary>
  TLspUriHelper = class
  public
    class function PercentEncodePath(const APath: string): string; static;
    class function HexDigitValue(Ch: Char): Integer; static;
    class function PercentDecode(const S: string): string; static;
  end;

{ TLspUriHelper }

class function TLspUriHelper.PercentEncodePath(const APath: string): string;
var
  I: Integer;
  Ch: Char;
begin
  Result := '';
  for I := 1 to Length(APath) do
  begin
    Ch := APath[I];
    case Ch of
      'A'..'Z', 'a'..'z', '0'..'9',
      '-', '_', '.', '~', '/', ':':
        Result := Result + Ch;
    else
      // Percent-encode any other character as its UTF-8 bytes.
      var Bytes := TEncoding.UTF8.GetBytes(Ch);
      for var B in Bytes do
        Result := Result + '%' + IntToHex(B, 2);
    end;
  end;
end;

class function TLspUriHelper.HexDigitValue(Ch: Char): Integer;
begin
  case Ch of
    '0'..'9': Result := Ord(Ch) - Ord('0');
    'A'..'F': Result := Ord(Ch) - Ord('A') + 10;
    'a'..'f': Result := Ord(Ch) - Ord('a') + 10;
  else
    Result := 0;
  end;
end;

class function TLspUriHelper.PercentDecode(const S: string): string;
var
  I: Integer;
  Bytes: TBytes;
  ByteCount: Integer;
begin
  SetLength(Bytes, Length(S));
  ByteCount := 0;
  I := 1;
  while I <= Length(S) do
  begin
    if (S[I] = '%') and (I + 2 <= Length(S)) then
    begin
      Bytes[ByteCount] := Byte(HexDigitValue(S[I+1]) * 16 + HexDigitValue(S[I+2]));
      Inc(ByteCount);
      Inc(I, 3);
    end
    else
    begin
      // Take ASCII characters as-is.
      Bytes[ByteCount] := Byte(Ord(S[I]));
      Inc(ByteCount);
      Inc(I);
    end;
  end;
  Result := TEncoding.UTF8.GetString(Bytes, 0, ByteCount);
end;

{ TLspUri }

class function TLspUri.PathToFileUri(const APath: string): string;
var
  AbsPath: string;
begin
  AbsPath := ExpandFileName(APath);
  // Backslashes -> forward slashes
  AbsPath := StringReplace(AbsPath, '\', '/', [rfReplaceAll]);
  // Percent-encode (but preserve '/' and ':')
  Result := 'file:///' + TLspUriHelper.PercentEncodePath(AbsPath);
end;

class function TLspUri.FileUriToPath(const AUri: string): string;
var
  Path: string;
begin
  Path := AUri;
  // Strip "file:///" prefix
  if Path.StartsWith('file:///', True) then
    Path := Copy(Path, 9)
  else if Path.StartsWith('file://', True) then
    Path := Copy(Path, 8);

  // Percent-decode
  Path := TLspUriHelper.PercentDecode(Path);

  // Forward slashes -> backslashes (Windows)
  Path := StringReplace(Path, '/', '\', [rfReplaceAll]);

  Result := Path;
end;

end.
