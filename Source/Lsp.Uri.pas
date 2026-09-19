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
  // UNC "\\server\share\x.pas" -> "file://server/share/x.pas": the server
  // is the URI's AUTHORITY. The old code produced "file://///server/...".
  if AbsPath.StartsWith('\\') then
  begin
    AbsPath := StringReplace(Copy(AbsPath, 3, MaxInt), '\', '/', [rfReplaceAll]);
    Exit('file://' + TLspUriHelper.PercentEncodePath(AbsPath));
  end;
  // Backslashes -> forward slashes
  AbsPath := StringReplace(AbsPath, '\', '/', [rfReplaceAll]);
  // Percent-encode (but preserve '/' and ':')
  Result := 'file:///' + TLspUriHelper.PercentEncodePath(AbsPath);
end;

class function TLspUri.FileUriToPath(const AUri: string): string;
var
  Path: string;
  IsUnc: Boolean;
begin
  // Decode FIRST: DelphiLSP sends "file:///c%3A/..." - with the colon still
  // encoded the drive test below would take every local path for UNC.
  Path := TLspUriHelper.PercentDecode(AUri);
  IsUnc := False;
  if Path.StartsWith('file://', True) then
  begin
    Path := Copy(Path, 8);   // after "file://": authority + path
    if Path.StartsWith('localhost/', True) then
      Path := Copy(Path, 10);   // "file://localhost/C:/x" = local
    var SlashRun := 0;
    while Path.StartsWith('/') do
    begin
      Delete(Path, 1, 1);
      Inc(SlashRun);
    end;
    // UNC only with the shape "server/share/...": the authority form
    // "file://server/share/x" (no slash left after "file://") or extra
    // slashes before the server ("file:////server/...", and the
    // "file://///server/..." this unit itself used to emit). A drive path
    // is local, and so is anything without a share part - DelphiLSP
    // sometimes answers with a bare "file:///Unit.pas", which must stay
    // a (relative) file name instead of becoming "\\Unit.pas".
    var IsDrive := (Length(Path) >= 2) and (Path[2] = ':') and
      CharInSet(Path[1], ['A'..'Z', 'a'..'z']);
    var Slash := Pos('/', Path);
    var HasShare := (Slash > 1) and (Slash < Length(Path));
    IsUnc := not IsDrive and HasShare and (SlashRun <> 1);
  end;

  // Forward slashes -> backslashes (Windows)
  Path := StringReplace(Path, '/', '\', [rfReplaceAll]);

  if IsUnc then
    Path := '\\' + Path;
  Result := Path;
end;

end.
