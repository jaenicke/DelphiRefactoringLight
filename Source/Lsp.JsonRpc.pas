(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Lsp.JsonRpc;

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.SyncObjs;

type
  /// <summary>
  ///  Reads and writes JSON-RPC 2.0 messages with Content-Length framing
  ///  over arbitrary TStream pairs (typically: pipes to a subprocess).
  /// </summary>
  TJsonRpcTransport = class
  private
    FInputStream: TStream;   // Read (stdout of the LSP server)
    FOutputStream: TStream;  // Write (stdin of the LSP server)
    FWriteLock: TCriticalSection;
    function ReadLine: RawByteString;
    function ReadBytes(Count: Integer): TBytes;
  public
    constructor Create(AInputStream, AOutputStream: TStream);
    destructor Destroy; override;

    /// <summary>Sends a JSON-RPC message with Content-Length header.</summary>
    procedure SendMessage(AMessage: TJSONObject);

    /// <summary>Reads the next JSON-RPC message. Blocks until a message is available.</summary>
    function ReadMessage: TJSONObject;
  end;

implementation

{ TJsonRpcTransport }

constructor TJsonRpcTransport.Create(AInputStream, AOutputStream: TStream);
begin
  inherited Create;
  FInputStream := AInputStream;
  FOutputStream := AOutputStream;
  FWriteLock := TCriticalSection.Create;
end;

destructor TJsonRpcTransport.Destroy;
begin
  FWriteLock.Free;
  inherited;
end;

function TJsonRpcTransport.ReadLine: RawByteString;
var
  B: Byte;
begin
  Result := '';
  while True do
  begin
    if FInputStream.Read(B, 1) <> 1 then
      raise EStreamError.Create('Unerwartetes Ende des Input-Streams');
    if B = 13 then // CR
    begin
      // Expect LF
      if FInputStream.Read(B, 1) = 1 then
      begin
        if B = 10 then
          Exit;
        // No LF - treat CR as part of the line
        Result := Result + RawByteString(AnsiChar(13)) + RawByteString(AnsiChar(B));
      end
      else
        Exit;
    end
    else if B = 10 then // LF only
      Exit
    else
      Result := Result + RawByteString(AnsiChar(B));
  end;
end;

function TJsonRpcTransport.ReadBytes(Count: Integer): TBytes;
var
  Offset, BytesRead: Integer;
begin
  SetLength(Result, Count);
  Offset := 0;
  while Offset < Count do
  begin
    BytesRead := FInputStream.Read(Result[Offset], Count - Offset);
    if BytesRead <= 0 then
      raise EStreamError.Create('Unerwartetes Ende des Input-Streams');
    Inc(Offset, BytesRead);
  end;
end;

procedure TJsonRpcTransport.SendMessage(AMessage: TJSONObject);
var
  JsonStr: string;
  JsonBytes: TBytes;
  Header: RawByteString;
begin
  JsonStr := AMessage.ToJSON;
  JsonBytes := TEncoding.UTF8.GetBytes(JsonStr);

  Header := RawByteString('Content-Length: ' + IntToStr(Length(JsonBytes)) + #13#10#13#10);

  FWriteLock.Enter;
  try
    FOutputStream.WriteBuffer(Header[1], Length(Header));
    FOutputStream.WriteBuffer(JsonBytes[0], Length(JsonBytes));
  finally
    FWriteLock.Leave;
  end;
end;

function TJsonRpcTransport.ReadMessage: TJSONObject;
var
  Line: RawByteString;
  ContentLength: Integer;
  Body: TBytes;
  JsonStr: string;
  JsonValue: TJSONValue;
begin
  ContentLength := -1;

  // Read headers until blank line
  while True do
  begin
    Line := ReadLine;
    if Line = '' then
      Break;
    if Pos(RawByteString('Content-Length:'), Line) = 1 then
    begin
      var ValStr := Trim(string(Copy(Line, 16, MaxInt)));
      ContentLength := StrToIntDef(ValStr, -1);
    end;
    // Ignore Content-Type and other headers
  end;

  if ContentLength < 0 then
    raise EStreamError.Create('Kein Content-Length Header empfangen');

  // Read body
  Body := ReadBytes(ContentLength);
  JsonStr := TEncoding.UTF8.GetString(Body);

  JsonValue := TJSONObject.ParseJSONValue(JsonStr);
  if not (JsonValue is TJSONObject) then
  begin
    JsonValue.Free;
    raise EJSONException.Create('Erwartetes JSON-Objekt, erhalten: ' + JsonStr);
  end;

  Result := TJSONObject(JsonValue);
end;

end.
