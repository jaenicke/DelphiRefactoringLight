(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
program RefactoringLightMcp;

// MCP server (stdio) for Claude Code & co. Claude Code STARTS this program
// itself - one process per session - after it was registered once:
//
//   claude mcp add --scope user delphi-refactoring-light -- "<path>\RefactoringLightMcp.exe"
//
// It talks MCP on stdin/stdout (one JSON-RPC message per line) and forwards
// the tool calls to the running RAD Studio IDEs through their named pipes
// (Mcp.Protocol). stdout carries NOTHING but protocol messages; diagnostics
// go to stderr.
//
// Command line (for humans, not for Claude Code):
//   --list     print the running IDEs as the bridge sees them, then exit
//   --version  print the version, then exit

{$APPTYPE CONSOLE}

uses
  Winapi.Windows,
  System.SysUtils,
  System.Classes,
  System.SyncObjs,
  System.JSON,
  Mcp.Protocol in '..\Source\Mcp.Protocol.pas',
  Mcp.Bridge in '..\Source\Mcp.Bridge.pas';

// '--list', '-list', '/list' alike
function HasArg(const AName: string): Boolean;
begin
  for var I := 1 to ParamCount do
    if SameText(ParamStr(I).TrimLeft(['-', '/']), AName) then Exit(True);
  Result := False;
end;

var
  // stdout is written by the request loop AND the tools watcher - one
  // message must never interleave with another.
  GOutLock: TCriticalSection;

procedure WriteOut(AHandle: THandle; const S: string);
var
  B: TBytes;
  Done, W: DWORD;
begin
  B := TEncoding.UTF8.GetBytes(S + #10);
  GOutLock.Enter;
  try
    Done := 0;
    while Done < DWORD(Length(B)) do
    begin
      if not WriteFile(AHandle, B[Done], DWORD(Length(B)) - Done, W, nil) then Exit;
      if W = 0 then Exit;
      Inc(Done, W);
    end;
  finally
    GOutLock.Leave;
  end;
end;

procedure RunList;
var
  Bridge: TMcpBridge;
  Resp: string;
  V: TJSONValue;
begin
  Bridge := TMcpBridge.Create(GetCurrentDir, TPipeTransport.Create);
  try
    Resp := Bridge.HandleLine('{"jsonrpc":"2.0","id":1,"method":"tools/call",' +
      '"params":{"name":"ide_instances","arguments":{}}}');
    V := TJSONObject.ParseJSONValue(Resp);
    try
      Writeln(TJSONObject(V).GetValue<string>('result.content[0].text', Resp));
    finally
      V.Free;
    end;
  finally
    Bridge.Free;
  end;
end;

procedure RunServer;
var
  InH, OutH, ErrH: THandle;
  Buf: array[0..65535] of Byte;
  Got: DWORD;
  Acc: TBytes;
  Len, I, Start: Integer;
  Bridge: TMcpBridge;
  Line, Resp: string;
begin
  InH := GetStdHandle(STD_INPUT_HANDLE);
  OutH := GetStdHandle(STD_OUTPUT_HANDLE);
  ErrH := GetStdHandle(STD_ERROR_HANDLE);
  Bridge := TMcpBridge.Create(GetCurrentDir, TPipeTransport.Create);
  var StopWatch := TEvent.Create(nil, True, False, '');
  // Tools watcher: an IDE started, closed or updated to a plugin with other
  // tools -> tell the client to fetch the list again.
  var Watcher := TThread.CreateAnonymousThread(
    procedure
    begin
      while StopWatch.WaitFor(5000) = wrTimeout do
        try
          if Bridge.CheckToolsChanged then
            WriteOut(OutH, '{"jsonrpc":"2.0","method":"notifications/tools/list_changed"}');
        except
          // a watcher round must never take the bridge down
        end;
    end);
  Watcher.FreeOnTerminate := False;
  Watcher.Start;
  try
    WriteOut(ErrH, 'RefactoringLightMcp ' + BridgeVersion + ' ready, working directory ' +
      Bridge.Cwd);
    Acc := nil;
    Len := 0;
    // Requests are handled one after the other: Claude Code waits for each
    // tool result anyway, and the IDE side serialises on its main thread.
    while ReadFile(InH, Buf[0], SizeOf(Buf), Got, nil) and (Got > 0) do
    begin
      SetLength(Acc, Len + Integer(Got));
      Move(Buf[0], Acc[Len], Got);
      Inc(Len, Got);
      Start := 0;
      for I := 0 to Len - 1 do
        if Acc[I] = 10 then
        begin
          Line := TEncoding.UTF8.GetString(Acc, Start, I - Start);
          Start := I + 1;
          try
            Resp := Bridge.HandleLine(Line.TrimRight([#13]));
          except
            on E: Exception do
            begin
              WriteOut(ErrH, 'error: ' + E.Message);
              Resp := '';
            end;
          end;
          if Resp <> '' then WriteOut(OutH, Resp);
        end;
      if Start > 0 then
      begin
        Acc := Copy(Acc, Start, Len - Start);
        Dec(Len, Start);
      end;
    end;
  finally
    StopWatch.SetEvent;
    Watcher.WaitFor;
    Watcher.Free;
    StopWatch.Free;
    Bridge.Free;
  end;
end;

begin
  GOutLock := TCriticalSection.Create;
  try
    if HasArg('version') then
      Writeln('RefactoringLightMcp ', BridgeVersion)
    else if HasArg('list') then
      RunList
    else
      RunServer;
  except
    on E: Exception do
    begin
      WriteOut(GetStdHandle(STD_ERROR_HANDLE), E.ClassName + ': ' + E.Message);
      ExitCode := 1;
    end;
  end;
end.
