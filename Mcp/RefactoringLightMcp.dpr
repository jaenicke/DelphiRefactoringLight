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
//   --no-relaunch  do not hand over to a newer exe (see RunServer)
//   --relaunch-hop=<n>  how often this chain has handed over already; the
//                  successor gets n+1, and MaxRelaunchHops is the backstop.
//                  1.33.0 passed --no-relaunch to every successor instead,
//                  which is permanent - so the SECOND install in one session
//                  left a stale bridge behind (2026-10-28).

{$APPTYPE CONSOLE}

uses
  Winapi.Windows,
  System.SysUtils,
  System.IOUtils,
  System.StrUtils,
  System.Classes,
  System.SyncObjs,
  System.JSON,
  Expert.Version in '..\Source\Expert.Version.pas',
  Mcp.Protocol in '..\Source\Mcp.Protocol.pas',
  Mcp.Bridge in '..\Source\Mcp.Bridge.pas';

// '--list', '-list', '/list' alike
function HasArg(const AName: string): Boolean;
begin
  for var I := 1 to ParamCount do
    if SameText(ParamStr(I).TrimLeft(['-', '/']), AName) then Exit(True);
  Result := False;
end;

// "--name=value" - the value, or '' when the argument is not there.
function ArgValue(const AName: string): string;
begin
  Result := '';
  for var I := 1 to ParamCount do
  begin
    var S := ParamStr(I).TrimLeft(['-', '/']);
    var Eq := Pos('=', S);
    if (Eq > 1) and SameText(Copy(S, 1, Eq - 1), AName) then
      Exit(Copy(S, Eq + 1, MaxInt));
  end;
end;

var
  // stdout is written by the request loop AND the tools watcher - one
  // message must never interleave with another.
  GOutLock: TCriticalSection;
  // How often this chain of bridges has handed its pipes on already - the
  // successor is started with this raised by one (see HandOverTo).
  GHop: Integer;

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

// Aside copies of OUR exe that nobody runs from. Best effort by
// definition: the one our own parent runs from is in use, so Windows
// refuses it and the NEXT bridge removes it.
procedure SweepAsideExes(AErrH: THandle);
var
  Mine, Canon: string;
  Gone: Integer;
begin
  Gone := 0;
  try
    Mine := ParamStr(0);
    Canon := CanonicalBridgeExePath(Mine);
    if Canon = '' then Canon := Mine;
    for var F in AsideBridgeExesToDelete(
      TDirectory.GetFiles(ExtractFilePath(Canon),
        ExtractFileName(Canon) + '.*.old'), Mine) do
      if DeleteFile(PChar(F)) then Inc(Gone);
  except
    // a cleanup must never cost the session its server
  end;
  if Gone > 0 then
    WriteOut(AErrH, Format('removed %d aside cop%s of an earlier install',
      [Gone, IfThen(Gone = 1, 'y', 'ies')]));
end;

// Runs "<exe> --version" and reads its first line. The successor is PROBED
// before it gets the real pipes: if it cannot start at all (a half-written
// file, a missing dependency), handing the conversation over would leave
// the session with no server, while simply carrying on with the old code
// costs nothing but the version skew we already report.
function AnswersItsVersion(const AExe: string; out AVersion: string): Boolean;
var
  Sec: TSecurityAttributes;
  RdH, WrH, NulH: THandle;
  SI: TStartupInfo;
  PI: TProcessInformation;
  Buf: array[0..511] of AnsiChar;
  Got: DWORD;
begin
  Result := False;
  AVersion := '';
  Sec := Default(TSecurityAttributes);
  Sec.nLength := SizeOf(Sec);
  Sec.bInheritHandle := True;
  if not CreatePipe(RdH, WrH, @Sec, 0) then Exit;
  // NUL as the probe's stdin: our own stdin belongs to the client and a
  // single byte read by the wrong process is a lost request.
  NulH := CreateFile('NUL', GENERIC_READ, FILE_SHARE_READ or FILE_SHARE_WRITE,
    @Sec, OPEN_EXISTING, 0, 0);
  try
    SetHandleInformation(RdH, HANDLE_FLAG_INHERIT, 0);
    SI := Default(TStartupInfo);
    SI.cb := SizeOf(SI);
    SI.dwFlags := STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
    SI.wShowWindow := SW_HIDE;
    SI.hStdInput := NulH;
    SI.hStdOutput := WrH;
    SI.hStdError := WrH;
    if not CreateProcess(nil, PChar('"' + AExe + '" --version'), nil, nil,
      True, CREATE_NO_WINDOW, nil, nil, SI, PI) then Exit;
    try
      // The write end must go here, or ReadFile would wait for a writer
      // that is this process itself.
      CloseHandle(WrH);
      WrH := 0;
      Got := 0;
      if ReadFile(RdH, Buf[0], SizeOf(Buf) - 1, Got, nil) and (Got > 0) then
      begin
        Buf[Got] := #0;
        AVersion := Trim(string(AnsiString(Buf)));
      end;
      WaitForSingleObject(PI.hProcess, 10000);
    finally
      CloseHandle(PI.hThread);
      CloseHandle(PI.hProcess);
    end;
    Result := AVersion.StartsWith('RefactoringLightMcp', True);
  finally
    if WrH <> 0 then CloseHandle(WrH);
    CloseHandle(RdH);
    if NulH <> INVALID_HANDLE_VALUE then CloseHandle(NulH);
  end;
end;

// Start AExe on OUR OWN std handles and wait for it. The client created
// those pipes and waits on THIS process, so this one stays alive until the
// successor is done - from the outside nothing happened but a slower start.
function HandOverTo(const AExe: string; AHop: Integer;
  out AError: string): Integer;
var
  SI: TStartupInfo;
  PI: TProcessInformation;
  Code: DWORD;
begin
  Result := -1;
  AError := '';
  SI := Default(TStartupInfo);
  SI.cb := SizeOf(SI);
  SI.dwFlags := STARTF_USESTDHANDLES;
  SI.hStdInput := GetStdHandle(STD_INPUT_HANDLE);
  SI.hStdOutput := GetStdHandle(STD_OUTPUT_HANDLE);
  SI.hStdError := GetStdHandle(STD_ERROR_HANDLE);
  SetHandleInformation(SI.hStdInput, HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT);
  SetHandleInformation(SI.hStdOutput, HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT);
  SetHandleInformation(SI.hStdError, HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT);
  // THE SUCCESSOR MAY HAND OVER AGAIN, and that is the whole correction:
  // the next install renames IT aside, and a permanent --no-relaunch meant
  // it could not react - the second install in one Claude Code session left
  // the stale bridge serving (2026-10-28). The count is the backstop.
  if not CreateProcess(nil,
    PChar(Format('"%s" --relaunch-hop=%d --announce-tools', [AExe, AHop + 1])),
    nil, nil, True, 0, nil, nil, SI, PI) then
  begin
    AError := SysErrorMessage(GetLastError);
    Exit;
  end;
  try
    WaitForSingleObject(PI.hProcess, INFINITE);
    if GetExitCodeProcess(PI.hProcess, Code) then Result := Integer(Code)
    else Result := 0;
  finally
    CloseHandle(PI.hThread);
    CloseHandle(PI.hProcess);
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
  Plan: TBridgeRelaunch;
  Ver, Err: string;
begin
  Plan := Default(TBridgeRelaunch);
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
    SweepAsideExes(ErrH);
    if HasArg('announce-tools') then
      // A hand-over means an install happened, so the IDE may offer other
      // tools than the client fetched from our predecessor. CheckToolsChanged
      // reports nothing until it has served one tools/list, so the ONE
      // notification that makes the client refetch comes from here.
      WriteOut(OutH, '{"jsonrpc":"2.0","method":"notifications/tools/list_changed"}');
    Acc := nil;
    Len := 0;
    // Requests are handled one after the other: Claude Code waits for each
    // tool result anyway, and the IDE side serialises on its main thread.
    while True do
    begin
      // BETWEEN two requests is the only safe moment to hand the pipes
      // over: nothing is half-read (Len = 0), the answer to the last
      // request is out, and the successor simply reads the next line.
      // Asked here rather than in the watcher thread, because only this
      // loop may act on it - so there is no shared state at all.
      if (Len = 0) and not HasArg('no-relaunch') then
      begin
        Plan := PlanBridgeRelaunch(CurrentProcessImagePath,
          function(APath: string): Boolean
          begin
            Result := TFile.Exists(APath);
          end, False, GHop);
        if Plan.Wanted then Break;
      end;
      if not (ReadFile(InH, Buf[0], SizeOf(Buf), Got, nil) and (Got > 0)) then
        Break;
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

  // Everything of ours is down now - the successor gets a clean field.
  if not Plan.Wanted then Exit;
  if not AnswersItsVersion(Plan.NewExe, Ver) then
  begin
    WriteOut(ErrH, Format('%s does not start (%s) - this session keeps the ' +
      'old bridge', [Plan.NewExe, IfThen(Ver = '', 'no answer', Ver)]));
    Exit;
  end;
  WriteOut(ErrH, Format('%s (now: %s)', [Plan.Reason, Ver]));
  ExitCode := HandOverTo(Plan.NewExe, GHop, Err);
  if Err <> '' then
    WriteOut(ErrH, 'could not start it: ' + Err + ' - this session has no ' +
      'bridge any more, start a new one');
end;

begin
  GOutLock := TCriticalSection.Create;
  // Read ONCE, here: every later reader is the request loop, which must not
  // parse the command line between two requests.
  GHop := StrToIntDef(ArgValue('relaunch-hop'), 0);
  if GHop < 0 then GHop := 0;
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
