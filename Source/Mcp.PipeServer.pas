(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Mcp.PipeServer;

// The listening end of the bridge <-> IDE pipe (see Mcp.Protocol). Pure
// Windows + RTL: no ToolsAPI, no VCL - the IDE side plugs in a handler,
// and the console tests run the very same server with a fake one.
//
// LIFECYCLE (the package is a BPL that can be unloaded):
// * ONE listener thread always keeps a pipe instance waiting, so a second
//   request never sees "no such pipe" while the first is being served
//   (which would make discovery miss the IDE);
// * every accepted connection runs on its own short-lived handler thread,
//   counted in FActive;
// * Stop signals FStopEvent - every wait in here and in the handler (see
//   StopEvent) watches it - and waits for the listener and all handlers
//   before returning. No code of this unit runs after Stop.

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes;

type
  /// <summary>One request line in, one response line out. Runs on a
  ///  handler thread; must watch AStop in long waits.</summary>
  TMcpRequestHandler = reference to function(const ARequest: string;
    AStop: THandle): string;

  /// <summary>A process that talked to this pipe, and when it last did.
  ///  </summary>
  TMcpClientInfo = record
    Pid: Cardinal;
    LastTick: UInt64;   // GetTickCount64 of its last request
    Requests: Integer;
  end;

  TMcpPipeServer = class
  private
    FPipeName: string;
    FHandler: TMcpRequestHandler;
    FStopEvent: THandle;
    FReadyEvent: THandle;   // set once a pipe instance is waiting
    FListener: TThread;
    FListening: Integer;
    FActive: Integer;
    FRequests: Integer;
    FLastError: string;
    FLock: TObject;
    FClients: TArray<TMcpClientInfo>;
    procedure ListenLoop;
    procedure NoteClient(APipe: THandle);
    procedure Serve(APipe: THandle);
    procedure SetLastError(const S: string);
  public
    constructor Create(const APipeName: string; const AHandler: TMcpRequestHandler);
    destructor Destroy; override;
    /// <summary>Starts the listener and waits briefly for its first pipe
    ///  instance. False = the pipe is NOT there (LastError says why); the
    ///  listener keeps retrying, so it can still come up later.</summary>
    function Start: Boolean;
    /// <summary>True while a pipe instance is waiting for clients.</summary>
    function Listening: Boolean;
    procedure Stop(ATimeoutMs: Cardinal = 5000);
    function LastError: string;
    /// <summary>The processes (MCP bridges) that sent a request within the
    ///  last AMaxAgeMs, most recent first.</summary>
    function RecentClients(AMaxAgeMs: Cardinal): TArray<TMcpClientInfo>;
    property PipeName: string read FPipeName;
    property StopEvent: THandle read FStopEvent;
    property RequestCount: Integer read FRequests;
  end;

/// <summary>Builds the pipe's security descriptor once and frees it again -
///  the step every pipe instance depends on, and the one that failed with
///  ERROR_NOACCESS (998) on a 64-bit IDE while the token buffer was a byte
///  array. For the test suites and for diagnostics.</summary>
function McpPipeSecurityProbe(out AErr: DWORD): Boolean;

implementation

uses
  System.SyncObjs, System.JSON, Mcp.Protocol;

const
  PIPE_REJECT_REMOTE_CLIENTS = $00000008;
  SDDL_REVISION_1 = 1;

function ConvertSidToStringSidW(Sid: PSID; var StringSid: PWideChar): BOOL;
  stdcall; external advapi32 name 'ConvertSidToStringSidW';
function ConvertStringSecurityDescriptorToSecurityDescriptorW(
  StringSecurityDescriptor: PWideChar; StringSDRevision: DWORD;
  var SecurityDescriptor: PSECURITY_DESCRIPTOR;
  SecurityDescriptorSize: PULONG): BOOL; stdcall;
  external advapi32 name 'ConvertStringSecurityDescriptorToSecurityDescriptorW';

// "D:P(A;;GA;;;<current user>)" - only this user (and SYSTEM) may open the
// pipe. The default DACL would additionally grant READ to Everyone.
function CurrentUserPipeSD(out ASD: PSECURITY_DESCRIPTOR; out AErr: DWORD): Boolean;
var
  Token: THandle;
  // ALIGNMENT MATTERS HERE. GetTokenInformation writes a TOKEN_USER, and
  // that structure STARTS WITH A POINTER. An array of Byte has alignment 1,
  // so its stack slot may land on an odd address - the API then refuses the
  // buffer with ERROR_NOACCESS (998), the security descriptor is never
  // built and the whole MCP pipe stays absent. Reported from a 64-bit IDE
  // (the 32-bit one happened to get an even slot every time). An array of
  // UInt64 is the same 512 bytes, guaranteed 8-byte aligned.
  // RULE: every buffer a Win32 API fills with a STRUCTURE must be aligned -
  // a byte array is only safe for raw bytes.
  Buf: array[0..63] of UInt64;
  Len: DWORD;
  SidStr: PWideChar;
  Sddl: string;
begin
  Result := False;
  ASD := nil;
  AErr := 0;
  if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, Token) then
  begin
    AErr := GetLastError;
    Exit;
  end;
  try
    if not GetTokenInformation(Token, TokenUser, @Buf[0], SizeOf(Buf), Len) then
    begin
      AErr := GetLastError;
      Exit;
    end;
    SidStr := nil;
    if not ConvertSidToStringSidW(PTokenUser(@Buf[0])^.User.Sid, SidStr) then
    begin
      AErr := GetLastError;
      Exit;
    end;
    try
      Sddl := 'D:P(A;;GA;;;' + SidStr + ')(A;;GA;;;SY)';
    finally
      LocalFree(HLOCAL(SidStr));
    end;
    Result := ConvertStringSecurityDescriptorToSecurityDescriptorW(PWideChar(Sddl),
      SDDL_REVISION_1, ASD, nil);
    if not Result then AErr := GetLastError;
  finally
    CloseHandle(Token);   // must not overwrite AErr - hence the captures above
  end;
end;

function McpPipeSecurityProbe(out AErr: DWORD): Boolean;
var
  SD: PSECURITY_DESCRIPTOR;
begin
  SD := nil;
  Result := CurrentUserPipeSD(SD, AErr);
  if Result then LocalFree(HLOCAL(SD));
end;

type
  TListenerThread = class(TThread)
  private
    FOwner: TMcpPipeServer;
  protected
    procedure Execute; override;
  end;

procedure TListenerThread.Execute;
begin
  FOwner.ListenLoop;
end;

{ TMcpPipeServer }

constructor TMcpPipeServer.Create(const APipeName: string;
  const AHandler: TMcpRequestHandler);
begin
  inherited Create;
  FPipeName := APipeName;
  FHandler := AHandler;
  FLock := TObject.Create;
  FStopEvent := CreateEvent(nil, True, False, nil);
  FReadyEvent := CreateEvent(nil, True, False, nil);
end;

destructor TMcpPipeServer.Destroy;
begin
  Stop;
  CloseHandle(FReadyEvent);
  CloseHandle(FStopEvent);
  FLock.Free;
  inherited;
end;

procedure TMcpPipeServer.SetLastError(const S: string);
begin
  TMonitor.Enter(FLock);
  try
    FLastError := S;
  finally
    TMonitor.Exit(FLock);
  end;
end;

function TMcpPipeServer.LastError: string;
begin
  TMonitor.Enter(FLock);
  try
    Result := FLastError;
  finally
    TMonitor.Exit(FLock);
  end;
end;

procedure TMcpPipeServer.NoteClient(APipe: THandle);
var
  Pid: ULONG;
  I: Integer;
begin
  if not GetNamedPipeClientProcessId(APipe, Pid) then Exit;
  TMonitor.Enter(FLock);
  try
    for I := 0 to High(FClients) do
      if FClients[I].Pid = Pid then
      begin
        FClients[I].LastTick := GetTickCount64;
        Inc(FClients[I].Requests);
        Exit;
      end;
    var C: TMcpClientInfo;
    C.Pid := Pid;
    C.LastTick := GetTickCount64;
    C.Requests := 1;
    // bounded: forget entries silent for more than an hour
    var Kept: TArray<TMcpClientInfo> := nil;
    for var X in FClients do
      if GetTickCount64 - X.LastTick < 3600000 then Kept := Kept + [X];
    FClients := Kept + [C];
  finally
    TMonitor.Exit(FLock);
  end;
end;

function TMcpPipeServer.RecentClients(AMaxAgeMs: Cardinal): TArray<TMcpClientInfo>;
var
  Now: UInt64;
begin
  Result := nil;
  Now := GetTickCount64;
  TMonitor.Enter(FLock);
  try
    for var C in FClients do
      if Now - C.LastTick <= AMaxAgeMs then Result := Result + [C];
  finally
    TMonitor.Exit(FLock);
  end;
  // most recent first (a handful of entries - insertion sort is plenty)
  for var I := 1 to High(Result) do
  begin
    var X := Result[I];
    var J := I - 1;
    while (J >= 0) and (Result[J].LastTick < X.LastTick) do
    begin
      Result[J + 1] := Result[J];
      Dec(J);
    end;
    Result[J + 1] := X;
  end;
end;

function TMcpPipeServer.Start: Boolean;
var
  T: TListenerThread;
begin
  Result := FListener <> nil;
  if Result then Exit;
  ResetEvent(FStopEvent);
  ResetEvent(FReadyEvent);
  T := TListenerThread.Create(True);
  T.FOwner := Self;
  T.FreeOnTerminate := False;
  FListener := T;
  T.Start;
  // The first instance is created within milliseconds. Waiting for it is
  // what makes a failure VISIBLE: the old version returned True whatever
  // happened, so a listener that had already given up looked like a
  // running server (that is how a 64-bit IDE ended up with no pipe and no
  // explanation beyond the status row).
  Result := WaitForSingleObject(FReadyEvent, 2000) = WAIT_OBJECT_0;
end;

function TMcpPipeServer.Listening: Boolean;
begin
  Result := TInterlocked.CompareExchange(FListening, 0, 0) <> 0;
end;

procedure TMcpPipeServer.Stop(ATimeoutMs: Cardinal);
var
  Deadline: UInt64;
begin
  if FListener = nil then Exit;
  SetEvent(FStopEvent);
  FListener.WaitFor;
  FreeAndNil(FListener);
  // Handlers watch the stop event in every wait; give them the time to
  // leave. A handler still running after the deadline would execute code
  // of an unloaded BPL, so the deadline is generous.
  Deadline := GetTickCount64 + ATimeoutMs;
  while (TInterlocked.CompareExchange(FActive, 0, 0) > 0) and
        (GetTickCount64 < Deadline) do
    Sleep(10);
end;

procedure TMcpPipeServer.ListenLoop;
var
  SD: PSECURITY_DESCRIPTOR;
  SA: TSecurityAttributes;
  Pipe: THandle;
  Ov: TOverlapped;
  Handles: array[0..1] of THandle;
  First: Boolean;
  Flags: DWORD;
  Got: DWORD;
  Err: DWORD;
  Attempt: Integer;
begin
  SD := nil;
  // A failing security descriptor used to END this thread - the IDE then
  // had no pipe until it was restarted, and Start reported success anyway.
  // It is treated like a failing CreateNamedPipe now: say why and try
  // again, so a transient cause heals itself.
  Attempt := 0;
  while not CurrentUserPipeSD(SD, Err) do
  begin
    Inc(Attempt);
    SetLastError(Format('cannot build the pipe security descriptor (attempt %d): %s (%d)',
      [Attempt, SysErrorMessage(Err), Err]));
    if WaitForSingleObject(FStopEvent, 5000) = WAIT_OBJECT_0 then Exit;
  end;
  FillChar(Ov, SizeOf(Ov), 0);
  Ov.hEvent := CreateEvent(nil, True, False, nil);
  try
    SA.nLength := SizeOf(SA);
    SA.lpSecurityDescriptor := SD;
    SA.bInheritHandle := False;
    First := True;
    while WaitForSingleObject(FStopEvent, 0) <> WAIT_OBJECT_0 do
    begin
      Flags := PIPE_ACCESS_DUPLEX or FILE_FLAG_OVERLAPPED;
      // FIRST_PIPE_INSTANCE: refuse to join a pipe somebody else created
      // under our name (pipe squatting).
      if First then Flags := Flags or FILE_FLAG_FIRST_PIPE_INSTANCE;
      Pipe := CreateNamedPipe(PChar(FPipeName), Flags,
        PIPE_TYPE_BYTE or PIPE_READMODE_BYTE or PIPE_WAIT or PIPE_REJECT_REMOTE_CLIENTS,
        PIPE_UNLIMITED_INSTANCES, 64 * 1024, 64 * 1024, 0, @SA);
      if Pipe = INVALID_HANDLE_VALUE then
      begin
        SetLastError('CreateNamedPipe failed: ' + SysErrorMessage(GetLastError));
        if First then Exit;   // somebody owns our name - give up
        if WaitForSingleObject(FStopEvent, 1000) = WAIT_OBJECT_0 then Exit;
        Continue;
      end;
      First := False;
      // An instance is waiting: the server really is reachable now, so a
      // previous failure message must not stay in the status row.
      if TInterlocked.Exchange(FListening, 1) = 0 then
      begin
        SetLastError('');
        SetEvent(FReadyEvent);
      end;

      ResetEvent(Ov.hEvent);
      var Connected := ConnectNamedPipe(Pipe, @Ov);
      if not Connected then
        case GetLastError of
          ERROR_PIPE_CONNECTED:
            Connected := True;
          ERROR_IO_PENDING:
            begin
              Handles[0] := Ov.hEvent;
              Handles[1] := FStopEvent;
              if WaitForMultipleObjects(2, @Handles[0], False, INFINITE) = WAIT_OBJECT_0 then
                Connected := GetOverlappedResult(Pipe, Ov, Got, False)
              else
              begin
                CancelIo(Pipe);
                GetOverlappedResult(Pipe, Ov, Got, True);
              end;
            end;
        end;
      if not Connected then
      begin
        CloseHandle(Pipe);
        Continue;
      end;

      // Hand the connection to its own thread and immediately offer the
      // next instance.
      TInterlocked.Increment(FActive);
      try
        var P := Pipe;
        TThread.CreateAnonymousThread(
          procedure
          begin
            try
              Serve(P);
            finally
              TInterlocked.Decrement(FActive);
            end;
          end).Start;
      except
        TInterlocked.Decrement(FActive);
        CloseHandle(Pipe);
      end;
    end;
  finally
    TInterlocked.Exchange(FListening, 0);
    CloseHandle(Ov.hEvent);
    LocalFree(HLOCAL(SD));
  end;
end;

procedure TMcpPipeServer.Serve(APipe: THandle);
var
  Req, Resp: string;
begin
  try
    if not PipeReadLine(APipe, FStopEvent, 10000, Req) then Exit;
    TInterlocked.Increment(FRequests);
    NoteClient(APipe);
    try
      Resp := FHandler(Req, FStopEvent);
    except
      on E: Exception do
      begin
        var O := TJSONObject.Create;
        try
          O.AddPair('ok', TJSONBool.Create(False));
          O.AddPair('error', E.ClassName + ': ' + E.Message);
          Resp := O.ToJSON;
        finally
          O.Free;
        end;
      end;
    end;
    PipeWriteLine(APipe, FStopEvent, 10000, Resp);
    FlushFileBuffers(APipe);
  finally
    DisconnectNamedPipe(APipe);
    CloseHandle(APipe);
  end;
end;

end.
