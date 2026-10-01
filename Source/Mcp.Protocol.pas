(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Mcp.Protocol;

// Shared between the IDE package, the MCP bridge (RefactoringLightMcp.exe)
// and the console tests - so it must stay free of ToolsAPI and the VCL.
//
// THE PICTURE:
//
//   Claude Code --stdio (MCP)--> RefactoringLightMcp.exe --named pipe--> IDE 1
//                                   (the bridge)          --named pipe--> IDE 2
//
// Every IDE running the package listens on its OWN pipe,
// \\.\pipe\DelphiRefactoringLight-mcp-<PID>. No ports, no conflicts between
// instances, and the pipe vanishes with the process - so enumerating
// \\.\pipe\ IS the discovery; there is no registry file that could go stale
// after a crash. The pipe's DACL admits the current user only.
//
// Between bridge and IDE: one connection per request, one line of UTF-8
// JSON each way (JSON escapes line breaks inside strings, so a raw #10
// always ends a message).
//   request : {"bridge":"1.15.2","method":"context"}
//             {"bridge":"...","method":"call","tool":"get_quick_fixes",...}
//   response: {"ok":true,"result":{...}}  /  {"ok":false,"error":"..."}
//
// "bridge" is the version of the bridge EXE that sent the request. Before
// 1.15.2 the bridge reported its version only to its MCP client, never here -
// so an install that silently left an OLD exe behind (install.cmd could not
// replace it) was undetectable from inside the IDE, and its tools behaved like
// the build they came from: nine days of wrong timeouts, diagnosed twice from
// the wrong end. A request without the field means "a bridge older than that".

interface

uses
  Winapi.Windows, System.SysUtils, System.JSON;

const
  McpPipePrefix = 'DelphiRefactoringLight-mcp-';
  /// <summary>Bumped when the bridge <-> IDE request format changes
  ///  incompatibly; both sides report it in the context.</summary>
  McpWireVersion = 1;

  McpServerName = 'delphi-refactoring-light';

function McpPipeName(APid: Cardinal): string;

/// <summary>PIDs of every process currently serving a Refactoring Light
///  pipe - enumerated from \\.\pipe\, so a crashed IDE simply is not there.
///  </summary>
function ListMcpPipePids: TArray<Cardinal>;

// ---- line framing over OVERLAPPED handles (both ends open them so) ----

/// <summary>Reads up to the next #10. False on timeout, on AStop being
///  signalled, on a broken pipe, or above AMaxBytes.</summary>
function PipeReadLine(AHandle: THandle; AStop: THandle; ATimeoutMs: Cardinal;
  out ALine: string; AMaxBytes: Integer = 32 * 1024 * 1024): Boolean;
function PipeWriteLine(AHandle: THandle; AStop: THandle; ATimeoutMs: Cardinal;
  const ALine: string): Boolean;

/// <summary>Client side: connect, send one request line, read one response
///  line. AError says why it failed (not running / busy / timeout).</summary>
function McpPipeRequest(APid: Cardinal; const ARequest: string;
  ATimeoutMs: Cardinal; out AResponse, AError: string): Boolean;

// ---- which bridge exe is talking to us (see the note at the top) ----

/// <summary>Writes AVersion into the request line as "bridge". Inserted as
///  TEXT right behind the '{' on purpose: a request can carry a whole unit
///  (buffer_edit), and parsing megabytes just to add one field would cost more
///  than the call itself. An empty AVersion or a non-object line is returned
///  unchanged.</summary>
function StampBridgeVersion(const ARequest, AVersion: string): string;

/// <summary>The version the sender stamped in, '' when the line carries none
///  (a bridge from before 1.15.2). Reads only the HEAD of the line, for the
///  same reason - the server must not parse a huge request twice.</summary>
function BridgeVersionOfRequest(const ARequest: string): string;

/// <summary>'' when the two versions agree (nothing to report); otherwise one
///  line naming BOTH numbers and the two things that cause a skew. Pure, so
///  the status row, get_status and the tests share one wording. Deliberately
///  does not compare which is newer: the answer is the same either way - the
///  install did not finish.</summary>
function BridgeVersionProblem(const ABridge, APlugin: string): string;

// ---- what an IDE instance tells the bridge about itself ----

type
  TMcpInstanceContext = record
    Pid: Cardinal;
    Bitness: Integer;          // 32 / 64
    IdeVersion: string;        // BDS version, e.g. '37.0'
    PluginVersion: string;     // e.g. '1.0'
    PluginBuild: string;       // build time of the loaded package
    WireVersion: Integer;
    ProjectGroup: string;
    ActiveProject: string;
    Projects: TArray<string>;
    ActiveFile: string;
    OpenFiles: TArray<string>;
    Busy: Boolean;             // a modal dialog is open in that IDE
    /// <summary>McpToolsHash of the tool list this IDE serves ('' = a
    ///  plugin from before the IDE announced its own tools).</summary>
    ToolsHash: string;
    function ToJson: TJSONObject;
    class function FromJson(AObj: TJSONObject): TMcpInstanceContext; static;
    function Describe: string;
  end;

/// <summary>True when APath is ADir itself or lies below it.</summary>
function IsPathWithin(const APath, ADir: string): Boolean;

/// <summary>How well an IDE instance matches a request. AFile (the file the
///  tool is about) weighs most - an IDE that has it OPEN is almost certainly
///  meant; ADir is the bridge's working directory, i.e. the folder the
///  Claude session runs in, compared against the IDE's projects.
///  0 = no relation at all.</summary>
function ScoreInstance(const ACtx: TMcpInstanceContext;
  const ADir, AFile: string): Integer;

/// <summary>Picks the instance for a request. A single running IDE is taken
///  as it is; otherwise the best score must be positive AND unique - an
///  ambiguous choice is refused (AReason lists why) rather than guessed.
///  </summary>
function PickInstance(const ACtxs: TArray<TMcpInstanceContext>;
  const ADir, AFile: string; out AIndex: Integer; out AReason: string): Boolean;

// ---- the MCP surface ----

/// <summary>The tool list the bridge announces (JSON schema per tool). The
///  IDE side validates names against the same list.</summary>
function McpToolDefinitions: TJSONArray;
/// <summary>Fingerprint of McpToolDefinitions - an IDE reports it in its
///  context, so the bridge sees a changed tool set without fetching it.
///  </summary>
function McpToolsHash: string;
function McpServerInstructions: string;

/// <summary>The argument names a tool declares in its schema, plus
///  'instance' - the bridge may add that one to any call, and most tools
///  do not list it.</summary>
function KnownToolArguments(const ATool: string): TArray<string>;
/// <summary>The incoming names the tool does NOT declare. A handler reads
///  its arguments BY NAME, so an unknown one is dropped without a word:
///  "from_line" instead of "start_line" makes buffer_read answer the whole
///  file and look like a tool that ignored half the request. Empty for a
///  tool this list does not know (the bridge's own two), because then
///  there is nothing to measure against.</summary>
function UnknownToolArguments(const ATool: string;
  const ANames: TArray<string>): TArray<string>;
/// <summary>The sentence the server puts into such an answer, '' when
///  every argument was understood.</summary>
function UnknownArgumentNote(const ATool: string;
  const ANames: TArray<string>): string;

implementation

uses
  System.Classes, System.StrUtils, System.IOUtils,
  System.Generics.Collections;

function McpPipeName(APid: Cardinal): string;
begin
  Result := '\\.\pipe\' + McpPipePrefix + UIntToStr(APid);
end;

function ListMcpPipePids: TArray<Cardinal>;
var
  FD: TWin32FindData;
  H: THandle;
  N: string;
  V: Cardinal;
begin
  Result := nil;
  H := FindFirstFile('\\.\pipe\*', FD);
  if H = INVALID_HANDLE_VALUE then Exit;
  try
    repeat
      N := FD.cFileName;
      if StartsText(McpPipePrefix, N) and
         TryStrToUInt(Copy(N, Length(McpPipePrefix) + 1, MaxInt), V) then
      begin
        var Dup := False;
        for var X in Result do
          if X = V then Dup := True;
        if not Dup then Result := Result + [V];
      end;
    until not FindNextFile(H, FD);
  finally
    Winapi.Windows.FindClose(H);
  end;
end;

// Waits for one overlapped operation. On timeout / stop the operation is
// CANCELLED and waited for, so the kernel never writes into a buffer the
// caller has already given up.
function WaitIo(AHandle: THandle; var AOv: TOverlapped; AStop: THandle;
  ATimeoutMs: Cardinal; out ABytes: DWORD): Boolean;
var
  Handles: array[0..1] of THandle;
  N, R: DWORD;
begin
  ABytes := 0;
  Handles[0] := AOv.hEvent;
  N := 1;
  if AStop <> 0 then
  begin
    Handles[1] := AStop;
    N := 2;
  end;
  R := WaitForMultipleObjects(N, @Handles[0], False, ATimeoutMs);
  if R <> WAIT_OBJECT_0 then
  begin
    CancelIo(AHandle);
    GetOverlappedResult(AHandle, AOv, ABytes, True);
    Exit(False);
  end;
  Result := GetOverlappedResult(AHandle, AOv, ABytes, False);
end;

function PipeReadLine(AHandle: THandle; AStop: THandle; ATimeoutMs: Cardinal;
  out ALine: string; AMaxBytes: Integer): Boolean;
var
  Buf: array[0..8191] of Byte;
  Acc: TBytes;
  Ov: TOverlapped;
  Got: DWORD;
  Deadline: UInt64;
  Remaining: Int64;
  I, Len: Integer;
begin
  Result := False;
  ALine := '';
  Acc := nil;
  Len := 0;
  Deadline := GetTickCount64 + ATimeoutMs;
  FillChar(Ov, SizeOf(Ov), 0);
  Ov.hEvent := CreateEvent(nil, True, False, nil);
  if Ov.hEvent = 0 then Exit;
  try
    while True do
    begin
      Remaining := Int64(Deadline) - Int64(GetTickCount64);
      if Remaining <= 0 then Exit;
      ResetEvent(Ov.hEvent);
      Got := 0;
      if not ReadFile(AHandle, Buf[0], SizeOf(Buf), Got, @Ov) then
      begin
        if GetLastError <> ERROR_IO_PENDING then Exit;   // broken pipe & co
        if not WaitIo(AHandle, Ov, AStop, Remaining, Got) then Exit;
      end
      else if not GetOverlappedResult(AHandle, Ov, Got, False) then
        Exit;
      if Got = 0 then Continue;
      for I := 0 to Integer(Got) - 1 do
        if Buf[I] = 10 then
        begin
          SetLength(Acc, Len + I);
          if I > 0 then Move(Buf[0], Acc[Len], I);
          ALine := TEncoding.UTF8.GetString(Acc);
          if ALine.EndsWith(#13) then SetLength(ALine, Length(ALine) - 1);
          Exit(True);
        end;
      if Len + Integer(Got) > AMaxBytes then Exit;
      SetLength(Acc, Len + Integer(Got));
      Move(Buf[0], Acc[Len], Got);
      Inc(Len, Got);
    end;
  finally
    CloseHandle(Ov.hEvent);
  end;
end;

function PipeWriteLine(AHandle: THandle; AStop: THandle; ATimeoutMs: Cardinal;
  const ALine: string): Boolean;
var
  Data: TBytes;
  Ov: TOverlapped;
  Done, Got: DWORD;
  Deadline: UInt64;
  Remaining: Int64;
begin
  Result := False;
  Data := TEncoding.UTF8.GetBytes(ALine + #10);
  Deadline := GetTickCount64 + ATimeoutMs;
  FillChar(Ov, SizeOf(Ov), 0);
  Ov.hEvent := CreateEvent(nil, True, False, nil);
  if Ov.hEvent = 0 then Exit;
  try
    Done := 0;
    while Done < DWORD(Length(Data)) do
    begin
      Remaining := Int64(Deadline) - Int64(GetTickCount64);
      if Remaining <= 0 then Exit;
      ResetEvent(Ov.hEvent);
      Got := 0;
      if not WriteFile(AHandle, Data[Done], DWORD(Length(Data)) - Done, Got, @Ov) then
      begin
        if GetLastError <> ERROR_IO_PENDING then Exit;
        if not WaitIo(AHandle, Ov, AStop, Remaining, Got) then Exit;
      end
      else if not GetOverlappedResult(AHandle, Ov, Got, False) then
        Exit;
      Inc(Done, Got);
    end;
    Result := True;
  finally
    CloseHandle(Ov.hEvent);
  end;
end;

function McpPipeRequest(APid: Cardinal; const ARequest: string;
  ATimeoutMs: Cardinal; out AResponse, AError: string): Boolean;
const
  SECURITY_SQOS_PRESENT = $00100000;
  SECURITY_IDENTIFICATION = $00010000;
var
  Name: string;
  H: THandle;
  Tries: Integer;
begin
  Result := False;
  AResponse := '';
  AError := '';
  Name := McpPipeName(APid);
  for Tries := 1 to 5 do
  begin
    // SECURITY_IDENTIFICATION: the server may learn who we are, but can
    // never act as us.
    H := CreateFile(PChar(Name), GENERIC_READ or GENERIC_WRITE, 0, nil,
      OPEN_EXISTING, FILE_FLAG_OVERLAPPED or SECURITY_SQOS_PRESENT or
      SECURITY_IDENTIFICATION, 0);
    if H <> INVALID_HANDLE_VALUE then Break;
    case GetLastError of
      ERROR_PIPE_BUSY:
        WaitNamedPipe(PChar(Name), 2000);
      ERROR_FILE_NOT_FOUND:
        begin
          AError := Format('no IDE with pid %d is serving Refactoring Light', [APid]);
          Exit;
        end;
    else
      AError := SysErrorMessage(GetLastError);
      Exit;
    end;
  end;
  if H = INVALID_HANDLE_VALUE then
  begin
    AError := 'the IDE is busy (pipe not available)';
    Exit;
  end;
  try
    if not PipeWriteLine(H, 0, 5000, ARequest) then
    begin
      AError := 'could not send the request to the IDE';
      Exit;
    end;
    if not PipeReadLine(H, 0, ATimeoutMs, AResponse) then
    begin
      AError := Format('no answer from the IDE within %d s (is it hanging ' +
        'or showing a dialog?)', [ATimeoutMs div 1000]);
      Exit;
    end;
    Result := True;
  finally
    CloseHandle(H);
  end;
end;

// ---------------------------------------------------------------------------
//  Instance context
// ---------------------------------------------------------------------------

function StrArrayToJson(const A: TArray<string>): TJSONArray;
begin
  Result := TJSONArray.Create;
  for var S in A do Result.Add(S);
end;

function JsonToStrArray(AVal: TJSONValue): TArray<string>;
begin
  Result := nil;
  if AVal is TJSONArray then
    for var V in TJSONArray(AVal) do
      Result := Result + [V.Value];
end;

function TMcpInstanceContext.ToJson: TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('pid', TJSONNumber.Create(Pid));
  Result.AddPair('bitness', TJSONNumber.Create(Bitness));
  Result.AddPair('ideVersion', IdeVersion);
  Result.AddPair('pluginVersion', PluginVersion);
  Result.AddPair('pluginBuild', PluginBuild);
  Result.AddPair('wireVersion', TJSONNumber.Create(WireVersion));
  Result.AddPair('projectGroup', ProjectGroup);
  Result.AddPair('activeProject', ActiveProject);
  Result.AddPair('projects', StrArrayToJson(Projects));
  Result.AddPair('activeFile', ActiveFile);
  Result.AddPair('openFiles', StrArrayToJson(OpenFiles));
  Result.AddPair('busy', TJSONBool.Create(Busy));
  Result.AddPair('toolsHash', ToolsHash);
end;

class function TMcpInstanceContext.FromJson(AObj: TJSONObject): TMcpInstanceContext;
begin
  Result := Default(TMcpInstanceContext);
  if AObj = nil then Exit;
  Result.Pid := AObj.GetValue<Cardinal>('pid', 0);
  Result.Bitness := AObj.GetValue<Integer>('bitness', 0);
  Result.IdeVersion := AObj.GetValue<string>('ideVersion', '');
  Result.PluginVersion := AObj.GetValue<string>('pluginVersion', '');
  Result.PluginBuild := AObj.GetValue<string>('pluginBuild', '');
  Result.WireVersion := AObj.GetValue<Integer>('wireVersion', 0);
  Result.ProjectGroup := AObj.GetValue<string>('projectGroup', '');
  Result.ActiveProject := AObj.GetValue<string>('activeProject', '');
  Result.Projects := JsonToStrArray(AObj.GetValue('projects'));
  Result.ActiveFile := AObj.GetValue<string>('activeFile', '');
  Result.OpenFiles := JsonToStrArray(AObj.GetValue('openFiles'));
  Result.Busy := AObj.GetValue<Boolean>('busy', False);
  Result.ToolsHash := AObj.GetValue<string>('toolsHash', '');
end;

function TMcpInstanceContext.Describe: string;
begin
  Result := Format('pid %d (%d-bit)', [Pid, Bitness]);
  if ActiveProject <> '' then
    Result := Result + ', active project ' + ExtractFileName(ActiveProject)
  else if ProjectGroup <> '' then
    Result := Result + ', group ' + ExtractFileName(ProjectGroup)
  else
    Result := Result + ', no project open';
end;

function StampBridgeVersion(const ARequest, AVersion: string): string;
var
  Field, Rest: string;
begin
  Result := ARequest;
  if (AVersion = '') or not ARequest.StartsWith('{') then Exit;
  // let the JSON writer quote it - the value is our own version constant, but
  // a field built by hand is exactly how invalid JSON gets on the wire
  var S := TJSONString.Create(AVersion);
  try
    Field := '"bridge":' + S.ToJSON;
  finally
    S.Free;
  end;
  Rest := ARequest.Substring(1);
  if Trim(Rest).StartsWith('}') then
    Result := '{' + Field + Rest          // '{}' has no member to separate from
  else
    Result := '{' + Field + ',' + Rest;
end;

function BridgeVersionOfRequest(const ARequest: string): string;
const
  Key = '"bridge"';
  HeadChars = 160;   // StampBridgeVersion puts it directly behind the '{'
var
  Head: string;
  P: Integer;
begin
  Result := '';
  Head := Copy(ARequest, 1, HeadChars);
  P := Pos(Key, Head);
  if P = 0 then Exit;
  P := P + Length(Key);
  while (P <= Length(Head)) and CharInSet(Head[P], [' ', #9]) do Inc(P);
  if (P > Length(Head)) or (Head[P] <> ':') then Exit;
  Inc(P);
  while (P <= Length(Head)) and CharInSet(Head[P], [' ', #9]) do Inc(P);
  if (P > Length(Head)) or (Head[P] <> '"') then Exit;
  Inc(P);
  while (P <= Length(Head)) and (Head[P] <> '"') do
  begin
    // a version needs no escapes; stop rather than misread something odd
    if Head[P] = '\' then Exit('');
    Result := Result + Head[P];
    Inc(P);
  end;
  if (P > Length(Head)) or (Head[P] <> '"') then Result := '';   // truncated
end;

function BridgeVersionProblem(const ABridge, APlugin: string): string;
begin
  Result := '';
  if (APlugin = '') or (ABridge = APlugin) then Exit;
  if ABridge = '' then
    Result := Format('the bridge exe does not report its version, so it is ' +
      'older than %s - if its tools behave oddly (timeouts, missing tools), ' +
      'run install.cmd and start a new Claude Code session', [APlugin])
  else
    Result := Format('bridge exe %s, plugin %s - these must match. Either ' +
      'install.cmd could not replace RefactoringLightMcp.exe (a running ' +
      'Claude Code session holds it: close them all and install again) or ' +
      'RAD Studio was not restarted after the install', [ABridge, APlugin]);
end;

function NormDir(const S: string): string;
begin
  Result := '';
  if S = '' then Exit;
  try
    Result := IncludeTrailingPathDelimiter(ExpandFileName(S));
  except
    Result := IncludeTrailingPathDelimiter(S);
  end;
end;

function IsPathWithin(const APath, ADir: string): Boolean;
var
  P, D: string;
begin
  if (APath = '') or (ADir = '') then Exit(False);
  P := NormDir(APath);
  D := NormDir(ADir);
  Result := StartsText(D, P);
end;

function ScoreInstance(const ACtx: TMcpInstanceContext;
  const ADir, AFile: string): Integer;
var
  FileScore, DirScore, S: Integer;
  PDir: string;
begin
  FileScore := 0;
  if AFile <> '' then
  begin
    if SameText(ExpandFileName(AFile), ExpandFileName(ACtx.ActiveFile)) then
      FileScore := 100
    else
      for var F in ACtx.OpenFiles do
        if SameText(ExpandFileName(AFile), ExpandFileName(F)) then
        begin
          FileScore := 90;
          Break;
        end;
    if FileScore = 0 then
      for var P in ACtx.Projects do
        if IsPathWithin(AFile, ExtractFilePath(P)) then
        begin
          FileScore := 60;
          Break;
        end;
  end;

  DirScore := 0;
  if ADir <> '' then
  begin
    for var P in ACtx.Projects do
    begin
      PDir := ExtractFilePath(P);
      S := 0;
      // the session runs inside a project's tree, or the project inside
      // the session's folder (a repository root with several projects)
      if IsPathWithin(ADir, PDir) or IsPathWithin(PDir, ADir) then
        if SameText(P, ACtx.ActiveProject) then S := 35 else S := 30;
      if S > DirScore then DirScore := S;
    end;
    if (DirScore = 0) and (ACtx.ProjectGroup <> '') and
       (IsPathWithin(ADir, ExtractFilePath(ACtx.ProjectGroup)) or
        IsPathWithin(ExtractFilePath(ACtx.ProjectGroup), ADir)) then
      DirScore := 25;
    if DirScore = 0 then
      for var F in ACtx.OpenFiles do
        if IsPathWithin(F, ADir) then
        begin
          DirScore := 10;
          Break;
        end;
  end;
  Result := FileScore + DirScore;
end;

function PickInstance(const ACtxs: TArray<TMcpInstanceContext>;
  const ADir, AFile: string; out AIndex: Integer; out AReason: string): Boolean;
var
  Best, BestCount, S, I: Integer;
  Scores: TArray<Integer>;
begin
  Result := False;
  AIndex := -1;
  AReason := '';
  if Length(ACtxs) = 0 then
  begin
    AReason := 'no RAD Studio IDE with Refactoring Light is running';
    Exit;
  end;
  SetLength(Scores, Length(ACtxs));
  Best := 0;
  BestCount := 0;
  for I := 0 to High(ACtxs) do
  begin
    S := ScoreInstance(ACtxs[I], ADir, AFile);
    Scores[I] := S;
    if S > Best then
    begin
      Best := S;
      BestCount := 1;
      AIndex := I;
    end
    else if (S = Best) and (S > 0) then
      Inc(BestCount);
  end;

  if (Best > 0) and (BestCount = 1) then
  begin
    if AFile <> '' then
      AReason := 'matched by the file ' + ExtractFileName(AFile)
    else
      AReason := 'matched by the working directory ' + ADir;
    Exit(True);
  end;
  if Length(ACtxs) = 1 then
  begin
    AIndex := 0;
    AReason := 'the only running IDE';
    Exit(True);
  end;

  AIndex := -1;
  if Best = 0 then
    AReason := 'none of the running IDEs has this file or folder open'
  else
    AReason := 'several running IDEs match equally well';
  AReason := AReason + ' - pass "instance" (the pid) or call select_ide:';
  for I := 0 to High(ACtxs) do
    AReason := AReason + sLineBreak + '  ' + ACtxs[I].Describe;
end;

// ---------------------------------------------------------------------------
//  Tool definitions
// ---------------------------------------------------------------------------

const
  InstanceProp =
    '"instance":{"type":"integer","description":"Process id of the IDE to ' +
    'use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}';
  FileProp =
    '"file":{"type":"string","description":"Absolute path of the unit. ' +
    'Default: the file active in the IDE editor."}';
  RefreshProp =
    '"refresh":{"type":"boolean","description":"Force a fresh analysis by the ' +
    'plugin''s own DelphiLSP session (slower, a few seconds). Done ' +
    'automatically when no current diagnostics exist for the buffer."}';
  ApplyProp =
    '"apply":{"type":"boolean","description":"Default false: nothing is ' +
    'written and the answer lists the changes (file, line, before, after) ' +
    'plus a token. true makes the change."}';
  TokenProp =
    '"token":{"type":"string","description":"The token of the preview this ' +
    'call applies. The change is refused when a file changed since then."}';

  ToolsJson =
    '[' +
    '{"name":"ide_instances",' +
    '"description":"Lists the running RAD Studio IDEs that have Refactoring ' +
    'Light loaded: process id, bitness, project group, projects, active file ' +
    'and open files - and which IDE this session uses and why.",' +
    '"inputSchema":{"type":"object","properties":{}}},' +

    '{"name":"select_ide",' +
    '"description":"Pins this session to one IDE (by process id). 0 returns ' +
    'to automatic selection.",' +
    '"inputSchema":{"type":"object","properties":{"instance":{"type":"integer",' +
    '"description":"Process id from ide_instances, 0 = automatic."}},' +
    '"required":["instance"]}},' +

    '{"name":"get_status",' +
    '"description":"The Refactoring Light status window as data: identifier ' +
    'index, the plugin''s DelphiLSP session, the live quick-fix checker and ' +
    'which diagnostics source answered (with the reason when a fix was ' +
    'declined), menus, process resources and memory, live blame, MCP ' +
    'endpoint. Use it to find out WHY something does not work.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"filter":{"type":"string","description":"Only rows whose item or section ' +
    'contains this text (case-insensitive), e.g. \"LSP\" or \"index\"."},' +
    InstanceProp + '}}},' +

    '{"name":"get_diagnostics",' +
    '"description":"Errors, warnings and hints for a Delphi unit as the IDE ' +
    'sees them in the CURRENT EDITOR BUFFER (unsaved changes included): ' +
    'Error Insight / Structure view, the last compile and the plugin''s own ' +
    'DelphiLSP session, merged and de-duplicated. Lines and columns are ' +
    '1-based.",' +
    '"inputSchema":{"type":"object","properties":{' + FileProp + ',' +
    RefreshProp + ',' + InstanceProp + '}}},' +

    '{"name":"get_quick_fixes",' +
    '"description":"The automatic fixes Refactoring Light offers for the ' +
    'unit''s current diagnostics (add a unit to the uses clause, declare a ' +
    'variable, fix a misspelled identifier, insert a missing semicolon, ' +
    'remove unused variables, ...). Each fix carries the diagnostic behind ' +
    'it (code, message, position), the affected line verbatim and - where ' +
    'the fix is a text edit - the changes it would make, so it can be judged ' +
    'before it is applied. Each fix has an id for apply_quick_fix; ids are ' +
    'bound to the buffer revision and expire with the next edit.",' +
    '"inputSchema":{"type":"object","properties":{' + FileProp + ',' +
    '"line":{"type":"integer","description":"Only fixes anchored to this ' +
    '1-based line."},' + RefreshProp + ',' + InstanceProp + '}}},' +

    '{"name":"apply_quick_fixes","description":"Applies SEVERAL quick fixes of one unit in one ' +
    'go: every fix of a kind (\"remove_var\", \"insert_semi\", ... - the kind names get_quick_f' +
    'ixes reports) or a list of fix ids. Applied bottom-up with the uses-clause fixes last; a f' +
    'ix whose line moved away while the batch ran is skipped, never applied elsewhere. Refused ' +
    'per fix (listed in not_applied): add_unit with several candidate units or a unit only on t' +
    'he browsing path, remove_private (may need a confirmation in the IDE). Call get_quick_fixe' +
    's for the file first; all its fix ids expire afterwards.","inputSchema":{"type":"object","' +
    'properties":{"file":{"type":"string","description":"Absolute path of the unit."},"kind":{"' +
    'type":"string","description":"Apply every fix of this kind, e.g. \"remove_var\"."},"fix_id' +
    's":{"type":"array","items":{"type":"string"},"description":"Apply these fixes (ids from ge' +
    't_quick_fixes)."},"instance":{"type":"integer","description":"Process id of the IDE to use' +
    '. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["' +
    'file"]}}' +
    ',' +
    '{"name":"apply_quick_fix",' +
    '"description":"Applies one fix from get_quick_fixes to the IDE editor ' +
    'buffer (the file is NOT saved). Fails when the buffer changed since ' +
    'the fix was listed - list again then.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path of the unit (as ' +
    'returned by get_quick_fixes)."},' +
    '"fix_id":{"type":"string","description":"Id from get_quick_fixes."},' +
    '"unit":{"type":"string","description":"For add-unit fixes: which of the ' +
    'candidate units to add. Default: the first candidate."},' +
    '"section":{"type":"string","enum":["interface","implementation"],' +
    '"description":"For add-unit fixes: target uses clause. Default: the ' +
    'section the fix proposes."},' + InstanceProp + '},' +
    '"required":["file","fix_id"]}},' +

    '{"name":"buffer_open",' +
    '"description":"Loads a file into the IDE: visible=true opens an editor tab ' +
    'like the user would; otherwise it is loaded HEADLESS - only in the IDE''s ' +
    'memory, no tab. Edits to a headless buffer (buffer_edit, apply_quick_fix) ' +
    'stay in memory until buffer_close saves or discards them, so changes can ' +
    'be tried and checked with get_diagnostics without touching the disk.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path."},' +
    '"visible":{"type":"boolean","description":"Open an editor tab instead of ' +
    'loading headless."},' + InstanceProp + '},"required":["file"]}},' +

    '{"name":"buffer_read",' +
    '"description":"Reads a file as the IDE sees it: the editor buffer when it ' +
    'is loaded (unsaved changes included), otherwise the disk. Returns the ' +
    'revision to pass to buffer_edit.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path."},' +
    '"start_line":{"type":"integer","description":"1-based, default 1."},' +
    '"end_line":{"type":"integer","description":"1-based, inclusive, default ' +
    'the last line."},' + InstanceProp + '},"required":["file"]}},' +

    '{"name":"buffer_edit",' +
    '"description":"Replaces one exact, unique piece of text in a LOADED IDE ' +
    'buffer (open in the editor or loaded by buffer_open). Only the changed ' +
    'lines are touched; nothing is saved.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path."},' +
    '"old_text":{"type":"string","description":"Exact text to replace; must ' +
    'occur exactly once."},' +
    '"new_text":{"type":"string","description":"Replacement text."},' +
    '"revision":{"type":"string","description":"Optional: revision from ' +
    'buffer_read; the edit is refused when the buffer changed since."},' +
    InstanceProp + '},"required":["file","old_text","new_text"]}},' +

    '{"name":"buffer_list",' +
    '"description":"Lists the files loaded in the IDE (modified or not, and ' +
    'whether this bridge loaded them headless) and the scratch units.",' +
    '"inputSchema":{"type":"object","properties":{' + InstanceProp + '}}},' +

    '{"name":"buffer_save",' +
    '"description":"Saves a loaded IDE buffer to disk, like Ctrl+S in the IDE ' +
    '- also a file the user opened. A headless buffer stays loaded (use ' +
    'buffer_close to release it). Nothing is written when the buffer has no ' +
    'unsaved changes.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path."},' +
    InstanceProp + '},"required":["file"]}},' +

    '{"name":"buffer_close",' +
    '"description":"Closes a HEADLESS buffer that buffer_open loaded: save=true ' +
    'writes it to disk, save=false discards the changes. Files the user opened ' +
    'are never closed from here.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path."},' +
    '"save":{"type":"boolean","description":"true = save to disk, false = ' +
    'discard."},' + InstanceProp + '},"required":["file","save"]}},' +

    '{"name":"scratch_analyze",' +
    '"description":"Creates or replaces a temporary unit that exists ONLY in ' +
    'memory (neither on disk nor in the IDE) and returns DelphiLSP''s ' +
    'diagnostics for it plus the fixes the plugin would offer. The unit is ' +
    'compiled in the context of the IDE''s active project, so it can use the ' +
    'project''s units - handy to test a snippet or a declaration before ' +
    'putting it into real code. The unit header must match the name ' +
    '(unit <name>;).",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"name":{"type":"string","description":"Unit name, e.g. ScratchTest."},' +
    '"content":{"type":"string","description":"Complete unit text."},' +
    InstanceProp + '},"required":["name","content"]}},' +

    '{"name":"scratch_close",' +
    '"description":"Removes a scratch unit (or all of them when name is ' +
    'omitted).",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"name":{"type":"string","description":"Unit name; omit for all."},' +
    InstanceProp + '}}},' +
    '{"name":"find_unit","description":"Which unit(s) declare an identifier - from the plugin''s' +
    ' identifier index over the library, browsing and project paths - and whether the compiler ' +
    'can see each unit (availability). partial=true searches by substring.","inputSchema":{"typ' +
    'e":"object","properties":{"identifier":{"type":"string"},"partial":{"type":"boolean"},"max' +
    '":{"type":"integer"},"instance":{"type":"integer","description":"Process id of the IDE to ' +
    'use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"required"' +
    ':["identifier"]}}' +
    ',' +
    '{"name":"extract_variable","description":"Extracts the selected expression into an ' +
    'inline variable (\"var LName := <expr>;\") right before the statement that contains it ' +
    'and replaces the expression with the name. Refused where hoisting would change behaviour ' +
    '(after a short-circuit and/or, in a loop condition, as the sole statement of a branch, ' +
    'inside a with). apply=false (the default) only shows the change.","inputSchema":{"type":' +
    '"object","properties":{"file":{"type":"string"},"line":{"type":"integer",' +
    '"description":"1-based line of the expression."},"expression":{"type":"string",' +
    '"description":"The text to extract; located on that line. Alternative to column + ' +
    'end_column."},"column":{"type":"integer","description":"1-based start column ' +
    '(with end_column), or where to start looking for \"expression\"."},"end_column":' +
    '{"type":"integer","description":"1-based, exclusive."},"name":{"type":"string",' +
    '"description":"Name of the new variable. Default: derived from the expression."},' +
    ApplyProp + ',' + TokenProp + ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"wrap_try_finally","description":"Wraps the lines from_line..to_line in a ' +
    'try..finally block. The cleanup is inferred from the statement right before them ' +
    '(\"X := TFoo.Create\" -> X.Free, BeginUpdate -> EndUpdate, Enter/Acquire/Lock -> the ' +
    'counterpart), otherwise a TODO comment is inserted - pass \"cleanup\" to say it ' +
    'yourself. Only wrapper lines are added. apply=false (the default) only shows the ' +
    'change.","inputSchema":{"type":"object","properties":{"file":{"type":"string"},' +
    '"from_line":{"type":"integer","description":"1-based first line."},"to_line":' +
    '{"type":"integer","description":"1-based last line. Default: from_line."},"cleanup":' +
    '{"type":"string","description":"The statement for the finally block, e.g. ' +
    '\"List.Free;\"."},' + ApplyProp + ',' + TokenProp + ',' + InstanceProp + '},' +
    '"required":["file","from_line"]}}' +
    ',' +
    '{"name":"move_to_unit","description":"Moves the declaration at a position (type, ' +
    'class incl. its method implementations, routine, const, var) into an EXISTING unit and ' +
    'updates the uses clauses of both units and of every unit using the symbol. Use ' +
    'move_to_new_unit when the target does not exist yet. apply=false (the default) reports ' +
    'the plan.","inputSchema":{"type":"object","properties":{"file":{"type":"string"},' +
    '"line":{"type":"integer","description":"1-based line of the identifier."},' +
    '"column":{"type":"integer","description":"1-based column."},"target_file":' +
    '{"type":"string","description":"The unit to move into (path, or a name next to ' +
    'the source)."},' + ApplyProp + ',' + TokenProp + ',' + InstanceProp + '},"required":' +
    '["file","line","column","target_file"]}}' +
    ',' +
    '{"name":"remove_with","description":"Rewrites \"with X do\" statements: every ' +
    'member of the body is qualified explicitly, verified per identifier with DelphiLSP. ' +
    'Pass file (+ line for the one statement enclosing it), files or project=true. An ' +
    'occurrence the rewriter cannot handle (several targets, unresolved type, inactive ' +
    '{$IFDEF} region) is listed with the reason and left alone. apply=false (the default) ' +
    'shows before/after per statement.","inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string"},"line":{"type":"integer","description":"1-based line ' +
    'inside the with-statement; without it the whole file is scanned."},"files":{"type":' +
    '"array","items":{"type":"string"}},"project":{"type":"boolean"},"inline_vars":' +
    '{"type":"boolean","description":"Introduce inline variables for complex targets ' +
    '(Delphi 10.3+). Default true."},' + ApplyProp + ',' + TokenProp + ',' + InstanceProp +
    '}}}' +
    ',' +
    '{"name":"cleanup_uses","description":"Runs what the uses-cleanup dialog runs: ' +
    'removes entries no identifier of which is used and (move_to_implementation) moves ' +
    'interface entries that are only needed in the implementation. Entries with ' +
    'initialization code, entries the FORM DESIGNER writes itself (asked from the loaded ' +
    'form - it would re-add them on the next save), entries on the user keep list and ' +
    'anything the analysis is unsure about are kept, and the answer says so per entry. In a ' +
    'form unit whose form is NOT loaded in the IDE, textually unused entries are reported ' +
    'as \"unverified\" and kept unless include_unverified is set. apply=false (the ' +
    'default) only shows the change.","inputSchema":{"type":"object","properties":' +
    '{"file":{"type":"string"},' +
    '"remove_unused":{"type":"boolean","description":"Default true."},' +
    '"move_to_implementation":{"type":"boolean","description":"Default false."},' +
    '"include_unverified":{"type":"boolean","description":"Default false. Also act on ' +
    'entries of a form unit whose designer could not be asked - outside the IDE nothing ' +
    're-adds them, so a wrong removal can break the build or the form streaming."},' +
    ApplyProp + ',' + TokenProp + ',' + InstanceProp + '},"required":["file"]}}' +
    ',' +
    '{"name":"find_unit_references","description":"Which units use the given unit, and ' +
    'where: every hit verified with DelphiLSP, plus one row with \"unused\": true per unit ' +
    'that lists it in its uses clause without referencing anything of it. Read-only.",' +
    '"inputSchema":{"type":"object","properties":{"file":{"type":"string",' +
    '"description":"Absolute path of the unit whose references you want."},"max":' +
    '{"type":"integer","description":"Maximum rows (default 400)."},' + InstanceProp +
    '},"required":["file"]}}' +
    ',' +
    '{"name":"find_original_symbol","description":"Where the identifier at a position is ' +
    'DECLARED - DelphiLSP first, then the type of its qualifier, then the identifier index. ' +
    'Answers in cases lsp_definition does not: a private/public overload pair (Delphi 13.1 ' +
    'RSS-5463) or a unit that is not open. Read-only.","inputSchema":{"type":"object",' +
    '"properties":{"file":{"type":"string"},"line":{"type":"integer","description":' +
    '"1-based."},"column":{"type":"integer","description":"1-based."},' + InstanceProp +
    '},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"extract_method","description":"Extracts the lines from_line..to_line ' +
    'into a new method of the enclosing class (or a local routine): which variables ' +
    'become parameters, which become locals and which one becomes the Result is resolved ' +
    'with DelphiLSP. apply=false (the default) answers with the generated code - the ' +
    'routine, the call that replaces the block and the declaration line - since that is ' +
    'what has to be judged; the write itself is not a line diff.","inputSchema":{"type":' +
    '"object","properties":{"file":{"type":"string"},"from_line":{"type":"integer",' +
    '"description":"1-based first line of the block."},"to_line":{"type":"integer",' +
    '"description":"1-based last line. Default: from_line."},"name":{"type":"string",' +
    '"description":"Name of the new method. Default ExtractedMethod."},' + ApplyProp + ',' +
    InstanceProp + '},"required":["file","from_line"]}}' +
    ',' +
    '{"name":"extract_interface","description":"Extracts an interface from the class at ' +
    'a line: a new unit with \"IXxx = interface\" plus a GUID, the class gets the ' +
    'interface in its ancestor list and both uses clauses are updated. ' +
    'add_to_existing=true adds the members to an interface that already exists ' +
    '(\"interface_name\"). apply=false (the default) answers with the interface text ' +
    'that would be written.","inputSchema":{"type":"object","properties":{"file":' +
    '{"type":"string"},"line":{"type":"integer","description":"1-based line inside ' +
    'or at the class declaration."},"interface_name":{"type":"string","description":' +
    '"Name of the interface. Default: the class name with T replaced by I. Required ' +
    'with add_to_existing."},"target_file":{"type":"string","description":"The new ' +
    'unit (path or a name next to the source)."},"members":{"type":"array","items":' +
    '{"type":"string"},"description":"Member names to include. Default: every public ' +
    'and published method and property."},"add_to_existing":{"type":"boolean"},' +
    ApplyProp + ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"add_iinterface","description":"Adds IInterface support to a class that ' +
    'does not descend from TInterfacedObject: IInterface in the ancestor list, an ' +
    'FRefCount field, QueryInterface / _AddRef / _Release and the NewInstance / ' +
    'AfterConstruction pair that mirrors TInterfacedObject''s initial-refcount trick. ' +
    'Afterwards the instance frees itself when the last interface reference drops. ' +
    'apply=false (the default) returns the code it would add.","inputSchema":{"type":' +
    '"object","properties":{"file":{"type":"string"},"line":{"type":"integer",' +
    '"description":"1-based line inside or at the class declaration."},' + ApplyProp +
    ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"signature_check","description":"Every declaration and implementation of ' +
    'the method at a position - interface, class, implementation header - and whether ' +
    'they agree. apply=true aligns the diverging ones with the majority signature ' +
    '(class declarations first, then the implementations, which keep their parameter ' +
    'NAMES because the body uses them).","inputSchema":{"type":"object","properties":' +
    '{"file":{"type":"string"},"line":{"type":"integer","description":"1-based."},' +
    '"column":{"type":"integer","description":"1-based."},' + ApplyProp + ',' +
    InstanceProp + '},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"dfm_events","description":"Checks every form of the project: an event ' +
    'assigned in the .dfm whose handler is MISSING in the unit, or whose parameter list ' +
    'does not match the event type (the cause of hard-to-find stack corruption). Each ' +
    'issue carries an id; apply=true with \"fix_ids\" generates the missing handlers / ' +
    'corrects the parameter lists. Nothing is fixed without naming ids - a generated ' +
    'empty handler shadows an inherited one.","inputSchema":{"type":"object",' +
    '"properties":{"fix_ids":{"type":"array","items":{"type":"string"},' +
    '"description":"Ids from a previous call."},' + ApplyProp + ',' + InstanceProp +
    '}}}' +
    ',' +
    '{"name":"interface_guids","description":"Every interface declaration of the ' +
    'project with its GUID, flagging DUPLICATE GUIDs (which make Supports / ' +
    'QueryInterface return the wrong object) and interfaces without one. An interface ' +
    'paired with a dispinterface on the same GUID is a type-library import and not ' +
    'flagged. Read-only.","inputSchema":{"type":"object","properties":' +
    '{"only_problems":{"type":"boolean","description":"Default true."},' +
    InstanceProp + '}}}' +
    ',' +
    '{"name":"add_unit","description":"Adds a unit to the uses clause of a file (interface or i' +
    'mplementation), minimal edit, IDE buffer when the file is open. Refuses units only on the ' +
    'browsing path. apply=false (the default) only shows the change.","inputSchema":{"type":"object","properties":{"file":{"type":"string","desc' +
    'ription":"Absolute path of the unit."},"unit":{"type":"string"},"section":{"type":"string"' +
    ',"enum":["interface","implementation"]},' + ApplyProp + ',' + TokenProp + ',' +
    '"instance":{"type":"integer","description":"Proces' +
    's id of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instan' +
    'ces)."}},"required":["file","unit"]}}' +
    ',' +
    '{"name":"remove_unit","description":"Removes a unit from the uses clause of a file. apply=' +
    'false (the default) only shows the change.","inpu' +
    'tSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute pa' +
    'th of the unit."},"unit":{"type":"string"},' + ApplyProp + ',' + TokenProp + ',' +
    '"instance":{"type":"integer","description":"Pro' +
    'cess id of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_ins' +
    'tances)."}},"required":["file","unit"]}}' +
    ',' +
    '{"name":"analyze_uses","description":"Uses-clause analysis of one unit: every entry with v' +
    'erdict used / unused / movable (only needed in the implementation) / init_code / ide_manag' +
    'ed (the form designer writes it itself - see reason) / kept_by_user / unverified (form not' +
    ' loaded, so nothing could confirm it) / unknown, its usage count and designerVerified for ' +
    'the file.","inputSchema":{"type":"object","properties":{"file":{"t' +
    'ype":"string","description":"Absolute path of the unit."},"instance":{"type":"integer","de' +
    'scription":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatical' +
    'ly (see ide_instances)."}},"required":["file"]}}' +
    ',' +
    '{"name":"find_implementations","description":"Implementations of the method, interface or ' +
    'property at a position: classes implementing an interface method, classes implementing / d' +
    'escending from a type (position on the type name), property accessors. Scans the project s' +
    'cope files on disk.","inputSchema":{"type":"object","properties":{"file":{"type":"string",' +
    '"description":"Absolute path of the unit."},"line":{"type":"integer","description":"1-base' +
    'd line."},"column":{"type":"integer","description":"1-based column (anywhere on the identi' +
    'fier)."},"instance":{"type":"integer","description":"Process id of the IDE to use. Normall' +
    'y omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["file","li' +
    'ne","column"]}}' +
    ',' +
    '{"name":"find_references","description":"All references to the identifier at a position: D' +
    'elphiLSP references when available, otherwise a project-wide text scan with every hit veri' +
    'fied through DelphiLSP GotoDefinition.","inputSchema":{"type":"object","properties":{"file' +
    '":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","d' +
    'escription":"1-based line."},"column":{"type":"integer","description":"1-based column (any' +
    'where on the identifier)."},"instance":{"type":"integer","description":"Process id of the ' +
    'IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"re' +
    'quired":["file","line","column"]}}' +
    ',' +
    '{"name":"uses_path","description":"Shortest chain of uses entries from one project unit to' +
    ' another - the answer to \"F2047 circular unit reference\": which chain would the new uses' +
    ' entry close.","inputSchema":{"type":"object","properties":{"from_unit":{"type":"string"},' +
    '"to_unit":{"type":"string"},"interface_only":{"type":"boolean","description":"Only interfa' +
    'ce-section uses (default true)."},"instance":{"type":"integer","description":"Process id o' +
    'f the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."' +
    '}},"required":["from_unit","to_unit"]}}' +
    ',' +
    '{"name":"uses_cycles","description":"Circular unit references of the active project (all c' +
    'ycles, or only those through one unit) plus the uses entries whose removal breaks the most' +
    ' cycles.","inputSchema":{"type":"object","properties":{"unit":{"type":"string"},"max":{"ty' +
    'pe":"integer"},"instance":{"type":"integer","description":"Process id of the IDE to use. N' +
    'ormally omitted - the IDE is chosen automatically (see ide_instances)."}}}}' +
    ',' +
    '{"name":"safe_delete","description":"Safe delete of the symbol at a position (method, rou' +
    'tine, field, property, variable, constant, single-line type): checks that NOTHING uses it ' +
    '(every occurrence in the project scope verified via DelphiLSP - an occurrence DelphiLSP ca' +
    'nnot resolve, e.g. in an inactive IFDEF branch, counts as a use; form files; interface imp' +
    'lementations) and refuses overloaded / virtual / override / published members. Without ap' +
    'ply it only reports; apply=true deletes declaration + implementation in the IDE buffer (no' +
    't saved) when the check allows it.","inputSchema":{"type":"object","properties":{"file":{"' +
    'type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","descr' +
    'iption":"1-based line of the identifier (a use or the declaration)."},"column":{"type":"in' +
    'teger","description":"1-based column."},"apply":{"type":"boolean","description":"Delete it' +
    ' when deletable (default false)."},"instance":{"type":"integer","description":"Process id ' +
    'of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances).' +
    '"}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"change_signature","description":"Change method signature of the method / routine ' +
    'at a position: add, remove, reorder or rename parameters, change type, modifier or default' +
    ' value. The whole FAMILY changes together (declaration + implementation, interface methods' +
    ' and every implementing class, the virtual/override chain) and every call site DelphiLSP v' +
    'erifies is rewritten (arguments reordered, values for new parameters inserted, defaults th' +
    'e calls relied on written out); renamed parameters are followed into the bodies. Blocked: ' +
    'overloaded / message methods, event handlers bound in a form, method references and proper' +
    'ty accessors when the change is more than a rename, removed parameters still used in a bod' +
    'y. Without params it only reports the current parameters, the family and the calls. With p' +
    'arams it returns the plan (errors, warnings, edits with the resulting lines); apply=true w' +
    'rites it (open units in the IDE buffer, not saved; closed units on disk).","inputSchema":{' +
    '"type":"object","properties":{"file":{"type":"string","description":"Absolute path of the ' +
    'unit."},"line":{"type":"integer","description":"1-based line of the method name (a call or' +
    ' the declaration)."},"column":{"type":"integer","description":"1-based column."},"params":' +
    '{"type":"array","description":"The NEW parameter list in order. Each item: name, type, mod' +
    'ifier (const/var/out/constref), default, from (the OLD parameter it replaces - omit for a ' +
    'new parameter; an old parameter keeps type/modifier/default unless given), value (new para' +
    'meters only: the argument existing calls get).","items":{"type":"object","properties":{"na' +
    'me":{"type":"string"},"type":{"type":"string"},"modifier":{"type":"string"},"default":{"ty' +
    'pe":"string"},"from":{"type":"string"},"value":{"type":"string"}},"required":["name"]}},"a' +
    'pply":{"type":"boolean","description":"Apply the plan when it has no errors (default false' +
    ')."},"instance":{"type":"integer","description":"Process id of the IDE to use. Normally om' +
    'itted - the IDE is chosen automatically (see ide_instances)."}},"required":["file","line",' +
    '"column"]}}' +
    ',' +
    '{"name":"semantic_replace","description":"Semantic replace with the rules of <project root' +
    '>\\semantic-replace.json (find/replace of dotted identifier paths, comment- and string-awa' +
    're, uses units added, optional local-var hoisting). Every match is VERIFIED through Delphi' +
    'LSP (the declaration of the last identifier): with a rule''s declaredIn it must be declared' +
    ' in that unit, otherwise all matches of a rule must lead to the same declaration (the most' +
    ' frequent one). Matches of another symbol are never replaced; matches DelphiLSP cannot res' +
    'olve (inactive IFDEF branch) only with include_unverified. Without apply it only reports."' +
    ',"inputSchema":{"type":"object","properties":{"files":{"type":"array","items":{"type":"str' +
    'ing"},"description":"Units to process (default: all project units)."},"apply":{"type":"boo' +
    'lean","description":"Replace (default false = report)."},"include_unverified":{"type":"boo' +
    'lean","description":"Also replace matches DelphiLSP gave no answer for (default false)."},' +
    '"max":{"type":"integer","description":"Max match rows in the result (default 100)."},"inst' +
    'ance":{"type":"integer","description":"Process id of the IDE to use. Normally omitted - th' +
    'e IDE is chosen automatically (see ide_instances)."}}}}' +
    ',' +
    '{"name":"move_to_new_unit","description":"Moves the declaration at a position (type, class' +
    ' incl. its method implementations, routine, const, var) into a NEW unit: creates <new_unit' +
    '>.pas (next to the file, or at the given path), adds it to the project, moves the code, ad' +
    'ds the needed uses (the source unit goes into the new unit''s implementation uses when only' +
    ' the moved implementation needs it) and updates the uses of the units using the symbol. Re' +
    'fused when the declaration itself needs identifiers of the source unit (circular unit refe' +
    'rence) - nothing is created then. apply=false (the default) runs every check and reports ' +
    'what would move, without creating the unit.","inputSchema":{"type":"object","properties":{"file":{"t' +
    'ype":"string","description":"Absolute path of the unit."},"line":{"type":"integer","descri' +
    'ption":"1-based line of the identifier."},"column":{"type":"integer","description":"1-base' +
    'd column."},"new_unit":{"type":"string","description":"Name of the new unit (e.g. \"Custom' +
    'er.List\") or a full path."},' + ApplyProp + ',' + TokenProp + ',' +
    '"instance":{"type":"integer","description":"Process id of the' +
    ' IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"r' +
    'equired":["file","line","column","new_unit"]}}' +
    ',' +
    '{"name":"expand_includes","description":"Writes the content of every {$I}/{$INCLUDE} f' +
    'ile IN PLACE into the including source (for debugging), framed by marker comments that ke' +
    'ep the original directive (// >>> include begin: ... / // <<< include end: ...). Nested in' +
    'cludes too. Files open in the IDE are changed in the editor buffer (not saved), closed fil' +
    'es ON DISK - revert with version control. Pass file, files, directory (recursive, .pas/.dp' +
    'r/.dpk) or project=true. apply=false (the default) only reports what would be expanded.","' +
    'inputSchema":{"type":"object","properties":{"file":{"type":"st' +
    'ring"},"files":{"type":"array","items":{"type":"string"}},"directory":{"type":"string"},"p' +
    'roject":{"type":"boolean"},' + ApplyProp + ',' +
    '"instance":{"type":"integer","description":"Process id of the I' +
    'DE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}}}}' +
    ',' +
    '{"name":"convert_properties","description":"Converts the properties declared on lines f' +
    'rom_line..to_line of a class/record between direct field access and getter/setter methods.' +
    ' to_accessors: read/write FX -> GetX/SetX, declarations added to the private section, impl' +
    'ementations after the type''s last method. to_fields: only TRIVIAL accessors (Result := FX' +
    ' / FX := Value), which are then removed; refused when used elsewhere, virtual/override/ove' +
    'rload. Every property gets a row with the result or the reason. apply=true changes the IDE' +
    ' buffer (not saved).","inputSchema":{"type":"object","properties":{"file":{"type":"string' +
    '"},"from_line":{"type":"integer","description":"1-based."},"to_line":{"type":"integer"},"d' +
    'irection":{"type":"string","enum":["to_accessors","to_fields"]},"getter":{"type":"boolean"' +
    '},"setter":{"type":"boolean"},"apply":{"type":"boolean"},"instance":{"type":"integer","des' +
    'cription":"Process id of the IDE to use. Normally omitted - the IDE is chosen automaticall' +
    'y (see ide_instances)."}},"required":["file","from_line"]}}' +
    ',' +
    '{"name":"debug_consistency","description":"Checks the active project for reasons why break' +
    'points are not hit or debug info does not match: duplicate sources on the search path, str' +
    'ay DCUs, LF line endings, debug options and directives, outdated executable / symbol files' +
    ', host application, duplicate DLL/BPL copies. Returns counts per kind and at most \"max\" ' +
    'rows (default 60, problems first).","inputSchema":{"type":"object","properties"' +
    ':{"max":{"type":"integer","description":"Rows to list, 0 = all."},"instance":{"type":"integer","description":"Process id of the IDE to use. Normally omitt' +
    'ed - the IDE is chosen automatically (see ide_instances)."}}}}' +
    ',' +
    '{"name":"blame","description":"git/svn blame of a file (revision, author, date, summary pe' +
    'r line) - the file on disk; at most 500 lines per call.","inputSchema":{"type":"object","p' +
    'roperties":{"file":{"type":"string","description":"Absolute path of the unit."},"start_lin' +
    'e":{"type":"integer"},"end_line":{"type":"integer"},"instance":{"type":"integer","descript' +
    'ion":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatically (se' +
    'e ide_instances)."}},"required":["file"]}}' +
    ',' +
    '{"name":"commit_info","description":"The commit that last changed a line: metadata, messag' +
    'e, changed files; diff_file adds the diff of the files whose path contains that text.","in' +
    'putSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute ' +
    'path of the unit."},"line":{"type":"integer","description":"1-based line."},"diff_file":{"' +
    'type":"string"},"instance":{"type":"integer","description":"Process id of the IDE to use. ' +
    'Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["fi' +
    'le","line"]}}' +
    ',' +
    '{"name":"rename_preview","description":"Previews renaming the identifier at a position wit' +
    'h the plugin''s rename pipeline (text scan + DelphiLSP verification, interface implementati' +
    'ons, .dfm/.fmx form files). SAVES ALL MODIFIED FILES FIRST, like the rename dialog. Return' +
    's the changes and a token for rename_apply.","inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"intege' +
    'r","description":"1-based line."},"column":{"type":"integer","description":"1-based column' +
    ' (anywhere on the identifier)."},"new_name":{"type":"string"},"scope":{"type":"string","en' +
    'um":["project","unit","method"]},"include_open_units":{"type":"boolean"},"include_used_uni' +
    'ts":{"type":"boolean"},"instance":{"type":"integer","description":"Process id of the IDE t' +
    'o use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"require' +
    'd":["file","line","column","new_name"]}}' +
    ',' +
    '{"name":"rename_apply","description":"Applies the rename previewed by rename_preview (toke' +
    'n) through the IDE editor - undoable, not saved.","inputSchema":{"type":"object","properti' +
    'es":{"token":{"type":"string"},"instance":{"type":"integer","description":"Process id of t' +
    'he IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},' +
    '"required":["token"]}}' +
    ',' +
    '{"name":"lsp_hover","description":"DelphiLSP hover text at a position (type / declaration ' +
    'info).","inputSchema":{"type":"object","properties":{"file":{"type":"string","description"' +
    ':"Absolute path of the unit."},"line":{"type":"integer","description":"1-based line."},"co' +
    'lumn":{"type":"integer","description":"1-based column (anywhere on the identifier)."},"ins' +
    'tance":{"type":"integer","description":"Process id of the IDE to use. Normally omitted - t' +
    'he IDE is chosen automatically (see ide_instances)."}},"required":["file","line","column"]' +
    '}}' +
    ',' +
    '{"name":"lsp_definition","description":"DelphiLSP go-to-definition for a position.","input' +
    'Schema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute pat' +
    'h of the unit."},"line":{"type":"integer","description":"1-based line."},"column":{"type":' +
    '"integer","description":"1-based column (anywhere on the identifier)."},"instance":{"type"' +
    ':"integer","description":"Process id of the IDE to use. Normally omitted - the IDE is chos' +
    'en automatically (see ide_instances)."}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"lsp_implementation","description":"DelphiLSP go-to-implementation for a position.' +
    '","inputSchema":{"type":"object","properties":{"file":{"type":"string","description":"Abso' +
    'lute path of the unit."},"line":{"type":"integer","description":"1-based line."},"column":' +
    '{"type":"integer","description":"1-based column (anywhere on the identifier)."},"instance"' +
    ':{"type":"integer","description":"Process id of the IDE to use. Normally omitted - the IDE' +
    ' is chosen automatically (see ide_instances)."}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"lsp_references","description":"DelphiLSP textDocument/references for a position (' +
    'when the server supports it).","inputSchema":{"type":"object","properties":{"file":{"type"' +
    ':"string","description":"Absolute path of the unit."},"line":{"type":"integer","descriptio' +
    'n":"1-based line."},"column":{"type":"integer","description":"1-based column (anywhere on ' +
    'the identifier)."},"include_declaration":{"type":"boolean"},"instance":{"type":"integer","' +
    'description":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatic' +
    'ally (see ide_instances)."}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"lsp_document_symbols","description":"DelphiLSP document symbols of a unit (raw LS' +
    'P JSON, 0-based ranges).","inputSchema":{"type":"object","properties":{"file":{"type":"str' +
    'ing","description":"Absolute path of the unit."},"instance":{"type":"integer","description' +
    '":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatically (see i' +
    'de_instances)."}},"required":["file"]}}' +
    ',' +
    '{"name":"lsp_signature_help","description":"DelphiLSP signature help inside a call (raw LS' +
    'P JSON).","inputSchema":{"type":"object","properties":{"file":{"type":"string","descriptio' +
    'n":"Absolute path of the unit."},"line":{"type":"integer","description":"1-based line."},"' +
    'column":{"type":"integer","description":"1-based column (anywhere on the identifier)."},"i' +
    'nstance":{"type":"integer","description":"Process id of the IDE to use. Normally omitted -' +
    ' the IDE is chosen automatically (see ide_instances)."}},"required":["file","line","column' +
    '"]}}' +
    ',' +
    '{"name":"lsp_completion","description":"DelphiLSP completion items at a position (label, d' +
    'etail, kind); prefix filters, max limits (default 100).","inputSchema":{"type":"object","p' +
    'roperties":{"file":{"type":"string","description":"Absolute path of the unit."},"line":{"t' +
    'ype":"integer","description":"1-based line."},"column":{"type":"integer","description":"1-' +
    'based column (anywhere on the identifier)."},"prefix":{"type":"string"},"max":{"type":"int' +
    'eger"},"instance":{"type":"integer","description":"Process id of the IDE to use. Normally ' +
    'omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["file","line' +
    '","column"]}}' +
    ',' +
    '{"name":"lsp_request","description":"Sends any request to the plugin''s DelphiLSP session a' +
    'nd returns the raw result (LSP conventions: 0-based positions, file URIs). \"file\" syncs ' +
    'that document first. Lifecycle and document-sync methods are refused.","inputSchema":{"typ' +
    'e":"object","properties":{"method":{"type":"string"},"params":{"type":"object"},"file":{"t' +
    'ype":"string","description":"Optional: sync this file''s current content before the request' +
    '."},"timeout_ms":{"type":"integer"},"instance":{"type":"integer","description":"Process id' +
    ' of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)' +
    '."}},"required":["method"]}}' +
    ']';

function McpToolDefinitions: TJSONArray;
begin
  Result := TJSONObject.ParseJSONValue(ToolsJson) as TJSONArray;
end;

function McpToolsHash: string;
var
  H: UInt64;
begin
  // FNV-1a, kept in 64-bit arithmetic and masked, so it never trips
  // overflow checking in whatever configuration the unit is built.
  H := 2166136261;
  for var C in ToolsJson do
    H := ((H xor Ord(C)) * 16777619) and $FFFFFFFF;
  Result := IntToHex(Cardinal(H), 8);
end;

// ---------------------------------------------------------------------------
//  Arguments a tool does not know
// ---------------------------------------------------------------------------
//  Every handler reads its arguments by name, so a name that is not in the
//  schema is simply not read - the call runs with a default instead, and
//  the answer looks like a tool that did something else than it was asked.
//  Found by making that mistake: buffer_read with "from_line"/"to_line"
//  (its schema says start_line/end_line) answered the whole file, with
//  nothing in the result saying why.
var
  GArgMapLock: TObject = nil;
  GArgMap: TDictionary<string, TArray<string>> = nil;

procedure EnsureArgMap;
var
  Arr: TJSONArray;
  Names: TArray<string>;
begin
  if GArgMap <> nil then Exit;
  GArgMap := TDictionary<string, TArray<string>>.Create;
  try
    Arr := McpToolDefinitions;
  except
    Arr := nil;    // a broken list must never break a tool call
  end;
  if Arr = nil then Exit;
  try
    for var V in Arr do
    begin
      if not (V is TJSONObject) then Continue;
      var O := TJSONObject(V);
      var Tool := O.GetValue<string>('name', '');
      if Tool = '' then Continue;
      Names := ['instance'];
      var Schema := O.GetValue('inputSchema');
      if Schema is TJSONObject then
      begin
        var Props := TJSONObject(Schema).GetValue('properties');
        if Props is TJSONObject then
          for var P in TJSONObject(Props) do
            if not MatchText(P.JsonString.Value, Names) then
              Names := Names + [P.JsonString.Value];
      end;
      GArgMap.AddOrSetValue(LowerCase(Tool), Names);
    end;
  finally
    Arr.Free;
  end;
end;

function KnownToolArguments(const ATool: string): TArray<string>;
begin
  Result := nil;
  if GArgMapLock = nil then Exit;    // finalized, or called from a DLL tail
  TMonitor.Enter(GArgMapLock);
  try
    EnsureArgMap;
    if not GArgMap.TryGetValue(LowerCase(ATool), Result) then
      Result := nil;
  finally
    TMonitor.Exit(GArgMapLock);
  end;
end;

function UnknownToolArguments(const ATool: string;
  const ANames: TArray<string>): TArray<string>;
var
  Known: TArray<string>;
begin
  Result := nil;
  Known := KnownToolArguments(ATool);
  // A tool the list does not describe (the bridge's own two, or an IDE
  // newer than this unit) says nothing - the alternative would be to
  // report every argument of it as unknown.
  if Length(Known) = 0 then Exit;
  for var N in ANames do
    if (N <> '') and not MatchText(N, Known) and not MatchText(N, Result) then
      Result := Result + [N];
end;

function UnknownArgumentNote(const ATool: string;
  const ANames: TArray<string>): string;
var
  Unknown: TArray<string>;
begin
  Result := '';
  Unknown := UnknownToolArguments(ATool, ANames);
  if Length(Unknown) = 0 then Exit;
  Result := Format(
    'IGNORED argument(s): %s. They are not in this tool''s schema, so ' +
    'nothing read them - the call ran as if they had been left out. %s ' +
    'takes: %s.',
    [string.Join(', ', Unknown), ATool,
     string.Join(', ', KnownToolArguments(ATool))]);
end;

function McpServerInstructions: string;
begin
  // Loaded into EVERY Claude Code session that has the server registered,
  // so it only says what the tool names cannot: how to treat buffers, fix
  // ids and IDE selection. No tool enumeration - the model gets the names
  // anyway (tool search keeps the schemas out until one is needed).
  Result :=
    'Delphi (RAD Studio) refactoring through running IDEs with the ' +
    'Refactoring Light plugin; answers come from DelphiLSP and the IDE, not ' +
    'from text matching. The IDE''s editor buffers are the truth, unsaved ' +
    'changes included: while a unit is open in the IDE, use buffer_read / ' +
    'buffer_edit / buffer_save or the quick fixes instead of editing the ' +
    'file on disk. Fix ids expire with every edit of the unit - list again ' +
    'after applying one. buffer_open without visible=true loads a file ' +
    'headless (memory only), so changes can be checked with get_diagnostics ' +
    'before buffer_close saves or discards them. When something does not ' +
    'work (no fixes, no diagnostics, index not ready), get_status says why. ' +
    'With several IDEs the target follows the file argument and the working ' +
    'directory; if that is ambiguous, use ide_instances and select_ide. ' +
    'While no IDE runs only ide_instances and select_ide exist; the other ' +
    'tools appear when an IDE with the plugin starts.';
end;

initialization
  GArgMapLock := TObject.Create;

finalization
  // The pipe handlers are stopped before this unit finalizes (StopMcpServer
  // waits for them), but a nil lock must still not be entered - same rule
  // as the main-thread guard's.
  FreeAndNil(GArgMap);
  FreeAndNil(GArgMapLock);

end.
