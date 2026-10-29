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

type
  /// <summary>What the bridge exe ON DISK is, compared with the process that
  ///  is talking to us. The two cases need OPPOSITE actions, which is why the
  ///  message must not name both (user, 2026-10-04: "version mismatch trotz
  ///  install.cmd und Claude Code Neustart" - the exe had been replaced
  ///  meanwhile, so only his session was old).</summary>
  TBridgeExeAge = (
    /// <summary>Not known - no process to ask, or Windows refused.</summary>
    beaUnknown,
    /// <summary>The exe was last written BEFORE that process started, so the
    ///  running image IS the installed file: the install did not replace
    ///  it.</summary>
    beaAsOld,
    /// <summary>The exe is NEWER than the process: it has been replaced and
    ///  only this session still runs the old image.</summary>
    beaReplaced);

/// <summary>'' when the two versions agree (nothing to report); otherwise one
///  line naming BOTH numbers and WHAT TO DO. Pure, so the status row,
///  get_status and the tests share one wording. It does not compare which
///  version is newer - that says nothing - but it does use AAGE, because
///  "start a new Claude Code session" and "the install did not replace the
///  exe" are different problems and only one of them is the user's.</summary>
function BridgeVersionProblem(const ABridge, APlugin: string;
  AAge: TBridgeExeAge = beaUnknown): string;

/// <summary>The exe an install moved aside, mapped back to the name it was
///  installed under: buildmcp.cmd renames the LIVE exe to
///  "&lt;exe&gt;.&lt;number&gt;.old" before copying the new one in, so a bridge that
///  started before that install reports the ASIDE name as its own image.
///  '' when the path is not such a name. MEASURED 2026-10-04: without this
///  the age probe compared the aside FILE's write time with the process
///  start - always older, of course - and the status row then said "the
///  installed exe is still the old one, run install.cmd" in exactly the
///  situation where the install HAD worked and only a new session was
///  needed. The user followed that advice twice.</summary>
function CanonicalBridgeExePath(const AImagePath: string): string;

type
  /// <summary>What a STARTING bridge does about the exe it runs from. An
  ///  install cannot overwrite a running exe, so it moves the old one aside
  ///  and writes the new one under the installed name - and the session that
  ///  was running keeps the old code, because no process can swap its own
  ///  image (1.22.1 / 1.22.3). It CAN hand its stdin/stdout to the new exe
  ///  and wait for it: the client keeps seeing one live process the whole
  ///  time, while the new code does the work (user, 2026-10-17).</summary>
  TBridgeRelaunch = record
    /// <summary>Start ANEWEXE on our own std handles and wait for it.</summary>
    Wanted: Boolean;
    /// <summary>The installed exe to hand over to, '' when there is none.</summary>
    NewExe: string;
    /// <summary>One line, in BOTH directions - "it did not hand over" is the
    ///  half that needs explaining, and stderr is where it can be read.</summary>
    Reason: string;
  end;

/// <summary>The image this process runs from, as the KERNEL knows it.
///  That is NOT the same question as ParamStr(0): MEASURED (scratchpad
///  whoami) - after a file is renamed under a running process,
///  GetModuleFileName and ParamStr(0) keep answering the name it was
///  STARTED with, because the loader caches it, while
///  QueryFullProcessImageName follows the rename. An install renames the
///  running bridge aside, so this is the ONLY way the process can notice
///  that it has been replaced - and the reason the check cannot live in
///  the startup code, where nothing has been renamed yet.</summary>
function CurrentProcessImagePath: string;

/// <summary>How often ONE bridge process may hand its pipes on. A chain is
///  finite by construction - every hop needs an install that renames the
///  running exe aside - so this is a backstop and not the rule.</summary>
const
  MaxRelaunchHops = 8;

/// <summary>The version from which a bridge hands over as OFTEN as it is
///  replaced. 1.33.0 to 2.9.4 did it once per session: the successor was
///  started with --no-relaunch, so the SECOND install in one Claude Code
///  session left a stale bridge behind (measured 2026-10-28, and the status
///  row's own sentence promised otherwise).</summary>
  FirstRepeatedHandOver = '2.9.5';

/// <summary>Whether this image should hand its stdio to the exe that
///  replaced it. AIMAGEPATH is CurrentProcessImagePath, never ParamStr(0)
///  - see there. Four reasons NOT to, each
///  named: this IS the installed exe (the normal case), the caller said so
///  (--no-relaunch, the explicit kill switch), it has handed over
///  MaxRelaunchHops times already, or the installed name is not on disk - a
///  packaged host redirects %LOCALAPPDATA%, so the aside copy can live in a
///  view the install never wrote to (1.22.3), and there is nothing to hand
///  over to.
///  AHOP is how often this chain has handed over already; the successor gets
///  it raised by one. A LOOP IS IMPOSSIBLE WITHOUT THAT COUNTER TOO - the
///  successor runs from the INSTALLED name, which this function refuses with
///  "this is the installed exe" - which is why 1.33.0's permanent
///  --no-relaunch only ever cost the thing it was meant to protect: the next
///  install renames THAT process aside, and then it may not react.</summary>
function PlanBridgeRelaunch(const AImagePath: string;
  const AFileExists: TFunc<string, Boolean>;
  ASuppressed: Boolean; AHop: Integer = 0): TBridgeRelaunch;

/// <summary>Of AFILES, the aside copies an install left behind that may be
///  deleted - every "&lt;exe&gt;.&lt;number&gt;.old" except AMYIMAGE, which is the
///  one this process runs from and which Windows refuses to delete anyway.
///  So the copy a hand-over leaves behind goes when the NEXT bridge
///  starts.</summary>
function AsideBridgeExesToDelete(const AFiles: TArray<string>;
  const AMyImage: string): TArray<string>;

// ---- a module a tool had to open must not stay behind invisibly ----

type
  /// <summary>What to do with a module that is loaded in the IDE when a tool
  ///  call ends. A write through the editor helper OPENS the file it edits
  ///  (OpenModuleSafe), so a file the IDE did not have loaded comes back as a
  ///  module with no tab: the edit then lives only in memory, the user cannot
  ///  see it, Ctrl+Z cannot reach it and the IDE asks about it when it
  ///  closes. The quick-fix tools have handled this since 2026-09-18 - this
  ///  is the same rule for every tool.</summary>
  TModuleSweepVerdict = (
    /// <summary>Not ours: the IDE had it before the call, the bridge opened
    ///  it headless on purpose, or the user can see it.</summary>
    msvLeave,
    /// <summary>The call opened it and succeeded: write it to DISK and
    ///  release it, so the result is where the user looks for it.</summary>
    msvSaveAndClose,
    /// <summary>The call opened it and failed: discard, which restores the
    ///  state before the call.</summary>
    msvDiscard);

/// <summary>The rule itself, pure so it can be tested and so there is only
///  ONE of it. The order of the tests is the point: ownership first (a module
///  the IDE already had, or one the bridge promised to keep, is never touched
///  - rename's "changed but not saved, undoable" stays exactly that), then
///  whether the user can SEE it, and only then the success of the call.
/// </summary>
function ModuleSweepVerdict(AWasLoadedBefore, ABridgeHeadless, AHasEditView,
  ACallSucceeded: Boolean): TModuleSweepVerdict;

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
type
  /// <summary>One group of the guide text: a job, the prose that says how
  ///  to do it, and the tools that prose names. The array is an INDEX of
  ///  the text, not a second truth - a test requires every name in it to
  ///  appear in the prose, every tool of McpToolDefinitions to be in
  ///  exactly one section, and no section to name a tool that does not
  ///  exist. So a tool added, renamed or dropped cannot leave the guide
  ///  silently stale.</summary>
  TMcpGuideSection = record
    Title: string;
    Text: string;
    Tools: TArray<string>;
  end;

/// <summary>What an EMPTY lsp_* answer says. APrefix is the tool''s own
///  sentence and is never repeated by the generic text. AAgoMs is how long
///  ago the file was (re)sent, -1 for "not by us"; AUnitAnalysed says
///  whether the session has really analysed this unit, which decides the
///  cause: a unit it never analysed is usually taken from a precompiled
///  .dcu, while a unit that answers hover but no definition is a symbol
///  DelphiLSP does not resolve (the private/public overload pair of
///  RSS-5463 is the known one) - and THEN find_original_symbol, which asks
///  the identifier index, is the way out. Blaming the .dcu in that case
///  sent a reviewing session looking in the wrong place.</summary>
function LspEmptyAnswerNote(const AEmptySentence, ADescription,
  AFile: string; AAgoMs: Integer;
  AUnitAnalysed, ASuggestIndex, AAboutSymbol: Boolean): string;

/// <summary>The sections the guide is built from.</summary>
function McpGuideSections: TArray<TMcpGuideSection>;
/// <summary>The whole guide as plain text - the answer of the "guide"
///  tool. It costs nothing until it is called, which is why it may say
///  what the always-loaded instructions deliberately leave out.</summary>
function McpGuideText: string;

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
  System.Generics.Collections,
  // the guide names the build it describes; Expert.Version is
  // the one place that knows the number and uses nothing itself
  Expert.Version;

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

function ModuleSweepVerdict(AWasLoadedBefore, ABridgeHeadless, AHasEditView,
  ACallSucceeded: Boolean): TModuleSweepVerdict;
begin
  if AWasLoadedBefore or ABridgeHeadless or AHasEditView then
    Exit(msvLeave);
  if ACallSucceeded then
    Result := msvSaveAndClose
  else
    Result := msvDiscard;
end;

function CanonicalBridgeExePath(const AImagePath: string): string;
var
  Base: string;
begin
  Result := '';
  if not EndsText('.old', AImagePath) then Exit;
  Base := Copy(AImagePath, 1, Length(AImagePath) - Length('.old'));
  // "<exe>.<number>" - the number is what makes the aside name unique, and
  // only our own install writes it. Anything else is not ours to interpret.
  var Dot := LastDelimiter('.', Base);
  if Dot <= 1 then Exit;
  var Num := Copy(Base, Dot + 1, MaxInt);
  if Num = '' then Exit;
  for var C in Num do
    if not CharInSet(C, ['0'..'9']) then Exit;
  Base := Copy(Base, 1, Dot - 1);
  if not EndsText('.exe', Base) then Exit;
  Result := Base;
end;

function QueryFullProcessImageNameW(hProcess: THandle; dwFlags: DWORD;
  lpExeName: PWideChar; var lpdwSize: DWORD): BOOL; stdcall;
  external kernel32 name 'QueryFullProcessImageNameW';

function CurrentProcessImagePath: string;
var
  Buf: array[0..1023] of Char;
  Len: DWORD;
begin
  Len := Length(Buf);
  if QueryFullProcessImageNameW(GetCurrentProcess, 0, @Buf[0], Len) then
    Result := string(PChar(@Buf[0]))
  else
    Result := ParamStr(0);
end;

function PlanBridgeRelaunch(const AImagePath: string;
  const AFileExists: TFunc<string, Boolean>;
  ASuppressed: Boolean; AHop: Integer): TBridgeRelaunch;
begin
  Result := Default(TBridgeRelaunch);
  Result.NewExe := CanonicalBridgeExePath(AImagePath);
  if Result.NewExe = '' then
  begin
    Result.Reason := 'this is the installed exe';
    Exit;
  end;
  if ASuppressed then
  begin
    Result.Reason := 'started from the aside copy on purpose (--no-relaunch)';
    Result.NewExe := '';
    Exit;
  end;
  if AHop >= MaxRelaunchHops then
  begin
    Result.Reason := Format('an install replaced this exe, but this chain has '
      + 'already handed the pipes on %d times in this session - not again',
      [AHop]);
    Result.NewExe := '';
    Exit;
  end;
  if not Assigned(AFileExists) or not AFileExists(Result.NewExe) then
  begin
    Result.Reason := Format('an install moved this exe aside, but %s is not ' +
      'there - nothing to hand over to', [Result.NewExe]);
    Result.NewExe := '';
    Exit;
  end;
  Result.Wanted := True;
  Result.Reason := Format('an install replaced this exe - handing stdin/' +
    'stdout over to %s', [Result.NewExe]);
end;

function AsideBridgeExesToDelete(const AFiles: TArray<string>;
  const AMyImage: string): TArray<string>;
var
  F: string;
begin
  Result := nil;
  for F in AFiles do
    if (CanonicalBridgeExePath(F) <> '') and not SameText(F, AMyImage) then
      Result := Result + [F];
end;

function BridgeVersionProblem(const ABridge, APlugin: string;
  AAge: TBridgeExeAge): string;
begin
  Result := '';
  if (APlugin = '') or (ABridge = APlugin) then Exit;
  if ABridge = '' then
    Result := Format('the bridge exe does not report its version, so it is ' +
      'older than %s - if its tools behave oddly (timeouts, missing tools), ' +
      'run install.cmd and start a new Claude Code session', [APlugin])
  else if AAge = beaReplaced then
    // The file is already right; a process cannot swap its own image.
    Result := Format('bridge exe %s, plugin %s - but the exe ON DISK has ' +
      'been replaced since this bridge started, so it is only the running ' +
      'session that is old. A bridge of %s or newer hands the pipes over to ' +
      'the new exe by itself with the next request, as often as it is ' +
      'replaced; 1.33.0 to 2.9.4 did it ONCE per session, so a SECOND ' +
      'install in one session needs a new Claude Code session - and so does ' +
      'anything older. And the exe on disk is only as new as the last ' +
      'install.cmd: the IDE''s own "Install" renews this BPL alone, so when ' +
      'a new session still reports a mismatch, that install is what is ' +
      'missing', [ABridge, APlugin, FirstRepeatedHandOver])
  else
    Result := Format('bridge exe %s, plugin %s - these must match. The ' +
      'installed exe is still the old one, so run install.cmd (it reports ' +
      'whether the bridge was really replaced) and then start a new Claude ' +
      'Code session', [ABridge, APlugin]);
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
  // What is true for EVERY writing tool. It used to end "plus a token",
  // which only HALF of them hand out - the three variants below say per
  // tool what protects an apply, because "is the preview pinned?" is a
  // question the caller cannot answer from the outside (reported by a
  // reviewing model, 2026-10-19).
  ApplyProp =
    '"apply":{"type":"boolean","description":"Default false: nothing is ' +
    'written and the answer lists the changes (file, line, before, after). ' +
    'true makes the change."}';
  // ...and it hands out a token that pins the files AND the arguments.
  ApplyTokenProp =
    '"apply":{"type":"boolean","description":"Default false: nothing is ' +
    'written and the answer lists the changes (file, line, before, after) ' +
    'plus a token. true makes the change - pass the token back with it and ' +
    'the change is refused when a file, or an argument of this call, ' +
    'differs from the preview."}';
  // ...and it needs no token, because it looks again itself.
  ApplyFreshProp =
    '"apply":{"type":"boolean","description":"Default false: nothing is ' +
    'written and the answer lists what would be written. true makes the ' +
    'change. There is no token and none is needed: apply re-runs this ' +
    'tool''s own analysis on the CURRENT buffers, so it can never carry ' +
    'out a plan made for a state that is gone - it answers for what it ' +
    'finds now, or refuses."}';
  // ...and it acts on the ids of a previous scan.
  ApplyIdsProp =
    '"apply":{"type":"boolean","description":"Default false: nothing is ' +
    'written and the findings are listed with an id each. true needs ' +
    'fix_ids and acts on exactly those. There is no token: apply ' +
    're-scans, and an id names the PLACE, so an id from an older scan ' +
    'finds nothing rather than something else - it comes back in ' +
    'unknownIds and nothing is done for it."}';
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

    '{"name":"guide",' +
    '"description":"What this server can do, as plain text: the three ' +
    'rules that save a round (the IDE buffer is the truth, apply=false ' +
    'first, an id or token belongs to ONE answer), where to start, and ' +
    'which tool for which job - all of them, grouped. Takes no ' +
    'arguments, needs no running IDE and changes nothing. Read this when ' +
    'you do not know this toolbox; a tool''s own description says how to ' +
    'CALL it.",' +
    '"inputSchema":{"type":"object","properties":{}}},' +
    '{"name":"settings",' +
    '"description":"Reads and writes this plugin''s own options - what ' +
    'the options page in the IDE offers, which no session can click. ' +
    'Without arguments it lists every setting with its current value, its ' +
    'default, its range or choices, what it does and WHERE the current ' +
    'value lives. With name + value + scope it sets one: ' +
    'scope=permanent writes the registry, so it holds for this IDE and ' +
    'every one started later; scope=session changes it for the RUNNING ' +
    'IDE only - it takes effect at once, the registry keeps what it had, ' +
    'and the value is gone when the IDE closes. There is deliberately no ' +
    'default scope: a change that may outlive the IDE is not one to guess ' +
    'at. LspLogging and VerifySession are read when a DelphiLSP session ' +
    'STARTS, so they apply to the next one.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"name":{"type":"string","description":"The setting to change, as the listing names it (case-insensitive). Omit it to only list."},' +
    '"value":{"type":"string","description":"The new value as text: true / false for a switch, a number inside its range, the text itself otherwise. Required together with name."},' +
    '"scope":{"type":"string","description":"permanent = store it in the registry; session = this running IDE only. Required together with name."},' +
    InstanceProp + '}}},' +
    '{"name":"lsp_restart",' +
    '"description":"Ends this plugin''s DelphiLSP session, so the next ' +
    'request that needs it starts a fresh one (the verification session ' +
    'goes with it). Closing and reopening the project does NOT restart ' +
    'it, so this is the way to pick up something the session only reads ' +
    'at startup - the .delphilsp.json, the library path, LspLogging - or ' +
    'to get rid of .dcu files it has loaded instead of their sources. ' +
    'Costs the usual cold start; get_status says when it is ready ' +
    'again. By default a new session is started right away - the lsp_* ' +
    'tools never cold-start one themselves, so without that they would ' +
    'answer \"not running\" until a search or a quick fix asks for ' +
    'it.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"start_again":{"type":"boolean","description":"Start a new session right away (default true). With false the session stays down until something asks for it."},' +
    InstanceProp + '}}},' +
    '{"name":"project_open",' +
    '"description":"Opens a project (.dproj / .dpk) or a project group ' +
    'in the IDE. Opening one closes what is open, so unsaved work is ' +
    'decided BEFORE: with save_modified=false (the default) the call is ' +
    'REFUSED and names the modified files, because the IDE would ask and ' +
    'nobody can answer that question through a pipe. The identifier ' +
    'index and the DelphiLSP session follow a switch by themselves but ' +
    'need time (get_status), and a session that was already running ' +
    'keeps answering for the old project until lsp_restart.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path of the .dproj, .dpk, .dpr or .groupproj."},' +
    '"new_group":{"type":"boolean","description":"Open it in a NEW project group (default false: it joins the current group the way the IDE does it)."},' +
    '"save_modified":{"type":"boolean","description":"Save every modified file first (default false = refuse while anything is unsaved)."},' +
    InstanceProp + '},"required":["file"]}},' +

    '{"name":"project_close",' +
    '"description":"Closes the project group and every open file - what ' +
    'File > Close All does, because the IDE has no \"close only the ' +
    'project\". Unsaved work is decided the same way as for ' +
    'project_open: refused unless save_modified=true. Our DelphiLSP ' +
    'session is not ended by it (lsp_restart does that), and without a ' +
    'project most other tools have nothing to work on.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"save_modified":{"type":"boolean","description":"Save every modified file first (default false = refuse while anything is unsaved)."},' +
    InstanceProp + '}}},' +

    '{"name":"project_files",' +
    '"description":"The files the active project lists, and adding or ' +
    'removing some. Without add / remove it only lists (unit and form ' +
    'name per entry). add takes paths that must EXIST - adding does not ' +
    'create a file and does not open it; remove takes paths the project ' +
    'lists - removing does not delete anything from disk. One call can ' +
    'do both, and \"changes\" says per file what happened or what ' +
    'would. The project FILE is not an editor buffer: a change to it is ' +
    'not undoable with Ctrl+Z, which is why it takes apply.",' +
    '"inputSchema":{"type":"object","properties":{' +
    '"add":{"type":"array","items":{"type":"string"},"description":"Paths to add to the project."},' +
    '"remove":{"type":"array","items":{"type":"string"},"description":"Paths to remove from the project."},' +
    '"save":{"type":"boolean","description":"Write the project file afterwards (default true). With false the change lives in the IDE''s model only and the IDE asks when the project closes."},' +
    ApplyFreshProp + ',' + InstanceProp + '}}},' +

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
    'e":"object","properties":{"identifier":{"type":"string","description":"The identifier to look up."},"partial":{"type":"boolean","description":"Substring search instead of the exact name (default false)."},"max' +
    '":{"type":"integer","description":"Max rows of a substring search (default 50)."},"instance":{"type":"integer","description":"Process id of the IDE to ' +
    'use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"required"' +
    ':["identifier"]}}' +
    ',' +
    '{"name":"extract_variable","description":"Extracts the selected expression into an ' +
    'inline variable (\"var LName := <expr>;\") right before the statement that contains it ' +
    'and replaces the expression with the name. Refused where hoisting would change behaviour ' +
    '(after a short-circuit and/or, in a loop condition, as the sole statement of a branch, ' +
    'inside a with). apply=false (the default) only shows the change.","inputSchema":{"type":' +
    '"object","properties":{"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer",' +
    '"description":"1-based line of the expression."},"expression":{"type":"string",' +
    '"description":"The text to extract; located on that line. Alternative to column + ' +
    'end_column."},"column":{"type":"integer","description":"1-based start column ' +
    '(with end_column), or where to start looking for \"expression\". Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"end_column":' +
    '{"type":"integer","description":"1-based, exclusive. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"name":{"type":"string",' +
    '"description":"Name of the new variable. Default: derived from the expression."},' +
    ApplyTokenProp + ',' + TokenProp + ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"wrap_try_finally","description":"Wraps the lines from_line..to_line in a ' +
    'try..finally block. The cleanup is inferred from the statement right before them ' +
    '(\"X := TFoo.Create\" -> X.Free, BeginUpdate -> EndUpdate, Enter/Acquire/Lock -> the ' +
    'counterpart), otherwise a TODO comment is inserted - pass \"cleanup\" to say it ' +
    'yourself. Only wrapper lines are added. apply=false (the default) only shows the ' +
    'change.","inputSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute path of the unit."},' +
    '"from_line":{"type":"integer","description":"1-based first line."},"to_line":' +
    '{"type":"integer","description":"1-based last line. Default: from_line."},"cleanup":' +
    '{"type":"string","description":"The statement for the finally block, e.g. ' +
    '\"List.Free;\"."},' + ApplyTokenProp + ',' + TokenProp + ',' + InstanceProp + '},' +
    '"required":["file","from_line"]}}' +
    ',' +
    '{"name":"move_to_unit","description":"Moves the declaration at a position (type, ' +
    'class incl. its method implementations, routine, const, var) into an EXISTING unit and ' +
    'updates the uses clauses of both units and of every unit using the symbol. Use ' +
    'move_to_new_unit when the target does not exist yet. apply=false (the default) reports ' +
    'the plan.","inputSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute path of the unit."},' +
    '"line":{"type":"integer","description":"1-based line of the identifier."},' +
    '"column":{"type":"integer","description":"1-based column. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"target_file":' +
    '{"type":"string","description":"The unit to move into (path, or a name next to ' +
    'the source)."},' + ApplyTokenProp + ',' + TokenProp + ',' + InstanceProp + '},"required":' +
    '["file","line","column","target_file"]}}' +
    ',' +
    '{"name":"remove_with","description":"Rewrites \"with X do\" statements: every ' +
    'member of the body is qualified explicitly, verified per identifier with DelphiLSP. ' +
    'Pass file (+ line for the one statement enclosing it), files or project=true. An ' +
    'occurrence the rewriter cannot handle (several targets, unresolved type, inactive ' +
    '{$IFDEF} region) is listed with the reason and left alone. apply=false (the default) ' +
    'shows before/after per statement.","inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","description":"1-based line ' +
    'inside the with-statement; without it the whole file is scanned."},"files":{"type":' +
    '"array","items":{"type":"string"},"description":"Several units instead of one (absolute paths)."},"project":{"type":"boolean","description":"Every unit of the project instead of one file (default false)."},"inline_vars":' +
    '{"type":"boolean","description":"Introduce inline variables for complex targets ' +
    '(Delphi 10.3+). Default true."},' + ApplyTokenProp + ',' + TokenProp + ',' + InstanceProp +
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
    '{"file":{"type":"string","description":"Absolute path of the unit."},' +
    '"remove_unused":{"type":"boolean","description":"Default true."},' +
    '"move_to_implementation":{"type":"boolean","description":"Default false."},' +
    '"include_unverified":{"type":"boolean","description":"Default false. Also act on ' +
    'entries of a form unit whose designer could not be asked - outside the IDE nothing ' +
    're-adds them, so a wrong removal can break the build or the form streaming."},' +
    ApplyTokenProp + ',' + TokenProp + ',' + InstanceProp + '},"required":["file"]}}' +
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
    'RSS-5463) or a symbol whose DECLARING unit the session has not analysed - only the unit you ask about is sent to DelphiLSP, so a declaration elsewhere can stay invisible to it while the identifier index knows it. Read-on' +
    'ly.","inputSchema":{"type":"object",' +
    '"properties":{"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","description":' +
    '"1-based."},"column":{"type":"integer","description":"1-based. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},' + InstanceProp +
    '},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"extract_method","description":"Extracts the lines from_line..to_line ' +
    'into a new method of the enclosing class (or a local routine): which variables ' +
    'become parameters, which become locals and which one becomes the Result is resolved ' +
    'with DelphiLSP. apply=false (the default) answers with the generated code - the ' +
    'routine, the call that replaces the block and the declaration line - since that is ' +
    'what has to be judged; the write itself is not a line diff.","inputSchema":{"type":' +
    '"object","properties":{"file":{"type":"string","description":"Absolute path of the unit."},"from_line":{"type":"integer",' +
    '"description":"1-based first line of the block."},"to_line":{"type":"integer",' +
    '"description":"1-based last line. Default: from_line."},"name":{"type":"string",' +
    '"description":"Name of the new method. Default ExtractedMethod."},' + ApplyFreshProp + ',' +
    InstanceProp + '},"required":["file","from_line"]}}' +
    ',' +
    '{"name":"extract_interface","description":"Extracts an interface from the class at ' +
    'a line: a new unit with \"IXxx = interface\" plus a GUID, the class gets the ' +
    'interface in its ancestor list and both uses clauses are updated. ' +
    'add_to_existing=true adds the members to an interface that already exists ' +
    '(\"interface_name\"). apply=false (the default) answers with the interface text ' +
    'that would be written. The answer also says whether the class can implement an ' +
    'interface at all - a TObject descendant provides no IInterface.",' +
    '"inputSchema":{"type":"object","properties":{"file":' +
    '{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","description":"1-based line inside ' +
    'or at the class declaration."},"interface_name":{"type":"string","description":' +
    '"Name of the interface. Default: the class name with T replaced by I. Required ' +
    'with add_to_existing."},"target_file":{"type":"string","description":"The new ' +
    'unit (path or a name next to the source)."},"members":{"type":"array","items":' +
    '{"type":"string"},"description":"Member names to include. Default: every public ' +
    'and published method and property."},"add_to_existing":{"type":"boolean","description":"Add the members to the interface target_file already declares instead of creating a new unit (default false)."},' +
    '"change_base_to":{"type":"string","description":"A base class to derive from ' +
    'instead, applied with the extraction. A class that does not provide IInterface ' +
    'cannot implement the extracted interface at all - the answer field ' +
    '\"iinterface\" says so and names what to pass here (usually ' +
    'TInterfacedObject), or points at add_iinterface when the ancestor must ' +
    'stay."},' +
    '"add_iinterface":{"type":"boolean","description":"Add IInterface ' +
    'support to the class ITSELF with the extraction - QueryInterface, ' +
    '_AddRef, _Release and the reference counting go into the class and its ' +
    'ancestor stays. The other half of what \"iinterface\" offers; use ' +
    'this where change_base_to would drop a base class that matters."},' +
    ApplyFreshProp + ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"add_iinterface","description":"Adds IInterface support to a class that ' +
    'does not descend from TInterfacedObject: IInterface in the ancestor list, an ' +
    'FRefCount field, QueryInterface / _AddRef / _Release and the NewInstance / ' +
    'AfterConstruction pair that mirrors TInterfacedObject''s initial-refcount trick. ' +
    'Afterwards the instance frees itself when the last interface reference drops. ' +
    'apply=false (the default) returns the code it would add.","inputSchema":{"type":' +
    '"object","properties":{"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer",' +
    '"description":"1-based line inside or at the class declaration."},' + ApplyFreshProp +
    ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"signature_check","description":"Every declaration and implementation of ' +
    'the method at a position - interface, class, implementation header - and whether ' +
    'they agree. apply=true aligns the diverging ones with the majority signature ' +
    '(class declarations first, then the implementations, which keep their parameter ' +
    'NAMES because the body uses them).","inputSchema":{"type":"object","properties":' +
    '{"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","description":"1-based."},' +
    '"column":{"type":"integer","description":"1-based. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},' + ApplyFreshProp + ',' +
    InstanceProp + '},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"dfm_events","description":"Checks every form of the project: an event ' +
    'assigned in the .dfm whose handler is MISSING in the unit, or whose parameter list ' +
    'does not match the event type (the cause of hard-to-find stack corruption). Each ' +
    'issue carries an id; apply=true with \"fix_ids\" generates the missing handlers / ' +
    'corrects the parameter lists. Nothing is fixed without naming ids - a generated ' +
    'empty handler shadows an inherited one. An issue with aliasOnly=true differs only ' +
    'in the NAME of a type (an alias, e.g. TcxImageIndex = TImageIndex): the handler is ' +
    'correct and needs no fix.","inputSchema":{"type":"object",' +
    '"properties":{"fix_ids":{"type":"array","items":{"type":"string"},' +
    '"description":"Ids from a previous call."},' + ApplyIdsProp + ',' + InstanceProp +
    '}}}' +
    ',' +
    '{"name":"interface_guids","description":"Every interface declaration of the ' +
    'project with its GUID, flagging DUPLICATE GUIDs (which make Supports / ' +
    'QueryInterface return the wrong object) and interfaces without one. An interface ' +
    'paired with a dispinterface on the same GUID is a type-library import and not ' +
    'flagged. Every declaration carries an id: apply=true with fix_ids gives exactly ' +
    'those a FRESH GUID, so for a duplicate the declaration you leave out is the one ' +
    'that keeps it - which of them owns the GUID is a decision the tool must not make ' +
    'for you.","inputSchema":{"type":"object","properties":' +
    '{"only_problems":{"type":"boolean","description":"Default true."},' +
    '"fix_ids":{"type":"array","items":{"type":"string"},"description":"Ids from a ' +
    'previous call - the declarations that get a new GUID. Required for apply."},' +
    ApplyIdsProp + ',' + InstanceProp + '}}}' +
    ',' +
    '{"name":"declare_constant","description":"Declares a number or a string literal as a const or a resourcestring and'+
    ' puts the name in its place. Pass literal (or column + end_column) to replace something that is there, or neither plus '+
    'value to define one from scratch - the line is then only the anchor that says which routine or class local and class mean. '+
    'The value must be a literal or an expression of literals; a name in it is refused, because it need not be in scope where the declaration goes."' +
    ',"inputSchema":{"type":"object","properties":' +
    '{"file":{"type":"string","description":"The unit."},' +
    '"line":{"type":"integer","description":"1-based line of the literal, or of the caret."},' +
    '"literal":{"type":"string","description":"The literal to replace, located on that line as code."},' +
    '"column":{"type":"integer","description":"1-based column; with end_column instead of literal. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},' +
    '"end_column":{"type":"integer","description":"1-based, end exclusive. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},' +
    '"value":{"type":"string","description":"What is declared. Defaults to the literal; pass it to declare something else (or from scratch)."},' +
    '"name":{"type":"string","description":"The constant name. Suggested from the value when omitted."},' +
    '"type":{"type":"string","description":"Pascal type; omitted = untyped, which is what a const usually wants."},' +
    '"section":{"type":"string","description":"const (default) or resourcestring - the latter only at unit level and only for a string."},' +
    '"place":{"type":"string","description":"local (default), class, interface or implementation."},' +
    '"replace_all":{"type":"boolean","description":"Also replace the other occurrences of the literal - the routine for local, the unit otherwise. Ignored without a literal."},' +
    ApplyTokenProp + ',' + TokenProp + ',' + InstanceProp + '},"required":["file","line"]}}' +
    ',' +
    '{"name":"file_encoding","description":"Project sources whose ENCODING is not the one you want,'+
    ' and the conversion for them. The default pair is the one that fixes a real defect: a unit that is valid UTF-8 '+
    'but carries NO BOM is read by the compiler as ANSI, so every non-ASCII character in it silently becomes two (measured). '+
    'Pure ASCII is never a finding unless you ask for it - its bytes are the same in every encoding - and form files and anything binary are never touched. '+
    'Every finding carries an id; apply=true with fix_ids rewrites exactly those files, keeping every character."' + ',"inputSchema":{"type":"object","properties":' +
    '{"from":{"type":"array","items":{"type":"string"},"description":"Which encodings to convert: ansi, utf8_no_bom, ascii, utf16. Default: ansi + utf8_no_bom."},' +
    '"to":{"type":"string","description":"utf8_bom (default), utf8_no_bom or ansi. A file whose text the target cannot hold is refused, never written with replacement characters."},' +
    '"fix_ids":{"type":"array","items":{"type":"string"},"description":"Ids from a previous call - the files to convert. Required for apply."},' +
    '"max":{"type":"integer","description":"Most rows to return (default 200)."},' +
    ApplyIdsProp + ',' + InstanceProp + '}}}' +
    ',' +
    '{"name":"unused_units","description":"Units the PROJECT FILE lists although nothing in the project names them any more - they still get compiled and still clutter the project tree. A unit counts as used when any other un' +
    'it uses it (units on the search path included, not only project ones) or when the project body names one of its identifiers, which is how a form reached only through Application.CreateForm is used; the project''s own uses' +
    ' clause is its unit LIST and never counts. Two kinds: unused, and unused_with_init_code - the latter is a HINT, because such a unit is often in the project exactly to register itself. READ-ONLY: it never touches the proj' +
    'ect file. Only ONE layer is reported, since two unused units that use each other keep each other alive. Packages are refused - the units a package contains are what it exists for."' +
    ',"inputSchema":{"type":"object","properties":' +
    '{"max":{"type":"integer","description":"Most rows to return (default 200). The counts are always complete."},' +
    InstanceProp + '}}}' +
    ',' +
    '{"name":"const_params","description":"String parameters of the ' +
    'project that are passed BY VALUE although the routine only reads them: ' +
    'every one of those copies the string on every call. A row is a PROPOSAL ' +
    'with its reason, never a sweep - a const string parameter aliases the ' +
    'variable the caller passes without raising its reference count, so a ' +
    'caller that releases that variable makes the parameter dangle, which no ' +
    'compiler warning catches. Refused and SAID: a body that writes or takes ' +
    'the address of the parameter, a virtual/override family, a method an ' +
    'interface declares (const is part of the signature - E2291), a ' +
    'published method, a routine whose name is used as a procedural VALUE ' +
    '(procedure(S: string) and procedure(const S: string) are incompatible - ' +
    'E2009), and anything whose two headers cannot both be read. A GROUP is ' +
    'the unit of work: one modifier serves \"A, B: string\". Every finding ' +
    'carries an id; apply=true with fix_ids writes const in front of exactly ' +
    'those groups, in BOTH headers."' +
    ',"inputSchema":{"type":"object","properties":' +
    '{"fix_ids":{"type":"array","items":{"type":"string"},"description":"Ids ' +
    'from a previous call - the parameter groups to change. Required for ' +
    'apply."},' +
    '"include_refused":{"type":"boolean","description":"Also list the ' +
    'parameters that were REFUSED, with the reason (default false). ' +
    '\"not proposed\" and \"never looked at\" must not read the same."},' +
    '"max":{"type":"integer","description":"Most rows to return (default ' +
    '200). The counts are always complete."},' +
    ApplyIdsProp + ',' + InstanceProp + '}}}' +
    ',' +
    '{"name":"line_endings","description":"Every project source (and the files ' +
    'it includes) whose line breaks are not plain CRLF - LF only, CR only, or ' +
    'MIXED, which is the one Delphi handles worst (breakpoints land on the ' +
    'wrong line). Form files and anything binary are never reported: the fix ' +
    'rewrites bytes. Every finding carries an id; apply=true with fix_ids ' +
    'rewrites exactly those files as CRLF, keeping the encoding and changing ' +
    'nothing else. A file open with unsaved changes is skipped and named."' +
    ',"inputSchema":{"type":"object","properties":' +
    '{"fix_ids":{"type":"array","items":{"type":"string"},"description":"Ids ' +
    'from a previous call - the files to rewrite. Required for apply."},' +
    '"max":{"type":"integer","description":"Most rows to return (default 200)."},' +
    ApplyIdsProp + ',' + InstanceProp + '}}}' +
    ',' +
    '{"name":"add_unit","description":"Adds a unit to the uses clause of a file (interface or i' +
    'mplementation), minimal edit, IDE buffer when the file is open. Refuses units only on the ' +
    'browsing path. apply=false (the default) only shows the change.","inputSchema":{"type":"object","properties":{"file":{"type":"string","desc' +
    'ription":"Absolute path of the unit."},"unit":{"type":"string","description":"The unit to add, spelled as it should appear in the clause."},"section":{"type":"string"' +
    ',"enum":["interface","implementation"],"description":"interface (default) or implementation."},' + ApplyTokenProp + ',' + TokenProp + ',' +
    '"instance":{"type":"integer","description":"Proces' +
    's id of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instan' +
    'ces)."}},"required":["file","unit"]}}' +
    ',' +
    '{"name":"remove_unit","description":"Removes a unit from the uses clause of a file. apply=' +
    'false (the default) only shows the change.","inpu' +
    'tSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute pa' +
    'th of the unit."},"unit":{"type":"string","description":"The unit to remove from the clause."},' + ApplyTokenProp + ',' + TokenProp + ',' +
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
    'fier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"instance":{"type":"integer","description":"Process id of the IDE to use. Normall' +
    'y omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["file","li' +
    'ne","column"]}}' +
    ',' +
    '{"name":"find_references","description":"All references to the identifier at a position: D' +
    'elphiLSP references when available, otherwise a project-wide text scan with every hit veri' +
    'fied through DelphiLSP GotoDefinition.","inputSchema":{"type":"object","properties":{"file' +
    '":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer","d' +
    'escription":"1-based line."},"column":{"type":"integer","description":"1-based column (any' +
    'where on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"instance":{"type":"integer","description":"Process id of the ' +
    'IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"re' +
    'quired":["file","line","column"]}}' +
    ',' +
    '{"name":"uses_path","description":"Shortest chain of uses entries from one project unit to' +
    ' another - the answer to \"F2047 circular unit reference\": which chain would the new uses' +
    ' entry close.","inputSchema":{"type":"object","properties":{"from_unit":{"type":"string","description":"Unit name the path starts at."},' +
    '"to_unit":{"type":"string","description":"Unit name the path leads to."},"interface_only":{"type":"boolean","description":"Only interfa' +
    'ce-section uses (default true)."},"instance":{"type":"integer","description":"Process id o' +
    'f the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."' +
    '}},"required":["from_unit","to_unit"]}}' +
    ',' +
    '{"name":"uses_cycles","description":"Circular unit references of the active project (all c' +
    'ycles, or only those through one unit) plus the uses entries whose removal breaks the most' +
    ' cycles.","inputSchema":{"type":"object","properties":{"unit":{"type":"string","description":"Only cycles running THROUGH this unit, which must be one of the project (an unknown name is an error). Without it, every cycle."},"max":{"ty' +
    'pe":"integer","description":"Max cycles in the answer (default 50)."},"instance":{"type":"integer","description":"Process id of the IDE to use. N' +
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
    'teger","description":"1-based column. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"apply":{"type":"boolean","description":"Delete it' +
    ' when deletable (default false). There is no token: apply compares the unit with the content this call analysed and refuses when it changed."},"instance":{"type":"integer","description":"Process id ' +
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
    ' the declaration)."},"column":{"type":"integer","description":"1-based column. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"params":' +
    '{"type":"array","description":"The NEW parameter list in order. Each item: name, type, mod' +
    'ifier (const/var/out/constref), default, from (the OLD parameter it replaces - omit for a ' +
    'new parameter; an old parameter keeps type/modifier/default unless given), value (new para' +
    'meters only: the argument existing calls get).","items":{"type":"object","properties":{"na' +
    'me":{"type":"string"},"type":{"type":"string"},"modifier":{"type":"string"},"default":{"ty' +
    'pe":"string"},"from":{"type":"string"},"value":{"type":"string"}},"required":["name"]}},"a' +
    'pply":{"type":"boolean","description":"Apply the plan when it has no errors (default false' +
    '). There is no token: apply reads every touched file again and compares it with what was analysed - all of them or none."},"instance":{"type":"integer","description":"Process id of the IDE to use. Normally om' +
    'itted - the IDE is chosen automatically (see ide_instances)."}},"required":["file","line",' +
    '"column"]}}' +
    ',' +
    '{"name":"edit_method","description":"Edit the methods of the class at a position: move mem' +
    'bers into ANOTHER CLASS (in this unit or in another one) and change their modifiers. The d' +
    'eclaration leaves the old class and enters the chosen section of the new one, the body tra' +
    'vels along and its header is requalified, the uses clauses of both units follow (the direc' +
    'tion that could close an interface cycle is reported, not risked). What the tool deliberat' +
    'ely does NOT do is reported instead: fields and methods the moved body uses stay in the ol' +
    'd class, and the CALL SITES are not rewritten - they are listed. A member that cannot trav' +
    'el alone (an overload, a virtual / override / abstract / message member, a published one) ' +
    'is refused with the reason. Without members/target_class it only reports the class, its me' +
    'mbers with movability and the occurrences. A parameter list is changed by change_signature' +
    ', which verifies and rewrites every call. It edits the ROUTINES a class declares - procedure, function, constructor, destructor. A field or a property is never listed and never moves, so a method that uses one leaves it ' +
    'behind; the answer names what stays.","inputSchema":{"type":"object","properties":{"f' +
    'ile":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"integer"' +
    ',"description":"1-based line inside the class, or on one of its members or method bodies."' +
    '},"column":{"type":"integer","description":"1-based column (optional; it only preselects t' +
    'he member at that position). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"members":{"type":"array","items":{"type":"string"},"descri' +
    'ption":"Names of the members to edit."},"target_class":{"type":"string","description":"The' +
    ' class the members move into. Omit to leave them where they are."},"target_unit":{"type":"' +
    'string","description":"Unit name or path of target_class when it is not in this unit."},"s' +
    'ection":{"type":"string","description":"private / protected / public / published - where t' +
    'he declaration lands (default private)."},"add_modifiers":{"type":"array","items":{"type":' +
    '"string"},"description":"Directives to add to every selected declaration (virtual, overloa' +
    'd, inline, static, reintroduce, override)."},"remove_modifiers":{"type":"array","items":{"' +
    'type":"string"},"description":"Directives to take out of every selected declaration."},"ap' +
    'ply":{"type":"boolean","description":"Write the plan (default false). There is no token: apply compares the source with the analysed state and re-runs the plan on the text that is there now."},"instance":{"type"' +
    ':"integer","description":"Process id of the IDE to use. Normally omitted - the IDE is chos' +
    'en automatically (see ide_instances)."}},"required":["file","line"]}}' +
    ',' +
    '{"name":"semantic_replace","description":"Semantic replace with the rules of <project root' +
    '>\\semantic-replace.json (find/replace of dotted identifier paths, comment- and string-awa' +
    're, uses units added, optional local-var hoisting). Every match is VERIFIED through Delphi' +
    'LSP (the declaration of the last identifier): with a rule''s declaredIn it must be declared' +
    ' in that unit, otherwise all matches of a rule must lead to the same declaration (the most' +
    ' frequent one). Matches of another symbol are never replaced; matches DelphiLSP cannot res' +
    'olve (inactive IFDEF branch) only with include_unverified. Without apply it only reports. THE RULES FILE, which nothing else ' +
    'describes: {\"rules\":[{\"find\":\"Manager.Config.WriteConfig\",\"replace\":\"TAppCentral.Get<IConfig>.WriteConfig\",' +
    '\"uses\":[\"AppCentral.Core\"],\"declaredIn\":\"AppCentral.Config\",\"localVar\":{\"name\":\"LConfig\",' +
    '\"type\":\"IConfig\",\"value\":\"TAppCentral.Get<IConfig>\",\"replace\":\"LConfig.WriteConfig\"}}]}. ' +
    'find and replace are dotted identifier paths, NOT regular expressions; the match is case-insensitive and never inside a comment ' +
    'or a string, and a rule without find is dropped. uses must be an ARRAY of unit names - they are added to every file that really gets ' +
    'an edit. declaredIn (optional) pins the unit the matched symbol must be DECLARED in. localVar (optional) fires only when a rule ' +
    'matches more than once inside ONE routine body: then \"var <name>: <type> := <value>;\" is hoisted right after its begin ' +
    'and localVar.replace is used in that routine instead of replace. The file is plain JSON, not a Delphi source, so write it with an ' +
    'ordinary file tool."' +
    ',"inputSchema":{"type":"object","properties":{"files":{"type":"array","items":{"type":"str' +
    'ing"},"description":"Units to process (default: all project units)."},"apply":{"type":"boo' +
    'lean","description":"Replace (default false = report). There is no token and no earlier preview to carry over: every call re-runs the scan and the DelphiLSP verification, and apply only decides whether the result is writ' +
    'ten."},"include_unverified":{"type":"boo' +
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
    'd column. Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"new_unit":{"type":"string","description":"Name of the new unit (e.g. \"Custom' +
    'er.List\") or a full path."},' + ApplyTokenProp + ',' + TokenProp + ',' +
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
    'ring","description":"Absolute path of the unit."},"files":{"type":"array","items":{"type":"string"},"description":"Several units (absolute paths)."},"directory":{"type":"string","description":"Every .pas, .dpr and .dpk b' +
    'elow this directory, recursively."},"p' +
    'roject":{"type":"boolean","description":"Every unit of the project (default false). Pass exactly one of file, files, directory or project."},' + ApplyFreshProp + ',' +
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
    '","description":"Absolute path of the unit."},"from_line":{"type":"integer","description":"1-based."},"to_line":{"type":"integer","description":"Last line of the range (default: from_line)."},"d' +
    'irection":{"type":"string","enum":["to_accessors","to_fields"],"description":"to_accessors (default) or to_fields."},"getter":{"type":"boolean"' +
    ',"description":"Create a getter (default true)."},"setter":{"type":"boolean","description":"Create a setter (default true)."},"apply":{"type":"boolean","description":"Default false: nothing is written and the answer show' +
    's the plan. true makes the change. There is no token: apply re-runs the plan on the file as it is now, in the same call."},"instance":{"type":"integer","des' +
    'cription":"Process id of the IDE to use. Normally omitted - the IDE is chosen automaticall' +
    'y (see ide_instances)."}},"required":["file","from_line"]}}' +
    ',' +
    '{"name":"debug_consistency","description":"Checks the active project for reasons why break' +
    'points are not hit or debug info does not match: duplicate sources on the search path, str' +
    'ay DCUs, LF line endings, debug options and directives, outdated executable / symbol files' +
    ', host application, duplicate DLL/BPL copies - and the one that makes ' +
    'DelphiLSP itself answer nothing: the DCU output its .delphilsp.json ' +
    'names not matching the sources. MEASURED: with that directory COMPLETE, ' +
    'go to definition across units answers with the source position (a unit ' +
    'loaded from a .dcu answers fine); with one unit missing from it, that ' +
    'unit comes from another build and the answer is empty, and the session ' +
    'log says \"definition found in file:///(-1:0)\". So the check reports ' +
    'what does not match - units missing there, a .dcu older than its own ' +
    'source - and the cure is a BUILD of that configuration. Returns counts ' +
    'per kind and at most \"max\" rows (default 60, problems first), each ' +
    'with an id. Only a stale .dcu whose source the compiler can reach ' +
    'instead can be fixed from here: apply=true with its id MOVES it into a ' +
    'subdirectory of its own folder (never deletes it); run lsp_restart ' +
    'afterwards.","inputSchema":{"type":"object","properties"' +
    ':{"max":{"type":"integer","description":"Rows to list, 0 = all."},' +
    '"fix_ids":{"type":"array","items":{"type":"string"},"description":"Ids from a previous call - the stale .dcu files to move aside. Only a finding the answer marks as fixable has that cure; for anything else the call says ' +
    'so and changes nothing. Required for apply."},' +
    ApplyIdsProp + ',"instance":{"type":"integer","description":"Process id of the IDE to use. Normally omitt' +
    'ed - the IDE is chosen automatically (see ide_instances)."}}}}' +
    ',' +
    '{"name":"blame","description":"git/svn blame of a file (revision, author, date, summary pe' +
    'r line) - the file on disk; at most 500 lines per call.","inputSchema":{"type":"object","p' +
    'roperties":{"file":{"type":"string","description":"Absolute path of the unit."},"start_lin' +
    'e":{"type":"integer","description":"First line (default 1)."},"end_line":{"type":"integer","description":"Last line; 0, the default, means to the end of the file. At most 500 lines come back per call."},"instance":{"type":"integer","descript' +
    'ion":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatically (se' +
    'e ide_instances)."}},"required":["file"]}}' +
    ',' +
    '{"name":"commit_info","description":"The commit that last changed a line: metadata, messag' +
    'e, changed files; diff_file adds the diff of the files whose path contains that text.","in' +
    'putSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute ' +
    'path of the unit."},"line":{"type":"integer","description":"1-based line."},"diff_file":{"' +
    'type":"string","description":"Also return the diff of the changed file whose path CONTAINS this text. Without it the answer lists the files of the commit and no diff."},"instance":{"type":"integer","description":"Process id of the IDE to use. ' +
    'Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["fi' +
    'le","line"]}}' +
    ',' +
    '{"name":"rename_preview","description":"Previews renaming the identifier at a position wit' +
    'h the plugin''s rename pipeline (text scan + DelphiLSP verification, interface implementati' +
    'ons, .dfm/.fmx form files). SAVES ALL MODIFIED FILES FIRST, like the rename dialog. Return' +
    's the changes and a token for rename_apply.","inputSchema":{"type":"object","properties":{' +
    '"file":{"type":"string","description":"Absolute path of the unit."},"line":{"type":"intege' +
    'r","description":"1-based line."},"column":{"type":"integer","description":"1-based column' +
    ' (anywhere on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"new_name":{"type":"string","description":"The new name."},"scope":{"type":"string","en' +
    'um":["project","unit","method"],"description":"How far the scan reaches: project (the default), unit or method. Anything else is read as project."},"include_open_units":{"type":"boolean","description":"Besides the projec' +
    't files, the units open in the editor (default: the setting on the options page)."},"include_used_uni' +
    'ts":{"type":"boolean","description":"Besides the project files, the units reachable through uses clauses (default: the setting on the options page)."},"instance":{"type":"integer","description":"Process id of the IDE t' +
    'o use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},"require' +
    'd":["file","line","column","new_name"]}}' +
    ',' +
    '{"name":"rename_apply","description":"Applies the rename previewed by rename_preview (toke' +
    'n) through the IDE editor - undoable, not saved.","inputSchema":{"type":"object","properti' +
    'es":{"token":{"type":"string","description":"The token of the rename_preview this call applies."},"instance":{"type":"integer","description":"Process id of t' +
    'he IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instances)."}},' +
    '"required":["token"]}}' +
    ',' +
    '{"name":"lsp_hover","description":"DelphiLSP hover text at a position (type / declaration ' +
    'info). The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its analysis before asking. What it cannot do is make DelphiLSP reso' +
    'lve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema":{"type":"object","properties":{"file":{"type":"string","descri' +
    'ption"' +
    ':"Absolute path of the unit."},"line":{"type":"integer","description":"1-based line."},"co' +
    'lumn":{"type":"integer","description":"1-based column (anywhere on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"ins' +
    'tance":{"type":"integer","description":"Process id of the IDE to use. Normally omitted - t' +
    'he IDE is chosen automatically (see ide_instances)."}},"required":["file","line","column"]' +
    '}}' +
    ',' +
    '{"name":"lsp_definition","description":"DelphiLSP go-to-definition for a position. The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and wa' +
    'its for its analysis before asking. What it cannot do is make DelphiLSP resolve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.",' +
    '"input' +
    'Schema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute pat' +
    'h of the unit."},"line":{"type":"integer","description":"1-based line."},"column":{"type":' +
    '"integer","description":"1-based column (anywhere on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"instance":{"type"' +
    ':"integer","description":"Process id of the IDE to use. Normally omitted - the IDE is chos' +
    'en automatically (see ide_instances)."}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"lsp_implementation","description":"DelphiLSP go-to-implementation for a position.' +
    ' The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its analysis before asking. What it cannot do is make DelphiLSP resolve a ' +
    'unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema":{"type":"object","properties":{"file":{"type":"string","description"' +
    ':"Abso' +
    'lute path of the unit."},"line":{"type":"integer","description":"1-based line."},"column":' +
    '{"type":"integer","description":"1-based column (anywhere on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"instance"' +
    ':{"type":"integer","description":"Process id of the IDE to use. Normally omitted - the IDE' +
    ' is chosen automatically (see ide_instances)."}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"lsp_references","description":"DelphiLSP textDocument/references for a position (' +
    'when the server supports it). The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its analysis before asking. What it cannot do' +
    ' is make DelphiLSP resolve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema":{"type":"object","properties":{"file":{' +
    '"type"' +
    ':"string","description":"Absolute path of the unit."},"line":{"type":"integer","descriptio' +
    'n":"1-based line."},"column":{"type":"integer","description":"1-based column (anywhere on ' +
    'the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"include_declaration":{"type":"boolean","description":"Include the declaration itself (default true)."},"instance":{"type":"integer","' +
    'description":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatic' +
    'ally (see ide_instances)."}},"required":["file","line","column"]}}' +
    ',' +
    '{"name":"lsp_document_symbols","description":"DelphiLSP document symbols of a unit (raw LS' +
    'P JSON, 0-based ranges). The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its analysis before asking. What it cannot do is m' +
    'ake DelphiLSP resolve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema":{"type":"object","properties":{"file":{"type' +
    '":"str' +
    'ing","description":"Absolute path of the unit."},"instance":{"type":"integer","description' +
    '":"Process id of the IDE to use. Normally omitted - the IDE is chosen automatically (see i' +
    'de_instances)."}},"required":["file"]}}' +
    ',' +
    '{"name":"lsp_signature_help","description":"DelphiLSP signature help inside a call (raw LS' +
    'P JSON). The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its analysis before asking. What it cannot do is make DelphiLSP re' +
    'solve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema":{"type":"object","properties":{"file":{"type":"string","desc' +
    'riptio' +
    'n":"Absolute path of the unit."},"line":{"type":"integer","description":"1-based line."},"' +
    'column":{"type":"integer","description":"1-based column (anywhere on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"i' +
    'nstance":{"type":"integer","description":"Process id of the IDE to use. Normally omitted -' +
    ' the IDE is chosen automatically (see ide_instances)."}},"required":["file","line","column' +
    '"]}}' +
    ',' +
    '{"name":"lsp_completion","description":"DelphiLSP completion items at a position (label, d' +
    'etail, kind); prefix filters, max limits (default 100). The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its analysis before' +
    ' asking. What it cannot do is make DelphiLSP resolve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema":{"type":"obje' +
    'ct","p' +
    'roperties":{"file":{"type":"string","description":"Absolute path of the unit."},"line":{"t' +
    'ype":"integer","description":"1-based line."},"column":{"type":"integer","description":"1-' +
    'based column (anywhere on the identifier). Characters of the line as stored - a tab counts as ONE, not as its width on screen."},"prefix":{"type":"string","description":"Only items whose label starts with this (case-insensitive)."},"max":{"type":"int' +
    'eger","description":"Max items (default 100)."},"instance":{"type":"integer","description":"Process id of the IDE to use. Normally ' +
    'omitted - the IDE is chosen automatically (see ide_instances)."}},"required":["file","line' +
    '","column"]}}' +
    ',' +
    '{"name":"lsp_request","description":"Sends any request to the plugin''s DelphiLSP session a' +
    'nd returns the raw result (LSP conventions: 0-based positions, file URIs). \"file\" syncs ' +
    'that document first. Lifecycle and document-sync methods are refused. The unit you ASK ABOUT need not be OPEN in the IDE: this reads it (editor buffer first, else disk), sends that text to the session and waits for its a' +
    'nalysis before asking. What it cannot do is make DelphiLSP resolve a unit its own compile does not reach - the answer is empty then, and find_original_symbol is the fallback that asks the identifier index.","inputSchema"' +
    ':{"typ' +
    'e":"object","properties":{"method":{"type":"string","description":"The LSP method, e.g. textDocument/documentHighlight. Lifecycle, did* and $/ methods are refused."},"params":{"type":"object","description":"The params of' +
    ' the request as JSON - sent as they are."},"file":{"t' +
    'ype":"string","description":"Optional: sync this file''s current content before the request' +
    '."},"timeout_ms":{"type":"integer","description":"How long to wait for the answer (default 20000)."},"instance":{"type":"integer","description":"Process id' +
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

function LspEmptyAnswerNote(const AEmptySentence, ADescription,
  AFile: string; AAgoMs: Integer;
  AUnitAnalysed, ASuggestIndex, AAboutSymbol: Boolean): string;

  procedure Add(const AText: string);
  begin
    if Result <> '' then Result := Result + ' ';
    Result := Result + AText;
  end;

begin
  Result := '';
  // WHAT AN EMPTY ANSWER MEANS comes first - in the tool's own words when
  // it has them, and generically otherwise. It is deliberately not the
  // same parameter as the DESCRIPTION of the answer: passing a description
  // there used to replace this sentence, so the one thing the note exists
  // to say was the one thing missing (issue #50.5).
  if AEmptySentence <> '' then
    Add(AEmptySentence)
  else
    Add('DelphiLSP returned nothing.');
  if ADescription <> '' then Add(ADescription);

  if (AAgoMs >= 0) and (AAgoMs < 60000) then
    Add(Format('%s was (re)sent to DelphiLSP %d s ago because its content ' +
      'changed, so it may still be analysing it: retry shortly.',
      [ExtractFileName(AFile), AAgoMs div 1000]));

  if AUnitAnalysed and not AAboutSymbol then
    // document symbols, signature help and completion do not resolve a
    // symbol at all: the paragraph below would be a wrong lead there.
    Add('The session HAS analysed this unit, so this is the server''s own '
      + 'answer here and not a timing problem.')
  else if AUnitAnalysed then
  begin
    // The unit we ASKED about is fine, so the silence is about the symbol -
    // and there are exactly two causes, which only the log tells apart.
    // MEASURED 2026-10-20: the second one produced this with hover
    // working all along, so naming only the first would be a wrong lead
    // (and naming only the .dcu, as before 2.1.1, was the other one).
    Add('The session HAS analysed this unit, so this is about the symbol, ' +
      'not about reaching the file. Two causes look identical from here: ' +
      'DelphiLSP does not resolve this symbol (a private / public overload ' +
      'pair is the known case - Delphi 13.1, RSS-5463), or it could not ' +
      'give the DECLARING unit a source position. MEASURED for the second ' +
      'one: it happens when the DCU output directory of the project''s ' +
      'configuration does not match the sources - a unit missing there ' +
      'comes from another build and loses its position, while a COMPLETE ' +
      'directory answers correctly (with or without debug information). ' +
      'get_status says whether it is complete, debug_consistency names ' +
      'what is missing, and a build of that configuration is the cure. ' +
      'The session log shows this case as "definition found in ' +
      'file:///(-1:0)": switch "Log this plugin''s DelphiLSP session" on ' +
      '(settings LspLogging, next session) and read ' +
      '%TEMP%\DelphiLSP\RefactoringLight*.log.');
    if ASuggestIndex then
      Add('find_original_symbol answers both cases from the identifier ' +
        'index - for a declaration that is usually the answer you want.');
  end
  else
  begin
    Add('The session has analysed nothing of this unit yet. While that ' +
      'lasts it answers nothing for it: get_status says whether it is ' +
      'still loading the project. When it persists, the usual cause is ' +
      'that the unit - or one it uses - is taken from a precompiled .dcu ' +
      'instead of its source: DCUs of the project''s own units on the IDE ' +
      'library path or in the DCU output directory its .delphilsp.json ' +
      'names. To see it, switch "Log this plugin''s DelphiLSP session" on ' +
      '(settings LspLogging, next session) and read %TEMP%\DelphiLSP\' +
      'RefactoringLight*.log - a unit taken from a .dcu answers ' +
      '"definition found in file:///(-1:0)".');
    if ASuggestIndex then
      Add('find_original_symbol does not depend on the session for its ' +
        'last step, so it may answer anyway.');
  end;
end;

function McpGuideSections: TArray<TMcpGuideSection>;

  function Sec(const ATitle: string; const ATools: TArray<string>;
    const ALines: TArray<string>): TMcpGuideSection;
  begin
    Result.Title := ATitle;
    Result.Tools := ATools;
    Result.Text := string.Join(#10, ALines);
  end;

begin
  // One section per JOB, and every tool of McpToolDefinitions is in
  // exactly one of them - TheGuideAndTheToolListAgree holds all three
  // directions, so a tool added, renamed or dropped cannot leave this
  // text silently stale.
  Result := [
    Sec('What this server is',
      [],
      [
        'Delphi / RAD Studio refactoring through a RUNNING IDE that has the',
        'Refactoring Light plugin loaded. The answers come from DelphiLSP and',
        'the IDE itself, not from text matching: a rename verifies every',
        'occurrence, a check reads the project model. Everything works on the',
        'editor buffers, unsaved changes included.'
      ]),
    Sec('Three rules that save a round',
      [],
      [
        '1. THE IDE BUFFER IS THE TRUTH. While a unit is open in the IDE, read',
        '   and write it with buffer_read / buffer_edit / buffer_save. A file',
        '   edited on disk behind the IDE desyncs its buffer, and the next',
        '   Ctrl+S writes the old text back over your change.',
        '2. CALL A WRITING TOOL WITH apply=false FIRST. Every one of them then',
        '   answers "changes" (before/after per line) plus what it refuses and',
        '   why, so the edit can be judged before it happens.',
        '3. AN ID OR TOKEN BELONGS TO ONE ANSWER. A quick-fix id, a preview',
        '   token and a check''s fix_ids expire with the next edit or the next',
        '   scan; a stale one is refused by name rather than applied to',
        '   something else. List again instead of reusing one.'
      ]),
    Sec('Where to start',
      ['guide', 'get_status', 'settings', 'ide_instances',
       'select_ide'],
      [
        'get_status is the first call when something does not work: it says',
        'whether the identifier index is ready, whether DelphiLSP answers,',
        'what the live checker found and whether the project has a',
        '.delphilsp.json at all. Most "no result" answers have their reason',
        'there.',
        'ide_instances lists the running IDEs; select_ide pins one for this',
        'session. Normally neither is needed - the target follows the file',
        'argument and the working directory - but with two IDEs open an',
        'ambiguous call is refused rather than guessed.',
        'settings lists the plugin''s own options with their current ',
        'values and what each one does, and sets one: permanently in the',
        'registry or for the RUNNING IDE only. It is the options page of',
        'the IDE as data - nothing else here can change them.',
        'guide is this text.'
      ]),
    Sec('Reading and writing a unit',
      ['buffer_open', 'buffer_read', 'buffer_edit', 'buffer_list', 
       'buffer_save', 'buffer_close', 'scratch_analyze', 'scratch_close'],
      [
        'buffer_read gives the current text (buffer when open, else disk) with',
        'a revision; buffer_edit replaces an exact, unique piece of text and',
        'writes only the lines that really change, so the IDE marks just that',
        'spot and one Ctrl+Z undoes it. buffer_save writes a loaded buffer to',
        'disk like Ctrl+S - for a FORM unit it also writes the .dfm and may',
        'extend the unit''s uses clause, and the answer names those files.',
        'buffer_open without visible=true loads a file HEADLESS (memory only,',
        'no editor tab): useful to check an edit with get_diagnostics before',
        'buffer_close saves or discards it. buffer_list says which files are',
        'loaded and which of them the user can see.',
        'scratch_analyze hands DelphiLSP a unit that exists only in memory and',
        'returns its diagnostics and quick fixes - the way to try something',
        'without touching the project; scratch_close drops it again.'
      ]),
    Sec('The project itself',
      ['project_open', 'project_close', 'project_files'],
      [
        'project_open opens a project or project group, project_close closes',
        'the group and every open file (the IDE has no "close only the',
        'project"). Both REFUSE while anything is unsaved and name the files:',
        'the IDE would ask, and nobody can answer that through a pipe - pass',
        'save_modified=true to save them first.',
        'project_files lists what the project carries and adds or removes',
        'entries. Adding does not create or open a file, removing does not',
        'delete one, and the project file is written unless you say',
        'save=false. After a switch, get_status says when the index and the',
        'LSP session have caught up.'
      ]),
    Sec('Errors and quick fixes',
      ['get_diagnostics', 'get_quick_fixes', 'apply_quick_fix', 
       'apply_quick_fixes'],
      [
        'get_diagnostics returns the compiler diagnostics for a file, merged',
        'from the sources that are fresh for its current content (the IDE''s',
        'Error Insight, the Messages window and the plugin''s own DelphiLSP',
        'session), with the revision they describe.',
        'get_quick_fixes turns them into offers: each carries its diagnostic,',
        'the affected line and the real edit as "changes", so a fix can be',
        'judged without applying it. apply_quick_fix applies one by id,',
        'apply_quick_fixes a whole kind or a list of ids in one pass',
        '(bottom-up, and a fix whose anchor text has moved is skipped and',
        'reported rather than applied elsewhere).',
        'Two fixes are deliberately refused through this server: adding a unit',
        'that only the browsing path can see, and removing a private member',
        'with a non-empty body - both need a human decision.'
      ]),
    Sec('Finding where something is used',
      ['find_references', 'find_implementations', 'find_original_symbol', 
       'find_unit', 'find_unit_references'],
      [
        'find_references is the verified search: a text scan over the project',
        'scope, then DelphiLSP for every candidate, plus what the sources',
        'themselves can decide (the type of a qualifier, the argument count of',
        'an overload). Each row carries its kind (declaration, call, write,',
        'read, ...) and, where it applies, that it was reached through an',
        'interface. An occurrence nobody could verify is LISTED and marked,',
        'never silently dropped.',
        'find_implementations answers for a method, an interface or a type',
        'header (every descendant), find_original_symbol where an identifier',
        'is DECLARED - including the cases lsp_definition stays silent about,',
        'because it falls back to the identifier index.',
        'find_unit answers which unit declares an identifier (that is the',
        'index, not the compiler), find_unit_references which units use a',
        'given unit - and which list it in their uses clause without using',
        'anything of it.'
      ]),
    Sec('Asking DelphiLSP directly',
      ['lsp_hover', 'lsp_definition', 'lsp_implementation', 
       'lsp_references', 'lsp_document_symbols', 'lsp_signature_help', 
       'lsp_completion', 'lsp_request', 'lsp_restart'],
      [
        'The raw server answers for ONE position: lsp_hover, lsp_definition,',
        'lsp_implementation, lsp_references, lsp_document_symbols,',
        'lsp_signature_help, lsp_completion - and lsp_request for any other',
        'method. The unit you ask about need not be open: the tool reads it,',
        'sends that text to the session and waits for its analysis.',
        'Use them to look at ONE position. For a question about a whole',
        'project, the find_* tools are more complete: they add the text scan,',
        'the interface links and the source resolution that the server alone',
        'does not give. textDocument/references is not even announced by',
        'DelphiLSP, so lsp_references usually answers nothing - that is why',
        'find_references verifies candidate by candidate.',
        'lsp_restart ends the session so the next request starts a fresh',
        'one - the way to pick up a changed .delphilsp.json or to get rid',
        'of .dcu files the session loaded instead of the sources. Closing',
        'the project does not restart it.'
      ]),
    Sec('Renaming, signatures, deleting',
      ['rename_preview', 'rename_apply', 'change_signature', 
       'safe_delete', 'edit_method'],
      [
        'rename_preview verifies every occurrence first and answers them with',
        'a token plus a warning when the declaration lies outside the chosen',
        'scope; rename_apply then writes exactly that preview. Form files are',
        'part of it: a renamed component or event handler is changed in the',
        '.dfm as well, and a form open in the designer is renamed THROUGH the',
        'designer.',
        'change_signature rewrites a routine''s parameter list together with',
        'the whole family (interface, overrides) and its call sites, and it',
        'refuses rather than guesses when a parameter the body still uses',
        'would go.',
        'safe_delete says whether a symbol can go: it lists every use it',
        'found, counts anything it could not verify as a use, and vetoes what',
        'is reachable without a textual call (virtual, published, an',
        'interface member).',
        'edit_method edits what a class declares - the ROUTINES only, never a',
        'field or a property: signature, modifiers, and the class or unit a',
        'method lives in. It does the mechanical half and NAMES the rest (the',
        'members of the old class the body still uses, the call sites), which',
        'is why it verifies those sites before the edit and lists them',
        'afterwards.'
      ]),
    Sec('Taking code apart and moving it',
      ['extract_method', 'extract_variable', 'extract_interface', 
       'add_iinterface', 'declare_constant', 'wrap_try_finally', 
       'convert_properties', 'move_to_unit', 'move_to_new_unit', 
       'remove_with', 'semantic_replace', 'expand_includes'],
      [
        'extract_method takes a line range into a new method (parameters, var',
        'and const decided from what the block reads and WRITES, Result',
        'handled as the function''s return value); extract_variable hoists an',
        'expression into a local; wrap_try_finally wraps a statement and',
        'guesses the cleanup call (Create -> Free, BeginUpdate -> EndUpdate).',
        'declare_constant declares a selected literal - or one you type - as a',
        'const or resourcestring, locally, in the class or at unit level, and',
        'can replace every other occurrence of that value.',
        'extract_interface builds an interface from a class (and adds it to an',
        'existing unit when the target is there); it also reports whether the',
        'class can implement an interface at all, and offers either the',
        'ancestor change or add_iinterface, which implements QueryInterface /',
        '_AddRef / _Release by delegation.',
        'move_to_unit and move_to_new_unit move a declaration with its body,',
        'take the units the moved code needs along and refuse an interface',
        'cycle. convert_properties turns field access into property accessors',
        'and back. remove_with rewrites a with-statement into qualified',
        'access, and refuses the cases where that would change behaviour (a',
        'record target that would become a copy, for one).',
        'semantic_replace replaces dotted identifier paths by the rules of',
        'semantic-replace.json, with every match verified against the symbol''s',
        'declaration. expand_includes writes {$I} files into the unit, marked',
        '- for reading, not as a refactoring.'
      ]),
    Sec('uses clauses',
      ['add_unit', 'remove_unit', 'analyze_uses', 'cleanup_uses', 
       'uses_path', 'uses_cycles'],
      [
        'add_unit / remove_unit edit one entry (interface or implementation,',
        'comment- and directive-aware). analyze_uses judges every entry of a',
        'file: used, unused, movable to the implementation, kept because the',
        'form DESIGNER writes it itself, or not analysable and why.',
        'cleanup_uses does what that verdict allows. For a FORM unit the',
        'designer is asked, and an entry it reported is never touched - the',
        'IDE would write it straight back.',
        'uses_path answers the shortest chain of uses entries from one unit to',
        'another - the question a circular-reference error poses -, uses_cycles',
        'the cycles of the project.'
      ]),
    Sec('Project-wide checks',
      ['signature_check', 'dfm_events', 'interface_guids', 
       'debug_consistency', 'unused_units', 'line_endings', 'const_params', 
       'file_encoding'],
      [
        'Each of these scans the project and reports findings with an id;',
        'those that can fix something do so only for the ids you NAME, because',
        'which of several places is the right one is not theirs to decide.',
        'signature_check finds declarations and implementations that disagree',
        'and can align them. dfm_events finds event handlers a form expects',
        'and the unit does not declare (or declares with another parameter',
        'list). interface_guids finds duplicate and missing GUIDs - a',
        'duplicate makes Supports/QueryInterface hand out the wrong object -',
        'and gives the declarations you name a fresh one.',
        'debug_consistency looks for the reasons a debugger stops at the wrong',
        'line: a second copy of a unit on the search path, a stray .dcu, mixed',
        'line endings, {$D-} and friends, a stale executable.',
        'unused_units lists units the project FILE carries although nothing',
        'names them any more (read-only). line_endings and file_encoding find',
        'files Delphi reads wrongly - mixed or LF-only breaks, and above all',
        'UTF-8 without a BOM, which the compiler reads as ANSI so one umlaut',
        'becomes two characters - and convert the ones you name.',
        'const_params finds string parameters passed by value although the',
        'routine only reads them: every call copies the string. It proposes,',
        'with a reason per row, and writes const only for the ids you name -',
        'a const string parameter aliases the caller''s variable, which is a',
        'decision no sweep may take.'
      ]),
    Sec('Version control',
      ['blame', 'commit_info'],
      [
        'blame answers per line who last changed it (git and svn), mapped onto',
        'the current buffer, so an unsaved edit does not shift the authors.',
        'commit_info gives one commit''s metadata, its changed files and the',
        'diff per file.'
      ]),
    Sec('What this server does not do',
      [],
      [
        'It has no parser of its own for name resolution: whether an',
        'occurrence belongs to your symbol is DelphiLSP''s answer, plus what',
        'the sources can decide without it. So when the server is silent for a',
        'unit - typically because that unit, or one it uses, is taken from a',
        'precompiled .dcu instead of its source - a verified search finds',
        'less, and it SAYS so instead of answering confidently.',
        'It never changes code without apply=true, it refuses a write while a',
        'modal dialog is open in the IDE, and it does not close an editor tab',
        'the user can see.'
      ])
  ];
end;

function McpGuideText: string;
var
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    SB.Append('REFACTORING LIGHT - what this server can do');
    SB.Append(#10).Append('Plugin / bridge version: ').Append(PluginVersion);
    for var S in McpGuideSections do
    begin
      SB.Append(#10#10).Append('== ').Append(S.Title);
      SB.Append(#10).Append(S.Text);
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
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
    'While no IDE runs only guide, ide_instances and select_ide exist; ' +
    'the other ' +
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
