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
//   request : {"method":"context"}
//             {"method":"call","tool":"get_quick_fixes","arguments":{...}}
//   response: {"ok":true,"result":{...}}  /  {"ok":false,"error":"..."}

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

implementation

uses
  System.Classes, System.StrUtils, System.IOUtils;

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
    'remove unused variables, ...). Each fix has an id for apply_quick_fix; ' +
    'ids are bound to the buffer revision and expire with the next edit.",' +
    '"inputSchema":{"type":"object","properties":{' + FileProp + ',' +
    '"line":{"type":"integer","description":"Only fixes anchored to this ' +
    '1-based line."},' + RefreshProp + ',' + InstanceProp + '}}},' +

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
    '{"name":"add_unit","description":"Adds a unit to the uses clause of a file (interface or i' +
    'mplementation), minimal edit, IDE buffer when the file is open. Refuses units only on the ' +
    'browsing path.","inputSchema":{"type":"object","properties":{"file":{"type":"string","desc' +
    'ription":"Absolute path of the unit."},"unit":{"type":"string"},"section":{"type":"string"' +
    ',"enum":["interface","implementation"]},"instance":{"type":"integer","description":"Proces' +
    's id of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_instan' +
    'ces)."}},"required":["file","unit"]}}' +
    ',' +
    '{"name":"remove_unit","description":"Removes a unit from the uses clause of a file.","inpu' +
    'tSchema":{"type":"object","properties":{"file":{"type":"string","description":"Absolute pa' +
    'th of the unit."},"unit":{"type":"string"},"instance":{"type":"integer","description":"Pro' +
    'cess id of the IDE to use. Normally omitted - the IDE is chosen automatically (see ide_ins' +
    'tances)."}},"required":["file","unit"]}}' +
    ',' +
    '{"name":"analyze_uses","description":"Uses-clause analysis of one unit: every entry with v' +
    'erdict used / unused / movable (only needed in the implementation) / init_code / ide_manag' +
    'ed / unknown and its usage count.","inputSchema":{"type":"object","properties":{"file":{"t' +
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
    '{"name":"expand_includes","description":"Writes the content of every {$I}/{$INCLUDE} f' +
    'ile IN PLACE into the including source (for debugging), framed by marker comments that ke' +
    'ep the original directive (// >>> include begin: ... / // <<< include end: ...). Nested in' +
    'cludes too. Files open in the IDE are changed in the editor buffer (not saved), closed fil' +
    'es ON DISK - revert with version control. Pass file, files, directory (recursive, .pas/.dp' +
    'r/.dpk) or project=true.","inputSchema":{"type":"object","properties":{"file":{"type":"st' +
    'ring"},"files":{"type":"array","items":{"type":"string"}},"directory":{"type":"string"},"p' +
    'roject":{"type":"boolean"},"instance":{"type":"integer","description":"Process id of the I' +
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

function McpServerInstructions: string;
begin
  Result :=
    'Bridge to running RAD Studio (Delphi) IDEs with the Refactoring Light ' +
    'plugin. Everything works on the IDE''s CURRENT EDITOR BUFFERS, ' +
    'including unsaved changes. ' +
    'Diagnostics: get_diagnostics merges Error Insight (Structure view), the ' +
    'messages of the last compile and the plugin''s own DelphiLSP session. ' +
    'Fixes: get_quick_fixes lists the plugin''s automatic fixes with ids, ' +
    'apply_quick_fix applies one to the editor buffer (not saved). Fix ids ' +
    'expire with every edit of the unit - after applying one, list again ' +
    'before applying the next. Prefer these fixes over editing the file on ' +
    'disk while the unit is open in the IDE. ' +
    'Buffers: buffer_read / buffer_edit work on the IDE''s editor buffers; ' +
    'buffer_save writes a loaded buffer to disk like Ctrl+S (no reload ' +
    'prompt in the IDE); prefer buffer_edit + buffer_save over editing files ' +
    'on disk while they are open in the IDE. ' +
    'buffer_open without visible=true loads a file HEADLESS (memory only), so ' +
    'edits and quick fixes can be tried and verified with get_diagnostics ' +
    'before buffer_close saves or discards them. scratch_analyze compiles a ' +
    'throw-away unit that exists only in memory against the active project - ' +
    'use it to test snippets. ' +
    'Refactoring: find_unit / add_unit / remove_unit / analyze_uses, ' +
    'find_references / find_implementations, rename_preview + rename_apply ' +
    '(the IDE plugin''s rename incl. form files), uses_path / uses_cycles, ' +
    'safe_delete (check that nothing uses a symbol, then delete it), ' +
    'change_signature (add / remove / reorder / rename parameters incl. every call), ' +
    'expand_includes (include files written into their units for debugging), ' +
    'convert_properties (field access <-> getter/setter), ' +
    'debug_consistency, blame / commit_info. The lsp_* tools talk to the ' +
    'plugin''s own DelphiLSP session directly. ' +
    'When something does not work as expected (no fixes, no diagnostics, ' +
    'index not ready), get_status shows the plugin''s own view of why. ' +
    'Several IDEs may run at once: the target is chosen automatically from ' +
    'the file argument (an IDE that has it open wins) and from this ' +
    'session''s working directory compared with each IDE''s projects. If the ' +
    'choice is ambiguous the tool says so - call ide_instances to see the ' +
    'IDEs and their projects, then select_ide or pass "instance".';
end;

end.
