(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.McpServer;

// The IDE end of the MCP bridge (IDE-only; see Mcp.Protocol for the
// picture). Serves \\.\pipe\DelphiRefactoringLight-mcp-<PID> through
// Mcp.PipeServer and implements the tools on top of the live checker's
// data.
//
// THREADING - the rules this plugin learned the hard way apply here too:
// * requests arrive on pipe HANDLER threads;
// * ToolsAPI (editor buffers included) is MAIN THREAD ONLY. Those pieces
//   run through RunOnMain, which POSTS a message to a hidden window - a
//   plain window message, NOT TThread.Queue: queued procs run inside
//   CheckSynchronize, where touching the IDE can deadlock against its
//   parser thread;
// * everything slow (waiting for DelphiLSP, resolving quick fixes against
//   the index) stays on the handler thread;
// * edits are refused while a modal dialog is open - the request answers
//   "busy" instead of landing in the middle of whatever the dialog does;
// * every wait also watches the server's stop event, so unloading the
//   package never waits for a hanging request.

interface

uses
  System.SysUtils, System.JSON;

type
  /// <summary>One MCP tool of the IDE side. Runs on a pipe HANDLER thread;
  ///  ToolsAPI only through McpRunOnMain, long waits watch AStop.</summary>
  TMcpToolHandler = reference to function(AArgs: TJSONObject; AStop: THandle): string;

/// <summary>Further tool units (Expert.McpMoreTools, Expert.McpLspTools)
///  register their handlers here from their initialization section. The
///  tool's DEFINITION belongs into Mcp.Protocol (single source).</summary>
procedure RegisterMcpTool(const AName: string; const AHandler: TMcpToolHandler);

/// <summary>Runs AProc on the IDE main thread (posted window message, never
///  TThread.Queue) and waits for it. AAllowModal=False refuses while a modal
///  dialog is open or another main-thread MCP call is still running.</summary>
function McpRunOnMain(const AProc: TProc; AAllowModal: Boolean; AStop: THandle;
  out AError: string; ATimeoutMs: Cardinal = 15000): Boolean;
function McpOk(AResult: TJSONValue): string;
function McpErr(const AMsg: string): string;
/// <summary>Editor buffer if loaded, else the file on disk. MAIN THREAD.</summary>
function McpReadContent(const AFile: string; out AContent: string): Boolean;

procedure StartMcpServer;
procedure StopMcpServer;
/// <summary>One line for the status window.</summary>
function McpServerStatus: string;
function McpServerPipe: string;
/// <summary>Which MCP bridges (Claude Code sessions) are connected: a
///  running bridge asks for this IDE's context every 5 s, so a request
///  within the last 15 s means "connected". ADetail lists them.</summary>
function McpConnectionStatus(out ADetail: string): string;

type
  /// <summary>What happened to one tool in this IDE session.</summary>
  TMcpToolStat = record
    Name: string;
    Calls, Errors, Running: Integer;
    LastTime: TDateTime;      // start of the last call, 0 = never called
    LastMs, TotalMs, MaxMs: Int64;
    LastOk: Boolean;
    LastError: string;        // message of the last FAILED call
  end;

/// <summary>Per-tool call statistics of this IDE session (thread-safe copy).
///  Tools that were never called are not in it.</summary>
function McpToolStats: TArray<TMcpToolStat>;
/// <summary>True when this IDE has a handler for the tool (the bridge's own
///  tools ide_instances / select_ide never reach an IDE).</summary>
function McpHandlesTool(const AName: string): Boolean;

implementation

uses
  Winapi.Windows, Winapi.Messages, System.Classes,
  System.IOUtils, System.Generics.Collections, System.SyncObjs,
  Vcl.Forms, ToolsAPI,
  Lsp.Protocol, Lsp.Client, Lsp.Uri, Expert.LspManager,
  Expert.EditorHelperIntf, Expert.AutoImport, Expert.DiagStore,
  Expert.McpTools, Expert.UnitIndex, Expert.UnitAvailability, Expert.UsesEditor,
  Mcp.Protocol, Mcp.PipeServer, Expert.StatusWindow, Expert.Version;

const
  WM_MCP_CALL = WM_APP + $3A1;
  MainCallTimeoutMs = 15000;
  LspWaitMs = 8000;

// ---------------------------------------------------------------------------
//  Main-thread dispatch
// ---------------------------------------------------------------------------

type
  TMainCall = class(TInterfacedObject)
  public
    Proc: TProc;
    AllowModal: Boolean;
    Done: THandle;
    Error: string;
    constructor Create(const AProc: TProc; AAllowModal: Boolean);
    destructor Destroy; override;
  end;

  TDispatcher = class
  private
    FWnd: HWND;
    procedure WndProc(var Msg: TMessage);
  public
    constructor Create;
    destructor Destroy; override;
  end;

var
  GServer: TMcpPipeServer = nil;
  GDispatcher: TDispatcher = nil;
  GWnd: HWND = 0;
  GFixLock: TCriticalSection = nil;
  GFixCache: TDictionary<string, TPair<Cardinal, TArray<TQuickFix>>> = nil;
  GLastTool: string;
  GToolRegistry: TDictionary<string, TMcpToolHandler> = nil;
  GStatLock: TCriticalSection = nil;
  GStats: TDictionary<string, TMcpToolStat> = nil;
  // Main-thread MCP calls currently executing. A long one (rename) may pump
  // messages; a WRITING call arriving inside that pump must not run nested.
  GMainDepth: Integer = 0;

constructor TMainCall.Create(const AProc: TProc; AAllowModal: Boolean);
begin
  inherited Create;
  Proc := AProc;
  AllowModal := AAllowModal;
  Done := CreateEvent(nil, True, False, nil);
end;

destructor TMainCall.Destroy;
begin
  CloseHandle(Done);
  inherited;
end;

constructor TDispatcher.Create;
begin
  inherited Create;
  FWnd := AllocateHWnd(WndProc);
  GWnd := FWnd;
end;

destructor TDispatcher.Destroy;
begin
  GWnd := 0;
  // Messages still queued for the window are dropped with it; their
  // callers have already given up (stop event) - the few call objects
  // they pinned are a deliberate, tiny leak at shutdown.
  DeallocateHWnd(FWnd);
  inherited;
end;

procedure TDispatcher.WndProc(var Msg: TMessage);
var
  Call: TMainCall;
begin
  if Msg.Msg <> WM_MCP_CALL then
  begin
    Msg.Result := DefWindowProc(FWnd, Msg.Msg, Msg.WParam, Msg.LParam);
    Exit;
  end;
  Call := TMainCall(Msg.LParam);
  try
    if not Call.AllowModal and (Application.ModalLevel > 0) then
      Call.Error := 'the IDE is busy: a modal dialog is open - close it and retry'
    else if not Call.AllowModal and (GMainDepth > 0) then
      Call.Error := 'the IDE is busy with another MCP request (e.g. a rename) - retry'
    else
    begin
      Inc(GMainDepth);
      try
        try
          Call.Proc();
        except
          on E: Exception do
            Call.Error := E.ClassName + ': ' + E.Message;
        end;
      finally
        Dec(GMainDepth);
      end;
    end;
  finally
    SetEvent(Call.Done);
    IInterface(Call)._Release;   // the reference PostMessage carried
  end;
end;

function RunOnMain(const AProc: TProc; AAllowModal: Boolean; AStop: THandle;
  out AError: string; ATimeoutMs: Cardinal = 15000): Boolean;
var
  Call: TMainCall;
  Ref: IInterface;
  Handles: array[0..1] of THandle;
begin
  AError := '';
  if GetCurrentThreadId = MainThreadID then
  begin
    try
      AProc();
      Exit(True);
    except
      on E: Exception do
      begin
        AError := E.Message;
        Exit(False);
      end;
    end;
  end;
  if GWnd = 0 then
  begin
    AError := 'the plugin is shutting down';
    Exit(False);
  end;
  Call := TMainCall.Create(AProc, AAllowModal);
  Ref := Call;
  IInterface(Call)._AddRef;   // travels with the message
  if not PostMessage(GWnd, WM_MCP_CALL, 0, LPARAM(Call)) then
  begin
    IInterface(Call)._Release;
    AError := 'cannot reach the IDE main thread';
    Exit(False);
  end;
  Handles[0] := Call.Done;
  Handles[1] := AStop;
  case WaitForMultipleObjects(2, @Handles[0], False, ATimeoutMs) of
    WAIT_OBJECT_0:
      begin
        AError := Call.Error;
        Result := AError = '';
      end;
    WAIT_OBJECT_0 + 1:
      begin
        AError := 'the plugin is shutting down';
        Result := False;
      end;
  else
    AError := Format('the IDE main thread did not respond within %d s',
      [ATimeoutMs div 1000]);
    Result := False;
  end;
end;

function McpRunOnMain(const AProc: TProc; AAllowModal: Boolean; AStop: THandle;
  out AError: string; ATimeoutMs: Cardinal): Boolean;
begin
  Result := RunOnMain(AProc, AAllowModal, AStop, AError, ATimeoutMs);
end;

procedure RegisterMcpTool(const AName: string; const AHandler: TMcpToolHandler);
begin
  if GToolRegistry = nil then
    GToolRegistry := TDictionary<string, TMcpToolHandler>.Create;
  GToolRegistry.AddOrSetValue(AName, AHandler);
end;

// ---------------------------------------------------------------------------
//  Main-thread pieces
// ---------------------------------------------------------------------------

function BuildContext: TMcpInstanceContext;
var
  MS: IOTAModuleServices;
  Group: IOTAProjectGroup;
  Svc: IOTAServices;
begin
  Result := Default(TMcpInstanceContext);
  Result.Pid := GetCurrentProcessId;
  Result.Bitness := SizeOf(Pointer) * 8;
  Result.WireVersion := McpWireVersion;
  Result.PluginVersion := PluginVersion;
  if Supports(BorlandIDEServices, IOTAServices, Svc) then
    Result.IdeVersion := ExtractFileName(ExcludeTrailingPathDelimiter(
      Svc.GetBaseRegistryKey));
  try
    Result.PluginBuild := FormatDateTime('yyyy-mm-dd hh:nn',
      TFile.GetLastWriteTime(GetModuleName(HInstance)));
  except
  end;
  if Supports(BorlandIDEServices, IOTAModuleServices, MS) then
  begin
    Group := MS.MainProjectGroup;
    if Group <> nil then
    begin
      Result.ProjectGroup := Group.FileName;
      for var I := 0 to Group.ProjectCount - 1 do
        if Group.Projects[I] <> nil then
          Result.Projects := Result.Projects + [Group.Projects[I].FileName];
      if Group.ActiveProject <> nil then
        Result.ActiveProject := Group.ActiveProject.FileName;
    end;
  end;
  if Editor <> nil then
  begin
    Result.ActiveFile := Editor.GetActiveFileName;
    Result.OpenFiles := Editor.GetOpenSourceFiles;
  end;
  Result.Busy := Application.ModalLevel > 0;
  Result.ToolsHash := McpToolsHash;
end;

function ReadContent(const AFile: string; out AContent: string): Boolean;
begin
  Result := (Editor <> nil) and Editor.ReadEditorContent(AFile, AContent);
  if not Result and FileExists(AFile) then
  begin
    AContent := TFile.ReadAllText(AFile);
    Result := True;
  end;
end;

function LoadedModule(const AFile: string): IOTAModule;
var
  MS: IOTAModuleServices;
begin
  Result := nil;
  if Supports(BorlandIDEServices, IOTAModuleServices, MS) then
    Result := MS.FindModule(AFile);
end;

// A fix on a file the IDE had NOT loaded must not stay behind as an
// invisible, modified module: the quick-fix appliers go through the editor
// helper, which opens such a file headless (no tab, not in the recent
// list) - the edit then lived only in memory and was lost or, worse,
// popped up as a save prompt when the IDE closed. The IDE never owned that
// file, so the result goes to DISK and the module is released again
// (discarding it when the fix failed).
function ReleaseHeadlessModule(const AFile: string; ASave: Boolean): Boolean;
var
  M: IOTAModule;
begin
  Result := True;
  M := LoadedModule(AFile);
  if M = nil then Exit;
  if ASave then
    Result := M.Save(False, True);
  M.CloseModule(True);
end;

function McpReadContent(const AFile: string; out AContent: string): Boolean;
begin
  Result := ReadContent(AFile, AContent);
end;

function AvailabilityName(const AUnit: string): string;
var
  Snap: IUnitSnapshot;
  Path: string;
begin
  Result := 'unknown';
  Snap := TUnitIndex.Instance.Snapshot;
  if (Snap = nil) or not Snap.TryGetUnitPath(AUnit, Path) then Exit;
  case CheckUnitAvailability(AUnit, Path) of
    uaInProject: Result := 'in_project';
    uaOnSearchPath: Result := 'search_path';
    uaDcuAvailable: Result := 'dcu';
    uaBrowsingOnly: Result := 'browsing_only';
  end;
end;

// ---------------------------------------------------------------------------
//  Tools
// ---------------------------------------------------------------------------

function OkResult(AResult: TJSONValue): string;
var
  O: TJSONObject;
begin
  O := TJSONObject.Create;
  try
    O.AddPair('ok', TJSONBool.Create(True));
    O.AddPair('result', AResult);
    Result := O.ToJSON;
  finally
    O.Free;
  end;
end;

function ErrResult(const AMsg: string): string;
var
  O: TJSONObject;
begin
  O := TJSONObject.Create;
  try
    O.AddPair('ok', TJSONBool.Create(False));
    O.AddPair('error', AMsg);
    Result := O.ToJSON;
  finally
    O.Free;
  end;
end;

function McpOk(AResult: TJSONValue): string;
begin
  Result := OkResult(AResult);
end;

function McpErr(const AMsg: string): string;
begin
  Result := ErrResult(AMsg);
end;

type
  TDiagSnapshot = record
    FileName: string;
    Content: string;
    Hash: Cardinal;
    Diags: TArray<TLspErrorDiag>;
    Sources: TArray<string>;
    Used, Stale, Note: string;
  end;

// Reads the buffer (main thread), refreshes the plugin's own LSP analysis
// when needed (this thread), and merges what the sources reported for the
// CURRENT content.
function CollectDiagnostics(AArgs: TJSONObject; AStop: THandle;
  out ASnap: TDiagSnapshot; out AError: string): Boolean;
var
  F, Active, C: string;
  Client: TLspClient;
  Refresh: Boolean;
begin
  Result := False;
  ASnap := Default(TDiagSnapshot);
  F := '';
  if AArgs <> nil then
  begin
    F := AArgs.GetValue<string>('file', '');
    Refresh := AArgs.GetValue<Boolean>('refresh', False);
  end
  else
    Refresh := False;
  Client := nil;
  var Found := False;
  if not RunOnMain(
    procedure
    begin
      Active := '';
      if Editor <> nil then Active := Editor.GetActiveFileName;
      if F = '' then F := Active;
      if F <> '' then
      begin
        F := ExpandFileName(F);
        Found := ReadContent(F, C);
      end;
      Client := TLspManager.Instance.PeekClient;
    end, True, AStop, AError) then Exit;
  if F = '' then
  begin
    AError := 'no file given and no file is active in the IDE editor';
    Exit;
  end;
  if not Found then
  begin
    AError := 'file not found: ' + F;
    Exit;
  end;
  ASnap.FileName := F;
  ASnap.Content := C;
  ASnap.Hash := DiagContentHash(C);

  ASnap.Diags := MergeFreshDiagnostics(StoredDiagnostics(F), ASnap.Hash,
    ASnap.Sources, ASnap.Used, ASnap.Stale);

  // Nothing current for this buffer state (or explicitly asked for): let
  // the plugin's own LSP session analyse it. For the ACTIVE file this is
  // normally not needed - the Structure view and the live checker keep it
  // current - and not refreshing it avoids racing the live checker's own
  // didClose/didOpen for the same document.
  if Refresh or (ASnap.Used = '') then
  begin
    if Client = nil then
      ASnap.Note := 'The plugin''s own DelphiLSP session is not running ' +
        '(it starts with the first opened project or refactoring action), so ' +
        'only the IDE''s own analysis of the ACTIVE editor file is available.'
    else
    begin
      try
        var Before := Client.GetFileDiagnosticsVersion(F);
        Client.RefreshDocumentWith(F, C);
        var TextDoc := TJSONObject.Create;
        TextDoc.AddPair('uri', TLspUri.PathToFileUri(F));
        var Params := TJSONObject.Create;
        Params.AddPair('textDocument', TextDoc);
        var ReqId := Client.SendRequestAsync('textDocument/documentSymbol', Params);
        if ReqId >= 0 then
          try Client.WaitForResponse(ReqId, LspWaitMs).Free; except end;
        var Waited := 0;
        while (Client.GetFileDiagnosticsVersion(F) <= Before) and (Waited < LspWaitMs) do
        begin
          if WaitForSingleObject(AStop, 50) = WAIT_OBJECT_0 then
          begin
            AError := 'the plugin is shutting down';
            Exit;
          end;
          Inc(Waited, 50);
        end;
        if Client.GetFileDiagnosticsVersion(F) > Before then
        begin
          StoreDiagnosticsHashed(F, SrcLsp, ASnap.Hash, Client.GetErrorDiagnostics(F));
          ASnap.Diags := MergeFreshDiagnostics(StoredDiagnostics(F), ASnap.Hash,
            ASnap.Sources, ASnap.Used, ASnap.Stale);
        end
        else
          ASnap.Note := 'DelphiLSP did not answer within 8 s - the result ' +
            'may be incomplete. Retry with refresh=true.';
      except
        on E: Exception do
          ASnap.Note := 'DelphiLSP analysis failed: ' + E.Message;
      end;
    end;
  end;
  if (ASnap.Used = '') and (ASnap.Note = '') then
    ASnap.Note := 'No source has analysed the current buffer state yet.';
  Result := True;
end;

procedure CacheFixes(const AFile: string; AHash: Cardinal; const AFixes: TArray<TQuickFix>);
begin
  GFixLock.Enter;
  try
    GFixCache.AddOrSetValue(UpperCase(AFile),
      TPair<Cardinal, TArray<TQuickFix>>.Create(AHash, AFixes));
  finally
    GFixLock.Leave;
  end;
end;

function CachedFixes(const AFile: string; AHash: Cardinal;
  out AFixes: TArray<TQuickFix>): Boolean;
var
  P: TPair<Cardinal, TArray<TQuickFix>>;
begin
  GFixLock.Enter;
  try
    Result := GFixCache.TryGetValue(UpperCase(AFile), P) and (P.Key = AHash);
    if Result then AFixes := P.Value;
  finally
    GFixLock.Leave;
  end;
end;

function ToolGetDiagnostics(AArgs: TJSONObject; AStop: THandle): string;
var
  Snap: TDiagSnapshot;
  Err: string;
begin
  if not CollectDiagnostics(AArgs, AStop, Snap, Err) then Exit(ErrResult(Err));
  Result := OkResult(DiagnosticsToJson(Snap.FileName, Snap.Hash, Snap.Diags,
    Snap.Sources, Snap.Used, Snap.Stale, Snap.Note));
end;

function ToolGetQuickFixes(AArgs: TJSONObject; AStop: THandle): string;
var
  Snap: TDiagSnapshot;
  Err, Note: string;
  Fixes: TArray<TQuickFix>;
  Avail: TDictionary<string, string>;
  Line1: Integer;
begin
  if not CollectDiagnostics(AArgs, AStop, Snap, Err) then Exit(ErrResult(Err));
  Line1 := 0;
  if AArgs <> nil then Line1 := AArgs.GetValue<Integer>('line', 0);
  // The resolver is thread-safe by design (the live checker runs it on
  // worker threads): index snapshot only, no ToolsAPI.
  Fixes := ResolveQuickFixes(Snap.Content, Snap.Diags);
  Note := Snap.Note;
  var Declined := LastResolveNote;
  if Declined <> '' then
  begin
    if Note <> '' then Note := Note + ' ';
    Note := Note + 'Declined: ' + Declined;
  end;
  CacheFixes(Snap.FileName, Snap.Hash, Fixes);

  // Whether the compiler can actually see a candidate unit needs the
  // project's search paths - ToolsAPI, so one trip to the main thread.
  Avail := TDictionary<string, string>.Create;
  try
    var Names: TArray<string> := nil;
    for var F in Fixes do
      if F.Kind = qfAddUnit then
        for var U in F.UnitNames do
          Names := Names + [U];
    if Length(Names) > 0 then
      RunOnMain(
        procedure
        begin
          for var U in Names do
            if not Avail.ContainsKey(UpperCase(U)) then
              Avail.Add(UpperCase(U), AvailabilityName(U));
        end, True, AStop, Err);
    Result := OkResult(QuickFixesToJson(Snap.FileName, Snap.Hash, Fixes, Line1,
      function(const AUnit: string): string
      begin
        if not Avail.TryGetValue(UpperCase(AUnit), Result) then Result := 'unknown';
      end,
      Length(Snap.Diags), Snap.Used, Snap.Stale, Note));
  finally
    Avail.Free;
  end;
end;

function ToolApplyQuickFix(AArgs: TJSONObject; AStop: THandle): string;
var
  F, FixId, UnitArg, SectionArg, Err: string;
  IdHash: Cardinal;
  Idx: Integer;
  Res: TJSONObject;
begin
  if AArgs = nil then Exit(ErrResult('arguments "file" and "fix_id" are required'));
  F := AArgs.GetValue<string>('file', '');
  FixId := AArgs.GetValue<string>('fix_id', '');
  UnitArg := AArgs.GetValue<string>('unit', '');
  SectionArg := AArgs.GetValue<string>('section', '');
  if (F = '') or (FixId = '') then
    Exit(ErrResult('arguments "file" and "fix_id" are required'));
  if not ParseFixId(FixId, IdHash, Idx) then
    Exit(ErrResult('malformed fix_id "' + FixId + '" - use an id from get_quick_fixes'));
  F := ExpandFileName(F);

  Res := nil;
  var Msg := '';
  var Ok := RunOnMain(
    procedure
    var
      C, Path: string;
      Fixes: TArray<TQuickFix>;
      Fix: TQuickFix;
      UnitIdx: Integer;
      Refused: Boolean;
      Saved: TRemoveConfirmFunc;
      Applied: Boolean;
    begin
      if not ReadContent(F, C) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      var Cur := DiagContentHash(C);
      if Cur <> IdHash then
      begin
        Msg := Format('stale fix id: the buffer changed since get_quick_fixes ' +
          '(revision %s, now %s) - list the fixes again',
          [IntToHex(IdHash, 8), IntToHex(Cur, 8)]);
        Exit;
      end;
      if not CachedFixes(F, Cur, Fixes) then
      begin
        Msg := 'unknown fix id - call get_quick_fixes for this file first';
        Exit;
      end;
      if Idx > High(Fixes) then
      begin
        Msg := 'unknown fix id "' + FixId + '"';
        Exit;
      end;
      Fix := Fixes[Idx];
      if SameText(SectionArg, 'interface') then Fix.Section := usInterface
      else if SameText(SectionArg, 'implementation') then Fix.Section := usImplementation;

      UnitIdx := -1;
      if Fix.Kind = qfAddUnit then
      begin
        UnitIdx := 0;
        if UnitArg <> '' then
        begin
          UnitIdx := -1;
          for var I := 0 to High(Fix.UnitNames) do
            if SameText(Fix.UnitNames[I], UnitArg) then UnitIdx := I;
          if UnitIdx < 0 then
          begin
            Msg := 'unit "' + UnitArg + '" is not a candidate of this fix: ' +
              string.Join(', ', Fix.UnitNames);
            Exit;
          end;
        end;
        if UnitIdx > High(Fix.UnitNames) then
        begin
          Msg := 'this fix has no candidate unit';
          Exit;
        end;
        // The IDE asks interactively when only the BROWSING path reaches the
        // unit (add the file to the project / the directory to the search
        // path?). A remote caller cannot answer that dialog - refuse and say
        // why, instead of producing a uses entry that will not compile.
        var U := Fix.UnitNames[UnitIdx];
        var Snap := TUnitIndex.Instance.Snapshot;
        if (Snap <> nil) and Snap.TryGetUnitPath(U, Path) and
           (CheckUnitAvailability(U, Path) = uaBrowsingOnly) then
        begin
          Msg := Format('%s is only on the IDE''s BROWSING path (%s) - the ' +
            'compiler would not find it. Add it in the IDE (the quick-fix popup ' +
            'offers to add the file to the project or its folder to the search ' +
            'path) or pick another candidate.', [U, Path]);
          Exit;
        end;
      end;

      // A non-empty body is deleted only after the user confirmed it in the
      // IDE (user requirement). Nobody can confirm from here.
      var WasLoaded := LoadedModule(F) <> nil;
      Refused := False;
      Saved := RemovePrivateConfirm;
      if Fix.Kind = qfRemovePrivate then
        RemovePrivateConfirm :=
          function(const AInfo: TPrivateMember): Boolean
          begin
            Refused := True;
            Result := False;
          end;
      try
        Applied := ApplyQuickFix(F, Fix, UnitIdx);
      finally
        RemovePrivateConfirm := Saved;
      end;
      var SavedToDisk := False;
      if not WasLoaded then
      begin
        SavedToDisk := Applied and not Refused;
        if not ReleaseHeadlessModule(F, SavedToDisk) then
        begin
          Msg := 'the fix was applied but ' + F + ' could not be saved';
          Exit;
        end;
      end;
      if Refused then
      begin
        Msg := 'refused: removing ' + Fix.Identifier + ' would delete a non-empty ' +
          'body - that needs a confirmation in the IDE (use the quick-fix popup there)';
        Exit;
      end;
      if not Applied then
      begin
        Msg := 'the fix could not be applied - the code at that place no longer ' +
          'matches (list the fixes again)';
        Exit;
      end;

      var After := '';
      ReadContent(F, After);
      Res := TJSONObject.Create;
      Res.AddPair('applied', TJSONBool.Create(True));
      Res.AddPair('file', F);
      Res.AddPair('fix', FixDescription(Fix));
      if Fix.Kind = qfAddUnit then Res.AddPair('unit', Fix.UnitNames[UnitIdx]);
      Res.AddPair('revision', IntToHex(DiagContentHash(After), 8));
      if SavedToDisk then
        Res.AddPair('note', 'The file was not open in the IDE, so the change ' +
          'was written to DISK. Fix ids of this file have expired - call ' +
          'get_quick_fixes again before applying another fix.')
      else
        Res.AddPair('note', 'The editor buffer was changed and is NOT saved. ' +
          'Fix ids of this file have expired; the IDE re-analyses within a few ' +
          'seconds - call get_quick_fixes again before applying another fix.');
      Res.AddPair('savedToDisk', TJSONBool.Create(SavedToDisk));
    end, False, AStop, Err);
  if not Ok then Exit(ErrResult(Err));
  if Msg <> '' then
  begin
    Res.Free;
    Exit(ErrResult(Msg));
  end;
  Result := OkResult(Res);
end;

// ---------------------------------------------------------------------------
//  Buffers: IDE editor buffers and in-memory scratch units
// ---------------------------------------------------------------------------
//
// IDE BUFFERS. buffer_open loads a file either VISIBLY (an editor tab, as if
// the user opened it) or HEADLESS (a module without a tab - it exists only in
// the IDE's memory). Headless modules are tracked in GHeadless: they are the
// only ones buffer_close may save or discard; a file the USER opened is
// never closed from here. buffer_edit and apply_quick_fix work on any
// loaded buffer and never save.
//
// SCRATCH UNITS never touch the IDE or the disk. Their content lives in
// GScratch and is handed to the plugin's own DelphiLSP session as an open
// document. Measured (scratchpad lspprobe\ProbeVirtual.dpr): DelphiLSP
// analyses a document that does not exist on disk as long as its FOLDER
// exists - so the scratch folder is created (and stays empty) - project
// units resolve from there, and the unit name must match the file name
// (else E1038).

var
  GHeadless: TList<string> = nil;               // main thread only
  GScratch: TDictionary<string, string> = nil;   // name -> content, GFixLock

function IsHeadless(const AFile: string): Boolean;
begin
  Result := (GHeadless <> nil) and GHeadless.Contains(UpperCase(AFile));
end;

function ModuleIsModified(const M: IOTAModule): Boolean;
begin
  Result := False;
  if M = nil then Exit;
  for var I := 0 to M.GetModuleFileCount - 1 do
    if (M.GetModuleFileEditor(I) <> nil) and M.GetModuleFileEditor(I).Modified then
      Exit(True);
end;

function SplitLinesLF(const S: string): TArray<string>;
begin
  Result := S.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

function RequireFile(AArgs: TJSONObject; out AFile, AError: string): Boolean;
begin
  AFile := '';
  if AArgs <> nil then AFile := AArgs.GetValue<string>('file', '');
  Result := AFile <> '';
  if Result then
    AFile := ExpandFileName(AFile)
  else
    AError := 'argument "file" is required';
end;

function ToolBufferOpen(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Msg: string;
  Res: TJSONObject;
begin
  if not RequireFile(AArgs, F, Err) then Exit(ErrResult(Err));
  var Visible := AArgs.GetValue<Boolean>('visible', False);
  Res := nil;
  Msg := '';
  if not RunOnMain(
    procedure
    var
      MS: IOTAModuleServices;
      AS_: IOTAActionServices;
      M: IOTAModule;
      C: string;
      How: string;
    begin
      if not FileExists(F) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      if not Supports(BorlandIDEServices, IOTAModuleServices, MS) then Exit;
      M := MS.FindModule(F);
      if Visible then
      begin
        if Supports(BorlandIDEServices, IOTAActionServices, AS_) and AS_.OpenFile(F) then
        begin
          // shown to the user now - no longer ours to close
          if IsHeadless(F) then GHeadless.Remove(UpperCase(F));
          How := 'visible';
        end
        else
        begin
          Msg := 'the IDE could not open ' + F;
          Exit;
        end;
      end
      else if M <> nil then
      begin
        if IsHeadless(F) then How := 'headless (already)' else How := 'already loaded by the IDE';
      end
      else
      begin
        M := MS.OpenModule(F);
        if M = nil then
        begin
          Msg := 'the IDE could not load ' + F;
          Exit;
        end;
        GHeadless.Add(UpperCase(F));
        How := 'headless';
      end;
      ReadContent(F, C);
      Res := TJSONObject.Create;
      Res.AddPair('file', F);
      Res.AddPair('state', How);
      Res.AddPair('revision', IntToHex(DiagContentHash(C), 8));
      Res.AddPair('lines', TJSONNumber.Create(Length(SplitLinesLF(C))));
      if How = 'headless' then
        Res.AddPair('note', 'Loaded into the IDE''s memory without an editor tab. ' +
          'Edits (buffer_edit, apply_quick_fix) stay in memory until buffer_close ' +
          'saves or discards them.');
    end, False, AStop, Err) then Exit(ErrResult(Err));
  if Msg <> '' then
  begin
    Res.Free;
    Exit(ErrResult(Msg));
  end;
  Result := OkResult(Res);
end;

function ToolBufferRead(AArgs: TJSONObject; AStop: THandle): string;
const
  MaxLines = 2000;
var
  F, Err, C: string;
  Loaded, Modified, Headless, Found: Boolean;
begin
  if not RequireFile(AArgs, F, Err) then Exit(ErrResult(Err));
  var First := AArgs.GetValue<Integer>('start_line', 1);
  var Last := AArgs.GetValue<Integer>('end_line', 0);
  if not RunOnMain(
    procedure
    begin
      var M := LoadedModule(F);
      Loaded := M <> nil;
      Modified := ModuleIsModified(M);
      Headless := IsHeadless(F);
      Found := ReadContent(F, C);
    end, True, AStop, Err) then Exit(ErrResult(Err));
  if not Found then Exit(ErrResult('file not found: ' + F));
  var Lines := SplitLinesLF(C);
  if First < 1 then First := 1;
  if (Last <= 0) or (Last > Length(Lines)) then Last := Length(Lines);
  var Truncated := Last - First + 1 > MaxLines;
  if Truncated then Last := First + MaxLines - 1;
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('revision', IntToHex(DiagContentHash(C), 8));
  if not Loaded then
    Res.AddPair('source', 'disk (not loaded in the IDE)')
  else if Headless then
    Res.AddPair('source', 'IDE buffer (headless, opened by this bridge)')
  else
    Res.AddPair('source', 'IDE buffer');
  Res.AddPair('modified', TJSONBool.Create(Modified));
  Res.AddPair('totalLines', TJSONNumber.Create(Length(Lines)));
  Res.AddPair('startLine', TJSONNumber.Create(First));
  var Arr := TJSONArray.Create;
  for var I := First to Last do
    Arr.Add(Lines[I - 1]);
  Res.AddPair('lines', Arr);
  if Truncated then
    Res.AddPair('note', Format('Only %d lines returned - pass start_line/end_line ' +
      'for the rest.', [MaxLines]));
  Result := OkResult(Res);
end;

function CountOccurrences(const S, Sub: string; out AFirst: Integer): Integer;
var
  P: Integer;
begin
  Result := 0;
  AFirst := 0;
  if Sub = '' then Exit;
  P := Pos(Sub, S);
  while P > 0 do
  begin
    if Result = 0 then AFirst := P;
    Inc(Result);
    P := Pos(Sub, S, P + Length(Sub));
  end;
end;

function ToolBufferEdit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Msg: string;
  Res: TJSONObject;
begin
  if not RequireFile(AArgs, F, Err) then Exit(ErrResult(Err));
  var OldText := AArgs.GetValue<string>('old_text', '').Replace(#13#10, #10);
  var NewText := AArgs.GetValue<string>('new_text', '').Replace(#13#10, #10);
  var Revision := AArgs.GetValue<string>('revision', '');
  if OldText = '' then Exit(ErrResult('argument "old_text" is required (the ' +
    'exact text to replace, unique in the buffer)'));
  Res := nil;
  Msg := '';
  if not RunOnMain(
    procedure
    var
      C, Norm: string;
      N, P: Integer;
      SL: TStringList;
    begin
      if LoadedModule(F) = nil then
      begin
        Msg := F + ' is not loaded in the IDE - call buffer_open first (headless ' +
          'keeps the change in memory, visible opens an editor tab)';
        Exit;
      end;
      ReadContent(F, C);
      var Cur := IntToHex(DiagContentHash(C), 8);
      if (Revision <> '') and not SameText(Revision, Cur) then
      begin
        Msg := Format('stale revision: the buffer is at %s, not %s - read it again',
          [Cur, Revision]);
        Exit;
      end;
      Norm := C.Replace(#13#10, #10);
      N := CountOccurrences(Norm, OldText, P);
      if N = 0 then
      begin
        Msg := 'old_text was not found in the buffer (whitespace and line breaks ' +
          'must match exactly)';
        Exit;
      end;
      if N > 1 then
      begin
        Msg := Format('old_text occurs %d times - add surrounding lines to make ' +
          'it unique', [N]);
        Exit;
      end;
      Delete(Norm, P, Length(OldText));
      Insert(NewText, Norm, P);
      SL := TStringList.Create;
      try
        SL.Text := Norm;
        // Line-level minimal edit through the editor: only the lines that
        // really change are touched (one undo step each, change bars only
        // there).
        if not ApplyLinesMinimal(F, SL, C) then
        begin
          Msg := 'the IDE refused the edit';
          Exit;
        end;
      finally
        SL.Free;
      end;
      var After := '';
      ReadContent(F, After);
      Res := TJSONObject.Create;
      Res.AddPair('file', F);
      Res.AddPair('revision', IntToHex(DiagContentHash(After), 8));
      Res.AddPair('line', TJSONNumber.Create(Length(SplitLinesLF(Copy(Norm, 1, P)))));
      Res.AddPair('saved', TJSONBool.Create(False));
    end, False, AStop, Err) then Exit(ErrResult(Err));
  if Msg <> '' then
  begin
    Res.Free;
    Exit(ErrResult(Msg));
  end;
  Result := OkResult(Res);
end;

function ToolBufferList(AArgs: TJSONObject; AStop: THandle): string;
var
  Err: string;
  Arr: TJSONArray;
begin
  Arr := TJSONArray.Create;
  if not RunOnMain(
    procedure
    var
      MS: IOTAModuleServices;
    begin
      if not Supports(BorlandIDEServices, IOTAModuleServices, MS) then Exit;
      for var I := 0 to MS.ModuleCount - 1 do
      begin
        var M := MS.Modules[I];
        if M = nil then Continue;
        var FN := M.FileName;
        var Ext := LowerCase(ExtractFileExt(FN));
        if (Ext = '.groupproj') or (Ext = '.dproj') then Continue;
        var O := TJSONObject.Create;
        O.AddPair('file', FN);
        O.AddPair('modified', TJSONBool.Create(ModuleIsModified(M)));
        O.AddPair('headless', TJSONBool.Create(IsHeadless(FN)));
        Arr.Add(O);
      end;
    end, True, AStop, Err) then
  begin
    Arr.Free;
    Exit(ErrResult(Err));
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('buffers', Arr);
  var Sc := TJSONArray.Create;
  GFixLock.Enter;
  try
    for var K in GScratch.Keys do Sc.Add(K);
  finally
    GFixLock.Leave;
  end;
  Res.AddPair('scratchUnits', Sc);
  Result := OkResult(Res);
end;

function ToolBufferClose(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Msg: string;
  Saved: Boolean;
begin
  if not RequireFile(AArgs, F, Err) then Exit(ErrResult(Err));
  if AArgs.GetValue('save') = nil then
    Exit(ErrResult('argument "save" is required: true writes the buffer to disk, ' +
      'false discards the changes'));
  var DoSave := AArgs.GetValue<Boolean>('save', False);
  Msg := '';
  Saved := False;
  if not RunOnMain(
    procedure
    begin
      if not IsHeadless(F) then
      begin
        if LoadedModule(F) = nil then
          Msg := F + ' is not loaded in the IDE'
        else
          Msg := F + ' was not opened headless by this bridge - it belongs to the ' +
            'user''s editor and is not closed from here';
        Exit;
      end;
      var M := LoadedModule(F);
      if M <> nil then
      begin
        if DoSave then
        begin
          Saved := M.Save(False, True);
          if not Saved then
          begin
            Msg := 'saving ' + F + ' failed - the buffer stays loaded';
            Exit;
          end;
        end;
        M.CloseModule(True);
      end;
      GHeadless.Remove(UpperCase(F));
    end, False, AStop, Err) then Exit(ErrResult(Err));
  if Msg <> '' then Exit(ErrResult(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('closed', TJSONBool.Create(True));
  Res.AddPair('saved', TJSONBool.Create(Saved));
  Result := OkResult(Res);
end;

// Saves ANY loaded buffer - one the user opened as well as a headless one
// (which stays loaded; buffer_close releases it). Same as Ctrl+S in the
// IDE: the IDE writes the file itself, so there is no "changed on disk"
// prompt afterwards.
function ToolBufferSave(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Msg: string;
  WasModified, Saved: Boolean;
begin
  if not RequireFile(AArgs, F, Err) then Exit(ErrResult(Err));
  Msg := '';
  WasModified := False;
  Saved := False;
  if not RunOnMain(
    procedure
    begin
      var M := LoadedModule(F);
      if M = nil then
      begin
        Msg := F + ' is not loaded in the IDE - there is no buffer to save';
        Exit;
      end;
      WasModified := ModuleIsModified(M);
      if not WasModified then Exit;
      Saved := M.Save(False, True);
      if not Saved then
        Msg := 'the IDE could not save ' + F + ' (read-only or locked?)';
    end, False, AStop, Err) then Exit(ErrResult(Err));
  if Msg <> '' then Exit(ErrResult(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('saved', TJSONBool.Create(Saved));
  if not WasModified then
    Res.AddPair('note', 'The buffer had no unsaved changes - nothing written.');
  Result := OkResult(Res);
end;

// The status window's rows - what the plugin is doing and why something
// does NOT work (index, LSP session, live checker and its sources, menus,
// resources, blame, ...). Collected on the main thread like the window
// does it; the window does not have to be open.
function ToolGetStatus(AArgs: TJSONObject; AStop: THandle): string;
var
  Rows: TArray<TStatusRow>;
  Err: string;
begin
  if not RunOnMain(
    procedure
    begin
      Rows := StatusSnapshot;
    end, True, AStop, Err) then Exit(ErrResult(Err));
  var Filter := '';
  if AArgs <> nil then Filter := Trim(AArgs.GetValue<string>('filter', ''));
  var Arr := TJSONArray.Create;
  var Section := '';
  for var R in Rows do
  begin
    var Item := Trim(R.Caption);
    // indented captions are sub-rows of the last top-level row
    var IsSub := (R.Caption <> '') and (R.Caption[1] = ' ');
    if not IsSub then Section := Item;
    if (Filter <> '') and (Pos(UpperCase(Filter), UpperCase(Item + ' ' + Section)) = 0) then
      Continue;
    var O := TJSONObject.Create;
    O.AddPair('item', Item);
    if IsSub then O.AddPair('section', Section);
    O.AddPair('status', R.Value);
    if R.Detail <> '' then O.AddPair('details', R.Detail);
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('rows', Arr);
  Result := OkResult(Res);
end;

function ScratchDir: string;
begin
  Result := TPath.Combine(TPath.GetTempPath,
    'DelphiRefactoringLight\scratch\' + IntToStr(GetCurrentProcessId));
end;

function ToolScratchAnalyze(AArgs: TJSONObject; AStop: THandle): string;
var
  Name, Content, Path, Err, Note: string;
  Client: TLspClient;
begin
  Name := '';
  Content := '';
  if AArgs <> nil then
  begin
    Name := Trim(AArgs.GetValue<string>('name', ''));
    Content := AArgs.GetValue<string>('content', '');
  end;
  if not IsValidIdent(Name, True) then
    Exit(ErrResult('argument "name" must be a unit name (identifier, dots allowed)'));
  if Content = '' then
    Exit(ErrResult('argument "content" is required (the complete unit text)'));
  Content := Content.Replace(#13#10, #10).Replace(#10, #13#10);

  Client := nil;
  if not RunOnMain(
    procedure
    begin
      Client := TLspManager.Instance.PeekClient;
    end, True, AStop, Err) then Exit(ErrResult(Err));
  if Client = nil then
    Exit(ErrResult('the plugin''s own DelphiLSP session is not running yet - it ' +
      'starts with the first opened project; open the project in the IDE and retry'));

  // DelphiLSP ignores documents whose FOLDER does not exist (measured).
  ForceDirectories(ScratchDir);
  Path := TPath.Combine(ScratchDir, Name + '.pas');
  GFixLock.Enter;
  try
    GScratch.AddOrSetValue(Name, Content);
  finally
    GFixLock.Leave;
  end;

  var Before := Client.GetFileDiagnosticsVersion(Path);
  try
    Client.RefreshDocumentWith(Path, Content);
    var TD := TJSONObject.Create;
    TD.AddPair('uri', TLspUri.PathToFileUri(Path));
    var P := TJSONObject.Create;
    P.AddPair('textDocument', TD);
    var Id := Client.SendRequestAsync('textDocument/documentSymbol', P);
    if Id >= 0 then
      try Client.WaitForResponse(Id, LspWaitMs).Free; except end;
  except
    on E: Exception do Exit(ErrResult('DelphiLSP analysis failed: ' + E.Message));
  end;
  var Waited := 0;
  while (Client.GetFileDiagnosticsVersion(Path) <= Before) and (Waited < LspWaitMs) do
  begin
    if WaitForSingleObject(AStop, 50) = WAIT_OBJECT_0 then
      Exit(ErrResult('the plugin is shutting down'));
    Inc(Waited, 50);
  end;
  if Client.GetFileDiagnosticsVersion(Path) <= Before then
    Note := 'DelphiLSP did not answer within 8 s - the result may be incomplete.';

  var Diags := Client.GetErrorDiagnostics(Path);
  var Sources: TArray<string>;
  SetLength(Sources, Length(Diags));
  for var I := 0 to High(Sources) do Sources[I] := 'lsp';
  var Hash := DiagContentHash(Content);
  var Res := DiagnosticsToJson(Path, Hash, Diags, Sources, 'lsp', '', Note);
  Res.AddPair('name', Name);
  Res.AddPair('onDisk', TJSONBool.Create(False));
  // What the plugin WOULD offer - there is no IDE buffer to apply it to.
  var Fixes := ResolveQuickFixes(Content, Diags);
  var FA := TJSONArray.Create;
  for var Fx in Fixes do
  begin
    var O := TJSONObject.Create;
    O.AddPair('kind', FixKindName(Fx.Kind));
    O.AddPair('line', TJSONNumber.Create(Fx.Line + 1));
    O.AddPair('description', FixDescription(Fx));
    if Fx.Kind = qfAddUnit then O.AddPair('units', string.Join(', ', Fx.UnitNames));
    FA.Add(O);
  end;
  Res.AddPair('suggestedFixes', FA);
  Result := OkResult(Res);
end;

function ToolScratchClose(AArgs: TJSONObject; AStop: THandle): string;
var
  Name, Err: string;
  Client: TLspClient;
  Names: TArray<string>;
begin
  Name := '';
  if AArgs <> nil then Name := Trim(AArgs.GetValue<string>('name', ''));
  GFixLock.Enter;
  try
    if Name = '' then
      Names := GScratch.Keys.ToArray
    else if GScratch.ContainsKey(Name) then
      Names := [Name]
    else
      Names := nil;
    for var N in Names do GScratch.Remove(N);
  finally
    GFixLock.Leave;
  end;
  if (Name <> '') and (Length(Names) = 0) then
    Exit(ErrResult('no scratch unit named "' + Name + '"'));
  Client := nil;
  RunOnMain(
    procedure
    begin
      Client := TLspManager.Instance.PeekClient;
    end, True, AStop, Err);
  if Client <> nil then
    for var N in Names do
      try
        Client.CloseDocument(TPath.Combine(ScratchDir, N + '.pas'));
      except
      end;
  var Res := TJSONObject.Create;
  var Arr := TJSONArray.Create;
  for var N in Names do Arr.Add(N);
  Res.AddPair('closed', Arr);
  Result := OkResult(Res);
end;

// ---- per-tool statistics (tools window) ----------------------------------

const
  BuiltinTools: array[0..11] of string = ('get_diagnostics', 'get_quick_fixes',
    'apply_quick_fix', 'buffer_open', 'buffer_read', 'buffer_edit',
    'buffer_list', 'buffer_close', 'buffer_save', 'get_status',
    'scratch_analyze', 'scratch_close');

function McpHandlesTool(const AName: string): Boolean;
begin
  for var B in BuiltinTools do
    if B = AName then Exit(True);
  Result := (GToolRegistry <> nil) and GToolRegistry.ContainsKey(AName);
end;

function StatBegin(const ATool: string): UInt64;
var
  S: TMcpToolStat;
begin
  Result := GetTickCount64;
  if (GStatLock = nil) or (ATool = '') then Exit;
  GStatLock.Enter;
  try
    if not GStats.TryGetValue(ATool, S) then
    begin
      S := Default(TMcpToolStat);
      S.Name := ATool;
    end;
    Inc(S.Calls);
    Inc(S.Running);
    S.LastTime := Now;
    GStats.AddOrSetValue(ATool, S);
  finally
    GStatLock.Leave;
  end;
end;

procedure StatEnd(const ATool: string; AStart: UInt64; const AResult: string);
var
  S: TMcpToolStat;
  Ms: Int64;
  Ok: Boolean;
  Msg: string;
begin
  if (GStatLock = nil) or (ATool = '') then Exit;
  Ms := Int64(GetTickCount64 - AStart);
  // Results are '{"ok":true,...' / '{"ok":false,"error":"..."}' - the
  // prefix decides, only a failure is parsed for its message.
  Ok := AResult.StartsWith('{"ok":true');
  Msg := '';
  if not Ok then
  begin
    var V := TJSONObject.ParseJSONValue(AResult);
    try
      if V is TJSONObject then Msg := TJSONObject(V).GetValue<string>('error', '');
    finally
      V.Free;
    end;
    if Msg = '' then Msg := Copy(AResult, 1, 200);
  end;
  GStatLock.Enter;
  try
    if not GStats.TryGetValue(ATool, S) then Exit;
    if S.Running > 0 then Dec(S.Running);
    S.LastMs := Ms;
    Inc(S.TotalMs, Ms);
    if Ms > S.MaxMs then S.MaxMs := Ms;
    S.LastOk := Ok;
    if not Ok then
    begin
      Inc(S.Errors);
      S.LastError := Msg;
    end;
    GStats.AddOrSetValue(ATool, S);
  finally
    GStatLock.Leave;
  end;
end;

function McpToolStats: TArray<TMcpToolStat>;
begin
  Result := nil;
  if GStatLock = nil then Exit;
  GStatLock.Enter;
  try
    Result := GStats.Values.ToArray;
  finally
    GStatLock.Leave;
  end;
end;

function HandleRequest(const ARequest: string; AStop: THandle): string;
var
  V: TJSONValue;
  Req, Args: TJSONObject;
  Method, Tool, Err: string;
  Ctx: TMcpInstanceContext;
begin
  V := TJSONObject.ParseJSONValue(ARequest);
  try
    if not (V is TJSONObject) then Exit(ErrResult('malformed request'));
    Req := TJSONObject(V);
    Method := Req.GetValue<string>('method', '');
    if Method = 'context' then
    begin
      if not RunOnMain(
        procedure
        begin
          Ctx := BuildContext;
        end, True, AStop, Err) then
        Exit(ErrResult(Err));
      Exit(OkResult(Ctx.ToJson));
    end;
    if Method = 'tools' then
    begin
      // The bridge builds its tool list from what the IDEs serve, so a new
      // tool ships with the plugin alone.
      var R := TJSONObject.Create;
      R.AddPair('tools', McpToolDefinitions);
      Exit(OkResult(R));
    end;
    if Method <> 'call' then Exit(ErrResult('unknown method "' + Method + '"'));
    Tool := Req.GetValue<string>('tool', '');
    Args := nil;
    if Req.GetValue('arguments') is TJSONObject then
      Args := TJSONObject(Req.GetValue('arguments'));
    GLastTool := Tool;
    var T0 := StatBegin(Tool);
    var Failed := True;
    try
    if Tool = 'get_diagnostics' then Result := ToolGetDiagnostics(Args, AStop)
    else if Tool = 'get_quick_fixes' then Result := ToolGetQuickFixes(Args, AStop)
    else if Tool = 'apply_quick_fix' then Result := ToolApplyQuickFix(Args, AStop)
    else if Tool = 'buffer_open' then Result := ToolBufferOpen(Args, AStop)
    else if Tool = 'buffer_read' then Result := ToolBufferRead(Args, AStop)
    else if Tool = 'buffer_edit' then Result := ToolBufferEdit(Args, AStop)
    else if Tool = 'buffer_list' then Result := ToolBufferList(Args, AStop)
    else if Tool = 'buffer_close' then Result := ToolBufferClose(Args, AStop)
    else if Tool = 'buffer_save' then Result := ToolBufferSave(Args, AStop)
    else if Tool = 'get_status' then Result := ToolGetStatus(Args, AStop)
    else if Tool = 'scratch_analyze' then Result := ToolScratchAnalyze(Args, AStop)
    else if Tool = 'scratch_close' then Result := ToolScratchClose(Args, AStop)
    else
    begin
      var Handler: TMcpToolHandler;
      if (GToolRegistry <> nil) and GToolRegistry.TryGetValue(Tool, Handler) then
        Result := Handler(Args, AStop)
      else
        Result := ErrResult('this IDE does not know the tool "' + Tool +
          '" - is the installed plugin older than the bridge?');
    end;
      Failed := False;
    finally
      if Failed then
        StatEnd(Tool, T0, '{"ok":false,"error":"exception in the tool handler"}')
      else
        StatEnd(Tool, T0, Result);
    end;
  finally
    V.Free;
  end;
end;

// ---------------------------------------------------------------------------

procedure StartMcpServer;
begin
  if GServer <> nil then Exit;
  GFixLock := TCriticalSection.Create;
  GFixCache := TDictionary<string, TPair<Cardinal, TArray<TQuickFix>>>.Create;
  GHeadless := TList<string>.Create;
  GScratch := TDictionary<string, string>.Create;
  GDispatcher := TDispatcher.Create;
  GServer := TMcpPipeServer.Create(McpPipeName(GetCurrentProcessId), HandleRequest);
  GServer.Start;
end;

procedure StopMcpServer;
begin
  // Server first: it signals the stop event and waits for every handler,
  // and the handlers' main-thread calls give up on that event. Only then
  // may the dispatch window go.
  if GServer <> nil then
  begin
    GServer.Stop;
    FreeAndNil(GServer);
  end;
  FreeAndNil(GDispatcher);
  // Headless buffers stay loaded: they may hold edits nobody decided about
  // yet, and the IDE asks about modified modules on its own when it closes.
  FreeAndNil(GHeadless);
  FreeAndNil(GScratch);
  FreeAndNil(GFixCache);
  FreeAndNil(GFixLock);
end;

function McpConnectionStatus(out ADetail: string): string;
const
  ConnectedMs = 15000;
begin
  ADetail := '';
  if GServer = nil then Exit('-');
  var Clients := GServer.RecentClients(ConnectedMs);
  if Length(Clients) = 0 then
  begin
    Result := 'no Claude Code session connected';
    var Older := GServer.RecentClients(3600000);
    if Length(Older) > 0 then
      ADetail := Format('last contact %d s ago (bridge pid %d)',
        [(GetTickCount64 - Older[0].LastTick) div 1000, Older[0].Pid])
    else
      ADetail := 'no bridge has contacted this IDE yet - register ' +
        'RefactoringLightMcp.exe in Claude Code (see install.cmd)';
    Exit;
  end;
  if Length(Clients) = 1 then
    Result := 'connected (1 Claude Code session)'
  else
    Result := Format('connected (%d Claude Code sessions)', [Length(Clients)]);
  for var C in Clients do
  begin
    if ADetail <> '' then ADetail := ADetail + ', ';
    ADetail := ADetail + Format('bridge pid %d: %d request(s), last %d s ago',
      [C.Pid, C.Requests, (GetTickCount64 - C.LastTick) div 1000]);
  end;
end;

function McpServerPipe: string;
begin
  Result := McpPipeName(GetCurrentProcessId);
end;

function McpServerStatus: string;
begin
  if GServer = nil then Exit('not running');
  if GServer.LastError <> '' then
    Result := 'ERROR: ' + GServer.LastError
  else
    Result := Format('listening, %d request(s)', [GServer.RequestCount]);
  if (GHeadless <> nil) and (GHeadless.Count > 0) then
    Result := Result + Format(', %d headless buffer(s)', [GHeadless.Count]);
  if GLastTool <> '' then Result := Result + ', last tool: ' + GLastTool;
end;

initialization
  GStatLock := TCriticalSection.Create;
  GStats := TDictionary<string, TMcpToolStat>.Create;

finalization
  FreeAndNil(GToolRegistry);
  FreeAndNil(GStats);
  FreeAndNil(GStatLock);

end.
