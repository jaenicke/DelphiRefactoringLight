(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.McpLspTools;

// Direct access to the plugin's own DelphiLSP session as MCP tools
// (IDE-only): hover, definition, implementation, references, symbols,
// signature help, completion - and lsp_request for anything else.
//
// BUFFER SYNC: DelphiLSP must see what the IDE editor shows. Before a
// position request the file's CURRENT content (buffer if loaded, else disk)
// is compared with what this unit last sent (per-file hash); only a change
// is sent again (didClose/didOpen with the text). The live checker keeps
// the active file current anyway; this covers every other file.
//
// POSITIONS: the tools take and return 1-based lines/columns like every
// other tool of the bridge. The RAW tools (symbols, signature help,
// lsp_request) return DelphiLSP's JSON unchanged - 0-based, as the LSP
// specification defines it; their descriptions say so.
//
// lsp_request refuses the lifecycle and document-sync methods: they would
// break the session the plugin itself depends on.

interface

uses
  Lsp.Client;

/// <summary>Hands the CURRENT content of AFiles (editor buffer when open,
///  else disk) to AClient's session, skipping files whose content it
///  already has. DelphiLSP answers GotoDefinition & co only for documents
///  opened in the session - a verification pass over files it has never
///  seen silently loses those hits. Handler-thread only (reads the buffers
///  through McpRunOnMain).</summary>
function McpSyncLspDocuments(AClient: TLspClient; const AFiles: TArray<string>;
  AStop: THandle; out AError: string): Boolean;

/// <summary>Hands AContent to the session ONLY when it differs from what the
///  session already has, and says whether it SENT - every didOpen makes
///  DelphiLSP analyse the unit again. After a send, McpWaitLspAnalysed waits
///  for the diagnostics push before the file is queried.</summary>
function McpSyncLspContent(AClient: TLspClient; const AFile, AContent: string): Boolean;

/// <summary>Waits until DelphiLSP has pushed diagnostics for AFile again,
///  i.e. its per-file push version moved past ABefore (take it with
///  GetFileDiagnosticsVersion BEFORE sending). The push comes when the
///  analysis is done - also with zero diagnostics. False on timeout/stop.
///  </summary>
function McpWaitLspAnalysed(AClient: TLspClient; const AFile: string;
  ABefore: Integer; AStop: THandle; ATimeoutMs: Cardinal = 60000): Boolean;

/// <summary>Milliseconds since this bridge last SENT AFile to AClient's
///  session, -1 when never - lets a caller give a freshly sent file a short
///  grace period before taking an empty answer at face value.</summary>
function McpLspSentAgoMs(AClient: TLspClient; const AFile: string): Int64;

implementation

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.JSON,
  System.Generics.Collections, System.SyncObjs,
  Expert.McpServer, Expert.EditorHelperIntf, Expert.LspManager,
  Lsp.Protocol, Lsp.Uri, Delphi.FileEncoding;

var
  GSyncLock: TCriticalSection = nil;
  GSentAt: TDictionary<string, UInt64> = nil;     // file -> tick of that send

function ArgStr(AArgs: TJSONObject; const AName: string; const ADefault: string = ''): string;
begin
  Result := ADefault;
  if AArgs <> nil then Result := AArgs.GetValue<string>(AName, ADefault);
end;

function ArgInt(AArgs: TJSONObject; const AName: string; ADefault: Integer = 0): Integer;
begin
  Result := ADefault;
  if AArgs <> nil then Result := AArgs.GetValue<Integer>(AName, ADefault);
end;

function ArgBool(AArgs: TJSONObject; const AName: string; ADefault: Boolean = False): Boolean;
begin
  Result := ADefault;
  if AArgs <> nil then Result := AArgs.GetValue<Boolean>(AName, ADefault);
end;

// Sends AContent unless the session already has exactly that content. The
// CLIENT remembers what it last sent per document (TLspClient
// .SyncDocumentWith) - shared with rename, safe delete, change signature
// and the live checker, so no second cache can drift from what the session
// really has; an LSP restart is a new client that knows no document.
function SyncOne(AClient: TLspClient; const AFile, AContent: string): Boolean;
begin
  Result := AClient.SyncDocumentWith(AFile, AContent);
  if not Result then Exit;
  var Key := UpperCase(AFile) + '|' + IntToHex(NativeInt(AClient), 8);
  GSyncLock.Enter;
  try
    GSentAt.AddOrSetValue(Key, GetTickCount64);
  finally
    GSyncLock.Leave;
  end;
end;

// The running client (never started from here) and, when AFile is given,
// the file synced into it.
function PrepareClient(const AFile: string; AStop: THandle; out AClient: TLspClient;
  out AError: string): Boolean;
var
  C: string;
  Found: Boolean;
  Client: TLspClient;
begin
  Result := False;
  Client := nil;
  Found := False;
  if not McpRunOnMain(
    procedure
    begin
      Client := TLspManager.Instance.PeekClient;
      if AFile <> '' then Found := McpReadContent(AFile, C);
    end, True, AStop, AError) then Exit;
  if Client = nil then
  begin
    AError := 'the plugin''s DelphiLSP session is not running yet - it starts ' +
      'with the first opened project (see get_status)';
    Exit;
  end;
  if AFile <> '' then
  begin
    if not Found then
    begin
      AError := 'file not found: ' + AFile;
      Exit;
    end;
    var Before := Client.GetFileDiagnosticsVersion(AFile);
    // a unit that was just (re)opened answers nothing until analysed
    if SyncOne(Client, AFile, C) then
      McpWaitLspAnalysed(Client, AFile, Before, AStop, 30000);   // bridge: 45 s
  end;
  AClient := Client;
  Result := True;
end;

function McpSyncLspContent(AClient: TLspClient; const AFile, AContent: string): Boolean;
begin
  Result := SyncOne(AClient, AFile, AContent);
end;

function McpLspSentAgoMs(AClient: TLspClient; const AFile: string): Int64;
var
  T: UInt64;
begin
  Result := -1;
  GSyncLock.Enter;
  try
    if GSentAt.TryGetValue(UpperCase(AFile) + '|' + IntToHex(NativeInt(AClient), 8), T) then
      Result := Int64(GetTickCount64 - T);
  finally
    GSyncLock.Leave;
  end;
end;

// Added to an EMPTY answer of the lsp_* tools when the file was sent a
// moment ago - DelphiLSP is silent in a unit while it analyses it.
procedure AddFreshNote(ARes: TJSONObject; AClient: TLspClient; const AFile: string);
begin
  var Ago := McpLspSentAgoMs(AClient, AFile);
  if (Ago >= 0) and (Ago < 60000) then
    ARes.AddPair('note', Format('%s was (re)sent to DelphiLSP %d s ago because ' +
      'its content changed - it may still be analysing it; retry shortly. If ' +
      'the unit stays silent, the usual cause is that a unit (or one it uses) is taken from a precompiled .dcu ' +
        'instead of its source - e.g. DCUs of the project''s own units on ' +
        'the IDE library path or in the DCU output directory. To see it: ' +
        'start the IDE with REFACTORINGLIGHT_LSP_ARGS=-LogModes 255 and read ' +
        '%%TEMP%%\DelphiLSP\Agent*.log.',
      [ExtractFileName(AFile), Ago div 1000]))
  else
    ARes.AddPair('note', 'DelphiLSP returned nothing. If the unit stays ' +
      'silent, the usual cause is that a unit (or one it uses) is taken from a precompiled .dcu ' +
        'instead of its source - e.g. DCUs of the project''s own units on ' +
        'the IDE library path or in the DCU output directory. To see it: ' +
        'start the IDE with REFACTORINGLIGHT_LSP_ARGS=-LogModes 255 and read ' +
        '%TEMP%\DelphiLSP\Agent*.log.');
end;

function McpWaitLspAnalysed(AClient: TLspClient; const AFile: string;
  ABefore: Integer; AStop: THandle; ATimeoutMs: Cardinal): Boolean;
begin
  var Deadline := GetTickCount64 + ATimeoutMs;
  repeat
    if AClient.GetFileDiagnosticsVersion(AFile) <> ABefore then Exit(True);
  until (WaitForSingleObject(AStop, 150) = WAIT_OBJECT_0) or
        (GetTickCount64 > Deadline);
  Result := False;
end;

function McpSyncLspDocuments(AClient: TLspClient; const AFiles: TArray<string>;
  AStop: THandle; out AError: string): Boolean;
var
  Contents: TArray<string>;
  Found: TArray<Boolean>;
begin
  Result := False;
  SetLength(Contents, Length(AFiles));
  SetLength(Found, Length(AFiles));
  if not McpRunOnMain(
    procedure
    begin
      for var I := 0 to High(AFiles) do
        Found[I] := McpReadContent(AFiles[I], Contents[I]);
    end, True, AStop, AError) then Exit;
  for var I := 0 to High(AFiles) do
  begin
    if WaitForSingleObject(AStop, 0) = WAIT_OBJECT_0 then
    begin
      AError := 'shutting down';
      Exit;
    end;
    if Found[I] then SyncOne(AClient, AFiles[I], Contents[I]);
  end;
  Result := True;
end;

function RequirePos(AArgs: TJSONObject; out AFile: string; out ALine0, ACol0: Integer;
  out AError: string): Boolean;
begin
  AFile := ArgStr(AArgs, 'file');
  ALine0 := ArgInt(AArgs, 'line') - 1;
  ACol0 := ArgInt(AArgs, 'column') - 1;
  Result := (AFile <> '') and (ALine0 >= 0) and (ACol0 >= 0);
  if Result then
    AFile := ExpandFileName(AFile)
  else
    AError := 'arguments "file", "line" and "column" (1-based) are required';
end;

function LocationsToJson(const ALocs: TArray<TLspLocation>): TJSONArray;
var
  Cache: TDictionary<string, TArray<string>>;
begin
  Result := TJSONArray.Create;
  Cache := TDictionary<string, TArray<string>>.Create;
  try
    for var L in ALocs do
    begin
      var F := TLspUri.FileUriToPath(L.Uri);
      var O := TJSONObject.Create;
      O.AddPair('file', F);
      O.AddPair('line', TJSONNumber.Create(L.Range.Start.Line + 1));
      O.AddPair('column', TJSONNumber.Create(L.Range.Start.Character + 1));
      var Lines: TArray<string>;
      if not Cache.TryGetValue(UpperCase(F), Lines) then
      begin
        try
          Lines := ReadDelphiFileLines(F);
        except
          Lines := nil;
        end;
        Cache.Add(UpperCase(F), Lines);
      end;
      if L.Range.Start.Line <= High(Lines) then
        O.AddPair('text', Trim(Lines[L.Range.Start.Line]));
      Result.Add(O);
    end;
  finally
    Cache.Free;
  end;
end;

function ToolHover(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  L0, C0: Integer;
  Client: TLspClient;
begin
  if not RequirePos(AArgs, F, L0, C0, Err) then Exit(McpErr(Err));
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var H := Client.GetHover(F, L0, C0);
  var Res := TJSONObject.Create;
  Res.AddPair('hover', H);
  if H = '' then AddFreshNote(Res, Client, F);
  if H = '' then Res.AddPair('note', 'DelphiLSP returned nothing for this position');
  Result := McpOk(Res);
end;

function ToolDefinition(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  L0, C0: Integer;
  Client: TLspClient;
begin
  if not RequirePos(AArgs, F, L0, C0, Err) then Exit(McpErr(Err));
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var Res := TJSONObject.Create;
  var Locs := Client.GotoDefinition(F, L0, C0);
  Res.AddPair('locations', LocationsToJson(Locs));
  if Length(Locs) = 0 then AddFreshNote(Res, Client, F);
  Result := McpOk(Res);
end;

function ToolImplementation(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  L0, C0: Integer;
  Client: TLspClient;
begin
  if not RequirePos(AArgs, F, L0, C0, Err) then Exit(McpErr(Err));
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var Res := TJSONObject.Create;
  Res.AddPair('locations', LocationsToJson(Client.GotoImplementation(F, L0, C0)));
  Res.AddPair('note', 'DelphiLSP''s own answer - for interface methods ' +
    'find_implementations (plugin scan) is usually more complete');
  Result := McpOk(Res);
end;

function ToolReferences(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  L0, C0: Integer;
  Client: TLspClient;
begin
  if not RequirePos(AArgs, F, L0, C0, Err) then Exit(McpErr(Err));
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  if not Client.SupportsReferences then
    Exit(McpErr('this DelphiLSP does not announce textDocument/references - ' +
      'use find_references (plugin scan with LSP verification)'));
  var Res := TJSONObject.Create;
  Res.AddPair('locations', LocationsToJson(
    Client.FindReferences(F, L0, C0, ArgBool(AArgs, 'include_declaration', True))));
  Result := McpOk(Res);
end;

function ToolDocumentSymbols(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  Client: TLspClient;
begin
  F := ArgStr(AArgs, 'file');
  if F = '' then Exit(McpErr('argument "file" is required'));
  F := ExpandFileName(F);
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var Arr := Client.GetDocumentSymbols(F, 20000);
  if Arr = nil then Arr := TJSONArray.Create;
  var Res := TJSONObject.Create;
  Res.AddPair('symbols', Arr);
  Res.AddPair('note', 'raw LSP DocumentSymbol[] - ranges are 0-based');
  Result := McpOk(Res);
end;

function RawRequest(AClient: TLspClient; const AMethod: string; AParams: TJSONValue;
  ATimeoutMs: Cardinal): string;
var
  Resp: TJSONObject;
begin
  Resp := AClient.SendRequest(AMethod, AParams, ATimeoutMs);
  try
    if Resp = nil then Exit(McpErr('no answer from DelphiLSP within the timeout'));
    var Err := Resp.GetValue('error');
    if Err <> nil then Exit(McpErr('DelphiLSP error: ' + Err.ToJSON));
    var R := Resp.GetValue('result');
    var Res := TJSONObject.Create;
    if R <> nil then
      Res.AddPair('result', R.Clone as TJSONValue)
    else
      Res.AddPair('result', TJSONNull.Create);
    Result := McpOk(Res);
  finally
    Resp.Free;
  end;
end;

// (no lsp_workspace_symbols: DelphiLSP answers workspace/symbol with
// "Method not found" - find_unit covers that search via the index.)

function ToolSignatureHelp(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  L0, C0: Integer;
  Client: TLspClient;
begin
  if not RequirePos(AArgs, F, L0, C0, Err) then Exit(McpErr(Err));
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var R := Client.GetSignatureHelp(F, L0, C0);
  var Res := TJSONObject.Create;
  if R <> nil then Res.AddPair('signatureHelp', R) else Res.AddPair('signatureHelp', TJSONNull.Create);
  Res.AddPair('note', 'raw LSP SignatureHelp (activeParameter/activeSignature 0-based)');
  Result := McpOk(Res);
end;

function ToolCompletion(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  L0, C0: Integer;
  Client: TLspClient;
begin
  if not RequirePos(AArgs, F, L0, C0, Err) then Exit(McpErr(Err));
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var MaxN := ArgInt(AArgs, 'max', 100);
  var Prefix := UpperCase(ArgStr(AArgs, 'prefix'));
  var R := Client.GetCompletion(F, L0, C0);
  try
    var Arr := TJSONArray.Create;
    var Total := 0;
    var ResV: TJSONValue := nil;
    if R <> nil then ResV := R.GetValue('result');
    var Items: TJSONArray := nil;
    if ResV is TJSONObject then Items := TJSONObject(ResV).GetValue('items') as TJSONArray
    else if ResV is TJSONArray then Items := TJSONArray(ResV);
    if Items <> nil then
      for var It in Items do
      begin
        if not (It is TJSONObject) then Continue;
        var Lbl := TJSONObject(It).GetValue<string>('label', '');
        if (Prefix <> '') and not UpperCase(Lbl).StartsWith(Prefix) then Continue;
        Inc(Total);
        if Arr.Count >= MaxN then Continue;
        var O := TJSONObject.Create;
        O.AddPair('label', Lbl);
        var D := TJSONObject(It).GetValue<string>('detail', '');
        if D <> '' then O.AddPair('detail', D);
        O.AddPair('kind', TJSONNumber.Create(TJSONObject(It).GetValue<Integer>('kind', 0)));
        Arr.Add(O);
      end;
    var Res := TJSONObject.Create;
    Res.AddPair('items', Arr);
    Res.AddPair('total', TJSONNumber.Create(Total));
    Res.AddPair('note', 'kind = LSP CompletionItemKind (2 method, 3 function, 5 field, ' +
      '6 variable, 7 class, 8 interface, 10 property, 13 enum, 14 keyword, 21 constant, ' +
      '22 struct)');
    Result := McpOk(Res);
  finally
    R.Free;
  end;
end;

const
  // Would break the session the plugin itself relies on.
  RefusedMethods: array[0..3] of string = ('initialize', 'shutdown', 'exit', 'initialized');

function ToolLspRequest(AArgs: TJSONObject; AStop: THandle): string;
var
  Method, F, Err: string;
  Client: TLspClient;
begin
  Method := Trim(ArgStr(AArgs, 'method'));
  if Method = '' then Exit(McpErr('argument "method" is required'));
  for var M in RefusedMethods do
    if SameText(Method, M) then
      Exit(McpErr('"' + Method + '" would break the plugin''s LSP session - refused'));
  if Method.StartsWith('textDocument/did', True) or Method.StartsWith('$/') or
     Method.StartsWith('workspace/did', True) then
    Exit(McpErr('notifications and document-sync methods are refused - the bridge ' +
      'keeps documents in sync itself (pass "file" to sync one)'));
  F := ArgStr(AArgs, 'file');
  if F <> '' then F := ExpandFileName(F);
  if not PrepareClient(F, AStop, Client, Err) then Exit(McpErr(Err));
  var P: TJSONValue;
  if (AArgs <> nil) and (AArgs.GetValue('params') <> nil) then
    P := AArgs.GetValue('params').Clone as TJSONValue
  else
    P := TJSONObject.Create;
  Result := RawRequest(Client, Method, P, Cardinal(ArgInt(AArgs, 'timeout_ms', 20000)));
end;

initialization
  GSyncLock := TCriticalSection.Create;
  GSentAt := TDictionary<string, UInt64>.Create;
  RegisterMcpTool('lsp_hover', ToolHover);
  RegisterMcpTool('lsp_definition', ToolDefinition);
  RegisterMcpTool('lsp_implementation', ToolImplementation);
  RegisterMcpTool('lsp_references', ToolReferences);
  RegisterMcpTool('lsp_document_symbols', ToolDocumentSymbols);
  RegisterMcpTool('lsp_signature_help', ToolSignatureHelp);
  RegisterMcpTool('lsp_completion', ToolCompletion);
  RegisterMcpTool('lsp_request', ToolLspRequest);

finalization
  FreeAndNil(GSentAt);
  FreeAndNil(GSyncLock);

end.
