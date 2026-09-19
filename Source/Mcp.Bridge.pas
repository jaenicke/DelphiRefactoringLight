(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Mcp.Bridge;

// The MCP server logic of RefactoringLightMcp.exe (stdio, JSON-RPC 2.0, one
// message per line). Kept out of the .dpr so the console tests can drive
// it message by message; the transport to the IDEs is injectable.
//
// The bridge answers ide_instances / select_ide itself and forwards every
// other tool to ONE IDE, chosen per call:
//   1. the "instance" argument (a pid),
//   2. the session's pin from select_ide,
//   3. PickInstance: the file argument, then the working directory against
//      each IDE's projects - an ambiguous choice is an error, never a guess.
// Every forwarded result names the IDE it came from ("ide"), so a wrong
// pick is visible instead of silently confusing.
//
// THE TOOL LIST COMES FROM THE IDEs. Only ide_instances / select_ide are
// the bridge's own; everything else is whatever the running IDEs serve
// ("tools" request), merged by name, newest plugin build first. The list
// compiled into the bridge is the fallback when no IDE runs or an IDE's
// plugin predates the "tools" request. So a new IDE tool needs a new
// PLUGIN only - not a new bridge. A running Claude session learns about a
// changed set through notifications/tools/list_changed: the exe's watcher
// thread calls CheckToolsChanged, which compares the IDEs' tool hashes
// with the ones behind the list last handed out.

interface

uses
  System.SysUtils, System.JSON, Mcp.Protocol;

type
  IMcpTransport = interface
    ['{6A0C1D7E-2B4F-4E53-9F1A-7C3D5E9B1A20}']
    function ListPids: TArray<Cardinal>;
    function Request(APid: Cardinal; const ARequest: string; ATimeoutMs: Cardinal;
      out AResponse, AError: string): Boolean;
  end;

  TMcpBridge = class
  private
    FTransport: IMcpTransport;
    FCwd: string;
    FPinned: Cardinal;
    FLock: TObject;
    FAnnounced: string;       // ToolsSignature behind the last tools/list
    FAnnouncedSet: Boolean;
    function ToolsSignature(const ACtxs: TArray<TMcpInstanceContext>): string;
    function EffectiveTools(const ACtxs: TArray<TMcpInstanceContext>): TJSONArray;
    function QueryContexts(out AErrors: string): TArray<TMcpInstanceContext>;
    function ToolResult(const AText: string; AIsError: Boolean): TJSONObject;
    function CallIdeInstances: TJSONObject;
    function CallSelectIde(AArgs: TJSONObject): TJSONObject;
    function CallForward(const ATool: string; AArgs: TJSONObject): TJSONObject;
    function HandleRequest(const AMethod: string; AParams: TJSONObject;
      out AResult: TJSONValue; out AErrCode: Integer; out AErrMsg: string): Boolean;
  public
    constructor Create(const ACwd: string; const ATransport: IMcpTransport);
    destructor Destroy; override;
    /// <summary>True (once per change) when the tool set of the running IDEs
    ///  differs from the one behind the last tools/list - the caller then
    ///  sends notifications/tools/list_changed. Thread-safe.</summary>
    function CheckToolsChanged: Boolean;
    /// <summary>One JSON-RPC message in; the response line out, or '' for
    ///  notifications.</summary>
    function HandleLine(const ALine: string): string;
    property Cwd: string read FCwd;
  end;

  /// <summary>The real transport: named pipes (Mcp.Protocol).</summary>
  TPipeTransport = class(TInterfacedObject, IMcpTransport)
  public
    function ListPids: TArray<Cardinal>;
    function Request(APid: Cardinal; const ARequest: string; ATimeoutMs: Cardinal;
      out AResponse, AError: string): Boolean;
  end;

const
  BridgeVersion = '1.2.0';

implementation

uses
  System.Generics.Collections, System.Generics.Defaults;

const
  ContextTimeoutMs = 5000;
  ToolTimeoutMs = 45000;
  SupportedProtocols: array[0..2] of string = ('2025-06-18', '2025-03-26', '2024-11-05');

{ TPipeTransport }

function TPipeTransport.ListPids: TArray<Cardinal>;
begin
  Result := ListMcpPipePids;
end;

function TPipeTransport.Request(APid: Cardinal; const ARequest: string;
  ATimeoutMs: Cardinal; out AResponse, AError: string): Boolean;
begin
  Result := McpPipeRequest(APid, ARequest, ATimeoutMs, AResponse, AError);
end;

// Rename, reference scans and the project-wide analyses can take minutes
// on a big project; everything else answers in seconds.
function ToolTimeout(const AName: string): Cardinal;
begin
  if AName.StartsWith('rename_') or (AName = 'find_references') or
     (AName = 'find_implementations') or AName.StartsWith('uses_') or
     (AName = 'analyze_uses') or (AName = 'debug_consistency') or
     (AName = 'blame') or (AName = 'commit_info') or (AName = 'safe_delete') then
    Result := 300000
  else
    Result := ToolTimeoutMs;
end;

function IsLocalTool(const AName: string): Boolean;
begin
  Result := (AName = 'ide_instances') or (AName = 'select_ide');
end;

// Everything but the bridge's own two tools goes to an IDE - which one
// knows it is decided there (an older plugin answers "unknown tool").
function IsForwardedTool(const AName: string): Boolean;
begin
  Result := (AName <> '') and not IsLocalTool(AName);
end;

{ TMcpBridge }

constructor TMcpBridge.Create(const ACwd: string; const ATransport: IMcpTransport);
begin
  inherited Create;
  FCwd := ACwd;
  FTransport := ATransport;
  FLock := TObject.Create;
end;

destructor TMcpBridge.Destroy;
begin
  FLock.Free;
  inherited;
end;

function TMcpBridge.ToolsSignature(const ACtxs: TArray<TMcpInstanceContext>): string;
var
  Hashes: TArray<string>;
  H: string;
begin
  Hashes := nil;
  for var C in ACtxs do
  begin
    H := C.ToolsHash;
    if H = '' then H := McpToolsHash;   // old plugin: the bridge's own list
    var Dup := False;
    for var X in Hashes do
      if X = H then Dup := True;
    if not Dup then Hashes := Hashes + [H];
  end;
  if Length(Hashes) = 0 then Exit(McpToolsHash);
  TArray.Sort<string>(Hashes);
  Result := string.Join(',', Hashes);
end;

function TMcpBridge.EffectiveTools(const ACtxs: TArray<TMcpInstanceContext>): TJSONArray;
var
  Seen: TArray<string>;
  Order: TArray<TMcpInstanceContext>;
  NeedStatic: Boolean;

  function Has(const AName: string): Boolean;
  begin
    for var S in Seen do
      if S = AName then Exit(True);
    Result := False;
  end;

  procedure AddFrom(AArr: TJSONArray; ALocal: Boolean);
  begin
    for var D in AArr do
      if D is TJSONObject then
      begin
        var N := TJSONObject(D).GetValue<string>('name', '');
        if (N = '') or Has(N) or (IsLocalTool(N) <> ALocal) then Continue;
        Seen := Seen + [N];
        Result.AddElement(D.Clone as TJSONValue);
      end;
  end;

var
  Static: TJSONArray;
  Resp, Err: string;
begin
  Result := TJSONArray.Create;
  Seen := nil;
  Static := McpToolDefinitions;
  try
    AddFrom(Static, True);     // the bridge's own two first
    // newest plugin build first: its description of a tool wins
    Order := Copy(ACtxs);
    TArray.Sort<TMcpInstanceContext>(Order, TComparer<TMcpInstanceContext>.Construct(
      function(const A, B: TMcpInstanceContext): Integer
      begin
        Result := -CompareStr(A.PluginBuild, B.PluginBuild);
      end));
    NeedStatic := Length(Order) = 0;
    for var C in Order do
    begin
      if C.ToolsHash = '' then
      begin
        NeedStatic := True;
        Continue;
      end;
      if not FTransport.Request(C.Pid, '{"method":"tools"}', ContextTimeoutMs, Resp, Err) then
      begin
        NeedStatic := True;
        Continue;
      end;
      var V := TJSONObject.ParseJSONValue(Resp);
      try
        var Arr: TJSONArray := nil;
        if (V is TJSONObject) and TJSONObject(V).GetValue<Boolean>('ok', False) then
          Arr := TJSONObject(V).FindValue('result.tools') as TJSONArray;
        if Arr <> nil then
          AddFrom(Arr, False)
        else
          NeedStatic := True;
      finally
        V.Free;
      end;
    end;
    if NeedStatic then AddFrom(Static, False);
  finally
    Static.Free;
  end;
end;

function TMcpBridge.CheckToolsChanged: Boolean;
var
  Errors, Sig: string;
begin
  Result := False;
  Sig := ToolsSignature(QueryContexts(Errors));
  TMonitor.Enter(FLock);
  try
    if FAnnouncedSet and (Sig <> FAnnounced) then
    begin
      FAnnounced := Sig;   // one notification per change
      Result := True;
    end;
  finally
    TMonitor.Exit(FLock);
  end;
end;

function TMcpBridge.QueryContexts(out AErrors: string): TArray<TMcpInstanceContext>;
var
  Resp, Err: string;
  V: TJSONValue;
begin
  Result := nil;
  AErrors := '';
  for var Pid in FTransport.ListPids do
  begin
    if not FTransport.Request(Pid, '{"method":"context"}', ContextTimeoutMs, Resp, Err) then
    begin
      AErrors := AErrors + Format('pid %d: %s', [Pid, Err]) + sLineBreak;
      Continue;
    end;
    V := TJSONObject.ParseJSONValue(Resp);
    try
      if (V is TJSONObject) and TJSONObject(V).GetValue<Boolean>('ok', False) and
         (TJSONObject(V).GetValue('result') is TJSONObject) then
        Result := Result + [TMcpInstanceContext.FromJson(
          TJSONObject(TJSONObject(V).GetValue('result')))]
      else
        AErrors := AErrors + Format('pid %d: unexpected answer', [Pid]) + sLineBreak;
    finally
      V.Free;
    end;
  end;
end;

function TMcpBridge.ToolResult(const AText: string; AIsError: Boolean): TJSONObject;
var
  Content: TJSONArray;
  Item: TJSONObject;
begin
  Result := TJSONObject.Create;
  Content := TJSONArray.Create;
  Item := TJSONObject.Create;
  Item.AddPair('type', 'text');
  Item.AddPair('text', AText);
  Content.Add(Item);
  Result.AddPair('content', Content);
  Result.AddPair('isError', TJSONBool.Create(AIsError));
end;

function TMcpBridge.CallIdeInstances: TJSONObject;
var
  Ctxs: TArray<TMcpInstanceContext>;
  Errors, Reason: string;
  Idx: Integer;
  Out_: TJSONObject;
  Arr: TJSONArray;
begin
  Ctxs := QueryContexts(Errors);
  Out_ := TJSONObject.Create;
  try
    Out_.AddPair('workingDirectory', FCwd);
    Arr := TJSONArray.Create;
    for var C in Ctxs do Arr.Add(C.ToJson);
    Out_.AddPair('instances', Arr);
    if FPinned <> 0 then
      Out_.AddPair('selection', Format('pinned to pid %d via select_ide', [FPinned]))
    else if PickInstance(Ctxs, FCwd, '', Idx, Reason) then
      Out_.AddPair('selection', Format('automatic: pid %d (%s)', [Ctxs[Idx].Pid, Reason]))
    else
      Out_.AddPair('selection', 'automatic, currently undecided: ' + Reason);
    if Errors <> '' then Out_.AddPair('unreachable', Trim(Errors));
    Result := ToolResult(Out_.Format(2), False);
  finally
    Out_.Free;
  end;
end;

function TMcpBridge.CallSelectIde(AArgs: TJSONObject): TJSONObject;
var
  Pid: Cardinal;
  Errors: string;
begin
  Pid := 0;
  if AArgs <> nil then Pid := AArgs.GetValue<Cardinal>('instance', 0);
  if Pid = 0 then
  begin
    FPinned := 0;
    Exit(ToolResult('Automatic IDE selection is active again.', False));
  end;
  for var C in QueryContexts(Errors) do
    if C.Pid = Pid then
    begin
      FPinned := Pid;
      Exit(ToolResult('This session now uses ' + C.Describe + '.', False));
    end;
  Result := ToolResult(Format('No running IDE with Refactoring Light has pid %d. ' +
    'Call ide_instances for the list.', [Pid]), True);
end;

function TMcpBridge.CallForward(const ATool: string; AArgs: TJSONObject): TJSONObject;
var
  Ctxs: TArray<TMcpInstanceContext>;
  Target: TMcpInstanceContext;
  Errors, Reason, FileArg, Resp, Err: string;
  Pid: Cardinal;
  Idx: Integer;
  Req: TJSONObject;
  V: TJSONValue;
begin
  Ctxs := QueryContexts(Errors);
  Pid := 0;
  FileArg := '';
  if AArgs <> nil then
  begin
    Pid := AArgs.GetValue<Cardinal>('instance', 0);
    FileArg := AArgs.GetValue<string>('file', '');
  end;
  Target := Default(TMcpInstanceContext);
  if Pid = 0 then Pid := FPinned;
  if Pid <> 0 then
  begin
    for var C in Ctxs do
      if C.Pid = Pid then Target := C;
    if Target.Pid = 0 then
      Exit(ToolResult(Format('The IDE with pid %d is not reachable (closed?). ' +
        'Call ide_instances, then select_ide.', [Pid]), True));
    if Pid = FPinned then Reason := 'pinned via select_ide' else Reason := 'instance argument';
  end
  else
  begin
    if not PickInstance(Ctxs, FCwd, FileArg, Idx, Reason) then
    begin
      if Errors <> '' then Reason := Reason + sLineBreak + 'Unreachable: ' + Trim(Errors);
      Exit(ToolResult(Reason, True));
    end;
    Target := Ctxs[Idx];
  end;

  Req := TJSONObject.Create;
  try
    Req.AddPair('method', 'call');
    Req.AddPair('tool', ATool);
    if AArgs <> nil then
      Req.AddPair('arguments', AArgs.Clone as TJSONObject)
    else
      Req.AddPair('arguments', TJSONObject.Create);
    if not FTransport.Request(Target.Pid, Req.ToJSON, ToolTimeout(ATool), Resp, Err) then
      Exit(ToolResult(Target.Describe + ': ' + Err, True));
  finally
    Req.Free;
  end;

  V := TJSONObject.ParseJSONValue(Resp);
  try
    if not (V is TJSONObject) then
      Exit(ToolResult('Malformed answer from the IDE.', True));
    if not TJSONObject(V).GetValue<Boolean>('ok', False) then
      Exit(ToolResult(TJSONObject(V).GetValue<string>('error', 'unknown error') +
        sLineBreak + '(IDE: ' + Target.Describe + ')', True));
    var R := TJSONObject(V).GetValue('result');
    if R is TJSONObject then
    begin
      var Ide := TJSONObject.Create;
      Ide.AddPair('pid', TJSONNumber.Create(Target.Pid));
      Ide.AddPair('description', Target.Describe);
      Ide.AddPair('selectedBy', Reason);
      TJSONObject(R).AddPair('ide', Ide);
      Result := ToolResult(TJSONObject(R).Format(2), False);
    end
    else if R <> nil then
      Result := ToolResult(R.ToJSON, False)
    else
      Result := ToolResult('{}', False);
  finally
    V.Free;
  end;
end;

function TMcpBridge.HandleRequest(const AMethod: string; AParams: TJSONObject;
  out AResult: TJSONValue; out AErrCode: Integer; out AErrMsg: string): Boolean;
begin
  Result := True;
  AResult := nil;
  AErrCode := 0;
  AErrMsg := '';
  if AMethod = 'initialize' then
  begin
    var Wanted := '';
    if AParams <> nil then Wanted := AParams.GetValue<string>('protocolVersion', '');
    var Version := SupportedProtocols[0];
    for var S in SupportedProtocols do
      if S = Wanted then Version := S;
    var R := TJSONObject.Create;
    R.AddPair('protocolVersion', Version);
    var Caps := TJSONObject.Create;
    var ToolCaps := TJSONObject.Create;
    ToolCaps.AddPair('listChanged', TJSONBool.Create(True));
    Caps.AddPair('tools', ToolCaps);
    R.AddPair('capabilities', Caps);
    var Info := TJSONObject.Create;
    Info.AddPair('name', McpServerName);
    Info.AddPair('version', BridgeVersion);
    R.AddPair('serverInfo', Info);
    R.AddPair('instructions', McpServerInstructions);
    AResult := R;
  end
  else if AMethod = 'ping' then
    AResult := TJSONObject.Create
  else if AMethod = 'tools/list' then
  begin
    var Errors: string;
    var Ctxs := QueryContexts(Errors);
    var R := TJSONObject.Create;
    R.AddPair('tools', EffectiveTools(Ctxs));
    AResult := R;
    TMonitor.Enter(FLock);
    try
      FAnnounced := ToolsSignature(Ctxs);
      FAnnouncedSet := True;
    finally
      TMonitor.Exit(FLock);
    end;
  end
  else if AMethod = 'tools/call' then
  begin
    var Name := '';
    var Args: TJSONObject := nil;
    if AParams <> nil then
    begin
      Name := AParams.GetValue<string>('name', '');
      if AParams.GetValue('arguments') is TJSONObject then
        Args := TJSONObject(AParams.GetValue('arguments'));
    end;
    if Name = 'ide_instances' then AResult := CallIdeInstances
    else if Name = 'select_ide' then AResult := CallSelectIde(Args)
    else if IsForwardedTool(Name) then AResult := CallForward(Name, Args)
    else
    begin
      AErrCode := -32602;
      AErrMsg := 'unknown tool: ' + Name;
      Result := False;
    end;
  end
  else
  begin
    AErrCode := -32601;
    AErrMsg := 'method not found: ' + AMethod;
    Result := False;
  end;
end;

function TMcpBridge.HandleLine(const ALine: string): string;
var
  V: TJSONValue;
  Msg, Resp, Err: TJSONObject;
  Method, ErrMsg: string;
  Id: TJSONValue;
  Res: TJSONValue;
  Code: Integer;
  Params: TJSONObject;
begin
  Result := '';
  if Trim(ALine) = '' then Exit;
  V := TJSONObject.ParseJSONValue(ALine);
  try
    if not (V is TJSONObject) then
    begin
      Result := '{"jsonrpc":"2.0","id":null,"error":{"code":-32700,"message":"parse error"}}';
      Exit;
    end;
    Msg := TJSONObject(V);
    Method := Msg.GetValue<string>('method', '');
    Id := Msg.GetValue('id');
    // notifications (initialized, cancelled, ...) and client responses
    // get no answer
    if (Id = nil) or (Method = '') then Exit;
    Params := nil;
    if Msg.GetValue('params') is TJSONObject then
      Params := TJSONObject(Msg.GetValue('params'));

    Resp := TJSONObject.Create;
    try
      Resp.AddPair('jsonrpc', '2.0');
      Resp.AddPair('id', Id.Clone as TJSONValue);
      try
        if HandleRequest(Method, Params, Res, Code, ErrMsg) then
          Resp.AddPair('result', Res)
        else
        begin
          Err := TJSONObject.Create;
          Err.AddPair('code', TJSONNumber.Create(Code));
          Err.AddPair('message', ErrMsg);
          Resp.AddPair('error', Err);
        end;
      except
        on E: Exception do
        begin
          Err := TJSONObject.Create;
          Err.AddPair('code', TJSONNumber.Create(-32603));
          Err.AddPair('message', E.ClassName + ': ' + E.Message);
          Resp.AddPair('error', Err);
        end;
      end;
      Result := Resp.ToJSON;
    finally
      Resp.Free;
    end;
  finally
    V.Free;
  end;
end;

end.
