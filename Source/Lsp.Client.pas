(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Lsp.Client;

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.SyncObjs, System.IOUtils, System.Generics.Collections, Winapi.Windows,
  Delphi.FileEncoding, Lsp.JsonRpc, Lsp.Protocol, Lsp.Uri;

type
  ELspError = class(Exception)
  public
    Code: Integer;
    constructor Create(ACode: Integer; const AMsg: string);
  end;

  ELspTimeout = class(Exception);

  TLspLogEvent = reference to procedure(const ADirection, AMethod, ABody: string);

  /// <summary>
  ///  Manages the DelphiLSP subprocess and the LSP communication.
  ///  Starts a background thread to read server responses.
  /// </summary>
  TLspClient = class
  private type
    TPendingRequest = class
      Event: TLightweightEvent;
      Response: TJSONObject;
      constructor Create;
      destructor Destroy; override;
    end;

    TReaderThread = class(TThread)
    private
      FOwner: TLspClient;
    protected
      procedure Execute; override;
    end;
  private
    FLspExePath: string;
    FProcessHandle: THandle;
    FStdinWrite: THandle;
    FStdoutRead: THandle;
    FStdinStream: THandleStream;
    FStdoutStream: THandleStream;
    FTransport: TJsonRpcTransport;
    FReaderThread: TReaderThread;
    FNextId: Integer;
    FPending: TObjectDictionary<Integer, TPendingRequest>;
    FPendingLock: TCriticalSection;
    FOnLog: TLspLogEvent;
    FVerbose: Boolean;
    FServerCapabilities: TJSONObject;

    function NextRequestId: Integer;
    procedure DispatchResponse(AMsg: TJSONObject);
    procedure HandleServerRequest(AMsg: TJSONObject);
    procedure Log(const ADirection, AMethod, ABody: string);
  public
    constructor Create(const ALspExePath: string);
    destructor Destroy; override;

    /// <summary>Starts the LSP server as a subprocess.</summary>
    procedure Start;

    /// <summary>Sends a request and waits for the response.</summary>
    function SendRequest(const AMethod: string; AParams: TJSONValue; ATimeoutMs: Cardinal = 60000): TJSONObject;

    /// <summary>Sends a request without waiting for the response.
    ///  Returns the request ID.</summary>
    function SendRequestAsync(const AMethod: string; AParams: TJSONValue): Integer;

    /// <summary>Waits for the response of a request previously sent via SendRequestAsync.</summary>
    function WaitForResponse(ARequestId: Integer; ATimeoutMs: Cardinal = 60000): TJSONObject;

    /// <summary>Sends a notification (no response expected).</summary>
    procedure SendNotification(const AMethod: string; AParams: TJSONValue = nil);

    /// <summary>Initializes the LSP session.</summary>
    procedure Initialize(const ARootPath, ADprojPath: string; const ASearchPath: string = '');

    /// <summary>Sends the project configuration from a .delphilsp.json file.</summary>
    procedure SendConfiguration(const ADelphiLspJsonPath: string);

    /// <summary>Opens a document on the LSP server.</summary>
    procedure OpenDocument(const AFilePath: string);

    /// <summary>Closes a document on the LSP server (didClose).</summary>
    procedure CloseDocument(const AFilePath: string);

    /// <summary>Refreshes a document at the LSP (didClose + didOpen).</summary>
    procedure RefreshDocument(const AFilePath: string);

    /// <summary>Checks whether a rename is possible at the position.</summary>
    function PrepareRename(const AFilePath: string; ALine, ACol: Integer): TLspPrepareRenameResult;

    /// <summary>Performs a rename and returns the WorkspaceEdit.</summary>
    function Rename(const AFilePath: string; ALine, ACol: Integer;
      const ANewName: string): TLspWorkspaceEdit;

    /// <summary>Finds all references of an identifier (incl. declaration).</summary>
    function FindReferences(const AFilePath: string; ALine, ACol: Integer;
      AIncludeDeclaration: Boolean = True): TArray<TLspLocation>;

    /// <summary>Jumps to the definition of the identifier at the position.</summary>
    function GotoDefinition(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>;

    /// <summary>Finds implementations (e.g. class methods that implement an interface).</summary>
    function GotoImplementation(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>;

    /// <summary>Gets hover info (type, description) for the identifier.</summary>
    function GetHover(const AFilePath: string; ALine, ACol: Integer): string;

    /// <summary>Requests code completion. Returns a JSON array of items.</summary>
    function GetCompletion(const AFilePath: string; ALine, ACol: Integer): TJSONObject;

    /// <summary>Returns the server capabilities as a JSON string (debugging).</summary>
    function GetServerCapabilities: string;

    /// <summary>Shuts the server down cleanly.</summary>
    procedure Shutdown;

    /// <summary>Indicates whether the server supports prepareRename.</summary>
    function SupportsPrepareRename: Boolean;

    /// <summary>Indicates whether the server supports references.</summary>
    function SupportsReferences: Boolean;

    property Verbose: Boolean read FVerbose write FVerbose;
    property OnLog: TLspLogEvent read FOnLog write FOnLog;
  end;

implementation

{ ELspError }

constructor ELspError.Create(ACode: Integer; const AMsg: string);
begin
  inherited Create(AMsg);
  Code := ACode;
end;

{ TLspClient.TPendingRequest }

constructor TLspClient.TPendingRequest.Create;
begin
  inherited;
  Event := TLightweightEvent.Create;
  Response := nil;
end;

destructor TLspClient.TPendingRequest.Destroy;
begin
  Event.Free;
  Response.Free;
  inherited;
end;

{ TLspClient.TReaderThread }

procedure TLspClient.TReaderThread.Execute;
var
  Msg: TJSONObject;
begin
  while not Terminated do
  begin
    try
      Msg := FOwner.FTransport.ReadMessage;
      FOwner.DispatchResponse(Msg);
    except
      on E: EStreamError do
      begin
        if not Terminated then
          FOwner.Log('<--', 'ERROR', E.Message);
        Exit;
      end;
    end;
  end;
end;

{ TLspClient }

constructor TLspClient.Create(const ALspExePath: string);
begin
  inherited Create;
  FLspExePath := ALspExePath;
  FNextId := 0;
  FPending := TObjectDictionary<Integer, TPendingRequest>.Create([doOwnsValues]);
  FPendingLock := TCriticalSection.Create;
  FProcessHandle := INVALID_HANDLE_VALUE;
  FStdinWrite := INVALID_HANDLE_VALUE;
  FStdoutRead := INVALID_HANDLE_VALUE;
end;

destructor TLspClient.Destroy;
begin
  if (FReaderThread <> nil) and not FReaderThread.Finished then
  begin
    FReaderThread.Terminate;
    // Closing the pipe unblocks the pending Read
    if FStdoutRead <> INVALID_HANDLE_VALUE then
      CloseHandle(FStdoutRead);
    FStdoutRead := INVALID_HANDLE_VALUE;
    FReaderThread.WaitFor;
  end;
  FreeAndNil(FReaderThread);
  FreeAndNil(FTransport);
  FreeAndNil(FStdinStream);
  FreeAndNil(FStdoutStream);

  if FStdinWrite <> INVALID_HANDLE_VALUE then
    CloseHandle(FStdinWrite);
  if FStdoutRead <> INVALID_HANDLE_VALUE then
    CloseHandle(FStdoutRead);
  if FProcessHandle <> INVALID_HANDLE_VALUE then
  begin
    TerminateProcess(FProcessHandle, 1);
    CloseHandle(FProcessHandle);
  end;

  FreeAndNil(FServerCapabilities);
  FPending.Free;
  FPendingLock.Free;
  inherited;
end;

function TLspClient.NextRequestId: Integer;
begin
  Result := TInterlocked.Increment(FNextId);
end;

procedure TLspClient.Log(const ADirection, AMethod, ABody: string);
begin
  if FVerbose and Assigned(FOnLog) then
    FOnLog(ADirection, AMethod, ABody);
end;

procedure TLspClient.DispatchResponse(AMsg: TJSONObject);
var
  IdValue: TJSONValue;
  Id: Integer;
  Pending: TPendingRequest;
begin
  IdValue := AMsg.GetValue('id');

  // Server notification (no id)
  if IdValue = nil then
  begin
    var Method := AMsg.GetValue<string>('method', '');
    Log('<--', Method, AMsg.ToJSON);
    AMsg.Free;
    Exit;
  end;

  // Server request (has method AND id)
  if AMsg.GetValue('method') <> nil then
  begin
    HandleServerRequest(AMsg);
    Exit;
  end;

  // Response to one of our requests
  Id := IdValue.AsType<Integer>;
  Log('<--', 'response#' + IntToStr(Id), AMsg.ToJSON);

  FPendingLock.Enter;
  try
    if FPending.TryGetValue(Id, Pending) then
    begin
      Pending.Response := AMsg; // Ownership passes to TPendingRequest
      Pending.Event.SetEvent;
    end
    else
      AMsg.Free;
  finally
    FPendingLock.Leave;
  end;
end;

procedure TLspClient.HandleServerRequest(AMsg: TJSONObject);
var
  Id: Integer;
  Response: TJSONObject;
begin
  // Respond to server requests (e.g. window/workDoneProgress/create) with an empty result
  Id := AMsg.GetValue<Integer>('id');
  var Method := AMsg.GetValue<string>('method', '');
  Log('<--', 'server-request: ' + Method, AMsg.ToJSON);

  Response := TJSONObject.Create;
  Response.AddPair('jsonrpc', '2.0');
  Response.AddPair('id', TJSONNumber.Create(Id));
  Response.AddPair('result', TJSONNull.Create);

  FTransport.SendMessage(Response);
  Response.Free;
  AMsg.Free;
end;

procedure TLspClient.Start;
var
  SA: TSecurityAttributes;
  SI: TStartupInfo;
  PI: TProcessInformation;
  hStdinRead, hStdoutWrite: THandle;
  CmdLine: string;
begin
  if not FileExists(FLspExePath) then
    raise EFileNotFoundException.Create('DelphiLsp.exe nicht gefunden: ' + FLspExePath);

  // Create anonymous pipes
  SA.nLength := SizeOf(SA);
  SA.bInheritHandle := True;
  SA.lpSecurityDescriptor := nil;

  if not CreatePipe(hStdinRead, FStdinWrite, @SA, 0) then
    RaiseLastOSError;
  if not CreatePipe(FStdoutRead, hStdoutWrite, @SA, 0) then
    RaiseLastOSError;

  // Make our ends non-inheritable
  SetHandleInformation(FStdinWrite, HANDLE_FLAG_INHERIT, 0);
  SetHandleInformation(FStdoutRead, HANDLE_FLAG_INHERIT, 0);

  // Start the process
  FillChar(SI, SizeOf(SI), 0);
  SI.cb := SizeOf(SI);
  SI.dwFlags := STARTF_USESTDHANDLES;
  SI.hStdInput := hStdinRead;
  SI.hStdOutput := hStdoutWrite;
  SI.hStdError := hStdoutWrite; // redirect stderr too

  CmdLine := '"' + FLspExePath + '"';

  if not CreateProcess(nil, PChar(CmdLine), nil, nil, True,
    CREATE_NO_WINDOW, nil, nil, SI, PI) then
    RaiseLastOSError;

  FProcessHandle := PI.hProcess;
  CloseHandle(PI.hThread);
  // Close child ends
  CloseHandle(hStdinRead);
  CloseHandle(hStdoutWrite);

  // Create streams and transport
  FStdinStream := THandleStream.Create(FStdinWrite);
  FStdoutStream := THandleStream.Create(FStdoutRead);
  FTransport := TJsonRpcTransport.Create(FStdoutStream, FStdinStream);

  // Start reader thread
  FReaderThread := TReaderThread.Create(True);
  FReaderThread.FreeOnTerminate := False;
  FReaderThread.FOwner := Self;
  FReaderThread.Start;
end;

function TLspClient.SendRequest(const AMethod: string; AParams: TJSONValue; ATimeoutMs: Cardinal): TJSONObject;
var
  Id: Integer;
  Msg: TJSONObject;
  Pending: TPendingRequest;
  WaitResult: TWaitResult;
  ErrorObj: TJSONObject;
begin
  Id := NextRequestId;

  Msg := TJSONObject.Create;
  Msg.AddPair('jsonrpc', '2.0');
  Msg.AddPair('id', TJSONNumber.Create(Id));
  Msg.AddPair('method', AMethod);
  if AParams <> nil then
    Msg.AddPair('params', AParams)
  else
    Msg.AddPair('params', TJSONObject.Create);

  Pending := TPendingRequest.Create;

  FPendingLock.Enter;
  try
    FPending.Add(Id, Pending);
  finally
    FPendingLock.Leave;
  end;

  Log('-->', AMethod, Msg.ToJSON);
  FTransport.SendMessage(Msg);
  Msg.Free;

  // Wait for response
  WaitResult := Pending.Event.WaitFor(ATimeoutMs);

  FPendingLock.Enter;
  try
    FPending.ExtractPair(Id); // Remove without Free (we now own Pending)
  finally
    FPendingLock.Leave;
  end;

  if WaitResult <> wrSignaled then
  begin
    Pending.Free;
    raise ELspTimeout.Create('Timeout bei LSP-Request: ' + AMethod);
  end;

  Result := Pending.Response;
  Pending.Response := nil; // Ownership to caller
  Pending.Free;

  // Check for error
  if Result.TryGetValue<TJSONObject>('error', ErrorObj) then
  begin
    var ErrCode := ErrorObj.GetValue<Integer>('code', -1);
    var ErrMsg := ErrorObj.GetValue<string>('message', 'Unbekannter LSP-Fehler');
    var Exc := ELspError.Create(ErrCode, ErrMsg);
    Result.Free;
    raise Exc;
  end;
end;

function TLspClient.SendRequestAsync(const AMethod: string; AParams: TJSONValue): Integer;
var
  Msg: TJSONObject;
  Pending: TPendingRequest;
begin
  Result := NextRequestId;

  Msg := TJSONObject.Create;
  Msg.AddPair('jsonrpc', '2.0');
  Msg.AddPair('id', TJSONNumber.Create(Result));
  Msg.AddPair('method', AMethod);
  if AParams <> nil then
    Msg.AddPair('params', AParams)
  else
    Msg.AddPair('params', TJSONObject.Create);

  Pending := TPendingRequest.Create;

  FPendingLock.Enter;
  try
    FPending.Add(Result, Pending);
  finally
    FPendingLock.Leave;
  end;

  Log('-->', AMethod + ' (async#' + IntToStr(Result) + ')', Msg.ToJSON);
  FTransport.SendMessage(Msg);
  Msg.Free;
end;

function TLspClient.WaitForResponse(ARequestId: Integer; ATimeoutMs: Cardinal): TJSONObject;
var
  Pending: TPendingRequest;
  WaitResult: TWaitResult;
  ErrorObj: TJSONObject;
begin
  FPendingLock.Enter;
  try
    if not FPending.TryGetValue(ARequestId, Pending) then
      raise ELspError.Create(-1, 'Kein ausstehender Request mit ID ' + IntToStr(ARequestId));
  finally
    FPendingLock.Leave;
  end;

  WaitResult := Pending.Event.WaitFor(ATimeoutMs);

  FPendingLock.Enter;
  try
    FPending.ExtractPair(ARequestId);
  finally
    FPendingLock.Leave;
  end;

  if WaitResult <> wrSignaled then
  begin
    Pending.Free;
    raise ELspTimeout.Create('Timeout bei async Request #' + IntToStr(ARequestId));
  end;

  Result := Pending.Response;
  Pending.Response := nil;
  Pending.Free;

  if Result.TryGetValue<TJSONObject>('error', ErrorObj) then
  begin
    var ErrCode := ErrorObj.GetValue<Integer>('code', -1);
    var ErrMsg := ErrorObj.GetValue<string>('message', 'Unbekannter LSP-Fehler');
    var Exc := ELspError.Create(ErrCode, ErrMsg);
    Result.Free;
    raise Exc;
  end;
end;

procedure TLspClient.SendNotification(const AMethod: string; AParams: TJSONValue);
var
  Msg: TJSONObject;
begin
  Msg := TJSONObject.Create;
  Msg.AddPair('jsonrpc', '2.0');
  Msg.AddPair('method', AMethod);
  if AParams <> nil then
    Msg.AddPair('params', AParams)
  else
    Msg.AddPair('params', TJSONObject.Create);

  Log('-->', AMethod, Msg.ToJSON);
  FTransport.SendMessage(Msg);
  Msg.Free;
end;

procedure TLspClient.Initialize(const ARootPath, ADprojPath: string; const ASearchPath: string);
var
  Params: TJSONObject;
  Response: TJSONObject;
  ResultObj: TJSONObject;
begin
  Params := TJSONObject.Create;
  Params.AddPair('processId', TJSONNumber.Create(GetCurrentProcessId));
  Params.AddPair('rootUri', TLspUri.PathToFileUri(ARootPath));
  Params.AddPair('rootPath', ARootPath);
  Params.AddPair('capabilities', TLspProtocol.BuildClientCapabilities);
  Params.AddPair('initializationOptions',
    TLspProtocol.BuildInitializationOptions(ADprojPath, ASearchPath));

  Response := SendRequest('initialize', Params);
  try
    // Store server capabilities
    if Response.TryGetValue<TJSONObject>('result', ResultObj) then
    begin
      var Caps := ResultObj.GetValue<TJSONObject>('capabilities');
      if Caps <> nil then
        FServerCapabilities := Caps.Clone as TJSONObject;
    end;
  finally
    Response.Free;
  end;

  // Send initialized notification
  SendNotification('initialized');
end;

procedure TLspClient.SendConfiguration(const ADelphiLspJsonPath: string);
var
  Content: string;
  JsonValue: TJSONValue;
  ConfigObj: TJSONObject;
begin
  if not FileExists(ADelphiLspJsonPath) then
    raise EFileNotFoundException.Create('.delphilsp.json nicht gefunden: ' + ADelphiLspJsonPath);

  Content := TFile.ReadAllText(ADelphiLspJsonPath, TEncoding.UTF8);
  JsonValue := TJSONObject.ParseJSONValue(Content);
  if JsonValue = nil then
    raise EJSONException.Create('Ungueltige JSON in: ' + ADelphiLspJsonPath);

  try
    // workspace/didChangeConfiguration expects { settings: { ... } }
    // The .delphilsp.json already has the format { settings: { ... } }
    if JsonValue is TJSONObject then
      ConfigObj := TJSONObject(JsonValue)
    else
    begin
      JsonValue.Free;
      raise EJSONException.Create('JSON-Objekt erwartet in: ' + ADelphiLspJsonPath);
    end;

    SendNotification('workspace/didChangeConfiguration', ConfigObj);
    // ConfigObj is not freed by SendNotification because it is embedded
    // as a parameter into the message. Do not free separately.
  except
    on E: EJSONException do
      raise;
    on E: Exception do
    begin
      JsonValue.Free;
      raise;
    end;
  end;
end;

procedure TLspClient.OpenDocument(const AFilePath: string);
var
  Params, TextDocObj: TJSONObject;
  Content: string;
  AbsPath: string;
begin
  AbsPath := ExpandFileName(AFilePath);
  Content := Delphi.FileEncoding.ReadDelphiFile(AbsPath);

  TextDocObj := TJSONObject.Create;
  TextDocObj.AddPair('uri', TLspUri.PathToFileUri(AbsPath));
  TextDocObj.AddPair('languageId', 'pascal');
  TextDocObj.AddPair('version', TJSONNumber.Create(1));
  TextDocObj.AddPair('text', Content);

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDocObj);

  SendNotification('textDocument/didOpen', Params);
end;

procedure TLspClient.CloseDocument(const AFilePath: string);
var
  Params, TextDocObj: TJSONObject;
  AbsPath: string;
begin
  AbsPath := ExpandFileName(AFilePath);
  TextDocObj := TJSONObject.Create;
  TextDocObj.AddPair('uri', TLspUri.PathToFileUri(AbsPath));
  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDocObj);
  SendNotification('textDocument/didClose', Params);
end;

procedure TLspClient.RefreshDocument(const AFilePath: string);
begin
  CloseDocument(AFilePath);
  Sleep(50);
  OpenDocument(AFilePath);
end;

function TLspClient.SupportsPrepareRename: Boolean;
var
  RenameProvider: TJSONValue;
begin
  Result := False;
  if FServerCapabilities = nil then
    Exit;
  RenameProvider := FServerCapabilities.GetValue('renameProvider');
  if RenameProvider = nil then
    Exit;
  if RenameProvider is TJSONObject then
    Result := TJSONObject(RenameProvider).GetValue<Boolean>('prepareProvider', False)
  else
    Result := True; // If renameProvider = true, there is no prepareRename
end;

function TLspClient.PrepareRename(const AFilePath: string; ALine, ACol: Integer): TLspPrepareRenameResult;
var
  Params: TJSONObject;
  TextDoc, Pos: TJSONObject;
  Response: TJSONObject;
begin
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);

  Response := SendRequest('textDocument/prepareRename', Params);
  try
    Result := TLspPrepareRenameResult.FromJSON(Response.GetValue('result'));
  finally
    Response.Free;
  end;
end;

function TLspClient.Rename(const AFilePath: string; ALine, ACol: Integer; const ANewName: string): TLspWorkspaceEdit;
var
  Params: TJSONObject;
  TextDoc, Pos: TJSONObject;
  Response: TJSONObject;
  ResultObj: TJSONObject;
begin
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);
  Params.AddPair('newName', ANewName);

  Response := SendRequest('textDocument/rename', Params);
  try
    if Response.TryGetValue<TJSONObject>('result', ResultObj) then
      Result := TLspWorkspaceEdit.FromJSON(ResultObj)
    else
    begin
      SetLength(Result.FileEdits, 0);
    end;
  finally
    Response.Free;
  end;
end;

function TLspClient.FindReferences(const AFilePath: string; ALine, ACol: Integer;
  AIncludeDeclaration: Boolean): TArray<TLspLocation>;
var
  Params: TJSONObject;
  TextDoc, Pos, Context: TJSONObject;
  Response: TJSONObject;
  ResultArr: TJSONArray;
  I: Integer;
begin
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Context := TJSONObject.Create;
  Context.AddPair('includeDeclaration', TJSONBool.Create(AIncludeDeclaration));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);
  Params.AddPair('context', Context);

  Response := SendRequest('textDocument/references', Params);
  try
    if Response.TryGetValue<TJSONArray>('result', ResultArr) then
    begin
      SetLength(Result, ResultArr.Count);
      for I := 0 to ResultArr.Count - 1 do
        Result[I] := TLspLocation.FromJSON(ResultArr.Items[I] as TJSONObject);
    end
    else
      SetLength(Result, 0);
  finally
    Response.Free;
  end;
end;

function TLspClient.GotoDefinition(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>;
var
  Params, TextDoc, Pos: TJSONObject;
  Response: TJSONObject;
  ResultValue: TJSONValue;
begin
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);

  Response := SendRequest('textDocument/definition', Params);
  try
    ResultValue := Response.GetValue('result');
    if ResultValue = nil then
    begin
      SetLength(Result, 0);
      Exit;
    end;

    // May be a single Location object or an array
    if ResultValue is TJSONArray then
    begin
      var Arr := TJSONArray(ResultValue);
      SetLength(Result, Arr.Count);
      for var I := 0 to Arr.Count - 1 do
        Result[I] := TLspLocation.FromJSON(Arr.Items[I] as TJSONObject);
    end
    else if ResultValue is TJSONObject then
    begin
      SetLength(Result, 1);
      Result[0] := TLspLocation.FromJSON(TJSONObject(ResultValue));
    end
    else
      SetLength(Result, 0);
  finally
    Response.Free;
  end;
end;

function TLspClient.GotoImplementation(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>;
var
  Params, TextDoc, Pos: TJSONObject;
  Response: TJSONObject;
  ResultValue: TJSONValue;
begin
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);

  Response := SendRequest('textDocument/implementation', Params);
  try
    ResultValue := Response.GetValue('result');
    if ResultValue = nil then
    begin
      SetLength(Result, 0);
      Exit;
    end;

    if ResultValue is TJSONArray then
    begin
      var Arr := TJSONArray(ResultValue);
      SetLength(Result, Arr.Count);
      for var I := 0 to Arr.Count - 1 do
        Result[I] := TLspLocation.FromJSON(Arr.Items[I] as TJSONObject);
    end
    else if ResultValue is TJSONObject then
    begin
      SetLength(Result, 1);
      Result[0] := TLspLocation.FromJSON(TJSONObject(ResultValue));
    end
    else
      SetLength(Result, 0);
  finally
    Response.Free;
  end;
end;

function TLspClient.GetHover(const AFilePath: string; ALine, ACol: Integer): string;
var
  Params, TextDoc, Pos: TJSONObject;
  Response: TJSONObject;
  ResultObj: TJSONObject;
  Contents: TJSONValue;
begin
  Result := '';
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);

  Response := SendRequest('textDocument/hover', Params);
  try
    if Response.TryGetValue<TJSONObject>('result', ResultObj) then
    begin
      Contents := ResultObj.GetValue('contents');
      if Contents is TJSONObject then
        Result := TJSONObject(Contents).GetValue<string>('value', '')
      else if Contents is TJSONString then
        Result := Contents.Value
      else if Contents is TJSONArray then
      begin
        // Multiple contents
        for var Item in TJSONArray(Contents) do
        begin
          if Result <> '' then
            Result := Result + #13#10;
          if Item is TJSONObject then
            Result := Result + TJSONObject(Item).GetValue<string>('value', '')
          else
            Result := Result + Item.Value;
        end;
      end;
    end;
  finally
    Response.Free;
  end;
end;

function TLspClient.GetCompletion(const AFilePath: string; ALine, ACol: Integer): TJSONObject;
var
  Params, TextDoc, Pos: TJSONObject;
  Response: TJSONObject;
begin
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);

  Response := SendRequest('textDocument/completion', Params);
  // Ownership passes to the caller
  Result := Response;
end;

function TLspClient.GetServerCapabilities: string;
begin
  if FServerCapabilities <> nil then
    Result := FServerCapabilities.ToJSON
  else
    Result := '(not available)';
end;

function TLspClient.SupportsReferences: Boolean;
begin
  Result := False;
  if FServerCapabilities <> nil then
    Result := FServerCapabilities.GetValue<Boolean>('referencesProvider', False);
end;

procedure TLspClient.Shutdown;
var
  Response: TJSONObject;
begin
  try
    Response := SendRequest('shutdown', nil, 15000);
    Response.Free;
  except
    // Tolerate shutdown errors
  end;

  try
    SendNotification('exit');
  except
  end;

  // Wait for process exit
  if FProcessHandle <> INVALID_HANDLE_VALUE then
  begin
    if WaitForSingleObject(FProcessHandle, 5000) = WAIT_TIMEOUT then
      TerminateProcess(FProcessHandle, 1);
  end;
end;

end.
