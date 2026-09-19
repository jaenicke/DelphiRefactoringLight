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
  Delphi.FileEncoding, Expert.EditorHelperIntf, Lsp.JsonRpc, Lsp.Protocol, Lsp.Uri;

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
    FExtraArgs: string;
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
    /// <summary>Map: uppercase absolute file path -> array of inactive
    ///  ranges. Populated by HandlePublishDiagnostics from DelphiLSP
    ///  diagnostics with code 'H2655'/'H2656' and tag=1.</summary>
    FInactiveRanges: TObjectDictionary<string, TList<TLspRange>>;
    /// <summary>Map: uppercase absolute file path -> array of error/warning
    ///  diagnostics (severity + code + range), for the auto-import feature
    ///  (undeclared identifiers = code 'E2003'). Guarded by FInactiveRangesLock.</summary>
    FErrorDiags: TObjectDictionary<string, TList<TLspErrorDiag>>;
    FInactiveRangesLock: TCriticalSection;
    // Partial-unit guard (see BeforePositionRequest)
    FPosLock: TCriticalSection;
    FLastPosFile: string;                        // file of the last position request
    FDocLines: TDictionary<string, Integer>;     // UPPER path -> lines last sent
    FDocHash: TDictionary<string, Cardinal>;     // UPPER path -> hash of the text last sent
    // UPPER file NAME -> full path of every document sent to the server
    // ('' when two different paths share the name) - see ResolveBareUris
    FKnownFiles: TDictionary<string, string>;
    FAutoCompleteUnits: Boolean;
    FDiagnosticsCount: Integer;
    // Pushes per file (see GetFileDiagnosticsVersion).
    FFileDiagVersion: TDictionary<string, Integer>;
    /// <summary>Set of uppercase file paths that have received at least
    ///  one publishDiagnostics notification.</summary>
    FFilesWithDiagnostics: TDictionary<string, Boolean>;
    /// <summary>The reader thread has ended - nothing will ever answer.</summary>
    FReaderDead: Boolean;

    procedure MarkReaderDead;
    procedure CheckConnected;
    function ResolveBareUris(const ALocs: TArray<TLspLocation>;
      const ARequestFile: string): TArray<TLspLocation>;
    function NextRequestId: Integer;
    procedure BeforePositionRequest(const AMethod: string; AParams: TJSONValue);
    function LineCountOf(const AFilePath: string): Integer;
    procedure DispatchResponse(AMsg: TJSONObject);
    procedure HandleServerRequest(AMsg: TJSONObject);
    procedure HandlePublishDiagnostics(AParams: TJSONObject);
    procedure Log(const ADirection, AMethod, ABody: string);
  public
    constructor Create(const ALspExePath: string);
    /// <summary>Extra command line for DelphiLsp.exe, e.g. '-LogModes 255'
    ///  (writes %TEMP%\DelphiLSP\DelphiLSP.log). Set before Start.</summary>
    property ExtraArgs: string read FExtraArgs write FExtraArgs;
    /// <summary>MEASURED (DelphiLSP 13): a position request (definition,
    ///  hover, ...) compiles the queried unit only UP TO THE CURSOR, and
    ///  units USING it then see that partial state - after a query at a
    ///  declaration in unit A, every query in a unit B that uses A answered
    ///  null until A was queried again further down. With this on (default)
    ///  the client queries the previous file once at its LAST line before
    ///  the first position request in ANOTHER file, which restores it
    ///  (a few ms per file switch). Off only for the A/B probe.</summary>
    property AutoCompleteUnits: Boolean read FAutoCompleteUnits write FAutoCompleteUnits;
    destructor Destroy; override;

    /// <summary>Starts the LSP server as a subprocess.</summary>
    procedure Start;

    /// <summary>The server process runs AND the reader thread is alive.
    ///  Cheap (no round trip). False means every request would fail.</summary>
    function IsConnected: Boolean;

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

    /// <summary>Refreshes a document at the LSP (didClose + didOpen).
    ///  Virtual (like GotoDefinition) so tests can script the server's
    ///  answers without starting one.</summary>
    procedure RefreshDocument(const AFilePath: string); virtual;
    /// <summary>RefreshDocument with the content HANDED IN. The plain
    ///  version reads the live editor buffer, which is ToolsAPI and
    ///  therefore MAIN-THREAD ONLY - a background worker must capture the
    ///  content before it starts and pass it here.</summary>
    procedure RefreshDocumentWith(const AFilePath, AContent: string);
    procedure OpenDocumentWith(const AFilePath, AContent: string);

    /// <summary>Hands AContent to the session ONLY when it differs from the
    ///  text last sent for this file (every send goes through
    ///  OpenDocumentWith, which records it). True = it was sent, i.e.
    ///  DelphiLSP analyses the unit again and answers NOTHING in it until
    ///  that is done - wait with WaitFileAnalysed. MEASURED: re-sending an
    ///  UNCHANGED unit (the old RefreshDocument habit) costs a full
    ///  re-analysis, 12 s for a big unit, and queries in between answer
    ///  null.</summary>
    function SyncDocumentWith(const AFilePath, AContent: string): Boolean;
    /// <summary>SyncDocumentWith with the live content (editor buffer on
    ///  the main thread, else disk).</summary>
    function SyncDocument(const AFilePath: string): Boolean;
    /// <summary>Waits until DelphiLSP pushed diagnostics for AFilePath again
    ///  (its per-file version moved past ABefore - take it BEFORE sending).
    ///  AKeepWaiting is called between the polls (pump messages there;
    ///  return False to give up). False on timeout / give-up.</summary>
    function WaitFileAnalysed(const AFilePath: string; ABefore: Integer;
      ATimeoutMs: Cardinal; const AKeepWaiting: TFunc<Boolean> = nil): Boolean;

    /// <summary>Checks whether a rename is possible at the position.</summary>
    function PrepareRename(const AFilePath: string; ALine, ACol: Integer): TLspPrepareRenameResult;

    /// <summary>Performs a rename and returns the WorkspaceEdit.</summary>
    function Rename(const AFilePath: string; ALine, ACol: Integer;
      const ANewName: string): TLspWorkspaceEdit;

    /// <summary>Finds all references of an identifier (incl. declaration).</summary>
    function FindReferences(const AFilePath: string; ALine, ACol: Integer;
      AIncludeDeclaration: Boolean = True): TArray<TLspLocation>;

    /// <summary>Jumps to the definition of the identifier at the position.</summary>
    function GotoDefinition(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>; virtual;

    /// <summary>Finds implementations (e.g. class methods that implement an interface).</summary>
    function GotoImplementation(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>;

    /// <summary>Gets hover info (type, description) for the identifier.</summary>
    function GetHover(const AFilePath: string; ALine, ACol: Integer): string;
    /// <summary>Calls textDocument/signatureHelp. Returns the response's
    ///  "result" object verbatim (caller owns and must Free). Returns
    ///  nil when the server returns null / no result / errors. The
    ///  result, when non-nil, has shape:
    ///    { signatures: [ { label, documentation, parameters[...] } ],
    ///      activeSignature: N, activeParameter: N }
    ///  Use SignatureHelpLabels for a convenience extractor.</summary>
    function GetSignatureHelp(const AFilePath: string; ALine, ACol: Integer): TJSONObject;

    /// <summary>Requests code completion. Returns a JSON array of items.</summary>
    function GetCompletion(const AFilePath: string; ALine, ACol: Integer): TJSONObject;

    /// <summary>Returns the document's symbol tree (textDocument/documentSymbol).
    ///  Caller owns the returned array and must free it.</summary>
    function GetDocumentSymbols(const AFilePath: string;
      ATimeoutMs: Cardinal = 60000): TJSONArray;

    /// <summary>Returns the server capabilities as a JSON string (debugging).</summary>
    function GetServerCapabilities: string;

    /// <summary>Returns the inactive {$IFDEF}-region ranges that DelphiLSP
    ///  has reported for AFilePath, as 0-based LSP positions. Empty when
    ///  the server has not yet pushed diagnostics for this file.</summary>
    function GetInactiveRanges(const AFilePath: string): TArray<TLspRange>;

    /// <summary>Returns the error/warning diagnostics DelphiLSP has reported
    ///  for AFilePath (0-based ranges). Empty when none / not yet analysed.
    ///  Use the 'E2003' code to find undeclared identifiers.</summary>
    function GetErrorDiagnostics(const AFilePath: string): TArray<TLspErrorDiag>;

    /// <summary>True iff the 0-based line ALine lies inside one of the
    ///  inactive ranges DelphiLSP reported for AFilePath.</summary>
    function IsLineInactive(const AFilePath: string; ALine: Integer): Boolean;

    /// <summary>Number of textDocument/publishDiagnostics notifications
    ///  received so far (incl. empty ones). Lets the caller detect that
    ///  the server actually pushes diagnostics. 0 == controller-mode
    ///  may not be active or the server hasn't analysed yet.</summary>
    function GetDiagnosticsCount: Integer;
    /// <summary>How many FILES we currently hold error diagnostics for -
    ///  the push counter alone cannot tell "123 pushes, but none for the
    ///  file you are looking at" (status window).</summary>
    function GetDiagnosticFileCount: Integer;
    /// <summary>How many diagnostic pushes arrived FOR THIS FILE. The
    ///  session-wide counter cannot answer "did the file I am waiting for
    ///  get its answer yet" - waiting on it made us read a stale (usually
    ///  empty) result whenever the push that arrived belonged to another
    ///  file.</summary>
    function GetFileDiagnosticsVersion(const AFilePath: string): Integer;

    /// <summary>Total inactive ranges across all known files.</summary>
    function GetInactiveRangesTotal: Integer;

    /// <summary>ESTIMATED heap bytes this client retains: diagnostics of
    ///  every file DelphiLSP ever pushed (all severities, with messages),
    ///  inactive ranges and the per-file bookkeeping. For the status
    ///  window - takes the diagnostics lock and walks everything.</summary>
    function EstimateRetainedBytes: Int64;

    /// <summary>True iff DelphiLSP has pushed at least one
    ///  publishDiagnostics notification for AFilePath (even if empty).
    ///  Used to decide whether IsLineInactive is meaningful or still
    ///  awaiting server analysis.</summary>
    function HasReceivedDiagnostics(const AFilePath: string): Boolean;

    /// <summary>Blocks until HasReceivedDiagnostics(AFilePath) becomes
    ///  True or ATimeoutMs elapses. Returns True on success.</summary>
    function WaitForDiagnostics(const AFilePath: string;
      ATimeoutMs: Cardinal): Boolean;

    /// <summary>Brings AFilePath into a state where hover / definition /
    ///  documentSymbol on it are reliable. Does the full warm-up sequence
    ///  uniformly for every wizard:
    ///   1. RefreshDocument (didOpen + didChange v2 with full content)
    ///   2. textDocument/documentSymbol with ASymbolTimeoutMs (blocks
    ///      until DelphiLSP has actually parsed the file or the timeout
    ///      hits - the request itself is the analysis trigger)
    ///   3. WaitForDiagnostics with ADiagnosticsTimeoutMs (publishDiag-
    ///      nostics typically arrive seconds after symbol resolution)
    ///
    ///  Idempotent: when HasReceivedDiagnostics(AFilePath) is already
    ///  true the routine returns immediately. AStatusCallback (optional)
    ///  receives a short user-facing string at each step so all dialogs
    ///  show the same wording. Exceptions inside the warm-up are
    ///  swallowed - the caller proceeds on best effort.</summary>
    procedure EnsureFileAnalysed(const AFilePath: string;
      ASymbolTimeoutMs, ADiagnosticsTimeoutMs: Cardinal;
      AStatusCallback: TProc<string> = nil);

    /// <summary>Shuts the server down cleanly.</summary>
    procedure Shutdown;

    /// <summary>Indicates whether the server supports prepareRename.</summary>
    function SupportsPrepareRename: Boolean;

    /// <summary>Indicates whether the server supports references.</summary>
    function SupportsReferences: Boolean;

    property Verbose: Boolean read FVerbose write FVerbose;
    property OnLog: TLspLogEvent read FOnLog write FOnLog;
  end;

type
  /// <summary>The source positions (file + 0-based line) that ARE one symbol:
  ///  its declaration and its implementation, for interface/virtual methods
  ///  also those of the implementing classes. Candidates of a text scan are
  ///  verified against this set instead of against the declaring FILE -
  ///  with ten 'Init' methods in one unit the file check accepted all of
  ///  them (forum report 2026-09: Find References 22 hits, and a rename
  ///  would have renamed every Init of the unit).</summary>
  TLspSymbolTargets = record
  private
    FKeys: TArray<string>;
    class function Key(const AFile: string; ALine: Integer): string; static;
  public
    procedure Add(const AFile: string; ALine: Integer);
    /// <summary>Adds the position AND where DelphiLSP's GotoDefinition takes
    ///  it (declaration <-> implementation).</summary>
    procedure AddWithPartner(AClient: TLspClient; const AFile: string;
      ALine, ACol: Integer);
    function Contains(const AFile: string; ALine: Integer): Boolean;
    /// <summary>Any position of the symbol in that FILE - the coarse test
    ///  for cases where a line cannot be pinned down (overloads).</summary>
    function ContainsFile(const AFile: string): Boolean;
    function Count: Integer;
    function Text: string;
  end;

implementation

// FNV-1a over the UTF-16 code units - identifies the text last sent per
// document (SyncDocumentWith); the Lsp units stay free of Expert.* units
function ContentHash(const S: string): Cardinal;
begin
  Result := 2166136261;
  for var I := 1 to Length(S) do
  begin
    Result := Result xor Ord(S[I]);
    // in 64 bit (< 2^57), so no {$Q} juggling that could leak into the unit
    Result := Cardinal((UInt64(Result) * 16777619) and $FFFFFFFF);
  end;
end;

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
const
  // A desynchronised stream produces nothing but bad frames - give up
  // after this many in a row instead of spinning forever.
  MaxBadFramesInRow = 20;
var
  Msg: TJSONObject;
  BadInRow: Integer;
begin
  BadInRow := 0;
  try
    while not Terminated do
    begin
      try
        Msg := FOwner.FTransport.ReadMessage;
      except
        on E: EStreamError do
        begin
          // The pipe is gone (process ended or handle closed): final.
          if not Terminated then
            FOwner.Log('<--', 'ERROR', E.Message);
          Exit;
        end;
        on E: Exception do
        begin
          // One malformed frame (EJSONException & co) used to kill this
          // thread silently - and every later request then sat out its
          // full timeout, for the rest of the IDE session.
          if not Terminated then
            FOwner.Log('<--', 'BAD FRAME', E.ClassName + ': ' + E.Message);
          Inc(BadInRow);
          if BadInRow >= MaxBadFramesInRow then Exit;
          Continue;
        end;
      end;
      BadInRow := 0;
      try
        FOwner.DispatchResponse(Msg);
      except
        on E: Exception do
          // e.g. an id of unexpected type - drop this message, keep reading
          if not Terminated then
            FOwner.Log('<--', 'DISPATCH ERROR', E.ClassName + ': ' + E.Message);
      end;
    end;
  finally
    FOwner.MarkReaderDead;
  end;
end;

procedure TLspClient.MarkReaderDead;
begin
  FReaderDead := True;
  // Wake every waiter NOW with no response: they raise "connection lost"
  // at once instead of waiting out their timeout.
  FPendingLock.Enter;
  try
    for var P in FPending.Values do
      P.Event.SetEvent;
  finally
    FPendingLock.Leave;
  end;
end;

function TLspClient.IsConnected: Boolean;
begin
  Result := (FReaderThread <> nil) and not FReaderDead and
    (FProcessHandle <> INVALID_HANDLE_VALUE) and
    (WaitForSingleObject(FProcessHandle, 0) = WAIT_TIMEOUT);
end;

procedure TLspClient.CheckConnected;
begin
  if FReaderDead then
    raise ELspError.Create(-32099, 'Connection to DelphiLSP lost');
end;

{ TLspClient }

constructor TLspClient.Create(const ALspExePath: string);
begin
  inherited Create;
  FLspExePath := ALspExePath;
  FNextId := 0;
  FPending := TObjectDictionary<Integer, TPendingRequest>.Create([doOwnsValues]);
  FPendingLock := TCriticalSection.Create;
  FInactiveRanges := TObjectDictionary<string, TList<TLspRange>>.Create([doOwnsValues]);
  FErrorDiags := TObjectDictionary<string, TList<TLspErrorDiag>>.Create([doOwnsValues]);
  FInactiveRangesLock := TCriticalSection.Create;
  FPosLock := TCriticalSection.Create;
  FDocLines := TDictionary<string, Integer>.Create;
  FDocHash := TDictionary<string, Cardinal>.Create;
  FKnownFiles := TDictionary<string, string>.Create;
  FAutoCompleteUnits := True;
  FFilesWithDiagnostics := TDictionary<string, Boolean>.Create;
  FFileDiagVersion := TDictionary<string, Integer>.Create;
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
  FreeAndNil(FDocLines);
  FreeAndNil(FDocHash);
  FreeAndNil(FKnownFiles);
  FreeAndNil(FPosLock);
  FInactiveRanges.Free;
  FErrorDiags.Free;
  FInactiveRangesLock.Free;
  FFilesWithDiagnostics.Free;
  FFileDiagVersion.Free;
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
    if SameText(Method, 'textDocument/publishDiagnostics') then
    begin
      // Count every publishDiagnostics notification we get, regardless
      // of whether the inner extraction succeeds. The count is the
      // "is the server talking to us" signal that the status line uses;
      // it must reflect actual traffic, not the strictness of our
      // inactive-range parser.
      TInterlocked.Increment(FDiagnosticsCount);
      var ParamsObj: TJSONObject;
      if AMsg.TryGetValue<TJSONObject>('params', ParamsObj) then
      try
        HandlePublishDiagnostics(ParamsObj);
      except
        // Diagnostic-Parsing-Fehler darf den Reader-Thread nicht killen
      end;
    end;
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
  if FExtraArgs <> '' then
    CmdLine := CmdLine + ' ' + FExtraArgs;

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

function TLspClient.LineCountOf(const AFilePath: string): Integer;
begin
  FPosLock.Enter;
  try
    if FDocLines.TryGetValue(UpperCase(ExpandFileName(AFilePath)), Result) then Exit;
  finally
    FPosLock.Leave;
  end;
  // never sent by us - the server has it from disk
  Result := 0;
  try
    var S := TFile.ReadAllText(AFilePath);
    Result := 1;
    for var Ch in S do
      if Ch = #10 then Inc(Result);
  except
  end;
end;

threadvar
  GInUnitReset: Boolean;   // the reset probe itself must not trigger one

procedure TLspClient.BeforePositionRequest(const AMethod: string; AParams: TJSONValue);
var
  Uri, FilePath, Prev: string;
begin
  if not FAutoCompleteUnits or GInUnitReset then Exit;
  if not ((AMethod = 'textDocument/definition') or (AMethod = 'textDocument/hover') or
          (AMethod = 'textDocument/implementation') or (AMethod = 'textDocument/references') or
          (AMethod = 'textDocument/signatureHelp') or (AMethod = 'textDocument/completion') or
          (AMethod = 'textDocument/prepareRename') or (AMethod = 'textDocument/rename') or
          (AMethod = 'textDocument/typeDefinition') or (AMethod = 'textDocument/declaration')) then
    Exit;
  if not (AParams is TJSONObject) then Exit;
  Uri := TJSONObject(AParams).GetValue<string>('textDocument.uri', '');
  if Uri = '' then Exit;
  FilePath := TLspUri.FileUriToPath(Uri);
  FPosLock.Enter;
  try
    Prev := FLastPosFile;
    FLastPosFile := FilePath;
  finally
    FPosLock.Leave;
  end;
  if (Prev = '') or SameText(ExpandFileName(Prev), ExpandFileName(FilePath)) then Exit;
  // A query at the LAST line compiles the previous unit completely again.
  var Lines := LineCountOf(Prev);
  if Lines <= 0 then Exit;
  GInUnitReset := True;
  try
    try
      GotoDefinition(Prev, Lines - 1, 0);
    except
      // only a reset probe - its answer and failure do not matter
    end;
  finally
    GInUnitReset := False;
  end;
end;

function TLspClient.SendRequest(const AMethod: string; AParams: TJSONValue; ATimeoutMs: Cardinal): TJSONObject;
var
  Id: Integer;
  Msg: TJSONObject;
  Pending: TPendingRequest;
  WaitResult: TWaitResult;
  ErrorObj: TJSONObject;
begin
  if FReaderDead then
  begin
    AParams.Free;   // the caller handed us ownership
    CheckConnected;
  end;
  BeforePositionRequest(AMethod, AParams);
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
    // the reader may have died meanwhile - never wait for nothing
    if FReaderDead then Pending.Event.SetEvent;
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
  if Result = nil then   // woken by MarkReaderDead
    raise ELspError.Create(-32099, 'Connection to DelphiLSP lost during ' + AMethod);

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
  if FReaderDead then
  begin
    AParams.Free;
    CheckConnected;
  end;
  BeforePositionRequest(AMethod, AParams);
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
    // the reader may have died meanwhile - never wait for nothing
    if FReaderDead then Pending.Event.SetEvent;
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
  if Result = nil then   // woken by MarkReaderDead
    raise ELspError.Create(-32099, 'Connection to DelphiLSP lost (request #' +
      IntToStr(ARequestId) + ')');

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

function ReadLiveContent(const AFilePath: string): string;
// Prefer the live editor buffer (via IEditorHelper) over disk so the
// LSP sees the user's unsaved edits. The IDE plugin's TIDEEditorHelper
// reads from IOTAEditBuffer; the standalone's TStandaloneEditorHelper
// reads from its in-memory Memo buffer. Without this hop the LSP gets
// stale content - the dot the user just typed isn't on disk yet, so
// completion at the new column returns scope-level suggestions
// instead of the dotted-member list.
var
  EditorContent: string;
begin
  // ToolsAPI IS MAIN-THREAD ONLY. Reading the edit buffer from a worker
  // thread (IOTASourceEditor.CreateReader/GetText) races the IDE's own
  // editing and parsing and is a documented way to hang the IDE - the
  // more LSP traffic runs in parallel, the more often. Off the main
  // thread we therefore read from DISK; a caller that needs the live
  // buffer captures it on the main thread and uses the ...With variants.
  if (Editor <> nil) and (TThread.CurrentThread.ThreadID = MainThreadID)
    and Editor.ReadEditorContent(AFilePath, EditorContent) then
    Result := EditorContent
  else
    Result := Delphi.FileEncoding.ReadDelphiFile(ExpandFileName(AFilePath));
end;

procedure TLspClient.OpenDocument(const AFilePath: string);
begin
  OpenDocumentWith(AFilePath, ReadLiveContent(AFilePath));
end;

procedure TLspClient.OpenDocumentWith(const AFilePath, AContent: string);
var
  Params, TextDocObj: TJSONObject;
  Content: string;
  AbsPath: string;
begin
  AbsPath := ExpandFileName(AFilePath);
  Content := AContent;
  var Lines := 1;
  for var Ch in Content do
    if Ch = #10 then Inc(Lines);
  FPosLock.Enter;
  try
    FDocLines.AddOrSetValue(UpperCase(AbsPath), Lines);
    FDocHash.AddOrSetValue(UpperCase(AbsPath), ContentHash(Content));
    var NameKey := UpperCase(ExtractFileName(AbsPath));
    var Known: string;
    if not FKnownFiles.TryGetValue(NameKey, Known) then
      FKnownFiles.Add(NameKey, AbsPath)
    else if (Known <> '') and not SameText(Known, AbsPath) then
      FKnownFiles[NameKey] := '';   // ambiguous
  finally
    FPosLock.Leave;
  end;

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
  FPosLock.Enter;
  try
    FDocHash.Remove(UpperCase(AbsPath));
  finally
    FPosLock.Leave;
  end;
  TextDocObj := TJSONObject.Create;
  TextDocObj.AddPair('uri', TLspUri.PathToFileUri(AbsPath));
  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDocObj);
  SendNotification('textDocument/didClose', Params);
end;

function TLspClient.SyncDocumentWith(const AFilePath, AContent: string): Boolean;
var
  Old: Cardinal;
begin
  var Key := UpperCase(ExpandFileName(AFilePath));
  var H := ContentHash(AContent);
  FPosLock.Enter;
  try
    Result := not (FDocHash.TryGetValue(Key, Old) and (Old = H));
  finally
    FPosLock.Leave;
  end;
  if Result then RefreshDocumentWith(AFilePath, AContent);
end;

function TLspClient.SyncDocument(const AFilePath: string): Boolean;
begin
  Result := SyncDocumentWith(AFilePath, ReadLiveContent(AFilePath));
end;

function TLspClient.WaitFileAnalysed(const AFilePath: string; ABefore: Integer;
  ATimeoutMs: Cardinal; const AKeepWaiting: TFunc<Boolean>): Boolean;
begin
  var Deadline := GetTickCount64 + ATimeoutMs;
  repeat
    if GetFileDiagnosticsVersion(AFilePath) <> ABefore then Exit(True);
    if Assigned(AKeepWaiting) and not AKeepWaiting() then Exit(False);
    Sleep(50);
  until GetTickCount64 > Deadline;
  Result := False;
end;

procedure TLspClient.RefreshDocument(const AFilePath: string);
begin
  RefreshDocumentWith(AFilePath, ReadLiveContent(AFilePath));
end;

procedure TLspClient.RefreshDocumentWith(const AFilePath, AContent: string);
var
  Params, TextDocObj: TJSONObject;
  ChangesArr: TJSONArray;
  ChangeObj: TJSONObject;
  Content: string;
  AbsPath: string;
begin
  CloseDocument(AFilePath);
  Sleep(50);
  OpenDocumentWith(AFilePath, AContent);
  // Im Anschluss noch ein didChange mit identischem Inhalt schicken.
  // DelphiLSP im controller-Modus beantwortet textDocument/hover sonst
  // konsistent mit -32603 "Internal server error" - die echte Delphi-
  // IDE schickt vor dem ersten Hover IMMER ein didChange, und nur in
  // diesem Modus wird der File-State so weit aufgesetzt, dass Hover
  // funktioniert. Identischer Text + Version 2 ist ein No-Op
  // semantisch, aber bringt den Server in den erwarteten Zustand.
  AbsPath := ExpandFileName(AFilePath);
  Content := AContent;
  TextDocObj := TJSONObject.Create;
  TextDocObj.AddPair('uri', TLspUri.PathToFileUri(AbsPath));
  TextDocObj.AddPair('version', TJSONNumber.Create(2));
  ChangeObj := TJSONObject.Create;
  ChangeObj.AddPair('text', Content);
  ChangesArr := TJSONArray.Create;
  ChangesArr.AddElement(ChangeObj);
  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDocObj);
  Params.AddPair('contentChanges', ChangesArr);
  SendNotification('textDocument/didChange', Params);
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
  Result := ResolveBareUris(Result, AFilePath);
end;

function TLspClient.ResolveBareUris(const ALocs: TArray<TLspLocation>;
  const ARequestFile: string): TArray<TLspLocation>;
var
  P, Full, Cand: string;
begin
  // MEASURED (DelphiLSP 13, right after a didOpen): a definition can come
  // back as "file:///Expert.LspManager.pas" - the file NAME only. Taken as
  // a relative path it resolved against the IDE's working directory, and
  // find references then rejected the symbol's own declaration. Resolve
  // the name against the documents this session has sent, then against
  // the requesting file's folder.
  Result := ALocs;
  for var I := 0 to High(Result) do
  begin
    P := TLspUri.FileUriToPath(Result[I].Uri);
    if (P = '') or (ExtractFilePath(P) <> '') then Continue;
    Full := '';
    FPosLock.Enter;
    try
      if not FKnownFiles.TryGetValue(UpperCase(P), Full) then Full := '';
    finally
      FPosLock.Leave;
    end;
    if Full = '' then
    begin
      Cand := ExtractFilePath(ExpandFileName(ARequestFile)) + P;
      if FileExists(Cand) then Full := Cand;
    end;
    if Full <> '' then
      Result[I].Uri := TLspUri.PathToFileUri(Full);
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
  Result := ResolveBareUris(Result, AFilePath);
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
  Result := ResolveBareUris(Result, AFilePath);
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

function TLspClient.GetSignatureHelp(const AFilePath: string;
  ALine, ACol: Integer): TJSONObject;
var
  Params, TextDoc, Pos, Context: TJSONObject;
  Response: TJSONObject;
  ResultVal: TJSONValue;
  ResultObj: TJSONObject;
begin
  Result := nil;
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));

  Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));

  // SignatureHelpContext (LSP 3.15+). triggerKind 2 = TriggerCharacter.
  Context := TJSONObject.Create;
  Context.AddPair('triggerKind', TJSONNumber.Create(2));
  Context.AddPair('triggerCharacter', '(');
  Context.AddPair('isRetrigger', TJSONBool.Create(False));

  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);
  Params.AddPair('position', Pos);
  Params.AddPair('context', Context);

  Response := SendRequest('textDocument/signatureHelp', Params);
  try
    ResultVal := Response.GetValue('result');
    if ResultVal is TJSONObject then
    begin
      ResultObj := TJSONObject(ResultVal);
      Result := ResultObj.Clone as TJSONObject;
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

function TLspClient.GetDocumentSymbols(const AFilePath: string;
  ATimeoutMs: Cardinal): TJSONArray;
var
  Params, TextDoc: TJSONObject;
  Response: TJSONObject;
  ResultVal: TJSONValue;
  Arr: TJSONArray;
begin
  Result := nil;
  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFilePath)));
  Params := TJSONObject.Create;
  Params.AddPair('textDocument', TextDoc);

  Response := SendRequest('textDocument/documentSymbol', Params, ATimeoutMs);
  try
    ResultVal := Response.GetValue('result');
    if ResultVal is TJSONArray then
    begin
      Arr := TJSONArray(ResultVal.Clone);
      Result := Arr;
    end;
  finally
    Response.Free;
  end;
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

procedure TLspClient.HandlePublishDiagnostics(AParams: TJSONObject);
// Erwartet das params-Objekt einer textDocument/publishDiagnostics-
// Notification. Extrahiert alle Diagnostics mit Source='DelphiLSP' und
// Code in {H2655,H2656} (oder allgemein Tag=1 Unnecessary) als
// "inactive ranges" und speichert sie pro Datei. Ueberschreibt bei
// erneutem Diagnostics-Push die alten Werte fuer diese Datei.
var
  Uri: string;
  Path: string;
  DiagArr: TJSONArray;
  DiagVal: TJSONValue;
  DiagObj: TJSONObject;
  RangeObj, StartObj, EndObj: TJSONObject;
  Source, Code: string;
  TagArr: TJSONArray;
  HasUnnecessaryTag: Boolean;
  Severity: Integer;
  R: TLspRange;
  List: TList<TLspRange>;
  ErrList: TList<TLspErrorDiag>;
  UpKey: string;
begin
  // The notification count is bumped by the dispatcher in
  // HandleNotification; here we do the deeper extraction (inactive
  // regions AND error/warning diagnostics for auto-import).
  if AParams = nil then Exit;
  if not AParams.TryGetValue<string>('uri', Uri) then Exit;
  if not AParams.TryGetValue<TJSONArray>('diagnostics', DiagArr) then Exit;

  Path := TLspUri.FileUriToPath(Uri);
  if Path = '' then Exit;
  // Same normalisation as every reader (ExpandFileName) - otherwise writer
  // and reader keys can differ (UNC, "..\" segments) and never match.
  UpKey := AnsiUpperCase(ExpandFileName(Path));

  FInactiveRangesLock.Enter;
  try
    FFilesWithDiagnostics.AddOrSetValue(UpKey, True);
    // Bump this file's version so a waiter can tell ITS answer apart
    // from a push for some other file.
    var Ver: Integer;
    if not FFileDiagVersion.TryGetValue(UpKey, Ver) then Ver := 0;
    FFileDiagVersion.AddOrSetValue(UpKey, Ver + 1);
    if FInactiveRanges.TryGetValue(UpKey, List) then List.Clear
    else begin List := TList<TLspRange>.Create; FInactiveRanges.Add(UpKey, List); end;
    if FErrorDiags.TryGetValue(UpKey, ErrList) then ErrList.Clear
    else begin ErrList := TList<TLspErrorDiag>.Create; FErrorDiags.Add(UpKey, ErrList); end;

    for DiagVal in DiagArr do
    begin
      if not (DiagVal is TJSONObject) then Continue;
      DiagObj := TJSONObject(DiagVal);
      Source := DiagObj.GetValue<string>('source', '');
      Code := DiagObj.GetValue<string>('code', '');
      Severity := DiagObj.GetValue<Integer>('severity', 0);

      // Range (shared by both classifications).
      if not DiagObj.TryGetValue<TJSONObject>('range', RangeObj) then Continue;
      if not RangeObj.TryGetValue<TJSONObject>('start', StartObj) then Continue;
      if not RangeObj.TryGetValue<TJSONObject>('end', EndObj) then Continue;
      R.Start.Line      := StartObj.GetValue<Integer>('line', -1);
      R.Start.Character := StartObj.GetValue<Integer>('character', 0);
      R.End_.Line       := EndObj.GetValue<Integer>('line', -1);
      R.End_.Character  := EndObj.GetValue<Integer>('character', 0);
      if (R.Start.Line < 0) or (R.End_.Line < 0) then Continue;

      // Inactive $IFDEF regions: Source='DelphiLSP' AND (H2655/H2656 OR tag=1).
      HasUnnecessaryTag := False;
      if DiagObj.TryGetValue<TJSONArray>('tags', TagArr) then
        for var T: TJSONValue in TagArr do
          if (T is TJSONNumber) and (TJSONNumber(T).AsInt = 1) then
          begin HasUnnecessaryTag := True; Break; end;
      if SameText(Source, 'DelphiLSP')
         and (HasUnnecessaryTag or SameText(Code, 'H2655') or SameText(Code, 'H2656')) then
        List.Add(R);

      // Diagnostics for the auto-import quick fixes: keep every severity -
      // Delphi hints map to LSP severity 4 (H2655 arrives as severity 4),
      // and hints matter too: H2443 ("inline function not expanded because
      // unit 'X' is not in USES") carries a directly actionable uses fix.
      // Only the inactive-region codes are excluded - they are the separate
      // channel handled above and would only be noise here.
      if (Severity >= 1) and not SameText(Code, 'H2655')
        and not SameText(Code, 'H2656') then
      begin
        var ED: TLspErrorDiag;
        ED.Range := R;
        ED.Severity := Severity;
        ED.Code := Code;
        ED.Message := DiagObj.GetValue<string>('message', '');
        ErrList.Add(ED);
      end;
    end;
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.GetFileDiagnosticsVersion(const AFilePath: string): Integer;
var
  UpKey: string;
begin
  Result := 0;
  if AFilePath = '' then Exit;
  UpKey := AnsiUpperCase(ExpandFileName(AFilePath));
  FInactiveRangesLock.Enter;
  try
    FFileDiagVersion.TryGetValue(UpKey, Result);
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.GetDiagnosticFileCount: Integer;
begin
  FInactiveRangesLock.Enter;
  try
    Result := FErrorDiags.Count;
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.GetErrorDiagnostics(const AFilePath: string): TArray<TLspErrorDiag>;
var
  UpKey: string;
  List: TList<TLspErrorDiag>;
begin
  SetLength(Result, 0);
  if AFilePath = '' then Exit;
  UpKey := AnsiUpperCase(ExpandFileName(AFilePath));
  FInactiveRangesLock.Enter;
  try
    if FErrorDiags.TryGetValue(UpKey, List) then
      Result := List.ToArray;
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.GetInactiveRanges(const AFilePath: string): TArray<TLspRange>;
var
  UpKey: string;
  List: TList<TLspRange>;
begin
  SetLength(Result, 0);
  if AFilePath = '' then Exit;
  UpKey := AnsiUpperCase(ExpandFileName(AFilePath));
  FInactiveRangesLock.Enter;
  try
    if FInactiveRanges.TryGetValue(UpKey, List) then
      Result := List.ToArray;
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.IsLineInactive(const AFilePath: string; ALine: Integer): Boolean;
var
  Ranges: TArray<TLspRange>;
begin
  Result := False;
  Ranges := GetInactiveRanges(AFilePath);
  for var R in Ranges do
    if (ALine >= R.Start.Line) and (ALine <= R.End_.Line) then
      Exit(True);
end;

function TLspClient.GetDiagnosticsCount: Integer;
begin
  Result := FDiagnosticsCount;
end;

function TLspClient.EstimateRetainedBytes: Int64;

  // Same 32-bit estimate as Expert.ResourceMonitor - repeated here, the Lsp
  // units stay free of Expert.* dependencies.
  function SB(const S: string): Int64;
  begin
    if S = '' then Exit(0);
    Result := ((12 + (Int64(Length(S)) + 1) * 2 + 4 + 7) div 8) * 8;
  end;

  function LB(ACapacity, AElem: Integer): Int64;   // TList object + buffer
  begin
    Result := 32;
    if ACapacity > 0 then
      Inc(Result, ((8 + Int64(ACapacity) * AElem + 4 + 7) div 8) * 8);
  end;

begin
  Result := 0;
  FInactiveRangesLock.Enter;
  try
    for var P in FErrorDiags do
    begin
      Inc(Result, SB(P.Key) + 24 + LB(P.Value.Capacity, SizeOf(TLspErrorDiag)));
      for var D in P.Value do
        Inc(Result, SB(D.Code) + SB(D.Message));
    end;
    for var P in FInactiveRanges do
      Inc(Result, SB(P.Key) + 24 + LB(P.Value.Capacity, SizeOf(TLspRange)));
    for var K in FFileDiagVersion.Keys do
      Inc(Result, SB(K) + 24);
    for var K in FFilesWithDiagnostics.Keys do
      Inc(Result, SB(K) + 24);
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.GetInactiveRangesTotal: Integer;
var
  L: TList<TLspRange>;
begin
  Result := 0;
  FInactiveRangesLock.Enter;
  try
    for L in FInactiveRanges.Values do
      Inc(Result, L.Count);
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.HasReceivedDiagnostics(const AFilePath: string): Boolean;
var
  UpKey: string;
begin
  Result := False;
  if AFilePath = '' then Exit;
  UpKey := AnsiUpperCase(ExpandFileName(AFilePath));
  FInactiveRangesLock.Enter;
  try
    Result := FFilesWithDiagnostics.ContainsKey(UpKey);
  finally
    FInactiveRangesLock.Leave;
  end;
end;

function TLspClient.WaitForDiagnostics(const AFilePath: string;
  ATimeoutMs: Cardinal): Boolean;
var
  Deadline: TDateTime;
begin
  Result := HasReceivedDiagnostics(AFilePath);
  if Result then Exit;
  Deadline := Now + ATimeoutMs / 86400000;
  while not Result do
  begin
    Sleep(100);
    Result := HasReceivedDiagnostics(AFilePath);
    if Now > Deadline then Break;
  end;
end;

procedure TLspClient.EnsureFileAnalysed(const AFilePath: string;
  ASymbolTimeoutMs, ADiagnosticsTimeoutMs: Cardinal;
  AStatusCallback: TProc<string>);

  procedure Status(const S: string);
  begin
    if Assigned(AStatusCallback) then AStatusCallback(S);
  end;

var
  FN: string;
  Sym: TJSONArray;
begin
  FN := ExtractFileName(AFilePath);

  // Idempotenz-Shortcut: wenn der Server fuer diese Datei schon einmal
  // publishDiagnostics gepusht hat, ist sie analysiert - kein Bedarf,
  // documentSymbol erneut zu erzwingen.
  if HasReceivedDiagnostics(AFilePath) then
  begin
    Status('LSP: ' + FN + ' already analysed.');
    Exit;
  end;

  Status('LSP: refreshing ' + FN + ' (didOpen + didChange v2)...');
  try
    RefreshDocument(AFilePath);
  except
    // Refresh-Fehler ist nicht fatal - wir versuchen den Rest trotzdem.
  end;

  Status(Format(
    'LSP: requesting symbol analysis for %s (up to %d s on cold files)...',
    [FN, ASymbolTimeoutMs div 1000]));
  Sym := nil;
  try
    try
      Sym := GetDocumentSymbols(AFilePath, ASymbolTimeoutMs);
    except
      // Timeout / Server-Fehler: wir machen mit Diagnostics-Wait weiter.
    end;
  finally
    if Sym <> nil then Sym.Free;
  end;

  Status(Format(
    'LSP: waiting for diagnostics on %s (up to %d s)...',
    [FN, ADiagnosticsTimeoutMs div 1000]));
  WaitForDiagnostics(AFilePath, ADiagnosticsTimeoutMs);

  if HasReceivedDiagnostics(AFilePath) then
    Status('LSP: ' + FN + ' ready (diagnostics received).')
  else
    Status('LSP: ' + FN + ' analysed but no diagnostics arrived '
      + '(inactive-region detection unavailable).');
end;


{ TLspSymbolTargets }

class function TLspSymbolTargets.Key(const AFile: string; ALine: Integer): string;
begin
  Result := UpperCase(ExpandFileName(AFile)) + '|' + IntToStr(ALine);
end;

procedure TLspSymbolTargets.Add(const AFile: string; ALine: Integer);
begin
  if (AFile = '') or (ALine < 0) or Contains(AFile, ALine) then Exit;
  FKeys := FKeys + [Key(AFile, ALine)];
end;

procedure TLspSymbolTargets.AddWithPartner(AClient: TLspClient;
  const AFile: string; ALine, ACol: Integer);
begin
  Add(AFile, ALine);
  if AClient = nil then Exit;
  try
    var D := AClient.GotoDefinition(AFile, ALine, ACol);
    if Length(D) > 0 then
      Add(TLspUri.FileUriToPath(D[0].Uri), D[0].Range.Start.Line);
  except
    // no partner - the position itself is still in the set
  end;
end;

function TLspSymbolTargets.Contains(const AFile: string; ALine: Integer): Boolean;
begin
  var K := Key(AFile, ALine);
  for var S in FKeys do
    if S = K then Exit(True);
  Result := False;
end;

function TLspSymbolTargets.ContainsFile(const AFile: string): Boolean;
begin
  Result := False;
  if AFile = '' then Exit;
  var Prefix := UpperCase(ExpandFileName(AFile)) + '|';
  for var S in FKeys do
    if S.StartsWith(Prefix) then Exit(True);
end;

function TLspSymbolTargets.Count: Integer;
begin
  Result := Length(FKeys);
end;

function TLspSymbolTargets.Text: string;
begin
  Result := '';
  for var S in FKeys do
  begin
    var P := LastDelimiter('|', S);
    Result := Result + '  ' + ExtractFileName(Copy(S, 1, P - 1)) + ':' +
      IntToStr(StrToIntDef(Copy(S, P + 1, MaxInt), -1) + 1) + sLineBreak;
  end;
end;

end.
