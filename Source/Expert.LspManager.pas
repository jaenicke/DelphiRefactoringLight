(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.LspManager;

interface

uses
  System.SysUtils, System.Classes, System.SyncObjs, Lsp.Client,
  System.Generics.Collections;

type
  /// <summary>
  ///  Singleton-Verwaltung fuer den DelphiLSP-Client.
  ///  Haelt den LSP-Prozess am Leben zwischen Rename-Operationen.
  ///  Erster Aufruf dauert je nach Projekt mehrere Sek. (Start + Indexierung).
  ///  Folgeaufrufe: sofort.
  /// </summary>
  /// <summary>Callback fuer Fortschritt beim Warmlaufen des LSP-Index.
  ///  ACurrent/ATotal sind 1-basiert. ACurrentFile ist der aktuell
  ///  bearbeitete Dateipfad (kann leer sein wenn abgeschlossen).</summary>
  TLspIndexProgress = reference to procedure(ACurrent, ATotal: Integer; const ACurrentFile: string);

  TLspManager = class
  private
    class var
      FInstance: TLspManager;
    var
      FClient: TLspClient;
      FCurrentProject: string;     // .delphilsp.json Pfad
      FCurrentRootPath: string;
      FIsReady: Boolean;
      FProjectIndexed: Boolean;    // Alle Projektdateien via didOpen bekannt?
      FLspExePath: string;
      // THREAD SAFETY: GetClient is called from worker threads (completion,
      // signature help, prewarmer, live checker) as well as the main thread.
      // FLock serialises every state change - without it a second caller
      // saw "not ready yet", restarted, and FREED the client the first one
      // was still initialising. A replaced client is never freed while a
      // worker may hold it: it is shut down (requests fail fast) and parked
      // in FRetired; the objects go only after RetiredGraceMs or when the
      // manager itself is destroyed (after the worker latch has drained).
      FLock: TCriticalSection;
      FRetired: TList<TPair<TLspClient, UInt64>>;
    constructor CreatePrivate;
    procedure RetireClient;
    procedure SweepRetired(AAll: Boolean);
  public
    destructor Destroy; override;

    /// <summary>Gibt den Singleton zurueck. Erstellt ihn bei Bedarf.</summary>
    class function Instance: TLspManager;

    /// <summary>Gibt den LSP-Client zurueck, initialisiert bei Bedarf.
    ///  Startet den LSP neu wenn sich das Projekt geaendert hat.</summary>
    function GetClient(const ARootPath, AProjectFile, ADelphiLspJson: string): TLspClient;

    /// <summary>Gibt den BEREITS laufenden Client zurueck, ohne einen
    ///  neuen zu starten. Fuer Hintergrund-Checks (Auto-Import live), die
    ///  keinen Kaltstart ausloesen und nicht mit einer parallelen
    ///  GetClient-Erzeugung kollidieren duerfen. Nil solange kein Client
    ///  initialisiert ist.</summary>
    function PeekClient: TLspClient;

    /// <summary>Sorgt dafuer, dass alle uebergebenen Projektdateien via
    ///  textDocument/didOpen im LSP bekannt sind. Das ist zwingend noetig
    ///  fuer projektweite Queries wie textDocument/implementation, die
    ///  ansonsten nur in der aktuellen Datei suchen.
    ///  Idempotent: wiederholte Aufrufe ohne Projektwechsel tun nichts.</summary>
    procedure EnsureProjectIndexed(const AProjectFiles: TArray<string>; AProgress: TLspIndexProgress = nil);

    /// <summary>Prueft ob der LSP-Client noch laeuft: Prozess lebt UND der
    ///  Lese-Thread laeuft (ohne Round-Trip).</summary>
    function IsAlive: Boolean;

    /// <summary>Beendet den LSP-Client (z.B. beim Entladen des Experts).</summary>
    procedure Shutdown;

    /// <summary>Setzt den LSP-Warmup-Status als Suffix der Dialog-
    ///  Caption an, damit der User beim Start einer Aktion sieht, ob
    ///  LSP bereit ist. ADialog muss eine TCustomForm-Instanz sein.
    ///  Ueberschreibt eine evtl. schon angeklebte Status-Klammer am
    ///  Ende der Caption.</summary>
    procedure ApplyStatusToCaption(ADialog: TObject);

    /// <summary>Kurzer menschenlesbarer Status fuer Dialog-Anzeigen:
    ///   "LSP not started"
    ///   "LSP cold-starting"
    ///   "LSP indexing - N diagnostics, M inactive ranges so far"
    ///   "LSP ready - N diagnostics, M inactive ranges"
    /// Soll an die erste Status-Zeile eines Wizard-Dialogs angeklebt
    /// werden, damit der User vor dem Start einer Aktion sieht, ob der
    /// LSP-Pre-Warmer schon durch ist.</summary>
    function GetWarmupStatusLine: string;

    /// <summary>Setzt den internen Zustand zurueck (erzwingt Neustart beim naechsten Aufruf).</summary>
    procedure Reset;

    /// <summary>Zeigt an, ob der Index (siehe EnsureProjectIndexed)
    ///  bereits aufgebaut wurde.</summary>
    property ProjectIndexed: Boolean read FProjectIndexed;

    class procedure FreeInstance; reintroduce;

    /// <summary>Ends a running DelphiLSP (if any) without creating the
    ///  manager. At unload this wakes every worker that waits for an
    ///  answer - their requests fail at once instead of timing out.</summary>
    class procedure ShutdownIfRunning;
  end;

implementation

uses
  Winapi.Windows, System.IOUtils, System.DateUtils, System.JSON,
  System.Win.Registry, System.TypInfo, Expert.PluginSettings
  {$IFNDEF STANDALONE_BUILD}, ToolsAPI, Expert.EditorHelper {$ENDIF};

const
  /// <summary>How long a replaced client object stays alive for workers
  ///  that may still hold it (the longest worker waits ~30 s).</summary>
  RetiredGraceMs = 10 * 60 * 1000;
  /// <summary>Absolute Notfall-Fallback, falls weder Registry noch ToolsAPI
  ///  eine brauchbare Antwort liefern. Wird nur benutzt wenn die IDE einen
  ///  voellig defekten Registry-Stand hat.</summary>
  FallbackLspPath = 'C:\Program Files (x86)\Embarcadero\Studio\37.0\bin\DelphiLsp.exe';

{ File-private helpers (class-wrapped to keep the unit free of globals) }

type
  TLspPathResolver = class
  public
    /// <summary>Reads the BDS RootDir from the registry (HKCU first,
    ///  then HKLM). Returns the value including a trailing backslash
    ///  or '' on failure.</summary>
    class function ReadBdsRootDirFromRegistry(const ABaseKey: string): string; static;

    /// <summary>Determines the path to DelphiLsp.exe. Order:
    ///    1. Registry via IOTAServices.GetBaseRegistryKey (HKCU -> HKLM)
    ///    2. IOTAServices.GetRootDirectory (running IDE)
    ///    3. FallbackLspPath (hard-coded).</summary>
    class function ResolveLspExePath: string; static;
  end;

class function TLspPathResolver.ReadBdsRootDirFromRegistry(const ABaseKey: string): string;

  function TryRead(ARootKey: HKEY; const AKey: string): string;
  var
    Reg: TRegistry;
  begin
    Result := '';
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := ARootKey;
      if Reg.OpenKeyReadOnly(AKey) then
      try
        if Reg.ValueExists('RootDir') then
          Result := Reg.ReadString('RootDir');
      finally
        Reg.CloseKey;
      end;
    finally
      Reg.Free;
    end;
  end;

begin
  Result := TryRead(HKEY_CURRENT_USER, ABaseKey);
  if Result = '' then
    Result := TryRead(HKEY_LOCAL_MACHINE, ABaseKey);
  if Result <> '' then
    Result := IncludeTrailingPathDelimiter(Result);
end;

class function TLspPathResolver.ResolveLspExePath: string;
var
  {$IFNDEF STANDALONE_BUILD}
  Services: IOTAServices;
  {$ENDIF}
  BaseKey, RootDir, Candidate: string;
begin
  // 1. Get the registry base-key via ToolsAPI (e.g. "Software\Embarcadero\BDS\37.0")
  BaseKey := '';
  {$IFNDEF STANDALONE_BUILD}
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  try
    BaseKey := Services.GetBaseRegistryKey;
  except
    BaseKey := '';
  end;
  {$ENDIF}

  // 1a. If ToolsAPI gives no key (or this is the standalone build), try a
  //     sensible registry probe (descending so newer BDS versions win).
  if BaseKey = '' then
  begin
    for var Ver in TArray<string>.Create('37.0', '23.0', '22.0', '21.0', '20.0') do
    begin
      RootDir := ReadBdsRootDirFromRegistry('Software\Embarcadero\BDS\' + Ver);
      if RootDir <> '' then Break;
    end;
  end
  else
    RootDir := ReadBdsRootDirFromRegistry(BaseKey);

  if RootDir <> '' then
  begin
    Candidate := RootDir + 'bin\DelphiLsp.exe';
    if TFile.Exists(Candidate) then
      Exit(Candidate);
  end;

  {$IFNDEF STANDALONE_BUILD}
  // 2. Ask ToolsAPI directly.
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  try
    RootDir := Services.GetRootDirectory;
    if RootDir <> '' then
    begin
      Candidate := IncludeTrailingPathDelimiter(RootDir) + 'bin\DelphiLsp.exe';
      if TFile.Exists(Candidate) then
        Exit(Candidate);
    end;
  except
    // ignorieren
  end;
  {$ENDIF}

  // 3. Notfall
  Result := FallbackLspPath;
end;

{ TLspManager }

constructor TLspManager.CreatePrivate;
begin
  inherited Create;
  FLock := TCriticalSection.Create;
  FRetired := TList<TPair<TLspClient, UInt64>>.Create;
  FLspExePath := TLspPathResolver.ResolveLspExePath;
  FIsReady := False;
  FProjectIndexed := False;
end;

destructor TLspManager.Destroy;
begin
  Shutdown;
  SweepRetired(True);   // the worker latch has drained by now
  FRetired.Free;
  FLock.Free;
  inherited;
end;

class function TLspManager.Instance: TLspManager;
var
  M: TLspManager;
begin
  // Workers call this too - create exactly one instance.
  if FInstance = nil then
  begin
    M := TLspManager.CreatePrivate;
    if AtomicCmpExchange(Pointer(FInstance), Pointer(M), nil) <> nil then
      M.Free;
  end;
  Result := FInstance;
end;

// Caller holds FLock.
procedure TLspManager.RetireClient;
begin
  if FClient <> nil then
  begin
    try
      FClient.Shutdown;   // ends the process: pending waits wake up and fail
    except
      // Shutdown-Fehler ignorieren
    end;
    FRetired.Add(TPair<TLspClient, UInt64>.Create(FClient, GetTickCount64));
    FClient := nil;
  end;
  FIsReady := False;
  FProjectIndexed := False;
  SweepRetired(False);
end;

// Caller holds FLock (or is the destructor).
procedure TLspManager.SweepRetired(AAll: Boolean);
begin
  for var I := FRetired.Count - 1 downto 0 do
    if AAll or (GetTickCount64 - FRetired[I].Value > RetiredGraceMs) then
    begin
      try
        FRetired[I].Key.Free;
      except
        // a broken client must not stop the sweep
      end;
      FRetired.Delete(I);
    end;
end;

class procedure TLspManager.FreeInstance;
begin
  FreeAndNil(FInstance);
end;

class procedure TLspManager.ShutdownIfRunning;
begin
  if FInstance <> nil then
    FInstance.Shutdown;
end;

function TLspManager.IsAlive: Boolean;
var
  C: TLspClient;
begin
  // A real probe now: the old version had an EMPTY try block and answered
  // "alive" for a dead reader thread, so every command then waited out its
  // full timeout. A replaced client is never freed under us (FRetired), so
  // reading the pointer without the lock is safe.
  C := FClient;
  Result := (C <> nil) and FIsReady and C.IsConnected;
end;

function TLspManager.PeekClient: TLspClient;
begin
  // Never block a poller: while another thread is (re)starting the client
  // under the lock, there simply is no ready client yet.
  if not FLock.TryEnter then Exit(nil);
  try
    if (FClient <> nil) and FIsReady and FClient.IsConnected then
      Result := FClient
    else
      Result := nil;
  finally
    FLock.Leave;
  end;
end;

function TLspManager.GetClient(const ARootPath, AProjectFile, ADelphiLspJson: string): TLspClient;
var
  NeedRestart: Boolean;
begin
  FLock.Enter;
  try
  NeedRestart := False;

  // Neustart noetig wenn:
  // 1. Kein Client vorhanden
  // 2. Projekt hat sich geaendert
  // 3. Client ist nicht mehr bereit
  if FClient = nil then
    NeedRestart := True
  else if not SameText(FCurrentProject, ADelphiLspJson) then
    NeedRestart := True
  else if not FIsReady then
    NeedRestart := True
  else if not FClient.IsConnected then
    NeedRestart := True;   // process or reader thread died - start afresh

  if NeedRestart then
  begin
    // Alten Client beenden (nicht freigeben - ein Worker koennte ihn halten)
    RetireClient;

    // Neuen Client starten
    FClient := TLspClient.Create(FLspExePath);
    // Diagnosis only: REFACTORINGLIGHT_LSP_ARGS=-LogModes 255 makes our
    // session write %TEMP%\DelphiLSP\DelphiLSP.log (+ trace) - the only
    // way to see WHY DelphiLSP answers nothing for a unit.
    FClient.ExtraArgs := GetEnvironmentVariable('REFACTORINGLIGHT_LSP_ARGS');
    // ... or from the options page (issue #13 asked for it): the log names
    // every request and how long it took, which is the only way to tell a
    // slow session from a broken one. Our own name keeps it apart from the
    // IDE's instance in the same folder.
    if (FClient.ExtraArgs = '') and TPluginSettings.LspLogging then
      FClient.ExtraArgs := '-LogModes 255 -Name RefactoringLight';
    try
      FClient.Start;
      FClient.Initialize(ARootPath, AProjectFile);
      FClient.SendConfiguration(ADelphiLspJson);

      FCurrentProject := ADelphiLspJson;
      FCurrentRootPath := ARootPath;
      FIsReady := True;
      FProjectIndexed := False; // neu gestartet -> Index muss neu aufgebaut werden
    except
      // Never handed out (other callers wait on FLock) - safe to free.
      FreeAndNil(FClient);
      FIsReady := False;
      FProjectIndexed := False;
      raise;
    end;
  end;

  Result := FClient;
  finally
    FLock.Leave;
  end;
end;

procedure TLspManager.EnsureProjectIndexed(const AProjectFiles: TArray<string>; AProgress: TLspIndexProgress);
var
  I, N: Integer;
  Client: TLspClient;
begin
  if FProjectIndexed then Exit;
  // A local reference: a restart meanwhile retires FClient but never frees
  // it while we may still use it.
  Client := FClient;
  if Client = nil then Exit;

  N := Length(AProjectFiles);
  if N = 0 then
  begin
    FProjectIndexed := True;
    Exit;
  end;

  // Jede Datei via didOpen dem LSP bekannt machen, damit projektweite
  // Queries (implementation, references) alle Dateien beruecksichtigen
  // koennen. Wichtig: hier NICHT RefreshDocument benutzen - das macht
  // didClose + 50ms Sleep + didOpen + didChange pro Datei. Auf einem
  // 1200-Datei-Projekt waeren das ueber 60 Sekunden allein an Sleep,
  // plus doppelter JSON-Traffic. Beim Cold-Start kennt der LSP die
  // Dateien noch gar nicht, also reicht ein blankes didOpen. Spaetere
  // Edits laufen weiter ueber RefreshDocument (das wurde gegen einen
  // controller-mode Hover-Bug eingefuehrt - der ist bei der initialen
  // Indexierung irrelevant).
  for I := 0 to N - 1 do
  begin
    if Assigned(AProgress) then
      AProgress(I + 1, N, AProjectFiles[I]);
    try
      Client.OpenDocument(AProjectFiles[I]);
    except
      // Einzelne Datei-Fehler ignorieren (z.B. fehlende Datei)
    end;
  end;

  // Aktiv warten, bis der LSP-Index wirklich nutzbar ist. didOpen liefert
  // sofort zurueck, aber DelphiLSP indexiert erst danach (Cold-Start dauert
  // bei brandneuen Projekten leicht 10-30 s, weil VCL eingelesen werden
  // muss). Zwei Stufen, weil documentSymbol frueher antwortet als
  // definition / references (beobachtet: erste GotoDefinition-Anfrage
  // schlaegt mit 'Server not responding' fehl, zweite klappt).
  //   Stufe 1: documentSymbol liefert nicht-leere Symbol-Liste
  //   Stufe 2: GotoDefinition fuer ein bekanntes Symbol antwortet ohne
  //            Timeout. Position = Anfang des ersten Symbols.
  // Max. 30 s gesamt, in 500 ms-Schritten.
  begin
    // Probe phase: didOpen was fast (~1s for 1200 files now that the
    // Sleep is gone), but DelphiLSP still needs time to actually build
    // its index. We report time-based progress here - elapsed seconds
    // of a 30s deadline - so the bar visibly keeps climbing instead of
    // sitting at 100% pretending to be done. Both stages share the
    // deadline; the elapsed counter just keeps counting through.
    var ProbeFile := AProjectFiles[0];
    var ProbeStart := Now;
    const ProbeMaxSec = 30;
    var Deadline := ProbeStart + ProbeMaxSec / SecsPerDay;
    var SymbolsReady := False;
    var ProbeLine: Integer := 0;
    var ProbeCol: Integer := 0;
    while (not SymbolsReady) and (Now < Deadline) do
    begin
      try
        var SymJson := Client.GetDocumentSymbols(ProbeFile);
        try
          if (SymJson <> nil) and (SymJson.Count > 0) then
          begin
            SymbolsReady := True;
            // Erste 'selectionRange.start' aus der Symbolliste merken
            // (wenn nicht da, bleibt es bei 0:0 - LSP antwortet eh nur
            // mit 0 Treffern, aber das ist auch ein gueltiges Signal).
            var First := SymJson.Items[0];
            if First is TJSONObject then
            begin
              var SR := TJSONObject(First).GetValue('selectionRange');
              if SR is TJSONObject then
              begin
                var SP := TJSONObject(SR).GetValue('start');
                if SP is TJSONObject then
                begin
                  ProbeLine := TJSONObject(SP).GetValue<Integer>('line', 0);
                  ProbeCol  := TJSONObject(SP).GetValue<Integer>('character', 0);
                end;
              end;
            end;
          end;
        finally
          SymJson.Free;
        end;
      except
        // Probe-Fehler nicht weiterreichen
      end;
      if not SymbolsReady then
      begin
        if Assigned(AProgress) then
        begin
          var Elapsed := Round((Now - ProbeStart) * SecsPerDay);
          AProgress(Elapsed, ProbeMaxSec,
            Format('Waiting for LSP symbols... (%d / %d s)', [Elapsed, ProbeMaxSec]));
        end;
        Sleep(500);
      end;
    end;

    // Stufe 2: GotoDefinition aktivieren. Wir akzeptieren JEDE Antwort,
    // die keine Exception/Timeout ist - leeres Locations-Array ist OK
    // (LSP arbeitet, kennt das Symbol vielleicht nur nicht).
    var DefReady := False;
    while (not DefReady) and (Now < Deadline) do
    begin
      try
        var Defs := Client.GotoDefinition(ProbeFile, ProbeLine, ProbeCol);
        // Kein Throw == Server antwortet. Locations-Count egal.
        DefReady := True;
        if Length(Defs) = 0 then ; // explizit ignorieren
      except
        // Server not responding -> nochmal warten
      end;
      if not DefReady then
      begin
        if Assigned(AProgress) then
        begin
          var Elapsed := Round((Now - ProbeStart) * SecsPerDay);
          AProgress(Elapsed, ProbeMaxSec,
            Format('Waiting for LSP index... (%d / %d s)', [Elapsed, ProbeMaxSec]));
        end;
        Sleep(500);
      end;
    end;
  end;

  if Assigned(AProgress) then
    AProgress(N, N, '');

  // only for the client we indexed - not for one started meanwhile
  if Client = FClient then
    FProjectIndexed := True;
end;

procedure TLspManager.Shutdown;
begin
  FLock.Enter;
  try
    RetireClient;
    FCurrentProject := '';
    FCurrentRootPath := '';
  finally
    FLock.Leave;
  end;
end;

procedure TLspManager.Reset;
begin
  FLock.Enter;
  try
    RetireClient;
    FCurrentProject := '';
  finally
    FLock.Leave;
  end;
end;

procedure TLspManager.ApplyStatusToCaption(ADialog: TObject);
var
  CapProp: string;
  StatusBracketStart: Integer;
begin
  if not (ADialog is TComponent) then Exit;
  // We use RTTI-free access via VCL.Forms TForm - but to avoid dragging
  // VCL.Forms into LspManager, we just use a published 'Caption' string
  // property via SetPropValue. Component-level access is fine for any
  // TForm/TFrame.
  if not IsPublishedProp(ADialog, 'Caption') then Exit;
  CapProp := GetStrProp(ADialog, 'Caption');
  // Strip a previously appended status (anything in trailing '  [...]').
  StatusBracketStart := LastDelimiter('[', CapProp);
  if (StatusBracketStart > 1)
    and (Copy(CapProp, StatusBracketStart - 2, 2) = '  ') then
    CapProp := TrimRight(Copy(CapProp, 1, StatusBracketStart - 1));
  SetStrProp(ADialog, 'Caption', CapProp + '  [' + GetWarmupStatusLine + ']');
end;

function TLspManager.GetWarmupStatusLine: string;
// Primary readiness signal is FProjectIndexed (set by EnsureProject-
// Indexed once every project file has been pushed through didOpen).
// Diagnostics are a *bonus* signal - DelphiLSP in single-process mode
// often never sends publishDiagnostics, so we must not gate "ready"
// on Diag > 0. Otherwise wizards that work just fine (rename, find-
// references, completion) look broken because the status reads
// "cold-starting" forever.
var
  Diag, Inactive: Integer;
  C: TLspClient;
begin
  C := FClient;   // never freed under us (see FRetired)
  if C = nil then
    Exit('LSP not started yet - first action will trigger a cold start (~10-30 s).');
  Diag := C.GetDiagnosticsCount;
  Inactive := C.GetInactiveRangesTotal;
  if not FProjectIndexed then
  begin
    if Diag = 0 then
      Result := 'LSP starting - indexing project files...'
    else
      Result := Format(
        'LSP starting - %d diagnostics, %d inactive regions so far.',
        [Diag, Inactive]);
  end
  else
  begin
    if Diag = 0 then
      Result := 'LSP ready (server did not publish diagnostics).'
    else
      Result := Format(
        'LSP ready - %d diagnostics, %d inactive regions analysed.',
        [Diag, Inactive]);
  end;
end;

initialization

finalization
  TLspManager.FreeInstance;

end.
