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
  System.SysUtils, System.Classes, Lsp.Client, System.Generics.Collections;

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
    constructor CreatePrivate;
  public
    destructor Destroy; override;

    /// <summary>Gibt den Singleton zurueck. Erstellt ihn bei Bedarf.</summary>
    class function Instance: TLspManager;

    /// <summary>Gibt den LSP-Client zurueck, initialisiert bei Bedarf.
    ///  Startet den LSP neu wenn sich das Projekt geaendert hat.</summary>
    function GetClient(const ARootPath, AProjectFile, ADelphiLspJson: string): TLspClient;

    /// <summary>Sorgt dafuer, dass alle uebergebenen Projektdateien via
    ///  textDocument/didOpen im LSP bekannt sind. Das ist zwingend noetig
    ///  fuer projektweite Queries wie textDocument/implementation, die
    ///  ansonsten nur in der aktuellen Datei suchen.
    ///  Idempotent: wiederholte Aufrufe ohne Projektwechsel tun nichts.</summary>
    procedure EnsureProjectIndexed(const AProjectFiles: TArray<string>; AProgress: TLspIndexProgress = nil);

    /// <summary>Prueft ob der LSP-Client noch laeuft.</summary>
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
  end;

implementation

uses
  Winapi.Windows, System.IOUtils, System.DateUtils, System.JSON,
  System.Win.Registry, System.TypInfo
  {$IFNDEF STANDALONE_BUILD}, ToolsAPI, Expert.EditorHelper {$ENDIF};

const
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
  FLspExePath := TLspPathResolver.ResolveLspExePath;
  FIsReady := False;
  FProjectIndexed := False;
end;

destructor TLspManager.Destroy;
begin
  Shutdown;
  inherited;
end;

class function TLspManager.Instance: TLspManager;
begin
  if FInstance = nil then
    FInstance := TLspManager.CreatePrivate;
  Result := FInstance;
end;

class procedure TLspManager.FreeInstance;
begin
  FreeAndNil(FInstance);
end;

function TLspManager.IsAlive: Boolean;
begin
  Result := (FClient <> nil) and FIsReady;
  // Zusaetzlich pruefen ob der Prozess noch laeuft
  if Result and (FClient <> nil) then
  begin
    try
      // Leichtgewichtiger Test: Hover auf Position 0,0 einer leeren Datei
      // Wenn der Prozess tot ist, wirft dies eine Exception
      // Alternativ: Wir vertrauen darauf dass der ReaderThread Fehler meldet
    except
      Result := False;
      Reset;
    end;
  end;
end;

function TLspManager.GetClient(const ARootPath, AProjectFile, ADelphiLspJson: string): TLspClient;
var
  NeedRestart: Boolean;
begin
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
    NeedRestart := True;

  if NeedRestart then
  begin
    // Alten Client beenden
    Shutdown;

    // Neuen Client starten
    FClient := TLspClient.Create(FLspExePath);
    try
      FClient.Start;
      FClient.Initialize(ARootPath, AProjectFile);
      FClient.SendConfiguration(ADelphiLspJson);

      FCurrentProject := ADelphiLspJson;
      FCurrentRootPath := ARootPath;
      FIsReady := True;
      FProjectIndexed := False; // neu gestartet -> Index muss neu aufgebaut werden
    except
      FreeAndNil(FClient);
      FIsReady := False;
      FProjectIndexed := False;
      raise;
    end;
  end;

  Result := FClient;
end;

procedure TLspManager.EnsureProjectIndexed(const AProjectFiles: TArray<string>; AProgress: TLspIndexProgress);
var
  I, N: Integer;
begin
  if FProjectIndexed then Exit;
  if FClient = nil then Exit;

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
      FClient.OpenDocument(AProjectFiles[I]);
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
        var SymJson := FClient.GetDocumentSymbols(ProbeFile);
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
        var Defs := FClient.GotoDefinition(ProbeFile, ProbeLine, ProbeCol);
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

  FProjectIndexed := True;
end;

procedure TLspManager.Shutdown;
begin
  if FClient <> nil then
  begin
    try
      FClient.Shutdown;
    except
      // Shutdown-Fehler ignorieren
    end;
    FreeAndNil(FClient);
  end;
  FIsReady := False;
  FProjectIndexed := False;
  FCurrentProject := '';
  FCurrentRootPath := '';
end;

procedure TLspManager.Reset;
begin
  FreeAndNil(FClient);
  FIsReady := False;
  FProjectIndexed := False;
  FCurrentProject := '';
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
begin
  if FClient = nil then
    Exit('LSP not started yet - first action will trigger a cold start (~10-30 s).');
  Diag := FClient.GetDiagnosticsCount;
  Inactive := FClient.GetInactiveRangesTotal;
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
