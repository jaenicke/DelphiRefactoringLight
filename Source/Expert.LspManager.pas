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
  System.SysUtils, System.Classes, Lsp.Client;

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

    /// <summary>Setzt den internen Zustand zurueck (erzwingt Neustart beim naechsten Aufruf).</summary>
    procedure Reset;

    /// <summary>Zeigt an, ob der Index (siehe EnsureProjectIndexed)
    ///  bereits aufgebaut wurde.</summary>
    property ProjectIndexed: Boolean read FProjectIndexed;

    class procedure FreeInstance; reintroduce;
  end;

implementation

uses
  Winapi.Windows, System.IOUtils, System.Win.Registry, ToolsAPI, Expert.EditorHelper;

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
  Services: IOTAServices;
  BaseKey, RootDir, Candidate: string;
begin
  // 1. Get the registry base-key via ToolsAPI (e.g. "Software\Embarcadero\BDS\37.0")
  BaseKey := '';
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  try
    BaseKey := Services.GetBaseRegistryKey;
  except
    BaseKey := '';
  end;

  // 1a. If ToolsAPI gives no key, try a sensible sequence
  //     (descending so newer BDS versions win).
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

  // Jede Datei via didOpen dem LSP bekannt machen, damit
  // projektweite Queries (implementation, references) alle Dateien
  // beruecksichtigen koennen. RefreshDocument = didClose + didOpen
  // mit kurzer Pause dazwischen (in Lsp.Client definiert).
  for I := 0 to N - 1 do
  begin
    if Assigned(AProgress) then
      AProgress(I + 1, N, AProjectFiles[I]);
    try
      FClient.RefreshDocument(AProjectFiles[I]);
    except
      // Einzelne Datei-Fehler ignorieren (z.B. fehlende Datei)
    end;
  end;

  // LSP noch kurz parsen lassen bevor die Query kommt
  Sleep(500);

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

initialization

finalization
  TLspManager.FreeInstance;

end.
