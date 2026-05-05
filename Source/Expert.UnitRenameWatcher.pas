(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnitRenameWatcher;

{
  Watches for unit renames in the Delphi IDE and, when one is detected,
  triggers the normal rename dialog pre-filled with (OldUnitName ->
  NewUnitName). The user sees the usual preview list and confirms the
  rename; replacements in 'uses' clauses (and qualified references like
  'OldUnit.Something') are applied via the same editor API as an ordinary
  identifier rename, so Ctrl+Z undoes them.

  Mechanism:
  * TModuleRenameNotifier implements IOTAModuleNotifier and is attached
    to every open module. When the IDE renames the module (File > Save
    As..., or project manager rename), ModuleRenamed(NewName) fires and
    gives us the new file name. The previous file name is captured at
    notifier-create time.
  * TIdeFileNotifier implements IOTAIDENotifier and listens for
    ofnFileOpened so that newly opened modules also receive a notifier.
  * TUnitRenameWatcher owns both notifiers and forwards the rename event
    to TLspRenameWizard.ExecuteForUnit.
}

interface

uses
  System.Classes, System.Generics.Collections, ToolsAPI;

type
  TUnitRenameWatcher = class;

  /// <summary>Per-module IOTAModuleNotifier. Captures the current file
  ///  name at creation time and, on ModuleRenamed, forwards (old, new)
  ///  to the owning watcher.</summary>
  TModuleRenameNotifier = class(TNotifierObject, IOTANotifier,
    IOTAModuleNotifier)
  private
    FOwner: TUnitRenameWatcher;
    FCurrentFileName: string;
  public
    constructor Create(AOwner: TUnitRenameWatcher; const ACurrentFileName: string);
    // IOTAModuleNotifier
    function CheckOverwrite: Boolean;
    procedure ModuleRenamed(const NewName: string);
  end;

  /// <summary>IDE-global IOTAIDENotifier. Catches ofnFileOpened so that
  ///  newly opened modules also receive a per-module notifier.</summary>
  TIdeFileNotifier = class(TNotifierObject, IOTANotifier, IOTAIDENotifier)
  private
    FOwner: TUnitRenameWatcher;
  public
    constructor Create(AOwner: TUnitRenameWatcher);
    // IOTAIDENotifier
    procedure FileNotification(NotifyCode: TOTAFileNotification; const FileName: string; var Cancel: Boolean);
    procedure BeforeCompile(const Project: IOTAProject; var Cancel: Boolean);
    procedure AfterCompile(Succeeded: Boolean);
  end;

  TUnitRenameWatcher = class
  private
    type
      TAttachedEntry = record
        Module: IOTAModule;
        NotifierIndex: Integer;
      end;
    var
      FIdeNotifier: IOTAIDENotifier;
      FIdeNotifierIndex: Integer;
      FAttached: TList<TAttachedEntry>;
    procedure AttachToAllOpenModules;
    function IsAttachedTo(const AFileName: string): Boolean;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Install;
    procedure Uninstall;

    /// <summary>Attach a per-module notifier to AModule. No-op if the
    ///  module is already tracked.</summary>
    procedure AttachToModule(AModule: IOTAModule);

    /// <summary>Called by TModuleRenameNotifier when the IDE renames a
    ///  module. Forwards the event to the rename wizard.</summary>
    procedure HandleRename(const AOldFileName, ANewFileName: string);

    /// <summary>Called by TIdeFileNotifier when a file is opened in the
    ///  IDE. Attaches a per-module notifier if possible.</summary>
    procedure HandleFileOpened(const AFileName: string);
  end;

var
  UnitRenameWatcherInstance: TUnitRenameWatcher;

implementation

uses
  System.SysUtils,
  Expert.RenameWizard;

{ TModuleRenameNotifier }

constructor TModuleRenameNotifier.Create(AOwner: TUnitRenameWatcher; const ACurrentFileName: string);
begin
  inherited Create;
  FOwner := AOwner;
  FCurrentFileName := ACurrentFileName;
end;

function TModuleRenameNotifier.CheckOverwrite: Boolean;
begin
  Result := True;
end;

procedure TModuleRenameNotifier.ModuleRenamed(const NewName: string);
var
  OldName: string;
begin
  OldName := FCurrentFileName;
  FCurrentFileName := NewName;
  if (FOwner <> nil) and (OldName <> '') and not SameText(OldName, NewName) then
    FOwner.HandleRename(OldName, NewName);
end;

{ TIdeFileNotifier }

constructor TIdeFileNotifier.Create(AOwner: TUnitRenameWatcher);
begin
  inherited Create;
  FOwner := AOwner;
end;

procedure TIdeFileNotifier.FileNotification(NotifyCode: TOTAFileNotification; const FileName: string; var Cancel: Boolean);
begin
  if (NotifyCode = ofnFileOpened) and (FOwner <> nil) then
    FOwner.HandleFileOpened(FileName);
end;

procedure TIdeFileNotifier.BeforeCompile(const Project: IOTAProject;
  var Cancel: Boolean);
begin
  // not needed
end;

procedure TIdeFileNotifier.AfterCompile(Succeeded: Boolean);
begin
  // not needed
end;

{ TUnitRenameWatcher }

constructor TUnitRenameWatcher.Create;
begin
  inherited Create;
  FAttached := TList<TAttachedEntry>.Create;
  FIdeNotifierIndex := -1;
end;

destructor TUnitRenameWatcher.Destroy;
begin
  Uninstall;
  FAttached.Free;
  inherited;
end;

procedure TUnitRenameWatcher.Install;
var
  Services: IOTAServices;
begin
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  begin
    FIdeNotifier := TIdeFileNotifier.Create(Self);
    FIdeNotifierIndex := Services.AddNotifier(FIdeNotifier);
  end;
  AttachToAllOpenModules;
end;

procedure TUnitRenameWatcher.Uninstall;
var
  Services: IOTAServices;
  Entry: TAttachedEntry;
begin
  for Entry in FAttached do
  begin
    if (Entry.Module <> nil) and (Entry.NotifierIndex >= 0) then
    begin
      try
        Entry.Module.RemoveNotifier(Entry.NotifierIndex);
      except
        // ignore - module may be closing
      end;
    end;
  end;
  FAttached.Clear;

  if (FIdeNotifierIndex >= 0) and
     Supports(BorlandIDEServices, IOTAServices, Services) then
  begin
    try
      Services.RemoveNotifier(FIdeNotifierIndex);
    except
      // ignore
    end;
  end;
  FIdeNotifierIndex := -1;
  FIdeNotifier := nil;
end;

procedure TUnitRenameWatcher.AttachToAllOpenModules;
var
  ModuleServices: IOTAModuleServices;
  I: Integer;
begin
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
    Exit;
  for I := 0 to ModuleServices.ModuleCount - 1 do
    AttachToModule(ModuleServices.Modules[I]);
end;

function TUnitRenameWatcher.IsAttachedTo(const AFileName: string): Boolean;
var
  Entry: TAttachedEntry;
  EntryName: string;
begin
  for Entry in FAttached do
  begin
    if Entry.Module = nil then Continue;
    EntryName := '';
    try
      EntryName := Entry.Module.FileName;
    except
      // module may be partially constructed/destructed - skip
      Continue;
    end;
    if SameText(EntryName, AFileName) then
      Exit(True);
  end;
  Result := False;
end;

procedure TUnitRenameWatcher.AttachToModule(AModule: IOTAModule);
var
  Notifier: IOTAModuleNotifier;
  Entry: TAttachedEntry;
  ModuleFileName: string;
begin
  if AModule = nil then Exit;

  ModuleFileName := '';
  try
    ModuleFileName := AModule.FileName;
  except
    // module not fully constructed yet
    Exit;
  end;
  if ModuleFileName = '' then Exit;
  if IsAttachedTo(ModuleFileName) then Exit;

  Notifier := TModuleRenameNotifier.Create(Self, ModuleFileName);
  Entry.Module := AModule;
  try
    Entry.NotifierIndex := AModule.AddNotifier(Notifier);
    FAttached.Add(Entry);
  except
    // attaching may fail for some module types; silently ignore
  end;
end;

procedure TUnitRenameWatcher.HandleFileOpened(const AFileName: string);
var
  FileNameCopy: string;
begin
  if AFileName = '' then Exit;
  // The IDE fires ofnFileOpened from inside TBaseProject.AfterConstruction,
  // before the module is fully usable. Reading IOTAModule.FileName too early
  // crashes in TCustomCodeIProject.IBaseModule_GetFileName. Defer the attach
  // to the next message-loop tick so construction is finished.
  FileNameCopy := AFileName;
  TThread.ForceQueue(nil,
    procedure
    var
      ModuleServices: IOTAModuleServices;
      Module: IOTAModule;
    begin
      if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
        Exit;
      try
        Module := ModuleServices.FindModule(FileNameCopy);
      except
        Module := nil;
      end;
      if Module <> nil then
        AttachToModule(Module);
    end);
end;

procedure TUnitRenameWatcher.HandleRename(const AOldFileName, ANewFileName: string);
var
  OldUnit, NewUnit: string;
begin
  // Only consider Pascal source files
  if not SameText(ExtractFileExt(ANewFileName), '.pas') then Exit;

  OldUnit := ChangeFileExt(ExtractFileName(AOldFileName), '');
  NewUnit := ChangeFileExt(ExtractFileName(ANewFileName), '');

  // Same unit name (e.g. moved to a different folder) - nothing to do.
  if SameText(OldUnit, NewUnit) or (OldUnit = '') or (NewUnit = '') then
    Exit;

  // Queue to the main message loop so we run after the IDE has finished
  // its rename processing.
  TThread.ForceQueue(nil,
    procedure
    begin
      if WizardInstance <> nil then
        WizardInstance.ExecuteForUnit(OldUnit, NewUnit);
    end);
end;

end.
