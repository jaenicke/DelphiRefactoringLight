(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.EditorHelper;

interface

uses
  System.SysUtils, System.IOUtils, System.Types, System.Classes, System.Generics.Collections, Xml.XMLDoc, Xml.XMLIntf, ToolsAPI;

type
  TEditorContext = record
    FileName: string;
    Line: Integer;       // 1-basiert
    Column: Integer;     // 1-basiert
    WordAtCursor: string;
    ProjectFile: string;
    ProjectRoot: string;
    IsValid: Boolean;
  end;

  TEditorHelper = class
    class function GetCurrentContext: TEditorContext;
    class procedure ReloadModifiedFiles(const FilePaths: TArray<string>);
    class function GetCurrentProjectDproj: string;
    class function GetProjectRoot: string;
    class function GetProjectSearchPaths: string;
    /// <summary>Speichert alle geaenderten Dateien in der IDE (File > Save All).</summary>
    class procedure SaveAllFiles;

    /// <summary>Wendet ein TextEdit ueber die IDE-Editor-API an (Undo-faehig).
    ///  ALine und ACol sind 0-basiert. Gibt True zurueck wenn erfolgreich.</summary>
    class function ApplyEditViaEditor(const AFilePath: string;
      ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;

    /// <summary>Fuegt Text am Anfang einer Zeile ein, OHNE Auto-Indent.
    ///  ALine ist 1-basiert. Text sollte bereits komplett formatiert sein.</summary>
    class function InsertTextAtLineStart(const AFilePath: string;
      ALine: Integer; const AText: string): Boolean;

    /// <summary>Ersetzt den gesamten Inhalt einer Datei im IDE-Editor (Undo-faehig).</summary>
    class function ReplaceFileContent(const AFilePath: string;
      const ANewContent: string): Boolean;

    /// <summary>Liest den aktuellen Inhalt aus dem Editor-Buffer (nicht von der Platte).</summary>
    class function ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;

    /// <summary>Ersetzt eine einzelne Zeile (1-basiert) im Editor-Buffer.</summary>
    class function ReplaceLineAt(const AFilePath: string; ALine: Integer;
      const ANewContent: string): Boolean;

    /// <summary>Loescht eine einzelne Zeile (1-basiert) im Editor-Buffer.</summary>
    class function DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;

    /// <summary>Informiert den IDE-Form-Designer ueber Aenderungen an der
    ///  Klassenstruktur, damit die RTTI-Caches aktualisiert werden.
    ///  Verhindert AV beim Speichern nach Text-Edits an Form-Klassen.</summary>
    class procedure NotifyClassStructureChanged(const AFilePath: string);

    /// <summary>Ersetzt einen Bereich (1-basiert) durch neuen Text OHNE Auto-Indent.</summary>
    class function ReplaceSelection(const AFilePath: string;
      AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      const ANewText: string): Boolean;

    /// <summary>Oeffnet die Datei im IDE-Editor und positioniert den Cursor.
    ///  ALine und ACol sind 0-basiert (LSP-Konvention). Optional wird ein
    ///  Token der Laenge AHighlightLen markiert. Gibt True bei Erfolg.</summary>
    class function GotoLocation(const AFilePath: string;
      ALine, ACol: Integer; AHighlightLen: Integer = 0): Boolean;

    /// <summary>Locates the .delphilsp.json for the current project.</summary>
    class function FindDelphiLspJson: string;
    /// <summary>Returns all source files (.pas/.dpr/.dpk) of the current
    ///  project. Uses IOTAProject.GetModule so files that live outside
    ///  the project directory but are bound into the project are
    ///  included.</summary>
    class function GetProjectSourceFiles: TArray<string>;
    /// <summary>Extracts the search paths from a .dproj file and
    ///  collects every directory that contains .pas files.</summary>
    class function BuildSearchPathFromProject(
      const ADprojPath, ARootPath: string): string;
  end;

implementation

class function TEditorHelper.GetCurrentContext: TEditorContext;
var
  EditorServices: IOTAEditorServices;
  EditBuffer: IOTAEditBuffer;
  EditView: IOTAEditView;
  EditPos: TOTAEditPos;
  EditPosition: IOTAEditPosition;
  Module: IOTAModule;
begin
  Result := Default(TEditorContext);
  Result.IsValid := False;

  if not Supports(BorlandIDEServices, IOTAEditorServices, EditorServices) then
    Exit;

  EditBuffer := EditorServices.TopBuffer;
  if EditBuffer = nil then
    Exit;

  EditView := EditBuffer.TopView;
  if EditView = nil then
    Exit;

  // Cursorposition (1-basiert)
  EditPos := EditView.CursorPos;
  Result.Line := EditPos.Line;
  Result.Column := EditPos.Col;

  // Dateiname
  Module := EditBuffer.Module;
  if Module <> nil then
    Result.FileName := Module.FileName;

  // Bezeichner unter dem Cursor via EditPosition
  EditPosition := EditBuffer.EditPosition;
  if EditPosition <> nil then
  begin
    EditPosition.Save;
    try
      // Zum Anfang des Wortes zurueck
      while EditPosition.IsWordCharacter and (EditPosition.Column > 1) do
        EditPosition.MoveRelative(0, -1);
      if not EditPosition.IsWordCharacter then
        EditPosition.MoveRelative(0, 1);
      // Wort zeichenweise lesen
      var Word := '';
      while EditPosition.IsWordCharacter do
      begin
        Word := Word + EditPosition.Character;
        if not EditPosition.MoveRelative(0, 1) then
          Break;
      end;
      Result.WordAtCursor := Word;
    finally
      EditPosition.Restore;
    end;
  end;

  // Projektdatei
  Result.ProjectFile := GetCurrentProjectDproj;
  Result.ProjectRoot := GetProjectRoot;
  Result.IsValid := (Result.FileName <> '') and (Result.WordAtCursor <> '');
end;

class function TEditorHelper.GetCurrentProjectDproj: string;
var
  ProjectGroup: IOTAProjectGroup;
  ModuleServices: IOTAModuleServices;
  Project: IOTAProject;
  I: Integer;
begin
  Result := '';
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
    Exit;

  // Aktives Projekt suchen
  for I := 0 to ModuleServices.ModuleCount - 1 do
  begin
    if Supports(ModuleServices.Modules[I], IOTAProjectGroup, ProjectGroup) then
    begin
      Project := ProjectGroup.ActiveProject;
      if Project <> nil then
      begin
        Result := Project.FileName;
        Exit;
      end;
    end;
  end;

  // Fallback: Erstes Projekt nehmen
  for I := 0 to ModuleServices.ModuleCount - 1 do
  begin
    if Supports(ModuleServices.Modules[I], IOTAProject, Project) then
    begin
      if not Supports(Project, IOTAProjectGroup) then
      begin
        Result := Project.FileName;
        Exit;
      end;
    end;
  end;
end;

class function TEditorHelper.GetProjectRoot: string;
var
  DprojPath: string;
begin
  DprojPath := GetCurrentProjectDproj;
  if DprojPath <> '' then
    Result := ExtractFilePath(DprojPath)
  else
    Result := '';
end;

class procedure TEditorHelper.ReloadModifiedFiles(const FilePaths: TArray<string>);
var
  ModuleServices: IOTAModuleServices;
  ActionServices: IOTAActionServices;
  Module: IOTAModule;
  FilePath: string;
begin
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
    Exit;

  for FilePath in FilePaths do
  begin
    Module := ModuleServices.FindModule(FilePath);
    if Module <> nil then
      Module.Refresh(True);
  end;

  // Editor-Ansicht aktualisieren
  if Supports(BorlandIDEServices, IOTAActionServices, ActionServices) then
    ActionServices.ReloadFile(''{ leer = aktuelle Datei });
end;

class procedure TEditorHelper.SaveAllFiles;
var
  ModuleServices: IOTAModuleServices;
begin
  if Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
    ModuleServices.SaveAll;
end;

class function TEditorHelper.GotoLocation(const AFilePath: string;
  ALine, ACol: Integer; AHighlightLen: Integer): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  EditView: IOTAEditView;
  EditPos: TOTAEditPos;
  CharPos: TOTACharPos;
  I: Integer;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
    Exit;

  // Modul oeffnen falls noetig
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then
    Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then
    Exit;

  // Datei sichtbar machen
  Module.Show;

  // SourceEditor finden
  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then
    Exit;

  // Source-Editor sichtbar machen (z.B. wenn Form-View aktiv ist)
  SourceEditor.Show;
  if SourceEditor.GetEditViewCount = 0 then
    Exit;
  EditView := SourceEditor.GetEditView(0);
  if EditView = nil then
    Exit;

  // EditPos ist 1-basiert, Eingabe (LSP) 0-basiert
  EditPos.Line := ALine + 1;
  EditPos.Col := ACol + 1;
  EditView.CursorPos := EditPos;

  // Auf Bildschirmmitte zentrieren
  CharPos.Line := EditPos.Line;
  CharPos.CharIndex := EditPos.Col - 1;
  EditView.Center(CharPos.Line, CharPos.CharIndex);

  // Optional Token markieren: Block setzen
  if AHighlightLen > 0 then
  begin
    var EditBlock := EditView.Block;
    if EditBlock <> nil then
    begin
      EditBlock.Reset;
      EditBlock.BeginBlock;
      EditPos.Col := ACol + 1;
      EditView.CursorPos := EditPos;
      EditPos.Col := ACol + 1 + AHighlightLen;
      EditView.CursorPos := EditPos;
      EditBlock.EndBlock;
    end;
  end;

  EditView.Paint;
  Result := True;
end;

class function TEditorHelper.ApplyEditViaEditor(const AFilePath: string;
  ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Reader: IOTAEditReader;
  Writer: IOTAEditWriter;
  I: Integer;
  BufSize: Integer;
  Buf: TBytes;
  BytesRead: Integer;
  LineNum, OldTextLen: Integer;
begin
  Result := False;

  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
    Exit;

  // Modul in der IDE oeffnen (falls nicht schon offen)
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then
    Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then
    Exit;

  // SourceEditor finden
  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
  begin
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  end;
  if SourceEditor = nil then
    Exit;

  // Lineare Position berechnen ueber den Reader
  // Der Editor-Buffer ist UTF-8 kodiert mit #13#10 Zeilenenden
  Reader := SourceEditor.CreateReader;
  if Reader = nil then
    Exit;

  // Buffer-Inhalt lesen (bis zu 4 MB)
  BufSize := SourceEditor.GetLinesInBuffer * 200; // Schaetzung
  if BufSize < 65536 then BufSize := 65536;
  if BufSize > 4 * 1024 * 1024 then BufSize := 4 * 1024 * 1024;
  SetLength(Buf, BufSize);
  BytesRead := Reader.GetText(0, @Buf[0], BufSize);
  Reader := nil; // Reader freigeben BEVOR Writer erstellt wird!

  // Lineare Position finden: Zeilen zaehlen (0-basiert)
  LineNum := 0;
  I := 0;
  while (I < BytesRead) and (LineNum < ALine) do
  begin
    if Buf[I] = 10 then // LF = Zeilenende
      Inc(LineNum);
    Inc(I);
  end;
  // I zeigt jetzt auf den Anfang von Zeile ALine
  var LinearPos := I + ACol; // Spalte addieren (0-basiert, Bytes)

  // Laenge des alten Textes in UTF-8 Bytes
  OldTextLen := Length(UTF8Encode(AOldText));

  // UndoableWriter: CopyTo(Start), DeleteTo(Ende), Insert(NeuText)
  Writer := SourceEditor.CreateUndoableWriter;
  if Writer = nil then
    Exit;
  try
    Writer.CopyTo(LinearPos);
    Writer.DeleteTo(LinearPos + OldTextLen);
    Writer.Insert(PAnsiChar(UTF8Encode(ANewText)));
  finally
    Writer := nil; // Flush + Release
  end;

  Result := True;
end;

class function TEditorHelper.InsertTextAtLineStart(const AFilePath: string;
  ALine: Integer; const AText: string): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Reader: IOTAEditReader;
  Writer: IOTAEditWriter;
  I, BufSize, BytesRead, LineNum: Integer;
  Buf: TBytes;
  LinearPos: Integer;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then Exit;

  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then Exit;

  // Lineare Position von Zeilenanfang berechnen
  Reader := SourceEditor.CreateReader;
  if Reader = nil then Exit;
  BufSize := SourceEditor.GetLinesInBuffer * 200;
  if BufSize < 65536 then BufSize := 65536;
  if BufSize > 8 * 1024 * 1024 then BufSize := 8 * 1024 * 1024;
  SetLength(Buf, BufSize);
  BytesRead := Reader.GetText(0, @Buf[0], BufSize);
  Reader := nil;

  LineNum := 1; // 1-basiert
  I := 0;
  while (I < BytesRead) and (LineNum < ALine) do
  begin
    if Buf[I] = 10 then Inc(LineNum);
    Inc(I);
  end;
  LinearPos := I;

  // Writer: CopyTo + Insert ohne Auto-Indent
  Writer := SourceEditor.CreateUndoableWriter;
  if Writer = nil then Exit;
  try
    Writer.CopyTo(LinearPos);
    Writer.Insert(PAnsiChar(UTF8Encode(AText)));
  finally
    Writer := nil;
  end;
  Result := True;
end;

class function TEditorHelper.ReplaceFileContent(const AFilePath: string;
  const ANewContent: string): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Writer: IOTAEditWriter;
  I, BufSize, BytesRead: Integer;
  Buf: TBytes;
  Reader: IOTAEditReader;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then Exit;

  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then Exit;

  // Aktuelle Laenge ermitteln
  Reader := SourceEditor.CreateReader;
  if Reader = nil then Exit;
  BufSize := SourceEditor.GetLinesInBuffer * 200;
  if BufSize < 65536 then BufSize := 65536;
  if BufSize > 8 * 1024 * 1024 then BufSize := 8 * 1024 * 1024;
  SetLength(Buf, BufSize);
  BytesRead := Reader.GetText(0, @Buf[0], BufSize);
  Reader := nil;

  // Writer: alles loeschen und neuen Inhalt einfuegen
  Writer := SourceEditor.CreateUndoableWriter;
  if Writer = nil then Exit;
  try
    Writer.DeleteTo(BytesRead);
    Writer.Insert(PAnsiChar(UTF8Encode(ANewContent)));
  finally
    Writer := nil;
  end;
  Result := True;
end;

class function TEditorHelper.ReplaceSelection(const AFilePath: string;
  AStartLine, AStartCol, AEndLine, AEndCol: Integer;
  const ANewText: string): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Reader: IOTAEditReader;
  Writer: IOTAEditWriter;
  I, BufSize, BytesRead: Integer;
  Buf: TBytes;
  StartPos, EndPos: Integer;

  function FindLinearPos(ALine, ACol: Integer): Integer;
  var
    Idx, LN: Integer;
  begin
    LN := 1;
    Idx := 0;
    while (Idx < BytesRead) and (LN < ALine) do
    begin
      if Buf[Idx] = 10 then Inc(LN);
      Inc(Idx);
    end;
    // Spalte (1-basiert, in Bytes) addieren
    Result := Idx + (ACol - 1);
  end;

begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then Exit;

  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then Exit;

  // Reader: Buffer lesen und Positionen berechnen
  Reader := SourceEditor.CreateReader;
  if Reader = nil then Exit;
  BufSize := SourceEditor.GetLinesInBuffer * 200;
  if BufSize < 65536 then BufSize := 65536;
  if BufSize > 16 * 1024 * 1024 then BufSize := 16 * 1024 * 1024;
  SetLength(Buf, BufSize);
  BytesRead := Reader.GetText(0, @Buf[0], BufSize);
  Reader := nil;

  StartPos := FindLinearPos(AStartLine, AStartCol);
  EndPos := FindLinearPos(AEndLine, AEndCol);
  if EndPos < StartPos then EndPos := StartPos;

  // Writer: Bereich loeschen und neuen Text einfuegen (ohne Auto-Indent)
  Writer := SourceEditor.CreateUndoableWriter;
  if Writer = nil then Exit;
  try
    Writer.CopyTo(StartPos);
    Writer.DeleteTo(EndPos);
    Writer.Insert(PAnsiChar(UTF8Encode(ANewText)));
  finally
    Writer := nil;
  end;
  Result := True;
end;

class procedure TEditorHelper.NotifyClassStructureChanged(const AFilePath: string);
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  I: Integer;
  FormEditor: IOTAFormEditor;
begin
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Exit;

  // 1. Form-Designer benachrichtigen (falls vorhanden)
  //    Der Form-Editor hat die Class-RTTI gecached, die aktualisiert werden muss.
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTAFormEditor, FormEditor) then
    begin
      // MarkModified signalisiert dem Form-Designer, dass sich etwas geaendert hat
      try
        FormEditor.MarkModified;
      except
        // Manche Formeditor-Implementierungen unterstuetzen MarkModified nicht
      end;
    end;

  // 2. Modul neu parsen ohne von Platte neu zu laden (False = Editor-Buffer behalten)
  try
    Module.Refresh(False);
  except
    // Refresh-Fehler tolerieren
  end;
end;

class function TEditorHelper.ReplaceLineAt(const AFilePath: string;
  ALine: Integer; const ANewContent: string): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Writer: IOTAEditWriter;
  I: Integer;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then Exit;

  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then Exit;

  Writer := SourceEditor.CreateUndoableWriter;
  if Writer = nil then Exit;
  try
    Writer.ReplaceLine(ALine, UTF8String(UTF8Encode(ANewContent)));
  finally
    Writer := nil;
  end;
  Result := True;
end;

class function TEditorHelper.DeleteLineAt(const AFilePath: string;
  ALine: Integer): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Writer: IOTAEditWriter;
  I: Integer;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Module := ModuleServices.OpenModule(AFilePath);
  if Module = nil then Exit;

  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then Exit;

  Writer := SourceEditor.CreateUndoableWriter;
  if Writer = nil then Exit;
  try
    Writer.DeleteLine(ALine);
  finally
    Writer := nil;
  end;
  Result := True;
end;

class function TEditorHelper.ReadEditorContent(const AFilePath: string;
  out AContent: string): Boolean;
var
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
  SourceEditor: IOTASourceEditor;
  Reader: IOTAEditReader;
  I, BufSize, BytesRead: Integer;
  Buf: TBytes;
begin
  Result := False;
  AContent := '';
  if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then Exit;
  Module := ModuleServices.FindModule(AFilePath);
  if Module = nil then Exit;

  SourceEditor := nil;
  for I := 0 to Module.GetModuleFileCount - 1 do
    if Supports(Module.GetModuleFileEditor(I), IOTASourceEditor, SourceEditor) then
      Break;
  if SourceEditor = nil then Exit;

  Reader := SourceEditor.CreateReader;
  if Reader = nil then Exit;
  BufSize := SourceEditor.GetLinesInBuffer * 200;
  if BufSize < 65536 then BufSize := 65536;
  if BufSize > 16 * 1024 * 1024 then BufSize := 16 * 1024 * 1024;
  SetLength(Buf, BufSize);
  BytesRead := Reader.GetText(0, @Buf[0], BufSize);
  Reader := nil;

  AContent := TEncoding.UTF8.GetString(Buf, 0, BytesRead);
  Result := True;
end;

class function TEditorHelper.FindDelphiLspJson: string;
var
  DprojPath, ProjDir: string;
  JsonFiles: TStringDynArray;
begin
  Result := '';
  DprojPath := GetCurrentProjectDproj;
  if DprojPath = '' then Exit;

  ProjDir := ExtractFilePath(DprojPath);

  // 1. Gleicher Name wie .dproj aber mit .delphilsp.json
  Result := ChangeFileExt(DprojPath, '.delphilsp.json');
  if FileExists(Result) then Exit;

  // 2. Gleicher Name wie .dpr (Projekte heissen oft ProjektName.dproj
  //    aber die .delphilsp.json heisst ProjektName.delphilsp.json vom .dpr)
  Result := ChangeFileExt(ChangeFileExt(DprojPath, ''), '.delphilsp.json');
  if FileExists(Result) then Exit;

  // 3. Suche im Projektverzeichnis
  JsonFiles := TDirectory.GetFiles(ProjDir, '*.delphilsp.json');
  if Length(JsonFiles) > 0 then
  begin
    Result := JsonFiles[0];
    Exit;
  end;

  Result := '';
end;

class function TEditorHelper.GetProjectSourceFiles: TArray<string>;
var
  ModuleServices: IOTAModuleServices;
  ProjectGroup: IOTAProjectGroup;
  Project: IOTAProject;
  ModInfo: IOTAModuleInfo;
  FileList: TList<string>;
  FileName, Ext: string;
  I: Integer;
begin
  FileList := TList<string>.Create;
  try
    if not Supports(BorlandIDEServices, IOTAModuleServices, ModuleServices) then
      Exit(nil);

    // Aktives Projekt finden
    Project := nil;
    for I := 0 to ModuleServices.ModuleCount - 1 do
    begin
      if Supports(ModuleServices.Modules[I], IOTAProjectGroup, ProjectGroup) then
      begin
        Project := ProjectGroup.ActiveProject;
        Break;
      end;
    end;

    // Fallback: erstes Projekt
    if Project = nil then
    begin
      for I := 0 to ModuleServices.ModuleCount - 1 do
      begin
        if Supports(ModuleServices.Modules[I], IOTAProject, Project) then
          if not Supports(Project, IOTAProjectGroup) then
            Break;
      end;
    end;

    if Project = nil then
      Exit(nil);

    // Alle Module des Projekts durchgehen
    for I := 0 to Project.GetModuleCount - 1 do
    begin
      ModInfo := Project.GetModule(I);
      if ModInfo = nil then Continue;
      FileName := ModInfo.FileName;
      if FileName = '' then Continue;

      Ext := LowerCase(ExtractFileExt(FileName));
      if (Ext = '.pas') or (Ext = '.dpr') or (Ext = '.dpk') then
      begin
        if not FileList.Contains(FileName) then
          FileList.Add(FileName);
      end;
    end;

    // Auch die Hauptdatei des Projekts (.dpr) hinzufuegen
    FileName := Project.FileName;
    if FileName <> '' then
    begin
      // .dproj -> .dpr
      var DprFile := ChangeFileExt(FileName, '.dpr');
      if FileExists(DprFile) and not FileList.Contains(DprFile) then
        FileList.Add(DprFile);
      // Oder .dpk
      DprFile := ChangeFileExt(FileName, '.dpk');
      if FileExists(DprFile) and not FileList.Contains(DprFile) then
        FileList.Add(DprFile);
    end;

    Result := FileList.ToArray;
  finally
    FileList.Free;
  end;
end;

class function TEditorHelper.GetProjectSearchPaths: string;
var
  DprojPath, RootPath: string;
begin
  DprojPath := GetCurrentProjectDproj;
  RootPath := GetProjectRoot;
  Result := TEditorHelper.BuildSearchPathFromProject(DprojPath, RootPath);
end;

class function TEditorHelper.BuildSearchPathFromProject(
  const ADprojPath, ARootPath: string): string;
var
  Dirs: TList<string>;
  DprojDir, Dir, AbsDir: string;
  XmlDoc: IXMLDocument;
  Root, ItemGroup, RefNode: IXMLNode;
  I, J: Integer;
  IncludeVal: string;
begin
  Dirs := TList<string>.Create;
  try
    DprojDir := ExtractFilePath(ADprojPath);

    // 1. Verzeichnis der .dproj selbst
    if DprojDir <> '' then
      Dirs.Add(ExcludeTrailingPathDelimiter(DprojDir));

    // 2. Root-Verzeichnis (falls anders als dproj-Dir)
    if (ARootPath <> '') and
       not SameText(ExcludeTrailingPathDelimiter(ARootPath),
                    ExcludeTrailingPathDelimiter(DprojDir)) then
      Dirs.Add(ExcludeTrailingPathDelimiter(ARootPath));

    // 3. Aus .dproj: DCCReference-Eintraege → deren Verzeichnisse
    if (ADprojPath <> '') and FileExists(ADprojPath) then
    begin
      try
        XmlDoc := TXMLDocument.Create(nil);
        XmlDoc.LoadFromFile(ADprojPath);
        XmlDoc.Active := True;
        Root := XmlDoc.DocumentElement;

        for I := 0 to Root.ChildNodes.Count - 1 do
        begin
          if Root.ChildNodes[I].NodeName = 'ItemGroup' then
          begin
            ItemGroup := Root.ChildNodes[I];
            for J := 0 to ItemGroup.ChildNodes.Count - 1 do
            begin
              RefNode := ItemGroup.ChildNodes[J];
              if (RefNode.NodeName = 'DCCReference') and
                 RefNode.HasAttribute('Include') then
              begin
                IncludeVal := RefNode.Attributes['Include'];
                // Relativen Pfad aufloesen
                AbsDir := ExtractFilePath(
                  ExpandFileName(TPath.Combine(DprojDir, IncludeVal)));
                AbsDir := ExcludeTrailingPathDelimiter(AbsDir);
                if (AbsDir <> '') and not Dirs.Contains(AbsDir) then
                  Dirs.Add(AbsDir);
              end;
            end;
          end;
        end;
      except
        // XML-Parse-Fehler ignorieren
      end;
    end;

    // 4. Alle Unterverzeichnisse von ARootPath die .pas enthalten
    if (ARootPath <> '') and TDirectory.Exists(ARootPath) then
    begin
      for Dir in TDirectory.GetDirectories(ARootPath, '*',
        TSearchOption.soAllDirectories) do
      begin
        AbsDir := ExcludeTrailingPathDelimiter(Dir);
        if not Dirs.Contains(AbsDir) then
        begin
          // Nur hinzufuegen wenn Verzeichnis .pas/.dpr Dateien enthaelt
          if (Length(TDirectory.GetFiles(Dir, '*.pas')) > 0) or
             (Length(TDirectory.GetFiles(Dir, '*.dpr')) > 0) then
            Dirs.Add(AbsDir);
        end;
      end;
    end;

    // Semikolon-getrennt zusammenbauen
    Result := '';
    for Dir in Dirs do
    begin
      if Result <> '' then
        Result := Result + ';';
      Result := Result + Dir;
    end;
  finally
    Dirs.Free;
  end;
end;

end.
