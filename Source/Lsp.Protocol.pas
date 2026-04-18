(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Lsp.Protocol;

interface

uses
  System.SysUtils, System.JSON, System.Generics.Collections;

type
  TLspPosition = record
    Line: Integer;      // 0-based
    Character: Integer; // 0-based
    function ToJSON: TJSONObject;
    class function FromJSON(AObj: TJSONObject): TLspPosition; static;
  end;

  TLspRange = record
    Start: TLspPosition;
    End_: TLspPosition;
    function ToJSON: TJSONObject;
    class function FromJSON(AObj: TJSONObject): TLspRange; static;
  end;

  TLspTextEdit = record
    Range: TLspRange;
    NewText: string;
    class function FromJSON(AObj: TJSONObject): TLspTextEdit; static;
  end;

  TLspTextDocumentIdentifier = record
    Uri: string;
    function ToJSON: TJSONObject;
  end;

  TLspTextDocumentItem = record
    Uri: string;
    LanguageId: string;
    Version: Integer;
    Text: string;
    function ToJSON: TJSONObject;
  end;

  TLspTextDocumentPositionParams = record
    TextDocument: TLspTextDocumentIdentifier;
    Position: TLspPosition;
    function ToJSON: TJSONObject;
  end;

  /// <summary>Edits for a single file.</summary>
  TLspFileEdits = record
    FilePath: string;
    Edits: TArray<TLspTextEdit>;
  end;

  /// <summary>A complete WorkspaceEdit spanning multiple files.</summary>
  TLspWorkspaceEdit = record
    FileEdits: TArray<TLspFileEdits>;
    class function FromJSON(AObj: TJSONObject): TLspWorkspaceEdit; static;
  end;

  /// <summary>Result of textDocument/prepareRename.</summary>
  TLspPrepareRenameResult = record
    Range: TLspRange;
    Placeholder: string;
    IsValid: Boolean;
    class function FromJSON(AValue: TJSONValue): TLspPrepareRenameResult; static;
  end;

  /// <summary>A single source location returned e.g. by
  ///  textDocument/references or textDocument/implementation.</summary>
  TLspLocation = record
    Uri: string;
    Range: TLspRange;
    class function FromJSON(AObj: TJSONObject): TLspLocation; static;
  end;

  /// <summary>Factory helpers for building the initialize-request payload.
  ///  Grouped into a class to keep the unit's public API free of global
  ///  routines.</summary>
  TLspProtocol = class
  public
    /// <summary>Builds the client capabilities object for the
    ///  initialize request.</summary>
    class function BuildClientCapabilities: TJSONObject; static;

    /// <summary>Builds the initializationOptions payload. The actual
    ///  per-project configuration is pushed later via
    ///  workspace/didChangeConfiguration.</summary>
    class function BuildInitializationOptions(const ADprojPath: string; const ASearchPath: string = ''): TJSONObject; static;
  end;

implementation

uses
  Lsp.Uri;

{ TLspPosition }

function TLspPosition.ToJSON: TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('line', TJSONNumber.Create(Line));
  Result.AddPair('character', TJSONNumber.Create(Character));
end;

class function TLspPosition.FromJSON(AObj: TJSONObject): TLspPosition;
begin
  Result.Line := AObj.GetValue<Integer>('line', 0);
  Result.Character := AObj.GetValue<Integer>('character', 0);
end;

{ TLspRange }

function TLspRange.ToJSON: TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('start', Start.ToJSON);
  Result.AddPair('end', End_.ToJSON);
end;

class function TLspRange.FromJSON(AObj: TJSONObject): TLspRange;
begin
  Result.Start := TLspPosition.FromJSON(AObj.GetValue<TJSONObject>('start'));
  Result.End_ := TLspPosition.FromJSON(AObj.GetValue<TJSONObject>('end'));
end;

{ TLspLocation }

class function TLspLocation.FromJSON(AObj: TJSONObject): TLspLocation;
begin
  Result.Uri := AObj.GetValue<string>('uri', '');
  Result.Range := TLspRange.FromJSON(AObj.GetValue<TJSONObject>('range'));
end;

{ TLspTextEdit }

class function TLspTextEdit.FromJSON(AObj: TJSONObject): TLspTextEdit;
begin
  Result.Range := TLspRange.FromJSON(AObj.GetValue<TJSONObject>('range'));
  Result.NewText := AObj.GetValue<string>('newText', '');
end;

{ TLspTextDocumentIdentifier }

function TLspTextDocumentIdentifier.ToJSON: TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('uri', Uri);
end;

{ TLspTextDocumentItem }

function TLspTextDocumentItem.ToJSON: TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('uri', Uri);
  Result.AddPair('languageId', LanguageId);
  Result.AddPair('version', TJSONNumber.Create(Version));
  Result.AddPair('text', Text);
end;

{ TLspTextDocumentPositionParams }

function TLspTextDocumentPositionParams.ToJSON: TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('textDocument', TextDocument.ToJSON);
  Result.AddPair('position', Position.ToJSON);
end;

{ TLspWorkspaceEdit }

class function TLspWorkspaceEdit.FromJSON(AObj: TJSONObject): TLspWorkspaceEdit;
var
  ChangesObj: TJSONObject;
  DocChangesArr: TJSONArray;
  FileList: TList<TLspFileEdits>;
  Pair: TJSONPair;
  EditArr: TJSONArray;
  FileEdit: TLspFileEdits;
  I: Integer;
begin
  FileList := TList<TLspFileEdits>.Create;
  try
    // Format 1: "changes" (URI -> TextEdit[])
    if AObj.TryGetValue<TJSONObject>('changes', ChangesObj) then
    begin
      for Pair in ChangesObj do
      begin
        FileEdit.FilePath := TLspUri.FileUriToPath(Pair.JsonString.Value);
        EditArr := Pair.JsonValue as TJSONArray;
        SetLength(FileEdit.Edits, EditArr.Count);
        for I := 0 to EditArr.Count - 1 do
          FileEdit.Edits[I] := TLspTextEdit.FromJSON(EditArr.Items[I] as TJSONObject);
        FileList.Add(FileEdit);
      end;
    end;

    // Format 2: "documentChanges" (TextDocumentEdit[])
    if AObj.TryGetValue<TJSONArray>('documentChanges', DocChangesArr) then
    begin
      for I := 0 to DocChangesArr.Count - 1 do
      begin
        var DocEdit := DocChangesArr.Items[I] as TJSONObject;
        var TextDocObj := DocEdit.GetValue<TJSONObject>('textDocument');
        FileEdit.FilePath := TLspUri.FileUriToPath(TextDocObj.GetValue<string>('uri'));
        EditArr := DocEdit.GetValue<TJSONArray>('edits');
        SetLength(FileEdit.Edits, EditArr.Count);
        for var J := 0 to EditArr.Count - 1 do
          FileEdit.Edits[J] := TLspTextEdit.FromJSON(EditArr.Items[J] as TJSONObject);
        FileList.Add(FileEdit);
      end;
    end;

    Result.FileEdits := FileList.ToArray;
  finally
    FileList.Free;
  end;
end;

{ TLspPrepareRenameResult }

class function TLspPrepareRenameResult.FromJSON(AValue: TJSONValue): TLspPrepareRenameResult;
var
  Obj: TJSONObject;
begin
  Result.IsValid := False;
  if AValue = nil then
    Exit;
  if not (AValue is TJSONObject) then
    Exit;

  Obj := TJSONObject(AValue);
  Result.IsValid := True;

  // May be {range, placeholder} or just {start, end}.
  if Obj.GetValue('range') <> nil then
  begin
    Result.Range := TLspRange.FromJSON(Obj.GetValue<TJSONObject>('range'));
    Result.Placeholder := Obj.GetValue<string>('placeholder', '');
  end
  else if Obj.GetValue('start') <> nil then
  begin
    Result.Range := TLspRange.FromJSON(Obj);
    Result.Placeholder := '';
  end
  else
    Result.IsValid := False;
end;

{ TLspProtocol }

class function TLspProtocol.BuildClientCapabilities: TJSONObject;
var
  TextDoc, Rename, Sync, Workspace, WsEdit: TJSONObject;
begin
  Rename := TJSONObject.Create;
  Rename.AddPair('prepareSupport', TJSONBool.Create(True));

  Sync := TJSONObject.Create;
  Sync.AddPair('didSave', TJSONBool.Create(True));
  Sync.AddPair('willSave', TJSONBool.Create(False));
  Sync.AddPair('willSaveWaitUntil', TJSONBool.Create(False));

  TextDoc := TJSONObject.Create;
  TextDoc.AddPair('rename', Rename);
  TextDoc.AddPair('synchronization', Sync);

  WsEdit := TJSONObject.Create;
  WsEdit.AddPair('documentChanges', TJSONBool.Create(True));

  Workspace := TJSONObject.Create;
  Workspace.AddPair('workspaceEdit', WsEdit);

  Result := TJSONObject.Create;
  Result.AddPair('textDocument', TextDoc);
  Result.AddPair('workspace', Workspace);
end;

class function TLspProtocol.BuildInitializationOptions(const ADprojPath: string; const ASearchPath: string = ''): TJSONObject;
begin
  Result := TJSONObject.Create;
  // No serverType/agentCount -> simple single-process mode.
  // The real configuration is pushed via workspace/didChangeConfiguration.
end;

end.
