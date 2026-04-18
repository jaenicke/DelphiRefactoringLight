(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.CompletionWizard;

interface

uses
  System.SysUtils, System.Types, System.Classes, System.JSON, Winapi.Windows, Vcl.Forms, Vcl.Controls, ToolsAPI,
  Expert.EditorHelper, Expert.LspManager, Expert.CompletionPopup, Lsp.Client;

type
  TLspCompletionWizard = class
  private
    FPopup: TCompletionPopup;
    procedure DoInsert(const AText: string);
    function GetCaretScreenPos: TPoint;
    /// <summary>Walks left from the current editor cursor collecting
    ///  word characters until a non-word character (or column 1) is
    ///  reached. Returns the collected characters as a prefix string.
    ///  Empty if the cursor is not immediately after a word character.</summary>
    function GetCurrentWordPrefix: string;
  public
    procedure Execute;
  end;

var
  CompletionWizardInstance: TLspCompletionWizard;

implementation

uses
  Delphi.FileEncoding;

function TLspCompletionWizard.GetCurrentWordPrefix: string;
var
  EditorServices: IOTAEditorServices;
  EditBuffer: IOTAEditBuffer;
  EditPosition: IOTAEditPosition;
  Buf: TStringBuilder;
begin
  Result := '';
  if not Supports(BorlandIDEServices, IOTAEditorServices, EditorServices) then Exit;
  EditBuffer := EditorServices.TopBuffer;
  if EditBuffer = nil then Exit;
  EditPosition := EditBuffer.EditPosition;
  if EditPosition = nil then Exit;

  Buf := TStringBuilder.Create;
  try
    EditPosition.Save;
    try
      // Walk left one column at a time. IOTAEditPosition.Character
      // returns the character AT the current column (i.e. just to the
      // right of the conceptual cursor). After MoveRelative(0, -1) the
      // cursor sits one column earlier, so Character is the character
      // immediately before the ORIGINAL cursor position on the first
      // iteration, two characters before on the second, and so on.
      while EditPosition.Column > 1 do
      begin
        EditPosition.MoveRelative(0, -1);
        if EditPosition.IsWordCharacter then
          Buf.Insert(0, EditPosition.Character)
        else
          Break;
      end;
    finally
      EditPosition.Restore;
    end;
    Result := Buf.ToString;
  finally
    Buf.Free;
  end;
end;

function TLspCompletionWizard.GetCaretScreenPos: TPoint;
var
  FocusHwnd: HWND;
  CaretPos: TPoint;
begin
  Result := Point(300, 300);

  FocusHwnd := GetFocus;
  if FocusHwnd <> 0 then
  begin
    if GetCaretPos(CaretPos) then
    begin
      ClientToScreen(FocusHwnd, CaretPos);
      Result := CaretPos;
      Inc(Result.Y, 20);
      Exit;
    end;
  end;

  GetCursorPos(Result);
  Inc(Result.Y, 20);
end;

procedure TLspCompletionWizard.Execute;
var
  Context: TEditorContext;
  DelphiLspJson: string;
  RootPath: string;
  Client: TLspClient;
  CompResponse: TJSONObject;
  Items: TArray<TCompletionItem>;
  PopupPos: TPoint;
  Prefix: string;
  QueryCol: Integer;
begin
  Context := TEditorHelper.GetCurrentContext;
  if not Context.IsValid then
    Exit;

  DelphiLspJson := TEditorHelper.FindDelphiLspJson;
  if DelphiLspJson = '' then
    Exit;

  RootPath := Context.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(Context.FileName);

  // If the cursor sits inside a partially typed identifier (e.g. "Sho"
  // in "foo.Sho|"), extract that prefix and query LSP at the START of
  // the word. That way LSP returns the full list (e.g. all members of
  // foo when preceded by a dot), and the popup filters locally by
  // prefix - instant response, and the user can backspace without
  // losing context.
  Prefix := GetCurrentWordPrefix;
  QueryCol := Context.Column - 1 - Length(Prefix); // 0-based

  PopupPos := GetCaretScreenPos;

  FreeAndNil(FPopup);

  FPopup := TCompletionPopup.CreatePopup(Application.MainForm);
  FPopup.OnInsert := DoInsert;
  FPopup.ShowLoading(PopupPos.X, PopupPos.Y);

  try
    Client := TLspManager.Instance.GetClient(
      RootPath, Context.ProjectFile, DelphiLspJson);
    Client.RefreshDocument(Context.FileName);

    CompResponse := Client.GetCompletion(
      Context.FileName, Context.Line - 1, QueryCol);
    try
      Items := TCompletionItems.Parse(CompResponse);
    finally
      CompResponse.Free;
    end;

    if Length(Items) = 0 then
    begin
      FPopup.ShowMessage('No suggestions.');
      Exit;
    end;

    FPopup.ShowItems(Items, Prefix);
  except
    on E: Exception do
    begin
      if FPopup <> nil then
        FPopup.ShowMessage('Error: ' + E.Message);
    end;
  end;
end;

procedure TLspCompletionWizard.DoInsert(const AText: string);
var
  EditorServices: IOTAEditorServices;
  EditBuffer: IOTAEditBuffer;
  EditPosition: IOTAEditPosition;
  StartCol: Integer;
  EndCol: Integer;
begin
  if not Supports(BorlandIDEServices, IOTAEditorServices, EditorServices) then
    Exit;

  EditBuffer := EditorServices.TopBuffer;
  if EditBuffer = nil then
    Exit;

  EditPosition := EditBuffer.EditPosition;
  if EditPosition = nil then
    Exit;

  EditPosition.Save;
  try
    while EditPosition.IsWordCharacter and (EditPosition.Column > 1) do
      EditPosition.MoveRelative(0, -1);
    if not EditPosition.IsWordCharacter then
      EditPosition.MoveRelative(0, 1);

    StartCol := EditPosition.Column;
    EditPosition.Restore;
    EndCol := EditPosition.Column;
    EditPosition.Move(EditPosition.Row, StartCol);
    EditPosition.Delete(EndCol - StartCol);
  except
    EditPosition.Restore;
  end;

  EditPosition.InsertText(AText);
end;

end.
