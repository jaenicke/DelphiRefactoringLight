(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Based on the idea and original implementation of pull request #9
 * by Dumach (github.com/Dumach) - reimplemented on the current
 * architecture (IEditorHelper abstraction, TLspManager prewarm pattern,
 * shared IDE + standalone menu integration).
 *)
unit Expert.FindOriginalSymbolWizard;

// "Find original symbol" (go to declaration) via the plugin's own LSP
// session - as an alternative to Ctrl+Click, which does not always work
// reliably in Delphi 13.1. Default shortcut: Ctrl+G (configurable).

interface

procedure FindOriginalSymbol;

implementation

uses
  System.SysUtils, System.IOUtils, System.Math,
  Vcl.Forms, Vcl.Controls, Vcl.Dialogs,
  Lsp.Protocol, Lsp.Client, Lsp.Uri, Expert.LspManager,
  Expert.EditorHelperIntf, Expert.DialogHelper,
  Expert.UnitIndex, Expert.FindUnitDialog;

// 0-based column of AWord as a whole word in ALine (case-insensitive),
// or -1. Word boundaries: identifier characters on either side disqualify.
function FindWholeWord(const ALine, AWord: string): Integer;
var
  U, NeedleU: string;
  P, AfterIdx: Integer;
begin
  Result := -1;
  if (ALine = '') or (AWord = '') then Exit;
  U := UpperCase(ALine);
  NeedleU := UpperCase(AWord);
  P := Pos(NeedleU, U);
  while P > 0 do
  begin
    AfterIdx := P + Length(NeedleU);
    if ((P = 1) or not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_']))
      and ((AfterIdx > Length(U)) or not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_'])) then
      Exit(P - 1);
    P := Pos(NeedleU, U, P + 1);
  end;
end;

// Index-based fallback for a symbol the LSP could not resolve. True when
// it handled the request (jumped, or opened the chooser).
function TryIndexFallback(const AIdent: string): Boolean;
var
  Hits: TArray<TFindUnitHit>;
  Seen: TArray<string>;
  Content: string;
  DeclLine: Integer;

  function AlreadySeen(const AUnit: string): Boolean;
  begin
    Result := False;
    for var S in Seen do
      if SameText(S, AUnit) then Exit(True);
  end;

begin
  Result := False;
  Hits := TUnitIndex.Instance.Lookup(AIdent);
  if Length(Hits) = 0 then Exit;

  // One declaring UNIT (the same unit can appear per declaration form) ->
  // go there directly.
  Seen := nil;
  for var H in Hits do
    if not AlreadySeen(H.UnitName) then
      Seen := Seen + [H.UnitName];

  if Length(Seen) > 1 then
  begin
    // Ambiguous - let the user pick in the dialog they know.
    FindUnitForIdentifier;
    Exit(True);
  end;

  if (Hits[0].Path = '') or not TFile.Exists(Hits[0].Path) then Exit;
  if not Editor.ReadEditorContent(Hits[0].Path, Content) then
    try
      Content := TFile.ReadAllText(Hits[0].Path);
    except
      Exit;
    end;

  DeclLine := FindDeclarationLine(Content, AIdent);
  if DeclLine < 0 then DeclLine := 0;
  var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  var Col := 0;
  if DeclLine <= High(Lines) then
  begin
    Col := FindWholeWord(Lines[DeclLine], AIdent);
    if Col < 0 then Col := 0;
  end;
  Result := Editor.GotoLocation(Hits[0].Path, DeclLine, Col, Length(AIdent));
end;

procedure FindOriginalSymbol;
var
  Ctx: TEditorContext;
  Json, Root, Proj, DefPath: string;
  Client: TLspClient;
  Locs: TArray<TLspLocation>;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or (Ctx.WordAtCursor = '') then
  begin
    ShowThemedMessage('Place the caret on an identifier first.');
    Exit;
  end;
  Json := Editor.FindDelphiLspJson;
  if Json = '' then
  begin
    ShowThemedMessage('DelphiLSP is not configured for this project ' +
      '(missing .delphilsp.json next to the project file).');
    Exit;
  end;
  Root := Editor.GetProjectRoot;
  if Root = '' then Root := ExtractFilePath(Ctx.FileName);
  Proj := Editor.GetCurrentProjectDproj;

  Locs := nil;
  Screen.Cursor := crHourGlass;
  try
    try
      // GetClient may cold-start + initialize the LSP and can raise -
      // this runs from a raw key-binding handler, so nothing may escape.
      Client := TLspManager.Instance.GetClient(Root, Proj, Json);
      if Client = nil then Exit;
      // Push the LIVE buffer so the position matches what the user sees.
      Client.RefreshDocument(Ctx.FileName);
      Locs := Client.GotoDefinition(Ctx.FileName, Ctx.Line - 1, Ctx.Column - 1);
    except
      on E: Exception do
      begin
        Screen.Cursor := crDefault;
        ShowThemedMessage('LSP request failed: ' + E.Message);
        Exit;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;

  if Length(Locs) = 0 then
  begin
    // FALLBACK: DelphiLSP frequently answers nothing for a symbol whose
    // unit is not (yet) open - hover and the identifier index do know it.
    // Use the index: one candidate -> jump straight to its declaration,
    // several -> hand over to the Find-Unit dialog (same search the user
    // would run manually).
    if TryIndexFallback(Ctx.WordAtCursor) then Exit;
    ShowThemedMessage(Format('No declaration found for "%s".'#13#10 +
      'The identifier index does not know it either.', [Ctx.WordAtCursor]));
    Exit;
  end;

  DefPath := TLspUri.FileUriToPath(Locs[0].Uri);
  if DefPath = '' then Exit;

  // DelphiLSP's definition range does not reliably point AT the identifier
  // (multi-line headers come back one line off, so a blind highlight from
  // Range.Start selects unrelated text). Locate the identifier around the
  // reported position and highlight the REAL occurrence; when it cannot be
  // found, just place the caret without a selection.
  var DefLine := Locs[0].Range.Start.Line;   // 0-based
  var Content: string;
  if not Editor.ReadEditorContent(DefPath, Content) then
    try
      Content := TFile.ReadAllText(DefPath);
    except
      Content := '';
    end;
  if Content <> '' then
  begin
    var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
    for var Off in [0, -1, 1, -2, 2] do
    begin
      var L := DefLine + Off;
      if (L < 0) or (L > High(Lines)) then Continue;
      var C := FindWholeWord(Lines[L], Ctx.WordAtCursor);
      if C >= 0 then
      begin
        Editor.GotoLocation(DefPath, L, C, Length(Ctx.WordAtCursor));
        Exit;
      end;
    end;
  end;
  Editor.GotoLocation(DefPath, DefLine,
    Max(0, Locs[0].Range.Start.Character), 0);
end;

end.
