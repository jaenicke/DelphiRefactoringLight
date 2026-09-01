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

// Current text of AFile - live editor buffer first, disk as fallback.
function ReadBuffer(const AFile: string; out AContent: string): Boolean;
begin
  AContent := '';
  Result := False;
  if Editor = nil then Exit;
  if Editor.ReadEditorContent(AFile, AContent) then Exit(True);
  try
    AContent := TFile.ReadAllText(AFile);
    Result := AContent <> '';
  except
    Result := False;
  end;
end;

// Qualifier of a dotted use site: for "TMyClass.MyClassProc" with the
// caret on the member this returns 'TMyClass'. '' when the identifier is
// not dot-qualified. ACol0 is the 0-based column of the identifier.
function QualifierBefore(const ALine: string; ACol0: Integer): string;
var
  I, EndP: Integer;
begin
  Result := '';
  I := ACol0;                      // 1-based index of the char BEFORE it
  while (I >= 1) and (I <= Length(ALine)) and CharInSet(ALine[I], [' ', #9]) do Dec(I);
  if (I < 1) or (I > Length(ALine)) or (ALine[I] <> '.') then Exit;
  Dec(I);
  while (I >= 1) and CharInSet(ALine[I], [' ', #9]) do Dec(I);
  EndP := I;
  while (I >= 1) and CharInSet(ALine[I],
    ['A'..'Z', 'a'..'z', '0'..'9', '_']) do Dec(I);
  if EndP > I then Result := Copy(ALine, I + 1, EndP - I);
end;

// Jumps to a MEMBER declaration of a dot-qualified use site. Class
// members (class procedure/function/property and their instance
// counterparts) are not in the identifier index - the index only holds
// top-level interface declarations - so resolve the QUALIFIER as a type
// and then walk that type's body. This is the case DelphiLSP most often
// cannot answer ("TMyClass.MyClassProc").
function TryQualifiedFallback(const AQualifier, AMember: string): Boolean;
var
  Hits: TArray<TFindUnitHit>;
  Content: string;
  Line, Col: Integer;
  Seen: TArray<string>;

  function AlreadyTried(const APath: string): Boolean;
  begin
    Result := False;
    for var S in Seen do
      if SameText(S, APath) then Exit(True);
  end;

begin
  Result := False;
  if (AQualifier = '') or (AMember = '') then Exit;
  Hits := TUnitIndex.Instance.Lookup(AQualifier);
  for var H in Hits do
  begin
    if (H.Path = '') or AlreadyTried(H.Path) then Continue;
    Seen := Seen + [H.Path];
    if not ReadBuffer(H.Path, Content) then Continue;
    Line := FindMemberDeclarationLine(Content, AQualifier, AMember);
    if Line < 0 then Continue;      // this unit declares the type but not
                                    // the member - try the next candidate
    var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
    Col := 0;
    if Line <= High(Lines) then
    begin
      Col := FindWholeWord(Lines[Line], AMember);
      if Col < 0 then Col := 0;
    end;
    Exit(Editor.GotoLocation(H.Path, Line, Col, Length(AMember)));
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
    //
    // DOT-QUALIFIED first ("TMyClass.MyClassProc"): resolving the member
    // through its type is both the case the LSP fails at most often AND
    // the safer answer - a plain lookup of the member name alone could
    // land on an unrelated global routine of the same name.
    var Content: string;
    if ReadBuffer(Ctx.FileName, Content) then
    begin
      var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
      var L0 := Ctx.Line - 1;
      if (L0 >= 0) and (L0 <= High(Lines)) then
      begin
        var C0 := FindWholeWord(Lines[L0], Ctx.WordAtCursor);
        if C0 >= 0 then
        begin
          var Qual := QualifierBefore(Lines[L0], C0);
          if (Qual <> '') and TryQualifiedFallback(Qual, Ctx.WordAtCursor) then Exit;
        end;
      end;
    end;

    // Unqualified: one declaring unit -> jump straight to its
    // declaration, several -> hand over to the Find-Unit dialog.
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
