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

uses
  Lsp.Client;

type
  /// <summary>One place the declaration of an identifier could be.
  ///  Line/Col are 0-based. Via says which step answered.</summary>
  TOriginalSymbolHit = record
    FilePath: string;
    Line: Integer;
    Col: Integer;
    Via: string;        // 'DelphiLSP' / 'type member' / 'identifier index'
    TypeName: string;   // set when the member was resolved through a type
  end;

/// <summary>Resolves the declaration of AIDENT at (AFile, ALine0, ACol0) -
///  DelphiLSP first, then the qualifier's type, then the identifier index -
///  WITHOUT opening anything. More than one hit means the index knows
///  several declaring units (the menu entry hands those to the Find-Unit
///  dialog). ANOTE carries the reason when nothing was found. Split out of
///  FindOriginalSymbol so the MCP tool answers from the same chain the menu
///  jumps with (user request 2026-09-29). MAIN THREAD (reads buffers).</summary>
function ResolveOriginalSymbol(AClient: TLspClient; const AFile: string;
  ALine0, ACol0: Integer; const AIdent: string;
  out AHits: TArray<TOriginalSymbolHit>; out ANote: string): Boolean;

procedure FindOriginalSymbol;

implementation

uses
  System.SysUtils, System.IOUtils, System.Math,
  Vcl.Forms, Vcl.Controls, Vcl.Dialogs,
  Lsp.Protocol, Lsp.Uri, Expert.LspManager,   // Lsp.Client: interface uses
  Expert.EditorHelperIntf, Expert.DialogHelper, Expert.PascalScanner,
  Expert.UnitIndex, Expert.InterfaceLinks, Expert.ScopeFiles,
  Expert.FindReferencesDialog, Delphi.FileEncoding;

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

// Jumps to a MEMBER declaration of a dot-qualified use site. Class
// members (class procedure/function/property and their instance
// counterparts) are not in the identifier index - the index only holds
// top-level interface declarations - so resolve the QUALIFIER as a type
// and then walk that type's body. This is the case DelphiLSP most often
// cannot answer ("TMyClass.MyClassProc").
// Resolves the member through the QUALIFIER's type: the units that declare
// the qualifier are asked for the member. Answers a location, never jumps -
// FindOriginalSymbol does that, and the MCP tool reports it instead.
function AlreadySeen2(const AHits: TArray<TOriginalSymbolHit>;
  const APath: string): Boolean;
begin
  Result := False;
  for var H in AHits do
    if SameText(H.FilePath, APath) then Exit(True);
end;

function ResolveQualifiedMember(const AQualifier, AMember: string;
  out AHit: TOriginalSymbolHit): Boolean;
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
  AHit := Default(TOriginalSymbolHit);
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
    AHit.FilePath := H.Path;
    AHit.Line := Line;
    AHit.Col := Col;
    AHit.Via := 'type member';
    AHit.TypeName := AQualifier;
    Exit(True);
  end;
end;

// Index-based fallback for a symbol the LSP could not resolve.
// Every failing branch leaves a reason in ANote. Without it, an index that
// KNOWS the identifier and a file it cannot read produce the same "the index
// does not know it either" - which sends the next report in the wrong
// direction (that is how the TcxButton case cost a round trip).
var
  GIndexFallbackNote: string = '';

function ResolveViaIndex(const AIdent: string;
  out AHits: TArray<TOriginalSymbolHit>; out ANote: string): Boolean;
var
  Hits: TArray<TFindUnitHit>;
  Units: TArray<string>;
  Content: string;
  DeclLine: Integer;

  function AlreadySeen(const AUnit: string): Boolean;
  begin
    Result := False;
    for var S in Units do
      if SameText(S, AUnit) then Exit(True);
  end;

begin
  Result := False;
  AHits := nil;
  ANote := '';
  if not TUnitIndex.Instance.Ready then
  begin
    ANote := 'the identifier index is not ready yet ('
      + TUnitIndex.Instance.StatusLine + ')';
    Exit;
  end;
  Hits := TUnitIndex.Instance.Lookup(AIdent);
  if Length(Hits) = 0 then
  begin
    ANote := 'no indexed unit declares it';
    Exit;
  end;

  // One declaring UNIT (the same unit can appear per declaration form) ->
  // that is the answer; several -> ambiguous, every candidate is reported.
  Units := nil;
  for var H in Hits do
    if not AlreadySeen(H.UnitName) then
      Units := Units + [H.UnitName];

  if Length(Units) > 1 then
  begin
    for var H in Hits do
      if not AlreadySeen2(AHits, H.Path) then
      begin
        var Amb := Default(TOriginalSymbolHit);
        Amb.FilePath := H.Path;
        Amb.Line := 0;
        Amb.Via := 'identifier index (several declaring units)';
        AHits := AHits + [Amb];
      end;
    Exit(Length(AHits) > 1);
  end;

  if Hits[0].Path = '' then
  begin
    ANote := Format('unit %s is indexed, but without a file path',
      [Hits[0].UnitName]);
    Exit;
  end;
  if not TFile.Exists(Hits[0].Path) then
  begin
    ANote := Format('the indexed file no longer exists: %s', [Hits[0].Path]);
    Exit;
  end;
  if not ReadBuffer(Hits[0].Path, Content) then
  begin
    ANote := Format('%s cannot be read', [Hits[0].Path]);
    Exit;
  end;

  DeclLine := FindDeclarationLine(Content, AIdent);
  if DeclLine < 0 then
  begin
    ANote := Format(
      'the declaration of "%s" was not recognised in %s',
      [AIdent, ExtractFileName(Hits[0].Path)]);
    DeclLine := 0;
  end;
  var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  var Col := 0;
  if DeclLine <= High(Lines) then
  begin
    Col := FindWholeWord(Lines[DeclLine], AIdent);
    if Col < 0 then Col := 0;
  end;
  var Hit := Default(TOriginalSymbolHit);
  Hit.FilePath := Hits[0].Path;
  Hit.Line := DeclLine;
  Hit.Col := Col;
  Hit.Via := 'identifier index';
  AHits := [Hit];
  Result := True;
end;

function ResolveOriginalSymbol(AClient: TLspClient; const AFile: string;
  ALine0, ACol0: Integer; const AIdent: string;
  out AHits: TArray<TOriginalSymbolHit>; out ANote: string): Boolean;
var
  Locs: TArray<TLspLocation>;
  Content: string;
begin
  Result := False;
  AHits := nil;
  ANote := '';
  if (AFile = '') or (AIdent = '') then
  begin
    ANote := 'no identifier at that position';
    Exit;
  end;
  Locs := nil;
  if AClient <> nil then
    try
      Locs := AClient.GotoDefinition(AFile, ALine0, ACol0);
    except
      on E: Exception do
        ANote := 'the DelphiLSP request failed: ' + E.Message;
    end;
  if Length(Locs) > 0 then
  begin
    var DefPath := TLspUri.FileUriToPath(Locs[0].Uri);
    if DefPath <> '' then
    begin
      var Hit := Default(TOriginalSymbolHit);
      Hit.FilePath := DefPath;
      Hit.Line := Locs[0].Range.Start.Line;
      Hit.Col := Max(0, Locs[0].Range.Start.Character);
      Hit.Via := 'DelphiLSP';
      // Its range does not reliably point AT the identifier (multi-line
      // headers come back one line off), so locate the name around it.
      if ReadBuffer(DefPath, Content) then
      begin
        var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
        for var Off in [0, -1, 1, -2, 2] do
        begin
          var L := Hit.Line + Off;
          if (L < 0) or (L > High(Lines)) then Continue;
          var C := FindWholeWord(Lines[L], AIdent);
          if C >= 0 then
          begin
            Hit.Line := L;
            Hit.Col := C;
            Break;
          end;
        end;
      end;
      AHits := [Hit];
      Exit(True);
    end;
  end;

  // FALLBACK: DelphiLSP frequently answers nothing for a symbol whose unit
  // is not (yet) open - hover and the identifier index do know it.
  //
  // DOT-QUALIFIED first ("TMyClass.MyClassProc"): resolving the member
  // through its type is both the case the LSP fails at most often AND the
  // safer answer - a plain lookup of the member name alone could land on an
  // unrelated global routine of the same name.
  var MemberNote := '';
  if ReadBuffer(AFile, Content) then
  begin
    var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
    if (ALine0 >= 0) and (ALine0 <= High(Lines)) then
    begin
      var C0 := FindWholeWord(Lines[ALine0], AIdent);
      if C0 >= 0 then
      begin
        var Qual := QualifierBefore(Lines[ALine0], C0);
        var QHit: TOriginalSymbolHit;
        if (Qual <> '') and ResolveQualifiedMember(Qual, AIdent, QHit) then
        begin
          AHits := [QHit];
          Exit(True);
        end;
        // The qualifier may be a VARIABLE, not a type ("lMyClassA.Init"):
        // resolve its declared type and look the member up there, ancestors
        // included. Since Delphi 13.1 this is the ONLY way for a
        // private/public overload pair - DelphiLSP answers nothing at all
        // for those (RSS-5463), not even hover or completion.
        if Qual <> '' then
        begin
          var Graph := TTypeGraph.Create([AFile], EditorOrDiskReader());
          try
            var Link: TMemberLink;
            if ResolveMemberUse(Graph, AFile, Content, ALine0, C0, AIdent,
              Link) <> murNone then
            begin
              var MHit := Default(TOriginalSymbolHit);
              MHit.FilePath := Link.FilePath;
              MHit.Line := Link.Line;
              MHit.Col := Link.Col;
              MHit.Via := 'type member';
              MHit.TypeName := Link.TypeName;
              AHits := [MHit];
              Exit(True);
            end;
          finally
            Graph.Free;
          end;
        end;
      end;
    end;
  end;

  Result := ResolveViaIndex(AIdent, AHits, ANote);
  if not Result and (ANote = '') and (MemberNote <> '') then ANote := MemberNote;
end;

// The ambiguous case: a read-only list of the declarations we resolved,
// each at its own line. Double-click / Enter / "Go to" navigate; nothing
// in this window changes code (audit #40, M20).
procedure ShowOriginalSymbolHits(const AIdent: string;
  const AHits: TArray<TOriginalSymbolHit>);
var
  Dlg: TFindReferencesDialog;
  Items: TFindReferenceItems;
  It: TFindReferenceItem;
  Lines: TArray<string>;
begin
  Items := nil;
  for var H in AHits do
  begin
    It := Default(TFindReferenceItem);
    It.FilePath := H.FilePath;
    It.Line := H.Line;
    It.Col := H.Col;
    It.Length := Length(AIdent);
    It.Kind := 'Declaration';
    It.Relation := H.Via;
    if H.TypeName <> '' then It.Relation := H.Via + ' (' + H.TypeName + ')';
    It.Preview := '';
    try
      Lines := ReadDelphiFileLines(H.FilePath);
      if (H.Line >= 0) and (H.Line <= High(Lines)) then
        It.Preview := Trim(Lines[H.Line]);
    except
      // the preview is a comfort, not the answer
    end;
    Items := Items + [It];
  end;
  Dlg := TFindReferencesDialog.CreateDialog(Application.MainForm, AIdent,
    'Declarations');
  PrepareDialog(Dlg, Application.MainForm);
  Dlg.SetItems(Items);
  Dlg.SetStatus(Format('%d unit(s) declare "%s" - double-click to go there. ' +
    'Nothing here changes code.', [Length(Items), AIdent]));
  Dlg.OnGotoLocation :=
    procedure(AItem: TFindReferenceItem)
    begin
      Editor.GotoLocation(AItem.FilePath, AItem.Line, AItem.Col, AItem.Length);
    end;
  Dlg.SetClosable;
  Dlg.Show;
end;

procedure FindOriginalSymbol;
var
  Ctx: TEditorContext;
  Json, Root, Proj: string;
  Client: TLspClient;
  Hits: TArray<TOriginalSymbolHit>;
  Note: string;
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
    ShowThemedMessage(LspConfigMissingHintLong);
    Exit;
  end;
  Root := Editor.GetProjectRoot;
  if Root = '' then Root := ExtractFilePath(Ctx.FileName);
  Proj := Editor.GetCurrentProjectDproj;

  Screen.Cursor := crHourGlass;
  try
    try
      // GetClient may cold-start + initialize the LSP and can raise - this
      // runs from a raw key-binding handler, so nothing may escape.
      Client := TLspManager.Instance.GetClient(Root, Proj, Json);
      if Client <> nil then
        // Push the LIVE buffer so the position matches what the user sees.
        Client.RefreshDocument(Ctx.FileName);
    except
      on E: Exception do
      begin
        Screen.Cursor := crDefault;
        ShowThemedMessage('LSP request failed: ' + E.Message);
        Exit;
      end;
    end;
    ResolveOriginalSymbol(Client, Ctx.FileName, Ctx.Line - 1, Ctx.Column - 1,
      Ctx.WordAtCursor, Hits, Note);
  finally
    Screen.Cursor := crDefault;
  end;

  if Length(Hits) > 1 then
  begin
    // Several declaring units. This used to drop the resolved hits and
    // open the Find-Unit dialog instead - whose DEFAULT button adds the
    // unit to the interface uses, so pressing Enter to "go there"
    // EDITED the uses clause, and its "Go to" opened the file at line 0
    // (audit #40, M20). The hits are already resolved to a line, so they
    // are shown in the navigation-only result window.
    ShowOriginalSymbolHits(Ctx.WordAtCursor, Hits);
    Exit;
  end;
  if Length(Hits) = 1 then
  begin
    if Editor.GotoLocation(Hits[0].FilePath, Hits[0].Line, Hits[0].Col,
      Length(Ctx.WordAtCursor)) then Exit;
    Note := Format('%s could not be opened in the editor',
      [ExtractFileName(Hits[0].FilePath)]);
  end;
  GIndexFallbackNote := Note;
  if Note <> '' then
    ShowThemedMessage(Format('No declaration found for "%s".'#13#10#13#10 +
      'Index fallback: %s.', [Ctx.WordAtCursor, Note]))
  else
    ShowThemedMessage(Format('No declaration found for "%s".'#13#10 +
      'The identifier index does not know it either.', [Ctx.WordAtCursor]));
end;

end.
