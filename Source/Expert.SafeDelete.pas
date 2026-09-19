(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SafeDelete;

// SAFE DELETE (idea: Ian Branch, issue #11) - deletes a declaration only
// after PROVING that nothing uses it:
//
//   1. DelphiLSP names the declaration of the identifier at the caret.
//   2. Expert.SafeDeletePlan decides what would be removed and vetoes
//      shapes that can be used WITHOUT a textual reference (overloads,
//      virtual / override / message methods, published members).
//   3. Every whole-word occurrence in the project scope (comments and
//      strings masked) is asked where it leads. THE POLARITY MATTERS: only
//      an occurrence that DelphiLSP proves to belong to ANOTHER symbol is
//      harmless. One that leads to our declaration is a use, and one that
//      gets NO answer is treated as a use too - that is exactly what a
//      call site in an inactive {$IFDEF} branch looks like, and deleting
//      the declaration would break the other configuration.
//   4. Form files (.dfm / .fmx) bind event handlers and components BY
//      NAME - any mention blocks.
//   5. A method that implements an interface method is kept.
//
// The analysis runs on a worker thread (reads files from disk, open
// buffers are captured on the main thread first); applying it re-checks
// that the declaring unit did not change in between and goes through
// ApplyLinesMinimal (undoable in an open editor).

interface

uses
  System.SysUtils, System.Classes, System.Types, Lsp.Client, Expert.SafeDeletePlan;

type
  TSafeDeleteFindingKind = (sfUse, sfUnverified, sfOther, sfForm, sfNote);

  TSafeDeleteFinding = record
    Kind: TSafeDeleteFindingKind;
    FilePath: string;
    Line: Integer;      // 0-based, -1 = not a position
    Col: Integer;       // 0-based
    Text: string;       // the source line
    Note: string;
  end;

  TSafeDeleteInput = record
    Client: TLspClient;
    FileName: string;
    Content: string;    // buffer of FileName
    Line0, Col0: Integer;
    ScopeFiles: TArray<string>;
    OpenFiles: TArray<string>;      // open editor buffers ...
    OpenContents: TArray<string>;   // ... and their content
    BdsRoot: string;
  end;

  TSafeDeleteResult = record
    Ok: Boolean;
    Error: string;
    Identifier: string;
    Symbol: TSafeDeleteSymbol;
    DeclFile: string;
    DeclContent: string;
    Candidates: Integer;
    FilesScanned: Integer;
    Truncated: Boolean;
    Findings: TArray<TSafeDeleteFinding>;
    function Blocking: Integer;
    function CanDelete: Boolean;
    function Summary: string;
  end;

/// <summary>Runs the whole check (worker thread; waits for DelphiLSP).
///  AStop cancels (an event handle).</summary>
function AnalyzeSafeDelete(const AIn: TSafeDeleteInput; AStop: THandle;
  const AProgress: TProc<Integer, Integer, string>): TSafeDeleteResult;

/// <summary>Main thread: applies the planned edits when the declaring unit
///  still has the analysed content. AError says why not.</summary>
function ApplySafeDelete(const ARes: TSafeDeleteResult; out AError: string): Boolean;

/// <summary>Editor entry point (menu "Safe delete...").</summary>
procedure SafeDeleteAtCursor;

// ---- shared with Expert.ChangeSignature -------------------------------------

/// <summary>The identifier at / directly left of (ALine0, ACol0);
///  AStartCol0 = its first column. '' when there is none.</summary>
function IdentifierAtPos(const ALines: TArray<string>; ALine0, ACol0: Integer;
  out AStartCol0: Integer): string;

/// <summary>Whole-word occurrences of AWord in CODE of ALines[AFirst..ALast]
///  (comments / strings masked): X = column, Y = line, 0-based.</summary>
procedure CodeWordHits(const ALines: TArray<string>; const AWord: string;
  AFirst, ALast: Integer; var AOut: TArray<TPoint>);

/// <summary>Main thread: the buffer at the position, the project scope, the
///  open buffers (captured for a worker) and the DelphiLSP client.</summary>
function GatherScanInput(const AFile: string; ALine0, ACol0: Integer;
  out AIn: TSafeDeleteInput; out AError: string): Boolean;

implementation

uses
  Winapi.Windows, System.Math, System.StrUtils, System.IOUtils, System.JSON, System.SyncObjs,
  System.Generics.Collections, System.Character,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.AutoImport, Expert.UsesEditor,
  Expert.ScopeFiles, Expert.LspManager, Expert.DiagStore, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort, Expert.WorkerLatch, Expert.McpServer,
  Expert.McpLspTools, Lsp.Protocol, Lsp.Uri, Delphi.FileEncoding, Expert.PascalScanner,
  Expert.IncludeExpansion, Expert.InterfaceLinks;

type
  TCand = record
    F: string;
    L, C: Integer;
    T: string;
  end;

const
  MaxCandidates = 400;
  FindingKindText: array[TSafeDeleteFindingKind] of string =
    ('USED', 'NOT VERIFIABLE', 'other symbol', 'FORM FILE', 'note');

function IdentifierAtPos(const ALines: TArray<string>; ALine0, ACol0: Integer;
  out AStartCol0: Integer): string;
var
  S: string;
  P, Q: Integer;
begin
  Result := '';
  AStartCol0 := -1;
  if (ALine0 < 0) or (ALine0 > High(ALines)) then Exit;
  S := ALines[ALine0];
  P := ACol0 + 1;
  if (P > Length(S)) or not IsIdentChar(S[P]) then
    if (P - 1 >= 1) and (P - 1 <= Length(S)) and IsIdentChar(S[P - 1]) then
      Dec(P)
    else
      Exit;
  Q := P;
  while (P > 1) and IsIdentChar(S[P - 1]) do Dec(P);
  while (Q < Length(S)) and IsIdentChar(S[Q + 1]) do Inc(Q);
  Result := Copy(S, P, Q - P + 1);
  if (Result <> '') and Result[1].IsDigit then Exit('');
  AStartCol0 := P - 1;
end;

// Whole-word occurrences of AWord in CODE (comments / strings masked):
// (line, col) pairs, 0-based.
procedure CodeWordHits(const ALines: TArray<string>; const AWord: string;
  AFirst, ALast: Integer; var AOut: TArray<TPoint>);
var
  M: TArray<string>;
  U, W: string;
  P, E: Integer;
begin
  W := UpperCase(AWord);
  M := MaskCommentsAndStrings(ALines);
  for var L := Max(AFirst, 0) to Min(ALast, High(M)) do
  begin
    U := UpperCase(M[L]);
    P := Pos(W, U);
    while P > 0 do
    begin
      E := P + Length(W);
      if ((P = 1) or not IsIdentChar(U[P - 1])) and ((E > Length(U)) or not IsIdentChar(U[E])) then
        AOut := AOut + [Point(P - 1, L)];
      P := Pos(W, U, P + Length(W));
    end;
  end;
end;

{ TSafeDeleteResult }

function TSafeDeleteResult.Blocking: Integer;
begin
  Result := Length(Symbol.Vetoes);
  for var F in Findings do
    if F.Kind in [sfUse, sfUnverified, sfForm] then Inc(Result);
  if Truncated then Inc(Result);
end;

function TSafeDeleteResult.CanDelete: Boolean;
begin
  Result := Ok and (Length(Symbol.Edits) > 0) and (Blocking = 0);
end;

function TSafeDeleteResult.Summary: string;
var
  Uses_, Unver, Forms: Integer;
begin
  if not Ok then Exit('Safe delete is not possible: ' + Error);
  Uses_ := 0; Unver := 0; Forms := 0;
  for var F in Findings do
    case F.Kind of
      sfUse: Inc(Uses_);
      sfUnverified: Inc(Unver);
      sfForm: Inc(Forms);
    end;
  Result := Format('%s %s%s - %d occurrence(s) checked in %d file(s)',
    [SafeDeleteKindText(Symbol.Kind),
     IfThen(Symbol.Container <> '', Symbol.Container + '.', ''), Symbol.Name,
     Candidates, FilesScanned]);
  if CanDelete then
    Result := Result + sLineBreak + 'Nothing uses it - it can be deleted.'
  else
  begin
    Result := Result + sLineBreak + 'NOT deletable:';
    for var V in Symbol.Vetoes do
      Result := Result + sLineBreak + '  - ' + V;
    if Uses_ > 0 then
      Result := Result + sLineBreak + Format('  - used %d time(s)', [Uses_]);
    if Unver > 0 then
      Result := Result + sLineBreak + Format('  - %d occurrence(s) DelphiLSP could ' +
        'not resolve (e.g. in an inactive {$IFDEF} branch) - counted as uses', [Unver]);
    if Forms > 0 then
      Result := Result + sLineBreak + Format('  - mentioned in %d form file line(s)', [Forms]);
    if Truncated then
      Result := Result + sLineBreak + Format('  - more than %d occurrences, the ' +
        'check stopped', [MaxCandidates]);
  end;
end;

// ---------------------------------------------------------------------------
//  Analysis
// ---------------------------------------------------------------------------

type
  TContentSource = class
  private
    FMap: TDictionary<string, string>;
  public
    constructor Create(const AIn: TSafeDeleteInput);
    destructor Destroy; override;
    function Get(const AFile: string; out AContent: string): Boolean;
  end;

constructor TContentSource.Create(const AIn: TSafeDeleteInput);
begin
  inherited Create;
  FMap := TDictionary<string, string>.Create;
  for var I := 0 to Min(High(AIn.OpenFiles), High(AIn.OpenContents)) do
    FMap.AddOrSetValue(UpperCase(ExpandFileName(AIn.OpenFiles[I])), AIn.OpenContents[I]);
  FMap.AddOrSetValue(UpperCase(ExpandFileName(AIn.FileName)), AIn.Content);
end;

destructor TContentSource.Destroy;
begin
  FMap.Free;
  inherited;
end;

function TContentSource.Get(const AFile: string; out AContent: string): Boolean;
var
  K: string;
begin
  K := UpperCase(ExpandFileName(AFile));
  if FMap.TryGetValue(K, AContent) then Exit(True);
  Result := False;
  AContent := '';
  if not FileExists(AFile) then Exit;
  try
    AContent := ReadDelphiFile(AFile);
    FMap.Add(K, AContent);
    Result := True;
  except
    Result := False;
  end;
end;

function AnalyzeSafeDelete(const AIn: TSafeDeleteInput; AStop: THandle;
  const AProgress: TProc<Integer, Integer, string>): TSafeDeleteResult;
var
  Res: TSafeDeleteResult;
  Src: TContentSource;
  StartLines, DeclLines: TArray<string>;
  IdentCol, DeclLine: Integer;
  Why: string;
  Targets: TLspSymbolTargets;
  IncCtx: TLspIncludeContext;

  procedure AddFinding(AKind: TSafeDeleteFindingKind; const AFile: string;
    ALine, ACol: Integer; const AText, ANote: string);
  var
    F: TSafeDeleteFinding;
  begin
    F.Kind := AKind;
    F.FilePath := AFile;
    F.Line := ALine;
    F.Col := ACol;
    F.Text := Trim(AText);
    F.Note := ANote;
    Res.Findings := Res.Findings + [F];
  end;

  function Stopped: Boolean;
  begin
    Result := (AStop <> 0) and (WaitForSingleObject(AStop, 0) = WAIT_OBJECT_0);
  end;

  function InEdits(ALine: Integer): Boolean;
  begin
    for var E in Res.Symbol.Edits do
      if (ALine >= E.FirstLine) and (ALine <= E.LastLine) then Exit(True);
    Result := False;
  end;

begin
  Res := Default(TSafeDeleteResult);
  IncCtx := nil;
  Src := TContentSource.Create(AIn);
  try
    StartLines := SplitContentLines(AIn.Content);
    Res.Identifier := IdentifierAtPos(StartLines, AIn.Line0, AIn.Col0, IdentCol);
    if Res.Identifier = '' then
    begin
      Res.Error := 'there is no identifier at the caret';
      Exit(Res);
    end;
    if AIn.Client = nil then
    begin
      Res.Error := 'the DelphiLSP session is not running yet (it starts with the ' +
        'first opened project)';
      Exit(Res);
    end;
    if Assigned(AProgress) then AProgress(0, 0, 'Asking DelphiLSP for the declaration...');

    // Positions inside {$I} include files are answered through the
    // INCLUDING unit, sent expanded; freeing the context restores it.
    IncCtx := TLspIncludeContext.Create(AIn.Client,
      function(const APath: string; out AContent: string): Boolean
      begin
        Result := Src.Get(APath, AContent);
      end);
    IncCtx.RegisterFiles(AIn.ScopeFiles + [AIn.FileName]);

    // 1. declaration
    if not IncCtx.OwnsDocument(AIn.FileName) then
    begin
      var Before := AIn.Client.GetFileDiagnosticsVersion(AIn.FileName);
      if McpSyncLspContent(AIn.Client, AIn.FileName, AIn.Content) then
        McpWaitLspAnalysed(AIn.Client, AIn.FileName, Before, AStop, 30000);
    end;
    var D := IncCtx.Definition(AIn.FileName, AIn.Line0, IdentCol);
    var Deadline := GetTickCount64 + 3000;
    while (Length(D) = 0) and (GetTickCount64 < Deadline) and not Stopped do
    begin
      Sleep(300);
      D := IncCtx.Definition(AIn.FileName, AIn.Line0, IdentCol);
    end;
    if Length(D) > 0 then
    begin
      Res.DeclFile := ExpandFileName(TLspUri.FileUriToPath(D[0].Uri));
      DeclLine := D[0].Range.Start.Line;
    end
    else if LineDeclaresName(StartLines[AIn.Line0], Res.Identifier) then
    begin
      // DelphiLSP answers nothing AT a declaration - the caret is it
      Res.DeclFile := ExpandFileName(AIn.FileName);
      DeclLine := AIn.Line0;
    end
    else
    begin
      Res.Error := 'DelphiLSP knows no declaration for "' + Res.Identifier + '"';
      Exit(Res);
    end;
    if (AIn.BdsRoot <> '') and Res.DeclFile.ToUpper.StartsWith(
      IncludeTrailingPathDelimiter(AIn.BdsRoot).ToUpper) then
    begin
      Res.Error := Res.Identifier + ' is declared in the RTL / VCL (' +
        ExtractFileName(Res.DeclFile) + ')';
      Exit(Res);
    end;
    if not Src.Get(Res.DeclFile, Res.DeclContent) then
    begin
      Res.Error := 'the declaring file cannot be read: ' + Res.DeclFile;
      Exit(Res);
    end;
    DeclLines := SplitContentLines(Res.DeclContent);

    // 2. what would be deleted
    if not PlanSafeDeleteSymbol(DeclLines, DeclLine, Res.Identifier, Res.Symbol, Why) then
    begin
      Res.Error := Why;
      Exit(Res);
    end;
    if (Res.Symbol.Kind = sdkMethod) and not Res.Symbol.IsInterfaceMember then
    begin
      // the interfaces of the class (and the ones they inherit from) - shared
      // with find references (Expert.InterfaceLinks)
      var Intf := '';
      var Graph := TTypeGraph.Create(AIn.ScopeFiles + [Res.DeclFile],
        function(const APath: string; out AContent: string): Boolean
        begin
          Result := Src.Get(APath, AContent);
        end);
      try
        var IL := Graph.InterfaceMethodsImplementedBy(Res.Symbol.Container, Res.Symbol.Name);
        if Length(IL) > 0 then Intf := IL[0].TypeName + '.' + Res.Symbol.Name;
      finally
        Graph.Free;
      end;
      if Intf <> '' then
        Res.Symbol.Vetoes := Res.Symbol.Vetoes + ['it implements the interface method ' +
          Intf + ' - calls through the interface have no textual reference to it'];
    end;
    Targets.Add(Res.DeclFile, Res.Symbol.DeclLine);
    if Res.Symbol.ImplLine >= 0 then Targets.Add(Res.DeclFile, Res.Symbol.ImplLine);

    // 3. every occurrence in the scope
    var Files: TArray<string>;
    if Res.Symbol.LocalFirst >= 0 then
      Files := [Res.DeclFile]
    else
    begin
      Files := [Res.DeclFile];
      for var F in AIn.ScopeFiles do
        if not SameText(ExpandFileName(F), Res.DeclFile) then Files := Files + [F];
    end;
    var Cands: TArray<TCand> := nil;
    for var F in Files do
    begin
      if Stopped then begin Res.Error := 'cancelled'; Exit(Res); end;
      var C: string;
      if not Src.Get(F, C) then Continue;
      Inc(Res.FilesScanned);
      var Lines := SplitContentLines(C);
      var Hits: TArray<TPoint> := nil;
      var IsDecl := SameText(ExpandFileName(F), Res.DeclFile);
      if Res.Symbol.LocalFirst >= 0 then
        CodeWordHits(Lines, Res.Identifier, Res.Symbol.LocalFirst, Res.Symbol.LocalLast, Hits)
      else
        CodeWordHits(Lines, Res.Identifier, 0, High(Lines), Hits);
      for var H in Hits do
      begin
        if IsDecl and InEdits(H.Y) then Continue;   // deleted with it
        if Length(Cands) >= MaxCandidates then
        begin
          Res.Truncated := True;
          Break;
        end;
        var Cd: TCand;
        Cd.F := F; Cd.L := H.Y; Cd.C := H.X; Cd.T := Lines[H.Y];
        Cands := Cands + [Cd];
      end;
      if Res.Truncated then Break;
    end;
    Res.Candidates := Length(Cands);

    // 4. ask DelphiLSP where each occurrence leads
    var Synced := TDictionary<string, Boolean>.Create;
    try
      for var I := 0 to High(Cands) do
      begin
        if Stopped then begin Res.Error := 'cancelled'; Exit(Res); end;
        var Cd := Cands[I];
        if Assigned(AProgress) then
          AProgress(I + 1, Length(Cands), Format('%s(%d)', [ExtractFileName(Cd.F), Cd.L + 1]));
        var Key := UpperCase(ExpandFileName(Cd.F));
        var FirstInFile := not Synced.ContainsKey(Key);
        if FirstInFile then
        begin
          Synced.Add(Key, True);
          var C: string;
          if not IncCtx.OwnsDocument(Cd.F) and Src.Get(Cd.F, C) then
          begin
            var B := AIn.Client.GetFileDiagnosticsVersion(Cd.F);
            if McpSyncLspContent(AIn.Client, Cd.F, C) then
              McpWaitLspAnalysed(AIn.Client, Cd.F, B, AStop, 30000);
          end;
        end;
        var Answer: TArray<TLspLocation> := nil;
        try
          Answer := IncCtx.Definition(Cd.F, Cd.L, Cd.C);
          // an EMPTY answer is retried briefly (the unit may still be in
          // analysis), a wrong one never
          var Dl := GetTickCount64 + UInt64(IfThen(FirstInFile, 3000, 600));
          while (Length(Answer) = 0) and (GetTickCount64 < Dl) and not Stopped do
          begin
            Sleep(300);
            Answer := IncCtx.Definition(Cd.F, Cd.L, Cd.C);
          end;
        except
          Answer := nil;
        end;
        if (Length(Answer) = 0) and IsIncludeFile(Cd.F) then
          AddFinding(sfUnverified, Cd.F, Cd.L, Cd.C, Cd.T,
            'DelphiLSP gives no answer inside this include file - counted as a use')
        else if Length(Answer) = 0 then
          AddFinding(sfUnverified, Cd.F, Cd.L, Cd.C, Cd.T,
            'DelphiLSP gives no answer here - an inactive {$IFDEF} branch looks like this')
        else
        begin
          var AF := TLspUri.FileUriToPath(Answer[0].Uri);
          var AL := Answer[0].Range.Start.Line;
          if Targets.Contains(AF, AL) then
            AddFinding(sfUse, Cd.F, Cd.L, Cd.C, Cd.T, '')
          else
            AddFinding(sfOther, Cd.F, Cd.L, Cd.C, Cd.T,
              Format('leads to %s(%d)', [ExtractFileName(AF), AL + 1]));
        end;
      end;
    finally
      Synced.Free;
    end;

    // 5. form files bind handlers and components by name
    if (Res.Symbol.Kind in [sdkMethod, sdkField]) and (Res.Symbol.Container <> '') then
      for var F in Files do
        for var Ext in ['.dfm', '.fmx'] do
        begin
          var FF := ChangeFileExt(F, Ext);
          if not FileExists(FF) then Continue;
          var Bytes: TBytes := nil;
          try
            Bytes := TFile.ReadAllBytes(FF);
          except
            Continue;
          end;
          if (Length(Bytes) > 0) and (Bytes[0] = $FF) then
          begin
            AddFinding(sfUnverified, FF, -1, 0, '', 'binary form file - cannot be checked');
            Continue;
          end;
          var Txt := ReadDelphiFile(FF);
          var FL := SplitContentLines(Txt);
          for var L in FormTextMentions(Txt, Res.Identifier) do
            AddFinding(sfForm, FF, L, 0, FL[L], 'the form refers to it by name');
        end;

    if Res.Symbol.LocalFirst >= 0 then
      AddFinding(sfNote, Res.DeclFile, -1, 0, '', 'local symbol - only its routine was checked')
    else
      AddFinding(sfNote, '', -1, 0, '', Format('checked the project scope (%d file(s)); ' +
        'code OUTSIDE it (other projects using this unit) is not seen', [Res.FilesScanned]));
    Res.Ok := True;
    Result := Res;
  finally
    IncCtx.Free;   // sends the original text of expanded units again
    Src.Free;
  end;
end;

function ApplySafeDelete(const ARes: TSafeDeleteResult; out AError: string): Boolean;
var
  Cur: string;
  SL: TStringList;
begin
  Result := False;
  AError := '';
  if not ARes.CanDelete then
  begin
    AError := 'the check did not allow deleting it';
    Exit;
  end;
  if (Editor = nil) or not Editor.ReadEditorContent(ARes.DeclFile, Cur) then
    try
      Cur := ReadDelphiFile(ARes.DeclFile);
    except
      on E: Exception do
      begin
        AError := E.Message;
        Exit;
      end;
    end;
  if DiagContentHash(Cur) <> DiagContentHash(ARes.DeclContent) then
  begin
    AError := ExtractFileName(ARes.DeclFile) + ' changed since the check - run it again';
    Exit;
  end;
  SL := TStringList.Create;
  try
    SL.Text := string.Join(sLineBreak,
      ApplySafeDeleteEdits(SplitContentLines(Cur), ARes.Symbol.Edits));
    Result := ApplyLinesMinimal(ARes.DeclFile, SL, Cur);
    if not Result then AError := 'the edit could not be applied';
  finally
    SL.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Main-thread input
// ---------------------------------------------------------------------------

function GatherScanInput(const AFile: string; ALine0, ACol0: Integer;
  out AIn: TSafeDeleteInput; out AError: string): Boolean;
begin
  Result := False;
  AIn := Default(TSafeDeleteInput);
  AIn.FileName := ExpandFileName(AFile);
  AIn.Line0 := ALine0;
  AIn.Col0 := ACol0;
  if Editor = nil then
  begin
    AError := 'no editor';
    Exit;
  end;
  if not Editor.ReadEditorContent(AIn.FileName, AIn.Content) then
    if FileExists(AIn.FileName) then
      AIn.Content := ReadDelphiFile(AIn.FileName)
    else
    begin
      AError := 'file not found: ' + AFile;
      Exit;
    end;
  AIn.ScopeFiles := ProjectScopeFiles(AIn.FileName);
  for var F in Editor.GetOpenSourceFiles do
  begin
    var C: string;
    if Editor.ReadEditorContent(F, C) then
    begin
      AIn.OpenFiles := AIn.OpenFiles + [F];
      AIn.OpenContents := AIn.OpenContents + [C];
    end;
  end;
  AIn.BdsRoot := FindBdsRoot;
  AIn.Client := TLspManager.Instance.PeekClient;
  Result := True;
end;

// ---------------------------------------------------------------------------
//  Dialog
// ---------------------------------------------------------------------------

type
  TSafeDeleteDialog = class(TForm)
  private
    FRes: TSafeDeleteResult;
    FList: TListView;
    FMemo: TMemo;
    FBtnDelete, FBtnGoto, FBtnClose: TButton;
    procedure DoGoto(Sender: TObject);
    procedure DoDelete(Sender: TObject);
  public
    Deleted: Boolean;
    constructor CreateDialog(AOwner: TComponent; const ARes: TSafeDeleteResult);
  end;

constructor TSafeDeleteDialog.CreateDialog(AOwner: TComponent; const ARes: TSafeDeleteResult);
var
  Col: TListColumn;
  Panel: TPanel;
  Item: TListItem;
begin
  inherited CreateNew(AOwner);
  FRes := ARes;
  Caption := 'Safe delete: ' + ARes.Identifier;
  Width := 820;
  Height := 480;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FMemo := TMemo.Create(Self);
  FMemo.Parent := Self;
  FMemo.Align := alTop;
  FMemo.AlignWithMargins := True;
  FMemo.Height := 110;
  FMemo.ReadOnly := True;
  FMemo.ScrollBars := ssVertical;
  FMemo.Text := ARes.Summary;

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.OnDblClick := DoGoto;
  Col := FList.Columns.Add; Col.Caption := 'Result'; Col.Width := 110;
  Col := FList.Columns.Add; Col.Caption := 'File';   Col.Width := 150;
  Col := FList.Columns.Add; Col.Caption := 'Line';   Col.Width := 50;
    Col.Alignment := taRightJustify;
  Col := FList.Columns.Add; Col.Caption := 'Code';   Col.Width := 300;
  Col := FList.Columns.Add; Col.Caption := 'Note';   Col.Width := 180;

  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Align := alBottom;
  Panel.Height := 40;
  Panel.BevelOuter := bvNone;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Align := alRight;
  FBtnClose.AlignWithMargins := True;
  FBtnClose.Cancel := True;
  FBtnClose.ModalResult := mrCancel;

  FBtnDelete := TButton.Create(Self);
  FBtnDelete.Parent := Panel;
  FBtnDelete.Caption := '&Delete';
  FBtnDelete.Align := alRight;
  FBtnDelete.AlignWithMargins := True;
  FBtnDelete.Enabled := ARes.CanDelete;
  FBtnDelete.OnClick := DoDelete;

  FBtnGoto := TButton.Create(Self);
  FBtnGoto.Parent := Panel;
  FBtnGoto.Caption := '&Go to';
  FBtnGoto.Align := alRight;
  FBtnGoto.AlignWithMargins := True;
  FBtnGoto.OnClick := DoGoto;

  // the declaration first, then the findings
  FList.Items.BeginUpdate;
  try
    for var E in ARes.Symbol.Edits do
    begin
      Item := FList.Items.Add;
      Item.Data := Pointer(NativeInt(-1 - E.FirstLine));
      Item.Caption := IfThen(E.HasReplacement, 'will change', 'will delete');
      Item.SubItems.Add(ExtractFileName(ARes.DeclFile));
      Item.SubItems.Add(IntToStr(E.FirstLine + 1));
      Item.SubItems.Add(Format('%d line(s)', [E.LastLine - E.FirstLine + 1]));
      Item.SubItems.Add(IfThen(E.HasReplacement, '-> ' + Trim(E.Replacement), ''));
    end;
    for var I := 0 to High(ARes.Findings) do
    begin
      var F := ARes.Findings[I];
      Item := FList.Items.Add;
      Item.Data := Pointer(NativeInt(I));
      Item.Caption := FindingKindText[F.Kind];
      Item.SubItems.Add(ExtractFileName(F.FilePath));
      Item.SubItems.Add(IfThen(F.Line >= 0, IntToStr(F.Line + 1), ''));
      Item.SubItems.Add(F.Text);
      Item.SubItems.Add(F.Note);
    end;
  finally
    FList.Items.EndUpdate;
  end;

  EnableListViewSorting(FList);
  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
  if ARes.CanDelete then ActiveControl := FBtnDelete else ActiveControl := FList;
end;

procedure TSafeDeleteDialog.DoGoto(Sender: TObject);
begin
  if FList.Selected = nil then Exit;
  var Idx := NativeInt(FList.Selected.Data);
  if Idx < 0 then
    Editor.GotoLocation(FRes.DeclFile, -1 - Idx, 0)
  else if (FRes.Findings[Idx].Line >= 0) and (FRes.Findings[Idx].FilePath <> '') then
  begin
    if SameText(ExtractFileExt(FRes.Findings[Idx].FilePath), '.pas') then
      Editor.GotoLocation(FRes.Findings[Idx].FilePath, FRes.Findings[Idx].Line,
        FRes.Findings[Idx].Col, Length(FRes.Identifier))
    else
      Exit;
  end
  else
    Exit;
  ModalResult := mrCancel;
end;

procedure TSafeDeleteDialog.DoDelete(Sender: TObject);
var
  Err: string;
begin
  if ApplySafeDelete(FRes, Err) then
  begin
    Deleted := True;
    ModalResult := mrOk;
  end
  else
    ShowThemedMessage('Safe delete failed: ' + Err);
end;

// ---------------------------------------------------------------------------
//  Editor entry point
// ---------------------------------------------------------------------------

type
  TSafeDeleteJob = class(TInterfacedObject)
  public
    Lock: TCriticalSection;
    Cur, Total: Integer;
    Text: string;
    Done: Boolean;
    Res: TSafeDeleteResult;
    Stop: TEvent;
    constructor Create;
    destructor Destroy; override;
  end;

constructor TSafeDeleteJob.Create;
begin
  inherited;
  Lock := TCriticalSection.Create;
  Stop := TEvent.Create(nil, True, False, '');
end;

destructor TSafeDeleteJob.Destroy;
begin
  Stop.Free;
  Lock.Free;
  inherited;
end;

procedure SafeDeleteAtCursor;
var
  Ctx: TEditorContext;
  Inp: TSafeDeleteInput;
  Err: string;
  Job: TSafeDeleteJob;
  JobRef: IInterface;
  Prog: TCheckProgressWindow;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or (Ctx.WordAtCursor = '') then
  begin
    ShowThemedMessage('Safe delete: place the caret on the identifier to delete.');
    Exit;
  end;
  Editor.SaveAllFiles;   // the scan reads closed files from disk
  if not GatherScanInput(Ctx.FileName, Ctx.Line - 1, Ctx.Column - 1, Inp, Err) then
  begin
    ShowThemedMessage('Safe delete: ' + Err);
    Exit;
  end;
  if Inp.Client = nil then
  begin
    ShowThemedMessage('Safe delete needs the DelphiLSP session, which is not running ' +
      'yet - open a project and try again in a moment.');
    Exit;
  end;

  Job := TSafeDeleteJob.Create;
  JobRef := Job;
  Prog := CreateCheckProgress('Safe delete', Application.MainForm,
    'Checking ' + Ctx.WordAtCursor + '...');
  try
    var ThreadRef: IInterface := JobRef;
    var StopHandle := Job.Stop.Handle;
    if not StartWorker(
      procedure
      var
        R: TSafeDeleteResult;
      begin
        try
          R := AnalyzeSafeDelete(Inp, StopHandle,
            procedure(ACur, ATotal: Integer; AText: string)
            begin
              Job.Lock.Enter;
              try
                Job.Cur := ACur; Job.Total := ATotal; Job.Text := AText;
              finally
                Job.Lock.Leave;
              end;
            end);
        except
          on E: Exception do
          begin
            R := Default(TSafeDeleteResult);
            R.Error := E.ClassName + ': ' + E.Message;
          end;
        end;
        Job.Lock.Enter;
        try
          Job.Res := R;
          Job.Done := True;
        finally
          Job.Lock.Leave;
        end;
        ThreadRef := nil;
      end) then
    begin
      ShowThemedMessage('Safe delete: the plugin is shutting down.');
      Exit;
    end;

    // wait, pumping; closing the progress window cancels
    while True do
    begin
      var C, T: Integer;
      var S: string;
      var D: Boolean;
      Job.Lock.Enter;
      try
        C := Job.Cur; T := Job.Total; S := Job.Text; D := Job.Done;
      finally
        Job.Lock.Leave;
      end;
      if D then Break;
      if not Prog.Visible then Job.Stop.SetEvent;
      Prog.Step(C, T, S);
      Sleep(40);
    end;
  finally
    Prog.Free;
  end;

  if not Job.Res.Ok then
  begin
    ShowThemedMessage(Job.Res.Summary);
    Exit;
  end;
  var Dlg := TSafeDeleteDialog.CreateDialog(Application.MainForm, Job.Res);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  MCP tool "safe_delete"
// ---------------------------------------------------------------------------

function FindingsToJson(const ARes: TSafeDeleteResult): TJSONArray;
begin
  Result := TJSONArray.Create;
  for var F in ARes.Findings do
  begin
    var O := TJSONObject.Create;
    O.AddPair('result', FindingKindText[F.Kind]);
    if F.FilePath <> '' then O.AddPair('file', F.FilePath);
    if F.Line >= 0 then
    begin
      O.AddPair('line', TJSONNumber.Create(F.Line + 1));
      O.AddPair('column', TJSONNumber.Create(F.Col + 1));
      O.AddPair('text', F.Text);
    end;
    if F.Note <> '' then O.AddPair('note', F.Note);
    Result.Add(O);
  end;
end;

function ToolSafeDelete(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, GatherErr: string;
  L1, C1: Integer;
  Inp: TSafeDeleteInput;
  Ok: Boolean;
begin
  F := AArgs.GetValue<string>('file', '');
  L1 := AArgs.GetValue<Integer>('line', 0);
  C1 := AArgs.GetValue<Integer>('column', 0);
  var DoApply := AArgs.GetValue<Boolean>('apply', False);
  if (F = '') or (L1 < 1) or (C1 < 1) then
    Exit(McpErr('arguments "file", "line" and "column" (1-based) are required'));
  F := ExpandFileName(F);
  Ok := False;
  if not McpRunOnMain(
    procedure
    var
      E: string;
    begin
      Ok := GatherScanInput(F, L1 - 1, C1 - 1, Inp, E);
      if not Ok then GatherErr := E;
    end, True, AStop, Err) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(GatherErr));
  var Res := AnalyzeSafeDelete(Inp, AStop, nil);
  var J := TJSONObject.Create;
  J.AddPair('identifier', Res.Identifier);
  J.AddPair('summary', Res.Summary);
  J.AddPair('deletable', TJSONBool.Create(Res.CanDelete));
  if Res.Ok then
  begin
    J.AddPair('kind', SafeDeleteKindText(Res.Symbol.Kind));
    if Res.Symbol.Container <> '' then J.AddPair('container', Res.Symbol.Container);
    J.AddPair('declaration', Format('%s:%d', [Res.DeclFile, Res.Symbol.DeclLine + 1]));
    if Res.Symbol.ImplLine >= 0 then
      J.AddPair('implementation', Format('%s:%d', [Res.DeclFile, Res.Symbol.ImplLine + 1]));
    var Ed := TJSONArray.Create;
    for var E in Res.Symbol.Edits do
    begin
      var O := TJSONObject.Create;
      O.AddPair('from_line', TJSONNumber.Create(E.FirstLine + 1));
      O.AddPair('to_line', TJSONNumber.Create(E.LastLine + 1));
      if E.HasReplacement then O.AddPair('replacement', E.Replacement);
      Ed.Add(O);
    end;
    J.AddPair('edits', Ed);
    var V := TJSONArray.Create;
    for var S in Res.Symbol.Vetoes do V.Add(S);
    J.AddPair('vetoes', V);
    J.AddPair('occurrences_checked', TJSONNumber.Create(Res.Candidates));
    J.AddPair('findings', FindingsToJson(Res));
  end;
  if DoApply then
  begin
    if not Res.CanDelete then
      J.AddPair('applied', TJSONBool.Create(False))
    else
    begin
      var Applied := False;
      var AErr := '';
      if not McpRunOnMain(
        procedure
        begin
          Applied := ApplySafeDelete(Res, AErr);
        end, False, AStop, Err) then
        AErr := Err;
      J.AddPair('applied', TJSONBool.Create(Applied));
      if AErr <> '' then J.AddPair('apply_error', AErr);
    end;
  end;
  Result := McpOk(J);
end;

initialization
  RegisterMcpTool('safe_delete', ToolSafeDelete);

end.
