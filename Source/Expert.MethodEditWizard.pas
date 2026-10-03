(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MethodEditWizard;

// EDIT METHODS (user request, 2026-10-03) - the IDE half of
// Expert.MethodEdit: the analysis around the caret, the dialog with its two
// tabs, the apply, and the MCP tool "edit_method".
//
// The division of labour is the whole point of this feature, and the user
// stated it: the tool does the MECHANICAL work - the declaration leaves the
// old class and enters the chosen section of another one (in this unit or in
// another), the body travels with it and is requalified, modifiers are
// edited on every selected member at once, a signature is rewritten on both
// of its headers - and it NAMES everything it does not do. What a moved body
// does with the fields it no longer has, and what the call sites need, stays
// with the user. That is why every call site is listed instead of rewritten:
// a text scan cannot tell a call from a method reference, and the verified
// path for a signature change is "Change signature...", which is offered
// right there.

interface

uses
  System.SysUtils, System.Classes, System.Types, Vcl.Forms,
  Expert.MethodEdit, Expert.SignatureEdit, Expert.SafeDelete;

type
  /// <summary>One occurrence of a member's name in the project scope. These
  ///  are TEXT hits, classified by Expert.ReferenceKind but NOT verified by
  ///  DelphiLSP - a same-named member of another class matches too. The
  ///  dialog says so; they exist to be looked at, not to be rewritten.</summary>
  TMethodEditCall = record
    FilePath: string;
    Line, Col: Integer;      // 0-based
    Text: string;            // the source line
    Member: string;
    Kind: string;            // 'call', 'method reference', 'form file', ...
  end;

  /// <summary>How often a member's name occurs in the scope, and how many of
  ///  those occurrences the list really holds.</summary>
  TMethodEditOcc = record
    Member: string;
    Total: Integer;
    Listed: Integer;
  end;

  TMethodEditAnalysis = record
    Ok: Boolean;
    Error: string;
    SourceFile: string;
    SourceContent: string;   // the state everything below was derived from
    OwnerType: string;
    CaretMember: string;     // the member at the caret, '' when none
    /// <summary>What the dialog ticks when it opens: the members the
    ///  SELECTION covers, or the caret's member when nothing is selected.
    ///  A selection of three methods must arrive as three ticks - that is
    ///  what the user marked.</summary>
    Preselected: TArray<string>;
    Members: TArray<TClassMemberInfo>;
    UnitFiles: TArray<string>;   // candidate target units, source unit first
    Calls: TArray<TMethodEditCall>;
    Occurrences: TArray<TMethodEditOcc>;
    Notes: TArray<string>;
    Truncated: Boolean;
    FilesScanned: Integer;
    function Summary: string;
    function CallsOf(const AMembers: TArray<string>): TArray<TMethodEditCall>;
    function OccurrenceOf(const AName: string): TMethodEditOcc;
    function MovableNames: TArray<string>;
    /// <summary>The parameter list and the result type of a member, read
    ///  from the declaration the analysis saw.</summary>
    function ParamsOf(const AName: string): TArray<TSigParam>;
    function ResultTypeOf(const AName: string): string;
  private
    function DeclHeader(const AName: string): string;
  public
  end;

/// <summary>Worker thread: the class at the caret, its members, the target
///  candidates and every text occurrence of the members in the project
///  scope. AStop (an event handle) cancels.
///  ASelFrom / ASelTo (0-based, inclusive) are the lines the editor
///  SELECTION covers; -1 means nothing is selected and the caret
///  decides.</summary>
function AnalyzeMethodEdit(const AIn: TSafeDeleteInput; AStop: THandle;
  const AProgress: TProc<Integer, Integer, string>;
  ASelFrom: Integer = -1; ASelTo: Integer = -1): TMethodEditAnalysis;

/// <summary>Main thread: re-reads both units, refuses when the source
///  changed since the analysis, re-plans on the fresh text and writes.</summary>
function ApplyMethodEdit(const A: TMethodEditAnalysis;
  const AReq: TMethodEditRequest; out AError: string): Boolean;

/// <summary>Main thread: the content of a unit - its editor buffer when it
///  is open, the file on disk otherwise.</summary>
function ReadUnitForEdit(const AFile: string): string;

/// <summary>Editor entry point (menu "Edit methods...").</summary>
procedure EditMethodsAtCursor;

/// <summary>Test seam: the dialog itself. Exported so its layout can be
///  RENDERED headless (scratchpad methodedit\RenderEdit.dpr) instead of
///  guessed - the same reason CreateCircularRefsDialog is exported. Nothing
///  of the production path calls it.</summary>
function CreateMethodEditDialog(AOwner: TComponent;
  const AAn: TMethodEditAnalysis): TForm;

implementation

uses
  Winapi.Windows, System.Math, System.StrUtils, System.IOUtils, System.JSON,
  System.SyncObjs, System.Generics.Collections, System.Generics.Defaults,
  System.UITypes, Vcl.Controls, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.Grids,
  Vcl.Graphics, Vcl.Dialogs,
  Expert.EditorHelperIntf, Expert.PascalScanner, Expert.UnitIndex,
  Expert.AutoImport, Expert.ImplementationFinder, Expert.ReferenceKind,
  Expert.SafeDeletePlan, Expert.DiagStore, Expert.DialogHelper, Expert.IdeThemes,
  Expert.WorkerLatch, Expert.ListViewSort, Expert.McpServer, Expert.McpTools,
  Expert.UsesEditor, Delphi.FileEncoding;

const
  MaxCalls = 800;
  // PER MEMBER, and that is the point: a run-wide cap alone lets ONE common
  // name eat it. Measured on this repository - a class with a 'Create' got
  // 767 of its 800 rows from that constructor, so the occurrences of every
  // other ticked member were truncated away before they were collected.
  MaxCallsPerMember = 60;
  // What the TOOL may put in one answer. The dialog shows a list someone
  // scrolls; an answer is read by a model, and the first live call returned
  // 194 KB for a class with a 'Create'. The counts say what there is, the
  // rows are for the members the caller actually named.
  MaxToolOccurrences = 40;
  SectionNames: array[0..3] of string = ('private', 'protected', 'public',
    'published');
  // The directives a declaration can carry that this dialog offers. Each one
  // is a tri-state: leave alone / add to all / remove from all.
  ModifierNames: array[0..5] of string = ('virtual', 'override', 'overload',
    'inline', 'static', 'reintroduce');

{ TMethodEditAnalysis }

function TMethodEditAnalysis.Summary: string;
var
  Movable: Integer;
begin
  if not Ok then Exit('Edit methods is not possible: ' + Error);
  Movable := 0;
  for var M in Members do
    if M.Movable then Inc(Movable);
  var Total := 0;
  for var O in Occurrences do Total := Total + O.Total;
  Result := Format('%s: %d member(s), %d of them can be moved on their own.',
    [OwnerType, Length(Members), Movable]) + sLineBreak +
    // The number FOUND, not the number kept - the budget is per member, so
    // a common name (a constructor) is listed in part while the others are
    // complete, and saying only the kept count hides both facts.
    Format('%d occurrence(s) of their names in %d file(s)%s - text matches, ' +
    'not verified by DelphiLSP.', [Total, FilesScanned,
    IfThen(Length(Calls) < Total, Format(', %d listed', [Length(Calls)]), '')]);
  for var N in Notes do
    Result := Result + sLineBreak + N;
end;

function TMethodEditAnalysis.CallsOf(const AMembers: TArray<string>): TArray<TMethodEditCall>;
begin
  Result := nil;
  for var C in Calls do
    for var N in AMembers do
      if SameText(C.Member, N) then
      begin
        Result := Result + [C];
        Break;
      end;
end;

function TMethodEditAnalysis.OccurrenceOf(const AName: string): TMethodEditOcc;
begin
  Result := Default(TMethodEditOcc);
  Result.Member := AName;
  for var O in Occurrences do
    if SameText(O.Member, AName) then Exit(O);
end;

function TMethodEditAnalysis.MovableNames: TArray<string>;
begin
  Result := nil;
  for var M in Members do
    if M.Movable then Result := Result + [M.Name];
end;

// The declaration's own text decides - never a second parse of the file.
function TMethodEditAnalysis.DeclHeader(const AName: string): string;
var
  Lines: TArray<string>;
  HdrEnd: Integer;
begin
  Result := '';
  Lines := SplitContentLines(SourceContent);
  for var M in Members do
    if SameText(M.Name, AName) and (M.DeclLine >= 0) and
       (M.DeclLine <= High(Lines)) then
      Exit(CollectHeader(Lines, M.DeclLine, HdrEnd));
end;

function TMethodEditAnalysis.ParamsOf(const AName: string): TArray<TSigParam>;
var
  Kind, Qual, Params, Ret: string;
  IsCM: Boolean;
begin
  Result := nil;
  var Hdr := DeclHeader(AName);
  if Hdr = '' then Exit;
  if not IsHeaderLine(Trim(Hdr), Kind, IsCM) then Exit;
  if not ParseHeader(Hdr, Kind, Qual, Params, Ret) then Exit;
  Result := ParseParamList(Params);
end;

function TMethodEditAnalysis.ResultTypeOf(const AName: string): string;
var
  Kind, Qual, Params, Ret: string;
  IsCM: Boolean;
begin
  Result := '';
  var Hdr := DeclHeader(AName);
  if Hdr = '' then Exit;
  if not IsHeaderLine(Trim(Hdr), Kind, IsCM) then Exit;
  if not ParseHeader(Hdr, Kind, Qual, Params, Ret) then Exit;
  Result := Ret;
end;

// ---------------------------------------------------------------------------
//  Analysis
// ---------------------------------------------------------------------------

type
  // The content of every file the scan touches: the editor buffers the main
  // thread handed over, everything else from disk, each file read once.
  TEditContentSource = class
  private
    FMap: TDictionary<string, string>;
  public
    constructor Create(const AIn: TSafeDeleteInput);
    destructor Destroy; override;
    function Get(const AFile: string; out AContent: string): Boolean;
  end;

constructor TEditContentSource.Create(const AIn: TSafeDeleteInput);
begin
  inherited Create;
  FMap := TDictionary<string, string>.Create;
  for var I := 0 to High(AIn.OpenFiles) do
    if I <= High(AIn.OpenContents) then
      FMap.AddOrSetValue(UpperCase(ExpandFileName(AIn.OpenFiles[I])),
        AIn.OpenContents[I]);
  if AIn.FileName <> '' then
    FMap.AddOrSetValue(UpperCase(ExpandFileName(AIn.FileName)), AIn.Content);
end;

destructor TEditContentSource.Destroy;
begin
  FMap.Free;
  inherited;
end;

function TEditContentSource.Get(const AFile: string; out AContent: string): Boolean;
var
  K: string;
begin
  K := UpperCase(ExpandFileName(AFile));
  if FMap.TryGetValue(K, AContent) then Exit(AContent <> '');
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

// Every whole-word occurrence of ANAMES in AMASKED - one pass over the file
// for all of them, which is what makes a class with thirty methods
// affordable. AOut[i] belongs to ANames[AWhich[i]].
procedure CollectNameHits(const AMasked: TArray<string>;
  const ANames: TArray<string>; out AOut: TArray<TPoint>;
  out AWhich: TArray<Integer>);
var
  Hits: TList<TPoint>;
  Which: TList<Integer>;
  L, P, N: Integer;
  Line, Up: string;
  Ups: TArray<string>;
begin
  AOut := nil;
  AWhich := nil;
  if Length(ANames) = 0 then Exit;
  SetLength(Ups, Length(ANames));
  for N := 0 to High(ANames) do Ups[N] := UpperCase(ANames[N]);
  Hits := TList<TPoint>.Create;
  Which := TList<Integer>.Create;
  try
    for L := 0 to High(AMasked) do
    begin
      Line := AMasked[L];
      if Line = '' then Continue;
      Up := UpperCase(Line);
      for N := 0 to High(Ups) do
      begin
        P := 1;
        while True do
        begin
          P := PosEx(Ups[N], Up, P);
          if P = 0 then Break;
          if ((P = 1) or not IsIdentChar(Line[P - 1])) and
             ((P + Length(Ups[N]) > Length(Line)) or
              not IsIdentChar(Line[P + Length(Ups[N])])) then
          begin
            Hits.Add(Point(P - 1, L));
            Which.Add(N);
          end;
          Inc(P, Length(Ups[N]));
        end;
      end;
    end;
    AOut := Hits.ToArray;
    AWhich := Which.ToArray;
  finally
    Which.Free;
    Hits.Free;
  end;
end;

function AnalyzeMethodEdit(const AIn: TSafeDeleteInput; AStop: THandle;
  const AProgress: TProc<Integer, Integer, string>;
  ASelFrom: Integer = -1; ASelTo: Integer = -1): TMethodEditAnalysis;
var
  Res: TMethodEditAnalysis;
  Src: TEditContentSource;
  Lines, Masked: TArray<string>;
  Names: TArray<string>;
  Kinds: TArray<TRefSymbolKind>;
  Found, Kept: TArray<Integer>;
  Calls: TList<TMethodEditCall>;
  Content: string;

  function Cancelled: Boolean;
  begin
    Result := (AStop <> 0) and (WaitForSingleObject(AStop, 0) = WAIT_OBJECT_0);
  end;

  // The sibling form file of a unit, '' when there is none.
  function FormFileOf(const AUnit: string): string;
  begin
    Result := ChangeFileExt(AUnit, '.dfm');
    if FileExists(Result) then Exit;
    Result := ChangeFileExt(AUnit, '.fmx');
    if FileExists(Result) then Exit;
    Result := '';
  end;

begin
  Res := Default(TMethodEditAnalysis);
  Res.SourceFile := AIn.FileName;
  Res.SourceContent := AIn.Content;
  Lines := SplitContentLines(AIn.Content);

  Res.OwnerType := TImplementationFinder.FindContainingTypeInLines(Lines, AIn.Line0);
  if Res.OwnerType = '' then
  begin
    Res.Error := 'the caret is not inside a class, record or interface - put ' +
      'it on a member of the class whose methods you want to edit';
    Exit(Res);
  end;
  Res.Members := ClassMembersOf(Lines, Res.OwnerType);
  if Length(Res.Members) = 0 then
  begin
    Res.Error := Format('%s declares no member this dialog can edit ' +
      '(its declaration is in another unit, or it has no methods)', [Res.OwnerType]);
    Exit(Res);
  end;

  // What the caret sits on gets preselected - usually exactly what the user
  // came for.
  var Col: Integer;
  var Ident := IdentifierAtPos(Lines, AIn.Line0, AIn.Col0, Col);
  for var M in Res.Members do
    if SameText(M.Name, Ident) then Res.CaretMember := M.Name;
  if Res.CaretMember = '' then
  begin
    // The caret may sit inside a BODY - then its header names the member.
    var HF, HL, HdrEnd: Integer;
    if FindEnclosingRoutineRangeIn(Lines, AIn.Line0, HF, HL) and
       SameText(TImplementationFinder.OwnerTypeFromImplLine(Lines[HF]),
         Res.OwnerType) then
    begin
      var Kind, Qual, Params, Ret: string;
      var IsCM: Boolean;
      var Hdr := CollectHeader(Lines, HF, HdrEnd);
      if (Hdr <> '') and IsHeaderLine(Trim(Hdr), Kind, IsCM) and
         ParseHeader(Hdr, Kind, Qual, Params, Ret) then
      begin
        var D := LastDelimiter('.', Qual);
        if D > 0 then Qual := Copy(Qual, D + 1, MaxInt);
        for var M in Res.Members do
          if SameText(M.Name, Qual) then Res.CaretMember := M.Name;
      end;
    end;
  end;

  // A SELECTION wins over the caret: marking three methods and opening the
  // dialog must tick those three (reported by the user - only the caret's
  // member was ticked). The caret is the fallback for "nothing selected",
  // and it stays the answer when the selection covers no member of this
  // class at all, so a stray selection cannot leave the dialog empty.
  if ASelFrom >= 0 then
    Res.Preselected := MembersInLineRange(Lines, Res.Members, Res.OwnerType,
      ASelFrom, ASelTo);
  if (Length(Res.Preselected) = 0) and (Res.CaretMember <> '') then
    Res.Preselected := [Res.CaretMember];

  // The target candidates: this unit first (a class in the same unit is the
  // most ordinary target of all), then the project scope.
  Res.UnitFiles := [AIn.FileName];
  for var F in AIn.ScopeFiles do
    if SameText(ExtractFileExt(F), '.pas') and
       not SameText(ExpandFileName(F), ExpandFileName(AIn.FileName)) then
      Res.UnitFiles := Res.UnitFiles + [F];

  // A member the FORM DESIGNER binds is a case the planner cannot see: the
  // .dfm names the handler and the designer resolves it against THIS class.
  var FormFile := FormFileOf(AIn.FileName);
  if FormFile <> '' then
  begin
    var FormText := '';
    try
      FormText := ReadDelphiFile(FormFile);
    except
      FormText := '';
    end;
    if FormText <> '' then
    begin
      var Bound := '';
      for var M in Res.Members do
        if (M.Kind <> 'field') and HasWholeWordCI(FormText, M.Name) then
          Bound := Bound + IfThen(Bound <> '', ', ', '') + M.Name;
      if Bound <> '' then
        Res.Notes := Res.Notes + [Format('%s names %s - the designer binds ' +
          'those to %s, so moving one leaves the form without its handler.',
          [ExtractFileName(FormFile), Bound, Res.OwnerType])];
    end;
  end;

  // ---- the occurrences --------------------------------------------------
  Names := nil;
  Kinds := nil;
  for var M in Res.Members do
    if M.Kind <> 'field' then
    begin
      Names := Names + [M.Name];
      if (M.DeclLine >= 0) and (M.DeclLine <= High(Lines)) then
        Kinds := Kinds + [SymbolKindFromDeclLine(Lines[M.DeclLine], M.Name)]
      else
        Kinds := Kinds + [rsUnknown];
    end;

  Calls := TList<TMethodEditCall>.Create;
  SetLength(Found, Length(Names));
  SetLength(Kept, Length(Names));
  Src := TEditContentSource.Create(AIn);
  try
    var Files := Res.UnitFiles;
    for var I := 0 to High(Files) do
    begin
      if Cancelled then Break;
      if Assigned(AProgress) then
        AProgress(I + 1, Length(Files), ExtractFileName(Files[I]));
      if not Src.Get(Files[I], Content) then Continue;
      Inc(Res.FilesScanned);
      var FLines := SplitContentLines(Content);
      Masked := MaskCommentsAndStrings(FLines);
      var Hits: TArray<TPoint>;
      var Which: TArray<Integer>;
      CollectNameHits(Masked, Names, Hits, Which);
      if Length(Hits) = 0 then Continue;
      // One classification pass per NAME, so the kinds really belong to the
      // symbol whose declaration decided them.
      for var N := 0 to High(Names) do
      begin
        var Pos0: TArray<TPoint> := nil;
        for var H := 0 to High(Hits) do
          if Which[H] = N then Pos0 := Pos0 + [Hits[H]];
        if Length(Pos0) = 0 then Continue;
        var RK := ClassifyReferences(Content, Pos0, Length(Names[N]), Kinds[N]);
        Inc(Found[N], Length(Pos0));
        for var H := 0 to High(Pos0) do
        begin
          if (Calls.Count >= MaxCalls) or (Kept[N] >= MaxCallsPerMember) then
          begin
            Res.Truncated := True;
            Break;
          end;
          var C := Default(TMethodEditCall);
          C.FilePath := Files[I];
          C.Line := Pos0[H].Y;
          C.Col := Pos0[H].X;
          C.Member := Names[N];
          if C.Line <= High(FLines) then C.Text := Trim(FLines[C.Line]);
          if H <= High(RK) then C.Kind := RefKindText(RK[H]) else C.Kind := 'use';
          Calls.Add(C);
          Inc(Kept[N]);
        end;
      end;
    end;
    Res.Calls := Calls.ToArray;
    for var N := 0 to High(Names) do
    begin
      var O := Default(TMethodEditOcc);
      O.Member := Names[N];
      O.Total := Found[N];
      O.Listed := Kept[N];
      Res.Occurrences := Res.Occurrences + [O];
    end;
  finally
    Src.Free;
    Calls.Free;
  end;

  Res.Ok := True;
  Result := Res;
end;

// ---------------------------------------------------------------------------
//  Plan and apply
// ---------------------------------------------------------------------------

function ReadUnitForEdit(const AFile: string): string;
begin
  Result := '';
  if AFile = '' then Exit;
  if (Editor <> nil) and Editor.ReadEditorContent(AFile, Result) and
     (Result <> '') then Exit;
  Result := '';
  if FileExists(AFile) then
    try
      Result := ReadDelphiFile(AFile);
    except
      Result := '';
    end;
end;

// The lines of a planned file in the shape ApplyLinesMinimal wants. The rule
// itself is JoinPlannedLines (pure, tested): assigning .Text drops the
// trailing empty element SplitContentLines keeps for the file's final break,
// while adding the elements one by one grew both units by a blank line on
// every apply. Same shape as LinesResult / ContentResult in
// Expert.McpMoreTools.
procedure FillPlannedLines(ASL: TStringList; const ALines: TArray<string>;
  const AOldContent: string);
begin
  ASL.Text := JoinPlannedLines(ALines, AOldContent);
end;

function ApplyMethodEdit(const A: TMethodEditAnalysis;
  const AReq: TMethodEditRequest; out AError: string): Boolean;
var
  Cur, TgtContent: string;
  Res: TMethodEditResult;
  SL: TStringList;
  Fresh: TMethodEditAnalysis;
begin
  AError := '';
  Cur := ReadUnitForEdit(A.SourceFile);
  if Cur = '' then
  begin
    AError := ExtractFileName(A.SourceFile) + ' could not be read.';
    Exit(False);
  end;
  if DiagContentHash(Cur) <> DiagContentHash(A.SourceContent) then
  begin
    AError := ExtractFileName(A.SourceFile) +
      ' changed since the analysis - run the command again.';
    Exit(False);
  end;
  TgtContent := '';
  if (AReq.TargetClass <> '') and (AReq.TargetFile <> '') and
     not SameText(ExpandFileName(AReq.TargetFile), ExpandFileName(A.SourceFile)) then
  begin
    TgtContent := ReadUnitForEdit(AReq.TargetFile);
    if TgtContent = '' then
    begin
      AError := ExtractFileName(AReq.TargetFile) + ' could not be read.';
      Exit(False);
    end;
  end;

  // Everything is planned on the text that is really there now.
  Fresh := A;
  Fresh.SourceContent := Cur;
  Res := PlanMethodEdit(Fresh.SourceFile, Fresh.SourceContent, Fresh.OwnerType,
    AReq, TgtContent);
  if not Res.Ok then
  begin
    AError := Res.Error;
    Exit(False);
  end;

  SL := TStringList.Create;
  try
    // THE TARGET FIRST: if that write fails, the source still declares the
    // member and the unit compiles. The other order would delete it and
    // leave it nowhere.
    if (Res.TargetFile <> '') and (Length(Res.TargetLines) > 0) then
    begin
      FillPlannedLines(SL, Res.TargetLines, TgtContent);
      if not ApplyLinesMinimal(Res.TargetFile, SL, TgtContent) then
      begin
        AError := 'nothing was written: ' + ExtractFileName(Res.TargetFile) +
          ' could not be changed.';
        Exit(False);
      end;
    end;
    FillPlannedLines(SL, Res.SourceLines, Cur);
    if not ApplyLinesMinimal(A.SourceFile, SL, Cur) then
    begin
      AError := Format('%s was changed, but %s could not be written - the ' +
        'member now exists twice. Undo in both editors and try again.',
        [IfThen(Res.TargetFile <> '', ExtractFileName(Res.TargetFile), '(nothing)'),
         ExtractFileName(A.SourceFile)]);
      Exit(False);
    end;
  finally
    SL.Free;
  end;
  Result := True;
end;

// ---------------------------------------------------------------------------
//  The dialog
// ---------------------------------------------------------------------------

type
  TMethodEditDialog = class(TForm)
  private
    FAn: TMethodEditAnalysis;
    FReq: TMethodEditRequest;
    FTargetContent: string;
    FRes: TMethodEditResult;
    FMembers: TListView;
    FTabs: TPageControl;
    FTabTarget, FTabSig: TTabSheet;
    FUnitCombo, FClassCombo, FSectionCombo: TComboBox;
    FBtnBrowse: TButton;
    FKeepHere: TCheckBox;
    FGrid: TStringGrid;
    FSigFor: TLabel;
    FSigAdd, FSigRemove, FSigUp, FSigDown: TButton;
    FModBoxes: array[0..High(ModifierNames)] of TCheckBox;
    FWarn: TMemo;
    FCalls: TListView;
    FCallsLabel: TLabel;
    FBtnApply, FBtnClose: TButton;
    FRows: TArray<TSigParam>;
    FSigResult: string;
    FSigMember: string;
    FFilling: Boolean;
    FTimer: TTimer;
    procedure FillMembers;
    procedure FillUnits;
    procedure FillClasses;
    procedure FillGrid;
    procedure ReadGrid;
    procedure FillCalls;
    function SelectedMembers: TArray<string>;
    procedure BuildRequest;
    procedure Replan;
    procedure Changed(Sender: TObject);
    procedure DoTimer(Sender: TObject);
    procedure SyncSignatureTab;
    procedure DoMemberChange(Sender: TObject; AItem: TListItem;
      AChange: TItemChange);
    procedure DoUnitChange(Sender: TObject);
    procedure DoBrowse(Sender: TObject);
    procedure DoSigAdd(Sender: TObject);
    procedure DoSigRemove(Sender: TObject);
    procedure DoSigMove(Sender: TObject);
    procedure DoCallDblClick(Sender: TObject);
    procedure DoSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
    procedure DoApply(Sender: TObject);
  public
    Applied: Boolean;
    constructor CreateDialog(AOwner: TComponent; const AAn: TMethodEditAnalysis);
  end;

const
  ColMod = 0;
  ColName = 1;
  ColType = 2;
  ColDefault = 3;

constructor TMethodEditDialog.CreateDialog(AOwner: TComponent;
  const AAn: TMethodEditAnalysis);

  function Btn(AParent: TWinControl; const ACaption: string; AAlign: TAlign;
    AOnClick: TNotifyEvent): TButton;
  begin
    Result := TButton.Create(Self);
    Result.Parent := AParent;
    Result.Caption := ACaption;
    Result.Align := AAlign;
    Result.AlignWithMargins := True;
    Result.OnClick := AOnClick;
  end;

  function Lbl(AParent: TWinControl; const ACaption: string;
    ALeft, ATop: Integer): TLabel;
  begin
    Result := TLabel.Create(Self);
    Result.Parent := AParent;
    Result.Caption := ACaption;
    Result.Left := ALeft;
    Result.Top := ATop + 3;
  end;

  function Combo(AParent: TWinControl; ALeft, ATop, AWidth: Integer): TComboBox;
  begin
    Result := TComboBox.Create(Self);
    Result.Parent := AParent;
    Result.Left := ALeft;
    Result.Top := ATop;
    Result.Width := AWidth;
    Result.Style := csDropDownList;
  end;

var
  Left, Right, Bottom, SigTop, SigSide, Mods: TPanel;
  Col: TListColumn;
begin
  inherited CreateNew(AOwner);
  FAn := AAn;
  Caption := 'Edit methods of ' + AAn.OwnerType;
  Width := 1080;
  Height := 720;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  Constraints.MinWidth := 820;
  Constraints.MinHeight := 560;

  // ---- the member list, outside the tabs: it is what both tabs act on ----
  Bottom := TPanel.Create(Self);
  Bottom.Parent := Self;
  Bottom.Align := alBottom;
  Bottom.Height := 40;
  Bottom.BevelOuter := bvNone;
  FBtnClose := Btn(Bottom, '&Close', alRight, nil);
  FBtnClose.Cancel := True;
  FBtnClose.ModalResult := mrCancel;
  FBtnApply := Btn(Bottom, 'A&pply', alRight, DoApply);

  Left := TPanel.Create(Self);
  Left.Parent := Self;
  Left.Align := alLeft;
  Left.Width := 360;
  Left.BevelOuter := bvNone;

  var Split := TSplitter.Create(Self);
  Split.Parent := Self;
  Split.Align := alLeft;
  Split.Width := 5;
  Split.MinSize := 200;

  Right := TPanel.Create(Self);
  Right.Parent := Self;
  Right.Align := alClient;
  Right.BevelOuter := bvNone;

  var Cap := TLabel.Create(Self);
  Cap.Parent := Left;
  Cap.Align := alTop;
  Cap.AlignWithMargins := True;
  Cap.Caption := 'Members of ' + AAn.OwnerType + ' (tick what you edit)';

  FMembers := TListView.Create(Self);
  FMembers.Parent := Left;
  FMembers.Align := alClient;
  FMembers.AlignWithMargins := True;
  FMembers.ViewStyle := vsReport;
  FMembers.ReadOnly := True;
  FMembers.RowSelect := True;
  FMembers.Checkboxes := True;
  FMembers.HideSelection := False;
  Col := FMembers.Columns.Add; Col.Caption := 'Member'; Col.Width := 130;
  Col := FMembers.Columns.Add; Col.Caption := 'Kind'; Col.Width := 80;
  // the reason a member cannot travel alone is the point of this column, so
  // it gets what is left - and the full text goes into the box on the right
  Col := FMembers.Columns.Add; Col.Caption := 'Cannot be moved because';
  Col.Width := 420;   // longer than the panel on purpose: the splitter below
                      // lets the reason be read without a second window
  FMembers.OnChange := DoMemberChange;

  // ---- the two tabs ------------------------------------------------------
  FTabs := TPageControl.Create(Self);
  FTabs.Parent := Right;
  FTabs.Align := alTop;
  FTabs.Height := 230;
  FTabs.AlignWithMargins := True;
  FTabTarget := TTabSheet.Create(Self);
  FTabTarget.PageControl := FTabs;
  FTabTarget.Caption := 'Target';
  FTabSig := TTabSheet.Create(Self);
  FTabSig.PageControl := FTabs;
  FTabSig.Caption := 'Signature';

  // Target tab: unit and class are picked separately - the target class may
  // live in another unit, and then the unit has to be named first.
  FKeepHere := TCheckBox.Create(Self);
  FKeepHere.Parent := FTabTarget;
  FKeepHere.Left := 12;
  FKeepHere.Top := 12;
  FKeepHere.Width := 520;
  FKeepHere.Caption := 'Leave the members where they are (only edit modifiers / signature)';
  FKeepHere.Checked := True;
  FKeepHere.OnClick := Changed;

  Lbl(FTabTarget, 'Unit:', 12, 48);
  FUnitCombo := Combo(FTabTarget, 90, 44, 420);
  FUnitCombo.OnChange := DoUnitChange;
  FBtnBrowse := TButton.Create(Self);
  FBtnBrowse.Parent := FTabTarget;
  FBtnBrowse.Left := 520;
  FBtnBrowse.Top := 43;
  FBtnBrowse.Width := 90;
  FBtnBrowse.Caption := 'Browse...';
  FBtnBrowse.OnClick := DoBrowse;

  Lbl(FTabTarget, 'Class:', 12, 84);
  FClassCombo := Combo(FTabTarget, 90, 80, 420);
  FClassCombo.OnChange := Changed;

  Lbl(FTabTarget, 'Section:', 12, 120);
  FSectionCombo := Combo(FTabTarget, 90, 116, 180);
  for var S in SectionNames do FSectionCombo.Items.Add(S);
  FSectionCombo.ItemIndex := 0;
  FSectionCombo.OnChange := Changed;

  var Hint := Lbl(FTabTarget, 'The body travels along and is requalified. ' +
    'Fields and methods of the old class that it uses stay behind - the box ' +
    'below the tabs names them, and so does everything else this will not do.',
    12, 150);
  Hint.AutoSize := False;
  Hint.WordWrap := True;
  Hint.Width := 600;
  Hint.Height := 34;

  // Signature tab: the grid edits ONE member, the modifiers reach all of the
  // selected ones - which is what the user asked for ("Modifier wie virtual
  // auch bei allen parallel").
  SigTop := TPanel.Create(Self);
  SigTop.Parent := FTabSig;
  SigTop.Align := alTop;
  SigTop.Height := 24;
  SigTop.BevelOuter := bvNone;
  FSigFor := TLabel.Create(Self);
  FSigFor.Parent := SigTop;
  FSigFor.Align := alLeft;
  FSigFor.Layout := tlCenter;
  FSigFor.Caption := 'Parameters of: (nothing selected)';

  Mods := TPanel.Create(Self);
  Mods.Parent := FTabSig;
  Mods.Align := alBottom;
  Mods.Height := 52;
  Mods.BevelOuter := bvNone;
  var ModHint := TLabel.Create(Self);
  ModHint.Parent := Mods;
  ModHint.Left := 8;
  ModHint.Top := 4;
  ModHint.Caption := 'Directives of EVERY ticked member:  grey = leave as it ' +
    'is,  ticked = add,  empty = remove';
  for var I := 0 to High(ModifierNames) do
  begin
    FModBoxes[I] := TCheckBox.Create(Self);
    FModBoxes[I].Parent := Mods;
    FModBoxes[I].Left := 8 + I * 118;
    FModBoxes[I].Top := 26;
    FModBoxes[I].Width := 114;
    FModBoxes[I].Caption := ModifierNames[I];
    FModBoxes[I].AllowGrayed := True;
    FModBoxes[I].State := cbGrayed;
    FModBoxes[I].OnClick := Changed;
  end;

  SigSide := TPanel.Create(Self);
  SigSide.Parent := FTabSig;
  SigSide.Align := alRight;
  SigSide.Width := 150;
  SigSide.BevelOuter := bvNone;
  FSigDown := Btn(SigSide, 'Move &down', alTop, DoSigMove);
  FSigUp := Btn(SigSide, 'Move &up', alTop, DoSigMove);
  FSigRemove := Btn(SigSide, '&Remove', alTop, DoSigRemove);
  FSigAdd := Btn(SigSide, '&Add', alTop, DoSigAdd);

  FGrid := TStringGrid.Create(Self);
  FGrid.Parent := FTabSig;
  FGrid.Align := alClient;
  FGrid.AlignWithMargins := True;
  FGrid.ColCount := 4;
  FGrid.FixedCols := 0;
  FGrid.FixedRows := 1;
  FGrid.RowCount := 2;
  FGrid.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine,
    goColSizing, goEditing, goTabs, goAlwaysShowEditor];
  FGrid.ColWidths[ColMod] := 80;
  FGrid.ColWidths[ColName] := 150;
  FGrid.ColWidths[ColType] := 200;
  FGrid.ColWidths[ColDefault] := 120;
  FGrid.Cells[ColMod, 0] := 'Modifier';
  FGrid.Cells[ColName, 0] := 'Name';
  FGrid.Cells[ColType, 0] := 'Type';
  FGrid.Cells[ColDefault, 0] := 'Default';
  FGrid.OnSetEditText := DoSetEditText;

  // ---- what this will do, and what it will not ---------------------------
  FWarn := TMemo.Create(Self);
  FWarn.Parent := Right;
  FWarn.Align := alTop;
  FWarn.Top := 400;
  FWarn.Height := 110;
  FWarn.AlignWithMargins := True;
  FWarn.ReadOnly := True;
  FWarn.ScrollBars := ssVertical;

  FCallsLabel := TLabel.Create(Self);
  FCallsLabel.Parent := Right;
  FCallsLabel.Align := alTop;
  FCallsLabel.Top := 520;
  FCallsLabel.AlignWithMargins := True;
  FCallsLabel.Caption := 'Occurrences';

  FCalls := TListView.Create(Self);
  FCalls.Parent := Right;
  FCalls.Align := alClient;
  FCalls.AlignWithMargins := True;
  FCalls.ViewStyle := vsReport;
  FCalls.ReadOnly := True;
  FCalls.RowSelect := True;
  Col := FCalls.Columns.Add; Col.Caption := 'Member'; Col.Width := 130;
  Col := FCalls.Columns.Add; Col.Caption := 'Kind'; Col.Width := 120;
  Col := FCalls.Columns.Add; Col.Caption := 'File'; Col.Width := 180;
  Col := FCalls.Columns.Add; Col.Caption := 'Line'; Col.Width := 60;
  Col := FCalls.Columns.Add; Col.Caption := 'Code'; Col.Width := 420;
  FCalls.OnDblClick := DoCallDblClick;
  EnableListViewSorting(FCalls);

  FTimer := TTimer.Create(Self);
  FTimer.Enabled := False;
  FTimer.Interval := 250;
  FTimer.OnTimer := DoTimer;

  FillMembers;
  FillUnits;
  SyncSignatureTab;   // FillMembers ticks without raising OnChange
  Replan;
  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
  ActiveControl := FMembers;
end;

procedure TMethodEditDialog.FillMembers;
begin
  FFilling := True;
  try
    FMembers.Items.BeginUpdate;
    try
      FMembers.Items.Clear;
      for var I := 0 to High(FAn.Members) do
      begin
        var M := FAn.Members[I];
        var It := FMembers.Items.Add;
        It.Caption := M.Name;
        It.SubItems.Add(M.Kind);
        if M.Movable then It.SubItems.Add('') else It.SubItems.Add(M.Why);
        It.Data := Pointer(NativeInt(I));
        It.Checked := False;
        for var P in FAn.Preselected do
          if SameText(M.Name, P) then
          begin
            It.Checked := True;
            Break;
          end;
        if It.Checked then It.Selected := True;
      end;
    finally
      FMembers.Items.EndUpdate;
    end;
  finally
    FFilling := False;
  end;
end;

procedure TMethodEditDialog.FillUnits;
begin
  FFilling := True;
  try
    FUnitCombo.Items.BeginUpdate;
    try
      FUnitCombo.Items.Clear;
      for var F in FAn.UnitFiles do
        FUnitCombo.Items.AddObject(ExtractFileName(F), nil);
    finally
      FUnitCombo.Items.EndUpdate;
    end;
    if FUnitCombo.Items.Count > 0 then FUnitCombo.ItemIndex := 0;
  finally
    FFilling := False;
  end;
  FillClasses;
end;

procedure TMethodEditDialog.FillClasses;
var
  Names: TArray<string>;
  Path: string;
begin
  FFilling := True;
  try
    FClassCombo.Items.Clear;
    FTargetContent := '';
    var I := FUnitCombo.ItemIndex;
    if (I < 0) or (I > High(FAn.UnitFiles)) then Exit;
    Path := FAn.UnitFiles[I];
    if SameText(ExpandFileName(Path), ExpandFileName(FAn.SourceFile)) then
      FTargetContent := FAn.SourceContent
    else
      FTargetContent := ReadUnitForEdit(Path);
    Names := ClassNamesOf(SplitContentLines(FTargetContent));
    for var N in Names do
      if not SameText(N, FAn.OwnerType) or
         not SameText(ExpandFileName(Path), ExpandFileName(FAn.SourceFile)) then
        FClassCombo.Items.Add(N);
    if FClassCombo.Items.Count > 0 then
    begin
      FClassCombo.ItemIndex := 0;
      for var K := 0 to FClassCombo.Items.Count - 1 do
        if IsClassType(SplitContentLines(FTargetContent),
          FClassCombo.Items[K]) then
        begin
          FClassCombo.ItemIndex := K;
          Break;
        end;
    end;
  finally
    FFilling := False;
  end;
end;

procedure TMethodEditDialog.FillGrid;
begin
  FFilling := True;
  try
    FGrid.RowCount := Max(2, Length(FRows) + 1);
    for var C := 0 to FGrid.ColCount - 1 do FGrid.Cells[C, 1] := '';
    for var I := 0 to High(FRows) do
    begin
      FGrid.Cells[ColMod, I + 1] := FRows[I].Modifier;
      FGrid.Cells[ColName, I + 1] := FRows[I].Name;
      FGrid.Cells[ColType, I + 1] := FRows[I].TypeText;
      FGrid.Cells[ColDefault, I + 1] := FRows[I].DefaultText;
    end;
  finally
    FFilling := False;
  end;
end;

procedure TMethodEditDialog.ReadGrid;
begin
  for var I := 0 to High(FRows) do
  begin
    FRows[I].Modifier := Trim(FGrid.Cells[ColMod, I + 1]);
    FRows[I].Name := Trim(FGrid.Cells[ColName, I + 1]);
    FRows[I].TypeText := Trim(FGrid.Cells[ColType, I + 1]);
    FRows[I].DefaultText := Trim(FGrid.Cells[ColDefault, I + 1]);
  end;
end;

procedure TMethodEditDialog.FillCalls;
var
  Calls: TArray<TMethodEditCall>;
begin
  Calls := FAn.CallsOf(SelectedMembers);
  FCalls.Items.BeginUpdate;
  try
    FCalls.Items.Clear;
    for var I := 0 to High(Calls) do
    begin
      var C := Calls[I];
      var It := FCalls.Items.Add;
      It.Caption := C.Member;
      It.SubItems.Add(C.Kind);
      It.SubItems.Add(ExtractFileName(C.FilePath));
      It.SubItems.Add(IntToStr(C.Line + 1));
      It.SubItems.Add(C.Text);
      It.Data := Pointer(NativeInt(I));
    end;
  finally
    FCalls.Items.EndUpdate;
  end;
  var Total := 0;
  for var N in SelectedMembers do Total := Total + FAn.OccurrenceOf(N).Total;
  FCallsLabel.Caption := Format('%s of the selected member(s) - text matches, ' +
    'NOT verified by DelphiLSP and NOT changed by Apply. Double-click to go ' +
    'there.', [IfThen(Total > Length(Calls),
    Format('%d of %d occurrence(s)', [Length(Calls), Total]),
    Format('%d occurrence(s)', [Length(Calls)]))]);
end;

function TMethodEditDialog.SelectedMembers: TArray<string>;
begin
  Result := nil;
  for var I := 0 to FMembers.Items.Count - 1 do
    if FMembers.Items[I].Checked then
      Result := Result + [FMembers.Items[I].Caption];
end;

procedure TMethodEditDialog.BuildRequest;
begin
  FReq := Default(TMethodEditRequest);
  FReq.Members := SelectedMembers;
  if not FKeepHere.Checked then
  begin
    var I := FUnitCombo.ItemIndex;
    if (I >= 0) and (I <= High(FAn.UnitFiles)) then FReq.TargetFile := FAn.UnitFiles[I];
    if FClassCombo.ItemIndex >= 0 then
      FReq.TargetClass := FClassCombo.Items[FClassCombo.ItemIndex];
    if FSectionCombo.ItemIndex >= 0 then
      FReq.Section := FSectionCombo.Items[FSectionCombo.ItemIndex];
  end;
  for var I := 0 to High(ModifierNames) do
    case FModBoxes[I].State of
      cbChecked: FReq.AddModifiers := FReq.AddModifiers + [ModifierNames[I]];
      cbUnchecked: FReq.RemoveModifiers := FReq.RemoveModifiers + [ModifierNames[I]];
    end;
  if FSigMember <> '' then
  begin
    ReadGrid;
    // Only when the list really differs - otherwise a rewrite of both
    // headers would be noise in the diff.
    var Old := FAn.ParamsOf(FSigMember);
    if FormatParamList(FRows) <> FormatParamList(Old) then
    begin
      FReq.SigMember := FSigMember;
      FReq.SigParams := FRows;
      FReq.SigResultType := FSigResult;
    end;
  end;
end;

procedure TMethodEditDialog.Replan;
begin
  BuildRequest;
  FUnitCombo.Enabled := not FKeepHere.Checked;
  FClassCombo.Enabled := not FKeepHere.Checked;
  FSectionCombo.Enabled := not FKeepHere.Checked;
  FBtnBrowse.Enabled := not FKeepHere.Checked;
  FRes := PlanMethodEdit(FAn.SourceFile, FAn.SourceContent, FAn.OwnerType,
    FReq, FTargetContent);
  // The analysis has things to say that no plan can know - above all the form
  // designer binding a handler. They belong in the SAME box; a warning the
  // dialog never shows is not a warning.
  var Txt := '';
  for var N in FAn.Notes do Txt := Txt + N + sLineBreak;
  FWarn.Text := Txt + FRes.Summary;
  FBtnApply.Enabled := FRes.Ok;
  FillCalls;
end;

procedure TMethodEditDialog.Changed(Sender: TObject);
begin
  if FFilling then Exit;
  FTimer.Enabled := False;
  FTimer.Enabled := True;
end;

procedure TMethodEditDialog.DoTimer(Sender: TObject);
begin
  FTimer.Enabled := False;
  Replan;
end;

// The parameter grid belongs to ONE member: the first ticked one.
procedure TMethodEditDialog.SyncSignatureTab;
begin
  var Sel := SelectedMembers;
  var Want := '';
  if Length(Sel) > 0 then Want := Sel[0];
  if SameText(Want, FSigMember) and (FGrid.RowCount = Max(2, Length(FRows) + 1)) then
    Exit;
  FSigMember := Want;
  FRows := FAn.ParamsOf(FSigMember);
  FSigResult := FAn.ResultTypeOf(FSigMember);
  FillGrid;
  if FSigMember = '' then
    FSigFor.Caption := 'Parameters of: (tick a member on the left)'
  else
    FSigFor.Caption := Format('Parameters of %s.%s%s', [FAn.OwnerType,
      FSigMember, IfThen(FSigResult <> '', ': ' + FSigResult, '')]);
  FSigAdd.Enabled := FSigMember <> '';
  FSigRemove.Enabled := FSigMember <> '';
  FSigUp.Enabled := FSigMember <> '';
  FSigDown.Enabled := FSigMember <> '';
end;

procedure TMethodEditDialog.DoMemberChange(Sender: TObject; AItem: TListItem;
  AChange: TItemChange);
begin
  if FFilling then Exit;
  SyncSignatureTab;
  Changed(Sender);
end;

procedure TMethodEditDialog.DoUnitChange(Sender: TObject);
begin
  if FFilling then Exit;
  FillClasses;
  Changed(Sender);
end;

procedure TMethodEditDialog.DoBrowse(Sender: TObject);
var
  Dlg: TOpenDialog;
begin
  Dlg := TOpenDialog.Create(Self);
  try
    Dlg.Filter := 'Delphi unit (*.pas)|*.pas';
    Dlg.Options := Dlg.Options + [ofFileMustExist];
    if ExtractFilePath(FAn.SourceFile) <> '' then
      Dlg.InitialDir := ExtractFilePath(FAn.SourceFile);
    if not Dlg.Execute(Handle) then Exit;
    var Found := -1;
    for var I := 0 to High(FAn.UnitFiles) do
      if SameText(ExpandFileName(FAn.UnitFiles[I]), ExpandFileName(Dlg.FileName)) then
        Found := I;
    if Found < 0 then
    begin
      FAn.UnitFiles := FAn.UnitFiles + [Dlg.FileName];
      FUnitCombo.Items.Add(ExtractFileName(Dlg.FileName));
      Found := High(FAn.UnitFiles);
    end;
    FUnitCombo.ItemIndex := Found;
    FillClasses;
    Changed(Sender);
  finally
    Dlg.Free;
  end;
end;

procedure TMethodEditDialog.DoSigAdd(Sender: TObject);
begin
  if FSigMember = '' then Exit;
  ReadGrid;
  var P := Default(TSigParam);
  P.Name := 'AValue';
  P.TypeText := 'Integer';
  FRows := FRows + [P];
  FillGrid;
  Changed(Sender);
end;

procedure TMethodEditDialog.DoSigRemove(Sender: TObject);
begin
  ReadGrid;
  var I := FGrid.Row - 1;
  if (I < 0) or (I > High(FRows)) then Exit;
  Delete(FRows, I, 1);
  FillGrid;
  Changed(Sender);
end;

procedure TMethodEditDialog.DoSigMove(Sender: TObject);
begin
  ReadGrid;
  var I := FGrid.Row - 1;
  var J := I + IfThen(Sender = FSigUp, -1, 1);
  if (I < 0) or (I > High(FRows)) or (J < 0) or (J > High(FRows)) then Exit;
  var T := FRows[I];
  FRows[I] := FRows[J];
  FRows[J] := T;
  FillGrid;
  FGrid.Row := J + 1;
  Changed(Sender);
end;

procedure TMethodEditDialog.DoCallDblClick(Sender: TObject);
begin
  if (FCalls.Selected = nil) or (Editor = nil) then Exit;
  var Calls := FAn.CallsOf(SelectedMembers);
  var I := NativeInt(FCalls.Selected.Data);
  if (I < 0) or (I > High(Calls)) then Exit;
  Editor.GotoLocation(Calls[I].FilePath, Calls[I].Line, Calls[I].Col,
    Length(Calls[I].Member));
end;

procedure TMethodEditDialog.DoSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: string);
begin
  if FFilling then Exit;
  Changed(Sender);
end;

procedure TMethodEditDialog.DoApply(Sender: TObject);
var
  Err: string;
begin
  Replan;
  if not FRes.Ok then Exit;
  if not ApplyMethodEdit(FAn, FReq, Err) then
  begin
    ShowThemedMessage('Edit methods: ' + Err);
    Exit;
  end;
  Applied := True;
  ModalResult := mrOk;
end;

// ---------------------------------------------------------------------------
//  Editor entry point
// ---------------------------------------------------------------------------

type
  TMethodEditJob = class(TInterfacedObject)
  public
    Lock: TCriticalSection;
    Cur, Total: Integer;
    Text: string;
    Done: Boolean;
    Res: TMethodEditAnalysis;
    Stop: TEvent;
    constructor Create;
    destructor Destroy; override;
  end;

constructor TMethodEditJob.Create;
begin
  inherited;
  Lock := TCriticalSection.Create;
  Stop := TEvent.Create(nil, True, False, '');
end;

destructor TMethodEditJob.Destroy;
begin
  Stop.Free;
  Lock.Free;
  inherited;
end;

function CreateMethodEditDialog(AOwner: TComponent;
  const AAn: TMethodEditAnalysis): TForm;
begin
  Result := TMethodEditDialog.CreateDialog(AOwner, AAn);
end;

procedure EditMethodsAtCursor;
var
  Inp: TSafeDeleteInput;
  Err: string;
  Job: TMethodEditJob;
  JobRef: IInterface;
  Prog: TCheckProgressWindow;
begin
  if Editor = nil then Exit;
  // WHAT IS SELECTED, FIRST - before anything else touches the editor.
  // GetCurrentContext walks the caret with IOTAEditPosition.MoveRelative to
  // find the word under it, and that DESTROYS the selection block (its
  // Save/Restore puts the position back, not the block). Reported by the
  // user: "leider verschwindet nach dem Klicken des Menueintrags vor der
  // Anzeige des Fensters die Markierung" - the marks were gone and the
  // dialog ticked the caret's member again. Extract method, extract
  // variable and the property converter all read the selection as their
  // first statement, which is exactly why they never had this bug.
  // Main thread only, so the worker gets two plain line numbers.
  var SelFrom := -1;
  var SelTo := -1;
  var SelFile := '';
  begin
    var SF, SC, EL, EC: Integer;
    var SelText: string;
    if Editor.GetSelection(SelFile, SF, SC, EL, EC, SelText) then
    begin
      SelFrom := Max(0, SF - 1);
      SelTo := Max(0, EL - 1);
      // A selection that ends in column 1 does not reach that line's text.
      if (EC <= 1) and (SelTo > SelFrom) then Dec(SelTo);
    end;
  end;

  // The caret through the CHEAP getters, never GetCurrentContext: that one
  // walks the edit position to read the word under the cursor and collapses
  // the selection on the way (Expert.EditorHelper says so where
  // GetActiveFileName is implemented). This command needs a file and a
  // caret, nothing else - the member at the caret is resolved from the
  // CONTENT by IdentifierAtPos, which steps back by itself when the caret
  // sits just behind the name.
  var CaretFile := Editor.GetActiveFileName;
  if CaretFile = '' then
  begin
    ShowThemedMessage('Edit methods: no file at the cursor.');
    Exit;
  end;
  var CaretLine := 1;
  var CaretCol := 1;
  Editor.GetCaretLineCol(CaretLine, CaretCol);
  // Only a selection in the file the caret is in can mean anything for this
  // class's member list.
  if (SelFile = '') or not SameText(ExpandFileName(SelFile),
       ExpandFileName(CaretFile)) then
  begin
    SelFrom := -1;
    SelTo := -1;
  end;

  Editor.SaveAllFiles;   // the scan reads the other units from disk
  if not GatherScanInput(CaretFile, Max(0, CaretLine - 1),
    Max(0, CaretCol - 1), Inp, Err) then
  begin
    ShowThemedMessage('Edit methods: ' + Err);
    Exit;
  end;

  Job := TMethodEditJob.Create;
  JobRef := Job;
  Prog := CreateCheckProgress('Edit methods', Application.MainForm,
    'Reading ' + ExtractFileName(CaretFile) + '...');
  try
    var ThreadRef: IInterface := JobRef;
    var StopHandle := Job.Stop.Handle;
    if not StartWorker(
      procedure
      var
        R: TMethodEditAnalysis;
      begin
        try
          R := AnalyzeMethodEdit(Inp, StopHandle,
            procedure(ACur, ATotal: Integer; AText: string)
            begin
              Job.Lock.Enter;
              try
                Job.Cur := ACur; Job.Total := ATotal; Job.Text := AText;
              finally
                Job.Lock.Leave;
              end;
            end, SelFrom, SelTo);
        except
          on E: Exception do
          begin
            R := Default(TMethodEditAnalysis);
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
      ShowThemedMessage('Edit methods: the plugin is shutting down.');
      Exit;
    end;
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
  var Dlg := TMethodEditDialog.CreateDialog(Application.MainForm, Job.Res);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  MCP tool "edit_method"
// ---------------------------------------------------------------------------

function StringArrayArg(AArgs: TJSONObject; const AName: string): TArray<string>;
var
  V: TJSONValue;
begin
  Result := nil;
  if AArgs = nil then Exit;
  V := AArgs.Values[AName];
  if V is TJSONArray then
  begin
    for var E in TJSONArray(V) do
      if Trim(E.Value) <> '' then Result := Result + [Trim(E.Value)];
  end
  else if (V <> nil) and (Trim(V.Value) <> '') then
    Result := [Trim(V.Value)];
end;

function ToolEditMethod(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, GatherErr, TargetUnit: string;
  L1, C1: Integer;
  Inp: TSafeDeleteInput;
  Ok: Boolean;
  An: TMethodEditAnalysis;
  Req: TMethodEditRequest;
  Res: TMethodEditResult;
  TgtContent: string;
begin
  F := AArgs.GetValue<string>('file', '');
  L1 := AArgs.GetValue<Integer>('line', 0);
  C1 := AArgs.GetValue<Integer>('column', 1);
  if (F = '') or (L1 < 1) then
    Exit(McpErr('arguments "file" and "line" (1-based) are required - put the ' +
      'position on a member of the class whose methods are edited'));
  F := ExpandFileName(F);
  Ok := False;
  if not McpRunOnMain(
    procedure
    var
      E: string;
    begin
      Ok := GatherScanInput(F, L1 - 1, Max(0, C1 - 1), Inp, E);
      if not Ok then GatherErr := E;
    end, True, AStop, Err) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(GatherErr));

  An := AnalyzeMethodEdit(Inp, AStop, nil);
  var J := TJSONObject.Create;
  J.AddPair('summary', An.Summary);
  if not An.Ok then
  begin
    J.AddPair('error', An.Error);
    Exit(McpOk(J));
  end;
  J.AddPair('class', An.OwnerType);
  if An.CaretMember <> '' then J.AddPair('member_at_position', An.CaretMember);
  var MA := TJSONArray.Create;
  for var M in An.Members do
  begin
    var O := TJSONObject.Create;
    O.AddPair('name', M.Name);
    O.AddPair('kind', M.Kind);
    O.AddPair('line', TJSONNumber.Create(M.DeclLine + 1));
    if M.Directives <> '' then O.AddPair('directives', M.Directives);
    O.AddPair('movable', TJSONBool.Create(M.Movable));
    if not M.Movable then O.AddPair('why_not', M.Why);
    MA.Add(O);
  end;
  J.AddPair('members', MA);
  for var N in An.Notes do
    J.AddPair('note', N);

  Req := Default(TMethodEditRequest);
  Req.Members := StringArrayArg(AArgs, 'members');
  Req.TargetClass := Trim(AArgs.GetValue<string>('target_class', ''));
  Req.Section := LowerCase(Trim(AArgs.GetValue<string>('section', 'private')));
  Req.AddModifiers := StringArrayArg(AArgs, 'add_modifiers');
  Req.RemoveModifiers := StringArrayArg(AArgs, 'remove_modifiers');
  TargetUnit := Trim(AArgs.GetValue<string>('target_unit', ''));

  // Nothing asked for = the list above IS the answer: which members exist,
  // which of them can travel alone, and what the form file says.
  // How often each name occurs - cheap, and it is what tells you whether
  // looking at the list is worth it. The ROWS come below, and only for the
  // members the caller named.
  var OC := TJSONArray.Create;
  for var Occ in An.Occurrences do
  begin
    var O := TJSONObject.Create;
    O.AddPair('member', Occ.Member);
    O.AddPair('occurrences', TJSONNumber.Create(Occ.Total));
    if Occ.Listed < Occ.Total then
      O.AddPair('listed', TJSONNumber.Create(Occ.Listed));
    OC.Add(O);
  end;
  J.AddPair('occurrenceCounts', OC);

  if (Length(Req.Members) = 0) and not Req.WantsSomething then
  begin
    J.AddPair('hint', 'pass "members" plus "target_class" (and "target_unit" ' +
      'when the class lives elsewhere) or "add_modifiers" / ' +
      '"remove_modifiers"; a parameter list is changed by change_signature, ' +
      'which verifies and rewrites the call sites. The occurrences of the ' +
      'members you name come with that call - they are text matches, not ' +
      'verified, and never rewritten');
    Exit(McpOk(J));
  end;

  // The target unit, by path or by unit name.
  if TargetUnit <> '' then
  begin
    var Found := '';
    for var U in An.UnitFiles do
      if SameText(ExtractFileName(U), TargetUnit) or
         SameText(ChangeFileExt(ExtractFileName(U), ''), TargetUnit) or
         SameText(ExpandFileName(U), ExpandFileName(TargetUnit)) then
        Found := U;
    if (Found = '') and FileExists(ExpandFileName(TargetUnit)) then
      Found := ExpandFileName(TargetUnit);
    if Found = '' then
    begin
      J.AddPair('error', Format('the target unit "%s" is not in the project ' +
        'scope and not a file on disk', [TargetUnit]));
      Exit(McpOk(J));
    end;
    Req.TargetFile := Found;
  end
  else if Req.TargetClass <> '' then
    Req.TargetFile := An.SourceFile;

  TgtContent := '';
  if (Req.TargetFile <> '') and
     not SameText(ExpandFileName(Req.TargetFile), ExpandFileName(An.SourceFile)) then
  begin
    var TC := '';
    if not McpRunOnMain(
      procedure
      begin
        TC := ReadUnitForEdit(Req.TargetFile);
      end, False, AStop, Err) then Exit(McpErr(Err));
    TgtContent := TC;
    if TgtContent = '' then
    begin
      J.AddPair('error', Req.TargetFile + ' could not be read');
      Exit(McpOk(J));
    end;
  end;

  var Sel := An.CallsOf(Req.Members);
  var CA := TJSONArray.Create;
  for var I := 0 to High(Sel) do
  begin
    if I >= MaxToolOccurrences then
    begin
      J.AddPair('occurrencesShown', TJSONNumber.Create(MaxToolOccurrences));
      Break;
    end;
    var O := TJSONObject.Create;
    O.AddPair('member', Sel[I].Member);
    O.AddPair('kind', Sel[I].Kind);
    O.AddPair('file', Sel[I].FilePath);
    O.AddPair('line', TJSONNumber.Create(Sel[I].Line + 1));
    O.AddPair('text', Sel[I].Text);
    CA.Add(O);
  end;
  J.AddPair('occurrences', CA);
  J.AddPair('occurrencesNote', 'text matches of the selected member(s), NOT ' +
    'verified by DelphiLSP and NOT changed by this tool');

  Res := PlanMethodEdit(An.SourceFile, An.SourceContent, An.OwnerType, Req,
    TgtContent);
  J.AddPair('plan', Res.Summary);
  if not Res.Ok then
  begin
    J.AddPair('error', Res.Error);
    Exit(McpOk(J));
  end;
  if Length(Res.Moved) > 0 then
    J.AddPair('moved', string.Join(', ', Res.Moved));
  var IA := TJSONArray.Create;
  for var I in Res.Issues do
  begin
    var O := TJSONObject.Create;
    case I.Kind of
      meiVeto: O.AddPair('kind', 'veto');
      meiOwnMember: O.AddPair('kind', 'stays_behind');
    else
      O.AddPair('kind', 'note');
    end;
    if I.Member <> '' then O.AddPair('member', I.Member);
    O.AddPair('text', I.Text);
    IA.Add(O);
  end;
  J.AddPair('issues', IA);

  // The edits, as every writing tool of this bridge reports them.
  var FA := TJSONArray.Create;
  var AddFile := procedure(const AFile, AOld: string; const ANew: TArray<string>)
    var
      T: Integer;
    begin
      if (AFile = '') or (Length(ANew) = 0) then Exit;
      var O := TJSONObject.Create;
      O.AddPair('file', AFile);
      O.AddPair('changes', ChangesToJson(DiffToChanges(AFile, AOld,
        string.Join(sLineBreak, ANew), 200, T)));
      O.AddPair('changed_lines', TJSONNumber.Create(T));
      FA.Add(O);
    end;
  AddFile(An.SourceFile, An.SourceContent, Res.SourceLines);
  if Res.TargetFile <> '' then
    AddFile(Res.TargetFile, TgtContent, Res.TargetLines);
  J.AddPair('files', FA);

  if not AArgs.GetValue<Boolean>('apply', False) then
  begin
    J.AddPair('applied', TJSONBool.Create(False));
    J.AddPair('next', 'call again with "apply": true to write this');
    Exit(McpOk(J));
  end;

  var AppErr := '';
  var Written := False;
  if not McpRunOnMain(
    procedure
    begin
      Written := ApplyMethodEdit(An, Req, AppErr);
    end, False, AStop, Err) then Exit(McpErr(Err));
  J.AddPair('applied', TJSONBool.Create(Written));
  if not Written then J.AddPair('error', AppErr);
  Result := McpOk(J);
end;

initialization
  RegisterMcpTool('edit_method', ToolEditMethod);

end.
