(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DfmEventCheck;

// Checks .dfm files for event references that have no matching handler
// method in the paired .pas file - the classic "Method 'BtnOkClick'
// not found" crash at form load time that the compiler never warns
// about. Also flags handlers whose parameter list does not match the
// event type's expected signature (for a table of well-known VCL
// events); a mismatched signature loads fine but corrupts the stack
// when the event fires.
//
// Scope / limitations (v1):
//   * Text DFMs only. Binary DFMs are skipped (reported as unparsed).
//   * The handler search covers the form class named in the DFM root
//     plus its ancestor classes as far as they can be resolved inside
//     the project files (visual form inheritance).
//   * Signature checking only applies to events in the built-in table;
//     unknown events are checked for handler existence only.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Lsp.Client;

type
  /// <summary>Batch state for the auto-fix: an LSP client plus a cache of
  ///  "type name -> declaring unit", so the correct unit can be added to a
  ///  form's uses clause when a corrected signature introduces a type that
  ///  is not yet reachable. Recurring types (TShiftState, TMouseButton, ...)
  ///  are resolved once and served from the cache thereafter.</summary>
  TFixContext = class
  private
    FTypeUnit: TDictionary<string, string>;    // UPPER(type) -> unit ('' unknown)
    FRefreshed: TDictionary<string, Boolean>;  // files already didChange'd
    FClient: TLspClient;
    FClientTried: Boolean;
    function GetClient: TLspClient;
  public
    constructor Create;
    destructor Destroy; override;
    /// <summary>Declaring unit of AType (resolved via LSP GotoDefinition at
    ///  the type's position in the event type declaration, then cached).
    ///  Returns '' for builtin/System types or when LSP is unavailable.</summary>
    function UnitOfType(const AType, AEventTypeFile: string;
      AEventTypeLine: Integer): string;
  end;

  TDfmEventIssueKind = (eikMissingHandler, eikSignatureMismatch);

  TDfmEventIssue = record
    DfmFile: string;
    PasFile: string;
    ComponentName: string;
    ComponentType: string;
    EventName: string;
    HandlerName: string;
    /// <summary>1-based line of the event assignment in the .dfm.</summary>
    DfmLine: Integer;
    /// <summary>1-based line of the handler declaration in the .pas
    ///  (0 when the handler is missing).</summary>
    PasLine: Integer;
    Kind: TDfmEventIssueKind;
    /// <summary>Human-readable expected / actual parameter lists,
    ///  filled for eikSignatureMismatch.</summary>
    Expected: string;
    Actual: string;
    /// <summary>The exact replacement parameter list WITH names
    ///  ("Sender: TBaseVirtualTree; Node: PVirtualNode"), only set when
    ///  the signature was resolved from source and the handler can be
    ///  auto-corrected. Empty = not auto-fixable.</summary>
    ExpectedRawParams: string;
    /// <summary>Form class the handler belongs to (for locating the
    ///  implementation "procedure TForm.Handler").</summary>
    FormClass: string;
    /// <summary>The event type name (e.g. TVTMeasureTextEvent) and the
    ///  location of its declaration, when resolved from source. Lets
    ///  the dialog jump to the "TXxx = procedure(...)" line.</summary>
    EventTypeName: string;
    EventTypeFile: string;
    EventTypeLine: Integer;
    /// <summary>Normalized (pipe-separated) expected / actual parameter
    ///  types, for token-by-token diff highlighting in the dialog.</summary>
    ExpectedNorm: string;
    ActualNorm: string;
  end;

  /// <summary>One event handler wired in a DFM, with the handler's
  ///  actual (normalized) signature. Collected across all forms so a
  ///  cross-project consistency check can flag the odd one out for
  ///  events whose type we cannot know (third-party, source-less
  ///  components like TVirtualStringTree).</summary>
  TEventAssignment = record
    CompType: string;
    EventName: string;
    HandlerName: string;
    Signature: string;
    ComponentName: string;
    DfmFile: string;
    PasFile: string;
    DfmLine: Integer;
    PasLine: Integer;
  end;

  TDfmEventChecker = class
  private
    /// <summary>File -> lines cache: every file is read at most once
    ///  per run. The ancestor resolution used to re-read files per
    ///  form, which made the check O(forms x files) on big projects.</summary>
    FLinesCache: TDictionary<string, TArray<string>>;
    /// <summary>All handler assignments for events NOT covered by the
    ///  built-in signature table - the pool the consistency pass
    ///  analyses.</summary>
    FAssignments: TList<TEventAssignment>;
    /// <summary>Signature resolver, built from reachable component
    ///  source (project + search-path .pas). Lets us check third-party
    ///  events (TVirtualStringTree, ...) exactly and version-correctly
    ///  by reading the actual .pas the project compiles against.
    ///   FClassAncestor:  UPPER(class)          -> ancestor class name
    ///   FEventPropType:  UPPER(class.event)    -> event type name
    ///   FProcSig:        UPPER(procTypeName)   -> normalized signature</summary>
    FClassAncestor: TDictionary<string, string>;
    FEventPropType: TDictionary<string, string>;
    FProcSig: TDictionary<string, string>;
    /// <summary>UPPER(procTypeName) -> raw parameter list WITH names
    ///  ("Sender: TBaseVirtualTree; Node: PVirtualNode"). Feeds the
    ///  auto-correction: the handler is rewritten to exactly this.</summary>
    FProcRawParams: TDictionary<string, string>;
    /// <summary>UPPER(procTypeName) -> "file|line" of its declaration,
    ///  so the dialog can jump to the event type.</summary>
    FProcLoc: TDictionary<string, string>;
    /// <summary>Simple type aliases "TFoo = TBar;" - VirtualTrees et al.
    ///  route the ancestor chain through IFDEF-guarded aliases
    ///  (TVTAncestor = TVTAncestorVcl). The resolver follows these.</summary>
    FClassAlias: TDictionary<string, string>;
    /// <summary>UPPER(classname) -> declaring file. Built in one pass
    ///  over all project sources; makes ancestor lookup O(1) instead
    ///  of a scan over every file. Classes NOT in the index (TForm,
    ///  TFrame, any VCL/RTL base) simply terminate the ancestor walk.</summary>
    FClassIndex: TDictionary<string, string>;
    function GetLines(const AFile: string): TArray<string>;
    procedure BuildClassIndex(const AFiles: TArray<string>);
    function CheckPairInternal(const ADfmFile, APasFile: string): TArray<TDfmEventIssue>;
    /// <summary>Consistency pass over FAssignments (run after all
    ///  pairs): flags the odd-one-out handler for events with no known
    ///  signature.</summary>
    function ConsistencyIssues: TArray<TDfmEventIssue>;
    /// <summary>Indexes class ancestry, event-property types and
    ///  procedural-type signatures from AFiles (component source).
    ///  AProgress (optional) ticks per file.</summary>
    procedure BuildSignatureIndex(const AFiles: TArray<string>;
      const AProgress: TProc<Integer, Integer, string> = nil);
    /// <summary>Resolves the expected (normalized) signature of an
    ///  event on a component by walking the class chain to the class
    ///  that declares the event property, then its procedural type.
    ///  False when not resolvable from indexed source.</summary>
    function TryResolveEventSignature(const ACompType, AEventName: string;
      out ASignature, ATypeName, ARawParams: string): Boolean;
    class function NormalizeParams(const AParamList: string): string;
    class function ExpectedSignature(const AEventName: string;
      out ASignature: string): Boolean;
    class procedure CollectClassMethods(const ALines: TArray<string>;
      const AClassName: string; AMethods: TDictionary<string, TPair<Integer, string>>;
      out AAncestor: string);
  public
    constructor Create;
    destructor Destroy; override;

    /// <summary>Runs the check for every project .pas that has a
    ///  sibling .dfm. AProgress (optional) reports (current, total,
    ///  filename). ASignatureFiles (optional) are extra component
    ///  source files (from the search path) indexed to resolve
    ///  third-party event signatures exactly.</summary>
    class function CheckProject(const AProjectFiles: TArray<string>;
      const AProgress: TProc<Integer, Integer, string> = nil;
      const ASignatureFiles: TArray<string> = nil): TArray<TDfmEventIssue>;

    /// <summary>Rewrites the handler in AIssue.PasFile - both the class
    ///  declaration and the implementation header - to the expected
    ///  parameter list (AIssue.ExpectedRawParams). Writes through the
    ///  IEditorHelper (undoable in the IDE). Returns False when the
    ///  issue is not auto-fixable or the handler could not be located.</summary>
    class function ApplyFix(const AIssue: TDfmEventIssue;
      out AFailReason: string; AContext: TFixContext = nil): Boolean;
  end;

implementation

uses
  System.IOUtils, System.StrUtils, System.Math,
  Expert.EditorHelperIntf, Delphi.FileEncoding,
  Expert.LspManager, Lsp.Uri, Lsp.Protocol;

// Unit name (possibly dotted, e.g. "Vcl.Controls") declared by AFile, or ''.
function ReadUnitNameOf(const AFile: string): string;
var
  Buf, U: string;
  SL: TStringList;
  I, P: Integer;
begin
  Result := '';
  if AFile = '' then Exit;
  if (Editor = nil) or not Editor.ReadEditorContent(AFile, Buf) then
  begin
    if not TFile.Exists(AFile) then Exit;
    try Buf := TFile.ReadAllText(AFile); except Exit; end;
  end;
  SL := TStringList.Create;
  try
    SL.Text := Buf;
    for I := 0 to Min(SL.Count - 1, 200) do
    begin
      U := Trim(SL[I]);
      if UpperCase(U).StartsWith('UNIT ') then
      begin
        U := Trim(Copy(U, 6, MaxInt));
        P := Pos(';', U);
        if P > 0 then U := Copy(U, 1, P - 1);
        Exit(Trim(U));
      end;
    end;
  finally
    SL.Free;
  end;
end;

// True when AToken sits at ALine[APos] as a whole identifier (its neighbours
// are not identifier characters).
function IsWholeWordAt(const ALine: string; APos, ALen: Integer): Boolean;
begin
  Result := True;
  if (APos > 1) and CharInSet(ALine[APos - 1], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
    Exit(False);
  if (APos + ALen <= Length(ALine)) and
     CharInSet(ALine[APos + ALen], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
    Exit(False);
end;

// First whole-word occurrence of AToken in AFile at/after AStartLine1 (1-based),
// within a short window (a type declaration spans few lines). Returns 0-based
// line/col for LSP.
function FindTokenLineCol(const AFile: string; AStartLine1: Integer;
  const AToken: string; out ALine0, ACol0: Integer): Boolean;
var
  Buf, Line: string;
  SL: TStringList;
  I, P: Integer;
begin
  Result := False;
  if (AToken = '') or (AStartLine1 <= 0) then Exit;
  if (Editor = nil) or not Editor.ReadEditorContent(AFile, Buf) then
  begin
    if not TFile.Exists(AFile) then Exit;
    try Buf := TFile.ReadAllText(AFile); except Exit; end;
  end;
  SL := TStringList.Create;
  try
    SL.Text := Buf;
    for I := AStartLine1 - 1 to Min(SL.Count - 1, AStartLine1 - 1 + 15) do
    begin
      Line := SL[I];
      P := Pos(UpperCase(AToken), UpperCase(Line));
      while P > 0 do
      begin
        if IsWholeWordAt(Line, P, Length(AToken)) then
        begin
          ALine0 := I;
          ACol0 := P - 1;
          Exit(True);
        end;
        P := PosEx(UpperCase(AToken), UpperCase(Line), P + 1);
      end;
    end;
  finally
    SL.Free;
  end;
end;

{ TFixContext }

constructor TFixContext.Create;
begin
  inherited Create;
  FTypeUnit := TDictionary<string, string>.Create;
  FRefreshed := TDictionary<string, Boolean>.Create;
end;

destructor TFixContext.Destroy;
begin
  FTypeUnit.Free;
  FRefreshed.Free;
  inherited;
end;

function TFixContext.GetClient: TLspClient;
var
  Json, Root, Proj: string;
begin
  if not FClientTried then
  begin
    FClientTried := True;
    try
      if Editor <> nil then
      begin
        Json := Editor.FindDelphiLspJson;
        if Json <> '' then
        begin
          Root := Editor.GetProjectRoot;
          Proj := Editor.GetCurrentProjectDproj;
          FClient := TLspManager.Instance.GetClient(Root, Proj, Json);
        end;
      end;
    except
      FClient := nil;
    end;
  end;
  Result := FClient;
end;

function TFixContext.UnitOfType(const AType, AEventTypeFile: string;
  AEventTypeLine: Integer): string;
var
  T, Key: string;
  Cli: TLspClient;
  L0, C0: Integer;
  Defs: TArray<TLspLocation>;
  Dummy: Boolean;
begin
  // Reduce "var TColor" / "const TFoo" / "array of Byte" to the bare type name.
  T := Trim(AType);
  for var Pre in ['VAR ', 'CONST ', 'OUT '] do
    if UpperCase(T).StartsWith(Pre) then T := Trim(Copy(T, Length(Pre) + 1, MaxInt));
  var Sp := T.LastIndexOf(' ');
  if Sp >= 0 then T := Trim(T.Substring(Sp + 1));   // "array of Byte" -> "Byte"
  if T = '' then Exit('');

  Key := UpperCase(T);
  if FTypeUnit.TryGetValue(Key, Result) then Exit;

  Result := '';
  FTypeUnit.Add(Key, '');   // negative-cache first; overwrite on success
  if (AEventTypeFile = '') or (AEventTypeLine <= 0) then Exit;

  Cli := GetClient;
  if Cli = nil then Exit;

  if not FindTokenLineCol(AEventTypeFile, AEventTypeLine, T, L0, C0) then Exit;
  try
    if not FRefreshed.TryGetValue(AEventTypeFile, Dummy) then
    begin
      Cli.RefreshDocument(AEventTypeFile);
      FRefreshed.Add(AEventTypeFile, True);
    end;
    Defs := Cli.GotoDefinition(AEventTypeFile, L0, C0);
  except
    Exit;
  end;
  if Length(Defs) = 0 then Exit;

  Result := ReadUnitNameOf(TLspUri.FileUriToPath(Defs[0].Uri));
  // Only the bare "System" unit is implicit and never needs a uses entry;
  // System.Classes / System.UITypes / ... are real units that must be listed.
  if SameText(Result, 'System') then Result := '';
  FTypeUnit[Key] := Result;
end;

function SplitLines(const AContent: string): TArray<string>;
begin
  // Normalize CRLF / lone LF / lone CR before splitting - files coming
  // out of git with LF-only endings would otherwise parse as one line.
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10)
    .Split([#10], TStringSplitOptions.None);
end;

function ReadFileLines(const AFile: string): TArray<string>;
var
  Content: string;
begin
  Result := nil;
  if (Editor <> nil) and Editor.ReadEditorContent(AFile, Content) then
    Exit(SplitLines(Content));
  if not TFile.Exists(AFile) then Exit;
  try
    Content := TDelphiFileEncoding.ReadAll(AFile);
    Result := SplitLines(Content);
  except
    Result := nil;
  end;
end;

function IsIdent(const S: string): Boolean;
var
  I: Integer;
begin
  Result := S <> '';
  for I := 1 to Length(S) do
    if not CharInSet(S[I], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
      Exit(False);
  if CharInSet(S[1], ['0'..'9']) then Result := False;
end;

{ ---------------- Event signature table ----------------
  Signatures are normalized as pipe-separated entries, one per
  parameter, each "[var |out ]TypeName" (const dropped - it does not
  affect the call ABI for these types and keeps false positives down).
  Multi-name groups (X, Y: Integer) expand to one entry per name. }

type
  TKnownEvent = record
    Name: string;       // event property name, case-insensitive
    Signature: string;  // normalized expected parameter list
  end;

const
  // (Sender: TObject) - the TNotifyEvent family. ONLY events whose
  // type is TNotifyEvent across the entire VCL AND common third-party
  // components belong here. Deliberately NOT listed (component-
  // specific signatures, checking them produced false-positive walls):
  //   OnChange          - TTreeView(Sender, Node), VirtualTree,
  //                       date pickers, ... all differ
  //   OnCellClick       - TDBGrid(Column), TIB_Grid(custom), ...
  //   OnChanged / OnSelectionChange / OnPaint - vary by component
  //   OnStartDock / OnStartDrag - have a var DragObject parameter
  NotifyEvents: array[0..14] of string = (
    'OnClick', 'OnDblClick', 'OnEnter', 'OnExit',
    'OnCreate', 'OnDestroy', 'OnShow', 'OnHide', 'OnActivate',
    'OnDeactivate', 'OnResize', 'OnPopup', 'OnExecute', 'OnUpdate',
    'OnTimer');

  // Component classes whose events from the tables below verifiably
  // carry the standard VCL signatures. The signature check ONLY runs
  // for these (and for the form itself, i.e. the DFM root object).
  // Anything else - third-party grids, VirtualTree, but also standard
  // VCL classes that redeclare an event name with a different type
  // (TUpDown.OnClick is TUDClickEvent!) - gets the existence check
  // only. Extend deliberately, never guess.
  SafeVclClasses: array[0..39] of string = (
    'TButton', 'TBitBtn', 'TSpeedButton', 'TCheckBox', 'TRadioButton',
    'TEdit', 'TLabeledEdit', 'TMaskEdit', 'TMemo', 'TComboBox',
    'TListBox', 'TCheckListBox', 'TPanel', 'TGroupBox', 'TLabel',
    'TStaticText', 'TImage', 'TShape', 'TBevel', 'TSplitter',
    'TScrollBox', 'TPageControl', 'TTabSheet', 'TTabControl',
    'TToolBar', 'TToolButton', 'TMenuItem', 'TMainMenu', 'TPopupMenu',
    'TAction', 'TActionList', 'TTimer', 'TStringGrid', 'TDrawGrid',
    'TListView', 'TTreeView', 'TRadioGroup', 'TTrackBar',
    'TProgressBar', 'TStatusBar');

  KnownEvents: array[0..11] of TKnownEvent = (
    (Name: 'OnKeyDown';    Signature: 'TObject|var Word|TShiftState'),
    (Name: 'OnKeyUp';      Signature: 'TObject|var Word|TShiftState'),
    (Name: 'OnKeyPress';   Signature: 'TObject|var Char'),
    (Name: 'OnMouseDown';  Signature: 'TObject|TMouseButton|TShiftState|Integer|Integer'),
    (Name: 'OnMouseUp';    Signature: 'TObject|TMouseButton|TShiftState|Integer|Integer'),
    (Name: 'OnMouseMove';  Signature: 'TObject|TShiftState|Integer|Integer'),
    (Name: 'OnClose';      Signature: 'TObject|var TCloseAction'),
    (Name: 'OnCloseQuery'; Signature: 'TObject|var Boolean'),
    (Name: 'OnMouseWheel'; Signature: 'TObject|TShiftState|Integer|TPoint|var Boolean'),
    (Name: 'OnDragOver';   Signature: 'TObject|TObject|Integer|Integer|TDragState|var Boolean'),
    (Name: 'OnDragDrop';   Signature: 'TObject|TObject|Integer|Integer'),
    (Name: 'OnCanResize';  Signature: 'TObject|var Integer|var Integer|var Boolean'));

class function TDfmEventChecker.ExpectedSignature(const AEventName: string;
  out ASignature: string): Boolean;
var
  I: Integer;
begin
  for I := Low(KnownEvents) to High(KnownEvents) do
    if SameText(KnownEvents[I].Name, AEventName) then
    begin
      ASignature := KnownEvents[I].Signature;
      Exit(True);
    end;
  for I := Low(NotifyEvents) to High(NotifyEvents) do
    if SameText(NotifyEvents[I], AEventName) then
    begin
      ASignature := 'TObject';
      Exit(True);
    end;
  ASignature := '';
  Result := False;
end;

/// <summary>Resolves compiler conditionals in a parameter string by
///  taking the FIRST (true) branch and stripping the directive markers
///  and any brace / (* *) comments. Makes conditional parameter types
///  ("{$if CompilerVersion >= 36}TVTDragDataObject{$else}IDataObject
///  {$ifend}") comparable to the concrete type the compiler picked.</summary>
function StripLineComment(const L: string): string;
// Remove a trailing "// ..." comment from a single source line, honouring
// single-quoted string literals so a '//' inside a default value stays.
// Must be applied per line BEFORE lines are joined into one signature
// string, otherwise a mid-signature comment would swallow the real
// parameters that follow on the next line.
var
  I, N: Integer;
  InStr: Boolean;
begin
  N := Length(L);
  InStr := False;
  I := 1;
  while I <= N do
  begin
    if L[I] = '''' then
      InStr := not InStr
    else if (not InStr) and (L[I] = '/') and (I < N) and (L[I + 1] = '/') then
      Exit(TrimRight(Copy(L, 1, I - 1)));
    Inc(I);
  end;
  Result := L;
end;

type
  TParamSpec = record
    Modifier, Name, Typ: string;   // e.g. "const", "aErrCode", "Integer"
  end;

// Split S on ASep, but only at the top level (ignoring separators nested in
// () or []). Used to break a parameter list into groups / a group into names.
function SplitTopLevel(const S: string; ASep: Char): TArray<string>;
var
  I, Depth, StartPos: Integer;
begin
  Result := nil;
  Depth := 0;
  StartPos := 1;
  for I := 1 to Length(S) do
  begin
    case S[I] of
      '(', '[': Inc(Depth);
      ')', ']': if Depth > 0 then Dec(Depth);
    else
      if (S[I] = ASep) and (Depth = 0) then
      begin
        Result := Result + [Copy(S, StartPos, I - StartPos)];
        StartPos := I + 1;
      end;
    end;
  end;
  Result := Result + [Copy(S, StartPos, MaxInt)];
end;

// Parse a Pascal parameter list into a flat, order-preserving list of
// (modifier, name, type). "X, Y: Integer" yields two entries sharing the type.
function ParseParamList(const S: string): TArray<TParamSpec>;
var
  Grp, NamesPart, Typ, Modifier, Low, Nm: string;
  I, Depth, ColonPos: Integer;
begin
  Result := nil;
  for Grp in SplitTopLevel(S, ';') do
  begin
    var G := Trim(Grp);
    if G = '' then Continue;
    // First top-level ':' separates the names from the type.
    ColonPos := 0; Depth := 0;
    for I := 1 to Length(G) do
    begin
      case G[I] of
        '(', '[': Inc(Depth);
        ')', ']': if Depth > 0 then Dec(Depth);
        ':': if Depth = 0 then begin ColonPos := I; Break; end;
      end;
    end;
    if ColonPos > 0 then
    begin
      NamesPart := Trim(Copy(G, 1, ColonPos - 1));
      Typ := Trim(Copy(G, ColonPos + 1, MaxInt));
    end
    else
    begin
      NamesPart := G;      // untyped ("var X") - rare in event signatures
      Typ := '';
    end;
    Modifier := '';
    Low := UpperCase(NamesPart);
    for var Pre in ['VAR ', 'CONST ', 'OUT '] do
      if Low.StartsWith(Pre) then
      begin
        Modifier := Trim(Copy(NamesPart, 1, Length(Pre)));
        NamesPart := Trim(Copy(NamesPart, Length(Pre) + 1, MaxInt));
        Break;
      end;
    for Nm in SplitTopLevel(NamesPart, ',') do
    begin
      var N := Trim(Nm);
      if N = '' then Continue;
      var Spec: TParamSpec;
      Spec.Modifier := Modifier;
      Spec.Name := N;
      Spec.Typ := Typ;
      Result := Result + [Spec];
    end;
  end;
end;

// Build a parameter list that keeps the handler's existing parameter NAMES
// but takes each parameter's modifier and TYPE from the expected signature.
// Falls back to AExpected unchanged when the arities differ (a real change of
// parameter count, where names cannot be mapped positionally).
function MergeParamNames(const ACurrent, AExpected: string): string;
var
  Cur, Exp: TArray<TParamSpec>;
  SB: TStringBuilder;
  I, J: Integer;
  Names: string;
begin
  Result := AExpected;
  Cur := ParseParamList(ACurrent);
  Exp := ParseParamList(AExpected);
  if (Length(Cur) = 0) or (Length(Cur) <> Length(Exp)) then Exit;

  SB := TStringBuilder.Create;
  try
    I := 0;
    while I < Length(Exp) do
    begin
      // Group consecutive parameters that share modifier + type.
      Names := Cur[I].Name;
      J := I + 1;
      while (J < Length(Exp)) and SameText(Exp[J].Modifier, Exp[I].Modifier)
            and SameText(Exp[J].Typ, Exp[I].Typ) do
      begin
        Names := Names + ', ' + Cur[J].Name;
        Inc(J);
      end;
      if SB.Length > 0 then SB.Append('; ');
      if Exp[I].Modifier <> '' then SB.Append(Exp[I].Modifier + ' ');
      SB.Append(Names);
      if Exp[I].Typ <> '' then SB.Append(': ' + Exp[I].Typ);
      I := J;
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function ResolveConditionals(const S: string): string;
var
  I, N: Integer;
  SB: TStringBuilder;
  SkipCount: Integer;
  KeepStack: TStack<Boolean>;
  Inner, InnerU: string;
  J: Integer;
begin
  SB := TStringBuilder.Create;
  KeepStack := TStack<Boolean>.Create;
  try
    SkipCount := 0;
    I := 1; N := Length(S);
    while I <= N do
    begin
      if (S[I] = '(') and (I < N) and (S[I + 1] = '*') then
      begin
        // (* ... *) comment - skip whole thing.
        J := I + 2;
        while (J < N) and not ((S[J] = '*') and (S[J + 1] = ')')) do Inc(J);
        I := J + 2;
        Continue;
      end;
      if S[I] = '{' then
      begin
        J := I + 1;
        while (J <= N) and (S[J] <> '}') do Inc(J);
        Inner := Copy(S, I + 1, J - I - 1);
        InnerU := UpperCase(Trim(Inner));
        if InnerU.StartsWith('$IF') then
          KeepStack.Push(True)                 // first branch: keep
        else if InnerU.StartsWith('$ELSE') or InnerU.StartsWith('$ELSEIF') then
        begin
          if (KeepStack.Count > 0) and KeepStack.Peek then
          begin
            KeepStack.Pop; KeepStack.Push(False); Inc(SkipCount);
          end;
        end
        else if InnerU.StartsWith('$ENDIF') or InnerU.StartsWith('$IFEND') then
        begin
          if KeepStack.Count > 0 then
            if not KeepStack.Pop then Dec(SkipCount);
        end;
        // else: a plain comment or unrelated directive - just drop it.
        I := J + 1;
        Continue;
      end;
      if SkipCount = 0 then
        SB.Append(S[I]);
      Inc(I);
    end;
    Result := SB.ToString;
  finally
    KeepStack.Free;
    SB.Free;
  end;
end;

class function TDfmEventChecker.NormalizeParams(const AParamList: string): string;
// "Sender: TObject; var Key: Word" -> "TObject|var Word"
// "X, Y: Integer" -> "Integer|Integer"
var
  Segs: TArray<string>;
  Seg, Prefix, Names, TypeName: string;
  ColonPos, NameCount, I: Integer;
  Parts: TArray<string>;
  Out_: TStringBuilder;
begin
  Out_ := TStringBuilder.Create;
  try
    Segs := ResolveConditionals(AParamList).Split([';']);
    for Seg in Segs do
    begin
      var T := Trim(Seg);
      if T = '' then Continue;
      Prefix := '';
      if StartsText('var ', T) then begin Prefix := 'var '; Delete(T, 1, 4); end
      else if StartsText('out ', T) then begin Prefix := 'out '; Delete(T, 1, 4); end
      else if StartsText('const ', T) then Delete(T, 1, 6);  // const dropped
      ColonPos := Pos(':', T);
      if ColonPos = 0 then
      begin
        // Untyped parameter (var X) - keep the prefix, type '?'.
        TypeName := '?';
        Names := Trim(T);
      end
      else
      begin
        Names := Trim(Copy(T, 1, ColonPos - 1));
        TypeName := Trim(Copy(T, ColonPos + 1, MaxInt));
        // Strip default values ("= nil") and unit prefixes
        // ("Vcl.Controls.TMouseButton" -> "TMouseButton").
        var EqPos := Pos('=', TypeName);
        if EqPos > 0 then TypeName := Trim(Copy(TypeName, 1, EqPos - 1));
        var DotPos := LastDelimiter('.', TypeName);
        if DotPos > 0 then TypeName := Copy(TypeName, DotPos + 1, MaxInt);
      end;
      Parts := Names.Split([',']);
      NameCount := 0;
      for I := 0 to High(Parts) do
        if Trim(Parts[I]) <> '' then Inc(NameCount);
      if NameCount = 0 then NameCount := 1;
      for I := 1 to NameCount do
      begin
        if Out_.Length > 0 then Out_.Append('|');
        Out_.Append(Prefix).Append(TypeName);
      end;
    end;
    Result := Out_.ToString;
  finally
    Out_.Free;
  end;
end;

/// <summary>Parses a class declaration line in any common spelling:
///  "TFoo = class(TBar)", "type TFoo = class", tabs / multiple spaces
///  around '='. Rejects forward declarations ("= class;") and
///  metaclasses ("= class of ..."). AAncestor returns the first
///  parenthesised entry, unit prefix stripped ('' if none).</summary>
function TryParseClassDecl(const ALine: string;
  out AName, AAncestor: string): Boolean;
var
  S, Rest, AfterKw: string;
  EqPos, P2: Integer;
begin
  Result := False;
  AName := ''; AAncestor := '';
  S := Trim(ALine.Replace(#9, ' '));
  // Optional leading "type" keyword on the same line
  // ("type  TFrmLDTexport = class(TPsyPraxForm)").
  if StartsText('type ', S) then
    S := Trim(Copy(S, 6, MaxInt));
  EqPos := Pos('=', S);
  if EqPos = 0 then Exit;
  AName := Trim(Copy(S, 1, EqPos - 1));
  if not IsIdent(AName) then Exit;
  Rest := Trim(Copy(S, EqPos + 1, MaxInt));
  if not StartsText('class', Rest) then Exit;
  AfterKw := Copy(Rest, 6, MaxInt);
  // Word boundary: "classes" etc. must not match.
  if (AfterKw <> '') and CharInSet(AfterKw[1], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then Exit;
  AfterKw := Trim(AfterKw);
  if AfterKw.StartsWith(';') then Exit;              // forward decl
  if StartsText('of ', AfterKw) or SameText(AfterKw, 'of') then Exit;  // metaclass
  // Optional class modifiers before the ancestor: "class abstract(...)",
  // "class sealed(...)". Skip them so the ancestor is still found.
  if StartsText('abstract', AfterKw) then AfterKw := Trim(Copy(AfterKw, 9, MaxInt))
  else if StartsText('sealed', AfterKw) then AfterKw := Trim(Copy(AfterKw, 7, MaxInt));
  // "class helper for TFoo" - not an ancestor relationship we track.
  if StartsText('helper', AfterKw) then Exit;
  if AfterKw.StartsWith('(') then
  begin
    P2 := Pos(')', AfterKw);
    if P2 > 1 then
    begin
      AAncestor := Trim(Copy(AfterKw, 2, P2 - 2));
      // Multi-interface lists: first entry is the class ancestor.
      var CommaPos := Pos(',', AAncestor);
      if CommaPos > 0 then AAncestor := Trim(Copy(AAncestor, 1, CommaPos - 1));
      var DotPos := LastDelimiter('.', AAncestor);
      if DotPos > 0 then AAncestor := Copy(AAncestor, DotPos + 1, MaxInt);
    end;
  end;
  Result := True;
end;

/// <summary>True once ADecl holds a complete method declaration: if a
///  '(' was opened, the matching ')' and a ';' after it must be
///  present; without parameters a plain ';' suffices. Needed because
///  multi-line declarations contain ';' INSIDE the parameter list
///  ("procedure X(Sender: TObject;" ...) - stopping at the first ';'
///  used to truncate the parameter list to nothing.</summary>
function DeclComplete(const ADecl: string): Boolean;
var
  P1, P2: Integer;
begin
  P1 := Pos('(', ADecl);
  if P1 = 0 then
    Result := Pos(';', ADecl) > 0
  else
  begin
    P2 := PosEx(')', ADecl, P1);
    Result := (P2 > 0) and (PosEx(';', ADecl, P2) > 0);
  end;
end;

class procedure TDfmEventChecker.CollectClassMethods(
  const ALines: TArray<string>; const AClassName: string;
  AMethods: TDictionary<string, TPair<Integer, string>>;
  out AAncestor: string);
// Finds "AClassName = class(...)" and collects every method
// declaration up to the closing "end;". Values are (line, normalized
// param list). AAncestor returns the parenthesised ancestor name (''
// if none / TObject-implied).
var
  I, J, Depth: Integer;
  L, U, Decl, MName, Params: string;
  InClass: Boolean;
  P1, P2: Integer;
  DName, DAnc: string;
begin
  AAncestor := '';
  InClass := False;
  Depth := 0;
  I := 0;
  while I < Length(ALines) do
  begin
    L := ALines[I];
    U := UpperCase(Trim(L));
    if not InClass then
    begin
      if TryParseClassDecl(L, DName, DAnc) and SameText(DName, AClassName) then
      begin
        InClass := True;
        AAncestor := DAnc;
      end;
    end
    else
    begin
      if (U = 'END;') and (Depth = 0) then Exit;
      // Rudimentary nested-structure tracking (records/classes inside).
      if U.EndsWith('= RECORD') or U.EndsWith('= CLASS') then Inc(Depth)
      else if (U = 'END;') and (Depth > 0) then Dec(Depth);

      if (Depth = 0) and (U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ')) then
      begin
        // Concatenate lines until the declaration is complete. NOTE:
        // multi-line declarations contain ';' inside the parameter
        // list, so "first ';'" is NOT a valid stop criterion.
        Decl := StripLineComment(Trim(L));
        J := I;
        while (not DeclComplete(Decl)) and (J + 1 < Length(ALines)) do
        begin
          Inc(J);
          Decl := Decl + ' ' + StripLineComment(Trim(ALines[J]));
        end;
        // Extract name + params.
        var SP := Pos(' ', Decl);
        var Rest := Trim(Copy(Decl, SP + 1, MaxInt));
        P1 := Pos('(', Rest);
        var SemiPos := Pos(';', Rest);
        if (P1 > 0) and ((SemiPos = 0) or (P1 < SemiPos)) then
        begin
          MName := Trim(Copy(Rest, 1, P1 - 1));
          P2 := Pos(')', Rest);
          if P2 > P1 then
            Params := Copy(Rest, P1 + 1, P2 - P1 - 1)
          else
            Params := '';
        end
        else
        begin
          if SemiPos > 0 then MName := Trim(Copy(Rest, 1, SemiPos - 1))
          else MName := Rest;
          // Strip ": ReturnType" from parameterless functions.
          var CPos := Pos(':', MName);
          if CPos > 0 then MName := Trim(Copy(MName, 1, CPos - 1));
          Params := '';
        end;
        if IsIdent(MName) and not AMethods.ContainsKey(UpperCase(MName)) then
          AMethods.Add(UpperCase(MName),
            TPair<Integer, string>.Create(I + 1, NormalizeParams(Params)));
        I := J;
      end;
    end;
    Inc(I);
  end;
end;

constructor TDfmEventChecker.Create;
begin
  inherited Create;
  FLinesCache := TDictionary<string, TArray<string>>.Create;
  FClassIndex := TDictionary<string, string>.Create;
  FAssignments := TList<TEventAssignment>.Create;
  FClassAncestor := TDictionary<string, string>.Create;
  FEventPropType := TDictionary<string, string>.Create;
  FProcSig := TDictionary<string, string>.Create;
  FProcRawParams := TDictionary<string, string>.Create;
  FProcLoc := TDictionary<string, string>.Create;
  FClassAlias := TDictionary<string, string>.Create;
end;

destructor TDfmEventChecker.Destroy;
begin
  FClassAlias.Free;
  FProcLoc.Free;
  FProcRawParams.Free;
  FProcSig.Free;
  FEventPropType.Free;
  FClassAncestor.Free;
  FAssignments.Free;
  FClassIndex.Free;
  FLinesCache.Free;
  inherited;
end;

function TDfmEventChecker.GetLines(const AFile: string): TArray<string>;
var
  Key: string;
begin
  Key := UpperCase(AFile);
  if FLinesCache.TryGetValue(Key, Result) then Exit;
  Result := ReadFileLines(AFile);
  FLinesCache.Add(Key, Result);
end;

procedure TDfmEventChecker.BuildClassIndex(const AFiles: TArray<string>);
// One pass over every project source: register "TFoo = class..." with
// its declaring file. First declaration wins; forwards ("= class;")
// and metaclasses ("= class of") are rejected by TryParseClassDecl.
var
  F, Name, Anc: string;
  Lines: TArray<string>;
  I: Integer;
begin
  for F in AFiles do
  begin
    if not SameText(ExtractFileExt(F), '.pas') then Continue;
    Lines := GetLines(F);
    for I := 0 to High(Lines) do
    begin
      // Cheap pre-filter before the full parse.
      if Pos('CLASS', UpperCase(Lines[I])) = 0 then Continue;
      if TryParseClassDecl(Lines[I], Name, Anc) then
        if not FClassIndex.ContainsKey(UpperCase(Name)) then
          FClassIndex.Add(UpperCase(Name), F);
    end;
  end;
end;

function TDfmEventChecker.CheckPairInternal(const ADfmFile, APasFile: string): TArray<TDfmEventIssue>;
var
  DfmLines, PasLines: TArray<string>;
  Issues: TList<TDfmEventIssue>;
  Methods: TDictionary<string, TPair<Integer, string>>;
  CompStack: TStack<TPair<string, string>>;  // name, type
  FormClass, Ancestor, NextClass, ProbeFile: string;
  I, Hops: Integer;
  Issue: TDfmEventIssue;
begin
  Result := nil;
  DfmLines := GetLines(ADfmFile);
  if Length(DfmLines) = 0 then Exit;
  // Binary DFM guard: text DFMs start with object/inherited/inline.
  var FirstWord := UpperCase(Trim(DfmLines[0]).Split([' '])[0]);
  if (FirstWord <> 'OBJECT') and (FirstWord <> 'INHERITED') and (FirstWord <> 'INLINE') then
    Exit;

  // Root line: "object Form1: TForm1"
  var RootParts := Trim(DfmLines[0]).Split([':']);
  if Length(RootParts) < 2 then Exit;
  FormClass := Trim(RootParts[1]);

  PasLines := GetLines(APasFile);
  if Length(PasLines) = 0 then Exit;

  Methods := TDictionary<string, TPair<Integer, string>>.Create;
  Issues := TList<TDfmEventIssue>.Create;
  CompStack := TStack<TPair<string, string>>.Create;
  try
    // Collect methods of the form class + project-resolvable ancestors
    // (visual form inheritance keeps handlers in the base class). The
    // class index makes each hop a dictionary lookup; ancestors outside
    // the project (TForm, TFrame, ...) are simply not in the index and
    // end the walk.
    CollectClassMethods(PasLines, FormClass, Methods, Ancestor);
    Hops := 0;
    NextClass := Ancestor;
    while (NextClass <> '') and (Hops < 8) do
    begin
      Inc(Hops);
      if not FClassIndex.TryGetValue(UpperCase(NextClass), ProbeFile) then Break;
      var NextAncestor: string;
      CollectClassMethods(GetLines(ProbeFile), NextClass, Methods, NextAncestor);
      NextClass := NextAncestor;
    end;

    // Walk the DFM: track the object stack, find "OnXxx = Handler".
    CompStack.Push(TPair<string, string>.Create(
      Trim(RootParts[0]).Split([' '])[High(Trim(RootParts[0]).Split([' ']))], FormClass));
    I := 1;
    while I < Length(DfmLines) do
    begin
      var L := Trim(DfmLines[I]);
      var U := UpperCase(L);
      if U.StartsWith('OBJECT ') or U.StartsWith('INHERITED ') or U.StartsWith('INLINE ') then
      begin
        var Parts := L.Split([' ', ':'], TStringSplitOptions.ExcludeEmpty);
        if Length(Parts) >= 3 then
          CompStack.Push(TPair<string, string>.Create(Parts[1], Parts[2]))
        else if Length(Parts) >= 2 then
          CompStack.Push(TPair<string, string>.Create(Parts[1], ''));
      end
      else if U = 'END' then
      begin
        if CompStack.Count > 1 then CompStack.Pop;
      end
      else
      begin
        // Multi-line property values: skip blocks opened by trailing
        // delimiters. Simplified but effective for real DFMs.
        if U.EndsWith('= {') or U.EndsWith('{') then
        begin
          while (I + 1 < Length(DfmLines)) and not Trim(DfmLines[I]).EndsWith('}') do Inc(I);
        end
        else if U.EndsWith('= (') then
        begin
          while (I + 1 < Length(DfmLines)) and not Trim(DfmLines[I]).EndsWith(')') do Inc(I);
        end
        else if U.EndsWith('= <') then
        begin
          // Collection property - items may contain events too, but
          // their handlers follow the same rules; keep scanning inside.
          // (No skip: item> lines parse as normal properties.)
        end
        else
        begin
          var EqPos := Pos(' = ', L);
          if EqPos > 0 then
          begin
            var PropName := Trim(Copy(L, 1, EqPos - 1));
            var PropValue := Trim(Copy(L, EqPos + 3, MaxInt));
            // Event heuristic: last dotted segment starts with 'On',
            // value is a bare identifier (no quotes, no digits-only,
            // not True/False/nil, no dots).
            var DotPos := LastDelimiter('.', PropName);
            if DotPos > 0 then PropName := Copy(PropName, DotPos + 1, MaxInt);
            if StartsText('On', PropName) and IsIdent(PropValue)
               and not SameText(PropValue, 'True') and not SameText(PropValue, 'False')
               and not SameText(PropValue, 'nil') then
            begin
              var MethodInfo: TPair<Integer, string>;
              if not Methods.TryGetValue(UpperCase(PropValue), MethodInfo) then
              begin
                Issue := Default(TDfmEventIssue);
                Issue.DfmFile := ADfmFile;
                Issue.PasFile := APasFile;
                if CompStack.Count > 0 then
                begin
                  Issue.ComponentName := CompStack.Peek.Key;
                  Issue.ComponentType := CompStack.Peek.Value;
                end;
                Issue.EventName := PropName;
                Issue.HandlerName := PropValue;
                Issue.DfmLine := I + 1;
                Issue.Kind := eikMissingHandler;
                Issues.Add(Issue);
              end
              else
              begin
                // Signature checking is only sound where we KNOW the
                // event type: on the form itself (stack depth 1) or on
                // whitelisted standard classes. Other components may
                // declare same-named events with different signatures.
                var SigCheckable := CompStack.Count <= 1;
                if not SigCheckable then
                  for var SafeIdx := Low(SafeVclClasses) to High(SafeVclClasses) do
                    if SameText(SafeVclClasses[SafeIdx], CompStack.Peek.Value) then
                    begin
                      SigCheckable := True;
                      Break;
                    end;
                var Expected: string;
                var HasTableSig := ExpectedSignature(PropName, Expected);
                if SigCheckable and HasTableSig
                   and not SameText(MethodInfo.Value, Expected) then
                begin
                  Issue := Default(TDfmEventIssue);
                  Issue.DfmFile := ADfmFile;
                  Issue.PasFile := APasFile;
                  if CompStack.Count > 0 then
                  begin
                    Issue.ComponentName := CompStack.Peek.Key;
                    Issue.ComponentType := CompStack.Peek.Value;
                  end;
                  Issue.EventName := PropName;
                  Issue.HandlerName := PropValue;
                  Issue.DfmLine := I + 1;
                  Issue.PasLine := MethodInfo.Key;
                  Issue.Kind := eikSignatureMismatch;
                  Issue.Expected := '(' + Expected.Replace('|', '; ') + ')';
                  Issue.Actual := '(' + MethodInfo.Value.Replace('|', '; ') + ')';
                  Issue.ExpectedNorm := Expected;
                  Issue.ActualNorm := MethodInfo.Value;
                  Issues.Add(Issue);
                end
                else if not HasTableSig and (CompStack.Count > 0)
                        and (CompStack.Peek.Value <> '') then
                begin
                  // No built-in table signature (third-party event).
                  // First try to resolve it EXACTLY from the component's
                  // own source (indexed from the search path). If that
                  // works, check against the real signature; otherwise
                  // pool it for the cross-project consistency check.
                  var ResolvedSig, TypeName, RawParams: string;
                  if TryResolveEventSignature(CompStack.Peek.Value, PropName,
                       ResolvedSig, TypeName, RawParams) then
                  begin
                    if not SameText(MethodInfo.Value, ResolvedSig) then
                    begin
                      Issue := Default(TDfmEventIssue);
                      Issue.DfmFile := ADfmFile;
                      Issue.PasFile := APasFile;
                      Issue.ComponentName := CompStack.Peek.Key;
                      Issue.ComponentType := CompStack.Peek.Value;
                      Issue.EventName := PropName;
                      Issue.HandlerName := PropValue;
                      Issue.DfmLine := I + 1;
                      Issue.PasLine := MethodInfo.Key;
                      Issue.Kind := eikSignatureMismatch;
                      Issue.Expected := Format('(%s)  [%s]',
                        [ResolvedSig.Replace('|', '; '), TypeName]);
                      Issue.Actual := '(' + MethodInfo.Value.Replace('|', '; ') + ')';
                      Issue.ExpectedNorm := ResolvedSig;
                      Issue.ActualNorm := MethodInfo.Value;
                      Issue.ExpectedRawParams := RawParams;   // enables auto-fix
                      Issue.FormClass := FormClass;
                      Issue.EventTypeName := TypeName;
                      var Loc: string;
                      if FProcLoc.TryGetValue(UpperCase(TypeName), Loc) then
                      begin
                        var Bar := Pos('|', Loc);
                        if Bar > 0 then
                        begin
                          Issue.EventTypeFile := Copy(Loc, 1, Bar - 1);
                          Issue.EventTypeLine := StrToIntDef(Copy(Loc, Bar + 1, MaxInt), 0);
                        end;
                      end;
                      Issues.Add(Issue);
                    end;
                  end
                  else
                  begin
                    var A: TEventAssignment;
                    A.CompType := CompStack.Peek.Value;
                    A.EventName := PropName;
                    A.HandlerName := PropValue;
                    A.Signature := MethodInfo.Value;
                    A.ComponentName := CompStack.Peek.Key;
                    A.DfmFile := ADfmFile;
                    A.PasFile := APasFile;
                    A.DfmLine := I + 1;
                    A.PasLine := MethodInfo.Key;
                    FAssignments.Add(A);
                  end;
                end;
              end;
            end;
          end;
        end;
      end;
      Inc(I);
    end;

    Result := Issues.ToArray;
  finally
    CompStack.Free;
    Issues.Free;
    Methods.Free;
  end;
end;

class function TDfmEventChecker.CheckProject(
  const AProjectFiles: TArray<string>;
  const AProgress: TProc<Integer, Integer, string>;
  const ASignatureFiles: TArray<string>): TArray<TDfmEventIssue>;
var
  Checker: TDfmEventChecker;
  All: TList<TDfmEventIssue>;
  Pairs: TList<string>;
  F, DfmFile: string;
  I, Total: Integer;
begin
  Checker := TDfmEventChecker.Create;
  All := TList<TDfmEventIssue>.Create;
  Pairs := TList<string>.Create;
  try
    for F in AProjectFiles do
    begin
      if not SameText(ExtractFileExt(F), '.pas') then Continue;
      if TFile.Exists(ChangeFileExt(F, '.dfm')) then
        Pairs.Add(F);
    end;

    // Progress spans both phases: indexing counts as one big step
    // (it reads every file once), then one step per form.
    Total := Pairs.Count + 1;
    if Assigned(AProgress) then
      AProgress(1, Total, 'Indexing project classes...');
    Checker.BuildClassIndex(AProjectFiles);

    // Signature index over component source reachable via the search
    // path (project files + ASignatureFiles). Enables exact checks for
    // third-party events (VirtualTrees, ...) read from their real .pas.
    Checker.BuildSignatureIndex(AProjectFiles,
      procedure(ACur, ATot: Integer; AFile: string)
      begin
        if Assigned(AProgress) then
          AProgress(ACur, ATot, 'Indexing project source: ' + ExtractFileName(AFile));
      end);
    if Length(ASignatureFiles) > 0 then
      Checker.BuildSignatureIndex(ASignatureFiles,
        procedure(ACur, ATot: Integer; AFile: string)
        begin
          if Assigned(AProgress) then
            AProgress(ACur, ATot, 'Indexing signatures: ' + ExtractFileName(AFile));
        end);

    for I := 0 to Pairs.Count - 1 do
    begin
      DfmFile := ChangeFileExt(Pairs[I], '.dfm');
      if Assigned(AProgress) then
        AProgress(I + 1, Pairs.Count, 'Checking forms: ' + ExtractFileName(Pairs[I]));
      All.AddRange(Checker.CheckPairInternal(DfmFile, Pairs[I]));
    end;

    // Cross-project consistency: for events with no known signature
    // (third-party components), the odd-one-out handler is very likely
    // a copy/paste mistake.
    if Assigned(AProgress) then
      AProgress(1, 1, 'Analyzing cross-project consistency...');
    All.AddRange(Checker.ConsistencyIssues);

    if Assigned(AProgress) then
      AProgress(1, 1, 'Building results...');

    Result := All.ToArray;
  finally
    Pairs.Free;
    All.Free;
    Checker.Free;
  end;
end;

procedure TDfmEventChecker.BuildSignatureIndex(const AFiles: TArray<string>;
  const AProgress: TProc<Integer, Integer, string>);

  function ExtractProcParams(const ALines: TArray<string>; AStart: Integer): string;
  // Gathers a "TType = procedure(...) of object" declaration (possibly
  // multi-line) starting at line AStart and returns the raw parameter
  // list between the outer ( ). '' for parameterless.
  var
    Decl: string;
    J, P1, P2: Integer;
  begin
    Result := '';
    Decl := StripLineComment(Trim(ALines[AStart]));
    J := AStart;
    // Gather until we have a matching ')' - but never more than a few
    // lines. Without this cap, a stray '(' (e.g. in a comment that got
    // past detection) would concatenate huge portions of the file into
    // one string, spiking memory.
    while (J + 1 < Length(ALines)) and (J - AStart < 12)
          and (Pos('(', Decl) > 0) and (PosEx(')', Decl, Pos('(', Decl)) = 0) do
    begin
      Inc(J);
      Decl := Decl + ' ' + StripLineComment(Trim(ALines[J]));
    end;
    P1 := Pos('(', Decl);
    if P1 = 0 then Exit;  // parameterless
    P2 := PosEx(')', Decl, P1);
    if P2 > P1 then
      Result := Copy(Decl, P1 + 1, P2 - P1 - 1);
  end;

var
  F, MethodOwnerClass, U, L, T: string;
  Lines: TArray<string>;
  I, Depth, ColonPos, K: Integer;
  InClass: Boolean;
  DName, DAnc: string;
begin
  var FileIdx := 0;
  for F in AFiles do
  begin
    Inc(FileIdx);
    if not SameText(ExtractFileExt(F), '.pas') then Continue;
    if Assigned(AProgress) and ((FileIdx mod 25 = 0) or (FileIdx = Length(AFiles))) then
      AProgress(FileIdx, Length(AFiles), F);
    // Read WITHOUT caching: the signature index scans thousands of
    // large library .pas (FastReport, TRichView, FHIR, ...). Caching
    // them all would exhaust the 32-bit IDE process's address space.
    // Reuse the cache only if the file is already there (project
    // files, cached by BuildClassIndex); otherwise read + discard.
    if not FLinesCache.TryGetValue(UpperCase(F), Lines) then
      Lines := ReadFileLines(F);
    InClass := False;
    Depth := 0;
    MethodOwnerClass := '';
    // Defensive: this scans arbitrary third-party source; a single
    // malformed file must never take down the host IDE.
    try
    for I := 0 to High(Lines) do
    begin
      L := Lines[I];
      U := UpperCase(Trim(L));
      // Skip full-line comments: raw lines are scanned without comment
      // awareness, so a "// TFoo = procedure(" would otherwise be
      // mis-detected as a real procedural type.
      if U.StartsWith('//') then Continue;
      if not InClass then
      begin
        if TryParseClassDecl(L, DName, DAnc) then
        begin
          InClass := True; Depth := 0;
          MethodOwnerClass := DName;
          if (DAnc <> '') and not FClassAncestor.ContainsKey(UpperCase(DName)) then
            FClassAncestor.Add(UpperCase(DName), DAnc);
        end
        else if (Pos(' = PROCEDURE', U) > 0) or (Pos(' = FUNCTION', U) > 0) then
        begin
          // Procedural TYPE (event type): "TXxx = procedure(...) of object;"
          if (Pos(' OF OBJECT', U) > 0) or (Pos('(', U) > 0) then
          begin
            var EqP := Pos(' = ', L);
            if EqP > 0 then
            begin
              T := Trim(Copy(L, 1, EqP - 1));
              if (T <> '') and CharInSet(T[1], ['A'..'Z', 'a'..'z', '_'])
                 and not FProcSig.ContainsKey(UpperCase(T)) then
              begin
                var RawP := ExtractProcParams(Lines, I);
                FProcSig.Add(UpperCase(T), NormalizeParams(RawP));
                FProcRawParams.Add(UpperCase(T), Trim(RawP));
                FProcLoc.Add(UpperCase(T), F + '|' + IntToStr(I + 1));
              end;
            end;
          end;
        end
        else
        begin
          // Simple type alias "TFoo = TBar;" (single identifier RHS).
          var EqP := Pos('=', L);
          if EqP > 0 then
          begin
            var Lhs := Trim(Copy(L, 1, EqP - 1));
            var Rhs := Trim(Copy(L, EqP + 1, MaxInt));
            if Rhs.EndsWith(';') then Rhs := Trim(Copy(Rhs, 1, Length(Rhs) - 1));
            if IsIdent(Lhs) and IsIdent(Rhs) then
            begin
              var DotP := LastDelimiter('.', Rhs);
              if DotP > 0 then Rhs := Copy(Rhs, DotP + 1, MaxInt);
              if not FClassAlias.ContainsKey(UpperCase(Lhs)) then
                FClassAlias.Add(UpperCase(Lhs), Rhs);
            end;
          end;
        end;
      end
      else
      begin
        if (U = 'END;') and (Depth = 0) then
        begin
          InClass := False;
          Continue;
        end;
        if U.EndsWith('= RECORD') or U.EndsWith('= CLASS') then Inc(Depth)
        else if (U = 'END;') and (Depth > 0) then Dec(Depth);
        if (Depth = 0) and U.StartsWith('PROPERTY ') then
        begin
          // "property OnXxx: TType read... write...;"
          // Use the TRIMMED line - property lines are indented, so
          // slicing the raw L at a fixed offset would be off by the
          // indentation.
          T := Trim(Copy(Trim(L), Length('property ') + 1, MaxInt));
          ColonPos := Pos(':', T);
          if ColonPos > 0 then
          begin
            var PName := Trim(Copy(T, 1, ColonPos - 1));
            // reject indexed/array properties "property Items[..]"
            if (Pos('[', PName) = 0) and StartsText('On', PName) then
            begin
              var Rest := Trim(Copy(T, ColonPos + 1, MaxInt));
              // type = first identifier token
              K := 1;
              while (K <= Length(Rest)) and
                    CharInSet(Rest[K], ['A'..'Z', 'a'..'z', '0'..'9', '_', '.']) do
                Inc(K);
              var TypeName := Copy(Rest, 1, K - 1);
              var DotPos := LastDelimiter('.', TypeName);
              if DotPos > 0 then TypeName := Copy(TypeName, DotPos + 1, MaxInt);
              if (TypeName <> '') and (MethodOwnerClass <> '') then
              begin
                var Key := UpperCase(MethodOwnerClass) + '.' + UpperCase(PName);
                if not FEventPropType.ContainsKey(Key) then
                  FEventPropType.Add(Key, TypeName);
              end;
            end;
          end;
        end;
      end;
    end;
    except
      // A single malformed third-party file must not abort the run
      // (or take down the host IDE).
      on E: Exception do ;
    end;
  end;
end;

function TDfmEventChecker.TryResolveEventSignature(
  const ACompType, AEventName: string;
  out ASignature, ATypeName, ARawParams: string): Boolean;
var
  Cls, TypeName, Sig: string;
  Hops: Integer;
begin
  Result := False;
  ASignature := ''; ATypeName := ''; ARawParams := '';
  Cls := ACompType;
  Hops := 0;
  while (Cls <> '') and (Hops < 40) do
  begin
    Inc(Hops);
    if FEventPropType.TryGetValue(UpperCase(Cls) + '.' + UpperCase(AEventName), TypeName) then
    begin
      ATypeName := TypeName;
      if FProcSig.TryGetValue(UpperCase(TypeName), Sig) then
      begin
        ASignature := Sig;
        FProcRawParams.TryGetValue(UpperCase(TypeName), ARawParams);
        Exit(True);
      end;
      // Property found but its type isn't a resolvable proc type
      // (declared elsewhere / not indexed). Give up - cannot verify.
      Exit(False);
    end;
    // Next hop: real ancestor first; if none, follow a type alias
    // (TVTAncestor = TVTAncestorVcl); else the chain ends.
    var Next: string;
    if FClassAncestor.TryGetValue(UpperCase(Cls), Next) then
      Cls := Next
    else if FClassAlias.TryGetValue(UpperCase(Cls), Next) then
      Cls := Next
    else
      Break;
  end;
end;

function TDfmEventChecker.ConsistencyIssues: TArray<TDfmEventIssue>;
// Groups all pooled assignments by (component type + event name). In a
// group where one signature is clearly the majority (>= 2 handlers
// agree AND it strictly outnumbers every other), the handlers that
// deviate are flagged. Requires >= 3 handlers, so a lone pair (1 vs 1,
// ambiguous) is never flagged - keeps false positives out.
var
  Groups: TDictionary<string, TList<Integer>>;
  I, K: Integer;
  Key: string;
  A: TEventAssignment;
  Issues: TList<TDfmEventIssue>;
  Pair: TPair<string, TList<Integer>>;
  SigCount: TDictionary<string, Integer>;
  MajSig: string;
  MajCount, SecondCount, C: Integer;
  Issue: TDfmEventIssue;
begin
  Groups := TDictionary<string, TList<Integer>>.Create;
  Issues := TList<TDfmEventIssue>.Create;
  try
    for I := 0 to FAssignments.Count - 1 do
    begin
      A := FAssignments[I];
      Key := UpperCase(A.CompType) + '.' + UpperCase(A.EventName);
      if not Groups.ContainsKey(Key) then
        Groups.Add(Key, TList<Integer>.Create);
      Groups[Key].Add(I);
    end;

    for Pair in Groups do
    begin
      if Pair.Value.Count < 3 then Continue;
      // Count signatures.
      SigCount := TDictionary<string, Integer>.Create;
      try
        for K in Pair.Value do
        begin
          SigCount.TryGetValue(FAssignments[K].Signature, C);
          SigCount.AddOrSetValue(FAssignments[K].Signature, C + 1);
        end;
        // Find top and second-best counts.
        MajSig := ''; MajCount := 0; SecondCount := 0;
        for var SP in SigCount do
          if SP.Value > MajCount then
          begin
            SecondCount := MajCount;
            MajCount := SP.Value;
            MajSig := SP.Key;
          end
          else if SP.Value > SecondCount then
            SecondCount := SP.Value;

        // Clear majority required: >= 2 agree and strictly outnumber.
        if (MajCount < 2) or (MajCount <= SecondCount) then Continue;

        for K in Pair.Value do
          if not SameText(FAssignments[K].Signature, MajSig) then
          begin
            A := FAssignments[K];
            Issue := Default(TDfmEventIssue);
            Issue.DfmFile := A.DfmFile;
            Issue.PasFile := A.PasFile;
            Issue.ComponentName := A.ComponentName;
            Issue.ComponentType := A.CompType;
            Issue.EventName := A.EventName;
            Issue.HandlerName := A.HandlerName;
            Issue.DfmLine := A.DfmLine;
            Issue.PasLine := A.PasLine;
            Issue.Kind := eikSignatureMismatch;
            Issue.Expected := Format('(%s)  [%d other handler(s) agree]',
              [MajSig.Replace('|', '; '), MajCount]);
            Issue.Actual := '(' + A.Signature.Replace('|', '; ') + ')';
            Issue.ExpectedNorm := MajSig;
            Issue.ActualNorm := A.Signature;
            Issues.Add(Issue);
          end;
      finally
        SigCount.Free;
      end;
    end;

    Result := Issues.ToArray;
  finally
    for Pair in Groups do
      Pair.Value.Free;
    Groups.Free;
    Issues.Free;
  end;
end;

class function TDfmEventChecker.ApplyFix(const AIssue: TDfmEventIssue;
  out AFailReason: string; AContext: TFixContext): Boolean;
var
  Content: string;
  Lines: TStringList;

  // A header is complete once its parameter list is closed and a ';'
  // follows - NOT at the first ';', which for multi-parameter headers
  // sits INSIDE the parameter list ("Sender: TObject; ACol: Integer").
  function HeaderComplete(const S: string): Boolean;
  var
    P1, Depth, K, PC: Integer;
  begin
    P1 := Pos('(', S);
    if P1 = 0 then Exit(Pos(';', S) > 0);   // parameterless: any ';' closes it
    Depth := 0; PC := 0;
    for K := P1 to Length(S) do
      if S[K] = '(' then Inc(Depth)
      else if S[K] = ')' then
      begin
        Dec(Depth);
        if Depth = 0 then begin PC := K; Break; end;
      end;
    if PC = 0 then Exit(False);            // '(' not yet closed
    Result := PosEx(';', S, PC) > 0;       // ';' after the closing ')'
  end;

  // Rewrites the method header starting at line index AIdx so its
  // parameter list becomes (AParams). Multi-line headers are collapsed
  // onto the start line. Returns True if a change was made.
  function RewriteHeaderAt(AIdx: Integer): Boolean;
  var
    Raw, Indent, Before, After, NewHeader, NewParams: string;
    J, P1, PC, Depth, SemiPos, K: Integer;
  begin
    Result := False;
    if (AIdx < 0) or (AIdx >= Lines.Count) then Exit;
    // Gather the full header (until the param list is closed and a ';'
    // follows), max 12 lines.
    Raw := Lines[AIdx];
    J := AIdx;
    while (J + 1 < Lines.Count) and (J - AIdx < 12) and
          (not HeaderComplete(Raw)) do
    begin
      Inc(J);
      Raw := Raw + ' ' + Trim(Lines[J]);
    end;
    // Preserve the original indentation of the start line.
    Indent := '';
    K := 1;
    while (K <= Length(Lines[AIdx])) and CharInSet(Lines[AIdx][K], [' ', #9]) do
    begin Indent := Indent + Lines[AIdx][K]; Inc(K); end;

    P1 := Pos('(', Raw);
    SemiPos := Pos(';', Raw);
    if (P1 > 0) and ((SemiPos = 0) or (P1 < SemiPos)) then
    begin
      // Find matching ')'.
      Depth := 0; PC := 0;
      for K := P1 to Length(Raw) do
      begin
        if Raw[K] = '(' then Inc(Depth)
        else if Raw[K] = ')' then
        begin
          Dec(Depth);
          if Depth = 0 then begin PC := K; Break; end;
        end;
      end;
      if PC = 0 then Exit;
      Before := Copy(Raw, 1, P1 - 1);
      After := Copy(Raw, PC + 1, MaxInt);
      // Keep the handler's own parameter NAMES; only correct the types.
      NewParams := MergeParamNames(Copy(Raw, P1 + 1, PC - P1 - 1),
                                   AIssue.ExpectedRawParams);
    end
    else
    begin
      // No parameter list yet - insert one before the ';'.
      if SemiPos = 0 then Exit;
      Before := Copy(Raw, 1, SemiPos - 1);
      After := Copy(Raw, SemiPos, MaxInt);
      NewParams := AIssue.ExpectedRawParams;   // nothing to preserve
    end;

    if Trim(AIssue.ExpectedRawParams) = '' then
      NewHeader := Trim(Before) + After
    else
      NewHeader := Trim(Before) + '(' + NewParams + ')' + After;

    // Replace the [AIdx..J] span with one rewritten, re-indented line.
    Lines[AIdx] := Indent + NewHeader;
    for K := J downto AIdx + 1 do
      Lines.Delete(K);
    Result := True;
  end;

  function FindImplIndex: Integer;
  // Locate the method implementation header "procedure <Class>.<Handler>(".
  // The class qualifier is preferred when known, but not required: the
  // form class captured from the DFM can be empty or differ from the
  // implementing class (inherited forms, frames), so falling back to the
  // handler name alone (with a word boundary) still finds the right one.
  var
    I, P, AfterIdx: Integer;
    U, HN: string;
    Fallback: Integer;
  begin
    Result := -1;
    Fallback := -1;
    HN := UpperCase(AIssue.HandlerName);
    if HN = '' then Exit;
    for I := 0 to Lines.Count - 1 do
    begin
      U := UpperCase(TrimLeft(Lines[I]));
      // Only implementation headers are class-qualified; the in-class
      // declaration has no '.' and must not be matched here.
      if not (U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ')) then Continue;
      U := UpperCase(Lines[I]);
      P := Pos('.' + HN, U);
      if P = 0 then Continue;
      // Word boundary: the char after the handler name must not continue
      // an identifier (so "OnClick" never matches "OnClickExtra").
      AfterIdx := P + 1 + Length(HN);
      if (AfterIdx <= Length(U)) and
         CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_']) then Continue;
      if (AIssue.FormClass <> '') and
         (Pos(UpperCase(AIssue.FormClass) + '.' + HN, U) > 0) then
        Exit(I);   // exact class match - best
      if Fallback < 0 then Fallback := I;
    end;
    Result := Fallback;
  end;

  function FindDeclIndex: Integer;
  // Locate the in-class DECLARATION header "procedure <Handler>(" - the one
  // that is NOT class-qualified (no '.'). Anchored by name, not by line
  // number: correcting an earlier handler in the same file collapses its
  // multi-line header and shifts every later line, so a stored line index
  // would be stale by the time this issue is fixed.
  var
    I, CutPos, K: Integer;
    U, HN, Rest, NamePart: string;
  begin
    Result := -1;
    HN := AIssue.HandlerName;
    if HN = '' then Exit;
    for I := 0 to Lines.Count - 1 do
    begin
      U := TrimLeft(Lines[I]);
      if UpperCase(U).StartsWith('PROCEDURE ') then
        Rest := Trim(Copy(U, Length('procedure ') + 1, MaxInt))
      else if UpperCase(U).StartsWith('FUNCTION ') then
        Rest := Trim(Copy(U, Length('function ') + 1, MaxInt))
      else
        Continue;
      // Name is everything up to the first '(' / ':' / ';'.
      CutPos := Length(Rest) + 1;
      for K := 1 to Length(Rest) do
        if CharInSet(Rest[K], ['(', ':', ';', ' ', #9]) then begin CutPos := K; Break; end;
      NamePart := Trim(Copy(Rest, 1, CutPos - 1));
      // In-class declaration only: qualified (T.Method) headers are the
      // implementation and are handled by FindImplIndex.
      if (Pos('.', NamePart) = 0) and SameText(NamePart, HN) then Exit(I);
    end;
  end;

  // True if AUnit already appears in any uses clause of Lines.
  function UnitInUses(const AUnit: string): Boolean;
  var
    I: Integer;
    InUses: Boolean;
    L, Low: string;
    Tok: string;
  begin
    Result := False;
    if AUnit = '' then Exit;
    InUses := False;
    for I := 0 to Lines.Count - 1 do
    begin
      L := StripLineComment(Lines[I]);
      Low := LowerCase(Trim(L));
      if not InUses then
      begin
        if (Low = 'uses') or Low.StartsWith('uses ') or Low.StartsWith('uses'#9) then
          InUses := True
        else
          Continue;
      end;
      for Tok in L.Split([',', ' ', #9, ';', '(', ')']) do
        if SameText(Trim(Tok), AUnit) then Exit(True);
      if Pos(';', L) > 0 then InUses := False;
    end;
  end;

  // Add AUnit to the interface uses clause (the class declaration lives in
  // the interface, so a parameter type must be visible there). Creates a
  // uses clause if the interface has none.
  procedure EnsureInterfaceUses(const AUnit: string);
  var
    I, IntfIdx, ImplIdx2, UsesIdx, P: Integer;
    Low: string;
  begin
    if (AUnit = '') or UnitInUses(AUnit) then Exit;
    IntfIdx := -1; ImplIdx2 := -1;
    for I := 0 to Lines.Count - 1 do
    begin
      Low := LowerCase(Trim(StripLineComment(Lines[I])));
      if (IntfIdx < 0) and (Low = 'interface') then IntfIdx := I
      else if Low = 'implementation' then begin ImplIdx2 := I; Break; end;
    end;
    if IntfIdx < 0 then Exit;
    if ImplIdx2 < 0 then ImplIdx2 := Lines.Count;
    // Find a uses clause inside the interface section.
    UsesIdx := -1;
    for I := IntfIdx + 1 to ImplIdx2 - 1 do
    begin
      Low := LowerCase(Trim(StripLineComment(Lines[I])));
      if (Low = 'uses') or Low.StartsWith('uses ') or Low.StartsWith('uses'#9) then
      begin UsesIdx := I; Break; end;
    end;
    if UsesIdx >= 0 then
    begin
      // Insert ", AUnit" before the ';' that closes the clause.
      for I := UsesIdx to ImplIdx2 - 1 do
      begin
        P := Pos(';', StripLineComment(Lines[I]));
        if P > 0 then
        begin
          P := Pos(';', Lines[I]);
          Lines[I] := Copy(Lines[I], 1, P - 1) + ', ' + AUnit +
                      Copy(Lines[I], P, MaxInt);
          Exit;
        end;
      end;
    end
    else
      Lines.Insert(IntfIdx + 1, 'uses ' + AUnit + ';');
  end;

var
  Changed: Boolean;
  ImplIdx, DeclIdx: Integer;
  TargetUnit, U: string;
  Needed: TStringList;
begin
  Result := False;
  AFailReason := '';
  if AIssue.ExpectedRawParams = '' then
  begin AFailReason := 'not auto-fixable (no expected params)'; Exit; end;
  if (Editor = nil) then
  begin AFailReason := 'no editor'; Exit; end;
  if not Editor.ReadEditorContent(AIssue.PasFile, Content) then
  begin
    if not TFile.Exists(AIssue.PasFile) then
    begin AFailReason := 'file not found: ' + AIssue.PasFile; Exit; end;
    try Content := TFile.ReadAllText(AIssue.PasFile);
    except AFailReason := 'read error: ' + AIssue.PasFile; Exit; end;
  end;

  Lines := TStringList.Create;
  try
    Lines.Text := Content;
    // Both the declaration AND the implementation must be rewritten, or
    // the two signatures would disagree and the unit would not compile.
    // If we cannot locate the implementation, change nothing.
    ImplIdx := FindImplIndex;
    if ImplIdx < 0 then
    begin
      AFailReason := Format('implementation "%s" not found in %s',
        [AIssue.HandlerName, ExtractFileName(AIssue.PasFile)]);
      Exit;
    end;
    // Implementation FIRST (it sits below the declaration, so collapsing
    // its continuation lines does not shift the declaration above it).
    Changed := RewriteHeaderAt(ImplIdx);
    // Then find the DECLARATION FRESH by name - never by a stored line
    // number: fixing earlier handlers in this same file already shifted
    // the line numbers.
    DeclIdx := FindDeclIndex;
    if DeclIdx < 0 then
    begin
      AFailReason := Format('declaration "%s" not found in %s',
        [AIssue.HandlerName, ExtractFileName(AIssue.PasFile)]);
      Exit;
    end;
    if RewriteHeaderAt(DeclIdx) then Changed := True;

    if not Changed then
    begin AFailReason := 'no header rewritten'; Exit; end;

    // The new parameter types may reference a unit not yet in uses (e.g.
    // VirtualTrees for TDimension, System.UITypes for TColor). The handler
    // is declared in the class (interface section), so each type must be
    // visible there. Resolve every parameter type's declaring unit via LSP
    // (cached across the batch) and ensure it is in the interface uses.
    TargetUnit := TPath.GetFileNameWithoutExtension(AIssue.PasFile);
    Needed := TStringList.Create;
    try
      Needed.Duplicates := dupIgnore;
      Needed.Sorted := True;
      if AContext <> nil then
        for var Ty in AIssue.ExpectedNorm.Split(['|']) do
        begin
          U := AContext.UnitOfType(Ty, AIssue.EventTypeFile, AIssue.EventTypeLine);
          if (U <> '') and not SameText(U, TargetUnit) then Needed.Add(U);
        end;
      // Backstop for the common case / when LSP is unavailable: the event
      // type's own unit (its parameter types usually live there too).
      U := ReadUnitNameOf(AIssue.EventTypeFile);
      if (U <> '') and not SameText(U, TargetUnit) and not SameText(U, 'System') then
        Needed.Add(U);
      for U in Needed do
        EnsureInterfaceUses(U);   // no-op when already present
    finally
      Needed.Free;
    end;

    Result := Editor.ReplaceFileContent(AIssue.PasFile, Lines.Text);
    if not Result then
      AFailReason := 'write failed: ' + AIssue.PasFile;
  finally
    Lines.Free;
  end;
end;

end.
