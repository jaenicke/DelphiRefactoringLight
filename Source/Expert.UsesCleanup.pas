(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UsesCleanup;

// "Uses cleanup" for the current unit: flags uses entries that are
//   * UNUSED   - no identifier in the file resolves to the unit -> remove
//   * MOVABLE  - listed under interface but every usage sits in the
//                implementation section -> move down (helps against
//                circular interface references)
// The analysis is driven by the background identifier index and is
// deliberately CONSERVATIVE:
//   * an identifier declared by several of the used units counts as a
//     usage of ALL of them (the compiler picks one, we cannot),
//   * a textual occurrence of the full unit name anywhere in the file
//     (qualified access, comments included) counts as a usage,
//   * units without indexed source (.dcu-only) are never flagged,
//   * uses clauses containing {$IFDEF}s are never analysed.
// Removing a unit can still change behavior (initialization sections,
// class helpers, operator visibility) - the dialog says so and lets the
// user decide per row.

interface

uses
  System.SysUtils, Expert.UsesEditor;

type
  TUsesVerdict = (uvUsed, uvUnused, uvMovable, uvUnknown,
    uvInitCode,     // textually unused, but the unit runs initialization/
                    // finalization code - removing it would change behavior
    uvIdeManaged);  // the IDE auto-manages this uses entry for form units
                    // (it would silently re-add it - don't fight it)

  TUsesEntryInfo = record
    UnitName: string;
    Section: TUsesSection;
    Verdict: TUsesVerdict;
    UsageCount: Integer;
    FirstUseLine: Integer;   // 0-based, -1 when unused
  end;

  /// <summary>Identifier -> declaring unit names (the index Lookup).</summary>
  TIdentLookup = reference to function(const AIdent: string): TArray<string>;
  /// <summary>Is the unit's source indexed (analysable at all)?</summary>
  TUnitKnown = reference to function(const AUnitName: string): Boolean;

/// <summary>Analyses AContent's uses entries. Pure function of the content
///  plus the injected lookups - unit-testable without an index.
///  AHasInitCode: does the unit run initialization/finalization code?</summary>
function AnalyzeUses(const AContent: string; const ALookup: TIdentLookup;
  const AUnitKnown: TUnitKnown;
  const AHasInitCode: TUnitKnown): TArray<TUsesEntryInfo>;

/// <summary>Opens the cleanup dialog for the active editor file.</summary>
procedure CleanupUsesCurrentUnit;

implementation

uses
  System.Classes, System.Generics.Collections, System.Math, System.StrUtils,
  System.UITypes,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Dialogs,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort;

// ---------------------------------------------------------------------------
//  Analysis
// ---------------------------------------------------------------------------

type
  TClauseInfo = record
    Section: TUsesSection;
    FirstLine, LastLine: Integer;   // 0-based span of the clause
    Names: TArray<string>;
    HasDirectives: Boolean;         // {$IFDEF} etc. inside - do not analyse
  end;

const
  // Units the IDE inserts and maintains ITSELF for form units (action
  // manager styles, touch keyboard, ...). Removing them just makes the
  // IDE silently re-add them - a pointless tug-of-war, so they are never
  // offered for removal or moving.
  IdeManagedUnits: array[0..3] of string = (
    'System.Actions', 'Vcl.XPStyleActnCtrls',
    'Vcl.PlatformDefaultStyleActnCtrls', 'Vcl.Touch.Keyboard');

function IsIdeManagedUnit(const AName: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := Low(IdeManagedUnits) to High(IdeManagedUnits) do
    if SameText(IdeManagedUnits[I], AName) then
      Exit(True);
end;

function SplitLines(const AContent: string): TArray<string>;
begin
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

function ImplLineOf(const ALines: TArray<string>): Integer;
var
  I: Integer;
begin
  for I := 0 to High(ALines) do
    if SameText(Trim(StripLineComment(ALines[I])), 'implementation') then
      Exit(I);
  Result := MaxInt;
end;

// Parses the uses clause of the section starting after ASectionLine.
function ParseClause(const ALines: TArray<string>; AFrom, ATo: Integer;
  ASection: TUsesSection; out AInfo: TClauseInfo): Boolean;
var
  I, J: Integer;
  T, Joined: string;
  InClause: Boolean;
begin
  Result := False;
  AInfo := Default(TClauseInfo);
  AInfo.Section := ASection;
  InClause := False;
  Joined := '';
  for I := AFrom to Min(ATo, High(ALines)) do
  begin
    T := Trim(StripLineComment(ALines[I]));
    if not InClause then
    begin
      if SameText(T, 'uses') or StartsText('uses ', T) then
      begin
        InClause := True;
        AInfo.FirstLine := I;
      end
      else
        Continue;
    end;
    if (Pos('{', T) > 0) or (Pos('(*', T) > 0) then
      AInfo.HasDirectives := True;
    Joined := Joined + ' ' + T;
    J := Pos(';', T);
    if J > 0 then
    begin
      AInfo.LastLine := I;
      // Strip the 'uses' keyword and the terminator, split the names.
      J := Pos(';', Joined);
      if J > 0 then Joined := Copy(Joined, 1, J - 1);
      J := Pos('uses', LowerCase(Joined));
      if J > 0 then Joined := Copy(Joined, J + 4, MaxInt);
      for var N in Joined.Split([',']) do
        if Trim(N) <> '' then
          AInfo.Names := AInfo.Names + [Trim(N)];
      Exit(True);
    end;
  end;
end;

function AnalyzeUses(const AContent: string; const ALookup: TIdentLookup;
  const AUnitKnown: TUnitKnown;
  const AHasInitCode: TUnitKnown): TArray<TUsesEntryInfo>;
var
  Lines: TArray<string>;
  ImplLine, I, CI: Integer;
  Clauses: array[0..1] of TClauseInfo;
  HasClause: array[0..1] of Boolean;
  Entries: TList<TUsesEntryInfo>;
  ByName: TDictionary<string, Integer>;       // UPPER(unit) -> entry index
  IntfUse: TDictionary<Integer, Boolean>;     // entry -> used from interface
  TokenCache: TDictionary<string, TArray<string>>;

  function InClauseSpan(ALine: Integer): Boolean;
  var
    K: Integer;
  begin
    Result := False;
    for K := 0 to 1 do
      if HasClause[K] and (ALine >= Clauses[K].FirstLine)
        and (ALine <= Clauses[K].LastLine) then
        Exit(True);
  end;

  procedure NoteUsage(AEntry, ALine: Integer);
  var
    E: TUsesEntryInfo;
  begin
    E := Entries[AEntry];
    Inc(E.UsageCount);
    if (E.FirstUseLine < 0) or (ALine < E.FirstUseLine) then
      E.FirstUseLine := ALine;
    Entries[AEntry] := E;
    if ALine < ImplLine then
      IntfUse.AddOrSetValue(AEntry, True);
  end;

  procedure ScanIdentifiers;
  var
    L, P, StartP: Integer;
    S, Token: string;
    InBrace, InParen, InStr: Boolean;
    Units: TArray<string>;
    EntryIdx: Integer;
  begin
    InBrace := False;
    InParen := False;
    for L := 0 to High(Lines) do
    begin
      if InClauseSpan(L) then Continue;
      S := Lines[L];
      InStr := False;
      P := 1;
      while P <= Length(S) do
      begin
        if InBrace then
        begin
          if S[P] = '}' then InBrace := False;
          Inc(P);
        end
        else if InParen then
        begin
          if (S[P] = '*') and (P < Length(S)) and (S[P + 1] = ')') then
          begin
            InParen := False;
            Inc(P);
          end;
          Inc(P);
        end
        else if InStr then
        begin
          if S[P] = '''' then InStr := False;
          Inc(P);
        end
        else if S[P] = '''' then
        begin
          InStr := True;
          Inc(P);
        end
        else if S[P] = '{' then
        begin
          InBrace := True;
          Inc(P);
        end
        else if (S[P] = '(') and (P < Length(S)) and (S[P + 1] = '*') then
        begin
          InParen := True;
          Inc(P, 2);
        end
        else if (S[P] = '/') and (P < Length(S)) and (S[P + 1] = '/') then
          Break   // rest of line is a comment
        else if CharInSet(S[P], ['A'..'Z', 'a'..'z', '_']) then
        begin
          StartP := P;
          while (P <= Length(S))
            and CharInSet(S[P], ['A'..'Z', 'a'..'z', '0'..'9', '_']) do
            Inc(P);
          Token := Copy(S, StartP, P - StartP);
          if not TokenCache.TryGetValue(UpperCase(Token), Units) then
          begin
            Units := ALookup(Token);
            TokenCache.Add(UpperCase(Token), Units);
          end;
          for var U in Units do
            if ByName.TryGetValue(UpperCase(U), EntryIdx) then
              NoteUsage(EntryIdx, L);
        end
        else
          Inc(P);
      end;
    end;
  end;

  procedure ScanQualifiedNames;
  var
    L, P, AfterIdx, EntryIdx: Integer;
    U, Needle: string;
  begin
    // A textual whole-word occurrence of the FULL unit name counts as a
    // usage (qualified access 'Vcl.Dialogs.MessageDlg'). Raw-line search -
    // hits in comments only ever KEEP a unit, never remove one.
    for EntryIdx := 0 to Entries.Count - 1 do
    begin
      Needle := UpperCase(Entries[EntryIdx].UnitName);
      for L := 0 to High(Lines) do
      begin
        if InClauseSpan(L) then Continue;
        U := UpperCase(Lines[L]);
        P := Pos(Needle, U);
        while P > 0 do
        begin
          AfterIdx := P + Length(Needle);
          if ((P = 1) or not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_', '.']))
            and ((AfterIdx > Length(U))
              or not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_'])) then
          begin
            NoteUsage(EntryIdx, L);
            Break;
          end;
          P := Pos(Needle, U, P + 1);
        end;
      end;
    end;
  end;

var
  E: TUsesEntryInfo;
begin
  Result := nil;
  Lines := SplitLines(AContent);
  ImplLine := ImplLineOf(Lines);

  HasClause[0] := ParseClause(Lines, 0, ImplLine - 1, usInterface, Clauses[0]);
  if ImplLine < MaxInt then
    HasClause[1] := ParseClause(Lines, ImplLine + 1, High(Lines),
      usImplementation, Clauses[1])
  else
    HasClause[1] := False;

  Entries := TList<TUsesEntryInfo>.Create;
  ByName := TDictionary<string, Integer>.Create;
  IntfUse := TDictionary<Integer, Boolean>.Create;
  TokenCache := TDictionary<string, TArray<string>>.Create;
  try
    for CI := 0 to 1 do
    begin
      if not HasClause[CI] then Continue;
      for var N in Clauses[CI].Names do
      begin
        E := Default(TUsesEntryInfo);
        E.UnitName := N;
        E.Section := Clauses[CI].Section;
        E.FirstUseLine := -1;
        if Clauses[CI].HasDirectives or not AUnitKnown(N) then
          E.Verdict := uvUnknown
        else
          E.Verdict := uvUsed;   // refined below
        Entries.Add(E);
        if not ByName.ContainsKey(UpperCase(N)) then
          ByName.Add(UpperCase(N), Entries.Count - 1);
      end;
    end;
    if Entries.Count = 0 then Exit;

    ScanIdentifiers;
    ScanQualifiedNames;

    for I := 0 to Entries.Count - 1 do
    begin
      E := Entries[I];
      if E.Verdict = uvUnknown then Continue;
      if E.UsageCount = 0 then
      begin
        if IsIdeManagedUnit(E.UnitName) then
          E.Verdict := uvIdeManaged
        else if Assigned(AHasInitCode) and AHasInitCode(E.UnitName) then
          E.Verdict := uvInitCode
        else
          E.Verdict := uvUnused;
      end
      else if (E.Section = usInterface) and not IntfUse.ContainsKey(I)
        and not IsIdeManagedUnit(E.UnitName) then
        E.Verdict := uvMovable
      else
        E.Verdict := uvUsed;
      Entries[I] := E;
    end;
    Result := Entries.ToArray;
  finally
    TokenCache.Free;
    IntfUse.Free;
    ByName.Free;
    Entries.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Dialog
// ---------------------------------------------------------------------------

type
  TUsesCleanupDialog = class(TForm)
  private
    FFile: string;
    FEntries: TArray<TUsesEntryInfo>;
    FList: TListView;
    FLbl: TLabel;
    FLblWarn: TLabel;
    FBtnApply: TButton;
    FBtnClose: TButton;
    FBtnAll: TButton;
    FBtnNone: TButton;
    FFilling: Boolean;
    procedure Fill;
    procedure DoApply(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure DoSelectAll(Sender: TObject);
    procedure DoSelectNone(Sender: TObject);
    procedure DoItemChecked(Sender: TObject; Item: TListItem);
    function IsActionable(AIndex: Integer): Boolean;
  public
    constructor CreateDialog(AOwner: TComponent; const AFile: string;
      const AEntries: TArray<TUsesEntryInfo>);
  end;

constructor TUsesCleanupDialog.CreateDialog(AOwner: TComponent;
  const AFile: string; const AEntries: TArray<TUsesEntryInfo>);
var
  Col: TListColumn;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  FFile := AFile;
  FEntries := AEntries;

  Caption := 'Uses cleanup - ' + ExtractFileName(AFile);
  Width := 720;
  Height := 480;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Top := 0;
  FLbl.Align := alTop;
  FLbl.AlignWithMargins := True;
  FLbl.Caption := 'Tick the entries to clean up: "unused" entries are ' +
    'removed, "movable" entries are moved to the implementation uses.';

  FLblWarn := TLabel.Create(Self);
  FLblWarn.Parent := Self;
  FLblWarn.Top := 20;
  FLblWarn.Align := alTop;
  FLblWarn.AlignWithMargins := True;
  FLblWarn.WordWrap := True;
  // AutoSize + WordWrap + alTop is a trap: the wrap height is computed
  // against the label's DEFAULT width (65px) before alTop stretches it,
  // yielding a ~260px invisible giant that squeezes the list away.
  FLblWarn.AutoSize := False;
  FLblWarn.Height := 48;
  FLblWarn.Caption := 'Conservative source-index analysis - ambiguous ' +
    'identifiers count as usage. CAUTION: removing a unit also removes its ' +
    'initialization side effects, class helpers and operator visibility, ' +
    'which this analysis cannot see. Review before applying; every change ' +
    'is undoable in the editor.';

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Top := 60;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.Checkboxes := True;
  Col := FList.Columns.Add; Col.Caption := 'Unit';       Col.Width := 220;
  Col := FList.Columns.Add; Col.Caption := 'Section';    Col.Width := 110;
  Col := FList.Columns.Add; Col.Caption := 'Status';     Col.Width := 200;
  Col := FList.Columns.Add; Col.Caption := 'Usages';     Col.Width := 60;
  Col := FList.Columns.Add; Col.Caption := 'First use';  Col.Width := 70;
  FList.OnItemChecked := DoItemChecked;

  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Top := 400;
  Panel.Align := alBottom;
  Panel.BevelOuter := bvNone;
  Panel.Height := 40;

  FBtnApply := TButton.Create(Self);
  FBtnApply.Parent := Panel;
  FBtnApply.Caption := '&Apply checked';
  FBtnApply.Left := 8; FBtnApply.Top := 8; FBtnApply.Width := 130;
  FBtnApply.OnClick := DoApply;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Left := 150; FBtnClose.Top := 8; FBtnClose.Width := 90;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  FBtnAll := TButton.Create(Self);
  FBtnAll.Parent := Panel;
  FBtnAll.Caption := 'Select a&ll';
  FBtnAll.Left := 270; FBtnAll.Top := 8; FBtnAll.Width := 90;
  FBtnAll.OnClick := DoSelectAll;

  FBtnNone := TButton.Create(Self);
  FBtnNone.Parent := Panel;
  FBtnNone.Caption := 'Select &none';
  FBtnNone.Left := 366; FBtnNone.Top := 8; FBtnNone.Width := 90;
  FBtnNone.OnClick := DoSelectNone;

  Fill;
  EnableListViewSorting(FList);
  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

function TUsesCleanupDialog.IsActionable(AIndex: Integer): Boolean;
begin
  Result := (AIndex >= 0) and (AIndex <= High(FEntries))
    and (FEntries[AIndex].Verdict in [uvUnused, uvMovable]);
end;

procedure TUsesCleanupDialog.DoItemChecked(Sender: TObject; Item: TListItem);
begin
  // Non-actionable rows (used / not analysable) cannot be ticked.
  // Rows map through Item.Data - the list is sortable, Index shifts.
  if FFilling or (Item = nil) then Exit;
  if Item.Checked and not IsActionable(NativeInt(Item.Data)) then
  begin
    FFilling := True;
    try
      Item.Checked := False;
    finally
      FFilling := False;
    end;
  end;
end;

procedure TUsesCleanupDialog.DoSelectAll(Sender: TObject);
begin
  FFilling := True;
  try
    for var I := 0 to FList.Items.Count - 1 do
      FList.Items[I].Checked := IsActionable(NativeInt(FList.Items[I].Data));
  finally
    FFilling := False;
  end;
end;

procedure TUsesCleanupDialog.DoSelectNone(Sender: TObject);
begin
  FFilling := True;
  try
    for var I := 0 to FList.Items.Count - 1 do
      FList.Items[I].Checked := False;
  finally
    FFilling := False;
  end;
end;

procedure TUsesCleanupDialog.Fill;
const
  SectionNames: array[TUsesSection] of string = ('interface', 'implementation');
var
  Item: TListItem;
begin
  FFilling := True;
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for var EIdx := 0 to High(FEntries) do
    begin
      var E := FEntries[EIdx];
      Item := FList.Items.Add;
      Item.Data := Pointer(NativeInt(EIdx));   // survives column sorting
      Item.Caption := E.UnitName;
      Item.SubItems.Add(SectionNames[E.Section]);
      case E.Verdict of
        uvUnused:  Item.SubItems.Add('unused - remove?');
        uvMovable: Item.SubItems.Add('only used in implementation - move?');
        uvUsed:    Item.SubItems.Add('used');
        uvUnknown: Item.SubItems.Add('not analysable (no indexed source / IFDEF)');
        uvInitCode:
          Item.SubItems.Add('no direct usage, but has initialization code - kept');
        uvIdeManaged:
          Item.SubItems.Add('managed by the IDE (auto re-added) - kept');
      end;
      Item.SubItems.Add(IntToStr(E.UsageCount));
      if E.FirstUseLine >= 0 then
        Item.SubItems.Add(IntToStr(E.FirstUseLine + 1))
      else
        Item.SubItems.Add('');
      // UNUSED rows are pre-ticked (that is what the dialog is for);
      // movable rows stay opt-in, non-actionable rows cannot be ticked.
      Item.Checked := E.Verdict = uvUnused;
    end;
  finally
    FList.Items.EndUpdate;
    FFilling := False;
  end;
end;

procedure TUsesCleanupDialog.DoApply(Sender: TObject);
var
  I, Done, FailCount: Integer;
  E: TUsesEntryInfo;
begin
  Done := 0;
  FailCount := 0;
  for I := 0 to FList.Items.Count - 1 do
  begin
    if not FList.Items[I].Checked then Continue;
    var Idx := NativeInt(FList.Items[I].Data);
    if (Idx < 0) or (Idx > High(FEntries)) then Continue;
    E := FEntries[Idx];
    case E.Verdict of
      uvUnused:
        if RemoveUnitFromUses(FFile, E.UnitName) then Inc(Done)
        else Inc(FailCount);
      uvMovable:
        if RemoveUnitFromUses(FFile, E.UnitName) then
        begin
          if AddUnitToUses(FFile, E.UnitName, usImplementation) then
            Inc(Done)
          else
          begin
            // Never leave the unit dropped entirely - restore it.
            AddUnitToUses(FFile, E.UnitName, usInterface);
            Inc(FailCount);
          end;
        end
        else
          Inc(FailCount);
    else
      Inc(FailCount);   // checked a non-actionable row
    end;
  end;
  ShowThemedMessage(Format('%d change(s) applied.%s', [Done,
    IfThen(FailCount > 0, Format(#13#10'%d entr%s could not be changed.',
      [FailCount, IfThen(FailCount = 1, 'y', 'ies')]), '')]));
  Close;
end;

procedure TUsesCleanupDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

// ---------------------------------------------------------------------------
//  Entry point
// ---------------------------------------------------------------------------

var
  GCleanupBusy: Boolean = False;

procedure CleanupUsesCurrentUnit;
var
  Ctx: TEditorContext;
  Content: string;
  Snap: IUnitSnapshot;
  Entries: TArray<TUsesEntryInfo>;
  Cycle0, Waited: Integer;
begin
  if (Editor = nil) or GCleanupBusy then Exit;
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or not SameText(ExtractFileExt(Ctx.FileName), '.pas') then
  begin
    ShowThemedMessage('Open a .pas unit first.');
    Exit;
  end;
  if not Editor.ReadEditorContent(Ctx.FileName, Content) then
  begin
    ShowThemedMessage('Could not read the editor buffer.');
    Exit;
  end;

  GCleanupBusy := True;
  Screen.Cursor := crHourGlass;
  try
    // Trigger a rescan and WAIT for the worker to complete one cycle over
    // the refreshed sources - otherwise the first run after an edit works
    // on the stale snapshot and only the second run sees the truth (the
    // index parses files from DISK, so unsaved edits stay invisible).
    Cycle0 := TUnitIndex.Instance.ScanCycle;
    TUnitIndex.Instance.RefreshSourcesFromEditor;
    Waited := 0;
    while (TUnitIndex.Instance.ScanCycle = Cycle0) and (Waited < 5000) do
    begin
      Sleep(50);
      Inc(Waited, 50);
      Application.ProcessMessages;
    end;

    Snap := TUnitIndex.Instance.Snapshot;
    if (Snap = nil) or (Snap.IdentCount = 0) then
    begin
      ShowThemedMessage('The identifier index is still building - try again in a ' +
        'few seconds.');
      Exit;
    end;

    Entries := AnalyzeUses(Content,
      function(const AIdent: string): TArray<string>
      var
        Hits: TArray<TFindUnitHit>;
      begin
        Result := nil;
        Hits := Snap.Lookup(AIdent);
        for var H in Hits do
          Result := Result + [H.UnitName];
      end,
      function(const AUnitName: string): Boolean
      begin
        Result := Snap.HasUnit(AUnitName);
      end,
      function(const AUnitName: string): Boolean
      begin
        Result := Snap.HasInitCode(AUnitName);
      end);
  finally
    Screen.Cursor := crDefault;
    GCleanupBusy := False;
  end;

  if Length(Entries) = 0 then
  begin
    ShowThemedMessage('No uses entries found.');
    Exit;
  end;
  var Dlg := TUsesCleanupDialog.CreateDialog(Application.MainForm,
    Ctx.FileName, Entries);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
