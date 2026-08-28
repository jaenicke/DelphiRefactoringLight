(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.AutoImport;

// "Add unit for a missing symbol" (auto-import), driven by DelphiLSP's own
// diagnostics: an undeclared identifier is reported as code E2003, from which
// we take the identifier, look it up in the background unit index, and add the
// declaring unit to the file's uses clause.
//
// Two entry points:
//   AddUnitForIdentifierAtCursor - resolves the identifier under the caret,
//     showing a small caret-anchored chooser popup (VS-like) when there is
//     more than one candidate unit or section choice.
//   ResolveMissingUnits          - a batch dialog listing every unresolved
//     identifier in the file with its suggested unit + target section.
//
// Live indicator ("lightbulb"): StartAutoImportLive watches the active
// buffer. Diagnostics come from either of two sources:
//   * IDE: Expert.StructureErrors pushes the IDE's own Error Insight
//     results (Structure view notifier) via LiveReportErrorDiags - primary,
//     no LSP round-trip of our own.
//   * Fallback / standalone: a low-frequency poller re-runs our LSP
//     diagnostics in the background after the buffer has been idle (all
//     ToolsAPI access on the main thread; the worker only waits on LSP).
// When the caret sits on a line with a missing identifier whose unit is
// known, a small non-focus-stealing button appears just below the caret;
// clicking it opens the chooser popup. The hint is bound to the error line:
// it hides when the caret moves elsewhere, the line scrolls out of view or
// the buffer changes.

interface

uses
  Lsp.Protocol;

procedure AddUnitForIdentifierAtCursor;
procedure ResolveMissingUnits;

procedure StartAutoImportLive;
procedure StopAutoImportLive;

/// <summary>True when the live checker holds FRESH results (matching the
///  current buffer content) for AFile; ACount is then the number of missing
///  identifiers with at least one known unit. Used by menus to enable /
///  annotate their quick-fix entry. False = no live data (caller should
///  keep the entry enabled and let the on-demand path decide).</summary>
function LiveFreshInfo(const AFile: string; out ACount: Integer): Boolean;

/// <summary>Feeds the live checker with error diagnostics obtained from an
///  EXTERNAL source - in the IDE that is the Structure view, which mirrors
///  Delphi's own Error Insight (see Expert.StructureErrors). AContent must
///  be the buffer content the diagnostics refer to. Marks the buffer state
///  as answered, so the LSP polling fallback skips its own analysis for
///  this state. Main thread only.</summary>
procedure LiveReportErrorDiags(const AFile, AContent: string;
  const ADiags: TArray<TLspErrorDiag>);

/// <summary>Opens the quick-fix chooser popup at the caret for AIdent
///  (or for the first missing identifier when AIdent is ''), taken from
///  the FRESH live results of AFile. False when there is no matching
///  fresh entry. Used by the Structure-view double-click integration.</summary>
function LiveShowFixAtCaret(const AFile, AIdent: string): Boolean;

implementation

uses
  System.SysUtils, System.Classes, System.Types, System.UITypes,
  System.Generics.Collections, System.IOUtils, System.Math, System.StrUtils,
  System.Hash, System.JSON,
  Winapi.Windows,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Graphics, Vcl.Dialogs,
  Lsp.Client, Lsp.Uri, Expert.LspManager,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.UsesEditor,
  Expert.DialogHelper;

type
  TMissingIdent = record
    Identifier: string;
    Line: Integer;        // 0-based (from the diagnostic range)
    Section: TUsesSection;// smart default (interface if used above impl)
    Units: TArray<TFindUnitHit>;
  end;

// ---------------------------------------------------------------------------
//  Engine
// ---------------------------------------------------------------------------

function SplitContentLines(const AContent: string): TArray<string>;
begin
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

// 0-based line index of the 'implementation' keyword, or MaxInt if none.
function ImplementationLineOf(const ALines: TArray<string>): Integer;
var I: Integer;
begin
  for I := 0 to High(ALines) do
    if SameText(Trim(ALines[I]), 'implementation') then Exit(I);
  Result := MaxInt;
end;

// The same unit name is often found under several paths (3rd-party libs
// typically ship duplicate source dirs). For a uses clause only the NAME
// matters, so collapse the hits to one per unit name (first hit wins -
// Lookup lists the project scope first).
function DedupeByUnitName(const AHits: TArray<TFindUnitHit>): TArray<TFindUnitHit>;
var
  Seen: TDictionary<string, Boolean>;
  Res: TList<TFindUnitHit>;
begin
  if Length(AHits) <= 1 then Exit(AHits);
  Seen := TDictionary<string, Boolean>.Create;
  Res := TList<TFindUnitHit>.Create;
  try
    for var H in AHits do
      if not Seen.ContainsKey(UpperCase(H.UnitName)) then
      begin
        Seen.Add(UpperCase(H.UnitName), True);
        Res.Add(H);
      end;
    Result := Res.ToArray;
  finally
    Res.Free;
    Seen.Free;
  end;
end;

function ReadCurrentContent(const AFile: string; out AContent: string): Boolean;
begin
  Result := (Editor <> nil) and Editor.ReadEditorContent(AFile, AContent);
  if not Result then
  begin
    if not TFile.Exists(AFile) then Exit(False);
    try AContent := TFile.ReadAllText(AFile); Result := True; except Result := False; end;
  end;
end;

// Turns the E2003 diagnostics for a buffer into missing-identifier records
// with candidate units. Pure function of (content, diagnostics) - safe on
// ANY thread (TUnitIndex.Lookup reads an immutable snapshot).
function ExtractMissing(const AContent: string;
  const ADiags: TArray<TLspErrorDiag>): TArray<TMissingIdent>;
var
  Lines: TArray<string>;
  Res: TList<TMissingIdent>;
  Seen: TDictionary<string, Boolean>;
  ImplLine: Integer;
begin
  Result := nil;
  Lines := SplitContentLines(AContent);
  ImplLine := ImplementationLineOf(Lines);

  Res := TList<TMissingIdent>.Create;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for var D in ADiags do
    begin
      if not SameText(D.Code, 'E2003') then Continue;   // undeclared identifier
      if (D.Range.Start.Line < 0) or (D.Range.Start.Line > High(Lines)) then Continue;
      // Extract the identifier text from the diagnostic range.
      var LnTxt := Lines[D.Range.Start.Line];
      var A := D.Range.Start.Character + 1;   // 1-based for Copy
      var B := D.Range.End_.Character;
      if D.Range.End_.Line <> D.Range.Start.Line then B := Length(LnTxt);
      if (A < 1) or (B < A) or (A > Length(LnTxt)) then Continue;
      var Ident := Trim(Copy(LnTxt, A, B - A + 1));
      if Ident = '' then Continue;
      if Seen.ContainsKey(UpperCase(Ident)) then Continue;
      Seen.Add(UpperCase(Ident), True);

      var Hits := DedupeByUnitName(TUnitIndex.Instance.Lookup(Ident));
      if Length(Hits) = 0 then Continue;   // nothing we can add - skip

      var M: TMissingIdent;
      M.Identifier := Ident;
      M.Line := D.Range.Start.Line;
      if M.Line < ImplLine then M.Section := usInterface else M.Section := usImplementation;
      M.Units := Hits;
      Res.Add(M);
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
    Seen.Free;
  end;
end;

var
  // True while an on-demand gather is running; the live poller skips its
  // tick then, so two didClose/didOpen sequences never interleave.
  GOnDemandBusy: Boolean = False;

// Runs DelphiLSP over AFile, collects the undeclared-identifier (E2003)
// diagnostics, resolves each identifier via the unit index and returns the
// ones that have at least one candidate unit. AStatus reports progress.
// MAIN THREAD ONLY (talks to the editor and pumps messages).
function GatherMissing(const AFile: string; AStatus: TProc<string>): TArray<TMissingIdent>;
var
  Root, Proj, Json, Content: string;
  Client: TLspClient;
  Diags: TArray<TLspErrorDiag>;
  Before, I: Integer;
begin
  Result := nil;
  if (Editor = nil) or (AFile = '') then Exit;
  Json := Editor.FindDelphiLspJson;
  if Json = '' then
  begin
    if Assigned(AStatus) then AStatus('DelphiLSP not configured for this project.');
    Exit;
  end;
  Root := Editor.GetProjectRoot;
  if Root = '' then Root := ExtractFilePath(AFile);
  Proj := Editor.GetCurrentProjectDproj;

  GOnDemandBusy := True;
  try
    // Keep the index warm and the file synced to disk (so LSP sees the buffer).
    TUnitIndex.Instance.RefreshSourcesFromEditor;
    Editor.SaveFile(AFile);

    if Assigned(AStatus) then AStatus('Starting LSP...');
    Client := TLspManager.Instance.GetClient(Root, Proj, Json);
    if Client = nil then Exit;

    // Force a FRESH analysis (RefreshDocument sends a didChange; documentSymbol
    // triggers parsing) and wait for the resulting diagnostics push.
    Before := Client.GetDiagnosticsCount;
    Client.RefreshDocument(AFile);
    try Client.GetDocumentSymbols(AFile, 15000).Free; except end;
    if Assigned(AStatus) then AStatus('Analysing (waiting for diagnostics)...');
    for I := 1 to 40 do
    begin
      if Client.GetDiagnosticsCount > Before then Break;
      Sleep(150);
      Application.ProcessMessages;
    end;
    Sleep(300);   // let this file's specific push settle
    Diags := Client.GetErrorDiagnostics(AFile);
  finally
    GOnDemandBusy := False;
  end;

  if not ReadCurrentContent(AFile, Content) then Exit;
  Result := ExtractMissing(Content, Diags);
end;

function SectionName(ASection: TUsesSection): string;
begin
  if ASection = usInterface then Result := 'interface' else Result := 'implementation';
end;

function ApplyOne(const AFile, AUnit: string; ASection: TUsesSection): Boolean;
begin
  Result := AddUnitToUses(AFile, AUnit, ASection);
end;

// ---------------------------------------------------------------------------
//  Caret-anchored chooser popup (VS-like) for the single-identifier quick fix
// ---------------------------------------------------------------------------

type
  TAutoImportPopup = class(TForm)
  private
    FFile: string;
    FMissing: TMissingIdent;
    FSection: TUsesSection;
    FLbl: TLabel;
    FList: TListBox;
    FSecBtn: TButton;
    FAddBtn: TButton;
    procedure DoAdd(Sender: TObject);
    procedure DoToggleSection(Sender: TObject);
    procedure DoDeactivate(Sender: TObject);
    procedure DoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoPopupClose(Sender: TObject; var Action: TCloseAction);
    procedure UpdateSectionCaption;
  public
    constructor CreatePopup(const AFile: string; const AMissing: TMissingIdent);
    procedure ShowAt(const APt: TPoint);
  end;

constructor TAutoImportPopup.CreatePopup(const AFile: string; const AMissing: TMissingIdent);
begin
  inherited CreateNew(nil);
  FFile := AFile;
  FMissing := AMissing;
  FSection := AMissing.Section;
  BorderStyle := bsNone;
  FormStyle := fsStayOnTop;
  Color := clWindow;
  Width := 340;
  KeyPreview := True;
  OnDeactivate := DoDeactivate;
  OnKeyDown := DoKeyDown;
  OnClose := DoPopupClose;

  // Explicit Top values BEFORE Align: with several alTop controls the VCL
  // stacks them by their current position, so without this the last-created
  // control (the button panel) would end up on top.
  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Top := 0;
  FLbl.Align := alTop;
  FLbl.AlignWithMargins := True;
  FLbl.Caption := Format('Add unit for "%s":', [FMissing.Identifier]);

  FList := TListBox.Create(Self);
  FList.Parent := Self;
  FList.Top := 50;
  FList.Align := alTop;
  FList.Height := 90;
  FList.AlignWithMargins := True;
  for var H in FMissing.Units do
    FList.Items.Add(H.UnitName);
  if FList.Items.Count > 0 then FList.ItemIndex := 0;
  FList.OnDblClick := DoAdd;   // double-click = insert straight away

  var Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Top := 200;
  Panel.Align := alTop;
  Panel.BevelOuter := bvNone;
  Panel.Height := 34;

  FSecBtn := TButton.Create(Self);
  FSecBtn.Parent := Panel;
  FSecBtn.Align := alLeft;
  FSecBtn.Width := 150;
  FSecBtn.AlignWithMargins := True;
  FSecBtn.OnClick := DoToggleSection;

  FAddBtn := TButton.Create(Self);
  FAddBtn.Parent := Panel;
  FAddBtn.Align := alRight;
  FAddBtn.Width := 80;
  FAddBtn.AlignWithMargins := True;
  FAddBtn.Caption := 'Add';
  FAddBtn.Default := True;
  FAddBtn.OnClick := DoAdd;

  UpdateSectionCaption;
  ClientHeight := FLbl.Height + FList.Height + Panel.Height + 20;
end;

procedure TAutoImportPopup.UpdateSectionCaption;
begin
  FSecBtn.Caption := 'Section: ' + SectionName(FSection);
end;

procedure TAutoImportPopup.DoToggleSection(Sender: TObject);
begin
  if FSection = usInterface then FSection := usImplementation else FSection := usInterface;
  UpdateSectionCaption;
end;

procedure TAutoImportPopup.DoAdd(Sender: TObject);
begin
  if (FList.ItemIndex < 0) or (FList.ItemIndex > High(FMissing.Units)) then Exit;
  var UnitName := FMissing.Units[FList.ItemIndex].UnitName;
  var Ok := ApplyOne(FFile, UnitName, FSection);
  Close;
  if not Ok then
    ShowMessage(Format('"%s" is already reachable in uses, or the clause ' +
      'could not be rewritten (e.g. IFDEFs inside it).', [UnitName]));
end;

procedure TAutoImportPopup.DoDeactivate(Sender: TObject);
begin
  Close;
end;

procedure TAutoImportPopup.DoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then Close;
end;

procedure TAutoImportPopup.DoPopupClose(Sender: TObject; var Action: TCloseAction);
begin
  // Non-modal, owner-less popup: Close alone would only hide it and leak
  // one instance per use. caFree releases it safely (deferred CM_RELEASE).
  Action := caFree;
end;

procedure TAutoImportPopup.ShowAt(const APt: TPoint);
begin
  // Keep the popup on the visible work area.
  var R := Screen.WorkAreaRect;
  Left := Min(APt.X, R.Right - Width - 4);
  Top := Min(APt.Y, R.Bottom - Height - 4);
  Show;
end;

// True when the focused window is a code editor we know how to anchor to:
// the IDE's editor class is 'TEditControl', the standalone's is a TMemo.
// Everything else (Structure pane, Object Inspector, ...) has no usable
// text caret - GetCaretPos would return garbage for those.
function FocusedEditorWindow(out AWnd: HWND): Boolean;
var
  Cls: array[0..63] of Char;
  ClsName: string;
begin
  Result := False;
  AWnd := GetFocus;
  if AWnd = 0 then Exit;
  if GetClassName(AWnd, Cls, Length(Cls)) = 0 then Exit;
  ClsName := string(Cls);
  Result := SameText(ClsName, 'TEditControl') or SameText(ClsName, 'TMemo');
end;

// Caret position of the focused EDITOR, in screen coordinates. False when
// the focus is not on an editor or the caret lies outside the visible
// client area (e.g. its line has been scrolled out of view).
function CaretScreenPos(out APt: TPoint): Boolean;
var
  H: HWND;
  P: TPoint;
  R: TRect;
begin
  Result := False;
  if not FocusedEditorWindow(H) then Exit;
  if not GetCaretPos(P) then Exit;
  if not GetClientRect(H, R) then Exit;
  if not PtInRect(R, P) then Exit;
  ClientToScreen(H, P);
  Inc(P.Y, 18);   // just below the caret line
  APt := P;
  Result := True;
end;

// ---------------------------------------------------------------------------
//  Batch dialog
// ---------------------------------------------------------------------------

type
  TAutoImportDialog = class(TForm)
  private
    FFile: string;
    FItems: TArray<TMissingIdent>;
    FSel: TArray<Integer>;      // chosen unit index per row
    FSecSel: TArray<TUsesSection>;
    FList: TListView;
    FCombo: TComboBox;
    FRadio: TRadioGroup;
    FBtnApply: TButton;
    FBtnClose: TButton;
    FLbl: TLabel;
    procedure Fill;
    procedure DoSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure DoComboChange(Sender: TObject);
    procedure DoRadioClick(Sender: TObject);
    procedure DoApply(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
  public
    constructor CreateDialog(AOwner: TComponent; const AFile: string;
      const AItems: TArray<TMissingIdent>);
  end;

constructor TAutoImportDialog.CreateDialog(AOwner: TComponent; const AFile: string;
  const AItems: TArray<TMissingIdent>);
var
  Col: TListColumn;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  FFile := AFile;
  FItems := AItems;
  SetLength(FSel, Length(FItems));
  SetLength(FSecSel, Length(FItems));
  for var I := 0 to High(FItems) do
  begin
    FSel[I] := 0;
    FSecSel[I] := FItems[I].Section;
  end;

  Caption := 'Resolve missing units';
  Width := 720;
  Height := 460;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Align := alTop;
  FLbl.AlignWithMargins := True;
  FLbl.Caption := Format('%d unresolved identifier(s) with a known unit. ' +
    'Tick the ones to add, adjust unit/section below.', [Length(FItems)]);

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.Checkboxes := True;
  FList.OnSelectItem := DoSelect;
  Col := FList.Columns.Add; Col.Caption := 'Identifier'; Col.Width := 200;
  Col := FList.Columns.Add; Col.Caption := 'Unit';       Col.Width := 220;
  Col := FList.Columns.Add; Col.Caption := 'Section';    Col.Width := 120;
  Col := FList.Columns.Add; Col.Caption := 'Options';    Col.Width := 90;

  // Bottom editing panel.
  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Align := alBottom;
  Panel.Height := 96;
  Panel.BevelOuter := bvNone;

  FCombo := TComboBox.Create(Self);
  FCombo.Parent := Panel;
  FCombo.Style := csDropDownList;
  FCombo.Left := 8; FCombo.Top := 8; FCombo.Width := 300;
  FCombo.OnChange := DoComboChange;

  FRadio := TRadioGroup.Create(Self);
  FRadio.Parent := Panel;
  FRadio.Left := 320; FRadio.Top := 4; FRadio.Width := 220; FRadio.Height := 52;
  FRadio.Caption := 'Target section';
  FRadio.Columns := 2;
  FRadio.Items.Add('interface');
  FRadio.Items.Add('implementation');
  FRadio.OnClick := DoRadioClick;

  FBtnApply := TButton.Create(Self);
  FBtnApply.Parent := Panel;
  FBtnApply.Caption := '&Apply checked';
  FBtnApply.Left := 8; FBtnApply.Top := 60; FBtnApply.Width := 130;
  FBtnApply.OnClick := DoApply;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Left := 150; FBtnClose.Top := 60; FBtnClose.Width := 90;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  Fill;
  PrepareDialog(Self, AOwner);
end;

procedure TAutoImportDialog.Fill;
var Item: TListItem;
begin
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for var I := 0 to High(FItems) do
    begin
      Item := FList.Items.Add;
      Item.Caption := FItems[I].Identifier;
      Item.Checked := True;
      Item.SubItems.Add(FItems[I].Units[FSel[I]].UnitName);
      Item.SubItems.Add(SectionName(FSecSel[I]));
      if Length(FItems[I].Units) > 1 then
        Item.SubItems.Add(Format('%d units', [Length(FItems[I].Units)]))
      else
        Item.SubItems.Add('');
    end;
  finally
    FList.Items.EndUpdate;
  end;
end;

procedure TAutoImportDialog.DoSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
begin
  if not Selected or (Item = nil) then Exit;
  var Idx := Item.Index;
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  FCombo.Items.BeginUpdate;
  try
    FCombo.Items.Clear;
    for var H in FItems[Idx].Units do FCombo.Items.Add(H.UnitName + '   (' + H.Path + ')');
  finally
    FCombo.Items.EndUpdate;
  end;
  FCombo.ItemIndex := FSel[Idx];
  FRadio.ItemIndex := Ord(FSecSel[Idx]);
end;

procedure TAutoImportDialog.DoComboChange(Sender: TObject);
begin
  if FList.Selected = nil then Exit;
  var Idx := FList.Selected.Index;
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if (FCombo.ItemIndex >= 0) and (FCombo.ItemIndex <= High(FItems[Idx].Units)) then
  begin
    FSel[Idx] := FCombo.ItemIndex;
    FList.Selected.SubItems[0] := FItems[Idx].Units[FSel[Idx]].UnitName;
  end;
end;

procedure TAutoImportDialog.DoRadioClick(Sender: TObject);
begin
  if FList.Selected = nil then Exit;
  var Idx := FList.Selected.Index;
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if FRadio.ItemIndex >= 0 then
  begin
    FSecSel[Idx] := TUsesSection(FRadio.ItemIndex);
    FList.Selected.SubItems[1] := SectionName(FSecSel[Idx]);
  end;
end;

procedure TAutoImportDialog.DoApply(Sender: TObject);
var Added, Failed: Integer;
begin
  Added := 0; Failed := 0;
  for var I := 0 to High(FItems) do
    if FList.Items[I].Checked then
    begin
      if ApplyOne(FFile, FItems[I].Units[FSel[I]].UnitName, FSecSel[I]) then Inc(Added)
      else Inc(Failed);
    end;
  ShowMessage(Format('%d unit(s) added.%s', [Added,
    IfThen(Failed > 0, Format(#13#10'%d already present / not written.', [Failed]), '')]));
  Close;
end;

procedure TAutoImportDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

// ---------------------------------------------------------------------------
//  Live quick-fix indicator ("lightbulb")
// ---------------------------------------------------------------------------
//
// A 600 ms UI timer watches the active editor buffer (via the CHEAP
// GetActiveFileName + ReadEditorContent - never GetCurrentContext, which
// would move the caret). When the content has been stable for ~1.5 s and
// differs from the last analysed state, the poller
//   1. on the MAIN thread: captures the content, sends RefreshDocument and
//      an async documentSymbol request (both are quick pipe writes),
//   2. on a WORKER thread: waits for the documentSymbol response and the
//      diagnostics push, then extracts the missing identifiers (pure
//      function + lock-free index snapshot - no editor access),
//   3. back on the main thread: stores the results and shows/hides the
//      hint button next to the caret.
// The hint window is WS_EX_NOACTIVATE, so clicking it never steals the
// editor focus mid-typing.

type
  TAutoImportHint = class(TForm)
  private
    FLbl: TLabel;
    FOnFix: TProc;
    procedure DoClick(Sender: TObject);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    constructor CreateHint;
    procedure SetInfo(const AIdent: string);
    procedure ShowNoActivateAt(const APt: TPoint);
    /// <summary>Counterpart to ShowNoActivateAt. The window is shown via
    ///  plain ShowWindow (bypassing VCL's Visible), so TForm.Hide would be
    ///  a NO-OP - it must be hidden the same WinAPI way.</summary>
    procedure HideHint;
    property OnFix: TProc read FOnFix write FOnFix;
  end;

constructor TAutoImportHint.CreateHint;
begin
  inherited CreateNew(nil);
  BorderStyle := bsNone;
  Color := clInfoBk;
  Height := 24;
  Width := 160;
  Cursor := crHandPoint;
  OnClick := DoClick;

  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Align := alClient;
  FLbl.Alignment := taCenter;
  FLbl.Layout := tlCenter;
  FLbl.Font.Color := clInfoText;
  FLbl.Cursor := crHandPoint;
  FLbl.OnClick := DoClick;
end;

procedure TAutoImportHint.CreateParams(var Params: TCreateParams);
begin
  inherited;
  // Never activate: the user keeps typing in the editor while we appear.
  Params.ExStyle := Params.ExStyle or WS_EX_NOACTIVATE or WS_EX_TOPMOST
    or WS_EX_TOOLWINDOW;
  Params.Style := (Params.Style and not WS_CHILD) or WS_POPUP;
end;

procedure TAutoImportHint.SetInfo(const AIdent: string);
begin
  FLbl.Caption := Format(#$1F4A1' Add unit for "%s"', [AIdent]);
  Width := Max(120, Canvas.TextWidth(FLbl.Caption) + 24);
end;

procedure TAutoImportHint.ShowNoActivateAt(const APt: TPoint);
var
  R: TRect;
begin
  R := Screen.WorkAreaRect;
  SetWindowPos(Handle, HWND_TOPMOST,
    Min(APt.X, R.Right - Width - 4), Min(APt.Y, R.Bottom - Height - 4),
    Width, Height, SWP_NOACTIVATE);
  if not IsWindowVisible(Handle) then
    ShowWindow(Handle, SW_SHOWNOACTIVATE);
end;

procedure TAutoImportHint.HideHint;
begin
  if HandleAllocated and IsWindowVisible(Handle) then
    ShowWindow(Handle, SW_HIDE);
end;

procedure TAutoImportHint.DoClick(Sender: TObject);
begin
  if Assigned(FOnFix) then FOnFix();
end;

type
  TAutoImportLive = class
  private
    FTimer: TTimer;
    FHint: TAutoImportHint;
    // Current buffer state (UI thread only).
    FFile: string;
    FHash: Integer;
    FDirty: Boolean;
    FDirtyTick: Cardinal;
    FAnalysing: Boolean;
    // Last analysis results (UI thread only).
    FResFile: string;
    FResHash: Integer;
    FResults: TArray<TMissingIdent>;
    // Buffer hash the EXTERNAL source (IDE Structure view) last answered
    // for. While it matches FHash the LSP polling fallback stays quiet -
    // the IDE's own Error Insight already delivered fresher data than an
    // extra LSP round-trip could.
    FExternAnsweredHash: Integer;
    procedure ApplyExternalDiags(const AFile, AContent: string;
      const ADiags: TArray<TLspErrorDiag>);
    procedure DoTick(Sender: TObject);
    procedure StartAnalysis(const AFile, AContent: string; AHash: Integer);
    procedure AnalysisDone(const AFile: string; AHash: Integer;
      const AResults: TArray<TMissingIdent>);
    procedure UpdateHint;
    procedure RunFix;
    function HasFresh(const AFile: string): Boolean;
    /// <summary>The missing identifier whose error LINE the caret is on -
    ///  the hint is bound to the error location (VS-style), not shown for
    ///  arbitrary caret positions in the file.</summary>
    function FindAtCaretLine(out AMissing: TMissingIdent): Boolean;
  public
    constructor Create;
    destructor Destroy; override;
  end;

var
  GLive: TAutoImportLive = nil;

const
  LiveTickMs  = 400;   // poll interval (also the max hint-update latency,
                       // since ONLY the tick may touch the hint window)
  LiveIdleMs  = 1500;  // buffer must be stable this long before analysing

constructor TAutoImportLive.Create;
begin
  inherited;
  FTimer := TTimer.Create(nil);
  FTimer.Interval := LiveTickMs;
  FTimer.OnTimer := DoTick;
  FTimer.Enabled := True;
end;

destructor TAutoImportLive.Destroy;
begin
  FreeAndNil(FTimer);
  FreeAndNil(FHint);
  inherited;
end;

function TAutoImportLive.HasFresh(const AFile: string): Boolean;
begin
  Result := (AFile <> '') and SameText(FResFile, AFile)
    and SameText(FResFile, FFile) and (FResHash = FHash);
end;

function TAutoImportLive.FindAtCaretLine(out AMissing: TMissingIdent): Boolean;
var
  Line, Col: Integer;
begin
  Result := False;
  if not HasFresh(FFile) then Exit;
  if (Editor = nil) or not Editor.GetCaretLineCol(Line, Col) then Exit;
  for var R in FResults do
    if R.Line = Line - 1 then   // R.Line is 0-based
    begin
      AMissing := R;
      Exit(True);
    end;
end;

procedure TAutoImportLive.DoTick(Sender: TObject);
var
  F, Content: string;
  H: Integer;
  Tick: Cardinal;
begin
  if (Editor = nil) or GOnDemandBusy then Exit;
  // No live churn while any modal dialog (incl. our own wizards) is open.
  if Application.ModalLevel > 0 then
  begin
    if FHint <> nil then FHint.HideHint;
    Exit;
  end;

  F := Editor.GetActiveFileName;
  if (F = '') or not SameText(ExtractFileExt(F), '.pas') then
  begin
    if FHint <> nil then FHint.HideHint;
    Exit;
  end;

  if not SameText(F, FFile) then
  begin
    FFile := F;
    FHash := 0;
    FDirty := False;   // hash below re-arms
    if FHint <> nil then FHint.HideHint;
  end;

  if not Editor.ReadEditorContent(F, Content) then
  begin
    if FHint <> nil then FHint.HideHint;
    Exit;
  end;

  H := THashBobJenkins.GetHashValue(Content);
  Tick := GetTickCount;
  if H <> FHash then
  begin
    FHash := H;
    FDirty := True;
    FDirtyTick := Tick;
    if FHint <> nil then FHint.HideHint;   // results are stale now
  end;

  if FDirty and (not FAnalysing) and (Tick - FDirtyTick >= LiveIdleMs)
    and (FExternAnsweredHash <> FHash) then
    StartAnalysis(F, Content, H);

  UpdateHint;
end;

procedure TAutoImportLive.ApplyExternalDiags(const AFile, AContent: string;
  const ADiags: TArray<TLspErrorDiag>);
begin
  if AFile = '' then Exit;
  // STATE ONLY - no window operations. This runs inside the IDE's
  // Structure-view notification, which the IDE dispatches from
  // CheckSynchronize during its LSP refresh. Showing/hiding a window
  // there triggers a synchronous activation cascade into
  // TEditWindow.ActivateModule -> ParserThread.CancelAndLock and can
  // DEADLOCK the IDE against its parser thread. The poll tick (a plain
  // WM_TIMER, safe context) picks these results up and updates the hint.
  FFile := AFile;
  FHash := THashBobJenkins.GetHashValue(AContent);
  FExternAnsweredHash := FHash;
  FDirty := False;

  FResFile := AFile;
  FResHash := FHash;
  FResults := ExtractMissing(AContent, ADiags);
end;

procedure TAutoImportLive.StartAnalysis(const AFile, AContent: string; AHash: Integer);
var
  Client: TLspClient;
  Params, TextDocObj: TJSONObject;
  Before, ReqId: Integer;
begin
  // NEVER cold-start the LSP from the background poller - only piggyback
  // on a client the prewarmer / a wizard has already brought up.
  Client := TLspManager.Instance.PeekClient;
  if Client = nil then
  begin
    FDirtyTick := GetTickCount;   // retry after another idle period
    Exit;
  end;

  Before := Client.GetDiagnosticsCount;
  try
    // Push the CURRENT buffer (RefreshDocument reads the live editor
    // content) and fire an async documentSymbol to force the analysis.
    Client.RefreshDocument(AFile);
    TextDocObj := TJSONObject.Create;
    TextDocObj.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFile)));
    Params := TJSONObject.Create;
    Params.AddPair('textDocument', TextDocObj);
    ReqId := Client.SendRequestAsync('textDocument/documentSymbol', Params);
  except
    Exit;   // client just died etc. - try again on a later tick
  end;

  FAnalysing := True;
  FDirty := False;

  TThread.CreateAnonymousThread(
    procedure
    var
      R: TArray<TMissingIdent>;
      I: Integer;
    begin
      R := nil;
      try
        try Client.WaitForResponse(ReqId, 20000).Free; except end;
        for I := 1 to 40 do
        begin
          if Client.GetDiagnosticsCount > Before then Break;
          Sleep(150);
        end;
        Sleep(300);   // let this file's push settle
        R := ExtractMissing(AContent, Client.GetErrorDiagnostics(AFile));
      except
        R := nil;
      end;
      TThread.Queue(nil,
        procedure
        begin
          if GLive <> nil then
            GLive.AnalysisDone(AFile, AHash, R);
        end);
    end).Start;
end;

procedure TAutoImportLive.AnalysisDone(const AFile: string; AHash: Integer;
  const AResults: TArray<TMissingIdent>);
begin
  // STATE ONLY (arrives via TThread.Queue, i.e. inside CheckSynchronize -
  // see ApplyExternalDiags for why no window operation may happen here).
  // The next poll tick shows/hides the hint.
  FAnalysing := False;
  // Only accept results that still match the current buffer.
  if SameText(AFile, FFile) and (AHash = FHash) then
  begin
    FResFile := AFile;
    FResHash := AHash;
    FResults := AResults;
  end;
end;

procedure TAutoImportLive.UpdateHint;
var
  Pt: TPoint;
  M: TMissingIdent;
begin
  // Bound to the ERROR LINE: shown only while the caret is on a line with
  // a missing identifier (and that line is scrolled into view -
  // CaretScreenPos rejects carets outside the visible client area).
  if FindAtCaretLine(M) and CaretScreenPos(Pt) then
  begin
    if FHint = nil then
    begin
      FHint := TAutoImportHint.CreateHint;
      FHint.OnFix := RunFix;
    end;
    FHint.SetInfo(M.Identifier);
    FHint.ShowNoActivateAt(Pt);
  end
  else if FHint <> nil then
    FHint.HideHint;
end;

procedure TAutoImportLive.RunFix;
var
  Pt: TPoint;
  M: TMissingIdent;
begin
  if not FindAtCaretLine(M) then Exit;
  if FHint <> nil then FHint.HideHint;
  var Popup := TAutoImportPopup.CreatePopup(FFile, M);
  if CaretScreenPos(Pt) then Popup.ShowAt(Pt)
  else Popup.ShowAt(Mouse.CursorPos);
end;

procedure StartAutoImportLive;
begin
  if GLive = nil then
    GLive := TAutoImportLive.Create;
end;

procedure StopAutoImportLive;
begin
  FreeAndNil(GLive);
end;

function LiveFreshInfo(const AFile: string; out ACount: Integer): Boolean;
begin
  ACount := 0;
  Result := (GLive <> nil) and GLive.HasFresh(AFile);
  if Result then
    ACount := Length(GLive.FResults);
end;

procedure LiveReportErrorDiags(const AFile, AContent: string;
  const ADiags: TArray<TLspErrorDiag>);
begin
  if GLive <> nil then
    GLive.ApplyExternalDiags(AFile, AContent, ADiags);
end;

function LiveShowFixAtCaret(const AFile, AIdent: string): Boolean;
var
  Pt: TPoint;
  M: TMissingIdent;
  Found: Boolean;
begin
  Result := False;
  if (GLive = nil) or not GLive.HasFresh(AFile) then Exit;
  Found := False;
  for var R in GLive.FResults do
    if (AIdent = '') or SameText(R.Identifier, AIdent) then
    begin
      M := R;
      Found := True;
      Break;
    end;
  if not Found then Exit;
  var Popup := TAutoImportPopup.CreatePopup(AFile, M);
  if CaretScreenPos(Pt) then Popup.ShowAt(Pt)
  else Popup.ShowAt(Mouse.CursorPos);
  Result := True;
end;

// ---------------------------------------------------------------------------
//  Public entry points
// ---------------------------------------------------------------------------

procedure WithWaitCursor(AProc: TProc);
begin
  Screen.Cursor := crHourGlass;
  try AProc; finally Screen.Cursor := crDefault; end;
end;

procedure AddUnitForIdentifierAtCursor;
var
  Ctx: TEditorContext;
  Missing: TArray<TMissingIdent>;
  Chosen: TMissingIdent;
  Found: Boolean;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if Ctx.FileName = '' then Exit;

  Missing := nil;
  WithWaitCursor(procedure begin Missing := GatherMissing(Ctx.FileName, nil); end);

  // Prefer the missing identifier at the caret; else the word under it.
  Found := False;
  for var M in Missing do
    if (M.Line + 1 = Ctx.Line) and SameText(M.Identifier, Ctx.WordAtCursor) then
    begin Chosen := M; Found := True; Break; end;
  if not Found then
    for var M in Missing do
      if SameText(M.Identifier, Ctx.WordAtCursor) then
      begin Chosen := M; Found := True; Break; end;

  if not Found then
  begin
    // Not flagged by LSP - offer it anyway if the index knows the word.
    var Hits := DedupeByUnitName(TUnitIndex.Instance.Lookup(Ctx.WordAtCursor));
    if Length(Hits) = 0 then
    begin
      ShowMessage(Format('No unit found for "%s" (and it is not reported as ' +
        'an undeclared identifier).', [Ctx.WordAtCursor]));
      Exit;
    end;
    Chosen.Identifier := Ctx.WordAtCursor;
    Chosen.Line := Ctx.Line - 1;
    Chosen.Units := Hits;
    Chosen.Section := usInterface;
  end;

  // One obvious choice -> apply straight away; else show the chooser popup.
  if Length(Chosen.Units) = 1 then
  begin
    if not ApplyOne(Ctx.FileName, Chosen.Units[0].UnitName, Chosen.Section) then
      ShowMessage(Format('"%s" is already reachable in uses, or the clause ' +
        'could not be rewritten (e.g. IFDEFs inside it).',
        [Chosen.Units[0].UnitName]));
    Exit;
  end;

  var Pt: TPoint;
  var Popup := TAutoImportPopup.CreatePopup(Ctx.FileName, Chosen);
  if CaretScreenPos(Pt) then Popup.ShowAt(Pt)
  else Popup.ShowAt(Mouse.CursorPos);
end;

procedure ResolveMissingUnits;
var
  Ctx: TEditorContext;
  Missing: TArray<TMissingIdent>;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if Ctx.FileName = '' then
  begin ShowMessage('No active editor file.'); Exit; end;

  Missing := nil;
  WithWaitCursor(procedure begin Missing := GatherMissing(Ctx.FileName, nil); end);
  if Length(Missing) = 0 then
  begin
    ShowMessage('No unresolved identifiers with a known unit were found.');
    Exit;
  end;
  var Dlg := TAutoImportDialog.CreateDialog(Application.MainForm, Ctx.FileName, Missing);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
