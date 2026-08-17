(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ExtractInterfaceDialog;

{
  Non-modal dialog for the Extract / Extend interface wizard.

  Layout (left to right):
    - Top: target unit / existing interface picker + interface name + GUID
    - Left pane: TCheckListBox of class members with preset buttons
      ("All public", "All published", "Methods only", "Properties only",
      "Fields only", "Everything", "None"). Each item shows
      "visibility: kind name : type".
    - Right pane: read-only memo with the live interface preview.
    - Bottom: OK / Cancel.
}

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.CheckLst, Vcl.ExtCtrls,
  Vcl.Dialogs, Vcl.ComCtrls, Vcl.Buttons,
  Expert.ExtractInterface;

type
  TExtractInterfaceDialog = class(TForm)
  private
    FInfo: TExtractInterfaceInfo;
    FMode: TInterfaceMode;

    FLblName: TLabel;
    FEdtName: TEdit;
    FLblGuid: TLabel;
    FEdtGuid: TEdit;
    FBtnNewGuid: TButton;
    FLblTarget: TLabel;
    FEdtTarget: TEdit;
    FBtnBrowse: TButton;
    FCmbExisting: TComboBox;

    FLblMembers: TLabel;
    FClb: TCheckListBox;
    FPnlPresets: TPanel;
    FBtnPresetAll: TButton;
    FBtnPresetPublic: TButton;
    FBtnPresetPublished: TButton;
    FBtnPresetMethods: TButton;
    FBtnPresetProperties: TButton;
    FBtnPresetFields: TButton;
    FBtnPresetNone: TButton;

    FLblPreview: TLabel;
    FMemoPreview: TMemo;

    FBtnOk: TButton;
    FBtnCancel: TButton;

    FExistingList: TArray<TInterfaceDeclLocation>;
    /// <summary>UPPERCASE names of every member already declared in the
    ///  currently selected target interface (Add-to-existing mode).
    ///  Empty / nil for Extract-New mode.</summary>
    FExistingMemberNames: TArray<string>;
    /// <summary>Per-interface (UPPERCASE name) member-checked state.
    ///  Allows the user to pick different members for different
    ///  interfaces and have all selections applied on OK.</summary>
    FSelectionByInterface: TDictionary<string, TArray<Boolean>>;
    /// <summary>UPPERCASE name of the interface whose selections are
    ///  currently shown in the CheckListBox.</summary>
    FCurrentInterfaceKey: string;
    FUpdating: Boolean;

    procedure BuildLayout;
    procedure PopulateList;
    procedure RefreshPreview;

    procedure DoCheckClick(Sender: TObject);
    procedure DoNameChange(Sender: TObject);
    procedure DoGuidChange(Sender: TObject);
    procedure DoTargetChange(Sender: TObject);
    procedure DoNewGuid(Sender: TObject);
    procedure DoBrowse(Sender: TObject);
    procedure DoExistingChange(Sender: TObject);
    /// <summary>Reads the currently selected entry of FCmbExisting and
    ///  syncs it into FInfo + the Name edit. Does NOT touch FClb /
    ///  preview - safe to call before those controls exist.</summary>
    procedure SyncExistingSelectionToInfo;
    /// <summary>Re-reads the selected target interface and populates
    ///  FExistingMemberNames. Then re-marks every CheckListBox row
    ///  whose member matches an existing entry: pre-checked + the row
    ///  text gets a "[in interface] " prefix so the user sees what's
    ///  already there.</summary>
    procedure RefreshExistingMemberMarks;
    /// <summary>True if Row index represents a member that is already
    procedure DoPresetAll(Sender: TObject);
    procedure DoPresetPublic(Sender: TObject);
    procedure DoPresetPublished(Sender: TObject);
    procedure DoPresetMethods(Sender: TObject);
    procedure DoPresetProperties(Sender: TObject);
    procedure DoPresetFields(Sender: TObject);
    procedure DoPresetNone(Sender: TObject);

    /// <summary>Toggles the visibility section that begins at the given
    ///  CheckListBox row index: walks forward until the next header (or
    ///  end of list) and sets all member rows in that section to the
    ///  new state. The new state is the inverse of the section's
    ///  current "all checked" status (if all members are already
    ///  checked, unchecks them all; otherwise checks them all).</summary>
    procedure ToggleSection(AHeaderRow: Integer);
    /// <summary>Toggles the rows whose underlying TClassMember matches
    ///  APredicate. If all matching rows are already checked, unchecks
    ///  them; otherwise checks them all. Leaves rows that do not match
    ///  the predicate untouched. This is the semantics the preset
    ///  buttons use: each button only affects its own group.</summary>
    procedure ToggleByPredicate(APredicate: TFunc<TClassMember, Boolean>);
    procedure SetAllChecked(AChecked: Boolean);

    /// <summary>True if the CheckListBox row at AIndex is a visibility
    ///  section header (no underlying class member).</summary>
    function IsHeaderRow(AIndex: Integer): Boolean;
    /// <summary>For a member row, returns the index into FInfo.Members.
    ///  For a header row, returns -1.</summary>
    function MemberIndex(ARow: Integer): Integer;

  public
    constructor CreateDialog(AOwner: TComponent; AMode: TInterfaceMode;
      const AInfo: TExtractInterfaceInfo;
      const AExistingInterfaces: TArray<TInterfaceDeclLocation>);
    destructor Destroy; override;
    function GetResult: TExtractInterfaceInfo;

    class function Choose(AOwner: TComponent; AMode: TInterfaceMode;
      const AInfo: TExtractInterfaceInfo;
      const AExistingInterfaces: TArray<TInterfaceDeclLocation>;
      out AResult: TExtractInterfaceInfo): Boolean;
  end;

implementation

uses
  System.UITypes, System.StrUtils,
  Expert.DialogHelper;

{ TExtractInterfaceDialog }

constructor TExtractInterfaceDialog.CreateDialog(AOwner: TComponent;
  AMode: TInterfaceMode; const AInfo: TExtractInterfaceInfo;
  const AExistingInterfaces: TArray<TInterfaceDeclLocation>);
begin
  inherited CreateNew(AOwner);
  FMode := AMode;
  FInfo := AInfo;
  FExistingList := AExistingInterfaces;
  FSelectionByInterface := TDictionary<string, TArray<Boolean>>.Create;
  if FMode = eimExtractNew then
    Caption := 'Extract interface from ' + FInfo.ClassName
  else
    Caption := 'Add to existing interface (' + FInfo.ClassName + ')';
  Position := poMainFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  ClientWidth := 1000;
  ClientHeight := 620;
  Constraints.MinWidth := 720;
  Constraints.MinHeight := 460;
  KeyPreview := True;

  BuildLayout;
  PopulateList;
  // Seed FCurrentInterfaceKey so the FIRST combo-change persists the
  // initial selection under the correct key.
  if FMode = eimAddToExisting then
    FCurrentInterfaceKey := UpperCase(FInfo.InterfaceName);
  RefreshExistingMemberMarks;
  RefreshPreview;
end;

destructor TExtractInterfaceDialog.Destroy;
begin
  FreeAndNil(FSelectionByInterface);
  inherited;
end;

procedure TExtractInterfaceDialog.BuildLayout;
const
  Pad = 8;
  RowH = 24;
var
  Y, ColLeftW: Integer;
  PnlTop: TPanel;
begin
  // Top panel with name / guid / target controls.
  PnlTop := TPanel.Create(Self);
  PnlTop.Parent := Self;
  PnlTop.Align := alTop;
  PnlTop.Height := 96;
  PnlTop.BevelOuter := bvNone;
  PnlTop.Padding.SetBounds(Pad, Pad, Pad, Pad);

  // Row 1: Interface name | GUID | New GUID button
  Y := Pad;
  FLblName := TLabel.Create(Self); FLblName.Parent := PnlTop;
  FLblName.SetBounds(Pad, Y + 5, 100, 16);
  FLblName.Caption := 'Interface name:';

  FEdtName := TEdit.Create(Self); FEdtName.Parent := PnlTop;
  FEdtName.SetBounds(Pad + 110, Y, 200, RowH);
  if FMode = eimExtractNew then
  begin
    FEdtName.Text := FInfo.InterfaceName;
    FEdtName.OnChange := DoNameChange;
  end
  else
  begin
    FEdtName.Text := '(select below)';
    FEdtName.Enabled := False;
  end;

  FLblGuid := TLabel.Create(Self); FLblGuid.Parent := PnlTop;
  FLblGuid.SetBounds(Pad + 330, Y + 5, 50, 16);
  FLblGuid.Caption := 'GUID:';

  FEdtGuid := TEdit.Create(Self); FEdtGuid.Parent := PnlTop;
  FEdtGuid.SetBounds(Pad + 380, Y, 270, RowH);
  FEdtGuid.Text := FInfo.Guid;
  FEdtGuid.OnChange := DoGuidChange;
  if FMode = eimAddToExisting then
    FEdtGuid.Enabled := False; // keep the existing interface's GUID

  FBtnNewGuid := TButton.Create(Self); FBtnNewGuid.Parent := PnlTop;
  FBtnNewGuid.SetBounds(Pad + 660, Y, 90, RowH);
  FBtnNewGuid.Caption := 'New GUID';
  FBtnNewGuid.OnClick := DoNewGuid;
  if FMode = eimAddToExisting then
    FBtnNewGuid.Enabled := False;

  // Row 2: Target file (extract) or Existing interface picker (extend)
  Inc(Y, RowH + Pad);
  if FMode = eimExtractNew then
  begin
    FLblTarget := TLabel.Create(Self); FLblTarget.Parent := PnlTop;
    FLblTarget.SetBounds(Pad, Y + 5, 100, 16);
    FLblTarget.Caption := 'Target file:';

    FEdtTarget := TEdit.Create(Self); FEdtTarget.Parent := PnlTop;
    FEdtTarget.SetBounds(Pad + 110, Y, 670, RowH);
    FEdtTarget.Text := FInfo.TargetFile;
    FEdtTarget.OnChange := DoTargetChange;

    FBtnBrowse := TButton.Create(Self); FBtnBrowse.Parent := PnlTop;
    FBtnBrowse.SetBounds(Pad + 790, Y, 90, RowH);
    FBtnBrowse.Caption := 'Browse...';
    FBtnBrowse.OnClick := DoBrowse;
  end
  else
  begin
    FLblTarget := TLabel.Create(Self); FLblTarget.Parent := PnlTop;
    FLblTarget.SetBounds(Pad, Y + 5, 100, 16);
    FLblTarget.Caption := 'Target interface:';

    FCmbExisting := TComboBox.Create(Self); FCmbExisting.Parent := PnlTop;
    FCmbExisting.SetBounds(Pad + 110, Y, 670, RowH);
    FCmbExisting.Style := csDropDownList;
    for var Loc in FExistingList do
      FCmbExisting.Items.Add(Loc.InterfaceName + '  -  ' +
        ExtractFileName(Loc.FileName) + '  line ' + IntToStr(Loc.DeclLine));
    if FCmbExisting.Items.Count > 0 then FCmbExisting.ItemIndex := 0;
    FCmbExisting.OnChange := DoExistingChange;
    // Only the FInfo / FEdtName sync here - FClb does not exist yet,
    // so we MUST NOT touch RefreshPreview from inside BuildLayout.
    // The constructor calls RefreshPreview after PopulateList anyway.
    SyncExistingSelectionToInfo;
  end;

  // Members header.
  Inc(Y, RowH + Pad);
  FLblMembers := TLabel.Create(Self); FLblMembers.Parent := PnlTop;
  FLblMembers.SetBounds(Pad, Y, 600, 16);
  FLblMembers.Caption := 'Class members - tick the ones you want exposed via the interface:';

  // Bottom button strip. MUST be its own alBottom panel so the buttons
  // get a reserved area; otherwise alClient panels added later cover
  // them up.
  var PnlBottom: TPanel := TPanel.Create(Self);
  PnlBottom.Parent := Self;
  PnlBottom.Align := alBottom;
  PnlBottom.Height := 44;
  PnlBottom.BevelOuter := bvNone;

  FBtnOk := TButton.Create(Self); FBtnOk.Parent := PnlBottom;
  FBtnOk.SetBounds(PnlBottom.Width - 200, 10, 90, RowH);
  FBtnOk.Anchors := [akRight, akTop];
  FBtnOk.Caption := 'OK';
  FBtnOk.Default := True;
  FBtnOk.ModalResult := mrOk;

  FBtnCancel := TButton.Create(Self); FBtnCancel.Parent := PnlBottom;
  FBtnCancel.SetBounds(PnlBottom.Width - 100, 10, 90, RowH);
  FBtnCancel.Anchors := [akRight, akTop];
  FBtnCancel.Caption := 'Cancel';
  FBtnCancel.Cancel := True;
  FBtnCancel.ModalResult := mrCancel;

  // Preset panel (below the member header).
  FPnlPresets := TPanel.Create(Self);
  FPnlPresets.Parent := Self;
  FPnlPresets.Align := alTop;
  FPnlPresets.Top := PnlTop.Top + PnlTop.Height;
  FPnlPresets.Height := 36;
  FPnlPresets.BevelOuter := bvNone;

  ColLeftW := 100;
  FBtnPresetAll := TButton.Create(Self); FBtnPresetAll.Parent := FPnlPresets;
  FBtnPresetAll.SetBounds(Pad, 6, 90, RowH);
  FBtnPresetAll.Caption := 'All'; FBtnPresetAll.OnClick := DoPresetAll;

  FBtnPresetPublic := TButton.Create(Self); FBtnPresetPublic.Parent := FPnlPresets;
  FBtnPresetPublic.SetBounds(Pad + 95, 6, 90, RowH);
  FBtnPresetPublic.Caption := 'Public+Pub.'; FBtnPresetPublic.OnClick := DoPresetPublic;

  FBtnPresetPublished := TButton.Create(Self); FBtnPresetPublished.Parent := FPnlPresets;
  FBtnPresetPublished.SetBounds(Pad + 190, 6, 90, RowH);
  FBtnPresetPublished.Caption := 'Published'; FBtnPresetPublished.OnClick := DoPresetPublished;

  FBtnPresetMethods := TButton.Create(Self); FBtnPresetMethods.Parent := FPnlPresets;
  FBtnPresetMethods.SetBounds(Pad + 285, 6, 90, RowH);
  FBtnPresetMethods.Caption := 'Methods'; FBtnPresetMethods.OnClick := DoPresetMethods;

  FBtnPresetProperties := TButton.Create(Self); FBtnPresetProperties.Parent := FPnlPresets;
  FBtnPresetProperties.SetBounds(Pad + 380, 6, 90, RowH);
  FBtnPresetProperties.Caption := 'Properties'; FBtnPresetProperties.OnClick := DoPresetProperties;

  FBtnPresetFields := TButton.Create(Self); FBtnPresetFields.Parent := FPnlPresets;
  FBtnPresetFields.SetBounds(Pad + 475, 6, 90, RowH);
  FBtnPresetFields.Caption := 'Fields'; FBtnPresetFields.OnClick := DoPresetFields;

  FBtnPresetNone := TButton.Create(Self); FBtnPresetNone.Parent := FPnlPresets;
  FBtnPresetNone.SetBounds(Pad + 570, 6, 90, RowH);
  FBtnPresetNone.Caption := 'None'; FBtnPresetNone.OnClick := DoPresetNone;
  if ColLeftW = 0 then ; // silence

  // Members + preview side-by-side.
  FClb := TCheckListBox.Create(Self); FClb.Parent := Self;
  FClb.Align := alLeft;
  FClb.Width := ClientWidth div 2;
  FClb.Top := FPnlPresets.Top + FPnlPresets.Height;
  FClb.OnClickCheck := DoCheckClick;
  FClb.Font.Name := 'Consolas';
  FClb.Font.Size := 9;

  var SplitPanel := TPanel.Create(Self);
  SplitPanel.Parent := Self;
  SplitPanel.Align := alClient;
  SplitPanel.BevelOuter := bvNone;
  SplitPanel.Padding.SetBounds(Pad, 0, Pad, 0);

  FLblPreview := TLabel.Create(Self); FLblPreview.Parent := SplitPanel;
  FLblPreview.Align := alTop; FLblPreview.AutoSize := False; FLblPreview.Height := 18;
  FLblPreview.Caption := 'Interface preview:';

  FMemoPreview := TMemo.Create(Self); FMemoPreview.Parent := SplitPanel;
  FMemoPreview.Align := alClient;
  FMemoPreview.ReadOnly := True;
  FMemoPreview.ScrollBars := ssBoth;
  FMemoPreview.WordWrap := False;
  FMemoPreview.Font.Name := 'Consolas';
  FMemoPreview.Font.Size := 10;
end;

procedure TExtractInterfaceDialog.PopulateList;
const
  VisStr: array[TMemberVisibility] of string =
    ('', 'strict private', 'private', 'strict protected', 'protected',
     'public', 'published');
  KindStr: array[TMemberKind] of string = ('field', 'property', 'method');
var
  I, Row: Integer;
  M: TClassMember;
  CurVis: TMemberVisibility;
  FirstSection: Boolean;
  HeaderText, S: string;
begin
  FClb.Items.BeginUpdate;
  try
    FClb.Items.Clear;
    CurVis := mvUnknown;
    FirstSection := True;
    for I := 0 to High(FInfo.Members) do
    begin
      M := FInfo.Members[I];
      if FirstSection or (M.Visibility <> CurVis) then
      begin
        HeaderText := VisStr[M.Visibility];
        if HeaderText = '' then HeaderText := 'default visibility';
        FirstSection := False;
        // Section header row. Objects[Idx] = nil distinguishes it from
        // member rows. Surfaced as a normal-looking line ("private")
        // that the user can click to toggle the whole section. We do
        // NOT use TCheckListBox.Header[] because that suppresses the
        // OnClickCheck event we rely on.
        Row := FClb.Items.AddObject(HeaderText, nil);
        FClb.Checked[Row] := False;
        CurVis := M.Visibility;
      end;
      // Member row. Objects[Row] = Pointer(NativeInt(I + 1)) so we can
      // resolve the row back to its FInfo.Members[] index later. We add
      // 1 because Objects[i] = nil is reserved for headers.
      S := Format('  %-9s  %s', [KindStr[M.Kind], M.Signature]);
      Row := FClb.Items.AddObject(S, TObject(NativeInt(I + 1)));
      FClb.Checked[Row] := M.Selected;
    end;
  finally
    FClb.Items.EndUpdate;
  end;
end;

function TExtractInterfaceDialog.IsHeaderRow(AIndex: Integer): Boolean;
begin
  Result := (AIndex >= 0) and (AIndex < FClb.Items.Count) and
    (FClb.Items.Objects[AIndex] = nil);
end;

function TExtractInterfaceDialog.MemberIndex(ARow: Integer): Integer;
var
  Obj: TObject;
begin
  Result := -1;
  if (ARow < 0) or (ARow >= FClb.Items.Count) then Exit;
  Obj := FClb.Items.Objects[ARow];
  if Obj = nil then Exit;
  Result := NativeInt(Obj) - 1;
end;

procedure TExtractInterfaceDialog.RefreshPreview;
var
  Row, MIdx: Integer;
begin
  if FUpdating then Exit;
  if (FClb = nil) or (FMemoPreview = nil) then Exit;
  // Sync FInfo.Members[].Selected from the list. Header rows are
  // skipped (MemberIndex returns -1).
  for Row := 0 to FClb.Items.Count - 1 do
  begin
    MIdx := MemberIndex(Row);
    if (MIdx >= 0) and (MIdx <= High(FInfo.Members)) then
      FInfo.Members[MIdx].Selected := FClb.Checked[Row];
  end;
  FMemoPreview.Lines.Text := TExtractInterfaceEngine.BuildInterfaceText(FInfo);
end;

procedure TExtractInterfaceDialog.DoCheckClick(Sender: TObject);
var
  Row: Integer;
begin
  Row := FClb.ItemIndex;
  // When the user clicked the checkbox of a header row, treat it as
  // "toggle the whole section". A header row's own Checked state has
  // no meaning - it's just a visual handle. Reset it back to False
  // after we used the click as a trigger.
  if IsHeaderRow(Row) then
  begin
    FUpdating := True;
    try
      FClb.Checked[Row] := False;
    finally
      FUpdating := False;
    end;
    ToggleSection(Row);
    Exit;
  end;
  RefreshPreview;
end;

procedure TExtractInterfaceDialog.DoNameChange(Sender: TObject);
begin
  if FMode = eimExtractNew then
  begin
    FInfo.InterfaceName := Trim(FEdtName.Text);
    // Re-suggest the target file when the user retypes the name and the
    // target field has not been hand-edited (we detect "hand-edited" by
    // comparing the basename heuristically).
    var OldBase := ChangeFileExt(ExtractFileName(FInfo.TargetFile), '');
    if SameText(OldBase, TExtractInterfaceEngine.SuggestInterfaceName(FInfo.ClassName))
       or (FEdtTarget.Modified = False) then
    begin
      FInfo.TargetFile := TExtractInterfaceEngine.SuggestTargetFile(
        FInfo.SourceFile, FInfo.InterfaceName);
      FUpdating := True;
      try FEdtTarget.Text := FInfo.TargetFile; finally FUpdating := False; end;
    end;
    RefreshPreview;
  end;
end;

procedure TExtractInterfaceDialog.DoGuidChange(Sender: TObject);
begin
  FInfo.Guid := FEdtGuid.Text;
  RefreshPreview;
end;

procedure TExtractInterfaceDialog.DoTargetChange(Sender: TObject);
begin
  if not FUpdating then FInfo.TargetFile := FEdtTarget.Text;
end;

procedure TExtractInterfaceDialog.DoNewGuid(Sender: TObject);
begin
  FInfo.Guid := TExtractInterfaceEngine.NewGuidLiteral;
  FEdtGuid.Text := FInfo.Guid;
end;

procedure TExtractInterfaceDialog.DoBrowse(Sender: TObject);
var
  SD: TSaveDialog;
begin
  SD := TSaveDialog.Create(Self);
  try
    SD.DefaultExt := 'pas';
    SD.Filter := 'Pascal source (*.pas)|*.pas';
    SD.FileName := FInfo.TargetFile;
    if SD.Execute then
    begin
      FInfo.TargetFile := SD.FileName;
      FEdtTarget.Text := FInfo.TargetFile;
    end;
  finally
    SD.Free;
  end;
end;

procedure TExtractInterfaceDialog.SyncExistingSelectionToInfo;
var
  Idx: Integer;
begin
  if FExistingList = nil then Exit;
  if FCmbExisting = nil then Exit;
  Idx := FCmbExisting.ItemIndex;
  if (Idx < 0) or (Idx > High(FExistingList)) then Exit;
  FInfo.ExistingFile := FExistingList[Idx].FileName;
  FInfo.ExistingDeclLine := FExistingList[Idx].DeclLine;
  FInfo.ExistingEndLine := FExistingList[Idx].EndLine;
  FInfo.InterfaceName := FExistingList[Idx].InterfaceName;
  if FEdtName <> nil then FEdtName.Text := FInfo.InterfaceName;
end;

procedure TExtractInterfaceDialog.DoExistingChange(Sender: TObject);
var
  Saved: TArray<Boolean>;
  Row, MIdx: Integer;
begin
  // Persist current visible selections to the OUTGOING interface's
  // slot in the per-interface map, then load the INCOMING interface's
  // previously-saved state (or, if none, initialise with defaults).
  if (FCurrentInterfaceKey <> '') and (FClb <> nil) then
  begin
    SetLength(Saved, Length(FInfo.Members));
    for Row := 0 to FClb.Items.Count - 1 do
    begin
      MIdx := MemberIndex(Row);
      if (MIdx >= 0) and (MIdx <= High(Saved)) then
        Saved[MIdx] := FClb.Checked[Row];
    end;
    FSelectionByInterface.AddOrSetValue(FCurrentInterfaceKey, Saved);
  end;

  SyncExistingSelectionToInfo;

  // Switch to the new interface.
  FCurrentInterfaceKey := UpperCase(FInfo.InterfaceName);
  if FSelectionByInterface.TryGetValue(FCurrentInterfaceKey, Saved) and
     (Length(Saved) = Length(FInfo.Members)) then
  begin
    FUpdating := True;
    try
      for Row := 0 to FClb.Items.Count - 1 do
      begin
        MIdx := MemberIndex(Row);
        if (MIdx >= 0) and (MIdx <= High(Saved)) then
          FClb.Checked[Row] := Saved[MIdx];
      end;
    finally
      FUpdating := False;
    end;
  end;

  RefreshExistingMemberMarks;
  RefreshPreview;
end;

procedure TExtractInterfaceDialog.RefreshExistingMemberMarks;
const
  Prefix = '[in interface] ';
var
  Row, MIdx: Integer;
  M: TClassMember;
  S: string;
begin
  if FMode <> eimAddToExisting then Exit;
  if FClb = nil then Exit;

  // 1. Re-parse the existing interface and store its identifier set.
  FExistingMemberNames := nil;
  if (FInfo.ExistingFile <> '') and (FInfo.ExistingEndLine > 0) then
    FExistingMemberNames := TExtractInterfaceEngine.ParseExistingInterfaceNames(
      FInfo.ExistingFile, FInfo.ExistingDeclLine, FInfo.ExistingEndLine);

  // 2. Walk every member row, mark clashes.
  FUpdating := True;
  try
    for Row := 0 to FClb.Items.Count - 1 do
    begin
      MIdx := MemberIndex(Row);
      if MIdx < 0 then Continue;
      M := FInfo.Members[MIdx];
      S := FClb.Items[Row];
      // Strip an earlier "[in interface] " prefix if present.
      if StartsText(Prefix, S) then S := Copy(S, Length(Prefix) + 1, MaxInt);
      if TExtractInterfaceEngine.ClashesWithExisting(M, FExistingMemberNames) then
      begin
        FClb.Items[Row] := Prefix + S;
        FClb.Checked[Row] := True;
        FClb.ItemEnabled[Row] := False;
      end
      else
      begin
        FClb.Items[Row] := S;
        FClb.ItemEnabled[Row] := True;
      end;
    end;
  finally
    FUpdating := False;
  end;
end;

procedure TExtractInterfaceDialog.ToggleSection(AHeaderRow: Integer);
var
  Row, AllChecked: Integer;
  HasAny: Boolean;
  NewState: Boolean;
begin
  // Determine state: if every member row in this section is already
  // checked, the toggle unchecks them all; otherwise it checks them
  // all. (Standard tri-state-style toggle behaviour.)
  HasAny := False;
  AllChecked := 1;
  Row := AHeaderRow + 1;
  while (Row < FClb.Items.Count) and not IsHeaderRow(Row) do
  begin
    HasAny := True;
    if not FClb.Checked[Row] then AllChecked := 0;
    Inc(Row);
  end;
  if not HasAny then Exit;
  NewState := AllChecked = 0;

  FUpdating := True;
  try
    Row := AHeaderRow + 1;
    while (Row < FClb.Items.Count) and not IsHeaderRow(Row) do
    begin
      FClb.Checked[Row] := NewState;
      Inc(Row);
    end;
  finally
    FUpdating := False;
  end;
  RefreshPreview;
end;

procedure TExtractInterfaceDialog.ToggleByPredicate(
  APredicate: TFunc<TClassMember, Boolean>);
var
  Row, MIdx: Integer;
  AllChecked, HasAny: Boolean;
  NewState: Boolean;
begin
  // Same "if all checked then uncheck else check" toggle, but only
  // for rows whose underlying TClassMember satisfies APredicate. Other
  // rows are left strictly as-is - this is what makes preset buttons
  // additive instead of exclusive.
  HasAny := False;
  AllChecked := True;
  for Row := 0 to FClb.Items.Count - 1 do
  begin
    MIdx := MemberIndex(Row);
    if MIdx < 0 then Continue;
    if not APredicate(FInfo.Members[MIdx]) then Continue;
    HasAny := True;
    if not FClb.Checked[Row] then begin AllChecked := False; Break; end;
  end;
  if not HasAny then Exit;
  NewState := not AllChecked;

  FUpdating := True;
  try
    for Row := 0 to FClb.Items.Count - 1 do
    begin
      MIdx := MemberIndex(Row);
      if MIdx < 0 then Continue;
      if APredicate(FInfo.Members[MIdx]) then
        FClb.Checked[Row] := NewState;
    end;
  finally
    FUpdating := False;
  end;
  RefreshPreview;
end;

procedure TExtractInterfaceDialog.SetAllChecked(AChecked: Boolean);
var
  Row, MIdx: Integer;
begin
  FUpdating := True;
  try
    for Row := 0 to FClb.Items.Count - 1 do
    begin
      MIdx := MemberIndex(Row);
      if MIdx >= 0 then FClb.Checked[Row] := AChecked;
    end;
  finally
    FUpdating := False;
  end;
  RefreshPreview;
end;

procedure TExtractInterfaceDialog.DoPresetAll(Sender: TObject);
var
  Row, MIdx: Integer;
  AllChecked: Boolean;
begin
  // "All" is additive too: tick everything; if everything is already
  // ticked, untick everything.
  AllChecked := True;
  for Row := 0 to FClb.Items.Count - 1 do
  begin
    MIdx := MemberIndex(Row);
    if (MIdx >= 0) and not FClb.Checked[Row] then
    begin
      AllChecked := False; Break;
    end;
  end;
  SetAllChecked(not AllChecked);
end;

procedure TExtractInterfaceDialog.DoPresetPublic(Sender: TObject);
begin
  ToggleByPredicate(function(M: TClassMember): Boolean
    begin Result := M.Visibility in [mvPublic, mvPublished]; end);
end;

procedure TExtractInterfaceDialog.DoPresetPublished(Sender: TObject);
begin
  ToggleByPredicate(function(M: TClassMember): Boolean
    begin Result := M.Visibility = mvPublished; end);
end;

procedure TExtractInterfaceDialog.DoPresetMethods(Sender: TObject);
begin
  ToggleByPredicate(function(M: TClassMember): Boolean
    begin Result := M.Kind = mkMethod; end);
end;

procedure TExtractInterfaceDialog.DoPresetProperties(Sender: TObject);
begin
  ToggleByPredicate(function(M: TClassMember): Boolean
    begin Result := M.Kind = mkProperty; end);
end;

procedure TExtractInterfaceDialog.DoPresetFields(Sender: TObject);
begin
  ToggleByPredicate(function(M: TClassMember): Boolean
    begin Result := M.Kind = mkField; end);
end;

procedure TExtractInterfaceDialog.DoPresetNone(Sender: TObject);
begin
  SetAllChecked(False);
end;

function TExtractInterfaceDialog.GetResult: TExtractInterfaceInfo;
var
  Row, MIdx, I: Integer;
  Saved: TArray<Boolean>;
  Loc: TInterfaceDeclLocation;
  Target: TInterfaceTarget;
begin
  // Final sync from list rows back into FInfo.Members AND into the
  // per-interface selection map (so the currently visible interface
  // gets persisted with the others).
  for Row := 0 to FClb.Items.Count - 1 do
  begin
    MIdx := MemberIndex(Row);
    if (MIdx >= 0) and (MIdx <= High(FInfo.Members)) then
      FInfo.Members[MIdx].Selected := FClb.Checked[Row];
  end;
  if (FMode = eimAddToExisting) and (FCurrentInterfaceKey <> '') then
  begin
    SetLength(Saved, Length(FInfo.Members));
    for I := 0 to High(FInfo.Members) do
      Saved[I] := FInfo.Members[I].Selected;
    FSelectionByInterface.AddOrSetValue(FCurrentInterfaceKey, Saved);
  end;

  FInfo.Mode := FMode;

  // For add-to-existing: build the Targets array from every interface
  // for which the user picked at least one member. A target with no
  // selection is dropped here so ApplyAddToExisting can iterate
  // Targets directly without per-entry filtering.
  if FMode = eimAddToExisting then
  begin
    FInfo.Targets := nil;
    for Loc in FExistingList do
    begin
      if not FSelectionByInterface.TryGetValue(
          UpperCase(Loc.InterfaceName), Saved) then Continue;
      if Length(Saved) <> Length(FInfo.Members) then Continue;
      var Any: Boolean := False;
      for I := 0 to High(Saved) do
        if Saved[I] then begin Any := True; Break; end;
      if not Any then Continue;
      Target.InterfaceName := Loc.InterfaceName;
      Target.FileName := Loc.FileName;
      Target.DeclLine := Loc.DeclLine;
      Target.EndLine := Loc.EndLine;
      Target.MemberSelected := Saved;
      FInfo.Targets := FInfo.Targets + [Target];
    end;
  end;
  Result := FInfo;
end;

class function TExtractInterfaceDialog.Choose(AOwner: TComponent;
  AMode: TInterfaceMode; const AInfo: TExtractInterfaceInfo;
  const AExistingInterfaces: TArray<TInterfaceDeclLocation>;
  out AResult: TExtractInterfaceInfo): Boolean;
var
  Dlg: TExtractInterfaceDialog;
begin
  Dlg := TExtractInterfaceDialog.CreateDialog(AOwner, AMode, AInfo, AExistingInterfaces);
  try
    Result := Dlg.ShowModal = mrOk;
    if Result then AResult := Dlg.GetResult;
  finally
    Dlg.Free;
  end;
end;

end.
