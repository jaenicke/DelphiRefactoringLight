(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindReferencesDialog;

interface

uses
  System.SysUtils, System.Classes, System.UITypes, System.Generics.Collections, Winapi.Windows, Winapi.Messages,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls;

type
  TFindReferenceItem = record
    FilePath: string;
    Line: Integer;     // 0-based (LSP)
    Col: Integer;      // 0-based
    Length: Integer;   // Length of the identifier
    Preview: string;   // Line containing the match
    /// <summary>'' = verified; otherwise why it could not be verified (the
    ///  hit is shown anyway - hiding it would look like "no reference").</summary>
    Note: string;
    /// <summary>How the hit belongs to the symbol when it is not the symbol
    ///  itself: "declared in interface IFoo", "call via interface IFoo",
    ///  "implemented by TFoo", "call via class TFoo". '' otherwise.</summary>
    Relation: string;
    /// <summary>How the symbol is used there - "Call", "Write", "Read",
    ///  "Declaration", "Implementation", "Type use", ... (see
    ///  Expert.ReferenceKind); '' when not classified.</summary>
    Kind: string;
  end;

  TFindReferenceItems = TArray<TFindReferenceItem>;

/// <summary>Fills Kind of every item: the symbol's own kind comes from its
///  declaration line (ADeclFile / ADeclLine, '' / -1 = unknown), the items
///  are classified per file in one pass. AReadContent returns a file's
///  current content ('' when unreadable) - editor buffer or disk, as the
///  caller's thread allows.</summary>
procedure AssignReferenceKinds(var AItems: TFindReferenceItems; const AName,
  ADeclFile: string; ADeclLine: Integer; const AReadContent: TFunc<string, string>);

type
  /// <summary>Modal dialog with a ListView of all occurrences of an identifier.
  ///  Double-click = jump to location (dialog stays open).
  ///  ENTER on a line = jump to location and close dialog.
  ///  ESC / Close button = close dialog.</summary>
  TFindReferencesDialog = class(TForm)
  private
    FStatusLabel: TLabel;
    FProgress: TProgressBar;
    FListView: TListView;
    FBtnGoto: TButton;
    FBtnClose: TButton;
    FItems: TFindReferenceItems;
    FOnGotoLocation: TProc<TFindReferenceItem>;
    FOnDialogClose: TNotifyEvent;
    FAllowFree: Boolean;
    FCloseRequested: Boolean;

    procedure CreateControls;
    procedure DoListDblClick(Sender: TObject);
    procedure DoBtnGotoClick(Sender: TObject);
    procedure DoBtnCloseClick(Sender: TObject);
    procedure DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function CommonPathPrefix(const AItems: TFindReferenceItems): string;
    procedure GotoSelected;
  public
    constructor CreateDialog(AOwner: TComponent; const AIdentifier: string;
      const ATitlePrefix: string = 'References'); reintroduce;

    /// <summary>Sets the matches and fills the ListView. Replaced previous items.</summary>
    procedure SetItems(const AItems: TFindReferenceItems);

    procedure SetStatus(const AText: string);
    procedure SetProgress(ACurrent, ATotal: Integer);

    /// <summary>Switches the dialog into "review mode": from now on,
    ///  closing the dialog actually frees it.</summary>
    procedure SetClosable;

    property OnGotoLocation: TProc<TFindReferenceItem> read FOnGotoLocation write FOnGotoLocation;
    /// <summary>Fired right before the dialog is destroyed.</summary>
    property OnDialogClose: TNotifyEvent read FOnDialogClose write FOnDialogClose;
  end;

implementation

uses
  System.IOUtils, System.Types, Expert.IdeThemes, Expert.DialogHelper, Expert.ListViewSort,
  Expert.ReferenceKind;

procedure AssignReferenceKinds(var AItems: TFindReferenceItems; const AName,
  ADeclFile: string; ADeclLine: Integer; const AReadContent: TFunc<string, string>);
var
  Cache: TDictionary<string, string>;
  Groups: TDictionary<string, TList<Integer>>;

  function Get(const AFile: string): string;
  begin
    if not Cache.TryGetValue(UpperCase(AFile), Result) then
    begin
      Result := '';
      if Assigned(AReadContent) then
        try
          Result := AReadContent(AFile);
        except
          Result := '';
        end;
      Cache.Add(UpperCase(AFile), Result);
    end;
  end;

begin
  if (Length(AItems) = 0) or (AName = '') then Exit;
  Cache := TDictionary<string, string>.Create;
  Groups := TDictionary<string, TList<Integer>>.Create;
  try
    var Sym := rsUnknown;
    if (ADeclFile <> '') and (ADeclLine >= 0) then
    begin
      var DL := Get(ADeclFile).Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
      if ADeclLine <= High(DL) then Sym := SymbolKindFromDeclLine(DL[ADeclLine], AName);
    end;
    for var I := 0 to High(AItems) do
    begin
      var K := UpperCase(AItems[I].FilePath);
      var L: TList<Integer>;
      if not Groups.TryGetValue(K, L) then
      begin
        L := TList<Integer>.Create;
        Groups.Add(K, L);
      end;
      L.Add(I);
    end;
    for var G in Groups do
    begin
      var Idx := G.Value.ToArray;
      var Content := Get(AItems[Idx[0]].FilePath);
      if Content = '' then Continue;
      var Pts: TArray<TPoint>;
      SetLength(Pts, Length(Idx));
      for var K := 0 to High(Idx) do
        Pts[K] := Point(AItems[Idx[K]].Col, AItems[Idx[K]].Line);
      var Kinds := ClassifyReferences(Content, Pts, Length(AName), Sym);
      for var K := 0 to High(Idx) do
        if K <= High(Kinds) then AItems[Idx[K]].Kind := RefKindText(Kinds[K]);
    end;
  finally
    for var G in Groups do G.Value.Free;
    Groups.Free;
    Cache.Free;
  end;
end;

{ TFindReferencesDialog }

constructor TFindReferencesDialog.CreateDialog(AOwner: TComponent; const AIdentifier: string; const ATitlePrefix: string);
begin
  inherited CreateNew(AOwner);
  Caption := ATitlePrefix + ': ' + AIdentifier;
  Position := poOwnerFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  Width := 800;
  Height := 500;
  Constraints.MinWidth := 480;
  Constraints.MinHeight := 280;
  KeyPreview := True;
  OnKeyDown := DoFormKeyDown;
  OnClose := DoFormClose;

  CreateControls;
  EnableListViewSorting(FListView);
  Expert.IdeThemes.EnableThemes(Self);

  PrepareDialog(Self, AOwner);
end;

procedure TFindReferencesDialog.CreateControls;
var
  Col: TListColumn;
begin
  // Status line at top
  FStatusLabel := TLabel.Create(Self);
  FStatusLabel.Parent := Self;
  FStatusLabel.Align := alTop;
  FStatusLabel.AlignWithMargins := True;
  FStatusLabel.Margins.SetBounds(8, 8, 8, 4);
  FStatusLabel.Caption := 'Searching...';

  // Progress
  FProgress := TProgressBar.Create(Self);
  FProgress.Parent := Self;
  FProgress.Align := alTop;
  FProgress.AlignWithMargins := True;
  FProgress.Margins.SetBounds(8, 0, 8, 4);
  FProgress.Height := 16;
  FProgress.Min := 0;
  FProgress.Max := 100;
  FProgress.Position := 0;

  // Button panel at bottom
  var BtnPanel := TPanel.Create(Self);
  BtnPanel.Parent := Self;
  BtnPanel.Align := alBottom;
  BtnPanel.Height := 40;
  BtnPanel.BevelOuter := bvNone;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := BtnPanel;
  FBtnClose.Caption := 'Close';
  FBtnClose.Width := 100;
  FBtnClose.Height := 28;
  FBtnClose.Top := 6;
  FBtnClose.Anchors := [akTop, akRight];
  FBtnClose.Left := BtnPanel.Width - FBtnClose.Width - 8;
  FBtnClose.OnClick := DoBtnCloseClick;
  FBtnClose.Cancel := True;

  FBtnGoto := TButton.Create(Self);
  FBtnGoto.Parent := BtnPanel;
  FBtnGoto.Caption := 'Go To';
  FBtnGoto.Width := 100;
  FBtnGoto.Height := 28;
  FBtnGoto.Top := 6;
  FBtnGoto.Anchors := [akTop, akRight];
  FBtnGoto.Left := FBtnClose.Left - FBtnGoto.Width - 6;
  FBtnGoto.OnClick := DoBtnGotoClick;
  FBtnGoto.Default := True;
  FBtnGoto.Enabled := False;

  // ListView in the middle
  FListView := TListView.Create(Self);
  FListView.Parent := Self;
  FListView.Align := alClient;
  FListView.AlignWithMargins := True;
  FListView.Margins.SetBounds(8, 4, 8, 4);
  FListView.ViewStyle := vsReport;
  FListView.RowSelect := True;
  FListView.ReadOnly := True;
  FListView.HideSelection := False;
  FListView.GridLines := True;
  FListView.OnDblClick := DoListDblClick;
  FListView.OnKeyDown := DoListKeyDown;

  Col := FListView.Columns.Add;
  Col.Caption := 'File';
  Col.Width := 220;

  Col := FListView.Columns.Add;
  Col.Caption := 'Line';
  Col.Width := 60;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Column';
  Col.Width := 60;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Kind';
  Col.Width := 110;

  Col := FListView.Columns.Add;
  Col.Caption := 'Preview';
  Col.Width := 420;

  Col := FListView.Columns.Add;
  Col.Caption := 'Note';
  Col.Width := 200;
end;

function TFindReferencesDialog.CommonPathPrefix(
  const AItems: TFindReferenceItems): string;
var
  P: string;
  I, J, MaxLen: Integer;
begin
  Result := '';
  if Length(AItems) = 0 then Exit;
  Result := ExtractFilePath(AItems[0].FilePath);
  for I := 1 to High(AItems) do
  begin
    P := ExtractFilePath(AItems[I].FilePath);
    MaxLen := Length(Result);
    if Length(P) < MaxLen then MaxLen := Length(P);
    J := 1;
    while (J <= MaxLen) and (UpCase(Result[J]) = UpCase(P[J])) do
      Inc(J);
    Result := Copy(Result, 1, J - 1);
    if Result = '' then Exit;
  end;
end;

procedure TFindReferencesDialog.SetItems(const AItems: TFindReferenceItems);
var
  I: Integer;
  LI: TListItem;
  Prefix, DisplayPath: string;
begin
  FItems := AItems;
  Prefix := CommonPathPrefix(AItems);

  FListView.Items.BeginUpdate;
  try
    FListView.Clear;
    for I := 0 to High(AItems) do
    begin
      LI := FListView.Items.Add;
      LI.Data := Pointer(NativeInt(I));  // FItems index; survives sorting
      DisplayPath := AItems[I].FilePath;
      if (Prefix <> '') and DisplayPath.StartsWith(Prefix, True) then
        DisplayPath := Copy(DisplayPath, Length(Prefix) + 1, MaxInt);
      LI.Caption := DisplayPath;
      LI.SubItems.Add(IntToStr(AItems[I].Line + 1));
      LI.SubItems.Add(IntToStr(AItems[I].Col + 1));
      LI.SubItems.Add(AItems[I].Kind);
      LI.SubItems.Add(AItems[I].Preview);
      if (AItems[I].Relation <> '') and (AItems[I].Note <> '') then
        LI.SubItems.Add(AItems[I].Relation + '; ' + AItems[I].Note)
      else
        LI.SubItems.Add(AItems[I].Relation + AItems[I].Note);
    end;
    if FListView.Items.Count > 0 then
    begin
      FListView.Items[0].Selected := True;
      FListView.Items[0].Focused := True;
    end;
  finally
    FListView.Items.EndUpdate;
  end;

  FBtnGoto.Enabled := Length(AItems) > 0;
end;

procedure TFindReferencesDialog.SetStatus(const AText: string);
begin
  FStatusLabel.Caption := AText;
end;

procedure TFindReferencesDialog.SetProgress(ACurrent, ATotal: Integer);
begin
  if ATotal <= 0 then
  begin
    FProgress.Position := 0;
    Exit;
  end;
  FProgress.Max := ATotal;
  if ACurrent > ATotal then
    ACurrent := ATotal;
  FProgress.Position := ACurrent;
end;

procedure TFindReferencesDialog.GotoSelected;
var
  Idx: Integer;
begin
  if not Assigned(FListView.Selected) then Exit;
  Idx := NativeInt(FListView.Selected.Data);
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if Assigned(FOnGotoLocation) then
    FOnGotoLocation(FItems[Idx]);
end;

procedure TFindReferencesDialog.DoListDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TFindReferencesDialog.DoBtnGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TFindReferencesDialog.DoBtnCloseClick(Sender: TObject);
begin
  // Defensive: route through DoFormClose-style logic directly. The
  // plain Close call has been reported as unresponsive while the
  // synchronous search loop is running - hiding the form here makes
  // the user feedback immediate; the actual free happens later when
  // SetClosable kicks in (or right away if we're already past it).
  FCloseRequested := True;
  Hide;
  if FAllowFree then
  begin
    if Assigned(FOnDialogClose) then FOnDialogClose(Self);
    Release;  // queued free; safe even when called from a button handler
  end;
end;

procedure TFindReferencesDialog.DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Close;
    Key := 0;
  end;
end;

procedure TFindReferencesDialog.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  if not FAllowFree then
  begin
    FCloseRequested := True;
    Hide;
    Action := caNone;
    Exit;
  end;
  if Assigned(FOnDialogClose) then FOnDialogClose(Self);
  Action := caFree;
end;

procedure TFindReferencesDialog.SetClosable;
begin
  FAllowFree := True;
  if FCloseRequested then
    Close;
end;

procedure TFindReferencesDialog.DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    GotoSelected;
    Key := 0;
  end;
end;

initialization
  RegisterDialogClass(TFindReferencesDialog);

end.
