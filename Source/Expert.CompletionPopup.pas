(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.CompletionPopup;

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.Generics.Collections, Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Graphics, Winapi.Windows, Winapi.Messages;

type
  TCompletionItem = record
    Label_: string;
    Detail: string;
    Kind: Integer;     // LSP CompletionItemKind
    SortText: string;
  end;

  TCompletionInsertEvent = procedure(const AText: string) of object;

  /// <summary>Helpers for parsing LSP completion payloads. Grouped into
  ///  a class to keep the unit free of global routines.</summary>
  TCompletionItems = class
  public
    /// <summary>Parses the LSP completion response into a flat array
    ///  of TCompletionItem records.</summary>
    class function Parse(AResponse: TJSONObject): TArray<TCompletionItem>; static;

    /// <summary>Returns a short textual icon-ish prefix for the given LSP
    ///  CompletionItemKind ('fn', 'var', 'prop', ...).</summary>
    class function KindToPrefix(AKind: Integer): string; static;
  end;

  /// <summary>Popup window for code completion. Shows a filterable list
  ///  of label + detail entries.</summary>
  TCompletionPopup = class(TForm)
  private
    FListBox: TListBox;
    FDetailLabel: TLabel;
    FAllItems: TArray<TCompletionItem>;
    FFilteredItems: TList<TCompletionItem>;
    FOnInsert: TCompletionInsertEvent;
    /// <summary>Current filter prefix. Driven externally by the
    ///  caller's editor input via SetPrefix - the popup itself does
    ///  not own an Edit control any more, focus stays with the
    ///  editor so the user can keep typing normally.</summary>
    FFilterText: string;
    /// <summary>Ticks while the popup shows "Loading...". Updates the
    ///  list item with the elapsed seconds and the host-supplied
    ///  status string (typically GetWarmupStatusLine from the LSP
    ///  manager) so the user sees that work is happening even when
    ///  the LSP itself takes 20+ seconds to respond.</summary>
    FLoadingTimer: TTimer;
    FLoadingStart: TDateTime;
    FLoadingStatusFn: TFunc<string>;
    procedure DoListBoxDblClick(Sender: TObject);
    procedure DoListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure DoLoadingTick(Sender: TObject);
    procedure UpdateLoadingLine;
    procedure ApplyFilter;
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    constructor CreatePopup(AOwner: TComponent);
    destructor Destroy; override;

    /// <summary>Shows the popup immediately with "Loading..." at the
    ///  given screen position.</summary>
    procedure ShowLoading(X, Y: Integer);
    /// <summary>Replaces "Loading..." with the actual items. If
    ///  AInitialFilter is non-empty, the filter is pre-filled with it
    ///  (used when the cursor sits inside a partially typed identifier).</summary>
    procedure ShowItems(const AItems: TArray<TCompletionItem>; const AInitialFilter: string = '');
    /// <summary>Shows a message instead of items (e.g. errors).</summary>
    procedure ShowMessage(const AMsg: string);

    /// <summary>Updates the filter as the user keeps typing in the
    ///  editor. Re-applies the prefix locally and updates the list.
    ///  Empty prefix = no filter (everything shown).</summary>
    procedure SetPrefix(const APrefix: string);
    /// <summary>Moves the selection by ADelta items (negative for up).
    ///  Clamped to range. Called by the host editor when the user
    ///  presses Up/Down while the popup is visible.</summary>
    procedure MoveSelection(ADelta: Integer);
    /// <summary>Commits the currently selected item: fires OnInsert
    ///  with its label, then closes the popup. Called by the host
    ///  editor on Enter / Tab.</summary>
    procedure InsertSelected;
    /// <summary>Closes the popup. Public so the editor can dismiss it
    ///  on Escape / non-word character / focus loss.</summary>
    procedure HidePopup;
    /// <summary>True when the popup is visible AND has items (i.e. the
    ///  LSP response has arrived). During the Loading state this is
    ///  False so the host editor still gets Enter / Up / Down for
    ///  normal keystroke handling.</summary>
    function IsActive: Boolean;

    property OnInsert: TCompletionInsertEvent read FOnInsert write FOnInsert;
    /// <summary>Optional callback invoked while the popup is in its
    ///  Loading state. Whatever the callback returns is appended to
    ///  the "Loading... (Xs)" line in the list, e.g.
    ///  "Loading suggestions... (12s) - LSP starting - indexing".
    ///  Wire it to TLspManager.Instance.GetWarmupStatusLine.</summary>
    property OnLoadingStatus: TFunc<string> read FLoadingStatusFn write FLoadingStatusFn;
  end;

implementation

uses
  Expert.IdeThemes;

{ TCompletionItems }

class function TCompletionItems.KindToPrefix(AKind: Integer): string;
begin
  case AKind of
    1:  Result := 'txt';  // Text
    2:  Result := 'fn';   // Method
    3:  Result := 'fn';   // Function
    4:  Result := 'ctor'; // Constructor
    5:  Result := 'fld';  // Field
    6:  Result := 'var';  // Variable
    7:  Result := 'cls';  // Class
    8:  Result := 'intf'; // Interface
    9:  Result := 'mod';  // Module
    10: Result := 'prop'; // Property
    11: Result := 'unit'; // Unit
    12: Result := 'val';  // Value
    13: Result := 'enum'; // Enum
    14: Result := 'kw';   // Keyword
    15: Result := 'snip'; // Snippet
    22: Result := 'type'; // Struct (= Record)
    25: Result := 'type'; // TypeParameter
  else
    Result := '   ';
  end;
end;

class function TCompletionItems.Parse(AResponse: TJSONObject): TArray<TCompletionItem>;
var
  ResultVal: TJSONValue;
  ItemsArr: TJSONArray;
  ItemObj: TJSONObject;
  Item: TCompletionItem;
  ItemList: TList<TCompletionItem>;
  I: Integer;
begin
  ItemList := TList<TCompletionItem>.Create;
  try
    ResultVal := AResponse.GetValue('result');
    if ResultVal = nil then
      Exit(nil);

    // The result may be an object with an "items" array, or the array directly.
    if ResultVal is TJSONObject then
    begin
      if not TJSONObject(ResultVal).TryGetValue<TJSONArray>('items', ItemsArr) then
        Exit(nil);
    end
    else if ResultVal is TJSONArray then
      ItemsArr := TJSONArray(ResultVal)
    else
      Exit(nil);

    for I := 0 to ItemsArr.Count - 1 do
    begin
      if not (ItemsArr.Items[I] is TJSONObject) then Continue;
      ItemObj := TJSONObject(ItemsArr.Items[I]);

      Item.Label_ := ItemObj.GetValue<string>('label', '');
      Item.Detail := ItemObj.GetValue<string>('detail', '');
      Item.Kind := ItemObj.GetValue<Integer>('kind', 1);
      Item.SortText := ItemObj.GetValue<string>('sortText', Item.Label_);
      ItemList.Add(Item);
    end;

    Result := ItemList.ToArray;
  finally
    ItemList.Free;
  end;
end;

{ TCompletionPopup }

constructor TCompletionPopup.CreatePopup(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);
  BorderStyle := bsNone;
  FormStyle := fsStayOnTop;
  Width := 450;
  Height := 320;
  Color := GetThemedColor(clWindow);
  // No more OnDeactivate / KeyPreview / OnKeyDown:
  //   * The popup never takes focus (WS_EX_NOACTIVATE + ShowWindow
  //     SW_SHOWNOACTIVATE), so OnDeactivate would either never fire or
  //     fire as the user moves between unrelated windows.
  //   * Key input is handled by the host editor; when the popup is
  //     visible, the host intercepts Up/Down/Enter/Escape and routes
  //     to the public MoveSelection / InsertSelected / HidePopup
  //     methods. Other keys keep typing into the editor and the host
  //     refreshes the filter via SetPrefix.

  FFilteredItems := TList<TCompletionItem>.Create;

  // Detail-Label unten
  FDetailLabel := TLabel.Create(Self);
  FDetailLabel.Parent := Self;
  FDetailLabel.Align := alBottom;
  FDetailLabel.Height := 22;
  FDetailLabel.Font.Name := 'Consolas';
  FDetailLabel.Font.Size := 9;
  FDetailLabel.Font.Color := GetThemedColor(clGrayText);
  FDetailLabel.AutoSize := False;
  FDetailLabel.Caption := '';
  FDetailLabel.Layout := tlCenter;

  // ListBox
  FListBox := TListBox.Create(Self);
  FListBox.Parent := Self;
  FListBox.Align := alClient;
  FListBox.Font.Name := 'Consolas';
  FListBox.Font.Size := 10;
  FListBox.Style := lbOwnerDrawFixed;
  FListBox.ItemHeight := 20;
  FListBox.OnDblClick := DoListBoxDblClick;
  FListBox.OnDrawItem := DoListBoxDrawItem;
  // Disabling the list box prevents it from taking focus / getting
  // selection ticks; we drive ItemIndex programmatically.
  // (TListBox does not have a TabStop = False default for owner-draw
  // mode.)
  FListBox.TabStop := False;

  Expert.IdeThemes.EnableThemes(Self);

  FLoadingTimer := TTimer.Create(Self);
  FLoadingTimer.Enabled := False;
  FLoadingTimer.Interval := 500;
  FLoadingTimer.OnTimer := DoLoadingTick;
end;

destructor TCompletionPopup.Destroy;
begin
  FFilteredItems.Free;
  inherited;
end;

procedure TCompletionPopup.CreateParams(var Params: TCreateParams);
begin
  inherited;
  // WS_EX_NOACTIVATE keeps the editor focused when the popup is
  // shown - that is what lets the user keep typing.
  Params.ExStyle := Params.ExStyle or WS_EX_NOACTIVATE or WS_EX_TOOLWINDOW;
end;

procedure TCompletionPopup.ShowLoading(X, Y: Integer);
begin
  SetLength(FAllItems, 0);
  FFilterText := '';
  FListBox.Items.Clear;
  FListBox.Items.Add('Loading suggestions...');
  FDetailLabel.Caption := '';

  Left := X;
  Top := Y;
  if Left + Width > Screen.Width then
    Left := Screen.Width - Width;
  if Top + Height > Screen.Height then
    Top := Y - Height - 20;

  // SW_SHOWNOACTIVATE preserves the editor's focus.
  ShowWindow(Handle, SW_SHOWNOACTIVATE);
  Visible := True;
  FLoadingStart := Now;
  FLoadingTimer.Enabled := True;
  UpdateLoadingLine;
end;

procedure TCompletionPopup.DoLoadingTick(Sender: TObject);
begin
  UpdateLoadingLine;
end;

procedure TCompletionPopup.UpdateLoadingLine;
// Refreshes the single "Loading..." entry with elapsed seconds and -
// if the caller supplied a status function - the LSP-side status it
// returns. Updating an existing TListBox item in place avoids the
// FreeAndNil + flicker of a Clear+Add cycle.
var
  Sec: Integer;
  StatusText, Line: string;
begin
  if FListBox.Items.Count = 0 then Exit;
  Sec := Round((Now - FLoadingStart) * SecsPerDay);
  Line := Format('Loading suggestions... (%d s)', [Sec]);
  if Assigned(FLoadingStatusFn) then
  try
    StatusText := FLoadingStatusFn();
  except
    StatusText := '';
  end;
  if StatusText <> '' then
    Line := Line + '  -  ' + StatusText;
  FListBox.Items[0] := Line;
end;

procedure TCompletionPopup.ShowItems(const AItems: TArray<TCompletionItem>;
  const AInitialFilter: string);
begin
  FLoadingTimer.Enabled := False;
  FAllItems := AItems;
  // If the host pushed a fresher prefix via SetPrefix while the LSP
  // was loading (user kept typing), respect that; otherwise the
  // wizard's initial filter.
  if FFilterText = '' then
    FFilterText := AInitialFilter;
  ApplyFilter;
  // No focus grab - editor keeps focus, user keeps typing. The popup
  // is already visible (ShowLoading); just re-render.
  if not Visible then
  begin
    ShowWindow(Handle, SW_SHOWNOACTIVATE);
    Visible := True;
  end;
end;

procedure TCompletionPopup.ShowMessage(const AMsg: string);
begin
  FLoadingTimer.Enabled := False;
  SetLength(FAllItems, 0);
  FFilteredItems.Clear;
  FListBox.Items.Clear;
  FListBox.Items.Add(AMsg);
  FDetailLabel.Caption := '';
end;

procedure TCompletionPopup.SetPrefix(const APrefix: string);
begin
  if APrefix = FFilterText then Exit;
  FFilterText := APrefix;
  // Skip during the loading phase - applying a filter would replace
  // the "Loading suggestions..." entry with the empty filtered list.
  if Length(FAllItems) > 0 then
    ApplyFilter;
end;

function TCompletionPopup.IsActive: Boolean;
begin
  Result := Visible and (Length(FAllItems) > 0);
end;

procedure TCompletionPopup.MoveSelection(ADelta: Integer);
var
  NewIdx: Integer;
begin
  if FFilteredItems.Count = 0 then Exit;
  NewIdx := FListBox.ItemIndex + ADelta;
  if NewIdx < 0 then NewIdx := 0;
  if NewIdx >= FFilteredItems.Count then NewIdx := FFilteredItems.Count - 1;
  FListBox.ItemIndex := NewIdx;
  FDetailLabel.Caption := '  ' + FFilteredItems[NewIdx].Detail;
end;

procedure TCompletionPopup.HidePopup;
begin
  FLoadingTimer.Enabled := False;
  if Visible then Hide;
end;

procedure TCompletionPopup.ApplyFilter;
var
  Filter: string;
  Item: TCompletionItem;
begin
  Filter := UpperCase(FFilterText);
  FFilteredItems.Clear;
  FListBox.Items.BeginUpdate;
  try
    FListBox.Items.Clear;
    for Item in FAllItems do
    begin
      // Prefix match (case-insensitive) - mirrors Delphi's built-in
      // Code Insight behavior: typing characters narrows the list to
      // identifiers starting with those characters.
      if (Filter = '') or UpperCase(Item.Label_).StartsWith(Filter) then
      begin
        FFilteredItems.Add(Item);
        FListBox.Items.Add(Item.Label_);
      end;
    end;
  finally
    FListBox.Items.EndUpdate;
  end;

  if FListBox.Items.Count > 0 then
    FListBox.ItemIndex := 0;
end;

procedure TCompletionPopup.DoListBoxDblClick(Sender: TObject);
begin
  InsertSelected;
end;

procedure TCompletionPopup.DoListBoxDrawItem(Control: TWinControl;
  Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  Canvas: TCanvas;
  Item: TCompletionItem;
  Prefix, LabelText, DetailText, FallbackText: string;
begin
  Canvas := FListBox.Canvas;

  // The ListBox is owner-draw, so we are responsible for every pixel.
  // While the popup is showing "Loading..." or a one-line message,
  // FFilteredItems is empty but FListBox.Items has one entry. Render
  // that entry as plain text so the user actually sees the status -
  // otherwise the slot stays the bare background, which used to be
  // confusing ("the popup is empty").
  if Index >= FFilteredItems.Count then
  begin
    if odSelected in State then
      Canvas.Brush.Color := GetThemedColor(clHighlight)
    else
      Canvas.Brush.Color := GetThemedColor(clWindow);
    Canvas.FillRect(Rect);
    if odSelected in State then
      Canvas.Font.Color := GetThemedColor(clHighlightText)
    else
      Canvas.Font.Color := GetThemedColor(clWindowText);
    if (Index >= 0) and (Index < FListBox.Items.Count) then
      FallbackText := FListBox.Items[Index]
    else
      FallbackText := '';
    Canvas.TextOut(Rect.Left + 6, Rect.Top + 2, FallbackText);
    Exit;
  end;
  Item := FFilteredItems[Index];

  // Hintergrund
  if odSelected in State then
    Canvas.Brush.Color := GetThemedColor(clHighlight)
  else
    Canvas.Brush.Color := GetThemedColor(clWindow);
  Canvas.FillRect(Rect);

  // Kind-Prefix (farbig)
  Prefix := TCompletionItems.KindToPrefix(Item.Kind);
  if odSelected in State then
    Canvas.Font.Color := GetThemedColor(clHighlightText)
  else
    Canvas.Font.Color := GetThemedColor(clGrayText);
  Canvas.Font.Style := [];
  Canvas.TextOut(Rect.Left + 4, Rect.Top + 2, Prefix);

  // Label
  if odSelected in State then
    Canvas.Font.Color := GetThemedColor(clHighlightText)
  else
    Canvas.Font.Color := GetThemedColor(clWindowText);
  Canvas.Font.Style := [fsBold];
  LabelText := Item.Label_;
  Canvas.TextOut(Rect.Left + 44, Rect.Top + 2, LabelText);

  // Detail (rechts, grau)
  if Item.Detail <> '' then
  begin
    DetailText := Item.Detail;
    if odSelected in State then
      Canvas.Font.Color := GetThemedColor(clHighlightText)
    else
      Canvas.Font.Color := GetThemedColor(clGrayText);
    Canvas.Font.Style := [];
    Canvas.TextOut(Rect.Left + 46 + Canvas.TextWidth(LabelText) + 8,
      Rect.Top + 2, DetailText);
  end;
end;

procedure TCompletionPopup.InsertSelected;
begin
  if (FListBox.ItemIndex >= 0) and (FListBox.ItemIndex < FFilteredItems.Count) then
  begin
    if Assigned(FOnInsert) then
      FOnInsert(FFilteredItems[FListBox.ItemIndex].Label_);
  end;
  HidePopup;
end;

end.
