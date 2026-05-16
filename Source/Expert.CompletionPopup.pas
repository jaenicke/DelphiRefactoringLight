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
    FFilterEdit: TEdit;
    FAllItems: TArray<TCompletionItem>;
    FFilteredItems: TList<TCompletionItem>;
    FOnInsert: TCompletionInsertEvent;
    procedure DoFilterChange(Sender: TObject);
    procedure DoListBoxDblClick(Sender: TObject);
    procedure DoListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure DoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoFormDeactivate(Sender: TObject);
    procedure ApplyFilter;
    procedure InsertSelected;
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

    property OnInsert: TCompletionInsertEvent read FOnInsert write FOnInsert;
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
  OnDeactivate := DoFormDeactivate;

  FFilteredItems := TList<TCompletionItem>.Create;

  // Filter-Eingabe oben
  FFilterEdit := TEdit.Create(Self);
  FFilterEdit.Parent := Self;
  FFilterEdit.Align := alTop;
  FFilterEdit.Font.Name := 'Consolas';
  FFilterEdit.Font.Size := 10;
  FFilterEdit.TextHint := 'Filter...';
  FFilterEdit.OnChange := DoFilterChange;
  FFilterEdit.OnKeyDown := DoKeyDown;

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
  FListBox.OnKeyDown := DoKeyDown;
  FListBox.OnDrawItem := DoListBoxDrawItem;

  Expert.IdeThemes.EnableThemes(Self);
end;

destructor TCompletionPopup.Destroy;
begin
  FFilteredItems.Free;
  inherited;
end;

procedure TCompletionPopup.ShowLoading(X, Y: Integer);
begin
  SetLength(FAllItems, 0);
  FFilterEdit.Text := '';
  FFilterEdit.Enabled := False;
  FListBox.Items.Clear;
  FListBox.Items.Add('Loading suggestions...');
  FDetailLabel.Caption := '';

  Left := X;
  Top := Y;
  if Left + Width > Screen.Width then
    Left := Screen.Width - Width;
  if Top + Height > Screen.Height then
    Top := Y - Height - 20;

  Show;
  Application.ProcessMessages;
end;

procedure TCompletionPopup.ShowItems(const AItems: TArray<TCompletionItem>; const AInitialFilter: string);
begin
  FAllItems := AItems;
  FFilterEdit.Text := AInitialFilter;
  FFilterEdit.Enabled := True;
  ApplyFilter;
  FFilterEdit.SetFocus;
  // Cursor at end, no selection, so further typing appends to the filter
  FFilterEdit.SelStart := Length(AInitialFilter);
  FFilterEdit.SelLength := 0;
end;

procedure TCompletionPopup.ShowMessage(const AMsg: string);
begin
  SetLength(FAllItems, 0);
  FFilteredItems.Clear;
  FListBox.Items.Clear;
  FListBox.Items.Add(AMsg);
  FDetailLabel.Caption := '';
  FFilterEdit.Enabled := False;
end;

procedure TCompletionPopup.ApplyFilter;
var
  Filter: string;
  Item: TCompletionItem;
begin
  Filter := UpperCase(FFilterEdit.Text);
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

procedure TCompletionPopup.DoFormDeactivate(Sender: TObject);
begin
  // Close when focus moves to another window (typically the user
  // clicked in the IDE editor).
  if Visible then
    Close;
end;

procedure TCompletionPopup.DoFilterChange(Sender: TObject);
begin
  ApplyFilter;
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
  Prefix, LabelText, DetailText: string;
begin
  Canvas := FListBox.Canvas;

  if Index >= FFilteredItems.Count then Exit;
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

procedure TCompletionPopup.DoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_RETURN:
      InsertSelected;
    VK_ESCAPE:
      Close;
    VK_DOWN:
      if (Sender = FFilterEdit) and (FListBox.Items.Count > 0) then
      begin
        FListBox.SetFocus;
        if FListBox.ItemIndex < FListBox.Items.Count - 1 then
          FListBox.ItemIndex := FListBox.ItemIndex + 1;
        Key := 0;
      end;
    VK_UP:
      if (Sender = FFilterEdit) then
      begin
        FListBox.SetFocus;
        Key := 0;
      end;
  end;

  // Detail aktualisieren bei Listbox-Navigation
  if (Sender = FListBox) and (FListBox.ItemIndex >= 0) and
     (FListBox.ItemIndex < FFilteredItems.Count) then
    FDetailLabel.Caption := '  ' + FFilteredItems[FListBox.ItemIndex].Detail;
end;

procedure TCompletionPopup.InsertSelected;
begin
  if (FListBox.ItemIndex >= 0) and (FListBox.ItemIndex < FFilteredItems.Count) then
  begin
    if Assigned(FOnInsert) then
      FOnInsert(FFilteredItems[FListBox.ItemIndex].Label_);
  end;
  Close;
end;

end.
