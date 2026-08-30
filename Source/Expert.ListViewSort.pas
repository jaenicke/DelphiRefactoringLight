(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ListViewSort;

{
  Click-to-sort for the plugin's result tables.

  EnableListViewSorting attaches a small helper component to a (non-
  virtual) TListView: clicking a column header sorts by that column,
  clicking it again reverses the order; the header shows the native
  sort arrow. Values that parse as numbers sort numerically, everything
  else case-insensitively as text.

  IMPORTANT for callers: sorting reorders Items, so Item.INDEX no longer
  corresponds to the fill order. Every consumer of a sortable list must
  map rows back to its data through Item.Data (set to the original array
  index at fill time), never through Item.Index.

  OwnerData (virtual) list views cannot be sorted this way - sort the
  underlying array instead and Invalidate.
}

interface

uses
  Vcl.ComCtrls;

procedure EnableListViewSorting(ALV: TListView);

/// <summary>Shows the native header sort arrow on column AColumn
///  (-1 = none) of any TListView - also usable by virtual list views
///  that sort their data array themselves.</summary>
procedure SetListViewSortArrow(ALV: TListView; AColumn: Integer;
  AAscending: Boolean);

implementation

uses
  System.SysUtils, System.Classes, System.Math,
  Winapi.Windows, Winapi.CommCtrl, Vcl.Controls;

type
  TListViewSorter = class(TComponent)
  private
    FLV: TListView;
    FSortColumn: Integer;    // -1 = unsorted
    FAscending: Boolean;
    procedure DoColumnClick(Sender: TObject; Column: TListColumn);
    function CellText(AItem: TListItem): string;
  end;

procedure SetListViewSortArrow(ALV: TListView; AColumn: Integer;
  AAscending: Boolean);
var
  Header: HWND;
  Item: THDItem;
  I: Integer;
begin
  if not ALV.HandleAllocated then Exit;
  Header := ListView_GetHeader(ALV.Handle);
  if Header = 0 then Exit;
  for I := 0 to ALV.Columns.Count - 1 do
  begin
    FillChar(Item, SizeOf(Item), 0);
    Item.Mask := HDI_FORMAT;
    if not Header_GetItem(Header, I, Item) then Continue;
    Item.fmt := Item.fmt and not (HDF_SORTUP or HDF_SORTDOWN);
    if I = AColumn then
    begin
      if AAscending then
        Item.fmt := Item.fmt or HDF_SORTUP
      else
        Item.fmt := Item.fmt or HDF_SORTDOWN;
    end;
    Header_SetItem(Header, I, Item);
  end;
end;

function TListViewSorter.CellText(AItem: TListItem): string;
begin
  if FSortColumn <= 0 then
    Result := AItem.Caption
  else if FSortColumn - 1 < AItem.SubItems.Count then
    Result := AItem.SubItems[FSortColumn - 1]
  else
    Result := '';
end;

// VCL CustomSort hands the TListItem instances through lParam1/lParam2.
function LVCompare(L1, L2, AParam: LPARAM): Integer; stdcall;
var
  Sorter: TListViewSorter;
  S1, S2: string;
  N1, N2: Double;
begin
  Sorter := TListViewSorter(AParam);
  S1 := Sorter.CellText(TListItem(L1));
  S2 := Sorter.CellText(TListItem(L2));
  // Numeric-aware: line numbers / counts sort as numbers, not text.
  if TryStrToFloat(S1, N1) and TryStrToFloat(S2, N2) then
    Result := CompareValue(N1, N2)
  else
    Result := CompareText(S1, S2);
  if Result = 0 then
    Result := CompareText(TListItem(L1).Caption, TListItem(L2).Caption);
  if not Sorter.FAscending then
    Result := -Result;
end;

procedure TListViewSorter.DoColumnClick(Sender: TObject; Column: TListColumn);
begin
  if FSortColumn = Column.Index then
    FAscending := not FAscending
  else
  begin
    FSortColumn := Column.Index;
    FAscending := True;
  end;
  FLV.CustomSort(@LVCompare, LPARAM(Self));
  SetListViewSortArrow(FLV, FSortColumn, FAscending);
end;

procedure EnableListViewSorting(ALV: TListView);
var
  Sorter: TListViewSorter;
begin
  if (ALV = nil) or ALV.OwnerData then Exit;
  Sorter := TListViewSorter.Create(ALV);   // freed with the list view
  Sorter.FLV := ALV;
  Sorter.FSortColumn := -1;
  Sorter.FAscending := True;
  ALV.OnColumnClick := Sorter.DoColumnClick;
end;

end.
