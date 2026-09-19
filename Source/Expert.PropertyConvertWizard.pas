(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.PropertyConvertWizard;

// The editor side of the property converter (Expert.PropertyConvert): menu
// entry "Convert properties...", the dialog, and the MCP tool
// "convert_properties". Works on the properties declared on the selected
// lines (or the caret's line). The result goes through ApplyLinesMinimal -
// in the editor buffer, undoable, not saved.

interface

/// <summary>Menu entry: the properties on the selected lines (or the caret's
///  line) of the active unit.</summary>
procedure ConvertPropertiesAtSelection;

implementation

uses
  System.SysUtils, System.Classes, System.JSON, System.Generics.Collections,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.PropertyConvert, Expert.ScopeFiles,
  Expert.UnitIndex, Expert.UsesEditor, Expert.DialogHelper, Expert.IdeThemes,
  Expert.McpServer, Expert.IncludeExpansion;

function SplitLines(const S: string): TArray<string>;
begin
  Result := S.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

type
  // the file cache of one check - freed together with the closure that
  // holds the interface reference
  TUseCache = class(TInterfacedObject)
  public
    Map: TDictionary<string, TArray<string>>;
    constructor Create;
    destructor Destroy; override;
  end;

constructor TUseCache.Create;
begin
  inherited;
  Map := TDictionary<string, TArray<string>>.Create;
end;

destructor TUseCache.Destroy;
begin
  Map.Free;
  inherited;
end;

// "used outside this unit?" - a whole-word occurrence in CODE of any other
// file of the scan scope counts (conservative: no LSP needed, and a false
// "used" only keeps an accessor). MAIN THREAD (reads editor buffers).
function MakeExternalUseCheck(const AFile: string): TExternalUseCheck;
var
  Files: TArray<string>;
  Cache: TUseCache;
  Holder: IInterface;
begin
  Files := ProjectScopeFiles(AFile);
  Cache := TUseCache.Create;
  Holder := Cache;
  Result :=
    function(const AName: string): Boolean
    var
      Lines: TArray<string>;
      C: string;
    begin
      var Keep := Holder;   // the closure owns the cache
      Result := False;
      for var F in Files do
      begin
        if SameText(ExpandFileName(F), ExpandFileName(AFile)) then Continue;
        if not Cache.Map.TryGetValue(UpperCase(F), Lines) then
        begin
          if EditorOrDiskReader()(F, C) then Lines := SplitLines(C) else Lines := nil;
          Cache.Map.Add(UpperCase(F), Lines);
        end;
        if Length(CodeWordLines(Lines, AName)) > 0 then Exit(True);
      end;
    end;
end;

// ---------------------------------------------------------------------------
//  Dialog
// ---------------------------------------------------------------------------

type
  TPropertyConvertDialog = class(TForm)
  private
    FFile: string;
    FContent: string;
    FLines: TArray<string>;
    FFirst, FLast: Integer;
    FExternal: TExternalUseCheck;
    FPlan: TPropConvPlan;
    FDirection: TRadioGroup;
    FChkGetter, FChkSetter: TCheckBox;
    FList: TListView;
    FStatus: TLabel;
    FBtnApply: TButton;
    procedure Replan(Sender: TObject);
    procedure DoApply(Sender: TObject);
  public
    constructor CreateDialog(AOwner: TComponent; const AFile, AContent: string;
      AFirst, ALast: Integer);
  end;

constructor TPropertyConvertDialog.CreateDialog(AOwner: TComponent; const AFile,
  AContent: string; AFirst, ALast: Integer);
var
  Col: TListColumn;
  Top, Bottom: TPanel;
  Btn: TButton;
begin
  inherited CreateNew(AOwner);
  FFile := AFile;
  FContent := AContent;
  FLines := SplitLines(AContent);
  FFirst := AFirst;
  FLast := ALast;
  FExternal := MakeExternalUseCheck(AFile);
  Caption := 'Convert properties - ' + ExtractFileName(AFile);
  Width := 900;
  Height := 440;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  Top := TPanel.Create(Self);
  Top.Parent := Self;
  Top.Align := alTop;
  Top.Height := 76;
  Top.BevelOuter := bvNone;

  FDirection := TRadioGroup.Create(Self);
  FDirection.Parent := Top;
  FDirection.SetBounds(8, 4, 360, 66);
  FDirection.Caption := 'Direction';
  FDirection.Items.Add('Field access -> getter / setter');
  FDirection.Items.Add('Getter / setter -> field access (trivial ones only)');
  FDirection.ItemIndex := 0;
  FDirection.OnClick := Replan;

  FChkGetter := TCheckBox.Create(Self);
  FChkGetter.Parent := Top;
  FChkGetter.SetBounds(384, 14, 200, 20);
  FChkGetter.Caption := 'Getter (read)';
  FChkGetter.Checked := True;
  FChkGetter.OnClick := Replan;

  FChkSetter := TCheckBox.Create(Self);
  FChkSetter.Parent := Top;
  FChkSetter.SetBounds(384, 40, 200, 20);
  FChkSetter.Caption := 'Setter (write)';
  FChkSetter.Checked := True;
  FChkSetter.OnClick := Replan;

  Bottom := TPanel.Create(Self);
  Bottom.Parent := Self;
  Bottom.Align := alBottom;
  Bottom.Height := 40;
  Bottom.BevelOuter := bvNone;

  Btn := TButton.Create(Self);
  Btn.Parent := Bottom;
  Btn.Caption := '&Close';
  Btn.Align := alRight;
  Btn.AlignWithMargins := True;
  Btn.Cancel := True;
  Btn.ModalResult := mrCancel;

  FBtnApply := TButton.Create(Self);
  FBtnApply.Parent := Bottom;
  FBtnApply.Caption := '&Apply';
  FBtnApply.Align := alRight;
  FBtnApply.AlignWithMargins := True;
  FBtnApply.Default := True;
  FBtnApply.OnClick := DoApply;

  FStatus := TLabel.Create(Self);
  FStatus.Parent := Bottom;
  FStatus.Align := alClient;
  FStatus.AlignWithMargins := True;
  FStatus.Layout := tlCenter;

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  Col := FList.Columns.Add; Col.Caption := 'Property'; Col.Width := 110;
  Col := FList.Columns.Add; Col.Caption := 'Line';     Col.Width := 45;
  Col.Alignment := taRightJustify;
  Col := FList.Columns.Add; Col.Caption := 'Result';   Col.Width := 360;
  Col := FList.Columns.Add; Col.Caption := 'What happens / why not'; Col.Width := 340;

  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
  Replan(nil);
end;

procedure TPropertyConvertDialog.Replan(Sender: TObject);
var
  Item: TListItem;
begin
  var Dir := pcToAccessors;
  if FDirection.ItemIndex = 1 then Dir := pcToFields;
  FChkGetter.Enabled := Dir = pcToAccessors;
  FChkSetter.Enabled := Dir = pcToAccessors;
  FPlan := PlanPropertyConversion(FLines, FFirst, FLast, Dir, FChkGetter.Checked,
    FChkSetter.Checked, FExternal);
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for var It in FPlan.Items do
    begin
      Item := FList.Items.Add;
      Item.Caption := It.Name;
      Item.SubItems.Add(IntToStr(It.Line + 1));
      if It.Ok then Item.SubItems.Add(Trim(It.After)) else Item.SubItems.Add('(unchanged)');
      Item.SubItems.Add(It.Note);
    end;
  finally
    FList.Items.EndUpdate;
  end;
  if FPlan.Error <> '' then
    FStatus.Caption := FPlan.Error
  else
    FStatus.Caption := Format('%d of %d propert(y/ies) will be converted.',
      [FPlan.OkCount, Length(FPlan.Items)]);
  FBtnApply.Enabled := FPlan.Changed;
end;

procedure TPropertyConvertDialog.DoApply(Sender: TObject);
var
  Cur: string;
  SL: TStringList;
begin
  // the buffer must still be what the plan was made from
  if not Editor.ReadEditorContent(FFile, Cur) or (Cur <> FContent) then
  begin
    ShowThemedMessage('The unit changed since the dialog opened - start the ' +
      'conversion again.');
    Exit;
  end;
  SL := TStringList.Create;
  try
    SL.Text := string.Join(sLineBreak, FPlan.NewLines);
    if ApplyLinesMinimal(FFile, SL, FContent) then
      ModalResult := mrOk
    else
      ShowThemedMessage('The change could not be applied.');
  finally
    SL.Free;
  end;
end;

procedure ConvertPropertiesAtSelection;
var
  F, Text, Content: string;
  SL, SC, EL, EC: Integer;
begin
  if Editor = nil then Exit;
  if Editor.GetSelection(F, SL, SC, EL, EC, Text) and (Trim(Text) <> '') then
  begin
    // a selection ending at column 1 of the next line does not include it
    if (EL > SL) and (EC <= 1) then Dec(EL);
  end
  else
  begin
    var Ctx := Editor.GetCurrentContext;
    F := Ctx.FileName;
    SL := Ctx.Line;
    EL := Ctx.Line;
  end;
  if (F = '') or not Editor.ReadEditorContent(F, Content) then
  begin
    ShowThemedMessage('Convert properties: no source file is active.');
    Exit;
  end;
  var Dlg := TPropertyConvertDialog.CreateDialog(Application.MainForm, F, Content,
    SL - 1, EL - 1);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  MCP tool "convert_properties"
// ---------------------------------------------------------------------------

function ToolConvertProperties(AArgs: TJSONObject; AStop: THandle): string;
var
  Plan: TPropConvPlan;
  Err, RunErr, Content: string;
  Applied: Boolean;
begin
  var F := AArgs.GetValue<string>('file', '');
  var L1 := AArgs.GetValue<Integer>('from_line', 0);
  var L2 := AArgs.GetValue<Integer>('to_line', L1);
  if (F = '') or (L1 < 1) then
    Exit(McpErr('arguments "file" and "from_line" (1-based) are required'));
  F := ExpandFileName(F);
  var DirS := LowerCase(AArgs.GetValue<string>('direction', 'to_accessors'));
  var Dir := pcToAccessors;
  if DirS = 'to_fields' then Dir := pcToFields
  else if DirS <> 'to_accessors' then
    Exit(McpErr('direction must be "to_accessors" or "to_fields"'));
  var Getter := AArgs.GetValue<Boolean>('getter', True);
  var Setter := AArgs.GetValue<Boolean>('setter', True);
  var DoApply := AArgs.GetValue<Boolean>('apply', False);
  RunErr := '';
  Applied := False;
  if not McpRunOnMain(
    procedure
    begin
      try
        if not McpReadContent(F, Content) then
        begin
          RunErr := 'file not found: ' + F;
          Exit;
        end;
        Plan := PlanPropertyConversion(SplitLines(Content), L1 - 1, L2 - 1, Dir, Getter,
          Setter, MakeExternalUseCheck(F));
        if DoApply and Plan.Changed then
        begin
          var SL := TStringList.Create;
          try
            SL.Text := string.Join(sLineBreak, Plan.NewLines);
            Applied := ApplyLinesMinimal(F, SL, Content);
          finally
            SL.Free;
          end;
        end;
      except
        on E: Exception do RunErr := E.ClassName + ': ' + E.Message;
      end;
    end, not DoApply, AStop, Err, 60000) then
    Exit(McpErr(Err));
  if RunErr <> '' then Exit(McpErr(RunErr));
  var J := TJSONObject.Create;
  J.AddPair('container', Plan.Container);
  if Plan.Error <> '' then J.AddPair('error', Plan.Error);
  var A := TJSONArray.Create;
  for var It in Plan.Items do
  begin
    var O := TJSONObject.Create;
    O.AddPair('property', It.Name);
    O.AddPair('line', TJSONNumber.Create(It.Line + 1));
    O.AddPair('converted', TJSONBool.Create(It.Ok));
    if It.Ok then O.AddPair('after', Trim(It.After));
    O.AddPair('note', It.Note);
    A.Add(O);
  end;
  J.AddPair('properties', A);
  if DoApply then J.AddPair('applied', TJSONBool.Create(Applied));
  Result := McpOk(J);
end;

initialization
  RegisterMcpTool('convert_properties', ToolConvertProperties);

end.
