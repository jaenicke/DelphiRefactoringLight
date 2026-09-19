(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DebugConsistencyDialog;

// IDE side of "Project checks > Debug consistency...": gathers what the
// check needs from the ACTIVE project through the ToolsAPI (active
// configuration AND platform - DCU output, target, host application, debug
// options, unit search path) and shows the findings. The checks themselves
// live in the pure Expert.DebugConsistency.

interface

uses
  Expert.DebugConsistency;

procedure CheckDebugConsistency;

/// <summary>The check's input for the ACTIVE project and configuration
///  (options, search order, target, open files). MAIN THREAD - ToolsAPI.
///  Also used by the MCP bridge.</summary>
function GatherDebugCheckInput(out AInput: TDebugCheckInput;
  out AError: string): Boolean;

/// <summary>Human-readable names of the issue kinds / severities.</summary>
function DebugIssueKindName(AKind: TDebugIssueKind): string;

implementation

uses
  Winapi.Windows, Winapi.ShellAPI,
  System.SysUtils, System.Classes, System.IOUtils, System.Math,
  System.Generics.Collections, System.Generics.Defaults,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Clipbrd,
  ToolsAPI, DCCStrs, CommonOptionStrs,
  Expert.EditorHelperIntf, Expert.DialogHelper, Expert.IdeThemes,
  Expert.ListViewSort, Expert.UnitIndex;

const
  KindNames: array[TDebugIssueKind] of string = ('duplicate source',
    'stray DCU', 'DCU only', 'line endings', 'project option',
    'source directive', 'target', 'symbol file', 'host application',
    'duplicate binary');
  SeverityNames: array[TDebugIssueSeverity] of string = ('3 note', '2 warning',
    '1 problem');

// ---------------------------------------------------------------------------
//  gathering the input
// ---------------------------------------------------------------------------

// A project option value as a directory/file path: the project's own
// platform and configuration first (Expert.UnitIndex.ExpandIdeVars always
// assumes Win32/Release), then the IDE variables, relative to the project.
// '' when a macro stays unresolved.
function ExpandProjectPath(const AValue, AProjectDir, APlatform, AConfig: string): string;
begin
  Result := Trim(AValue);
  if Result = '' then Exit;
  Result := StringReplace(Result, '$(Platform)', APlatform, [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '$(Config)', AConfig, [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '$(ProjectDir)',
    IncludeTrailingPathDelimiter(AProjectDir), [rfReplaceAll, rfIgnoreCase]);
  Result := ExpandIdeVars(Result, FindBdsRoot, nil);
  if Pos('$(', Result) > 0 then Exit('');
  try
    if not TPath.IsPathRooted(Result) then
      Result := TPath.Combine(AProjectDir, Result);
    Result := ExcludeTrailingPathDelimiter(TPath.GetFullPath(Result));
  except
    Result := '';
  end;
end;

procedure AddDir(var ADirs: TArray<string>; const ADir: string);
begin
  if ADir = '' then Exit;
  for var D in ADirs do
    if SameText(D, ADir) then Exit;
  ADirs := ADirs + [ADir];
end;

function DebugIssueKindName(AKind: TDebugIssueKind): string;
begin
  Result := KindNames[AKind];
end;

function GatherDebugCheckInput(out AInput: TDebugCheckInput;
  out AError: string): Boolean;
var
  Project: IOTAProject;
  Cfgs: IOTAProjectOptionsConfigurations;
  Cfg: IOTABuildConfiguration;
  Plat, CfgName, Dir: string;
  Search: TArray<string>;
begin
  Result := False;
  AError := '';
  AInput := Default(TDebugCheckInput);
  try
    Project := GetActiveProject;
  except
    Project := nil;
  end;
  if Project = nil then
  begin
    AError := 'No active project.';
    Exit;
  end;
  try
    AInput.ProjectFile := Project.FileName;
    AInput.ProjectDir := ExcludeTrailingPathDelimiter(ExtractFilePath(Project.FileName));
    Plat := Project.CurrentPlatform;
    CfgName := Project.CurrentConfiguration;
    AInput.ConfigName := CfgName;
    if Plat <> '' then AInput.ConfigName := AInput.ConfigName + ' / ' + Plat;
    AInput.ProjectSources := Editor.GetProjectSourceFiles;
    AInput.OpenFiles := Editor.GetOpenSourceFiles;
    try
      AInput.TargetFile := Project.ProjectOptions.TargetName;
    except
      AInput.TargetFile := '';
    end;

    Search := [AInput.ProjectDir];
    Cfg := nil;
    if Supports(Project.ProjectOptions, IOTAProjectOptionsConfigurations, Cfgs) then
    begin
      Cfg := Cfgs.ActiveConfiguration;
      // the platform-specific configuration holds the effective values
      // (it inherits from the configuration, which inherits from Base)
      if (Cfg <> nil) and (Plat <> '') then
        try
          var PlatCfg := Cfg.PlatformConfiguration[Plat];
          if PlatCfg <> nil then Cfg := PlatCfg;
        except
        end;
    end;
    if Cfg <> nil then
    begin
      AInput.DebugInformation := Cfg.Value[DCCStrs.sDebugInformation];
      AInput.LocalDebugSymbols := Cfg.Value[DCCStrs.sLocalDebugSymbols];
      AInput.Optimize := Cfg.Value[DCCStrs.sOptimize];
      AInput.DcuOutputDir := ExpandProjectPath(Cfg.Value[DCCStrs.sDcuOutput],
        AInput.ProjectDir, Plat, CfgName);
      var Host := Cfg.Value[CommonOptionStrs.sDebugger_HostApplication];
      if Trim(Host) <> '' then
        AInput.HostApplication := ExpandProjectPath(Host, AInput.ProjectDir, Plat, CfgName);
      // compiler order: project directory, the project's search path, then
      // the IDE library path
      for var E in Cfg.Value[DCCStrs.sUnitSearchPath].Split([';']) do
        if (Trim(E) <> '') and (Pos('$(DCC_', E) = 0) then
          AddDir(Search, ExpandProjectPath(E, AInput.ProjectDir, Plat, CfgName));
      AddDir(AInput.BinarySearchDirs,
        ExpandProjectPath(Cfg.Value[DCCStrs.sBplOutput], AInput.ProjectDir, Plat, CfgName));
    end;
    for Dir in GatherCompileSearchDirs do
      AddDir(Search, ExcludeTrailingPathDelimiter(Dir));
    AInput.UnitSearchDirs := Search;

    for var E in GetEnvironmentVariable('PATH').Split([';']) do
      if Trim(E) <> '' then
        AddDir(AInput.BinarySearchDirs, ExcludeTrailingPathDelimiter(Trim(E)));

    Dir := FindBdsRoot;
    if Dir <> '' then AInput.IgnoreDirs := [ExcludeTrailingPathDelimiter(Dir)];
    Result := True;
  except
    on E: Exception do
      AError := 'Reading the project options failed: ' + E.Message;
  end;
end;

// ---------------------------------------------------------------------------
//  dialog
// ---------------------------------------------------------------------------

type
  TDebugConsistencyDialog = class(TForm)
  private
    FInput: TDebugCheckInput;
    FIssues: TArray<TDebugIssue>;
    FDone: TArray<Boolean>;
    FSummary: TLabel;
    FList: TListView;
    FDetail: TMemo;
    FBtnGoto, FBtnExplorer, FBtnConvert, FBtnCopy, FBtnClose: TButton;
    procedure FillRows;
    function SelectedIndex: Integer;
    procedure UpdateButtons;
    procedure DoSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure DoDblClick(Sender: TObject);
    procedure DoGoto(Sender: TObject);
    procedure DoExplorer(Sender: TObject);
    procedure DoConvert(Sender: TObject);
    procedure DoCopy(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
  public
    constructor CreateDialog(AOwner: TComponent; const AInput: TDebugCheckInput;
      const AIssues: TArray<TDebugIssue>);
  end;

function IsSourceFile(const AFile: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(AFile));
  Result := (Ext = '.pas') or (Ext = '.dpr') or (Ext = '.dpk') or (Ext = '.inc');
end;

constructor TDebugConsistencyDialog.CreateDialog(AOwner: TComponent;
  const AInput: TDebugCheckInput; const AIssues: TArray<TDebugIssue>);
var
  Col: TListColumn;
  Panel: TPanel;
  Problems, Warnings, Notes: Integer;

  function AddButton(const ACaption: string; AWidth: Integer; AClick: TNotifyEvent;
    AAlign: TAlign): TButton;
  begin
    Result := TButton.Create(Self);
    Result.Parent := Panel;
    Result.Caption := ACaption;
    Result.Width := AWidth;
    Result.Align := AAlign;
    Result.AlignWithMargins := True;
    Result.OnClick := AClick;
  end;

begin
  inherited CreateNew(AOwner);
  Caption := 'Debug consistency';
  Width := 1100;
  Height := 600;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  OnClose := DoFormClose;
  FInput := AInput;
  FIssues := AIssues;
  SetLength(FDone, Length(FIssues));

  Problems := 0;
  Warnings := 0;
  Notes := 0;
  for var I in FIssues do
    case I.Severity of
      dsProblem: Inc(Problems);
      dsWarning: Inc(Warnings);
    else
      Inc(Notes);
    end;

  FSummary := TLabel.Create(Self);
  FSummary.Parent := Self;
  FSummary.Top := 0;
  FSummary.Align := alTop;
  FSummary.AlignWithMargins := True;
  FSummary.Margins.SetBounds(8, 8, 8, 4);
  FSummary.Caption := Format('%d problem(s), %d warning(s), %d note(s)  -  %s, ' +
    'configuration %s, target %s', [Problems, Warnings, Notes,
    ExtractFileName(FInput.ProjectFile), FInput.ConfigName, FInput.TargetFile]);

  // Explicit Top values BEFORE Align: alBottom controls stack by position.
  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Top := 10000;
  Panel.Align := alBottom;
  Panel.Height := 40;
  Panel.BevelOuter := bvNone;

  FDetail := TMemo.Create(Self);
  FDetail.Parent := Self;
  FDetail.Top := 9000;
  FDetail.Align := alBottom;
  FDetail.Height := 90;
  FDetail.AlignWithMargins := True;
  FDetail.Margins.SetBounds(8, 0, 8, 4);
  FDetail.ReadOnly := True;
  FDetail.ScrollBars := ssVertical;
  FDetail.WordWrap := True;

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.Margins.SetBounds(8, 4, 8, 4);
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.MultiSelect := True;
  FList.HideSelection := False;
  FList.OnDblClick := DoDblClick;
  FList.OnSelectItem := DoSelectItem;
  Col := FList.Columns.Add; Col.Caption := 'Severity'; Col.Width := 80;
  Col := FList.Columns.Add; Col.Caption := 'Kind'; Col.Width := 120;
  Col := FList.Columns.Add; Col.Caption := 'File'; Col.Width := 170;
  Col := FList.Columns.Add; Col.Caption := 'Line'; Col.Width := 50;
  Col := FList.Columns.Add; Col.Caption := 'Reason'; Col.Width := 460;
  Col := FList.Columns.Add; Col.Caption := 'Folder'; Col.Width := 300;

  FBtnClose := AddButton('&Close', 90, DoCloseClick, alRight);
  FBtnClose.Cancel := True;
  FBtnCopy := AddButton('Copy &report', 110, DoCopy, alRight);
  FBtnConvert := AddButton('Convert to C&RLF', 130, DoConvert, alLeft);
  FBtnExplorer := AddButton('Show in &Explorer', 130, DoExplorer, alLeft);
  FBtnGoto := AddButton('&Go to', 90, DoGoto, alLeft);
  FBtnGoto.Default := True;

  FillRows;
  EnableListViewSorting(FList);
  UpdateButtons;

  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

procedure TDebugConsistencyDialog.FillRows;
var
  Item: TListItem;
begin
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for var I := 0 to High(FIssues) do
    begin
      Item := FList.Items.Add;
      Item.Data := Pointer(NativeInt(I));   // sortable: map via Data, never Item.Index
      Item.Caption := SeverityNames[FIssues[I].Severity];
      Item.SubItems.Add(KindNames[FIssues[I].Kind]);
      Item.SubItems.Add(ExtractFileName(FIssues[I].FileName));
      if FIssues[I].Line > 0 then
        Item.SubItems.Add(IntToStr(FIssues[I].Line))
      else
        Item.SubItems.Add('');
      if FDone[I] then
        Item.SubItems.Add('converted to CRLF - ' + FIssues[I].Reason)
      else
        Item.SubItems.Add(FIssues[I].Reason);
      Item.SubItems.Add(ExtractFileDir(FIssues[I].FileName));
    end;
    if FList.Items.Count > 0 then
      FList.Items[0].Selected := True;
  finally
    FList.Items.EndUpdate;
  end;
end;

function TDebugConsistencyDialog.SelectedIndex: Integer;
begin
  if FList.Selected = nil then Exit(-1);
  Result := NativeInt(FList.Selected.Data);
end;

procedure TDebugConsistencyDialog.UpdateButtons;
var
  Idx: Integer;
  AnyFixable: Boolean;
begin
  Idx := SelectedIndex;
  FBtnGoto.Enabled := (Idx >= 0) and IsSourceFile(FIssues[Idx].FileName)
    and TFile.Exists(FIssues[Idx].FileName);
  FBtnExplorer.Enabled := (Idx >= 0) and (FIssues[Idx].FileName <> '');
  AnyFixable := False;
  for var Item in FList.Items do
    if Item.Selected then
    begin
      var I := NativeInt(Item.Data);
      if FIssues[I].Fixable and not FDone[I] then AnyFixable := True;
    end;
  FBtnConvert.Enabled := AnyFixable;
  if Idx >= 0 then
    FDetail.Text := FIssues[Idx].FileName + sLineBreak +
      FIssues[Idx].Reason + sLineBreak + 'What to do: ' + FIssues[Idx].Hint
  else
    FDetail.Text := '';
end;

procedure TDebugConsistencyDialog.DoSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  UpdateButtons;
end;

procedure TDebugConsistencyDialog.DoDblClick(Sender: TObject);
begin
  if FBtnGoto.Enabled then DoGoto(Sender) else if FBtnExplorer.Enabled then DoExplorer(Sender);
end;

procedure TDebugConsistencyDialog.DoGoto(Sender: TObject);
var
  Idx: Integer;
begin
  Idx := SelectedIndex;
  if (Idx < 0) or not IsSourceFile(FIssues[Idx].FileName) then Exit;
  Editor.GotoLocation(FIssues[Idx].FileName, Max(FIssues[Idx].Line - 1, 0), 0);
end;

procedure TDebugConsistencyDialog.DoExplorer(Sender: TObject);
var
  Idx: Integer;
  F: string;
begin
  Idx := SelectedIndex;
  if Idx < 0 then Exit;
  F := FIssues[Idx].FileName;
  if TFile.Exists(F) then
    ShellExecute(0, 'open', 'explorer.exe', PChar('/select,"' + F + '"'), nil, SW_SHOWNORMAL)
  else if TDirectory.Exists(ExtractFileDir(F)) then
    ShellExecute(0, 'open', PChar(ExtractFileDir(F)), nil, nil, SW_SHOWNORMAL)
  else
    ShowThemedMessage('Neither the file nor its folder exists:' + sLineBreak + F);
end;

procedure TDebugConsistencyDialog.DoConvert(Sender: TObject);
var
  Converted, Skipped: Integer;
  Content, Names: string;
begin
  Converted := 0;
  Skipped := 0;
  Names := '';
  for var Item in FList.Items do
  begin
    if not Item.Selected then Continue;
    var I := NativeInt(Item.Data);
    if not FIssues[I].Fixable or FDone[I] then Continue;
    // opened in the meantime? The IDE buffer owns the file then - writing
    // it on disk would be overwritten (or lose the user's edits).
    if (Editor <> nil) and Editor.ReadEditorContent(FIssues[I].FileName, Content) then
    begin
      Inc(Skipped);
      Names := Names + sLineBreak + '  ' + ExtractFileName(FIssues[I].FileName) + ' (open in the editor)';
      Continue;
    end;
    if ConvertFileToCrLf(FIssues[I].FileName) then
    begin
      FDone[I] := True;
      Inc(Converted);
    end
    else
    begin
      Inc(Skipped);
      Names := Names + sLineBreak + '  ' + ExtractFileName(FIssues[I].FileName) + ' (could not be written)';
    end;
  end;
  FillRows;
  UpdateButtons;
  if Skipped > 0 then
    ShowThemedMessage(Format('%d file(s) converted to CRLF, %d skipped:%s',
      [Converted, Skipped, Names]))
  else
    ShowThemedMessage(Format('%d file(s) converted to CRLF. Rebuild the project so ' +
      'the debug information matches the new line numbering.', [Converted]));
end;

procedure TDebugConsistencyDialog.DoCopy(Sender: TObject);
var
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    SB.AppendLine('Debug consistency - ' + FInput.ProjectFile + ' (' + FInput.ConfigName + ')');
    SB.AppendLine('Target: ' + FInput.TargetFile);
    SB.AppendLine;
    for var I := 0 to High(FIssues) do
    begin
      SB.Append(Copy(SeverityNames[FIssues[I].Severity], 3, MaxInt)).Append(' | ')
        .Append(KindNames[FIssues[I].Kind]).Append(' | ').Append(FIssues[I].FileName);
      if FIssues[I].Line > 0 then SB.Append('(').Append(FIssues[I].Line).Append(')');
      SB.AppendLine;
      SB.Append('  ').AppendLine(FIssues[I].Reason);
      SB.Append('  -> ').AppendLine(FIssues[I].Hint);
    end;
    Clipboard.AsText := SB.ToString;
  finally
    SB.Free;
  end;
end;

procedure TDebugConsistencyDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TDebugConsistencyDialog.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

// ---------------------------------------------------------------------------
//  entry point
// ---------------------------------------------------------------------------

procedure CheckDebugConsistency;
var
  Input: TDebugCheckInput;
  Err: string;
  Issues: TArray<TDebugIssue>;
  Progress: TCheckProgressWindow;
begin
  if not GatherDebugCheckInput(Input, Err) then
  begin
    ShowThemedMessage(Err);
    Exit;
  end;
  Progress := CreateCheckProgress('Debug consistency', nil);
  try
    Screen.Cursor := crHourGlass;
    try
      Issues := RunDebugConsistencyCheck(Input,
        function(ACurrent, ATotal: Integer; const AText: string): Boolean
        begin
          if (ACurrent mod 5 = 0) or (ACurrent = ATotal) then
            Progress.Step(ACurrent, ATotal, AText);
          Result := True;
        end);
    finally
      Screen.Cursor := crDefault;
    end;
  finally
    Progress.Free;
  end;
  if Length(Issues) = 0 then
  begin
    ShowThemedMessage(Format('No debug consistency problems found.' + sLineBreak +
      sLineBreak + '%d project file(s), %d open file(s) and %d search path director(y/ies) ' +
      'checked (configuration %s).', [Length(Input.ProjectSources), Length(Input.OpenFiles),
      Length(Input.UnitSearchDirs), Input.ConfigName]));
    Exit;
  end;
  // problems first, then by kind and file
  TArray.Sort<TDebugIssue>(Issues, TComparer<TDebugIssue>.Construct(
    function(const A, B: TDebugIssue): Integer
    begin
      Result := Ord(B.Severity) - Ord(A.Severity);
      if Result = 0 then Result := Ord(A.Kind) - Ord(B.Kind);
      if Result = 0 then Result := CompareText(A.FileName, B.FileName);
    end));
  TDebugConsistencyDialog.CreateDialog(Application.MainForm, Input, Issues).Show;
end;

end.
