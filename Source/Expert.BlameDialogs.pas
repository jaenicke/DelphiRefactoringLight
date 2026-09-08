(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.BlameDialogs;

{
  The two things the gutter column makes you want next:

  * "What ELSE did this commit change?"  -> ShowCommitOfCaretLine
    runs "git show" / "svn log -v + svn diff -c" for the revision of the
    caret's line and shows the output read-only.
  * "Show me the whole file annotated" -> ShowBlameForCurrentFile
    lists every line with revision, author, date and summary; a
    double-click jumps to that line in the editor.

  Both are USER ACTIONS, so they may take their time - but not silently:
  the VCS call runs with an hour-glass and a bounded timeout, and any
  failure is reported instead of leaving an empty window.
}

interface

/// <summary>Commit details for the line the caret is on.</summary>
procedure ShowCommitOfCaretLine;

/// <summary>Annotated view of the whole active file.</summary>
procedure ShowBlameForCurrentFile;

/// <summary>True when the caret's file has blame data at all - "show
///  commit of this line" needs the REVISION of that line, which only our
///  own data carries.</summary>
function BlameAvailableForCaretFile: Boolean;

/// <summary>True when a blame VIEW can be opened: either we have data,
///  or Tortoise is going to do it and the file is in a working copy.
///  (Handing the job over needs no data of our own.)</summary>
function BlameViewAvailableForCaretFile: Boolean;

/// <summary>Which viewer the two menu entries will actually open -
///  "built-in", "TortoiseGit" or "TortoiseSVN". Shown in the status
///  window, because the answer depends on a setting AND on what is
///  installed.</summary>
function BlameViewerName: string;

/// <summary>Small LIVE adjuster for the gutter column: every move of a
///  slider is applied and repainted at once, so the column can be lined
///  up with whatever else draws in the gutter (the Parnassus Navigator,
///  for one) BY LOOKING AT IT. Cancel puts the previous values back.
///  Opened from the status window's "Live blame" row.</summary>
procedure AdjustBlameColumn;

implementation

uses
  System.SysUtils, System.Classes, System.DateUtils, System.Math,
  System.IOUtils,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.ExtCtrls,
  Vcl.Menus,
  Expert.EditorHelperIntf, Expert.VcsBlame, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort, Expert.PluginSettings,
  Expert.BlameGutter;

function VcsName(AKind: TVcsKind): string;
begin
  case AKind of
    vcsGit: Result := 'git';
    vcsSvn: Result := 'svn';
  else
    Result := 'the VCS';
  end;
end;

function CaretFileAndLine(out AFile: string; out ALine: Integer): Boolean;
var
  Col: Integer;
begin
  AFile := '';
  ALine := 0;
  Result := Editor <> nil;
  if not Result then Exit;
  AFile := Editor.GetActiveFileName;
  Result := (AFile <> '') and Editor.GetCaretLineCol(ALine, Col) and (ALine > 0);
end;

function BlameAvailableForCaretFile: Boolean;
var
  F: string;
  Line: Integer;
  Lines: TBlameLines;
begin
  Result := CaretFileAndLine(F, Line) and BlameForFile(F, Lines)
    and (Length(Lines) > 0);
end;

// What the two menu entries will actually open - worth stating in the
// status window, because the answer depends on a SETTING and on what is
// installed on this machine.
function BlameViewerName: string;
var
  F: string;
  Line: Integer;
  Kind: TVcsKind;
begin
  Result := 'built-in';
  if not TPluginSettings.BlameUseTortoise then Exit;
  if not CaretFileAndLine(F, Line) then Exit;
  Kind := DetectVcs(F);
  if not TortoiseAvailable(Kind) then Exit;
  case Kind of
    vcsGit: Result := 'TortoiseGit';
    vcsSvn: Result := 'TortoiseSVN';
  end;
end;


function BlameViewAvailableForCaretFile: Boolean;
var
  F: string;
  Line: Integer;
begin
  if BlameAvailableForCaretFile then Exit(True);
  Result := TPluginSettings.BlameUseTortoise and CaretFileAndLine(F, Line)
    and (DetectVcs(F) <> vcsNone) and TortoiseAvailable(DetectVcs(F));
end;

// ---------------------------------------------------------------------------
//  Live adjuster for the gutter column
// ---------------------------------------------------------------------------

type
  TBlameAdjustForm = class(TThemedToolForm)
  private
    FOffset, FWidth: TTrackBar;
    FLblOffset, FLblWidth: TLabel;
    FInfo: TComboBox;
    procedure DoChange(Sender: TObject);
  end;

procedure TBlameAdjustForm.DoChange(Sender: TObject);
begin
  // Write through and repaint on EVERY move - that is the whole point of
  // this dialog: the numbers mean nothing, the picture does.
  TPluginSettings.BlameColumnOffset := FOffset.Position;
  TPluginSettings.BlameColumnWidth := FWidth.Position;
  TPluginSettings.BlameInfo := Max(0, FInfo.ItemIndex);
  FLblOffset.Caption := Format('Offset: %d px', [FOffset.Position]);
  FLblWidth.Caption := Format('Width: %d px', [FWidth.Position]);
  ApplyBlameSettings;
end;

procedure AdjustBlameColumn;
var
  Form: TBlameAdjustForm;
  OldOffset, OldWidth, OldInfo: Integer;
  Btn: TButton;

  function Slider(ATop, AMax, AValue: Integer): TTrackBar;
  begin
    Result := TTrackBar.Create(Form);
    Result.Parent := Form;
    Result.SetBounds(12, ATop, 360, 30);
    Result.Min := 0;
    Result.Max := AMax;
    Result.Frequency := 25;
    Result.Position := EnsureRange(AValue, 0, AMax);
    Result.OnChange := Form.DoChange;
  end;

  function Cap(ATop: Integer; const AText: string): TLabel;
  begin
    Result := TLabel.Create(Form);
    Result.Parent := Form;
    Result.SetBounds(12, ATop, 200, 15);
    Result.Caption := AText;
  end;

begin
  OldOffset := TPluginSettings.BlameColumnOffset;
  OldWidth := TPluginSettings.BlameColumnWidth;
  OldInfo := TPluginSettings.BlameInfo;

  Form := TBlameAdjustForm.CreateNew(Application.MainForm);
  try
    Form.Caption := 'Live blame - gutter column';
    Form.BorderStyle := bsToolWindow;
    Form.ClientWidth := 384;
    Form.ClientHeight := 226;
    // Bottom right, so the gutter it adjusts stays visible.
    Form.Position := poDesigned;
    Form.Left := Screen.WorkAreaRect.Right - Form.Width - 40;
    Form.Top := Screen.WorkAreaRect.Bottom - Form.Height - 60;

    Form.FLblOffset := Cap(10, '');
    Form.FOffset := Slider(28, 400, OldOffset);
    Form.FLblWidth := Cap(66, '');
    Form.FWidth := Slider(84, 400, OldWidth);

    Cap(122, 'Show:');
    Form.FInfo := TComboBox.Create(Form);
    Form.FInfo.Parent := Form;
    Form.FInfo.SetBounds(70, 118, 302, 23);
    Form.FInfo.Style := csDropDownList;
    Form.FInfo.Items.Add('Revision only');
    Form.FInfo.Items.Add('Revision and author');
    Form.FInfo.Items.Add('Revision, author and age');
    Form.FInfo.ItemIndex := EnsureRange(OldInfo, 0, 2);
    Form.FInfo.OnChange := Form.DoChange;

    Cap(152, 'Width 0 = do not widen the gutter at all.');

    Btn := TButton.Create(Form);
    Btn.Parent := Form;
    Btn.SetBounds(196, 180, 84, 28);
    Btn.Caption := 'Keep';
    Btn.ModalResult := mrOk;
    Btn.Default := True;

    Btn := TButton.Create(Form);
    Btn.Parent := Form;
    Btn.SetBounds(288, 180, 84, 28);
    Btn.Caption := 'Cancel';
    Btn.ModalResult := mrCancel;
    Btn.Cancel := True;

    Form.DoChange(nil);      // show the current values
    EnableThemes(Form);
    PrepareDialog(Form, Application.MainForm);

    if Form.ShowModal = mrOk then
      TPluginSettings.Save
    else
    begin
      TPluginSettings.BlameColumnOffset := OldOffset;
      TPluginSettings.BlameColumnWidth := OldWidth;
      TPluginSettings.BlameInfo := OldInfo;
      ApplyBlameSettings;
    end;
  finally
    Form.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  A plain read-only text window (themed like the rest)
// ---------------------------------------------------------------------------

procedure ShowTextWindow(const ACaption, AText: string);
var
  Form: TThemedToolForm;
  Memo: TMemo;
  Btn: TButton;
begin
  Form := TThemedToolForm.CreateNew(Application.MainForm);
  try
    Form.Caption := ACaption;
    Form.BorderStyle := bsSizeable;
    Form.Position := poOwnerFormCenter;
    Form.ClientWidth := 900;
    Form.ClientHeight := 560;

    Memo := TMemo.Create(Form);
    Memo.Parent := Form;
    Memo.Align := alClient;
    Memo.ScrollBars := ssBoth;
    Memo.WordWrap := False;
    Memo.ReadOnly := True;
    Memo.ParentFont := False;
    Memo.Font.Name := 'Consolas';
    Memo.Font.Size := 9;
    Memo.Font.Color := GetThemedColor(clWindowText);
    Memo.Color := GetThemedColor(clWindow);
    Memo.Lines.Text := AText;

    Btn := TButton.Create(Form);
    Btn.Parent := Form;
    Btn.Align := alBottom;
    Btn.Height := 28;
    Btn.Caption := 'Close';
    Btn.ModalResult := mrOk;
    Btn.Cancel := True;
    Btn.Default := True;

    EnableThemes(Form);
    PrepareDialog(Form, Application.MainForm);
    Form.ShowModal;
  finally
    Form.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Commit details
// ---------------------------------------------------------------------------

const
  EMPTYSTR   = '';
  FONTMONO   = 'Consolas';
  CAPACTION  = 'Action';
  CAPFILES   = 'Changed files (%d)';
  CAPCOMMIT  = 'Commit %s - %s';
  NODIFFTEXT = '(no diff for this entry - a directory, a property change, or a binary file)';
  HEADERFMT  = 'Revision: %s'#13#10 + 'Author:   %s'#13#10 +
    'Date:     %s'#13#10 + '--------------------------------------------'#13#10 + '%s';
  MSGNOLINE  = 'No active editor line.';
  MSGNODATA  = 'No blame data for this file yet.'#13#10 + 'Switch live blame on and give it a moment.';
  MSGLOCAL   = 'This line is not committed yet.';
  MSGNOREV   = 'No revision known for this line.';
  MSGNOTHING = 'The version control client returned nothing.';
  MSGSLOW    = 'The version control client did not answer within %d seconds.';
  CommitFetchTimeoutMs = 60000;

type
  // Reference-counted result holder for the worker below.
  TCommitFetch = class(TInterfacedObject)
  public
    Info: TCommitInfo;
    Ok: Boolean;
    Done: Boolean;
  end;

type
  // Log view in the shape TortoiseGit / TortoiseSVN use, because that is
  // what everyone here already reads: metadata and message on top, the
  // files the commit touched in the middle, the selected file's diff
  // below.
  TCommitViewForm = class(TThemedToolForm)
  private
    FHeader: TMemo;
    FFiles: TListView;
    FDiff: TMemo;
    FInfo: TCommitInfo;
    procedure DoFileSelect(Sender: TObject);
  end;

procedure TCommitViewForm.DoFileSelect(Sender: TObject);
var
  Idx: Integer;
begin
  FDiff.Lines.Clear;
  if FFiles.Selected = nil then Exit;
  Idx := NativeInt(FFiles.Selected.Data);
  if (Idx < 0) or (Idx > High(FInfo.Files)) then Exit;
  if FInfo.Files[Idx].Diff = EMPTYSTR then
    FDiff.Lines.Text := NODIFFTEXT
  else
    FDiff.Lines.Text := FInfo.Files[Idx].Diff;
end;

procedure ShowCommitView(const AInfo: TCommitInfo);
var
  Form: TCommitViewForm;
  Split1: TSplitter;
  LI: TListItem;
  I: Integer;
  Msg: string;
begin
  Form := TCommitViewForm.CreateNew(Application.MainForm);
  try
    Form.FInfo := AInfo;
    Form.Caption := Format(CAPCOMMIT, [AInfo.Revision, AInfo.Author]);
    Form.BorderStyle := bsSizeable;
    Form.Position := poOwnerFormCenter;
    Form.ClientWidth := 1000;
    Form.ClientHeight := 700;

    // --- header: who, when, and the message
    Form.FHeader := TMemo.Create(Form);
    Form.FHeader.Parent := Form;
    Form.FHeader.Align := alTop;
    Form.FHeader.Height := 130;
    Form.FHeader.ReadOnly := True;
    Form.FHeader.ScrollBars := ssVertical;
    Form.FHeader.Color := GetThemedColor(clWindow);
    Form.FHeader.Font.Color := GetThemedColor(clWindowText);
    Msg := Trim(AInfo.Subject);
    if Trim(AInfo.Body) <> EMPTYSTR then Msg := Msg + sLineBreak + Trim(AInfo.Body);
    Form.FHeader.Lines.Text := Format(HEADERFMT,
      [AInfo.Revision, AInfo.Author, AInfo.DateStr, Msg]);

    // --- the files the commit touched
    Form.FFiles := TListView.Create(Form);
    Form.FFiles.Parent := Form;
    Form.FFiles.Align := alTop;
    Form.FFiles.Height := 200;
    Form.FFiles.ViewStyle := vsReport;
    Form.FFiles.ReadOnly := True;
    Form.FFiles.RowSelect := True;
    Form.FFiles.OnClick := Form.DoFileSelect;
    Form.FFiles.OnKeyUp := nil;
    Form.FFiles.Columns.Add.Caption := CAPACTION;
    Form.FFiles.Columns[0].Width := 60;
    Form.FFiles.Columns.Add.Caption := Format(CAPFILES, [Length(AInfo.Files)]);
    Form.FFiles.Columns[1].Width := 900;
    for I := 0 to High(AInfo.Files) do
    begin
      LI := Form.FFiles.Items.Add;
      LI.Caption := AInfo.Files[I].Action;
      LI.SubItems.Add(AInfo.Files[I].Path);
      LI.Data := Pointer(NativeInt(I));
    end;
    EnableListViewSorting(Form.FFiles);

    Split1 := TSplitter.Create(Form);
    Split1.Parent := Form;
    Split1.Align := alTop;
    Split1.Top := Form.FFiles.Top + Form.FFiles.Height;
    Split1.Height := 4;

    // --- the diff of whatever is selected above
    Form.FDiff := TMemo.Create(Form);
    Form.FDiff.Parent := Form;
    Form.FDiff.Align := alClient;
    Form.FDiff.ReadOnly := True;
    Form.FDiff.ScrollBars := ssBoth;
    Form.FDiff.WordWrap := False;
    Form.FDiff.ParentFont := False;
    Form.FDiff.Font.Name := FONTMONO;
    Form.FDiff.Font.Size := 9;
    Form.FDiff.Color := GetThemedColor(clWindow);
    Form.FDiff.Font.Color := GetThemedColor(clWindowText);

    if Form.FFiles.Items.Count > 0 then
    begin
      Form.FFiles.ItemIndex := 0;
      Form.DoFileSelect(nil);
    end;

    EnableThemes(Form);
    PrepareDialog(Form, Application.MainForm);
    Form.ShowModal;
  finally
    Form.Free;
  end;
end;

procedure ShowCommitOfCaretLine;
var
  F: string;
  Line: Integer;
  Lines: TBlameLines;
  Info: TBlameLine;
  Holder: TCommitFetch;
  HolderRef: IInterface;
  Prog: TCheckProgressWindow;
  Waited: Integer;
begin
  if not CaretFileAndLine(F, Line) then
  begin
    ShowThemedMessage(MSGNOLINE);
    Exit;
  end;
  if not BlameForFile(F, Lines) or (Line > Length(Lines)) then
  begin
    ShowThemedMessage(MSGNODATA);
    Exit;
  end;
  Info := Lines[Line - 1];
  if Info.IsUncommitted then
  begin
    ShowThemedMessage(MSGLOCAL);
    Exit;
  end;
  if Info.Hash = EMPTYSTR then
  begin
    ShowThemedMessage(MSGNOREV);
    Exit;
  end;

  // TortoiseGit / TortoiseSVN first, when the user wants it and the
  // client is installed: its log window is richer than ours, and it is
  // what most people here already have to hand. Anything that goes wrong
  // (not installed, refused to start) falls through to the built-in view
  // rather than leaving the user with nothing.
  if TPluginSettings.BlameUseTortoise and TortoiseShowLog(F, Info) then Exit;

  // The VCS call runs on a WORKER: a big commit means a big diff, and
  // svn in particular can take seconds. The main thread keeps pumping
  // messages behind a progress window, so the IDE stays alive.
  // The holder is reference-counted: if the wait gives up before the
  // worker does, the LAST of the two frees it - never the waiter.
  Holder := TCommitFetch.Create;
  HolderRef := Holder;
  Prog := CreateCheckProgress('Reading commit...', Application.MainForm,
    Format('Asking %s for %s...',
      [VcsName(Info.Kind), Info.ShortHash]));
  try
    var ThreadRef: IInterface := HolderRef;   // keeps it alive
    TThread.CreateAnonymousThread(
      procedure
      var
        Res: TCommitInfo;
        Ok: Boolean;
      begin
        Ok := False;
        try
          Ok := GetCommitInfo(F, Info, Res);
        except
          Ok := False;
        end;
        TThread.Queue(nil,
          procedure
          begin
            Holder.Info := Res;
            Holder.Ok := Ok;
            Holder.Done := True;
          end);
        ThreadRef := nil;
      end).Start;

    Waited := 0;
    while (not Holder.Done) and (Waited < CommitFetchTimeoutMs) do
    begin
      Application.ProcessMessages;   // the queued result arrives here
      Sleep(30);
      Inc(Waited, 30);
      Prog.Step(Waited div 1000, CommitFetchTimeoutMs div 1000,
        Format('Asking %s for %s...',
          [VcsName(Info.Kind), Info.ShortHash]));
    end;
  finally
    Prog.Free;
  end;

  if not Holder.Done then
  begin
    ShowThemedMessage(Format(MSGSLOW, [CommitFetchTimeoutMs div 1000]));
    Exit;
  end;
  if not Holder.Ok then
  begin
    ShowThemedMessage(MSGNOTHING + sLineBreak + BlameStatus);
    Exit;
  end;
  ShowCommitView(Holder.Info);
end;

// ---------------------------------------------------------------------------
//  Annotated file
// ---------------------------------------------------------------------------

type
  TBlameListForm = class(TThemedToolForm)
  private
    FList: TListView;
    FFile: string;
    procedure DoDblClick(Sender: TObject);
  end;

procedure TBlameListForm.DoDblClick(Sender: TObject);
var
  L: Integer;
begin
  if FList.Selected = nil then Exit;
  L := StrToIntDef(FList.Selected.Caption, 0);
  if (L > 0) and (Editor <> nil) then
    Editor.GotoLocation(FFile, L - 1, 0, 0);
end;

procedure ShowBlameForCurrentFile;
var
  F: string;
  Line, I: Integer;
  Lines: TBlameLines;
  Form: TBlameListForm;
  LI: TListItem;
  SrcLines: TArray<string>;
  Content: string;
begin
  if not CaretFileAndLine(F, Line) then
  begin
    ShowThemedMessage('No active editor file.');
    Exit;
  end;
  // Same here: the Tortoise blame window opens on the caret's line.
  if TPluginSettings.BlameUseTortoise
    and TortoiseShowBlame(F, DetectVcs(F), Line) then Exit;

  if not BlameForFile(F, Lines) or (Length(Lines) = 0) then
  begin
    ShowThemedMessage('No blame data for this file yet.'#13#10 + BlameStatus);
    Exit;
  end;

  // The source as the blame belongs to it: the editor buffer when the
  // file is open (the painter pauses while it is modified anyway),
  // otherwise from disk.
  if not Editor.ReadEditorContent(F, Content) then
    try
      Content := TFile.ReadAllText(F);
    except
      Content := EMPTYSTR;
    end;
  SrcLines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);

  Form := TBlameListForm.CreateNew(Application.MainForm);
  try
    Form.FFile := F;
    Form.Caption := Format('Blame - %s (%d lines)',
      [ExtractFileName(F), Length(Lines)]);
    Form.BorderStyle := bsSizeable;
    Form.Position := poOwnerFormCenter;
    Form.ClientWidth := 1000;
    Form.ClientHeight := 600;

    Form.FList := TListView.Create(Form);
    Form.FList.Parent := Form;
    Form.FList.Align := alClient;
    Form.FList.ViewStyle := vsReport;
    Form.FList.ReadOnly := True;
    Form.FList.RowSelect := True;
    Form.FList.OnDblClick := Form.DoDblClick;
    Form.FList.Columns.Add.Caption := 'Line';
    Form.FList.Columns[0].Width := 60;
    Form.FList.Columns.Add.Caption := 'Revision';
    Form.FList.Columns[1].Width := 90;
    Form.FList.Columns.Add.Caption := 'Author';
    Form.FList.Columns[2].Width := 160;
    Form.FList.Columns.Add.Caption := 'Date';
    Form.FList.Columns[3].Width := 130;
    Form.FList.Columns.Add.Caption := 'Summary';
    Form.FList.Columns[4].Width := 260;
    // Without the LINE ITSELF this is a table of numbers - the code is
    // what one is actually looking at (the Tortoise blame views read
    // that way too).
    Form.FList.Columns.Add.Caption := 'Source';
    Form.FList.Columns[5].Width := 600;

    Form.FList.Items.BeginUpdate;
    try
      for I := 0 to High(Lines) do
      begin
        LI := Form.FList.Items.Add;
        LI.Caption := IntToStr(I + 1);
        LI.Data := Pointer(NativeInt(I));   // sortable rows map via Data
        if Lines[I].IsUncommitted then
          LI.SubItems.Add('(local)')
        else
          LI.SubItems.Add(Lines[I].ShortHash);
        LI.SubItems.Add(Lines[I].Author);
        if Lines[I].AuthorTime > 0 then
          LI.SubItems.Add(FormatDateTime('yyyy-mm-dd hh:nn', Lines[I].AuthorTime))
        else
          LI.SubItems.Add('');
        LI.SubItems.Add(Lines[I].Summary);
        if I <= High(SrcLines) then
          LI.SubItems.Add(SrcLines[I].Replace(#9, '  '))
        else
          LI.SubItems.Add(EMPTYSTR);
      end;
    finally
      Form.FList.Items.EndUpdate;
    end;

    // Start on the caret's line, so the view opens where the user was.
    if (Line >= 1) and (Line <= Form.FList.Items.Count) then
    begin
      Form.FList.ItemIndex := Line - 1;
      Form.FList.Items[Line - 1].MakeVisible(False);
    end;

    EnableListViewSorting(Form.FList);
    EnableThemes(Form);
    PrepareDialog(Form, Application.MainForm);
    Form.ShowModal;
  finally
    Form.Free;
  end;
end;

end.
