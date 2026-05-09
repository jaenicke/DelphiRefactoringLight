(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.WithRefactorWizard;

{
  Orchestrates the project-wide "remove with" refactoring:

    1) Resolve project files via TEditorHelper.GetProjectSourceFiles.
    2) Save unsaved editor buffers to disk so the LSP and the file
       reads agree on content.
    3) Start LSP via TLspManager and ensure it has indexed the project.
    4) For each file, run TWithScanner.ScanFile.
    5) For each occurrence, run TWithRewriter.Rewrite to produce the
       before/after texts.
    6) Show TWithRefactorDialog modal. The user reviews the entries
       and clicks "Apply selected", "Apply all" or "Close".
    7) Apply the chosen edits via TEditorHelper.ApplyEditViaEditor.
       For multiple edits in the same file, apply bottom-up so earlier
       offsets stay stable.

  This wizard does NOT need an editor cursor position - it operates on
  the whole project. It can therefore be invoked even when no source
  file is currently focused.
}

interface

uses
  System.Classes, System.SysUtils, System.Generics.Collections,
  Vcl.Forms, ToolsAPI;

type
  TLspWithRefactorWizard = class(TNotifierObject, IOTAWizard, IOTAMenuWizard)
  public
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    procedure Execute;
    function GetMenuText: string;
  end;

var
  WithRefactorInstance: TLspWithRefactorWizard;

implementation

uses
  System.UITypes, System.IOUtils, System.Math,
  Vcl.Dialogs, Vcl.Controls,
  Expert.EditorHelper, Expert.LspManager, Lsp.Client,
  Expert.WithScanner, Expert.WithRewriter, Expert.WithRefactorDialog;

procedure TLspWithRefactorWizard.AfterSave; begin end;
procedure TLspWithRefactorWizard.BeforeSave; begin end;
procedure TLspWithRefactorWizard.Destroyed; begin end;
procedure TLspWithRefactorWizard.Modified; begin end;

function TLspWithRefactorWizard.GetIDString: string;
begin
  Result := 'DelphiRefactoringLight.WithRefactorWizard';
end;

function TLspWithRefactorWizard.GetName: string;
begin
  Result := 'Delphi Refactoring Light - Remove with';
end;

function TLspWithRefactorWizard.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

function TLspWithRefactorWizard.GetMenuText: string;
begin
  Result := 'Remove with (project-wide)...';
end;

{ ---------- Apply phase ---------- }

/// <summary>Sorts edits by file (asc) and within a file by replace
///  range start (DESCENDING) so applying them top-to-bottom does not
///  invalidate offsets of later edits.</summary>
procedure SortForApply(var AEdits: TArray<TWithRewriteResult>);

  procedure SwapAt(I, J: Integer);
  var
    Tmp: TWithRewriteResult;
  begin
    Tmp := AEdits[I];
    AEdits[I] := AEdits[J];
    AEdits[J] := Tmp;
  end;

  function Less(const A, B: TWithRewriteResult): Boolean;
  var
    Cmp: Integer;
  begin
    Cmp := CompareText(A.FileName, B.FileName);
    if Cmp <> 0 then Exit(Cmp < 0);
    if A.ReplaceRange.StartPos.Line <> B.ReplaceRange.StartPos.Line then
      Exit(A.ReplaceRange.StartPos.Line > B.ReplaceRange.StartPos.Line);
    Result := A.ReplaceRange.StartPos.Col > B.ReplaceRange.StartPos.Col;
  end;

var
  I, J: Integer;
begin
  // Tiny insertion sort - the list is small (one entry per applied
  // with-statement, typically <100).
  for I := 1 to High(AEdits) do
  begin
    J := I;
    while (J > 0) and Less(AEdits[J], AEdits[J - 1]) do
    begin
      SwapAt(J, J - 1);
      Dec(J);
    end;
  end;
end;

/// <summary>Applies one rewrite using the IDE editor APIs so the change
///  is undoable. Returns True on success.</summary>
function ApplyOneEdit(const AItem: TWithRewriteResult): Boolean;
begin
  // ApplyEditViaEditor wants 0-based line / 0-based column on the start
  // position, plus the literal old text and new text.
  Result := TEditorHelper.ApplyEditViaEditor(
    AItem.FileName,
    AItem.ReplaceRange.StartPos.Line - 1,
    AItem.ReplaceRange.StartPos.Col - 1,
    AItem.OriginalText,
    AItem.NewText);
end;

procedure ApplyEdits(const AItems: TArray<TWithRewriteResult>;
  out AOk, AFailed: Integer);
var
  Sorted: TArray<TWithRewriteResult>;
  I: Integer;
begin
  AOk := 0;
  AFailed := 0;
  Sorted := Copy(AItems);
  SortForApply(Sorted);
  for I := 0 to High(Sorted) do
    if ApplyOneEdit(Sorted[I]) then
      Inc(AOk)
    else
      Inc(AFailed);
end;

{ ---------- Scan phase ---------- }

/// <summary>Returns True iff the file path looks like a Pascal source
///  we should scan (.pas / .dpr / .dpk). Other extensions are skipped.</summary>
function IsScannableSource(const APath: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(APath));
  Result := (Ext = '.pas') or (Ext = '.dpr') or (Ext = '.dpk');
end;

procedure TLspWithRefactorWizard.Execute;
var
  Dialog: TWithRefactorDialog;
  ProjFiles: TArray<string>;
  DelphiLspJson, RootPath, ProjFile: string;
  Client: TLspClient;
  Results: TList<TWithRewriteResult>;
  ScanFiles: TList<string>;
  I, FileIdx: Integer;
  Source: string;
  Occs: TArray<TWithOccurrence>;
  Occ: TWithOccurrence;
  Rewrite: TWithRewriteResult;
  AutoCount, ApplyOk, ApplyFailed, ModalRes, SelIdx: Integer;
  ItemsToApply: TArray<TWithRewriteResult>;
begin
  Dialog := TWithRefactorDialog.CreateDialog(Application.MainForm);
  try
    Dialog.SetStatus('Resolving project source files...');
    Dialog.Show;
    Application.ProcessMessages;

    ProjFiles := TEditorHelper.GetProjectSourceFiles;
    ScanFiles := TList<string>.Create;
    try
      for I := 0 to High(ProjFiles) do
        if IsScannableSource(ProjFiles[I]) and TFile.Exists(ProjFiles[I]) then
          ScanFiles.Add(ProjFiles[I]);

      if ScanFiles.Count = 0 then
      begin
        Dialog.SetStatus('No project source files found.');
        Dialog.Hide;
        Dialog.ShowModal;
        Exit;
      end;

      // Save editors so on-disk content matches what the LSP sees.
      Dialog.SetStatus('Saving open editor buffers...');
      Application.ProcessMessages;
      TEditorHelper.SaveAllFiles;

      // Resolve LSP config.
      DelphiLspJson := TEditorHelper.FindDelphiLspJson;
      if DelphiLspJson = '' then
      begin
        Dialog.SetStatus('No .delphilsp.json found - cannot resolve target types. ' +
          'Enable Tools > Options > Editor > Language > Code Insight > "Generate LSP Config".');
        Dialog.Hide;
        Dialog.ShowModal;
        Exit;
      end;

      RootPath := TEditorHelper.GetProjectRoot;
      ProjFile := TEditorHelper.GetCurrentProjectDproj;
      if RootPath = '' then
        RootPath := ExtractFilePath(ProjFile);

      Dialog.SetStatus('Starting LSP server...');
      Application.ProcessMessages;
      try
        Client := TLspManager.Instance.GetClient(RootPath, ProjFile, DelphiLspJson);
      except
        on E: Exception do
        begin
          Dialog.SetStatus('LSP startup failed: ' + E.Message);
          Dialog.Hide;
          Dialog.ShowModal;
          Exit;
        end;
      end;

      Dialog.SetStatus('Indexing project...');
      Application.ProcessMessages;
      try
        TLspManager.Instance.EnsureProjectIndexed(ScanFiles.ToArray,
          procedure(ACurrent, ATotal: Integer; const ACurrentFile: string)
          begin
            Dialog.SetProgress(ACurrent, ATotal);
            if ACurrentFile <> '' then
              Dialog.SetStatus(Format('Indexing %d/%d: %s',
                [ACurrent, ATotal, ExtractFileName(ACurrentFile)]));
            Application.ProcessMessages;
          end);
      except
        on E: Exception do
          Dialog.SetStatus('Indexing partial: ' + E.Message);
      end;

      // Scan + rewrite.
      Results := TList<TWithRewriteResult>.Create;
      try
        Dialog.SetProgress(0, ScanFiles.Count);
        for FileIdx := 0 to ScanFiles.Count - 1 do
        begin
          Dialog.SetStatus(Format('Scanning %d/%d: %s',
            [FileIdx + 1, ScanFiles.Count, ExtractFileName(ScanFiles[FileIdx])]));
          Dialog.SetProgress(FileIdx + 1, ScanFiles.Count);
          if (FileIdx mod 3 = 0) then Application.ProcessMessages;

          try
            Source := TFile.ReadAllText(ScanFiles[FileIdx]);
          except
            Continue;
          end;

          Occs := TWithScanner.ScanSource(Source);
          if Length(Occs) = 0 then Continue;

          for Occ in Occs do
          begin
            try
              Rewrite := TWithRewriter.Rewrite(Client,
                ScanFiles[FileIdx], Source, Occ);
            except
              on E: Exception do
              begin
                Rewrite := Default(TWithRewriteResult);
                Rewrite.FileName := ScanFiles[FileIdx];
                Rewrite.Occurrence := Occ;
                Include(Rewrite.Issues, wriTypeUnresolved);
              end;
            end;
            Results.Add(Rewrite);
          end;
        end;

        // Compose status summary.
        AutoCount := 0;
        for I := 0 to Results.Count - 1 do
          if Results[I].IsAutoRewritable then Inc(AutoCount);

        Dialog.SetStatus(Format('Found %d with-statement(s) - %d auto-rewritable.',
          [Results.Count, AutoCount]));
        Dialog.SetProgress(0, 0);
        Dialog.SetItems(Results.ToArray);

        // Switch to modal for review.
        Dialog.Hide;
        ModalRes := Dialog.ShowModal;

        // Apply.
        ItemsToApply := nil;
        case ModalRes of
          mrOk:
            begin
              SetLength(ItemsToApply, AutoCount);
              var K := 0;
              for I := 0 to High(Dialog.Items) do
                if Dialog.Items[I].IsAutoRewritable then
                begin
                  ItemsToApply[K] := Dialog.Items[I];
                  Inc(K);
                end;
            end;
          mrApplySelected:
            begin
              SelIdx := Dialog.SelectedIndex;
              if (SelIdx >= 0) and (SelIdx <= High(Dialog.Items))
                and Dialog.Items[SelIdx].IsAutoRewritable then
              begin
                SetLength(ItemsToApply, 1);
                ItemsToApply[0] := Dialog.Items[SelIdx];
              end;
            end;
        end;

        if Length(ItemsToApply) > 0 then
        begin
          ApplyEdits(ItemsToApply, ApplyOk, ApplyFailed);
          if ApplyFailed > 0 then
            MessageDlg(Format('Applied %d edit(s); %d failed.',
              [ApplyOk, ApplyFailed]), mtWarning, [mbOK], 0);
        end;
      finally
        Results.Free;
      end;
    finally
      ScanFiles.Free;
    end;
  finally
    Dialog.Free;
  end;
end;

end.
