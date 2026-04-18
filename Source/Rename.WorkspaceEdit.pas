(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Rename.WorkspaceEdit;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Generics.Collections, System.Generics.Defaults, Delphi.FileEncoding,
  Lsp.Protocol;

type
  /// <summary>
  ///  Applies an LSP WorkspaceEdit to the affected files.
  ///  Supports dry run (preview) and backup creation.
  /// </summary>
  TWorkspaceEditApplier = class
  private
    FEdit: TLspWorkspaceEdit;
    FBackupDir: string;
    FNoBackup: Boolean;
    procedure SortEditsReverse(var AEdits: TArray<TLspTextEdit>);
    function ApplyEditsToContent(const AContent: string; const AEdits: TArray<TLspTextEdit>): string;
  public
    constructor Create(const AEdit: TLspWorkspaceEdit);

    /// <summary>Returns a preview of the changes as text.</summary>
    function Preview: string;

    /// <summary>Creates backups of the affected files.</summary>
    procedure CreateBackups;

    /// <summary>Applies all changes to the files.</summary>
    procedure Apply;

    /// <summary>Returns the number of affected files.</summary>
    function FileCount: Integer;

    /// <summary>Returns the total number of edits.</summary>
    function EditCount: Integer;

    /// <summary>Returns the paths of all affected files.</summary>
    function GetAffectedFiles: TArray<string>;

    property BackupDir: string read FBackupDir write FBackupDir;
    property NoBackup: Boolean read FNoBackup write FNoBackup;
  end;

implementation

{ TWorkspaceEditApplier }

constructor TWorkspaceEditApplier.Create(const AEdit: TLspWorkspaceEdit);
begin
  inherited Create;
  FEdit := AEdit;
  FBackupDir := '';
  FNoBackup := False;
end;

procedure TWorkspaceEditApplier.SortEditsReverse(var AEdits: TArray<TLspTextEdit>);
begin
  TArray.Sort<TLspTextEdit>(AEdits,
    TComparer<TLspTextEdit>.Construct(
      function(const L, R: TLspTextEdit): Integer
      begin
        // Sort in reverse: largest position first
        Result := R.Range.Start.Line - L.Range.Start.Line;
        if Result = 0 then
          Result := R.Range.Start.Character - L.Range.Start.Character;
      end));
end;

function TWorkspaceEditApplier.ApplyEditsToContent(const AContent: string; const AEdits: TArray<TLspTextEdit>): string;
var
  Lines: TStringList;
  Edit: TLspTextEdit;
  SortedEdits: TArray<TLspTextEdit>;
  StartLine, EndLine, StartChar, EndChar: Integer;
  Prefix, Suffix: string;
  NewLines: TArray<string>;
  I: Integer;
begin
  Lines := TStringList.Create;
  try
    Lines.Text := AContent;
    // Make sure an empty trailing line is preserved
    if AContent.EndsWith(sLineBreak) then
      Lines.Add('');

    SortedEdits := Copy(AEdits);
    SortEditsReverse(SortedEdits);

    for Edit in SortedEdits do
    begin
      StartLine := Edit.Range.Start.Line;
      EndLine := Edit.Range.End_.Line;
      StartChar := Edit.Range.Start.Character;
      EndChar := Edit.Range.End_.Character;

      // Check bounds
      if StartLine >= Lines.Count then
        Continue;
      if EndLine >= Lines.Count then
        EndLine := Lines.Count - 1;

      // Prefix (text before the edit on the start line)
      Prefix := Copy(Lines[StartLine], 1, StartChar);
      // Suffix (text after the edit on the end line)
      Suffix := Copy(Lines[EndLine], EndChar + 1);

      // Insert new text
      NewLines := Edit.NewText.Split([#10]);
      // Remove CR if present
      for I := 0 to High(NewLines) do
        if NewLines[I].EndsWith(#13) then
          NewLines[I] := Copy(NewLines[I], 1, Length(NewLines[I]) - 1);

      // Remove old lines
      for I := EndLine downto StartLine do
        Lines.Delete(I);

      // Insert new lines
      if Length(NewLines) = 0 then
      begin
        Lines.Insert(StartLine, Prefix + Suffix);
      end
      else if Length(NewLines) = 1 then
      begin
        Lines.Insert(StartLine, Prefix + NewLines[0] + Suffix);
      end
      else
      begin
        // First line with prefix
        Lines.Insert(StartLine, Prefix + NewLines[0]);
        // Middle lines
        for I := 1 to High(NewLines) - 1 do
          Lines.Insert(StartLine + I, NewLines[I]);
        // Last line with suffix
        Lines.Insert(StartLine + High(NewLines), NewLines[High(NewLines)] + Suffix);
      end;
    end;

    // Back to string - remove the last artificial empty string
    if (Lines.Count > 0) and (Lines[Lines.Count - 1] = '') and AContent.EndsWith(sLineBreak) then
      Lines.Delete(Lines.Count - 1);

    Result := Lines.Text;
  finally
    Lines.Free;
  end;
end;

function TWorkspaceEditApplier.Preview: string;
var
  SB: TStringBuilder;
  FE: TLspFileEdits;
  Edit: TLspTextEdit;
  Lines: TStringList;
  Content: string;
  TotalEdits: Integer;
begin
  SB := TStringBuilder.Create;
  Lines := TStringList.Create;
  try
    TotalEdits := 0;

    for FE in FEdit.FileEdits do
    begin
      SB.AppendLine('--- ' + FE.FilePath + ' ---');

      if TFile.Exists(FE.FilePath) then
      begin
        Content := ReadDelphiFile(FE.FilePath);
        Lines.Text := Content;
      end
      else
      begin
        SB.AppendLine('  (Datei nicht gefunden)');
        Continue;
      end;

      for Edit in FE.Edits do
      begin
        Inc(TotalEdits);
        var LineNo := Edit.Range.Start.Line;
        if LineNo < Lines.Count then
        begin
          SB.AppendFormat('  Zeile %d: "%s"', [LineNo + 1,
            Copy(Lines[LineNo], Edit.Range.Start.Character + 1, Edit.Range.End_.Character - Edit.Range.Start.Character)]);
          SB.AppendLine;
          SB.AppendFormat('        -> "%s"', [Edit.NewText]);
          SB.AppendLine;
        end;
      end;
      SB.AppendLine;
    end;

    SB.AppendFormat('Zusammenfassung: %d Aenderung(en) in %d Datei(en)', [TotalEdits, Length(FEdit.FileEdits)]);

    Result := SB.ToString;
  finally
    Lines.Free;
    SB.Free;
  end;
end;

procedure TWorkspaceEditApplier.CreateBackups;
var
  FE: TLspFileEdits;
  BackupPath: string;
begin
  if FNoBackup then
    Exit;

  if FBackupDir = '' then
    FBackupDir := TPath.Combine(TPath.GetTempPath, 'rename_backup_' + FormatDateTime('yyyymmdd_hhnnss', Now));

  for FE in FEdit.FileEdits do
  begin
    if not TFile.Exists(FE.FilePath) then
      Continue;

    BackupPath := TPath.Combine(FBackupDir, ExtractFileName(FE.FilePath));
    ForceDirectories(ExtractFilePath(BackupPath));
    TFile.Copy(FE.FilePath, BackupPath, True);
  end;
end;

procedure TWorkspaceEditApplier.Apply;
var
  FE: TLspFileEdits;
  Content, NewContent: string;
begin
  for FE in FEdit.FileEdits do
  begin
    if not TFile.Exists(FE.FilePath) then
    begin
      Writeln('WARNUNG: Datei nicht gefunden, uebersprungen: ', FE.FilePath);
      Continue;
    end;

    var OrigEncoding := DetectFileEncoding(FE.FilePath);
    Content := ReadDelphiFile(FE.FilePath);
    NewContent := ApplyEditsToContent(Content, FE.Edits);
    WriteDelphiFile(FE.FilePath, NewContent, OrigEncoding);
  end;
end;

function TWorkspaceEditApplier.FileCount: Integer;
begin
  Result := Length(FEdit.FileEdits);
end;

function TWorkspaceEditApplier.EditCount: Integer;
var
  FE: TLspFileEdits;
begin
  Result := 0;
  for FE in FEdit.FileEdits do
    Inc(Result, Length(FE.Edits));
end;

function TWorkspaceEditApplier.GetAffectedFiles: TArray<string>;
var
  I: Integer;
begin
  SetLength(Result, Length(FEdit.FileEdits));
  for I := 0 to High(FEdit.FileEdits) do
    Result[I] := FEdit.FileEdits[I].FilePath;
end;

end.
