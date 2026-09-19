(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.IncludeExpander;

// "Expand include files" (user request 2026-09-19): writes the content of
// every {$I}/{$INCLUDE} file IN PLACE into the including source, framed by
// marker comments (Expert.IncludeExpansion.ExpandIncludesMarked) - so the
// code can be debugged where it runs. Going back is the version control
// system's job; the markers show where each block came from.
//
// Four reaches, like "Remove with": current unit, selected units, a
// directory (recursive) and the whole project. Open files are changed in
// the editor buffer (undoable, not saved), closed ones on disk - both
// through ApplyLinesMinimal, so only the changed lines are touched.

interface

type
  TIncludeExpandResult = record
    FilesChanged: Integer;
    Includes: Integer;
    Failed: Integer;
    Report: string;     // one line per changed / failed file
  end;

/// <summary>Expands the includes of AFiles (main thread - editor buffers).
///  Files without an include directive are left alone.</summary>
function ExpandIncludesInFiles(const AFiles: TArray<string>): TIncludeExpandResult;

/// <summary>AFiles that contain at least one include directive (cheap text
///  check, buffer or disk).</summary>
function FilesWithIncludes(const AFiles: TArray<string>): TArray<string>;

procedure ExpandIncludesCurrentUnit;
procedure ExpandIncludesSelectedUnits;
procedure ExpandIncludesInDirectory;
procedure ExpandIncludesProjectWide;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Types,
  Vcl.Forms, Vcl.Dialogs,
  Expert.EditorHelperIntf, Expert.IncludeExpansion, Expert.ScopeFiles,
  Expert.UsesEditor, Expert.DialogHelper, Expert.WithRefactorDialog,
  Expert.McpServer, Delphi.FileEncoding, System.JSON;

function ReadAny(const AFile: string; out AContent: string): Boolean;
begin
  Result := EditorOrDiskReader()(AFile, AContent);
end;

function HasIncludeDirective(const AText: string): Boolean;
begin
  // the exact test is the expansion itself - this only avoids work
  var U := UpperCase(AText);
  Result := (Pos('{$I ', U) > 0) or (Pos('{$I'#9, U) > 0) or (Pos('{$INCLUDE', U) > 0) or
    (Pos('(*$I', U) > 0);
end;

function FilesWithIncludes(const AFiles: TArray<string>): TArray<string>;
var
  C: string;
begin
  Result := nil;
  for var F in AFiles do
    if ReadAny(F, C) and HasIncludeDirective(C) then
      Result := Result + [F];
end;

function ExpandIncludesInFiles(const AFiles: TArray<string>): TIncludeExpandResult;
var
  Original, Expanded: string;
  Count: Integer;
  SL: TStringList;
begin
  Result := Default(TIncludeExpandResult);
  for var F in AFiles do
  begin
    if not ReadAny(F, Original) then Continue;
    Expanded := ExpandIncludesMarked(Original, ExtractFileDir(ExpandFileName(F)),
      EditorOrDiskReader(), Count);
    if Count = 0 then Continue;
    SL := TStringList.Create;
    try
      SL.Text := Expanded;
      if ApplyLinesMinimal(F, SL, Original) then
      begin
        Inc(Result.FilesChanged);
        Inc(Result.Includes, Count);
        Result.Report := Result.Report + Format('%s: %d include(s) expanded',
          [ExtractFileName(F), Count]) + sLineBreak;
      end
      else
      begin
        Inc(Result.Failed);
        Result.Report := Result.Report + ExtractFileName(F) + ': could NOT be written' + sLineBreak;
      end;
    finally
      SL.Free;
    end;
  end;
end;

procedure RunOn(const AFiles: TArray<string>; const AWhat: string; AConfirm: Boolean);
begin
  var Todo := FilesWithIncludes(AFiles);
  if Length(Todo) = 0 then
  begin
    ShowThemedMessage('Expand include files: ' + AWhat + ' contains no include directive.');
    Exit;
  end;
  if AConfirm and not AskThemedConfirm(Format(
    'Expand the include files in %d file(s) (%s)?' + sLineBreak + sLineBreak +
    'Every {$I} / {$INCLUDE} is replaced by the file''s content between marker ' +
    'comments. Files open in the editor are changed in the buffer (undoable, not ' +
    'saved), all others ON DISK - revert them with your version control system.',
    [Length(Todo), AWhat]), 'Expand') then Exit;
  var R := ExpandIncludesInFiles(Todo);
  var Msg := Format('%d include(s) expanded in %d file(s).', [R.Includes, R.FilesChanged]);
  if R.Failed > 0 then
    Msg := Msg + Format(' %d file(s) could NOT be written.', [R.Failed]);
  ShowThemedMessage(Msg + sLineBreak + sLineBreak + Trim(R.Report));
end;

procedure ExpandIncludesCurrentUnit;
begin
  if Editor = nil then Exit;
  var F := Editor.GetActiveFileName;
  if F = '' then
  begin
    ShowThemedMessage('Expand include files: no source file is active.');
    Exit;
  end;
  // one file, open in the editor: undoable - no confirmation needed
  RunOn([F], ExtractFileName(F), False);
end;

procedure ExpandIncludesSelectedUnits;
var
  Chosen: TArray<string>;
begin
  if Editor = nil then Exit;
  // offer only what can change
  var Candidates := FilesWithIncludes(Editor.GetProjectSourceFiles);
  if Length(Candidates) = 0 then
  begin
    ShowThemedMessage('Expand include files: no project file contains an include directive.');
    Exit;
  end;
  if not TWithRefactorDialog.PickFiles(Application.MainForm, Candidates, Chosen) or
     (Length(Chosen) = 0) then Exit;
  RunOn(Chosen, Format('%d selected unit(s)', [Length(Chosen)]), True);
end;

procedure ExpandIncludesInDirectory;
var
  Dlg: TFileOpenDialog;
  Dir: string;
begin
  Dlg := TFileOpenDialog.Create(nil);
  try
    Dlg.Title := 'Expand include files - choose a directory (sub-directories included)';
    Dlg.Options := Dlg.Options + [fdoPickFolders, fdoPathMustExist];
    if Editor <> nil then
      Dlg.DefaultFolder := ExtractFileDir(Editor.GetActiveFileName);
    if not Dlg.Execute then Exit;
    Dir := Dlg.FileName;
  finally
    Dlg.Free;
  end;
  var Files: TArray<string> := nil;
  for var Mask in ['*.pas', '*.dpr', '*.dpk'] do
    Files := Files + TArray<string>(TDirectory.GetFiles(Dir, Mask, TSearchOption.soAllDirectories));
  RunOn(Files, Dir, True);
end;

procedure ExpandIncludesProjectWide;
begin
  if Editor = nil then Exit;
  RunOn(Editor.GetProjectSourceFiles, 'the whole project', True);
end;

// MCP tool "expand_includes": file | files | directory | project=true
function ToolExpandIncludes(AArgs: TJSONObject; AStop: THandle): string;
var
  Files: TArray<string>;
  R: TIncludeExpandResult;
  Err, RunErr: string;
begin
  Files := nil;
  var F := AArgs.GetValue<string>('file', '');
  if F <> '' then Files := [ExpandFileName(F)];
  var Arr := AArgs.GetValue<TJSONArray>('files', nil);
  if Arr <> nil then
    for var V in Arr do Files := Files + [ExpandFileName(V.Value)];
  var Dir := AArgs.GetValue<string>('directory', '');
  if Dir <> '' then
  begin
    if not TDirectory.Exists(Dir) then Exit(McpErr('directory not found: ' + Dir));
    for var Mask in ['*.pas', '*.dpr', '*.dpk'] do
      Files := Files + TArray<string>(TDirectory.GetFiles(Dir, Mask, TSearchOption.soAllDirectories));
  end;
  var Project := AArgs.GetValue<Boolean>('project', False);
  if (Length(Files) = 0) and not Project then
    Exit(McpErr('pass "file", "files", "directory" or "project": true'));
  RunErr := '';
  if not McpRunOnMain(
    procedure
    begin
      try
        var All := Files;
        if Project and (Editor <> nil) then All := All + Editor.GetProjectSourceFiles;
        R := ExpandIncludesInFiles(All);
      except
        on E: Exception do RunErr := E.Message;
      end;
    end, False, AStop, Err) then Exit(McpErr(Err));
  if RunErr <> '' then Exit(McpErr(RunErr));
  var J := TJSONObject.Create;
  J.AddPair('files_changed', TJSONNumber.Create(R.FilesChanged));
  J.AddPair('includes_expanded', TJSONNumber.Create(R.Includes));
  J.AddPair('failed', TJSONNumber.Create(R.Failed));
  J.AddPair('report', Trim(R.Report));
  J.AddPair('note', 'open files were changed in the editor buffer (not saved), closed ' +
    'files on disk; revert with the version control system');
  Result := McpOk(J);
end;

initialization
  RegisterMcpTool('expand_includes', ToolExpandIncludes);

end.
