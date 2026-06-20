(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SemanticReplaceWizard;

// Wizard plumbing for "Semantic Replace".
//
// Four entry points (one rules editor + three scope-bound applies):
//   - EditSemanticReplaceRules
//   - ApplySemanticReplacements_CurrentUnit
//   - ApplySemanticReplacements_SelectedUnits
//   - ApplySemanticReplacements_Project
//
// All three apply paths share the same pipeline:
//   1. Load rules from <project_root>/semantic-replace.json (offer to
//      create a starter file if missing)
//   2. Pick the source file list (current unit / multi-select / all
//      project files)
//   3. Per file: dry-run apply, collect stats and matches
//   4. Build a human-readable preview and show it in the preview dialog
//   5. On confirm: rewrite each modified file through TEditorHelper.
//      ReplaceFileContent (IOTAEditWriter, undoable, instant) and
//      augment its interface-uses clause with the units declared on
//      the rules that fired.

interface

procedure EditSemanticReplaceRules;
procedure ApplySemanticReplacements_CurrentUnit;
procedure ApplySemanticReplacements_SelectedUnits;
procedure ApplySemanticReplacements_Project;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.StrUtils,
  System.Generics.Collections,
  Vcl.Dialogs, Vcl.Forms, Vcl.Controls,
  ToolsAPI,
  Expert.EditorHelper, Expert.SemanticReplace, Expert.SemanticReplaceDialogs,
  Delphi.FileEncoding;

const
  CRulesFileName = 'semantic-replace.json';

function RulesFilePath(out APath: string): Boolean;
var
  Root: string;
begin
  Result := False;
  Root := TEditorHelper.GetProjectRoot;
  if Root = '' then Exit;
  APath := IncludeTrailingPathDelimiter(Root) + CRulesFileName;
  Result := True;
end;

function EnsureRulesLoaded(out ARules: TArray<TSemanticReplaceRule>;
  out APath: string): Boolean;
var
  Err: string;
begin
  Result := False;
  ARules := nil;
  if not RulesFilePath(APath) then
  begin
    ShowMessage('No active project root - cannot locate the rules file.');
    Exit;
  end;
  if not TFile.Exists(APath) then
  begin
    if MessageDlg(
      'No semantic-replace.json found at:' + sLineBreak + APath + sLineBreak +
      sLineBreak +
      'Create a starter file with one example rule and open it in the rules editor?',
      mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;
    try
      TSemanticReplaceEngine.WriteExampleRules(APath);
    except
      on E: Exception do
      begin
        ShowMessage('Could not create rules file: ' + E.Message); Exit;
      end;
    end;
  end;
  ARules := TSemanticReplaceEngine.LoadRules(APath, Err);
  if Err <> '' then
  begin
    ShowMessage('Rules file is invalid:' + sLineBreak + Err); Exit;
  end;
  if Length(ARules) = 0 then
  begin
    if MessageDlg(
      'Rules file contains no rules. Open the rules editor?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if TSemanticReplaceRulesListDialog.Edit(Application.MainForm, ARules) then
      begin
        TSemanticReplaceEngine.SaveRules(APath, ARules);
        Result := Length(ARules) > 0;
      end;
    end;
    Exit;
  end;
  Result := True;
end;

procedure EditSemanticReplaceRules;
var
  Path, Err: string;
  Rules: TArray<TSemanticReplaceRule>;
begin
  if not RulesFilePath(Path) then
  begin
    ShowMessage('No active project root - cannot locate the rules file.');
    Exit;
  end;
  if TFile.Exists(Path) then
    Rules := TSemanticReplaceEngine.LoadRules(Path, Err)
  else
    Rules := nil;
  if TSemanticReplaceRulesListDialog.Edit(Application.MainForm, Rules) then
    TSemanticReplaceEngine.SaveRules(Path, Rules);
end;

function ReadSourceText(const AFile: string): string;
var
  Tmp: string;
begin
  if TEditorHelper.ReadEditorContent(AFile, Tmp) then
    Result := Tmp
  else
    Result := TDelphiFileEncoding.ReadAll(AFile);
end;

procedure WriteSourceText(const AFile, AContent: string);
var
  Enc: TEncoding;
begin
  if not TEditorHelper.ReplaceFileContent(AFile, AContent) then
  begin
    if TFile.Exists(AFile) then Enc := TDelphiFileEncoding.Detect(AFile)
    else Enc := TEncoding.UTF8;
    TDelphiFileEncoding.WriteAll(AFile, AContent, Enc);
  end;
end;

function AddUsesToInterfaceClause(var AContent: string;
  const AUnitsToAdd: TArray<string>): Boolean;
// Appends AUnitsToAdd (deduped, case-insensitive) to the END of the
// content's interface-section uses clause.
var
  Lines: TArray<string>;
  I, IfaceLine, UsesLine, EndLine: Integer;
  U, Acc, LL, Before, Indent, Sep, Tail: string;
  Buf: TStringList;
  ToAdd: TArray<string>;
  PresentSet: TDictionary<string, Boolean>;
  SemPos, K: Integer;
begin
  Result := False;
  if Length(AUnitsToAdd) = 0 then Exit;
  Lines := AContent.Split([sLineBreak], TStringSplitOptions.None);
  IfaceLine := -1;
  UsesLine := -1;
  for I := 0 to High(Lines) do
  begin
    U := UpperCase(Trim(Lines[I]));
    if U = 'INTERFACE' then IfaceLine := I
    else if U = 'IMPLEMENTATION' then Break
    else if (IfaceLine >= 0) and StartsText('USES', U) then
    begin
      UsesLine := I; Break;
    end;
  end;
  if (IfaceLine < 0) or (UsesLine < 0) then Exit;

  EndLine := UsesLine;
  Acc := '';
  for I := UsesLine to High(Lines) do
  begin
    Acc := Acc + ' ' + Lines[I];
    if Pos(';', Lines[I]) > 0 then begin EndLine := I; Break; end;
  end;
  PresentSet := TDictionary<string, Boolean>.Create;
  try
    for var Tok in Acc.Split([',', ';', ' ']) do
    begin
      var T := Trim(Tok);
      if (T = '') or SameText(T, 'uses') then Continue;
      PresentSet.AddOrSetValue(UpperCase(T), True);
    end;
    ToAdd := nil;
    for U in AUnitsToAdd do
      if not PresentSet.ContainsKey(UpperCase(U)) then ToAdd := ToAdd + [U];
    if Length(ToAdd) = 0 then Exit;
  finally
    PresentSet.Free;
  end;
  LL := Lines[EndLine];
  SemPos := Pos(';', LL);
  if SemPos = 0 then Exit;
  Before := Copy(LL, 1, SemPos - 1);
  while (Before <> '') and (Before[Length(Before)] = ' ') do
    Before := Copy(Before, 1, Length(Before) - 1);
  if EndLine = UsesLine then
  begin
    Sep := ', ';
    LL := Before + Sep + string.Join(', ', ToAdd) + Copy(LL, SemPos, MaxInt);
    Lines[EndLine] := LL;
  end
  else
  begin
    Indent := '';
    K := 1;
    while (K <= Length(LL)) and (LL[K] = ' ') do
    begin
      Indent := Indent + ' '; Inc(K);
    end;
    if Indent = '' then Indent := '  ';
    Lines[EndLine] := Before + ',';
    Buf := TStringList.Create;
    try
      for I := 0 to EndLine do Buf.Add(Lines[I]);
      for I := 0 to High(ToAdd) do
      begin
        if I = High(ToAdd) then Tail := Copy(LL, SemPos, MaxInt)
        else Tail := ',';
        Buf.Add(Indent + ToAdd[I] + Tail);
      end;
      for I := EndLine + 1 to High(Lines) do Buf.Add(Lines[I]);
      SetLength(Lines, Buf.Count);
      for I := 0 to Buf.Count - 1 do Lines[I] := Buf[I];
    finally
      Buf.Free;
    end;
  end;
  AContent := string.Join(sLineBreak, Lines);
  Result := True;
end;

type
  TFilePlan = record
    FileName: string;
    NewContent: string;
    UsesToAdd: TArray<string>;
    Stats: TSemanticReplaceStats;
    Matches: TArray<TSemanticReplaceMatch>;
    Original: string;
  end;

function BuildPreviewText(const APlans: TArray<TFilePlan>;
  const ARules: TArray<TSemanticReplaceRule>): string;
var
  SB: TStringBuilder;
  P: TFilePlan;
  M: TSemanticReplaceMatch;
  Line, Col: Integer;
  Orig, NewLine: string;
  RuleMap: TDictionary<Integer, Integer>;   // ruleIdx -> count
  WithVarMap: TDictionary<Int64, Boolean>;
begin
  SB := TStringBuilder.Create;
  RuleMap := TDictionary<Integer, Integer>.Create;
  try
    for P in APlans do
    begin
      SB.Append('=== ').Append(ExtractFileName(P.FileName)).Append(' ').AppendLine;
      SB.Append('    ').Append(P.FileName).AppendLine.AppendLine;
      if Length(P.UsesToAdd) > 0 then
      begin
        SB.Append('    uses += ').Append(string.Join(', ', P.UsesToAdd))
          .AppendLine.AppendLine;
      end;
      // Recompute rule-in-method counts to mirror local-var logic.
      RuleMap.Clear;
      for M in P.Matches do
      begin
        var Cnt: Integer;
        if RuleMap.TryGetValue(M.RuleIdx, Cnt) then RuleMap[M.RuleIdx] := Cnt + 1
        else RuleMap.Add(M.RuleIdx, 1);
      end;
      for M in P.Matches do
      begin
        TSemanticReplaceEngine.OffsetToLineCol(P.Original, M.Offset, Line, Col);
        Orig := TSemanticReplaceEngine.LineAtOffset(P.Original, M.Offset);
        var R := ARules[M.RuleIdx];
        var Replacement := R.Replace;
        if (R.LocalVarName <> '') and (R.LocalVarType <> '') and
           (R.LocalVarValue <> '') and (R.ReplaceWhenLocalVar <> '') and
           RuleMap.ContainsKey(M.RuleIdx) and (RuleMap[M.RuleIdx] >= 2) then
          Replacement := R.ReplaceWhenLocalVar;
        // Replace the actual matched text inside the line for the
        // "after" view; keep surrounding code intact.
        NewLine := StringReplace(Orig, R.Find, Replacement, []);
        SB.Append('    L').Append(Line).Append(':').AppendLine;
        SB.Append('      - ').Append(TrimLeft(Orig)).AppendLine;
        SB.Append('      + ').Append(TrimLeft(NewLine)).AppendLine.AppendLine;
      end;
      if P.Stats.LocalVarsIntroduced > 0 then
        SB.Append('    -- ').Append(P.Stats.LocalVarsIntroduced)
          .Append(' local var(s) will be hoisted right after BEGIN.')
          .AppendLine.AppendLine;
    end;
    Result := SB.ToString;
  finally
    RuleMap.Free;
    SB.Free;
  end;
  if WithVarMap = nil then ;
end;

procedure RunReplaceOver(const AFiles: TArray<string>);
var
  Rules: TArray<TSemanticReplaceRule>;
  Path: string;
  Plans: TList<TFilePlan>;
  TotalEdits, TotalLocalVars, TotalFiles: Integer;
  Preview, Summary: string;
  P: TFilePlan;
begin
  if Length(AFiles) = 0 then
  begin
    ShowMessage('No source files to scan.'); Exit;
  end;
  if not EnsureRulesLoaded(Rules, Path) then Exit;
  TEditorHelper.SaveAllFiles;

  Plans := TList<TFilePlan>.Create;
  TotalEdits := 0;
  TotalLocalVars := 0;
  TotalFiles := 0;
  try
    Screen.Cursor := crHourGlass;
    try
      for var F in AFiles do
      begin
        var Plan: TFilePlan;
        Plan.FileName := F;
        try
          Plan.Original := ReadSourceText(F);
        except
          Continue;
        end;
        Plan.NewContent := TSemanticReplaceEngine.ApplyToText(
          Plan.Original, Rules, Plan.Stats);
        if Plan.Stats.Occurrences = 0 then Continue;
        Plan.Matches := TSemanticReplaceEngine.FindAllMatches(Plan.Original, Rules);
        // Per-file uses to add: every rule that hit, deduped.
        Plan.UsesToAdd := nil;
        var Seen: TDictionary<string, Boolean> :=
          TDictionary<string, Boolean>.Create;
        try
          for var Rh in Plan.Stats.RuleHits do
            for var U in Rules[Rh].UsesToAdd do
              if not Seen.ContainsKey(UpperCase(U)) then
              begin
                Seen.Add(UpperCase(U), True);
                Plan.UsesToAdd := Plan.UsesToAdd + [U];
              end;
        finally
          Seen.Free;
        end;
        Plans.Add(Plan);
        Inc(TotalEdits, Plan.Stats.Occurrences);
        Inc(TotalLocalVars, Plan.Stats.LocalVarsIntroduced);
        Inc(TotalFiles);
      end;
    finally
      Screen.Cursor := crDefault;
    end;

    if TotalFiles = 0 then
    begin
      ShowMessage('No matches found.'); Exit;
    end;

    Preview := BuildPreviewText(Plans.ToArray, Rules);
    Summary := Format(
      '%d file(s), %d occurrence(s), %d local var(s) to hoist.',
      [TotalFiles, TotalEdits, TotalLocalVars]);

    if not TSemanticReplacePreviewDialog.Confirm(Application.MainForm,
      Summary, Preview) then Exit;

    Screen.Cursor := crHourGlass;
    try
      for P in Plans do
      begin
        var Content: string := P.NewContent;
        if Length(P.UsesToAdd) > 0 then
          AddUsesToInterfaceClause(Content, P.UsesToAdd);
        WriteSourceText(P.FileName, Content);
      end;
    finally
      Screen.Cursor := crDefault;
    end;
    ShowMessage(Format('Applied to %d file(s), %d occurrence(s) replaced.',
      [TotalFiles, TotalEdits]));
  finally
    Plans.Free;
  end;
end;

procedure ApplySemanticReplacements_CurrentUnit;
var
  Ctx: TEditorContext;
begin
  Ctx := TEditorHelper.GetCurrentContext;
  if not Ctx.IsValid then
  begin
    ShowMessage('No file at cursor.'); Exit;
  end;
  RunReplaceOver([Ctx.FileName]);
end;

procedure ApplySemanticReplacements_SelectedUnits;
var
  AllFiles, Chosen: TArray<string>;
begin
  AllFiles := TEditorHelper.GetProjectSourceFiles;
  if Length(AllFiles) = 0 then
  begin
    ShowMessage('Project source file list is empty.'); Exit;
  end;
  if not TSemanticReplaceUnitsDialog.Choose(Application.MainForm, AllFiles, Chosen) then
    Exit;
  if Length(Chosen) = 0 then
  begin
    ShowMessage('No units selected.'); Exit;
  end;
  RunReplaceOver(Chosen);
end;

procedure ApplySemanticReplacements_Project;
begin
  RunReplaceOver(TEditorHelper.GetProjectSourceFiles);
end;

end.
