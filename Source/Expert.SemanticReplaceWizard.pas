(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
//   5. On confirm: rewrite each modified file through Editor.
//      ReplaceFileContent (IOTAEditWriter, undoable, instant) and
//      augment its interface-uses clause with the units declared on
//      the rules that fired.

interface

uses
  Expert.SemanticReplace, Lsp.Client;

type
  TSemanticFilePlan = record
    FileName: string;
    Original: string;
    Matches: TArray<TSemanticReplaceMatch>;
    /// <summary>Parallel to Matches; empty = not verified (no DelphiLSP).</summary>
    Verdicts: TArray<TMatchVerdict>;
    Targets: TArray<TMatchTarget>;
  end;

  /// <summary>Progress + cancel: False stops the verification.</summary>
  TSemanticReplaceProgress = reference to function(ACur, ATotal: Integer;
    const AText: string): Boolean;

procedure EditSemanticReplaceRules;
procedure ApplySemanticReplacements_CurrentUnit;
procedure ApplySemanticReplacements_SelectedUnits;
procedure ApplySemanticReplacements_Project;

/// <summary><project root>\semantic-replace.json ('' without a project).</summary>
function SemanticRulesPath: string;

/// <summary>Asks DelphiLSP for the declaration of every match's LAST
///  identifier and decides per match (VerifyVerdicts over all matches of
///  a rule). Sends each file's Original text (so the offsets fit) - safe
///  on any thread. False = cancelled.</summary>
function VerifySemanticPlans(AClient: TLspClient; var APlans: TArray<TSemanticFilePlan>;
  const ARules: TArray<TSemanticReplaceRule>; const AProgress: TSemanticReplaceProgress;
  out ADominant: TArray<string>): Boolean;

/// <summary>Offsets of the matches that must NOT be replaced.</summary>
function SemanticSkipOffsets(const APlan: TSemanticFilePlan; AIncludeUnverified: Boolean): TArray<Integer>;

/// <summary>Human-readable verdict of one match.</summary>
function SemanticVerdictText(AVerdict: TMatchVerdict; const ATarget: TMatchTarget;
  const ARule: TSemanticReplaceRule): string;

/// <summary>Main thread: writes APlan (skipping the unverified matches as
///  asked) incl. the rules' uses units. Returns the number of replaced
///  occurrences (0 = nothing written).</summary>
function ApplySemanticPlan(const APlan: TSemanticFilePlan;
  const ARules: TArray<TSemanticReplaceRule>; AIncludeUnverified: Boolean): Integer;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.StrUtils, System.UITypes,
  System.Generics.Collections,
  Vcl.Dialogs, Vcl.Forms, Vcl.Controls,
  Winapi.Windows, System.Math,
  Expert.EditorHelperIntf, Expert.DialogHelper,
  Expert.SemanticReplaceDialogs, Expert.LspManager,
  Lsp.Protocol, Lsp.Uri, Delphi.FileEncoding;

const
  CRulesFileName = 'semantic-replace.json';

function SemanticRulesPath: string;
begin
  Result := '';
  var Root := Editor.GetProjectRoot;
  if Root <> '' then Result := IncludeTrailingPathDelimiter(Root) + CRulesFileName;
end;

function RulesFilePath(out APath: string): Boolean;
var
  Root: string;
begin
  Result := False;
  Root := Editor.GetProjectRoot;
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
    ShowThemedMessage('No active project root - cannot locate the rules file.');
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
        ShowThemedMessage('Could not create rules file: ' + E.Message); Exit;
      end;
    end;
  end;
  ARules := TSemanticReplaceEngine.LoadRules(APath, Err);
  if Err <> '' then
  begin
    ShowThemedMessage('Rules file is invalid:' + sLineBreak + Err); Exit;
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
    ShowThemedMessage('No active project root - cannot locate the rules file.');
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
  if Editor.ReadEditorContent(AFile, Tmp) then
    Result := Tmp
  else
    Result := TDelphiFileEncoding.ReadAll(AFile);
end;

procedure WriteSourceText(const AFile, AContent: string);
var
  Enc: TEncoding;
begin
  if not Editor.ReplaceFileContent(AFile, AContent) then
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

function SemanticVerdictText(AVerdict: TMatchVerdict; const ATarget: TMatchTarget;
  const ARule: TSemanticReplaceRule): string;
begin
  var Where := '';
  if ATarget.TargetFile <> '' then
    Where := Format('%s:%d', [ExtractFileName(ATarget.TargetFile), ATarget.TargetLine + 1]);
  case AVerdict of
    mvVerified: Result := 'verified -> ' + Where;
    mvOtherSymbol: Result := 'SKIPPED: another symbol of that name -> ' + Where;
    mvWrongUnit: Result := Format('SKIPPED: declared in %s, the rule expects %s',
      [Where, ARule.DeclaredIn]);
    mvNoAnswer: Result := 'NOT VERIFIED: DelphiLSP gave no answer (replaced only ' +
      'with the checkbox)';
  end;
end;

function VerifySemanticPlans(AClient: TLspClient; var APlans: TArray<TSemanticFilePlan>;
  const ARules: TArray<TSemanticReplaceRule>; const AProgress: TSemanticReplaceProgress;
  out ADominant: TArray<string>): Boolean;
var
  All: TArray<TMatchTarget>;

  function Go(ACur, ATotal: Integer; const AText: string): Boolean;
  begin
    Result := (not Assigned(AProgress)) or AProgress(ACur, ATotal, AText);
  end;

begin
  Result := False;
  ADominant := nil;
  var Total := 0;
  for var P in APlans do Inc(Total, Length(P.Matches));
  var Done := 0;
  All := nil;
  for var PI := 0 to High(APlans) do
  begin
    var F := APlans[PI].FileName;
    var Before := AClient.GetFileDiagnosticsVersion(F);
    // the text the offsets belong to - not a re-read that could differ
    if AClient.SyncDocumentWith(F, APlans[PI].Original) then
    begin
      var Cb: TSemanticReplaceProgress := AProgress;   // a nested routine cannot be captured
      var Cur := Done;
      AClient.WaitFileAnalysed(F, Before, 30000,
        function: Boolean
        begin
          Result := (not Assigned(Cb)) or Cb(Cur, Total, 'Waiting for DelphiLSP to analyse ' +
            ExtractFileName(F) + '...');
        end);
    end;
    if not Go(Done, Total, ExtractFileName(F)) then Exit;
    SetLength(APlans[PI].Targets, Length(APlans[PI].Matches));
    for var MI := 0 to High(APlans[PI].Matches) do
    begin
      Inc(Done);
      if not Go(Done, Total, Format('%s (%d/%d)', [ExtractFileName(F), Done, Total])) then Exit;
      var M := APlans[PI].Matches[MI];
      var T := Default(TMatchTarget);
      T.RuleIdx := M.RuleIdx;
      T.TargetLine := -1;
      var L, C: Integer;
      TSemanticReplaceEngine.OffsetToLineCol(APlans[PI].Original,
        TSemanticReplaceEngine.VerifyOffset(ARules[M.RuleIdx], M), L, C);
      var D: TArray<TLspLocation> := nil;
      try
        D := AClient.GotoDefinition(F, L - 1, C - 1);
        // an EMPTY answer is retried briefly (the unit may still be in
        // analysis), a wrong one never
        var Dl := GetTickCount64 + UInt64(IfThen(MI = 0, 3000, 600));
        while (Length(D) = 0) and (GetTickCount64 < Dl) do
        begin
          Sleep(150);
          if not Go(Done, Total, Format('%s (%d/%d) - waiting for an answer',
            [ExtractFileName(F), Done, Total])) then Exit;
          D := AClient.GotoDefinition(F, L - 1, C - 1);
        end;
      except
        D := nil;
      end;
      if Length(D) > 0 then
      begin
        T.TargetFile := TLspUri.FileUriToPath(D[0].Uri);
        T.TargetLine := D[0].Range.Start.Line;
      end;
      APlans[PI].Targets[MI] := T;
      All := All + [T];
    end;
  end;
  // decide over ALL matches of a rule at once (the majority is project-wide)
  var V := TSemanticReplaceEngine.VerifyVerdicts(ARules, All, ADominant);
  var K := 0;
  for var PI := 0 to High(APlans) do
  begin
    SetLength(APlans[PI].Verdicts, Length(APlans[PI].Matches));
    for var MI := 0 to High(APlans[PI].Matches) do
    begin
      APlans[PI].Verdicts[MI] := V[K];
      Inc(K);
    end;
  end;
  Result := True;
end;

// The wizard's verification: the client of the project, a progress window.
// ANote explains a verification that could not run. False = cancelled.
function VerifyPlans(var APlans: TArray<TSemanticFilePlan>;
  const ARules: TArray<TSemanticReplaceRule>; out ADominant: TArray<string>;
  out ANote: string): Boolean;
var
  Client: TLspClient;
  Prog: TCheckProgressWindow;
begin
  Result := True;
  ANote := '';
  ADominant := nil;
  var Json := Editor.FindDelphiLspJson;
  var Root := Editor.GetProjectRoot;
  var Dproj := Editor.GetCurrentProjectDproj;
  if (Json = '') or (Dproj = '') then
  begin
    ANote := 'No .delphilsp.json / project - the matches are NOT verified (text only).';
    Exit;
  end;
  try
    Client := TLspManager.Instance.GetClient(Root, Dproj, Json);
  except
    on E: Exception do
    begin
      ANote := 'DelphiLSP could not be started (' + E.Message + ') - the matches ' +
        'are NOT verified (text only).';
      Exit;
    end;
  end;
  Prog := CreateCheckProgress('Semantic replace', Application.MainForm,
    'Verifying the matches with DelphiLSP...');
  try
    Result := VerifySemanticPlans(Client, APlans, ARules,
      function(ACur, ATotal: Integer; const AText: string): Boolean
      begin
        Prog.Step(ACur, ATotal, AText);
        Result := Prog.Visible;
      end, ADominant);
  finally
    Prog.Free;
  end;
end;

// Offsets of the matches that must NOT be replaced.
function SemanticSkipOffsets(const APlan: TSemanticFilePlan; AIncludeUnverified: Boolean): TArray<Integer>;
begin
  Result := nil;
  if Length(APlan.Verdicts) = 0 then Exit;   // not verified at all: text only
  for var I := 0 to High(APlan.Matches) do
    case APlan.Verdicts[I] of
      mvOtherSymbol, mvWrongUnit: Result := Result + [APlan.Matches[I].Offset];
      mvNoAnswer: if not AIncludeUnverified then Result := Result + [APlan.Matches[I].Offset];
    end;
end;

function BuildPreviewText(const APlans: TArray<TSemanticFilePlan>;
  const ARules: TArray<TSemanticReplaceRule>; const ADominant: TArray<string>;
  const ANote: string): string;
var
  SB: TStringBuilder;
  Line, Col: Integer;
  Orig, NewLine: string;
begin
  SB := TStringBuilder.Create;
  try
    if ANote <> '' then SB.Append('!!! ').Append(ANote).AppendLine.AppendLine;
    for var R := 0 to High(ARules) do
      if (R <= High(ADominant)) and (ADominant[R] <> '') then
      begin
        if ARules[R].DeclaredIn <> '' then
          SB.AppendFormat('Rule "%s": the symbol must be declared in %s', [ARules[R].Find,
            ARules[R].DeclaredIn]).AppendLine
        else
          SB.AppendFormat('Rule "%s": the symbol is the one at %s (most matches lead there)',
            [ARules[R].Find, ADominant[R]]).AppendLine;
      end;
    if Length(ADominant) > 0 then SB.AppendLine;
    for var P in APlans do
    begin
      SB.Append('=== ').Append(ExtractFileName(P.FileName)).AppendLine;
      SB.Append('    ').Append(P.FileName).AppendLine.AppendLine;
      for var I := 0 to High(P.Matches) do
      begin
        var M := P.Matches[I];
        TSemanticReplaceEngine.OffsetToLineCol(P.Original, M.Offset, Line, Col);
        Orig := TSemanticReplaceEngine.LineAtOffset(P.Original, M.Offset);
        var R := ARules[M.RuleIdx];
        NewLine := StringReplace(Orig, R.Find, R.Replace, []);
        SB.Append('    L').Append(Line);
        if Length(P.Verdicts) > 0 then
          SB.Append('  [').Append(SemanticVerdictText(P.Verdicts[I], P.Targets[I], R)).Append(']');
        SB.AppendLine;
        SB.Append('      - ').Append(TrimLeft(Orig)).AppendLine;
        if (Length(P.Verdicts) = 0) or (P.Verdicts[I] in [mvVerified, mvNoAnswer]) then
          SB.Append('      + ').Append(TrimLeft(NewLine)).AppendLine;
        SB.AppendLine;
      end;
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

// Per-file uses to add: every rule that hit, deduped.
function UsesForHits(const ARules: TArray<TSemanticReplaceRule>;
  const AStats: TSemanticReplaceStats): TArray<string>;
begin
  Result := nil;
  var Seen := TDictionary<string, Boolean>.Create;
  try
    for var Rh in AStats.RuleHits do
      for var U in ARules[Rh].UsesToAdd do
        if not Seen.ContainsKey(UpperCase(U)) then
        begin
          Seen.Add(UpperCase(U), True);
          Result := Result + [U];
        end;
  finally
    Seen.Free;
  end;
end;

function ApplySemanticPlan(const APlan: TSemanticFilePlan;
  const ARules: TArray<TSemanticReplaceRule>; AIncludeUnverified: Boolean): Integer;
var
  Stats: TSemanticReplaceStats;
begin
  var Content := TSemanticReplaceEngine.ApplyToText(APlan.Original, ARules,
    SemanticSkipOffsets(APlan, AIncludeUnverified), Stats);
  Result := Stats.Occurrences;
  if Result = 0 then Exit;
  var UsesToAdd := UsesForHits(ARules, Stats);
  if Length(UsesToAdd) > 0 then
    AddUsesToInterfaceClause(Content, UsesToAdd);
  WriteSourceText(APlan.FileName, Content);
end;

procedure RunReplaceOver(const AFiles: TArray<string>);
var
  Rules: TArray<TSemanticReplaceRule>;
  Path, Note: string;
  Plans: TArray<TSemanticFilePlan>;
  Dominant: TArray<string>;
begin
  if Length(AFiles) = 0 then
  begin
    ShowThemedMessage('No source files to scan.'); Exit;
  end;
  if not EnsureRulesLoaded(Rules, Path) then Exit;
  Editor.SaveAllFiles;

  // 1. the text matches
  Plans := nil;
  Screen.Cursor := crHourGlass;
  try
    for var F in AFiles do
    begin
      var Plan := Default(TSemanticFilePlan);
      Plan.FileName := F;
      try
        Plan.Original := ReadSourceText(F);
      except
        Continue;
      end;
      Plan.Matches := TSemanticReplaceEngine.FindAllMatches(Plan.Original, Rules);
      if Length(Plan.Matches) > 0 then Plans := Plans + [Plan];
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  if Length(Plans) = 0 then
  begin
    ShowThemedMessage('No matches found.'); Exit;
  end;

  // 2. which of them ARE the symbol (DelphiLSP)
  if not VerifyPlans(Plans, Rules, Dominant, Note) then
  begin
    ShowThemedMessage('Semantic replace cancelled.');
    Exit;
  end;

  // 3. preview
  var Verified := 0;
  var Other := 0;
  var NoAnswer := 0;
  var Unchecked := 0;
  for var P in Plans do
    if Length(P.Verdicts) = 0 then
      Inc(Unchecked, Length(P.Matches))
    else
      for var V in P.Verdicts do
        case V of
          mvVerified: Inc(Verified);
          mvOtherSymbol, mvWrongUnit: Inc(Other);
          mvNoAnswer: Inc(NoAnswer);
        end;
  var Summary: string;
  if Unchecked > 0 then
    Summary := Format('%d file(s), %d occurrence(s) - NOT verified (text only).',
      [Length(Plans), Unchecked])
  else
    Summary := Format('%d file(s): %d occurrence(s) verified and replaced, %d skipped ' +
      '(another symbol), %d not verifiable.', [Length(Plans), Verified, Other, NoAnswer]);
  var IncludeUnverified: Boolean;
  if not TSemanticReplacePreviewDialog.Confirm(Application.MainForm, Summary,
    BuildPreviewText(Plans, Rules, Dominant, Note), NoAnswer, IncludeUnverified) then Exit;

  // 4. apply
  var TotalFiles := 0;
  var TotalEdits := 0;
  Screen.Cursor := crHourGlass;
  try
    for var P in Plans do
    begin
      var N := ApplySemanticPlan(P, Rules, IncludeUnverified);
      if N = 0 then Continue;
      Inc(TotalFiles);
      Inc(TotalEdits, N);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  ShowThemedMessage(Format('Applied to %d file(s), %d occurrence(s) replaced.',
    [TotalFiles, TotalEdits]));
end;

procedure ApplySemanticReplacements_CurrentUnit;
var
  Ctx: TEditorContext;
begin
  Ctx := Editor.GetCurrentContext;
  if not Ctx.IsValid then
  begin
    ShowThemedMessage('No file at cursor.'); Exit;
  end;
  RunReplaceOver([Ctx.FileName]);
end;

procedure ApplySemanticReplacements_SelectedUnits;
var
  AllFiles, Chosen: TArray<string>;
begin
  AllFiles := Editor.GetProjectSourceFiles;
  if Length(AllFiles) = 0 then
  begin
    ShowThemedMessage('Project source file list is empty.'); Exit;
  end;
  if not TSemanticReplaceUnitsDialog.Choose(Application.MainForm, AllFiles, Chosen) then
    Exit;
  if Length(Chosen) = 0 then
  begin
    ShowThemedMessage('No units selected.'); Exit;
  end;
  RunReplaceOver(Chosen);
end;

procedure ApplySemanticReplacements_Project;
begin
  RunReplaceOver(Editor.GetProjectSourceFiles);
end;

end.
