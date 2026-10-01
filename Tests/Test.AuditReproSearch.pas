(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Audit repro tests for the area "search, rename, references, semantic
///  replace" (issue #22). Each test names the audit item it reproduces.
///  RED means the item is still open; a test can be deleted once it goes
///  green. The tests need no IDE and no DelphiLSP.
/// </summary>
unit Test.AuditReproSearch;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TAuditReproSearchTests = class
  public
    [Test] procedure M16_RenameDialog_CancelEnabledAgainAfterStop;
    [Test] procedure M17_RenameDialog_ScopeBoxesDisabledWhileBusy;
    [Test] procedure M18b_SemanticReplace_MatchesIgnoreCase;
    [Test] procedure M21a_ImplementationScan_SkipsHeaderInBlockComment;
    [Test] procedure L5b_UnitsDialog_KeepsTicksAcrossFilter;
    [Test] procedure L5m_KeyBindingHeader_SpellsTheAuthorsName;
    [Test] procedure L5n_Rename_KeepsLfLineBreaks;
    [Test] procedure L5n_Rename_AddsNoFinalLineBreak;
    [Test] procedure L6b_BareRaiseMessage_ShowsTheCount;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, Vcl.StdCtrls, Vcl.CheckLst,
  Delphi.FileEncoding, Lsp.Protocol, Rename.WorkspaceEdit, Expert.RenameDialog,
  Expert.SemanticReplace, Expert.SemanticReplaceDialogs, Expert.ImplementationFinder,
  Expert.FindReferencesDialog, Expert.SelectionValidator;

// the repository root: the first folder upwards that has Source\Expert.Version.pas
function RepoRoot: string;
begin
  Result := ExtractFileDir(ParamStr(0));
  while Result <> '' do
  begin
    if TFile.Exists(TPath.Combine(Result, 'Source\Expert.Version.pas')) then Exit;
    var Up := ExtractFileDir(Result);
    if SameText(Up, Result) then Break;
    Result := Up;
  end;
  Result := '';
end;

function TempPasFile(const AContent: string): string;
begin
  Result := TPath.Combine(TPath.GetTempPath, 'AuditRepro_' + TGUID.NewGuid.ToString + '.pas');
  TFile.WriteAllText(Result, AContent, TEncoding.ASCII);
end;

// line breaks made visible, so a failure message shows CR/LF differences
function Visible(const AText: string): string;
begin
  Result := AText.Replace(#13, '\r').Replace(#10, '\n');
end;

function FindButton(AOwner: TComponent; const ACaption: string): TButton;
begin
  for var I := 0 to AOwner.ComponentCount - 1 do
    if (AOwner.Components[I] is TButton) and (TButton(AOwner.Components[I]).Caption = ACaption) then
      Exit(TButton(AOwner.Components[I]));
  Result := nil;
end;

function FindCheckBox(AOwner: TComponent; const ACaptionStart: string): TCheckBox;
begin
  for var I := 0 to AOwner.ComponentCount - 1 do
    if (AOwner.Components[I] is TCheckBox) and
       string(TCheckBox(AOwner.Components[I]).Caption).StartsWith(ACaptionStart) then
      Exit(TCheckBox(AOwner.Components[I]));
  Result := nil;
end;

// Renames 'foo' (line 1, columns 0..3) to 'Bar' through the public Apply
// and returns the file content afterwards.
function RenameLineOneThroughApply(const AContent: string): string;
var
  Edit: TLspWorkspaceEdit;
  TE: TLspTextEdit;
begin
  var F := TempPasFile(AContent);
  try
    TE.Range.Start.Line := 1;
    TE.Range.Start.Character := 0;
    TE.Range.End_.Line := 1;
    TE.Range.End_.Character := 3;
    TE.NewText := 'Bar';
    SetLength(Edit.FileEdits, 1);
    Edit.FileEdits[0].FilePath := F;
    Edit.FileEdits[0].Edits := [TE];
    var Applier := TWorkspaceEditApplier.Create(Edit);
    try
      Applier.NoBackup := True;
      Applier.Apply;
    finally
      Applier.Free;
    end;
    Result := ReadDelphiFile(F);
  finally
    TFile.Delete(F);
  end;
end;

/// <summary>M16: after one Stop the button is Cancel again but stays
///  disabled for the rest of the dialog's life.</summary>
procedure TAuditReproSearchTests.M16_RenameDialog_CancelEnabledAgainAfterStop;
begin
  var Dlg := TRenameDialog.CreateDialog(nil, 'Foo');
  try
    Dlg.SetBusy(True);
    var Btn := FindButton(Dlg, 'Stop');
    Assert.IsNotNull(Btn, 'the Cancel button reads Stop while busy');
    Btn.Click;
    Assert.IsTrue(Dlg.ScanCancelled, 'Stop cancels the scan');
    Dlg.SetBusy(False);
    Assert.AreEqual('Cancel', Btn.Caption, False, 'the button is Cancel again');
    Assert.IsTrue(Btn.Enabled, 'M16: after one Stop, Cancel stays disabled once the scan is over');
  finally
    Dlg.Free;
  end;
end;

/// <summary>M17: SetBusy disables scope combo, pick button and name edit,
///  but not the two "plus units" boxes - ticking one during a scan clears
///  the preview, which the running scan then refills for the old scope.</summary>
procedure TAuditReproSearchTests.M17_RenameDialog_ScopeBoxesDisabledWhileBusy;
begin
  var Dlg := TRenameDialog.CreateDialog(nil, 'Foo');
  try
    var ChkOpen := FindCheckBox(Dlg, 'plus units open');
    var ChkUsed := FindCheckBox(Dlg, 'plus units reachable');
    Assert.IsNotNull(ChkOpen, 'open-units box exists');
    Assert.IsNotNull(ChkUsed, 'used-units box exists');
    // the idle state for the whole-project scope (UpdateScopeControls), set
    // directly so the remembered scope of this machine does not matter
    ChkOpen.Enabled := True;
    ChkUsed.Enabled := True;
    Dlg.SetBusy(True);
    try
      Assert.IsFalse(ChkOpen.Enabled or ChkUsed.Enabled,
        'M17: the "plus units" scope boxes stay enabled during a scan');
    finally
      Dlg.SetBusy(False);
    end;
  finally
    Dlg.Free;
  end;
end;

/// <summary>M18b: a Pascal identifier is case-insensitive, the rule match
///  is not - 'foo' and 'FOO' stay on the old API without a message.</summary>
procedure TAuditReproSearchTests.M18b_SemanticReplace_MatchesIgnoreCase;
var
  Rule: TSemanticReplaceRule;
  Stats: TSemanticReplaceStats;
begin
  Rule := Default(TSemanticReplaceRule);
  Rule.Find := 'Foo';
  Rule.Replace := 'Bar';
  var Res := TSemanticReplaceEngine.ApplyToText('  foo := 1;'#13#10'  FOO(x);', [Rule], Stats);
  Assert.AreEqual('  Bar := 1;'#13#10'  Bar(x);', Res, False,
    'M18b: rule Foo -> Bar must also match foo and FOO');
  Assert.AreEqual(2, Stats.Occurrences, 'M18b: two occurrences');
end;

/// <summary>M21a: the project scan runs over raw lines, so a header inside
///  a block comment is listed as an implementation (and becomes a rename
///  target).</summary>
procedure TAuditReproSearchTests.M21a_ImplementationScan_SkipsHeaderInBlockComment;
begin
  var F := TempPasFile(string.Join(#13#10, TArray<string>.Create(
    'unit U;',                     // 0
    'interface',                   // 1
    'type',                        // 2
    '  TFoo = class',              // 3
    '    procedure Bar;',          // 4
    '  end;',                      // 5
    'implementation',              // 6
    '{ old version',               // 7
    'procedure TFoo.Bar;',         // 8  inside the comment
    '}',                           // 9
    'procedure TFoo.Bar;',         // 10 the real one
    'begin',
    'end;',
    'end.')));
  try
    var Items := TImplementationFinder.FindByProjectScan([F], 'Bar', '');
    var Lines := '';
    for var It in Items do
      Lines := Lines + ' ' + IntToStr(It.Line);
    Assert.AreEqual(1, Integer(Length(Items)), 'M21a: header inside { } listed as implementation, hits at lines' + Lines);
    Assert.AreEqual(10, Items[0].Line, 'M21a: the real implementation');
  finally
    TFile.Delete(F);
  end;
end;

/// <summary>L5b: RefreshList clears the list and re-adds every item
///  unticked, so typing in the filter loses all ticks.</summary>
procedure TAuditReproSearchTests.L5b_UnitsDialog_KeepsTicksAcrossFilter;
begin
  var Dlg := TSemanticReplaceUnitsDialog.CreateDialog(nil, ['C:\p\Alpha.pas', 'C:\p\Beta.pas']);
  try
    var Btn := FindButton(Dlg, 'Select all');
    Assert.IsNotNull(Btn, 'Select all button exists');
    Btn.Click;
    Assert.AreEqual(2, Integer(Length(Dlg.SelectedFiles)), 'both units ticked');
    var Filter: TEdit := nil;
    var List: TCheckListBox := nil;
    for var I := 0 to Dlg.ComponentCount - 1 do
      if Dlg.Components[I] is TEdit then
        Filter := TEdit(Dlg.Components[I])
      else if Dlg.Components[I] is TCheckListBox then
        List := TCheckListBox(Dlg.Components[I]);
    Assert.IsNotNull(Filter, 'filter edit exists');
    Assert.IsNotNull(List, 'unit list exists');
    Filter.Text := 'Alpha';
    Assert.AreEqual(1, List.Items.Count, 'the filter applies');
    Filter.Text := '';
    Assert.AreEqual(2, List.Items.Count, 'the filter is cleared');
    Assert.AreEqual(2, Integer(Length(Dlg.SelectedFiles)),
      'L5b: ticks are lost when the filter changes');
  finally
    Dlg.Free;
  end;
end;

/// <summary>L5m: line 2 of Expert.KeyBinding.pas holds EF BF BD (U+FFFD)
///  in place of the a-umlaut in the author's name.</summary>
procedure TAuditReproSearchTests.L5m_KeyBindingHeader_SpellsTheAuthorsName;
begin
  var Root := RepoRoot;
  Assert.IsTrue(Root <> '', 'repository root found from the test executable');
  var Lines := TFile.ReadAllLines(TPath.Combine(Root, 'Source\Expert.KeyBinding.pas'), TEncoding.UTF8);
  Assert.IsTrue(Length(Lines) > 1, 'file read');
  Assert.IsTrue(Lines[1].Contains('J' + #$00E4 + 'nicke'),
    Format('L5m: line 2 should spell the name with U+00E4, it holds U+FFFD: %s',
    [BoolToStr(Lines[1].Contains(#$FFFD), True)]));
end;

/// <summary>L5n: ApplyEditsToContent joins with sLineBreak, so a one-token
///  rename turns an LF-only file into CRLF.</summary>
procedure TAuditReproSearchTests.L5n_Rename_KeepsLfLineBreaks;
begin
  var Res := RenameLineOneThroughApply('a'#10'foo'#10'c'#10);
  Assert.AreEqual(Visible('a'#10'Bar'#10'c'#10), Visible(Res), False,
    'L5n: LF-only file rewritten with CRLF (whole-file diff for one token)');
end;

/// <summary>L5n: ... and a file without a final line break gains one.</summary>
procedure TAuditReproSearchTests.L5n_Rename_AddsNoFinalLineBreak;
begin
  var Res := RenameLineOneThroughApply('a'#13#10'foo');
  Assert.AreEqual(Visible('a'#13#10'Bar'), Visible(Res), False, 'L5n: a final line break was added');
end;

/// <summary>L6b: the bare-raise message passes the count to Format but has
///  no %d, unlike the Exit/Break/Continue messages beside it.</summary>
procedure TAuditReproSearchTests.L6b_BareRaiseMessage_ShowsTheCount;
begin
  var R := TSelectionValidator.Validate('raise;'#13#10'raise;', nil, 1, 2, 5, '');
  var Msg := '';
  for var Issue in R.Issues do
    if Issue.Message.Contains('"raise;"') then
      Msg := Issue.Message;
  Assert.IsTrue(Msg <> '', 'the bare-raise error is reported');
  Assert.IsTrue(Msg.Contains('2'), 'L6b: the bare-raise count is not shown: ' + Msg);
end;

initialization
  TDUnitX.RegisterTestFixture(TAuditReproSearchTests);

end.
