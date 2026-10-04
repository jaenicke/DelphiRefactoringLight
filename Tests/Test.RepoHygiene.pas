(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Repository hygiene (issue #11, suggestion 8 by Ian Branch): every Delphi
///  source is UTF-8 WITH BOM and uses CRLF only, text form files use CRLF,
///  and .gitattributes pins the line endings. LF-only sources are not
///  cosmetic - the debugger can put breakpoints on the wrong lines. The
///  tests find the repository from the test executable's folder upwards.
/// </summary>
unit Test.RepoHygiene;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TRepoHygieneTests = class
  public
    [Test] procedure Sources_HaveBomAndCrLfOnly;
    [Test] procedure TextForms_UseCrLfOnly;
    [Test] procedure GitAttributes_PinDelphiLineEndings;
    /// <summary>
    ///  Every label block of the install scripts must END with a goto or an
    ///  exit. A block that falls into the NEXT label is invisible in the
    ///  source and loud on screen: from 1.15.2 to 1.16.30 install.cmd
    ///  printed the framed "the MCP BRIDGE IS NOT installed" warning after
    ///  every SUCCESSFUL install, because the "bridge already registered"
    ///  block had no terminator and ran straight into :mcp_failed.
    /// </summary>
    [Test] procedure InstallScripts_HaveNoLabelFallThrough;
    /// <summary>User, 2026-10-04: the bridge stayed four versions behind
    ///  although every install.cmd succeeded. The aside name was FIXED, so an
    ///  .old a bridge still ran from could not be deleted, the move had
    ///  nowhere to go and the live exe could not be overwritten - and an
    ///  install while a Claude Code session runs is the normal case.</summary>
    [Test] procedure TheBridgeIsMovedAsideUnderAUniqueName;
    /// <summary>
    ///  GetCurrentContext walks the edit position to read the word under the
    ///  cursor, and that COLLAPSES an active selection - the helper says so
    ///  at GetActiveFileName and the context menu says it again. A routine
    ///  that asks for the context FIRST and the selection afterwards
    ///  therefore never sees one: "Edit methods" shipped that way in 1.18.5
    ///  and the user watched their three marked methods disappear when they
    ///  clicked the menu entry. So no routine may call GetCurrentContext
    ///  above its own GetSelection.
    /// </summary>
    [Test] procedure NoRoutineReadsTheContextBeforeTheSelection;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Types,
  Expert.PascalScanner;

const
  SourceDirs: array[0..5] of string = ('Source', 'Packages', 'Standalone', 'Tests', 'Mcp', 'dih');

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

// build output, IDE history and the like are not sources
function Skipped(const AFile: string): Boolean;
begin
  var U := '\' + UpperCase(AFile);
  Result := U.Contains('\WIN32\') or U.Contains('\WIN64\') or U.Contains('\DCU\') or
    U.Contains('\__HISTORY\') or U.Contains('\__RECOVERY\') or U.Contains('\OUTPUT\') or
    U.Contains('\BIN\') or U.Contains('\BACKUP\');
end;

function Files(const AMasks: array of string): TArray<string>;
begin
  Result := nil;
  var Root := RepoRoot;
  if Root = '' then Exit;
  for var D in SourceDirs do
  begin
    var Dir := TPath.Combine(Root, D);
    if not TDirectory.Exists(Dir) then Continue;
    for var M in AMasks do
      for var F in TDirectory.GetFiles(Dir, M, TSearchOption.soAllDirectories) do
        if not Skipped(F) then Result := Result + [F];
  end;
end;

// '' when the bytes use CRLF only, else what is wrong
function LineEndingProblem(const ABytes: TBytes): string;
begin
  Result := '';
  for var I := 0 to High(ABytes) do
    if (ABytes[I] = 10) and ((I = 0) or (ABytes[I - 1] <> 13)) then
      Exit('LF without CR')
    else if (ABytes[I] = 13) and ((I = High(ABytes)) or (ABytes[I + 1] <> 10)) then
      Exit('CR without LF');
end;

procedure TRepoHygieneTests.Sources_HaveBomAndCrLfOnly;
var
  Bad: TStringList;
begin
  var List := Files(['*.pas', '*.inc', '*.dpr', '*.dpk']);
  Assert.IsTrue(Length(List) > 50, 'the repository sources were not found (' + RepoRoot + ')');
  Bad := TStringList.Create;
  try
    for var F in List do
    begin
      var B := TFile.ReadAllBytes(F);
      if (Length(B) < 3) or (B[0] <> $EF) or (B[1] <> $BB) or (B[2] <> $BF) then
        Bad.Add(ExtractRelativePath(RepoRoot + '\', F) + ': no UTF-8 BOM')
      else
      begin
        var P := LineEndingProblem(B);
        if P <> '' then Bad.Add(ExtractRelativePath(RepoRoot + '\', F) + ': ' + P);
      end;
    end;
    Assert.AreEqual(0, Bad.Count, 'hygiene: ' + Bad.CommaText);
  finally
    Bad.Free;
  end;
end;

procedure TRepoHygieneTests.TextForms_UseCrLfOnly;
var
  Bad: TStringList;
begin
  Bad := TStringList.Create;
  try
    for var F in Files(['*.dfm', '*.fmx']) do
    begin
      var B := TFile.ReadAllBytes(F);
      if (Length(B) > 0) and (B[0] = $FF) then Continue;   // a binary form
      var P := LineEndingProblem(B);
      if P <> '' then Bad.Add(ExtractRelativePath(RepoRoot + '\', F) + ': ' + P);
    end;
    Assert.AreEqual(0, Bad.Count, 'hygiene: ' + Bad.CommaText);
  finally
    Bad.Free;
  end;
end;

procedure TRepoHygieneTests.GitAttributes_PinDelphiLineEndings;
begin
  var F := TPath.Combine(RepoRoot, '.gitattributes');
  Assert.IsTrue(TFile.Exists(F), '.gitattributes is missing');
  var Rules := TFile.ReadAllText(F);
  for var Ext in ['*.pas', '*.inc', '*.dpr', '*.dpk', '*.dproj'] do
    Assert.IsTrue(Pos(Ext, Rules) > 0, Ext + ' has no rule in .gitattributes');
  Assert.IsTrue(Pos('eol=crlf', Rules) > 0, '.gitattributes does not pin CRLF');
  Assert.IsTrue(Pos('*.res        binary', Rules) > 0, '.res files must be binary');
end;

const
  // every batch script of the repository that uses labels
  ScriptFiles: array[0..4] of string = ('install.cmd', 'uninstall.cmd',
    'rebuild.cmd', 'Mcp\buildmcp.cmd', 'dih\builddih.cmd');

// A label line is ':name' - ':: text' is a comment, which is why the second
// character decides.
function IsLabelLine(const ALine: string): Boolean;
begin
  var T := Trim(ALine);
  Result := (Length(T) > 1) and (T[1] = ':') and (T[2] <> ':');
end;

// the last line of a block that cmd really executes ('' when there is none)
function LastExecutedLine(ALines: TStrings; AFrom, ATo: Integer): string;
begin
  Result := '';
  for var I := AFrom to ATo do
  begin
    var T := Trim(ALines[I]);
    if (T = '') or T.StartsWith('::') or T.StartsWith('rem ', True) then Continue;
    Result := T;
  end;
end;

procedure TRepoHygieneTests.TheBridgeIsMovedAsideUnderAUniqueName;
var
  Lines: TArray<string>;
  MoveLine: string;
begin
  var F := TPath.Combine(TPath.Combine(RepoRoot, 'Mcp'), 'buildmcp.cmd');
  Assert.IsTrue(TFile.Exists(F), 'buildmcp.cmd must be there: ' + F);
  Lines := TFile.ReadAllText(F).Replace(#13#10, #10).Split([#10]);
  for var L in Lines do
    if L.TrimLeft.StartsWith('if exist "%TARGET%" move', True) or
       L.TrimLeft.StartsWith('move /y "%TARGET%"', True) then
      MoveLine := L.Trim;
  Assert.AreNotEqual('', MoveLine,
    'the live exe must still be moved aside before the copy');
  Assert.IsTrue(MoveLine.Contains('%RANDOM%'),
    'and under a UNIQUE name, or an undeletable .old blocks the whole ' +
    'install: ' + MoveLine);
  Assert.IsFalse(MoveLine.Contains('"%TARGET%.old"'),
    'the fixed name is exactly what broke it: ' + MoveLine);
end;

procedure TRepoHygieneTests.InstallScripts_HaveNoLabelFallThrough;
var
  Checked: Integer;
  Bad: string;
begin
  var Root := RepoRoot;
  Assert.IsTrue(Root <> '', 'the repository must be reachable from ' + ParamStr(0));
  Checked := 0;
  Bad := '';
  var SL := TStringList.Create;
  try
    for var Rel in ScriptFiles do
    begin
      var F := TPath.Combine(Root, Rel);
      Assert.IsTrue(TFile.Exists(F), 'the script is missing: ' + Rel);
      SL.Text := TEncoding.ANSI.GetString(TFile.ReadAllBytes(F));
      var Labels: TArray<Integer> := nil;
      for var I := 0 to SL.Count - 1 do
        if IsLabelLine(SL[I]) then Labels := Labels + [I];
      // The LAST block has nothing to fall into, so only the others count.
      for var K := 0 to High(Labels) - 1 do
      begin
        Inc(Checked);
        var Last := LowerCase(LastExecutedLine(SL, Labels[K] + 1, Labels[K + 1] - 1));
        if Last.StartsWith('goto ') or Last.StartsWith('exit') then Continue;
        if Bad = '' then
          Bad := Format('%s %s ends with "%s" and falls into %s',
            [Rel, Trim(SL[Labels[K]]), Last, Trim(SL[Labels[K + 1]])]);
      end;
    end;
  finally
    SL.Free;
  end;
  Assert.IsTrue(Checked >= 10,
    Format('the check must really see label blocks (%d)', [Checked]));
  Assert.AreEqual('', Bad, Bad);
end;

procedure TRepoHygieneTests.NoRoutineReadsTheContextBeforeTheSelection;
var
  Offenders: TStringList;
  Checked: Integer;
begin
  var Root := RepoRoot;
  Assert.IsTrue(Root <> '', 'the repository root must be found');
  var Dir := TPath.Combine(Root, 'Source');
  Assert.IsTrue(TDirectory.Exists(Dir), 'Source must be there: ' + Dir);

  Offenders := TStringList.Create;
  try
    Checked := 0;
    for var F in TDirectory.GetFiles(Dir, '*.pas') do
    begin
      var Lines := TFile.ReadAllLines(F);
      // A routine starts at column 1 (the style of every unit here); its
      // own nested routines are indented, so they belong to it.
      var Routine := '';
      var SawContext := False;
      for var I := 0 to High(Lines) do
      begin
        // CODE only: the comment above the fixed call names
        // GetCurrentContext to say why it is not used, and a sweep that
        // counts prose finds the fix instead of the defect.
        var L := StripLineComment(Lines[I]);
        var U := UpperCase(L);
        if (L <> '') and (L[1] <> ' ') and
           (U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ') or
            U.StartsWith('CONSTRUCTOR ') or U.StartsWith('DESTRUCTOR ')) then
        begin
          Routine := Trim(Copy(L, Pos(' ', L) + 1, MaxInt));
          SawContext := False;
        end;
        if Pos('GETCURRENTCONTEXT', U) > 0 then SawContext := True;
        if Pos('.GETSELECTION(', U) > 0 then
        begin
          Inc(Checked);
          if SawContext then
            Offenders.Add(Format('%s: %s reads GetCurrentContext before ' +
              'GetSelection (line %d)',
              [ExtractFileName(F), Routine, I + 1]));
        end;
      end;
    end;
    // The sweep is only worth anything if it really looked at the call
    // sites - a renamed method would otherwise make it pass vacuously.
    Assert.IsTrue(Checked >= 4,
      Format('the sweep must find the GetSelection call sites, found %d',
        [Checked]));
    Assert.AreEqual('', Offenders.Text.Trim, Offenders.Text);
  finally
    Offenders.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TRepoHygieneTests);

end.
