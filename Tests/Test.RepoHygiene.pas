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
  end;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Types;

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

initialization
  TDUnitX.RegisterTestFixture(TRepoHygieneTests);

end.
