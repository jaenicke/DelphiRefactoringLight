(*
 * Written by Ian Branch for his fork gllDelphiRefactoringLight and
 * contributed together with his code audit (GitHub issue #10); adapted
 * to this code base by Sebastian Jänicke.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Regression tests for Expert.UsesEditor.ApplyLinesMinimal - the
///  minimal-line write path the 2026-08-30 upstream merge introduced.
///
///  This is the riskiest single routine the merge brought in: every
///  uses-clause edit and every DFM signature auto-fix now goes through it,
///  and it rewrites the user's source. Its whole point is to touch ONLY the
///  changed lines so the IDE's change bars mark the edit rather than the
///  whole unit - which means "the file ends up correct" is only half the
///  contract. WHICH editor operations it used is the other half, and a
///  regression there is invisible to a content-only assertion: falling back
///  to a whole-file write produces byte-identical output and silently loses
///  the entire feature.
///
///  So TFakeEditor counts operations as well as applying them, and the tests
///  assert on both. It implements the line operations with the same contract
///  as the two real helpers: InsertTextAtLineStart inserts RAW text at the
///  start of a line, so a caller passing "text" + CRLF inserts a whole line.
///  (Standalone.EditorHelper got that wrong until the same merge fixed it -
///  it used Lines.Insert, which treats the argument as one item and renders
///  the embedded CRLF as an extra blank line.)
/// </summary>
unit Test.UsesEditorMinimalWrite;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TUsesEditorMinimalWriteTests = class
  private
    /// <summary>Joins ALines with CRLF and a trailing CRLF, the shape a
    ///  source file actually has.</summary>
    function Src(const ALines: array of string): string;
    /// <summary>Installs a fresh fake editor holding AContent for AFile as
    ///  an OPEN buffer, and returns it.</summary>
    procedure NewEditor(const AFile, AContent: string);
    /// <summary>Runs ApplyLinesMinimal against the fake, with ALines as the
    ///  edited copy of AOriginal.</summary>
    function Apply(const AFile, AOriginal: string;
      const ALines: array of string): Boolean;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    [Test] procedure NoChange_TouchesNothing;
    [Test] procedure OneChangedLine_UsesOneLineReplace;
    [Test] procedure TwoChangedLinesApart_SkipsTheIdenticalLineBetween;
    [Test] procedure PureInsertion_UsesLineInsert_NotAWholeFileWrite;
    [Test] procedure TwoSeparateInsertions_BothLandOnTheirOwnLine;
    [Test] procedure PureDeletion_UsesLineDelete;
    [Test] procedure InsertionAtEndOfFile_FallsBackToWholeFileWrite;
    [Test] procedure AboveTheOperationCap_FallsBackToWholeFileWrite;
    [Test] procedure FileNotOpenInEditor_FallsBackToWholeFileWrite;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  Expert.EditorHelperIntf, Expert.UsesEditor;

const
  NL = #13#10;

type
  /// <summary>In-memory IEditorHelper. Only the five members
  ///  ApplyLinesMinimal actually calls carry behaviour; the rest satisfy the
  ///  interface and are never reached (a test that reaches one is a test
  ///  that is not doing what it says).</summary>
  TFakeEditor = class(TInterfacedObject, IEditorHelper)
  private
    FBuffers: TDictionary<string, string>;
    FWholeWrites: Integer;
    FLineReplaces: Integer;
    FLineDeletes: Integer;
    FLineInserts: Integer;
    function SplitOf(const AFile: string): TStringList;
    procedure StoreFrom(const AFile: string; ASL: TStringList);
  public
    // IEditorHelper members this fake does not need
    function GetOpenSourceFiles: TArray<string>;
    function IsFormInDesigner(const APasFile: string): Boolean;
    function RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
      AIsMethod: Boolean; out AMessage: string): Boolean;
    constructor Create;
    destructor Destroy; override;
    procedure OpenBuffer(const AFile, AContent: string);
    function Content(const AFile: string): string;
    property WholeWrites: Integer read FWholeWrites;
    property LineReplaces: Integer read FLineReplaces;
    property LineDeletes: Integer read FLineDeletes;
    property LineInserts: Integer read FLineInserts;
    function LineOps: Integer;

    // ---- the members under test ----
    function ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;
    function ReplaceFileContent(const AFilePath: string; const ANewContent: string): Boolean;
    function ReplaceLineAt(const AFilePath: string; ALine: Integer;
      const ANewContent: string): Boolean;
    function DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;
    function InsertTextAtLineStart(const AFilePath: string; ALine: Integer;
      const AText: string): Boolean;

    // ---- never reached by ApplyLinesMinimal ----
    function GetCurrentContext: TEditorContext;
    function GetActiveFileName: string;
    function GetCaretLineCol(out ALine, ACol: Integer): Boolean;
    function GetCurrentProjectDproj: string;
    function GetProjectRoot: string;
    function GetProjectSearchPaths: string;
    function AddProjectSearchPath(const ADir: string): Boolean;
    function GetProjectSourceFiles: TArray<string>;
    function BuildSearchPathFromProject(const ADprojPath, ARootPath: string): string;
    function FindDelphiLspJson: string;
    function ReplaceSelection(const AFilePath: string;
      AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      const ANewText: string): Boolean;
    function ApplyEditViaEditor(const AFilePath: string;
      ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;
    procedure SaveAllFiles;
    function SaveFile(const AFilePath: string): Boolean;
    procedure ReloadModifiedFiles(const FilePaths: TArray<string>);
    procedure NotifyClassStructureChanged(const AFilePath: string);
    function GotoLocation(const AFilePath: string;
      ALine, ACol: Integer; AHighlightLen: Integer = 0): Boolean;
    function AddFileToActiveProject(const AFilePath: string): Boolean;
    function GetSelection(out AFilePath: string;
      out AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      out AText: string): Boolean;
  end;

var
  GFake: TFakeEditor;   // weak view of the installed helper, for assertions

{ TFakeEditor }

constructor TFakeEditor.Create;
begin
  inherited Create;
  FBuffers := TDictionary<string, string>.Create;
end;

destructor TFakeEditor.Destroy;
begin
  FBuffers.Free;
  inherited;
end;

function TFakeEditor.GetOpenSourceFiles: TArray<string>;
begin
  Result := nil;
end;

function TFakeEditor.IsFormInDesigner(const APasFile: string): Boolean;
begin
  Result := False;
end;

function TFakeEditor.RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
  AIsMethod: Boolean; out AMessage: string): Boolean;
begin
  AMessage := '';
  Result := False;
end;

procedure TFakeEditor.OpenBuffer(const AFile, AContent: string);
begin
  FBuffers.AddOrSetValue(AFile, AContent);
end;

function TFakeEditor.Content(const AFile: string): string;
begin
  if not FBuffers.TryGetValue(AFile, Result) then Result := '';
end;

function TFakeEditor.LineOps: Integer;
begin
  Result := FLineReplaces + FLineDeletes + FLineInserts;
end;

function TFakeEditor.SplitOf(const AFile: string): TStringList;
begin
  Result := TStringList.Create;
  Result.Text := Content(AFile);
end;

procedure TFakeEditor.StoreFrom(const AFile: string; ASL: TStringList);
begin
  FBuffers.AddOrSetValue(AFile, ASL.Text);
end;

function TFakeEditor.ReadEditorContent(const AFilePath: string;
  out AContent: string): Boolean;
begin
  Result := FBuffers.TryGetValue(AFilePath, AContent);
  if not Result then AContent := '';
end;

function TFakeEditor.ReplaceFileContent(const AFilePath: string;
  const ANewContent: string): Boolean;
begin
  Inc(FWholeWrites);
  FBuffers.AddOrSetValue(AFilePath, ANewContent);
  Result := True;
end;

function TFakeEditor.ReplaceLineAt(const AFilePath: string; ALine: Integer;
  const ANewContent: string): Boolean;
var
  SL: TStringList;
begin
  SL := SplitOf(AFilePath);
  try
    Result := (ALine >= 1) and (ALine <= SL.Count);
    if not Result then Exit;
    SL[ALine - 1] := ANewContent;
    StoreFrom(AFilePath, SL);
    Inc(FLineReplaces);
  finally
    SL.Free;
  end;
end;

function TFakeEditor.DeleteLineAt(const AFilePath: string;
  ALine: Integer): Boolean;
var
  SL: TStringList;
begin
  SL := SplitOf(AFilePath);
  try
    Result := (ALine >= 1) and (ALine <= SL.Count);
    if not Result then Exit;
    SL.Delete(ALine - 1);
    StoreFrom(AFilePath, SL);
    Inc(FLineDeletes);
  finally
    SL.Free;
  end;
end;

function TFakeEditor.InsertTextAtLineStart(const AFilePath: string;
  ALine: Integer; const AText: string): Boolean;
var
  SL: TStringList;
begin
  SL := SplitOf(AFilePath);
  try
    Result := (ALine >= 1) and (ALine <= SL.Count + 1);
    if not Result then Exit;
    // RAW text at the line's start - NOT SL.Insert. A caller passing
    // "unit," + CRLF inserts a whole line; SL.Insert would keep the CRLF
    // inside one item and produce a blank line on the next round-trip.
    if ALine <= SL.Count then
      SL[ALine - 1] := AText + SL[ALine - 1]
    else
      SL.Add(AText.TrimRight([#13, #10]));
    StoreFrom(AFilePath, SL);
    Inc(FLineInserts);
  finally
    SL.Free;
  end;
end;

function TFakeEditor.GetCurrentContext: TEditorContext;
begin
  Result := Default(TEditorContext);
end;

function TFakeEditor.GetActiveFileName: string;
begin
  Result := '';
end;

function TFakeEditor.GetCaretLineCol(out ALine, ACol: Integer): Boolean;
begin
  ALine := 0; ACol := 0; Result := False;
end;

function TFakeEditor.GetCurrentProjectDproj: string;
begin
  Result := '';
end;

function TFakeEditor.GetProjectRoot: string;
begin
  Result := '';
end;

function TFakeEditor.GetProjectSearchPaths: string;
begin
  Result := '';
end;

function TFakeEditor.AddProjectSearchPath(const ADir: string): Boolean;
begin
  Result := False;
end;

function TFakeEditor.GetProjectSourceFiles: TArray<string>;
begin
  Result := nil;
end;

function TFakeEditor.BuildSearchPathFromProject(
  const ADprojPath, ARootPath: string): string;
begin
  Result := '';
end;

function TFakeEditor.FindDelphiLspJson: string;
begin
  Result := '';
end;

function TFakeEditor.ReplaceSelection(const AFilePath: string;
  AStartLine, AStartCol, AEndLine, AEndCol: Integer;
  const ANewText: string): Boolean;
begin
  Result := False;
end;

function TFakeEditor.ApplyEditViaEditor(const AFilePath: string;
  ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;
begin
  Result := False;
end;

procedure TFakeEditor.SaveAllFiles;
begin
end;

function TFakeEditor.SaveFile(const AFilePath: string): Boolean;
begin
  Result := True;
end;

procedure TFakeEditor.ReloadModifiedFiles(const FilePaths: TArray<string>);
begin
end;

procedure TFakeEditor.NotifyClassStructureChanged(const AFilePath: string);
begin
end;

function TFakeEditor.GotoLocation(const AFilePath: string;
  ALine, ACol: Integer; AHighlightLen: Integer = 0): Boolean;
begin
  Result := False;
end;

function TFakeEditor.AddFileToActiveProject(const AFilePath: string): Boolean;
begin
  Result := False;
end;

function TFakeEditor.GetSelection(out AFilePath: string;
  out AStartLine, AStartCol, AEndLine, AEndCol: Integer;
  out AText: string): Boolean;
begin
  AFilePath := ''; AStartLine := 0; AStartCol := 0;
  AEndLine := 0; AEndCol := 0; AText := '';
  Result := False;
end;

{ TUsesEditorMinimalWriteTests }

procedure TUsesEditorMinimalWriteTests.Setup;
begin
  GFake := nil;
end;

procedure TUsesEditorMinimalWriteTests.TearDown;
begin
  SetEditorImpl(nil);
  GFake := nil;
end;

function TUsesEditorMinimalWriteTests.Src(const ALines: array of string): string;
var
  L: string;
begin
  Result := '';
  for L in ALines do
    Result := Result + L + NL;
end;

procedure TUsesEditorMinimalWriteTests.NewEditor(const AFile, AContent: string);
var
  Fake: TFakeEditor;
begin
  Fake := TFakeEditor.Create;
  Fake.OpenBuffer(AFile, AContent);
  GFake := Fake;              // the interface reference below owns it
  SetEditorImpl(Fake);
end;

function TUsesEditorMinimalWriteTests.Apply(const AFile, AOriginal: string;
  const ALines: array of string): Boolean;
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    SL.Text := Src(ALines);
    Result := ApplyLinesMinimal(AFile, SL, AOriginal);
  finally
    SL.Free;
  end;
end;

procedure TUsesEditorMinimalWriteTests.NoChange_TouchesNothing;
var
  Orig: string;
begin
  Orig := Src(['unit Foo;', '', 'interface', '', 'implementation', '', 'end.']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig,
    ['unit Foo;', '', 'interface', '', 'implementation', '', 'end.']),
    'an unchanged apply still reports success');
  Assert.AreEqual<Integer>(0, GFake.WholeWrites, 'no whole-file write');
  Assert.AreEqual<Integer>(0, GFake.LineOps, 'no line operations either');
end;

procedure TUsesEditorMinimalWriteTests.OneChangedLine_UsesOneLineReplace;
var
  Orig: string;
begin
  Orig := Src(['unit Foo;', '', 'uses', '  System.SysUtils;', '', 'implementation']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig,
    ['unit Foo;', '', 'uses', '  System.SysUtils, System.Classes;', '',
     'implementation']), 'apply succeeds');
  Assert.AreEqual<Integer>(0, GFake.WholeWrites,
    'a one-line edit must NOT rewrite the whole file - that is the feature');
  Assert.AreEqual<Integer>(1, GFake.LineReplaces, 'exactly one line replaced');
  Assert.AreEqual<Integer>(0, GFake.LineInserts + GFake.LineDeletes,
    'and nothing inserted or deleted');
  Assert.AreEqual(
    Src(['unit Foo;', '', 'uses', '  System.SysUtils, System.Classes;', '',
         'implementation']),
    GFake.Content('Foo.pas'), 'resulting content');
end;

procedure TUsesEditorMinimalWriteTests.TwoChangedLinesApart_SkipsTheIdenticalLineBetween;
var
  Orig: string;
begin
  // The interface->implementation move touches BOTH clauses, so the diff
  // window spans two edits with untouched lines between them. Those must
  // not be rewritten - each rewrite is another change bar.
  Orig := Src(['uses', '  A;', '', 'implementation', '', 'uses', '  B;']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig,
    ['uses', '  A, X;', '', 'implementation', '', 'uses', '  B, Y;']),
    'apply succeeds');
  Assert.AreEqual<Integer>(0, GFake.WholeWrites, 'no whole-file write');
  Assert.AreEqual<Integer>(2, GFake.LineReplaces,
    'exactly the two changed lines - the identical lines between them are skipped');
end;

procedure TUsesEditorMinimalWriteTests.PureInsertion_UsesLineInsert_NotAWholeFileWrite;
var
  Orig: string;
begin
  // One-unit-per-line clause: adding a unit is a pure insertion, and it is
  // the single most common edit this routine performs.
  Orig := Src(['uses', '  System.SysUtils,', '  System.Classes;', '',
               'implementation']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig,
    ['uses', '  System.SysUtils,', '  System.Math,', '  System.Classes;', '',
     'implementation']), 'apply succeeds');
  Assert.AreEqual<Integer>(0, GFake.WholeWrites, 'no whole-file write');
  Assert.AreEqual<Integer>(1, GFake.LineInserts, 'one line inserted');
  Assert.AreEqual<Integer>(0, GFake.LineReplaces + GFake.LineDeletes,
    'and nothing replaced or deleted - the lines below merely shifted');
  Assert.AreEqual(
    Src(['uses', '  System.SysUtils,', '  System.Math,', '  System.Classes;',
         '', 'implementation']),
    GFake.Content('Foo.pas'), 'the new line lands in the right place');
end;

procedure TUsesEditorMinimalWriteTests.TwoSeparateInsertions_BothLandOnTheirOwnLine;
var
  Orig: string;
begin
  // Applied bottom-up, so the upper insertion's original line number is
  // still valid when it runs. Top-down would put the second one one line
  // too low - and the content assertion is what catches that.
  Orig := Src(['A;', 'B;', 'C;', 'D;']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig, ['A;', 'NEW1;', 'B;', 'C;', 'NEW2;', 'D;']),
    'apply succeeds');
  Assert.AreEqual<Integer>(0, GFake.WholeWrites, 'no whole-file write');
  Assert.AreEqual<Integer>(2, GFake.LineInserts, 'two line inserts');
  Assert.AreEqual(Src(['A;', 'NEW1;', 'B;', 'C;', 'NEW2;', 'D;']),
    GFake.Content('Foo.pas'), 'both insertions land on their own line');
end;

procedure TUsesEditorMinimalWriteTests.PureDeletion_UsesLineDelete;
var
  Orig: string;
begin
  // Removing a unit from a one-per-line clause (uses cleanup).
  Orig := Src(['uses', '  A,', '  B,', '  C;', '', 'implementation']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig,
    ['uses', '  A,', '  C;', '', 'implementation']), 'apply succeeds');
  Assert.AreEqual<Integer>(0, GFake.WholeWrites, 'no whole-file write');
  Assert.AreEqual<Integer>(1, GFake.LineDeletes, 'one line deleted');
  Assert.AreEqual(Src(['uses', '  A,', '  C;', '', 'implementation']),
    GFake.Content('Foo.pas'), 'the right line went');
end;

procedure TUsesEditorMinimalWriteTests.InsertionAtEndOfFile_FallsBackToWholeFileWrite;
var
  Orig: string;
begin
  // No common suffix (S = 0), so the routine cannot anchor the insertion
  // against a following line and deliberately takes the single write.
  Orig := Src(['A;', 'B;']);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig, ['A;', 'B;', 'C;']), 'apply succeeds');
  Assert.AreEqual<Integer>(1, GFake.WholeWrites,
    'an append with no common suffix falls back to one whole-file write');
  Assert.AreEqual<Integer>(0, GFake.LineOps, 'and uses no line operations');
  Assert.AreEqual(Src(['A;', 'B;', 'C;']), GFake.Content('Foo.pas'),
    'the fallback still produces the right content');
end;

procedure TUsesEditorMinimalWriteTests.AboveTheOperationCap_FallsBackToWholeFileWrite;
var
  Before, After: TArray<string>;
  I: Integer;
  Orig: string;
begin
  // 20 changed lines > the 16-operation cap: one undo step beats twenty.
  SetLength(Before, 22);
  SetLength(After, 22);
  Before[0] := 'HEAD;'; After[0] := 'HEAD;';
  for I := 1 to 20 do
  begin
    Before[I] := Format('  Old%d;', [I]);
    After[I] := Format('  New%d;', [I]);
  end;
  Before[21] := 'TAIL;'; After[21] := 'TAIL;';
  Orig := Src(Before);
  NewEditor('Foo.pas', Orig);
  Assert.IsTrue(Apply('Foo.pas', Orig, After), 'apply succeeds');
  Assert.AreEqual<Integer>(1, GFake.WholeWrites,
    'above the cap it takes a single whole-file write');
  Assert.AreEqual<Integer>(0, GFake.LineOps, 'and no line operations at all');
  Assert.AreEqual(Src(After), GFake.Content('Foo.pas'), 'content still correct');
end;

procedure TUsesEditorMinimalWriteTests.FileNotOpenInEditor_FallsBackToWholeFileWrite;
var
  Orig: string;
begin
  // Closed files are written straight to disk: line operations would have to
  // OpenModule first and would flood the IDE with tabs on a batch fix, and
  // change bars mean nothing for a file nobody is looking at.
  Orig := Src(['uses', '  A;', '', 'implementation']);
  NewEditor('Open.pas', Orig);            // a DIFFERENT file is open
  Assert.IsTrue(Apply('Closed.pas', Orig, ['uses', '  A, B;', '', 'implementation']),
    'apply succeeds for a file with no editor buffer');
  Assert.AreEqual<Integer>(1, GFake.WholeWrites, 'one whole-file write');
  Assert.AreEqual<Integer>(0, GFake.LineOps, 'no line operations');
  Assert.AreEqual(Src(['uses', '  A, B;', '', 'implementation']),
    GFake.Content('Closed.pas'), 'and the closed file received the new content');
end;

initialization
  TDUnitX.RegisterTestFixture(TUsesEditorMinimalWriteTests);

end.
