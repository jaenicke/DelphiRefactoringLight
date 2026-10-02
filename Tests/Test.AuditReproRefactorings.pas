(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Audit repro tests, refactorings area (issue #22): move to unit, extract
///  interface, delegate IInterface, extract method and change signature.
///  One test per open item, named after the item id. RED means the item is
///  still open; a test can be deleted once it goes green. They use only the
///  public API, with an in-memory IEditorHelper standing in for the IDE.
/// </summary>
unit Test.AuditReproRefactorings;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TAuditReproRefactoringsTests = class
  private
    FDir: string;
    function WriteUnit(const AName: string; const ALines: array of string): string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;
    // move to unit
    [Test] procedure M39a_ForwardDeclarationIsRemovedWithTheType;
    [Test] procedure M39b_OverloadedRoutineIsNotSplit;
    [Test] procedure M39c_RefusedSourceWriteIsNoSuccess;
    [Test] procedure M39d_PreviewSavesNothing;
    [Test] procedure L7j_DirectiveAboveTheBodyIsKept;
    [Test] procedure L7k_MultiLineStringKeepsItsBlankLines;
    [Test] procedure L7k_FinalLineBreakIsKept;
    [Test] procedure L7l_FailedMoveToNewUnitLeavesNoUnit;
    // extract interface / delegate IInterface
    [Test] procedure M38a_NestedClassAndCommentEndDoNotEndTheClass;
    [Test] procedure M38b_DefaultSelectionSkipsWhatAnInterfaceCannotHold;
    [Test] procedure M38c_ClassMethodIsNotEmittedAsInstanceMethod;
    [Test] procedure M38d_InjectionGoesBeforeTheMethodsOwnEnd;
    [Test] procedure M38d_InjectionFindsTheExactMethod;
    [Test] procedure M38e_ClassAbstractGetsACompilableAncestorList;
    [Test] procedure M38e_SubstringOfAnotherInterfaceIsNoDuplicate;
    [Test] procedure M38f_LfOnlyBufferIsReadAsLines;
    [Test] procedure L7g_UsesSeedSkipsComments;
    [Test] procedure L7h_FailedExtractLeavesNoUnitInTheProject;
    // extract method
    [Test] procedure M37b_PreviewSavesNothing;
    // change signature
    [Test] procedure L7m_GenericCallWithoutArguments;
    [Test] procedure L7n_NestedCallOfTheSameRoutineIsRefusedClearly;
  end;

implementation

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Types, System.IOUtils,
  System.Generics.Collections,
  Expert.EditorHelperIntf, Expert.MoveToUnit, Expert.ExtractInterface,
  Expert.ExtractInterfaceWizard, Expert.ExtractMethod, Expert.SignatureEdit;

const
  NL = #13#10;

type
  /// <summary>In-memory IEditorHelper: a buffer per written file (reads fall
  ///  through to the disk until a file is written), a list of files added to
  ///  the project, a SaveAllFiles counter, and one file whose writes are
  ///  refused. No DelphiLSP: FindDelphiLspJson is '' unless a test sets it.</summary>
  TAuditFakeEditor = class(TInterfacedObject, IEditorHelper)
  private
    FBuffers: TDictionary<string, string>;
  public
    Saves: Integer;
    Added: TArray<string>;
    RefuseWritesTo: string;
    LspJson: string;
    RaiseOnProject: Boolean;
    constructor Create;
    destructor Destroy; override;
    procedure OpenBuffer(const AFile, AContent: string);
    function Content(const AFile: string): string;
    function GetCurrentContext: TEditorContext;
    function GetActiveFileName: string;
    function GetCaretLineCol(out ALine, ACol: Integer): Boolean;
    function RawColumn(const AFile: string; ALine, ADisplayCol: Integer): Integer;
    function GetCurrentProjectDproj: string;
    function GetProjectRoot: string;
    function GetProjectSearchPaths: string;
    function GetProjectSourceFiles: TArray<string>;
    function GetOpenSourceFiles: TArray<string>;
    function BuildSearchPathFromProject(const ADprojPath, ARootPath: string): string;
    function FindDelphiLspJson: string;
    function ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;
    function ReplaceFileContent(const AFilePath: string; const ANewContent: string): Boolean;
    function ReplaceSelection(const AFilePath: string;
      AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      const ANewText: string): Boolean;
    function ReplaceLineAt(const AFilePath: string; ALine: Integer;
      const ANewContent: string): Boolean;
    function DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;
    function InsertTextAtLineStart(const AFilePath: string; ALine: Integer;
      const AText: string): Boolean;
    function ApplyEditViaEditor(const AFilePath: string;
      ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;
    procedure SaveAllFiles;
    function SaveFile(const AFilePath: string): Boolean;
    procedure ReloadModifiedFiles(const FilePaths: TArray<string>);
    procedure NotifyClassStructureChanged(const AFilePath: string);
    function IsFormInDesigner(const APasFile: string): Boolean;
    function GetDesignerRequiredUnits(const APasFile: string;
      out AUnits: TArray<TDesignerRequiredUnit>;
      out AComplete: Boolean): Boolean;
    function RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
      AIsMethod: Boolean; out AMessage: string): Boolean;
    function GotoLocation(const AFilePath: string;
      ALine, ACol: Integer; AHighlightLen: Integer = 0): Boolean;
    function AddFileToActiveProject(const AFilePath: string): Boolean;
    function AddProjectSearchPath(const ADir: string): Boolean;
    function GetSelection(out AFilePath: string;
      out AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      out AText: string): Boolean;
  end;

var
  GFake: TAuditFakeEditor;   // weak view of the installed helper

{ TAuditFakeEditor }

constructor TAuditFakeEditor.Create;
begin
  inherited Create;
  FBuffers := TDictionary<string, string>.Create;
end;

destructor TAuditFakeEditor.Destroy;
begin
  FBuffers.Free;
  inherited;
end;

procedure TAuditFakeEditor.OpenBuffer(const AFile, AContent: string);
begin
  FBuffers.AddOrSetValue(UpperCase(AFile), AContent);
end;

function TAuditFakeEditor.Content(const AFile: string): string;
begin
  if not FBuffers.TryGetValue(UpperCase(AFile), Result) then
    Result := '';
end;

function TAuditFakeEditor.GetCurrentContext: TEditorContext;
begin
  Result := Default(TEditorContext);
end;

function TAuditFakeEditor.GetActiveFileName: string;
begin
  Result := '';
end;

function TAuditFakeEditor.RawColumn(const AFile: string; ALine,
  ADisplayCol: Integer): Integer;
begin
  Result := ADisplayCol;   // no tab expansion in a test stand-in
end;

function TAuditFakeEditor.GetCaretLineCol(out ALine, ACol: Integer): Boolean;
begin
  ALine := 0;
  ACol := 0;
  Result := False;
end;

function TAuditFakeEditor.GetCurrentProjectDproj: string;
begin
  // lets a test stop a refactoring right before it would start DelphiLSP
  if RaiseOnProject then
    raise EAbort.Create('test: no project');
  Result := '';
end;

function TAuditFakeEditor.GetProjectRoot: string;
begin
  Result := '';
end;

function TAuditFakeEditor.GetProjectSearchPaths: string;
begin
  Result := '';
end;

function TAuditFakeEditor.GetProjectSourceFiles: TArray<string>;
begin
  Result := nil;
end;

function TAuditFakeEditor.GetOpenSourceFiles: TArray<string>;
begin
  Result := nil;
end;

function TAuditFakeEditor.BuildSearchPathFromProject(const ADprojPath, ARootPath: string): string;
begin
  Result := '';
end;

function TAuditFakeEditor.FindDelphiLspJson: string;
begin
  Result := LspJson;
end;

function TAuditFakeEditor.ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;
begin
  Result := FBuffers.TryGetValue(UpperCase(AFilePath), AContent);
end;

function TAuditFakeEditor.ReplaceFileContent(const AFilePath: string; const ANewContent: string): Boolean;
begin
  if (RefuseWritesTo <> '') and SameText(AFilePath, RefuseWritesTo) then
    Exit(False);
  FBuffers.AddOrSetValue(UpperCase(AFilePath), ANewContent);
  Result := True;
end;

function TAuditFakeEditor.ReplaceSelection(const AFilePath: string;
  AStartLine, AStartCol, AEndLine, AEndCol: Integer; const ANewText: string): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.ReplaceLineAt(const AFilePath: string; ALine: Integer;
  const ANewContent: string): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.InsertTextAtLineStart(const AFilePath: string; ALine: Integer;
  const AText: string): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.ApplyEditViaEditor(const AFilePath: string;
  ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;
begin
  Result := False;
end;

procedure TAuditFakeEditor.SaveAllFiles;
begin
  Inc(Saves);
end;

function TAuditFakeEditor.SaveFile(const AFilePath: string): Boolean;
begin
  Result := True;
end;

procedure TAuditFakeEditor.ReloadModifiedFiles(const FilePaths: TArray<string>);
begin
end;

procedure TAuditFakeEditor.NotifyClassStructureChanged(const AFilePath: string);
begin
end;

function TAuditFakeEditor.IsFormInDesigner(const APasFile: string): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.GetDesignerRequiredUnits(const APasFile: string;
  out AUnits: TArray<TDesignerRequiredUnit>; out AComplete: Boolean): Boolean;
begin
  AUnits := nil;
  AComplete := True;
  Result := False;
end;

function TAuditFakeEditor.RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
  AIsMethod: Boolean; out AMessage: string): Boolean;
begin
  AMessage := '';
  Result := False;
end;

function TAuditFakeEditor.GotoLocation(const AFilePath: string;
  ALine, ACol: Integer; AHighlightLen: Integer): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.AddFileToActiveProject(const AFilePath: string): Boolean;
begin
  Added := Added + [AFilePath];
  Result := True;
end;

function TAuditFakeEditor.AddProjectSearchPath(const ADir: string): Boolean;
begin
  Result := False;
end;

function TAuditFakeEditor.GetSelection(out AFilePath: string;
  out AStartLine, AStartCol, AEndLine, AEndCol: Integer; out AText: string): Boolean;
begin
  AFilePath := '';
  AStartLine := 0;
  AStartCol := 0;
  AEndLine := 0;
  AEndCol := 0;
  AText := '';
  Result := False;
end;

{ helpers }

function Join(const ALines: array of string): string;
begin
  Result := '';
  for var L in ALines do
    Result := Result + L + NL;
end;

/// <summary>0-based index of the first line whose trimmed text is ATrimmed
///  at or after AFrom, -1 = none.</summary>
function IndexOfLine(const ALines: TArray<string>; const ATrimmed: string; AFrom: Integer = 0): Integer;
begin
  for var I := AFrom to High(ALines) do
    if Trim(ALines[I]) = ATrimmed then Exit(I);
  Result := -1;
end;

function LinesOf(const AText: string): TArray<string>;
begin
  Result := AText.Replace(#13#10, #10).Split([#10]);
end;

function CountOf(const ASub, AText: string): Integer;
begin
  Result := 0;
  var P := Pos(ASub, AText);
  while P > 0 do
  begin
    Inc(Result);
    P := Pos(ASub, AText, P + Length(ASub));
  end;
end;

{ TAuditReproRefactoringsTests }

procedure TAuditReproRefactoringsTests.Setup;
var
  Fake: TAuditFakeEditor;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'AuditRepro' + TGUID.NewGuid.ToString.Substring(1, 8));
  TDirectory.CreateDirectory(FDir);
  Fake := TAuditFakeEditor.Create;
  GFake := Fake;              // the interface reference below owns it
  SetEditorImpl(Fake);
end;

procedure TAuditReproRefactoringsTests.TearDown;
begin
  SetEditorImpl(nil);
  GFake := nil;
  for var F in TDirectory.GetFiles(FDir) do
    SetFileAttributes(PChar(F), FILE_ATTRIBUTE_NORMAL);
  TDirectory.Delete(FDir, True);
end;

function TAuditReproRefactoringsTests.WriteUnit(const AName: string;
  const ALines: array of string): string;
begin
  Result := TPath.Combine(FDir, AName + '.pas');
  TFile.WriteAllText(Result, Join(ALines));
end;

{ move to unit }

procedure TAuditReproRefactoringsTests.M39a_ForwardDeclarationIsRemovedWithTheType;
var
  Plan: TMovePlan;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class;',
    '  TFoo = class', '  end;', 'implementation', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  Assert.IsTrue(TLspMoveToUnit.BuildPlan('TFoo', Src, Tgt, Plan), Plan.ProblemDetail);
  if TLspMoveToUnit.ApplyPlan(Plan) then
    Assert.IsFalse(GFake.Content(Src).Contains('TFoo = class;'),
      'M39a: the forward declaration stays in the source without its type (E2086)');
end;

procedure TAuditReproRefactoringsTests.M39b_OverloadedRoutineIsNotSplit;
var
  Plan: TMovePlan;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface',
    'procedure Log(const S: string); overload;', 'procedure Log(N: Integer); overload;',
    'implementation', 'procedure Log(N: Integer);', 'begin', 'end;',
    'procedure Log(const S: string);', 'begin', 'end;', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  // refusing is fine; moving must take the whole family. (Adapted while
  // fixing the item: the refusal branch has to ASSERT something, or DUnitX
  // fails the test for making no assertion at all.)
  if TLspMoveToUnit.BuildPlan('Log', Src, Tgt, Plan) then
  begin
    if TLspMoveToUnit.ApplyPlan(Plan) then
      Assert.IsFalse(GFake.Content(Src).Contains('procedure Log'),
        'M39b: one overload moved, the other stays behind - and the moved ' +
        'declaration and body belong to different overloads');
  end
  else
    Assert.Contains(LowerCase(Plan.ProblemDetail), 'overload',
      'M39b: refusing is fine, but the reason must name the overloads: ' +
      Plan.ProblemDetail);
end;

procedure TAuditReproRefactoringsTests.M39c_RefusedSourceWriteIsNoSuccess;
var
  Plan: TMovePlan;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'const', '  CFoo = 1;',
    'implementation', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  Assert.IsTrue(TLspMoveToUnit.BuildPlan('CFoo', Src, Tgt, Plan), Plan.ProblemDetail);
  GFake.RefuseWritesTo := Src;
  Assert.IsFalse(TLspMoveToUnit.ApplyPlan(Plan),
    'M39c: the source write was refused (CFoo is now in both units) and ApplyPlan reports success');
end;

procedure TAuditReproRefactoringsTests.M39d_PreviewSavesNothing;
var
  Plan: TMovePlan;
  Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'const', '  CFoo = 1;',
    'implementation', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  Assert.IsTrue(TLspMoveToUnit.ExecuteToExistingUnit('CFoo', Src, Tgt, True, Plan, Err), Err);
  Assert.AreEqual(0, GFake.Saves,
    'M39d: a move_to_unit PREVIEW saved every modified file (SaveAllFiles)');
end;

procedure TAuditReproRefactoringsTests.L7j_DirectiveAboveTheBodyIsKept;
var
  Plan: TMovePlan;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class',
    '    procedure Run;', '  end;', 'implementation', '{$IFDEF USE_TFoo}',
    'procedure TFoo.Run;', 'begin', 'end;', '{$ENDIF}', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  Assert.IsTrue(TLspMoveToUnit.BuildPlan('TFoo', Src, Tgt, Plan), Plan.ProblemDetail);
  Assert.IsTrue(TLspMoveToUnit.ApplyPlan(Plan), 'the move itself');
  var Left := GFake.Content(Src);
  Assert.AreEqual(CountOf('{$ENDIF}', Left), CountOf('{$IFDEF', Left),
    'L7j: {$IFDEF USE_TFoo} was taken for a "{ TFoo }" marker and deleted, its {$ENDIF} ' +
    'stays - the source is unbalanced:' + NL + Left);
end;

procedure TAuditReproRefactoringsTests.L7k_MultiLineStringKeepsItsBlankLines;
var
  Plan: TMovePlan;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class',
    '  end;', 'const', '  Banner = ''''''', '    a', '', '', '    b', '    '''''';',
    'implementation', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  Assert.IsTrue(TLspMoveToUnit.BuildPlan('TFoo', Src, Tgt, Plan), Plan.ProblemDetail);
  Assert.IsTrue(TLspMoveToUnit.ApplyPlan(Plan), 'the move itself');
  Assert.IsTrue(GFake.Content(Src).Contains('    a' + NL + NL + NL + '    b'),
    'L7k: the blank-line cleanup changed the value of a multi-line string it did not move');
end;

procedure TAuditReproRefactoringsTests.L7k_FinalLineBreakIsKept;
var
  Plan: TMovePlan;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'const', '  CFoo = 1;',
    '  CBar = 2;', 'implementation', 'end.']);
  var Tgt := WriteUnit('UTgt', ['unit UTgt;', 'interface', 'implementation', 'end.']);
  Assert.IsTrue(TLspMoveToUnit.BuildPlan('CFoo', Src, Tgt, Plan), Plan.ProblemDetail);
  Assert.IsTrue(TLspMoveToUnit.ApplyPlan(Plan), 'the move itself');
  Assert.IsTrue(GFake.Content(Src).EndsWith('end.' + NL),
    'L7k: the source unit lost its final line break');
  Assert.IsTrue(GFake.Content(Tgt).EndsWith('end.' + NL),
    'L7k: the target unit lost its final line break');
end;

procedure TAuditReproRefactoringsTests.L7l_FailedMoveToNewUnitLeavesNoUnit;
var
  Plan: TMovePlan;
  Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'const', '  CFoo = 1;',
    'implementation', 'end.']);
  var NewFile := TPath.Combine(FDir, 'UNew.pas');
  GFake.RefuseWritesTo := NewFile;   // the move into it does not land
  var Moved := TLspMoveToUnit.ExecuteToNewUnit('CFoo', Src, NewFile, False, Plan, Err);
  Assert.IsFalse(Moved, 'the refused write must fail the move');
  Assert.IsTrue((not FileExists(NewFile)) and (Length(GFake.Added) = 0),
    'L7l: the move failed (' + Err + ') but the empty UNew.pas stays on disk and in the project');
end;

{ extract interface / delegate IInterface }

procedure TAuditReproRefactoringsTests.M38a_NestedClassAndCommentEndDoNotEndTheClass;

  function HasRun(const ALines: array of string): Boolean;
  var
    Info: TExtractInterfaceInfo;
    Lines: TArray<string>;
  begin
    SetLength(Lines, Length(ALines));
    for var I := 0 to High(ALines) do Lines[I] := ALines[I];
    Assert.IsTrue(TExtractInterfaceEngine.ParseClassAtLine(Lines, 'U.pas', 4, Info), 'no class');
    Result := False;
    for var M in Info.Members do
      if SameText(M.Name, 'Run') then Exit(True);
  end;

begin
  Assert.IsTrue(HasRun(['unit U;', 'interface', 'type', '  TOuter = class', '  type',
    '    TInner = class', '    end;', '  public', '    procedure Run;', '  end;',
    'implementation', 'end.']),
    'M38a: the nested class''s "end" ended TOuter - Run is not offered');
  Assert.IsTrue(HasRun(['unit U;', 'interface', 'type', '  TOuter = class',
    '    { the old end }', '  public', '    procedure Run;', '  end;',
    'implementation', 'end.']),
    'M38a: an "end" inside a { } comment ended TOuter - Run is not offered');
end;

const
  CtorClass: array[0..14] of string = ('unit USrc;', 'interface', 'type',
    '  TFoo = class', '  public', '    constructor Create;', '    destructor Destroy; override;',
    '    class function Make: TFoo;', '    procedure Run;', '  end;', 'implementation',
    'constructor TFoo.Create; begin end;', 'destructor TFoo.Destroy; begin end;',
    'class function TFoo.Make: TFoo; begin Result := nil; end;', 'end.');

procedure TAuditReproRefactoringsTests.M38b_DefaultSelectionSkipsWhatAnInterfaceCannotHold;
var
  Info: TExtractInterfaceInfo;
  Text, Err: string;
begin
  var Src := WriteUnit('USrc', CtorClass);
  Assert.IsTrue(ExtractInterfaceHeadless(Src, 4, False, 'IFoo', '', [], True, Info, Text, Err), Err);
  var Picked := '';
  for var M in Info.Members do
    if M.Selected and (SameText(M.Name, 'Create') or SameText(M.Name, 'Destroy') or
       SameText(M.Name, 'Make')) then
      Picked := Picked + ' ' + M.Name;
  Assert.AreEqual('', Picked,
    'M38b: the default selection ticks members an interface cannot hold');
end;

procedure TAuditReproRefactoringsTests.M38c_ClassMethodIsNotEmittedAsInstanceMethod;
var
  Info: TExtractInterfaceInfo;
  Text, Err: string;
begin
  var Src := WriteUnit('USrc', CtorClass);
  ExtractInterfaceHeadless(Src, 4, False, 'IFoo', '', ['Run', 'Make'], True, Info, Text, Err);
  Assert.IsFalse(Text.Contains('Make'),
    'M38c: "class function Make" went into the interface as an instance method ' +
    '(the class then does not implement it, E2291):' + NL + Text);
end;

procedure TAuditReproRefactoringsTests.M38d_InjectionGoesBeforeTheMethodsOwnEnd;
var
  Report, Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class(TObject)',
    '  private', '    FFlag: Boolean;', '  public', '    procedure AfterConstruction; override;',
    '  end;', 'implementation', 'procedure TFoo.AfterConstruction;', 'begin',
    '  if FFlag then begin', '    FFlag := False;', '  end;', '  inherited;', 'end;', 'end.']);
  Assert.IsTrue(DelegateIInterfaceHeadless(Src, 4, False, Report, Err), Err);
  var Out_ := LinesOf(GFake.Content(Src));
  var Inh := IndexOfLine(Out_, 'inherited;');
  var Dec_ := IndexOfLine(Out_, 'AtomicDecrement(FRefCount);');
  Assert.IsTrue((Inh >= 0) and (Dec_ > Inh),
    'M38d: AtomicDecrement(FRefCount) was injected inside the "if" (it closed at the ' +
    'if''s "end;"), so it runs only when FFlag is set');
end;

procedure TAuditReproRefactoringsTests.M38d_InjectionFindsTheExactMethod;
var
  Report, Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class(TObject)',
    '  public', '    procedure AfterConstructionHelper;', '    procedure AfterConstruction; override;',
    '  end;', 'implementation', 'procedure TFoo.AfterConstructionHelper;', 'begin', 'end;',
    'procedure TFoo.AfterConstruction;', 'begin', '  inherited;', 'end;', 'end.']);
  Assert.IsTrue(DelegateIInterfaceHeadless(Src, 4, False, Report, Err), Err);
  var Out_ := LinesOf(GFake.Content(Src));
  var Header := IndexOfLine(Out_, 'procedure TFoo.AfterConstruction;');
  var Dec_ := IndexOfLine(Out_, 'AtomicDecrement(FRefCount);');
  Assert.IsTrue((Header >= 0) and (Dec_ > Header),
    'M38d: AtomicDecrement(FRefCount) went into AfterConstructionHelper (header matched by substring)');
end;

/// <summary>The class header line of the source after an apply-mode Extract Interface.</summary>
function ExtractAndReadHeader(const ASrc, ATarget: string): string;
var
  Info: TExtractInterfaceInfo;
  Text, Err: string;
begin
  Assert.IsTrue(ExtractInterfaceHeadless(ASrc, 4, False, 'IFoo', ATarget, [], False, Info,
    Text, Err), Err);
  Result := '';
  for var L in LinesOf(GFake.Content(ASrc)) do
    if L.Contains('TFoo =') then Exit(Trim(L));
end;

procedure TAuditReproRefactoringsTests.M38e_ClassAbstractGetsACompilableAncestorList;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class abstract',
    '  public', '    procedure Run;', '  end;', 'implementation', 'procedure TFoo.Run;',
    'begin', 'end;', 'end.']);
  var Header := ExtractAndReadHeader(Src, TPath.Combine(FDir, 'IFooUnit.pas'));
  Assert.IsFalse(UpperCase(Header).Contains(') ABSTRACT'),
    'M38e: "' + Header + '" does not compile - abstract must precede the ancestor list');
end;

procedure TAuditReproRefactoringsTests.M38e_SubstringOfAnotherInterfaceIsNoDuplicate;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type',
    '  TFoo = class(TInterfacedObject, IFooBar)', '  public', '    procedure Run;', '  end;',
    'implementation', 'procedure TFoo.Run;', 'begin', 'end;', 'end.']);
  var Header := ExtractAndReadHeader(Src, TPath.Combine(FDir, 'IFooUnit.pas'));
  var Listed := False;
  var Open_ := Pos('(', Header);
  var Close_ := Pos(')', Header);
  if (Open_ > 0) and (Close_ > Open_) then
    for var E in Copy(Header, Open_ + 1, Close_ - Open_ - 1).Split([',']) do
      if SameText(Trim(E), 'IFoo') then Listed := True;
  Assert.IsTrue(Listed,
    'M38e: IFoo was taken as already listed because IFooBar contains it: "' + Header + '"');
end;

procedure TAuditReproRefactoringsTests.M38f_LfOnlyBufferIsReadAsLines;
var
  Info: TExtractInterfaceInfo;
  Text, Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class',
    '  public', '    procedure Run;', '  end;', 'implementation', 'end.']);
  // the IDE buffer of a unit checked out with LF line ends
  GFake.OpenBuffer(Src, TFile.ReadAllText(Src).Replace(#13#10, #10));
  Assert.IsTrue(ExtractInterfaceHeadless(Src, 4, False, 'IFoo', '', [], True, Info, Text, Err),
    'M38f: an LF-only buffer is read as one line - ' + Err);
end;

procedure TAuditReproRefactoringsTests.L7g_UsesSeedSkipsComments;
var
  Info: TExtractInterfaceInfo;
  Text, Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'uses',
    '  System.SysUtils, // the RTL', '  System.Classes;', 'type', '  TFoo = class', '  public',
    '    procedure Run;', '  end;', 'implementation', 'procedure TFoo.Run;', 'begin', 'end;',
    'end.']);
  var Target := TPath.Combine(FDir, 'IFooUnit.pas');
  Assert.IsTrue(ExtractInterfaceHeadless(Src, 7, False, 'IFoo', Target, [], False, Info,
    Text, Err), Err);
  var NewUnit := TFile.ReadAllText(Target);
  var UsesStart := Pos('uses', NewUnit);
  var UsesText := Copy(NewUnit, UsesStart, Pos(';', NewUnit, UsesStart) - UsesStart + 1);
  Assert.IsFalse(UsesText.Contains('//'),
    'L7g: the uses seed copied a comment, which now swallows the rest of the clause ' +
    'including its ";":' + NL + UsesText);
end;

procedure TAuditReproRefactoringsTests.L7h_FailedExtractLeavesNoUnitInTheProject;
var
  Info: TExtractInterfaceInfo;
  Text, Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'type', '  TFoo = class',
    '  public', '    procedure Run;', '  end;', 'implementation', 'procedure TFoo.Run;',
    'begin', 'end;', 'end.']);
  var Target := TPath.Combine(FDir, 'IFooUnit.pas');
  // the source edit fails: the editor refuses it and the file is read-only
  GFake.RefuseWritesTo := Src;
  SetFileAttributes(PChar(Src), FILE_ATTRIBUTE_READONLY);
  var Done := ExtractInterfaceHeadless(Src, 4, False, 'IFoo', Target, [], False, Info, Text, Err);
  Assert.IsFalse(Done, 'the source write could not land, the extraction must fail');
  Assert.AreEqual(0, Integer(Length(GFake.Added)),
    'L7h: the extraction failed (' + Err + ') but IFooUnit.pas was already added to the project');
end;

{ extract method }

procedure TAuditReproRefactoringsTests.M37b_PreviewSavesNothing;
var
  Preview: TExtractMethodPreview;
  Err: string;
begin
  var Src := WriteUnit('USrc', ['unit USrc;', 'interface', 'implementation', 'procedure Run;',
    'var', '  I: Integer;', 'begin', '  I := 1;', '  I := I + 1;', 'end;', 'end.']);
  GFake.LspJson := TPath.Combine(FDir, 'none.delphilsp.json');
  GFake.RaiseOnProject := True;   // stops the run before DelphiLSP would start
  ExtractMethodHeadless(Src, 8, 8, 'Step', {AApply:} False, Preview, Err);
  Assert.AreEqual(0, GFake.Saves,
    'M37b: extract_method with apply=false saved every modified file (SaveAllFiles)');
end;

{ change signature }

procedure TAuditReproRefactoringsTests.L7m_GenericCallWithoutArguments;
var
  HD, HI: TSigHeader;
  Why: string;
  S: TArray<TSigSource>;
begin
  var Content := string.Join(#13#10, ['unit U5;', 'interface', 'type', '  TFoo = class',
    '    procedure Run<T>;', '  end;', 'implementation', 'procedure TFoo.Run<T>;', 'begin', 'end;',
    'procedure Use(F: TFoo);', 'begin', '  F.Run<Integer>;', 'end;', 'end.']);
  var Lines := LinesOf(Content);
  Assert.IsTrue(LocateSigHeader(Content, 4, 'Run', True, HD, Why), Why);
  Assert.IsTrue(LocateSigHeader(Content, 7, 'Run', False, HI, Why), Why);
  var Calls := LocateSigCalls(Content, [Point(Pos('Run<', Lines[12]) - 1, 12)], 3, 0);
  SetLength(S, 1);
  S[0].FilePath := 'U5.pas';
  S[0].Content := Content;
  var New_: TArray<TNewParam>;
  SetLength(New_, 1);
  New_[0].Param := ParseParamList('A: Integer')[0];
  New_[0].OldIndex := -1;
  New_[0].CallValue := '5';
  var Plan := PlanSignatureEdits(S, [HD, HI], Calls, New_);
  Assert.IsTrue(Plan.Ok, 'L7m: ' + string.Join('|', Plan.Errors));
  var Out_ := ApplySigEdits(Content, Plan.Edits, 0).Split([#10]);
  Assert.AreEqual('  F.Run<Integer>(5);', Out_[12],
    'L7m: the new argument list went between the name and its generic arguments');
end;

procedure TAuditReproRefactoringsTests.L7n_NestedCallOfTheSameRoutineIsRefusedClearly;
var
  HD, HI: TSigHeader;
  Why: string;
  S: TArray<TSigSource>;
begin
  var Content := string.Join(#13#10, ['unit U6;', 'interface',
    'function Foo(A: Integer): Integer;', 'implementation', 'function Foo(A: Integer): Integer;',
    'begin', '  Result := A;', 'end;', 'procedure Use;', 'begin', '  Foo(Foo(1));', 'end;', 'end.']);
  var Lines := LinesOf(Content);
  Assert.IsTrue(LocateSigHeader(Content, 2, 'Foo', True, HD, Why), Why);
  Assert.IsTrue(LocateSigHeader(Content, 4, 'Foo', False, HI, Why), Why);
  var Outer := Pos('Foo(', Lines[10]);
  var Inner := Pos('Foo(', Lines[10], Outer + 1);
  var Calls := LocateSigCalls(Content, [Point(Outer - 1, 10), Point(Inner - 1, 10)], 3, 0);
  SetLength(S, 1);
  S[0].FilePath := 'U6.pas';
  S[0].Content := Content;
  var New_ := UnchangedSignature(HD.Params);
  SetLength(New_, 2);
  New_[1].Param := ParseParamList('B: Integer')[0];
  New_[1].OldIndex := -1;
  New_[1].CallValue := '0';
  var Plan := PlanSignatureEdits(S, [HD, HI], Calls, New_);
  var Errors := string.Join('|', Plan.Errors);
  Assert.IsFalse(Errors.Contains('please report'),
    'L7n: ordinary code (a call nested in a call of the same routine) is reported as a ' +
    'planner bug: ' + Errors);
end;

initialization
  TDUnitX.RegisterTestFixture(TAuditReproRefactoringsTests);

end.
