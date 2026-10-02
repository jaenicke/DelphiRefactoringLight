(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Audit repro tests (issue #22), area "project checks, quick fixes, blame,
///  auto-import": one test per open item that can be shown headless through
///  the existing public API. RED means the item is still open; a test can be
///  deleted once it goes green. Test names and messages carry the item id.
/// </summary>
unit Test.AuditReproChecks;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TAuditReproChecksTests = class
  private
    FDir: string;
    function WriteFile(const AName, AText: string): string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    /// <summary>M23: a DFM whose first line is blank must not raise out of
    ///  CheckProject (the binary-DFM guard indexes an empty Split result).</summary>
    [Test] procedure M23_DfmCheck_BlankFirstLine_DoesNotRaise;
    /// <summary>M24: an interface without a GUID must not be reported with
    ///  the GUID of the interface declared after it.</summary>
    [Test] procedure M24_InterfaceGuid_MissingGuid_DoesNotBorrowTheNextOne;
    /// <summary>M25: an implementation must be blocked while the class
    ///  declaration of its unit still diverges - with entries shaped the way
    ///  Collect produces them (implementation Container = '').</summary>
    [Test] procedure M25_AlignBlocker_ImplementationWaitsForItsClassDeclaration;
    /// <summary>L3b: ApplyFix rewrites the declaration in the FORM class, not
    ///  a same-named method of another class declared earlier in the unit.</summary>
    [Test] procedure L3b_DfmApplyFix_RewritesTheFormClassDeclaration;
    /// <summary>L3d: interfaces inside { } and (* *) comments are not real
    ///  declarations.</summary>
    [Test] procedure L3d_InterfaceGuid_CommentedOutDeclarations_AreIgnored;
    /// <summary>L3d: "type IFoo = interface" declares IFoo, not "type IFoo".</summary>
    [Test] procedure L3d_InterfaceGuid_OneLineTypeDeclaration_NameIsTheInterface;
    /// <summary>L4f: the column sort must be a total order - the same cells
    ///  must come out in the same order whatever order they went in.</summary>
    [Test] procedure L4f_ListViewSort_OrderDoesNotDependOnInputOrder;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Generics.Collections,
  Vcl.Forms, Vcl.ComCtrls,
  Expert.EditorHelperIntf, Expert.DfmEventCheck, Expert.InterfaceGuidCheck,
  Expert.SignatureCheck, Expert.ListViewSort;

const
  NL = #13#10;

type
  /// <summary>In-memory IEditorHelper: open buffers live in a dictionary;
  ///  only reading and the line/whole-file writes carry behaviour.</summary>
  TFakeEditor = class(TInterfacedObject, IEditorHelper)
  private
    FBuffers: TDictionary<string, string>;
    function Edit(const AFile: string; ALine: Integer;
      const AProc: TProc<TStringList, Integer>; AAllowAppend: Boolean = False): Boolean;
  public
    constructor Create;
    destructor Destroy; override;
    procedure OpenBuffer(const AFile, AContent: string);
    function Content(const AFile: string): string;

    function GetOpenSourceFiles: TArray<string>;
    function IsFormInDesigner(const APasFile: string): Boolean;
    function GetDesignerRequiredUnits(const APasFile: string;
      out AUnits: TArray<TDesignerRequiredUnit>;
      out AComplete: Boolean): Boolean;
    function RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
      AIsMethod: Boolean; out AMessage: string): Boolean;
    function ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;
    function ReplaceFileContent(const AFilePath: string; const ANewContent: string): Boolean;
    function ReplaceLineAt(const AFilePath: string; ALine: Integer;
      const ANewContent: string): Boolean;
    function DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;
    function InsertTextAtLineStart(const AFilePath: string; ALine: Integer;
      const AText: string): Boolean;
    function GetCurrentContext: TEditorContext;
    function GetActiveFileName: string;
    function GetCaretLineCol(out ALine, ACol: Integer): Boolean;
    function RawColumn(const AFile: string; ALine, ADisplayCol: Integer): Integer;
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
  GFake: TFakeEditor;   // weak view of the installed helper

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

procedure TFakeEditor.OpenBuffer(const AFile, AContent: string);
begin
  FBuffers.AddOrSetValue(AFile, AContent);
end;

function TFakeEditor.Content(const AFile: string): string;
begin
  if not FBuffers.TryGetValue(AFile, Result) then Result := '';
end;

function TFakeEditor.Edit(const AFile: string; ALine: Integer;
  const AProc: TProc<TStringList, Integer>; AAllowAppend: Boolean): Boolean;
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    SL.Text := Content(AFile);
    Result := (ALine >= 1) and ((ALine <= SL.Count) or (AAllowAppend and (ALine = SL.Count + 1)));
    if not Result then Exit;
    AProc(SL, ALine - 1);
    FBuffers.AddOrSetValue(AFile, SL.Text);
  finally
    SL.Free;
  end;
end;

function TFakeEditor.ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;
begin
  Result := FBuffers.TryGetValue(AFilePath, AContent);
  if not Result then AContent := '';
end;

function TFakeEditor.ReplaceFileContent(const AFilePath, ANewContent: string): Boolean;
begin
  FBuffers.AddOrSetValue(AFilePath, ANewContent);
  Result := True;
end;

function TFakeEditor.ReplaceLineAt(const AFilePath: string; ALine: Integer;
  const ANewContent: string): Boolean;
begin
  Result := Edit(AFilePath, ALine,
    procedure(SL: TStringList; I: Integer)
    begin
      SL[I] := ANewContent;
    end);
end;

function TFakeEditor.DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;
begin
  Result := Edit(AFilePath, ALine,
    procedure(SL: TStringList; I: Integer)
    begin
      SL.Delete(I);
    end);
end;

function TFakeEditor.InsertTextAtLineStart(const AFilePath: string; ALine: Integer;
  const AText: string): Boolean;
begin
  // RAW text at the line's start, as the IDE does it.
  Result := Edit(AFilePath, ALine,
    procedure(SL: TStringList; I: Integer)
    begin
      if I < SL.Count then
        SL[I] := AText + SL[I]
      else
        SL.Add(AText.TrimRight([#13, #10]));
    end, True);
end;

function TFakeEditor.GetOpenSourceFiles: TArray<string>;
begin
  Result := nil;
end;

function TFakeEditor.IsFormInDesigner(const APasFile: string): Boolean;
begin
  Result := False;
end;

function TFakeEditor.GetDesignerRequiredUnits(const APasFile: string;
  out AUnits: TArray<TDesignerRequiredUnit>; out AComplete: Boolean): Boolean;
begin
  AUnits := nil;
  AComplete := False;
  Result := False;
end;

function TFakeEditor.RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
  AIsMethod: Boolean; out AMessage: string): Boolean;
begin
  AMessage := '';
  Result := False;
end;

function TFakeEditor.GetCurrentContext: TEditorContext;
begin
  Result := Default(TEditorContext);
end;

function TFakeEditor.GetActiveFileName: string;
begin
  Result := '';
end;

function TFakeEditor.RawColumn(const AFile: string; ALine,
  ADisplayCol: Integer): Integer;
begin
  Result := ADisplayCol;   // no tab expansion in a test stand-in
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

function TFakeEditor.BuildSearchPathFromProject(const ADprojPath, ARootPath: string): string;
begin
  Result := '';
end;

function TFakeEditor.FindDelphiLspJson: string;
begin
  Result := '';
end;

function TFakeEditor.ReplaceSelection(const AFilePath: string;
  AStartLine, AStartCol, AEndLine, AEndCol: Integer; const ANewText: string): Boolean;
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
  ALine, ACol: Integer; AHighlightLen: Integer): Boolean;
begin
  Result := False;
end;

function TFakeEditor.AddFileToActiveProject(const AFilePath: string): Boolean;
begin
  Result := False;
end;

function TFakeEditor.GetSelection(out AFilePath: string;
  out AStartLine, AStartCol, AEndLine, AEndCol: Integer; out AText: string): Boolean;
begin
  AFilePath := ''; AStartLine := 0; AStartCol := 0;
  AEndLine := 0; AEndCol := 0; AText := '';
  Result := False;
end;

{ TAuditReproChecksTests }

procedure TAuditReproChecksTests.Setup;
var
  Fake: TFakeEditor;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'RLAuditChecks_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(FDir);
  Fake := TFakeEditor.Create;
  GFake := Fake;              // the interface reference below owns it
  SetEditorImpl(Fake);
end;

procedure TAuditReproChecksTests.TearDown;
begin
  SetEditorImpl(nil);
  GFake := nil;
  if TDirectory.Exists(FDir) then
    TDirectory.Delete(FDir, True);
end;

function TAuditReproChecksTests.WriteFile(const AName, AText: string): string;
begin
  Result := TPath.Combine(FDir, AName);
  TFile.WriteAllBytes(Result, TEncoding.UTF8.GetBytes(AText));
end;

const
  /// <summary>A form unit whose handler carries one parameter too many.</summary>
  FORM_UNIT =
    'unit UForm;' + NL +
    'interface' + NL +
    'uses Vcl.Forms;' + NL +
    'type' + NL +
    '  TForm1 = class(TForm)' + NL +
    '    procedure Btn1Click(Sender: TObject; X: Integer);' + NL +
    '  end;' + NL +
    'implementation' + NL +
    'procedure TForm1.Btn1Click(Sender: TObject; X: Integer);' + NL +
    'begin' + NL +
    'end;' + NL +
    'end.' + NL;

procedure TAuditReproChecksTests.M23_DfmCheck_BlankFirstLine_DoesNotRaise;
var
  Pas: string;
begin
  Pas := WriteFile('UBlank.pas', FORM_UNIT);
  WriteFile('UBlank.dfm', '' + NL + 'object Form1: TForm1' + NL + 'end' + NL);
  Assert.WillNotRaiseAny(
    procedure
    begin
      TDfmEventChecker.CheckProject([Pas]);
    end, 'M23: a blank first DFM line raised out of CheckProject');
end;

procedure TAuditReproChecksTests.M24_InterfaceGuid_MissingGuid_DoesNotBorrowTheNextOne;
var
  F: string;
  E: TArray<TInterfaceGuidEntry>;
begin
  F := WriteFile('UIntf.pas',
    'unit UIntf;' + NL + 'interface' + NL + 'type' + NL +
    '  IMarker = interface' + NL +
    '  end;' + NL +
    '  IOther = interface [''{11111111-2222-3333-4444-555555555555}'']' + NL +
    '  end;' + NL +
    'implementation' + NL + 'end.' + NL);
  E := TInterfaceGuidChecker.Scan([F]);
  Assert.AreEqual(2, Integer(Length(E)), 'M24: two declarations');
  Assert.AreEqual('IMarker', E[0].InterfaceName, False);
  Assert.IsFalse(E[0].HasGuid, 'M24: IMarker has no GUID of its own but was given ' + E[0].Guid);
  Assert.IsFalse(E[1].IsDuplicate, 'M24: IOther is reported as a duplicate of IMarker');
end;

procedure TAuditReproChecksTests.M25_AlignBlocker_ImplementationWaitsForItsClassDeclaration;
const
  UNIT_FILE = 'C:\p\UFoo.pas';
var
  E: TSignatureEntries;
  Ref: string;

  function Entry(const AFile: string; ARole: TSignatureRole;
    const AContainer, ASig: string): TSignatureEntry;
  begin
    Result := Default(TSignatureEntry);
    Result.FilePath := AFile;
    Result.Role := ARole;
    Result.Container := AContainer;
    Result.Name := 'Bar';
    Result.RawSignature := ASig;
    Result.Normalized := TSignatureChecker.Normalize(ASig);
  end;

begin
  Ref := TSignatureChecker.Normalize('procedure Bar(A: Integer; B: string)');
  // As Collect produces them: a declaration carries its type as Container,
  // an implementation (top-level symbol "TFoo.Bar") carries none.
  E := [
    Entry('C:\p\UIntf.pas', srInterfaceDecl, 'IFoo', 'procedure Bar(A: Integer; B: string)'),
    Entry(UNIT_FILE, srClassDecl, 'TFoo', 'procedure Bar(A: Integer)'),
    Entry(UNIT_FILE, srImplementation, '', 'procedure TFoo.Bar(A: Integer)')];
  Assert.AreEqual('align the class declaration first',
    TSignatureChecker.AlignBlocker(E, 2, Ref), False,
    'M25: the implementation may be aligned while TFoo''s declaration still diverges');
end;

procedure TAuditReproChecksTests.L3b_DfmApplyFix_RewritesTheFormClassDeclaration;
const
  TWO_CLASS_UNIT =
    'unit UTwo;' + NL +
    'interface' + NL +
    'uses Vcl.Forms;' + NL +
    'type' + NL +
    '  TFrameA = class(TFrame)' + NL +
    '    procedure Btn1Click(Sender: TObject; X: Integer);' + NL +
    '  end;' + NL +
    '  TForm1 = class(TForm)' + NL +
    '    procedure Btn1Click(Sender: TObject; X: Integer);' + NL +
    '  end;' + NL +
    'implementation' + NL +
    'procedure TFrameA.Btn1Click(Sender: TObject; X: Integer);' + NL +
    'begin' + NL +
    'end;' + NL +
    'procedure TForm1.Btn1Click(Sender: TObject; X: Integer);' + NL +
    'begin' + NL +
    'end;' + NL +
    'end.' + NL;
var
  Pas, Why: string;
  Issue: TDfmEventIssue;
  Lines: TStringList;
begin
  Pas := TPath.Combine(FDir, 'UTwo.pas');
  GFake.OpenBuffer(Pas, TWO_CLASS_UNIT);
  Issue := Default(TDfmEventIssue);
  Issue.PasFile := Pas;
  Issue.Kind := eikSignatureMismatch;
  Issue.HandlerName := 'Btn1Click';
  Issue.FormClass := 'TForm1';
  Issue.ExpectedRawParams := 'Sender: TObject';
  Issue.ExpectedNorm := 'TObject';
  Assert.IsTrue(TDfmEventChecker.ApplyFix(Issue, Why), 'L3b: fix applied: ' + Why);
  Lines := TStringList.Create;
  try
    Lines.Text := GFake.Content(Pas);
    Assert.AreEqual('    procedure Btn1Click(Sender: TObject; X: Integer);', Lines[5], False,
      'L3b: TFrameA''s declaration was rewritten');
    Assert.AreEqual('    procedure Btn1Click(Sender: TObject);', Lines[8], False,
      'L3b: TForm1''s declaration was not rewritten');
  finally
    Lines.Free;
  end;
end;

procedure TAuditReproChecksTests.L3d_InterfaceGuid_CommentedOutDeclarations_AreIgnored;
var
  F: string;
  E: TArray<TInterfaceGuidEntry>;
begin
  F := WriteFile('UCmt.pas',
    'unit UCmt;' + NL + 'interface' + NL + 'type' + NL +
    '  {' + NL +
    '  IOld = interface' + NL +
    '    [''{11111111-2222-3333-4444-555555555555}'']' + NL +
    '  }' + NL +
    '  (*' + NL +
    '  IOlder = interface' + NL +
    '    [''{11111111-2222-3333-4444-555555555555}'']' + NL +
    '  *)' + NL +
    '  IReal = interface' + NL +
    '    [''{11111111-2222-3333-4444-555555555555}'']' + NL +
    '  end;' + NL +
    'implementation' + NL + 'end.' + NL);
  E := TInterfaceGuidChecker.Scan([F]);
  Assert.AreEqual(1, Integer(Length(E)),
    'L3d: interfaces inside { } and (* *) are reported as declarations');
  Assert.IsFalse(E[0].IsDuplicate, 'L3d: the commented-out copies make IReal a duplicate');
end;

procedure TAuditReproChecksTests.L3d_InterfaceGuid_OneLineTypeDeclaration_NameIsTheInterface;
var
  Name: string;
  IsDisp: Boolean;
begin
  Assert.IsTrue(IsInterfaceDeclLine('type IFoo = interface', Name, IsDisp), 'L3d: is a declaration');
  Assert.AreEqual('IFoo', Name, False, 'L3d: the name of "type IFoo = interface"');
end;

procedure TAuditReproChecksTests.L4f_ListViewSort_OrderDoesNotDependOnInputOrder;
const
  // '9' < '10' as numbers, '10' < '1a' and '1a' < '9' as text: a cycle.
  Perms: array[0..5, 0..2] of string = (
    ('9', '10', '1a'), ('9', '1a', '10'), ('10', '9', '1a'),
    ('10', '1a', '9'), ('1a', '9', '10'), ('1a', '10', '9'));
var
  Form: TForm;
  LV: TListView;
  Seen: TStringList;
begin
  Seen := TStringList.Create;
  Form := TForm.CreateNew(nil);
  try
    Seen.Sorted := True;
    Seen.Duplicates := dupIgnore;
    for var P := 0 to High(Perms) do
    begin
      LV := TListView.Create(Form);
      try
        LV.Parent := Form;
        LV.ViewStyle := vsReport;
        LV.Columns.Add.Caption := 'Cell';
        for var C := 0 to 2 do
          LV.Items.Add.Caption := Perms[P, C];
        EnableListViewSorting(LV);
        LV.OnColumnClick(LV, LV.Columns[0]);
        var Order := '';
        for var I := 0 to LV.Items.Count - 1 do
          Order := Order + LV.Items[I].Caption + ' ';
        Seen.Add(Trim(Order));
      finally
        LV.Free;
      end;
    end;
    Assert.AreEqual(1, Seen.Count,
      'L4f: the same cells sort differently by input order: ' + Seen.CommaText);
  finally
    Form.Free;
    Seen.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TAuditReproChecksTests);

end.
