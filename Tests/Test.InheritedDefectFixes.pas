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
///  Regression tests for the defects this fork inherited from upstream and
///  fixed on 2026-09-03 after the audit of `jaenicke/DelphiRefactoringLight`
///  at `a03d564`. Each fixture pins ONE fixed behaviour, and each was
///  confirmed to fail against the pre-fix code.
///
///  Only the pure, IDE-free surface is covered here - which is exactly the
///  surface the audit identified as the highest-leverage test target:
///  `CanTakeSemicolon`, `TLspUri` and `TDelphiFileEncoding` need neither a
///  running IDE nor an LSP session, yet three of the fixed CRITICALs are
///  reachable through them from a plain string or a temp file.
///
///  Deliberately NOT covered: the editor write paths
///  (`ReadWholeEditorBuffer`, the diagnostics-cache key) need
///  `IOTASourceEditor` / a live LSP push and so belong to an integration
///  harness, not this suite.
/// </summary>
unit Test.InheritedDefectFixes;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TCanTakeSemicolonTests = class
  public
    // ---- block openers that must REFUSE a trailing ';' ----
    [Test] procedure ForHeaderBrokenBeforeDo_IsRefused;
    [Test] procedure TrailingDivOperator_IsRefused;
    [Test] procedure TrailingAsOperator_IsRefused;
    [Test] procedure RecordOpener_IsRefused;
    [Test] procedure WhileHeader_IsRefused;
    [Test] procedure VisibilitySectionIsNotAStatement_IsRefused;

    // ---- trailing comments must not defeat the guard ----
    [Test] procedure BraceCommentAfterThen_IsStillRefused;
    [Test] procedure ParenStarCommentAfterBegin_IsStillRefused;
    [Test] procedure BraceCommentAfterStatement_IsStillAccepted;

    // ---- genuine statements must still be ACCEPTED ----
    [Test] procedure PlainAssignment_IsAccepted;
    [Test] procedure InheritedAlone_IsAccepted;
    [Test] procedure AssignmentOfNil_IsAccepted;
  end;

  [TestFixture]
  TLspUriUncTests = class
  public
    [Test] procedure DriveLetterUri_RoundTrips;
    [Test] procedure UncUri_KeepsItsLeadingBackslashes;
    [Test] procedure UncPath_IsEmittedAsTheTwoSlashForm;
    [Test] procedure UncPath_RoundTrips;
  end;

  [TestFixture]
  TFileEncodingTests = class
  private
    FDir: string;
    function WriteRaw(const AName: string; const ABytes: array of Byte): string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    [Test] procedure BomlessUtf8_IsDetectedAsUtf8_NotAnsi;
    [Test] procedure BomlessUtf8_RoundTripsWithoutLoss;
    [Test] procedure BomlessUtf8_DoesNotGainABom;
    [Test] procedure InvalidUtf8_IsStillDetectedAsAnsi;
    [Test] procedure Utf8WithBom_KeepsItsBom;
    [Test] procedure WriteAll_LeavesNoTempFileBehind;
  end;

implementation

uses
  System.SysUtils, System.IOUtils, System.Classes,
  Expert.AutoImport, Lsp.Uri, Delphi.FileEncoding;

{ TCanTakeSemicolonTests }

// A missing semicolon always follows a STATEMENT. Every case below ends in
// something a statement cannot end with, so offering to append ';' there
// would emit code that does not compile.

procedure TCanTakeSemicolonTests.ForHeaderBrokenBeforeDo_IsRefused;
begin
  // The last word is '10', not a keyword - which is why the original list
  // (BEGIN/THEN/ELSE/DO/OF/... but no FOR) let this through and offered a
  // ';' inside a for-header.
  Assert.IsFalse(CanTakeSemicolon('  for I := 0 to 10'),
    'a for-header wrapped before "do" cannot take a semicolon');
end;

procedure TCanTakeSemicolonTests.TrailingDivOperator_IsRefused;
begin
  // AND/OR/NOT/IN were listed; the identically-behaving DIV/MOD/SHL/SHR/
  // XOR/AS/IS were not.
  Assert.IsFalse(CanTakeSemicolon('  X := A div'),
    'a line ending in a binary operator continues on the next line');
end;

procedure TCanTakeSemicolonTests.TrailingAsOperator_IsRefused;
begin
  Assert.IsFalse(CanTakeSemicolon('  Y := Obj as'),
    '"as" is an operator, not a statement end');
end;

procedure TCanTakeSemicolonTests.RecordOpener_IsRefused;
begin
  Assert.IsFalse(CanTakeSemicolon('  TFoo = record'),
    '"record" opens a block');
end;

procedure TCanTakeSemicolonTests.WhileHeader_IsRefused;
begin
  Assert.IsFalse(CanTakeSemicolon('  while not Eof(F)'),
    'a while-header wrapped before "do" cannot take a semicolon');
end;

procedure TCanTakeSemicolonTests.VisibilitySectionIsNotAStatement_IsRefused;
begin
  Assert.IsFalse(CanTakeSemicolon('  published'),
    'a visibility section is not a statement');
end;

procedure TCanTakeSemicolonTests.BraceCommentAfterThen_IsStillRefused;
begin
  // The single most common formatting in real code, and it defeated the
  // whole guard: the line then ended in '}', the last-word scan yielded '',
  // and '' is in no keyword list - so the guard returned True.
  Assert.IsFalse(CanTakeSemicolon('  if X then { TODO }'),
    'a trailing brace comment must not hide the block opener');
end;

procedure TCanTakeSemicolonTests.ParenStarCommentAfterBegin_IsStillRefused;
begin
  Assert.IsFalse(CanTakeSemicolon('  begin (* los gehts *)'),
    'a trailing (* *) comment must not hide the block opener');
end;

procedure TCanTakeSemicolonTests.BraceCommentAfterStatement_IsStillAccepted;
begin
  // The comment strip must not make the guard over-refuse either: with the
  // comment removed this is an ordinary statement.
  Assert.IsTrue(CanTakeSemicolon('  Foo := Bar { hinweis }'),
    'stripping the comment must still leave a statement visible');
end;

procedure TCanTakeSemicolonTests.PlainAssignment_IsAccepted;
begin
  Assert.IsTrue(CanTakeSemicolon('  X := 1'),
    'an assignment is exactly where a semicolon can be missing');
end;

procedure TCanTakeSemicolonTests.InheritedAlone_IsAccepted;
begin
  // 'inherited;' is legal, so INHERITED must stay OUT of the opener list.
  Assert.IsTrue(CanTakeSemicolon('  inherited'),
    '"inherited" is a complete statement');
end;

procedure TCanTakeSemicolonTests.AssignmentOfNil_IsAccepted;
begin
  // 'X := nil;' is legal, so NIL must stay out of the list too.
  Assert.IsTrue(CanTakeSemicolon('  FList := nil'),
    'an assignment of nil is a complete statement');
end;

{ TLspUriUncTests }

procedure TLspUriUncTests.DriveLetterUri_RoundTrips;
begin
  Assert.AreEqual('C:\Temp\X.pas',
    TLspUri.FileUriToPath('file:///C:/Temp/X.pas'),
    'the ordinary drive-letter form is unaffected by the UNC fix');
end;

procedure TLspUriUncTests.UncUri_KeepsItsLeadingBackslashes;
begin
  // "file://" with TWO slashes is the UNC form. Stripping all seven
  // characters produced "SERVER\SHARE\X.PAS" - and because the diagnostics
  // cache in Lsp.Client writes under that path but reads via
  // ExpandFileName, writer and reader keys never matched again.
  Assert.AreEqual('\\server\share\X.pas',
    TLspUri.FileUriToPath('file://server/share/X.pas'),
    'a UNC uri must keep its two leading backslashes');
end;

procedure TLspUriUncTests.UncPath_IsEmittedAsTheTwoSlashForm;
begin
  Assert.AreEqual('file://server/share/X.pas',
    TLspUri.PathToFileUri('\\server\share\X.pas'),
    'a UNC path must not be emitted as file:///// ');
end;

procedure TLspUriUncTests.UncPath_RoundTrips;
const
  UncPath = '\\server\share\Sub Dir\X.pas';
begin
  Assert.AreEqual(UncPath,
    TLspUri.FileUriToPath(TLspUri.PathToFileUri(UncPath)),
    'path -> uri -> path must be identity for UNC, spaces included');
end;

{ TFileEncodingTests }

procedure TFileEncodingTests.Setup;
begin
  FDir := TPath.Combine(TPath.GetTempPath,
    'RefLightEncTest_' + FormatDateTime('yyyymmddhhnnsszzz', Now));
  TDirectory.CreateDirectory(FDir);
end;

procedure TFileEncodingTests.TearDown;
begin
  if (FDir <> '') and TDirectory.Exists(FDir) then
    try
      TDirectory.Delete(FDir, True);
    except
      // a locked temp file must not fail the test run
    end;
end;

function TFileEncodingTests.WriteRaw(const AName: string;
  const ABytes: array of Byte): string;
var
  Stream: TFileStream;
begin
  Result := TPath.Combine(FDir, AName);
  Stream := TFileStream.Create(Result, fmCreate);
  try
    if Length(ABytes) > 0 then
      Stream.WriteBuffer(ABytes[0], Length(ABytes));
  finally
    Stream.Free;
  end;
end;

procedure TFileEncodingTests.BomlessUtf8_IsDetectedAsUtf8_NotAnsi;
var
  Path: string;
begin
  // "Gr" + C3 B6 ("oe") + "sse" - valid UTF-8, no BOM.
  Path := WriteRaw('bomless.pas',
    [$47, $72, $C3, $B6, $73, $73, $65]);
  Assert.AreEqual(65001, TDelphiFileEncoding.Detect(Path).CodePage,
    'a BOM-less file whose bytes are valid UTF-8 must be read as UTF-8');
end;

procedure TFileEncodingTests.BomlessUtf8_RoundTripsWithoutLoss;
var
  Path, Text: string;
begin
  // The defect: Detect returned ANSI for this file, so the content was
  // decoded as CP1252 AND written back through WideCharToMultiByte, which
  // substitutes '?' for anything it cannot represent - unrecoverable.
  Path := WriteRaw('roundtrip.pas',
    [$27, $52, $C3, $B6, $6E, $74, $67, $65, $6E, $20, $E2, $86, $92,
     $20, $C3, $98, $20, $32, $30, $20, $C2, $B5, $6D, $27]);
  Text := TDelphiFileEncoding.ReadAll(Path);
  Assert.AreEqual('''Röntgen → Ø 20 µm''', Text,
    'the arrow and the diameter sign must survive reading');

  TDelphiFileEncoding.WriteAll(Path, Text, TDelphiFileEncoding.Detect(Path));
  Assert.AreEqual(Text, TDelphiFileEncoding.ReadAll(Path),
    'and must survive a write-back through the detected encoding');
  Assert.IsFalse(TDelphiFileEncoding.ReadAll(Path).Contains('?'),
    'no character may be replaced by a question mark');
end;

procedure TFileEncodingTests.BomlessUtf8_DoesNotGainABom;
var
  Path: string;
  Bytes: TBytes;
begin
  // Reading BOM-less UTF-8 correctly must not cost a BOM on write-back:
  // that would be a whole-file diff on every refactoring touch.
  Path := WriteRaw('nobom.pas', [$41, $C3, $A4, $42]);
  TDelphiFileEncoding.WriteAll(Path, TDelphiFileEncoding.ReadAll(Path),
    TDelphiFileEncoding.Detect(Path));
  Bytes := TFile.ReadAllBytes(Path);
  Assert.IsTrue(Length(Bytes) >= 3, 'file must not be empty');
  Assert.IsFalse((Bytes[0] = $EF) and (Bytes[1] = $BB) and (Bytes[2] = $BF),
    'a file that had no BOM must not acquire one');
end;

procedure TFileEncodingTests.InvalidUtf8_IsStillDetectedAsAnsi;
var
  Path: string;
begin
  // $E4 alone is 'ä' in CP1252 but an incomplete UTF-8 sequence, so the
  // ANSI fallback must still be reachable - the fix must not blanket-assume
  // UTF-8.
  Path := WriteRaw('ansi.pas', [$47, $72, $E4, $73, $73, $65]);
  Assert.AreEqual(TEncoding.Default.CodePage,
    TDelphiFileEncoding.Detect(Path).CodePage,
    'bytes that are not valid UTF-8 must still be treated as ANSI');
end;

procedure TFileEncodingTests.Utf8WithBom_KeepsItsBom;
var
  Path: string;
  Bytes: TBytes;
begin
  Path := WriteRaw('withbom.pas', [$EF, $BB, $BF, $41, $C3, $A4, $42]);
  TDelphiFileEncoding.WriteAll(Path, TDelphiFileEncoding.ReadAll(Path),
    TDelphiFileEncoding.Detect(Path));
  Bytes := TFile.ReadAllBytes(Path);
  Assert.IsTrue((Bytes[0] = $EF) and (Bytes[1] = $BB) and (Bytes[2] = $BF),
    'a file that had a BOM must keep it');
end;

procedure TFileEncodingTests.WriteAll_LeavesNoTempFileBehind;
var
  Path: string;
begin
  // WriteAll now writes beside the target and swaps, so that a failed write
  // cannot leave the user's unit truncated to zero bytes. The scratch file
  // must not survive a successful write.
  Path := WriteRaw('swap.pas', [$41, $42, $43]);
  TDelphiFileEncoding.WriteAll(Path, 'XYZ', TEncoding.UTF8);
  Assert.IsFalse(TFile.Exists(Path + '.rl~tmp'),
    'the scratch file must be gone after a successful write');
  Assert.AreEqual('XYZ', TDelphiFileEncoding.ReadAll(Path),
    'and the new content must be in place');
end;

initialization
  TDUnitX.RegisterTestFixture(TCanTakeSemicolonTests);
  TDUnitX.RegisterTestFixture(TLspUriUncTests);
  TDUnitX.RegisterTestFixture(TFileEncodingTests);

end.
