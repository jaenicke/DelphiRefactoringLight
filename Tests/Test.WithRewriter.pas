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
///  Fixture-driven tests for TWithRewriter - the layer that decides whether a
///  with-statement can be rewritten at all.
///
///  The rewriter reaches DelphiLSP through exactly two calls, GotoDefinition
///  and RefreshDocument, both virtual. TFakeLspClient below overrides them and
///  scripts the answers, so these tests run headless with no language server.
///  TLspClient.Create only initialises fields - it launches nothing - so the
///  fake is safe to construct and free without ever calling Start.
///
///  What is deliberately NOT claimed: these tests do not verify that a
///  successful rewrite is semantically correct. They verify the far more
///  important property that the rewriter REFUSES when it does not actually
///  know - which is where the silent-rebind bugs lived.
/// </summary>
unit Test.WithRewriter;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TWithRewriterTests = class
  public
    [Test] procedure UnresolvableTarget_IsNotAutoRewritable;
    [Test] procedure UnresolvableTarget_RaisesTypeUnresolved;
    [Test] procedure TargetDeclaredInFileNotOnDisk_IsNotAutoRewritable;
    [Test] procedure TargetDeclaredInFileNotOnDisk_RaisesTargetRangeUnusable;
    [Test] procedure Rewrite_NeverRaises_OnHostileLspAnswers;
  end;

implementation

uses
  System.SysUtils,
  Lsp.Client, Lsp.Protocol, Lsp.Uri,
  Expert.WithScanner, Expert.WithRewriter;

type
  /// <summary>How the fake should answer GotoDefinition.</summary>
  TFakeMode = (
    /// <summary>Always return an empty array - LSP knows nothing.</summary>
    fmNothing,
    /// <summary>Return a location in a file that does not exist on disk.
    ///  This is the real-world shape that caused the trouble: a type from a
    ///  precompiled package whose source the server cannot reach.</summary>
    fmMissingFile,
    /// <summary>Raise, to prove the rewriter never lets an exception out.</summary>
    fmRaise);

  /// <summary>A TLspClient that answers from a script instead of a server.</summary>
  TFakeLspClient = class(TLspClient)
  private
    FMode: TFakeMode;
  public
    constructor Create(AMode: TFakeMode); reintroduce;
    /// <summary>Scripted stand-in; never contacts a server.</summary>
    function GotoDefinition(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>; override;
    /// <summary>No-op; the fake has no document state to refresh.</summary>
    procedure RefreshDocument(const AFilePath: string); override;
  end;

const
  /// <summary>A path that is guaranteed not to exist.</summary>
  MISSING_FILE = 'C:\__rl_no_such_folder__\NoSuchUnit.pas';

{ TFakeLspClient }

constructor TFakeLspClient.Create(AMode: TFakeMode);
begin
  inherited Create('');  // no exe path - Start is never called
  FMode := AMode;
end;

function TFakeLspClient.GotoDefinition(const AFilePath: string;
  ALine, ACol: Integer): TArray<TLspLocation>;
var
  Loc: TLspLocation;
begin
  Result := nil;
  case FMode of
    fmNothing:
      Exit;
    fmMissingFile:
      begin
        Loc := Default(TLspLocation);
        Loc.Uri := TLspUri.PathToFileUri(MISSING_FILE);
        Loc.Range.Start.Line := 10;
        Loc.Range.Start.Character := 2;
        Loc.Range.End_ := Loc.Range.Start;
        Result := [Loc];
      end;
    fmRaise:
      raise Exception.Create('scripted LSP failure');
  end;
end;

procedure TFakeLspClient.RefreshDocument(const AFilePath: string);
begin
  // deliberately nothing
end;

/// <summary>Scans a source, asserts exactly one with-statement, and rewrites
///  it against a fake client in the given mode.</summary>
function RewriteOnly(const ASource: string; AMode: TFakeMode): TWithRewriteResult;
var
  Occs: TArray<TWithOccurrence>;
  Client: TFakeLspClient;
begin
  Occs := TWithScanner.ScanSource(ASource);
  Assert.AreEqual<Integer>(1, Length(Occs), 'fixture should contain exactly one with-statement');

  Client := TFakeLspClient.Create(AMode);
  try
    Result := TWithRewriter.Rewrite(Client, 'C:\fixture\Unit1.pas', ASource,
      Occs[0], TWithRewriteSettings.Defaults);
  finally
    Client.Free;
  end;
end;

/// <summary>The fixture used by every test here: one simple with-statement
///  whose body touches two members.</summary>
function Fixture: string;
begin
  Result :=
    'procedure TForm1.Refresh;'#13#10 +
    'begin'#13#10 +
    '  with FQuery do'#13#10 +
    '    begin'#13#10 +
    '    Close;'#13#10 +
    '    Open;'#13#10 +
    '    end;'#13#10 +
    'end;'#13#10;
end;

{ TWithRewriterTests }

procedure TWithRewriterTests.UnresolvableTarget_IsNotAutoRewritable;
var
  R: TWithRewriteResult;
begin
  // If the server cannot resolve the target at all, the rewriter must not
  // offer a rewrite. Emitting the body unqualified would rebind every member
  // access to whatever the enclosing scope happens to provide.
  R := RewriteOnly(Fixture, fmNothing);
  Assert.IsFalse(R.IsAutoRewritable,
    'an unresolvable target must never be reported as auto-rewritable');
end;

procedure TWithRewriterTests.UnresolvableTarget_RaisesTypeUnresolved;
var
  R: TWithRewriteResult;
begin
  R := RewriteOnly(Fixture, fmNothing);
  Assert.IsTrue(wriTypeUnresolved in R.Issues,
    'wriTypeUnresolved should be set when the target type cannot be resolved');
end;

procedure TWithRewriterTests.TargetDeclaredInFileNotOnDisk_IsNotAutoRewritable;
var
  R: TWithRewriteResult;
begin
  // THE REGRESSION THAT MATTERS. A target whose declaring file is not on disk
  // (an ElevateDB/LMD/TMS type from a precompiled package) used to take a
  // "partial" path that set Resolved := True with an empty member list. No
  // body identifier could then be tested for membership, so the with was
  // stripped and every member access went out BARE - and the dialog still
  // said "ok". That is exactly what happened to
  // "with MainForm.SystemDatabaseQuery do" in DBiManager\publishdlg.pas.
  R := RewriteOnly(Fixture, fmMissingFile);
  Assert.IsFalse(R.IsAutoRewritable,
    'a target with no usable class range must never be reported as auto-rewritable');
end;

procedure TWithRewriterTests.TargetDeclaredInFileNotOnDisk_RaisesTargetRangeUnusable;
var
  R: TWithRewriteResult;
begin
  R := RewriteOnly(Fixture, fmMissingFile);
  Assert.IsTrue((wriClassRangeUnknown in R.Issues) or (wriTypeUnresolved in R.Issues),
    'an unusable target range must raise an issue rather than pass silently');
end;

procedure TWithRewriterTests.Rewrite_NeverRaises_OnHostileLspAnswers;
var
  R: TWithRewriteResult;
begin
  // Documented contract: Rewrite never raises; on LSP error it sets an issue
  // and returns a partial result. A wizard that let an exception out would
  // take it into the IDE's dispatch.
  Assert.WillNotRaise(
    procedure
    begin
      R := RewriteOnly(Fixture, fmRaise);
      Assert.IsFalse(R.IsAutoRewritable,
        'a raising LSP must not yield an auto-rewritable result');
    end,
    Exception,
    'Rewrite must never let an exception escape');
end;

initialization
  TDUnitX.RegisterTestFixture(TWithRewriterTests);

end.
