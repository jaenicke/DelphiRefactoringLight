(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Tests for the standalone host's state and editor helper - the parts
///  that do not need the main form.
/// </summary>
unit Test.StandaloneState;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TStandaloneStateTests = class
  public
    [Test] procedure GotoLocation_TellsTheFormWhereToGo;
  end;

implementation

uses
  System.SysUtils, Expert.EditorHelperIntf, Standalone.EditorHelper;

procedure TStandaloneStateTests.GotoLocation_TellsTheFormWhereToGo;
// GotoLocation used to move only the active file. Nothing observed that,
// so the Memo kept the previous file's text while Ctrl+S wrote it to the
// new one. The form has to be told, with the 1-based position.
var
  State: TStandaloneProjectState;
  Helper: IEditorHelper;
  GotFile: string;
  GotLine, GotCol, GotLen, Calls: Integer;
begin
  Calls := 0;
  GotLine := 0;
  GotCol := 0;
  GotLen := 0;
  State := TStandaloneProjectState.Create;
  try
    State.OnNavigate :=
      procedure(const AFile: string; ALine, ACol, AHighlightLen: Integer)
      begin
        Inc(Calls);
        GotFile := AFile;
        GotLine := ALine;
        GotCol := ACol;
        GotLen := AHighlightLen;
      end;
    Helper := TStandaloneEditorHelper.Create(State);
    Assert.IsTrue(Helper.GotoLocation('C:\X\B.pas', 4, 2, 3));
    Assert.AreEqual<Integer>(1, Calls, 'the form is told exactly once');
    Assert.AreEqual('C:\X\B.pas', GotFile);
    Assert.AreEqual<Integer>(5, GotLine, 'LSP line 4 is editor line 5');
    Assert.AreEqual<Integer>(3, GotCol, 'LSP column 2 is editor column 3');
    Assert.AreEqual<Integer>(3, GotLen, 'the highlight length is passed on');
    Assert.AreEqual('C:\X\B.pas', State.ActiveFile);
  finally
    Helper := nil;
    State.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TStandaloneStateTests);

end.
