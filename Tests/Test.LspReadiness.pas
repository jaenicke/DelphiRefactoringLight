(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Test.LspReadiness;
// A forum report of 2026-09-30 came with the logs of two Find-References
// runs on the same position: the first, right after the IDE had started,
// listed 29 of 327 references (two of them wrong), the second one two
// minutes later was correct.
//
// The logs name the cause. Every readiness check the scan does ran on the
// MAIN session, while the answering was done by a FRESHLY STARTED agent
// session - and an agent pushes no diagnostics at all by design, so no
// counter could show that it was still loading the project. It answered
// null for about 55 seconds; the declaration query fell into that window,
// the caret was taken as the declaration, and every candidate the server
// later resolved CORRECTLY to the real declaration was dropped as
// "leads to another symbol".
//
// TLspClient.WaitUnitParsed is the signal that works without diagnostics:
// documentSymbol can only be answered once the unit is parsed. These tests
// script that answer, because the behaviour under test is precisely
// "not ready yet -> ready".

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TLspReadinessTests = class
  public
    [Test] procedure NotReady_WhileTheServerAnswersNothing;
    [Test] procedure Ready_AsSoonAsTheSymbolsArrive;
    [Test] procedure Cancel_EndsTheWaitImmediately;
    [Test] procedure AnEmptySymbolListIsNotReady;
  end;

implementation

uses
  System.SysUtils, System.JSON, Lsp.Client;

type
  /// <summary>Answers documentSymbol from a script: nothing for the first
  ///  AQuietCalls calls, then a one-entry symbol list - the shape of a
  ///  server that is loading the project and then becomes usable.</summary>
  TScriptedSymbolClient = class(TLspClient)
  private
    FQuietCalls: Integer;
    FCalls: Integer;
    FEmptyForever: Boolean;
  public
    constructor Create(AQuietCalls: Integer; AEmptyForever: Boolean = False); reintroduce;
    function GetDocumentSymbols(const AFilePath: string;
      ATimeoutMs: Cardinal = 60000): TJSONArray; override;
    property Calls: Integer read FCalls;
  end;

constructor TScriptedSymbolClient.Create(AQuietCalls: Integer; AEmptyForever: Boolean);
begin
  inherited Create('');      // no exe path - Start is never called
  FQuietCalls := AQuietCalls;
  FEmptyForever := AEmptyForever;
end;

function TScriptedSymbolClient.GetDocumentSymbols(const AFilePath: string;
  ATimeoutMs: Cardinal): TJSONArray;
begin
  Inc(FCalls);
  if FEmptyForever then
    Exit(TJSONArray.Create);          // answers, but knows no symbol
  if FCalls <= FQuietCalls then
    Exit(nil);                        // still loading: no answer at all
  Result := TJSONArray.Create;
  Result.Add(TJSONObject.Create.AddPair('name', 'TFormOptionSelection'));
end;

{ TLspReadinessTests }

procedure TLspReadinessTests.NotReady_WhileTheServerAnswersNothing;
var
  C: TScriptedSymbolClient;
begin
  C := TScriptedSymbolClient.Create(1000);   // never answers within the budget
  try
    // 1.2 s budget, 500 ms between attempts: the probe must give up and say
    // NOT ready - that is the answer the scan needs in order not to mistake
    // null answers for "no declaration"
    Assert.IsFalse(C.WaitUnitParsed('C:\x\uGlue.pas', 1200));
    Assert.IsTrue(C.Calls >= 2, 'it keeps asking while it waits');
  finally
    C.Free;
  end;
end;

procedure TLspReadinessTests.Ready_AsSoonAsTheSymbolsArrive;
var
  C: TScriptedSymbolClient;
begin
  C := TScriptedSymbolClient.Create(2);      // ready on the third attempt
  try
    Assert.IsTrue(C.WaitUnitParsed('C:\x\uGlue.pas', 30000));
    Assert.AreEqual(3, C.Calls, 'it stops asking once the answer arrives');
  finally
    C.Free;
  end;
end;

procedure TLspReadinessTests.Cancel_EndsTheWaitImmediately;
var
  C: TScriptedSymbolClient;
  Polls: Integer;
begin
  C := TScriptedSymbolClient.Create(1000);
  Polls := 0;
  try
    // the window's close is a cancel: the wait must end at the first poll,
    // not sit out its budget (a scan that keeps the IDE busy after its
    // window is gone was a tester report of its own)
    Assert.IsFalse(C.WaitUnitParsed('C:\x\uGlue.pas', 60000,
      function: Boolean
      begin
        Inc(Polls);
        Result := False;
      end));
    Assert.AreEqual(1, Polls);
    Assert.AreEqual(1, C.Calls);
  finally
    C.Free;
  end;
end;

procedure TLspReadinessTests.AnEmptySymbolListIsNotReady;
var
  C: TScriptedSymbolClient;
begin
  // An empty result is what the server sends for a unit it has not parsed;
  // treating it as ready would put us back where the report started.
  C := TScriptedSymbolClient.Create(0, True);
  try
    Assert.IsFalse(C.WaitUnitParsed('C:\x\uGlue.pas', 1200));
  finally
    C.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TLspReadinessTests);

end.
