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
    [Test] procedure TheBudgetScalesWithTheProject;
    [Test] procedure StatusColoursAreReadableInBothThemes;
    /// <summary>The status window in the IDE's DARK theme was a patchwork
    ///  (report 2026-09-30): the theming service does not reach a frame's
    ///  children, so the item column stayed white while the custom-drawn cells
    ///  went dark. We paint the list ourselves now - body AND header.</summary>
    [Test] procedure ListHeaderIsVisibleAgainstItsOwnBody;
    [Test] procedure ThemedListOnlyRepaintsWhenSomethingChanged;
  end;

implementation

uses
  System.SysUtils, System.Math, System.JSON, Vcl.Graphics,
  Lsp.Client, Expert.IdeThemes;

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

procedure TLspReadinessTests.TheBudgetScalesWithTheProject;
begin
  // There is no honest single number: DelphiLsp loads the whole project
  // before it answers anything, so the bound has to scale with it. 120 s was
  // a guess from ONE measurement (566 files, ~56 s) - the user asked whether
  // it is always enough, and it is not.
  Assert.AreEqual(Cardinal(60000), LspReadinessBudgetMs(0), 'the floor');
  Assert.AreEqual(Cardinal(60000), LspReadinessBudgetMs(-5), 'nonsense counts as none');
  // 566 files: the reporter's project - comfortably above the ~56 s measured
  Assert.IsTrue(LspReadinessBudgetMs(566) > 150000, 'the reported project');
  // monotone, and capped so nothing waits forever
  Assert.IsTrue(LspReadinessBudgetMs(2000) > LspReadinessBudgetMs(566));
  Assert.AreEqual(Cardinal(600000), LspReadinessBudgetMs(100000), 'capped at 10 min');
  Assert.AreEqual(Cardinal(600000), LspReadinessBudgetMs(MaxInt), 'no overflow');
end;

procedure TLspReadinessTests.StatusColoursAreReadableInBothThemes;

  function Lum(C: TColor): Integer;
  var
    RGBVal: Cardinal;
  begin
    RGBVal := ColorToRGB(C);
    Result := (Integer(RGBVal and $FF) * 30 +
      Integer((RGBVal shr 8) and $FF) * 59 +
      Integer((RGBVal shr 16) and $FF) * 11) div 100;
  end;

const
  Light = TColor($00FFFFFF);
  Dark = TColor($001E1E1E);
begin
  // The status window says the LSP state in colour now. A fixed palette
  // would be unreadable in one of the two themes - the colours are derived
  // from the background, so assert the CONTRAST, not the values.
  for var L in [slGood, slWait, slBad] do
  begin
    Assert.IsTrue(Abs(Lum(StatusLevelColor(L, Light, clBlack)) - Lum(Light)) > 60,
      'readable on a light background');
    Assert.IsTrue(Abs(Lum(StatusLevelColor(L, Dark, clWhite)) - Lum(Dark)) > 60,
      'readable on a dark background');
    // and the two themes really get different colours
    Assert.AreNotEqual(StatusLevelColor(L, Light, clBlack),
      StatusLevelColor(L, Dark, clWhite));
  end;
  // a plain row keeps the theme's text colour
  Assert.AreEqual(clBlack, StatusLevelColor(slNeutral, Light, clBlack));
  Assert.AreEqual(clWhite, StatusLevelColor(slNeutral, Dark, clWhite));
end;

procedure TLspReadinessTests.ListHeaderIsVisibleAgainstItsOwnBody;

  function Lum(C: TColor): Integer;
  var
    RGBVal: Cardinal;
  begin
    RGBVal := ColorToRGB(C);
    Result := (Integer(RGBVal and $FF) * 30 +
      Integer((RGBVal shr 8) and $FF) * 59 +
      Integer((RGBVal shr 16) and $FF) * 11) div 100;
  end;

var
  Head, Line: TColor;
begin
  // The column header is a separate native control that no Color property
  // reaches - we paint it, so its colour must be derived from the BODY. A
  // theme whose button face equals its window colour would otherwise give a
  // header nobody can see.
  for var Back in [TColor($00FFFFFF), TColor($001E1E1E), TColor($00322F2D)] do
  begin
    ListHeaderColors(Back, Head, Line);
    Assert.IsTrue(Abs(Lum(Head) - Lum(Back)) >= 10,
      'the header must stand out from the rows');
    Assert.IsTrue(Abs(Lum(Line) - Lum(Head)) >= 15,
      'and its separator from the header');
    // a dark body gets a LIGHTER header, a light body a darker one
    if Lum(Back) < 128 then
      Assert.IsTrue(Lum(Head) > Lum(Back), 'lighter on a dark theme')
    else
      Assert.IsTrue(Lum(Head) < Lum(Back), 'darker on a light theme');
  end;
end;

procedure TLspReadinessTests.ThemedListOnlyRepaintsWhenSomethingChanged;
var
  LV: TThemedListView;
begin
  LV := TThemedListView.Create(nil);
  try
    Assert.IsFalse(LV.Themed, 'untouched until it is given colours');
    Assert.IsTrue(LV.ApplyColors(TColor($00322F2D), TColor($00E8E8E8)),
      'the first call changes something');
    Assert.AreEqual(TColor($00322F2D), LV.Color);
    // the native grid lines are a fixed light grey - a bright cage around
    // every cell is exactly what a dark theme must not have
    Assert.IsFalse(LV.GridLines, 'no grid lines on a dark body');
    // the poll tick calls this every few seconds: it must be a no-op then
    Assert.IsFalse(LV.ApplyColors(TColor($00322F2D), TColor($00E8E8E8)),
      'the same colours again must not repaint the window');
    Assert.IsTrue(LV.ApplyColors(clWindow, clWindowText), 'a theme switch does');
    Assert.IsTrue(LV.GridLines, 'and brings the grid lines back on a light body');
  finally
    LV.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TLspReadinessTests);

end.
