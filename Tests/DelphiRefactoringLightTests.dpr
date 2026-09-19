/// <summary>
///  Headless DUnitX runner for the IDE-free layer of Delphi Refactoring
///  Light: scanners, parsers, the uses editor, the quick-fix providers,
///  file encoding, URI conversion. No IDE and no DelphiLSP needed.
///
///  The Test.WithScanner / WithRewriter / UnitIndexGenerics /
///  UsesEditorMinimalWrite / AutoImportGenericResolve / InheritedDefectFixes
///  fixtures were written by Ian Branch (issue #10, MPL-2.0);
///  Test.RegressionSuite ports cases from the project's console suite.
///
///  Build and run: run-tests.cmd (or open the .dproj). Exit code 0 = all
///  passed, 1 = failures, 2 = the runner itself failed.
/// </summary>
program DelphiRefactoringLightTests;

{$APPTYPE CONSOLE}
{$STRONGLINKTYPES ON}

uses
  System.SysUtils,
  DUnitX.Loggers.Console,
  DUnitX.Loggers.Xml.NUnit,
  DUnitX.TestFramework,
  Delphi.FileEncoding in '..\Source\Delphi.FileEncoding.pas',
  Lsp.Uri in '..\Source\Lsp.Uri.pas',
  Expert.WithScanner in '..\Source\Expert.WithScanner.pas',
  Expert.WithRewriter in '..\Source\Expert.WithRewriter.pas',
  Expert.UnitIndex in '..\Source\Expert.UnitIndex.pas',
  Expert.UsesEditor in '..\Source\Expert.UsesEditor.pas',
  Expert.AutoImport in '..\Source\Expert.AutoImport.pas',
  Expert.VcsBlame in '..\Source\Expert.VcsBlame.pas',
  Expert.WorkerLatch in '..\Source\Expert.WorkerLatch.pas',
  Rename.WorkspaceEdit in '..\Source\Rename.WorkspaceEdit.pas',
  Test.WithScanner in 'Test.WithScanner.pas',
  Test.WithRewriter in 'Test.WithRewriter.pas',
  Test.UnitIndexGenerics in 'Test.UnitIndexGenerics.pas',
  Test.UsesEditorMinimalWrite in 'Test.UsesEditorMinimalWrite.pas',
  Test.AutoImportGenericResolve in 'Test.AutoImportGenericResolve.pas',
  Test.InheritedDefectFixes in 'Test.InheritedDefectFixes.pas',
  Expert.SignatureCheck in '..\Source\Expert.SignatureCheck.pas',
  Expert.StatementRefactor in '..\Source\Expert.StatementRefactor.pas',
  Expert.SafeDeletePlan in '..\Source\Expert.SafeDeletePlan.pas',
  Expert.PascalScanner in '..\Source\Expert.PascalScanner.pas',
  Expert.IncludeExpansion in '..\Source\Expert.IncludeExpansion.pas',
  Expert.InterfaceLinks in '..\Source\Expert.InterfaceLinks.pas',
  Expert.PropertyConvert in '..\Source\Expert.PropertyConvert.pas',
  Test.UserIdeas13 in 'Test.UserIdeas13.pas',
  Test.RegressionSuite in 'Test.RegressionSuite.pas',
  Test.Issue11 in 'Test.Issue11.pas';

var
  Runner: ITestRunner;
  Results: IRunResults;
  ConsoleLogger: ITestLogger;
  NUnitLogger: ITestLogger;
begin
  try
    TDUnitX.CheckCommandLine;

    Runner := TDUnitX.CreateRunner;
    // Fixtures are registered explicitly in each test unit's initialization.
    // RTTI discovery would find them a second time and run every test twice.
    Runner.UseRTTI := False;
    Runner.FailsOnNoAsserts := True;

    ConsoleLogger := TDUnitXConsoleLogger.Create(True);
    Runner.AddLogger(ConsoleLogger);

    NUnitLogger := TDUnitXXMLNUnitFileLogger.Create(TDUnitX.Options.XMLOutputFile);
    Runner.AddLogger(NUnitLogger);

    Results := Runner.Execute;

    if Results.AllPassed then
      ExitCode := 0
    else
      ExitCode := 1;

    if TDUnitX.Options.ExitBehavior = TDUnitXExitBehavior.Pause then
    begin
      Write('Done - press <Enter> to quit.');
      Readln;
    end;
  except
    on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      ExitCode := 2;
    end;
  end;
end.
