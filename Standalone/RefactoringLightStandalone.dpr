program RefactoringLightStandalone;

uses
  Vcl.Forms,
  Standalone.MainForm in 'Standalone.MainForm.pas',
  Standalone.EditorHelper in 'Standalone.EditorHelper.pas',
  Expert.EditorHelperIntf in '..\Source\Expert.EditorHelperIntf.pas',
  Delphi.FileEncoding in '..\Source\Delphi.FileEncoding.pas',
  Expert.SemanticReplace in '..\Source\Expert.SemanticReplace.pas',
  Expert.SemanticReplaceDialogs in '..\Source\Expert.SemanticReplaceDialogs.pas',
  Expert.SemanticReplaceWizard in '..\Source\Expert.SemanticReplaceWizard.pas',
  Expert.RenameWizard in '..\Source\Expert.RenameWizard.pas',
  Expert.FindReferencesWizard in '..\Source\Expert.FindReferencesWizard.pas',
  Expert.FindImplementationsWizard in '..\Source\Expert.FindImplementationsWizard.pas',
  Expert.SignatureCheckWizard in '..\Source\Expert.SignatureCheckWizard.pas',
  Expert.WithRefactorWizard in '..\Source\Expert.WithRefactorWizard.pas',
  Expert.MoveToUnitWizard in '..\Source\Expert.MoveToUnitWizard.pas',
  Expert.UnitReferencesWizard in '..\Source\Expert.UnitReferencesWizard.pas',
  Expert.ExtractInterfaceWizard in '..\Source\Expert.ExtractInterfaceWizard.pas',
  Expert.ExtractMethod in '..\Source\Expert.ExtractMethod.pas',
  Expert.ExtractMethodDialog in '..\Source\Expert.ExtractMethodDialog.pas',
  Expert.CompletionWizard in '..\Source\Expert.CompletionWizard.pas',
  Expert.CompletionPopup in '..\Source\Expert.CompletionPopup.pas',
  Expert.SignatureHelpWizard in '..\Source\Expert.SignatureHelpWizard.pas',
  Expert.RenameDialog in '..\Source\Expert.RenameDialog.pas',
  Lsp.Protocol in '..\Source\Lsp.Protocol.pas',
  Expert.ExtractInterface in '..\Source\Expert.ExtractInterface.pas',
  Expert.ExtractInterfaceDialog in '..\Source\Expert.ExtractInterfaceDialog.pas',
  Expert.FindReferencesDialog in '..\Source\Expert.FindReferencesDialog.pas',
  Expert.IdentifierCheck in '..\Source\Expert.IdentifierCheck.pas',
  Expert.IdeThemes in '..\Source\Expert.IdeThemes.pas',
  Expert.ImplementationFinder in '..\Source\Expert.ImplementationFinder.pas',
  Expert.LspManager in '..\Source\Expert.LspManager.pas',
  Expert.MoveToUnit in '..\Source\Expert.MoveToUnit.pas',
  Expert.MoveToUnitDialog in '..\Source\Expert.MoveToUnitDialog.pas',
  Expert.RestartHint in '..\Source\Expert.RestartHint.pas',
  Expert.SelectionValidator in '..\Source\Expert.SelectionValidator.pas',
  Expert.SignatureCheck in '..\Source\Expert.SignatureCheck.pas',
  Expert.SignatureCheckDialog in '..\Source\Expert.SignatureCheckDialog.pas',
  Expert.UnitReferencesDialog in '..\Source\Expert.UnitReferencesDialog.pas',
  Expert.WithRefactorDialog in '..\Source\Expert.WithRefactorDialog.pas',
  Expert.WithRewriter in '..\Source\Expert.WithRewriter.pas',
  Expert.WithScanner in '..\Source\Expert.WithScanner.pas',
  Lsp.Client in '..\Source\Lsp.Client.pas',
  Lsp.JsonRpc in '..\Source\Lsp.JsonRpc.pas',
  Lsp.Uri in '..\Source\Lsp.Uri.pas',
  Rename.WorkspaceEdit in '..\Source\Rename.WorkspaceEdit.pas',
  Expert.DialogHelper in '..\Source\Expert.DialogHelper.pas';

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
