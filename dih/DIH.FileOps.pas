(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.FileOps;

interface

uses
  System.SysUtils, System.IOUtils, System.Types,
  DIH.Types, DIH.Logger, DIH.Placeholders;

type
  TDIHFileOperations = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    FBaseDir: string;
    procedure CopyFileWithLog(const ASource, ATarget: string);
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver; const ABaseDir: string);
    /// <summary>Copies the entries whose AfterBuild flag equals
    ///  AAfterBuild - inputs before the build, artefacts after it.</summary>
    function CopyFiles(const AFiles: TArray<TDIHFileEntry>; APlatform: TDIHPlatform;
      AAfterBuild: Boolean; out AFailed: Integer): Integer;
    function DeleteCopiedFiles(const AFiles: TArray<TDIHFileEntry>; APlatform: TDIHPlatform): Integer;
  end;

/// <summary>Copies ASource over ATarget, even when the target is a file
///  somebody is RUNNING: Windows refuses to overwrite it, but it does
///  allow RENAMING it, so the old one is moved aside under a name nothing
///  can hold yet and the copy is made in its place. AMovedAside receives
///  that name ('' when the target could simply be overwritten).
///
///  The aside name is UNIQUE on purpose. A fixed "&lt;name&gt;.old" is
///  held by the process that was started FROM it after the previous
///  install, and then the move has nowhere to go - which is how an
///  installer can report success four versions in a row without replacing
///  anything (Refactoring Light 1.22.1).
///
///  Raises when the copy fails for any other reason: an installer that
///  swallows that is worse than one that stops.</summary>
function ReplaceFilePossiblyInUse(const ASource, ATarget: string;
  out AMovedAside: string): Boolean;

implementation

function ReplaceFilePossiblyInUse(const ASource, ATarget: string;
  out AMovedAside: string): Boolean;
const
  MaxAsideTries = 20;
var
  Aside: string;
  Tries: Integer;
begin
  AMovedAside := '';
  if not TFile.Exists(ATarget) then
  begin
    TFile.Copy(ASource, ATarget, True);
    Exit(True);
  end;

  try
    TFile.Copy(ASource, ATarget, True);
    Exit(True);
  except
    // Not "any error is a lock": only try the rename when the target is
    // really still there. Everything else is re-raised below.
    on E: Exception do
      if not TFile.Exists(ATarget) then raise;
  end;

  // The name has to be FREE, and in practice free on the first try: the
  // previous install's aside copy is still there (a process was started
  // from it, so nothing can delete it). A timestamp differs per install,
  // the try counter within one - Random would need Randomize, and without
  // it every process walks the same sequence from the same first value.
  // DIGITS ONLY, deliberately: a consumer tells its own aside copies from
  // somebody else's file by the "<name>.<digits>.old" shape.
  Tries := 0;
  repeat
    Inc(Tries);
    Aside := Format('%s.%s%d.old',
      [ATarget, FormatDateTime('yyyymmddhhnnsszzz', Now), Tries]);
  until (not TFile.Exists(Aside)) or (Tries >= MaxAsideTries);

  // Saying so beats renaming over a file somebody may still be running.
  if TFile.Exists(Aside) then
    raise Exception.CreateFmt('Cannot move %s aside: no unused name after ' +
      '%d tries (last tried: %s)', [ATarget, Tries, Aside]);

  // A rename is allowed while the file is open; a delete is not, so the
  // aside copy is LEFT BEHIND rather than removed. It is evidence that
  // the install worked, and the next one picks another name.
  TFile.Move(ATarget, Aside);
  try
    TFile.Copy(ASource, ATarget, True);
  except
    // Put the old one back, or the caller is left with nothing at all.
    if not TFile.Exists(ATarget) then
      TFile.Move(Aside, ATarget);
    raise;
  end;
  AMovedAside := Aside;
  Result := True;
end;

{ TDIHFileOperations }

constructor TDIHFileOperations.Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver; const ABaseDir: string);
begin
  inherited Create;
  FLogger := ALogger;
  FResolver := AResolver;
  FBaseDir := ABaseDir;
end;

procedure TDIHFileOperations.CopyFileWithLog(const ASource, ATarget: string);
var
  TargetDir, MovedAside: string;
begin
  TargetDir := ExtractFilePath(ATarget);
  if not TDirectory.Exists(TargetDir) then
    TDirectory.CreateDirectory(TargetDir);

  ReplaceFilePossiblyInUse(ASource, ATarget, MovedAside);
  if MovedAside <> '' then
    FLogger.Detail('Copied: %s -> %s (the running one was moved aside as %s)',
      [ExtractFileName(ASource), ATarget, ExtractFileName(MovedAside)])
  else
    FLogger.Detail('Copied: %s -> %s', [ExtractFileName(ASource), ATarget]);
end;

function TDIHFileOperations.CopyFiles(const AFiles: TArray<TDIHFileEntry>;
  APlatform: TDIHPlatform; AAfterBuild: Boolean; out AFailed: Integer): Integer;
var
  FileEntry: TDIHFileEntry;
  SourcePath, TargetDir, SourceDir, Mask: string;
  Files: TStringDynArray;
  FileName: string;
begin
  Result := 0;
  AFailed := 0;
  for FileEntry in AFiles do
  begin
    if not (APlatform in FileEntry.Platforms) then
      Continue;
    // Inputs before the build, artefacts after it.
    if FileEntry.AfterBuild <> AAfterBuild then
      Continue;

    SourcePath := FResolver.Resolve(FileEntry.Source);
    TargetDir := FResolver.Resolve(FileEntry.Target);

    // Make source path absolute if relative
    if not TPath.IsPathRooted(SourcePath) then
      SourcePath := IncludeTrailingPathDelimiter(FBaseDir) + SourcePath;

    // Make target path absolute if relative
    if not TPath.IsPathRooted(TargetDir) then
      TargetDir := IncludeTrailingPathDelimiter(FBaseDir) + TargetDir;

    // Check for wildcards
    if SourcePath.Contains('*') or SourcePath.Contains('?') then
    begin
      SourceDir := ExtractFilePath(SourcePath);
      Mask := ExtractFileName(SourcePath);
      if TDirectory.Exists(SourceDir) then
      begin
        Files := TDirectory.GetFiles(SourceDir, Mask);
        for FileName in Files do
        begin
          try
            CopyFileWithLog(FileName, IncludeTrailingPathDelimiter(TargetDir) + ExtractFileName(FileName));
            Inc(Result);
          except
            on E: Exception do
            begin
              FLogger.Error('Failed to copy %s: %s', [FileName, E.Message]);
              Inc(AFailed);
            end;
          end;
        end;
      end
      else
      begin
        FLogger.Error('Source directory not found: %s', [SourceDir]);
        Inc(AFailed);
      end;
    end
    else
    begin
      // Single file
      if FileExists(SourcePath) then
      begin
        try
          if TDirectory.Exists(TargetDir) or (ExtractFileExt(TargetDir) = '') then
            CopyFileWithLog(SourcePath, IncludeTrailingPathDelimiter(TargetDir) + ExtractFileName(SourcePath))
          else
            CopyFileWithLog(SourcePath, TargetDir);
          Inc(Result);
        except
          on E: Exception do
          begin
            FLogger.Error('Failed to copy %s: %s', [SourcePath, E.Message]);
            Inc(AFailed);
          end;
        end;
      end
      else
      begin
        FLogger.Error('Source file not found: %s', [SourcePath]);
        Inc(AFailed);
      end;
    end;
  end;
end;

function TDIHFileOperations.DeleteCopiedFiles(const AFiles: TArray<TDIHFileEntry>; APlatform: TDIHPlatform): Integer;
var
  FileEntry: TDIHFileEntry;
  SourcePath, TargetDir, TargetFile, SourceDir, Mask: string;
  Files: TStringDynArray;
  FileName: string;
begin
  Result := 0;
  for FileEntry in AFiles do
  begin
    if not (APlatform in FileEntry.Platforms) then
      Continue;

    SourcePath := FResolver.Resolve(FileEntry.Source);
    TargetDir := FResolver.Resolve(FileEntry.Target);

    if not TPath.IsPathRooted(SourcePath) then
      SourcePath := IncludeTrailingPathDelimiter(FBaseDir) + SourcePath;
    if not TPath.IsPathRooted(TargetDir) then
      TargetDir := IncludeTrailingPathDelimiter(FBaseDir) + TargetDir;

    if SourcePath.Contains('*') or SourcePath.Contains('?') then
    begin
      SourceDir := ExtractFilePath(SourcePath);
      Mask := ExtractFileName(SourcePath);
      if TDirectory.Exists(SourceDir) then
      begin
        Files := TDirectory.GetFiles(SourceDir, Mask);
        for FileName in Files do
        begin
          TargetFile := IncludeTrailingPathDelimiter(TargetDir) + ExtractFileName(FileName);
          if FileExists(TargetFile) then
          begin
            try
              TFile.Delete(TargetFile);
              FLogger.Detail('Deleted: %s', [TargetFile]);
              Inc(Result);
            except
              on E: Exception do
                FLogger.Error('Failed to delete %s: %s', [TargetFile, E.Message]);
            end;
          end;
        end;
      end;
    end
    else
    begin
      TargetFile := IncludeTrailingPathDelimiter(TargetDir) + ExtractFileName(SourcePath);
      if FileExists(TargetFile) then
      begin
        try
          TFile.Delete(TargetFile);
          FLogger.Detail('Deleted: %s', [TargetFile]);
          Inc(Result);
        except
          on E: Exception do
            FLogger.Error('Failed to delete %s: %s', [TargetFile, E.Message]);
        end;
      end;
    end;
  end;
end;

end.
