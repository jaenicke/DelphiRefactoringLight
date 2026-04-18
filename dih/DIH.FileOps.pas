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
    function CopyFiles(const AFiles: TArray<TDIHFileEntry>; APlatform: TDIHPlatform): Integer;
    function DeleteCopiedFiles(const AFiles: TArray<TDIHFileEntry>; APlatform: TDIHPlatform): Integer;
  end;

implementation

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
  TargetDir: string;
begin
  TargetDir := ExtractFilePath(ATarget);
  if not TDirectory.Exists(TargetDir) then
    TDirectory.CreateDirectory(TargetDir);

  TFile.Copy(ASource, ATarget, True);
  FLogger.Detail('Copied: %s -> %s', [ExtractFileName(ASource), ATarget]);
end;

function TDIHFileOperations.CopyFiles(const AFiles: TArray<TDIHFileEntry>; APlatform: TDIHPlatform): Integer;
var
  FileEntry: TDIHFileEntry;
  SourcePath, TargetDir, SourceDir, Mask: string;
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
              FLogger.Error('Failed to copy %s: %s', [FileName, E.Message]);
          end;
        end;
      end
      else
        FLogger.Warning('Source directory not found: %s', [SourceDir]);
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
            FLogger.Error('Failed to copy %s: %s', [SourcePath, E.Message]);
        end;
      end
      else
        FLogger.Warning('Source file not found: %s', [SourcePath]);
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
