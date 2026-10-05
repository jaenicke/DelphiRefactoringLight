(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Builder;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, Winapi.Windows, Winapi.Messages, Winapi.CommCtrl,
  DIH.Types, DIH.Logger, DIH.Placeholders;

type
  /// <summary>What a line of bds.exe's .err file means.</summary>
  TBdsErrSeverity = (
    /// plain build output (command lines, "Erfolg", timings)
    besPlain,
    /// a bracketed IDE message without a compiler code
    besIdeMessage,
    /// a compiler ERROR or FATAL error - this build did not produce code
    besCompilerError);

/// <summary>The severity of one line of a bds.exe .err file. The IDE
///  writes the severity LOCALIZED ("[Fataler Fehler]" / "[Fatal Error]"),
///  so the word is useless - the COMPILER CODE is language-independent: a
///  letter class plus exactly four digits (F2063, E2003, W1036, H2164).
///  A bracketed line WITHOUT a code is an IDE message and must NEVER fail
///  an install: a design-time package the IDE loads at startup and cannot
///  find on disk reports exactly "[Fataler Fehler] Package X.bpl kann
///  nicht geladen werden", and the build that follows is what creates that
///  file (user log 2026-10-05). W and H codes do not fail a build either.
///  Exported so it can be checked against real .err output.</summary>
function BdsErrLineSeverity(const ALine: string): TBdsErrSeverity;

type
  TDIHBuilder = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    FBaseDir: string;
    FUseBds: Boolean;
    FRsVarsPath: string;
    function ExecuteProcess(const ACommand: string; AAutoCloseDialogs: Boolean = False): Integer;
    function ExecuteWithRsVars(const ACommand: string): Integer;
    function BuildWithMSBuild(const AProjectPath: string; APlatform: TDIHPlatform;
      const ABuildConfig, AExtraParams: string): Boolean;
    function BuildWithBds(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform;
      const ABuildConfig: string): Boolean;
    function BuildWithBdsProjects(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform;
      const ABuildConfig: string): Boolean;
    procedure SetEnv(const AName, AValue: string);
    procedure CreateSingleProjectGroupProj(const AGroupProjPath, AProjectPath: string; APlatform: TDIHPlatform;
      const ABuildConfig, ADcuDir, ABplDir, ADcpDir: string);
    function ReadBdsErrFile(const AErrPath: string): Boolean;
    procedure CleanupBdsTempFiles(const AGroupProjPath: string);
    function RelocateBdsArtifacts(const AProjectPath, ABplDir, ADcpDir: string;
      AStart: TDateTime): Boolean;
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver; const ABaseDir: string; AUseBds: Boolean);
    function Build(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform; const ABuildConfig: string): Boolean;
  end;

implementation

{ TDIHBuilder }

constructor TDIHBuilder.Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver; const ABaseDir: string; AUseBds: Boolean);
begin
  inherited Create;
  FLogger := ALogger;
  FResolver := AResolver;
  FBaseDir := ABaseDir;
  FUseBds := AUseBds;
  FRsVarsPath := FResolver.GetRsVarsPath;
end;

type
  TDialogCloserInfo = record
    ProcessId: DWORD;
    Running: Boolean;
  end;
  PDialogCloserInfo = ^TDialogCloserInfo;

function EnumWindowsCallback(Wnd: HWND; LParam: LPARAM): BOOL; stdcall;
var
  WndProcessId: DWORD;
  ClassName: array[0..255] of Char;
begin
  Result := True; // continue enumeration
  GetWindowThreadProcessId(Wnd, WndProcessId);
  if WndProcessId <> PDialogCloserInfo(LParam)^.ProcessId then
    Exit;

  // Check if this is a dialog window (#32770 is the Windows dialog class, used by both classic and Task Dialogs)
  GetClassName(Wnd, ClassName, Length(ClassName));
  if ClassName <> '#32770' then
    Exit;

  if not IsWindowVisible(Wnd) then
    Exit;

  // Try Task Dialog message first (TDM_CLICK_BUTTON), then fall back to WM_COMMAND.
  // Task Dialogs (DirectUIHWND/CtrlNotifySink) require TDM_CLICK_BUTTON = WM_USER + 102.
  PostMessage(Wnd, TDM_CLICK_BUTTON, IDNO, 0);
  Result := False; // stop enumeration
end;

function DialogCloserThread(Parameter: Pointer): Integer;
var
  Info: PDialogCloserInfo;
begin
  Result := 0;
  Info := PDialogCloserInfo(Parameter);
  while Info^.Running do
  begin
    EnumWindows(@EnumWindowsCallback, LPARAM(Info));
    Sleep(500);
  end;
end;

function TDIHBuilder.ExecuteProcess(const ACommand: string; AAutoCloseDialogs: Boolean): Integer;
var
  SI: TStartupInfo;
  PI: TProcessInformation;
  SA: TSecurityAttributes;
  ReadPipe, WritePipe: THandle;
  Buffer: array[0..4095] of AnsiChar;
  BytesRead: DWORD;
  ExitCode: DWORD;
  CmdLine: string;
  Output: AnsiString;
  Lines: TStringList;
  Line: string;
  CloserInfo: TDialogCloserInfo;
  CloserThread: THandle;
  ThreadId: DWORD;
begin
  Result := -1;

  SA.nLength := SizeOf(SA);
  SA.bInheritHandle := True;
  SA.lpSecurityDescriptor := nil;

  if not CreatePipe(ReadPipe, WritePipe, @SA, 0) then
    Exit;

  try
    ZeroMemory(@SI, SizeOf(SI));
    SI.cb := SizeOf(SI);
    SI.dwFlags := STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
    SI.hStdOutput := WritePipe;
    SI.hStdError := WritePipe;
    SI.wShowWindow := SW_HIDE;

    CmdLine := ACommand;
    UniqueString(CmdLine);

    if not CreateProcess(nil, PChar(CmdLine), nil, nil, True, CREATE_NO_WINDOW, nil, PChar(FBaseDir), SI, PI) then
    begin
      FLogger.Error('Failed to execute: %s (Error: %d)', [ACommand, GetLastError]);
      Exit;
    end;

    // Start dialog closer thread if requested (for bds.exe save dialogs)
    CloserThread := 0;
    if AAutoCloseDialogs then
    begin
      CloserInfo.ProcessId := PI.dwProcessId;
      CloserInfo.Running := True;
      CloserThread := BeginThread(nil, 0, @DialogCloserThread, @CloserInfo, 0, ThreadId);
    end;

    CloseHandle(WritePipe);
    WritePipe := 0;

    Output := '';
    while ReadFile(ReadPipe, Buffer, SizeOf(Buffer) - 1, BytesRead, nil) and (BytesRead > 0) do
    begin
      Buffer[BytesRead] := #0;
      Output := Output + Buffer;
    end;

    WaitForSingleObject(PI.hProcess, INFINITE);

    // Stop dialog closer thread
    if CloserThread <> 0 then
    begin
      CloserInfo.Running := False;
      WaitForSingleObject(CloserThread, 2000);
      CloseHandle(CloserThread);
    end;

    GetExitCodeProcess(PI.hProcess, ExitCode);
    Result := ExitCode;

    CloseHandle(PI.hProcess);
    CloseHandle(PI.hThread);

    // Log output via CompilerOutput (respects verbose settings)
    Lines := TStringList.Create;
    try
      Lines.Text := String(Output);
      for Line in Lines do
      begin
        if not Line.Trim.IsEmpty then
          FLogger.CompilerOutput(Line);
      end;
    finally
      Lines.Free;
    end;
  finally
    if ReadPipe <> 0 then
      CloseHandle(ReadPipe);
    if WritePipe <> 0 then
      CloseHandle(WritePipe);
  end;
end;

function TDIHBuilder.ExecuteWithRsVars(const ACommand: string): Integer;
var
  WrappedCmd: string;
begin
  if FileExists(FRsVarsPath) then
  begin
    WrappedCmd := Format('cmd.exe /c "call "%s" && %s"', [FRsVarsPath, ACommand]);
    FLogger.Detail('Using rsvars.bat: %s', [FRsVarsPath]);
  end
  else
  begin
    FLogger.Warning('rsvars.bat not found: %s - calling msbuild without it', [FRsVarsPath]);
    WrappedCmd := Format('cmd.exe /c "%s"', [ACommand]);
  end;
  Result := ExecuteProcess(WrappedCmd);
end;

function TDIHBuilder.BuildWithMSBuild(const AProjectPath: string; APlatform: TDIHPlatform;
  const ABuildConfig, AExtraParams: string): Boolean;
var
  Cmd, FullProjectPath, DcuDir, BplDir, DcpDir: string;
begin
  FullProjectPath := AProjectPath;
  if not TPath.IsPathRooted(FullProjectPath) then
    FullProjectPath := IncludeTrailingPathDelimiter(FBaseDir) + FullProjectPath;

  DcuDir := FResolver.Resolve('{#DcuTargetDir}');
  BplDir := FResolver.Resolve('{#BplTargetDir}');
  DcpDir := FResolver.Resolve('{#DcpTargetDir}');

  Cmd := Format('msbuild.exe "%s" /t:Build /p:Platform=%s /p:Config=%s /p:DCC_DcuOutput="%s" /p:DCC_BplOutput="%s" /p:DCC_DcpOutput="%s"',
    [FullProjectPath, APlatform.ToString, ABuildConfig, DcuDir, BplDir, DcpDir]);

  if not AExtraParams.IsEmpty then
    Cmd := Cmd + ' ' + AExtraParams;

  FLogger.Detail('Executing: %s', [Cmd]);
  Result := ExecuteWithRsVars(Cmd) = 0;
end;

procedure TDIHBuilder.CreateSingleProjectGroupProj(const AGroupProjPath, AProjectPath: string; APlatform: TDIHPlatform;
  const ABuildConfig, ADcuDir, ABplDir, ADcpDir: string);
var
  SL: TStringList;
  ProjName, Props: string;
begin
  ProjName := ChangeFileExt(ExtractFileName(AProjectPath), '');
  Props := Format('Platform=%s;Config=%s;DCC_DcuOutput=%s;DCC_BplOutput=%s;DCC_DcpOutput=%s',
    [APlatform.ToString, ABuildConfig, ADcuDir, ABplDir, ADcpDir]);

  SL := TStringList.Create;
  try
    SL.Add('<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">');
    SL.Add('    <PropertyGroup>');
    SL.Add('        <ProjectGuid>{D1E00002-0002-0002-0002-D1E000000002}</ProjectGuid>');
    SL.Add('    </PropertyGroup>');
    SL.Add('    <ItemGroup>');
    SL.Add(Format('        <Projects Include="%s">', [AProjectPath]));
    SL.Add('            <Dependencies/>');
    SL.Add('        </Projects>');
    SL.Add('    </ItemGroup>');
    SL.Add('    <ProjectExtensions>');
    SL.Add('        <Borland.Personality>Default.Personality.12</Borland.Personality>');
    SL.Add('        <Borland.ProjectType/>');
    SL.Add('        <BorlandProject>');
    SL.Add('            <Default.Personality/>');
    SL.Add('        </BorlandProject>');
    SL.Add('    </ProjectExtensions>');
    SL.Add(Format('    <Target Name="%s">', [ProjName]));
    SL.Add(Format('        <MSBuild Projects="%s" Properties="%s"/>', [AProjectPath, Props]));
    SL.Add('    </Target>');
    SL.Add(Format('    <Target Name="%s:Clean">', [ProjName]));
    SL.Add(Format('        <MSBuild Projects="%s" Targets="Clean"/>', [AProjectPath]));
    SL.Add('    </Target>');
    SL.Add(Format('    <Target Name="%s:Make">', [ProjName]));
    SL.Add(Format('        <MSBuild Projects="%s" Targets="Make" Properties="%s"/>', [AProjectPath, Props]));
    SL.Add('    </Target>');
    SL.Add('    <Target Name="Build">');
    SL.Add(Format('        <CallTarget Targets="%s"/>', [ProjName]));
    SL.Add('    </Target>');
    SL.Add('    <Target Name="Clean">');
    SL.Add(Format('        <CallTarget Targets="%s:Clean"/>', [ProjName]));
    SL.Add('    </Target>');
    SL.Add('    <Target Name="Make">');
    SL.Add(Format('        <CallTarget Targets="%s:Make"/>', [ProjName]));
    SL.Add('    </Target>');
    SL.Add('</Project>');
    SL.SaveToFile(AGroupProjPath, TEncoding.UTF8);
  finally
    SL.Free;
  end;
end;

function BdsErrLineSeverity(const ALine: string): TBdsErrSeverity;

  // A compiler code is a letter and EXACTLY four digits, standing on its
  // own - "-K00400000" or a path like "Win32\Release" must not match.
  function CodeLetter: Char;
  var
    I, D: Integer;
  begin
    Result := #0;
    for I := 1 to Length(ALine) do
    begin
      if not CharInSet(ALine[I], ['A'..'Z']) then Continue;
      if (I > 1) and (CharInSet(ALine[I - 1], ['A'..'Z', 'a'..'z', '0'..'9', '_', '-', '/'])) then
        Continue;
      D := 0;
      while (I + D + 1 <= Length(ALine)) and CharInSet(ALine[I + D + 1], ['0'..'9']) do
        Inc(D);
      if D <> 4 then Continue;
      if (I + 5 <= Length(ALine)) and
         CharInSet(ALine[I + 5], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then Continue;
      Exit(ALine[I]);
    end;
  end;

var
  T: string;
  L: Char;
begin
  Result := besPlain;
  T := ALine.TrimLeft;
  if not T.StartsWith('[') then Exit;
  if Pos(']', T) <= 1 then Exit;
  Result := besIdeMessage;
  L := CodeLetter;
  if CharInSet(L, ['E', 'F']) then
    Result := besCompilerError;
end;

// True when the output names a COMPILER error. bds.exe returns 0 even when
// its own output ends in a fatal error, so the exit code alone must not
// decide whether a build succeeded (user log 2026-10-05).
function TDIHBuilder.ReadBdsErrFile(const AErrPath: string): Boolean;
var
  Lines: TStringList;
  Line, FirstError, FirstNote: string;
  Notes: Integer;
begin
  Result := False;
  if not FileExists(AErrPath) then
    Exit;

  Notes := 0;
  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(AErrPath);
    for Line in Lines do
    begin
      if Line.Trim.IsEmpty then Continue;
      FLogger.CompilerOutput(Line);
      case BdsErrLineSeverity(Line) of
        besCompilerError:
          begin
            Result := True;
            if FirstError = '' then FirstError := Line.Trim;
          end;
        besIdeMessage:
          begin
            Inc(Notes);
            if FirstNote = '' then FirstNote := Line.Trim;
          end;
      end;
    end;
  finally
    Lines.Free;
  end;

  if Result then
    FLogger.Error('The IDE reported a compiler error: %s', [FirstError])
  else if Notes > 0 then
    FLogger.Warning('%d IDE message(s) during the build, no compiler code - ' +
      'not a compile error: %s', [Notes, FirstNote]);
end;

procedure TDIHBuilder.CleanupBdsTempFiles(const AGroupProjPath: string);
var
  BasePath, ErrPath, TvsPath: string;
begin
  // bds.exe creates additional files alongside the groupproj:
  //   _dih_temp.err              - build error output
  //   _dih_temp_prjgroup.tvsconfig - IDE tree view state
  BasePath := ChangeFileExt(AGroupProjPath, '');

  ErrPath := BasePath + '.err';
  TvsPath := BasePath + '_prjgroup.tvsconfig';

  if FileExists(AGroupProjPath) then
    System.SysUtils.DeleteFile(AGroupProjPath);
  if FileExists(ErrPath) then
    System.SysUtils.DeleteFile(ErrPath);
  if FileExists(TvsPath) then
    System.SysUtils.DeleteFile(TvsPath);
end;

function TDIHBuilder.RelocateBdsArtifacts(const AProjectPath, ABplDir, ADcpDir: string;
  AStart: TDateTime): Boolean;
// bds.exe builds the temp groupproj with the IDE's OWN project builder,
// which ignores the MSBuild property overrides (DCC_BplOutput etc.) baked
// into it - the artifacts land in the .dproj's own output directories
// (e.g. Packages\Output). The registered package path expects them under
// the BDS common Bpl/Dcp dirs, so without this step the IDE cannot load
// the package on installations without a command-line compiler (Delphi
// Community Edition - the only setups that take the bds.exe path at all).
var
  ProjectDir, Target, TargetDir, Mask: string;
  Kind: Integer;
begin
  Result := True;
  ProjectDir := ExtractFilePath(AProjectPath);
  if ProjectDir = '' then Exit;

  for Kind := 0 to 1 do
  begin
    if Kind = 0 then
    begin
      TargetDir := ABplDir;
      Mask := '*.bpl';
    end
    else
    begin
      TargetDir := ADcpDir;
      Mask := '*.dcp';
    end;
    if TargetDir.IsEmpty then Continue;
    TargetDir := IncludeTrailingPathDelimiter(TargetDir);
    try
      TDirectory.CreateDirectory(TargetDir);
      for var F in TDirectory.GetFiles(ProjectDir, Mask, TSearchOption.soAllDirectories) do
      begin
        // Only artifacts THIS build produced (2 min slack against
        // filesystem/clock granularity); anything already in the target
        // directory stays untouched.
        if TFile.GetLastWriteTime(F) < AStart - (2 / (24 * 60)) then Continue;
        if SameText(IncludeTrailingPathDelimiter(ExtractFilePath(F)), TargetDir) then Continue;
        Target := TargetDir + ExtractFileName(F);
        TFile.Copy(F, Target, True);
        FLogger.Detail('Relocated: %s -> %s', [F, Target]);
      end;
    except
      on E: Exception do
      begin
        FLogger.Error('Could not relocate build artifacts to %s: %s',
          [TargetDir, E.Message]);
        Result := False;
      end;
    end;
  end;
end;

// '' removes the variable, so the project's own default applies again.
procedure TDIHBuilder.SetEnv(const AName, AValue: string);
begin
  if AValue = '' then
    Winapi.Windows.SetEnvironmentVariable(PChar(AName), nil)
  else
    Winapi.Windows.SetEnvironmentVariable(PChar(AName), PChar(AValue));
end;

function TDIHBuilder.BuildWithBds(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform;
  const ABuildConfig: string): Boolean;
// MSBuild - and with it the IDE's OWN project builder - reads EVERY
// environment variable as a property; that is how the groupproj's
// DIH_ExeOutput reaches a bds build, while the Properties baked into that
// groupproj are ignored (see RelocateBdsArtifacts above). So an inherited
// "Config" or "Platform" decides what this build does, and nothing on the
// command line can correct it: a caller's `set CONFIG=<config file>` - the
// shape our own install.cmd had - turned DCC_DcuOutput into
// ".\Win32\E:\...\DelphiRefactoringLight.xml" and the build died in
// MakeDir (user report 2026-10-05), while a Visual Studio prompt exports
// Platform=x64, which no Delphi project has. The msbuild path passes both
// with /p: and is therefore immune. Pin them to what is really being
// built, and put the caller's environment back afterwards.
var
  OldConfig, OldPlatform: string;
begin
  OldConfig := System.SysUtils.GetEnvironmentVariable('Config');
  OldPlatform := System.SysUtils.GetEnvironmentVariable('Platform');
  if not SameText(OldConfig, ABuildConfig) and (OldConfig <> '') then
    FLogger.Detail('Inherited Config=%s replaced by %s for the bds build',
      [OldConfig, ABuildConfig]);
  SetEnv('Config', ABuildConfig);
  SetEnv('Platform', APlatform.ToString);
  try
    Result := BuildWithBdsProjects(AProjects, APlatform, ABuildConfig);
  finally
    SetEnv('Config', OldConfig);
    SetEnv('Platform', OldPlatform);
  end;
end;

function TDIHBuilder.BuildWithBdsProjects(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform;
  const ABuildConfig: string): Boolean;
var
  BdsExe, BdsProfile, Cmd, FullProjectPath, DcuDir, BplDir, DcpDir, GroupProjPath, ErrPath: string;
  Proj: TDIHBuildProject;
  I: Integer;
  BuildStart: TDateTime;
begin
  Result := True;
  BdsExe := IncludeTrailingPathDelimiter(FResolver.Resolve('{#BDSRootDir}')) + 'bin' + PathDelim + 'bds.exe';
  BdsProfile := FResolver.Resolve('{#BDSProfileName}');
  DcuDir := FResolver.Resolve('{#DcuTargetDir}');
  BplDir := FResolver.Resolve('{#BplTargetDir}');
  DcpDir := FResolver.Resolve('{#DcpTargetDir}');

  // Build each project individually so we can pass output directories
  for I := 0 to High(AProjects) do
  begin
    Proj := AProjects[I];

    FullProjectPath := Proj.ProjectPath;
    if not TPath.IsPathRooted(FullProjectPath) then
      FullProjectPath := IncludeTrailingPathDelimiter(FBaseDir) + FullProjectPath;

    // Create a single-project groupproj with the correct properties
    GroupProjPath := IncludeTrailingPathDelimiter(FBaseDir) + '_dih_temp.groupproj';
    ErrPath := IncludeTrailingPathDelimiter(FBaseDir) + '_dih_temp.err';
    CreateSingleProjectGroupProj(GroupProjPath, FullProjectPath, APlatform, ABuildConfig, DcuDir, BplDir, DcpDir);

    try
      FLogger.ClearBuildOutput;
      FLogger.Info('Building (bds.exe): %s', [ExtractFileName(Proj.ProjectPath)]);

      if not SameText(BdsProfile, 'BDS') then
        Cmd := Format('"%s" -b -ns -r %s "%s"', [BdsExe, BdsProfile, GroupProjPath])
      else
        Cmd := Format('"%s" -b -ns "%s"', [BdsExe, GroupProjPath]);
      FLogger.Detail('Executing: %s', [Cmd]);

      BuildStart := Now;
      // Auto-close dialogs: bds.exe may show a save dialog if it upgrades the .dproj ProjectVersion
      if ExecuteProcess(Cmd, True) <> 0 then
      begin
        ReadBdsErrFile(ErrPath);
        FLogger.FlushBuildOutputToConsole;
        FLogger.Error('Build failed (bds.exe): %s', [Proj.ProjectPath]);
        Result := False;
      end
      // bds.exe returns 0 even when its own output ends in a fatal error,
      // so the .err has the last word (user log 2026-10-05: an install
      // that reported OK right under "[Fataler Fehler]").
      else if ReadBdsErrFile(ErrPath) then
      begin
        FLogger.FlushBuildOutputToConsole;
        FLogger.Error('Build failed (bds.exe reported success, its output did ' +
          'not): %s', [Proj.ProjectPath]);
        Result := False;
      end
      else
      begin
        FLogger.Success('Build succeeded (bds.exe): %s', [ExtractFileName(Proj.ProjectPath)]);
        // bds.exe ignored the output-dir overrides - move the BPL/DCP to
        // the registered target dirs.
        if not RelocateBdsArtifacts(FullProjectPath, BplDir, DcpDir, BuildStart) then
          Result := False;
      end;
    finally
      CleanupBdsTempFiles(GroupProjPath);
    end;
  end;
end;

function TDIHBuilder.Build(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform; const ABuildConfig: string): Boolean;
var
  Proj: TDIHBuildProject;
  PlatformProjects: TArray<TDIHBuildProject>;
  Count: Integer;
begin
  Result := True;

  // Filter projects for current platform
  Count := 0;
  SetLength(PlatformProjects, Length(AProjects));
  for Proj in AProjects do
  begin
    if APlatform in Proj.Platforms then
    begin
      PlatformProjects[Count] := Proj;
      Inc(Count);
    end;
  end;
  SetLength(PlatformProjects, Count);

  if Count = 0 then
    Exit;

  if FUseBds then
    Result := BuildWithBds(PlatformProjects, APlatform, ABuildConfig)
  else
  begin
    for Proj in PlatformProjects do
    begin
      FLogger.ClearBuildOutput;
      FLogger.Info('Building: %s', [ExtractFileName(Proj.ProjectPath)]);
      if not BuildWithMSBuild(Proj.ProjectPath, APlatform, ABuildConfig, Proj.ExtraParams) then
      begin
        FLogger.FlushBuildOutputToConsole;
        FLogger.Error('Build failed: %s', [Proj.ProjectPath]);
        Result := False;
      end
      else
        FLogger.Success('Build succeeded: %s', [ExtractFileName(Proj.ProjectPath)]);
    end;
  end;
end;

end.
