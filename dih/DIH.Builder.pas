(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
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
    procedure CreateSingleProjectGroupProj(const AGroupProjPath, AProjectPath: string; APlatform: TDIHPlatform;
      const ABuildConfig, ADcuDir, ABplDir, ADcpDir: string);
    procedure ReadBdsErrFile(const AErrPath: string);
    procedure CleanupBdsTempFiles(const AGroupProjPath: string);
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

procedure TDIHBuilder.ReadBdsErrFile(const AErrPath: string);
var
  Lines: TStringList;
  Line: string;
begin
  if not FileExists(AErrPath) then
    Exit;

  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(AErrPath);
    for Line in Lines do
    begin
      if not Line.Trim.IsEmpty then
        FLogger.CompilerOutput(Line);
    end;
  finally
    Lines.Free;
  end;
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

function TDIHBuilder.BuildWithBds(const AProjects: TArray<TDIHBuildProject>; APlatform: TDIHPlatform;
  const ABuildConfig: string): Boolean;
var
  BdsExe, BdsProfile, Cmd, FullProjectPath, DcuDir, BplDir, DcpDir, GroupProjPath, ErrPath: string;
  Proj: TDIHBuildProject;
  I: Integer;
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

      // Auto-close dialogs: bds.exe may show a save dialog if it upgrades the .dproj ProjectVersion
      if ExecuteProcess(Cmd, True) <> 0 then
      begin
        ReadBdsErrFile(ErrPath);
        FLogger.FlushBuildOutputToConsole;
        FLogger.Error('Build failed (bds.exe): %s', [Proj.ProjectPath]);
        Result := False;
      end
      else
      begin
        ReadBdsErrFile(ErrPath);
        FLogger.Success('Build succeeded (bds.exe): %s', [ExtractFileName(Proj.ProjectPath)]);
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
