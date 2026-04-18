(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.PathManager;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.StrUtils,
  System.Win.Registry, Winapi.Windows,
  Xml.XMLIntf, Xml.XMLDoc, Xml.omnixmldom, Xml.Xmldom,
  DIH.Types, DIH.Logger, DIH.Placeholders;

type
  TDIHPathManager = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    function GetRegistryKey(APlatform: TDIHPlatform): string;
    function GetRegistryValueName(APathType: TDIHPathType): string;
    function GetEnvOptionsValueName(APathType: TDIHPathType): string;
    function GetEnvOptionsPath: string;
    function ReadCurrentPaths(const ARegKey, AValueName: string): TStringList;
    procedure WritePathList(const ARegKey, AValueName: string; APaths: TStringList);
    procedure UpdateEnvOptionsXml(const APlatform, AValueName, AContents: string);
    function ReadEnvOptionsPaths(const APlatform, AValueName: string): string;
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
    procedure AddPaths(const APaths: TArray<TDIHPathEntry>; APlatform: TDIHPlatform);
    procedure RemovePaths(const APaths: TArray<TDIHPathEntry>; APlatform: TDIHPlatform);
  end;

implementation

{ TDIHPathManager }

constructor TDIHPathManager.Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
begin
  inherited Create;
  FLogger := ALogger;
  FResolver := AResolver;
end;

function TDIHPathManager.GetRegistryKey(APlatform: TDIHPlatform): string;
begin
  Result := FResolver.Resolve('{#BDS}') + '\Library\' + APlatform.ToString;
end;

function TDIHPathManager.GetRegistryValueName(APathType: TDIHPathType): string;
begin
  case APathType of
    ptNamespace: Result := 'Namespace Search Path';
    ptSearch:    Result := 'Search Path';
    ptBrowsing:  Result := 'Browsing Path';
  end;
end;

function TDIHPathManager.GetEnvOptionsValueName(APathType: TDIHPathType): string;
begin
  case APathType of
    ptNamespace: Result := 'DelphiNamespaceSearchPath';
    ptSearch:    Result := 'DelphiLibraryPath';
    ptBrowsing:  Result := 'DelphiBrowsingPath';
  end;
end;

function TDIHPathManager.GetEnvOptionsPath: string;
var
  ProfileName, BDSVersion: string;
begin
  ProfileName := FResolver.Resolve('{#BDSProfileName}');
  BDSVersion := FResolver.Resolve('{#BDSVersion}');
  Result := TPath.Combine(TPath.GetHomePath,
    'Embarcadero' + PathDelim + ProfileName + PathDelim + BDSVersion + PathDelim + 'EnvOptions.proj');
end;

function TDIHPathManager.ReadCurrentPaths(const ARegKey, AValueName: string): TStringList;
var
  Reg: TRegistry;
  Value: string;
begin
  Result := TStringList.Create;
  Result.StrictDelimiter := True;
  Result.Delimiter := ';';
  Result.CaseSensitive := False;

  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly(ARegKey) then
    begin
      try
        if Reg.ValueExists(AValueName) then
        begin
          Value := Reg.ReadString(AValueName);
          Result.DelimitedText := Value;
        end;
      finally
        Reg.CloseKey;
      end;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TDIHPathManager.WritePathList(const ARegKey, AValueName: string; APaths: TStringList);
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey(ARegKey, True) then
    begin
      try
        Reg.WriteString(AValueName, APaths.DelimitedText);
      finally
        Reg.CloseKey;
      end;
    end
    else
      FLogger.Error('Failed to write registry key: %s', [ARegKey]);
  finally
    Reg.Free;
  end;
end;

function TDIHPathManager.ReadEnvOptionsPaths(const APlatform, AValueName: string): string;
var
  XmlFileName: string;
  XmlConfig: IXMLDocument;
  RootElement, CurrentChildNode, ValueNode: IXMLNode;
  I: Integer;
begin
  Result := '';
  XmlFileName := GetEnvOptionsPath;
  if not FileExists(XmlFileName) then
    Exit;

  DefaultDOMVendor := sOmniXmlVendor;
  XmlConfig := TXMLDocument.Create(nil);
  XmlConfig.LoadFromFile(XmlFileName);
  RootElement := XmlConfig.DocumentElement;

  for I := 0 to RootElement.ChildNodes.Count - 1 do
  begin
    CurrentChildNode := RootElement.ChildNodes[I];
    if (CurrentChildNode.NodeName = 'PropertyGroup')
      and (CurrentChildNode.Attributes['Condition'] = '''$(Platform)''==''' + APlatform + '''') then
    begin
      ValueNode := CurrentChildNode.ChildNodes.FindNode(AValueName);
      if Assigned(ValueNode) then
        Result := ValueNode.Text;
      Break;
    end;
  end;
end;

procedure TDIHPathManager.UpdateEnvOptionsXml(const APlatform, AValueName, AContents: string);
var
  XmlFileName, XmlData: string;
  XmlConfig: IXMLDocument;
  RootElement, CurrentChildNode: IXMLNode;
  I: Integer;
begin
  XmlFileName := GetEnvOptionsPath;
  if not FileExists(XmlFileName) then
  begin
    FLogger.Warning('EnvOptions.proj not found: %s', [XmlFileName]);
    Exit;
  end;

  DefaultDOMVendor := sOmniXmlVendor;
  XmlConfig := TXMLDocument.Create(nil);
  XmlConfig.LoadFromFile(XmlFileName);
  RootElement := XmlConfig.DocumentElement;

  for I := 0 to RootElement.ChildNodes.Count - 1 do
  begin
    CurrentChildNode := RootElement.ChildNodes[I];
    if (CurrentChildNode.NodeName = 'PropertyGroup')
      and (CurrentChildNode.Attributes['Condition'] = '''$(Platform)''==''' + APlatform + '''') then
    begin
      CurrentChildNode.ChildValues[AValueName] := AContents;
      XmlConfig.SaveToXML(XmlData);
      TFile.WriteAllText(XmlFileName, ReplaceStr(FormatXMLData(XmlData), '&apos;', ''''));
      FLogger.Detail('EnvOptions.proj updated: %s [%s]', [AValueName, APlatform]);
      Exit;
    end;
  end;

  FLogger.Warning('Platform "%s" not found in EnvOptions.proj', [APlatform]);
end;

procedure TDIHPathManager.AddPaths(const APaths: TArray<TDIHPathEntry>; APlatform: TDIHPlatform);
var
  PathEntry: TDIHPathEntry;
  RegKey, RegValueName, EnvValueName, ResolvedPath, PlatStr: string;
  CurrentPaths: TStringList;
  EnvCurrentStr: string;
  EnvPaths: TStringList;
  Modified: Boolean;
begin
  PlatStr := APlatform.ToString;
  RegKey := GetRegistryKey(APlatform);

  for PathEntry in APaths do
  begin
    if not (APlatform in PathEntry.Platforms) then
      Continue;

    RegValueName := GetRegistryValueName(PathEntry.PathType);
    EnvValueName := GetEnvOptionsValueName(PathEntry.PathType);
    ResolvedPath := FResolver.Resolve(PathEntry.Path);

    // 1. Update Registry
    CurrentPaths := ReadCurrentPaths(RegKey, RegValueName);
    try
      if CurrentPaths.IndexOf(ResolvedPath) = -1 then
      begin
        CurrentPaths.Add(ResolvedPath);
        WritePathList(RegKey, RegValueName, CurrentPaths);
        FLogger.Detail('Registry: Added %s path: %s', [RegValueName, ResolvedPath]);
      end
      else
        FLogger.Detail('Registry: Path already exists in %s: %s', [RegValueName, ResolvedPath]);
    finally
      CurrentPaths.Free;
    end;

    // 2. Update EnvOptions.proj
    EnvCurrentStr := ReadEnvOptionsPaths(PlatStr, EnvValueName);
    EnvPaths := TStringList.Create;
    try
      EnvPaths.StrictDelimiter := True;
      EnvPaths.Delimiter := ';';
      EnvPaths.CaseSensitive := False;
      if not EnvCurrentStr.IsEmpty then
        EnvPaths.DelimitedText := EnvCurrentStr;

      Modified := False;
      if EnvPaths.IndexOf(ResolvedPath) = -1 then
      begin
        EnvPaths.Add(ResolvedPath);
        Modified := True;
      end;

      if Modified then
        UpdateEnvOptionsXml(PlatStr, EnvValueName, EnvPaths.DelimitedText);
    finally
      EnvPaths.Free;
    end;
  end;
end;

procedure TDIHPathManager.RemovePaths(const APaths: TArray<TDIHPathEntry>; APlatform: TDIHPlatform);
var
  PathEntry: TDIHPathEntry;
  RegKey, RegValueName, EnvValueName, ResolvedPath, PlatStr: string;
  CurrentPaths: TStringList;
  Idx: Integer;
  EnvCurrentStr: string;
  EnvPaths: TStringList;
  Modified: Boolean;
begin
  PlatStr := APlatform.ToString;
  RegKey := GetRegistryKey(APlatform);

  for PathEntry in APaths do
  begin
    if not (APlatform in PathEntry.Platforms) then
      Continue;

    RegValueName := GetRegistryValueName(PathEntry.PathType);
    EnvValueName := GetEnvOptionsValueName(PathEntry.PathType);
    ResolvedPath := FResolver.Resolve(PathEntry.Path);

    // 1. Update Registry
    CurrentPaths := ReadCurrentPaths(RegKey, RegValueName);
    try
      Idx := CurrentPaths.IndexOf(ResolvedPath);
      if Idx >= 0 then
      begin
        CurrentPaths.Delete(Idx);
        WritePathList(RegKey, RegValueName, CurrentPaths);
        FLogger.Detail('Registry: Removed %s path: %s', [RegValueName, ResolvedPath]);
      end;
    finally
      CurrentPaths.Free;
    end;

    // 2. Update EnvOptions.proj
    EnvCurrentStr := ReadEnvOptionsPaths(PlatStr, EnvValueName);
    EnvPaths := TStringList.Create;
    try
      EnvPaths.StrictDelimiter := True;
      EnvPaths.Delimiter := ';';
      EnvPaths.CaseSensitive := False;
      if not EnvCurrentStr.IsEmpty then
        EnvPaths.DelimitedText := EnvCurrentStr;

      Modified := False;
      Idx := EnvPaths.IndexOf(ResolvedPath);
      if Idx >= 0 then
      begin
        EnvPaths.Delete(Idx);
        Modified := True;
      end;

      if Modified then
        UpdateEnvOptionsXml(PlatStr, EnvValueName, EnvPaths.DelimitedText);
    finally
      EnvPaths.Free;
    end;
  end;
end;

end.
