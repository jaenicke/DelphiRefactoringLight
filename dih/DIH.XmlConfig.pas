(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.XmlConfig;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Xml.XMLIntf, Xml.XMLDoc, Xml.omnixmldom, Xml.Xmldom,
  DIH.Types;

type
  TDIHXmlConfig = class
  private
    FEntries: TObjectList<TDIHEntry>;
    FBaseDir: string;
    procedure ParseEntry(ANode: IXMLNode);
    procedure ParsePaths(ANode: IXMLNode; AEntry: TDIHEntry);
    procedure ParsePathGroup(ANode: IXMLNode; APathType: TDIHPathType; AEntry: TDIHEntry);
    procedure ParseFiles(ANode: IXMLNode; AEntry: TDIHEntry);
    procedure ParsePackages(ANode: IXMLNode; AEntry: TDIHEntry);
    procedure ParseExperts(ANode: IXMLNode; AEntry: TDIHEntry);
    procedure ParseBuild(ANode: IXMLNode; AEntry: TDIHEntry);
    procedure ParseRegistry(ANode: IXMLNode; AEntry: TDIHEntry);
    procedure ParseEvents(ANode: IXMLNode; AEntry: TDIHEntry; AIsPost: Boolean);
    procedure LoadFromFile(const AFileName: string);
    function GetAttr(ANode: IXMLNode; const AName: string; const ADefault: string = ''): string;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Load(const AFileName: string);
    property Entries: TObjectList<TDIHEntry> read FEntries;
    property BaseDir: string read FBaseDir;
  end;

implementation

{ TDIHXmlConfig }

constructor TDIHXmlConfig.Create;
begin
  inherited;
  FEntries := TObjectList<TDIHEntry>.Create(True);
  DefaultDOMVendor := sOmniXmlVendor;
end;

destructor TDIHXmlConfig.Destroy;
begin
  FEntries.Free;
  inherited;
end;

function TDIHXmlConfig.GetAttr(ANode: IXMLNode; const AName: string; const ADefault: string): string;
var
  AttrNode: IXMLNode;
begin
  AttrNode := ANode.AttributeNodes.FindNode(AName);
  if Assigned(AttrNode) then
    Result := AttrNode.Text
  else
    Result := ADefault;
end;

procedure TDIHXmlConfig.Load(const AFileName: string);
begin
  FBaseDir := ExtractFilePath(ExpandFileName(AFileName));
  LoadFromFile(AFileName);
end;

procedure TDIHXmlConfig.LoadFromFile(const AFileName: string);
var
  Doc: IXMLDocument;
  Root, Node: IXMLNode;
  I: Integer;
begin
  Doc := TXMLDocument.Create(nil);
  Doc.LoadFromFile(AFileName);
  Doc.Active := True;

  Root := Doc.DocumentElement;
  if not SameText(Root.NodeName, 'dih') then
    raise Exception.CreateFmt('Invalid root element: expected "dih", got "%s"', [Root.NodeName]);

  for I := 0 to Root.ChildNodes.Count - 1 do
  begin
    Node := Root.ChildNodes[I];
    if SameText(Node.NodeName, 'entry') then
      ParseEntry(Node);
  end;
end;

procedure TDIHXmlConfig.ParseEntry(ANode: IXMLNode);
var
  Entry: TDIHEntry;
  SourceFile: string;
  I: Integer;
  Child: IXMLNode;
begin
  Entry := TDIHEntry.Create;
  Entry.Id := GetAttr(ANode, 'id');
  Entry.Description := GetAttr(ANode, 'description');

  // Check for source reference to external XML
  SourceFile := '';
  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'source') then
    begin
      SourceFile := Child.Text;
      if SourceFile.IsEmpty then
        SourceFile := GetAttr(Child, '', '');
      Break;
    end;
  end;

  // Also check for source= attribute syntax (as in the example XML)
  if SourceFile.IsEmpty then
    SourceFile := GetAttr(ANode, 'source', '');

  // Check child nodes for source element with attributes
  if SourceFile.IsEmpty then
  begin
    for I := 0 to ANode.ChildNodes.Count - 1 do
    begin
      Child := ANode.ChildNodes[I];
      if SameText(Child.NodeName, 'source') then
      begin
        SourceFile := Child.Text;
        if SourceFile.IsEmpty then
          SourceFile := GetAttr(Child, 'file', '');
        Break;
      end;
    end;
  end;

  if not SourceFile.IsEmpty then
  begin
    Entry.SourceFile := SourceFile;
    // Load from external file
    var FullPath := IncludeTrailingPathDelimiter(FBaseDir) + SourceFile;
    if FileExists(FullPath) then
      LoadFromFile(FullPath);
    Entry.Free;
    Exit;
  end;

  // Parse child elements
  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'paths') then
      ParsePaths(Child, Entry)
    else if SameText(Child.NodeName, 'files') then
      ParseFiles(Child, Entry)
    else if SameText(Child.NodeName, 'packages') then
      ParsePackages(Child, Entry)
    else if SameText(Child.NodeName, 'experts') then
      ParseExperts(Child, Entry)
    else if SameText(Child.NodeName, 'build') then
      ParseBuild(Child, Entry)
    else if SameText(Child.NodeName, 'registry') then
      ParseRegistry(Child, Entry)
    else if SameText(Child.NodeName, 'pre') then
      ParseEvents(Child, Entry, False)
    else if SameText(Child.NodeName, 'post') then
      ParseEvents(Child, Entry, True);
  end;

  FEntries.Add(Entry);
end;

procedure TDIHXmlConfig.ParsePaths(ANode: IXMLNode; AEntry: TDIHEntry);
var
  I: Integer;
  Child: IXMLNode;
begin
  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'namespace') then
      ParsePathGroup(Child, ptNamespace, AEntry)
    else if SameText(Child.NodeName, 'search') then
      ParsePathGroup(Child, ptSearch, AEntry)
    else if SameText(Child.NodeName, 'browsing') then
      ParsePathGroup(Child, ptBrowsing, AEntry);
  end;
end;

procedure TDIHXmlConfig.ParsePathGroup(ANode: IXMLNode; APathType: TDIHPathType; AEntry: TDIHEntry);
var
  Platforms: TDIHPlatforms;
  I: Integer;
  Child: IXMLNode;
  PathEntry: TDIHPathEntry;
begin
  Platforms := ParsePlatforms(GetAttr(ANode, 'platforms'));

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'path') or SameText(Child.NodeName, 'entry') then
    begin
      PathEntry.PathType := APathType;
      PathEntry.Path := Child.Text;
      PathEntry.Recursive := SameText(GetAttr(Child, 'recursive', 'False'), 'True');
      PathEntry.Platforms := Platforms;
      AEntry.Paths.Add(PathEntry);
    end;
  end;
end;

procedure TDIHXmlConfig.ParseFiles(ANode: IXMLNode; AEntry: TDIHEntry);
var
  Platforms: TDIHPlatforms;
  I: Integer;
  Child: IXMLNode;
  FileEntry: TDIHFileEntry;
begin
  Platforms := ParsePlatforms(GetAttr(ANode, 'platforms'));

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'file') then
    begin
      FileEntry.Source := Child.Text;
      FileEntry.Target := GetAttr(Child, 'target');
      FileEntry.Platforms := Platforms;
      AEntry.Files.Add(FileEntry);
    end;
  end;
end;

procedure TDIHXmlConfig.ParsePackages(ANode: IXMLNode; AEntry: TDIHEntry);
var
  Platforms: TDIHPlatforms;
  I: Integer;
  Child: IXMLNode;
  PkgEntry: TDIHPackageEntry;
begin
  Platforms := ParsePlatforms(GetAttr(ANode, 'platforms'));

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'package') then
    begin
      PkgEntry.BplPath := Child.Text;
      PkgEntry.Description := GetAttr(Child, 'description');
      PkgEntry.Platforms := Platforms;
      AEntry.Packages.Add(PkgEntry);
    end;
  end;
end;

procedure TDIHXmlConfig.ParseExperts(ANode: IXMLNode; AEntry: TDIHEntry);
var
  Platforms: TDIHPlatforms;
  I: Integer;
  Child: IXMLNode;
  Expert: TDIHExpertEntry;
begin
  Platforms := ParsePlatforms(GetAttr(ANode, 'platforms'));

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'expert') then
    begin
      Expert.BplPath := Child.Text;
      Expert.Name := GetAttr(Child, 'name');
      Expert.Platforms := Platforms;
      AEntry.Experts.Add(Expert);
    end;
  end;
end;

procedure TDIHXmlConfig.ParseBuild(ANode: IXMLNode; AEntry: TDIHEntry);
var
  Platforms: TDIHPlatforms;
  I: Integer;
  Child: IXMLNode;
  BuildProj: TDIHBuildProject;
  ExtraParams: string;
begin
  Platforms := ParsePlatforms(GetAttr(ANode, 'platforms'));
  ExtraParams := GetAttr(ANode, 'params');

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'project') then
    begin
      BuildProj.ProjectPath := Child.Text;
      BuildProj.Platforms := Platforms;
      BuildProj.ExtraParams := ExtraParams;
      // Allow per-project params override
      var ProjParams := GetAttr(Child, 'params');
      if not ProjParams.IsEmpty then
        BuildProj.ExtraParams := ProjParams;
      AEntry.BuildProjects.Add(BuildProj);
    end;
  end;
end;

procedure TDIHXmlConfig.ParseRegistry(ANode: IXMLNode; AEntry: TDIHEntry);
var
  RegPath: string;
  I: Integer;
  Child: IXMLNode;
  RegValue: TDIHRegistryValue;
begin
  RegPath := GetAttr(ANode, 'path');

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'value') then
    begin
      RegValue.Path := RegPath;
      RegValue.Name := GetAttr(Child, 'name');
      RegValue.Value := Child.Text;
      RegValue.ValueType := GetAttr(Child, 'type', 'String');
      AEntry.RegistryValues.Add(RegValue);
    end;
  end;
end;

procedure TDIHXmlConfig.ParseEvents(ANode: IXMLNode; AEntry: TDIHEntry; AIsPost: Boolean);
var
  Platforms: TDIHPlatforms;
  I: Integer;
  Child: IXMLNode;
  EventEntry: TDIHEventEntry;
begin
  Platforms := ParsePlatforms(GetAttr(ANode, 'platforms'));

  for I := 0 to ANode.ChildNodes.Count - 1 do
  begin
    Child := ANode.ChildNodes[I];
    if SameText(Child.NodeName, 'exec') then
    begin
      EventEntry.Command := Child.Text;
      EventEntry.Description := GetAttr(Child, 'description');
      EventEntry.Platforms := Platforms;
      if AIsPost then
        AEntry.PostEvents.Add(EventEntry)
      else
        AEntry.PreEvents.Add(EventEntry);
    end;
  end;
end;

end.
