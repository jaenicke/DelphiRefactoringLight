(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnitAvailability;

// The identifier index deliberately scans the IDE's BROWSING paths too
// (RTL/VCL sources live there) - but the compiler never looks at browsing
// paths. Adding a uses entry for a unit that is only reachable via a
// browsing path would produce F2613 at the next compile. This unit
// detects that case before the uses edit and offers the two remedies:
// add the unit file to the project, or add its directory to the
// project's unit search path.
//
// NEITHER OF THOSE IS THE REMEDY for a unit that is shipped COMPILED inside
// a package - DesignIntf and the rest of source\ToolsAPI are in
// designide.dcp and have no .dcu at all (measured, see
// PackageShippingUnit). There the way in is the package: a package
// project names it in its REQUIRES clause, and nothing else is shipped to
// do it with. So the question "is it reachable" is asked of the requires
// clause as well, and a project that already requires the package is not
// warned at all - which is what the report was about: the warning fired
// although the .dpk said "requires designide".

interface

type
  TUnitAvailability = (
    uaInProject,      // the unit's .pas is part of the active project
    uaOnSearchPath,   // its directory is compile-reachable (project search
                      // path, IDE library path, or the project directory)
    uaDcuAvailable,   // a compiled .dcu sits on a compile-reachable path
                      // (RTL/VCL: sources on browsing, DCUs on the lib path)
    uaViaPackage,     // shipped inside a PACKAGE the project already
                      // requires - reachable, and nothing to warn about
    uaPackageMissing, // shipped inside a package the project does NOT
                      // require: the requires clause is the only way in
    uaBrowsingOnly);  // ONLY the index can see it - the compiler cannot

const
  /// <summary>The verdicts the compiler cannot follow. A caller that gates
  ///  an "add unit" edit asks for THIS, never for one value: a unit shipped
  ///  inside a package is just as unreachable as a browsing-only one, and
  ///  it was the other half of this set that was missing from the answer
  ///  (user report, 2026-10-28).</summary>
  UnitNotReachable = [uaPackageMissing, uaBrowsingOnly];

/// <summary>Classifies how the compiler can reach the unit declared in
///  APasPath. AUnitName is the (dotted) unit name for the .dcu probe.</summary>
function CheckUnitAvailability(const AUnitName, APasPath: string): TUnitAvailability;

/// <summary>Gate for every "add unit to uses" action: when the unit is
///  reachable only via browsing paths, asks the user how to make it
///  available (add file to project / add dir to search path / just add
///  the uses entry). True = proceed with the uses edit; False = the user
///  cancelled. Units the index does not know pass through unasked.</summary>
function EnsureUnitAvailable(const AUnitName: string): Boolean;

implementation

uses
  System.SysUtils, System.IOUtils, System.Classes,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls,
  Delphi.FileEncoding,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.UsesEditor,
  Expert.DialogHelper, Expert.IdeThemes;

// The active project's PACKAGE source, or '' when it is no package project.
function ActivePackageSource: string;
var
  Dproj: string;
begin
  Result := '';
  if Editor = nil then Exit;
  Dproj := Editor.GetCurrentProjectDproj;
  if Dproj = '' then Exit;
  Result := ChangeFileExt(Dproj, '.dpk');
  if not FileExists(Result) then Result := '';
end;

// Its text - the BUFFER first, because a requires clause the user has just
// added is the answer, and only on the main thread: ToolsAPI is main-thread
// only, and this runs from an MCP handler too (the same rule
// Lsp.Client.ReadLiveContent follows).
function ReadPackageSource(const APath: string; out AText: string): Boolean;
begin
  AText := '';
  if APath = '' then Exit(False);
  if (Editor <> nil) and (TThread.CurrentThread.ThreadID = MainThreadID)
    and Editor.ReadEditorContent(APath, AText) then Exit(True);
  try
    AText := TDelphiFileEncoding.ReadAll(APath);
    Result := True;
  except
    Result := False;
  end;
end;

function ActiveProjectRequires(const APackage: string): Boolean;
var
  Text: string;
begin
  Result := False;
  if APackage = '' then Exit;
  if not ReadPackageSource(ActivePackageSource, Text) then Exit;
  for var P in PackageRequiresNames(Text) do
    if SameText(P, APackage) then Exit(True);
end;

function CheckUnitAvailability(const AUnitName, APasPath: string): TUnitAvailability;
var
  Dir, D, Pkg: string;
  Dirs: TArray<string>;
begin
  Dir := ExcludeTrailingPathDelimiter(ExtractFilePath(APasPath));

  if Editor <> nil then
    for var F in Editor.GetProjectSourceFiles do
      if SameText(F, APasPath) then
        Exit(uaInProject);

  Dirs := GatherCompileSearchDirs;
  for D in Dirs do
    if SameText(D, Dir) then
      Exit(uaOnSearchPath);
  for D in Dirs do
    if FileExists(IncludeTrailingPathDelimiter(D) + AUnitName + '.dcu') then
      Exit(uaDcuAvailable);

  // No .dcu and no directory: for a unit that is shipped inside a package
  // that is not a defect of the project's paths, and the paths cannot cure
  // it either.
  Pkg := PackageShippingUnit(APasPath, FindBdsRoot);
  if Pkg <> '' then
  begin
    if ActiveProjectRequires(Pkg) then Exit(uaViaPackage);
    Exit(uaPackageMissing);
  end;

  Result := uaBrowsingOnly;
end;

type
  TAvailabilityChoice = (acAddFile, acAddPath, acUsesOnly, acCancel);

function AskOutsideProject(const AUnitName, APasPath: string): TAvailabilityChoice;
var
  Dlg: TThemedToolForm;
  Lbl: TLabel;

  function AddButton(const ACaption: string; ATop, AModal: Integer): TButton;
  begin
    Result := TButton.Create(Dlg);
    Result.Parent := Dlg;
    Result.SetBounds(16, ATop, 328, 27);
    Result.Caption := ACaption;
    Result.ModalResult := AModal;
  end;

begin
  Dlg := TThemedToolForm.CreateNew(nil);
  try
    Dlg.BorderStyle := bsDialog;
    Dlg.Caption := 'Unit outside project scope';
    Dlg.Position := poScreenCenter;
    Dlg.ClientWidth := 360;

    Lbl := TLabel.Create(Dlg);
    Lbl.Parent := Dlg;
    Lbl.AutoSize := False;
    Lbl.WordWrap := True;
    Lbl.ShowAccelChar := False;
    Lbl.SetBounds(16, 12, 328, 92);
    Lbl.Caption := Format(
      '"%s" was found via the IDE browsing path only:'#13#10'%s'#13#10#13#10 +
      'The compiler will NOT find it there (F2613). Make it available first:',
      [AUnitName, ExtractFilePath(APasPath)]);

    AddButton('Add unit file to the &project', 110, mrYes).Default := True;
    AddButton('Add directory to the project &search path', 144, mrNo);
    AddButton('Only add the &uses entry (I''ll handle it)', 178, mrIgnore);
    with AddButton('Cancel', 212, mrCancel) do Cancel := True;

    Dlg.ClientHeight := 250;
    EnableThemes(Dlg);
    PrepareDialog(Dlg, nil);
    case Dlg.ShowModal of
      mrYes:    Result := acAddFile;
      mrNo:     Result := acAddPath;
      mrIgnore: Result := acUsesOnly;
    else
      Result := acCancel;
    end;
  finally
    Dlg.Free;
  end;
end;

// A unit that only a package can bring in: the file and the path are not
// offered, because neither of them is a way in - adding the directory makes
// the compiler RECOMPILE a design-time unit into the project (measured:
// 2192 lines for DesignIntf), and adding the file does the same.
function AskPackageMissing(const AUnitName, APackage: string;
  AIsPackageProject: Boolean): Boolean;
var
  What: string;
begin
  if AIsPackageProject then
    What := Format('Add'#13#10'  requires %s;'#13#10'to the package source, then '
      + 'the uses entry compiles.', [APackage])
  else
    What := Format('Only a package can bring it in: this project would have '
      + 'to be built with the runtime package %s.', [APackage]);
  Result := AskThemedConfirm(
    Format('"%s" is shipped compiled inside the %s package - there is no '
      + '.dcu for it anywhere, so neither adding its file to the project nor '
      + 'adding its folder to the search path makes the compiler see it '
      + '(both would recompile a design-time unit into your project).'
      + #13#10#13#10 + '%s', [AUnitName, APackage, What]),
    'Add the uses entry');
end;

function EnsureUnitAvailable(const AUnitName: string): Boolean;
var
  Snap: IUnitSnapshot;
  Path, Pkg: string;
begin
  Result := True;
  if (Editor = nil) or (AUnitName = '') then Exit;
  Snap := TUnitIndex.Instance.Snapshot;
  if (Snap = nil) or not Snap.TryGetUnitPath(AUnitName, Path) then
    Exit;   // not indexed (dcu-only RTL etc.) - nothing to warn about
  var Avail := CheckUnitAvailability(AUnitName, Path);
  if not (Avail in UnitNotReachable) then
    Exit;   // reachable - including through a package the project requires
  if Avail = uaPackageMissing then
  begin
    Pkg := PackageShippingUnit(Path, FindBdsRoot);
    Exit(AskPackageMissing(AUnitName, Pkg, ActivePackageSource <> ''));
  end;

  case AskOutsideProject(AUnitName, Path) of
    acAddFile:
      if not Editor.AddFileToActiveProject(Path) then
        ShowThemedMessage('The unit could not be added to the project - ' +
          'add it manually (Project > Add to project).');
    acAddPath:
      if not Editor.AddProjectSearchPath(
           ExcludeTrailingPathDelimiter(ExtractFilePath(Path))) then
        ShowThemedMessage('The search path could not be modified - add'#13#10 +
          ExtractFilePath(Path) + #13#10'to the project''s unit search path manually.');
    acUsesOnly: ;
  else
    Result := False;   // cancelled - skip the uses edit as well
  end;
end;

end.
