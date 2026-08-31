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

interface

type
  TUnitAvailability = (
    uaInProject,      // the unit's .pas is part of the active project
    uaOnSearchPath,   // its directory is compile-reachable (project search
                      // path, IDE library path, or the project directory)
    uaDcuAvailable,   // a compiled .dcu sits on a compile-reachable path
                      // (RTL/VCL: sources on browsing, DCUs on the lib path)
    uaBrowsingOnly);  // ONLY the index can see it - the compiler cannot

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
  System.SysUtils, System.IOUtils,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.DialogHelper,
  Expert.IdeThemes;

function CheckUnitAvailability(const AUnitName, APasPath: string): TUnitAvailability;
var
  Dir, D: string;
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

function EnsureUnitAvailable(const AUnitName: string): Boolean;
var
  Snap: IUnitSnapshot;
  Path: string;
begin
  Result := True;
  if (Editor = nil) or (AUnitName = '') then Exit;
  Snap := TUnitIndex.Instance.Snapshot;
  if (Snap = nil) or not Snap.TryGetUnitPath(AUnitName, Path) then
    Exit;   // not indexed (dcu-only RTL etc.) - nothing to warn about
  if CheckUnitAvailability(AUnitName, Path) <> uaBrowsingOnly then
    Exit;

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
