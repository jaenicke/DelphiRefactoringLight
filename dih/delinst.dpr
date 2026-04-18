(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
program delinst;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  DIH.Types in 'DIH.Types.pas',
  DIH.Logger in 'DIH.Logger.pas',
  DIH.CommandLine in 'DIH.CommandLine.pas',
  DIH.Placeholders in 'DIH.Placeholders.pas',
  DIH.XmlConfig in 'DIH.XmlConfig.pas',
  DIH.Registry in 'DIH.Registry.pas',
  DIH.FileOps in 'DIH.FileOps.pas',
  DIH.PathManager in 'DIH.PathManager.pas',
  DIH.Builder in 'DIH.Builder.pas',
  DIH.Packages in 'DIH.Packages.pas',
  DIH.Engine in 'DIH.Engine.pas';

var
  Engine: TDIHEngine;
begin
  try
    Engine := TDIHEngine.Create;
    try
      ExitCode := Engine.Run;
    finally
      Engine.Free;
    end;
  except
    on E: Exception do
    begin
      Writeln(ErrOutput, 'Fatal error: ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
