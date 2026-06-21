unit Expert.IdeThemes;

interface

uses
  System.UITypes,
  Vcl.Controls,
  Vcl.Forms;

//function IdeThemesEnabled: Boolean;
//function IsDarkMode: Boolean;
function GetThemedColor(AColor: TColor): TColor;
procedure EnableThemes(AForm: TCustomForm);

implementation

uses
  {$IFNDEF STANDALONE_BUILD} ToolsApi, {$ENDIF}
  System.SysUtils;

{$IFDEF STANDALONE_BUILD}
// In standalone, theming is a no-op: the IDE theme service is not
// available and our VCL forms use their own colors.
function GetThemedColor(AColor: TColor): TColor;
begin Result := AColor; end;
procedure EnableThemes(AForm: TCustomForm);
begin end;
{$ELSE}

function IdeThemesEnabled: Boolean;
var
  Service: IOTAIDEThemingServices;
begin
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    Result := Service.IDEThemingEnabled
  end
  else
    Result := False;
end;

function IsDarkMode: Boolean;
var
  Service: IOTAIDEThemingServices;
begin
  Result := False;
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    if Service.IDEThemingEnabled then
    begin
      Result := Service.ActiveTheme.Contains('Dark', True);
    end;
  end;
end;

procedure EnableThemes(AForm: TCustomForm);
var
  Service: IOTAIDEThemingServices;
begin
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    if Service.IDEThemingEnabled then
    begin
      Service.RegisterFormClass(TCustomFormClass(AForm.ClassType));
      Service.ApplyTheme(AForm);
    end;
  end;
end;

function GetThemedColor(AColor: TColor): TColor;
var
  Service: IOTAIDEThemingServices;
begin
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    if (Service.IDEThemingEnabled) and Assigned(Service.StyleServices) then
      Result := Service.StyleServices.GetSystemColor(AColor)
    else
      Result := AColor;
  end
  else
    Result := AColor;
end;
{$ENDIF}

end.
