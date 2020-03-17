{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

unit vgr_IDEAddon;

{$I vtk.inc}

interface

uses
//  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Dialogs, Classes, menus, ToolsAPI, SysUtils;

const
  sRootItemCaption = 'vtkTools';
  sRootItemName = 'vtkToolsMenu';

  sLocalizeFormMenuItemCaption = '&Localize form...';
  sLocalizeFormMenuItemName = 'LocalizeFormItem';

type
  /////////////////////////////////////////////////
  //
  // TvgrLocalizeExpert
  //
  /////////////////////////////////////////////////
  TvgrLocalizeExpert = class(TNotifierObject, IOTAWizard)
  private
    FMenuItem: TMenuItem;
    procedure OnMenuItemClick(Sender: TObject);
  protected
    { IOTANotifier }
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    { IOTAWizrd }
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    procedure Execute;
  public
    constructor Create;
    destructor Destroy; override;
  end;

var
  RootMenuItem: TMenuItem;

implementation

uses
  vgr_Functions, vgr_LocalizeExpertForm;

{$IFDEF VTK_D7}
procedure InitRootMenuItem;
var
  ANTAServices: INTAServices;
  AMainMenu: TMainMenu;
begin
  if RootMenuItem <> nil then
    exit;

  if Supports(BorlandIDEServices, INTAServices, ANTAServices) then
  begin
    AMainMenu := ANTAServices.MainMenu;
    RootMenuItem := TMenuItem.Create(AMainMenu);
    RootMenuItem.Caption := sRootItemCaption;
    RootMenuItem.Name := sRootItemName;
    AMainMenu.Items[4].Add(RootMenuItem);
  end;
end;

procedure DeInitRootMenuItem;
begin
  if RootMenuItem <> nil then
  begin
    RootMenuItem.Free;
    RootMenuItem := nil;
  end;
end;
{$ENDIF}

/////////////////////////////////////////////////
//
// TvgrLocalizeExpert
//
/////////////////////////////////////////////////
constructor TvgrLocalizeExpert.Create;
begin
  inherited Create;
  if RootMenuItem <> nil then
  begin
    FMenuItem := TMenuItem.Create(RootMenuItem.Owner);
    FMenuItem.Caption := sLocalizeFormMenuItemCaption;
    FMenuItem.Name := sLocalizeFormMenuItemName;
    FMenuItem.OnClick := OnMenuItemClick;
    RootMenuItem.Add(FMenuItem);
  end;
end;

destructor TvgrLocalizeExpert.Destroy;
begin
  if FMenuItem <> nil then
    FMenuItem.Free;
  inherited;
end;

procedure TvgrLocalizeExpert.AfterSave;
begin
end;

procedure TvgrLocalizeExpert.BeforeSave;
begin
end;

procedure TvgrLocalizeExpert.Destroyed;
begin
end;

procedure TvgrLocalizeExpert.Modified;
begin
end;

function TvgrLocalizeExpert.GetIDString: string;
begin
  Result := 'vtkTools.LocalizeExpert';
end;

function TvgrLocalizeExpert.GetName: string;
begin
  Result := 'vtkTools form localize expert';
end;

function TvgrLocalizeExpert.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

procedure TvgrLocalizeExpert.Execute;
begin
  TvgrLocalizeExpertForm.Create(nil).Execute;
end;

procedure TvgrLocalizeExpert.OnMenuItemClick(Sender: TObject);
begin
  Execute;
end;


initialization

{$IFDEF VTK_D7}
  InitRootMenuItem;
{$ENDIF}

finalization

{$IFDEF VTK_D7}
  DeInitRootMenuItem;
{$ENDIF}

end.

