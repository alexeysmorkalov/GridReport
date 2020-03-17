{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains the definitions of classes TvgrForm and TvgrDialogForm<br>
TvgrForm - base class of all forms in GridReport.<br>
TvgrDialogForm - base class of all dialog forms in GridReport. }
unit vgr_Form;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, Forms, Menus, Messages, SysUtils, Windows,

  vgr_IniStorage;

const
{Default name of GridReport ini file, that is used by GridReport to store settings.}
  sDefaultIniFileName = 'gr.ini';
{Default registry path, that is used by GridReport to store settings.}
  sDefaultRegistryPath = 'Software\GridReport';
{This is name of OS enviroment variable, that defines the full name of GridReport ini file (including the directory name).}
  sIniFileEnvVariableName = 'GR_INIFILENAME';

type

  /////////////////////////////////////////////////
  //
  // TPopupListAccess
  //
  /////////////////////////////////////////////////
  TPopupListAccess = class(TPopupList)
  end;

  /////////////////////////////////////////////////
  //
  // rvgrShortCutItem
  //
  /////////////////////////////////////////////////
{Internal structure, used to store information about shortcut.
Syntax:
  rvgrShortCutItem = record
    MenuItem: TMenuItem;
    ShortCut: TShortCut;
  end;}
  rvgrShortCutItem = record
    MenuItem: TMenuItem;
    ShortCut: TShortCut;
  end;
{Pointer to a rvgrShortCutItem structure.}
  pvgrShortCutItem = ^rvgrShortCutItem;

  /////////////////////////////////////////////////
  //
  // TvgrShortCutsList
  //
  /////////////////////////////////////////////////
{Internal class, used to store information about shortcuts within form.
List contains then pointers to rvgrShortCutItem structures.}
  TvgrShortCutsList = class(TList)
  public
{Saves the stortcuts to a list. 
Parameters:
  AForm - TForm
  AStartPopupIndex - Integer}
    procedure SaveShortCuts(AForm: TForm; AStartPopupIndex: Integer);
{Removes the shortcuts from controls of form.}
    procedure RemoveShortCuts;
{Restores the shortcuts to controls of form.}
    procedure RestoreShortCuts;
    destructor Destroy; override;
  end;

{Describes the method to store various settings of GridReport (forms positions, etc)
Items:
  vgrssmNone - The settings is not stored.
  vgrssmIniFile - The settings is stored in the ini file,
name of ini file is specified by the DefaultStoreSettingsIniFileName variable.
  vgrssmRegistry - The settings is stored in the registry,
name of registry path is specified by the DefaultStoreSettingsRegistryPath variable.}
  TvgrStoreSettingsMethod = (vgrssmNone, vgrssmIniFile, vgrssmRegistry);
  /////////////////////////////////////////////////
  //
  // TvgrForm
  //
  /////////////////////////////////////////////////
{Base class for all forms in GridReport.}
  TvgrForm = class(TForm)
  private
    FShortCuts: TvgrShortCutsList;
    FStartPopupIndex: Integer;
    FSettingsStorage: TvgrIniStorage;
    FOldOnClose: TCloseEvent;

    procedure WMEnterMenuLoop(var Msg: TWMEnterMenuLoop); message WM_ENTERMENULOOP;
    procedure WMExitMenuLoop(var Msg: TWMExitMenuLoop); message WM_EXITMENULOOP;
    procedure PopupListWndProc(var Message: TMessage);
    procedure OnInitPopupMenu;
    procedure OnUnInitPopupMenu;
    procedure CreateSettingsStorage;
  protected
    function GetFixupShortcuts: Boolean; virtual;
    function GetStoreSettingsMethod: TvgrStoreSettingsMethod; virtual;
    function GetStorageSection: string; virtual;
    procedure DoLocalize; virtual;

    procedure RestoreSettings; virtual;
    procedure SaveSettings; virtual;

    procedure OnCloseEvent(Sender: TObject; var Action: TCloseAction); virtual;

    procedure Loaded; override;

    property FixupShortcuts: Boolean read GetFixupShortcuts;
    property StoreSettingsMethod: TvgrStoreSettingsMethod read GetStoreSettingsMethod;
    property SettingsStorage: TvgrIniStorage read FSettingsStorage;
    property StorageSection: string read GetStorageSection;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrDialogForm
  //
  /////////////////////////////////////////////////
{Base class of all dialog forms in GridReport.}
  TvgrDialogForm = class(TvgrForm)
  protected
    function GetStoreSettingsMethod: TvgrStoreSettingsMethod; override;
  end;

var
{Specifies the method to store settings of GridReport. For example to disable store settings
set this variable to vgrssmNone.<br>
Default value - vgrssmRegistry.}
  DefaultStoreSettingsMethod: TvgrStoreSettingsMethod;
{Specifies the ini file name, to store settings of GridReport.
Default value of this variable calculated based on this rules:<br>
1. Searches the OS enviroment variable with name sIniFileEnvVariableName,
if variable exists and not empty - it is used.<br>
2. Searches the sDefaultIniFileName file
in application dir, system dir, windows dir, if file not found
creates it in application dir.}
  DefaultStoreSettingsIniFileName: string;
{Specifies the registry path to store settings of GridReport.<br>
Path is relative to HKEY_CURRENT_USER.}
  DefaultStoreSettingsRegistryPath: string;

  procedure BuildComponentsList(ARootComponent: TComponent; AList: TStrings; AValidClass: TClass);
  function TextToComponent(ARootComponent: TComponent; const AText: string; AValidClass: TClass): TComponent; overload;
  function TextToComponent(ARootComponent: TComponent; const AText: string; AValidClass: TClass; AComponents: TStrings): TComponent; overload;
  function ComponentToText(ARootComponent: TComponent; AComponent: TComponent; AComponents: TStrings): string; overload;
  function ComponentToText(ARootComponent: TComponent; AComponent: TComponent): string; overload;
  
implementation

uses
  vgr_FormLocalizer, vgr_Functions;

procedure BuildComponentsList(ARootComponent: TComponent; AList: TStrings; AValidClass: TClass);
var
  I: Integer;
{$IFDEF VTK_D6_OR_D7}
  J: Integer;
{$ENDIF}

  procedure AddComponent(const AComponentName: string; AComponent: TComponent);
  begin
    if AComponentName = '' then
      exit;
    if (AValidClass <> nil) and not AComponent.InheritsFrom(AValidClass) then
      exit;
    if AList.IndexOfObject(AComponent) <> -1 then
      exit;
      
    AList.AddObject(AComponentName, AComponent);
  end;

  procedure AddComponents(AComponent: TComponent);
  var
    I: integer;
  begin
    AddComponent(AComponent.Name, AComponent);

    for I := 0 to AComponent.ComponentCount - 1 do
      if (AComponent = ARootComponent.Owner) or (AComponent = ARootComponent) then
        AddComponent(AComponent.Components[I].Name, AComponent.Components[I])
      else
        AddComponent(AComponent.Name + '.' + AComponent.Components[I].Name, AComponent.Components[I])
  end;
  
begin
  AddComponent(ARootComponent.Name, ARootComponent);
  if ARootComponent.Owner <> nil then
    AddComponents(ARootComponent.Owner);

{$IFDEF VTK_D6_OR_D7}
  if csDesigning in ARootComponent.ComponentState then
    begin
      for I := 0 to Screen.CustomFormCount - 1 do
        if AnsiCompareText(Screen.CustomForms[I].ClassName, 'TDataModuleForm') = 0 then
          for J := 0 to Screen.CustomForms[I].ComponentCount - 1 do
            if Screen.CustomForms[I].Components[J] is TDataModule then
              AddComponents(Screen.CustomForms[I].Components[J])
    end
  else
    begin
      for i:=0 to Screen.DataModuleCount-1 do
        AddComponents(Screen.DataModules[i]);
    end;
{$ELSE}
    for i:=0 to Screen.DataModuleCount-1 do
      AddComponents(Screen.DataModules[i]);
{$ENDIF}

  for I := 0 to Screen.DataModuleCount - 1 do
    AddComponents(Screen.DataModules[I]);

  for I := 0 to Screen.FormCount - 1 do
    AddComponents(Screen.Forms[I]);
end;

function TextToComponent(ARootComponent: TComponent; const AText: string; AValidClass: TClass; AComponents: TStrings): TComponent;
var
  I: Integer;
begin
  I := IndexInStrings(AText, AComponents);
  if (I = -1) or (not AComponents.Objects[I].InheritsFrom(AValidClass)) then
    Result := nil
  else
    Result := TComponent(AComponents.Objects[I]);
end; 

function TextToComponent(ARootComponent: TComponent; const AText: string; AValidClass: TClass): TComponent;
var
  AComponents: TStringList;
begin
  AComponents := TStringList.Create;
  try
    BuildComponentsList(ARootComponent, AComponents, AValidClass);
    Result := TextToComponent(ARootComponent, AText, AValidClass, AComponents)
  finally
    AComponents.Free;
  end;
end;

function ComponentToText(ARootComponent: TComponent; AComponent: TComponent; AComponents: TStrings): string;
var
  I: Integer;
begin
  if AComponent = nil then
    Result := ''
  else
  begin
    I := AComponents.IndexOfObject(AComponent);
    if I = -1 then
      Result := AComponent.Name
    else
      Result := AComponents[I];
  end;
end;

function ComponentToText(ARootComponent: TComponent; AComponent: TComponent): string;
var
  AComponents: TStringList;
begin
  AComponents := TStringList.Create;
  try
    BuildComponentsList(ARootComponent, AComponents, nil);
    ComponentToText(ARootComponent, AComponent, AComponents);
  finally
    AComponents.Free;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrShortCutsList
//
/////////////////////////////////////////////////
destructor TvgrShortCutsList.Destroy;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    FreeMem(Items[I]);
  inherited;
end;

procedure TvgrShortCutsList.SaveShortCuts(AForm: TForm; AStartPopupIndex: Integer);
var
  I, J: Integer;

  procedure AddShortCut(Item: TMenuItem);
  var
    I: Integer;
    P: pvgrShortCutItem;
  begin
    if Item.ShortCut <> 0 then
    begin
      GetMem(P, sizeof(rvgrShortCutItem));
      P.MenuItem := Item;
      P.ShortCut := Item.ShortCut;
      Add(P);
    end;
    for I := 0 to Item.Count - 1 do
      AddShortCut(Item.Items[I]);
  end;
  
begin
  if AForm.Menu <> nil then
    for I := 0 to AForm.Menu.Items.Count - 1 do
      AddShortCut(AForm.Menu.Items[I]);
  for I := AStartPopupIndex to PopupList.Count - 1 do
    with TPopupMenu(PopupList[I]) do
      for J := 0 to Count - 1 do
        AddShortCut(Items[J]);
end;

procedure TvgrShortCutsList.RemoveShortCuts;
var
  I: integer;
begin
  for I := 0 to Count - 1 do
    pvgrShortCutItem(Items[I]).MenuItem.ShortCut := 0;
end;

procedure TvgrShortCutsList.RestoreShortCuts;
var
  I: integer;
begin
  for I := 0 to Count - 1 do
    pvgrShortCutItem(Items[I]).MenuItem.ShortCut := pvgrShortCutItem(Items[I]).ShortCut;
end;

/////////////////////////////////////////////////
//
// TvgrForm
//
/////////////////////////////////////////////////
constructor TvgrForm.Create(AOwner: TComponent);
begin
  inherited;
  if FixupShortcuts then
    FShortCuts := TvgrShortCutsList.Create;
  CreateSettingsStorage;
  FStartPopupIndex := PopupList.Count;
end;

destructor TvgrForm.Destroy;
begin
  if FixupShortcuts then
    FreeAndNil(FShortCuts);
  if FSettingsStorage <> nil then
    FreeAndNil(FSettingsStorage);
  inherited;
end;

function TvgrForm.GetFixupShortcuts: Boolean;
begin
  Result := False;
end;

function TvgrForm.GetStoreSettingsMethod: TvgrStoreSettingsMethod;
begin
  Result := vgrssmNone;
end;

function TvgrForm.GetStorageSection: string;
begin
  Result := ClassName;
end;

procedure TvgrForm.RestoreSettings;
var
  AWindowState: TWindowState;
begin
  AWindowState := TWindowState(SettingsStorage.ReadInteger(StorageSection, 'WindowState', Integer(wsNormal)));
  if AWindowState <> wsMinimized then
  begin
    if AWindowState = wsNormal then
    begin
      Left := SettingsStorage.ReadInteger(StorageSection, 'Left', Left);
      Top := SettingsStorage.ReadInteger(StorageSection, 'Top', Top);
      if BorderStyle in [bsSizeToolWin, bsSizeable] then
      begin
        Width := SettingsStorage.ReadInteger(StorageSection, 'Width', Width);
        Height := SettingsStorage.ReadInteger(StorageSection, 'Height', Height);
      end
    end
    else
      WindowState := AWindowState;
  end;
end;

procedure TvgrForm.SaveSettings;
begin
  SettingsStorage.WriteInteger(StorageSection, 'WindowState', Integer(WindowState));
  SettingsStorage.WriteInteger(StorageSection, 'Left', Left);
  SettingsStorage.WriteInteger(StorageSection, 'Top', Top);
  if BorderStyle in [bsSizeToolWin,bsSizeable] then
  begin
    SettingsStorage.WriteInteger(StorageSection, 'Width', Width);
    SettingsStorage.WriteInteger(StorageSection, 'Height', Height);
  end
end;

procedure TvgrForm.CreateSettingsStorage;
begin
  case StoreSettingsMethod of
    vgrssmIniFile: FSettingsStorage := TvgrIniFileStorage.Create(DefaultStoreSettingsIniFileName);
    vgrssmRegistry: FSettingsStorage := TvgrRegistryStorage.Create(DefaultStoreSettingsRegistryPath);
  end;
end;

procedure TvgrForm.DoLocalize;
var
  I: Integer;
begin
  for I := 0 to ComponentCount - 1 do
    if Components[I] is TvgrFormLocalizer then
      TvgrFormLocalizer(Components[I]).DoLocalize({Self});
end;

procedure TvgrForm.Loaded;
begin
  inherited;
  DoLocalize;
  FOldOnClose := Self.OnClose;
  Self.OnClose := OnCloseEvent;
end;

procedure TvgrForm.AfterConstruction;
begin
  if SettingsStorage <> nil then
    RestoreSettings;
  if FixupShortcuts then
  begin
    FShortCuts.SaveShortCuts(Self, FStartPopupIndex);
    FShortCuts.RemoveShortCuts;
    with TPopupListAccess(PopupList) do
{$IFDEF VTK_D6_OR_D7}
      SetWindowLong(Window, GWL_WNDPROC, Longint(Classes.MakeObjectInstance(PopupListWndProc)));
{$ELSE}
      SetWindowLong(Window, GWL_WNDPROC, Longint(Forms.MakeObjectInstance(PopupListWndProc)));
{$ENDIF}
  end;

  //
  inherited;
end;

procedure TvgrForm.BeforeDestruction;
begin
  if SettingsStorage <> nil then
    SaveSettings;
  if FixupShortcuts then
  begin
    with TPopupListAccess(PopupList) do
{$IFDEF VTK_D6_OR_D7}
      SetWindowLong(Window, GWL_WNDPROC, Longint(Classes.MakeObjectInstance(WndProc)));
{$ELSE}
      SetWindowLong(Window, GWL_WNDPROC, Longint(Forms.MakeObjectInstance(WndProc)));
{$ENDIF}
  end;
  inherited;
end;

procedure TvgrForm.OnCloseEvent(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
  if Assigned(FOldOnClose) then
    FOldOnClose(Sender, Action);
end;

procedure TvgrForm.WMEnterMenuLoop(var Msg: TWMEnterMenuLoop);
begin
  if FixupShortCuts and (FShortCuts <> nil) then
    FShortCuts.RestoreShortCuts;
end;

procedure TvgrForm.WMExitMenuLoop(var Msg: TWMExitMenuLoop);
begin
  if FixupShortCuts and (FShortCuts <> nil) then
    FShortCuts.RemoveShortCuts;
end;

procedure TvgrForm.PopupListWndProc(var Message: TMessage);
begin
  case Message.Msg of
    WM_INITMENUPOPUP: OnInitPopupMenu;
    WM_UNINITMENUPOPUP: OnUnInitPopupMenu;
  end;
  TPopupListAccess(PopupList).WndProc(Message);
end;

procedure TvgrForm.OnInitPopupMenu;
begin
  if FixupShortCuts and (FShortCuts <> nil) then
    FShortCuts.RestoreShortCuts;
end;

procedure TvgrForm.OnUnInitPopupMenu;
begin
  if FixupShortCuts and (FShortCuts <> nil) then
    FShortCuts.RemoveShortCuts;
end;

/////////////////////////////////////////////////
//
// TvgrDialogForm
//
/////////////////////////////////////////////////
function TvgrDialogForm.GetStoreSettingsMethod: TvgrStoreSettingsMethod;
begin
  Result := DefaultStoreSettingsMethod;
end;

var
  buf: array [0..255] of char;
  S: string;

initialization

  DefaultStoreSettingsMethod := vgrssmRegistry; //vgrssmIniFile;

  // get ini file name for store settings
  if GetEnvironmentVariable(sIniFileEnvVariableName, buf, sizeof(buf)) > 0 then
  begin
    S := Trim(StrPas(buf));
    if S = '' then
      DefaultStoreSettingsIniFileName := GetFindFileName(sDefaultIniFileName);
  end
  else
    DefaultStoreSettingsIniFileName := GetFindFileName(sDefaultIniFileName);

  DefaultStoreSettingsRegistryPath := sDefaultRegistryPath;

end.

