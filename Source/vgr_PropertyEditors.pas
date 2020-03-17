{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains the property editors and the component editors, used by GridReport components at design time.}
unit vgr_PropertyEditors;

{$I vtk.inc}

interface

uses
//  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Messages, {$IFDEF VTK_D6_OR_D7} Types, {$ENDIF}
  Classes, Graphics, Menus, Controls, Forms, StdCtrls,
  {$IFDEF VTK_D6_OR_D7} VCLEditors, DesignIntf, DesignEditors, DesignMenus, {$ELSE} dsgnintf, {$ENDIF}
  ActnList, TypInfo,
  vgr_DataStorage, vgr_ScriptEdit, vgr_AliasManager, vgr_WorkbookDesigner, vgr_ReportDesigner, vgr_AliasManagerDesigner;

type

  /////////////////////////////////////////////////
  //
  // TvgrFixedFontNameProperty
  //
  /////////////////////////////////////////////////
{Represents the property editor for selecting name of fixed font.}
  TvgrFixedFontNameProperty = class(TFontNameProperty)
  public
{Provides the enumerated values of the property to a callback procedure.
Parameters:
  Proc - callback procedure.}
    procedure GetValues(Proc: TGetStrProc); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrComponentEditor
  //
  /////////////////////////////////////////////////
{Base class for components editors, adds one item into the context menu displaying version of GridReport.}
  TvgrComponentEditor = class(TComponentEditor)
  public
{Returns the string that corresponds to the specified position of the component’s context menu.
The form designer calls GetVerb iteratively to build the context menu that appears when the user
right-clicks the component.
TvgrWorkbookEditor support only one verb - "Edit"
Parameters:
  Index - index of the verb.
Return value:
  Returns name of the verb.
See also:
  GetVerbCount, ExecuteVerb}
    function GetVerb(Index: Integer): string; override;
{Returns the number of menu strings added to the context menu by the component editor.
The form designer calls GetVerbCount when the user right-clicks the component in order
to determine the number of menu items that should be added to the context menu.
It then adds the menu items by calling GetVerb for every value from 0 to GetVerbCount - 1.
TvgrWorkbookDesigner returns 1.
Return value:
  Count of items in the component popup menu, 1 in this case.
See also:
  GetVerb, ExecuteVerb.}
    function GetVerbCount: Integer; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWorkbookEditor
  //
  /////////////////////////////////////////////////
{Represents the component editor for editing of the workbook (TvgrWorkbook object).
This editor uses a TvgrWorkbookDesigner component.
See also:
  TvgrWorkbook, TvgrWorkbookDesigner}
  TvgrWorkbookEditor = class(TvgrComponentEditor)
  protected
    function GetDesignerClass: TvgrCustomWorkbookDesignerClass; virtual;
    function GetVerbEditCaption: string; virtual;
    procedure OnDesignerChange(Sender: TObject); virtual;
  public
{Responds when the user double-clicks the component in the form designer.
Creates instance of TvgrWorkbookDesigner and show designer as modal window.}
    procedure Edit; override;
{Performs the action of a specified verb.
The form designer calls ExecuteVerb when the user selects the verb that is returned by GetVerb(Index).
This verb is the string that appears in the corresponding position of the component’s context menu.
TvgrWorkbookEditor support only one verb and simple calls Edit method.
Parameters:
  Index - index of the verb.
See also:
  Edit}
    procedure ExecuteVerb(Index: Integer); override;
{Returns the string that corresponds to the specified position of the component’s context menu.
The form designer calls GetVerb iteratively to build the context menu that appears when the user
right-clicks the component.
TvgrWorkbookEditor support only one verb - "Edit"
Parameters:
  Index - index of the verb.
Return value:
  Returns name of the verb.
See also:
  GetVerbCount, ExecuteVerb}
    function GetVerb(Index: Integer): string; override;
{Returns the number of menu strings added to the context menu by the component editor.
The form designer calls GetVerbCount when the user right-clicks the component in order
to determine the number of menu items that should be added to the context menu.
It then adds the menu items by calling GetVerb for every value from 0 to GetVerbCount - 1.
TvgrWorkbookDesigner returns 1.
Return value:
  Count of items in the component popup menu, 1 in this case.
See also:
  GetVerb, ExecuteVerb.}
    function GetVerbCount: Integer; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrReportEditor
  //
  /////////////////////////////////////////////////
{Represents the component editor for editing of the report template (TvgrReportTemplate object).
This editor uses a TvgrReportDesigner component.
See also:
  TvgrReportTemplate, TvgrReportDesigner}
  TvgrReportEditor = class(TvgrWorkbookEditor)
  protected
    function GetDesignerClass: TvgrCustomWorkbookDesignerClass; override;
    function GetVerbEditCaption: string; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrAliasNodesProperty
  //
  /////////////////////////////////////////////////
{Represents the property editor for editing Items property of the TvgrAliasManager component.
See also:
  TvgrAliasManager, TvgrAliasManagerDesigner, TvgrAliasEditor}
  TvgrAliasNodesProperty = class(TClassProperty)
  public
{Responds when the user double-clicks the component in the form designer.
Creates instance of TvgrAliasManagerDesigner and show designer as modal window.}
    procedure Edit; override;
{Describes the property so the Object Inspector provides the appropriate controls.
Return value:
  Returns [paDialog].}
    function GetAttributes: TPropertyAttributes; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrAliasEditor
  //
  /////////////////////////////////////////////////
{Represents the component editor for editing contents of the alias manager (TvgrAliasManager component).
See also:
  TvgrAliasManager, TvgrAliasManagerDesigner, TvgrAliasNodesProperty}
  TvgrAliasEditor = class(TvgrComponentEditor)
  public
{Responds when the user double-clicks the component in the form designer.
Creates instance of TvgrAliasManagerDesigner and show designer as modal window.}
    procedure Edit; override;
{Performs the action of a specified verb.
The form designer calls ExecuteVerb when the user selects the verb that is returned by GetVerb(Index).
This verb is the string that appears in the corresponding position of the component’s context menu.
TvgrAliasEditor support only one verb and simple calls Edit method.
Parameters:
  Index - index of the verb.
See also:
  Edit}
    procedure ExecuteVerb(Index: Integer); override;
{Returns the string that corresponds to the specified position of the component’s context menu.
The form designer calls GetVerb iteratively to build the context menu that appears when the user
right-clicks the component. TvgrAliasEditor support only one verb - "Edit"
Parameters:
  Index - index of the verb.
Return value:
  Returns name of the verb.
See also:
  GetVerbCount, ExecuteVerb}
    function GetVerb(Index: Integer): string; override;
{Returns the number of menu strings added to the context menu by the component editor.
The form designer calls GetVerbCount when the user right-clicks the component in order
to determine the number of menu items that should be added to the context menu.
It then adds the menu items by calling GetVerb for every value from 0 to GetVerbCount - 1.
TvgrAliasEditor returns 1.
Return value:
  Count of items in the component popup menu, 1 in this case.
See also:
  GetVerb, ExecuteVerb.}
    function GetVerbCount: Integer; override;
  end;

implementation

uses
  vgr_StringIDs, vgr_Functions, vgr_Localize;

{$R ..\res\vgr_PropertyEditorsStrings.res}


/////////////////////////////////////////////////
//
// TvgrComponentEditor
//
/////////////////////////////////////////////////
function TvgrComponentEditor.GetVerb(Index: Integer): string;
begin
  case Index of
    0: Result := 'GridReport ' + vgrGridReportVersion;
    1: Result := '-';
    else Result := '';
  end;
end;

function TvgrComponentEditor.GetVerbCount: Integer;
begin
  Result := 2;
end;

/////////////////////////////////////////////////
//
// TvgrFixedFontNameProperty
//
/////////////////////////////////////////////////
procedure TvgrFixedFontNameProperty.GetValues(Proc: TGetStrProc);
var
  I : Integer;
  Font: TFont;
begin
  for I := 0 to Screen.Fonts.Count - 1 do
  begin
    Font := TFont.Create;
    Font.Name := Screen.Fonts[I];
    if Font.Pitch = fpFixed then
      Proc(Screen.Fonts[I]);
    Font.Free;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrWorkbookEditor
//
/////////////////////////////////////////////////
function TvgrWorkbookEditor.GetDesignerClass: TvgrCustomWorkbookDesignerClass;
begin
  Result := TvgrWorkbookDesigner;
end;

function TvgrWorkbookEditor.GetVerbEditCaption: string;
begin
  Result := vgrLoadStr(svgrid_vgr_PropertyEditors_EditWorkbook);
end;

procedure TvgrWorkbookEditor.OnDesignerChange(Sender: TObject);
begin
  Designer.Modified;
end;

procedure TvgrWorkbookEditor.Edit;
var
  AWorkbook : TvgrWorkbook;
  ADesigner: TvgrCustomWorkbookDesigner;
begin
  AWorkbook := Component as TvgrWorkbook;
  ADesigner := GetDesignerClass.Create(nil);
  ADesigner.OnChange := OnDesignerChange;
  try
    ADesigner.Workbook := AWorkbook;
    ADesigner.Design(True);
  finally
    ADesigner.Free;
  end;
end;

procedure TvgrWorkbookEditor.ExecuteVerb(Index: Integer);
begin
  case Index of
    2: Edit;
  end;
end;

function TvgrWorkbookEditor.GetVerb(Index: Integer): string;
begin
  case Index of
     2: Result := GetVerbEditCaption;
    else Result := inherited GetVerb(Index);
  end;
end;

function TvgrWorkbookEditor.GetVerbCount: Integer;
begin
  Result := 3;
end;

/////////////////////////////////////////////////
//
// TvgrReportEditor
//
/////////////////////////////////////////////////
function TvgrReportEditor.GetDesignerClass: TvgrCustomWorkbookDesignerClass;
begin
  Result := TvgrReportDesigner;
end;

function TvgrReportEditor.GetVerbEditCaption: string;
begin
  Result := vgrLoadStr(svgrid_vgr_PropertyEditors_EditTemplate);
end;

/////////////////////////////////////////////////
//
// TvgrAliasNodesProperty
//
/////////////////////////////////////////////////
procedure TvgrAliasNodesProperty.Edit;
var
  AAliasManager : TvgrAliasManager;
  ADesigner: TvgrAliasManagerDesignerDialog;
begin
  AAliasManager := TvgrAliasNodes(GetOrdValue).Owner as TvgrAliasManager;
  ADesigner := TvgrAliasManagerDesignerDialog.Create(nil);
  try
    ADesigner.AliasManager := AAliasManager;
    ADesigner.AvailableComponentsProvider.Root := AAliasManager.Owner;
    ADesigner.Execute;
    Modified;
  finally
    ADesigner.Free;
  end;
end;

function TvgrAliasNodesProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paDialog];
end;

/////////////////////////////////////////////////
//
// TvgrAliasEditor
//
/////////////////////////////////////////////////
procedure TvgrAliasEditor.Edit;
var
  AAliasManager : TvgrAliasManager;
  ADesigner: TvgrAliasManagerDesignerDialog;
begin
  AAliasManager := Component as TvgrAliasManager;
  ADesigner := TvgrAliasManagerDesignerDialog.Create(nil);
  try
    ADesigner.AliasManager := AAliasManager;
    ADesigner.AvailableComponentsProvider.Root := AAliasManager.Owner;
    ADesigner.Execute;
    if Designer <> nil then
      Designer.Modified;
  finally
    ADesigner.Free;
  end;
end;

procedure TvgrAliasEditor.ExecuteVerb(Index: Integer);
begin
  case Index of
    2: Edit;
  end;
end;

function TvgrAliasEditor.GetVerb(Index: Integer): string;
begin
  case Index of
    2: Result := vgrLoadStr(svgrid_vgr_PropertyEditors_AliasManager);
    else Result := inherited GetVerb(Index);
  end;
end;

function TvgrAliasEditor.GetVerbCount: Integer;
begin
  Result := 3;
end;

end.
