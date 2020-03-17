{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains classes, that can used for editing of the report template.
The editor of the report template can be shown as modal whether or not.
See also:
  TvgrReportDesigner, TvgrReportDesignerForm}
unit vgr_ReportDesigner;

{$I vtk.inc}        

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils,
  {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ImgList, clipbrd,
  ActnList, ComCtrls, StdCtrls,
  ToolWin, typinfo, CommCtrl,

  vgr_DataStorage, vgr_CommonClasses, vgr_WorkbookDesigner, vgr_Report,
  vgr_CellPropertiesDialog, vgr_WorkbookGrid, vgr_FontComboBox,
  vgr_ControlBar, ExtCtrls, vgr_ScriptEdit, vgr_PageSetupDialog,
  vgr_WorkbookPreviewForm, vgr_Event, vgr_ScriptParser, vgr_FormLocalizer;

type

  TvgrReportDesignerForm = class;
  /////////////////////////////////////////////////
  //
  // TvgrNodeData
  //
  /////////////////////////////////////////////////
{Internal class. Describes data of the node of the report explorer.
For each node added to the report explorer tree the appropriate TvgrNodeData object is created.
Reference to the instance of this class contains in the TTreeNode.Data property.
The TvgrNodeData class has some descendants (TvgrPersistentNodeData, TvgrBandsNodeData and so on)
which creates for the appropriate objects of the report template.}
  TvgrNodeData = class(TObject)
  private
    FParent: TvgrNodeData;
    FDesignerForm: TvgrReportDesignerForm;
  protected
    function GetNodeImageIndex: Integer; virtual;
    function GetNodeCaption: string; virtual;
    function GetNodeBold: Boolean; virtual;

    property Parent: TvgrNodeData read FParent;
    property NodeImageIndex: Integer read GetNodeImageIndex;
    property NodeCaption: string read GetNodeCaption;
    property NodeBold: Boolean read GetNodeBold;
    property DesignerForm: TvgrReportDesignerForm read FDesignerForm;
  public
{Creates instance of the TvgrNodeData.
Parameters:
  AParent - the parent TvgrNodeData object.}
    constructor Create(ADesignerForm: TvgrReportDesignerForm; AParent: TvgrNodeData);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPersistentNodeData
  //
  /////////////////////////////////////////////////
{Internal class. The instance of this class is created for the objects of the report template derived from the TPersistent class.}
  TvgrPersistentNodeData = class(TvgrNodeData)
  private
    FPersistent: TPersistent;
  protected
    function GetNodeCaption: string; override;
    function GetNodeImageIndex: Integer; override;

    property Persistent: TPersistent read FPersistent;
  public
{Creates instance of the TvgrPersistentNodeData.
Parameters:
  AParent - the parent TvgrNodeData object.
  APersistent - the appropriate TPersistent object.}
    constructor Create(ADesignerForm: TvgrReportDesignerForm;
                       AParent: TvgrNodeData;
                       APersistent: TPersistent);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBandsNodeData
  //
  /////////////////////////////////////////////////
{Internal class. The instance of this class is created for the lists of the bands in the report template (TvgrBands objects).}
  TvgrBandsNodeData = class(TvgrNodeData)
  private
    FBands: TvgrBands;
    FIsVertical: Boolean;
  protected
    function GetNodeCaption: string; override;
    function GetNodeImageIndex: Integer; override;

    property Bands: TvgrBands read FBands;
    property IsVertical: Boolean read FIsVertical;
  public
{Creates instance of the TvgrBandsNodeData.
Parameters:
  AParent - the parent TvgrNodeData object.
  ABands - the appropriate TvgrBands object.
  AIsVertical - true if TvgrBands is the list of the vertical bands.}
    constructor Create(ADesignerForm: TvgrReportDesignerForm;
                       AParent: TvgrNodeData;
                       ABands: TvgrBands;
                       AIsVertical: Boolean);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrEventsNodeData
  //
  /////////////////////////////////////////////////
{Internal class. The instance of this class is created for the objects of the report template derived from the TvgrScriptEvents class.}
  TvgrEventsNodeData = class(TvgrPersistentNodeData)
  private
    FEventsPropInfo: PPropInfo;
    function GetEvents: TvgrScriptEvents;
  protected
    function GetNodeCaption: string; override;
    function GetNodeImageIndex: Integer; override;
    function GetMainObject: TPersistent;
    function GetEventsPropName: string;

    property Events: TvgrScriptEvents read GetEvents;
    property EventsPropInfo: PPropInfo read FEventsPropInfo;
    property MainObject: TPersistent read GetMainObject;
    property EventsPropName: string read GetEventsPropName;
  public
{Creates instance of the TvgrEventsNodeData.
Parameters:
  AParent - the parent TvgrNodeData object.
  APersistent - the appropriate TvgrScriptEvents object.
  AEventsPropInfo - pointer to the TPropInfo structure, that describes appropriate property in the objects holds TvgrScriptEvents object.}
    constructor Create(ADesignerForm: TvgrReportDesignerForm;
                       AParent: TvgrNodeData;
                       APersistent: TPersistent;
                       AEventsPropInfo: PPropInfo);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrEventNodeData
  //
  /////////////////////////////////////////////////
{Internal class. The instance of this class is created for the each event in the TvgrScriptEvents class.
For example, the TvgrBandGenerateEvents class contains two events: BeforeGenerate and AfterGenerate, so
two TvgrEventNodeData objects will be created and added to the appropriate TvgrEventsNodeData node.}
  TvgrEventNodeData = class(TvgrNodeData)
  private
    FPropInfo: PPropInfo;
    function GetEventProcedureName: string;
    procedure SetEventProcedureName(const Value: string);
    function GetDescription: string;
  protected
    function GetNodeCaption: string; override;
    function GetNodeImageIndex: Integer; override;
    function GetNodeBold: Boolean; override;
    function GetEventPropName: string;
    function GetEvents: TvgrScriptEvents;
    function GetMainObject: TPersistent;
    function GetEventsPropName: string;

    function GetAutoProcedureName: string;
    property PropInfo: PPropInfo read FPropInfo;
    property EventPropName: string read GetEventPropName;
    property EventsPropName: string read GetEventsPropName;
    property Events: TvgrScriptEvents read GetEvents;
    property MainObject: TPersistent read GetMainObject;
    property Description: string read GetDescription;
    property EventProcedureName: string read GetEventProcedureName write SetEventProcedureName;
  public
{Creates instance of the TvgrEventsNodeData.
Parameters:
  AParent - the parent TvgrNodeData object.
  AEventsPropInfo - pointer to the TPropInfo structure, that describes the appropriate event property.}  
    constructor Create(ADesignerForm: TvgrReportDesignerForm;
                       AParent: TvgrNodeData;
                       APropInfo: PPropInfo);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrReportDesignerForm
  //
  /////////////////////////////////////////////////
{Implements a form for editing the report template contents.
The form can be shown as modal whether or not.}
  TvgrReportDesignerForm = class(TvgrWorkbookDesignerForm)
    N4: TMenuItem;
    mInsertHorzBand: TMenuItem;
    mInsertVertBand: TMenuItem;
    aFormatHorzBand: TAction;
    aFormatVertBand: TAction;
    N5: TMenuItem;
    Horizontalband1: TMenuItem;
    Verticalband1: TMenuItem;
    TVExplorer: TTreeView;
    ExplorerImageList: TImageList;
    Splitter: TSplitter;
    PScript: TTabSheet;
    ScriptEdit: TvgrScriptEdit;
    PMExplorer: TPopupMenu;
    ExplorerActions: TActionList;
    aProperties: TAction;
    Properties1: TMenuItem;
    aDeleteEvent: TAction;
    N10: TMenuItem;
    Deleteeventhandler1: TMenuItem;
    tbScript: TToolBar;
    edProcedureName: TComboBox;
    aReportOptions: TAction;
    N11: TMenuItem;
    Reportoptions1: TMenuItem;
    ScriptParserTimer: TTimer;
    ToolButton26: TToolButton;
    aSyntaxCheck: TAction;
    ToolButton27: TToolButton;
    ReportDesignerLocalizer: TvgrFormLocalizer;
    aReportSheetProperties: TAction;
    Reportoptions2: TMenuItem;
    procedure aFormatHorzBandUpdate(Sender: TObject);
    procedure aFormatHorzBandExecute(Sender: TObject);
    procedure aFormatVertBandUpdate(Sender: TObject);
    procedure aFormatVertBandExecute(Sender: TObject);
    procedure WorkbookGridSectionDblClick(Sender: TObject;
      ASections: TvgrSections; ASectionIndex: Integer);
    procedure TVExplorerDblClick(Sender: TObject);
    procedure ScriptEditChange(Sender: TObject);
    procedure ScriptEditCaretMove(Sender: TObject; X, Y: Integer);
    procedure ScriptEditEditorModeChange(Sender: TObject);
    procedure TVExplorerDeletion(Sender: TObject; Node: TTreeNode);
    procedure aPropertiesUpdate(Sender: TObject);
    procedure aPropertiesExecute(Sender: TObject);
    procedure aDeleteEventUpdate(Sender: TObject);
    procedure aDeleteEventExecute(Sender: TObject);
    procedure aReportOptionsExecute(Sender: TObject);
    procedure aReportOptionsUpdate(Sender: TObject);
    procedure WorkbookGridActionsUpdate(Action: TBasicAction;
      var Handled: Boolean);
    procedure ScriptParserTimerTimer(Sender: TObject);
    procedure edProcedureNameClick(Sender: TObject);
    procedure aSyntaxCheckUpdate(Sender: TObject);
    procedure aSyntaxCheckExecute(Sender: TObject);
    procedure TVExplorerCustomDrawItem(Sender: TCustomTreeView;
      Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure aReportSheetPropertiesUpdate(Sender: TObject);
    procedure aReportSheetPropertiesExecute(Sender: TObject);
  private
    { Private declarations }
    FScriptModified: Boolean;
    FScriptParser: TvgrScriptParser;
    FScriptInfo: TvgrScriptProgramInfo;
    
    function GetTemplate: TvgrReportTemplate;
    function GetWorksheet: TvgrReportTemplateWorksheet;
    function GetSelectedVertBand: TvgrBand;
    function GetSelectedHorzBand: TvgrBand;
    function GetSelectedData: TvgrNodeData;
    function GetScriptParseDelay: Integer;
    procedure SetScriptParseDelay(Value: Integer);
  protected
    // IvgrWorkbookHandler
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); override;
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); override;
    procedure DeleteItemInterface(AItem: IvgrWBListItem); override;

    procedure SetWorkbook(AValue: TvgrWorkbook); override;

    function GetModified: Boolean; override;
    procedure SetModified(Value: Boolean); override;

    function GetNewDocCaption: string; override;
    function GetTabSheetCaption: string; override;
    function GetFormCaption: string; override;

    function IsCutAllowed: Boolean; override;
    function IsPasteAllowed: Boolean; override;
    function IsCopyAllowed: Boolean; override;

    procedure OnInsertHorzBand(Sender: TObject);
    procedure OnInsertVertBand(Sender: TObject);
    procedure SetupDialogs(AOpenDialog: TOpenDialog; ASaveDialog: TSaveDialog); override;

    function GetExplorerNodeEditText(AObject: TObject): string;
    procedure UpdateExplorer;
    procedure RefreshExplorer;

    procedure UpdateScriptParser;
    procedure UpdateProgramInfo;
    procedure StartScriptParserTimer;
    procedure StopScriptParserTimer;

    procedure BuildStatusBar; override;
    procedure UpdateStatusBar; override;
    procedure InternalOnSave; override;
    procedure InternalOnLoad; override;
    procedure SetInplaceEditValue(Sender: TObject; const AValue: String; ARange: IvgrRange; ANewRange: Boolean); override;

    procedure EditExplorerNode(ANode: TTreeNode);

    procedure RestoreSettings; override;
    procedure SaveSettings; override;

    property ScriptParser: TvgrScriptParser read FScriptParser;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;

{Cuts current selection to the clipboard.}
    procedure CutToClipboard; override;
{Copy current selection to the clipboard.}
    procedure CopyToClipboard; override;
{Pastes data from the clipboard to the current position.}
    procedure PasteFromClipboard; override;

{Showes the dialog form, that allow to enter the script procedure name, that must be used
for processing of the appropriate event.
Parameters:
  AEventNodeData - the TvgrEventNodeData object that describes event.
See also:
  TvgrEventNodeData}
    procedure EditEventProcedure(AEventNodeData: TvgrEventNodeData);
{Found the script procedure, creates them if not found and goes to them,
after this activates the script page and set focus to the script editor.
Parameters:
  AEventNodeData - the TvgrEventNodeData object that describes event.
See also:
  TvgrEventNodeData}
    procedure GotoToEvent(AEventNodeData: TvgrEventNodeData);
{Found the script procedure and goes to them.
Parameters:
  AProcedureInfo - the TvgrScriptProcedureInfo object, that describes the needed script procedure.
See also:
  TvgrScriptProcedureInfo, ScriptInfo, TvgrScriptProgramInfo}
    procedure GotoProcedure(AProcedureInfo: TvgrScriptProcedureInfo);
{Activates the script page and set focus to the script editor.}
    procedure ShowScriptPage;

{Sets or gets the edited TvgrReportTemplate object.}
    property Template: TvgrReportTemplate read GetTemplate;
{Returns the currently selected worksheet.}
    property ActiveWorksheet: TvgrReportTemplateWorksheet read GetWorksheet;
{Returns the currently selected vertically band.
The selected band is the band within of which cursor are placed.}
    property SelectedVertBand: TvgrBand read GetSelectedVertBand;
{Returns the currently selected horizontally band.
The selected band is the band within of which cursor are placed.}
    property SelectedHorzBand: TvgrBand read GetSelectedHorzBand;
{Returns the TvgrNodeData object that are owned by the selected node in the report explorer.}
    property SelectedData: TvgrNodeData read GetSelectedData;

{After each changes in the script editor the text is parsed and the lists of the script procedures are created.
This property specifies the delay in milliseconds that must be ended before starting of the text parsing.
See also:
  ScriptInfo}
    property ScriptParseDelay: Integer read GetScriptParseDelay write SetScriptParseDelay;

{Returns TvgrScriptProgramInfo object that describes the parsed script.}
    property ScriptInfo: TvgrScriptProgramInfo read FScriptInfo;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrReportDesigner
  //
  /////////////////////////////////////////////////
{Non visual component, realizes the designer of the report template (TvgrReportTemplate object).
The designer does not appear at runtime until it is activated by a call to the Design method.
The designer can be shown as modal whether or not.
See also:
  TvgrReportDesignerForm}
  TvgrReportDesigner = class(TvgrCustomWorkbookDesigner)
  private
    function GetTemplate: TvgrReportTemplate;
    procedure SetTemplate(Value: TvgrReportTemplate);
    function GetForm: TvgrReportDesignerForm;
  protected
    function GetDesignerFormClass: TvgrWorkbookDesignerFormClass; override;
  public
{Returns TvgrReportDesignerForm that is created by this component.}
    property Form: TvgrReportDesignerForm read GetForm;
  published
{Sets or gets the edited report template (TvgrReportTemplate object).}
    property Template: TvgrReportTemplate read GetTemplate write SetTemplate;
  end;

implementation

{$R *.dfm}
{$R ..\res\vgr_ReportDesignerStrings.res}

uses
  vgr_SheetFormatDialog, vgr_StringIDs,
  vgr_Functions, vgr_Dialogs, vgr_BandFormatDialog,
  vgr_ScriptEventEditDialog, vgr_ReportOptionsDialog, vgr_Form, vgr_Localize,
  vgr_ScriptControl, vgr_ReportSheetFormatDialog;

const
  sPostFix = 'ScriptProcName';

function FindTreeNodeByData(ATreeView: TTreeView; AData: Pointer): TTreeNode;
var
  I: Integer;
begin
  I := 0;
  while (I < ATreeView.Items.Count) and (ATreeView.Items[I].Data <> AData) do Inc(I);
  if I < ATreeView.Items.Count then
    Result := ATreeView.Items[I]
  else
    Result := nil;
end;

/////////////////////////////////////////////////
//
// TvgrNodeData
//
/////////////////////////////////////////////////
constructor TvgrNodeData.Create(ADesignerForm: TvgrReportDesignerForm; AParent: TvgrNodeData);
begin
  inherited Create;
  FDesignerForm := ADesignerForm;
  FParent := AParent;
end;

function TvgrNodeData.GetNodeImageIndex: Integer;
begin
  Result := -1;
end;

function TvgrNodeData.GetNodeCaption: string;
begin
  Result := '';
end;

function TvgrNodeData.GetNodeBold: Boolean;
begin
  Result := False;
end;

/////////////////////////////////////////////////
//
// TvgrPersistentNodeData
//
/////////////////////////////////////////////////
constructor TvgrPersistentNodeData.Create(ADesignerForm: TvgrReportDesignerForm; AParent: TvgrNodeData; APersistent: TPersistent);
begin
  inherited Create(ADesignerForm, AParent);
  FPersistent := APersistent;
end;

function TvgrPersistentNodeData.GetNodeImageIndex: Integer;
begin
  if Persistent is TvgrReportTemplate then
    Result := 0
  else
    if Persistent is TvgrReportTemplateWorksheet then
      Result := 1
    else
      if Persistent is TvgrBand then
        Result := 6 + BandClasses.GetBandClassIndex(TvgrBand(Persistent))
      else
        if Persistent <> nil then
          Result := 6
        else
          Result := inherited GetNodeImageIndex;
end;

function TvgrPersistentNodeData.GetNodeCaption: string;
begin
  if Persistent is TvgrReportTemplate then
    Result := vgrLoadStr(svgrid_vgr_ReportDesigner_ReportTreeNodeCaption)
  else
    if Persistent is TvgrReportTemplateWorksheet then
      with TvgrReportTemplateWorksheet(Persistent) do
        Result := Format('%s (%s)', [Title, Name])
    else
      if Persistent is TvgrBand then
        with TvgrBand(Persistent) do
          Result := Format('%s (%s)', [GetCaption, Name])
      else
        if Persistent <> nil then
          Result := Persistent.ClassName
        else
          Result := inherited GetNodeCaption;
end;

/////////////////////////////////////////////////
//
// TvgrBandsNodeData
//
/////////////////////////////////////////////////
constructor TvgrBandsNodeData.Create(ADesignerForm: TvgrReportDesignerForm; AParent: TvgrNodeData; ABands: TvgrBands; AIsVertical: Boolean);
begin
  inherited Create(ADesignerForm, AParent);
  FBands := ABands;
  FIsVertical := AIsVertical;
end;

function TvgrBandsNodeData.GetNodeCaption: string;
begin
  if IsVertical then
    Result := vgrLoadStr(svgrid_vgr_ReportDesigner_VertBandsTreeNodeCaption)
  else
    Result := vgrLoadStr(svgrid_vgr_ReportDesigner_HorzBandsTreeNodeCaption);
end;

function TvgrBandsNodeData.GetNodeImageIndex: Integer;
begin
  if IsVertical then
    Result := 3
  else
    Result := 2;
end;

/////////////////////////////////////////////////
//
// TvgrEventsNodeData
//
/////////////////////////////////////////////////
constructor TvgrEventsNodeData.Create(ADesignerForm: TvgrReportDesignerForm;
                                      AParent: TvgrNodeData;
                                      APersistent: TPersistent;
                                      AEventsPropInfo: PPropInfo);
begin
  inherited Create(ADesignerForm, AParent, APersistent);
  FEventsPropInfo := AEventsPropInfo;
end;

function TvgrEventsNodeData.GetNodeCaption: string;
begin
  Result := EventsPropInfo.Name;
end;

function TvgrEventsNodeData.GetNodeImageIndex: Integer;
begin
  Result := 4;
end;

function TvgrEventsNodeData.GetEvents: TvgrScriptEvents;
begin
  Result := TvgrScriptEvents(inherited Persistent);
end;

function TvgrEventsNodeData.GetMainObject: TPersistent;
begin
  if Parent is TvgrPersistentNodeData then
    Result := TvgrPersistentNodeData(Parent).Persistent
  else
    Result := nil;
end;

function TvgrEventsNodeData.GetEventsPropName: string;
begin
  Result := Trim(EventsPropInfo.Name);
end;

/////////////////////////////////////////////////
//
// TvgrEventNodeData
//
/////////////////////////////////////////////////
constructor TvgrEventNodeData.Create(ADesignerForm: TvgrReportDesignerForm; AParent: TvgrNodeData; APropInfo: PPropInfo);
begin
  inherited Create(ADesignerForm, AParent);
  FPropInfo := APropInfo;
end;

function TvgrEventNodeData.GetEventProcedureName: string;
begin
  Result := Trim(GetStrProp(Events, PropInfo));
end;

procedure TvgrEventNodeData.SetEventProcedureName(const Value: string);
begin
  SetStrProp(Events, PropInfo, Value);
end;

function TvgrEventNodeData.GetNodeImageIndex: Integer;
var
  S: string;
begin
  S := GetEventProcedureName;
  if S <> '' then
  begin
    if DesignerForm.ScriptInfo.IndexOfProcedure(S) = -1 then
      Result := 8
    else
      Result := 7
  end
  else
    Result := 5;
end;

function TvgrEventNodeData.GetNodeBold: Boolean;
begin
  Result := EventProcedureName <> '';
end;

function TvgrEventNodeData.GetEventPropName: string;
begin
  Result := Trim(PropInfo.Name);
end;

function TvgrEventNodeData.GetNodeCaption: string;
begin
  Result := EventPropName;
  if Copy(Result, Length(Result) - Length(sPostFix) + 1, Length(sPostFix)) = sPostFix then
    Delete(Result, Length(Result) - Length(sPostFix) + 1, Length(sPostFix));
end;

function TvgrEventNodeData.GetEvents: TvgrScriptEvents;
begin
  if Parent is TvgrEventsNodeData then
    Result := TvgrEventsNodeData(Parent).Events
  else
    Result := nil;
end;

function TvgrEventNodeData.GetMainObject: TPersistent;
begin
  if Parent is TvgrEventsNodeData then
    Result := TvgrEventsNodeData(Parent).MainObject
  else
    Result := nil;
end;

function TvgrEventNodeData.GetEventsPropName: string;
begin
  if Parent is TvgrEventsNodeData then
    Result := TvgrEventsNodeData(Parent).EventsPropName
  else
    Result := '';
end;

function TvgrEventNodeData.GetDescription: string;
begin
  if MainObject is TComponent then
    Result := Format('%s.%s.%s', [TComponent(MainObject).Name, EventsPropName, NodeCaption])
  else
    Result := Format('%s.%s.%s', [MainObject.ClassName, EventsPropName, NodeCaption]);
end;

function TvgrEventNodeData.GetAutoProcedureName: string;
begin
  if MainObject is TComponent then
    Result := Format('%s_%s_%s', [TComponent(MainObject).Name, EventsPropName, NodeCaption])
  else
    Result := Format('%s_%s_%s', [MainObject.ClassName, EventsPropName, NodeCaption]);
end;

/////////////////////////////////////////////////
//
// TvgrReportDesigner
//
/////////////////////////////////////////////////
function TvgrReportDesigner.GetDesignerFormClass: TvgrWorkbookDesignerFormClass;
begin
  Result := TvgrReportDesignerForm;
end;

function TvgrReportDesigner.GetTemplate: TvgrReportTemplate;
begin
  Result := TvgrReportTemplate(Workbook);
end;

procedure TvgrReportDesigner.SetTemplate(Value: TvgrReportTemplate);
begin
  Workbook := Value;
end;

function TvgrReportDesigner.GetForm: TvgrReportDesignerForm;
begin
  Result := TvgrReportDesignerForm(inherited Form);
end;

/////////////////////////////////////////////////
//
// TvgrReportDesignerForm
//
/////////////////////////////////////////////////
function TvgrReportDesignerForm.GetWorksheet: TvgrReportTemplateWorksheet;
begin
  Result := TvgrReportTemplateWorksheet(inherited ActiveWorksheet);
end;

function TvgrReportDesignerForm.GetTemplate: TvgrReportTemplate;
begin
  Result := TvgrReportTemplate(Workbook);
end;

procedure TvgrReportDesignerForm.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); 
begin
  inherited;
end;

procedure TvgrReportDesignerForm.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcNewWorksheet, vgrwcNewHorzSection, vgrwcNewVertSection,
    vgrwcDeleteWorksheet, vgrwcDeleteHorzSection, vgrwcDeleteVertSection,
    vgrwcChangeHorzSection, vgrwcChangeVertSection:
      UpdateExplorer;
    vgrwcChangeWorksheet:
      RefreshExplorer;
    vgrwcUpdateAll:
      begin
        UpdateProgramInfo;
        UpdateExplorer;
      end;
  end;
  inherited;
end;

procedure TvgrReportDesignerForm.DeleteItemInterface(AItem: IvgrWBListItem);
begin
end;

function TvgrReportDesignerForm.GetModified: Boolean;
begin
  Result := inherited GetModified or FScriptModified;
end;

procedure TvgrReportDesignerForm.SetModified(Value: Boolean);
begin
  FScriptModified := Value;
  inherited;
end;

function TvgrReportDesignerForm.GetNewDocCaption: string;
begin
  Result := vgrLoadStr(svgrid_vgr_ReportDesigner_NewTemplateCaption);
end;

function TvgrReportDesignerForm.GetTabSheetCaption: string;
begin
  Result := vgrLoadStr(svgrid_vgr_ReportDesigner_ReportTabSheetCaption);
end;

function TvgrReportDesignerForm.GetFormCaption: string;
begin
  Result := Format('%s - [%s]', [Application.Title, DocCaption])
end;

procedure TvgrReportDesignerForm.OnInsertHorzBand(Sender: TObject);
var
  ABand: TvgrBand;
begin
  ABand := ActiveWorksheet.CreateHorzBand(BandClasses[TMenuItem(Sender).Tag]);
  with WorkbookGrid.LastSelection do
  begin
    ABand.EndPos := Bottom;
    ABand.StartPos := Top;
  end;
end;

procedure TvgrReportDesignerForm.OnInsertVertBand(Sender: TObject);
var
  ABand: TvgrBand;
begin
  ABand := ActiveWorksheet.CreateVertBand(BandClasses[TMenuItem(Sender).Tag]);
  with WorkbookGrid.LastSelection do
  begin
    ABand.EndPos := Right;
    ABand.StartPos := Left;
  end;
end;

procedure TvgrReportDesignerForm.SetupDialogs(AOpenDialog: TOpenDialog; ASaveDialog: TSaveDialog);
begin
  with AOpenDialog do
  begin
    Filter := vgrLoadStr(svgrid_vgr_ReportDesigner_TemplateOpenSaveDialogFilter);
    DefaultExt := vgrLoadStr(svgrid_vgr_ReportDesigner_TemplateOpenSaveDefaultExt);
    Options := [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing];
  end;
  with ASaveDialog do
  begin
    Filter := vgrLoadStr(svgrid_vgr_ReportDesigner_TemplateOpenSaveDialogFilter);
    DefaultExt := vgrLoadStr(svgrid_vgr_ReportDesigner_TemplateOpenSaveDefaultExt);
    Options := [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing];
  end;
end;

function TvgrReportDesignerForm.GetExplorerNodeEditText(AObject: TObject): string;
begin
  if AObject is TvgrReportTemplate then
    Result := TvgrReportTemplate(AObject).Name
  else
    if AObject is TvgrReportTemplateWorksheet then
      Result := TvgrReportTemplateWorksheet(AObject).Name
    else
      if AObject is TvgrBand then
        Result := TvgrBand(AObject).Name
      else
        Result := '';
end;

procedure TvgrReportDesignerForm.RefreshExplorer;
var
  I: Integer;
begin
  for I := 0 to TVExplorer.Items.Count - 1 do
    if TVExplorer.Items[I].Data <> nil then
    begin
      with TVExplorer.Items[I] do
      begin
        Text := TvgrNodeData(Data).NodeCaption;
        ImageIndex := TvgrNodeData(Data).NodeImageIndex;
        SelectedIndex := ImageIndex;
      end;
    end;
end;

procedure TvgrReportDesignerForm.UpdateExplorer;
var
  I: Integer;
  ANode: TTreeNode;
  ASaveSelectedNodeData: TvgrNodeData;

  procedure AddEventsNodes(AParentNode: TTreeNode; const ACaption: string; AEvents: TvgrScriptEvents; AEventsPropInfo: PPropInfo);
  var
    I: Integer;
    ANode: TTreeNode;
    AEventNode: TTreeNode;
    APropsList: PPropList;
    APropsCount: Integer;
    AEventsNodeData: TvgrEventsNodeData;
    AEventNodeData: TvgrEventNodeData;
  begin
    AEventsNodeData := TvgrEventsNodeData.Create(Self, AParentNode.Data, AEvents, AEventsPropInfo);
    ANode := TVExplorer.Items.AddChildObject(AParentNode, AEventsNodeData.NodeCaption, AEventsNodeData);
    ANode.ImageIndex := AEventsNodeData.NodeImageIndex;
    ANode.SelectedIndex := ANode.ImageIndex;
    APropsCount := GetPropList(AEvents.ClassInfo, [tkString, tkLString, tkWString], nil);
    if APropsCount > 0 then
    begin
      GetMem(APropsList, APropsCount * SizeOf(PPropInfo));
      try
        GetPropList(AEvents.ClassInfo, [tkString, tkLString, tkWString], APropsList);
        for I := 0 to APropsCount - 1 do
        begin
          AEventNodeData := TvgrEventNodeData.Create(Self, AEventsNodeData, APropsList[I]);
          AEventNode := TVExplorer.Items.AddChildObject(ANode, AEventNodeData.NodeCaption, AEventNodeData);
          with AEventNode do
          begin
            ImageIndex := AEventNodeData.NodeImageIndex;
            SelectedIndex := ImageIndex;
          end;
        end;
      finally
        FreeMem(APropsList);
      end;
    end;
  end;

  function AddObjectNode(AParentNode: TTreeNode; AObject: TPersistent): TTreeNode;
  var
    I, APropsCount: Integer;
    APropsList: PPropList;
    ASub: TObject;
    ANodeData: TvgrPersistentNodeData;
  begin
    if AParentNode <> nil then
      ANodeData := TvgrPersistentNodeData.Create(Self, AParentNode.Data, AObject)
    else
      ANodeData := TvgrPersistentNodeData.Create(Self, nil, AObject);
    Result := TVExplorer.Items.AddChildObject(AParentNode, ANodeData.NodeCaption, ANodeData);
    Result.ImageIndex := ANodeData.NodeImageIndex;
    Result.SelectedIndex := Result.ImageIndex;

    if AObject <> nil then
    begin
      // add events
      APropsCount := GetPropList(AObject.ClassInfo, [tkClass], nil);
      if APropsCount > 0 then
      begin
        GetMem(APropsList, APropsCount * SizeOf(PPropInfo));
        try
          GetPropList(AObject.ClassInfo, [tkClass], APropsList);
          for I := 0 to APropsCount - 1 do
          begin
            ASub := TPersistent(GetOrdProp(AObject, APropsList[I]));
            if ASub is TvgrScriptEvents then
              AddEventsNodes(Result, APropsList[I].Name, TvgrScriptEvents(ASub), APropsList[I]);
          end
        finally
          FreeMem(APropsList);
        end;
      end;
    end;
  end;
  
  function AddWorksheetNode(AParent: TTreeNode; AWorksheet: TvgrReportTemplateWorksheet): TTreeNode;

    function AddBandsNodes(AParent: TTreeNode; ABands: TvgrBands; AIsVertical: Boolean): TTreeNode;
    var
      ABandsNodeData: TvgrBandsNodeData;

      procedure AddBandNodes(AParentNode: TTreeNode; ABands: TvgrBands; AParentBand: TvgrBand);
      var
        I: Integer;
      begin
        for I := 0 to ABands.Count - 1 do
          if ABands.ByIndex[I].ParentBand = AParentBand then
            AddBandNodes(AddObjectNode(AParentNode, ABands.ByIndex[I]),
                         ABands, ABands.ByIndex[I]);
      end;

    begin
      ABandsNodeData := TvgrBandsNodeData.Create(Self, AParent.Data, ABands, AIsVertical);
      Result := TVExplorer.Items.AddChildObject(AParent, ABandsNodeData.NodeCaption, ABandsNodeData);
      Result.ImageIndex := ABandsNodeData.NodeImageIndex;
      Result.SelectedIndex := ANode.ImageIndex;

      AddBandNodes(Result, ABands, nil);
    end;

  begin
    Result := AddObjectNode(AParent, AWorksheet);
    AddBandsNodes(Result, AWorksheet.HorzBands, False);
    AddBandsNodes(Result, AWorksheet.VertBands, True);
  end;

begin
  TVExplorer.Items.BeginUpdate;
  try
    ASaveSelectedNodeData := SelectedData;

    TVExplorer.Items.Clear;
    ANode := AddObjectNode(nil, Template);
    for I := 0 to Template.WorksheetsCount - 1 do
      AddWorksheetNode(ANode, Template.Worksheets[I]);

    ANode.Expand(True);
    ANode :=  FindTreeNodeByData(TVExplorer, ASaveSelectedNodeData);
    if ANode <> nil then
      TVExplorer.Selected := ANode
    else
      TVExplorer.Selected := TVExplorer.Items[0];
  finally
    TVExplorer.Items.EndUpdate;
  end;
end;

procedure TvgrReportDesignerForm.AfterConstruction;
var
  I: Integer;

  procedure FillInsertBandMenuItem(AMenuItem: TMenuItem; AOnClick: TNotifyEvent);
  var
    I: Integer;
  begin
    AMenuItem.Clear;
    if BandClasses.Count = 0 then
      AMenuItem.Enabled := False
    else
    begin
      AMenuItem.Enabled := True;
      for I := 0 to BandClasses.Count - 1 do
        AddMenuItem(MainMenu, AMenuItem, BandClasses[I].GetCaption, AOnClick, I, BandClasses[I].GetBitmapResName);
    end;
  end;

begin
  inherited;
  ScriptParserTimer.Interval := 1000;
  
  FScriptInfo := TvgrScriptProgramInfo.Create;
  
  FillInsertBandMenuItem(mInsertHorzBand, OnInsertHorzBand);
  FillInsertBandMenuItem(mInsertVertBand, OnInsertVertBand);

  for I := 0 to BandClasses.Count - 1 do
    ExplorerImageList.AddMasked(BandClasses.Bitmaps[I], clPurple);
end;

procedure TvgrReportDesignerForm.BeforeDestruction;
begin
  inherited;
  StopScriptParserTimer;
  FreeAndNil(FScriptParser);
  FreeAndNil(FScriptInfo);
end;

procedure TvgrReportDesignerForm.SetWorkbook(AValue: TvgrWorkbook);
begin
  inherited;
  UpdateScriptParser;
  UpdateProgramInfo;
  if Template = nil then
    ScriptEdit.Script.Language := ''
  else
  begin
    ScriptEdit.Lines.Assign(Template.Script.Script);
    ScriptEdit.Script.Language := Template.Script.Language;
    FScriptModified := False;
    UpdateExplorer;
  end;
end;

function TvgrReportDesignerForm.GetSelectedVertBand: TvgrBand;
begin
  with WorkbookGrid.LastSelection do
    Result := ActiveWorksheet.FindVertBandAt(Left, Right);
end;

function TvgrReportDesignerForm.GetSelectedHorzBand: TvgrBand;
begin
  with WorkbookGrid.LastSelection do
    Result := ActiveWorksheet.FindHorzBandAt(Top, Bottom);
end;

function TvgrReportDesignerForm.GetSelectedData: TvgrNodeData;
begin
  if TVExplorer.Selected = nil then
    Result := nil
  else
    Result := TVExplorer.Selected.Data;
end;

function TvgrReportDesignerForm.GetScriptParseDelay: Integer;
begin
  Result := ScriptParserTimer.Interval;
end;

procedure TvgrReportDesignerForm.SetScriptParseDelay(Value: Integer);
begin
  ScriptParserTimer.Interval := Value;
end;

procedure TvgrReportDesignerForm.aFormatHorzBandUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit and (SelectedHorzBand <> nil);
end;

procedure TvgrReportDesignerForm.aFormatHorzBandExecute(Sender: TObject);
begin
  TvgrBandFormatForm.Create(nil).Execute(SelectedHorzBand);
end;

procedure TvgrReportDesignerForm.aFormatVertBandUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit and (SelectedVertBand <> nil);
end;

procedure TvgrReportDesignerForm.aFormatVertBandExecute(Sender: TObject);
begin
  TvgrBandFormatForm.Create(nil).Execute(SelectedVertBand);
end;

procedure TvgrReportDesignerForm.WorkbookGridSectionDblClick(
  Sender: TObject; ASections: TvgrSections; ASectionIndex: Integer);
begin
  TvgrBandFormatForm.Create(nil).Execute(TvgrBands(ASections).ByIndex[ASectionIndex]);
end;

procedure TvgrReportDesignerForm.TVExplorerDblClick(Sender: TObject);
begin
  if SelectedData is TvgrEventNodeData then
    GotoToEvent(TvgrEventNodeData(SelectedData))
  else
    EditExplorerNode(TVExplorer.Selected);
end;

function TvgrReportDesignerForm.IsCutAllowed: Boolean;
begin
  if PageControl.ActivePage = PGrid then
    Result := inherited IsCutAllowed
  else
    Result := True;
end;

function TvgrReportDesignerForm.IsPasteAllowed: Boolean;
begin
  if PageControl.ActivePage = PGrid then
    Result := inherited IsPasteAllowed
  else
    Result := Clipboard.HasFormat(CF_TEXT);
end;

function TvgrReportDesignerForm.IsCopyAllowed: Boolean;
begin
  if PageControl.ActivePage = PGrid then
    Result := inherited IsCopyAllowed
  else
    Result := True;
end;

procedure TvgrReportDesignerForm.CutToClipboard;
begin
  if PageControl.ActivePage = PGrid then
    inherited
  else
    ScriptEdit.ClipBoardCut;
end;

procedure TvgrReportDesignerForm.CopyToClipboard;
begin
  if PageControl.ActivePage = PGrid then
    inherited
  else
    ScriptEdit.ClipBoardCopy;
end;

procedure TvgrReportDesignerForm.PasteFromClipboard;
begin
  if PageControl.ActivePage = PGrid then
    inherited
  else
    ScriptEdit.ClipBoardPaste;
end;

procedure TvgrReportDesignerForm.BuildStatusBar;
begin
  inherited;
  if PageControl.ActivePage = PScript then
  begin
    with StatusBar.Panels.Add do
      Width := 100;
    with StatusBar.Panels.Add do
      Width := 60;
  end;
end;

procedure TvgrReportDesignerForm.UpdateStatusBar;
begin
  inherited;
  if PageControl.ActivePage = PScript then
  begin
    if StatusBar.Panels.Count > 1 then
      StatusBar.Panels[1].Text := Format('%d : %d', [ScriptEdit.CaretPos.Y + 1, ScriptEdit.CaretPos.X + 1]);
    if StatusBar.Panels.Count > 2 then
      case ScriptEdit.EditorMode of
        vmmInsert: StatusBar.Panels[2].Text := vgrLoadStr(svgrid_vgr_ReportDesigner_Insert);
        else StatusBar.Panels[2].Text := vgrLoadStr(svgrid_vgr_ReportDesigner_Overwrite);
      end;
  end;
end;

procedure TvgrReportDesignerForm.ScriptEditChange(Sender: TObject);
begin
  FScriptModified := True;
  StartScriptParserTimer;
end;

procedure TvgrReportDesignerForm.ScriptEditCaretMove(Sender: TObject; X,
  Y: Integer);
begin
  UpdateStatusBar;
end;

procedure TvgrReportDesignerForm.ScriptEditEditorModeChange(
  Sender: TObject);
begin
  UpdateStatusBar;
end;

procedure TvgrReportDesignerForm.InternalOnSave;
begin
  if FScriptModified then
    Template.Script.Script := ScriptEdit.Lines;
end;

procedure TvgrReportDesignerForm.InternalOnLoad;
begin
  ScriptEdit.Lines.Assign(Template.Script.Script);
end;

procedure TvgrReportDesignerForm.SetInplaceEditValue(Sender: TObject; const AValue: String; ARange: IvgrRange; ANewRange: Boolean); 
begin
  ARange.SimpleStringValue := AValue;
  if ANewRange then
    ARange.WordWrap := (pos(#13, AValue) <> 0) or (pos(#10, AValue) <> 0);
end;

procedure TvgrReportDesignerForm.UpdateScriptParser;
var
  AScriptParserClass: TvgrScriptParserClass;
begin
  FreeAndNil(FScriptParser);
  if Template <> nil then
  begin
    AScriptParserClass := GetScriptParserClass(Template.Script.Language);
    if AScriptParserClass <> nil then
      FScriptParser := AScriptParserClass.Create;
  end;
end;

procedure TvgrReportDesignerForm.UpdateProgramInfo;
var
  I: Integer;
begin
  if ScriptParser = nil then
    ScriptInfo.Clear
  else
    ScriptParser.ParseProgram(ScriptEdit.Lines.Text, ScriptInfo);
  edProcedureName.Items.BeginUpdate;
  try
    edProcedureName.Items.Clear;
    for I := 0 to ScriptInfo.ProcedureCount - 1 do
      edProcedureName.Items.AddObject(ScriptInfo.Procedures[I].Name, ScriptInfo.Procedures[I]);
    edProcedureName.ItemIndex := -1;
  finally
    edProcedureName.Items.EndUpdate;
  end;
end;

procedure TvgrReportDesignerForm.StartScriptParserTimer;
begin
  ScriptParserTimer.Enabled := False;
  ScriptParserTimer.Enabled := True;
end;

procedure TvgrReportDesignerForm.StopScriptParserTimer;
begin
  ScriptParserTimer.Enabled := False;
end;

procedure TvgrReportDesignerForm.TVExplorerDeletion(Sender: TObject;
  Node: TTreeNode);
begin
  if TObject(Node.Data) is TvgrNodeData then
    TvgrNodeData(Node.Data).Free;
end;

procedure TvgrReportDesignerForm.EditExplorerNode(ANode: TTreeNode);
begin
  if ANode <> nil then
  begin
    if (TvgrNodeData(ANode.Data) is TvgrPersistentNodeData) and
       (TvgrPersistentNodeData(ANode.Data).Persistent is TvgrReportTemplateWorksheet) then
      WorkbookGrid.EditWorksheet(TvgrReportTemplateWorksheet(TvgrPersistentNodeData(ANode.Data).Persistent))
    else
      if (TvgrNodeData(ANode.Data) is TvgrPersistentNodeData) and
         (TvgrPersistentNodeData(ANode.Data).Persistent is TvgrBand) then
        TvgrBandFormatForm.Create(nil).Execute(TvgrBand(TvgrPersistentNodeData(ANode.Data).Persistent))
      else
        if TvgrNodeData(ANode.Data) is TvgrEventNodeData then
          EditEventProcedure(TvgrEventNodeData(ANode.Data));
  end;
end;

procedure TvgrReportDesignerForm.RestoreSettings;
begin
  inherited;
  TVExplorer.Width := SettingsStorage.ReadInteger(StorageSection, 'ExplorerWidth', TVExplorer.Width);
end;

procedure TvgrReportDesignerForm.SaveSettings;
begin
  inherited;
  SettingsStorage.WriteInteger(StorageSection, 'ExplorerWidth', TVExplorer.Width);
end;

procedure TvgrReportDesignerForm.aPropertiesUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := SelectedData <> nil;
end;

procedure TvgrReportDesignerForm.aPropertiesExecute(Sender: TObject);
begin
  EditExplorerNode(TVExplorer.Selected);
end;

procedure TvgrReportDesignerForm.aDeleteEventUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := (SelectedData is TvgrEventNodeData) and
                             (GetStrProp(TvgrEventNodeData(SelectedData).Events,
                                         TvgrEventNodeData(SelectedData).PropInfo.Name) <> '');
end;

procedure TvgrReportDesignerForm.aDeleteEventExecute(Sender: TObject);
begin
  with TvgrEventNodeData(SelectedData) do
  begin
    SetStrProp(Events, PropInfo, '');
    RefreshExplorer;
  end;
end;

procedure TvgrReportDesignerForm.aReportOptionsExecute(Sender: TObject);
begin
  if TvgrReportOptionsDialogForm.Create(nil).Execute(Template) then
  begin
    ScriptEdit.Script.Language := Template.Script.Language;
    // Update script parser
    UpdateScriptParser;
    UpdateProgramInfo;
  end;
end;

procedure TvgrReportDesignerForm.aReportOptionsUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrReportDesignerForm.WorkbookGridActionsUpdate(
  Action: TBasicAction; var Handled: Boolean);
begin
  inherited;
  if Action is TAction then
  begin
    edProcedureName.Enabled := not WorkbookGrid.IsInplaceEdit and (ScriptParser <> nil);
  end;
end;

procedure TvgrReportDesignerForm.ScriptParserTimerTimer(Sender: TObject);
begin
  StopScriptParserTimer;
  UpdateProgramInfo;
  RefreshExplorer;
end;

procedure TvgrReportDesignerForm.GotoProcedure(AProcedureInfo: TvgrScriptProcedureInfo);
var
  ARow, ACol: Integer;
begin
  if PosToRowAndCol(ScriptEdit.Lines, AProcedureInfo.StartPos, ARow, ACol) then
  begin
    ScriptEdit.SetCaretPos(ACol - 1, ARow - 1);
    ShowScriptPage;
  end;
end;

procedure TvgrReportDesignerForm.GotoToEvent(AEventNodeData: TvgrEventNodeData);
var
  S, AProcedureText: string;
  I, J: Integer;
  AStrings : TStringList;
begin
  if ScriptParser <> nil then
  begin
    S := AEventNodeData.EventProcedureName;
    if S <> '' then
    begin
      I := ScriptInfo.IndexOfProcedure(S);
      if I <> -1 then
      begin
        GotoProcedure(ScriptInfo.Procedures[I]);
        exit;
      end;
    end;
    // generate new procedure
    I := RegisteredScriptParsers.IndexOfEvent(AEventNodeData.MainObject.ClassType,
                                              AEventNodeData.EventsPropName,
                                              AEventNodeData.EventPropName);
    if I = -1 then
      // event not registered
      EditEventProcedure(AEventNodeData)
    else
    begin
      // generate event text and go to it
      if S = '' then
      begin
        S := AEventNodeData.GetAutoProcedureName;
        AEventNodeData.EventProcedureName := S;
      end;
      J := ScriptInfo.IndexOfProcedure(S);
      if J = -1 then
      begin
        ScriptParser.GenerateEventProcedureText(RegisteredScriptParsers.Events[I], S, AProcedureText);
        AStrings := TStringList.Create;
        try
          AStrings.Text := AProcedureText;
          ScriptEdit.Lines.AddStrings(AStrings);
          ScriptEdit.Lines.Add('');
//          ScriptEdit.Update;
        finally
          AStrings.Free;
        end;
        UpdateProgramInfo;
        RefreshExplorer;
        J := ScriptInfo.IndexOfProcedure(S);
      end;
      if J <> -1 then
        GotoProcedure(ScriptInfo.Procedures[J]);
    end;
  end;
end;

procedure TvgrReportDesignerForm.EditEventProcedure(AEventNodeData: TvgrEventNodeData);
var
  S: string;
begin
  S := AEventNodeData.EventProcedureName;
  if TvgrScriptEventEditForm.Create(nil).Execute(AEventNodeData.Description, ScriptInfo, S) then
  begin
    AEventNodeData.EventProcedureName := S;
  end;
end;

procedure TvgrReportDesignerForm.edProcedureNameClick(Sender: TObject);
begin
  if edProcedureName.ItemIndex >= 0 then
    GotoProcedure(TvgrScriptProcedureInfo(edProcedureName.Items.Objects[edProcedureName.ItemIndex]));
end;

procedure TvgrReportDesignerForm.aSyntaxCheckUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrReportDesignerForm.ShowScriptPage;
begin
  PageControl.ActivePage := PScript;
  ActiveControl := ScriptEdit;
end;

procedure TvgrReportDesignerForm.aSyntaxCheckExecute(Sender: TObject);
begin
  ShowScriptPage;
  ScriptEdit.CheckSyntax;
end;

procedure TvgrReportDesignerForm.TVExplorerCustomDrawItem(
  Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  if (TObject(Node.Data) is TvgrNodeData) and TvgrNodedata(Node.Data).NodeBold then
  begin
    TVExplorer.Canvas.Font.Style := TVExplorer.Canvas.Font.Style + [fsBold];
  end
  else
  begin
    TVExplorer.Canvas.Font.Style := TVExplorer.Canvas.Font.Style - [fsBold];
  end;
end;

procedure TvgrReportDesignerForm.aReportSheetPropertiesUpdate(
  Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrReportDesignerForm.aReportSheetPropertiesExecute(
  Sender: TObject);
begin
  TvgrReportSheetFormatForm.Create(nil).Execute(WorkbookGrid.ActiveWorksheet as TvgrReportTemplateWorksheet);
end;

initialization

  RegisterScriptEvent(TvgrBand, 'Events', 'BeforeGenerateScriptProcName', vgrsptProcedure, 'Band,val');
  RegisterScriptEvent(TvgrBand, 'Events', 'AfterGenerateScriptProcName', vgrsptProcedure, 'Band,val');

  RegisterScriptEvent(TvgrDataBand, 'Events', 'BeforeGenerateScriptProcName', vgrsptProcedure, 'Band,val');
  RegisterScriptEvent(TvgrDataBand, 'Events', 'AfterGenerateScriptProcName', vgrsptProcedure, 'Band,val');

  RegisterScriptEvent(TvgrDetailBand, 'Events', 'BeforeGenerateScriptProcName', vgrsptProcedure, 'Band,val');
  RegisterScriptEvent(TvgrDetailBand, 'Events', 'AfterGenerateScriptProcName', vgrsptProcedure, 'Band,val');
  RegisterScriptEvent(TvgrDetailBand, 'Events', 'InitDatasetScriptProcName', vgrsptProcedure, 'Band,val,Initialized,var');

  RegisterScriptEvent(TvgrGroupBand, 'Events', 'BeforeGenerateScriptProcName', vgrsptProcedure, 'Band,val');
  RegisterScriptEvent(TvgrGroupBand, 'Events', 'AfterGenerateScriptProcName', vgrsptProcedure, 'Band,val');
  
  RegisterScriptEvent(TvgrReportTemplateWorksheet, 'Events', 'BeforeGenerateScriptProcName', vgrsptProcedure, 'TemplateWorksheet,val');
  RegisterScriptEvent(TvgrReportTemplateWorksheet, 'Events', 'AfterGenerateScriptProcName', vgrsptProcedure, 'TemplateWorksheet,val,WorkbookWorksheet,val');
  RegisterScriptEvent(TvgrReportTemplateWorksheet, 'Events', 'AddWorkbookWorksheetScriptProcName', vgrsptProcedure, 'TemplateWorksheet,val,WorkbookWorksheet,val');
  RegisterScriptEvent(TvgrReportTemplateWorksheet, 'Events', 'CustomGenerateScriptProcName', vgrsptProcedure, 'TemplateWorksheet,val,WorkbookWorksheet,val,Done,var');

end.

