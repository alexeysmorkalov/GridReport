{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains the built-in workbook designer, the workbook designer can be shown as modal or not.}
unit vgr_WorkbookDesigner;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ActnList, ComCtrls, ToolWin, ImgList, StdCtrls,

  vgr_ControlBar, vgr_WorkbookGrid, vgr_DataStorage, vgr_CommonClasses,
  vgr_Dialogs, vgr_CellPropertiesDialog, vgr_ScriptEdit,
  vgr_FontComboBox, vgr_DataStorageTypes, vgr_ColorButton,
  vgr_PageSetupDialog, vgr_WorkbookPreviewForm, vgr_Consts, vgr_Form,
  vgr_FormLocalizer;

type
  TvgrCustomWorkbookDesigner = class;
  TvgrCustomWorkbookDesignerClass = class of TvgrCustomWorkbookDesigner;
  TvgrWorkbookDesignerFormClass = class of TvgrWorkbookDesignerForm;

  /////////////////////////////////////////////////
  //
  // TvgrWorkbookDesignerForm
  //
  /////////////////////////////////////////////////
{Implements a form for editing the workbook contents.
The form can be shown as modal whether or not.}
  TvgrWorkbookDesignerForm = class(TvgrForm, IvgrWorkbookHandler)
    MainMenu: TMainMenu;
    CommonActions: TActionList;
    aOpen: TAction;
    aNew: TAction;
    aSave: TAction;
    aSaveAs: TAction;
    aExit: TAction;
    File1: TMenuItem;
    New1: TMenuItem;
    Open1: TMenuItem;
    N1: TMenuItem;
    Save1: TMenuItem;
    Saveas1: TMenuItem;
    N2: TMenuItem;
    Exit1: TMenuItem;
    View1: TMenuItem;
    aViewHeaders: TAction;
    aViewFormula: TAction;
    aOptions: TAction;
    Formula1: TMenuItem;
    Headers1: TMenuItem;
    mToolbars: TMenuItem;
    N3: TMenuItem;
    Options1: TMenuItem;
    Insert1: TMenuItem;
    aInsertRow: TAction;
    aInsertColumn: TAction;
    aInsertSheet: TAction;
    Row1: TMenuItem;
    Column1: TMenuItem;
    Sheet1: TMenuItem;
    Format1: TMenuItem;
    aFormatCell: TAction;
    aRowHeight: TAction;
    aRowAutoHeight: TAction;
    aRowHide: TAction;
    aRowShow: TAction;
    aColWidth: TAction;
    aColAutoWidth: TAction;
    aColHide: TAction;
    aColShow: TAction;
    Cell1: TMenuItem;
    Row2: TMenuItem;
    Column2: TMenuItem;
    Height1: TMenuItem;
    Autowidth1: TMenuItem;
    Hide1: TMenuItem;
    Show1: TMenuItem;
    Width1: TMenuItem;
    Autowidth2: TMenuItem;
    Hide2: TMenuItem;
    Show2: TMenuItem;
    Sheet2: TMenuItem;
    aSheetRename: TAction;
    Rename1: TMenuItem;
    Edit1: TMenuItem;
    aDelRow: TAction;
    aDelCol: TAction;
    aDelSheet: TAction;
    Deleterow1: TMenuItem;
    Deletecolumn1: TMenuItem;
    Delsheet1: TMenuItem;
    TopControlBar: TvgrControlBar;
    LeftControlBar: TvgrControlBar;
    RightControlBar: TvgrControlBar;
    StatusBar: TStatusBar;
    PageControl: TPageControl;
    PGrid: TTabSheet;
    WorkbookGrid: TvgrWorkbookGrid;
    WorkbookGridActions: TActionList;
    ScriptEditActions: TActionList;
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    ViewActions: TActionList;
    CellPropertiesDialog: TvgrCellPropertiesDialog;
    aCopy: TAction;
    aPaste: TAction;
    aCut: TAction;
    N6: TMenuItem;
    Cut1: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    tbStandard: TToolBar;
    ImageList: TImageList;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    tbFormatting: TToolBar;
    ToolButton8: TToolButton;
    edFontSize: TComboBox;
    edFontName: TvgrFontComboBox;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    bFillBackColor: TToolButton;
    bFontColor: TToolButton;
    aFontBold: TAction;
    aFontItalic: TAction;
    aFontUnderline: TAction;
    aAlignLeft: TAction;
    aAlignHorzCenter: TAction;
    aAlignRight: TAction;
    aMerge: TAction;
    aFormatCurrency: TAction;
    aFormatPercent: TAction;
    aFormatSeparator: TAction;
    aFormatDownDec: TAction;
    aFormatUpDec: TAction;
    aFillBackColor: TAction;
    aFontColor: TAction;
    PMFontColor: TPopupMenu;
    PMFillBackColor: TPopupMenu;
    PopupMenu1: TPopupMenu;
    Clear1: TMenuItem;
    aClearAll: TAction;
    aClearDisplayFormat: TAction;
    aClearAllFormats: TAction;
    aClearValue: TAction;
    aClearBorders: TAction;
    All1: TMenuItem;
    Value1: TMenuItem;
    Format2: TMenuItem;
    Displayformats1: TMenuItem;
    N7: TMenuItem;
    Borders1: TMenuItem;
    ControlBarManager: TvgrControlBarManager;
    PageSetupDialog: TvgrPageSetupDialog;
    aPageSetup: TAction;
    N8: TMenuItem;
    Pagesetup1: TMenuItem;
    WorkbookPreviewer: TvgrWorkbookPreviewer;
    aPrintPreview: TAction;
    N9: TMenuItem;
    Printpreview1: TMenuItem;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    WorkbookDesignerLocalizer: TvgrFormLocalizer;
    procedure aNewExecute(Sender: TObject);
    procedure aOpenExecute(Sender: TObject);
    procedure aSaveExecute(Sender: TObject);
    procedure aSaveUpdate(Sender: TObject);
    procedure aSaveAsExecute(Sender: TObject);
    procedure aExitExecute(Sender: TObject);
    procedure aViewHeadersUpdate(Sender: TObject);
    procedure aViewHeadersExecute(Sender: TObject);
    procedure aViewFormulaUpdate(Sender: TObject);
    procedure aViewFormulaExecute(Sender: TObject);
    procedure aDelRowExecute(Sender: TObject);
    procedure aDelColExecute(Sender: TObject);
    procedure aDelSheetExecute(Sender: TObject);
    procedure WorkbookGridActionsUpdate(Action: TBasicAction;
      var Handled: Boolean);
    procedure aInsertRowExecute(Sender: TObject);
    procedure aInsertSheetExecute(Sender: TObject);
    procedure aInsertColumnExecute(Sender: TObject);
    procedure aFormatCellExecute(Sender: TObject);
    procedure aRowHeightExecute(Sender: TObject);
    procedure aColWidthExecute(Sender: TObject);
    procedure aRowHideExecute(Sender: TObject);
    procedure aRowShowExecute(Sender: TObject);
    procedure aColHideExecute(Sender: TObject);
    procedure aColShowExecute(Sender: TObject);
    procedure aSheetRenameExecute(Sender: TObject);
    procedure aFontBoldUpdate(Sender: TObject);
    procedure aFontBoldExecute(Sender: TObject);
    procedure aFontItalicUpdate(Sender: TObject);
    procedure aFontItalicExecute(Sender: TObject);
    procedure aFontUnderlineUpdate(Sender: TObject);
    procedure aFontUnderlineExecute(Sender: TObject);
    procedure aAlignLeftUpdate(Sender: TObject);
    procedure aAlignLeftExecute(Sender: TObject);
    procedure aMergeUpdate(Sender: TObject);
    procedure aMergeExecute(Sender: TObject);
    procedure aFormatCurrencyExecute(Sender: TObject);
    procedure aFormatPercentExecute(Sender: TObject);
    procedure aFormatSeparatorExecute(Sender: TObject);
    procedure aFormatDownDecExecute(Sender: TObject);
    procedure aFormatUpDecExecute(Sender: TObject);
    procedure PMFontColorPopup(Sender: TObject);
    procedure aFillBackColorUpdate(Sender: TObject);
    procedure aCutUpdate(Sender: TObject);
    procedure aCopyUpdate(Sender: TObject);
    procedure aPasteUpdate(Sender: TObject);
    procedure WorkbookGridSelectionChange(Sender: TObject);
    procedure edFontNameClick(Sender: TObject);
    procedure edFontSizeKeyPress(Sender: TObject; var Key: Char);
    procedure edFontSizeClick(Sender: TObject);
    procedure aRowAutoHeightExecute(Sender: TObject);
    procedure aColAutoWidthExecute(Sender: TObject);
    procedure aOpenUpdate(Sender: TObject);
    procedure aDelRowUpdate(Sender: TObject);
    procedure aFormatCurrencyUpdate(Sender: TObject);
    procedure aClearAllUpdate(Sender: TObject);
    procedure aClearAllExecute(Sender: TObject);
    procedure aClearAllFormatsExecute(Sender: TObject);
    procedure aClearDisplayFormatExecute(Sender: TObject);
    procedure aClearValueExecute(Sender: TObject);
    procedure aClearBordersExecute(Sender: TObject);
    procedure WorkbookGridCellProperties(Sender: TObject);
    procedure aCutExecute(Sender: TObject);
    procedure aCopyExecute(Sender: TObject);
    procedure aPasteExecute(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure View1Click(Sender: TObject);
    procedure aPageSetupUpdate(Sender: TObject);
    procedure aPageSetupExecute(Sender: TObject);
    procedure aPrintPreviewUpdate(Sender: TObject);
    procedure aPrintPreviewExecute(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    FDesigner: TvgrCustomWorkbookDesigner;
    FDocCaption: string;
    FPaletteForm: TvgrColorPaletteForm;
    FFontOtherColor: TColor;
    FFillOtherColor: TColor;
    FSettingsRestored: Boolean;

    function GetWorkbook: TvgrWorkbook;
    procedure SetDesigner(Value: TvgrCustomWorkbookDesigner);
    procedure SetDocCaption(const Value: string);
    function GetActiveWorksheet: TvgrWorksheet;
    function GetSelectionFormat: IvgrRangesFormat;
    procedure FontColorSetColor(Sender: TObject; AColor: TColor);
    procedure FillBackColorSetColor(Sender: TObject; AColor: TColor);
    procedure UpdateFont;
    procedure UpdateToolbarsMenu;
  protected
    // IvgrWorkbookHandler
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); virtual;
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); virtual;
    procedure DeleteItemInterface(AItem: IvgrWBListItem); virtual;

    procedure SetWorkbook(Value: TvgrWorkbook); virtual;

    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
    procedure OnCloseEvent(Sender: TObject; var Action: TCloseAction); override;
    function CheckSaveChanges: Boolean;
    function DoNew: Boolean;
    function DoOpen: Boolean;
    function DoSave(var ANeedSaveAs: Boolean; var ASaveFileName: string): Boolean;
    function DoSaveAs(var ASaveFileName, ASaveInitialDir: string): Boolean;

    function GetModified: Boolean; virtual;
    procedure SetModified(Value: Boolean); virtual;

    function GetNewDocCaption: string; virtual;
    function GetTabSheetCaption: string; virtual;
    function GetFormCaption: string; virtual;
    procedure SetupDialogs(AOpenDialog: TOpenDialog; ASaveDialog: TSaveDialog); virtual;

    function IsCutAllowed: Boolean; virtual;
    function IsPasteAllowed: Boolean; virtual;
    function IsCopyAllowed: Boolean; virtual;

    procedure BuildStatusBar; virtual;
    procedure UpdateStatusBar; virtual;

    procedure OnToolbarClick(Sender: TObject); virtual;

    procedure InternalOnSave; virtual;
    procedure InternalOnLoad; virtual;

    procedure SetInplaceEditValue(Sender: TObject; const AValue: String; ARange: IvgrRange; ANewRange: Boolean); virtual;

    function GetFixupShortcuts: Boolean; override;
    function GetStoreSettingsMethod: TvgrStoreSettingsMethod; override;

    procedure RestoreSettings; override;
    procedure SaveSettings; override;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
    function CloseQuery: Boolean; override;

{Cuts the current selection to the clipboard.}
    procedure CutToClipboard; virtual;
{Copies the current selection to the clipboard.}
    procedure CopyToClipboard; virtual;
{Pastes data from the clipboard to the current position.}
    procedure PasteFromClipboard; virtual;

{Sets or gets the edited TvgrWorkbook object.}
    property Workbook: TvgrWorkbook read GetWorkbook write SetWorkbook;
{Gets the currently selected worksheet.}
    property ActiveWorksheet: TvgrWorksheet read GetActiveWorksheet;
{Returns the TvgrCustomWorkbookDesigner object that creates this form or nil.}
    property Designer: TvgrCustomWorkbookDesigner read FDesigner write SetDesigner;
{Returns the currently edited file name.}
    property DocCaption: string read FDocCaption write SetDocCaption;
{Returns the IvgrRangesFormat interface, that describes properties of the current selection.
See also:
  IvgrRangesFormat}
    property SelectionFormat: IvgrRangesFormat read GetSelectionFormat;

{Sets or gets the boolean value indicates whether the workbook is modified.}
    property Modified: Boolean read GetModified write SetModified;
  end;

{Used for firing information about some action in the designer
(the opening of the existing file, the creating of the new file).
Syntax:
  TvgrDesignerActionEvent = procedure (Sender: TObject; var ADone: Boolean) of object;
Parameters:
  Sender - the TvgrCustomWorkbookDesigner object.
  ADone - if true will be returned, the designer will not undertake any additional operations.}
  TvgrDesignerActionEvent = procedure (Sender: TObject; var ADone: Boolean) of object;
{Used for firing information about "Save" action in the designer.
Syntax:
  TvgrDesignerSaveEvent = procedure (Sender: TObject; var ADone: Boolean; var ANeedSaveAs: Boolean; var ASaveFileName: string) of object;
Parameters:
  Sender - the TvgrCustomWorkbookDesigner object.
  ADone - if true will be returned, the designer will not undertake any additional operations.
  ANeedSaveAs - if true will be returned, the designer show the "Save as..." dialog to select the file name.
  ASaveFileName - you can return the name of the file in this parameter.}
  TvgrDesignerSaveEvent = procedure (Sender: TObject; var ADone: Boolean; var ANeedSaveAs: Boolean; var ASaveFileName: string) of object;
{Used for firing information about "Save as" action in the designer.
Syntax:
  TvgrDesignerSaveAsEvent = procedure (Sender: TObject; var ADone: Boolean; var ASaveFileName, ASaveInitialDir: string) of object;
Parameters:
  Sender - the TvgrCustomWorkbookDesigner object.
  ADone - if true will be returned, the designer will not undertake any additional operations.
  ASaveFileName - you can return the initial name of the file in this parameter.
  ASaveInitialDir - you can return the initial directory in this parameter.}
  TvgrDesignerSaveAsEvent = procedure (Sender: TObject; var ADone: Boolean; var ASaveFileName, ASaveInitialDir: string) of object;
  /////////////////////////////////////////////////
  //
  // TvgrCustomWorkbookDesigner
  //
  /////////////////////////////////////////////////
{Non visual component, realizes the designer of the workvbook (TvgrWorkbook object).
The designer does not appear at runtime until it is activated by a call to the Design method.
The designer can be shown as modal whether or not.
See also:
  TvgrWorkbookDesignerForm}
  TvgrCustomWorkbookDesigner = class(TComponent)
  private
    FWorkbook: TvgrWorkbook;
    FOnNew: TvgrDesignerActionEvent;
    FOnOpen: TvgrDesignerActionEvent;
    FOnSave: TvgrDesignerSaveEvent;
    FOnSaveAs: TvgrDesignerSaveAsEvent;
    FOnChange: TNotifyEvent;
    FDocumentName: string;
    FForm: TvgrWorkbookDesignerForm;
    procedure SetWorkbook(Value: TvgrWorkbook);
  protected
    function GetDesignerFormClass: TvgrWorkbookDesignerFormClass; virtual;
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
    function DoNew: Boolean; virtual;
    function DoOpen: Boolean; virtual;
    function DoSave(var ANeedSaveAs: Boolean; var ASaveFileName: string): Boolean; virtual;
    function DoSaveAs(var ASaveFileName, ASaveInitialDir: string): Boolean; virtual;
    procedure DoChange;
    procedure SetDocumentName(const Value: string);

    property Form: TvgrWorkbookDesignerForm read FForm;
  public
{Creates instance of the TvgrCustomWorkbookDesigner.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees instance of a TvgrCustomWorkbookDesigner.}
    destructor Destroy; override;

{Design opens the designer form.
Parameters:
  AModal - if true the form shown as modal.}
    procedure Design(AModal: Boolean);
{Showes the created designer form.
Parameters:
  AModal - if true the form shown as modal.}
    procedure ActivateForm(AModal: Boolean);
{Closes the designer form.}
    procedure CloseForm;

{Sets or gets the edited workbook(TvgrWorkbook object).}
    property Workbook: TvgrWorkbook read FWorkbook write SetWorkbook;
{Sets or gets the currently edited file name.}
    property DocumentName: string read FDocumentName write SetDocumentName;
  published
{This event occurs when user creates the new workbook file.
See also:
  TvgrDesignerActionEvent}
    property OnNew: TvgrDesignerActionEvent read FOnNew write FOnNew;
{This event occurs when user opens the existing workbook file.
See also:
  TvgrDesignerActionEvent}
    property OnOpen: TvgrDesignerActionEvent read FOnOpen write FOnOpen;
{This event occurs when user saves the workbook to the file.
See also:
  TvgrDesignerSaveEvent}
    property OnSave: TvgrDesignerSaveEvent read FOnSave write FOnSave;
{This event occurs when user saves the workbook to the file and
the workbook file name is not specified.
See also:
  DocumentName, TvgrDesignerSaveEvent}
    property OnSaveAs: TvgrDesignerSaveAsEvent read FOnSaveAs write FOnSaveAs;
{This event occurs when workbook are changed.}
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWorkbookDesigner
  //
  /////////////////////////////////////////////////
{Non visual component, realizes the designer of the workvbook (TvgrWorkbook object).
The designer does not appear at runtime until it is activated by a call to the Design method.
The designer can be shown as modal whether or not.
See also:
  TvgrWorkbookDesignerForm}
  TvgrWorkbookDesigner = class(TvgrCustomWorkbookDesigner)
  published
{Sets or gets the edited workbook(TvgrWorkbook object).}
    property Workbook;
  end;

implementation

uses
  vgr_Functions, vgr_StringIDs, vgr_RowColSizeDialog,
  vgr_SheetFormatDialog, vgr_PageProperties, vgr_Localize;

{$R *.dfm}
{$R ..\res\vgr_WorkbookDesignerStrings.res}

/////////////////////////////////////////////////
//
// TvgrCustomWorkbookDesigner
//
/////////////////////////////////////////////////
constructor TvgrCustomWorkbookDesigner.Create(AOwner: TComponent);
begin
  inherited;
end;

destructor TvgrCustomWorkbookDesigner.Destroy;
begin
  FreeAndNil(FForm);
  inherited;
end;

procedure TvgrCustomWorkbookDesigner.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if (AOperation = opRemove) and (AComponent = Workbook) then
    Workbook := nil;
end;

function TvgrCustomWorkbookDesigner.GetDesignerFormClass: TvgrWorkbookDesignerFormClass;
begin
  Result := TvgrWorkbookDesignerForm;
end;

procedure TvgrCustomWorkbookDesigner.SetWorkbook(Value: TvgrWorkbook);
begin
  if FWorkbook <> Value then
  begin
    FWorkbook := Value;
    if FForm <> nil then
      FForm.Workbook := Value;
  end;
end;

procedure TvgrCustomWorkbookDesigner.Design(AModal: Boolean);
begin
  if Workbook <> nil then
  begin
    if FForm = nil then
    begin
      FForm := GetDesignerFormClass.Create(Application);
      FForm.Workbook := Workbook;
      FForm.Designer := Self;
      FForm.DocCaption := DocumentName; 
    end;
    ActivateForm(AModal);
  end;
end;

procedure TvgrCustomWorkbookDesigner.ActivateForm(AModal: Boolean);
begin
  if FForm <> nil then
  begin
    if not FForm.Visible then
    begin
      if AModal then
        FForm.ShowModal
      else
      begin
        FForm.Show;
        if FForm.WindowState = wsMinimized then
          FForm.WindowState := wsNormal;
        FForm.SetFocus;
      end;
    end;
  end;
end;

procedure TvgrCustomWorkbookDesigner.CloseForm;
begin
  if FForm <> nil then
    FForm.Close;
end;

function TvgrCustomWorkbookDesigner.DoNew: Boolean;
begin
  Result := False;
  if Assigned(FOnNew) then
    FOnNew(Self, Result);
end;

function TvgrCustomWorkbookDesigner.DoOpen: Boolean;
begin
  Result := False;
  if Assigned(FOnOpen) then
    FOnOpen(Self, Result);
end;

function TvgrCustomWorkbookDesigner.DoSave(var ANeedSaveAs: Boolean; var ASaveFileName: string): Boolean;
begin
  Result := False;
  if Assigned(FOnSave) then
    FOnSave(Self, Result, ANeedSaveAs, ASaveFileName);
end;

function TvgrCustomWorkbookDesigner.DoSaveAs(var ASaveFileName, ASaveInitialDir: string): Boolean;
begin
  Result := False;
  if Assigned(FOnSaveAs) then
    FOnSaveAs(Self, Result, ASaveFileName, ASaveInitialDir);
end;

procedure TvgrCustomWorkbookDesigner.DoChange;
begin
  if Assigned(FOnChange) then
    FOnChange(Self);
end;

procedure TvgrCustomWorkbookDesigner.SetDocumentName(const Value: string);
begin
  if FDocumentName <> Value then
  begin
    FDocumentName := Value;
    if Form <> nil then
      Form.DocCaption := DocumentName;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrWorkbookDesignerForm
//
/////////////////////////////////////////////////
function TvgrWorkbookDesignerForm.GetFixupShortcuts: Boolean;
begin
  Result := True;
end;

function TvgrWorkbookDesignerForm.GetStoreSettingsMethod: TvgrStoreSettingsMethod;
begin
  Result := DefaultStoreSettingsMethod;
end;

procedure TvgrWorkbookDesignerForm.RestoreSettings;
begin
  inherited;
end;

procedure TvgrWorkbookDesignerForm.SaveSettings;
begin
  inherited;
  ControlBarManager.WriteToStorage(SettingsStorage, StorageSection + '_TB');
end;

function TvgrWorkbookDesignerForm.GetWorkbook: TvgrWorkbook;
begin
  if WorkbookGrid = nil then
    Result := nil
  else
    Result := TvgrWorkbook(WorkbookGrid.Workbook);
end;

procedure TvgrWorkbookDesignerForm.SetWorkbook(Value: TvgrWorkbook);
begin
  if WorkbookGrid <> nil then
  begin
    if Workbook <> nil then
      Workbook.DisconnectHandler(Self);

    WorkbookGrid.Workbook := Value;

    if Workbook <> nil then
      with Workbook do
      begin
        ConnectHandler(Self);
        if WorksheetsCount = 0 then
          AddWorksheet;
        WorkbookGrid.ActiveWorksheetIndex := 0;
        DocCaption := GetNewDocCaption;
        UpdateStatusBar;
      end;
  end;
end;

procedure TvgrWorkbookDesignerForm.SetDesigner(Value: TvgrCustomWorkbookDesigner);
begin
  if FDesigner <> Value then
  begin
    FDesigner := Value;
  end;
end;

procedure TvgrWorkbookDesignerForm.SetDocCaption(const Value: string);
begin
  if FDocCaption <> Value then
  begin
    FDocCaption := Value;
    Caption := GetFormCaption;
  end;
end;

function TvgrWorkbookDesignerForm.GetActiveWorksheet: TvgrWorksheet;
begin
  if Workbook <> nil then
    Result := WorkbookGrid.ActiveWorksheet
  else
    Result := nil;
end;

function TvgrWorkbookDesignerForm.GetSelectionFormat: IvgrRangesFormat;
begin
  Result := WorkbookGrid.SelectionFormat;
  if Result <> nil then
    Result.ForceBeginUpdateEndUpdate := True;
end;

procedure TvgrWorkbookDesignerForm.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if (AOperation = opRemove) and (AComponent = Workbook) then
    Workbook := nil;
end;

procedure TvgrWorkbookDesignerForm.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
end;

procedure TvgrWorkbookDesignerForm.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  UpdateStatusBar;
  if Designer <> nil then
    Designer.DoChange;
end;

procedure TvgrWorkbookDesignerForm.DeleteItemInterface(AItem: IvgrWBListItem);
begin
end;

procedure TvgrWorkbookDesignerForm.OnCloseEvent(Sender: TObject; var Action: TCloseAction);
begin
  InternalOnSave;
  inherited;
end;

function TvgrWorkbookDesignerForm.CloseQuery: Boolean;
begin
  Result := inherited CloseQuery;
  if Result and not (csDesigning in Workbook.ComponentState) then
    Result := CheckSaveChanges;
end;

procedure TvgrWorkbookDesignerForm.UpdateToolbarsMenu;
var
  I: Integer;
begin
  ClearMenuItem(mToolbars);
  for I := 0 to ComponentCount - 1 do
    if Components[I] is TToolbar then
      with TToolbar(Components[I]) do
        AddMenuItem(MainMenu, mToolbars, Caption, OnToolbarClick, Integer(Self.Components[I]), '', True, Visible);
end;

procedure TvgrWorkbookDesignerForm.AfterConstruction;
begin
  inherited;
  SetupDialogs(OpenDialog, SaveDialog);
  PGrid.Caption := GetTabSheetCaption;
  PageControl.ActivePageIndex := 0;
  ActiveControl := WorkbookGrid;
  BuildStatusBar;
  UpdateStatusBar;

  WorkbookGrid.OnSetInplaceEditValue := SetInplaceEditValue;
end;

procedure TvgrWorkbookDesignerForm.BeforeDestruction;
begin
  inherited;
  Workbook := nil;
  if Designer <> nil then
    Designer.FForm := nil;
end;

procedure TvgrWorkbookDesignerForm.SetupDialogs(AOpenDialog: TOpenDialog; ASaveDialog: TSaveDialog);
begin
  with AOpenDialog do
  begin
    DefaultExt := vgrLoadStr(svgrid_vgr_WorkbookDesigner_WorkbookOpenSaveDefaultExt);
    Filter := vgrLoadStr(svgrid_vgr_WorkbookDesigner_WorkbookOpenSaveDialogFilter);
    Options := [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing];
  end;
  with ASaveDialog do
  begin
    DefaultExt := vgrLoadStr(svgrid_vgr_WorkbookDesigner_WorkbookOpenSaveDefaultExt);
    Filter := vgrLoadStr(svgrid_vgr_WorkbookDesigner_WorkbookOpenSaveDialogFilter);
    Options := [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing];
  end;
end;

function TvgrWorkbookDesignerForm.CheckSaveChanges: Boolean;
begin
  Result := not Modified;
  if not Result then
  begin
    case MBoxFmt(vgrLoadStr(svgrid_vgr_WorkbookDesigner_SaveChanges), MB_YESNOCANCEL or MB_ICONQUESTION, [DocCaption]) of
      IDYES:
        begin
          aSave.Execute;
          Result := True;
        end;
      IDNO:
        Result := True;
      IDCANCEL: ;
    end;
  end;
end;

procedure TvgrWorkbookDesignerForm.aNewExecute(Sender: TObject);
begin
  if CheckSaveChanges then
    if not DoNew then
    begin
      Workbook.Clear;
      Workbook.AddWorksheet;
      WorkbookGrid.ActiveWorksheetIndex := 0;
      Modified := False;
      DocCaption := GetNewDocCaption;
    end;
end;

function TvgrWorkbookDesignerForm.DoNew: Boolean;
begin
  Result := (Designer <> nil) and Designer.DoNew;
end;

function TvgrWorkbookDesignerForm.DoOpen: Boolean;
begin
  Result := (Designer <> nil) and Designer.DoOpen;
end;

function TvgrWorkbookDesignerForm.DoSave(var ANeedSaveAs: Boolean; var ASaveFileName: string): Boolean;
begin
  Result := (Designer <> nil) and Designer.DoSave(ANeedSaveAs, ASaveFileName);
end;

function TvgrWorkbookDesignerForm.DoSaveAs(var ASaveFileName, ASaveInitialDir: string): Boolean;
begin
  Result := (Designer <> nil) and Designer.DoSaveAs(ASaveFileName, ASaveInitialDir);
end;

function TvgrWorkbookDesignerForm.GetModified: Boolean;
begin
  Result := (Workbook <> nil) and Workbook.Modified;
end;

procedure TvgrWorkbookDesignerForm.SetModified(Value: Boolean);
begin
  Workbook.Modified := Value;
  UpdateStatusBar;
end;

function TvgrWorkbookDesignerForm.GetNewDocCaption: string;
begin
  Result := vgrLoadStr(svgrid_vgr_WorkbookDesigner_NewDocCaption);
end;

function TvgrWorkbookDesignerForm.GetTabSheetCaption: string;
begin
  Result := vgrLoadStr(svgrid_vgr_WorkbookDesigner_WorkbookTabSheetCaption);
end;

function TvgrWorkbookDesignerForm.GetFormCaption: string;
begin
  Result := Format('%s - [%s]', [Application.Title, FDocCaption])
end;

procedure TvgrWorkbookDesignerForm.aOpenExecute(Sender: TObject);
begin
  if CheckSaveChanges then
    if not DoOpen then
    begin
      OpenDialog.FileName := DocCaption;
      if OpenDialog.Execute then
      begin
        Workbook.LoadFromFile(OpenDialog.FileName);
        InternalOnLoad;
        WorkbookGrid.ActiveWorksheetIndex := 0;
        Modified := False;
        DocCaption := OpenDialog.FileName;
      end;
    end;
end;

procedure TvgrWorkbookDesignerForm.aSaveUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit and Modified;
end;

procedure TvgrWorkbookDesignerForm.aSaveExecute(Sender: TObject);
var
  ANeedSaveAs: Boolean;
  ASaveFileName: string;
begin
  ANeedSaveAs := (DocCaption = '') or (DocCaption = GetNewDocCaption);
  ASaveFileName := DocCaption;
  if not DoSave(ANeedSaveAs, ASaveFileName) then
  begin
    if ANeedSaveAs then
      aSaveAs.Execute
    else
    begin
      InternalOnSave;
      Workbook.SaveToFile(ASaveFileName);
      Modified := False;
    end;
  end;
end;

procedure TvgrWorkbookDesignerForm.aSaveAsExecute(Sender: TObject);
var
  ASaveFileName, ASaveInitialDir: string;
begin
  ASaveFileName := DocCaption;
  ASaveInitialDir := SaveDialog.InitialDir;
  if not DoSaveAs(ASaveFileName, ASaveInitialDir) then
  begin
    SaveDialog.FileName := ASaveFileName;
    SaveDialog.InitialDir := ASaveInitialDir;
    if SaveDialog.Execute then
    begin
      InternalOnSave;
      Workbook.SaveToFile(SaveDialog.FileName);
      DocCaption := SaveDialog.FileName;
      Modified := False;
    end;
  end;
end;

procedure TvgrWorkbookDesignerForm.aExitExecute(Sender: TObject);
begin
  Close;
end;

procedure TvgrWorkbookDesignerForm.aViewHeadersUpdate(Sender: TObject);
begin
  TAction(Sender).Checked := WorkbookGrid.OptionsCols.Visible and WorkbookGrid.OptionsRows.Visible;
end;

procedure TvgrWorkbookDesignerForm.aViewHeadersExecute(Sender: TObject);
begin
  with WorkbookGrid do
  begin
    OptionsCols.Visible := not OptionsCols.Visible;
    OptionsRows.Visible := not OptionsRows.Visible;
  end;
end;

procedure TvgrWorkbookDesignerForm.aViewFormulaUpdate(Sender: TObject);
begin
  TAction(Sender).Checked := WorkbookGrid.OptionsFormulaPanel.Visible;
end;

procedure TvgrWorkbookDesignerForm.aViewFormulaExecute(Sender: TObject);
begin
  WorkbookGrid.OptionsFormulaPanel.Visible := not WorkbookGrid.OptionsFormulaPanel.Visible;
end;

procedure TvgrWorkbookDesignerForm.aDelRowExecute(Sender: TObject);
begin
  WorkbookGrid.DeleteRows;
end;

procedure TvgrWorkbookDesignerForm.aDelColExecute(Sender: TObject);
begin
  WorkbookGrid.DeleteColumns;
end;

procedure TvgrWorkbookDesignerForm.aDelSheetExecute(Sender: TObject);
begin
  if Workbook.WorksheetsCount < 2 then
    MBox(vgrLoadStr(svgrid_vgr_WorkbookDesigner_DeleteLastSheetErrorMessage), MB_OK or MB_ICONEXCLAMATION)
  else
    if ActiveWorksheet.RangesList.Count <> 0 then
    begin
      if MBox(vgrLoadStr(svgrid_vgr_WorkbookDesigner_DeleteNotEmptySheetWarning), MB_OKCANCEL or MB_ICONEXCLAMATION) = ID_OK then
        Workbook.DeleteWorksheet(WorkbookGrid.ActiveWorksheetIndex);
    end
    else
      Workbook.DeleteWorksheet(WorkbookGrid.ActiveWorksheetIndex);
end;

procedure TvgrWorkbookDesignerForm.WorkbookGridActionsUpdate(
  Action: TBasicAction; var Handled: Boolean);
begin
  if Action is TAction then
  begin
    edFontName.Enabled := not WorkbookGrid.IsInplaceEdit;
    edFontSize.Enabled := not WorkbookGrid.IsInplaceEdit;
  end;
end;

procedure TvgrWorkbookDesignerForm.aInsertRowExecute(Sender: TObject);
begin
  WorkbookGrid.InsertRow;
end;

procedure TvgrWorkbookDesignerForm.aInsertColumnExecute(Sender: TObject);
begin
  WorkbookGrid.InsertColumn;
end;

procedure TvgrWorkbookDesignerForm.aInsertSheetExecute(Sender: TObject);
begin
  WorkbookGrid.InsertWorksheet;
end;

procedure TvgrWorkbookDesignerForm.aFormatCellExecute(Sender: TObject);
begin
  CellPropertiesDialog.Execute(ActiveWorksheet, WorkbookGrid.SelectionRects);
end;

procedure TvgrWorkbookDesignerForm.aRowHeightExecute(Sender: TObject);
var
  AHeight: Integer;
begin
  AHeight := WorkbookGrid.GetSelectedRowsHeight;
  if TvgrRowColSizeForm.Create(nil).Execute(vgrLoadStr(svgrid_vgr_WorkbookDesigner_RowHeightFormCaption), vgrLoadStr(svgrid_vgr_WorkbookDesigner_RowHeightFormText), AHeight) then
    WorkbookGrid.SetSelectedRowsHeight(AHeight);
end;

procedure TvgrWorkbookDesignerForm.aColWidthExecute(Sender: TObject);
var
  AWidth: Integer;
begin
  AWidth := WorkbookGrid.GetSelectedColsWidth;
  if TvgrRowColSizeForm.Create(nil).Execute(vgrLoadStr(svgrid_vgr_WorkbookDesigner_ColWidthFormCaption), vgrLoadStr(svgrid_vgr_WorkbookDesigner_ColWidthFormText), AWidth) then
    WorkbookGrid.SetSelectedColsWidth(AWidth);
end;

procedure TvgrWorkbookDesignerForm.aRowHideExecute(Sender: TObject);
begin
  WorkbookGrid.ShowHideSelectedRows(False);
end;

procedure TvgrWorkbookDesignerForm.aRowShowExecute(Sender: TObject);
begin
  WorkbookGrid.ShowHideSelectedRows(True);
end;

procedure TvgrWorkbookDesignerForm.aColHideExecute(Sender: TObject);
begin
  WorkbookGrid.ShowHideSelectedColumns(False);
end;

procedure TvgrWorkbookDesignerForm.aColShowExecute(Sender: TObject);
begin
  WorkbookGrid.ShowHideSelectedColumns(True);
end;

procedure TvgrWorkbookDesignerForm.aSheetRenameExecute(Sender: TObject);
begin
  TvgrSheetFormatForm.Create(nil).Execute(WorkbookGrid.ActiveWorksheet);
end;

procedure TvgrWorkbookDesignerForm.aFontBoldUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
  TAction(Sender).Checked := (SelectionFormat <> nil) and
                             (vgrrfFontStyleBold in SelectionFormat.HasFlags) and
                             SelectionFormat.FontStyleBold;
end;

procedure TvgrWorkbookDesignerForm.aFontBoldExecute(Sender: TObject);
begin
  Workbook.BeginUpdate;
  try
    SelectionFormat.FontStyleBold := not TAction(Sender).Checked;
  finally
    Workbook.EndUpdate;
  end;
end;

procedure TvgrWorkbookDesignerForm.aFontItalicUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
  TAction(Sender).Checked := (SelectionFormat <> nil) and
                             (vgrrfFontStyleItalic in SelectionFormat.HasFlags) and
                             SelectionFormat.FontStyleItalic;
end;

procedure TvgrWorkbookDesignerForm.aFontItalicExecute(Sender: TObject);
begin
  Workbook.BeginUpdate;
  try
    SelectionFormat.FontStyleItalic := not TAction(Sender).Checked;
  finally
    Workbook.EndUpdate;
  end;
end;

procedure TvgrWorkbookDesignerForm.aFontUnderlineUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
  TAction(Sender).Checked := (SelectionFormat <> nil) and
                             (vgrrfFontStyleUnderline in SelectionFormat.HasFlags) and
                             SelectionFormat.FontStyleUnderline;
end;

procedure TvgrWorkbookDesignerForm.aFontUnderlineExecute(Sender: TObject);
begin
  Workbook.BeginUpdate;
  try
    SelectionFormat.FontStyleUnderline := not TAction(Sender).Checked;
  finally
    Workbook.EndUpdate;
  end;
end;

procedure TvgrWorkbookDesignerForm.aAlignLeftUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
  TAction(Sender).Checked := (SelectionFormat <> nil) and
                             (vgrrfHorzAlign in SelectionFormat.HasFlags) and
                             (SelectionFormat.HorzAlign = TvgrRangeHorzAlign(TAction(Sender).Tag))
end;

procedure TvgrWorkbookDesignerForm.aAlignLeftExecute(Sender: TObject);
begin
  Workbook.BeginUpdate;
  try
    SelectionFormat.HorzAlign := TvgrRangeHorzAlign(TAction(Sender).Tag);
  finally
    Workbook.EndUpdate;
  end;
end;

procedure TvgrWorkbookDesignerForm.aMergeUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
  TAction(Sender).Checked := WorkbookGrid.IsMergeSelected; 
end;

procedure TvgrWorkbookDesignerForm.aMergeExecute(Sender: TObject);
begin
  WorkbookGrid.MergeUnMergeSelection;
end;

procedure TvgrWorkbookDesignerForm.aFormatCurrencyExecute(Sender: TObject);
begin
  if SelectionFormat <> nil then
    SelectionFormat.DisplayFormat := vgr_Functions.GetCurrencyFormat;
end;

procedure TvgrWorkbookDesignerForm.aFormatPercentExecute(Sender: TObject);
begin
  if SelectionFormat <> nil then
    SelectionFormat.DisplayFormat := '0%';
end;

procedure TvgrWorkbookDesignerForm.aFormatSeparatorExecute(Sender: TObject);
begin
  if SelectionFormat <> nil then
    SelectionFormat.DisplayFormat := '#.##';
end;

procedure TvgrWorkbookDesignerForm.aFormatDownDecExecute(Sender: TObject);
begin
  if SelectionFormat <> nil then
  begin
    if (vgrrfDisplayFormat in SelectionFormat.HasFlags) and
       (Pos('0.0', SelectionFormat.DisplayFormat) <> 0) then
      SelectionFormat.DisplayFormat := SelectionFormat.DisplayFormat + '0'
    else
      SelectionFormat.DisplayFormat := '0.0';
  end;
end;

procedure TvgrWorkbookDesignerForm.aFormatUpDecExecute(Sender: TObject);
var
  ADisplayFormat: string;
begin
  if SelectionFormat <> nil then
  begin
    if (vgrrfDisplayFormat in SelectionFormat.HasFlags) and
       (Pos('0.0', SelectionFormat.DisplayFormat) <> 0) then
    begin
      ADisplayFormat := SelectionFormat.DisplayFormat;
      Delete(ADisplayFormat, Length(ADisplayFormat), 1);
      if Length(ADisplayFormat) = 2 then
        Delete(ADisplayFormat, Length(ADisplayFormat), 1);
      SelectionFormat.DisplayFormat := ADisplayFormat;
    end
    else
      SelectionFormat.DisplayFormat := '0.0';
  end;
end;

procedure TvgrWorkbookDesignerForm.FontColorSetColor(Sender: TObject; AColor: TColor);
begin
  SelectionFormat.FontColor := AColor;
end;

procedure TvgrWorkbookDesignerForm.FillBackColorSetColor(Sender: TObject; AColor: TColor);
begin
  SelectionFormat.FillBackColor := AColor;
end;

procedure TvgrWorkbookDesignerForm.PMFontColorPopup(Sender: TObject);
var
  APoint: TPoint;
  AFontColor: TColor;
  AFillColor: TColor;
begin
  with WorkbookGrid do
  begin
    if vgrrfFontColor in SelectionFormat.HasFlags then
      AFontColor := SelectionFormat.FontColor
    else
      AFontColor := clNone;
    if vgrrfFillBackColor in SelectionFormat.HasFlags then
      AFillColor := SelectionFormat.FillBackColor
    else
      AFillColor := clNone;
  end;

  if Sender = PMFontColor then
    begin
      APoint := bFontColor.ClientToScreen(Point(0, 0));
      FPaletteForm := PopupPaletteForm(Self,
                                       APoint.x,
                                       APoint.y + bFontColor.Height,
                                       AFontColor,
                                       FFontOtherColor,
                                       FontColorSetColor,
                                       vgrLoadStr(svgrid_Common_ColorButtonOther),
                                       vgrLoadStr(svgrid_Common_ColorButtonNone))
    end
  else
    if Sender = PMFillBackColor then
      begin
        APoint := bFillBackColor.ClientToScreen(Point(0,0));
        FPaletteForm := PopupPaletteForm(Self,
                                         APoint.x,
                                         APoint.y + bFillBackColor.Height,
                                         AFillColor,
                                         FFillOtherColor,
                                         FillBackColorSetColor,
                                         vgrLoadStr(svgrid_Common_ColorButtonOther),
                                         vgrLoadStr(svgrid_Common_ColorButtonNone))
      end;
end;

procedure TvgrWorkbookDesignerForm.aFillBackColorUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := True;
end;

procedure TvgrWorkbookDesignerForm.aCutUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := IsCutAllowed;
end;

procedure TvgrWorkbookDesignerForm.aCopyUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := IsCopyAllowed;
end;

procedure TvgrWorkbookDesignerForm.aPasteUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := IsPasteAllowed;
end;

procedure TvgrWorkbookDesignerForm.UpdateFont;
begin
  if SelectionFormat <> nil then
    with SelectionFormat do
    begin
      if vgrrfFontName in HasFlags then
        edFontName.ItemIndex := edFontName.Items.IndexOf(FontName)
      else
        edFontName.ItemIndex := edFontName.Items.IndexOf(vgrDefaultRangeFontName);
      if vgrrfFontSize in HasFlags then
        edFontSize.Text := IntToStr(FontSize)
      else
        edFontSize.Text := IntToStr(vgrDefaultRangeFontSize);
    end;
end;

procedure TvgrWorkbookDesignerForm.WorkbookGridSelectionChange(
  Sender: TObject);
begin
  UpdateFont;
end;

procedure TvgrWorkbookDesignerForm.edFontNameClick(Sender: TObject);
begin
  if SelectionFormat <> nil then
    SelectionFormat.FontName := edFontName.Items[edFontName.ItemIndex];
end;

procedure TvgrWorkbookDesignerForm.edFontSizeKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (SelectionFormat <> nil) then
  begin
    SelectionFormat.FontSize := StrToIntDef(edFontSize.Text, vgrDefaultRangeFontSize);
    edFontSize.Text := IntToStr(StrToIntDef(edFontSize.Text, vgrDefaultRangeFontSize));
    Key := #0;
    ActiveControl := WorkbookGrid;
  end;
end;

procedure TvgrWorkbookDesignerForm.edFontSizeClick(Sender: TObject);
begin
  SelectionFormat.FontSize := StrToIntDef(edFontSize.Text, vgrDefaultRangeFontSize);
  ActiveControl := WorkbookGrid;
end;

procedure TvgrWorkbookDesignerForm.aRowAutoHeightExecute(Sender: TObject);
begin
  WorkbookGrid.SetAutoHeightForSelectedRows;
end;

procedure TvgrWorkbookDesignerForm.aColAutoWidthExecute(Sender: TObject);
begin
  WorkbookGrid.SetAutoWidthForSelectedColumns;
end;

procedure TvgrWorkbookDesignerForm.aOpenUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrWorkbookDesignerForm.aDelRowUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrWorkbookDesignerForm.aFormatCurrencyUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrWorkbookDesignerForm.aClearAllUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrWorkbookDesignerForm.aClearAllExecute(Sender: TObject);
begin
  WorkbookGrid.ClearSelection([vgrccBorders, vgrccRanges]);
end;

procedure TvgrWorkbookDesignerForm.aClearAllFormatsExecute(
  Sender: TObject);
begin
  WorkbookGrid.ClearSelection([vgrccFormat]);
end;

procedure TvgrWorkbookDesignerForm.aClearDisplayFormatExecute(
  Sender: TObject);
begin
  WorkbookGrid.ClearSelection([vgrccDisplayFormat]);
end;

procedure TvgrWorkbookDesignerForm.aClearValueExecute(Sender: TObject);
begin
  WorkbookGrid.ClearSelection([vgrccValue]);
end;

procedure TvgrWorkbookDesignerForm.aClearBordersExecute(Sender: TObject);
begin
  WorkbookGrid.ClearSelection([vgrccBorders]);
end;

procedure TvgrWorkbookDesignerForm.WorkbookGridCellProperties(
  Sender: TObject);
begin
  CellPropertiesDialog.Execute(ActiveWorksheet, WorkbookGrid.SelectionRects);
end;

function TvgrWorkbookDesignerForm.IsCutAllowed: Boolean;
begin
  Result := not WorkbookGrid.IsInplaceEdit and WorkbookGrid.IsCutAllowed;
end;

function TvgrWorkbookDesignerForm.IsPasteAllowed: Boolean;
begin
  Result := not WorkbookGrid.IsInplaceEdit and WorkbookGrid.IsPasteAllowed;
end;

function TvgrWorkbookDesignerForm.IsCopyAllowed: Boolean;
begin
  Result := not WorkbookGrid.IsInplaceEdit and WorkbookGrid.IsCopyAllowed;
end;

procedure TvgrWorkbookDesignerForm.CutToClipboard;
begin
  WorkbookGrid.CutToClipboard;
end;

procedure TvgrWorkbookDesignerForm.CopyToClipboard;
begin
  WorkbookGrid.CopyToClipboard;
end;

procedure TvgrWorkbookDesignerForm.PasteFromClipboard;
begin
  WorkbookGrid.PasteFromClipboard;
end;

procedure TvgrWorkbookDesignerForm.aCutExecute(Sender: TObject);
begin
  CutToClipboard;
end;

procedure TvgrWorkbookDesignerForm.aCopyExecute(Sender: TObject);
begin
  CopyToClipboard;
end;

procedure TvgrWorkbookDesignerForm.aPasteExecute(Sender: TObject);
begin
  PasteFromClipboard;
end;

procedure TvgrWorkbookDesignerForm.BuildStatusBar;
begin
  StatusBar.Panels.Clear;
  with StatusBar.Panels.Add do
    Width := 60;
end;

procedure TvgrWorkbookDesignerForm.UpdateStatusBar;
begin
  if StatusBar.Panels.Count > 0 then
    if Modified then
      StatusBar.Panels[0].Text := vgrLoadStr(svgrid_vgr_WorkbookDesigner_Modified)
    else
      StatusBar.Panels[0].Text := vgrLoadStr(svgrid_vgr_WorkbookDesigner_NotModified);
end;

procedure TvgrWorkbookDesignerForm.PageControlChange(Sender: TObject);
begin
  BuildStatusBar;
  UpdateStatusBar;
end;

procedure TvgrWorkbookDesignerForm.OnToolbarClick(Sender: TObject);
begin
  ControlBarManager.ShowHideToolBar(TToolbar(Pointer(TMenuItem(Sender).Tag)), not TToolbar(Pointer(TMenuItem(Sender).Tag)).Visible);
end;

procedure TvgrWorkbookDesignerForm.View1Click(Sender: TObject);
begin
  UpdateToolbarsMenu;
end;

procedure TvgrWorkbookDesignerForm.InternalOnSave;
begin
end;

procedure TvgrWorkbookDesignerForm.InternalOnLoad;
begin
end;

procedure TvgrWorkbookDesignerForm.SetInplaceEditValue(Sender: TObject; const AValue: String; ARange: IvgrRange; ANewRange: Boolean); 
begin
  ARange.StringValue := AValue;
  if ANewRange then
    ARange.WordWrap := (pos(#13, AValue) <> 0) or (pos(#10, AValue) <> 0);
end;

procedure TvgrWorkbookDesignerForm.aPageSetupUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit and (WorkbookGrid.ActiveWorksheet <> nil);
end;

procedure TvgrWorkbookDesignerForm.aPageSetupExecute(Sender: TObject);
begin
  PageSetupDialog.Execute(WorkbookGrid.ActiveWorksheet.PageProperties);
end;

procedure TvgrWorkbookDesignerForm.aPrintPreviewUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not WorkbookGrid.IsInplaceEdit;
end;

procedure TvgrWorkbookDesignerForm.aPrintPreviewExecute(Sender: TObject);
begin
  WorkbookPreviewer.Caption := GetFormCaption;
  WorkbookPreviewer.BoundsRect := BoundsRect;
  try
    WorkbookPreviewer.Workbook := Workbook;
    WorkbookPreviewer.Preview(True);
  finally
    WorkbookPreviewer.Workbook := nil;
  end;
end;

procedure TvgrWorkbookDesignerForm.FormKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key = ord('N')) and (Shift = [ssCtrl]) and aNew.Update then
  begin
    aNew.Execute;
    Key := 0;
  end
  else
  if (Key = ord('O')) and (Shift = [ssCtrl]) and aOpen.Update then
  begin
    aOpen.Execute;
    Key := 0;
  end
  else
  if (Key = ord('S')) and (Shift = [ssCtrl]) and aSave.Update then
  begin
    aSave.Execute;
    Key := 0;
  end;
end;

procedure TvgrWorkbookDesignerForm.FormShow(Sender: TObject);
begin
  if not FSettingsRestored then
  begin
    ControlBarManager.ReadFromStorage(SettingsStorage, StorageSection + '_TB');
    FSettingsRestored := True;
  end;
end;

end.

