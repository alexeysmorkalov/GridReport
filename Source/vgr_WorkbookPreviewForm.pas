{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains built-in workbook preview form.}
unit vgr_WorkbookPreviewForm;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ToolWin, ComCtrls, StdCtrls, ActnList,

  vgr_DataStorage, vgr_WorkbookPreview,
  vgr_PageSetupDialog, vgr_PrintSetupDialog, vgr_PrintEngine, vgr_Form,
  vgr_FormLocalizer, Buttons, vgr_ColorButton, vgr_MultiPageButton,
  vgr_Button;

type
  TvgrCustomWorkbookPreviewer = class;
  TvgrWorkbookPreviewFormClass = class of TvgrWorkbookPreviewForm;

  /////////////////////////////////////////////////
  //
  // TvgrWorkbookPreviewForm
  //
  /////////////////////////////////////////////////
{Implements the form for previewing workbook.
This form is constructed with use of components TvgrWorkbookPreview,
TvgrPrintEngine, TvgrPrintSetupDialog, TvgrPageSetupDialog.
See also:
  TvgrWorkbookPreview, TvgrPrintEngine, TvgrPrintSetupDialog, TvgrPageSetupDialog}
  TvgrWorkbookPreviewForm = class(TvgrForm)
    ToolBar: TToolBar;
    ImageList: TImageList;
    StatusBar: TStatusBar;
    WorkbookPreview: TvgrWorkbookPreview;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    edScalePercent: TComboBox;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ActionList: TActionList;
    aPrint: TAction;
    aEditPageParams: TAction;
    aZoomPageWidth: TAction;
    aZoomWholePage: TAction;
    aZoom100Percent: TAction;
    aFirstPage: TAction;
    aPriorPage: TAction;
    aNextPage: TAction;
    aLastPage: TAction;
    ToolButton13: TToolButton;
    PageSetupDialog: TvgrPageSetupDialog;
    PrintSetupDialog: TvgrPrintSetupDialog;
    PrintEngine: TvgrPrintEngine;
    vgrFormLocalizer1: TvgrFormLocalizer;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    aHandTool: TAction;
    aZoomInTool: TAction;
    aZoomOutTool: TAction;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    aNoneTool: TAction;
    ToolButton19: TToolButton;
    aClose: TAction;
    CloseSpeedButton: TSpeedButton;
    ToolButton20: TToolButton;
    aZoomTwoPages: TAction;
    vgrMultiPageButton1: TvgrMultiPageButton;
    procedure aPrintUpdate(Sender: TObject);
    procedure aPrintExecute(Sender: TObject);
    procedure aEditPageParamsUpdate(Sender: TObject);
    procedure aEditPageParamsExecute(Sender: TObject);
    procedure aZoomPageWidthUpdate(Sender: TObject);
    procedure aZoomPageWidthExecute(Sender: TObject);
    procedure aZoomWholePageUpdate(Sender: TObject);
    procedure aZoomWholePageExecute(Sender: TObject);
    procedure aZoom100PercentUpdate(Sender: TObject);
    procedure aZoom100PercentExecute(Sender: TObject);
    procedure aFirstPageUpdate(Sender: TObject);
    procedure aFirstPageExecute(Sender: TObject);
    procedure aPriorPageUpdate(Sender: TObject);
    procedure aNextPageUpdate(Sender: TObject);
    procedure aNextPageExecute(Sender: TObject);
    procedure aPriorPageExecute(Sender: TObject);
    procedure aLastPageUpdate(Sender: TObject);
    procedure aLastPageExecute(Sender: TObject);
    procedure WorkbookPreviewScalePercentChanged(Sender: TObject);
    procedure edScalePercentKeyPress(Sender: TObject; var Key: Char);
    procedure WorkbookPreviewActivePageChanged(Sender: TObject);
    procedure WorkbookPreviewActiveWorksheetChanged(Sender: TObject);
    procedure WorkbookPreviewNeedPageParams(Sender: TObject);
    procedure WorkbookPreviewNeedPrint(Sender: TObject);
    procedure WorkbookPreviewActiveWorksheetPagesCountChanged(
      Sender: TObject);
    procedure aHandToolExecute(Sender: TObject);
    procedure aZoomInToolExecute(Sender: TObject);
    procedure aZoomOutToolExecute(Sender: TObject);
    procedure aHandToolUpdate(Sender: TObject);
    procedure aZoomInToolUpdate(Sender: TObject);
    procedure aZoomOutToolUpdate(Sender: TObject);
    procedure aNoneToolExecute(Sender: TObject);
    procedure aNoneToolUpdate(Sender: TObject);
    procedure aCloseExecute(Sender: TObject);
    procedure aZoomTwoPagesExecute(Sender: TObject);
    procedure aZoomTwoPagesUpdate(Sender: TObject);
    procedure vgrMultiPageButton1PageSelected(Sender: TObject;
      Pages: TPoint);
    procedure edScalePercentClick(Sender: TObject);
  private
    { Private declarations }
    FPreviewer: TvgrCustomWorkbookPreviewer;

    procedure SetPreviewer(Value: TvgrCustomWorkbookPreviewer);
    function GetWorkbook: TvgrWorkbook;
    procedure SetWorkbook(Value: TvgrWorkbook);
    function GetActiveWorksheet: TvgrWorksheet;
    function GetActivePage: Integer;
    procedure SetActivePage(Value: Integer);
    function GetPageCount: Integer;
  protected
    procedure UpdateStatusBar; virtual;

    function GetFixupShortcuts: Boolean; override;
    function GetStoreSettingsMethod: TvgrStoreSettingsMethod; override;

    property Previewer: TvgrCustomWorkbookPreviewer read FPreviewer write SetPreviewer;
  public
    procedure BeforeDestruction; override;

{Sets or gets the TvgrWorkbook object, that are linked to this form.}
    property Workbook: TvgrWorkbook read GetWorkbook write SetWorkbook;
{Returns the active worksheet.
See also:
  TvgrWorkbookPreview.ActiveWorksheet}
    property ActiveWorksheet: TvgrWorksheet read GetActiveWorksheet;
{Returns the active page.
See also:
  TvgrWorkbookPreview.ActivePage}
    property ActivePage: Integer read GetActivePage write SetActivePage;
{Returns the amount of the pages in the active worksheet.}
    property PageCount: Integer read GetPageCount;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrCustomWorkbookPreviewer
  //
  /////////////////////////////////////////////////
{Non visual component, realizes the preview window of the workvbook.
The preview window does not appear at runtime until it is activated by a call to the Preview method.
The preview window can be shown as modal whether or not.
See also:
  TvgrWorkbookPreviewer}
  TvgrCustomWorkbookPreviewer = class(TComponent)
  private
    FBoundsRect: TRect;
    FCaption: string;
    FWorkbook: TvgrWorkbook;
    FForm: TvgrWorkbookPreviewForm;
    procedure SetWorkbook(Value: TvgrWorkbook);
    procedure SetCaption(const Value: string);
    procedure SetBoundsrect(const ARect: TRect);
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
    function GetPreviewFormClass: TvgrWorkbookPreviewFormClass; virtual;
    property Form: TvgrWorkbookPreviewForm read FForm;
  public
{Creates a instance of the TvgrCustomWorkbookPreviewer class.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees a instance of the TvgrCustomWorkbookPreviewer class.}
    destructor Destroy; override;

{Preview opens the preview form.
Parameters:
  AModal - if true the form shown as modal.}
    procedure Preview(AModal: Boolean);
{Showes the created preview form.
Parameters:
  AModal - if true the form shown as modal.}
    procedure ActivateForm(AModal: Boolean);
{Closes the preview form.}
    procedure CloseForm;

{Sets or gets the workbook for previewing.}
    property Workbook: TvgrWorkbook read FWorkbook write SetWorkbook;
{Sets or gets the title of the form.}
    property Caption: string read FCaption write SetCaption;
{Sets or gets the bounds of the preview form.}
    property BoundsRect: TRect read FBoundsRect write SetBoundsRect;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWorkbookPreviewer
  //
  /////////////////////////////////////////////////
{Non visual component, realizes the preview window of the workvbook.
The preview window does not appear at runtime until it is activated by a call to the Preview method.
The preview window can be shown as modal whether or not.}
  TvgrWorkbookPreviewer = class(TvgrCustomWorkbookPreviewer)
  published
{Sets or gets the workbook for previewing.}
    property Workbook;
{Sets or gets the title of the form.}
    property Caption;
  end;

implementation

uses
  {$IFDEF VTK_D6_OR_D7} Types, {$ENDIF} vgr_Functions, vgr_StringIDs, vgr_Localize;

{$R *.dfm}
{$R ..\res\vgr_WorkbookPreviewFormStrings.res}

/////////////////////////////////////////////////
//
// TvgrCustomWorkbookPreviewer
//
/////////////////////////////////////////////////
constructor TvgrCustomWorkbookPreviewer.Create(AOwner: TComponent);
begin
  inherited;
end;

destructor TvgrCustomWorkbookPreviewer.Destroy;
begin
  FreeAndNil(FForm);
  inherited;
end;

procedure TvgrCustomWorkbookPreviewer.SetWorkbook(Value: TvgrWorkbook);
begin
  if FWorkbook <> Value then
  begin
    FWorkbook := Value;
    if FForm <> nil then
      FForm.Workbook := Value;
  end;
end;

procedure TvgrCustomWorkbookPreviewer.SetCaption(const Value: string);
begin
  if FCaption <> Value then
  begin
    FCaption := Value;
    if FForm <> nil then
      FForm.Caption := FCaption;
  end;
end;

procedure TvgrCustomWorkbookPreviewer.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if (AOperation = opRemove) and (AComponent = Workbook) then
    Workbook := nil;
end;

function TvgrCustomWorkbookPreviewer.GetPreviewFormClass: TvgrWorkbookPreviewFormClass;
begin
  Result := TvgrWorkbookPreviewForm;
end;

procedure TvgrCustomWorkbookPreviewer.Preview(AModal: Boolean);
begin
  if Workbook <> nil then
  begin
    if FForm = nil then
    begin
      FForm := GetPreviewFormClass.Create(Application);
      FForm.Caption := Caption;
      FForm.Workbook := Workbook;
      FForm.Previewer := Self;
      if not IsRectEmpty(BoundsRect) then
        FForm.BoundsRect := BoundsRect;
    end;
    ActivateForm(AModal);
  end;
end;

procedure TvgrCustomWorkbookPreviewer.ActivateForm(AModal: Boolean);
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

procedure TvgrCustomWorkbookPreviewer.CloseForm;
begin
  if FForm <> nil then
    FForm.Close;
end;

procedure TvgrCustomWorkbookPreviewer.SetBoundsRect(const ARect: TRect);
begin
  if FForm <> nil then
    FForm.BoundsRect := ARect;
  FBoundsRect := ARect;
end;

/////////////////////////////////////////////////
//
// TvgrWorkbookPreviewForm
//
/////////////////////////////////////////////////
function TvgrWorkbookPreviewForm.GetFixupShortcuts: Boolean;
begin
  Result := True;
end;

function TvgrWorkbookPreviewForm.GetStoreSettingsMethod: TvgrStoreSettingsMethod;
begin
  Result := DefaultStoreSettingsMethod;
end;

procedure TvgrWorkbookPreviewForm.BeforeDestruction;
begin
  inherited;
  Workbook := nil;
  if Previewer <> nil then
    Previewer.FForm := nil;
end;

procedure TvgrWorkbookPreviewForm.SetPreviewer(Value: TvgrCustomWorkbookPreviewer);
begin
  if FPreviewer <> Value then
  begin
    FPreviewer := Value;
  end;
end;

function TvgrWorkbookPreviewForm.GetWorkbook: TvgrWorkbook;
begin
  if WorkbookPreview = nil then
    Result := nil
  else
    Result := WorkbookPreview.Workbook;
end;

procedure TvgrWorkbookPreviewForm.SetWorkbook(Value: TvgrWorkbook);
begin
  if WorkbookPreview <> nil then
    WorkbookPreview.Workbook := Value;
end;

function TvgrWorkbookPreviewForm.GetActiveWorksheet: TvgrWorksheet;
begin
  if WorkbookPreview = nil then
    Result := nil
  else
    Result := WorkbookPreview.ActiveWorksheet;
end;

function TvgrWorkbookPreviewForm.GetActivePage: Integer;
begin
  if WorkbookPreview = nil then
    Result := -1
  else
    Result := WorkbookPreview.ActivePage;
end;

procedure TvgrWorkbookPreviewForm.SetActivePage(Value: Integer);
begin
  if WorkbookPreview <> nil then
    WorkbookPreview.ActivePage := Value;
end;

function TvgrWorkbookPreviewForm.GetPageCount: Integer;
begin
  if WorkbookPreview = nil then
    Result := -1
  else
    Result := WorkbookPreview.PageCount;
end;

procedure TvgrWorkbookPreviewForm.UpdateStatusBar;
begin
  StatusBar.SimpleText := Format(vgrLoadStr(svgrid_vgr_WorkbookPreviewForm_StatusBarFormatMask), [ActivePage + 1, PageCount]);
end;

procedure TvgrWorkbookPreviewForm.aPrintUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaPrint);
end;

procedure TvgrWorkbookPreviewForm.aPrintExecute(Sender: TObject);
begin
  WorkbookPreview.Print;
end;

procedure TvgrWorkbookPreviewForm.aEditPageParamsUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaPageParams);
end;

procedure TvgrWorkbookPreviewForm.aEditPageParamsExecute(Sender: TObject);
begin
  WorkbookPreview.EditPageParams;
end;

procedure TvgrWorkbookPreviewForm.aZoomPageWidthUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaZoomPageWidth);
  TAction(Sender).Checked := WorkbookPreview.ScaleMode = vgrpsmPageWidth;
end;

procedure TvgrWorkbookPreviewForm.aZoomPageWidthExecute(Sender: TObject);
begin
  WorkbookPreview.ScaleMode := vgrpsmPageWidth;
end;

procedure TvgrWorkbookPreviewForm.aZoomWholePageUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaZoomWholePage);
  TAction(Sender).Checked := WorkbookPreview.ScaleMode = vgrpsmWholePage;
end;

procedure TvgrWorkbookPreviewForm.aZoomWholePageExecute(Sender: TObject);
begin
  WorkbookPreview.ScaleMode := vgrpsmWholePage;
end;

procedure TvgrWorkbookPreviewForm.aZoom100PercentUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaZoom100Percent);
  TAction(Sender).Checked := WorkbookPreview.ScaleMode = vgrpsmPercent;
end;

procedure TvgrWorkbookPreviewForm.aZoom100PercentExecute(Sender: TObject);
begin
  WorkbookPreview.ScaleMode := vgrpsmPercent;
  WorkbookPreview.ScalePercent := 100;
end;

procedure TvgrWorkbookPreviewForm.aFirstPageUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaFirstPage);
end;

procedure TvgrWorkbookPreviewForm.aFirstPageExecute(Sender: TObject);
begin
  WorkbookPreview.DoPreviewAction(vgrpaFirstPage);
end;

procedure TvgrWorkbookPreviewForm.aPriorPageUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaPriorPage);
end;

procedure TvgrWorkbookPreviewForm.aPriorPageExecute(Sender: TObject);
begin
  WorkbookPreview.DoPreviewAction(vgrpaPriorPage);
end;

procedure TvgrWorkbookPreviewForm.aNextPageUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaNextPage);
end;

procedure TvgrWorkbookPreviewForm.aNextPageExecute(Sender: TObject);
begin
  WorkbookPreview.DoPreviewAction(vgrpaNextPage);
end;

procedure TvgrWorkbookPreviewForm.aLastPageUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaLastPage);
end;

procedure TvgrWorkbookPreviewForm.aLastPageExecute(Sender: TObject);
begin
  WorkbookPreview.DoPreviewAction(vgrpaLastPage);
end;

procedure TvgrWorkbookPreviewForm.WorkbookPreviewScalePercentChanged(
  Sender: TObject);
begin
  edScalePercent.Text := FormatFloat('0', WorkbookPreview.ScalePercent) + '%';
end;

procedure TvgrWorkbookPreviewForm.edScalePercentKeyPress(Sender: TObject;
  var Key: Char);
var
  AScalePercent: Extended;
begin
  if Key = #13 then
  begin
    if TextToFloat(PChar(edScalePercent.Text), AScalePercent, fvExtended) then
    begin
      WorkbookPreview.ScaleMode := vgrpsmPercent;
      WorkbookPreview.ScalePercent := AScalePercent;
    end;
    Key := #0;
    ActiveControl := WorkbookPreview;
  end;
end;

procedure TvgrWorkbookPreviewForm.WorkbookPreviewActivePageChanged(
  Sender: TObject);
begin
  UpdateStatusBar;
end;

procedure TvgrWorkbookPreviewForm.WorkbookPreviewActiveWorksheetChanged(
  Sender: TObject);
begin
  UpdateStatusBar;
end;

procedure TvgrWorkbookPreviewForm.WorkbookPreviewNeedPageParams(
  Sender: TObject);
begin
  if WorkbookPreview.ActiveWorksheet <> nil then
    PageSetupDialog.Execute(WorkbookPreview.ActiveWorksheet.PageProperties);
end;

procedure TvgrWorkbookPreviewForm.WorkbookPreviewNeedPrint(
  Sender: TObject);
var
  I: Integer;
begin
  if Workbook <> nil then
  begin
    if PrintSetupDialog.Execute(PrintEngine.PrintProperties, Workbook, True) then
    begin
      with PrintEngine do
      begin
        Workbook := Self.Workbook;
        PrinterName := PrintSetupDialog.PrinterName;
        DocumentTitle := Caption;

        if PrintProperties.PrintPageMode = vgrpmCurrent then
        begin
          for I := 1 to PrintProperties.Copies do
            PrintPage(WorkbookPreview.ActivePrintPage, WorkbookPreview.ActiveWorksheet, WorkbookPreview.ActiveWorksheet.PageProperties);
          EndDocument;
        end
        else
          PrintEngine.Print;
      end;
    end;
  end;
end;

procedure TvgrWorkbookPreviewForm.WorkbookPreviewActiveWorksheetPagesCountChanged(
  Sender: TObject);
begin
  UpdateStatusBar;
end;

procedure TvgrWorkbookPreviewForm.aHandToolExecute(Sender: TObject);
begin
  WorkbookPreview.ActionMode := vgrpamHand;
end;

procedure TvgrWorkbookPreviewForm.aZoomInToolExecute(Sender: TObject);
begin
    WorkbookPreview.ActionMode := vgrpamZoomIn;
end;

procedure TvgrWorkbookPreviewForm.aZoomOutToolExecute(Sender: TObject);
begin
  WorkbookPreview.ActionMode := vgrpamZoomOut;
end;

procedure TvgrWorkbookPreviewForm.aHandToolUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := True;
  TAction(Sender).Checked := WorkbookPreview.ActionMode = vgrpamHand;
end;

procedure TvgrWorkbookPreviewForm.aZoomInToolUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := True;
  TAction(Sender).Checked := WorkbookPreview.ActionMode = vgrpamZoomIn;
end;

procedure TvgrWorkbookPreviewForm.aZoomOutToolUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := True;
  TAction(Sender).Checked := WorkbookPreview.ActionMode = vgrpamZoomOut;
end;

procedure TvgrWorkbookPreviewForm.aNoneToolExecute(Sender: TObject);
begin
  WorkbookPreview.ActionMode := vgrpamNone;
end;

procedure TvgrWorkbookPreviewForm.aNoneToolUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := True;
  TAction(Sender).Checked := WorkbookPreview.ActionMode = vgrpamNone;
end;

procedure TvgrWorkbookPreviewForm.aCloseExecute(Sender: TObject);
begin
  Close;
end;

procedure TvgrWorkbookPreviewForm.aZoomTwoPagesExecute(Sender: TObject);
begin
  WorkbookPreview.ScaleMode := vgrpsmTwoPages;
end;

procedure TvgrWorkbookPreviewForm.aZoomTwoPagesUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := WorkbookPreview.IsPreviewActionAllowed(vgrpaZoomTwoPages);
  TAction(Sender).Checked := WorkbookPreview.ScaleMode = vgrpsmTwoPages;
end;

procedure TvgrWorkbookPreviewForm.vgrMultiPageButton1PageSelected(
  Sender: TObject; Pages: TPoint);
begin
  WorkbookPreview.SetMultiPagePreview(Pages.X,Pages.Y);
end;

procedure TvgrWorkbookPreviewForm.edScalePercentClick(Sender: TObject);
var
  AKey: Char;
begin
  AKey := #13;
  edScalePercentKeyPress(Self,  AKey);
end;

end.

