{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains TvgrPrintSetupDialog - non visual component and TvgrPrintSetupDialogForm form.
These classes realize a dialog form,
that allows users to select a printer and choose which portions of the workbook to print.
See also:
  TvgrPageSetupDialog, TvgrPageSetupDialogForm, TvgrPageProperties}
unit vgr_PrintSetupDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls,

  vgr_PrintEngine, vgr_PrinterComboBox, vgr_Label, vgr_Printer, vgr_DataStorage,
  vgr_Functions, vgr_GUIFunctions, vgr_Dialogs, vgr_Form, vgr_FormLocalizer;

type

  /////////////////////////////////////////////////
  //
  // TvgrPrintSetupDialogForm
  //
  /////////////////////////////////////////////////
{Implements a dialog form for selecting printer and portions of the workbook to print.}
  TvgrPrintSetupDialogForm = class(TvgrDialogForm)
    vgrBevelLabel1: TvgrBevelLabel;
    Label1: TLabel;
    edPrinterName: TvgrPrinterComboBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    LStatus: TLabel;
    LComment: TLabel;
    LType: TLabel;
    bProperties: TButton;
    vgrBevelLabel2: TvgrBevelLabel;
    vgrBevelLabel3: TvgrBevelLabel;
    Label5: TLabel;
    edCopies: TEdit;
    udCopies: TUpDown;
    bOk: TButton;
    bCancel: TButton;
    Bevel1: TBevel;
    vgrBevelLabel4: TvgrBevelLabel;
    pWorksheetMode: TPanel;
    rbWorksheetModeAll: TRadioButton;
    rbWorksheetModeOne: TRadioButton;
    edPrintWorksheet: TComboBox;
    pPageMode: TPanel;
    rbPageModeAll: TRadioButton;
    rbPageModeRange: TRadioButton;
    rbPageModeCurrent: TRadioButton;
    rbPageModeList: TRadioButton;
    Label6: TLabel;
    edFromPage: TEdit;
    Label7: TLabel;
    edToPage: TEdit;
    edPrintPages: TEdit;
    imgCollate: TImage;
    cbCollate: TCheckBox;
    imgNoCollate: TImage;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure edPrinterNameClick(Sender: TObject);
    procedure edFromPageChange(Sender: TObject);
    procedure cbCollateClick(Sender: TObject);
    procedure rbWorksheetModeAllClick(Sender: TObject);
    procedure rbPageModeAllClick(Sender: TObject);
    procedure edCopiesChange(Sender: TObject);
    procedure edPrintWorksheetClick(Sender: TObject);
    procedure edPrintPagesChange(Sender: TObject);
    procedure bOkClick(Sender: TObject);
    procedure bPropertiesClick(Sender: TObject);
  private
    { Private declarations }
    FPrintProperties: TvgrPrintProperties;
    FWorkbook: TvgrWorkbook;
    FUpdateCount: Integer;
    procedure SetPrintProperties(Value: TvgrPrintProperties);
    procedure BeginUpdate;
    procedure EndUpdate;
    function GetEnableUpdate: Boolean;
    function GetSelectedPrinter: TvgrPrinter;
    procedure UpdateEnabled;

    property EnableUpdate: Boolean read GetEnableUpdate;
  protected
    procedure PrinterSetup;
  public
{Opens the dialog.
Parameters:
  APrintProperties - TvgrPrintProperties object, that
specifies information about how a workbook is printed: range of the pages, amount of copies, etc
  AWorkbook - TvgrWorkbook object, that provides list of worksheets available for printing.
  ACurrentEnabled - value, that indicating whether the "Current" radio button is enabled.
  ASelectedPrinterName - On exit contains name of the printer that are selected by the user.
Return value:
  Returns true when the user edit properties and clicks OK, or false when the user cancels.
See also:
  TvgrPrintProperties, TvgrWorkbook}
    function Execute(APrintProperties: TvgrPrintProperties;
                     AWorkbook: TvgrWorkbook;
                     ACurrentEnabled: Boolean;
                     var ASelectedPrinterName: string): Boolean;

{Gets and sets edited TvgrPrintProperties object, passed to the Execute method.
See also:
  TvgrPrintProperties}
    property PrintProperties: TvgrPrintProperties read FPrintProperties write SetPrintProperties;
{Gets TvgrWorbook object, passed to the Execute method.
See also:
  TvgrWorkbook}
    property Workbook: TvgrWorkbook read FWorkbook;
{Gets TvgrPrinter object, that represents currently selected printer.
See also:
  TvgrPrinter}
    property SelectedPrinter: TvgrPrinter read GetSelectedPrinter;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPrintSetupDialog
  //
  /////////////////////////////////////////////////
{Allows users to select a printer and choose which portions of the workbook to print.
The dialog does not appear at runtime until it is activated by a call to the Execute method.
When the user clicks OK, the dialog closes and all changes is stored to the passed parameters.
See also:
  TvgrPrintSetupDialogForm, TvgrPrintProperties}
  TvgrPrintSetupDialog = class(TComponent)
  private
    FForm: TvgrPrintSetupDialogForm;
    FPrinterName: string;
    procedure SetPrinterName(const Value: string);
  public
{Creates instance of TvgrPrintSetupDialog.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Opens the dialog. All changes of the user stored to the passed parameters,
name of the printer stored to PrinterName property.
Parameters:
  APrintProperties - TvgrPrintProperties object, that
specifies information about how a workbook is printed: range of the pages, amount of copies, etc
  AWorkbook - TvgrWorkbook object, that provides list of worksheets available for printing.
  ACurrentEnabled - value, that indicating whether the "Current" radio button is enabled.
Return value:
  Returns true when the user edit properties and clicks OK, or false when the user cancels.
Example:
  var
    AWorkbook: TvgrWorkbook;
    APrintEngine: TvgrPrintEngine
  begin
  ...
    APrintEngine := TvgrPrintEngine.Create(nil);
    APrintEngine.Workbook := AWorkbook;
    try
      vgrPrintSetupDialog1.Execute(APrintEngine.PrintProperties, AWorkbook, false);
      APrintEngine.Print;
    finally
      APrintEngine.Free;
    end;
  ...
  end;
See also:
  TvgrPrintProperties, TvgrWorkbook
}
    function Execute(APrintProperties: TvgrPrintProperties;
                     AWorkbook: TvgrWorkbook;
                     ACurrentEnabled: Boolean): Boolean;
{Returns reference to a TvgrPrintSetupDialogForm object, created by component in Execute method.}
    property Form: TvgrPrintSetupDialogForm read FForm;
{Gets and sets printer, selected in the dialog.}
    property PrinterName: string read FPrinterName write SetPrinterName;
  end;

implementation

uses vgr_StringIDs, vgr_Localize;

{$R *.dfm}
{$R ..\res\vgr_PrintSetupDialogStrings.res}

/////////////////////////////////////////////////
//
// TvgrPrintSetupDialog
//
/////////////////////////////////////////////////
constructor TvgrPrintSetupDialog.Create(AOwner: TComponent);
begin
  inherited;
end;

procedure TvgrPrintSetupDialog.SetPrinterName(const Value: string);
begin
  if FPrinterName <> Value then
  begin
    if FForm <> nil then
    begin
      FForm.edPrinterName.SelectedPrinterName := Value;
      FPrinterName := FForm.edPrinterName.SelectedPrinterName;
    end
    else
      FPrinterName := Value;
  end;
end;

function TvgrPrintSetupDialog.Execute(APrintProperties: TvgrPrintProperties;
                                      AWorkbook: TvgrWorkbook;
                                      ACurrentEnabled: Boolean): Boolean;
begin
  FForm := TvgrPrintSetupDialogForm.Create(nil);
  Result := FForm.Execute(APrintProperties, AWorkbook, ACurrentEnabled, FPrinterName);
  FForm := nil;
end;

/////////////////////////////////////////////////
//
// TvgrPrintSetupDialogForm
//
/////////////////////////////////////////////////
procedure TvgrPrintSetupDialogForm.SetPrintProperties(Value: TvgrPrintProperties);
begin
  PrintProperties.Assign(Value);
end;

procedure TvgrPrintSetupDialogForm.BeginUpdate;
begin
  Inc(FUpdateCount);
end;

procedure TvgrPrintSetupDialogForm.EndUpdate;
begin
  Dec(FUpdateCount);
end;

function TvgrPrintSetupDialogForm.GetSelectedPrinter: TvgrPrinter;
begin
  Result := edPrinterName.SelectedPrinter;
end;

function TvgrPrintSetupDialogForm.GetEnableUpdate: Boolean;
begin
  Result := FUpdateCount = 0;
end;

function TvgrPrintSetupDialogForm.Execute(APrintProperties: TvgrPrintProperties;
                                          AWorkbook: TvgrWorkbook;
                                          ACurrentEnabled: Boolean;
                                          var ASelectedPrinterName: string): Boolean;
var
  I: Integer;
begin
  FWorkbook := AWorkbook;
  PrintProperties.Assign(APrintProperties);

  if ASelectedPrinterName <> '' then
    edPrinterName.SelectedPrinterName := ASelectedPrinterName;

  if Workbook <> nil then
  begin
    edPrintWorksheet.Items.BeginUpdate;
    try
      edPrintWorksheet.Clear;
      for I := 0 to Workbook.WorksheetsCount - 1 do
        edPrintWorksheet.Items.AddObject(Workbook.Worksheets[I].Title, Workbook.Worksheets[I]);
    finally
      edPrintWorksheet.Items.EndUpdate;
    end;
  end;

  BeginUpdate;
  try
    with PrintProperties do
    begin
      rbWorksheetModeAll.Checked := PrintWorksheetMode = vgrwmAll;
      rbWorksheetModeOne.Checked := PrintWorksheetMode = vgrwmDefined;
      edPrintWorksheet.ItemIndex := edPrintWorksheet.Items.IndexOfObject(PrintWorksheet);

      rbPageModeAll.Checked := PrintPageMode = vgrpmAll;
      rbPageModeCurrent.Checked := PrintPageMode = vgrpmCurrent;
      rbPageModeRange.Checked := PrintPageMode = vgrpmRange;
      rbPageModeList.Checked := PrintPageMode = vgrpmList;

      edFromPage.Text := IntToStr(FromPage);
      edToPage.Text := IntToStr(ToPage);
      edPrintPages.Text := PrintPages;

      udCopies.Position := Copies;
      cbCollate.Checked := Collate;
    end;
  finally
    EndUpdate;
  end;

  rbPageModeCurrent.Enabled := ACurrentEnabled;
  UpdateEnabled;
  edPrinterNameClick(nil);
  cbCollateClick(nil);

  Result := ShowModal = mrOk;
  if Result then
  begin
    APrintProperties.Assign(PrintProperties);
    if SelectedPrinter = nil then
      ASelectedPrinterName := ''
    else
      ASelectedPrinterName := SelectedPrinter.PrinterName;
  end;
end;

procedure TvgrPrintSetupDialogForm.FormCreate(Sender: TObject);
begin
  FPrintProperties := TvgrPrintProperties.Create;
end;

procedure TvgrPrintSetupDialogForm.FormDestroy(Sender: TObject);
begin
  FPrintProperties.Free;
end;

procedure TvgrPrintSetupDialogForm.edPrinterNameClick(Sender: TObject);
var
  AOldCursor: TCursor;
begin
  if SelectedPrinter = nil then
  begin
    LStatus.Caption := '';
    LComment.Caption := '';
    LType.Caption := '';
  end
  else
    with SelectedPrinter do
    begin
      AOldCursor := Screen.Cursor;
      Screen.Cursor := crHourGlass;
      try
        LStatus.Caption := Status;
        LComment.Caption := Comment;
        LType.Caption := DriverName;
      finally
        Screen.Cursor := AOldCursor;
      end;
    end;
  UpdateEnabled;
end;

procedure TvgrPrintSetupDialogForm.UpdateEnabled;
begin
  bProperties.Enabled := SelectedPrinter <> nil;
  rbWorksheetModeOne.Enabled := (Workbook <> nil) and (Workbook.WorksheetsCount > 1);
  edPrintWorksheet.Enabled := (Workbook <> nil) and (Workbook.WorksheetsCount > 1);
  cbCollate.Enabled := (PrintProperties.Copies > 1) and not rbPageModeCurrent.Checked;
end;

procedure TvgrPrintSetupDialogForm.edFromPageChange(Sender: TObject);
var
  AValue: Integer;
begin
  if EnableUpdate then
  begin
    BeginUpdate;
    try
      rbPageModeRange.Checked := True;
      PrintProperties.PrintPageMode := vgrpmRange;
      AValue := StrToIntDef(TEdit(Sender).Text, -1);
      if AValue <> -1 then
      begin
        if Sender = edFromPage then
          PrintProperties.FromPage := AValue
        else
          if Sender = edToPage then
            PrintProperties.ToPage := AValue;
      end;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TvgrPrintSetupDialogForm.cbCollateClick(Sender: TObject);
begin
  PrintProperties.Collate := cbCollate.Checked;
  imgCollate.Visible := cbCollate.Checked;
  imgNoCollate.Visible := not cbCollate.Checked;
end;

procedure TvgrPrintSetupDialogForm.rbWorksheetModeAllClick(
  Sender: TObject);
begin
  if EnableUpdate then
  begin
    if rbWorksheetModeAll.Checked then
      PrintProperties.PrintWorksheetMode := vgrwmAll
    else
      if rbWorksheetModeOne.Checked then
        PrintProperties.PrintWorksheetMode := vgrwmDefined;

    if Sender = rbWorksheetModeOne then
      ActiveControl := edPrintWorksheet;
  end;
end;

procedure TvgrPrintSetupDialogForm.rbPageModeAllClick(Sender: TObject);
begin
  if EnableUpdate then
  begin
    if rbPageModeAll.Checked then
      PrintProperties.PrintPageMode := vgrpmAll
    else
      if rbPageModeRange.Checked then
        PrintProperties.PrintPageMode := vgrpmRange
      else
        if rbPageModeCurrent.Checked then
          PrintProperties.PrintPageMode := vgrpmCurrent
        else
          if rbPageModeList.Checked then
            PrintProperties.PrintPageMode := vgrpmList;

    //
    if Sender = rbPageModeRange then
      ActiveControl := edFromPage
    else
      if Sender = rbPageModeList then
        ActiveControl := edPrintPages;
    UpdateEnabled;
  end;
end;

procedure TvgrPrintSetupDialogForm.edCopiesChange(Sender: TObject);
begin
  if EnableUpdate then
  begin
    if StrToIntDef(edCopies.Text, -1) <> -1 then
    begin
      PrintProperties.Copies := StrToInt(edCopies.Text);
      UpdateEnabled;
    end;
  end;
end;

procedure TvgrPrintSetupDialogForm.edPrintWorksheetClick(Sender: TObject);
begin
  if EnableUpdate then
  begin
    BeginUpdate;
    try
      rbWorksheetModeOne.Checked := True;
      if edPrintWorksheet.ItemIndex >= 0 then
        PrintProperties.PrintWorksheet := TvgrWorksheet(edPrintWorksheet.Items.Objects[edPrintWorksheet.ItemIndex])
      else
        PrintProperties.PrintWorksheet := nil;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TvgrPrintSetupDialogForm.edPrintPagesChange(Sender: TObject);
begin
  if EnableUpdate then
  begin
    BeginUpdate;
    try
      rbPageModeList.Checked := True;
      PrintProperties.PrintPageMode := vgrpmList;
      if CheckPageList(edPrintPages.Text) then
        PrintProperties.PrintPages := edPrintPages.Text;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TvgrPrintSetupDialogForm.bOkClick(Sender: TObject);
begin
  if PrintProperties.FromPage > PrintProperties.ToPage then
  begin
    MBox(vgrLoadStr(svgrid_vgr_PrintSetupDialog_InvalidPagesRange), MB_OK or MB_ICONERROR);
    ActiveControl := edFromPage;
    exit;
  end;

  if not CheckPageList(edPrintPages.Text) then
  begin
    MBox(vgrLoadStr(svgrid_vgr_PrintSetupDialog_InvalidPagesList), MB_OK or MB_ICONERROR);
    ActiveControl := edPrintPages;
    exit;
  end;

  if (PrintProperties.PrintWorksheetMode = vgrwmDefined) and (PrintProperties.PrintWorksheet = nil) then
  begin
    MBox(vgrLoadStr(svgrid_vgr_PrintSetupDialog_PrintWorksheetNotDefined), MB_OK or MB_ICONERROR);
    ActiveControl := edPrintWorksheet;
    exit;
  end; 

  ModalResult := mrOk;
end;

procedure TvgrPrintSetupDialogForm.PrinterSetup;
begin
  if SelectedPrinter <> nil then
    SelectedPrinter.ShowPropertiesDialog(Self.Handle);
end;

procedure TvgrPrintSetupDialogForm.bPropertiesClick(Sender: TObject);
begin
  PrinterSetup;
end;

end.
