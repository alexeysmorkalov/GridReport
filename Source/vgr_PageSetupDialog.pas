{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains TvgrPageSetupDialog - non visual component and TvgrPageSetupDialogForm form.
These classes realize a dialog form
for editing properties of the page for printing or previewing.
See also:
  TvgrPageSetupDialog, TvgrPageSetupDialogForm, TvgrPageProperties}
unit vgr_PageSetupDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, StdCtrls,

  vgr_PageProperties, vgr_FormLocalizer, vgr_PrinterComboBox,
  vgr_Label, vgr_Form, 
  vgr_CommonClasses;

type

  TvgrPageSetupDialogForm = class;
  TvgrRegisteredHeaderFooterFormat = class;

  /////////////////////////////////////////////////
  //
  // TvgrPageSetupDialog
  //
  /////////////////////////////////////////////////
{Represents a dialog box that allows users to manipulate page settings, including margins and paper orientation.
The TvgrPageSetupDialog component displays a dialog box for editing properties of the page.
The dialog does not appear at runtime until it is activated by a call to the Execute method.
When the user clicks OK, the dialog closes and all changes is stored in the APageProperties parameter.
See also:
  TvgrPageSetupDialogForm}
  TvgrPageSetupDialog = class(TComponent)
  private
    FForm: TvgrPageSetupDialogForm;
  public
    constructor Create(AOwner: TComponent); override;
{Opens the dialog.
Parameters:
  APageProperties - specifies a TvgrPageProperties object, which should be edited.
Return value:
  Returns true when the user edit properties and clicks OK, or false when the user cancels.
Example:
  var
    sheet: TvgrWorksheet;
  begin
  ...
  vgrPageSetupDialog1.Execute(sheet.PageProperties);
  ...
  end;
See also:
  TvgrPageProperties, TvgrWorksheet}
    function Execute(APageProperties: TvgrPageProperties): Boolean;
{Returns reference to a TvgrCellPropertiesForm object, created by component in Execute method.}
    property Form: TvgrPageSetupDialogForm read FForm;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPageSetupDialogForm
  //
  /////////////////////////////////////////////////
{Implements a dialog form for editing page properties.}
  TvgrPageSetupDialogForm = class(TvgrDialogForm)
    PageControl: TPageControl;
    Panel1: TPanel;
    bOk: TButton;
    bCancel: TButton;
    PPage: TTabSheet;
    PMargins: TTabSheet;
    Label2: TvgrBevelLabel;
    rbPredefinedPageSize: TRadioButton;
    rbCustomPageSize: TRadioButton;
    edPrefferedPageSize: TComboBox;
    edCustomPageWidth: TEdit;
    Label3: TLabel;
    edCustomPageHeight: TEdit;
    Label4: TvgrBevelLabel;
    Label1: TLabel;
    Panel2: TPanel;
    Label5: TvgrBevelLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    edPrinterMarginLeft: TEdit;
    edPrinterMarginTop: TEdit;
    edPrinterMarginRight: TEdit;
    edPrinterMarginBottom: TEdit;
    edMarginBottom: TEdit;
    edMarginRight: TEdit;
    Label10: TLabel;
    Label11: TLabel;
    edMarginLeft: TEdit;
    edMarginTop: TEdit;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TvgrBevelLabel;
    edPrinterName: TvgrPrinterComboBox;
    edMeasurementSystem: TComboBox;
    Panel3: TPanel;
    rbPagePortrait: TRadioButton;
    rbPageLandscape: TRadioButton;
    PagePreview: TPaintBox;
    Label15: TLabel;
    PDefaults: TTabSheet;
    Label16: TLabel;
    Label17: TLabel;
    edDefaultColWidth: TEdit;
    edDefaultRowHeight: TEdit;
    vgrFormLocalizer1: TvgrFormLocalizer;
    PHeaderFooter: TTabSheet;
    vgrBevelLabel1: TvgrBevelLabel;
    vgrBevelLabel2: TvgrBevelLabel;
    edHeader: TComboBox;
    bCustomHeader: TButton;
    bCustomFooter: TButton;
    edFooter: TComboBox;
    Label18: TLabel;
    edHeaderHeight: TEdit;
    Label19: TLabel;
    edFooterHeight: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure edPrinterNameClick(Sender: TObject);
    procedure edMeasurementSystemClick(Sender: TObject);
    procedure rbPredefinedPageSizeClick(Sender: TObject);
    procedure rbCustomPageSizeClick(Sender: TObject);
    procedure edPrefferedPageSizeClick(Sender: TObject);
    procedure edCustomPageWidthChange(Sender: TObject);
    procedure edCustomPageHeightChange(Sender: TObject);
    procedure rbPagePortraitClick(Sender: TObject);
    procedure edMarginLeftChange(Sender: TObject);
    procedure PagePreviewPaint(Sender: TObject);
    procedure edCustomPageWidthExit(Sender: TObject);
    procedure edMarginLeftExit(Sender: TObject);
    procedure edDefaultColWidthChange(Sender: TObject);
    procedure edDefaultColWidthExit(Sender: TObject);
    procedure bCustomHeaderClick(Sender: TObject);
    procedure bCustomFooterClick(Sender: TObject);
    procedure edHeaderHeightChange(Sender: TObject);
    procedure edHeaderHeightExit(Sender: TObject);
    procedure edHeaderClick(Sender: TObject);
    procedure edHeaderDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
  private
    FCustomHeaderFormat: TvgrRegisteredHeaderFooterFormat;
    FCustomFooterFormat: TvgrRegisteredHeaderFooterFormat;
    FPageProperties: TvgrPageProperties;
    FUpdateCount: Integer;
    FPagePreviewKoef: Double;
    FSystemPrefix: array [TvgrMeasurementSystem] of string;

    procedure UpdateEnabled;
    procedure UpdateCustomPageSizeEdits;
    procedure UpdatePageOrientation;
    procedure UpdatePageMargins;
    procedure UpdatePagePreview;
    procedure UpdateDefaults;
    procedure UpdateHeaderFooter;
    procedure BeginUpdate;
    procedure EndUpdate;
    function GetEnableUpdate: Boolean;
    function FormatNumber(const S: string): string;
    function TextToFloat(const S: string; var AValue: extended): Boolean;
    procedure SetHeaderFooterFormat(AHeaderFooter: TvgrPageHeaderFooter; edCombo: TComboBox; ACustomFormat: TvgrRegisteredHeaderFooterFormat);
    property EnableUpdate: Boolean read GetEnableUpdate;
    property PagePreviewKoef: Double read FPagePreviewKoef;
  protected
    procedure RestoreSettings; override;
    procedure SaveSettings; override;

    procedure DoLocalize; override;
  public
{Opens the dialog.
Parameters:
  APageProperties - specifies a TvgrPageProperties object, which should be edited.
Return value:
  Returns true when the user edit properties and clicks OK, or false when the user cancels.}
    function Execute(APageProperties: TvgrPageProperties): Boolean;
{Edited TvgrPageProperties object.
Syntax:
  property PageProperties: TvgrPageProperties read;}
    property PageProperties: TvgrPageProperties read FPageProperties;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredHeaderFooterFormat
  //
  /////////////////////////////////////////////////
{Represents the predefined format of page's header or footer.
See also:
  TvgrRegisteredHeaderFooterFormats}
  TvgrRegisteredHeaderFooterFormat = class(TObject)
  private
    FLeftSection: string;
    FCenterSection: string;
    FRightSection: string;
  public
{Specifies the text of left section.
See also:
  TvgPageHeaderFooter.LeftSection}
    property LeftSection: string read FLeftSection;
{Specifies the text of center section.
See also:
  TvgPageHeaderFooter.CenterSection}
    property CenterSection: string read FCenterSection;
{Specifies the text of right section.
See also:
  TvgPageHeaderFooter.RightSection}
    property RightSection: string read FRightSection; 
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredHeaderFooterFormats
  //
  /////////////////////////////////////////////////
{Maintains the list of predefined formats of header or footer of page.
See also:
  TvgrRegisteredHeaderFooterFormat}
  TvgrRegisteredHeaderFooterFormats = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrRegisteredHeaderFooterFormat;
  public
{Registers a new format, if format is already registered then it will be ignored.
Parameters:
  ALeftSection - Format of the left section.
  ACenterSection - Format of the center section.
  ARightSection - Format of the right section.
Example:
  begin
    ...
    // this format will display the "Page X of X" text in the right section of header or footer.
    HeaderFooterFormats.RegisterFormat('', '', 'Page &[Page] of &[Pages]');
    ...
  end;}
    procedure RegisterFormat(const ALeftSection, ACenterSection, ARightSection: string);
{Lists the registered formats.
Parameters:
  Index - Specifies the index of format, starts from 0.}    
    property Items[Index: Integer]: TvgrRegisteredHeaderFooterFormat read GetItem; default;
  end;

var
{Global variable, that holds the list of predefined formats of page's footer or header.
Syntax:
  HeaderFooterFormats: TvgrRegisteredHeaderFooterFormats;
See also:
  TvgrRegisteredHeaderFooterFormats}
  HeaderFooterFormats: TvgrRegisteredHeaderFooterFormats;

implementation

uses
  Math, vgr_GUIFunctions, vgr_Functions, vgr_StringIDs, vgr_Printer,
  vgr_ReportGUIFunctions, vgr_Localize, vgr_HeaderFooterOptionsDialog,
  vgr_DataStorageTypes;

{$R *.dfm}
{$R ..\res\vgr_PageSetupDialogStrings.res}

/////////////////////////////////////////////////
//
// TvgrPageSetupDialog
//
/////////////////////////////////////////////////
constructor TvgrPageSetupDialog.Create(AOwner: TComponent);
begin
  inherited;
end;

function TvgrPageSetupDialog.Execute(APageProperties: TvgrPageProperties): Boolean;
begin
  FForm := TvgrPageSetupDialogForm.Create(nil);
  Result := FForm.Execute(APageProperties);
  FForm := nil;
end;

/////////////////////////////////////////////////
//
// TvgrPageSetupDialogForm
//
/////////////////////////////////////////////////
procedure TvgrPageSetupDialogForm.RestoreSettings;
begin
  inherited;
  edPrinterName.SelectedPrinterName := SettingsStorage.ReadString(StorageSection, 'PrinterName', '');
end;

procedure TvgrPageSetupDialogForm.SaveSettings;
begin
  inherited;
  SettingsStorage.WriteString(StorageSection, 'PrinterName', edPrinterName.SelectedPrinterName);
end;

procedure TvgrPageSetupDialogForm.UpdateEnabled;
begin
  edPrefferedPageSize.Enabled := rbPredefinedPageSize.Checked;
  edCustomPageWidth.Enabled := rbCustomPageSize.Checked;
  edCustomPageHeight.Enabled := rbCustomPageSize.Checked;
  Label3.Enabled := rbCustomPageSize.Checked;
end;

procedure TvgrPageSetupDialogForm.UpdateCustomPageSizeEdits;
begin
  BeginUpdate;
  try
    edCustomPageWidth.Text := FormatNumber(PageProperties.UnitsWidthStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
    edCustomPageHeight.Text := FormatNumber(PageProperties.UnitsHeightStr[ASystemToUnits[PageProperties.MeasurementSystem]]);

    Label15.Caption := Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_PaperSizeCaptionMask),
                              [PageProperties.UnitsWidthStr[ASystemToUnits[PageProperties.MeasurementSystem]],
                               PageProperties.UnitsHeightStr[ASystemToUnits[PageProperties.MeasurementSystem]],
                               FSystemPrefix[PageProperties.MeasurementSystem]]);

    UpdatePagePreview;
  finally
    EndUpdate;
  end;
end;

procedure TvgrPageSetupDialogForm.UpdatePageOrientation;
begin
  BeginUpdate;
  try
    rbPagePortrait.Checked := PageProperties.Width <= PageProperties.Height;
    rbPageLandscape.Checked := PageProperties.Width > PageProperties.Height;
    UpdatePagePreview;
  finally
    EndUpdate;
  end;
end;

procedure TvgrPageSetupDialogForm.UpdatePagePreview;
var
  AWidth, AHeight: Integer;
  AClientWidth, AClientHeight: Integer;
  ARect: TRect;
begin
  AWidth := Round(PageProperties.UnitsWidth[vgruPixels]);
  AHeight := Round(PageProperties.UnitsHeight[vgruPixels]);
  if (AWidth <> 0) and (AHeight <> 0) then
  begin
    AClientWidth := Panel2.ClientWidth - 8;
    AClientHeight := Label15.Top - 8;
    if AClientWidth / AWidth < AClientHeight / AHeight then
      FPagePreviewKoef := AClientWidth / AWidth
    else
      FPagePreviewKoef := AClientHeight / AHeight;
    ARect := Rect(0, 4, Round(AWidth * FPagePreviewKoef), 4 + Round(AHeight * FPagePreviewKoef));
    OffsetRect(ARect, (Panel2.ClientWidth - ARect.Right) div 2, 0);
  end
  else
    ARect := Rect(0, 0, 0, 0);

  PagePreview.BoundsRect := ARect;
end;

procedure TvgrPageSetupDialogForm.UpdatePageMargins;
begin
  BeginUpdate;
  try
    with PageProperties.Margins do
    begin
      edMarginLeft.Text := FormatNumber(UnitsLeftStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
      edMarginTop.Text := FormatNumber(UnitsTopStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
      edMarginRight.Text := FormatNumber(UnitsRightStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
      edMarginBottom.Text := FormatNumber(UnitsBottomStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
    end;
    if edPrinterName.SelectedPrinter = nil then
    begin
      edPrinterMarginLeft.Text := '';
      edPrinterMarginTop.Text := '';
      edPrinterMarginRight.Text := '';
      edPrinterMarginBottom.Text := '';
    end
    else
      with edPrinterName.SelectedPrinter.PageMargins do
      begin
        edPrinterMarginLeft.Text := FormatNumber(UnitsLeftStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
        edPrinterMarginTop.Text := FormatNumber(UnitsTopStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
        edPrinterMarginRight.Text := FormatNumber(UnitsRightStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
        edPrinterMarginBottom.Text := FormatNumber(UnitsBottomStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
      end;
    PagePreview.Invalidate;
  finally
    EndUpdate;
  end;
end;

procedure TvgrPageSetupDialogForm.UpdateDefaults;
begin
  BeginUpdate;
  try
    with PageProperties.Defaults do
    begin
      edDefaultColWidth.Text := FormatNumber(UnitsColWidthStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
      edDefaultRowHeight.Text := FormatNumber(UnitsRowHeightStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
    end;
  finally
    EndUpdate;
  end;
end;

procedure TvgrPageSetupDialogForm.UpdateHeaderFooter;
begin
  BeginUpdate;
  try
    with PageProperties.Header do
      edHeaderHeight.Text := FormatNumber(UnitsHeightStr[ASystemToUnits[PageProperties.MeasurementSystem]]);

    with PageProperties.Footer do
      edFooterHeight.Text := FormatNumber(UnitsHeightStr[ASystemToUnits[PageProperties.MeasurementSystem]]);
  finally
    EndUpdate;
  end;
end;

procedure TvgrPageSetupDialogForm.BeginUpdate;
begin
  Inc(FUpdateCount);
end;

procedure TvgrPageSetupDialogForm.EndUpdate;
begin
  Dec(FUpdateCount);
end;

function TvgrPageSetupDialogForm.GetEnableUpdate: Boolean;
begin
  Result := FUpdateCount = 0;
end;

function TvgrPageSetupDialogForm.FormatNumber(const S: string): string;
begin
  Result := S + ' ' + FSystemPrefix[PageProperties.MeasurementSystem];
end;

function TvgrPageSetupDialogForm.TextToFloat(const S: string; var AValue: extended): Boolean;
var
  I: Integer;
  Temp: string;
begin
  Temp := '';
  for I := 1 to Length(S) do
    if S[I] in ['0'..'9', DecimalSeparator] then
      Temp := Temp + S[I];
  Result := SysUtils.TextToFloat(PChar(Temp), AValue, fvExtended);
end;

procedure TvgrPageSetupDialogForm.SetHeaderFooterFormat(AHeaderFooter: TvgrPageHeaderFooter; edCombo: TComboBox; ACustomFormat: TvgrRegisteredHeaderFooterFormat);
var
  I, J: Integer;
  AFormat: TvgrRegisteredHeaderFooterFormat;
begin
  for I := 0 to edCombo.Items.Count - 1 do
    if (edCombo.Items.Objects[I] is TvgrRegisteredHeaderFooterFormat) and
       (edCombo.Items.Objects[I] <> ACustomFormat) then
      with TvgrRegisteredHeaderFooterFormat(edCombo.Items.Objects[I]) do
        if (AHeaderFooter.LeftSection = LeftSection) and
           (AHeaderFooter.CenterSection = CenterSection) and
           (AHeaderFooter.RightSection = RightSection) then
        begin
          AFormat := TvgrRegisteredHeaderFooterFormat(edCombo.Items.Objects[I]);

          J := edCombo.Items.IndexOfObject(ACustomFormat);
          if J <> -1 then
            edCombo.Items.Delete(J);

          edCombo.ItemIndex := edCombo.Items.IndexOfObject(AFormat);
          edCombo.Invalidate;
          exit;
        end;

  // predefined format is not found
  I := 0;
  while (I < edCombo.Items.Count) and (edCombo.Items.Objects[I] <> ACustomFormat) do Inc(I);
  if I >= edCombo.Items.Count then
    I := edCombo.Items.AddObject(' ', ACustomFormat);

  with TvgrRegisteredHeaderFooterFormat(edCombo.Items.Objects[I]) do
  begin
    FLeftSection := AHeaderFooter.LeftSection;
    FCenterSection := AHeaderFooter.CenterSection;
    FRightSection := AHeaderFooter.RightSection;
  end;
  edCombo.ItemIndex := I;
  edCombo.Invalidate;
end;

function TvgrPageSetupDialogForm.Execute(APageProperties: TvgrPageProperties): Boolean;
var
  I: Integer;
begin
  FPageProperties.Assign(APageProperties);
  // add predefined formats of header / footer
  for I := 0 to HeaderFooterFormats.Count - 1 do
  begin
    edHeader.Items.AddObject(' ', HeaderFooterFormats[I]);
    edFooter.Items.AddObject(' ', HeaderFooterFormats[I]);
  end;
  SetHeaderFooterFormat(PageProperties.Header, edHeader, FCustomHeaderFormat);
  SetHeaderFooterFormat(PageProperties.Footer, edFooter, FCustomFooterFormat);

  edMeasurementSystem.ItemIndex := Integer(PageProperties.MeasurementSystem);
  edPrinterNameClick(nil);
  UpdatePageOrientation;
  UpdateDefaults;
  UpdateHeaderFooter;

  Result := ShowModal = mrOk;
  if Result then
    APageProperties.Assign(PageProperties);
end;

procedure TvgrPageSetupDialogForm.DoLocalize;
begin
  inherited;

  edMeasurementSystem.Items.Clear;
  edMeasurementSystem.Items.Add(vgrLoadStr(svgrid_vgr_PageSetupDialog_MetricSystem));
  edMeasurementSystem.Items.Add(vgrLoadStr(svgrid_vgr_PageSetupDialog_USASystem));

  FSystemPrefix[vgrmsMetric] := vgrLoadStr(svgrid_Common_Mm);
  FSystemPrefix[vgrmsUSA] := vgrLoadStr(svgrid_Common_Inch);
end;

procedure TvgrPageSetupDialogForm.FormCreate(Sender: TObject);
begin
  FPageProperties := TvgrPageProperties.Create;
  FCustomHeaderFormat := TvgrRegisteredHeaderFooterFormat.Create;
  FCustomFooterFormat := TvgrRegisteredHeaderFooterFormat.Create;
end;

procedure TvgrPageSetupDialogForm.FormDestroy(Sender: TObject);
begin
  FPageProperties.Free;
  FCustomHeaderFormat.Free;
  FCustomFooterFormat.Free;
end;

procedure TvgrPageSetupDialogForm.edPrinterNameClick(Sender: TObject);
var
  I: Integer;
  APaperIndex: Integer;
  APaperSize: Integer;
  APaperName: string;
  APaperOrientation: TvgrPageOrientation;
  APaperDimensions: TPoint;
  AOldCursor: TCursor;
begin
  AOldCursor := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  BeginUpdate;
  try
    edPrefferedPageSize.Items.BeginUpdate;
    try
      edPrefferedPageSize.Items.Clear;
      if edPrinterName.SelectedPrinter = nil then
      begin
        for I := 0 to DEF_PAPERCOUNT - 1 do
          edPrefferedPageSize.Items.AddObject(PaperInfo[I].Name, Pointer(PaperInfo[I].Typ));
      end
      else
      begin
        with edPrinterName.SelectedPrinter do
          for I := 0 to PaperCount - 1 do
            edPrefferedPageSize.Items.AddObject(PaperNames[I], Pointer(PaperSizes[I]));
      end;
    finally
      edPrefferedPageSize.Items.EndUpdate;
    end;

    Printers.FindPaper(PageProperties,
                       edPrinterName.SelectedPrinter,
                       APaperIndex,
                       APaperSize,
                       APaperName,
                       APaperOrientation,
                       APaperDimensions);
    if APaperIndex = -1 then
    begin
      rbPredefinedPageSize.Checked := False;
      rbCustomPageSize.Checked := True;
      edPrefferedPageSize.ItemIndex := -1;
    end
    else
    begin
      rbPredefinedPageSize.Checked := True;
      rbCustomPageSize.Checked := False;
      edPrefferedPageSize.ItemIndex := edPrefferedPageSize.Items.IndexOf(APaperName);
    end;
    rbPagePortrait.Checked := APaperOrientation = vgrpoPortrait;
    rbPageLandscape.Checked := APaperOrientation = vgrpoLandscape;
    UpdateCustomPageSizeEdits;
    UpdatePageMargins;
    UpdateHeaderFooter;
    UpdateEnabled;
  finally
    EndUpdate;
    Screen.Cursor := AOldCursor;
  end;
end;

procedure TvgrPageSetupDialogForm.edMeasurementSystemClick(
  Sender: TObject);
begin
  PageProperties.MeasurementSystem := TvgrMeasurementSystem(edMeasurementSystem.ItemIndex);
  UpdateCustomPageSizeEdits;
  UpdatePageMargins;
  UpdateHeaderFooter;
  UpdateDefaults;
end;

procedure TvgrPageSetupDialogForm.rbPredefinedPageSizeClick(
  Sender: TObject);
begin
  UpdateEnabled;
end;

procedure TvgrPageSetupDialogForm.rbCustomPageSizeClick(Sender: TObject);
begin
  UpdateEnabled;
end;

procedure TvgrPageSetupDialogForm.edPrefferedPageSizeClick(
  Sender: TObject);
var
  ASize: TPoint;
begin
  ASize := Printers.GetPaperDimensionsBySize(Integer(edPrefferedPageSize.Items.Objects[edPrefferedPageSize.ItemIndex]), edPrinterName.SelectedPrinter);
  if rbPagePortrait.Checked then
  begin
    PageProperties.UnitsWidth[vgruTenthsMMs] := ASize.X;
    PageProperties.UnitsHeight[vgruTenthsMMs] := ASize.Y;
  end
  else
  begin
    PageProperties.UnitsWidth[vgruTenthsMMs] := ASize.Y;
    PageProperties.UnitsHeight[vgruTenthsMMs] := ASize.X;
  end;
  UpdateCustomPageSizeEdits;
  UpdatePageOrientation;
end;

procedure TvgrPageSetupDialogForm.edCustomPageWidthChange(Sender: TObject);
var
  AValue: Extended;
begin
  if EnableUpdate and TextToFloat(edCustomPageWidth.Text, AValue) then
  begin
    PageProperties.UnitsWidth[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue;
    UpdatePageOrientation;
  end;
end;

procedure TvgrPageSetupDialogForm.edCustomPageHeightChange(
  Sender: TObject);
var
  AValue: Extended;
begin
  if EnableUpdate and TextToFloat(edCustomPageHeight.Text, AValue) then
  begin
    PageProperties.UnitsHeight[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue;
    UpdatePageOrientation;
  end;
end;

procedure TvgrPageSetupDialogForm.rbPagePortraitClick(Sender: TObject);
var
  I: Integer;
begin
  if EnableUpdate then
  begin
    if rbPagePortrait.Checked then
    begin
      if PageProperties.Width > PageProperties.Height then
      begin
        I := PageProperties.Width;
        PageProperties.Width := PageProperties.Height;
        PageProperties.Height := I;
      end;
    end
    else
    begin
      if PageProperties.Height > PageProperties.Width then
      begin
        I := PageProperties.Width;
        PageProperties.Width := PageProperties.Height;
        PageProperties.Height := I;
      end;
    end;
    UpdateCustomPageSizeEdits;
  end;
end;

procedure TvgrPageSetupDialogForm.edMarginLeftChange(Sender: TObject);
var
  AValue: Extended;
begin
  if EnableUpdate and TextToFloat(TEdit(Sender).Text, AValue) then
  begin
    if Sender = edMarginLeft then
      PageProperties.Margins.UnitsLeft[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue
    else
      if Sender = edMarginTop then
        PageProperties.Margins.UnitsTop[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue
      else
        if Sender = edMarginRight then
          PageProperties.Margins.UnitsRight[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue
        else
          if Sender = edMarginBottom then
            PageProperties.Margins.UnitsBottom[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue;
    PagePreview.Invalidate;
  end;
end;

procedure TvgrPageSetupDialogForm.PagePreviewPaint(Sender: TObject);
var
  APageRect, ARect: TRect;
begin
  with PagePreview.Canvas do
  begin
    ARect := PagePreview.ClientRect;
    DrawFrame(PagePreview.Canvas, ARect, clBlack);
    InflateRect(ARect, -1, -1);

    APageRect := PagePreview.ClientRect;
    APageRect.Left := APageRect.Left + Round(PageProperties.Margins.UnitsLeft[vgruPixels] * PagePreviewKoef);
    APageRect.Top := APageRect.Top + Round(PageProperties.Margins.UnitsTop[vgruPixels] * PagePreviewKoef);
    APageRect.Right := APageRect.Right - Round(PageProperties.Margins.UnitsRight[vgruPixels] * PagePreviewKoef);
    APageRect.Bottom := APageRect.Bottom - Round(PageProperties.Margins.UnitsBottom[vgruPixels] * PagePreviewKoef);

    // margins
    Brush.Color := clSilver;
    FillRect(Rect(ARect.Left, ARect.Top, APageRect.Left, ARect.Bottom));
    FillRect(Rect(ARect.Left, ARect.Top, ARect.Right, APageRect.Top));
    FillRect(Rect(APageRect.Right, ARect.Top, ARect.Right, ARect.Bottom));
    FillRect(Rect(ARect.Left, APageRect.Bottom, ARect.Right, ARect.Bottom));

    InflateRect(APageRect, 1, 1);
    DrawFrame(PagePreview.Canvas, APageRect, clBlack);

    Brush.Color := clWhite;
    InflateRect(APageRect, -1, -1);
    FillRect(APageRect);

    if edPrinterName.SelectedPrinter <> nil then
    begin
      APageRect := PagePreview.ClientRect;
      with edPrinterName.SelectedPrinter.PageMargins do
      begin
        APageRect.Left := APageRect.Left + Round(UnitsLeft[vgruPixels] * PagePreviewKoef);
        APageRect.Top := APageRect.Top + Round(UnitsTop[vgruPixels] * PagePreviewKoef);
        APageRect.Right := APageRect.Right - Round(UnitsRight[vgruPixels] * PagePreviewKoef);
        APageRect.Bottom := APageRect.Bottom - Round(UnitsBottom[vgruPixels] * PagePreviewKoef);
      end;
      DrawFrame(PagePreview.Canvas, APageRect, clRed);
    end;
  end;
end;

procedure TvgrPageSetupDialogForm.edCustomPageWidthExit(Sender: TObject);
begin
  UpdateCustomPageSizeEdits;
end;

procedure TvgrPageSetupDialogForm.edMarginLeftExit(Sender: TObject);
begin
  UpdatePageMargins;
end;

procedure TvgrPageSetupDialogForm.edDefaultColWidthChange(Sender: TObject);
var
  AValue: Extended;
begin
  if EnableUpdate and TextToFloat(TEdit(Sender).Text, AValue) then
  begin
    if Sender = edDefaultColWidth then
      PageProperties.Defaults.UnitsColWidth[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue
    else
      if Sender = edDefaultRowHeight then
        PageProperties.Defaults.UnitsRowHeight[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue;
  end;
end;

procedure TvgrPageSetupDialogForm.edDefaultColWidthExit(Sender: TObject);
begin
  UpdateDefaults;
end;

procedure TvgrPageSetupDialogForm.bCustomHeaderClick(Sender: TObject);
begin
  if TvgrHeaderFooterOptionsDialogForm.Create(Self).Execute(PageProperties.Header) then
  begin
    SetHeaderFooterFormat(PageProperties.Header, edHeader, FCustomHeaderFormat);
  end;
end;

procedure TvgrPageSetupDialogForm.bCustomFooterClick(Sender: TObject);
begin
  if TvgrHeaderFooterOptionsDialogForm.Create(Self).Execute(PageProperties.Footer) then
  begin
    SetHeaderFooterFormat(PageProperties.Footer, edFooter, FCustomFooterFormat);
  end;
end;

procedure TvgrPageSetupDialogForm.edHeaderHeightChange(Sender: TObject);
var
  AValue: Extended;
begin
  if EnableUpdate and TextToFloat(TEdit(Sender).Text, AValue) then
  begin
    if Sender = edHeaderHeight then
      PageProperties.Header.UnitsHeight[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue
    else
      if Sender = edFooterHeight then
        PageProperties.Footer.UnitsHeight[ASystemToUnits[PageProperties.MeasurementSystem]] := AValue;
  end;
end;

procedure TvgrPageSetupDialogForm.edHeaderHeightExit(Sender: TObject);
begin
  UpdateHeaderFooter;
end;

/////////////////////////////////////////////////
//
// TvgrRegisteredHeaderFooterFormats
//
/////////////////////////////////////////////////
function TvgrRegisteredHeaderFooterFormats.GetItem(Index: Integer): TvgrRegisteredHeaderFooterFormat;
begin
  Result := TvgrRegisteredHeaderFooterFormat(inherited Items[Index]);
end;

procedure TvgrRegisteredHeaderFooterFormats.RegisterFormat(const ALeftSection, ACenterSection, ARightSection: string);
var
  I: Integer;
  AFormat: TvgrRegisteredHeaderFooterFormat;
begin
  for I := 0 to Count - 1 do
    with Items[I] do
      if (LeftSection = ALeftSection) and (CenterSection = ACenterSection) and (RightSection = ARightSection) then
        exit;

  AFormat := TvgrRegisteredHeaderFooterFormat.Create;
  with AFormat do
  begin
    FLeftSection := ALeftSection;
    FCenterSection := ACenterSection;
    FRightSection := ARightSection;
  end;
  Add(AFormat);
end;

procedure TvgrPageSetupDialogForm.edHeaderClick(Sender: TObject);

  procedure SetParams(AHeaderFooter: TvgrPageHeaderFooter; edCombo: TComboBox);
  begin
    if edCombo.Items.Objects[edCombo.ItemIndex] is TvgrRegisteredHeaderFooterFormat then
      with TvgrRegisteredHeaderFooterFormat(edCombo.Items.Objects[edCombo.ItemIndex]) do
      begin
        AHeaderFooter.LeftSection := LeftSection;
        AHeaderFooter.CenterSection := CenterSection;
        AHeaderFooter.RightSection := RightSection;
      end;
  end;

begin
  //
  if (Sender = edHeader) and (edHeader.ItemIndex >= 0) then
    SetParams(PageProperties.Header, edHeader)
  else
    if (Sender = edFooter) and (edFooter.ItemIndex >= 0) then
      SetParams(PageProperties.Footer, edFooter);
end;

procedure TvgrPageSetupDialogForm.edHeaderDrawItem(Control: TWinControl;
  Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  ALeft, ACenter, ARight: string;

  function GetText(const S: string): string;
  begin
    Result := ParseHeaderFooterText(S, 1, 2, 'Sheet1');
  end;
  
begin
  with TComboBox(Control) do
  begin
    if (Index < 0) or (Index >= Items.Count) then
      exit;

    if Items.Objects[Index] is TvgrRegisteredHeaderFooterFormat then
      with TvgrRegisteredHeaderFooterFormat(Items.Objects[Index]) do
      begin
        ALeft := LeftSection;
        ACenter := CenterSection;
        ARight := RightSection;
      end
    else
      exit;

    Canvas.FillRect(Rect);
    InflateRect(Rect, -2, -1);
    if ALeft <> '' then
    begin
      vgrDrawText(Canvas, GetText(ALeft), Rect, true, vgrhaLeft, vgrvaTop, 0);
    end;
    if ACenter <> '' then
    begin
      vgrDrawText(Canvas, GetText(ACenter), Rect, true, vgrhaCenter, vgrvaTop, 0);
    end;
    if ARight <> '' then
    begin
      vgrDrawText(Canvas, GetText(ARight), Rect, true, vgrhaRight, vgrvaTop, 0);
    end;
  end;
end;

initialization

  HeaderFooterFormats := TvgrRegisteredHeaderFooterFormats.Create;
  // [] [Page 1] []
  HeaderFooterFormats.RegisterFormat('', Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_HF_Page), [cHFKeyPage]), '');
  // [] [Page 1 of X] []
  HeaderFooterFormats.RegisterFormat('', Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_HF_PageOf), [cHFKeyPage, cHFKeyPages]), '');
  // [] [] [Page 1 of X]
  HeaderFooterFormats.RegisterFormat('', '', Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_HF_PageOf), [cHFKeyPage, cHFKeyPages]));
  // [] [Sheet1] []
  HeaderFooterFormats.RegisterFormat('', cHFKeyTab, '');
  // [] [Sheet1] [Page 1 of X]
  HeaderFooterFormats.RegisterFormat('', cHFKeyTab, Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_HF_PageOf), [cHFKeyPage, cHFKeyPages]));
  // [] [21.10.2003] [Page 1 of X]
  HeaderFooterFormats.RegisterFormat('', cHFKeyDate, Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_HF_PageOf), [cHFKeyPage, cHFKeyPages]));
  // [10:00] [Sheet1] [Page 1 of X]
  HeaderFooterFormats.RegisterFormat(cHFKeyTime, cHFKeyTab, Format(vgrLoadStr(svgrid_vgr_PageSetupDialog_HF_PageOf), [cHFKeyPage, cHFKeyPages]));

finalization

  HeaderFooterFormats.Free;

end.

