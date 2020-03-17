{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains TvgrBandFormatForm class, which realize a dialog form for editing ranges and borders of TvgrWorksheet.
Also this module contains a TvgrCellPropertiesDialog class - non
visual component which manage by TvgrBandFormatForm form. }
unit vgr_CellPropertiesDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils,
  {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, vgr_WorkbookGrid, vgr_ColorButton, ExtCtrls, StdCtrls, ComCtrls,
  Buttons, vgr_CommonClasses, vgr_DataStorage, vgr_DataStorageTypes,
  vgr_Label, vgr_Form, vgr_FormLocalizer, vgr_Button;

type
{Specifies a type of range value, selected by user.
Syntax:
  TvgrFormatCategory = (vgrfcNumeric, vgrfcDate, vgrfcTime, vgrfcDateTime, vgrfcText);}
  TvgrFormatCategory = (vgrfcNumeric, vgrfcDate, 
                          vgrfcTime, vgrfcDateTime, vgrfcText);
  
type
  TvgrCellPropertiesForm = class;
  
{Specifies the types of changes, which can be made by user
Items:
  vgrdcHorzAlign          - horizontal text aligment are changed
  vgrdcVertAlign          - vertical text alignment are changed
  vgrdcWordWrap           - WordWrap are changed
  vgrdcMerge              - merge of cells are changed
  vgrdcAngle              - angle of text rotation are changed
  vgrdcFontName           - name of font are changed
  vgrdcFontSize           - size of font are changed
  vgrdcFontStyleBold      - font bold are changed
  vgrdcFontStyleItalic    - font italic are changed
  vgrdcFontStyleUnderline - font underline are changed
  vgrdcFontStyleStrikeout - font strikeout are changed
  vgrdcFontColor          - color of font are changed
  vgrdcLeftBorder         - properties of left border are changed
  vgrdcVertCenterBorder   - properties of internal vertical border are changed
  vgrdcRightBorder        - properties of right border are changed
  vgrdcTopBorder          - properties of top border are changed
  vgrdcHorzCenterBorder   - properties of internal horizontal border are changed
  vgrdcBottomBorder       - properties of bottom border are changed
  vgrdcFillBackColor      - background fill color are changed
  vgrdcFillForeColor      - foregraund fill color are changed
  vgrdcFillPattern        - pattern of fill are changed
  vgrdcFontCharset        - charset of font are changed
  vgrdcDisplayFormat      - display format are changed}
  TvgeCellPropertiesDialogChange = (vgrdcHorzAlign,
                                    vgrdcVertAlign,
                                    vgrdcWordWrap,
                                    vgrdcMerge,
                                    vgrdcAngle,
                                    vgrdcFontName,
                                    vgrdcFontSize,
                                    vgrdcFontStyleBold,
                                    vgrdcFontStyleItalic,
                                    vgrdcFontStyleUnderline,
                                    vgrdcFontStyleStrikeout,
                                    vgrdcFontColor,
                                    vgrdcLeftBorder,
                                    vgrdcVertCenterBorder,
                                    vgrdcRightBorder,
                                    vgrdcTopBorder,
                                    vgrdcHorzCenterBorder,
                                    vgrdcBottomBorder,
                                    vgrdcFillBackColor,
                                    vgrdcFillForeColor,
                                    vgrdcFillPattern,
                                    vgrdcFontCharset,
                                    vgrdcDisplayFormat);
{Describes the types of changes, which can be made by user.
Syntax:
  TvgeCellPropertiesDialogChanges = set of TvgeCellPropertiesDialogChange;}
  TvgeCellPropertiesDialogChanges = set of TvgeCellPropertiesDialogChange;

  /////////////////////////////////////////////////
  //
  // TvgrCellPropertiesDialog
  //
  /////////////////////////////////////////////////
{The TvgrCellPropertiesDialog component displays a dialog box for editing properties of ranges.
The dialog does not appear at runtime until it is activated by a call to the Execute method.
When the user clicks OK, the dialog closes and all changes is stored in the ranges.}
  TvgrCellPropertiesDialog = class(TComponent)
  private
    FForm: TvgrCellPropertiesForm;
  public
{Execute opens the dialog.
Parameters:
  AWorksheet - worksheet for editing
  ACellsRects - specifies a area of worksheet for editing.
Return value:
  Boolean, returning true when the user edit properties and clicks OK,
or false when the user cancels.
Example:
  begin
  ...
  vgrCellPropertiesDialog1.Execute(vgrWorkbookGrid1.ActiveWorksheet,
                                   vgrWorkbookGrid1.SelectionRects);
  ...
  end;}
    function Execute(AWorksheet: TvgrWorksheet; const ACellsRects: TvgrRectArray): Boolean;
{Returns reference to a TvgrCellPropertiesForm form, created by component in Execute method.}
    property Form: TvgrCellPropertiesForm read FForm;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrValueFormat
  //
  /////////////////////////////////////////////////
{Internal class, used to store properties of DisplayFormat.}
  TvgrValueFormat = class(TObject)
  private
    FCategory: TvgrFormatCategory;
    FFormat: string;
    FIsBuiltIn: Boolean;
  public
    constructor Create(ACategory: TvgrFormatCategory; const AFormat: string; AIsBuiltIn: Boolean);
    { Category of DisplayFormat. }
    property Category: TvgrFormatCategory read FCategory;
    { Value of DisplayFormat. }
    property Format: string read FFormat;
    { Returns True is format are builtin, and is not entered by the user. }
    property IsBuiltIn: Boolean read FIsBuiltIn;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrValueFormats
  //
  /////////////////////////////////////////////////
{Internal class, used to store a list of TvgrValueFormat objects.}
  TvgrValueFormats = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrValueFormat;
  public
    function RegisterValueFormat(ACategory: TvgrFormatCategory; const AFormat: string; AIsBuiltIn: Boolean): TvgrValueFormat;
    function IndexOfValueFormat(ACategory: TvgrFormatCategory; const AFormat: string): Integer;
    function IndexByFormat(const AFormat: string): Integer;
    property Items[Index: Integer]: TvgrValueFormat read GetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrCellPropertiesForm
  //
  /////////////////////////////////////////////////
{Implements a dialog form for editing ranges and borders of TvgrWorksheet.}
  TvgrCellPropertiesForm = class(TvgrDialogForm)
    PreviewBordersWorkbook: TvgrWorkbook;
    PreviewFillWorkBook: TvgrWorkbook;
    Panel1: TPanel;
    bOk: TButton;
    bCancel: TButton;
    PageControl1: TPageControl;
    Alignment: TTabSheet;
    Label1: TvgrBevelLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TvgrBevelLabel;
    Bevel6: TBevel;
    bLeftTop: TSpeedButton;
    bCenterTop: TSpeedButton;
    bRightTop: TSpeedButton;
    bLeftCenter: TSpeedButton;
    bCenterCenter: TSpeedButton;
    bRightCenter: TSpeedButton;
    bLeftBottom: TSpeedButton;
    bCenterBottom: TSpeedButton;
    bRightBottom: TSpeedButton;
    bDefault: TSpeedButton;
    cbWrapText: TCheckBox;
    cbMergeCells: TCheckBox;
    edHorzAlign: TComboBox;
    edVertAlign: TComboBox;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    udAngle: TUpDown;
    edAngle: TEdit;
    TabSheet1: TTabSheet;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Bevel5: TBevel;
    edFontName: TEdit;
    edFontSize: TEdit;
    lbFontName: TListBox;
    lbFontSize: TListBox;
    GroupBox3: TGroupBox;
    cbFontStyleBold: TCheckBox;
    cbFontStyleItalic: TCheckBox;
    cbFontStyleUnderline: TCheckBox;
    cbFontStyleStrikeout: TCheckBox;
    bFontColor: TvgrColorButton;
    TabSheet2: TTabSheet;
    Label14: TvgrBevelLabel;
    bNoneBorders: TSpeedButton;
    bOutlineBorders: TSpeedButton;
    bInsideBorders: TSpeedButton;
    Label20: TLabel;
    Label16: TvgrBevelLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    bTopBorder: TSpeedButton;
    bCenterHorzBorder: TSpeedButton;
    bBottomBorder: TSpeedButton;
    bLeftBorder: TSpeedButton;
    bCenterVertBorder: TSpeedButton;
    bRightBorder: TSpeedButton;
    PreviewBordersGrid: TvgrWorkbookGrid;
    GroupBox4: TGroupBox;
    Label21: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    bBorderColor: TvgrColorButton;
    edBorderStyle: TComboBox;
    edBorderWidth: TEdit;
    udBorderWidth: TUpDown;
    Patterns: TTabSheet;
    Label22: TvgrBevelLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label23: TLabel;
    bFillBackColor: TvgrColorButton;
    bFillForeColor: TvgrColorButton;
    edFillPattern: TComboBox;
    GroupBox2: TGroupBox;
    PreviewFillGrid: TvgrWorkbookGrid;
    Label5: TLabel;
    FontPreview: TLabel;
    edCharSet: TComboBox;
    Label24: TLabel;
    vgrFormLocalizer1: TvgrFormLocalizer;
    Formats: TTabSheet;
    vgrBevelLabel1: TvgrBevelLabel;
    lbFormats: TListBox;
    GroupBox5: TGroupBox;
    edExample: TEdit;
    vgrBevelLabel2: TvgrBevelLabel;
    edType: TEdit;
    lbType: TListBox;
    bAddFormat: TButton;
    bDeleteFormat: TButton;
    Label25: TLabel;
    Label26: TLabel;
    procedure edHorzAlignClick(Sender: TObject);
    procedure edVertAlignClick(Sender: TObject);
    procedure bLeftTopClick(Sender: TObject);
    procedure cbWrapTextClick(Sender: TObject);
    procedure cbMergeCellsClick(Sender: TObject);
    procedure edFontNameChange(Sender: TObject);
    procedure lbFontNameClick(Sender: TObject);
    procedure cbFontStyleBoldClick(Sender: TObject);
    procedure cbFontStyleItalicClick(Sender: TObject);
    procedure cbFontStyleUnderlineClick(Sender: TObject);
    procedure cbFontStyleStrikeoutClick(Sender: TObject);
    procedure edFontSizeChange(Sender: TObject);
    procedure bNoneBordersClick(Sender: TObject);
    procedure bOutlineBordersClick(Sender: TObject);
    procedure bInsideBordersClick(Sender: TObject);
    procedure bTopBorderClick(Sender: TObject);
    procedure edFillPatternClick(Sender: TObject);
    procedure bFontColorColorSelected(Sender: TObject; Color: TColor);
    procedure bFillBackColorColorSelected(Sender: TObject; Color: TColor);
    procedure bFillForeColorColorSelected(Sender: TObject; Color: TColor);
    procedure lbFontSizeClick(Sender: TObject);
    procedure edAngleChange(Sender: TObject);
    procedure edCharSetClick(Sender: TObject);
    procedure lbFormatsClick(Sender: TObject);
    procedure lbTypeClick(Sender: TObject);
    procedure edTypeChange(Sender: TObject);
    procedure bAddFormatClick(Sender: TObject);
    procedure bDeleteFormatClick(Sender: TObject);
  private
    { Private declarations }
    FWorksheet: TvgrWorksheet;
    FCellsRects: TvgrRectArray;
    FLockCount: Integer;
    FPreviewFillFormat: IvgrRangesFormat;
    FPreviewBordersFormat: IvgrRangesFormat;
    FFoundedRanges: TInterfaceList;
    FFoundedRangesSquare: Integer;
    FChanges: TvgeCellPropertiesDialogChanges;

    function GetPreviewFill: TvgrWorksheet;
    function GetPreviewBorders: TvgrWorksheet;
    procedure UpdateFontSize(AFontSize: Integer);
    procedure UpdateFontName(const AFontName: string);
    procedure UpdateFontCharset(const AFontName: string; ACharset: Integer);
    procedure UpdateFontStyle;
    procedure UpdateFontColor(AFontColor: TColor);
    procedure UpdateFillBackColor(AFillBackColor: TColor);
    procedure UpdateFillForeColor(AFillForeColor: TColor);
    procedure UpdateFillPattern(AFillPattern: TBrushStyle);
    procedure UpdateBorderButtons;
    procedure UpdateAlignButtons;
    procedure FindMergeRangesCallBack(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure FindMergeRanges(const ARect: TRect);
    function MergeWarningProc: Boolean;
    function GetEventsEnabled: Boolean;
    procedure SetBordersToPreview(ACellsBorders: TvgrBorderTypes;
                                  ABorderStyle: TvgrBorderStyle;
                                  ABorderWidth: Integer;
                                  ABorderColor: TColor);
    function GetSelectedFormatCategory: TvgrFormatCategory;
    function GetDefaultFormatValue(ACategory: Integer = -1): Variant;
    procedure UpdateFormatControls(AVisible, AEnabled: Boolean);
    function AddFormatType(AFormat: TvgrValueFormat): Integer;
  protected
    procedure DoShow; override;

    procedure BeginUpdate;
    procedure EndUpdate;

    procedure DoLocalize; override;

    property Changes: TvgeCellPropertiesDialogChanges read FChanges;
    property PreviewFill: TvgrWorksheet read GetPreviewFill;
    property PreviewBorders: TvgrWorksheet read GetPreviewBorders;
    property PreviewFillFormat: IvgrRangesFormat read FPreviewFillFormat;
    property PreviewBordersFormat: IvgrRangesFormat read FPreviewBordersFormat;
    property EventsEnabled: Boolean read GetEventsEnabled;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;

{Execute opens the dialog.
Parameters:
  AWorksheet  - worksheet for editing
  ACellsRects - specifies a area of worksheet for editing.
Return value:
  Boolean, returning true when the user edit properties and clicks OK,
or false when the user cancels.}
    function Execute(AWorksheet: TvgrWorksheet; const ACellsRects: TvgrRectArray): Boolean;
  end;

implementation

uses Math, vgr_GUIFunctions, vgr_Functions, vgr_Dialogs, vgr_StringIDs, vgr_Localize;

const
  vgrFormatCategoryAll = Integer(High(TvgrFormatCategory)) + 1;
  vgrFormatCategoryCommon = vgrFormatCategoryAll + 1;

var
  RegisteredValueFormats: TvgrValueFormats;

{$R *.dfm}
{$R ..\res\vgr_CellPropertiesDialogStrings.res}

function RegisterValueFormat(ACategory: TvgrFormatCategory; const AFormat: string; AIsBuiltIn: Boolean): TvgrValueFormat;
begin
  Result := RegisteredValueFormats.RegisterValueFormat(ACategory, AFormat, AIsBuiltIn);
end;

/////////////////////////////////////////////////
//
// TvgrValueFormat
//
/////////////////////////////////////////////////
constructor TvgrValueFormat.Create(ACategory: TvgrFormatCategory; const AFormat: string; AIsBuiltIn: Boolean);
begin
  inherited Create;
  FCategory := ACategory;
  FFormat := AFormat;
  FIsBuiltIn := AIsBuiltIn;
end;

/////////////////////////////////////////////////
//
// TvgrValueFormats
//
/////////////////////////////////////////////////
function TvgrValueFormats.GetItem(Index: Integer): TvgrValueFormat;
begin
  Result := TvgrValueFormat(inherited Items[Index]);
end;

function TvgrValueFormats.IndexOfValueFormat(ACategory: TvgrFormatCategory; const AFormat: string): Integer;
begin
  Result := 0;
  while (Result < Count) and not ((Items[Result].Category = ACategory) and (Items[Result].Format = AFormat)) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrValueFormats.IndexByFormat(const AFormat: string): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].Format <> AFormat) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrValueFormats.RegisterValueFormat(ACategory: TvgrFormatCategory; const AFormat: string; AIsBuiltIn: Boolean): TvgrValueFormat;
begin
  if IndexOfValueFormat(ACategory, AFormat) = -1 then
    Result := Items[Add(TvgrValueFormat.Create(ACategory, AFormat, AIsBuiltIn))]
  else
    Result := nil;
end;

/////////////////////////////////////////////////
//
// TvgrCellPropertiesDialog
//
/////////////////////////////////////////////////
function TvgrCellPropertiesDialog.Execute(AWorksheet: TvgrWorksheet; const ACellsRects: TvgrRectArray): Boolean;
begin
  FForm := TvgrCellPropertiesForm.Create(nil);
  Result := FForm.Execute(AWorksheet, ACellsRects);
  FForm := nil;
end;

/////////////////////////////////////////////////
//
// TvgrCellPropertiesForm
//
/////////////////////////////////////////////////
procedure TvgrCellPropertiesForm.DoShow;
begin
  FChanges := [];
end;

procedure TvgrCellPropertiesForm.AfterConstruction;
begin
  inherited;
  FFoundedRanges := TInterfaceList.Create;

  PreviewFillWorkBook.AddWorksheet;
  with PreviewFill.Ranges[1, 1, 1, 1] do
  begin
    Value := 'Text';
    HorzAlign := vgrhaCenter;
    VertAlign := vgrvaCenter;
  end;
  with PreviewFill do
  begin
    Rows[0].Height := 100;
    Rows[1].Height := 1300;
    Rows[2].Height := 100;
    Cols[0].Width := 100;
    Cols[1].Width := 2035;
    Cols[2].Width := 100;
  end;
  PreviewFillGrid.ActiveWorksheetIndex := 0;
  PreviewFillGrid.SetSelection(Rect(3, 3, 3, 3));

  PreviewBordersWorkbook.AddWorksheet;
  with PreviewBorders do
  begin
    Rows[0].Height := 100;
    Rows[1].Height := 670;
    Rows[2].Height := 670;
    Rows[3].Height := 100;
    Cols[0].Width := 100;
    Cols[1].Width := 1475;
    Cols[2].Width := 1475;
    Cols[3].Width := 100;
  end;
  PreviewBordersGrid.ActiveWorksheetIndex := 0;
  PreviewBordersGrid.SetSelection(Rect(4, 4, 4, 4));

  lbFontName.Items := Screen.Fonts;

  edBorderStyle.ItemIndex := 0;

  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryCommon), Pointer(vgrFormatCategoryCommon));
  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryNumeric),
                            Pointer(vgrfcNumeric));
  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryDate),
                            Pointer(vgrfcDate));
  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryTime),
                            Pointer(vgrfcTime));
  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryDateTime),
                            Pointer(vgrfcDateTime));
  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryText),
                            Pointer(vgrfcText));
  lbFormats.Items.AddObject(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_FormatCategoryAll), Pointer(vgrFormatCategoryAll));
end;

procedure TvgrCellPropertiesForm.BeforeDestruction;
begin
  inherited;
  FreeAndNil(FFoundedRanges);
end;

function TvgrCellPropertiesForm.GetPreviewFill: TvgrWorksheet;
begin
  Result := PreviewFillWorkbook.Worksheets[0];
end;

function TvgrCellPropertiesForm.GetPreviewBorders: TvgrWorksheet;
begin
  Result := PreviewBordersWorkbook.Worksheets[0];
end;

procedure TvgrCellPropertiesForm.UpdateFontSize(AFontSize: Integer);
begin
  BeginUpdate;
  try
    edFontSize.Text := IntToStr(AFontSize);
    lbFontSize.ItemIndex := lbFontSize.Items.IndexOf(IntToStr(AFontSize));
    FontPreview.Font.Size := AFontSize;
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFontName(const AFontName: string);
begin
  BeginUpdate;
  try
    edFontName.Text := AFontName;
    lbFontName.ItemIndex := lbFontName.Items.IndexOf(AFontName);
    FontPreview.Font.Name := AFontName;
    UpdateFontCharset(AFontName, -1);
  finally
    EndUpdate;
  end;
end;

function FontEnumProc(var LogFont: TLogFont; var TextMetric: TTextMetric;
  FontType: Integer; Data: Pointer): Integer; stdcall;
begin
  with TvgrCellPropertiesForm(Data) do
  begin
    if edCharset.Items.IndexOfObject(Pointer(LogFont.lfCharSet)) = -1 then
      edCharSet.Items.AddObject(vgrCharsetToIdent(LogFont.lfCharSet), Pointer(LogFont.lfCharSet));
  end;
  Result := 1;
end;

procedure TvgrCellPropertiesForm.UpdateFontCharset(const AFontName: string; ACharset: Integer);
var
  ALogFont: TLogFont;
  ADC: HDC;
  I, AOldCharset: Integer;
begin
  BeginUpdate;
  ADC := GetDC(0);
  try
    if ACharset <> -1 then
      AOldCharset := ACharset
    else
    begin
      if edCharSet.ItemIndex >= 0 then
        AOldCharset := Integer(edCharSet.Items.Objects[edCharSet.ItemIndex])
      else
        AOldCharset := -1;
    end;
    edCharSet.Items.Clear;
    edCharSet.Items.AddObject(vgrCharsetToIdent(DEFAULT_CHARSET), Pointer(DEFAULT_CHARSET));
    if (ACharset <> -1) and (ACharset <> DEFAULT_CHARSET) then
      edCharSet.Items.AddObject(vgrCharsetToIdent(ACharset), Pointer(ACharset));

    FillChar(ALogFont, SizeOf(ALogFont), 0);
    StrPCopy(ALogFont.lfFaceName, AFontName);
    ALogFont.lfCharSet := DEFAULT_CHARSET;
    EnumFontFamiliesEx(ADC, ALogFont, @FontEnumProc, Integer(Self), 0);

    if AOldCharSet = -1  then
      edCharSet.ItemIndex := 0
    else
    begin
      I := edCharSet.Items.IndexOfObject(Pointer(AOldCharSet));
      if I <> -1 then
        edCharSet.ItemIndex := I
      else
        edCharSet.ItemIndex := 0;
    end;
  finally
    ReleaseDC(0, ADC);
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFontStyle;

  procedure CheckButton(AButton: TCheckBox; AItem: TFontStyle);
  begin
    if AButton.State <> cbGrayed then
    begin
      if AButton.Checked then
        FontPreview.Font.Style := FontPreview.Font.Style + [AItem]
      else
        FontPreview.Font.Style := FontPreview.Font.Style - [AItem];
    end;
  end;

begin
  BeginUpdate;
  try
    CheckButton(cbFontStyleBold, fsBold);
    CheckButton(cbFontStyleItalic, fsItalic);
    CheckButton(cbFontStyleUnderline, fsUnderline);
    CheckButton(cbFontStyleStrikeout, fsStrikeOut);
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFontColor(AFontColor: TColor);
begin
  BeginUpdate;
  try
    bFontColor.SelectedColor := AFontColor;
    FontPreview.Font.Color := AFontColor;
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFillBackColor(AFillBackColor: TColor);
begin
  BeginUpdate;
  try
    bFillBackColor.SelectedColor := AFillBackColor;
    PreviewFillFormat.FillBackColor := AFillBackColor;
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFillForeColor(AFillForeColor: TColor);
begin
  BeginUpdate;
  try
    bFillForeColor.SelectedColor := AFillForeColor;
    PreviewFillFormat.FillForeColor := AFillForeColor;
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFillPattern(AFillPattern: TBrushStyle);
begin
  BeginUpdate;
  try
    edFillPattern.ItemIndex := Integer(AFillPattern);
    PreviewFillFormat.FillPattern := AFillPattern;
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateBorderButtons;

  procedure UpdateButton(AButton: TSpeedButton; ABorder: TvgrRangesFormatBorder);
  begin
    AButton.Down := ABorder.HasWidth and (ABorder.Width > 0);
  end;

begin
  BeginUpdate;
  try
    with PreviewBordersFormat.Borders do
    begin
      UpdateButton(bTopBorder, Top);
      UpdateButton(bCenterHorzBorder, Middle);
      UpdateButton(bBottomBorder, Bottom);

      UpdateButton(bLeftBorder, Left);
      UpdateButton(bCenterVertBorder, Center);
      UpdateButton(bRightBorder, Right);
    end;
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.UpdateAlignButtons;
begin
  BeginUpdate;
  try
    bLeftTop.Down := (edHorzAlign.ItemIndex = 1) and (edVertAlign.ItemIndex = 0);
    bCenterTop.Down := (edHorzAlign.ItemIndex = 2) and (edVertAlign.ItemIndex = 0);
    bRightTop.Down := (edHorzAlign.ItemIndex = 3) and (edVertAlign.ItemIndex = 0);

    bLeftCenter.Down := (edHorzAlign.ItemIndex = 1) and (edVertAlign.ItemIndex = 1);
    bCenterCenter.Down := (edHorzAlign.ItemIndex = 2) and (edVertAlign.ItemIndex = 1);
    bRightCenter.Down := (edHorzAlign.ItemIndex = 3) and (edVertAlign.ItemIndex = 1);

    bLeftBottom.Down := (edHorzAlign.ItemIndex = 1) and (edVertAlign.ItemIndex = 2);
    bCenterBottom.Down := (edHorzAlign.ItemIndex = 2) and (edVertAlign.ItemIndex = 2);
    bRightBottom.Down := (edHorzAlign.ItemIndex = 3) and (edVertAlign.ItemIndex = 2);
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.FindMergeRangesCallBack(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrRange do
    if (Left <> Right) or (Top <> Bottom) then
    begin
      FFoundedRanges.Add(AItem as IvgrRange);
      FFoundedRangesSquare := FFoundedRangesSquare + (Right - Left + 1) * (Bottom - Top + 1);
    end;
end;

procedure TvgrCellPropertiesForm.FindMergeRanges(const ARect: TRect);
begin
  FFoundedRanges.Clear;
  FFoundedRangesSquare := 0;
  FWorksheet.RangesList.FindAndCallBack(ARect, FindMergeRangesCallback, nil);
end;

procedure TvgrCellPropertiesForm.BeginUpdate;
begin
  Inc(FLockCount);
end;

procedure TvgrCellPropertiesForm.EndUpdate;
begin
  Dec(FLockCount);
end;

function TvgrCellPropertiesForm.MergeWarningProc: Boolean;
begin
  Result := MBox(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_MergeWarning), MB_OKCANCEL or MB_ICONEXCLAMATION) = IDOK;
end;

function TvgrCellPropertiesForm.GetEventsEnabled: Boolean;
begin
  Result := FLockCount = 0;
end;

function TvgrCellPropertiesForm.Execute(AWorksheet: TvgrWorksheet; const ACellsRects: TvgrRectArray): Boolean;
const
  //  vgrfcNumeric, vgrfcDate, vgrfcTime, vgrfcDateTime, vgrfcText
  AConvert: array [TvgrFormatCategory] of TvgrRangeValueType = (rvtExtended, rvtDateTime, rvtDateTime, rvtDateTime, rvtString);
var
  I, ASquare: Integer;
  AFormat: IvgrRangesFormat;
  ARects: TvgrRectArray;
  ADisplayFormat, AFontName: string;
  AValueFormat: TvgrValueFormat;

  function GetCheckBoxState(ARangesFlag: TvgrRangesFlag; AValue: Boolean): TCheckBoxState;
  const
    AState : Array [Boolean] of TCheckBoxState = (cbUnchecked, cbChecked);
  begin
    if ARangesFlag in AFormat.HasFlags then
      Result := AState[AValue]
    else
      Result := cbGrayed;
  end;
  
begin
  FWorksheet := AWorksheet;
  FCellsRects := ACellsRects;

  SetLength(ARects, 1);
  // Init Preview fill workbook
  ARects[0] := Rect(1, 1, 1, 1);
  FPreviewFillFormat := PreviewFill.RangesFormat[ARects, False];

  // Init Preview borders workbook
  ARects[0] := Rect(1, 1, 2, 2);
  FPreviewBordersFormat := PreviewBorders.RangesFormat[ARects, True];

  AFormat := AWorksheet.RangesFormat[FCellsRects, True];

  BeginUpdate;
  try
    with AFormat do
    begin
      if vgrrfFontSize in HasFlags then
        UpdateFontSize(FontSize);
      if vgrrfFontName in HasFlags then
      begin
        UpdateFontName(FontName);
        AFontName := FontName;
      end
      else
        AFontName := '';
      if vgrrfFontCharset in HasFlags then
        UpdateFontCharset(AFontName, FontCharset);
      cbFontStyleBold.State := GetCheckBoxState(vgrrfFontStyleBold, FontStyleBold);
      cbFontStyleItalic.State := GetCheckBoxState(vgrrfFontStyleItalic, FontStyleItalic);
      cbFontStyleUnderline.State := GetCheckBoxState(vgrrfFontStyleUnderline, FontStyleUnderline);
      cbFontstyleStrikeout.State := GetCheckBoxState(vgrrfFontStyleStrikeout, FontStyleStrikeout);
      if vgrrfFontColor in HasFlags then
        UpdateFontColor(FontColor);

      cbWrapText.State := GetCheckBoxState(vgrrfWordWrap, WordWrap);

      if vgrrfFillBackColor in HasFlags then
        UpdateFillBackColor(FillBackColor);
      if vgrrfFillForeColor in HasFlags then
        UpdateFillForeColor(FillForeColor);
      if vgrrfFillPattern in HasFlags then
        UpdateFillPattern(FillPattern);

      PreviewBordersFormat.Borders.AssignBordersStyles(AFormat.Borders);
      UpdateBorderButtons;

      if vgrrfHorzAlign in HasFlags then
        edHorzAlign.ItemIndex := Integer(HorzAlign);
      if vgrrfVertAlign in HasFlags then
        edVertAlign.ItemIndex := Integer(VertAlign);
      UpdateAlignButtons;

      if vgrrfAngle in HasFlags then
        udAngle.Position := Angle;
    end;

    // Merge
    ASquare := 0;
    for I := 0 to High(FCellsRects) do
    begin
      FindMergeRanges(FCellsRects[I]);
      with FCellsRects[I] do
        ASquare := ASquare + (Right - Left + 1) * (Bottom - Top + 1);
    end;
    if FFoundedRanges.Count > 0 then
    begin
      if FFoundedRangesSquare >= ASquare then
        cbMergeCells.State := cbChecked
      else
        cbMergeCells.State := cbGrayed;
    end
    else
      cbMergeCells.State := cbUnchecked;
    FFoundedRanges.Clear;

    // DisplayFormat
    ADisplayFormat := AFormat.DisplayFormat;
    if (ADisplayFormat = '') or not (vgrrfDisplayFormat in AFormat.HasFlags) then
    begin
      lbFormats.ItemIndex := lbFormats.Items.IndexOfObject(Pointer(vgrFormatCategoryCommon));
      lbFormatsClick(nil);
    end
    else
    begin
      I := RegisteredValueFormats.IndexByFormat(ADisplayFormat);
      if I = -1 then
      begin
        lbFormats.ItemIndex := lbFormats.Items.IndexOfObject(Pointer(vgrFormatCategoryAll));
        lbFormatsClick(nil);
        edType.Text := ADisplayFormat;
        edTypeChange(nil);
      end
      else
      begin
        AValueFormat := RegisteredValueFormats[I];
        lbFormats.ItemIndex := lbFormats.Items.IndexOfObject(Pointer(AValueFormat.Category));
        lbFormatsClick(nil);
        lbType.ItemIndex := lbType.Items.IndexOfObject(AValueFormat);
        lbTypeClick(nil);
      end;
    end;

  finally
    EndUpdate;
  end;

  Result := ShowModal = mrOk;
  if Result then
  begin
    AWorksheet.Workbook.BeginUpdate;
    try
      AFormat := AWorksheet.RangesFormat[FCellsRects, True];
      with AFormat do
      begin
        if vgrdcHorzAlign in Changes then
          HorzAlign := TvgrRangeHorzAlign(edHorzAlign.ItemIndex);
        if vgrdcVertAlign in Changes then
          VertAlign := TvgrRangeVertAlign(edVertAlign.ItemIndex);
        if vgrdcWordWrap in Changes then
          WordWrap := cbWrapText.Checked;
        if vgrdcMerge in Changes then
          for I := 0 to High(FCellsRects) do
          begin
            if cbMergeCells.Checked then
              FWorksheet.Merge(FCellsRects[I], MergeWarningProc)
            else
              FWorksheet.UnMerge(FCellsRects[I])
          end;
        if vgrdcAngle in Changes then
          Angle := StrToIntDef(edAngle.Text, Angle);

        if vgrdcFontName in Changes then
          FontName := edFontName.Text;
        if vgrdcFontSize in Changes then
          FontSize := StrToIntDef(edFontSize.Text, 12);
        if vgrdcFontStyleBold in Changes then
          FontStyleBold := cbFontStyleBold.Checked;
        if vgrdcFontStyleItalic in Changes then
          FontStyleItalic := cbFontStyleItalic.Checked;
        if vgrdcFontStyleUnderline in Changes then
          FontStyleUnderline := cbFontStyleUnderline.Checked;
        if vgrdcFontStyleStrikeout in Changes then
          FontStyleStrikeout := cbFontStyleStrikeout.Checked;
        if vgrdcFontColor in Changes then
          FontColor := bFontColor.SelectedColor;
        if (vgrdcFontCharset in Changes) and (edCharset.ItemIndex >= 0) then
          FontCharset := Integer(edCharset.Items.Objects[edCharset.ItemIndex]);

        if vgrdcLeftBorder in Changes then
          Borders.Left.AssignBorderStyle(PreviewBordersFormat.Borders.Left);
        if vgrdcVertCenterBorder in Changes then
          Borders.Center.AssignBorderStyle(PreviewBordersFormat.Borders.Center);
        if vgrdcRightBorder in Changes then
          Borders.Right.AssignBorderStyle(PreviewBordersFormat.Borders.Right);
        if vgrdcTopBorder in Changes then
          Borders.Top.AssignBorderStyle(PreviewBordersFormat.Borders.Top);
        if vgrdcHorzCenterBorder in Changes then
          Borders.Middle.AssignBorderStyle(PreviewBordersFormat.Borders.Middle);
        if vgrdcBottomBorder in Changes then
          Borders.Bottom.AssignBorderStyle(PreviewBordersFormat.Borders.Bottom);

        if vgrdcFillBackColor in Changes then
          FillBackColor := bFillBackColor.SelectedColor;
        if vgrdcFillForeColor in Changes then
          FillForeColor := bFillForeColor.SelectedColor;
        if vgrdcFillPattern in Changes then
          FillPattern := TBrushStyle(edFillPattern.ItemIndex);

        if vgrdcDisplayFormat in Changes then
        begin
          DisplayFormat := edType.Text;
          // we must also change type of value in ranges
          if (lbFormats.ItemIndex >= 0) and (Integer(lbFormats.Items.Objects[lbFormats.ItemIndex]) <> vgrFormatCategoryCommon) then
            AWorksheet.ChangeValueTypeOfRanges(ACellsRects, AConvert[GetSelectedFormatCategory]);
        end;
      end;
    finally
      AWorksheet.Workbook.EndUpdate;
    end;
  end;
end;

procedure TvgrCellPropertiesForm.edHorzAlignClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcHorzAlign);
    UpdateAlignButtons;
  end;
end;

procedure TvgrCellPropertiesForm.edVertAlignClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcVertAlign);
    UpdateAlignButtons;
  end;
end;

procedure TvgrCellPropertiesForm.bLeftTopClick(Sender: TObject);
type
  rAlign = record
    Horz: TvgrRangeHorzAlign;
    Vert: TvgrRangeVertAlign;
  end;
const
  AAlignments: Array [0..8] of rAlign =
               ((Horz: vgrhaLeft; Vert: vgrvaTop),
                (Horz: vgrhaCenter; Vert: vgrvaTop),
                (Horz: vgrhaRight; Vert: vgrvaTop),
                (Horz: vgrhaLeft; Vert: vgrvaCenter),
                (Horz: vgrhaCenter; Vert: vgrvaCenter),
                (Horz: vgrhaRight; Vert: vgrvaCenter),
                (Horz: vgrhaLeft; Vert: vgrvaBottom),
                (Horz: vgrhaCenter; Vert: vgrvaBottom),
                (Horz: vgrhaRight; Vert: vgrvaBottom));
begin
  if EventsEnabled then
  begin
    if edHorzAlign.ItemIndex <> Integer(AAlignments[TSpeedButton(Sender).Tag].Horz) then
    begin
      edHorzAlign.ItemIndex := Integer(AAlignments[TSpeedButton(Sender).Tag].Horz);
      Include(FChanges, vgrdcHorzAlign);
    end;
    if edVertAlign.ItemIndex <> Integer(AAlignments[TSpeedButton(Sender).Tag].Vert) then
    begin
      edVertAlign.ItemIndex := Integer(AAlignments[TSpeedButton(Sender).Tag].Vert);
      Include(FChanges, vgrdcVertAlign);
    end;
    UpdateAlignButtons;
  end;
end;

procedure TvgrCellPropertiesForm.cbWrapTextClick(Sender: TObject);
begin
  if EventsEnabled then
    Include(FChanges, vgrdcWordWrap);
end;

procedure TvgrCellPropertiesForm.cbMergeCellsClick(Sender: TObject);
begin
  if EventsEnabled then
    Include(FChanges, vgrdcMerge);
end;

procedure TvgrCellPropertiesForm.edFontNameChange(Sender: TObject);
var
  I: Integer;
begin
  if EventsEnabled then
  begin
    I := 0;
    while (I < lbFontName.Items.Count) and
          (AnsiCompareText(Copy(lbFontName.Items[I], 1, Length(edFontName.Text)), edFontName.Text) <> 0) do Inc(I);
    if I < lbFontName.Items.Count then
    begin
      UpdateFontName(edFontName.Text);
      Include(FChanges, vgrdcFontName);
    end;
  end;
end;

procedure TvgrCellPropertiesForm.lbFontNameClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    UpdateFontName(lbFontName.Items[lbFontName.ItemIndex]);
    Include(FChanges, vgrdcFontName);
  end;
end;

procedure TvgrCellPropertiesForm.cbFontStyleBoldClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcFontStyleBold);
    UpdateFontStyle;
  end;
end;

procedure TvgrCellPropertiesForm.cbFontStyleItalicClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcFontStyleItalic);
    UpdateFontStyle;
  end;
end;

procedure TvgrCellPropertiesForm.cbFontStyleUnderlineClick(
  Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcFontStyleUnderline);
    UpdateFontStyle;
  end;
end;

procedure TvgrCellPropertiesForm.cbFontStyleStrikeoutClick(
  Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcFontStyleStrikeout);
    UpdateFontStyle;
  end;
end;

procedure TvgrCellPropertiesForm.edFontSizeChange(Sender: TObject);
begin
  if EventsEnabled then
  begin
    Include(FChanges, vgrdcFontSize);
    UpdateFontSize(StrToIntDef(edFontSize.Text, 8));
  end;
end;

procedure TvgrCellPropertiesForm.SetBordersToPreview(ACellsBorders: TvgrBorderTypes;
                                                     ABorderStyle: TvgrBorderStyle;
                                                     ABorderWidth: Integer;
                                                     ABorderColor: TColor);

  procedure SetBorder(ABorder: TvgrRangesFormatBorder);
  begin
    with ABorder do
    begin
      Width := ABorderWidth;
      if ABorderWidth > 0 then
      begin
        Pattern := ABorderStyle;
        Color := ABorderColor;
      end;
    end;
  end;

begin
  if EventsEnabled then
  begin
    if vgrbtLeft in ACellsBorders then
    begin
      SetBorder(PreviewBordersFormat.Borders.Left);
      Include(FChanges, vgrdcLeftBorder);
    end;
    if vgrbtCenter in ACellsBorders then
    begin
      SetBorder(PreviewBordersFormat.Borders.Center);
      Include(FChanges, vgrdcVertCenterBorder);
    end;
    if vgrbtRight in ACellsBorders then
    begin
      SetBorder(PreviewBordersFormat.Borders.Right);
      Include(FChanges, vgrdcRightBorder);
    end;

    if vgrbtTop in ACellsBorders then
    begin
      SetBorder(PreviewBordersFormat.Borders.Top);
      Include(FChanges, vgrdcTopBorder);
    end;
    if vgrbtMiddle in ACellsBorders then
    begin
      SetBorder(PreviewBordersFormat.Borders.Middle);
      Include(FChanges, vgrdcHorzCenterBorder);
    end;
    if vgrbtBottom in ACellsBorders then
    begin
      SetBorder(PreviewBordersFormat.Borders.Bottom);
      Include(FChanges, vgrdcBottomBorder);
    end;
    UpdateBorderButtons;
  end;
end;

procedure TvgrCellPropertiesForm.bNoneBordersClick(Sender: TObject);
begin
  SetBordersToPreview([vgrbtLeft, vgrbtCenter, vgrbtRight, vgrbtTop, vgrbtMiddle, vgrbtBottom],
                      vgrbsSolid,
                      0,
                      clNone);
end;

procedure TvgrCellPropertiesForm.bOutlineBordersClick(Sender: TObject);
begin
  SetBordersToPreview([vgrbtLeft, vgrbtRight, vgrbtTop, vgrbtBottom],
                      TvgrBorderStyle(edBorderStyle.ItemIndex),
                      ConvertPixelsToTwipsX(StrToIntDef(edBorderWidth.Text, 1)),
                      bBorderColor.SelectedColor);
end;

procedure TvgrCellPropertiesForm.bInsideBordersClick(Sender: TObject);
begin
  SetBordersToPreview([vgrbtCenter, vgrbtMiddle],
                      TvgrBorderStyle(edBorderStyle.ItemIndex),
                      ConvertPixelsToTwipsX(StrToIntDef(edBorderWidth.Text, 1)),
                      bBorderColor.SelectedColor);
end;

procedure TvgrCellPropertiesForm.bTopBorderClick(Sender: TObject);
begin
  with TSpeedButton(Sender) do
    if Down then
      SetBordersToPreview([TvgrBorderType(Tag)],
                          TvgrBorderStyle(edBorderStyle.ItemIndex),
                          ConvertPixelsToTwipsX(StrToIntDef(edBorderWidth.Text, 1)),
                          bBorderColor.SelectedColor)
    else
      SetBordersToPreview([TvgrBorderType(Tag)],
                          vgrbsSolid,
                          0,
                          clNone);
end;

procedure TvgrCellPropertiesForm.edFillPatternClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    UpdateFillPattern(TBrushStyle(edFillPattern.ItemIndex));
    Include(FChanges, vgrdcFillPattern);
  end;
end;

procedure TvgrCellPropertiesForm.bFontColorColorSelected(Sender: TObject;
  Color: TColor);
begin
  if EventsEnabled then
  begin
    UpdateFontColor(Color);
    Include(FChanges, vgrdcFontColor);
  end;
end;

procedure TvgrCellPropertiesForm.bFillBackColorColorSelected(
  Sender: TObject; Color: TColor);
begin
  if EventsEnabled then
  begin
    UpdateFillBackColor(Color);
    Include(FChanges, vgrdcFillBackColor);
  end;
end;

procedure TvgrCellPropertiesForm.bFillForeColorColorSelected(
  Sender: TObject; Color: TColor);
begin
  if EventsEnabled then
  begin
    UpdateFillForeColor(Color);
    Include(FChanges, vgrdcFillForeColor);
  end;
end;

procedure TvgrCellPropertiesForm.lbFontSizeClick(Sender: TObject);
begin
  if EventsEnabled then
  begin
    UpdateFontSize(StrToIntDef(lbFontSize.Items[lbFontSize.ItemIndex], 8));
    Include(FChanges, vgrdcFontSize);
  end;
end;

procedure TvgrCellPropertiesForm.edAngleChange(Sender: TObject);
begin
  if EventsEnabled then
    Include(FChanges, vgrdcAngle);
end;

procedure TvgrCellPropertiesForm.edCharSetClick(Sender: TObject);
begin
  if edCharSet.ItemIndex >= 0 then
  begin
    FontPreview.Font.Charset := Integer(edCharSet.Items.Objects[edCharSet.ItemIndex]);
    Include(FChanges, vgrdcFontCharset);
  end;
end;

procedure TvgrCellPropertiesForm.UpdateFormatControls(AVisible, AEnabled: Boolean);
begin
  GroupBox5.Visible := AVisible;
  GroupBox5.Enabled := AEnabled;
  vgrBevelLabel2.Visible := AVisible;
  vgrBevelLabel2.Enabled := AEnabled;
  edType.Visible := AVisible;
  edType.Enabled := AEnabled;
  lbType.Visible := AVisible;
  lbType.Enabled := AEnabled;
  bAddFormat.Visible := AVisible;
  bAddFormat.Enabled := AEnabled;
  bDeleteFormat.Visible := AVisible;
  bDeleteFormat.Enabled := AEnabled;
end;

function TvgrCellPropertiesForm.AddFormatType(AFormat: TvgrValueFormat): Integer;
begin
  Result := lbType.Items.AddObject(Format('%s (%s)', [AFormat.Format,
                                                      vgrFormatValue(AFormat.Format,
                                                                     GetDefaultFormatValue(Integer(AFormat.Category)))]),
                                   AFormat);
end;

procedure TvgrCellPropertiesForm.lbFormatsClick(Sender: TObject);
var
  I: Integer;
  ACategory: Integer;
begin
  if lbFormats.ItemIndex = -1 then
  begin
    UpdateFormatControls(True, False);

    edExample.Text := '';
    edType.Text := '';
    lbType.Clear;

    Label25.Caption := vgrLoadStr(svgrid_vgr_CellPropertiesDialog_SelectCategoryMessage);
  end
  else
  begin
    if Integer(lbFormats.Items.Objects[lbFormats.ItemIndex]) = vgrFormatCategoryCommon then
    begin
      UpdateFormatControls(False, False);

      edExample.Text := '';
      edType.Text := '';
      lbType.Clear;

      Label25.Caption := vgrLoadStr(svgrid_vgr_CellPropertiesDialog_NoFormatMessage);
    end
    else
    begin
      UpdateFormatControls(True, True);

      ACategory := Integer(lbFormats.Items.Objects[lbFormats.ItemIndex]);
      if ACategory <> vgrFormatCategoryAll then
      begin
        FChanges := FChanges + [vgrdcDisplayFormat];
        Label26.Visible := True;
      end;
      lbType.Items.Clear;
      for I := 0 to RegisteredValueFormats.Count - 1 do
        if (ACategory = vgrFormatCategoryAll) or (TvgrFormatCategory(ACategory) = RegisteredValueFormats[I].Category) then
          AddFormatType(RegisteredValueFormats[I]);
      if lbType.Items.Count > 0 then
      begin
        bAddFormat.Enabled := ACategory <> vgrFormatCategoryAll;
        lbType.ItemIndex := 0;
        lbTypeClick(nil);
      end
      else
      begin
        edType.Text := '';
        edExample.Text := '';
        bAddFormat.Enabled := False;
        bDeleteFormat.Enabled := False;
      end;

      Label25.Caption := '';
    end;
  end;
end;

procedure TvgrCellPropertiesForm.lbTypeClick(Sender: TObject);
begin
  if lbType.ItemIndex <> -1 then
    with TvgrValueFormat(lbType.Items.Objects[lbType.ItemIndex]) do
    begin
      edType.Text := Format;
      bAddFormat.Enabled := False;
      bDeleteFormat.Enabled := not IsBuiltIn;
    end;
end;

function TvgrCellPropertiesForm.GetSelectedFormatCategory: TvgrFormatCategory;
begin
  if (lbFormats.ItemIndex = -1) or
     (Integer(lbFormats.Items.Objects[lbFormats.ItemIndex]) = vgrFormatCategoryAll) then
  begin
    if lbType.ItemIndex <> -1 then
      Result := TvgrValueFormat(lbType.Items.Objects[lbType.ItemIndex]).Category
    else
      Result := vgrfcNumeric;
  end
  else
    Result := TvgrFormatCategory(lbFormats.Items.Objects[lbFormats.ItemIndex]);
end;

function TvgrCellPropertiesForm.GetDefaultFormatValue(ACategory: Integer = -1): Variant;
begin
  if ACategory = -1 then
    ACategory := Integer(GetSelectedFormatCategory);

  case TvgrFormatCategory(ACategory) of
    vgrfcNumeric: Result := 1234.5678;
    vgrfcDate: Result := Date;
    vgrfcTime: Result := Time;
    vgrfcDateTime: Result := Now;
    vgrfcText: Result := 'Text';
  end;
end;

procedure TvgrCellPropertiesForm.edTypeChange(Sender: TObject);
begin
  if EventsEnabled then
  begin
    FChanges := FChanges + [vgrdcDisplayFormat];
    Label26.Visible := True;
  end;

  BeginUpdate;
  try
    edExample.Text := vgrFormatValue(edType.Text, GetDefaultFormatValue);
    bAddFormat.Enabled := (edType.Text <> '') and (lbType.Items.IndexOf(edType.Text) = -1);
  finally
    EndUpdate;
  end;
end;

procedure TvgrCellPropertiesForm.bAddFormatClick(Sender: TObject);
var
  ACategory: Integer;
  AFormat: TvgrValueFormat;
begin
  if (lbFormats.ItemIndex >= 0) and (Trim(edType.Text) <> '') then
  begin
    ACategory := Integer(lbFormats.Items.Objects[lbFormats.ItemIndex]);
    if (ACategory <> vgrFormatCategoryAll) and (ACategory <> vgrFormatCategoryCommon) then
    begin
      AFormat := RegisterValueFormat(TvgrFormatCategory(ACategory), edType.Text, False);
      if AFormat <> nil then
      begin
        lbType.ItemIndex := AddFormatType(AFormat);
        bDeleteFormat.Enabled := True;
      end;

      bAddFormat.Enabled := False;
    end;
  end;
end;

procedure TvgrCellPropertiesForm.bDeleteFormatClick(Sender: TObject);
begin
  if (lbFormats.ItemIndex >= 0) and (lbType.Items.Objects[lbType.ItemIndex] <> nil) then
  begin
    RegisteredValueFormats.Remove(lbType.Items.Objects[lbType.ItemIndex]);
    lbType.Items.Delete(lbType.ItemIndex);
    if lbType.Items.Count > 0 then
    begin
      lbType.ItemIndex := 0;
      lbTypeClick(nil);
    end
    else
    begin
      edType.Text := '';
      edExample.Text := '';
      bAddFormat.Enabled := False;
      bDeleteFormat.Enabled := False;
    end;
  end;
end;

procedure TvgrCellPropertiesForm.DoLocalize;
begin
  inherited;
  edHorzAlign.Items.Clear;
  edHorzAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrhaAuto));
  edHorzAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrhaLeft));
  edHorzAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrhaCenter));
  edHorzAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrhaRight));

  edVertAlign.Items.Clear;
  edVertAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrvaTop));
  edVertAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrvaCenter));
  edVertAlign.Items.Add(vgrLoadStr(svgrid_vgr_CellPropertiesDialog_vgrvaBottom));

  edBorderStyle.Items.Clear;
  edBorderStyle.Items.Add(vgrLoadStr(svgrid_Common_vgrbsSolid));
  edBorderStyle.Items.Add(vgrLoadStr(svgrid_Common_vgrbsDash));
  edBorderStyle.Items.Add(vgrLoadStr(svgrid_Common_vgrbsDot));
  edBorderStyle.Items.Add(vgrLoadStr(svgrid_Common_vgrbsDashDot));
  edBorderStyle.Items.Add(vgrLoadStr(svgrid_Common_vgrbsDashDotDot));

  edFillPattern.Items.Clear;
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsSolid));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsClear));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsHorizontal));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsVertical));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsFDiagonal));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsBDiagonal));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsCross));
  edFillPattern.Items.Add(vgrLoadStr(svgrid_Common_bsDiagCross));
end;

initialization

  RegisteredValueFormats := TvgrValueFormats.Create;
  RegisterValueFormat(vgrfcNumeric, '0', True);
  RegisterValueFormat(vgrfcNumeric, '0.00', True);
  RegisterValueFormat(vgrfcNumeric, '#,##0', True);
  RegisterValueFormat(vgrfcNumeric, '#,##0.00', True);
  RegisterValueFormat(vgrfcNumeric, '0%', True);
  RegisterValueFormat(vgrfcNumeric, '0.00%', True);
  RegisterValueFormat(vgrfcNumeric, '0,00E+00', True);
  RegisterValueFormat(vgrfcNumeric, '##0,0E+0', True);

  RegisterValueFormat(vgrfcDate, 'dd/mm/yyyy', True);
  RegisterValueFormat(vgrfcDate, 'dd/mm/yy', True);
  RegisterValueFormat(vgrfcDate, 'dd/mmm/yy', True);
  RegisterValueFormat(vgrfcDate, 'dd/mmm', True);
  RegisterValueFormat(vgrfcDate, 'mmm/yy', True);

  RegisterValueFormat(vgrfcTime, 'h:nn', True);
  RegisterValueFormat(vgrfcTime, 'hh:nn', True);
  RegisterValueFormat(vgrfcTime, 'h:nn ampm', True);
  RegisterValueFormat(vgrfcTime, 'hh:nn:ss', True);
  RegisterValueFormat(vgrfcTime, 't', True);
  RegisterValueFormat(vgrfcTime, 'tt', True);

  RegisterValueFormat(vgrfcDateTime, 'dd/mm/yyyy h:nn', True);
  RegisterValueFormat(vgrfcDateTime, 'dd/mm/yy h:nn', True);
  RegisterValueFormat(vgrfcDateTime, 'dd/mm/yyyy h:nn:ss', True);

finalization

  FreeAndNil(RegisteredValueFormats);

end.

