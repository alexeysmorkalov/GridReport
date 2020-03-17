{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes for describing various page properties, margins, sizes etc.
  TvgrPageDefaults - specifies default width of column and default width of row.<br>
  TvgrPageMargins - specifies margins of page.<br>
  TvgrPageProperties - specifies sizes of page.<br>
See also:
  TvgrPageDefaults, TvgrPageMargins, TvgrPageProperties}
unit vgr_PageProperties;

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, {$IFDEF VTK_D6_OR_D7} Types, {$ELSE} Windows, {$ENDIF}
  SysUtils, Graphics, 

  vgr_Consts, vgr_Functions, vgr_CommonClasses, vgr_Localize;


const
{Specifies the key which displays the page number in the header or footer.}
  cHFKeyPage = '&[Page]';
{Specifies the key which displays the number of pages in the header or footer.}
  cHFKeyPages = '&[Pages]';
{Specifies the key which displays the current date in the header or footer.}
  cHFKeyDate = '&[Date]';
{Specifies the key which displays the current time in the header or footer.}
  cHFKeyTime = '&[Time]';
{Specifies the key which displays the caption of worksheet in the header or footer.}
  cHFKeyTab = '&[Tab]';

type

{Specifies page orientation.
Items:
  vgrpoPortrait - portrait orientation.
  vgrpoLandscape - landscape orientation.}
  TvgrPageOrientation = (vgrpoPortrait, vgrpoLandscape);

  /////////////////////////////////////////////////
  //
  // TvgrPageDefaults
  //
  ////////////////////////////////////////////////
{Specifies default column width and default row height for page.
If width of column on page was not hardcoded, then TvgrPageDefaults.ColWidth is used.
If height of row on page was not hardcoded, then TvgrPageDefaults.RowHeight is used.
ColWidth and RowHeight are specified in twips (1/1440 of inch),
also you can use UnitsColWidth and UnitsRowHeight properties to specify
these values in other measurement units.}
  TvgrPageDefaults = class(TvgrPersistent)
  private
    FData: array[0..1] of Integer;
    function GetValue(Index: Integer): Integer;
    procedure SetValue(Index: Integer; Value: Integer);

    function GetUnitsColWidth(Index: TvgrUnits): Extended;
    procedure SetUnitsColWidth(Index: TvgrUnits; Value: Extended);
    function GetUnitsRowHeight(Index: TvgrUnits): Extended;
    procedure SetUnitsRowHeight(Index: TvgrUnits; Value: Extended);

    function GetUnitsColWidthStr(Index: TvgrUnits): string;
    function GetUnitsRowHeightStr(Index: TvgrUnits): string;
  public
    constructor Create; override;
    procedure Assign(Source: TPersistent); override;

{Use this property to specify default column width (ColWidth property)
in various measurements units:
pixels, millimeters, centimeters, inches, etc.
Syntax:
  property UnitsColWidth[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units.
See alse:
  ColWidth}
    property UnitsColWidth[Index: TvgrUnits]: Extended read GetUnitsColWidth write SetUnitsColWidth;
{Use this property to specify default row height (RowHeight property)
in various measurement units:
pixels, millimeters, centimeters, inches, etc.
Syntax:
  property UnitsRowHeight[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units.
See also:
  RowHeight}
    property UnitsRowHeight[Index: TvgrUnits]: Extended read GetUnitsRowHeight write SetUnitsRowHeight;

{Returns as string value of the ColWidth property, string is formatted appropriate
to specified measurement units.
Syntax:
  property UnitsColWidthStr[Index: TvgrUnits]: string read
Parameters:
  Index - the measurements units.
See also:
  ColWidth, UnitsColWidth}
    property UnitsColWidthStr[Index: TvgrUnits]: string read GetUnitsColWidthStr;
{Returns as string value of the RowHeight property, string is formatted appropriate
to specified measurement units.
Syntax:
  property UnitsRowHeightStr[Index: TvgrUnits]: string read
Parameters:
  Index - the measurements units.
See also:
  RowHeight, UnitsRowHeight}
    property UnitsRowHeightStr[Index: TvgrUnits]: string read GetUnitsRowHeightStr;
  published
{Default column width on page (in twips, 1/1440 of inch), these property is used if
width of column on page was not hardcoded.
Syntax:
  property ColWidth: Integer read write
See also:
  UnitsColWidth, RowHeight, UnitsRowHeight}
    property ColWidth: Integer Index 0 read GetValue write SetValue default cDefaultColWidthTwips;
{Default row height on page (in twips, 1/1440 of inch), these property is used if
height of row on page was not hardcoded.
Syntax:
  property RowHeight: Integer read write
See also:
  UnitsRowHeight, ColWidth, UnitsColWidth}
    property RowHeight: Integer Index 1 read GetValue write SetValue default cDefaultRowHeightTwips;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPageMargins
  //
  ////////////////////////////////////////////////
{This class specifies margins of the page.
Main properties is: Left, Top, Right and Bottom,
all these properties are specified in twips (1/1440 of inch), also you can use
UnitsXXXX properties to specify margins in various measurement units.}
  TvgrPageMargins = class(TvgrPersistent)
  private
    FData: array[0..3] of Integer;
    function GetValue(Index: Integer): Integer;
    procedure SetValue(Index: Integer; Value: Integer);

    function GetUnitsLeft(Index: TvgrUnits): Extended;
    procedure SetUnitsLeft(Index: TvgrUnits; Value: Extended);
    function GetUnitsTop(Index: TvgrUnits): Extended;
    procedure SetUnitsTop(Index: TvgrUnits; Value: Extended);
    function GetUnitsRight(Index: TvgrUnits): Extended;
    procedure SetUnitsRight(Index: TvgrUnits; Value: Extended);
    function GetUnitsBottom(Index: TvgrUnits): Extended;
    procedure SetUnitsBottom(Index: TvgrUnits; Value: Extended);

    function GetUnitsLeftStr(Index: TvgrUnits): string;
    function GetUnitsTopStr(Index: TvgrUnits): string;
    function GetUnitsRightStr(Index: TvgrUnits): string;
    function GetUnitsBottomStr(Index: TvgrUnits): string;
  public
    constructor Create; override;
    procedure Assign(Source: TPersistent); override;

{Left margin of the page, can be specified in various measurement units.
Syntax:
  property UnitsLeft[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsLeft[Index: TvgrUnits]: Extended read GetUnitsLeft write SetUnitsLeft;
{Top margin of the page, can be specified in various measurement units.
Syntax:
  property UnitsTop[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsTop[Index: TvgrUnits]: Extended read GetUnitsTop write SetUnitsTop;
{Right margin of the page, can be specified in various measurement units.
Syntax:
  property UnitsRight[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsRight[Index: TvgrUnits]: Extended read GetUnitsRight write SetUnitsRight;
{Bottom margin of the page, can be specified in various measurement units.
Syntax:
  property UnitsBottom[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsBottom[Index: TvgrUnits]: Extended read GetUnitsBottom write SetUnitsBottom;

{Returns as string the value of the UnitsLeft property.
Syntax:
  property UnitsLeftStr[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsLeftStr[Index: TvgrUnits]: string read GetUnitsLeftStr;
{Returns as string the value of the UnitsTop property.
Syntax:
  property UnitsTopStr[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsTopStr[Index: TvgrUnits]: string read GetUnitsTopStr;
{Returns as string the value of the UnitsRight property.
Syntax:
  property UnitsRightStr[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsRightStr[Index: TvgrUnits]: string read GetUnitsRightStr;
{Returns as string the value of the UnitsBottom property.
Syntax:
  property UnitsBottomStr[Index: TvgrUnits]: Extended read write
Parameters:
  Index - the measurement units}
    property UnitsBottomStr[Index: TvgrUnits]: string read GetUnitsBottomStr;
  published
{Left margin of the page, is specified in twips (1440 of inch).
Syntax:
  property Left: Integer read write}
    property Left: Integer Index 0 read GetValue write SetValue default cTwipsPerCm;
{Top margin of the page, is specified in twips (1440 of inch).
Syntax:
  property Top: Integer read write}
    property Top: Integer Index 1 read GetValue write SetValue default cTwipsPerCm;
{Right margin of the page, is specified in twips (1440 of inch).
Syntax:
  property Right: Integer read write}
    property Right: Integer Index 2 read GetValue write SetValue default cTwipsPerCm;
{Bottom margin of the page, is specified in twips (1440 of inch).
Syntax:
  property Bottom: Integer read write}
    property Bottom: Integer Index 3 read GetValue write SetValue default cTwipsPerCm;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPageHeaderFooter
  //
  ////////////////////////////////////////////////
{Represents the page header or footer.
Each header or footer consists of three sections:<ul>
<li>Left - will be printed on left side.</li>
<li>Center - will be printed in center.</li>
<li>Right - will be printed on right side.</li><ul>}
  TvgrPageHeaderFooter = class(TvgrPersistent)
  private
    FLeftSection: string;
    FCenterSection: string;
    FRightSection: string;
    FHeight: Integer;
    FLeftFont: TFont;
    FCenterFont: TFont;
    FRightFont: TFont;
    FBackColor: TColor;

    function GetUnitsHeight(Index: TvgrUnits): Extended;
    procedure SetUnitsHeight(Index: TvgrUnits; Value: Extended);
    function GetUnitsHeightStr(Index: TvgrUnits): string;
    procedure SetLeftSection(Value: string);
    procedure SetCenterSection(Value: string);
    procedure SetRightSection(Value: string);
    procedure SetHeight(Value: Integer);
    procedure SetLeftFont(Value: TFont);
    procedure SetCenterFont(Value: TFont);
    procedure SetRightFont(Value: TFont);
    procedure OnChange(Sender: TObject);
    function IsFontStored(AFont: TFont): Boolean;
    function IsLeftFontStored: Boolean;
    function IsCenterFontStored: Boolean;
    function IsRightFontStored: Boolean;
    procedure SetBackColor(Value: TColor);
  public
{Creates an instance of the TvgrPageHeaderFooter class.}
    constructor Create; override;
{Frees an instance of the TvgrPageHeaderFooter class.}
    destructor Destroy; override;
{Copies the contents of another, similar object.
Parameters:
  Source - The source object.}
    procedure Assign(Source: TPersistent); override;

{Height of header or footer, can be specified in various measurement units.
Parameters:
  Index - The measurement units}
    property UnitsHeight[Index: TvgrUnits]: Extended read GetUnitsHeight write SetUnitsHeight;
{Returns as string the value of the UnitsHeight property.
Parameters:
  Index - The measurement units}
    property UnitsHeightStr[Index: TvgrUnits]: string read GetUnitsHeightStr;
  published
{Specifies the background color of header or footer.}
    property BackColor: TColor read FBackColor write SetBackColor default clNone;
{Specifies the font for left section of header or footer.}
    property LeftFont: TFont read FLeftFont write SetLeftFont stored IsLeftFontStored;
{Specifies the font for center section of header or footer.}
    property CenterFont: TFont read FCenterFont write SetCenterFont stored IsCenterFontStored;
{Specifies the font for right section of header or footer.}
    property RightFont: TFont read FRightFont write SetRightFont stored IsRightFontStored;
{Specifies the text to display or print the header in the top-left corner of
the worksheet or the footer in the bottom-left corner of the worksheet.
See also:
  CenterSection, RightSection}
    property LeftSection: string read FLeftSection write SetLeftSection;
{Specifies the text to display or print the header or footer centered
at the bottom of the worksheet.
See also:
  LeftSection, RightSection}
    property CenterSection: string read FCenterSection write SetCenterSection;
{Specifies the text to display or print the header in the top-right corner of
the worksheet or the footer in the bottom-right corner of the worksheet.
See also:
  CenterSection, LeftSection}
    property RightSection: string read FRightSection write SetRightSection;
{Specifies the height of header in twips (1/1440 inch).}
    property Height: Integer read FHeight write SetHeight default 0;
  end;

{Specifies the measurement system, used to specify sizes of the page in the
"Page setup" dialog window.
Items:
  vgrmsMetric - metric.
  vgrmsUSA - USA.
See also:
  TvgrPageProperties, TvgrPageSetupDialog}
  TvgrMeasurementSystem = (vgrmsMetric, vgrmsUSA);

  /////////////////////////////////////////////////
  //
  // TvgrPageProperties
  //
  ////////////////////////////////////////////////
{This class describes sizes of page in printing or previewing.
Instance of this class creates by TvgrWorksheet.
See also:
  TvgrPageMargins, TvgrPageDefaults, TvgrWorksheet}
  TvgrPageProperties = class(TvgrPersistent)
  private
    FHeight: Integer;
    FWidth: Integer;
    FMargins: TvgrPageMargins;
    FDefaults: TvgrPageDefaults;
    FHeader: TvgrPageHeaderFooter;
    FFooter: TvgrPageHeaderFooter;
    FMeasurementSystem: TvgrMeasurementSystem;
    function GetHeight: Integer;
    function GetWidth: Integer;
    function GetClientHeight: Integer;
    function GetClientWidth: Integer;
    function GetMargins: TvgrPageMargins;
    function GetDefaults: TvgrPageDefaults;
    function GetMeasurementSystem: TvgrMeasurementSystem;
    procedure SetHeight(Value: Integer);
    procedure SetWidth(Value: Integer);
    procedure SetMargins(Value: TvgrPageMargins);
    procedure SetDefaults(Value: TvgrPageDefaults);
    procedure SetMeasurementSystem(Value: TvgrMeasurementSystem);
    procedure SetDefaultMeasurementSystem;
    procedure SetHeader(Value: TvgrPageHeaderFooter);
    procedure SetFooter(Value: TvgrPageHeaderFooter);

    function GetTenthsMMWidth: Integer;
    procedure SetTenthsMMWidth(Value: Integer);
    function GetTenthsMMHeight: Integer;
    procedure SetTenthsMMHeight(Value: Integer);

    function GetUnitsWidth(Index: TvgrUnits): Extended;
    procedure SetUnitsWidth(Index: TvgrUnits; Value: Extended);
    function GetUnitsHeight(Index: TvgrUnits): Extended;
    procedure SetUnitsHeight(Index: TvgrUnits; Value: Extended);
    function GetUnitsWidthStr(Index: TvgrUnits): string;
    function GetUnitsHeightStr(Index: TvgrUnits): string;
  protected
    procedure OnChange(Sender: TObject);
  public
    constructor Create; override;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;

{Returns client height of the page, excluding the top and bottom margins, value
returns in twips (1/1440 of inch).
Syntax:
  property ClientHeight: Integer read;}
    property ClientHeight: Integer read GetClientHeight;
{Returns client width of the page, excluding the left and right margins, value
returns in twips (1/1440 of inch).
Syntax:
  property ClientWidth: Integer read;}
    property ClientWidth: Integer read GetClientWidth;

{Gets or sets width of the page in 1/10 of millimeter.
Syntax:
  property TenthsMMWidth: Integer read write;}
    property TenthsMMWidth: Integer read GetTenthsMMWidth write SetTenthsMMWidth;
{Gets or sets height of the page in 1/10 of millimeter.
Syntax:
  property TenthsMMHeight: Integer read write;}
    property TenthsMMHeight: Integer read GetTenthsMMHeight write SetTenthsMMHeight;
{Gets or sets width of the page, in the specified measurement units.
Syntax:
  property UnitsWidth[Index: TvgrUnits]: Extended read write;
Parameters:
  Index - the measurement units.}
    property UnitsWidth[Index: TvgrUnits]: Extended read GetUnitsWidth write SetUnitsWidth;
{Gets or sets height of the page, in the specified measurement units.
Syntax:
  property UnitsHeight[Index: TvgrUnits]: Extended read write;
Parameters:
  Index - the measurement units.}
    property UnitsHeight[Index: TvgrUnits]: Extended read GetUnitsHeight write SetUnitsHeight;
{Returns as string the value of the UnitsWidth property.
Syntax:
  property UnitsWidthStr[Index: TvgrUnits]: string read;
Parameters:
  Index - the measurement units.}
    property UnitsWidthStr[Index: TvgrUnits]: string read GetUnitsWidthStr;
{Returns as string the value of the UnitsHeight property.
Syntax:
  property UnitsHeightStr[Index: TvgrUnits]: string read;
Parameters:
  Index - the measurement units.}
    property UnitsHeightStr[Index: TvgrUnits]: string read GetUnitsHeightStr;
  published
{Specifies the page header.
See also:
  TvgrPageHeaderFooter}
    property Header: TvgrPageHeaderFooter read FHeader write SetHeader;
{Specifies the page footer.
See also:
  TvgrPageHeaderFooter}
    property Footer: TvgrPageHeaderFooter read FFooter write SetFooter;
{Gets or sets height of the page in twips (1/1440 of inch).
Syntax:
  property Height: Integer read write;}
    property Height: Integer read GetHeight write SetHeight;
{Gets or sets width of the page in twips (1/1440 of inch).
Syntax:
  property Width: Integer read write;}
    property Width: Integer read GetWidth write SetWidth;
{Specifies margins of the page.
Syntax:
  property Margins: TvgrPageMargins read write;
See also:
  TvgrPageMargins}
    property Margins: TvgrPageMargins read GetMargins write SetMargins;
{Specifies default width of the column and default height of the row on page.
Syntax:
  property Defaults: TvgrPageDefaults read write;
See also:
  TvgrPageDefaults}
    property Defaults: TvgrPageDefaults read GetDefaults write SetDefaults;
{Gets or sets measurement system, used to specify sizes of the page in the
"Page setup" dialog window.
Syntax:
  property MeasurementSystem: TvgrMeasurementSystem read write;
See also:
  TvgrMeasurementSystem, TvgrPageSetupDialog}
    property MeasurementSystem: TvgrMeasurementSystem read GetMeasurementSystem write SetMeasurementSystem;
  end;

  function ParseHeaderFooterText(const AText: string;
                                 APageNo: Integer;
                                 APageCount: Integer;
                                 const AWorksheetCaption: string): string;

var
{Is used to getting the measurement units for the measurement system.}
  ASystemToUnits: array [TvgrMeasurementSystem] of TvgrUnits = (vgruMMs, vgruInches);

implementation

function ParseHeaderFooterText(const AText: string;
                               APageNo: Integer;
                               APageCount: Integer;
                               const AWorksheetCaption: string): string;
begin
  Result := StringReplace(AText, cHFKeyPage, IntToStr(APageNo), [rfReplaceAll]);
  Result := StringReplace(Result, cHFKeyPages, IntToStr(APageCount), [rfReplaceAll]);
  Result := StringReplace(Result, cHFKeyTab, AWorksheetCaption, [rfReplaceAll]);
  Result := StringReplace(Result, cHFKeyDate, DateToStr(Date), [rfReplaceAll]);
  Result := StringReplace(Result, cHFKeyTime, TimeToStr(Time), [rfReplaceAll]);
end;

/////////////////////////////////////////////////
//
// TvgrPageDefaults
//
////////////////////////////////////////////////
constructor TvgrPageDefaults.Create;
begin
  inherited;
  FData[0] := cDefaultColWidthTwips;
  FData[1] := cDefaultRowHeightTwips;
end;

function TvgrPageDefaults.GetValue(Index: Integer): Integer;
begin
  Result := FData[Index];
end;

procedure TvgrPageDefaults.SetValue(Index: Integer; Value: Integer);
begin
  if FData[Index] <> Value then
  begin
    FData[Index] := Value;
    DoChange;
  end;
end;

function TvgrPageDefaults.GetUnitsColWidth(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Colwidth, Index);
end;

procedure TvgrPageDefaults.SetUnitsColWidth(Index: TvgrUnits; Value: Extended);
begin
  ColWidth := ConvertUnitsToTwips(Value, Index);
end;

function TvgrPageDefaults.GetUnitsRowHeight(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(RowHeight, Index);
end;

procedure TvgrPageDefaults.SetUnitsRowHeight(Index: TvgrUnits; Value: Extended);
begin
  RowHeight := ConvertUnitsToTwips(Value, Index);
end;

function TvgrPageDefaults.GetUnitsColWidthStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(ColWidth, Index);
end;

function TvgrPageDefaults.GetUnitsRowHeightStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(RowHeight, Index);
end;

procedure TvgrPageDefaults.Assign(Source: TPersistent);
begin
  if Source is TvgrPageDefaults then
  begin
    with TvgrPageDefaults(Source) do
      Self.FData := FData;
    DoChange;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrPageMargins
//
////////////////////////////////////////////////
constructor TvgrPageMargins.Create;
begin
  inherited;
  FData[0] := cTwipsPerCm;
  FData[1] := cTwipsPerCm;
  FData[2] := cTwipsPerCm;
  FData[3] := cTwipsPerCm;
end;

function TvgrPageMargins.GetValue(Index: Integer): Integer;
begin
  Result := FData[Index];
end;

procedure TvgrPageMargins.SetValue(Index: Integer; Value: Integer);
begin
  FData[Index] := Value;
  DoChange;
end;

function TvgrPageMargins.GetUnitsLeft(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Left, Index)
end;

procedure TvgrPageMargins.SetUnitsLeft(Index: TvgrUnits; Value: Extended);
begin
  Left := ConvertUnitsToTwips(Value, Index)
end;

function TvgrPageMargins.GetUnitsTop(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Top, Index)
end;

procedure TvgrPageMargins.SetUnitsTop(Index: TvgrUnits; Value: Extended);
begin
  Top := ConvertUnitsToTwips(Value, Index)
end;

function TvgrPageMargins.GetUnitsRight(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Right, Index)
end;

procedure TvgrPageMargins.SetUnitsRight(Index: TvgrUnits; Value: Extended);
begin
  Right := ConvertUnitsToTwips(Value, Index)
end;

function TvgrPageMargins.GetUnitsBottom(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Bottom, Index)
end;

procedure TvgrPageMargins.SetUnitsBottom(Index: TvgrUnits; Value: Extended);
begin
  Bottom := ConvertUnitsToTwips(Value, Index)
end;

function TvgrPageMargins.GetUnitsLeftStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(Left, Index)
end;

function TvgrPageMargins.GetUnitsTopStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(Top, Index)
end;

function TvgrPageMargins.GetUnitsRightStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(Right, Index)
end;

function TvgrPageMargins.GetUnitsBottomStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(bottom, Index)
end;

procedure TvgrPageMargins.Assign(Source: TPersistent);
begin
  if Source is TvgrPageMargins then
  begin
    with TvgrPageMargins(Source) do
      Self.FData := FData;
    DoChange;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrPageHeaderFooter
//
////////////////////////////////////////////////
constructor TvgrPageHeaderFooter.Create;

  procedure InitFont(AFont: TFont);
  begin
    with AFont do
    begin
      Name := 'Arial';
      Size := 10;
      Charset := DEFAULT_CHARSET;
      Style := [];
      OnChange := Self.OnChange;
    end;
  end;

begin
  inherited;
  FHeight := 0;
  FBackColor := clNone;
  FLeftFont := TFont.Create;
  FCenterFont := TFont.Create;
  FRightFont := TFont.Create;
  InitFont(FLeftFont);
  InitFont(FCenterFont);
  InitFont(FRightFont);
end;

destructor TvgrPageHeaderFooter.Destroy;
begin
  FLeftFont.Free;
  FCenterFont.Free;
  FRightFont.Free;
  inherited;
end;

procedure TvgrPageHeaderFooter.Assign(Source: TPersistent);
begin
  if Source is TvgrPageHeaderFooter then
  begin
    with TvgrPageHeaderFooter(Source) do
    begin
      BeginUpdate;
      try
        Self.FLeftSection := LeftSection;
        Self.FRightSection := RightSection;
        Self.FCenterSection := CenterSection;
        Self.FHeight := Height;
        Self.FBackColor := BackColor;
        Self.FLeftFont.Assign(LeftFont);
        Self.FCenterFont.Assign(CenterFont);
        Self.FRightFont.Assign(RightFont);
      finally
        EndUpdate;
      end;
    end;
    DoChange;
  end;
end;

function TvgrPageHeaderFooter.GetUnitsHeight(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Height, Index);
end;

procedure TvgrPageHeaderFooter.SetUnitsHeight(Index: TvgrUnits; Value: Extended);
begin
  Height := ConvertUnitsToTwips(Value, Index);
end;

function TvgrPageHeaderFooter.GetUnitsHeightStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(Height, Index);
end;

procedure TvgrPageHeaderFooter.SetLeftSection(Value: string);
begin
  if FLeftSection <> Value then
  begin
    FLeftSection := Value;
    DoChange;
  end;
end;

procedure TvgrPageHeaderFooter.SetCenterSection(Value: string);
begin
  if FCenterSection <> Value then
  begin
    FCenterSection := Value;
    DoChange;
  end;
end;

procedure TvgrPageHeaderFooter.SetRightSection(Value: string);
begin
  if FRightSection <> Value then
  begin
    FRightSection := Value;
    DoChange;
  end;
end;

procedure TvgrPageHeaderFooter.SetHeight(Value: Integer);
begin
  if FHeight <> Value then
  begin
    FHeight := Value;
    DoChange;
  end;
end;

procedure TvgrPageHeaderFooter.SetLeftFont(Value: TFont);
begin
  FLeftFont.Assign(Value);
end;

procedure TvgrPageHeaderFooter.SetCenterFont(Value: TFont);
begin
  FCenterFont.Assign(Value);
end;

procedure TvgrPageHeaderFooter.SetRightFont(Value: TFont);
begin
  FRightFont.Assign(Value);
end;

procedure TvgrPageHeaderFooter.OnChange(Sender: TObject);
begin
  DoChange;
end;

function TvgrPageHeaderFooter.IsFontStored(AFont: TFont): Boolean;
begin
  Result := (AFont.Name <> 'Arial') or
            (AFont.Size <> 10) or
            (AFont.Charset <> DEFAULT_CHARSET) or
            (AFont.Style <> []);
end;

function TvgrPageHeaderFooter.IsLeftFontStored: Boolean;
begin
  Result := IsFontStored(FLeftFont);
end;

function TvgrPageHeaderFooter.IsCenterFontStored: Boolean;
begin
  Result := IsFontStored(FCenterFont);
end;

function TvgrPageHeaderFooter.IsRightFontStored: Boolean;
begin
  Result := IsFontStored(FRightFont);
end;

procedure TvgrPageHeaderFooter.SetBackColor(Value: TColor);
begin
  if FBackColor <> Value then
  begin
    FBackColor := Value;
    DoChange;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrPageProperties
//
////////////////////////////////////////////////
constructor TvgrPageProperties.Create;
begin
  inherited Create;
  FWidth := ConvertUnitsToTwips(A4_PaperWidth, vgruTenthsMMs);
  FHeight := ConvertUnitsToTwips(A4_PaperHeight, vgruTenthsMMs);
  FMargins := TvgrPageMargins.Create(OnChange);
  FDefaults := TvgrPageDefaults.Create(OnChange);
  FHeader := TvgrPageHeaderFooter.Create(OnChange);
  FFooter := TvgrPageHeaderFooter.Create(OnChange);
  SetDefaultMeasurementSystem;
end;

destructor TvgrPageProperties.Destroy;
begin
  FMargins.Free;
  FDefaults.Free;
  FHeader.Free;
  FFooter.Free;
  inherited;
end;

procedure TvgrPageProperties.OnChange(Sender: TObject);
begin
  if ((Sender = FMargins) or (Sender = FDefaults) or
      (Sender = FHeader) or (Sender = FFooter)) and EnableUpdate then
    DoChange;
end;

function TvgrPageProperties.GetTenthsMMWidth: Integer;
begin
  Result := Round(ConvertTwipsToUnits(Width, vgruTenthsMMs));
end;

procedure TvgrPageProperties.SetTenthsMMWidth(Value: Integer);
begin
  Width := ConvertUnitsToTwips(Value, vgruTenthsMMs);
end;

function TvgrPageProperties.GetTenthsMMHeight: Integer;
begin
  Result := Round(ConvertTwipsToUnits(Height, vgruTenthsMMs));
end;

procedure TvgrPageProperties.SetTenthsMMHeight(Value: Integer);
begin
  Height := ConvertUnitsToTwips(Value, vgruTenthsMMs);
end;

function TvgrPageProperties.GetUnitsWidth(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Width, Index);
end;

procedure TvgrPageProperties.SetUnitsWidth(Index: TvgrUnits; Value: Extended);
begin
  Width := ConvertUnitsToTwips(Value, Index);
end;

function TvgrPageProperties.GetUnitsHeight(Index: TvgrUnits): Extended;
begin
  Result := ConvertTwipsToUnits(Height, Index);
end;

procedure TvgrPageProperties.SetUnitsHeight(Index: TvgrUnits; Value: Extended);
begin
  Height := ConvertUnitsToTwips(Value, Index);
end;

function TvgrPageProperties.GetUnitsWidthStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(Width, Index);
end;

function TvgrPageProperties.GetUnitsHeightStr(Index: TvgrUnits): string;
begin
  Result := ConvertTwipsToUnitsStr(Height, Index);
end;

function TvgrPageProperties.GetHeight: Integer;
begin
  Result := FHeight;
end;

function TvgrPageProperties.GetWidth: Integer;
begin
  Result := FWidth;
end;

function TvgrPageProperties.GetClientHeight: Integer;
begin
  Result := Height - Margins.Top - Margins.Bottom - Header.Height - Footer.Height;
end;

function TvgrPageProperties.GetClientWidth: Integer;
begin
  Result := Width - Margins.Left - Margins.Right;
end;

function TvgrPageProperties.GetMargins: TvgrPageMargins;
begin
  Result := FMargins;
end;

function TvgrPageProperties.GetDefaults: TvgrPageDefaults;
begin
  Result := FDefaults;
end;

function TvgrPageProperties.GetMeasurementSystem: TvgrMeasurementSystem;
begin
  Result := FMeasurementSystem;
end;

procedure TvgrPageProperties.SetHeight(Value: Integer);
begin
  FHeight := Value;
  DoChange;
end;

procedure TvgrPageProperties.SetWidth(Value: Integer);
begin
  FWidth := Value;
  DoChange;
end;

procedure TvgrPageProperties.SetMargins(Value: TvgrPageMargins);
begin
  FMargins.Assign(Value);
  DoChange;
end;

procedure TvgrPageProperties.SetDefaults(Value: TvgrPageDefaults);
begin
  FDefaults.Assign(Value);
  DoChange;
end;

procedure TvgrPageProperties.SetMeasurementSystem(Value: TvgrMeasurementSystem);
begin
  FMeasurementSystem := Value;
  DoChange;
end;

procedure TvgrPageProperties.SetDefaultMeasurementSystem;
var
 AMeasure: array [0..0] of Char;
begin
  GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, AMeasure, SizeOf(AMeasure));
  if AMeasure[0] = '0' then
    FMeasurementSystem := vgrmsMetric
  else
    FMeasurementSystem := vgrmsUSA
end;

procedure TvgrPageProperties.SetHeader(Value: TvgrPageHeaderFooter);
begin
  FHeader.Assign(Value);
end;

procedure TvgrPageProperties.SetFooter(Value: TvgrPageHeaderFooter);
begin
  FFooter.Assign(Value);
end;

procedure TvgrPageProperties.Assign(Source: TPersistent);
begin
  if Source is TvgrPageProperties then
  begin
    with TvgrPageProperties(Source) do
    begin
      Self.BeginUpdate;
      try
        Self.FHeight := Height;
        Self.FWidth := Width;
        Self.FMeasurementSystem := MeasurementSystem;
        Self.FMargins.Assign(Margins);
        Self.FDefaults.Assign(Defaults);
        Self.FHeader.Assign(Header);
        Self.FFooter.Assign(Footer);
      finally
        Self.EndUpdate;
      end;
    end;
    DoChange;
  end;
end;

initialization


end.
