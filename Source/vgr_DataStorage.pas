{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{This module contains all base classes, which implement the storage of data.
TvgrWorkbook is a central component to work with library,
it contains Worksheets, like Excel's workbook contains woksheets.
Worksheets contains Ranges, Borders, Rows and Cols.
Range is a rectangle of one or more cells.
See also:
  TvgrWorkbook, TvgrWorksheet, TvgrSection, TvgrRow, TvgrCol, TvgrRange}
unit vgr_DataStorage;

interface

{$I vtk.inc}

//{$DEFINE DS_FIND_DEBUG}
{$DEFINE VTK_SECTIONS_DBG}
{$DEFINE VTK_REALLOC_IF_DELETE}
//{$DEFINE VTK_DSSAVE_DBG}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  Classes, Graphics, SysUtils, ActiveX, Windows,
  {$IFDEF VTK_D6_OR_D7} Types, Variants, {$ENDIF} Math, clipbrd,

  vgr_ExcelFormula, vgr_ReportGUIFunctions, vgr_ReportFunctions,
  vgr_CommonClasses, vgr_DataStorageRecords, vgr_Functions, vgr_Consts,
  vgr_DataStorageTypes, vgr_FormulaCalculator, vgr_PageProperties,
  vgr_ScriptDispIDs;

const
{Default font name for painting ranges
Syntax:
  vgrDefaultRangeFontName = 'Arial';}
  vgrDefaultRangeFontName = 'Arial';
{Default font size for painting ranges
Syntax:
  vgrDefaultRangeFontSize = 10;}
  vgrDefaultRangeFontSize = 10;

(*GUID for IvgrSectionExt interface
Syntax:
  svgrSectionExtGUID: TGUID = '{540E34BA-97AD-400D-946E-9C01D3FF78B1}';*)
  svgrSectionExtGUID: TGUID = '{540E34BA-97AD-400D-946E-9C01D3FF78B1}';
(*GUID for IvgrRow interface
Syntax:
  svgrRowGUID: TGUID = '672CB801-EB6D-48A8-BC73-3ECE5D30888C';*)
  svgrRowGUID: TGUID = '{672CB801-EB6D-48A8-BC73-3ECE5D30888C}';
(*GUID for IvgrCol interface
Syntax:
  svgrColGUID: TGUID = '9DB9E237-4113-4FC9-A59F-E83454A4D099';*)
  svgrColGUID: TGUID = '{9DB9E237-4113-4FC9-A59F-E83454A4D099}';
(*GUID for IvgrPageVector interface
Syntax:
  IID_IvgrPageVector: TGUID = '{AF4FAF66-C5BA-45F4-962D-9E8605F1EAEB}';*)
  IID_IvgrPageVector: TGUID = '{AF4FAF66-C5BA-45F4-962D-9E8605F1EAEB}';

{DataStorage version.
For internal use.
Syntax:
  vgrDataStorageVersion = 1;}
  vgrDataStorageVersion = 1;
{GridReport version.}
  vgrGridReportVersion = '1.2.3';

type

  TvgrRange = class;
  TvgrBorder = class;
  TvgrWBList = class;
  TvgrWBListItem = class;
  TvgrWorksheet = class;
  TvgrWorkbook = class;
  TvgrWorksheets = class;
  TvgrFormulasList = class;

  /////////////////////////////////////////////////
  //
  // INTERFACES
  //
  /////////////////////////////////////////////////
  /////////////////////////////////////////////////
  //
  // IvgrWBListItem
  //
  /////////////////////////////////////////////////
{Common interface of differ Workbook lists items
Used access to common properties for Workbook elements
- Rows, Cols, Ranges etc.}
  IvgrWBListItem = interface
  ['{6F7DF4CD-421B-4AFD-951D-5B6E918D897D}']
{Returns the value of the Worksheet property.
See also:
  Worksheet, TvgrWorksheet}
    function GetWorksheet: TvgrWorksheet;
{Returns the value of the Workbook property.
See also:
  Workbook, TvgrWorkbook}
    function GetWorkbook: TvgrWorkbook;
{Returns the value of the ItemData property.
See also:
  ItemData}
    function GetItemData: Pointer;
{Returns the the value of the ItemIndex property.
See also:
  ItemIndex}
    function GetItemIndex: Integer;
{Returns the value of the StyleData property.}
    function GetStyleData: Pointer;
{Returns the value of the Valid property.}
    function GetValid: Boolean;
{Sets the Valid property to False.
It is only for internal use, do not use this property.}
    procedure SetInvalid;

{Returns the TvgrWorksheet object which contains this item.}
    property Worksheet: TvgrWorksheet read GetWorksheet;
{Returns the TvgrWorkbook object which contains this item.}
    property Workbook: TvgrWorkbook read GetWorkbook;
{Returns the pointer to the item's data.
It is only for internal use, do not use this property.}
    property ItemData: Pointer read GetItemData;
{Returns the item's index in the list.
It is only for internal use, do not use this property.}
    property ItemIndex: Integer read GetItemIndex;
{Returns the pointer to the item's style data.
It is only for internal use, do not use this property.}
    property StyleData: Pointer read GetStyleData;
{Returns the true value if this interface is valid and can be used.}
    property Valid: Boolean read GetValid;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrWorkbookHandler
  //
  /////////////////////////////////////////////////
{This interface can be used for receiving events when Workbook is changed.
The object which realizes this interface must be registered with using of
the TvgrWorkbook.ConnectHandler method. Use the TvgrWorkbook.DisconnectHandler
method to unregister.}
  IvgrWorkbookHandler = interface
{This method is called before workbook is changed.
Parameters:
  ChangeInfo - Structure with information about changes.
See also:
  TvgrWorkbookChangeInfo}
    procedure BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
{This method is called after workbook is changed.
Parameters:
  ChangeInfo - Structure with information about changes.
See also:
  TvgrWorkbookChangeInfo}
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
{This method is called when AItem object is deleted from workbook.
Parameters:
  AItem - Interface to item which is deleted.
See also:
  IvgrWBListItem}
    procedure DeleteItemInterface(AItem: IvgrWBListItem);
  end;

  /////////////////////////////////////////////////
  //
  // IvgrVector
  //
  /////////////////////////////////////////////////
{Base interface for worksheet's rows and columns.
See also:
  IvgrWBListItem, IvgrPageVector}
  IvgrVector = interface(IvgrWBListItem)
  ['{1297B8B6-03B1-4B66-A9A1-85AE488BB6A0}']
{Returns the value of the Number property.}
    function GetNumber: Integer;
{Returns the value of the Size property.}
    function GetSize: Integer;
{Sets the value of the Size property.}
    procedure SetSize(Value: Integer);
{Returns the value of the Visible property.}
    function GetVisible: Boolean;
{Sets the value of the Visible property.}
    procedure SetVisible(Value: Boolean);
{Copies the properties of another vector.
Parameters:
  ASource - The source vector to assign to the this vector.}
    procedure Assign(ASource: IvgrVector);
{Specifies the vector's index in the worksheet's list (columns or rows).}
    property Number: Integer read GetNumber;
{Specifies the vector's size in twips.}
    property Size: Integer read GetSize write SetSize;
{Specifies the visibility of the vector.}
    property Visible: Boolean read GetVisible write SetVisible;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrPageVector
  //
  /////////////////////////////////////////////////
{Interface extents IvgrVector with page breaks
See also:
  IvgrVector, IvgrCol, IvgrRow}
  IvgrPageVector = interface(IvgrVector)
  ['{AF4FAF66-C5BA-45F4-962D-9E8605F1EAEB}']
{Returns the value of the PageBreak property.}
    function GetPageBreak: Boolean;
{Sets the value of the PageBreak property.}
    procedure SetPageBreak(Value: Boolean);
{Specifies the value indicating need to start new page after this vector,
when workbook is printed or previewed.}
    property PageBreak: Boolean read GetPageBreak write SetPageBreak;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrCol
  //
  /////////////////////////////////////////////////
{Interface for the worksheet's column.
See also:
  IvgrPageVector, IvgrRow, TvgrWorksheet.Rows}
  IvgrCol = interface(IvgrPageVector)
  ['{9DB9E237-4113-4FC9-A59F-E83454A4D099}']
{The column's width, in twips.}
    property Width: Integer read GetSize write SetSize;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrRow
  //
  /////////////////////////////////////////////////
{Interface for the worksheet's row. 
See also:
  IvgrPageVector, IvgrCol, TvgrWorksheet.Cols}
  IvgrRow = interface(IvgrPageVector)
  ['{672CB801-EB6D-48A8-BC73-3ECE5D30888C}']
{The row's height, in twips}
    property Height: Integer read GetSize write SetSize;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrRangeFont
  //
  /////////////////////////////////////////////////
{Encapsulates a TFont object properties
See also:
  IvgrRange}
  IvgrRangeFont = interface
  ['{757A227E-A966-438A-ABEB-3AAD49A6AD54}']
{Returns the value of the FontSize property.
See also:
  FontSize}
    function GetFontSize: Integer;
{Sets the value of the FontSize property.
See also:
  FontSize}
    procedure SetFontSize(Value: Integer);
{Returns the value of the FontPitch property.
See also:
  FontPitch}
    function GetFontPitch: TFontPitch;
{Sets the value of the FontPitch property.
See also:
  FontPitch}
    procedure SetFontPitch(Value: TFontPitch);
{Returns the value of the FontStyle property.
See also:
  FontStyle}
    function GetFontStyle: TFontStyles;
{Sets the value of the FontStyle property.
See also:
  FontStyle}
    procedure SetFontStyle(Value: TFontStyles);
{Returns the value of the FontCharset property.
See also:
  FontCharset}
    function GetFontCharset: TFontCharset;
{Sets the value of the FontCharset property.
See also:
  FontCharset}
    procedure SetFontCharset(Value: TFontCharset);
{Returns the value of the FontName property.
See also:
  FontName}
    function GetFontName: string;
{Sets the value of the FontName property.
See also:
  FontName}
    procedure SetFontName(const Value: string);
{Returns the value of the FontColor property.
See also:
  FontColor}
    function GetFontColor: TColor;
{Sets the value of the FontColor property.
See also:
  FontColor}
    procedure SetFontColor(Value: TColor);
{Copies the properties from TFont object.
Parameters:
  AFont - TFont object, which properties is copied.}
    procedure Assign(AFont: TFont);
{Copies the properties from IvgrRangeFont interface.
Parameters:
  AFont - IvgrRangeFont interface, which properties is copied.}
    procedure AssignRangeFont(AFont: IvgrRangeFont);
{Copies the properties of this object to a destination TFont object.
Parameters:
  AFont - TFont object to fill with properties of current object,
supported IvgrRangeFont interface}
    procedure AssignTo(AFont: TFont);

{Specifies the height of the font in points.}
    property Size: Integer read GetFontSize write SetFontSize;
{Specifies whether the characters in the font all have the same width.}
    property Pitch: TFontPitch read GetFontPitch write SetFontPitch;
{Determines whether the font is normal, italic, underlined, bold, and so on.}
    property Style: TFontStyles read GetFontStyle write SetFontStyle;
{Specifies the character set of the font.}
    property Charset: TFontCharset read GetFontCharset write SetFontCharset;
{Identifies the typeface of the font.}
    property Name: string read GetFontName write SetFontName;
{Specifies the color of the text.}
    property Color: TColor read GetFontColor write SetFontColor;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrRange
  //
  /////////////////////////////////////////////////
{Interface for the single cell or the range of cells.
See also:
  IvgrWBListItem}
  IvgrRange = interface(IvgrWBListItem)
  ['{06F7635A-96C5-4B74-B865-89CB2D788582}']
{Returns the value of the Left property.
See also:
  Left}
    function GetLeft: Integer;
{Returns the value of the Top property.
See also:
  Top}
    function GetTop: Integer;
{Returns the value of the Right property.
See also:
  Right}
    function GetRight: Integer;
{Returns the value of the Bottom property.
See also:
  Bottom}
    function GetBottom: Integer;
{Returns the value of the Place property.
See also:
  Place}
    function GetPlace: TRect;
{Returns the value of the Value property.
See also:
  Value}
    function GetValue: Variant;
{Sets the value of the Value property.
See also:
  Value}
    procedure SetValue(const Value: Variant);
{Returns the value of the DisplayText property.
See also:
  DisplayText}
    function GetDisplayText: string;
{Returns the value of the Font property.
See also:
  Font}
    function GetFont: IvgrRangeFont;
{Sets the value of the Font property.
See also:
  Font}
    procedure SetFont(Value: IvgrRangeFont);
{Returns the value of the FillBackColor property.
See also:
  FillBackColor}
    function GetFillBackColor: TColor;
{Sets the value of the FillBackColor property.
See also:
  FillBackColor}
    procedure SetFillBackColor(Value: TColor);
{Returns the value of the FillForeColor property.
See also:
  FillForeColor}
    function GetFillForeColor: TColor;
{Sets the value of the FillForeColor property.
See also:
  FillForeColor}
    procedure SetFillForeColor(Value: TColor);
{Returns the value of the FillPattern property.
See also:
  FillPattern}
    function GetFillPattern: TBrushStyle;
{Sets the value of the FillPattern property.
See also:
  FillPattern}
    procedure SetFillPattern(Value: TBrushStyle);
{Returns the value of the DisplayFormat property.
See also:
  DisplayFormat}
    function GetDisplayFormat: string;
{Sets the value of the DisplayFormat property.
See also:
  DisplayFormat}
    procedure SetDisplayFormat(const Value: string);
{Returns the value of the HorzAlign property.
See also:
  HorzAlign}
    function GetHorzAlign: TvgrRangeHorzAlign;
{Sets the value of the HorzAlign property.
See also:
  HorzAlign}
    procedure SetHorzAlign(Value: TvgrRangeHorzAlign);
{Returns the value of the VertAlign property.
See also:
  VertAlign}
    function GetVertAlign: TvgrRangeVertAlign;
{Sets the value of the VertAlign property.
See also:
  VertAlign}
    procedure SetVertAlign(Value: TvgrRangeVertAlign);
{Returns the value of the Angle property.
See also:
  Angle}
    function GetAngle: Word;
{Sets the value of the Angle property.
See also:
  Angle}
    procedure SetAngle(Value: Word);
{Returns the value of the Flags property.
See also:
  Flags}
    function GetFlags: Word;
{Sets the value of the Flags property.
See also:
  Flags}
    procedure SetFlags(Value: Word);
{Returns the value of the WordWrap property.
See also:
  WordWrap}
    function GetWordWrap: Boolean;
{Sets the value of the WordWrap property.
See also:
  WordWrap}
    procedure SetWordWrap(Value: Boolean);
{Returns the value of the ValueType property.
See also:
  ValueType}
    function GetValueType: TvgrRangeValueType;
{Returns the value of the StringValue property.
See also:
  StringValue}
    function GetStringValue: string;
{Sets the value of the StringValue property.
See also:
  StringValue}
    procedure SetStringValue(const Value: string);
{Returns the value of the SimpleStringValue property.
See also:
  SimpleStringValue}
    function GetSimpleStringValue: string;
{Sets the value of the SimpleStringValue property.
See also:
  SimpleStringValue}
    procedure SetSimpleStringValue(const Value: string);
{Returns the value of the Formula property.
See also:
  Formula}
    function GetFormula: string;
{Sets the value of the Formula property.
See also:
  Formula}
    procedure SetFormula(const Value: string);
{Returns the workbook's strings list object.
See also:}
    function GetWBStrings: TvgrWBStrings;
{Copies the properties from another range.
Parameters:
  ASource - The source IvgrRange interface.}
    procedure Assign(ASource: IvgrRange);
{Copies the visual properties of another range. (it is all properties excluding the Value property.)
ASource:
  ASource - The source IvgrRange interface.}
    procedure AssignStyle(ASource: IvgrRange);
{Returns value of the ValueData property.
See also:
  ValueData}
    function GetValueData: pvgrRangeValue;
{Forces range to default style.
It is only for internal use, do not use this property.}
    procedure SetDefaultStyle;
{Called when workbook is exported.
It is only for internal use, do not use this property.}
    function GetExportData: Pointer;
{Called when workbook is exported.
It is only for internal use, do not use this property.}
    procedure SetExportData(Value: Pointer);
{ This procedure are changing type of range value to specified value,
if conversion can not to be made - value not changing. }
    procedure ChangeValueType(ANewValueType: TvgrRangeValueType);

{Returns the X coordinate of the range's top-left corner.}
    property Left: Integer read GetLeft;
{Returns the Y coordinate of the range's top-left corner.}
    property Top: Integer read GetTop;
{Returns the X coordinate of the range's bottom-right corner.}
    property Right: Integer read GetRight;
{Returns the Y coordinate of the range's bottom-right corner.}
    property Bottom: Integer read GetBottom;
{Returns the placement of the range.}
    property Place: TRect read GetPlace;
{Specifies the range value.}
    property Value: Variant read GetValue write SetValue;
{Returns the range’s value as it is displayed in a grid control.
This property returns the formatted Value property, the DisplayFormat
property is used as format.}
    property DisplayText: string read GetDisplayText;
{Specifies the range's font.}
    property Font: IvgrRangeFont read GetFont write SetFont;
{Specifies the background color of the range's fill pattern.}
    property FillBackColor: TColor read GetFillBackColor write SetFillBackColor;
{Specifies the foreground color of the range's fill pattern.}
    property FillForeColor: TColor read GetFillForeColor write SetFillForeColor;
{Specifies the range's fill pattern.}
    property FillPattern: TBrushStyle read GetFillPattern write SetFillPattern;
{Determines how a range’s value is formatted.
Format of the format string is determined by the type of the range value.
For information on the format strings, see "Format Strings" in Delphi help.}
    property DisplayFormat: string read GetDisplayFormat write SetDisplayFormat;
{Specifies the horizontal text alignment.
See also:
  TvgrRangeHorzAlign}
    property HorzAlign: TvgrRangeHorzAlign read GetHorzAlign write SetHorzAlign;
{Specifies the vertical text alignment.
See also:
  TvgrRangeVertAlign}
    property VertAlign: TvgrRangeVertAlign read GetVertAlign write SetVertAlign;
{Specifies the rotation angle of the range's text in degrees.}
    property Angle: Word read GetAngle write SetAngle;
{Specifies additional information about Range object.}
    property Flags: Word read GetFlags write SetFlags;
{Specifies whether the range text wraps when it is too long for the width of the range.}
    property WordWrap: Boolean read GetWordWrap write SetWordWrap;
{Returns the data type of the range object.
See also:
  TvgrRangeValueType}
    property ValueType: TvgrRangeValueType read GetValueType;
{Represents the value of the range as a string value.
This property usually used for getting the range's value for editing.
  - If range contains the formula this property returns a formula string starting with '='.
    '=A1 + A2' for example.
  - If range contains the value of non-string type this property convert it
    to the string as VarToStr(Value).
When string is being assigned to this property:
  - If the string strats with '=' it is being converted to the formula.
  - If the string can be converted to the integer then the integer value is stored in the range.
  - If the string can be converted to the real number then the extended value is stored in the range.
  - If the string can be converted to the datetime then the datatime value is stored in the range.
Example:
  begin
    ...
    with worksheet.Ranges[0, 0, 0, 0] do
      StringValue := '=A2 + A3'; // stores a formula in the range

    with worksheet.Ranges[0, 0, 0, 0] do
      Formula := 'A2 + A3'; // stores a formula in the range (other variant)

    with worksheet.Ranges[0, 1, 0, 1] do
      StringValue := '123'; // stores the integer value in the range

    with worksheet.Ranges[0, 2, 0, 2] do
      Value := 123; // stores the integer value in the range (other variant)
    ...
  end;
See also:
  Value, SimpleStringValue}
    property StringValue: string read GetStringValue write SetStringValue;
{Represents the value of the range as a string value.
If the range contains a formula then the formula is calculated and its value is returned.
Example:
  begin
    ...
    with worksheet.Ranges[0, 0, 0, 0] do
      SimpleStringValue := '=A2 + A3'; // string will be stored, not a formula !!!

    with worksheet.Ranges[0, 1, 0, 1] do
      SimpleStringValue := '123'; // stores the integer value in the range

    with worksheet.Ranges[0, 2, 0, 2] do
      Value := 123; // stores the integer value in the range (other variant)
    ...
  end;
See also:
  Value, StringValue}
    property SimpleStringValue: string read GetSimpleStringValue write SetSimpleStringValue;
{Returns the pointer to the data structure.
It is only for internal use, do not use this property.}
    property ValueData: pvgrRangeValue read GetValueData;
{Specifies the range formula, if range has no formula returns the empty string.
Example:
  begin
    with worksheet.Ranges[10, 10, 10, 10] do
      Formula := 'A4 + A6';
  end;

See also:
  Value, StringValue, SimpleStringValue}
    property Formula: string read GetFormula write SetFormula;
{Returns the pointer to the temporary data for exporting.
It is only for internal use, do not use this property.}
    property ExportData: Pointer read GetExportData write SetExportData;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrBorder
  //
  /////////////////////////////////////////////////
{Interface for the border of a range.
See also:
  IvgrWBListItem}
  IvgrBorder = interface(IvgrWBListItem)
  ['{19ABDE2C-6926-42AC-B91D-CA9046CF6A13}']
{Returns the value of the Left property.
See also:
  Left}
    function GetLeft: Integer;
{Returns the value of the Top property.
See also:
  Top}
    function GetTop: Integer;
{Returns the value of the Orientation property.
See also:
  Orientation}
    function GetOrientation: TvgrBorderOrientation;
{Returns the value of the Width property.
See also:
  Width}
    function GetWidth: Integer;
{Sets the value of the Width property.
See also:
  Width}
    procedure SetWidth(Value: Integer);
{Returns the value of the Color property.
See also:
  Color}
    function GetColor: TColor;
{Sets the value of the Color property.
See also:
  Color}
    procedure SetColor(Value: TColor);
{Returns the value of the Pattern property.
See also:
  Pattern}
    function GetPattern: TvgrBorderStyle;
{Sets the value of the Pattern property.
See also:
  Pattern}
    procedure SetPattern(Value: TvgrBorderStyle);
{Sets the all visual border's properties.
Parameters:
  AWidth - The border's width in twips.
  AColor - The border's color.
  APattern - The border's style.
See also:
  Width, Color, Pattern, TvgrBorderStyle}
    procedure SetStyles(AWidth: Integer; AColor: TColor; APattern: TvgrBorderStyle);
{Copies the properties of another border.
Parameters:
  ASource - The source border.}
    procedure Assign(ASource: IvgrBorder);

{Returns the X coordinate for the top-left point of the border.}
    property Left: Integer read GetLeft;
{Returns the Y coordinate for the top-left point of the border.}
    property Top: Integer read GetTop;
{Specifies the border's orientation - horizontal or vertical.
See also:
  TvgrBorderOrientation}
    property Orientation: TvgrBorderOrientation read GetOrientation;
{Specifies the border's width in twips.}
    property Width: Integer read GetWidth write SetWidth;
{Specifies the border's color.}
    property Color: TColor read GetColor write SetColor;
{Specifies the border's pattern.
See also:
  TvgrBorderStyle}
    property Pattern: TvgrBorderStyle read GetPattern write SetPattern;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrSection
  //
  /////////////////////////////////////////////////
{Interface for the section of a worksheet.
Section can be vertical or horizontal.
The vertical section includes one or more columns, the horizontal -
one or more rows.
The section is specified by two coordinates StartPos and EndPos,
StartPos - defines number of a line / column with which section begins and
EndPos - defines number of a line / column on which section ends.
Sections can form multilevel hierarchical structure.
So for example the section with coordinates [1, 4] is a parent for section
with coordinates [2, 3].
Section can be used in TvgrWorkbookGrid to quickly hides / shows some
rows or columns.
See also:
  IvgrWBListItem, IvgrRow, IvgrCol}
  IvgrSection = interface(IvgrWBListItem)
  ['{08B03EC4-C359-44A1-8DF1-7E8A0BF52A1B}']
{Returns the value of the StartPos property.
See also:
  StartPos}
    function GetStartPos: Integer;
{Returns the value of the EndPos property.
See also:
  EndPos}
    function GetEndPos: Integer;
{Returns the value of the Level property.
See also:
  Level}
    function GetLevel: Integer;
{Returns the value of the Parent property.
See also:
  Parent}
    function GetParent: IvgrSection;
{Returns the value of the RepeatOnPageTop property.
See also:
  RepeatOnPageTop}
    function GetRepeatOnPageTop: Boolean;
{Sets the value of the RepeatOnPageTop property.
See also:
  RepeatOnPageTop}
    procedure SetRepeatOnPageTop(Value: Boolean);
{Returns the value of the RepeatOnPageBottom property.
See also:
  RepeatOnPageTop}
    function GetRepeatOnPageBottom: Boolean;
{Sets the value of the RepeatOnPageBottom property.
See also:
  RepeatOnPageTop}
    procedure SetRepeatOnPageBottom(Value: Boolean);
{Returns the value of the PrintWithNextSection property.
See also:
  PrintWithNextSection}
    function GetPrintWithNextSection: Boolean;
{Sets the value of the PrintWithNextSection property.
See also:
  PrintWithNextSection}
    procedure SetPrintWithNextSection(Value: Boolean);
{Returns the value of the PrintWithPreviosSection property.
See also:
  PrintWithPreviosSection}
    function GetPrintWithPreviosSection: Boolean;
{Sets the value of the PrintWithPreviosSection property.
See also:
  PrintWithPreviosSection}
    procedure SetPrintWithPreviosSection(Value: Boolean);

{Copies the properties of another section.
Parameters:
  ASource - The source section to assign to the this section.}
    procedure Assign(ASource: IvgrSection);

{Specifies the starting row for horizontal sections and the starting
column for vertical sections.
See also:
  EndPos}
    property StartPos: Integer read GetStartPos;
{Specifies the ending row for horizontal sections and the ending
column for vertical sections.
See also:
  StartPos}
    property EndPos: Integer read GetEndPos;
{Returns the section's level, this property is calculated automatically
on the base of StartPos and EndPos properties.
The section with coordinates [1, 4] is a parent for section
with coordinates [2, 3] for example.
See also:
  StartPos, EndPos}
    property Level: Integer read GetLevel;
{Returns the interface to the parent section for this section.}
    property Parent: IvgrSection read GetParent;

{Gets or sets the value indicating whether the section's columns or rows will be
reprinted at the top of each page.}
    property RepeatOnPageTop: Boolean read GetRepeatOnPageTop write SetRepeatOnPageTop;
{Gets or sets the value indicating whether the section's columns or rows will be
reprinted at the bottom of each page.}
    property RepeatOnPageBottom: Boolean read GetRepeatOnPageBottom write SetRepeatOnPageBottom;
{Gets or sets the value indicating that even one row of the following section of the same level
should be on the current page, otherwise the section will be moved to the new page.
This property can be used, for prevent situations in which the group header is printed
separately from the group data.}
    property PrintWithNextSection: Boolean read GetPrintWithNextSection write SetPrintWithNextSection;
{Gets or sets the value indicating that even one row of the previous section of the same level
should be on the current page, otherwise the last row or column of the previous section
will be moved to the new page. This property can be used, for prevent situations in which
the group footer is printed separately from the group data.}
    property PrintWithPreviosSection: Boolean read GetPrintWithPreviosSection write SetPrintWithPreviosSection;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrSectionExt
  //
  /////////////////////////////////////////////////
{Extention for IvgrSection interface.}
  IvgrSectionExt = interface(IvgrWBListItem)
  ['{540E34BA-97AD-400D-946E-9C01D3FF78B1}']
{Returns the value of the Text property.
See also:
  Text}
    function GetText: string;
{Returns the string, associated with section, for displaying in the grid.}
    property Text: string read GetText;
  end;

  /////////////////////////////////////////////////
  //
  // Classes
  //
  /////////////////////////////////////////////////

  /////////////////////////////////////////////////
  //
  // TvgrWBListItem
  //
  /////////////////////////////////////////////////
{TvgrWBListItemClass defines the metaclass for TvgrWBListItem.}
  TvgrWBListItemClass = class of TvgrWBListItem;
{TvgrWBListItem implements the IvgrWBListItem interface.}
  TvgrWBListItem = class(TObject, IvgrWBListItem, IDispatch)
  private
    FList: TvgrWBList;
    FValid: Boolean;
    FRefCount: Integer;
    { Pointer to item data }
    FData: Pointer;
    { Index of item data }
    FDataIndex: Integer; // !!!
  protected
    procedure BeforeChangeProperty;
    procedure AfterChangeProperty;
    function GetEditChangesType: TvgrWorkbookChangesType; virtual; abstract;

    // IUnknown
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    // IDispatch
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual;  stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    function GetDispIdOfName(const AName: String): Integer; virtual;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; virtual;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; virtual;

    function GetWorksheet: TvgrWorksheet;
    function GetWorkbook: TvgrWorkbook;
    function GetItemData: Pointer; virtual;
    function GetStyleData: Pointer; virtual;
    function GetValid: Boolean;
    procedure SetInvalid;

    function GetItemIndex: Integer;
    property ItemData: Pointer read GetItemData;
    property ItemIndex: Integer read GetItemIndex;
    property StyleData: Pointer read GetStyleData;
    property Valid: Boolean read GetValid;
    property Worksheet: TvgrWorksheet read GetWorksheet;
    property Workbook: TvgrWorkbook read GetWorkbook;
  public
{Creates an instance of the TvgrWBListItem class.}
    constructor Create(AList: TvgrWBList; AData: Pointer; ADataIndex: Integer);
{Frees an instance of the TvgrWBListItem class.}
    destructor Destroy; override;
{$IFDEF VGR_DS_DEBUG}
    property RefCount: Integer read FRefCount;
{$ENDIF}
  end;

{Syntax:
  TvgrCellShiftState = (vgrcsSuccess, vgrcsLargeRange)}
  TvgrCellShiftState = (vgrcsSuccess, vgrcsLargeRange);
{Syntax:
  TvgrCellShift = (vgrcsLeft, vgrcsTop)}
  TvgrCellShift = (vgrcsLeft, vgrcsTop);
{For internal use.}
  TvgrConvertRecordProc = procedure(AOldRecord, ANewRecord: Pointer; AOldSize, ANewSize, ADataStorageVersion: Integer) of object;
{This call back procedure is called each time when the new object
is found in the list derived from the TvgrWBList class.
Parameters:
  AItem - The interface to the found object.
  AItemIndex - The index of the object in the list.
  AData - The user-defined data is passed to the find procedure.
See also:
  IvgrWBListItem}
  TvgrFindListItemCallBackProc = procedure(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer) of object;
  /////////////////////////////////////////////////
  //
  // TvgrWBList
  //
  /////////////////////////////////////////////////
{Base class for all lists of the worksheet (TvgrRanges, TvgrCols, TvgrBorders and so on).}
  TvgrWBList = class(TObject)
  private
    FWorksheet: TvgrWorksheet;
    FRecs: Pointer;
    FRecSize: Integer;
    FCount: Integer;
    FCapacity: Integer;
    FGrowSize: Integer;
    FItems: TList;

    FListFieldStatistic1 : TvgrListFieldStatistic;
    FListFieldStatistic2 : TvgrListFieldStatistic;
    function GetIntfItem(Index: Integer): TvgrWBListItem;
    function GetIntfItemCount: Integer;
    function GetDataListItem(Index: Integer): Pointer;
    function GetWorkbook: TvgrWorkbook;
    procedure UpdateDSListItems(AStartIndex: Integer; AOffset: Integer);

    procedure CalcRangeOfSearching(Idx1,Idx2,Size1,Size2 : Integer; out LowLev1,LowLev2,HiLev1,HiLev2 : Integer; DeleteIfCrossOver : Boolean);
    procedure SetLowRulerSpeedy(LowLev1,LowLev2 : Integer; var LowRuler : Integer);
    function CheckSearchRule(Index,LowLev1,LowLev2 : Integer) : boolean;
    function CheckOuterSearchRule(Index,HiLev1,HiLev2 : Integer) : boolean;
    procedure CorrectLowRuler(ANumber1,ANumber2,CurPos : Integer; var LowRuler : Integer);

    function ItemGetIndexField1(AIndex: Integer): Integer; virtual;
    function ItemGetIndexField2(AIndex: Integer): Integer; virtual;
    function ItemGetIndexField3(AIndex: Integer): Integer; virtual;
    function ItemGetSizeField1(AIndex: Integer): Integer; virtual;
    function ItemGetSizeField2(AIndex: Integer): Integer; virtual;

    procedure ItemSetIndexField1(AIndex: Integer; Value: Integer); virtual;
    procedure ItemSetIndexField2(AIndex: Integer; Value: Integer); virtual;
    procedure ItemSetIndexField3(AIndex: Integer; Value: Integer); virtual;
    procedure ItemSetSizeField1(AIndex: Integer; Value: Integer); virtual;
    procedure ItemSetSizeField2(AIndex: Integer; Value: Integer); virtual;

    function ItemInRect(AIndex: Integer; const ARect: TRect): Boolean;
    function ItemInsideRect(AIndex: Integer; const ARect: TRect): Boolean;
    function ItemInside(AIndex: Integer; AStartPos, AEndPos, ADimension: Integer): Boolean;
    procedure InternalDeleteLines(AStartPos, AEndPos: Integer);
    procedure InternalDeleteCols(AStartPos, AEndPos: Integer);
    procedure InternalInsertLines(AIndexBefore, ACount: Integer); virtual;
    procedure InternalInsertCols(AIndexBefore, ACount: Integer);  virtual;
    procedure MoveItem(AIndex, AStartPos, AEndPos, ADimension: Integer);
    procedure RecalcStatistics;
  protected
    function GetCount: Integer; virtual;
    function GetCreateChangesType: TvgrWorkbookChangesType; virtual; abstract;
    function GetDeleteChangesType: TvgrWorkbookChangesType; virtual; abstract;
    function GetWBListItemClass: TvgrWBListItemClass; virtual; abstract;
    function GetRecSize: Integer; virtual; abstract;
    function GetGrowSize: Integer; virtual;
    function CreateWBListItem(Index: Integer): TvgrWBListItem;
    procedure SaveToStream(AStream: TStream); virtual;
    class procedure LoadFromStreamConvertRecord(AOldRecord, ANewRecord: Pointer; AOldSize, ANewSize, ADataStorageVersion: Integer); virtual;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); virtual;
    function Add: Integer;
    procedure Insert(AIndex: Integer);
    procedure Delete(AIndex: Integer);
    procedure InternalDelete(Index : Integer);
    procedure InternalAdd(ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer; LowRuler : Integer);
    function InternalFind(ANumber1, ANumber2, ANumber3 : Integer;
                          ASizeField1, ASizeField2 : Integer;
                          DeleteIfCrossOver : Boolean;
                          AutoCreateItem : Boolean) : TvgrWBListItem;
    procedure InternalFindItemsInRect(ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc; DeleteItem : boolean; AData: Pointer);
    procedure InternalFindAndCallBack(const ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer);
    procedure DoDelete(AData: Pointer); virtual;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); virtual;
    procedure BeforeChangeItem(const ChangeInfo: TvgrWorkbookChangeInfo);
    procedure AfterChangeItem(const ChangeInfo: TvgrWorkbookChangeInfo);
    procedure InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer); virtual;
    procedure InternalExtendedActionInsertCols(AIndexBefore, ACount: Integer); virtual;

{$IFNDEF VGR_DS_DEBUG}
    property IntfItems[Index: Integer]: TvgrWBListItem read GetIntfItem;
    property IntfItemCount: Integer read GetIntfItemCount;
{$ENDIF}
    property DataList[Index: Integer]: Pointer read GetDataListItem;
    property ListFieldStatistic1 : TvgrListFieldStatistic read FListFieldStatistic1;
    property ListFieldStatistic2 : TvgrListFieldStatistic read FListFieldStatistic2;
  public
{Creates an instance ot the TvgrWBList class.
Parameters:
  AWorksheet - The TvgrWorksheet object which contains this list.}
    constructor Create(AWorksheet: TvgrWorksheet); virtual;
{Frees an instance of the TvgrWBList class.}
    destructor Destroy; override;
    procedure Clear; virtual;
{$IFDEF VGR_DS_DEBUG}
    function DebugInfo: TvgrDebugInfo;
    property IntfItems[Index: Integer]: TvgrWBListItem read GetIntfItem;
    property IntfItemCount: Integer read GetIntfItemCount;
{$ENDIF}
{Returns the amount of the objects in the list.}
    property Count: Integer read GetCount;
{Returns the TvgrWorksheet object which contains this list.}
    property Worksheet: TvgrWorksheet read FWorksheet;
{Returns the TvgrWorkbook object which contains this list.}
    property Workbook: TvgrWorkbook read GetWorkbook;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrVector
  //
  /////////////////////////////////////////////////
{TvgrVector implements the IvgrWBListItem interface.}
  TvgrVector = class(TvgrWBListItem, IvgrVector)
  private
    function GetData: pvgrVector;
  protected
//    function GetIndexField1 : Integer; override;

    // IvgrVector
    function GetSize: Integer;
    procedure SetSize(Value: Integer);
    function GetVisible: Boolean;
    procedure SetVisible(Value: Boolean);
    function GetNumber: Integer;
    { Copies the properties of another vector. }
    procedure Assign(ASource: IvgrVector); reintroduce; virtual;

    function GetDispIdOfName(const AName: String): Integer; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    property Data: pvgrVector read GetData;
    property Number: Integer read GetNumber;
    property Size: Integer read GetSize write SetSize;
    property Visible: Boolean read GetVisible write SetVisible;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrVectors
  //
  /////////////////////////////////////////////////
{Base class for the list of columns (TvgrCols) and list of rows (TvgrRows).
This list contains the TvgrVector objects.
See also:
  TvgrVector, TvgrCol, TvgrRow}
  TvgrVectors = class(TvgrWBList)
  private
    function GetItem(Number: Integer): IvgrVector;
    function GetByIndex(Index: Integer): IvgrVector;
  protected
    function ItemGetIndexField1(AIndex: Integer): Integer; override;
    procedure ItemSetIndexField1(AIndex: Integer; Value: Integer); override;

    function GetWBListItemClass: TvgrWBListItemClass; override;
    function GetRecSize: Integer; override;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
    procedure Delete(AStartPos, AEndPos: Integer);
    procedure Insert(AIndexBefore, ACount: Integer);
    procedure InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer); override;
  public
{Use this property to access to the vector in the list by its number.
If vector is not found it is created.
Parameters:
  Number - The vector's number.}
    property Items[Number: Integer]: IvgrVector read GetItem; default;
{Lists the vectors in the list.
Parameters:
  Index - The vector's index.}
    property ByIndex[Index: Integer]: IvgrVector read GetByIndex;
{Finds the vectors in the specified interval.
Parameters:
  AStartPos - The start of the interval.
  AEndPos - The end of the interval.
  CallBackProc - The callback procedure which is called for the each found object.
  AData - The user defined data, which will be passed to the callback procedure.}
    procedure FindAndCallBack(AStartPos, AEndPos: Integer; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer);
{Finds the vector by its number.
Parameters:
  Number - The vector's number.
See also:
  Returns the interface to the vector or nil if vector is not found.}
    function Find(Number: Integer): IvgrVector;
  end;


  /////////////////////////////////////////////////
  //
  // TvgrPageVectors
  //
  /////////////////////////////////////////////////
  TvgrPageVectors = class(TvgrVectors)
  protected
    function GetWBListItemClass: TvgrWBListItemClass; override;
    function GetRecSize: Integer; override;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPageVector
  //
  /////////////////////////////////////////////////
{TvgrPageVector implements the IvgrPageVector interface.
See also:
  IvgrPageVector}
  TvgrPageVector = class(TvgrVector, IvgrPageVector)
  private
    function GetFlags(Index: Integer): Boolean; virtual; abstract;
    procedure SetFlags(Index: Integer; Value: Boolean); virtual; abstract;
    function GetPageBreak: Boolean;
    procedure SetPageBreak(Value: Boolean);
  protected
    procedure Assign(ASource: IvgrVector); override;

    function GetDispIdOfName(const AName: string) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    property PageBreak: Boolean Index 0 read GetFlags write SetFlags;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrCol
  //
  /////////////////////////////////////////////////
{TvgrCol implements the IvgrCol interface.
See also:
  TvgrCols}
  TvgrCol = class(TvgrPageVector, IvgrCol)
  private
    function GetData: pvgrCol;
    function GetFlags(Index: Integer): Boolean; override;
    procedure SetFlags(Index: Integer; Value: Boolean); override;
  protected
    function GetEditChangesType: TvgrWorkbookChangesType; override;

    function GetDispIdOfName(const AName: string): Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    property Data: pvgrCol read GetData;

    property Width: Integer read GetSize write SetSize;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrCols
  //
  /////////////////////////////////////////////////
{Represents the list of worksheet's columns.
Each TvgrWorksheet object contains one instance of the TvgrCols object.
Main properties:
  - Use the Items property to add new columns in the list.
  - Use the Find procedure to find the column in the list by its number.
See also:
  TvgrWorksheet, TvgrWorksheet.Cols}
  TvgrCols = class(TvgrPageVectors)
  private
    function GetItem(Number: Integer): IvgrCol;
    function GetByIndex(Index: Integer): IvgrCol;
  protected
    function GetRecSize: Integer; override;
    function GetCreateChangesType: TvgrWorkbookChangesType; override;
    function GetDeleteChangesType: TvgrWorkbookChangesType; override;
    function GetWBListItemClass: TvgrWBListItemClass; override;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
  public
{Searches the column with specified number and returns interface to it,
if the object is not found then creates it and returns the interface to it.
Parameters:
  Number - The number of column.
Example:
  var
    I: Integer;
  begin
    for I := 0 to 10 do
      with Worksheet.Cols[I] do
        Width := ConvertUnitsToTwips(1, vgruInches);
  end;}
    property Items[Number: Integer]: IvgrCol read GetItem; default;
{Returns the column's interface by the column's index in the list.
Parameters:
  Index - The index of the column, must be from 0 to Count - 1}
    property ByIndex[Index: Integer]: IvgrCol read GetByIndex;
{Searches the column by its number and returns interface to it,
if column is not found returns nil.
Parameters:
  Number - The number of the column.}
    function Find(Number: Integer): IvgrCol;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRow
  //
  /////////////////////////////////////////////////
{TvgrRow implements the IvgrRow interface.
See also:
  TvgrRows}
  TvgrRow = class(TvgrPageVector, IvgrRow)
  private
    function GetData: pvgrRow;
    function GetFlags(Index: Integer): Boolean; override;
    procedure SetFlags(Index: Integer; Value: Boolean); override;
  protected
    function GetEditChangesType: TvgrWorkbookChangesType; override;

    function GetDispIdOfName(const AName: string) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    property Data: pvgrRow read GetData;
  public
    property Height: Integer read GetSize write SetSize;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRows
  //
  /////////////////////////////////////////////////
{Represents the list of worksheet's rows.
Each TvgrWorksheet object contains one instance of the TvgrRows object.
Main properties:<br>
  - Use the Items property to add new rows in the list.<br>
  - Use the Find procedure to find the row in the list by its number.<br>
See also:
  TvgrWorksheet, TvgrWorksheet.Rows}
  TvgrRows = class(TvgrPageVectors)
  private
    function GetItem(Number: Integer): IvgrRow;
    function GetByIndex(Index: Integer): IvgrRow;
  protected
    function GetRecSize: Integer; override;
    function GetCreateChangesType: TvgrWorkbookChangesType; override;
    function GetDeleteChangesType: TvgrWorkbookChangesType; override;
    function GetWBListItemClass: TvgrWBListItemClass; override;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
  public
{Searches the row with specified number and returns interface to it,
if the row is not found then creates it and returns its interface.
Parameters:
  Number - The number of row.
Example:
  var
    I: Integer;
  begin
    for I := 0 to 10 do
      with Worksheet.Rows[I] do
        Height := ConvertUnitsToTwips(1, vgruInches);
  end;}
    property Items[Number: Integer]: IvgrRow read GetItem; default;
{Returns the row's interface by the row's index in the list.
Parameters:
  Index - The index of the row, must be from 0 to Count - 1}
    property ByIndex[Index: Integer]: IvgrRow read GetByIndex;
{Searches the row by its number and returns interface to it,
if row is not found returns nil.
Parameters:
  Number - The number of the row.}
    function Find(Number: Integer): IvgrRow;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrStylesList
  //
  /////////////////////////////////////////////////
{Base class for the list of styles.
This is an internal class.}
  TvgrStylesList = class(TObject)
  private
    FItems: TList;
    FStyleItemSize: Integer;
    function GetItem(Index: Integer): Pointer;
    function GetCount: Integer;
    property Items[Index: Integer]: Pointer read GetItem; default;
  protected
    procedure SaveToStream(AStream: TStream); virtual;
    procedure LoadFromStreamConvertRecord(AOldData: Pointer;
                                          ANewData: Pointer;
                                          AOldHeaderSize, AOldDataSize,
                                          ANewHeaderSize, ANewDataSize,
                                          ADataStorageVersion: Integer); virtual;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); virtual;
    function GetStyleItemSize: Integer; virtual; abstract;
    procedure InitStyle(AStyle: Pointer); virtual;
    function InternalAdd(var AStyle: pvgrStyleHeader): Integer;
    function InternalFindOrAdd(Data: Pointer): Integer;

    procedure Clear;
    procedure ClearAll;
    procedure Release(AIndex: Integer);
    procedure AddRef(AIndex: Integer);
  public
{Creates an instance of the TvgrStylesList class.}
    constructor Create;
{Frees an instance of the TvgrStylesList class.}
    destructor Destroy; override;
{$IFDEF VGR_DEBUG}
    function DebugInfo: TvgrDebugInfo;
{$ENDIF}
{Returns the amount of the items in the list.}
    property Count: Integer read GetCount;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRangeStylesList
  //
  /////////////////////////////////////////////////
{Represents the list of the ranges' styles.
This is an internal class.}
  TvgrRangeStylesList = class(TvgrStylesList)
  private
    function GetItem(Index: Integer): pvgrRangeStyle;
  protected
    function GetStyleItemSize: Integer; override;
    procedure InitStyle(AStyle: Pointer); override;

    function FindOrAdd(const Style: rvgrRangeStyle; AOldIndex: Integer): Integer;
    property Items[Index: Integer]: pvgrRangeStyle read GetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRange
  //
  /////////////////////////////////////////////////
{TvgrRange implements IvgrRange and IvgrRangeFont interfaces.}
  TvgrRange = class(TvgrWBListItem, IvgrRange, IvgrRangeFont)
  private
    procedure _AssignStyle(ASource: IvgrRange);
    function GetData: pvgrRange;
    function GetStyle: pvgrRangeStyle;
    function GetStyles: TvgrRangeStylesList;
  protected
    function GetEditChangesType: TvgrWorkbookChangesType; override;

//    function GetIndexField1 : Integer; override;
//    function GetIndexField2 : Integer; override;
//    function GetSizeField1 : Integer; override;
//    function GetSizeField2 : Integer; override;
//    procedure SetSizeField1(Value : Integer); override;
//    procedure SetSizeField2(Value : Integer); override;

    function GetStyleData: Pointer; override;

    function GetLeft: Integer;
    function GetTop: Integer;
    function GetRight: Integer;
    function GetBottom: Integer;
    function GetPlace: TRect;
    function GetValue: Variant;
    procedure SetValue(const Value: Variant);
    function GetDisplayText: string;
    function GetFont: IvgrRangeFont;

    procedure ChangeValueType(ANewValueType: TvgrRangeValueType);
    { Copies the properties of another range. }
    procedure Assign(ASource: IvgrRange); reintroduce;
    procedure AssignStyle(ASource: IvgrRange);
    function GetValueData: pvgrRangeValue;

    procedure SetFont(Value: IvgrRangeFont);
    function GetFillBackColor: TColor;
    procedure SetFillBackColor(Value: TColor);
    function GetFillForeColor: TColor;
    procedure SetFillForeColor(Value: TColor);
    function GetFillPattern: TBrushStyle;
    procedure SetFillPattern(Value: TBrushStyle);
    function GetDisplayFormat: string;
    procedure SetDisplayFormat(const Value: string);
    function GetHorzAlign: TvgrRangeHorzAlign;
    procedure SetHorzAlign(Value: TvgrRangeHorzAlign);
    function GetVertAlign: TvgrRangeVertAlign;
    procedure SetVertAlign(Value: TvgrRangeVertAlign);
    function GetAngle: Word;
    procedure SetAngle(Value: Word);
    function GetFlags: Word;
    procedure SetFlags(Value: Word);
    function GetWordWrap: Boolean;
    procedure SetWordWrap(Value: Boolean);
    function GetFormula: string;
    procedure SetFormula(const Value: string);
    function GetValueType: TvgrRangeValueType;
    function GetStringValue: string;
    procedure SetStringValue(const Value: string);
    function GetSimpleStringValue: string;
    procedure SetSimpleStringValue(const Value: string);
    function GetWBStrings: TvgrWBStrings;
    // IvgrRangeFont
    function GetFontSize: Integer;
    procedure SetFontSize(Value: Integer);
    function GetFontPitch: TFontPitch;
    procedure SetFontPitch(Value: TFontPitch);
    function GetFontStyle: TFontStyles;
    procedure SetFontStyle(Value: TFontStyles);
    function GetFontCharset: TFontCharset;
    procedure SetFontCharset(Value: TFontCharset);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontColor: TColor;
    procedure SetFontColor(Value: TColor);
    function GetExportData: Pointer;
    procedure SetExportData(Value: Pointer);

    procedure FontAssign(AFont: TFont);
    procedure FontAssignRangeFont(AFont: IvgrRangeFont);
    procedure FontAssignTo(AFont: TFont);

    procedure IvgrRangeFont.Assign = FontAssign;
    procedure IvgrRangeFont.AssignRangeFont = FontAssignRangeFont;
    procedure IvgrRangeFont.AssignTo = FontAssignTo;

    procedure SetDefaultStyle;

    function GetDispIdOfName(const AName: string) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    property Data: pvgrRange read GetData;
    property Style: pvgrRangeStyle read GetStyle;
    property Styles: TvgrRangeStylesList read GetStyles;
    property WBStrings: TvgrWBStrings read GetWBStrings;

    property Left: Integer read GetLeft;
    property Top: Integer read GetTop;
    property Right: Integer read GetRight;
    property Bottom: Integer read GetBottom;
    property Place: TRect read GetPlace;
    property Value: Variant read GetValue write SetValue;
    property DisplayText: string read GetDisplayText;
    property Font: IvgrRangeFont read GetFont write SetFont;
    property FillBackColor: TColor read GetFillBackColor write SetFillBackColor;
    property FillForeColor: TColor read GetFillForeColor write SetFillForeColor;
    property FillPattern: TBrushStyle read GetFillPattern write SetFillPattern;
    property DisplayFormat: string read GetDisplayFormat write SetDisplayFormat;
    property HorzAlign: TvgrRangeHorzAlign read GetHorzAlign write SetHorzAlign;
    property VertAlign: TvgrRangeVertAlign read GetVertAlign write SetVertAlign;
    property Angle: Word read GetAngle write SetAngle;
    property WordWrap: Boolean read GetWordWrap write SetWordWrap;
    property ValueData: pvgrRangeValue read GetValueData;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPerformDeleteCellInfo
  //
  /////////////////////////////////////////////////
{For internal use.}
  TvgrPerformDeleteCellInfo = record
    Success: Boolean;
    AllowBreakRanges: Boolean;
    SearchRect: TRect;
    Shift: TvgrCellShift;
    ShiftTo: Integer;
    BottomRange: Integer;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRanges
  //
  /////////////////////////////////////////////////
{Represents the list of worksheet's ranges.
Each range can contain one or more cells, and is defined on a worksheet
by a rectangle which contains coordinate top-left and bottom-right
corners of a range. Coordinates are specified since zero.
The ranges can not overlap, for example: if you create range with coordinates
[0, 0, 4, 4] and then create a range with coordinates [2, 2, 2, 2] - the
first range will be deleted.
Main properties of the TvgrRanges class:<br>
  - Use the Items property to add new ranges in the list.<br>
  - Use the Find procedure to find the existing ranges in the list.<br>
See also:
  TvgrWorksheet, TvgrWorksheet.RangesList}
  TvgrRanges = class(TvgrWBList)
  private
    FTemp: IvgrRange;
    FTempShift: IvgrRange;
    FTempShiftIndex: Integer;
    FPerformDeleteCellInfo: TvgrPerformDeleteCellInfo;

    function GetItem(Left,Top,Right,Bottom : integer): IvgrRange;
    function GetByIndex(Index: Integer): IvgrRange;
    procedure CallBackFindAtCell(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    function GetStyles: TvgrRangeStylesList;
    function GetWBStrings: TvgrWBStrings;
    function GetFormulas: TvgrFormulasList;
    procedure DeleteRows(AStartPos, AEndPos: Integer);
    procedure DeleteCols(AStartPos, AEndPos: Integer);
    function PerformDeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean): Boolean;
    procedure PerformDeleteCellsCallBack(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure DeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean);
    procedure DeleteCellsCallBack(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure ShiftToTop;
    function FindAtCellForShift(X, Y: Integer): IvgrRange;
    procedure CallBackFindAtCellForShift(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  protected
    procedure SaveToStream(AStream: TStream); override;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); override;
    function ItemGetIndexField1(AIndex: Integer): Integer; override;
    function ItemGetIndexField2(AIndex: Integer): Integer; override;
    function ItemGetSizeField1(AIndex: Integer): Integer; override;
    function ItemGetSizeField2(AIndex: Integer): Integer; override;

    procedure ItemSetIndexField1(AIndex: Integer; Value: Integer); override;
    procedure ItemSetIndexField2(AIndex: Integer; Value: Integer); override;
    procedure ItemSetSizeField1(AIndex: Integer; Value: Integer); override;
    procedure ItemSetSizeField2(AIndex: Integer; Value: Integer); override;

    procedure ReleaseReferences(ARange: pvgrRange);
    procedure InternalInsertLines(AIndexBefore, ACount: Integer); override;
    procedure InternalInsertCols(AIndexBefore, ACount: Integer); override;

    function GetCreateChangesType: TvgrWorkbookChangesType; override;
    function GetDeleteChangesType: TvgrWorkbookChangesType; override;
    function GetWBListItemClass: TvgrWBListItemClass; override;
    function GetRecSize: Integer; override;
    function GetGrowSize: Integer; override;
    procedure DoDelete(AData: Pointer); override;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
    property Styles: TvgrRangeStylesList read GetStyles;
    property WBStrings: TvgrWBStrings read GetWBStrings;
    property Formulas: TvgrFormulasList read GetFormulas;
  public
{Searches the range with specified area and returns interface to it,
if the range is not found then creates it and returns its interface.
Parameters:
  Left - The X coordinate of the top-left range corner.
  Top - The Y coordinate of the top-left range corner.
  Right - The X coordinate of the bottom-right range corner.
  Bottom - The Y coordinate of the bottom-right range corner.
Example:
  var
    I: Integer;
  begin
    // create the rectangle of cells
    for I := 0 to 10 do
      for J := 0 to 10 do
        with Worksheet.Ranges[I, J, I, J] do
          Value := Format('Cell %d, %d', [I, J]);
  end;}
    property Items[Left,Top,Right,Bottom : integer]: IvgrRange read GetItem; default;
{Returns the range's interface by the range's index in the list.
Parameters:
  Index - The index of the range, must be from 0 to Count - 1
Example:
  var
    I: Integer;
  begin
    // set font to bold for all ranges on worksheet
    for I := 0 to AWorksheet.RangesList.Count - 1 do
      with AWorksheet.RangesList[I] do
        FontStyle := FontStyle + [fsBold];
  end;}
    property ByIndex[Index: Integer]: IvgrRange read GetByIndex;
{Searches the range by its area and returns interface to it,
if range is not found returns nil.
Parameters:
  Left - The X coordinate of the top-left range corner.
  Top - The Y coordinate of the top-left range corner.
  Right - The X coordinate of the bottom-right range corner.
  Bottom - The Y coordinate of the bottom-right range corner.
Return value:
  The interface to the found range.}
    function Find(Left, Top, Right, Bottom : integer) : IvgrRange;
{Searches the range that contains a cell with specified coordinates,
if range is not found returns nil.
Parameters:
  X - The X coordinate of the cell.
  Y - The Y coordinate of the cell.
Return value:
  The interface to the found range.}
    function FindAtCell(X, Y: Integer): IvgrRange;
{Searches the ranges within the specified area.
Parameters:
  ARect - The area of searching.
  CallBackProc - The procedure which will be called for each found range.
  AData - The user-data passed to the CallbackProc.
Example:
  procedure TMyForm.DumpRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  var
    s: string;
  begin
    with AItem as IvgrRange do
      s := Format('Range: [%d, %d, %d, %d], Value: [%s]', [Left, Top, Right, Bottom, StringValue]);
    OutputDebugString(PChar(s));
  end;

  procedure TMyForm.DumpRanges(AWorksheet: TvgrWorksheet; const AArea: TRect);
  begin
    AWorksheet.RangesList.FindAndCallback(AArea, FindCallback, nil);
  end;
See also:
  TvgrFindListItemCallBackProc}
    procedure FindAndCallBack(const ARect: TRect; CallBackProc: TvgrFindListItemCallBackProc; AData: Pointer); overload;
{Fills the AFont object with the default font data.
See also:
  AFont - The TFont object to fill.}
    procedure GetDefaultFont(AFont: TFont);
{Clears the list.}
    procedure Clear; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBorderStylesList
  //
  /////////////////////////////////////////////////
{Represents the list of the borders' styles.
This is an internal class.}
  TvgrBorderStylesList = class(TvgrStylesList)
  private
    function GetItem(Index: Integer): pvgrBorderStyle;
  protected
    function GetStyleItemSize: Integer; override;
    procedure InitStyle(AStyle: Pointer); override;

    function FindOrAdd(const Style: rvgrBorderStyle; AOldIndex: Integer): Integer;
    property Items[Index: Integer]: pvgrBorderStyle read GetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBorder
  //
  /////////////////////////////////////////////////
{TvgrBorder implemets the IvgrBorder interface.}
  TvgrBorder = class(TvgrWBListItem, IvgrBorder)
  private
    function GetData: pvgrBorder;
    function GetStyle: pvgrBorderStyle;
    function GetStyles: TvgrBorderStylesList;
  protected
//    function GetIndexField1 : Integer; override;
//    function GetIndexField2 : Integer; override;
//    function GetIndexField3 : Integer; override;

    function GetStyleData: Pointer; override;

    function GetLeft: Integer;
    function GetTop: Integer;
    function GetOrientation: TvgrBorderOrientation;
    function GetWidth: Integer;
    procedure SetWidth(Value: Integer);
    function GetColor: TColor;
    procedure SetColor(Value: TColor);
    function GetPattern: TvgrBorderStyle;
    procedure SetPattern(Value: TvgrBorderStyle);
    function GetEditChangesType: TvgrWorkbookChangesType; override;
    procedure Assign(ASource: IvgrBorder); reintroduce;

    function GetDispIdOfName(const AName: string) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;
    procedure SetStyles(AWidth: Integer; AColor: TColor; APattern: TvgrBorderStyle);

    property Data: pvgrBorder read GetData;
    property Style: pvgrBorderStyle read GetStyle;
    property Styles: TvgrBorderStylesList read GetStyles;

    property Left: Integer read GetLeft;
    property Top: Integer read GetTop;
    property Orientation: TvgrBorderOrientation read GetOrientation;
    property Width: Integer read GetWidth write SetWidth;
    property Color: TColor read GetColor write SetColor;
    property Pattern: TvgrBorderStyle read GetPattern write SetPattern;
  end;

{Describes the type of borders for formatting.
Items:
  vgrbtLeft - The vertical borders on the left edge.
  vgrbtCenter - The vertical borders within area.
  vgrbtRight - The vertical borders on the right edge.
  vgrbtTop - The horizontal borders on the top edge.
  vgrbtMiddle - The horizontal borders within area.
  vgrbtBottom - The horizontal borders on the bottom edge.
Syntax:
  TvgrBorderType = (vgrbtLeft, vgrbtCenter, vgrbtRight, vgrbtTop, vgrbtMiddle, vgrbtBottom)}
  TvgrBorderType = (vgrbtLeft, vgrbtCenter, vgrbtRight, vgrbtTop, vgrbtMiddle, vgrbtBottom);
{Describes the set of borders' types for formatting.
Syntax:
  TvgrBorderTypes = set of TvgrBorderType}
  TvgrBorderTypes = set of TvgrBorderType;
  /////////////////////////////////////////////////
  //
  // TvgrBorders
  //
  /////////////////////////////////////////////////
{Represents the list or worksheet's borders.
In GridReport the borders are not linked with cells ranges and
stored in the separate list. Each border is defined by coordinate of its
top-left corner and orientation - vertical or horizontal.
Coordinates are specified from zero.
Main properties of the TvgrBorders class:<br>
  - Use the Items property to add new borders in the list.<br>
  - Use the Find procedure to find the existing borders in the list.<br>
See also:
  TvgrWorksheet, TvgrWorksheet.BordersList}
  TvgrBorders = class(TvgrWBList)
  private
    function GetItem(Left,Top: Integer; Orientation: TvgrBorderOrientation): IvgrBorder;
    function GetByIndex(Index: Integer): IvgrBorder;
    function GetStyles: TvgrBorderStylesList;
    procedure DeleteRows(AStartPos, AEndPos: Integer);
    procedure DeleteCols(AStartPos, AEndPos: Integer);
    procedure DeleteCells(const ARect: TRect; ACellShift: TvgrCellShift);
    procedure FindAtRectAndCallBackCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure DeleteBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  protected
    procedure SaveToStream(AStream: TStream); override;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); override;
    function ItemGetIndexField1(AIndex: Integer): Integer; override;
    function ItemGetIndexField2(AIndex: Integer): Integer; override;
    function ItemGetIndexField3(AIndex: Integer): Integer; override;
    procedure ItemSetIndexField1(AIndex: Integer; Value: Integer); override;
    procedure ItemSetIndexField2(AIndex: Integer; Value: Integer); override;
    procedure ItemSetIndexField3(AIndex: Integer; Value: Integer); override;

    procedure ReleaseReferences(ABorder: pvgrBorder);
    procedure InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer); override;
    procedure InternalExtendedActionInsertCols(AIndexBefore, ACount: Integer); override;

    function GetCreateChangesType: TvgrWorkbookChangesType; override;
    function GetDeleteChangesType: TvgrWorkbookChangesType; override;
    function GetWBListItemClass: TvgrWBListItemClass; override;
    function GetRecSize: Integer; override;
    function GetGrowSize: Integer; override;
    procedure DoDelete(AData: Pointer); override;
    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
    property Styles: TvgrBorderStylesList read GetStyles;
  public
{Searches the border with specified coordinates and returns interface to it,
if the range is not found then creates it and returns its interface.
Parameters:
  Left - The X coordinate of the top-left border corner.
  Top - The Y coordinate of the top-left border corner.
  Orientation - The border type - vertical or horizontal.
Example:
  var
    I: Integer;
  begin
    // create the line of borders
    for I := 0 to 10 do
      with Worksheet.Borders[1, I, vgrboTop] do
      begin
        Width := ConvertUnitsToTwips(1, vgruMms);
        if (I mod 2) = 0 then
          Color := clRed
        else
          Color := clBlue;
      end;
  end;}
    property Items[Left, Top: Integer; Orientation: TvgrBorderOrientation]: IvgrBorder read GetItem; default;
{Returns the border's interface by the border's index in the list.
Parameters:
  Index - The index of the border, must be from 0 to Count - 1
Example:
  var
    I: Integer;
  begin
    // make all borders green
    for I := 0 to AWorksheet.BordersList.Count - 1 do
      with AWorksheet.BordersList[I] do
        Color := clGreen;
  end;}
    property ByIndex[Index: Integer]: IvgrBorder read GetByIndex;
{Searches the border by its coordinates and returns interface to it,
if range is not found returns nil.
Parameters:
  Left - The X coordinate of the top-left range corner.
  Top - The Y coordinate of the top-left range corner.
  Orientation - The type of border - vertical or horizontal.
Return value:
  The interface to the found border.}
    function Find(Left, Top: Integer; Orientation: TvgrBorderOrientation) : IvgrBorder;
{Searches the ranges within the specified rectangle.
This method does not found vertical borders which is placed on the right edge of rectangle
and does not found horizontal borders which is placed on the bottom edge of rectangle.
Parameters:
  ARect - The area of searching.
  CallBackProc - The procedure which will be called for each found range.
  AData - The user-data passed to the CallbackProc.
Example:
  procedure TMyForm.DumpBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  var
    s: string;
  begin
    with AItem as IvgrBorder do
      s := Format('Color: [%d], Width: [%d]', [Color, Width]);
    OutputDebugString(PChar(s));
  end;

  procedure TMyForm.DumpBorders(AWorksheet: TvgrWorksheet; const AArea: TRect);
  begin
    AWorksheet.BordersList.FindAtRectAndCallBack(AArea, DumpBordersCallback, nil);
  end;
See also:
  TvgrFindListItemCallBackProc}
    procedure FindAndCallBack(const ARect: TRect; CallBackProc: TvgrFindListItemCallBackProc; AData: Pointer);
{Searches the ranges within the specified rectangle.
Parameters:
  ARect - The area of searching.
  CallBackProc - The procedure which will be called for each found range.
  AData - The user-data passed to the CallbackProc.
Example:
  procedure TMyForm.DumpBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  var
    s: string;
  begin
    with AItem as IvgrBorder do
      s := Format('Color: [%d], Width: [%d]', [Color, Width]);
    OutputDebugString(PChar(s));
  end;

  procedure TMyForm.DumpBorders(AWorksheet: TvgrWorksheet; const AArea: TRect);
  begin
    AWorksheet.BordersList.FindAtRectAndCallBack(AArea, DumpBordersCallback, nil);
  end;
See also:
  TvgrFindListItemCallBackProc}
    procedure FindAtRectAndCallBack(const ARect: TRect; CallBackProc: TvgrFindListItemCallBackProc; AData: Pointer);
{Clears the list.}
    procedure Clear; override;

{Deletes borders within specified rectangle.
Parameters:
  ARect -  rectangle, which defines the area within of which the borders must be deleted.}
    procedure DeleteBorders(const ARect: TRect); overload;
{Deletes borders of specified type within specified rectangle.
Parameters:
  ARect -  rectangle, which defines the area within of which the borders must be deleted.
  ABorderTypes - specifies types of borders, which must be deleted.}
    procedure DeleteBorders(const ARect: TRect; ABorderTypes: TvgrBorderTypes); overload;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSection
  //
  /////////////////////////////////////////////////
{TvgrSection implements the IvgrSection interface.}
  TvgrSection = class(TvgrWBListItem, IvgrSection)
  private
    function GetStartPos: Integer;
    function GetEndPos: Integer;
    function GetLevel: Integer;
    function GetData: pvgrSection;
    function GetParent: IvgrSection;
    function GetFlags(Index: Integer): Boolean;
    procedure SetFlags(Index: Integer; Value: Boolean);
  protected
    function GetRepeatOnPageTop: Boolean;
    procedure SetRepeatOnPageTop(Value: Boolean);
    function GetRepeatOnPageBottom: Boolean;
    procedure SetRepeatOnPageBottom(Value: Boolean);
    function GetPrintWithNextSection: Boolean;
    procedure SetPrintWithNextSection(Value: Boolean);
    function GetPrintWithPreviosSection: Boolean;
    procedure SetPrintWithPreviosSection(Value: Boolean);
    function GetEditChangesType: TvgrWorkbookChangesType; override;

    function GetDispIdOfName(const AName: string) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    procedure Assign(ASource: IvgrSection); reintroduce;

    property Data: pvgrSection read GetData;

    property StartPos: Integer read GetStartPos;
    property EndPos: Integer read GetEndPos;
    property Level: Integer read GetLevel;
    property RepeatOnPageTop: Boolean Index 0 read GetFlags write SetFlags;
    property RepeatOnPageBottom: Boolean Index 1 read GetFlags write SetFlags;
    property PrintWithNextSection: Boolean Index 2 read GetFlags write SetFlags;
    property PrintWithPreviosSection: Boolean Index 3 read GetFlags write SetFlags;
  end;

{TvgrSectionsClass defines the metaclass for TvgrSections.}
  TvgrSectionsClass = class of TvgrSections;

{For internal use.}
  TvgrSectionsNotifyEvents = record
    NewEvent: TvgrWorkbookChangesType;
    DeleteEvent: TvgrWorkbookChangesType;
    EditEvent: TvgrWorkbookChangesType;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSections
  //
  /////////////////////////////////////////////////
{Represents the list or worksheet's sections (vertical and horizontal).
Each section includes one or more rows or columns.
The section is specified by two coordinates StartPos and EndPos,
StartPos - defines number of a line / column with which section begins and
EndPos - defines number of a line / column on which section comes to an end.
Sections can form multilevel hierarchical structure.
So for example the section with coordinates [1, 4] is a parent for section with coordinates [2, 3].
Section can be used in TvgrWorkbookGrid to quickly hides / shows some rows or columns.
See also:
  TvgrWorksheet, TvgrWorksheet.VertSectionsList, TvgrWorksheet.HorzSectionsList}
  TvgrSections = class(TvgrWBList)
  private
    FTempStartPos: Integer;
    FTempEndPos: Integer;
    FTempLevel: Integer;
    FMaxStatistic: TvgrListFieldStatistic;
    FItemExists: Boolean;
    FSectionsNotifyEvents: TvgrSectionsNotifyEvents;

    FCurrentOuterIndex: Integer;
    FCurrentOuterLevel: Integer;
    FDirectOuterIndex: Integer;
    FDirectOuterLevel: Integer;
    FDirectOuterLevelFor: Integer;
    FSearchInnersPresent: Boolean;

    FDeleteStartPos: Integer;
    FDeleteEndPos: Integer;
    FDeleteItemIndex: Integer;

    FItemsToDelete: Array of Integer;
    FOuterItems: Array of Integer;

    procedure AddLevelRef(ALevel: Integer);
    procedure ReleaseLevelRef(ALevel: Integer);
    procedure SearchInnersCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    function SearchInners(AStartPos, AEndPos: Integer): Boolean;
    procedure SearchCurrentOuterCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure SearchDirectOuterCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure DownLevel(AStartPos, AEndPos: Integer; var ALevel: Integer);
    procedure UpLevel(AStartPos, AEndPos: Integer; ALevel: Integer);
    procedure AddItemToDelete(AIndex: Integer);
    procedure AddOuterItem(AIndex: Integer);
    procedure CheckOutersLevel;
    procedure DeleteEnemyItems;
    procedure ProcessIntersectedSections(AStartPos, AEndPos, ALevelIncrement: Integer);
    procedure ProcessIntersectCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure ProcessDeleteCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  protected
    procedure DeleteRows(AStartPos, AEndPos: Integer); virtual;
    procedure SaveToStream(AStream: TStream); override;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); override;
    procedure InitNotifyEvents(ANew, ADelete, AEdit: TvgrWorkbookChangesType);

    procedure DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer); override;
    procedure DoDelete(AData: Pointer); override;

    function ItemGetIndexField1(AIndex: Integer): Integer; override;
    function ItemGetSizeField1(AIndex: Integer): Integer; override;

    procedure ItemSetIndexField1(AIndex: Integer; Value: Integer); override;
    procedure ItemSetSizeField1(AIndex: Integer; Value: Integer); override;

    function GetCreateChangesType: TvgrWorkbookChangesType; override;
    function GetDeleteChangesType: TvgrWorkbookChangesType; override;
    function GetEditChangesType: TvgrWorkbookChangesType; virtual;
    function GetWBListItemClass: TvgrWBListItemClass; override;
    function GetRecSize: Integer; override;

    function GetItem(StartPos, EndPos: Integer): IvgrSection; virtual;
    function GetByIndex(Index: Integer): IvgrSection; virtual;
    function GetMaxLevel: Integer; virtual;

    procedure InternalInsertLines(AIndexBefore, ACount: Integer); override;
    procedure InternalInsertCols(AIndexBefore, ACount: Integer);  override;
    procedure InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer); override;
  public
{Creates an instance of the TvgrSections class.}
    constructor Create(AWorksheet: TvgrWorksheet); override;
{Frees an instance of the TvgrSections class.}
    destructor Destroy; override;

{Searches the sections within the specified interval.
Parameters:
  AStartPos - The begin of the interval.
  AEndPos - The end of the interval.
  CallBackProc - The procedure which will be called for each found section.
  AData - The user-data passed to the CallbackProc.
Example:
  procedure TMyForm.DumpSectionsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  var
    s: string;
  begin
    with AItem as IvgrSection do
      s := Format('StartPos: [%d], EndPos: [%d], Level: [%d]', [StartPos, EndPos, Level]);
    OutputDebugString(PChar(s));
  end;

  procedure TMyForm.DumpSections(ASections: TvgrSections);
  begin
    ASections.FindAtRectAndCallBack(0, 10, DumpSectionsCallback, nil);
  end;
See also:
  TvgrFindListItemCallBackProc}
    procedure FindAndCallBack(AStartPos, AEndPos: Integer; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer); overload; virtual;
{Searches the section by its coordinates and returns interface to it,
if section is not found returns nil.
Parameters:
  StartPos - The begin of the interval.
  EndPos - The end of the interval.
Return value:
  The interface to the found section.}  
    function Find(StartPos, EndPos: Integer): IvgrSection; virtual;
{Deletes the section with specified coordinates.
See also:
  AStartPos - The begin of the section.
  AEndPos - The end of the section.}
    procedure Delete(AStartPos, AEndPos: Integer); virtual;
{Searches the section placing over interval from AStartPos to
AEndPos and at specified level.
Parameters:
  AStartPos - Specifies the starting row (for horizontal sections) or column (for
vertical) of the interval in which section must be.
  AEndPos - Specifies the ending row (for horizontal sections) or column (for
vertical) of the interval in which section must be.
  ALevel - Specifies the level of the section.
Return value:
  Returns index of the found section or -1.}
    function SearchCurrentOuter(AStartPos, AEndPos, ALevel: Integer): Integer; virtual;
{Searches the section with least level, which is placed over interval from
AStartPos to AEndPos and has level, greater than specified.
Parameters:
  AStartPos - Specifies the starting row (for horizontal bands) or column (for
vertical) of the interval in which section must be.
  AEndPos - Specifies the ending row (for horizontal bands) or column (for
vertical) of the interval in which section must be.
  ALevel - Specifies the level of the section.
Return value:
  Returns index of the found section or -1.}
    function SearchDirectOuter(AStartPos, AEndPos, ALevel: Integer): Integer; virtual;

{Searches the section with specified coordinates and returns interface to it,
if the section is not found then creates it and returns its interface.
Parameters:
  StartPos - The begin of the section.
The starting row for horizontal sections and the starting column for vertical.
  EndPos - The end of the section.
The ending row for horizontal sections and the ending column for vertical.
Example:
  begin
    // create the page header
    with Items[0, 2] do
    begin
      RepeatOnPageTop := True;
    end;

    // create the page footer
    with Items[100, 100] do
    begin
      RepeatOnPageBottom := True;
    end;
  end;}
    property Items[StartPos, EndPos: Integer]: IvgrSection read GetItem; default;
{Returns the section's interface by the section's index in the list.
Parameters:
  Index - The index of the section, must be from 0 to Count - 1.}
    property ByIndex[Index: Integer]: IvgrSection read GetByIndex;
{Returns a level of section of the highest level.}
    property MaxLevel: Integer read GetMaxLevel;
  end;

{Specifies the dynamical array of the TRect structures.}
  TvgrRectArray = array of TRect;

{Describes the operation with worksheet's borders.
For internal use.}
  TvgrRangesFormatBorderOperation = (vgrfoGet, vgrfoSetColor, vgrfoSetPattern, vgrfoSetWidth);
{Describes the set of operations with worksheet's borders.
For internal use.}
  TvgrRangesFormatBorderOperations = set of TvgrRangesFormatBorderOperation;

{The temporary object, for internal use.}
  TvgrRangesFormatBorderTmpObject = record
    Count: Integer;
    Operations: TvgrRangesFormatBorderOperations;
    Color: TColor;
    Width: Integer;
    Pattern: TvgrBorderStyle;
    HasColor: Boolean;
    HasWidth: Boolean;
    HasPattern: Boolean;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRangesFormatBorder
  //
  /////////////////////////////////////////////////
{TvgrRangesFormatBorder represents a class for editing the borders within the specified area.
TvgrRangesFormatBorder object can edit or create the borders of the type
which is specified by the ABorderType parameter in the constructor.
Please do not create instances of this class, use the property
TvgrWorksheet.RangesFormat for changing the borders format.
See also:
  TvgrRangesFormatBorders, TvgrWorksheet.RangesFormat}
  TvgrRangesFormatBorder = class
  private
    FAutoCreate: Boolean;
    FBorderType: TvgrBorderType;
    FBorders: TvgrBorders;
    FRectArray: TvgrRectArray;
    FTempObject: TvgrRangesFormatBorderTmpObject;
    function SameBorder(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation; ABorderType: TvgrBorderType): Boolean;
    procedure InitTempObject;
    function GetRect(ARect: TRect): TRect;
    procedure Execute;
    procedure ExecuteCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure ExecuteCallbackSet(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    function GetWidth: Integer;
    procedure SetWidth(Value: Integer);
    function GetColor: TColor;
    procedure SetColor(Value: TColor);
    function GetPattern: TvgrBorderStyle;
    procedure SetPattern(Value: TvgrBorderStyle);
    function GetHasWidth: Boolean;
    function GetHasColor: Boolean;
    function GetHasPattern: Boolean;
    procedure SavePropertyToList(AList: TStrings; const AName: string; const AValue: Variant);
    function LoadPropertyFromList(AList: TStrings; const AName: string; var AValue: Variant): Boolean;
    procedure SaveToList(AList: TStrings);
    procedure LoadFromList(AList: TStrings);
  public
{Creates an instance of the TvgrRangesFormatBorder class.
Parameters:
  ARects - The area within of which borders are edited.
  ABorders - The TvgrBorders object which contains edited borders.
  ABorderType - The type of borders within specified area. 
  AAutoCreate - Automatically create borders or change the properties of existing objects only.
See also:
  TvgrBorderType}
    constructor Create(ARects: TvgrRectArray; ABorders: TvgrBorders; ABorderType: TvgrBorderType; AAutoCreate: Boolean);
{Returns the borders' format.
Parameters:
  AColor - The color of the borders.
  AWidth - The width of the borders in twips.
  APattern - The borders' style.
  AHasColor - Contains the true if all borders within specified area have the equal color.
  AHasWidth - Contains the true if all borders within specified area have the equal width.
  AHasPattern - Contains the true if all borders within specified area have the equal pattern.
See also:
  TvgrBorderStyle}
    procedure GetFormat(out AColor: TColor; out AWidth: Integer; out APattern: TvgrBorderStyle; out AHasColor, AHasWidth, AHasPattern: Boolean);
{Sets the borders' format.
Parameters:
  AColor - The color of the borders.
  AWidth - The width of the borders in twips.
  APattern - The borders' style.
See also:
  TvgrBorderStyle}
    procedure SetFormat(AColor: TColor; AWidth: Integer; APattern: TvgrBorderStyle);
{Copies the properties from another TvgrRangesFormatBorder object.
Parameters:
  ASource - The source TvgrRangesFormatBorder object.}
    procedure AssignBorderStyle(ASource: TvgrRangesFormatBorder);

{Specifies the size of the borders in twips.}
    property Width: Integer read GetWidth write SetWidth;
{Specifies the color of the borders.}
    property Color: TColor read GetColor write SetColor;
{Specifies the pattern of the borders.
See also:
  TvgrBorderStyle}
    property Pattern: TvgrBorderStyle read GetPattern write SetPattern;
{Returns the true value if all borders within specified area have the equal width.}
    property HasWidth: Boolean read GetHasWidth;
{Returns the true value if all borders within specified area have the equal color.}
    property HasColor: Boolean read GetHasColor;
{Returns the true value if all borders within specified area have the equal pattern.}
    property HasPattern: Boolean read GetHasPattern;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRangesFormatBorders
  //
  /////////////////////////////////////////////////
{TvgrRangesFormatBorder represents a class for editing the borders within the specified area.
With using of this object you can get access to all borders within rectangular area.
Please do not create instances of this class, use the property
TvgrWorksheet.RangesFormat for changinh the borders format.
See also:
  TvgrRangesFormatBorder, TvgrWorksheet.RangesFormat}
  TvgrRangesFormatBorders = class
  private
    FAutoCreate: Boolean;
    FBorders: TvgrBorders;
    FRectArray: TvgrRectArray;
    FLeft: TvgrRangesFormatBorder;
    FCenter: TvgrRangesFormatBorder;
    FRight: TvgrRangesFormatBorder;
    FTop: TvgrRangesFormatBorder;
    FMiddle: TvgrRangesFormatBorder;
    FBottom: TvgrRangesFormatBorder;
    procedure SaveToList(AList: TStrings);
    procedure LoadFromList(AList: TStrings);
  public
{Creates an instance of the TvgrRangesFormatBorders class.
Parameters:
  ARects - The area within of which borders are edited.
  ABorders - The TvgrBorders object which contains edited borders.
  AAutoCreate - Automatically create borders or change the properties of existing objects only.}
    constructor Create(ARects: TvgrRectArray; ABorders: TvgrBorders; AAutoCreate: Boolean);
{Frees an instance of the TvgrRangesFormatBorders class.}
    destructor Destroy; override;
{Copies the properties from another TvgrRangesFormatBorders object.
Parameters:
  ASource - The source TvgrRangesFormatBorders object.}
    procedure AssignBordersStyles(ASource: TvgrRangesFormatBorders);
{Returns the TvgrRangesFormatBorder object for vertical borders on the left edge.}
    property Left: TvgrRangesFormatBorder read FLeft;
{Returns the TvgrRangesFormatBorder object for vertical borders within area.}
    property Center: TvgrRangesFormatBorder read FCenter;
{Returns the TvgrRangesFormatBorder object for vertical borders on the right edge.}
    property Right: TvgrRangesFormatBorder read FRight;
{Returns the TvgrRangesFormatBorder object for horizontal borders on the top edge.}
    property Top: TvgrRangesFormatBorder read FTop;
{Returns the TvgrRangesFormatBorder object for horizontal borders within area.}
    property Middle: TvgrRangesFormatBorder read FMiddle;
{Returns the TvgrRangesFormatBorder object for horizontal borders on the bottom edge.}
    property Bottom: TvgrRangesFormatBorder read FBottom;
  end;

{Describes the properties which have the equal values for each range.
Items:
  vgrrfFontName - All ranges have equal FontName.
  vgrrfFontSize - All ranges have equal FontSize.
  vgrrfFontStyleBold - All ranges have bold font.
  vgrrfFontStyleItalic - All ranges have italic font.
  vgrrfFontStyleUnderline - All ranges have underline font.
  vgrrfFontStyleStrikeout - All ranges have strikeout font.
  vgrrfFontColor - All ranges have equal FontColor.
  vgrrfFillBackColor - All ranges have equal FillBackColor.
  vgrrfFillForeColor - All ranges have equal FillForeColor.
  vgrrfFillPattern - All ranges have equal FillPattern.
  vgrrfDisplayFormat - All ranges have equal DisplayFormat.
  vgrrfHorzAlign - All ranges have equal HorzAlign.
  vgrrfVertAlign - All ranges have equal VertAlign.
  vgrrfAngle - All ranges have equal Angle.
  vgrrfWordWrap - All ranges have equal WordWrap.
  vgrrfFontCharset - All ranges have equal FontCharset.
Syntax:
  TvgrRangesFlag = (
    vgrrfFontName,
    vgrrfFontSize,
    vgrrfFontStyleBold,
    vgrrfFontStyleItalic,
    vgrrfFontStyleUnderline,
    vgrrfFontStyleStrikeout,
    vgrrfFontColor,
    vgrrfFillBackColor,
    vgrrfFillForeColor,
    vgrrfFillPattern,
    vgrrfDisplayFormat,
    vgrrfHorzAlign,
    vgrrfVertAlign,
    vgrrfAngle,
    vgrrfWordWrap,
    vgrrfFontCharset);
}
  TvgrRangesFlag = (
    vgrrfFontName,
    vgrrfFontSize,
    vgrrfFontStyleBold,
    vgrrfFontStyleItalic,
    vgrrfFontStyleUnderline,
    vgrrfFontStyleStrikeout,
    vgrrfFontColor,
    vgrrfFillBackColor,
    vgrrfFillForeColor,
    vgrrfFillPattern,
    vgrrfDisplayFormat,
    vgrrfHorzAlign,
    vgrrfVertAlign,
    vgrrfAngle,
    vgrrfWordWrap,
    vgrrfFontCharset);
{Describes the set of the properties which have the equal values for each range.}
  TvgrRangesFlags = set of TvgrRangesFlag;

  /////////////////////////////////////////////////
  //
  // IvgrRangesFormat
  //
  /////////////////////////////////////////////////
{Interface for formatting the area of the worksheet.
With using of this interface you can:<br>
  - Change the visual properties of the ranges and create the new ranges with specified properties.<br>
  - Change the visual properties of the borders and create the new borders with specified properties. <br>
See also:
  TvgrWorksheet.RangesFormat}
  IvgrRangesFormat = interface
  ['{E13B6B6B-D5E0-4017-A8A4-59CABD58F6CB}']
{Returns the value of the Borders property.
See also:
  Borders}
    function GetBorders: TvgrRangesFormatBorders;
{Returns the value of the HasFlags property.
See also:
  HasFlags}
    function GetHasFlags: TvgrRangesFlags;
{Returns the value of the FontName property.
See also:
  FontName}
    function GetFontName: TFontName;
{Returns the value of the FontSize property.
See also:
  FontSize}
    function GetFontSize: Integer;
{Returns the value of the FontStyle property.
See also:
  FontStyle}
    function GetFontStyle: TFontStyles;
{Returns the value of the FontStyleBold property.
See also:
  FontStyleBold}
    function GetFontStyleBold: Boolean;
{Returns the value of the FontStyleItalic property.
See also:
  FontStyleItalic}
    function GetFontStyleItalic: Boolean;
{Returns the value of the FontStyleUnderline property.
See also:
  FontStyleUnderline}
    function GetFontStyleUnderline: Boolean;
{Returns the value of the FontStyleStrikeout property.
See also:
  FontStyleStrikeout}
    function GetFontStyleStrikeout: Boolean;
{Returns the value of the FontColor property.
See also:
  FontColor}
    function GetFontColor: TColor;
{Returns the value of the FontCharset property.
See also:
  FontCharset}
    function GetFontCharset: TFontCharset;
{Returns the value of the FillBackColor property.
See also:
  FillBackColor}
    function GetFillBackColor: TColor;
{Returns the value of the FillForeColor property.
See also:
  FillForeColor}
    function GetFillForeColor: TColor;
{Returns the value of the FillPattern property.
See also:
  FillPattern}
    function GetFillPattern: TBrushStyle;
{Returns the value of the DisplayFormat property.
See also:
  DisplayFormat}
    function GetDisplayFormat: string;
{Returns the value of the HorzAlign property.
See also:
  HorzAlign}
    function GetHorzAlign: TvgrRangeHorzAlign;
{Returns the value of the VertAlign property.
See also:
  VertAlign}
    function GetVertAlign: TvgrRangeVertAlign;
{Returns the value of the Angle property.
See also:
  Angle}
    function GetAngle: Word;
{Returns the value of the WordWrap property.
See also:
  WordWrap}
    function GetWordWrap: Boolean;

{Sets the value of the FontName property.
See also:
  FontName}
    procedure SetFontName(Value: TFontName);
{Sets the value of the FontSize property.
See also:
  FontSize}
    procedure SetFontSize(Value: Integer);
{Sets the value of the FontStyle property.
See also:
  FontStyle}
    procedure SetFontStyle(Value: TFontStyles);
{Sets the value of the FontStyleBold property.
See also:
  FontStyleBold}
    procedure SetFontStyleBold(Value: Boolean);
{Sets the value of the FontStyleItalic property.
See also:
  FontStyleItalic}
    procedure SetFontStyleItalic(Value: Boolean);
{Sets the value of the FontStyleUnderline property.
See also:
  FontStyleUnderline}
    procedure SetFontStyleUnderline(Value: Boolean);
{Sets the value of the FontStyleStrikeout property.
See also:
  FontStyleStrikeout}
    procedure SetFontStyleStrikeout(Value: Boolean);
{Sets the value of the FontColor property.
See also:
  FontColor}
    procedure SetFontColor(Value: TColor);
{Sets the value of the FontCharset property.
See also:
  FontCharset}
    procedure SetFontCharset(Value: TFontCharSet);
{Sets the value of the FillBackColor property.
See also:
  FillBackColor}
    procedure SetFillBackColor(Value: TColor);
{Sets the value of the FillForeColor property.
See also:
  FillForeColor}
    procedure SetFillForeColor(Value: TColor);
{Sets the value of the FillPattern property.
See also:
  FillPattern}
    procedure SetFillPattern(Value: TBrushStyle);
{Sets the value of the DisplayFormat property.
See also:
  DisplayFormat}
    procedure SetDisplayFormat(Value: string);
{Sets the value of the HorzAlign property.
See also:
  HorzAlign}
    procedure SetHorzAlign(Value: TvgrRangeHorzAlign);
{Sets the value of the VertAlign property.
See also:
  VertAlign}
    procedure SetVertAlign(Value: TvgrRangeVertAlign);
{Sets the value of the Angle property.
See also:
  Angle}
    procedure SetAngle(Value: Word);
{Sets the value of the WordWrap property.
See also:
  WordWrap}
    procedure SetWordWrap(Value: Boolean);
{Returns all visual properties of ranges.
See also:
  TvgrRangesFlags, TvgrRangeHorzAlign, TvgrRangeVertAlign}
    procedure GetFormat(
      out AFontName: TFontName;
      out AFontSize: Integer;
      out AFontStyle: TFontStyles;
      out AFontCharset: TFontCharSet;
      out AFontColor, AFillBackColor, AFillForeColor: TColor;
      out AFillPattern: TBrushStyle;
      out ADisplayFormat: string;
      out AHorzAlign: TvgrRangeHorzAlign;
      out AVertAlign: TvgrRangeVertAlign;
      out AAngle: Word;
      out AWordWrap: Boolean;
      out AHasFlags: TvgrRangesFlags);
{Sets all visual properties of ranges.}
    procedure SetFormat(
      AFontName: TFontName;
      AFontSize: Integer;
      AFontStyle: TFontStyles;
      AFontCharset: TFontCharSet;
      AFontColor, AFillBackColor, AFillForeColor: TColor;
      AFillPattern: TBrushStyle;
      ADisplayFormat: string;
      AHorzAlign: TvgrRangeHorzAlign;
      AVertAlign: TvgrRangeVertAlign;
      AAngle: Word;
      AWordWrap: Boolean);
{Saves the formatting information to the stream.
Parameters:
  AStream - The destenation TStream object.}
    procedure SaveToStream(AStream: TStream);
{Loads the formatting information from the stream and apply it to the worksheet area.
Parameters:
  AStream - The source TStream object.}
    procedure LoadFromStream(AStream: TStream);
{Returns the value of the ForceBeginUpdateEndUpdate property.}
    function GetForceBeginUpdateEndUpdate: Boolean;
{Sets the value of the ForceBeginUpdateEndUpdate property.}
    procedure SetForceBeginUpdateEndUpdate(Value: Boolean);

{Specifies the font name for all ranges within area.
This property returns the empty string if not all ranges have equal FontName property.}
    property FontName: TFontName read GetFontName write SetFontName;
{Specifies the font size for all ranges within area.}
    property FontSize: Integer read GetFontSize write SetFontSize;
{Specifies the font style for all ranges within area.}
    property FontStyle: TFontStyles read GetFontStyle write SetFontStyle;
{Specifies the value indicating BOLD font for all ranges within area.
This property return the true value if all ranges have the bold font.}
    property FontStyleBold: Boolean read GetFontStyleBold write SetFontStyleBold;
{Specifies the value indicating ITALIC font for all ranges within area.
This property return the true value if all ranges have the italic font.}
    property FontStyleItalic: Boolean read GetFontStyleItalic write SetFontStyleItalic;
{Specifies the value indicating UNDERLINE font for all ranges within area.
This property return the true value if all ranges have the underline font.}
    property FontStyleUnderline: Boolean read GetFontStyleUnderline write SetFontStyleUnderline;
{Specifies the value indicating STRIKEOUT font for all ranges within area.
This property return the true value if all ranges have the strikeout font.}
    property FontStyleStrikeout: Boolean read GetFontStyleStrikeout write SetFontStyleStrikeout;
{Specifies the font color for all ranges within area.}
    property FontColor: TColor read GetFontColor write SetFontColor;
{Specifies the font charset for all ranges within area.}
    property FontCharset: TFontCharset read GetFontCharset write SetFontCharset;
{Specifies the background color for all ranges within area.}
    property FillBackColor: TColor read GetFillBackColor write SetFillBackColor;
{Specifies the foreground fill color for all ranges within area.}
    property FillForeColor: TColor read GetFillForeColor write SetFillForeColor;
{Specifies the fill style for all ranges within area.}
    property FillPattern: TBrushStyle read GetFillPattern write SetFillPattern;
{Specifies the display format for all ranges within area.}
    property DisplayFormat: string read GetDisplayFormat write SetDisplayFormat;
{Specifies the horizontal text alignment for all ranges within area.}
    property HorzAlign: TvgrRangeHorzAlign read GetHorzAlign write SetHorzAlign;
{Specifies the vertical text alignment for all ranges within area.}
    property VertAlign: TvgrRangeVertAlign read GetVertAlign write SetVertAlign;
{Specifies the angle of text rotation for all ranges within area.}
    property Angle: Word read GetAngle write SetAngle;
{Specifies the wordwrap property for all ranges within area.}
    property WordWrap: Boolean read GetWordWrap write SetWordWrap;

{Returns the value indicating which properties is equal for all ranges.
See also:
  TvgrRangesFlags}
    property HasFlags: TvgrRangesFlags read GetHasFlags;

{Returns the TvgrRangesFormatBorders object which can be used for formatting the borders.
See also:
  TvgrRangesFormatBorders}
    property Borders: TvgrRangesFormatBorders read GetBorders;

{Specifies the value indicating whether the BeginUpdate / EndUpdate methods is called
before the changes in workbook.
Set this property to true if you want to increase speed of work when you change
the large area of worksheet.}
    property ForceBeginUpdateEndUpdate: Boolean read GetForceBeginUpdateEndUpdate write SetForceBeginUpdateEndUpdate; 
  end;

{For internal use.}
  TvgrRangesFormatOperation = (vgrroGet,
    vgrroSetFontName,
    vgrroSetFontSize,
    vgrroSetFontStyleBold,
    vgrroSetFontStyleItalic,
    vgrroSetFontStyleUnderline,
    vgrroSetFontStyleStrikeout,
    vgrroSetFontColor,
    vgrroSetFillBackColor,
    vgrroSetFillForeColor,
    vgrroSetFillPattern,
    vgrroSetDisplayFormat,
    vgrroSetHorzAlign,
    vgrroSetVertAlign,
    vgrroSetAngle,
    vgrroSetWordWrap,
    vgrroSetFontCharset);

{For internal use.}
  TvgrRangesFormatOperations = set of TvgrRangesFormatOperation;

{The temporary object, for internal use.}
  TvgrRangesFormatTmpObject = class(TObject)
  public
    Count: Integer;
    Operations: TvgrRangesFormatOperations;
    FontName: TFontName;
    FontSize: Integer;
    FontStyle: TFontStyles;
    FontCharset: TFontCharset;
    FontColor: TColor;
    FillBackColor: TColor;
    FillForeColor: TColor;
    FillPattern: TBrushStyle;
    DisplayFormat: string;
    HorzAlign: TvgrRangeHorzAlign;
    VertAlign: TvgrRangeVertAlign;
    Angle: Word;
    WordWrap: Boolean;
    HasFlags: TvgrRangesFlags;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRangesFormat
  //
  /////////////////////////////////////////////////
{TvgrRangesFormat implements the IvgrRangesFormat interface.}
  TvgrRangesFormat = class(TInterfacedObject, IvgrRangesFormat)
  private
    FForceBeginUpdateEndUpdate: Boolean;
    FRectArray: TvgrRectArray;
    FBorders: TvgrRangesFormatBorders;
    FWorksheet: TvgrWorksheet;
    FAutoCreate: Boolean;
    FTempObject: TvgrRangesFormatTmpObject;
    function GetBorders: TvgrRangesFormatBorders;
    procedure InitTempObject;
    procedure Execute;
    procedure ExecuteCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure ExecuteCallbackSet(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    function GetHasFlags: TvgrRangesFlags;
    function GetFontName: TFontName;
    function GetFontSize: Integer;
    function GetFontStyle: TFontStyles;
    function GetFontStyleBold: Boolean;
    function GetFontStyleItalic: Boolean;
    function GetFontStyleUnderline: Boolean;
    function GetFontStyleStrikeout: Boolean;
    function GetFontCharset: TFontCharset;
    function GetFontColor: TColor;
    function GetFillBackColor: TColor;
    function GetFillForeColor: TColor;
    function GetFillPattern: TBrushStyle;
    function GetDisplayFormat: string;
    function GetHorzAlign: TvgrRangeHorzAlign;
    function GetVertAlign: TvgrRangeVertAlign;
    function GetAngle: Word;
    function GetWordWrap: Boolean;

    procedure SetFontName(Value: TFontName);
    procedure SetFontSize(Value: Integer);
    procedure SetFontStyle(Value: TFontStyles);
    procedure SetFontStyleBold(Value: Boolean);
    procedure SetFontStyleItalic(Value: Boolean);
    procedure SetFontStyleUnderline(Value: Boolean);
    procedure SetFontStyleStrikeout(Value: Boolean);
    procedure SetFontCharset(Value: TFontCharset);
    procedure SetFontColor(Value: TColor);
    procedure SetFillBackColor(Value: TColor);
    procedure SetFillForeColor(Value: TColor);
    procedure SetFillPattern(Value: TBrushStyle);
    procedure SetDisplayFormat(Value: string);
    procedure SetHorzAlign(Value: TvgrRangeHorzAlign);
    procedure SetVertAlign(Value: TvgrRangeVertAlign);
    procedure SetAngle(Value: Word);
    procedure SetWordWrap(Value: Boolean);

    procedure SavePropertyToList(AList: TStrings; const AName: string; const AValue: Variant);
    function LoadPropertyFromList(AList: TStrings; const AName: string; var AValue: Variant): Boolean;
    procedure SaveToStream(AStream: TStream);
    procedure LoadFromStream(AStream: TStream);
    function GetForceBeginUpdateEndUpdate: Boolean;
    procedure SetForceBeginUpdateEndUpdate(Value: Boolean);
  protected
    procedure GetFormat(
      out AFontName: TFontName;
      out AFontSize: Integer;
      out AFontStyle: TFontStyles;
      out AFontCharset: TFontcharset;
      out AFontColor, AFillBackColor, AFillForeColor: TColor;
      out AFillPattern: TBrushStyle;
      out ADisplayFormat: string;
      out AHorzAlign: TvgrRangeHorzAlign;
      out AVertAlign: TvgrRangeVertAlign;
      out AAngle: Word;
      out AWordWrap: Boolean;
      out AHasFlags: TvgrRangesFlags);
    procedure SetFormat(
      AFontName: TFontName;
      AFontSize: Integer;
      AFontStyle: TFontStyles;
      AFontCharset: TFontCharset;
      AFontColor, AFillBackColor, AFillForeColor: TColor;
      AFillPattern: TBrushStyle;
      ADisplayFormat: string;
      AHorzAlign: TvgrRangeHorzAlign;
      AVertAlign: TvgrRangeVertAlign;
      AAngle: Word;
      AWordWrap: Boolean);

    { Font name of the ranges }
    property FontName: TFontName read GetFontName write SetFontName;
    { Font size of the ranges }
    property FontSize: Integer read GetFontSize write SetFontSize;
    property FontStyleBold: Boolean read GetFontStyleBold write SetFontStyleBold;
    property FontStyle: TFontStyles read GetFontStyle write SetFontStyle;
    property FontStyleItalic: Boolean read GetFontStyleItalic write SetFontStyleItalic;
    property FontStyleUnderline: Boolean read GetFontStyleUnderline write SetFontStyleUnderline;
    property FontStyleStrikeout: Boolean read GetFontStyleStrikeout write SetFontStyleStrikeout;
    { Font charset of the ranges }
    property FontCharset: TFontCharset read GetFontCharset write SetFontCharset;
    { Font color of the ranges }
    property FontColor: TColor read GetFontColor write SetFontColor;
    { Fill background of the ranges }
    property FillBackColor: TColor read GetFillBackColor write SetFillBackColor;
    { Fill foreground of the ranges }
    property FillForeColor: TColor read GetFillForeColor write SetFillForeColor;
    { Fill pattern of the ranges }
    property FillPattern: TBrushStyle read GetFillPattern write SetFillPattern;
    { Display format of the ranges }
    property DisplayFormat: string read GetDisplayFormat write SetDisplayFormat;
    { Horizontal alignment of the ranges }
    property HorzAlign: TvgrRangeHorzAlign read GetHorzAlign write SetHorzAlign;
    { Vertical alignment of the ranges }
    property VertAlign: TvgrRangeVertAlign read GetVertAlign write SetVertAlign;
    { Text angle of the ranges }
    property Angle: Word read GetAngle write SetAngle;
    { Wordwrap property of the ranges }
    property WordWrap: Boolean read GetWordWrap write SetWordWrap;

    property HasFlags: TvgrRangesFlags read GetHasFlags;
  public
{Creates an instance of the TvgrRangesFormat class.
Please do not create instances of this class, it will be created automatically
when you use the TvgrWorksheet.RangesFormat property.
Parameters:
  ARects - The worksheet's area which will be formatted by this object.
  AWorksheet - The TvgrWorksheet object.
  AAutoCreate - If this parameter has the true value ranges and borders
will be automatically created when needed. If false - only existing objects will be edited.}
    constructor Create(ARects: TvgrRectArray; AWorksheet: TvgrWorksheet; AAutoCreate: Boolean);
{Frees an instance of the TvgrRangesFormat class.}
    destructor Destroy; override;
  end;

{Callback procedure which is called when the row height or column width is calculated,
this procedure must calculate the OPTIMAL size of range in twips. 
Parameters:
  ARange - The interface to the worksheet's range for which the OPTIMAL size in twips must be calculated.
Return value:
  The OPTIMAL size of the range in twips.}
  TvgrRangeCalcSizeProc = function(ARange: IvgrRange): TSize of object;
{Callback procedure which is called when the merging of the non empty ranges occurs.
Return value:
  Must return true if the merging is allowed.}
  TvgrMergeWarningProc = function: Boolean of object;

{Describes the borders' types within the rectangular area.
Items:
  vgrsbLeft - The vertical borders on the left edge.
  vgrsbRight - The vertical borders on the right edge.
  vgrsbTop - The horizontal borders on the top edge.
  vgrsbBottom - The horizontal borders on the bottom edge.
  vgrsbInside - The vertical and horizontal borders within area.
Syntax:
  TvgrSettedBorder = (vgrsbLeft, vgrsbRight, vgrsbTop, vgrsbBottom, vgrsbInside)
}
  TvgrSettedBorder = (vgrsbLeft, vgrsbRight, vgrsbTop, vgrsbBottom, vgrsbInside);
{Describes the set of borders' types within the rectangular area.
Syntax:
  TvgrSettedBorders = set of TvgrSettedBorder
See also:
  TvgrSettedBorder}
  TvgrSettedBorders = set of TvgrSettedBorder;

{TvgrWorksheetClass defines the metaclass for TvgrWorksheet.}
  TvgrWorksheetClass = class of TvgrWorksheet;

{Describes a part of the worksheet's contents to delete.
Items:
  vgrccDisplayFormat - The ranges' display format.
  vgrccValue - The ranges' values.
  vgrccFormat - The ranges' format.
  vgrccBorders - The borders.
  vgrccRanges - The whole ranges will be deleted.
Syntax:
  TvgrClearContentFlag = (vgrccDisplayFormat, vgrccValue, vgrccFormat, vgrccBorders, vgrccRanges)}
  TvgrClearContentFlag = (vgrccDisplayFormat, vgrccValue, vgrccFormat, vgrccBorders, vgrccRanges);
{Describes a part of the worksheet's contents to delete.
Syntax:
  TvgrClearContentFlags = set of TvgrClearContentFlag
See also:
  TvgrClearContentFlag}
  TvgrClearContentFlags = set of TvgrClearContentFlag;

{Describes a part of the worksheet's contents to paste from cllipboard or copy to clipboard.
Items:
  vgrptRangeValue - The ranges' values.
  vgrptRangeStyle - The ranges' visual styles.
  vgrptBorders - The borders.
  vgrptCols - The columns.
  vgrptRows - The rows.
Syntax:
  TvgrPasteType = (vgrptRangeValue, vgrptRangeStyle, vgrptBorders, vgrptCols, vgrptRows)}
  TvgrPasteType = (vgrptRangeValue, vgrptRangeStyle, vgrptBorders, vgrptCols, vgrptRows);
{Describes a part of the worksheet's contents to paste from cllipboard or copy to clipboard.
Syntax:
  TvgrPasteTypeSet = set of TvgrPasteType
See also:
  TvgrPasteType}
  TvgrPasteTypeSet = set of TvgrPasteType;

  /////////////////////////////////////////////////
  //
  // TvgrWorksheet
  //
  /////////////////////////////////////////////////
{Represents the separate sheet of workbook.
Worksheet contains properties and methods
for working with ranges, borders, sections, columns and rows.
See also:
  TvgrRanges, TvgrBorders, TvgrCols, TvgrRows, TvgrSections}
  TvgrWorksheet = class(TvgrComponent)
  private
    FPageProperties: TvgrPageProperties;
    FWorksheets : TvgrWorksheets;
    FRanges : TvgrRanges;
    FCols : TvgrCols;
    FRows : TvgrRows;
    FBorders : TvgrBorders;
    FHorzSections: TvgrSections;
    FVertSections: TvgrSections;
    FTitle : string;
    FMergeList: TInterfaceList;
    FCalculateCanvas: TCanvas;

    FExportData: array of Pointer;
    FDimensions: TRect;
    FDimensionsValid: Boolean;
    FDimensionsEmpty: Boolean;

    procedure ValidateDimensions;
    procedure CheckMinDimensions;
    procedure RangeCheckDimensions(const APlace: TRect);
    procedure BorderCheckDimensions(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation);
    procedure NewRowCheckDimensions(AIndexBefore, ACount: Integer);
    procedure NewColCheckDimensions(AIndexBefore, ACount: Integer);
    procedure DelColsCheckDimensions(AStartPos, AEndPos: Integer);
    procedure DelRowsCheckDimensions(AStartPos, AEndPos: Integer);
    procedure RangeOrBorderDeleted;

    procedure ReadHorzSectionsData(Stream: TStream);
    procedure WriteHorzSectionsData(Stream: TStream);
    procedure ReadVertSectionsData(Stream: TStream);
    procedure WriteVertSectionsData(Stream: TStream);
    procedure ReadColsData(Stream: TStream);
    procedure WriteColsData(Stream: TStream);
    procedure ReadRowsData(Stream: TStream);
    procedure WriteRowsData(Stream: TStream);
    procedure ReadBordersData(Stream: TStream);
    procedure WriteBordersData(Stream: TStream);
    procedure ReadRangesData(Stream: TStream);
    procedure WriteRangesData(Stream: TStream);
    function GetWorkbook: TvgrWorkbook;
    procedure SetTitle(const Value : string);
    function GetCol(ColNumber : integer) : IvgrCol;
    function GetColByIndex(Index : integer) : IvgrCol;
    function GetColsCount : integer;
    function GetRow(RowNumber : integer) : IvgrRow;
    function GetRowByIndex(Index : integer) : IvgrRow;
    function GetRowsCount : integer;
    function GetRange(Left,Top,Right,Bottom : integer) : IvgrRange;
    function GetRangeByIndex(Index : integer) : IvgrRange;
    function GetRangesCount : integer;
    function GetBorder(Left,Top : integer; Orientation : TvgrBorderOrientation) : IvgrBorder;
    function GetBorderByIndex(Index : integer) : IvgrBorder;
    function GetBordersCount : integer;
    function GetHorzSection(StartPos, EndPos: Integer): IvgrSection;
    function GetHorzSectionByIndex(Index: Integer): IvgrSection;
    function GetHorzSectionCount: Integer;
    function GetVertSection(StartPos, EndPos: Integer): IvgrSection;
    function GetVertSectionByIndex(Index: Integer): IvgrSection;
    function GetVertSectionCount: Integer;
    procedure SetPageProperties(Value: TvgrPageProperties);

    function GetIndexInWorkbook: Integer;
    procedure SetIndexInWorkbook(Value: Integer);

    procedure MergeFindMergedCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure MergeFindAllCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure GetRowAutoHeightCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure GetColAutoWidthCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure ClearContentsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    procedure CopyToClipboardRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure CopyToClipboardBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure CopyToClipboardColsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure CopyToClipboardRowsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    procedure ChangeValueTypeOfRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    procedure PasteFromPoint(AClipboardObject: TObject; const APastePoint: TPoint; APasteType: TvgrPasteTypeSet);
    function ExportDataIsActive: Boolean;
    function GetCalculateCanvas: TCanvas;
    procedure GetRangePixelsRect(const ARangeRect: TRect; var ARangePixelsRect: TRect);
    procedure GetRangeInternalPixelsRect(const ARangeRect: TRect; var ARangePixelsRect: TRect);
    function DefaultGetRangeSizeProc(ARange: IvgrRange): TSize;
  protected
    procedure BeforeChangeProperty;
    procedure AfterChangeProperty;
    procedure BeginUpdate;
    procedure EndUpdate;

    procedure SetName(const NewName: TComponentName); override;
    function GetDimensions : TRect; virtual;
    function GetSectionsClass: TvgrSectionsClass; virtual;
    function GetChildOwner: TComponent; override;

    procedure DefineProperties(Filer: TFiler); override;

    function GetDispIdOfName(const AName: string): Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    procedure OnPagePropertiesChange(Sender: TObject);

    procedure PrepareExportData;
    procedure ClearExportData;

    property CalculateCanvas: TCanvas read GetCalculateCanvas;
  public
{Creates an instance of the TvgrWorksheet class.
Parameters:
  AOwner - The component - owner.}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of the TvgrWorksheet class.}
    destructor Destroy; override;
{Returns the associated workbook.}
    function GetParentComponent: TComponent; override;
{Always returns true.}
    function HasParent: Boolean; override;
{If Value is TvgrWorkbook, sets Value as Parent component.}
    procedure SetParentComponent(Value: TComponent); override;

{Creates the borders or changes the properties of existing borders in the specified area.
Parameters:
  ARect - Specifies the area, in which borders is created.
  ASettedBorders - The parameter determining what borders in relation to a cell will be created (left, right, top, bottom).
  ABorderStyle - The style of the created borders.
  ABorderWidth - The width of the created borders in twips.
  ABorderColor - The color of the created borders.
Example:
  Worksheet.SetBorders(Rect(0,0,10,10), [vgrsbLeft], vgrbsSolid, 70, clRed);
See also:
  TvgrSettedBorders}
    procedure SetBorders(const ARect: TRect; ASettedBorders: TvgrSettedBorders; ABorderStyle: TvgrBorderStyle; ABorderWidth: Integer; ABorderColor: TColor);

{Inserts the rows in the specified position.
Parameters:
  AIndexBefore - The number of row before which the new rows will be inserted.
  ACount - The amount of the inserted rows.
See also:
  InsertCols}
    procedure InsertRows(AIndexBefore, ACount: Integer);
{Inserts the columns in the specified position.
Parameters:
  AIndexBefore - The number of clumn before which the new columns will be inserted.
  ACount - The amount of the inserted columns.
See also:
  InsertRows}
    procedure InsertCols(AIndexBefore, ACount: Integer);
{Deletes the rows in the specified range.
Parameters:
  AStartPos - The number of first deleted row.
  AEndPos - The number of last deleted row.
Example:
  Worksheet.DeleteRows(5, 10); // deletes rows: 5, 6, 7, 8, 9, 10}
    procedure DeleteRows(AStartPos, AEndPos: Integer);
{Deletes the columns in the specified range.
Parameters:
  AStartPos - The number of first deleted column.
  AEndPos - The number of last deleted column.
Example:
  Worksheet.DeleteCols(5, 10); // deletes columns: 5, 6, 7, 8, 9, 10 }
    procedure DeleteCols(AStartPos, AEndPos: Integer);
{Checks an opportunity of the cells' deleting.
Function may return false, if it will be found a range, covering area which cannot be correctly
processed at the set parameters of operation.
Parameters:
  ARect - The area of the cells.
  ACellShift - The direction of shift of the next cells after deleting.
  ABreakRanges - Shows, that if necessary it is possible to break big ranges.
Return value:
  True - if operation with the set conditions is possible, false - if not. }
    function PerformDeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean): Boolean;
{Deletes the cells in the specified rect.
Parameters:
  ARect - The area of the cells.
  ACellShift - The direction of shift of the next cells after deleting.
  ABreakRanges - Shows, that if necessary it is possible to break big ranges.
Return value:
  True - success, false - if not. }
    function DeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean): TvgrCellShiftState;
{Returns the IvgrRangesFormat interface for formatting the specified area.
Parameters:
  ARects - Defines the formatting area.
  AAutoCreate - If this parameter has the true value ranges and borders
will be automatically created. If false - only existing objects will be edited.
See also:
  IvgrRangesFormat}
    function GetRangesFormat(ARects: TvgrRectArray; AAutoCreate: Boolean): IvgrRangesFormat;

{Merges all ranges in the specified rectangle and creates the one range.
Parameters:
  AMergeRect - Specifies the area for merging.
  AWarningProc - This procedure is called when in specified area not one range
is found which has the non empty value. This parameter may be nil, in this case
the superfluous ranges will be deleted.
Example:

  function TMyForm.MergeWarningProc: Boolean;
  begin
    Result := MBox('The slection contains multiple data values. Merging into one cell will keep the#13upper-left most data only',
                   MB_OKCANCEL or MB_ICONEXCLAMATION) = IDOK;
  end;

  procedure TMyForm.Merge(AWorksheet: TvgrWorksheet; const ARect: TRect);
  begin
    AWorksheet.Merge(ARect, MergeWarningProc)
  end;
See also:
  UnMerge, MergeUnmerge, TvgrMergeWarningProc}
    procedure Merge(const AMergeRect: TRect; AWarningProc: TvgrMergeWarningProc = nil);
{Unmerges the all ranges in the specified rectangle and creates the some ranges.
Parameters:
  AMergeRect - Specifies the area for unmerging.
See also:
  Merge, MergeUnmerge}
    procedure UnMerge(const AMergeRect: TRect);
{If the specified area contains the merged ranges (more than one cell) - unmerges them,
else - merges the ranges within specified rectangle.
Parameters:
  AMergeRect - Specifies the area for merging.
  AWarningProc - This procedure is called when in specified area not one range
is found which has the non empty value. This parameter may be nil, in this case
the superfluous ranges will be deleted.
See also:
  Merge, UnMerge, IsMergeInRect}
    procedure MergeUnMerge(const AMergeRect: TRect; AWarningProc: TvgrMergeWarningProc = nil);
{Tests the specified area and returns the true value if within the area
merged ranges exist.
Parameters:
  ARect - Defines area to test.
See also:
  Merge, UnMerge, MergeUnmerge}
    function IsMergeInRect(const ARect: TRect): Boolean;

{Returns the OPTIMUM height for specified row (in twips).
Parameters:
  ARowNumber - The number of row.
  ACalcRangeSizeProc - The callback procedure which must calculate size of the separate range.
Return value:
  The OPTIMUM row's height in twips.
See also:
  TvgrRangeCalcSizeProc}
    function GetRowAutoHeight(ARowNumber: Integer; ACalcRangeSizeProc: TvgrRangeCalcSizeProc): Integer; overload;
{Returns the OPTIMUM height for specified row (in twips).
The screen dc is used to calculate the sizes of ranges.
Parameters:
  ARowNumber - The number of row.
Return value:
  The OPTIMUM row's height in twips.}
    function GetRowAutoHeight(ARowNumber: Integer): Integer; overload;
{Returns the OPTIMUM width for specified column (in twips).
Parameters:
  AColNumber - The number of column.
  ACalcRangeSizeProc - The callback procedure which must calculate size of the separate range.
Return value:
  The OPTIMUM column's width in twips.
See also:
  TvgrRangeCalcSizeProc}
    function GetColAutoWidth(AColNumber: Integer; ACalcRangeSizeProc: TvgrRangeCalcSizeProc): Integer; overload;
{Returns the OPTIMUM width for specified column (in twips).
The screen dc is used to calculate the sizes of ranges.
Parameters:
  AColNumber - The number of column.
Return value:
  The OPTIMUM column's width in twips.}
    function GetColAutoWidth(AColNumber: Integer): Integer; overload;

{Cuts the worksheet's contents into clipboard.
Parameters:
  ARect - Specifies the area to cut.}
    procedure CutToClipboard(const ARects: TvgrRectArray);
{Copies the worksheet's contents into clipboard.
Parameters:
  ARect - Specifies the area to copy.}
    procedure CopyToClipboard(const ARect: TRect); overload;
{Copies the worksheet's contents into clipboard.
Parameters:
  ARects - Specifies the area to copy.
See also:
  TvgrRectArray}
    procedure CopyToClipboard(const ARects: TvgrRectArray); overload;
{Pastes the clipboard contents into the specified area of worksheet.
Parameters:
  ARect - Specifies the worksheet's area to paste.
  APasteType - Specifies the part of the contents to paste.
See also:
  TvgrPasteTypeSet}
    procedure PasteFromClipboard(const ARect: TRect; APasteType: TvgrPasteTypeSet);
{Clears the worksheet's area.
Parameters:
  ARect - Specifies the worksheet's area.
  AClearFlags - Specifies the part of worksheet's contents for clearing.
See also:
  TvgrClearContentFlags}
    procedure ClearContents(const ARect: TRect; AClearFlags: TvgrClearContentFlags); overload;
{Clears the worksheet's area.
Parameters:
  ARects - Specifies the worksheet's area.
  AClearFlags - Specifies the part of worksheet's contents for clearing.
See also:
  TvgrRectArray, TvgrClearContentFlags}
    procedure ClearContents(const ARects: TvgrRectArray; AClearFlags: TvgrClearContentFlags); overload;
{Deletes all borders within the specified area.
Parameters:
  ARect - Specifies the worksheet's area.}
    procedure ClearBorders(const ARect: TRect); overload;
{Deletes the specified borders within the specified area.
Parameters:
  ARect - Specifies the worksheet's area.
  ABorderTypes - Defines the borders' types.
See also:
  TvgrBorderTypes}
    procedure ClearBorders(const ARect: TRect; ABorderTypes: TvgrBorderTypes); overload;
{Deletes all ranges within the specified area.
Parameters:
  ARect - Specifies the worksheet's area.}
    procedure ClearRanges(const ARect: TRect);
{Clears the specified properties of the worksheet's ranges.
Parameters:
  ARect - Specifies the worksheet's area.
  AClearFlags - Specified the ranges' properties to clear. Items: vgrccBorders and vgrccRanges
are ignored by this procedure.
See also:
  TvgrClearContentFlags}
    procedure ClearRangesContents(const ARect: TRect; AClearFlags: TvgrClearContentFlags);

{Creates a copy of worksheet and inserts it into the specified position in the list of worksheets.
Parameters:
  AInsertIndex - Specifies the position of the created worksheet in the list of worksheets.
Return value:
  Returns the created TvgrWorksheet object.}
    function Copy(AInsertIndex: Integer): TvgrWorksheet;

{Changes the type of ranges' values, to the specified type,
if conversion can not be made - value is not changing.
Parameters:
  ARect - Defines the area of ranges.
  ANewValueType - Defines the new type of ranges.
See also:
  TvgrRangeValueType}
    procedure ChangeValueTypeOfRanges(const ARect: TRect; ANewValueType: TvgrRangeValueType); overload;
{Changes the type of ranges' values, to the specified type,
if conversion can not be made - value is not changing.
Parameters:
  ACellsRects - Defines the area of ranges.
  ANewValueType - Defines the new type of ranges.
See also:
  TvgrRectArray, TvgrRangeValueType}
    procedure ChangeValueTypeOfRanges(const ACellsRects: TvgrRectArray; ANewValueType: TvgrRangeValueType); overload;

{Returns the TvgrWorkbook object containing this worksheet.
See also:
  TvgrWorkbook}
    property Workbook: TvgrWorkbook read GetWorkbook;
{Returns the list of worksheet's columns.
See also:
  IvgrCol, TvgrCols}
    property ColsList: TvgrCols read FCols;
{Searches the column with specified number and returns interface to it,
if the column is not found then creates it and returns the interface to it.
Parameters:
  Number - The number of column.
Example:
  var
    I: Integer;
  begin
    // Create the 11 columns and set it width to 1 inch. 
    for I := 0 to 10 do
      with Worksheet.Cols[I] do
        Width := ConvertUnitsToTwips(1, vgruInches);
  end;
See also:
  TvgrCols.Items, IvgrCol}
    property Cols[ColNumber: Integer]: IvgrCol read GetCol;
{Returns the column's interface by the column's index in the list.
Parameters:
  Index - The index of the column, must be from 0 to ColsCount - 1.
See also:
  TvgrCols.ByIndex, IvgrCol}
    property ColByIndex[Index: Integer]: IvgrCol read GetColByIndex;
{Returns the amount of columns.}
    property ColsCount: Integer read GetColsCount;
{Returns the list of worksheet's rows.
See also:
  IvgrRow, TvgrRows}
    property RowsList: TvgrRows read FRows;
{Searches the row with specified number and returns interface to it,
if the row is not found then creates it and returns its interface.
Parameters:
  Number - The number of row.
Example:
  var
    I: Integer;
  begin
    for I := 0 to 10 do
      with Worksheet.Rows[I] do
        Height := ConvertUnitsToTwips(1, vgruInches);
  end;
See also:
  IvgrRow, TvgrRows, TvgrRows.Items}
    property Rows[RowNumber: Integer]: IvgrRow read GetRow;
{Returns the row's interface by the row's index in the list.
Parameters:
  Index - The index of the row, must be from 0 to RowsCount - 1}
    property RowByIndex[Index: Integer]: IvgrRow read GetRowByIndex;
{Returns the amount of rows.}
    property RowsCount: Integer read GetRowsCount;
{Returns the list of worksheet's ranges.
See also:
  IvgrRange, TvgrRanges}
    property RangesList : TvgrRanges read FRanges;
{Searches the range with specified area and returns interface to it,
if the range is not found then creates it and returns its interface.
Parameters:
  Left - The X coordinate of the top-left range corner.
  Top - The Y coordinate of the top-left range corner.
  Right - The X coordinate of the bottom-right range corner.
  Bottom - The Y coordinate of the bottom-right range corner.
Example:
  var
    I: Integer;
  begin
    // create the rectangle of cells
    for I := 0 to 10 do
      for J := 0 to 10 do
        with Worksheet.Ranges[I, J, I, J] do
          Value := Format('Cell %d, %d', [I, J]);
  end;
See also:
  IvgrRange, TvgrRanges, TvgrRanges.Items}
    property Ranges[Left, Top, Right, Bottom: Integer]: IvgrRange read GetRange;
{Returns the range's interface by the range's index in the RangesList list.
Parameters:
  Index - The index of the range, must be from 0 to RangesCount - 1
Example:
  var
    I: Integer;
  begin
    // set font to bold for all ranges on worksheet
    for I := 0 to AWorksheet.RangesList.Count - 1 do
      with AWorksheet.RangesList[I] do
        FontStyle := FontStyle + [fsBold];
  end;
See also:
  IvgrRange, TvgrRanges, TvgrRanges.ByIndex}
    property RangeByIndex[Index: Integer]: IvgrRange read GetRangeByIndex;
{Returns the amount of ranges.}
    property RangesCount : integer read GetRangesCount;
{Returns the list of worksheet's borders.
See also:
  IvgrBorder, TvgrBorders}
    property BordersList : TvgrBorders read FBorders;
{Searches the border with specified coordinates and returns interface to it,
if the border is not found then creates it and returns its interface.
Parameters:
  Left - The X coordinate of the top-left border corner.
  Top - The Y coordinate of the top-left border corner.
  Orientation - The border type - vertical or horizontal.
Example:
  var
    I: Integer;
  begin
    // create the line of borders
    for I := 0 to 10 do
      with Worksheet.Borders[1, I, vgrboLeft] do
      begin
        Width := ConvertUnitsToTwips(1, vgruMms);
        if (I mod 2) = 0 then
          Color := clRed
        else
          Color := clBlue;
      end;
  end;
See also:
  IvgrBorder, TvgrBorders, TvgrBorders.Items}
    property Borders[Left, Top: Integer; Orientation: TvgrBorderOrientation]: IvgrBorder read GetBorder;
{Returns the border's interface by the border's index in the BordersList list.
Parameters:
  Index - The index of the border, must be from 0 to BordersCount - 1
Example:
  var
    I: Integer;
  begin
    // make all borders green
    for I := 0 to AWorksheet.BordersList.Count - 1 do
      with AWorksheet.BordersList[I] do
        Color := clGreen;
  end;}
    property BorderByIndex[Index : integer] : IvgrBorder read GetBorderByIndex;
{Returns the amount of borders.}
    property BordersCount : integer read GetBordersCount;
{Searches the horizontal section with specified coordinates and returns interface to it,
if the section is not found then creates it and returns its interface.
Parameters:
  StartPos - The starting row of section.
  EndPos - The ending row of section.
Example:
  begin
    // create the horizontal page header
    with Worksheet.HorzSections[0, 2] do
    begin
      RepeatOnPageTop := True;
    end;

    // create the horizontal page footer
    with Worksheet.HorzSections[100, 100] do
    begin
      RepeatOnPageBottom := True;
    end;
  end;
See also:
  IvgrSection, TvgrSections, TvgrSections.Items}
    property HorzSections[StartPos, EndPos: Integer]: IvgrSection read GetHorzSection;
{Returns the TvgrSections object containing the horizontal sections.
See also:
  IvgrSection, TvgrSections}
    property HorzSectionsList: TvgrSections read FHorzSections;
{Returns the section's interface by the section's index in the HorzSectionsList list.
Parameters:
  Index - The index of the section, must be from 0 to HorzSectionCount - 1.}
    property HorzSectionByIndex[Index: Integer]: IvgrSection read GetHorzSectionByIndex;
{Returns the amount of the horizontal sections.}
    property HorzSectionCount: Integer read GetHorzSectionCount;
{Searches the vertical section with specified coordinates and returns interface to it,
if the section is not found then creates it and returns its interface.
Parameters:
  StartPos - The starting row of section.
  EndPos - The ending row of section.
Example:
  begin
    // create the vertical page header
    with Worksheet.VertSections[0, 2] do
    begin
      RepeatOnPageTop := True;
    end;

    // create the vertical page footer
    with Worksheet.VertSections[100, 100] do
    begin
      RepeatOnPageBottom := True;
    end;
  end;
See also:
  IvgrSection, TvgrSections, TvgrSections.Items}
    property VertSections[StartPos, EndPos: Integer]: IvgrSection read GetVertSection;
{Returns the TvgrSections object containing the vertical sections.
See also:
  IvgrSection, TvgrSections}
    property VertSectionsList: TvgrSections read FVertSections;
{Returns the section's interface by the section's index in the VertSectionsList list.
Parameters:
  Index - The index of the section, must be from 0 to HorzSectionCount - 1.}
    property VertSectionByIndex[Index: Integer]: IvgrSection read GetVertSectionByIndex;
{Returns the amount of the vertical sections.}
    property VertSectionCount: Integer read GetVertSectionCount;

{Gets or sets the index of the worksheet within workbook.}
    property IndexInWorkbook: Integer read GetIndexInWorkbook write SetIndexInWorkbook;
{Returns the worksheet's area within which all ranges and borders of worksheet are placed.}
    property Dimensions: TRect read GetDimensions;
{Returns the IvgrRangesFormat interface for formatting the specified worksheet's area.
Parameters:
  ARects - Specifies the worksheet's area.
  AAutoCreate - If this parameter has the true value ranges and borders
will be automatically created when needed. If false - only existing objects will be edited.
Example:

  procedure TForm1.CreateTable(AWorksheet: TvgrWorksheet; const ATableRect: TRect; AHeaderFontStyle: TFontStyles);
  var
    ARects: TvgrRectArray;
    I: Integer;
  begin
    SetLength(ARects, 1);
    with ATableRect do
      ARects[0] := Rect(Left, Top, Right, Top);
  
    // Creates the table header
    with AWorksheet.RangesFormat[ARects, True] do
    begin
      FillBackColor := clSilver;
      FontStyle := AHeaderFontStyle;
    end;
  
    // Creates the table body
    for I := ATableRect.Top + 1 to ATableRect.Bottom do
    begin
      ARects[0] := Rect(ATableRect.Left, I, ATableRect.Right, I);
      with AWorksheet.RangesFormat[ARects, True] do
        if I mod 2 = 0 then
          FillBackColor := $00D8FEFE
        else
          FillBackColor := $008BFCFC;
    end;
  
    // Creates the table borders
    ARects[0] := ATableRect;
    with AWorksheet.RangesFormat[ARects, True].Borders do
    begin
      Left.Width := ConvertPixelsToTwipsX(1);
      Center.Width := ConvertPixelsToTwipsX(1);
      Right.Width := ConvertPixelsToTwipsX(1);
      Top.Width := ConvertPixelsToTwipsY(1);
      Middle.Width := ConvertPixelsToTwipsY(1);
      bottom.Width := ConvertPixelsToTwipsY(1);
    end;
  end;
  
See also:
  IvgrRangesFormat}
    property RangesFormat[ARects: TvgrRectArray; AAutoCreate: Boolean]: IvgrRangesFormat read GetRangesFormat;
  published
{Gets or sets the worksheet's caption.}
    property Title: string read FTitle write SetTitle;
{Specifies the worksheet's page properties.
See also:
  TvgrPageProperties}
    property PageProperties: TvgrPageProperties read FPageProperties write SetPageProperties;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWorksheets
  //
  /////////////////////////////////////////////////
{Store and maintains a list of worksheets.
See also:
  TvgrWorksheet}
  TvgrWorksheets = class(TList)
  private
    FWorkbook : TvgrWorkbook;
    function GetItm(Index : integer) : TvgrWorksheet;
    procedure BeforeChange(ChangeInfo : TvgrWorkbookChangeInfo);
    procedure AfterChange(ChangeInfo : TvgrWorkbookChangeInfo);
    procedure AddInList(AWorksheet: TvgrWorksheet);
  public
{Creates an instance of the TvgrWorksheets class.
Parameters:
  AWorkbook - The TvgrWorkbook object containing this list.}
    constructor Create(AWorkbook : TvgrWorkbook);
{Frees an instance of the TvgrWorksheet class.}
    destructor Destroy; override;
{Creates a new Worksheet and adds it to list.
Return value:
  The created TvgrWorksheet object.}
    function Add: TvgrWorksheet;
{Creates a new Worksheet and inserts it to list at specified position
Parameters:
  Index - The position to inserting.
Return value:
  The created TvgrWorksheet object.}
    function Insert(Index: Integer): TvgrWorksheet;
{Removes the worksheet with specified Index from a list.
Parameters:
  Index - The index of the removed worksheet.}
    procedure Delete(Index: Integer); overload;
{Clears the list.}
    procedure Clear; reintroduce;
{Removes the specified TvgrWorksheet object from workbook.
{Parameters:
  AWorksheet - The TvgrWorksheet object to remove.}
    procedure RemoveWorksheet(AWorksheet: TvgrWorksheet);
{List the worksheets in the TvgrWorksheets.
Parameters:
  Index - The worksheet's index, from 0 to Count - 1.}
    property Items[Index: Integer]: TvgrWorksheet read GetItm;
{Returns the TvgrWorkbook object containing this list.}    
    property Workbook: TvgrWorkbook read FWorkbook;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormulasList
  //
  /////////////////////////////////////////////////
{Store and maintains a list of formulas.
TvgrFormulasList provides properties and methods to working with formulas within worksheets.
This is an internal class.}
  TvgrFormulasList = class(TObject, IUnknown, IvteFormulaCompilerOwner, IvgrFormulaCalculatorOwner)
  private
    FWorkbook: TvgrWorkbook;
    FWBStrings: TvgrWBStrings;
    FItems: TList;
    FCompiler: TvteExcelFormulaCompiler;
    FCalculator: TvgrFormulaCalculator;
    FCalculatorStack: TvgrCalculatorStack;
    function GetItem(Index: Integer): pvgrFormulaHeader;
    function GetCount: Integer;
    function GetFormula(Index: Integer): pvteFormula;
    function GetFormulaSize(Index: Integer): Integer;
    property Items[Index: Integer]: pvgrFormulaHeader read GetItem;
  protected
    // IUnknown
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
    // IvteFormulaCompilerOwner
    function GetStringIndex(const S: string): Integer;
    function GetString(AStringIndex: Integer): string;

    function GetExternalSheetIndex(const AWorkbookName, ASheetName: string): Integer;
    function GetExternalSheetName(const AWorkbookName: string; AIndex: Integer): string;

    function GetExternalWorkbookIndex(const AWorkbookName: string): Integer;
    function GetExternalWorkbookName(AIndex: Integer): string;

    function GetFunctionName(AId: Integer): string;
    function GetFunctionId(const AFunctionName: string): Integer;
    // IvgrFormulaCalculatorOwner
    procedure EnumRangeValuesCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    function RangeValueToFormulaValue(ACalculator: TvgrFormulaCalculator; ARange: IvgrRange; ARangeSheet: Integer; var AValue: rvteFormulaValue): Boolean;
    function GetCellValue(ACalculator: TvgrFormulaCalculator;
                          var AValue: rvteFormulaValue;
                          ASheet, ACol, ARow: Integer): Boolean;
    function EnumRangeValues(ACalculator: TvgrFormulaCalculator;
                             var AEnumRec: rvgrEnumStackRec;
                             ACallback: TvgrFormulaCalculatorEnumStackProc;
                             ASheet: Integer;
                             const ARangeRect: TRect): Boolean;

    procedure InternalDelete(AIndex: Integer);
    function InternalAdd(var AHeader: pvgrFormulaHeader; AFormulaSize: Integer): Integer;
    function InternalFindOrAdd(AFormula: pvteFormula; AFormulaSize: Integer): Integer;

    procedure SaveToStream(AStream: TStream); virtual;
    procedure LoadFromStreamConvertRecord(AOldData, ANewData: Pointer;
                                          AItemCount,
                                          AOldHeaderSize, AOldItemSize,
                                          ANewHeaderSize, ANewItemSize,
                                          ADataStorageVersion: Integer); virtual;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); virtual;

{Decrements the reference count for formula with given index.
All formulas in list has reference count, which indicate,
in how many places is used this formula.
Parameters:
  AIndex - index ot formula in list, which reference count is decremented.}
    procedure Release(AIndex: Integer);
{Incerments the reference count for formula with given index.
All formulas in list has reference count, which indicate,
in how many places is used this formula.
Parameters:
  AIndex - index ot formula in list, which reference count is incremented.}
    procedure AddRef(AIndex: Integer);
{Searches formula in a list, and if not found, add it to list.
Parameters:
  AFormula - pointer to array with formula items.
  AFormulaSize - count of items in formula
  AOldIndex - index of formula, which is replaced with new formula. -1 if none.
  ASourceWBStrings - The TvgrWBStrings object that is used specified in the AFormula parameter.
Return value:
  Index of specified formula in a list.  }
    function FindOrAdd(AFormula: pvteFormula; AFormulaSize, AOldIndex: Integer; ASourceWBStrings: TvgrWBStrings): Integer; overload;
{Searches formula in a list, and if not found, add it to list.
Parameters:
  AFormula - pointer to array with formula items.
  AFormulaSize - count of items in formula
  AOldIndex - index of formula, which is replaced with new formula. -1 if none.
Return value:
  Index of specified formula in a list.  }
    function FindOrAdd(AFormula: pvteFormula; AFormulaSize, AOldIndex: Integer): Integer; overload;
{Searches formula in a list, and if not found, add it to list.
Parameters:
  AFormulaText - string representation of formula.
  AOldIndex  - index of formula, which is replaced with new formula. -1 if none.
  AFormulaCol - cell column, in which contains formula
  AFormulaRow - cell row, in which contains formula
Return value:
  Index of specified formula in a list.  }
    function FindOrAdd(const AFormulaText: string; AOldIndex: Integer; AFormulaCol, AFormulaRow: Integer): Integer; overload;
{Returns formula by Index as a string.
Function carries out decompiling and returns string  representation of the formula
Parameters:
  Index - index of formula in list
  AFormulaCol - cell column, in which contains formula
  AFormulaRow - cell row, in which contains formula  }
    function GetFormulaText(Index: Integer; AFormulaCol, AFormulaRow: Integer): string;
{Recalculates all values of ranges, which contains formulas.}
    procedure Calculate;
{Deletes all formulas from the list.}
    procedure Clear;
{Sets or returns the associated workbook.}
    property Workbook: TvgrWorkbook read FWorkbook;
{Specifies the number of formulas in the list.}
    property Count: Integer read GetCount;
{Returns formula by given Index.
See also:
  pvteFormula }
    property Formulas[Index: Integer]: pvteFormula read GetFormula;
{ Returns number of items, contained in the formula with given Index }
    property FormulasSize[Index: Integer]: Integer read GetFormulaSize;
{Returns the TvgrWBStrings object that is used by this object, typically this is a Workbook.WBStrings.}
    property WBStrings: TvgrWBStrings read FWBStrings;
  public
{Creates an instance of the TvgrFormulasList class.
Parameters:
  AWorkbook - The TvgrWorkbook object containing this object.}
    constructor Create(AWorkbook: TvgrWorkbook);
{Frees an instance of the TvgrFormulasList class.}
    destructor Destroy; override;
{$IFDEF VGR_DEBUG}
    function DebugInfo: TvgrDebugInfo;
{$ENDIF}
  end;

{Is used for event which notify about deleted items.
Parameters:
  Sender - The TvgrWorkbook object firing an event.
  AItem - The IvgrWBListItem interface to the deleted item.}
  TvgrItemInterfaceDeleteNotify = procedure(Sender: TObject; AItem: IvgrWBListItem) of object;
  /////////////////////////////////////////////////
  //
  // TvgrWorkbook
  //
  /////////////////////////////////////////////////
{TvgrWorkbook is the most important component of GridReport library.
It represents the data storage with structure very similar to the book of MS Excel.
The TvgrWorkbook contains the worksheets which in turn contain cells,
borders, formulas and so on.
Visual components, such as TvgrWorkbookGrid, TvgrWorkbookPreview display the content
of the TvgrWorkbook component connected to them.
See also:
  TvgrWorksheet}
  TvgrWorkbook = class(TvgrComponent)
  private
    FDataStorageVersion: Integer;

    FWorkbookHandlers : TInterfaceList;
    FWorksheets : TvgrWorksheets;
    FRangeStyles: TvgrRangeStylesList;
    FBorderStyles: TvgrBorderStylesList;
    FWBStrings: TvgrWBStrings;
    FFormulas: TvgrFormulasList;
    FLockCount: Integer;
    FModified: Boolean;

    FOnItemInterfaceDelete: TvgrItemInterfaceDeleteNotify;
    FOnChanged: TNotifyEvent;

    function GetWorksheetsCount : integer;
    function GetWorksheet(Index : integer) : TvgrWorksheet;
    procedure SetModified(Value: Boolean);

    procedure ReadFormulasData(Stream: TStream);
    procedure WriteFormulasData(Stream: TStream);
    procedure ReadRangeStylesData(Stream: TStream);
    procedure WriteRangeStylesData(Stream: TStream);
    procedure ReadBorderStylesData(Stream: TStream);
    procedure WriteBorderStylesData(Stream: TStream);
    procedure ReadStringsData(Stream: TStream);
    procedure WriteStringsData(Stream: TStream);
    procedure ReadDataStorageVersion(Reader: TReader);
    procedure WriteDataStorageVersion(Writer: TWriter);
    procedure ReadSystemInfo(Reader: TReader);
    procedure WriteSystemInfo(Writer: TWriter);
  protected
    procedure DefineProperties(Filer: TFiler); override;

    function GetWorksheetClass: TvgrWorksheetClass; virtual;
    procedure BeforeChange(ChangeInfo : TvgrWorkbookChangeInfo); virtual;
    procedure AfterChange(ChangeInfo : TvgrWorkbookChangeInfo); virtual;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    function GetChildOwner: TComponent; override;

    function GetDispIdOfName(const AName: string) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    procedure DoChanged;

    property DataStorageVersion: Integer read FDataStorageVersion;

    property RangeStyles: TvgrRangeStylesList read FRangeStyles;
    property BorderStyles: TvgrBorderStylesList read FBorderStyles;
    property WBStrings: TvgrWBStrings read FWBStrings;
    property Formulas: TvgrFormulasList read FFormulas;
  public
{Returns the TvgrWorksheet object by specified index.
Parameters:
  Index - The index of the TvgrWorksheet object from 0 to WorksheetsCount - 1.
See also:
  TvgrWorksheet}
    property Worksheets[Index: Integer]: TvgrWorksheet read GetWorksheet;
{Returns the number of worksheets in workbook.
See also:
  TvgrWorksheet, Worksheets}
    property WorksheetsCount: Integer read GetWorksheetsCount;
{Clears the workbook.}
    procedure Clear; virtual;
{Removes the worksheet with specified Index from the workbook.
Example:
  Workbook.DeleteWorksheet(0);}
    procedure DeleteWorksheet(Index: Integer);
{Removes the specified worksheet from workbook.}
    procedure RemoveWorksheet(Worksheet: TvgrWorksheet);
{Returns index of worksheet with given title.
If worksheet with specified title is not exists, function returns -1.
Example:
  WorksheetIndex := Workbook.WorksheetIndexByTitle('Sheet1');
  if WorksheetIndex >= 0 then
    CurrentWorksheet := Workbook.Worksheets[WorksheetIndex];}
    function WorksheetIndexByTitle(const ATitle: string): Integer;
{Creates a new worksheet instance and adds it to the Worksheets collection of workbook.
Example:
  NewWorksheet := Workbook.AddWorksheet;
See also:
  Worksheets, TvgrWorksheet}
    function AddWorksheet : TvgrWorksheet;
{Creates a new worksheet instance and inserts it to the Worksheets collection of workbook.
Parameters:
  AInsertIndex - The position to inserting.
Example:
  NewWorksheet := Workbook.InsertWorksheet(0);}
    function InsertWorksheet(AInsertIndex: Integer): TvgrWorksheet;
{Attachs the Value to this workbook for receiving events about changes in workbook
Parameters:
  Value - The interface to object, which will receive events about changes in the book.
See also:
  DisconnectHandler}
    procedure ConnectHandler(Value: IvgrWorkbookHandler);
{Releases the Value from reception of events about changes in the workbook.
See also:
  ConnectHandler}
    procedure DisconnectHandler(Value: IvgrWorkbookHandler);
{Calls procedure Proc for all worksheets in workbook.}
    procedure GetChildren(Proc: TGetChildProc; Root: TComponent); override;

{Prevents the updating of the WorkbookGrid until the EndUpdate method is called.
Call BeginUpdate before making multiple changes.
Make sure that EndUpdate is called after all changes have been made.
Example:
  Workbook1.BeginUpdate;
  for i := 0 to 1000 do
    Workbook1.Worksheets[0].Ranges[0,i,0,i].Value = i;
  Workbook1.EndUpdate;
See also:
  EndUpdate}
    procedure BeginUpdate;
{Signals the end of an update operation.
See also:
  BeginUpdate}
    procedure EndUpdate;

{Saves the workbook into a stream.
Parameters:
  AStream - Specifies the stream into which the workbook will saved.
Example:
  Workbook.SaveToStream(stream);
See also:
  LoadFromStream, SaveToFile}
    procedure SaveToStream(AStream: TStream); virtual;
{Saves the workbook into file.
Parameters:
  AFileName - Specifies the file, into which the workbook will saved.
Example:
  Workbook.SaveToFile('Wb1.grw');
See also:
  SaveToStream, LoadFromStream, LoadFromFile}
    procedure SaveToFile(const AFileName: string); virtual;
{Loads the workbook from a stream into the TvgrWorkbook object.
Parameters:
  AStream - Specifies the stream from which the workbook is loaded.
Example:
  Workbook.LoadFromStream(stream);
See also:
  SaveToStream, LoadFromFile, SaveToFile}
    procedure LoadFromStream(AStream: TStream); virtual;
{Loads the workbook from file into the TvgrWorkbook object.
Parameters:
  AFileName - Specifies the file, from which the workbook is loaded.
Example:
  Workbook.LoadFromFile('Wb1.grw');
See also:
  SaveToStream, SaveToFile, LoadFromStream}
    procedure LoadFromFile(const AFileName: string); virtual;

{Creates an instance of the TvgrWorkbook class.
Parameters:
  AOwner - The component - owner.
Example:
  Workbook := TvgrWorkbook.Create(nil);}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of the TvgrWorkbook class.}
    destructor Destroy; override;
{Shows, that the book changes at present.}
    property Modified: Boolean read FModified write SetModified;
  published
{This event happens, when developer should release the interface to an element
(range, border, etc.) in connection with that it should be removed.
See also:
  TvgrItemInterfaceDeleteNotify}
    property OnItemInterfaceDelete: TvgrItemInterfaceDeleteNotify read FonItemInterfaceDelete write FonItemInterfaceDelete;
{Occurs after workbook is changed.}
    property OnChanged: TNotifyEvent read FOnChanged write FOnChanged;
  end;

{ This function formats the value.
Parameters:
  AFormat - is the format string.
  AValue - value.
Return value:
  Depending on a type of value, the mode of formatting is determined.
  If AValue is a numeric value FormatFloat function used,
  If AValue is a datetime value FormatDateTime function used,
  If AValue is a currency value FormatCurr function used,
  If AValue is a string value formatting not used. }
  function vgrFormatValue(const AFormat: string; const AValue: Variant): string;

implementation

{$IFDEF VTK_D6_OR_D7} uses DateUtils{, SysInit}; {$ENDIF}

const
  vgrDefaultWBListGrowSize = $1000;
  sSheetTitlePrefix = 'Sheet';
  vgrCF_WORKSHEET_RECT = CF_MAX + 1;

type
  /////////////////////////////////////////////////
  //
  // rvgrCalculatorRec
  //
  /////////////////////////////////////////////////
  rvgrCalculatorRec = record
    Calculator: TvgrFormulaCalculator;
    EnumRec: pvgrEnumStackRec;
    Callback: TvgrFormulaCalculatorEnumStackProc;
    RangeRect: TRect;
    RangeSheet: Integer;
    Result: Boolean;
  end;
  pvgrCalculatorRec = ^rvgrCalculatorRec;

function CheckDateFormat(const ADisplayFormat: string): boolean;
begin
  Result := (Pos('y',ADisplayFormat)>0)
    or (Pos('m',ADisplayFormat)>0)
    or (Pos('d',ADisplayFormat)>0)
    or (Pos('h',ADisplayFormat)>0)
    or (Pos('n',ADisplayFormat)>0)
    or (Pos('s',ADisplayFormat)>0)
    or (Pos('z',ADisplayFormat)>0);
end;


function vgrFormatValue(const AFormat: string; const AValue: Variant): string;
begin
  case VarType(AValue) of
    {$IFDEF VGR_D6_OR_D7}varShortInt, varWord, varLongWord, varInt64, {$ENDIF}
    varByte, varSmallint, varInteger, varSingle, varDouble, varCurrency:
      if not CheckDateFormat(AFormat) then
        Result := FormatFloat(AFormat, AValue)
      else
        Result := FormatDateTime(AFormat, AValue);
    varDate:
      Result := FormatDateTime(AFormat, AValue);
    varOleStr, varStrArg, varString:
      Result := VarToStr(AValue);
  else
    Result := VarToStr(AValue);
  end;
end;

function GetHashCode(const Buffer; Count: Integer): Word; assembler;
asm
        MOV     ECX,EDX
        MOV     EDX,EAX
        XOR     EAX,EAX
@@1:    ROL     AX,5
        XOR     AL,[EDX]
        INC     EDX
        DEC     ECX
        JNE     @@1
end;

function RangeValueToVariant(AValue: pvgrRangeValue; AWBStrings: TvgrWBStrings): Variant;
begin
  case AValue.ValueType of
    rvtInteger:
      Result := AValue.vInteger;
    rvtExtended:
      Result := AValue.vExtended;
    rvtString:
      Result := AWBStrings[AValue.vString];
    rvtDateTime:
      Result := AValue.vDateTime;
  else
    Result := Null;
  end;
end;

procedure VariantToRangeValue(var ARangeValue: rvgrRangeValue; const AVariant: Variant; AWBStrings: TvgrWBStrings);
var
  AVarType: Integer;
begin
  if ARangeValue.ValueType = rvtString then
    AWBStrings.Release(ARangeValue.vString);
    
  AVarType := VarType(AVariant);
  case AVarType of
    {$IFDEF VGR_D6_OR_D7}varShortInt, varWord, varLongWord, varInt64, {$ENDIF} varByte, varSmallint, varInteger:
      begin
        ARangeValue.ValueType := rvtInteger;
        ARangeValue.vInteger := AVariant;
      end;
    varCurrency, varSingle, varDouble:
      begin
        ARangeValue.ValueType := rvtExtended;
        ARangeValue.vExtended := AVariant;
      end;
    varDate:
      begin
        ARangeValue.ValueType := rvtDateTime;
        ARangeValue.vDateTime := AVariant;
      end;
    varOleStr, varStrArg, varString:
      begin
        ARangeValue.ValueType := rvtString;
        ARangeValue.vString := AWBStrings.FindOrAdd(VarToStr(AVariant));
      end;
    varBoolean:
      begin
        ARangeValue.ValueType := rvtInteger;
        ARangeValue.vInteger := AVariant;
      end;
  else
    ARangeValue.ValueType := rvtNull;
  end;
end;

procedure StringToRangeValue(var ARangeValue: rvgrRangeValue; const AString: string; AWBStrings: TvgrWBStrings);
var
  ACode: Integer;
  AInteger: Integer;
  AExtended: Extended;
  ADateTime: TDateTime;
begin
  if ARangeValue.ValueType = rvtString then
    AWBStrings.Release(ARangeValue.vString);
  if AString = '' then
    ARangeValue.ValueType := rvtNull
  else
  begin
    Val(AString, AInteger, ACode);
    if ACode = 0 then
    begin
      ARangeValue.ValueType := rvtInteger;
      ARangeValue.vInteger := AInteger;
    end
    else
      if TextToFloat(PChar(AString), AExtended, fvExtended) then
      begin
        ARangeValue.ValueType := rvtExtended;
        ARangeValue.vExtended := AExtended;
      end
      else
      begin
        if TryStrToDateTime(AString, ADateTime) then
        begin
          ARangeValue.vDateTime := ADateTime;
          ARangeValue.ValueType := rvtDateTime;
        end
        else
        begin
          ARangeValue.ValueType := rvtString;
          ARangeValue.vString := AWBStrings.FindOrAdd(AString);
        end;
      end;
  end;
end;

type
  /////////////////////////////////////////////////
  //
  // TvgrClipboardList
  //
  /////////////////////////////////////////////////
  TvgrClipboardList = class
  private
    FData: Pointer;
    FCount: Integer;
    FItemSize: Integer;
    FConvertProc: TvgrConvertRecordProc;
    function GetItem(Index: Integer): Pointer;
  public
    constructor Create(AItemSize: Integer; AConvertProc: TvgrConvertRecordProc);
    destructor Destroy; override;
    function AddItem(ASource: Pointer): Pointer;
    procedure SaveToStream(AStream: TStream);
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
    property ItemSize: Integer read FItemSize;
    property ConvertProc: TvgrConvertRecordProc read FConvertProc;
    property Count: Integer read FCount;
    property Items[Index: Integer]: Pointer read GetItem; default;
  end;
  
  /////////////////////////////////////////////////
  //
  // TvgrClipboardObject
  //
  /////////////////////////////////////////////////
{Internal class. Copies the part of worksheet to clipboard.}
  TvgrClipboardObject = class(TObject)
  private
    FWorkbook: TvgrWorkbook;
    FCopyPoint: TPoint;
    FCopySize: TSize;
    FRanges: TvgrClipboardList;
    FBorders: TvgrClipboardList;
    FCols: TvgrClipboardList;
    FRows: TvgrClipboardList;
    FWBStrings: TvgrWBStrings;
    FRangeStyles: TvgrRangeStylesList;
    FBorderStyles: TvgrBorderStylesList;
    FFormulas: TvgrFormulasList;
    function GetRangeCount: Integer;
    function GetRange(Index: Integer): pvgrRange;
    function GetBorderCount: Integer;
    function GetBorder(Index: Integer): pvgrBorder;
    function GetColCount: Integer;
    function GetCol(Index: Integer): pvgrCol;
    function GetRowCount: Integer;
    function GetRow(Index: Integer): pvgrRow;
  protected
    procedure AddRange(ARange: pvgrRange);
    procedure AddBorder(ABorder: pvgrBorder);
    procedure AddCol(ACol: pvgrCol);
    procedure AddRow(ARow: pvgrRow);

    procedure SaveToStream(AStream: TStream);
    procedure LoadFromStream(AStream: TStream);
    procedure CopyToClipboard;
    procedure PasteFromClipboard;

    property Workbook: TvgrWorkbook read FWorkbook;
    property CopyPoint: TPoint read FCopyPoint write FCopyPoint;
    property CopySize: TSize read FCopySize write FCopySize;
    property WBStrings: TvgrWBStrings read FWBStrings;
    property RangeStyles: TvgrRangeStylesList read FRangeStyles;
    property BorderStyles: TvgrBorderStylesList read FBorderStyles;
    property Formulas: TvgrFormulasList read FFormulas;
    property RangeCount: Integer read GetRangeCount;
    property Ranges[Index: Integer]: pvgrRange read GetRange;
    property BorderCount: Integer read GetBorderCount;
    property Borders[Index: Integer]: pvgrBorder read GetBorder;
    property ColCount: Integer read GetColCount;
    property Cols[Index: Integer]: pvgrCol read GetCol;
    property RowCount: Integer read GetRowCount;
    property Rows[Index: Integer]: pvgrRow read GetRow;
  public
    constructor Create(AWorkbook: TvgrWorkbook);
    destructor Destroy; override;
  end;

/////////////////////////////////////////////////
//
// TvgrWBListItem
//
/////////////////////////////////////////////////
constructor TvgrWBListItem.Create(AList: TvgrWBList; AData: Pointer; ADataIndex: Integer);
begin
  inherited Create;
  FList := AList;
  FData := AData;
  FDataIndex := ADataIndex;
  FValid:= True;
end;

destructor TvgrWBListItem.Destroy;
begin
  inherited;
end;

function TvgrWBListItem.GetItemData: Pointer;
begin
  Result := FData;
end;

function TvgrWBListItem.GetStyleData: Pointer;
begin
  Result := nil;
end;

procedure TvgrWBListItem.BeforeChangeProperty;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := GetEditChangesType;
  AChangeInfo.ChangedInterface := IvgrWBListItem(Self);
  Workbook.BeforeChange(AChangeInfo);
end;

procedure TvgrWBListItem.AfterChangeProperty;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := GetEditChangesType;
  AChangeInfo.ChangedInterface := IvgrWBListItem(Self);
  Workbook.AfterChange(AChangeInfo);
end;

function TvgrWBListItem.GetWorksheet: TvgrWorksheet;
begin
  Result := FList.Worksheet;
end;

function TvgrWBListItem.GetWorkbook: TvgrWorkbook;
begin
  Result := FList.Workbook;
end;

function TvgrWBListItem.GetValid: Boolean;
begin
  Result := FValid;
end;

procedure TvgrWBListItem.SetInvalid;
begin
  FValid := False;
end;

function TvgrWBListItem.QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
begin
  if GetInterface(IID, Obj) then
    Result := S_OK
  else
    Result := E_NOINTERFACE
end;

function TvgrWBListItem._AddRef: Integer;
begin
  Inc(FRefCount);
  Result := FRefCount;
end;

function TvgrWBListItem._Release: Integer;
begin
  Dec(FRefCount);
  Result := FRefCount;
end;

function TvgrWBListItem.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Count := 0;
  Result := S_OK;
end;

function TvgrWBListItem.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; 
begin
  Result := S_OK;
end;

function TvgrWBListItem.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
{
type
  PNamesArray = ^TNamesArray;
  TNamesArray = array[0..0] of PWideChar;

  PDispIdsArray = ^TDispIdsArray;
  TDispIdsArray = array[0..0] of Integer;
var
  PropName: String;
  PropInfo: Integer;
}
begin
  Result := Common_GetIDsOfNames(IID, Names, NameCount, LocaleID, DispIDs, GetDispIdOfName);
{
  PropName := PNamesArray(Names)[0];
  PropInfo := GetDispIdOfName(PropName);
  if PropInfo > 0 then
  begin
    PDispIdsArray(DispIds)[0] := PropInfo;
    Result := S_OK;
  end
  else
    Result := DISP_E_UNKNOWNNAME;
}
end;

function TvgrWBListItem.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
{
var
  dps : TDispParams absolute Params;
  HasParams : boolean;
  pDispIds : PDispIdList;
  iDispIdsSize : integer;
  WS: WideString;
  I: Integer;
}
begin
  Result := Common_Invoke(DispID,
                          IID,
                          LocaleID,
                          Flags,
                          Params,
                          VarResult,
                          ExcepInfo,
                          ArgErr,
                          ClassName,
                          DoInvoke,
                          DoCheckScriptInfo);
{
  pDispIds := NIL;
  HasParams := (dps.cArgs > 0);
  if HasParams then
  begin
    iDispIdsSize := dps.cArgs * SizeOf(TDispId);
    GetMem(pDispIds, iDispIdsSize);
  end;
  try
    if HasParams then
      for I := 0 to dps.cArgs - 1 do
        pDispIds^[I] := dps.cArgs - 1 - I;
    try
      Result := DoInvoke(DispId, IID, LocaleID, Flags, dps, pDispIds, VarResult, ExcepInfo, ArgErr);
    except
      on E: EInvalidParamCount do Result := DISP_E_BADPARAMCOUNT;
      on E: EInvalidParamType do  Result := DISP_E_BADVARTYPE;
      on E: Exception do begin
        if Assigned(ExcepInfo) then begin
          FillChar(ExcepInfo^, SizeOf(TExcepInfo), 0);
          TExcepInfo(ExcepInfo^).wCode := 1001;
          TExcepInfo(ExcepInfo^).BStrSource := ClassName;
          WS := E.Message;
          TExcepInfo(ExcepInfo^).bstrDescription := SysAllocString(PWideChar(WS));
        end;
        Result := DISP_E_EXCEPTION;
      end;
    end;
  finally
    if HasParams then
      FreeMem(pDispIds);
  end;
}
end;

function TvgrWBListItem.GetDispIdOfName(const AName: String): Integer; 
begin
  Result := -1;
end;

function TvgrWBListItem.DoInvoke(DispId: Integer;
                                 Flags: Integer;
                                 var AParameters: TvgrOleVariantDynArray;
                                 var AResult: OleVariant): HResult;
begin
  Result := E_UNEXPECTED;
end;

function TvgrWBListItem.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := S_OK;
end;

function TvgrWBListItem.GetItemIndex: Integer;
begin
  Result := FDataIndex;
end;

/////////////////////////////////////////////////
//
// TvgrWBList
//
/////////////////////////////////////////////////
constructor TvgrWBList.Create;
begin
  inherited Create;
  FWorksheet := AWorksheet;
  FItems := TList.Create;
  FRecSize := GetRecSize;
  FGrowSize := GetGrowSize;
  FListFieldStatistic1 := TvgrListFieldStatistic.Create;
  FListFieldStatistic2 := TvgrListFieldStatistic.Create;
end;

destructor TvgrWBList.Destroy;
begin
  Clear;
  FItems.Free;
  FListFieldStatistic1.Free;
  FListFieldStatistic2.Free;
  inherited;
end;

{$IFDEF VGR_DS_DEBUG}
function TvgrWBList.DebugInfo: TvgrDebugInfo;
var
  I, AEmptyItemCount: Integer;
begin
  Result := TvgrDebugInfo.Create;
  Result.Add('FRecSize', FRecSize, 'Размер записи в списке');
  Result.Add('FCount', FCount, 'Количество элементов в списке');
  Result.Add('FCapacity', FCapacity, 'Выделено памяти под столько элементов');
  Result.Add('FGrowSize', FGrowSize, 'Столько записей добавляется при увеличении размера списка');
  Result.Add('FItems.Count', FItems.Count, 'Количество отданных наружу интерфейсов');
  AEmptyItemCount := 0;
  for I := 0 to IntfItemCount - 1 do
    if IntfItems[I].RefCount = 0 then
      Inc(AEmptyItemCount);
  Result.Add('Свободных интерфейсов', AEmptyItemCount, 'Количество пустых интерфейсов (неиспользуемых)');
end;
{$ENDIF}

procedure TvgrWBList.Clear;
var
  I: Integer;
begin
  for I := 0 to IntfItemCount - 1 do
    IntfItems[I].Free;
  FItems.Clear;
  FreeMem(FRecs);
  FCount := 0;
  FCapacity := 0;
end;

function TvgrWBList.GetGrowSize: Integer;
begin
  Result := vgrDefaultWBListGrowSize;
end;

function TvgrWBList.GetCount: Integer;
begin
  Result := FCount;
end;

function TvgrWBList.GetIntfItem(Index: Integer): TvgrWBListItem;
begin
  Result := TvgrWBListItem(FItems[Index]);
end;

function TvgrWBList.GetIntfItemCount: Integer;
begin
  Result := FItems.Count;
end;

function TvgrWBList.GetDataListItem(Index: Integer): Pointer;
begin
  Result := Pchar(FRecs) + FRecSize * Index;
end;

function TvgrWBList.GetWorkbook: TvgrWorkbook;
begin
  Result := FWorksheet.Workbook;
end;

function TvgrWBList.CreateWBListItem(Index: Integer): TvgrWBListItem;
var
  I: Integer;
begin
  if (Index >= 0) and (Index < Count) then
  begin
    for I := 0 to FItems.Count - 1 do
      if IntfItems[I].FRefCount = 0 then
      begin
        Result := IntfItems[I];
        Result.FData := DataList[Index];
        Result.FDataIndex := Index;  // !!!
        exit;
      end;
    Result := GetWBListItemClass.Create(Self, DataList[Index], Index);
    FItems.Add(Result);
  end
  else
    Result := nil;
end;

procedure TvgrWBList.SaveToStream(AStream: TStream);
var
  Buf: Integer;
begin
  // Size of one record
  AStream.Write(FRecSize, 4);
  // Count of records
  Buf := Count;
  AStream.Write(Buf, 4);
  // Records
  AStream.Write(FRecs^, FRecSize * Count);
end;

class procedure TvgrWBList.LoadFromStreamConvertRecord(AOldRecord, ANewRecord: Pointer; AOldSize, ANewSize, ADataStorageVersion: Integer);
begin
  if AOldSize >= ANewSize then
    System.Move(AOldRecord^, ANewRecord^, ANewSize)
  else
    System.Move(AOldRecord^, ANewRecord^, AOldSize);
end;

procedure TvgrWBList.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
var
  I, Buf: Integer;
  ARecBuf: Pointer;
begin
  AStream.Read(Buf, 4);
  AStream.Read(FCount, 4);
  FCapacity := FCount;
  ReallocMem(FRecs, FCount * FRecSize);
  if ADataStorageVersion = vgrDataStorageVersion then
  begin
    // version of DataStorage not changed
    AStream.Read(FRecs^, FRecSize * FCount)
  end
  else
  begin
    FillChar(FRecs^, FCount * FRecSize, #0);

    GetMem(ARecBuf, Buf);
    try
      for I := 0 to FCount - 1 do
      begin
        AStream.Read(ARecBuf^, Buf);
        LoadFromStreamConvertRecord(ARecBuf, PChar(FRecs) + I * FRecSize, Buf, FRecSize, ADataStorageVersion);
      end;
    finally
      FreeMem(ARecBuf);
    end;
  end;
end;

procedure TvgrWBList.RecalcStatistics;
var
  I: Integer;
begin
  for I := 0 to FCount - 1 do
  begin
    FListFieldStatistic1.Add(ItemGetSizeField1(I));
    FListFieldStatistic2.Add(ItemGetSizeField2(I));
  end;
end;

procedure TvgrWBList.UpdateDSListItems(AStartIndex: Integer; AOffset: Integer);
var
  I, AHandler: Integer;
begin
  for I := 0 to IntfItemCount - 1 do
    with IntfItems[I] do
      if FRefCount > 0 then
      begin
        if (FDataIndex = AStartIndex) and (AOffset = -1) then
        begin
          SetInvalid;
          if Assigned(Workbook.OnItemInterfaceDelete) then
            Workbook.OnItemInterfaceDelete(Self, IntfItems[I]);
          with Workbook.FWorkbookHandlers do
            for AHandler := 0 to Count - 1 do
              IvgrWorkbookHandler(Items[AHandler]).DeleteItemInterface(IntfItems[I]);
        end;
        if FDataIndex >= AStartIndex then
          FDataIndex := FDataIndex + AOffset;
        FData := DataList[FDataIndex];
      end;
end;

procedure TvgrWBList.CalcRangeOfSearching(Idx1,Idx2,Size1,Size2 : Integer; out LowLev1,LowLev2,HiLev1,HiLev2 : Integer; DeleteIfCrossOver : Boolean);
begin
  LowLev1 := Idx1;
  LowLev2 := Idx2;
  HiLev1  := Idx1;
  HiLev2  := Idx2;
  if DeleteIfCrossOver then
  begin
    Dec(LowLev1,FListFieldStatistic1.Max + Size1 - 2);
    Dec(LowLev2,FListFieldStatistic2.Max + Size2 - 2);
    Inc(HiLev1,FListFieldStatistic1.Max + Size1 - 2);
    Inc(HiLev2,FListFieldStatistic2.Max + Size2 - 2);
    if HiLev1 < 0 then
      HiLev1 := MaxInt;
    if HiLev2 < 0 then
      HiLev2 := MaxInt;
    if LowLev1 > Idx1 then
      LowLev1 := 0;
    if LowLev2 > Idx2 then
      LowLev2 := 0;
  end;
end;

procedure TvgrWBList.SetLowRulerSpeedy(LowLev1,LowLev2 : Integer; var LowRuler : Integer);
var
  LowMarker, HiMarker, MiddleMarker : Integer;
begin
  if (Count = 0) then
    LowRuler := 0
  else
  if (Count = 1) then
  begin
    if  CheckSearchRule(0,LowLev1,LowLev2) then
      LowRuler := Count
    else
      LowRuler := 0;
  end
  else
  if CheckSearchRule(Count-1,LowLev1,LowLev2) then
      LowRuler := Count
  else
  begin
    LowMarker := 0;
    HiMarker := Count - 1;
    repeat
      MiddleMarker:=(LowMarker + HiMarker) div 2;
      if CheckSearchRule(MiddleMarker,LowLev1,LowLev2) then
        LowMarker := MiddleMarker + 1
      else
        HiMarker := MiddleMarker;
    until LowMarker = HiMarker;
    LowRuler := LowMarker;
  end;
end;

function TvgrWBList.CheckSearchRule(Index,LowLev1,LowLev2 : Integer) : boolean;
begin
  Result := (ItemGetIndexField1(Index) < LowLev1) or
    ((ItemGetIndexField1(Index) = LowLev1) and
    (ItemGetIndexField2(Index) < LowLev2));
end;

function TvgrWBList.CheckOuterSearchRule(Index,HiLev1,HiLev2 : Integer) : boolean;
begin
  if Index < Count then
      Result := (ItemGetIndexField1(Index) < HiLev1) or
        ((ItemGetIndexField1(Index) = HiLev1) and (ItemGetIndexField2(Index) <= HiLev2))
  else
    Result := False;
end;

procedure TvgrWBList.CorrectLowRuler(ANumber1, ANumber2, CurPos : Integer; var LowRuler : Integer);
begin
  if (LowRuler < Count) and ((ItemGetIndexField1(LowRuler) < ANumber1) or
    ((ItemGetIndexField1(LowRuler) = ANumber1) and
    (ItemGetIndexField2(LowRuler) <= ANumber2))) then
    LowRuler := CurPos;
end;

function TvgrWBList.ItemGetIndexField1(AIndex: Integer): Integer;
begin
  Result := 0;
end;

function TvgrWBList.ItemGetIndexField2(AIndex: Integer): Integer;
begin
  Result := 0;
end;

function TvgrWBList.ItemGetIndexField3(AIndex: Integer): Integer;
begin
  Result := 0;
end;

function TvgrWBList.ItemGetSizeField1(AIndex: Integer): Integer;
begin
  Result := 1;
end;

function TvgrWBList.ItemGetSizeField2(AIndex: Integer): Integer;
begin
  Result := 1;
end;

procedure TvgrWBList.ItemSetIndexField1(AIndex: Integer; Value: Integer);
begin
end;

procedure TvgrWBList.ItemSetIndexField2(AIndex: Integer; Value: Integer);
begin
end;

procedure TvgrWBList.ItemSetIndexField3(AIndex: Integer; Value: Integer);
begin
end;

procedure TvgrWBList.ItemSetSizeField1(AIndex: Integer; Value: Integer);
begin
end;

procedure TvgrWBList.ItemSetSizeField2(AIndex: Integer; Value: Integer);
begin
end;

function TvgrWBList.ItemInRect(AIndex: Integer; const ARect: TRect): Boolean;
var
  RectItem : TRect;
  R1, R2 : TRect;
begin
  R1 := ARect;
  Inc(R1.Right);
  Inc(R1.Bottom);
  RectItem := Rect(ItemGetIndexField2(AIndex),ItemGetIndexField1(AIndex),ItemGetIndexField2(AIndex) + ItemGetSizeField2(AIndex) - 1, ItemGetIndexField1(AIndex) + ItemGetSizeField1(AIndex) - 1);
  R2 := RectItem;
  Inc(R2.Right);
  Inc(R2.Bottom);
  Result := (PointInCut(RectItem.Left,ARect.Left,ARect.Right)
    or PointInCut(RectItem.Right,ARect.Left,ARect.Right)
    or PointInCut(ARect.Left,RectItem.Left,RectItem.Right))
    and
    (PointInCut(RectItem.Top,ARect.Top,ARect.Bottom)
    or PointInCut(RectItem.Bottom,ARect.Top,ARect.Bottom)
    or PointInCut(ARect.Top,RectItem.Top,RectItem.Bottom));
end;

function TvgrWBList.ItemInsideRect(AIndex: Integer; const ARect: TRect): Boolean;
begin
  Result := ItemInside(AIndex, ARect.Left, ARect.Right, 2) and
    ItemInside(AIndex, ARect.Top, ARect.Bottom, 1);
end;

function TvgrWBList.ItemInside(AIndex: Integer; AStartPos, AEndPos, ADimension: Integer): Boolean;
var
  AItemStart, AItemEnd: Integer;
begin
  if ADimension = 1 then
  begin
    AItemStart := ItemGetIndexField1(AIndex);
    AItemEnd   := ItemGetIndexField1(AIndex) + ItemGetSizeField1(AIndex) - 1;
  end
  else
  begin
    AItemStart := ItemGetIndexField2(AIndex);
    AItemEnd   := ItemGetIndexField2(AIndex) + ItemGetSizeField2(AIndex) - 1;
  end;
  Result := (AItemStart >= AStartPos) and (AItemEnd <= AEndPos);
end;

procedure ReallocMem(var P: Pointer; Size: Integer);
begin
  System.ReallocMem(P, Size);
end;

function TvgrWBList.Add: Integer;
begin
  if FCount = FCapacity then
  begin
    ReallocMem(FRecs, (FCount + GetGrowSize) * FRecSize);
    FCapacity := FCapacity + GetGrowSize;
  end;
  Result := FCount;
  Inc(FCount);
  UpdateDSListItems(0, 0);
end;

procedure TvgrWBList.Insert(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex > FCount) then
    EListError.Create('');
  if FCount = FCapacity then
  begin
    ReallocMem(FRecs, (FCount + GetGrowSize) * FRecSize);
    FCapacity := FCapacity + GetGrowSize;
  end;
  Move(DataList[AIndex]^, DataList[AIndex + 1]^, (FCount - AIndex) * FRecSize);
  Inc(FCount);
  UpdateDSListItems(AIndex, 1);
end;

procedure TvgrWBList.Delete(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex >= FCount) then
    EListError.Create('');
  System.Move(DataList[AIndex + 1]^, DataList[AIndex]^, (FCapacity - AIndex - 1) * FRecSize);
  {$IFDEF VTK_REALLOC_IF_DELETE}
  ReallocMem(FRecs, (FCapacity - 1) * FRecSize);
  {$ENDIF}
  Dec(FCount);
  {$IFDEF VTK_REALLOC_IF_DELETE}
  Dec(FCapacity);
  {$ENDIF}
  UpdateDSListItems(AIndex, -1);
end;

procedure TvgrWBList.DoDelete(AData: Pointer);
begin
end;

procedure TvgrWBList.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
end;

procedure TvgrWBList.InternalDelete(Index : Integer);
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := GetDeleteChangesType;
  AChangeInfo.ChangedInterface := IvgrWBListItem(CreateWBListItem(Index));
  Workbook.BeforeChange(AChangeInfo);

  ListFieldStatistic1.Delete(ItemGetSizeField1(Index));
  ListFieldStatistic2.Delete(ItemGetSizeField2(Index));
  DoDelete(DataList[Index]);
  Delete(Index);

  AChangeInfo.ChangedInterface := nil;
  Workbook.AfterChange(AChangeInfo);
end;

procedure TvgrWBList.InternalAdd(ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer; LowRuler : Integer);
var
  AIndex: Integer;
  AData: Pointer;
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  {$IFDEF DS_FIND_DEBUG}
  OutputDebugString(PChar(Format('Add %d %d %d %d',[LowRuler, ANumber1, ANumber2, ANumber3])));
  {$ENDIF}
  // fire event Before inserting
  AChangeInfo.ChangesType := GetCreateChangesType;
  AChangeInfo.ChangedInterface := nil;
  Workbook.BeforeChange(AChangeInfo);

  ListFieldStatistic1.Add(ASizeField1);
  ListFieldStatistic2.Add(ASizeField2);
  if LowRuler >= Count then
    AIndex := Add
  else
  begin
    AIndex := LowRuler;
    Insert(AIndex);
  end;
  AData := DataList[AIndex];
  DoAdd(AData, ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2);

  // fire event after inserting
  AChangeInfo.ChangedInterface := IvgrWBListItem((CreateWBListItem(AIndex))) ;//IvgrRange(TvgrRange(CreateWBListItem(AIndex)));
  Workbook.AfterChange(AChangeInfo);
end;

function TvgrWBList.InternalFind(ANumber1, ANumber2, ANumber3 : Integer;
                                 ASizeField1, ASizeField2 : Integer;
                                 DeleteIfCrossOver : Boolean;
                                 AutoCreateItem : Boolean) : TvgrWBListItem;
var
  i : Integer;
  FLowLev1 : Integer;
  FLowLev2 : Integer;
  FHiLev1 : Integer;
  FHiLev2 : Integer;
  LowRuler : Integer;
begin
  Result := nil;
  CalcRangeOfSearching(ANumber1,ANumber2,ASizeField1,ASizeField2,FLowLev1,FLowLev2,FHiLev1,FHiLev2, DeleteIfCrossOver);
  SetLowRulerSpeedy(FLowLev1,FLowLev2,LowRuler);
  i := LowRuler;
  while CheckOuterSearchRule(i,FHiLev1,FHiLev2) do
  begin
    if (ANumber1 = ItemGetIndexField1(i)) and
       (ANumber2 = ItemGetIndexField2(i)) and
       (ANumber3 = ItemGetIndexField3(i)) and
       (ASizeField1 = ItemGetSizeField1(i)) and
       (ASizeField2 = ItemGetSizeField2(i)) then
    begin
      Result := CreateWBListItem(i);
      if not DeleteIfCrossOver then
        Break;
    end
    else
      if (DeleteIfCrossOver) and (ANumber3 = ItemGetIndexField3(i)) and ItemInRect(i,Rect(ANumber2, ANumber1, ANumber2 + ASizeField2 - 1, ANumber1 + ASizeField1 - 1)) then
      begin
        InternalDelete(i);
        Dec(i);
      end;
    Inc(i);
    CorrectLowRuler(ANumber1,ANumber2,i, LowRuler);
  end;

  if AutoCreateItem and (Result = nil) then
  begin
    InternalAdd(ANumber1,ANumber2,ANumber3,ASizeField1,ASizeField2,LowRuler);
    Result := CreateWBListItem(LowRuler);
  end;
end;

procedure TvgrWBList.InternalFindItemsInRect(ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc; DeleteItem : boolean; AData: Pointer);
var
  i : Integer;
  Item : TvgrWBListItem;
  FLowLev1 : Integer;
  FLowLev2 : Integer;
  FHiLev1 : Integer;
  FHiLev2 : Integer;
  LowRuler : Integer;
  ASize1, ASize2 : Integer;
begin
  FLowLev1 := ARect.Top;
  FHiLev1  := ARect.Bottom;
  FLowLev2 := ARect.Left;
  FHiLev2  := ARect.Right;

  ASize1 := FHiLev1 - FLowLev1 + 1;
  if ASize1 < 0 then
    ASize1 := MaxInt;
  ASize2 := FHiLev2 - FLowLev2 + 1;
  if ASize2 < 0 then
    ASize2 := MaxInt;
  CalcRangeOfSearching(FLowLev1, FLowLev2, ASize1, ASize2, FLowLev1, FLowLev2, FHiLev1, FHiLev2, True);

  SetLowRulerSpeedy(FLowLev1, FLowLev2, LowRuler);
  i := LowRuler;
  while CheckOuterSearchRule(i,FHiLev1,FHiLev2) do
  begin
    if ItemInRect(i, ARect) then
      if DeleteItem then
      begin
        InternalDelete(i);
        Dec(i);
      end
      else
        if Assigned(CallBackProc) then
        begin
          Item := CreateWBListItem(i);
          CallBackProc(Item, I, AData);
        end;
    Inc(i);
  end;
end;

procedure TvgrWBList.InternalFindAndCallBack(const ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer);
begin
  InternalFindItemsInRect(ARect, CallBackProc, False, AData);
end;

procedure TvgrWBList.InternalDeleteLines(AStartPos, AEndPos: Integer);
var
  I : Integer;
  ALowLev1 : Integer;
  ALowLev2 : Integer;
  ALowRuler : Integer;
begin
  ALowLev1 := AStartPos;
  ALowLev2 := 0;

  Dec(ALowLev1, FListFieldStatistic1.Max + 1);

  SetLowRulerSpeedy(ALowLev1, ALowLev2, ALowRuler);
  i := ALowRuler;
  while I < Count do
  begin
    if ItemInside(i, AStartPos, AEndPos, 1) then
    begin
      InternalDelete(i);
      Dec(i);
    end
    else
      MoveItem(i, AStartPos, AEndPos, 1);
    Inc(I);
  end;
end;

procedure TvgrWBList.InternalDeleteCols(AStartPos, AEndPos: Integer);
var
  I : Integer;
begin
  i := 0;
  while I < Count do
  begin
    if ItemInside(i, AStartPos, AEndPos, 2) then
    begin
      InternalDelete(i);
      Dec(i);
    end
    else
      MoveItem(i, AStartPos, AEndPos, 2);
    Inc(I);
  end;
end;

procedure  TvgrWBList.InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer);
begin
end;

procedure  TvgrWBList.InternalExtendedActionInsertCols(AIndexBefore, ACount: Integer);
begin
end;

procedure TvgrWBList.InternalInsertLines(AIndexBefore, ACount: Integer);
var
  I: Integer;
  AItemStartPos: Integer;
begin
  I := Count-1;
  while (I >= 0) do
  begin
    AItemStartPos := ItemGetIndexField1(I);
    if (AItemStartPos >= AIndexBefore) then
      ItemSetIndexField1(I, AItemStartPos + ACount);
    Dec(I);
  end;
  InternalExtendedActionInsertLines(AIndexBefore, ACount);
  UpdateDSListItems(0, -1);
end;

procedure TvgrWBList.InternalInsertCols(AIndexBefore, ACount: Integer);
var
  I: Integer;
  AItemStartPos: Integer;
begin
  I := Count-1;
  while (I >= 0) do
  begin
    AItemStartPos := ItemGetIndexField2(I);
    if (AItemStartPos >= AIndexBefore) then
      ItemSetIndexField2(I, AItemStartPos + ACount);
    Dec(I);
  end;
  InternalExtendedActionInsertCols(AIndexBefore, ACount);
  UpdateDSListItems(0, -1);
end;

procedure TvgrWBList.MoveItem(AIndex, AStartPos, AEndPos, ADimension: Integer);
var
  AItemStart, AItemEnd, AItemSize: Integer;
  AItemStartOld: Integer;

  procedure Processing;
  begin
    if AItemEnd >= AStartPos then
    begin
      AItemEnd := AItemEnd - (AEndPos - AStartPos + 1);
      AItemStartOld := AItemStart;
      if AItemStart > AStartPos then
        AItemStart := Max(AStartPos, AItemStart - (AEndPos - AStartPos + 1));
    end;
    AItemSize := AItemEnd - AItemStart + 1;
  end;

begin
  if ADimension = 1 then
  begin
    AItemStart := ItemGetIndexField1(AIndex);
    AItemEnd   := ItemGetIndexField1(AIndex) + ItemGetSizeField1(AIndex) - 1;
    Processing;
    ItemSetIndexField1(AIndex, AItemStart);
    ItemSetSizeField1(AIndex, AItemSize);
  end
  else
  begin
    AItemStart := ItemGetIndexField2(AIndex);
    AItemEnd   := ItemGetIndexField2(AIndex) + ItemGetSizeField2(AIndex) - 1;
    Processing;
    ItemSetIndexField2(AIndex, AItemStart);
    ItemSetSizeField2(AIndex, AItemSize);
  end;

end;

procedure TvgrWBList.BeforeChangeItem(const ChangeInfo: TvgrWorkbookChangeInfo);
begin
  Workbook.BeforeChange(ChangeInfo);
end;

procedure TvgrWBList.AfterChangeItem(const ChangeInfo: TvgrWorkbookChangeInfo);
begin
  Workbook.AfterChange(ChangeInfo);
end;

/////////////////////////////////////////////////
//
// TvgrVector
//
/////////////////////////////////////////////////
{
function TvgrVector.GetIndexField1 : Integer;
begin
  Result := Number;
end;
}

function TvgrVector.GetData: pvgrVector;
begin
  Result := pvgrVector(FData);
end;

function TvgrVector.GetSize: Integer;
begin
  Result := Data.Size;
end;

procedure TvgrVector.SetSize(Value: Integer);
begin
  with Data^ do
    if Size <> Value then
    begin
      BeforeChangeProperty;
      Size := Value;
      AfterChangeProperty;
    end;
end;

function TvgrVector.GetVisible: Boolean;
begin
  Result := Data.Visible;
end;

procedure TvgrVector.SetVisible(Value: Boolean);
begin
  with Data^ do
    if Visible <> Value then
    begin
      BeforeChangeProperty;
      Visible := Value;
      AfterChangeProperty;
    end;
end;

function TvgrVector.GetNumber: Integer;
begin
  Result := Data.Number;
end;

procedure TvgrVector.Assign(ASource: IvgrVector);
begin
  BeforeChangeProperty;
  Data.Size := pvgrVector(ASource.ItemData).Size;
  Data.Visible := pvgrVector(ASource.ItemData).Visible;
  AfterChangeProperty;
end;

function TvgrVector.GetDispIdOfName(const AName: String) : Integer;
begin
  if AnsiCompareText(AName, 'Number') = 0 then
    Result := cs_TvgrVector_Number
  else
  if AnsiCompareText(AName, 'Size') = 0 then
    Result := cs_TvgrVector_Size
  else
  if AnsiCompareText(AName, 'Visible') = 0 then
    Result := cs_TvgrVector_Visible
  else
    Result := -1;
end;

function TvgrVector.DoInvoke(DispId: Integer;
                             Flags: Integer;
                             var AParameters: TvgrOleVariantDynArray;
                             var AResult: OleVariant): HResult;
begin
  Result := CheckScriptInfo(DispId, Flags, Length(AParameters), @siTvgrVector, siTvgrVectorLength);
  if Result <> S_OK then
    exit; 
  case DispId of
    cs_TvgrVector_Number:
      AResult := Data.Number;
    cs_TvgrVector_Size:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Data.Size
      else
        Size := AParameters[0];
    cs_TvgrVector_Visible:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Data.Visible
      else
        Visible := AParameters[0];
  end;
end;

/////////////////////////////////////////////////
//
// TvgrVectors
//
/////////////////////////////////////////////////
function TvgrVectors.GetItem(Number: Integer): IvgrVector;
begin
  Result := (IvgrVector(TvgrVector(InternalFind(Number, 0, 0, 1, 1, False, True))));
end;

function TvgrVectors.GetByIndex(Index: Integer): IvgrVector;
begin
  Result := (IvgrVector(TvgrVector(CreateWBListItem(Index))));
end;

function TvgrVectors.ItemGetIndexField1(AIndex: Integer): Integer;
begin
  Result := pvgrVector(DataList[AIndex]).Number;
end;

procedure TvgrVectors.ItemSetIndexField1(AIndex: Integer; Value: Integer);
begin
  pvgrVector(DataList[AIndex]).Number := Value;
end;

function TvgrVectors.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrVector;
end;

function TvgrVectors.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrVector);
end;

function TvgrVectors.Find(Number: Integer): IvgrVector;
begin
  Result := IvgrVector(TvgrVector(InternalFind(Number, 0, 0, 1, 1, False, False)));
end;

procedure TvgrVectors.Delete(AStartPos, AEndPos: Integer);
begin
  InternalDeleteLines(AStartPos, AEndPos);
end;

procedure TvgrVectors.Insert(AIndexBefore, ACount: Integer);
begin
  InternalInsertLines(AIndexBefore, ACount);
end;

procedure TvgrVectors.FindAndCallBack(AStartPos, AEndPos: Integer; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer);
begin
  InternalFindAndCallBack(Rect(0, AStartPos, 0, AEndPos), CallBackProc, AData);
end;

procedure TvgrVectors.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  with pvgrVector(AData)^ do
  begin
    Number := ANumber1;
    Size := 0;
    Visible := True;
  end;
end;

procedure TvgrVectors.InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer);
var
  AVector:  IvgrVector;
  I : Integer;
begin
  AVector := Find(AIndexBefore + ACount);
  if (AVector <> nil) and (ACount > 0) then
    for I := 0 to ACount - 1 do
      Items[AIndexBefore + I].Size := AVector.Size;
end;

/////////////////////////////////////////////////
//
// TvgrPageVectors
//
/////////////////////////////////////////////////
function TvgrPageVectors.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrPageVector;
end;

function TvgrPageVectors.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrPageVector);
end;

procedure TvgrPageVectors.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  inherited;
  with pvgrPageVector(AData)^ do
  begin
    Flags := 0;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrPageVector
//
/////////////////////////////////////////////////
function TvgrPageVector.GetPageBreak: Boolean;
begin
  Result := PageBreak;
end;

procedure TvgrPageVector.SetPageBreak(Value: Boolean);
begin
  PageBreak := Value;
end;

function TvgrPageVector.GetDispIdOfName(const AName: string) : Integer;
begin
  if AnsiCompareText(AName, 'PageBreak') = 0 then
    Result := cs_TvgrPageVector_PageBreak
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrPageVector.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrPageVector, siTvgrPageVectorLength);
end;

function TvgrPageVector.DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrPageVector_PageBreak:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := PageBreak
      else
        PageBreak := AParameters[0];
    else
      Result := DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

procedure TvgrPageVector.Assign(ASource: IvgrVector);
var
  AObj: IvgrPageVector;
begin
  if (ASource.QueryInterface(IID_IvgrPageVector, AObj) = S_OK) and (AObj <> nil)then
  begin
    BeforeChangeProperty;
    Data.Size := pvgrVector(ASource.ItemData).Size;
    Data.Visible := pvgrVector(ASource.ItemData).Visible;
    PageBreak := AObj.PageBreak;
    AfterChangeProperty;
  end
  else
    inherited Assign(ASource);
end;

/////////////////////////////////////////////////
//
// TvgrCol
//
/////////////////////////////////////////////////
function TvgrCol.GetEditChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcChangeCol;
end;

function TvgrCol.GetData: pvgrCol;
begin
  Result := pvgrCol(FData);
end;

function TvgrCol.GetFlags(Index: Integer): Boolean;
begin
  Result := (Data.Flags and (1 shl Index)) <> 0;
end;

procedure TvgrCol.SetFlags(Index: Integer; Value: Boolean);
var
  AData: rvgrCol;
  AMask: Word;
begin
  AData := Data^;
  AMask := 1 shl Index;
  if ((AData.Flags and AMask) <> 0) xor Value then
  begin
    BeforeChangeProperty;
    if Value then
      Data^.Flags := AData.Flags or AMask
    else
      Data^.Flags := AData.Flags and not AMask;
    AfterChangeProperty;
  end;
end;

function TvgrCol.GetDispIdOfName(const AName: String) : Integer;
begin
  if AnsiCompareText(AName, 'Width') = 0 then
    Result := cs_TvgrCol_Width
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrCol.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrCol, siTvgrColLength);
end;

function TvgrCol.DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrCol_Width:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Width
      else
        Width := AParameters[0];
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrCols
//
/////////////////////////////////////////////////
function TvgrCols.GetItem(Number: Integer): IvgrCol;
begin
  Result := IvgrCol(TvgrCol(InternalFind(Number, 0, 0, 1, 1, False, True)));
end;

function TvgrCols.GetByIndex(Index: Integer): IvgrCol;
begin
  Result := IvgrCol(TvgrCol(CreateWBListItem(Index)));
end;

function TvgrCols.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrCol);
end;

function TvgrCols.GetCreateChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcNewCol;
end;

function TvgrCols.GetDeleteChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcDeleteCol;
end;

function TvgrCols.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrCol;
end;

procedure TvgrCols.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  inherited DoAdd(AData, ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2);
  with pvgrCol(AData)^ do
  begin
    Width := Worksheet.PageProperties.Defaults.ColWidth;
  end;
end;

function TvgrCols.Find(Number: Integer): IvgrCol;
begin
  Result := IvgrCol(TvgrCol(InternalFind(Number, 0, 0, 1, 1, False, False)));
end;

/////////////////////////////////////////////////
//
// TvgrRow
//
/////////////////////////////////////////////////
function TvgrRow.GetEditChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcChangeRow;
end;

function TvgrRow.GetData: pvgrRow;
begin
  Result := pvgrRow(FData);
end;

function TvgrRow.GetFlags(Index: Integer): Boolean;
begin
  Result := (Data.Flags and (1 shl Index)) <> 0;
end;

procedure TvgrRow.SetFlags(Index: Integer; Value: Boolean);
var
  AData: rvgrRow;
  AMask: Word;
begin
  AData := Data^;
  AMask := 1 shl Index;
  if ((AData.Flags and AMask) <> 0) xor Value then
  begin
    BeforeChangeProperty;
    if Value then
      Data^.Flags := AData.Flags or AMask
    else
      Data^.Flags := AData.Flags and not AMask;
    AfterChangeProperty;
  end;
end;

function TvgrRow.GetDispIdOfName(const AName: string) : Integer;
begin
  if AnsiCompareText(AName, 'Height') = 0 then
    Result := cs_TvgrRow_Height
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrRow.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrRow, siTvgrRowLength);
end;

function TvgrRow.DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrRow_Height:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Height
      else
        Height := AParameters[0];
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrRows
//
/////////////////////////////////////////////////
procedure TvgrRows.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  inherited DoAdd(AData, ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2);
  with pvgrRow(AData)^ do
  begin
    Height := Worksheet.PageProperties.Defaults.RowHeight;
  end;
end;

function TvgrRows.GetItem(Number: Integer): IvgrRow;
begin
  Result := IvgrRow(TvgrRow(InternalFind(Number, 0, 0, 1, 1, False, True)));
end;

function TvgrRows.GetByIndex(Index: Integer): IvgrRow;
begin
  Result := IvgrRow(TvgrRow(CreateWBListItem(Index)));
end;

function TvgrRows.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrRow);
end;

function TvgrRows.GetCreateChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcNewRow;
end;

function TvgrRows.GetDeleteChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcDeleterow;
end;

function TvgrRows.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrRow;
end;

function TvgrRows.Find(Number: Integer): IvgrRow;
begin
  Result := IvgrRow(TvgrRow(InternalFind(Number,0,0,1,1,False, False)));
end;

/////////////////////////////////////////////////
//
// TvgrStylesList
//
/////////////////////////////////////////////////
constructor TvgrStylesList.Create;
var
  AStyle: pvgrStyleHeader;
begin
  inherited Create;
  FItems := TList.Create;
  FStyleItemSize := GetStyleItemSize;
  InternalAdd(AStyle);
  InitStyle(PChar(AStyle) + SizeOf(rvgrStyleHeader));
  AStyle.RefCount := 1;
  AStyle.Hash := GetHashCode((PChar(AStyle) + SizeOf(rvgrStyleHeader))^, FStyleItemSize)
end;

destructor TvgrStylesList.Destroy;
begin
  ClearAll;
  FItems.Free;
  inherited;
end;

{$IFDEF VGR_DEBUG}
function TvgrStylesList.DebugInfo: TvgrDebugInfo;
var
  I, ARefCountZeroCount: Integer;
begin
  Result := TvgrDebugInfo.Create;
  Result.Add('SizeOf(rvgrStyleHeader)', SizeOf(rvgrStyleHeader), 'Размер заголовка для каждой записи стиля');
  Result.Add('FRecSize', FStyleItemSize, 'Размер данных стиля');
  Result.Add('Count', Count, 'Количество стилей');
  ARefCountZeroCount := 0;
  for I := 0 to Count - 1 do
    if pvgrStyleHeader(FItems[I]).RefCount = 0 then
      Inc(ARefCountZeroCount);
  Result.Add('(RefCount = 0)', ARefCountZeroCount, 'Количество свободных элементов, у которых RefCount = 0');
end;
{$ENDIF}

procedure TvgrStylesList.ClearAll;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do // delete all items
    FreeMem(FItems[I]);
  FItems.Clear;
end;

procedure TvgrStylesList.Clear;
var
  I: Integer;
begin
  for I := Count - 1 downto 1 do // do not delete first item !!!
  begin
    FreeMem(FItems[I]);
    FItems.Delete(I);
  end;
end;

procedure TvgrStylesList.Release(AIndex: Integer);
begin
  if AIndex > 0 then
    Dec(pvgrStyleHeader(FItems[AIndex]).RefCount);
end;

procedure TvgrStylesList.AddRef(AIndex: Integer);
begin
  Inc(pvgrStyleHeader(FItems[AIndex]).RefCount);
end;

function TvgrStylesList.GetItem(Index: Integer): Pointer;
begin
  Result := PChar(FItems[Index]) + SizeOf(rvgrStyleHeader);
end;

function TvgrStylesList.GetCount: Integer;
begin
  Result := FItems.Count;
end;

procedure TvgrStylesList.SaveToStream(AStream: TStream);
var
  I, Buf: Integer;
begin
  // HeaderSize
  Buf := SizeOf(rvgrStyleHeader);
  AStream.Write(Buf, 4);

  // DataSize
  AStream.Write(FStyleItemSize, 4);

  // Count of records
  Buf := Count;
  AStream.Write(Buf, 4);

  for I := 0 to Count - 1 do
    AStream.Write(FItems[I]^, GetStyleItemSize + SizeOf(rvgrStyleHeader))
end;

procedure TvgrStylesList.LoadFromStreamConvertRecord(AOldData: Pointer;
                                                     ANewData: Pointer;
                                                     AOldHeaderSize, AOldDataSize,
                                                     ANewHeaderSize, ANewDataSize,
                                                     ADataStorageVersion: Integer);
begin
  if AOldHeaderSize + AOldDataSize > ANewHeaderSize + ANewDataSize then
    System.Move(AOldData^, ANewData^, ANewHeaderSize + ANewDataSize)
  else
    System.Move(AOldData^, ANewData^, AOldHeaderSize + AOldDataSize);
end;

procedure TvgrStylesList.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
var
  I, ACount, ASize, AHeaderSize, ADataSize: Integer;
  AReadBuf, P: Pointer;
begin
  ClearAll;
  AStream.Read(AHeaderSize, 4);
  AStream.Read(ADataSize, 4);
  AStream.Read(ACount, 4);
  FItems.Count := ACount;
  ASize := FStyleItemSize + SizeOf(rvgrStyleHeader);
  if ADataStorageVersion = vgrDataStorageVersion then
    for I := 0 to ACount - 1 do
    begin
      GetMem(P, ASize);
      AStream.Read(P^, ASize);
      FItems[I] := P;
    end
  else
  begin
    GetMem(AReadBuf, AHeaderSize + ADataSize);
    try
      for I := 0 to ACount - 1 do
      begin
        P := AllocMem(ASize);
        AStream.Read(AReadBuf^, AHeaderSize + ADataSize);
        LoadFromStreamConvertRecord(AReadBuf,
                                    P,
                                    AHeaderSize,
                                    ADataSize,
                                    SizeOf(rvgrStyleHeader),
                                    FStyleItemSize,
                                    ADataStorageVersion);
        FItems[I] := P;
      end;
    finally
      FreeMem(AReadBuf);
    end;
  end;
end;

procedure TvgrStylesList.InitStyle(AStyle: Pointer);
begin
end;

function TvgrStylesList.InternalAdd(var AStyle: pvgrStyleHeader): Integer;
begin
  GetMem(AStyle, SizeOf(rvgrStyleHeader) + FStyleItemSize);
  Result := FItems.Add(AStyle);
end;

function TvgrStylesList.InternalFindOrAdd(Data: Pointer): Integer;
var
  I: Integer;
  AHash: Integer;
  AStyle: pvgrStyleHeader;
  ADest: Pointer;
begin
  AHash := GetHashCode(Data^, FStyleItemSize);
  AStyle := nil;
  Result := 0;
  for I := 0 to Count - 1 do
    with pvgrStyleHeader(FItems[I])^ do
    begin
      if (Hash = AHash) and CompareMem(Items[I], Data, FStyleItemSize) then
      begin
        Result := I;
        Inc(RefCount);
        exit;
      end
      else
        if RefCount = 0 then
        begin
          AStyle := pvgrStyleHeader(FItems[I]);
          Result := I;
          break;
        end;
    end;
  if AStyle = nil then
    Result := InternalAdd(AStyle);
  AStyle.Hash := AHash;
  AStyle.RefCount := 1;
  ADest := PChar(AStyle) + SizeOf(rvgrStyleHeader);
  System.Move(Data^, ADest^, FStyleItemSize);
end;

/////////////////////////////////////////////////
//
// TvgrRangeStylesList
//
/////////////////////////////////////////////////
function TvgrRangeStylesList.GetItem(Index: Integer): pvgrRangeStyle;
begin
  Result := pvgrRangeStyle(inherited Items[Index]);
end;

function TvgrRangeStylesList.GetStyleItemSize: Integer;
begin
  Result := SizeOf(rvgrRangeStyle);
end;

procedure TvgrRangeStylesList.InitStyle(AStyle: Pointer);
begin
  with pvgrRangeStyle(AStyle)^ do
  begin
    FillBackColor := clNone;
    FillForeColor := clBlack;
    FillPattern := bsSolid;
    DisplayFormat := 0;
    Font.Size := vgrDefaultRangeFontSize;
    Font.Charset := DEFAULT_CHARSET;
    Font.Color := clWindowText;
    Font.Pitch := fpDefault;
    Font.Name := vgrDefaultRangeFontName;
    Font.Style := [];
    HorzAlign := vgrhaAuto;
    VertAlign := vgrvaTop;
    Angle := 0;
    Flags := vgrmask_RangeFlagsWordWrap;
  end;
end;

function TvgrRangeStylesList.FindOrAdd(const Style: rvgrRangeStyle; AOldIndex: Integer): Integer;
begin
  Release(AOldIndex);
  Result := InternalFindOrAdd(@Style);
end;

/////////////////////////////////////////////////
//
// TvgrRange
//
/////////////////////////////////////////////////
function TvgrRange.GetStyleData: Pointer;
begin
  Result := GetStyle;
end;

function TvgrRange.GetEditChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcChangeRange;
end;

function TvgrRange.GetData: pvgrRange;
begin
  Result := pvgrRange(FData);
end;

function TvgrRange.GetStyle: pvgrRangeStyle;
begin
  Result := Styles[Data.Style];
end;

function TvgrRange.GetStyles: TvgrRangeStylesList;
begin
  Result := Workbook.RangeStyles;
end;

procedure TvgrRange.ChangeValueType(ANewValueType: TvgrRangeValueType);
var
  S: string;
  ACode, AInteger: Integer;
  AExtended: Extended;
  ADateTime: TDateTime;
begin
  if (ANewValueType <> Data.Value.ValueType) and (Data.Value.ValueType <> rvtNull) then
  begin
    BeforeChangeProperty;
    with Data^.Value do
    begin
      if ValueType = rvtString then
        S := WBStrings[vString];

      case ANewValueType of
        rvtNull: ValueType := rvtNull;
        rvtInteger:
          case ValueType of
            rvtExtended:
              if IsZero(Trunc(vExtended)) and (vExtended >= Low(Integer)) and (vExtended <= MaxInt) then
              begin
                vInteger := Round(vExtended);
                ValueType := rvtInteger;
              end;
            rvtString:
              begin
                val(S, AInteger, ACode);
                if ACode = 0 then
                begin
                  WBStrings.Release(vString);
                  vInteger := AInteger;
                  ValueType := rvtInteger;
                end;
              end;
          end;
        rvtExtended:
          case ValueType of
            rvtInteger:
              begin
                vExtended := vInteger;
                ValueType := rvtExtended;
              end;
            rvtString:
              if TextToFloat(PChar(S), AExtended, fvExtended) then
              begin
                WBStrings.Release(vString);
                vExtended := AExtended;
                ValueType := rvtExtended;
              end;
            rvtDateTime:
              begin
                vExtended := vDateTime;
                ValueType := rvtExtended;
              end;
          end;
        rvtString:
          begin
            if ValueType = rvtString then
              WBStrings.Release(vString);
            case ValueType of
              rvtInteger: S := IntToStr(vInteger);
              rvtExtended: S := FloatToStr(vExtended);
              rvtDateTime: S := DateTimeToStr(vDateTime);
            end;
            ValueType := rvtString;
            vString := WBStrings.FindOrAdd(S);
          end;
        rvtDateTime:
          case ValueType of
            rvtInteger:
              begin
                vDateTime := vInteger;
                ValueType := rvtDateTime;
              end;
            rvtExtended:
              begin
                vDateTime := vExtended;
                ValueType := rvtDateTime;
              end;
            rvtString:
              if TryStrToDateTime(S, ADateTime) then
              begin
                WBStrings.Release(vString);
                vDateTime := ADateTime;
                ValueType := rvtDateTime;
              end;
          end;
      end;
    end;
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetWBStrings: TvgrWBStrings;
begin
  Result := Workbook.WBStrings;
end;

{
function TvgrRange.GetIndexField1 : Integer;
begin
  Result := Top;
end;

function TvgrRange.GetIndexField2 : Integer;
begin
  Result := Left;
end;

function TvgrRange.GetSizeField1 : Integer;
begin
  Result := Bottom - Top + 1;
end;

function TvgrRange.GetSizeField2 : Integer;
begin
  Result := Right - Left + 1;
end;

procedure TvgrRange.SetSizeField1(Value : Integer);
begin
  Data.Place.Bottom := Top + Value - 1;
end;

procedure TvgrRange.SetSizeField2(Value : Integer);
begin
  Data.Place.Right := Left + Value - 1;
end;
}

function TvgrRange.GetLeft: Integer;
begin
  Result := Data.Place.Left;
end;

function TvgrRange.GetTop: Integer;
begin
  Result := Data.Place.Top;
end;

function TvgrRange.GetRight: Integer;
begin
  Result := Data.Place.Right;
end;

function TvgrRange.GetBottom: Integer;
begin
  Result := Data.Place.Bottom;
end;

function TvgrRange.GetPlace: TRect;
begin
  Result := Data.Place;
end;

function TvgrRange.GetValue: Variant;
begin
  Result := RangeValueToVariant(@Data.Value, WBStrings);
end;

procedure TvgrRange.SetValue(const Value: Variant);
begin
  BeforeChangeProperty;
  if Data.Formula <> -1 then
  begin
    Workbook.Formulas.Release(Data.Formula);
    Data.Formula := -1;
  end;
  VariantToRangeValue(Data.Value, Value, WBStrings);
  AfterChangeProperty;
end;

function TvgrRange.GetDisplayText: string;
var
  ADisplayFormat: string;
begin
  ADisplayFormat := GetDisplayFormat;
  case Data.Value.ValueType of
    rvtInteger:
      if not CheckDateFormat(ADisplayFormat) then
        Result := FormatFloat(ADisplayFormat, Data.Value.vInteger)
      else
        Result := FormatDateTime(ADisplayFormat, Data.Value.vInteger);
    rvtExtended:
      if not CheckDateFormat(ADisplayFormat) then
        Result := FormatFloat(ADisplayFormat, Data.Value.vExtended)
      else
        Result := FormatDateTime(ADisplayFormat, Data.Value.vExtended);
    rvtDateTime:
      if (ADisplayFormat = '') or CheckDateFormat(ADisplayFormat) then
        Result := FormatDateTime(GetDisplayFormat, Data.Value.vDateTime)
      else
        Result := FormatFloat(ADisplayFormat, Data.Value.vDateTime);
    rvtString:
      Result := WBStrings[Data.Value.vString];
  else
    Result := '';
  end;
end;

function TvgrRange.GetFont: IvgrRangeFont;
begin
  Result := Self;
end;

procedure TvgrRange.Assign(ASource: IvgrRange);
begin
  BeforeChangeProperty;
  _AssignStyle(ASource);
  VariantToRangeValue(Data.Value, ASource.Value, WBStrings);
  AfterChangeProperty;
end;

procedure TvgrRange._AssignStyle(ASource: IvgrRange);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  with AStyle do
  begin
//    FontAssignRangeFont(ASource.Font); ????
    HorzAlign := ASource.HorzAlign;
    VertAlign := ASource.VertAlign;
    Flags := ASource.Flags;
    FillBackColor := ASource.FillBackColor;
    FillForeColor := ASource.FillForeColor;
    FillPattern := ASource.FillPattern;
    Angle := ASource.Angle;
    with Font do
    begin
      Size := ASource.Font.Size;
      Pitch := ASource.Font.Pitch;
      Style := ASource.Font.Style;
      Charset := ASource.Font.Charset;
      Name := ASource.Font.Name;
      Color := ASource.Font.Color;
    end;
    WBStrings.Release(DisplayFormat);
    DisplayFormat := WBStrings.FindOrAdd(ASource.DisplayFormat);

    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
  end;
end;

procedure TvgrRange.AssignStyle(ASource: IvgrRange);
begin
  BeforeChangeProperty;
  _AssignStyle(ASource);
  AfterChangeProperty;
end;

function TvgrRange.GetValueData: pvgrRangeValue;
begin
  Result := @Data^.Value;
end;

procedure TvgrRange.SetFont(Value: IvgrRangeFont);
begin
  FontAssignRangeFont(Value);
end;

function TvgrRange.GetFillBackColor: TColor;
begin
  Result := Style.FillBackColor;
end;

procedure TvgrRange.SetFillBackColor(Value: TColor);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.FillBackColor <> Value then
  begin
    BeforeChangeProperty;
    AStyle.FillBackColor := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFillForeColor: TColor;
begin
  Result := Style.FillForeColor;
end;

procedure TvgrRange.SetFillForeColor(Value: TColor);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.FillForeColor <> Value then
  begin
    BeforeChangeProperty;
    AStyle.FillForeColor := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFillPattern: TBrushStyle;
begin
  Result := Style.FillPattern;
end;

procedure TvgrRange.SetFillPattern(Value: TBrushStyle);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.FillPattern <> Value then
  begin
    BeforeChangeProperty;
    AStyle.FillPattern := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetDisplayFormat: string;
begin
  Result := WBStrings[Style.DisplayFormat];
end;

procedure TvgrRange.SetDisplayFormat(const Value: string);
var
  AStyle: rvgrRangeStyle;
begin
  if GetDisplayFormat <> Value then
  begin
    AStyle := Style^;
    BeforeChangeProperty;
    WBStrings.Release(AStyle.DisplayFormat);
    AStyle.DisplayFormat := WBStrings.FindOrAdd(Value);
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetHorzAlign: TvgrRangeHorzAlign;
begin
  Result := Style.HorzAlign;
end;

procedure TvgrRange.SetHorzAlign(Value: TvgrRangeHorzAlign);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.HorzAlign <> Value then
  begin
    BeforeChangeProperty;
    AStyle.HorzAlign := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetVertAlign: TvgrRangeVertAlign;
begin
  Result := Style.VertAlign;
end;

procedure TvgrRange.SetVertAlign(Value: TvgrRangeVertAlign);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.VertAlign <> Value then
  begin
    BeforeChangeProperty;
    AStyle.VertAlign := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetAngle: Word;
begin
  Result := Style.Angle;
end;

procedure TvgrRange.SetAngle(Value: Word);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Angle <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Angle := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFlags: Word;
begin
  Result := Style.Flags;
end;

procedure TvgrRange.SetFlags(Value: Word);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Flags <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Flags := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetWordWrap: Boolean;
begin
  Result := (Style.Flags and vgrmask_RangeFlagsWordWrap) <> 0;
end;

procedure TvgrRange.SetWordWrap(Value: Boolean);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if ((AStyle.Flags and vgrmask_RangeFlagsWordWrap) = 1) xor Value then
  begin
    BeforeChangeProperty;
    if Value then
      AStyle.Flags := AStyle.Flags or vgrmask_RangeFlagsWordWrap
    else
      AStyle.Flags := AStyle.Flags and not vgrmask_RangeFlagsWordWrap;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFormula: string;
begin
  if Data.Formula = -1 then
    Result := ''
  else
    Result := Workbook.Formulas.GetFormulaText(Data.Formula, Data.Place.Left, Data.Place.Top);
end;

procedure TvgrRange.SetFormula(const Value: string);
begin
  BeforeChangeProperty;
  Data.Formula := Workbook.Formulas.FindOrAdd(Value, Data.Formula, Data.Place.Left, Data.Place.Top);
  AfterChangeProperty;
end;

function TvgrRange.GetValueType: TvgrRangeValueType;
begin
  Result := Data.Value.ValueType;
end;

function TvgrRange.GetStringValue: string;
begin
  if Data.Formula = -1 then
    Result := VarToStr(Value)
  else
    Result := '='+Workbook.Formulas.GetFormulaText(Data.Formula, Data.Place.Left, Data.Place.Top);
end;

procedure TvgrRange.SetStringValue(const Value: string);
begin
  BeforeChangeProperty;
  if (Length(Value) > 0) and (Value[1] = '=') then
    Data.Formula := Workbook.Formulas.FindOrAdd(Copy(Value, 2, Length(Value) - 1),
                                                Data.Formula,
                                                Data.Place.Left,
                                                Data.Place.Top)
  else
  begin
    if Data.Formula <> -1 then
    begin
      Workbook.Formulas.Release(Data.Formula);
      Data.Formula := -1;
    end;
    StringToRangeValue(Data.Value, Value, WBStrings);
  end;
  AfterChangeProperty;
end;

function TvgrRange.GetSimpleStringValue: string;
begin
  Result := VarToStr(Value)
end;

procedure TvgrRange.SetSimpleStringValue(const Value: string);
begin
  BeforeChangeProperty;
  if Data.Formula <> -1 then
  begin
    Workbook.Formulas.Release(Data.Formula);
    Data.Formula := -1;
  end;
  StringToRangeValue(Data.Value, Value, WBStrings);
  AfterChangeProperty;
end;

function TvgrRange.GetFontSize: Integer;
begin
  Result := Style.Font.Size;
end;

procedure TvgrRange.SetFontSize(Value: Integer);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Font.Size <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Font.Size := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFontPitch: TFontPitch;
begin
  Result := Style.Font.Pitch;
end;

procedure TvgrRange.SetFontPitch(Value: TFontPitch);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Font.Pitch <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Font.Pitch := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFontStyle: TFontStyles;
begin
  Result := Style.Font.Style;
end;

procedure TvgrRange.SetFontStyle(Value: TFontStyles);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Font.Style <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Font.Style := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFontCharset: TFontCharset;
begin
  Result := Style.Font.Charset;
end;

procedure TvgrRange.SetFontCharset(Value: TFontCharset);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Font.Charset <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Font.Charset := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFontName: string;
begin
  Result := Style.Font.Name;
end;

procedure TvgrRange.SetFontName(const Value: string);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Font.Name <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Font.Name := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetFontColor: TColor;
begin
  Result := Style.Font.Color;
end;

procedure TvgrRange.SetFontColor(Value: TColor);
var
  AStyle: rvgrRangeStyle;
begin
  AStyle := Style^;
  if AStyle.Font.Color <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Font.Color := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrRange.GetExportData: Pointer;
begin
  if Worksheet.ExportDataIsActive then
    Result := Worksheet.FExportData[GetItemIndex]
  else
    Result := nil;
end;

procedure TvgrRange.SetExportData(Value: Pointer);
begin
  if Worksheet.ExportDataIsActive then
    Worksheet.FExportData[GetItemIndex] := Value;
end;

procedure TvgrRange.FontAssign(AFont: TFont);
var
  AStyle: rvgrRangeStyle;
begin
  BeforeChangeProperty;
  AStyle := Style^;
  with AStyle.Font do
  begin
    Size := AFont.Size;
    Pitch := AFont.Pitch;
    Style := AFont.Style;
    Charset := AFont.Charset;
    Name := AFont.Name;
    Color := AFont.Color;
  end;
  Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
  AfterChangeProperty;
end;

procedure TvgrRange.FontAssignRangeFont(AFont: IvgrRangeFont);
var
  AStyle: rvgrRangeStyle;
begin
  BeforeChangeProperty;
  AStyle := Style^;
  with AStyle.Font do
  begin
    Size := AFont.Size;
    Pitch := AFont.Pitch;
    Style := AFont.Style;
    Charset := AFont.Charset;
    Name := AFont.Name;
    Color := AFont.Color;
  end;
  Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
  AfterChangeProperty;
end;

procedure TvgrRange.FontAssignTo(AFont: TFont);
begin
  with Style^.Font do
  begin
    AFont.Size := Size;
    AFont.Pitch := Pitch;
    AFont.Style := Style;
    AFont.Charset := Charset;
    AFont.Name := Name;
    AFont.Color := Color;
  end;
end;

function TvgrRange.GetDispIdOfName(const AName: String) : Integer;
begin
  if AnsiCompareText(AName,'Left') = 0 then
    Result := cs_TvgrRange_Left
  else
  if AnsiCompareText(AName,'Top') = 0 then
    Result := cs_TvgrRange_Top
  else
  if AnsiCompareText(AName,'Right') = 0 then
    Result := cs_TvgrRange_Right
  else
  if AnsiCompareText(AName,'Bottom') = 0 then
    Result := cs_TvgrRange_Bottom
  else
  if AnsiCompareText(AName,'Value') = 0 then
    Result := cs_TvgrRange_Value
  else
  if AnsiCompareText(AName,'Font') = 0 then
    Result := cs_TvgrRange_Font
  else
  if AnsiCompareText(AName,'FillForeColor') = 0 then
    Result := cs_TvgrRange_FillForeColor
  else
  if AnsiCompareText(AName,'FillBackColor') = 0 then
    Result := cs_TvgrRange_FillBackColor
  else
  if AnsiCompareText(AName,'FillPattern') = 0 then
    Result := cs_TvgrRange_FillPattern
  else
  if AnsiCompareText(AName,'DisplayFormat') = 0 then
    Result := cs_TvgrRange_DisplayFormat
  else
  if AnsiCompareText(AName,'HorzAlign') = 0 then
    Result := cs_TvgrRange_HorzAlign
  else
  if AnsiCompareText(AName,'VertAlign') = 0 then
    Result := cs_TvgrRange_VertAlign
  else
  if AnsiCompareText(AName,'Angle') = 0 then
    Result := cs_TvgrRange_Angle
  else
  if AnsiCompareText(AName,'WordWrap') = 0 then
    Result := cs_TvgrRange_WordWrap
  else
  if AnsiCompareText(AName,'Name') = 0 then
    Result := cs_TvgrRange_Name
  else
  if AnsiCompareText(AName,'Style') = 0 then
    Result := cs_TvgrRange_Style
  else
  if AnsiCompareText(AName,'Color') = 0 then
    Result := cs_TvgrRange_Color
  else
  if AnsiCompareText(AName,'Size') = 0 then
    Result := cs_TvgrRange_Size
  else
  if AnsiCompareText(AName, 'Formula') = 0 then
    Result := cs_TvgrRange_Formula
  else
  if AnsiCompareText(AName, 'Charset') = 0 then
    Result := cs_TvgrRange_Charset
  else
    Result := 0;
end;

function TvgrRange.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrRange, siTvgrRangeLength);
end;

procedure TvgrRange.SetDefaultStyle;
begin
  BeforeChangeProperty;
  Styles.Release(Data.Style);
  Data.Style := 0;
  Styles.AddRef(Data.Style);
  AfterChangeProperty;
end;

function TvgrRange.DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
var
  AFontStyle: TFontStyles;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrRange_Left: AResult := Place.Left;
    cs_TvgrRange_Top: AResult := Place.Top;
    cs_TvgrRange_Right: AResult := Place.Right;
    cs_TvgrRange_Bottom: AResult := Place.Bottom;
    cs_TvgrRange_Value:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Value
      else
        Value := AParameters[0];
    cs_TvgrRange_Font:
      AResult := IvgrRange(Self);
    cs_TvgrRange_FillForeColor:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := FillForeColor
      else
        FillForeColor := AParameters[0];
    cs_TvgrRange_FillBackColor:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := FillBackColor
      else
        FillBackColor := AParameters[0];
    cs_TvgrRange_FillPattern:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := FillPattern
      else
        FillPattern := AParameters[0];
    cs_TvgrRange_DisplayFormat:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := DisplayFormat
      else
        DisplayFormat := AParameters[0];
    cs_TvgrRange_HorzAlign:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := HorzAlign
      else
        HorzAlign := AParameters[0];
    cs_TvgrRange_VertAlign:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := VertAlign
      else
        VertAlign := AParameters[0];
    cs_TvgrRange_Angle:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Angle
      else
        Angle := AParameters[0];
    cs_TvgrRange_WordWrap:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := WordWrap
      else
        WordWrap := AParameters[0];
    cs_TvgrRange_Name:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Font.Name
      else
        Font.Name := AParameters[0];
    cs_TvgrRange_Style:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
      begin
        AFontStyle := Font.Style;
        AResult := PInteger(@AFontStyle)^;
      end
      else
      begin
        PInteger(@AFontStyle)^ := AParameters[0];
        Font.Style := AFontStyle;
      end;
    cs_TvgrRange_Color:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Font.Color
      else
        Font.Color := AParameters[0];
    cs_TvgrRange_Size:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Font.Size
      else
        Font.Size := AParameters[0];
    cs_TvgrRange_Formula:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := GetFormula
      else
        SetFormula(AParameters[0]);
    cs_TvgrRange_Charset:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Font.Charset
      else
        Font.Charset := AParameters[0];
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrRanges
//
/////////////////////////////////////////////////
procedure TvgrRanges.ReleaseReferences(ARange: pvgrRange);
begin
  Formulas.Release(ARange^.Formula);
  WBStrings.Release(Styles[ARange^.Style].DisplayFormat);
  if ARange^.Value.ValueType = rvtString then
    WBStrings.Release(ARange^.Value.vString);
  Styles.Release(ARange^.Style);
end;

procedure TvgrRanges.InternalInsertLines(AIndexBefore, ACount: Integer);
var
  I, ACol, ARow: Integer;
  AItemStartPos, AItemEndPos: Integer;
  PRange: pvgrRange;
  ARangeEtalon, ARange: IvgrRange;
  ADimensions: TRect;
begin
  for I := 0 to Count - 1 do
  begin
    PRange := pvgrRange(DataList[I]);
    AItemStartPos := PRange.Place.Top;
    AItemEndPos := PRange.Place.Bottom;
    if ((AItemStartPos >= AIndexBefore) or (AItemEndPos >= AIndexBefore)) then
    begin
      Inc(PRange.Place.Top);
      Inc(PRange.Place.Bottom);
    end;
  end;
  ADimensions := Worksheet.Dimensions;
  for ACol := ADimensions.Left to ADimensions.Right do
  begin
    ARangeEtalon := Find(ACol, AIndexBefore+ACount, ACol, AIndexBefore+ACount);
    if (ARangeEtalon <> nil) then
      for ARow := AIndexBefore to (AIndexBefore + ACount - 1) do
      begin
        ARange := Items[ACol, ARow, ACol, ARow];
        ARange.AssignStyle(ARangeEtalon);
      end;
  end;
end;

procedure TvgrRanges.InternalInsertCols(AIndexBefore, ACount: Integer);
var
  I, ACol, ARow: Integer;
  AItemStartPos, AItemEndPos: Integer;
  PRange: pvgrRange;
  ARangeEtalon, ARange: IvgrRange;
  ADimensions: TRect;
begin
  for I := 0 to Count - 1 do
  begin
    PRange := pvgrRange(DataList[I]);
    AItemStartPos := PRange.Place.Left;
    AItemEndPos := PRange.Place.Right;
    if ((AItemStartPos >= AIndexBefore) or (AItemEndPos >= AIndexBefore)) then
    begin
      Inc(PRange.Place.Left);
      Inc(PRange.Place.Right);
    end;
  end;
  ADimensions := Worksheet.Dimensions;
  for ARow := ADimensions.Left to ADimensions.Right do
  begin
    ARangeEtalon := Find(AIndexBefore+ACount, ARow, AIndexBefore+ACount, ARow);
    if (ARangeEtalon <> nil) then
      for ACol := AIndexBefore to (AIndexBefore + ACount - 1) do
      begin
        ARange := Items[ACol, ARow, ACol, ARow];
        ARange.AssignStyle(ARangeEtalon);
      end;
  end;
end;

procedure TvgrRanges.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    ReleaseReferences(pvgrRange(DataList[I]));
  inherited;
end;

function TvgrRanges.GetItem(Left,Top,Right,Bottom : integer): IvgrRange;
begin
  Result := IvgrRange(TvgrRange(InternalFind(Top, Left, 0, Bottom - Top + 1, Right - Left + 1, True, True)));
end;

function TvgrRanges.GetByIndex(Index: Integer): IvgrRange;
begin
  Result := IvgrRange(TvgrRange(CreateWBListItem(Index)));
end;

function TvgrRanges.ItemGetIndexField1(AIndex: Integer): Integer;
begin
  Result := pvgrRange(DataList[AIndex]).Place.Top;
end;

function TvgrRanges.ItemGetIndexField2(AIndex: Integer): Integer;
begin
  Result := pvgrRange(DataList[AIndex]).Place.Left;
end;

function TvgrRanges.ItemGetSizeField1(AIndex: Integer): Integer;
begin
  with pvgrRange(DataList[AIndex]).Place do
    Result := Bottom - Top + 1;
end;

function TvgrRanges.ItemGetSizeField2(AIndex: Integer): Integer;
begin
  with pvgrRange(DataList[AIndex]).Place do
    Result := Right - Left + 1;
end;

procedure TvgrRanges.ItemSetIndexField1(AIndex: Integer; Value: Integer);
begin
  with pvgrRange(DataList[AIndex]).Place do
    Top := Value;
end;

procedure TvgrRanges.ItemSetIndexField2(AIndex: Integer; Value: Integer);
begin
  with pvgrRange(DataList[AIndex]).Place do
     Left := Value;
end;

procedure TvgrRanges.ItemSetSizeField1(AIndex: Integer; Value: Integer);
begin
  with pvgrRange(DataList[AIndex]).Place do
    Bottom := Top + Value  - 1;
end;

procedure TvgrRanges.ItemSetSizeField2(AIndex: Integer; Value: Integer);
begin
  with pvgrRange(DataList[AIndex]).Place do
    Right := Left + Value  - 1;
end;

function TvgrRanges.GetCreateChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcNewRange;
end;

function TvgrRanges.GetDeleteChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcDeleteRange;
end;

function TvgrRanges.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrRange;
end;

function TvgrRanges.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrRange);
end;

function TvgrRanges.GetGrowSize: Integer;
begin
  Result := (Count div 5 + 1 );
end;

procedure TvgrRanges.DoDelete(AData: Pointer);
begin
  ReleaseReferences(pvgrRange(AData));
  Worksheet.RangeOrBorderDeleted;
end;

procedure TvgrRanges.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  with pvgrRange(AData)^ do
  begin
    Place.Top := ANumber1;
    Place.Left := ANumber2;
    Place.Bottom := ANumber1 + ASizeField1 - 1;
    Place.Right := ANumber2 + ASizeField2 - 1;
    Style := 0;
    Value.ValueType := rvtNull;
    Value.vString := 0;
    Styles.AddRef(0);
    Formula := -1;
    Worksheet.RangeCheckDimensions(Place);
  end;
end;

procedure TvgrRanges.GetDefaultFont(AFont: TFont);
begin
  with AFont do
  begin
    Size := 10;
    Charset := DEFAULT_CHARSET;
    Color := clWindowText;
    Pitch := fpDefault;
    Name := vgrDefaultRangeFontName;
    Style := [];
  end;
end;

function TvgrRanges.Find(Left,Top,Right,Bottom : integer) : IvgrRange;
begin
  Result := IvgrRange(TvgrRange(InternalFind(Top, Left, 0, Bottom - Top + 1, Right - Left + 1, False, False)));
end;

procedure TvgrRanges.CallBackFindAtCell(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  FTemp := AItem as IvgrRange;
end;

function TvgrRanges.FindAtCell(X, Y: Integer): IvgrRange;
begin
  FTemp := nil;
  InternalFindAndCallback(Rect(X, Y, X, Y), CallbackFindAtCell, nil);
  Result := FTemp;
  FTemp := nil;
end;

procedure TvgrRanges.FindAndCallBack(const ARect: TRect; CallBackProc: TvgrFindListItemCallBackProc; AData: Pointer);
begin
  InternalFindAndCallBack(ARect, CallBackProc, AData);
end;

function TvgrRanges.GetStyles: TvgrRangeStylesList;
begin
  Result := Workbook.RangeStyles;
end;

function TvgrRanges.GetWBStrings: TvgrWBStrings;
begin
  Result := Workbook.WBStrings;
end;

function TvgrRanges.GetFormulas: TvgrFormulasList;
begin
  Result := Workbook.Formulas;
end;

procedure TvgrRanges.DeleteRows(AStartPos, AEndPos: Integer);
begin
  InternalDeleteLines(AStartPos, AEndPos);
end;

procedure TvgrRanges.DeleteCols(AStartPos, AEndPos: Integer);
begin
  InternalDeleteCols(AStartPos, AEndPos);
end;

procedure TvgrRanges.PerformDeleteCellsCallBack(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with FPerformDeleteCellInfo do
    if not ItemInsideRect(AItemIndex, SearchRect) then
    begin
      if AllowBreakRanges then
      begin
        ItemSetSizeField1(AItemIndex, 1);
        ItemSetSizeField2(AItemIndex, 1);
      end
      else
        Success := False;
    end;
end;

function TvgrRanges.PerformDeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean): Boolean;
begin
  with FPerformDeleteCellInfo do
  begin
    Success := True;
    AllowBreakRanges := ABreakRanges;
    Shift := ACellShift;
    if ACellShift = vgrcsLeft then
      SearchRect := Rect(ARect.Left, ARect.Top, MaxInt, ARect.Bottom)
    else
      SearchRect := Rect(ARect.Left, ARect.Top, ARect.Right, MaxInt);
    FindAndCallBack(SearchRect, PerformDeleteCellsCallBack, nil);
    Result := Success;
  end;
end;

procedure TvgrRanges.DeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean);
begin
  InternalFindItemsInRect(ARect, nil, True, nil);
  FPerformDeleteCellInfo.BottomRange := ARect.Bottom;
  with FPerformDeleteCellInfo do
  begin
    Shift := ACellShift;
    if ACellShift = vgrcsLeft then
    begin
      ShiftTo :=  ARect.Right - ARect.Left + 1;
      SearchRect := Rect(ARect.Left + ShiftTo, ARect.Top, MaxInt, ARect.Bottom);
    end
    else
    begin
      ShiftTo :=  ARect.Bottom - ARect.Top + 1;
      SearchRect := Rect(ARect.Left, ARect.Top + ShiftTo, ARect.Right, MaxInt);
    end;
    FindAndCallBack(SearchRect, DeleteCellsCallBack, nil);
    if ACellShift = vgrcsTop then
      ShiftToTop;
  end;
end;

procedure TvgrRanges.DeleteCellsCallBack(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ASize: Integer;
begin
  if FPerformDeleteCellInfo.Shift = vgrcsTop then
  begin
    FPerformDeleteCellInfo.BottomRange := Max(FPerformDeleteCellInfo.BottomRange, ItemGetIndexField1(AItemIndex));
  end
  else
  begin
    ASize := ItemGetSizeField2(AItemIndex);
    ItemSetIndexField2(AItemIndex, ItemGetIndexField2(AItemIndex) - FPerformDeleteCellInfo.ShiftTo);
    ItemSetSizeField2(AItemIndex, ASize);
  end;
end;

procedure TvgrRanges.ShiftToTop;
var
  ARow, ACol: Integer;
  ARangeItem: IvgrRange;
  ARange: rvgrRange;
  P : Pointer;
begin
  with FPerformDeleteCellInfo do
   for ARow := SearchRect.Top to BottomRange do
     for ACol := SearchRect.Left to SearchRect.Right do
     begin
       ARangeItem := FindAtCellForShift(ACol, ARow);
       if ARangeItem <> nil then
       begin
         ARange := rvgrRange(ARangeItem.ItemData^);
         ARangeItem := nil;
         if (ARange.Place.Top = ARow) and (ARange.Place.Left = ACol) then
         begin
           Dec(ARange.Place.Top, ShiftTo);
           Dec(ARange.Place.Bottom, ShiftTo);
           Delete(FTempShiftIndex);
           with IvgrWBListItem(Items[ARange.Place.Left, ARange.Place.Top, ARange.Place.Right, ARange.Place.Bottom]) do
           begin
             P := ItemData;
//             Inc(P, SizeOf());
             Move(ARange, P^, SizeOf(rvgrRange));
           end;
         end;
       end;
     end;
end;

procedure TvgrRanges.CallBackFindAtCellForShift(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  FTempShift := AItem as IvgrRange;
  FTempShiftIndex := AItemIndex;
end;

function TvgrRanges.FindAtCellForShift(X, Y: Integer): IvgrRange;
begin
  FTempShift := nil;
  InternalFindAndCallback(Rect(X, Y, X, Y), CallbackFindAtCellForShift, nil);
  Result := FTempShift;
  FTempShift := nil;
end;

procedure TvgrRanges.SaveToStream(AStream: TStream);
begin
  inherited;
end;

procedure TvgrRanges.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
begin
  inherited;
  RecalcStatistics;
  Worksheet.RangeOrBorderDeleted;
end;

/////////////////////////////////////////////////
//
// TvgrBorderStylesList
//
/////////////////////////////////////////////////
function TvgrBorderStylesList.GetItem(Index: Integer): pvgrBorderStyle;
begin
  Result := pvgrBorderStyle(inherited Items[Index]);
end;

function TvgrBorderStylesList.GetStyleItemSize: Integer;
begin
  Result := SizeOf(rvgrBorderStyle);
end;

procedure TvgrBorderStylesList.InitStyle(AStyle: Pointer);
begin
  with pvgrBorderStyle(AStyle)^ do
  begin
    Width := 0;
    Pattern := vgrbsSolid;
    Color := clBlack;
  end;
end;

function TvgrBorderStylesList.FindOrAdd(const Style: rvgrBorderStyle; AOldIndex: Integer): Integer;
begin
  if AOldIndex = -1 then
    Release(AOldIndex);
  Result := InternalFindOrAdd(@Style);
end;

/////////////////////////////////////////////////
//
// TvgrBorder
//
/////////////////////////////////////////////////
function TvgrBorder.GetEditChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcChangeBorder;
end;

procedure TvgrBorder.Assign(ASource: IvgrBorder);
var
  AStyle: rvgrBorderStyle;
begin
  AStyle := Style^;
  BeforeChangeProperty;
  AStyle.Width := ASource.Width;
  AStyle.Color := ASource.Color;
  AStyle.Pattern := ASource.Pattern;
  Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
  AfterChangeProperty;
end;

function TvgrBorder.GetStyleData: Pointer;
begin
  Result := GetStyle;
end;

function TvgrBorder.GetData: pvgrBorder;
begin
  Result := pvgrBorder(FData);
end;

function TvgrBorder.GetStyles: TvgrBorderStylesList;
begin
  Result := Workbook.BorderStyles;
end;

function TvgrBorder.GetStyle: pvgrBorderStyle;
begin
  Result := Styles[Data.Style];
end;

{
function TvgrBorder.GetIndexField1 : Integer;
begin
  Result := Top;
end;

function TvgrBorder.GetIndexField2 : Integer;
begin
  Result := Left;
end;

function TvgrBorder.GetIndexField3 : Integer;
begin
  Result := Integer(Orientation);
end;
}

function TvgrBorder.GetLeft: Integer;
begin
  Result := Data.Left;
end;

function TvgrBorder.GetTop: Integer;
begin
  Result := Data.Top;
end;

function TvgrBorder.GetOrientation: TvgrBorderOrientation;
begin
  Result := Data.Orientation;
end;

function TvgrBorder.GetWidth: Integer;
begin
  Result := Style.Width;
end;

procedure TvgrBorder.SetWidth(Value: Integer);
var
  AStyle: rvgrBorderStyle;
begin
  AStyle := Style^;
  if AStyle.Width <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Width := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrBorder.GetColor: TColor;
begin
  Result := Style.Color;
end;

procedure TvgrBorder.SetColor(Value: TColor);
var
  AStyle: rvgrBorderStyle;
begin
  AStyle := Style^;
  if AStyle.Color <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Color := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

function TvgrBorder.GetPattern: TvgrBorderStyle;
begin
  Result := Style.Pattern;
end;

procedure TvgrBorder.SetPattern(Value: TvgrBorderStyle);
var
  AStyle: rvgrBorderStyle;
begin
  AStyle := Style^;
  if AStyle.Pattern <> Value then
  begin
    BeforeChangeProperty;
    AStyle.Pattern := Value;
    Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
    AfterChangeProperty;
  end;
end;

procedure TvgrBorder.SetStyles(AWidth: Integer; AColor: TColor; APattern: TvgrBorderStyle);
var
  AStyle: rvgrBorderStyle;
begin
  AStyle := Style^;
  BeforeChangeProperty;
  AStyle.Width := AWidth;
  AStyle.Color := AColor;
  AStyle.Pattern := APattern;
  Data.Style := Styles.FindOrAdd(AStyle, Data.Style);
  AfterChangeProperty;
end;

function TvgrBorder.GetDispIdOfName(const AName: string) : Integer;
begin
  if AnsiCompareText(AName,'Left') = 0 then
    Result := cs_TvgrBorder_Left
  else
  if AnsiCompareText(AName,'Top') = 0 then
    Result := cs_TvgrBorder_Top
  else
  if AnsiCompareText(AName,'Orientation') = 0 then
    Result := cs_TvgrBorder_Orientation
  else
  if AnsiCompareText(AName,'Width') = 0 then
    Result := cs_TvgrBorder_Width
  else
  if AnsiCompareText(AName,'Color') = 0 then
    Result := cs_TvgrBorder_Color
  else
  if AnsiCompareText(AName,'Pattern') = 0 then
    Result := cs_TvgrBorder_Pattern
  else
    Result := 0;
end;

function TvgrBorder.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrBorder, siTvgrBorderLength);
end;

function TvgrBorder.DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrBorder_Left: AResult := Left;
    cs_TvgrBorder_Top: AResult := Top;
    cs_TvgrBorder_Orientation: AResult := Orientation;
    cs_TvgrBorder_Width: // Width
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Width
      else
        Width := AParameters[0];
    cs_TvgrBorder_Color: // Color
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Color
      else
        Color := AParameters[0];
    cs_TvgrBorder_Pattern: // Pattern
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := Pattern
      else
        Pattern := AParameters[0];
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrBorders
//
/////////////////////////////////////////////////
procedure TvgrBorders.ReleaseReferences(ABorder: pvgrBorder);
begin
  Styles.Release(ABorder.Style);
end;

procedure TvgrBorders.InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer);
var
  ABorderEtalon, ABorder:  IvgrBorder;
  ARow, ACol : Integer;
  ADimensions: TRect;
begin
  ADimensions := Worksheet.Dimensions;
  for ACol := ADimensions.Left to ADimensions.Right do
  begin
    ABorderEtalon := Find(ACol, AIndexBefore+ACount, vgrboTop);
    if (ABorderEtalon <> nil) then
      for ARow := AIndexBefore to (AIndexBefore + ACount - 1) do
      begin
        ABorder := Items[ACol, ARow, vgrboTop];
        ABorder.Assign(ABorderEtalon);
      end;
  end;
  for ACol := ADimensions.Left to ADimensions.Right do
  begin
    ABorderEtalon := Find(ACol, AIndexBefore+ACount, vgrboLeft);
    if (ABorderEtalon <> nil) then
      for ARow := AIndexBefore to (AIndexBefore + ACount - 1) do
      begin
        ABorder := Items[ACol, ARow, vgrboLeft];
        ABorder.Assign(ABorderEtalon);
      end;
  end;
end;

procedure TvgrBorders.InternalExtendedActionInsertCols(AIndexBefore, ACount: Integer);
var
  ABorderEtalon, ABorder:  IvgrBorder;
  ARow, ACol : Integer;
  ADimensions: TRect;
begin
  ADimensions := Worksheet.Dimensions;
  for ARow := ADimensions.Top to ADimensions.Bottom do
  begin
    ABorderEtalon := Find(AIndexBefore+ACount, ARow, vgrboTop);
    if (ABorderEtalon <> nil) then
      for ACol := AIndexBefore to (AIndexBefore + ACount - 1) do
      begin
        ABorder := Items[ACol, ARow, vgrboTop];
        ABorder.Assign(ABorderEtalon);
      end;
  end;
  for ARow := ADimensions.Top to ADimensions.Bottom do
  begin
    ABorderEtalon := Find(AIndexBefore+ACount, ARow, vgrboLeft);
    if (ABorderEtalon <> nil) then
      for ACol := AIndexBefore to (AIndexBefore + ACount - 1) do
      begin
        ABorder := Items[ACol, ARow, vgrboLeft];
        ABorder.Assign(ABorderEtalon);
      end;
  end;
end;

procedure TvgrBorders.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    ReleaseReferences(pvgrBorder(DataList[I]));
  inherited;
end;

function TvgrBorders.GetItem(Left,Top: Integer; Orientation: TvgrBorderOrientation): IvgrBorder;
begin
  Result := IvgrBorder(TvgrBorder(InternalFind(Top, Left, Integer(Orientation), 1, 1, False, True)));
end;

function TvgrBorders.GetByIndex(Index: Integer): IvgrBorder;
begin
  Result := IvgrBorder(TvgrBorder(CreateWBListItem(Index)));
end;

procedure TvgrBorders.SaveToStream(AStream: TStream);
begin
  inherited;
end;

procedure TvgrBorders.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
begin
  inherited;
  RecalcStatistics;
  Worksheet.RangeOrBorderDeleted;
end;

function TvgrBorders.ItemGetIndexField1(AIndex: Integer): Integer;
begin
  Result := pvgrBorder(DataList[AIndex]).Top;
end;

function TvgrBorders.ItemGetIndexField2(AIndex: Integer): Integer;
begin
  Result := pvgrBorder(DataList[AIndex]).Left;
end;

function TvgrBorders.ItemGetIndexField3(AIndex: Integer): Integer;
begin
  Result := Integer(pvgrBorder(DataList[AIndex]).Orientation);
end;

procedure TvgrBorders.ItemSetIndexField1(AIndex: Integer; Value: Integer);
begin
  pvgrBorder(DataList[AIndex]).Top := Value;
end;

procedure TvgrBorders.ItemSetIndexField2(AIndex: Integer; Value: Integer);
begin
  pvgrBorder(DataList[AIndex]).Left := Value;
end;

procedure TvgrBorders.ItemSetIndexField3(AIndex: Integer; Value: Integer);
begin
  pvgrBorder(DataList[AIndex]).Orientation := TvgrBorderOrientation(Value);
end;

function TvgrBorders.GetCreateChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcNewBorder;
end;

function TvgrBorders.GetDeleteChangesType: TvgrWorkbookChangesType;
begin
  Result := vgrwcDeleteBorder;
end;

function TvgrBorders.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrBorder;
end;

function TvgrBorders.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrBorder);
end;

function TvgrBorders.GetGrowSize: Integer;
begin
  Result := (Count div 5 + 1 );
end;

function TvgrBorders.Find(Left,Top: Integer; Orientation: TvgrBorderOrientation) : IvgrBorder;
begin
  Result := IvgrBorder(TvgrBorder(InternalFind(Top, Left, Integer(Orientation), 1, 1, False, False)));
end;

procedure TvgrBorders.FindAndCallBack(const ARect : TRect; CallBackProc: TvgrFindListItemCallBackProc; AData: Pointer);
begin
  InternalFindAndCallBack(ARect, CallBackProc, AData);
end;

type
  rvgrBordersFindAtRect = record
    Rect: TRect;
    Proc: TvgrFindListItemCallBackProc;
    Data: Pointer;
  end;
  pvgrBordersFindAtRect = ^rvgrBordersFindAtRect;

procedure TvgrBorders.FindAtRectAndCallBackCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrBorder, pvgrBordersFindAtRect(AData)^ do
    if ((Left < Rect.Right) or (Orientation = vgrboLeft)) and
       ((Top < Rect.Bottom) or (Orientation = vgrboTop)) then
      Proc(AItem, AItemIndex, Data);
end;

procedure TvgrBorders.FindAtRectAndCallBack(const ARect: TRect; CallBackProc: TvgrFindListItemCallBackProc; AData: Pointer);
var
  ARec: rvgrBordersFindAtRect;
begin
  ARec.Proc := CallBackProc;
  ARec.Rect := Rect(ARect.Left, ARect.Top, ARect.Right + 1, ARect.Bottom + 1){ARect};
  ARec.Data := AData;
  with ARect do
    FindAndCallBack(ARec.Rect{Rect(Left, Top, Right + 1, Bottom + 1)}, FindAtRectAndCallBackCallback, @ARec);
end;

function TvgrBorders.GetStyles: TvgrBorderStylesList;
begin
  Result := Workbook.BorderStyles;
end;

procedure TvgrBorders.DeleteRows(AStartPos, AEndPos: Integer);
begin
  InternalDeleteLines(AStartPos, AEndPos);
end;

procedure TvgrBorders.DeleteCols(AStartPos, AEndPos: Integer);
begin
  InternalDeleteCols(AStartPos, AEndPos);
end;

procedure TvgrBorders.DeleteCells(const ARect: TRect; ACellShift: TvgrCellShift);
begin

end;

procedure TvgrBorders.DoDelete(AData: Pointer);
begin
  ReleaseReferences(pvgrBorder(AData));
  Worksheet.RangeOrBorderDeleted;
end;

procedure TvgrBorders.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  with pvgrBorder(AData)^ do
  begin
    Left := ANumber2;
    Top := ANumber1;
    Orientation := TvgrBorderOrientation(ANumber3);
    Style := 0;
    Styles.AddRef(0);
    Worksheet.BorderCheckDimensions(Left, Top, Orientation);
  end;
end;

type
  rvgrClearBorders = record
    Orientation: TvgrBorderOrientation;
    Borders: Array of Integer;
    BordersCount: Integer;
  end;
  pvgrClearBorders = ^rvgrClearBorders;
  
procedure TvgrBorders.DeleteBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  AClearData: pvgrClearBorders;
begin
  AClearData := pvgrClearBorders(AData);
  with AItem as IvgrBorder do
  begin
    if Orientation = AClearData.Orientation then
    begin
      if AClearData.BordersCount = Length(AClearData.Borders) then
        SetLength(AClearData.Borders, AClearData.BordersCount + 1);
      AClearData.Borders[AClearData.BordersCount] := AItemIndex;
      Inc(AClearData.BordersCount);
    end;
  end;
end;

procedure TvgrBorders.DeleteBorders(const ARect: TRect);
var
  I: Integer;
  AClearData: rvgrClearBorders;
begin
  InternalFindItemsInRect(ARect, nil, True, nil);

  SetLength(AClearData.Borders, 100);
  AClearData.BordersCount := 0;

  AClearData.Orientation := vgrboLeft;
  FindAndCallBack(Rect(ARect.Right + 1,
                       ARect.Top,
                       ARect.Right + 1,
                       ARect.Bottom), DeleteBordersCallback, @AClearData);
  AClearData.Orientation := vgrboTop;
  FindAndCallBack(Rect(ARect.Left,
                       ARect.Bottom + 1,
                       ARect.Right,
                       ARect.Bottom + 1), DeleteBordersCallback, @AClearData);

  // ???
  for I := AClearData.BordersCount - 1 downto 0 do
    InternalDelete(AClearData.Borders[I]);
end;

procedure TvgrBorders.DeleteBorders(const ARect: TRect; ABorderTypes: TvgrBorderTypes);
var
  I: Integer;
  AClearData: rvgrClearBorders;
begin
  SetLength(AClearData.Borders, 100);
  AClearData.BordersCount := 0;

  AClearData.Orientation := vgrboLeft;
  if vgrbtLeft in ABorderTypes then
    FindAndCallBack(Rect(ARect.Left,
                         ARect.Top,
                         ARect.Left,
                         ARect.Bottom + 1), DeleteBordersCallback, @AClearData);

  if vgrbtCenter in ABorderTypes then
    FindAndCallBack(Rect(ARect.Left + 1,
                         ARect.Top,
                         ARect.Right,
                         ARect.Bottom + 1), DeleteBordersCallback, @AClearData);

  if vgrbtRight in ABorderTypes then
    FindAndCallBack(Rect(ARect.Right + 1,
                         ARect.Top,
                         ARect.Right + 1,
                         ARect.Bottom + 1), DeleteBordersCallback, @AClearData);

  AClearData.Orientation := vgrboTop;
  if vgrbtTop in ABorderTypes then
    FindAndCallBack(Rect(ARect.Left,
                         ARect.Top,
                         ARect.Right + 1,
                         ARect.Top), DeleteBordersCallback, @AClearData);

  if vgrbtMiddle in ABorderTypes then
    FindAndCallBack(Rect(ARect.Left,
                         ARect.Top + 1,
                         ARect.Right + 1,
                         ARect.Bottom), DeleteBordersCallback, @AClearData);

  if vgrbtBottom in ABorderTypes then
    FindAndCallBack(Rect(ARect.Left,
                         ARect.Bottom + 1,
                         ARect.Right + 1,
                         ARect.Bottom + 1), DeleteBordersCallback, @AClearData);

  // ???
  for I := AClearData.BordersCount - 1 downto 0 do
    InternalDelete(AClearData.Borders[I]);
end;

/////////////////////////////////////////////////
//
// TvgrSection
//
/////////////////////////////////////////////////
function TvgrSection.GetEditChangesType: TvgrWorkbookChangesType;
begin
  Result := TvgrSections(FList).GetEditChangesType;
end;

function TvgrSection.GetData: pvgrSection;
begin
  Result := pvgrSection(FData);
end;

function TvgrSection.GetParent: IvgrSection;
begin
  Result := IvgrSection(TvgrSection(FList.CreateWBListItem(TvgrSections(FList).SearchDirectOuter(StartPos, EndPos, Level))));
end;

function TvgrSection.GetFlags(Index: Integer): Boolean;
begin
  Result := (Data.Flags and (1 shl Index)) <> 0;
end;

procedure TvgrSection.SetFlags(Index: Integer; Value: Boolean);
var
  AData: rvgrSection;
  AMask: Word;
begin
  AData := Data^;
  AMask := 1 shl Index;
  if ((AData.Flags and AMask) <> 0) xor Value then
  begin
    BeforeChangeProperty;
    if Value then
      Data^.Flags := AData.Flags or AMask
    else
      Data^.Flags := AData.Flags and not AMask;
    AfterChangeProperty;
  end;
end;

function TvgrSection.GetRepeatOnPageTop: Boolean;
begin
  Result := RepeatOnPageTop;
end;

procedure TvgrSection.SetRepeatOnPageTop(Value: Boolean);
begin
  RepeatOnPageTop := Value;
end;

function TvgrSection.GetRepeatOnPageBottom: Boolean;
begin
  Result := RepeatOnPageBottom;
end;

procedure TvgrSection.SetRepeatOnPageBottom(Value: Boolean);
begin
  RepeatOnPageBottom := Value;
end;

function TvgrSection.GetPrintWithNextSection: Boolean;
begin
  Result := PrintWithNextSection;
end;

procedure TvgrSection.SetPrintWithNextSection(Value: Boolean);
begin
  PrintWithNextSection := Value;
end;

function TvgrSection.GetPrintWithPreviosSection: Boolean;
begin
  Result := PrintWithPreviosSection;
end;

procedure TvgrSection.SetPrintWithPreviosSection(Value: Boolean);
begin
  PrintWithPreviosSection := Value;
end;

function TvgrSection.GetDispIdOfName(const AName: String) : Integer;
begin
  if AnsiCompareText(AName,'StartPos') = 0 then
    Result := 1
  else
  if AnsiCompareText(AName,'EndPos') = 0 then
    Result := 2
  else
  if AnsiCompareText(AName,'Level') = 0 then
    Result := 3
  else
  if AnsiCompareText(AName, 'RepeatOnPageTop') = 0 then
    Result := 4
  else
  if AnsiCompareText(AName, 'RepeatOnPageBottom') = 0 then
    Result := 5
  else
  if AnsiCompareText(AName, 'PrintWithNextSection') = 0 then
    Result := 6
  else
  if AnsiCompareText(AName, 'PrintWithPreviosSection') = 0 then
    Result := 7
  else
    Result := -1;
end;

function TvgrSection.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrSection, siTvgrSectionLength);
end;

function TvgrSection.DoInvoke (DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrSection_StartPos: AResult := StartPos;
    cs_TvgrSection_EndPos: AResult := EndPos;
    cs_TvgrSection_Level: AResult := Level;
    cs_TvgrSection_RepeatOnPageTop:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := RepeatOnPageTop
      else
        RepeatOnPageTop := AParameters[0];
    cs_TvgrSection_RepeatOnPageBottom:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := RepeatOnPageBottom
      else
        RepeatOnPageBottom := AParameters[0];
    cs_TvgrSection_PrintWithNextSection:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := PrintWithNextSection
      else
        PrintWithNextSection := AParameters[0];
    cs_TvgrSection_PrintWithPreviosSection:
      if Flags and DISPATCH_PROPERTYPUT = 0 then
        AResult := PrintWithPreviosSection
      else
        PrintWithPreviosSection := AParameters[0];
    else
      inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

procedure TvgrSection.Assign(ASource: IvgrSection);
begin
  BeforeChangeProperty;
  Data.Flags := pvgrSection(ASource.ItemData).Flags;
  AfterChangeProperty;
end;

function TvgrSection.GetStartPos: Integer;
begin
  Result := Data.StartPos;
end;

function TvgrSection.GetEndPos: Integer;
begin
  Result := Data.EndPos;
end;

function TvgrSection.GetLevel: Integer;
begin
  Result := Data.Level;
end;

/////////////////////////////////////////////////
//
// TvgrSections
//
/////////////////////////////////////////////////
constructor TvgrSections.Create(AWorksheet: TvgrWorksheet);
begin
  inherited;
  FMaxStatistic := TvgrListFieldStatistic.Create;
end;

destructor TvgrSections.Destroy;
begin
  FMaxStatistic.Free;
  inherited;
end;

procedure TvgrSections.AddLevelRef(ALevel: Integer);
begin
  FMaxStatistic.Add(ALevel + 1);
end;

procedure TvgrSections.ReleaseLevelRef(ALevel: Integer);
begin
  FMaxStatistic.Delete(ALevel + 1);
end;

procedure TvgrSections.SearchInnersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
  begin
    if (FCurrentOuterLevel = (Level+1)) then
    begin
      FSearchInnersPresent := True;
    end
  end;
end;

function TvgrSections.SearchInners(AStartPos, AEndPos: Integer): Boolean;
var
  R : TRect;
begin
  FSearchInnersPresent := False;
  R := Rect(0, AStartPos, 0, AEndPos);
  InternalFindAndCallBack(R, SearchInnersCallback, nil);
  Result := FSearchInnersPresent;
end;

procedure TvgrSections.SearchCurrentOuterCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
  begin
    if (FCurrentOuterLevel = Level) then
    begin
      FCurrentOuterIndex := AItemIndex;
    end
  end;
end;

procedure TvgrSections.SearchDirectOuterCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
  begin
    if (Level < FDirectOuterLevel) and (Level > FDirectOuterLevelFor) then
    begin
      FDirectOuterIndex := AItemIndex;
      FDirectOuterLevel := Level;
    end
  end;
end;

function TvgrSections.SearchCurrentOuter(AStartPos, AEndPos, ALevel: Integer): Integer;
var
  R : TRect;
begin
  FCurrentOuterIndex := -1;
  FCurrentOuterLevel := ALevel;

  R := Rect(0,AStartPos,0,AEndPos);

  InternalFindAndCallBack(R, SearchCurrentOuterCallback, nil);
  Result := FCurrentOuterIndex;
end;

function TvgrSections.SearchDirectOuter(AStartPos, AEndPos, ALevel: Integer): Integer;
var
  R : TRect;
begin
  FDirectOuterIndex := -1;
  FDirectOuterLevel := MaxInt;
  FDirectOuterLevelFor := ALevel;
  R := Rect(0,AStartPos,0,AEndPos);

  InternalFindAndCallBack(R, SearchDirectOuterCallback, nil);
  Result := FDirectOuterIndex;
end;

procedure TvgrSections.DownLevel(AStartPos, AEndPos: Integer; var ALevel: Integer);
var
  ACurrentOuterIndex: Integer;
  AItem: TvgrSection;
begin
  ReleaseLevelRef(ALevel);
  Dec(ALevel);
  AddLevelRef(ALevel);
  ACurrentOuterIndex := SearchCurrentOuter(AStartPos, AEndPos, ALevel+2);
  if ACurrentOuterIndex >= 0 then
  begin
    AItem := TvgrSection(CreateWBListItem(ACurrentOuterIndex));
    with AItem, pvgrSection(AItem.FData)^ do
    begin
      if not SearchInners(StartPos, EndPos) then
      begin
        BeforeChangeProperty;
        DownLevel(StartPos, EndPos, Level);
        AfterChangeProperty;
      end;
    end;
  end;
end;

procedure TvgrSections.UpLevel(AStartPos, AEndPos: Integer; ALevel: Integer);
var
  ACurrentOuterIndex: Integer;
  AItem: TvgrSection;
begin
  ACurrentOuterIndex := SearchCurrentOuter(AStartPos, AEndPos, ALevel+1);
  if ACurrentOuterIndex >= 0 then
  begin
    AItem := TvgrSection(CreateWBListItem(ACurrentOuterIndex));
    with AItem, pvgrSection(AItem.FData)^ do
    begin
      BeforeChangeProperty;
      UpLevel(StartPos, EndPos, Level);
      Level := Level + 1;
      ReleaseLevelRef(Level-1);
      AddLevelRef(Level);
      AfterChangeProperty;
    end;
  end;
end;

procedure TvgrSections.AddItemToDelete(AIndex: Integer);
begin
  SetLength(FItemsToDelete, Length(FItemsToDelete)+1);
  FItemsToDelete[Length(FItemsToDelete)-1] := AIndex;
end;

procedure TvgrSections.AddOuterItem(AIndex: Integer);
begin
  SetLength(FOuterItems, Length(FOuterItems)+1);
  FOuterItems[Length(FOuterItems)-1] := AIndex;
end;

procedure TvgrSections.DeleteEnemyItems;
var
  I : Integer;
begin
  for I := 0 to Length(FItemsToDelete) - 1 do
    InternalDelete(FItemsToDelete[I] - I);
end;

procedure TvgrSections.CheckOutersLevel;
begin
  UpLevel(FTempStartPos, FTempEndPos, FTempLevel-1);
end;

function TvgrSections.GetItem(StartPos, EndPos: Integer): IvgrSection;
var
  ASection: TvgrSection;
begin
  ProcessIntersectedSections(StartPos, EndPos, 1);
  ASection := TvgrSection(InternalFind(StartPos, 0, 0, EndPos - StartPos + 1, 1, False, True));
  Result := IvgrSection(ASection);
end;

function TvgrSections.GetByIndex(Index: Integer): IvgrSection;
begin
  Result := IvgrSection(TvgrSection(CreateWBListItem(Index)));
end;

function TvgrSections.GetMaxLevel: Integer;
begin
  Result := FMaxStatistic.Max - 1;
end;

procedure TvgrSections.InternalInsertLines(AIndexBefore, ACount: Integer);
var
  I: Integer;
  AItemStartPos, AItemEndPos: Integer;
  PSection: pvgrSection;
begin
  for I := 0 to Count - 1 do
  begin
    PSection := pvgrSection(DataList[I]);
    AItemStartPos := PSection.StartPos;
    AItemEndPos := PSection.EndPos;
    if (AItemStartPos >= (AIndexBefore+1)) then
      Inc(PSection.StartPos);
    if (AItemEndPos >= AIndexBefore) then
      Inc(PSection.EndPos);
  end;
end;

procedure TvgrSections.InternalInsertCols(AIndexBefore, ACount: Integer);
begin
end;

procedure TvgrSections.ProcessIntersectedSections(AStartPos, AEndPos, ALevelIncrement: Integer);
var
  R : TRect;
begin
  FTempStartPos := AStartPos;
  FTempEndPos   := AEndPos;
  FTempLevel := 0;
  FItemExists := False;

  SetLength(FItemsToDelete,0);
  SetLength(FOuterItems,0);
  R := Rect(0,AStartPos,0,AEndPos);
  InternalFindAndCallBack(R, ProcessIntersectCallback, nil);
  if not FItemExists then
  begin
    CheckOutersLevel;
    DeleteEnemyItems;
  end;
end;

procedure TvgrSections.ProcessIntersectCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
  begin
    if not ((StartPos = FTempStartPos) and (EndPos = FTempEndPos)) then
    begin
      if ((StartPos < FTempStartPos) and
        (EndPos >= FTempStartPos) and
        (EndPos < FTempEndPos)) or
        ((FTempStartPos < StartPos) and
        (FTempEndPos >= StartPos) and
        (FTempEndPos < EndPos))  then
        AddItemToDelete(AItemIndex)
      else
      if (StartPos < FTempStartPos) or (EndPos > FTempEndPos) then
        AddOuterItem(AItemIndex)
      else
      if (StartPos > FTempStartPos) or (EndPos < FTempEndPos) then
        FTempLevel := Max(Level+1, FTempLevel)
    end
    else
      FItemExists := True;
  end;
end;

procedure TvgrSections.ProcessDeleteCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
    if (StartPos = FDeleteStartPos) and (EndPos = FDeleteEndPos) then
      FDeleteItemIndex := AItemIndex;
end;

procedure TvgrSections.DeleteRows(AStartPos, AEndPos: Integer);
begin
  InternalDeleteLines(AStartPos, AEndPos);
end;

procedure TvgrSections.SaveToStream(AStream: TStream);
begin
  inherited;
end;

procedure TvgrSections.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
begin
  inherited;
  RecalcStatistics;
end;

procedure TvgrSections.DoDelete(AData: Pointer);
begin
  with pvgrSection(AData)^ do
  begin
    ReleaseLevelRef(Level);
    DownLevel(StartPos, EndPos, Level);
  end;
  inherited;
end;

function TvgrSections.ItemGetIndexField1(AIndex: Integer): Integer;
begin
  Result := pvgrSection(DataList[AIndex]).StartPos;
end;

function TvgrSections.ItemGetSizeField1(AIndex: Integer): Integer;
begin
  with pvgrSection(DataList[AIndex])^ do
    Result := EndPos - StartPos + 1;
end;

procedure TvgrSections.ItemSetIndexField1(AIndex: Integer; Value: Integer);
begin
  with pvgrSection(DataList[AIndex])^ do
    StartPos := Value;
end;

procedure TvgrSections.ItemSetSizeField1(AIndex: Integer; Value: Integer);
begin
  with pvgrSection(DataList[AIndex])^ do
    EndPos := StartPos + Value - 1;
end;

function TvgrSections.GetCreateChangesType: TvgrWorkbookChangesType;
begin
  Result := FSectionsNotifyEvents.NewEvent;
end;

function TvgrSections.GetDeleteChangesType: TvgrWorkbookChangesType;
begin
  Result := FSectionsNotifyEvents.DeleteEvent;
end;

function TvgrSections.GetEditChangesType: TvgrWorkbookChangesType;
begin
  Result := FSectionsNotifyEvents.EditEvent;
end;

function TvgrSections.GetWBListItemClass: TvgrWBListItemClass;
begin
  Result := TvgrSection;
end;

function TvgrSections.GetRecSize: Integer;
begin
  Result := SizeOf(rvgrSection);
end;

procedure TvgrSections.DoAdd(AData: Pointer; ANumber1, ANumber2, ANumber3, ASizeField1, ASizeField2: Integer);
begin
  with pvgrSection(AData)^ do
  begin
    StartPos := ANumber1;
    EndPos := ANumber1 + ASizeField1 - 1;
    Level := FTempLevel;
    Flags := 0;
    AddLevelRef(Level);
  end;
end;

procedure TvgrSections.InitNotifyEvents(ANew, ADelete, AEdit: TvgrWorkbookChangesType);
begin
  with FSectionsNotifyEvents do
  begin
    NewEvent := ANew;
    DeleteEvent := ADelete;
    EditEvent := AEdit;
  end;
end;

procedure TvgrSections.InternalExtendedActionInsertLines(AIndexBefore, ACount: Integer);
begin
end;

procedure TvgrSections.FindAndCallBack(AStartPos, AEndPos: Integer; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer);
begin
  InternalFindAndCallBack(Rect(0,AStartPos,0,AEndPos), CallBackProc, AData);
end;

function TvgrSections.Find(StartPos, EndPos: Integer) : IvgrSection;
begin
  Result := IvgrSection(TvgrSection(InternalFind(StartPos, 0, 0, EndPos - StartPos + 1, 1, False, False)));
end;

procedure TvgrSections.Delete(AStartPos, AEndPos: Integer);
begin
  FDeleteStartPos := AStartPos;
  FDeleteEndPos := AEndPos;
  FDeleteItemIndex := -1;
  FindAndCallBack(AStartPos, AEndPos, ProcessDeleteCallback, nil);
  if FDeleteItemIndex >= 0 then
    InternalDelete(FDeleteItemIndex);
end;

/////////////////////////////////////////////////
//
// TvgrRangesFormatBorder
//
/////////////////////////////////////////////////
constructor TvgrRangesFormatBorder.Create(ARects: TvgrRectArray; ABorders: TvgrBorders; ABorderType: TvgrBorderType; AAutoCreate: Boolean);
begin
  FRectArray := ARects;
  FBorders := ABorders;
  FBorderType := ABorderType;
  FAutoCreate := AAutoCreate;
end;

procedure TvgrRangesFormatBorder.GetFormat(out AColor: TColor; out AWidth: Integer; out APattern: TvgrBorderStyle; out AHasColor, AHasWidth, AHasPattern: Boolean);
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  APattern := FTempObject.Pattern;
  AColor := FTempObject.Color;
  AWidth := FTempObject.Width;
  AHasPattern := FTempObject.HasPattern;
  AHasColor := FTempObject.HasColor;
  AHasWidth := FTempObject.HasWidth;
end;

procedure TvgrRangesFormatBorder.SetFormat(AColor: TColor; AWidth: Integer; APattern: TvgrBorderStyle);
var
  I: Integer;
begin
  if AWidth <= 0 then
  begin
    // delete borders
    for I := 0 to High(FRectArray) do
      FBorders.DeleteBorders(FRectArray[I], [FBorderType]);
  end
  else
  begin
    InitTempObject;
    FTempObject.Operations := [vgrfoSetWidth, vgrfoSetColor, vgrfoSetPattern];
    FTempObject.Pattern := APattern;
    FTempObject.Color := AColor;
    FTempObject.Width := AWidth;
    Execute;
  end;
end;

procedure TvgrRangesFormatBorder.AssignBorderStyle(ASource: TvgrRangesFormatBorder);
var
  AWidth: Integer;
begin
  if ASource.HasWidth then
  begin
    AWidth := ASource.Width;
    Width := AWidth;
    if AWidth <= 0 then
      exit;
  end;
  if ASource.HasColor then
    Color := ASource.Color;
  if ASource.HasPattern then
    Pattern := ASource.Pattern;
end;

function TvgrRangesFormatBorder.SameBorder(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation; ABorderType: TvgrBorderType): Boolean;
begin
  Result := ((ABorderType in [vgrbtLeft, vgrbtCenter, vgrbtRight]) and (AOrientation = vgrboLeft)) or
    ((ABorderType in [vgrbtTop, vgrbtMiddle, vgrbtBottom]) and (AOrientation = vgrboTop));
end;

procedure TvgrRangesFormatBorder.InitTempObject;
begin
  ZeroMemory(@FTempObject, SizeOf(FTempObject));
end;

function TvgrRangesFormatBorder.GetRect(ARect: TRect): TRect;
begin
  case FBorderType of
    vgrbtLeft: Result := Rect(ARect.Left, ARect.Top, ARect.Left, ARect.Bottom);
    vgrbtCenter: Result := Rect(ARect.Left + 1, ARect.Top, ARect.Right, ARect.Bottom);
    vgrbtRight: Result := Rect(ARect.Right + 1, ARect.Top, ARect.Right + 1, ARect.Bottom);
    vgrbtTop: Result := Rect(ARect.Left, ARect.Top, ARect.Right, ARect.Top);
    vgrbtMiddle: Result := Rect(ARect.Left, ARect.Top + 1, ARect.Right, ARect.Bottom);
    vgrbtBottom: Result := Rect(ARect.Left, ARect.Bottom + 1, ARect.Right, ARect.Bottom + 1);
  end;
end;

procedure TvgrRangesFormatBorder.Execute;
var
  ARectIndex: Integer;
  ACellsCount: Integer;
  ARow, ACol: Integer;
  ARect: TRect;
  ADefBorderStyle: rvgrBorderStyle;
  AOrientation: TvgrBorderOrientation;
begin
  ACellsCount := 0;
  for ARectIndex := 0 to Length(FRectArray)-1 do
  begin
    ARect := GetRect(FRectArray[ARectIndex]);
    if (ARect.Left <= ARect.Right) and (ARect.Top <= ARect.Bottom) then
    begin
      Inc(ACellsCount, (ARect.Bottom - ARect.Top + 1)*(ARect.Right - ARect.Left + 1));
      if FTempObject.Operations = [vgrfoGet] then
        FBorders.FindAndCallBack(ARect, ExecuteCallback, nil)
      else
      begin
        if FAutoCreate then
        begin
          if FBorderType in [vgrbtLeft, vgrbtCenter, vgrbtRight] then
            AOrientation := vgrboLeft
          else
            AOrientation := vgrboTop;
          for ACol := ARect.Left to ARect.Right do
            for ARow := ARect.Top to ARect.Bottom do
            begin
              ExecuteCallbackSet(FBorders.Items[ACol, ARow, AOrientation], 0, nil);
            end;
        end
        else
          FBorders.FindAndCallBack(ARect, ExecuteCallbackSet, nil)
      end;
    end;
  end;

  if (FTempObject.Operations = [vgrfoGet]) and (FAutoCreate) and (ACellsCount > FTempObject.Count) then
  begin
    FBorders.Styles.InitStyle(@ADefBorderStyle);
    with ADefBorderStyle do
    begin
      if FTempObject.Count = 0 then
      begin
        FTempObject.Color := Color;
        FTempObject.Width := Width;
        FTempObject.Pattern := Pattern;
        FTempObject.HasColor := True;
        FTempObject.HasWidth := True;
        FTempObject.HasPattern := True;
      end
      else
      begin
        FTempObject.HasColor := FTempObject.HasColor and (FTempObject.Color = Color);
        FTempObject.HasWidth := FTempObject.HasWidth and (FTempObject.Width = Width);
        FTempObject.HasPattern := FTempObject.HasPattern and (FTempObject.Pattern = Pattern);
      end;
      Inc(FTempObject.Count);
    end;
  end;
end;

procedure TvgrRangesFormatBorder.ExecuteCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with (AItem as IvgrBorder) do
  begin
    if SameBorder(Left, Top, Orientation, FBorderType) then
    begin
      if FTempObject.Count = 0 then
      begin
        FTempObject.Color := Color;
        FTempObject.Width := Width;
        FTempObject.Pattern := Pattern;
        FTempObject.HasColor := True;
        FTempObject.HasWidth := True;
        FTempObject.HasPattern := True;
      end
      else
      begin
        FTempObject.HasColor := FTempObject.HasColor and (FTempObject.Color = Color);
        FTempObject.HasWidth := FTempObject.HasWidth and (FTempObject.Width = Width);
        FTempObject.HasPattern := FTempObject.HasPattern and (FTempObject.Pattern = Pattern);
      end;
      Inc(FTempObject.Count);
    end;
  end;
end;

procedure TvgrRangesFormatBorder.ExecuteCallbackSet(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with (AItem as IvgrBorder) do
  begin
    if SameBorder(Left, Top, Orientation, FBorderType) then
    begin
      if vgrfoSetColor in FTempObject.Operations then
        Color := FTempObject.Color;
      if vgrfoSetWidth in FTempObject.Operations then
        Width := FTempObject.Width;
      if vgrfoSetPattern in FTempObject.Operations then
        Pattern := FTempObject.Pattern;
    end;
  end;
end;

function TvgrRangesFormatBorder.GetWidth: Integer;
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  Result := FTempObject.Width;
end;

procedure TvgrRangesFormatBorder.SetWidth(Value: Integer);
var
  I: Integer;
begin
  if Value <= 0 then
  begin
    // delete borders
    for I := 0 to High(FRectArray) do
      FBorders.DeleteBorders(FRectArray[I], [FBorderType]);
  end
  else
  begin
    InitTempObject;
    FTempObject.Operations := [vgrfoSetWidth];
    FTempObject.Width := Value;
    Execute;
  end;
end;

function TvgrRangesFormatBorder.GetColor: TColor;
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  Result := FTempObject.Color;
end;

procedure TvgrRangesFormatBorder.SetColor(Value: TColor);
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoSetColor];
  FTempObject.Color := Value;
  Execute;
end;

function TvgrRangesFormatBorder.GetPattern: TvgrBorderStyle;
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  Result := FTempObject.Pattern;
end;

procedure TvgrRangesFormatBorder.SetPattern(Value: TvgrBorderStyle);
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoSetPattern];
  FTempObject.Pattern := Value;
  Execute;
end;

function TvgrRangesFormatBorder.GetHasWidth: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  Result := FTempObject.HasWidth;
end;

function TvgrRangesFormatBorder.GetHasColor: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  Result := FTempObject.HasColor;
end;

function TvgrRangesFormatBorder.GetHasPattern: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrfoGet];
  Execute;
  Result := FTempObject.HasPattern;
end;

procedure TvgrRangesFormatBorder.SavePropertyToList(AList: TStrings; const AName: string; const AValue: Variant);
begin
  AList.Add(Format('border%d_%s=%s',[Integer(FBorderType), AName, AValue]));
end;

function TvgrRangesFormatBorder.LoadPropertyFromList(AList: TStrings; const AName: string; var AValue: Variant): Boolean;
begin
  AValue := AList.Values[Format('border%d_%s',[Integer(FBorderType),AName])];
  Result := (AValue <> '') or (AList.IndexOf(AName+'=') >= 0);
end;

procedure TvgrRangesFormatBorder.SaveToList(AList: TStrings);
begin
  if HasColor then
    SavePropertyToList(AList, 'Color', Color);
  if HasPattern then
    SavePropertyToList(AList, 'Pattern', Pattern);
  if HasWidth then
    SavePropertyToList(AList, 'Width', Width);
end;

procedure TvgrRangesFormatBorder.LoadFromList(AList: TStrings);
var
  APropValue: Variant;
begin
  if LoadPropertyFromList(AList, 'Color', APropValue) then
    Color := APropValue;
  if LoadPropertyFromList(AList, 'Width', APropValue) then
    Width := APropValue;
  if LoadPropertyFromList(AList, 'Pattern', APropValue) then
    Pattern := APropValue;
end;

/////////////////////////////////////////////////
//
// TvgrRangesFormatBorders
//
/////////////////////////////////////////////////
constructor TvgrRangesFormatBorders.Create(ARects: TvgrRectArray; ABorders: TvgrBorders; AAutoCreate: Boolean);
begin
  inherited Create;
  FBorders := ABorders;
  FRectArray := ARects;
  FAutoCreate := AAutoCreate;
  FLeft := TvgrRangesFormatBorder.Create(ARects, ABorders, vgrbtLeft, AAutoCreate);
  FCenter := TvgrRangesFormatBorder.Create(ARects, ABorders, vgrbtCenter, AAutoCreate);
  FRight := TvgrRangesFormatBorder.Create(ARects, ABorders, vgrbtRight, AAutoCreate);
  FTop := TvgrRangesFormatBorder.Create(ARects, ABorders, vgrbtTop, AAutoCreate);
  FMiddle := TvgrRangesFormatBorder.Create(ARects, ABorders, vgrbtMiddle, AAutoCreate);
  FBottom := TvgrRangesFormatBorder.Create(ARects, ABorders, vgrbtBottom, AAutoCreate);
end;

destructor TvgrRangesFormatBorders.Destroy;
begin
  FLeft.Free;
  FCenter.Free;
  FRight.Free;
  FTop.Free;
  FMiddle.Free;
  FBottom.Free;
  inherited;
end;

procedure TvgrRangesFormatBorders.AssignBordersStyles(ASource: TvgrRangesFormatBorders);
begin
  Left.AssignBorderStyle(ASource.Left);
  Center.AssignBorderStyle(ASource.Center);
  Right.AssignBorderStyle(ASource.Right);
  Top.AssignBorderStyle(ASource.Top);
  Middle.AssignBorderStyle(ASource.Middle);
  Bottom.AssignBorderStyle(ASource.Bottom);
end;

procedure TvgrRangesFormatBorders.SaveToList(AList: TStrings);
begin
  Left.SaveToList(AList);
  Center.SaveToList(AList);
  Right.SaveToList(AList);
  Top.SaveToList(AList);
  Middle.SaveToList(AList);
  Bottom.SaveToList(AList);
end;

procedure TvgrRangesFormatBorders.LoadFromList(AList: TStrings);
begin
  Left.LoadFromList(AList);
  Center.LoadFromList(AList);
  Right.LoadFromList(AList);
  Top.LoadFromList(AList);
  Middle.LoadFromList(AList);
  Bottom.LoadFromList(AList);
end;

/////////////////////////////////////////////////
//
// TvgrRangesFormat
//
/////////////////////////////////////////////////
constructor TvgrRangesFormat.Create(ARects: TvgrRectArray; AWorksheet: TvgrWorksheet; AAutoCreate: Boolean);
begin
  inherited Create;
  FForceBeginUpdateEndUpdate := False;
  FTempObject := TvgrRangesFormatTmpObject.Create;
  FRectArray := ARects;
  FWorksheet := AWorksheet;
  FAutoCreate := AAutoCreate;
  FBorders := TvgrRangesFormatBorders.Create(ARects, AWorksheet.BordersList, AAutoCreate);
end;

destructor TvgrRangesFormat.Destroy;
begin
  FTempObject.Free;
  FBorders.Free;
  inherited;
end;

procedure TvgrRangesFormat.GetFormat(
      out AFontName: TFontName;
      out AFontSize: Integer;
      out AFontStyle: TFontStyles;
      out AFontCharset: TFontCharset;
      out AFontColor, AFillBackColor, AFillForeColor: TColor;
      out AFillPattern: TBrushStyle;
      out ADisplayFormat: string;
      out AHorzAlign: TvgrRangeHorzAlign;
      out AVertAlign: TvgrRangeVertAlign;
      out AAngle: Word;
      out AWordWrap: Boolean;
      out AHasFlags: TvgrRangesFlags);
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  with FTempObject do
  begin
    AFontName := FontName;
    AFontSize := FontSize;
    AFontStyle := FontStyle;
    AFontCharset := FontCharset;
    AFontColor := FontColor;
    AFillBackColor := FillBackColor;
    AFillForeColor := FillForeColor;
    AFillPattern := FillPattern;
    ADisplayFormat := DisplayFormat;
    AHorzAlign := HorzAlign;
    AVertAlign := VertAlign;
    AAngle := Angle;
    AWordWrap := WordWrap;
    AHasFlags := HasFlags;
  end;
end;

procedure TvgrRangesFormat.SetFormat(
      AFontName: TFontName;
      AFontSize: Integer;
      AFontStyle: TFontStyles;
      AFontCharset: TFontCharset;
      AFontColor, AFillBackColor, AFillForeColor: TColor;
      AFillPattern: TBrushStyle;
      ADisplayFormat: string;
      AHorzAlign: TvgrRangeHorzAlign;
      AVertAlign: TvgrRangeVertAlign;
      AAngle: Word;
      AWordWrap: Boolean);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [
    vgrroSetFontName,
    vgrroSetFontSize,
    vgrroSetFontStyleBold,
    vgrroSetFontStyleItalic,
    vgrroSetFontStyleUnderline,
    vgrroSetFontStyleStrikeout,
    vgrroSetFontColor,
    vgrroSetFillBackColor,
    vgrroSetFillForeColor,
    vgrroSetFillPattern,
    vgrroSetDisplayFormat,
    vgrroSetHorzAlign,
    vgrroSetVertAlign,
    vgrroSetAngle,
    vgrroSetWordWrap,
    vgrroSetFontCharset];

    FontName := AFontName;
    FontSize := AFontSize;
    FontStyle := AFontStyle;
    FontCharset := AFontCharset;    
    FontColor := AFontColor;
    FillBackColor := AFillBackColor;
    FillForeColor := AFillForeColor;
    FillPattern := AFillPattern;
    DisplayFormat := ADisplayFormat;
    HorzAlign := AHorzAlign;
    VertAlign := AVertAlign;
    Angle := AAngle;
    WordWrap := AWordWrap;
  end;
  Execute;
end;

function TvgrRangesFormat.GetHasFlags: TvgrRangesFlags;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.HasFlags;
end;

function TvgrRangesFormat.GetFontName: TFontName;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FontName;
end;

function TvgrRangesFormat.GetFontSize: Integer;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FontSize;
end;

function TvgrRangesFormat.GetFontStyle: TFontStyles;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FontStyle;
end;

function TvgrRangesFormat.GetFontStyleBold: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := fsBold in FTempObject.FontStyle;
end;

function TvgrRangesFormat.GetFontStyleItalic: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := fsItalic in FTempObject.FontStyle;
end;

function TvgrRangesFormat.GetFontStyleUnderline: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := fsUnderline in FTempObject.FontStyle;
end;

function TvgrRangesFormat.GetFontStyleStrikeout: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := fsStrikeout in FTempObject.FontStyle;
end;

function TvgrRangesFormat.GetFontCharset: TFontCharset;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FontCharset;
end;

function TvgrRangesFormat.GetFontColor: TColor;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FontColor;
end;

function TvgrRangesFormat.GetFillBackColor: TColor;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FillBackColor;
end;

function TvgrRangesFormat.GetFillForeColor: TColor;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FillForeColor;
end;

function TvgrRangesFormat.GetFillPattern: TBrushStyle;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.FillPattern;
end;

function TvgrRangesFormat.GetDisplayFormat: string;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.DisplayFormat;
end;

function TvgrRangesFormat.GetHorzAlign: TvgrRangeHorzAlign;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.HorzAlign;
end;

function TvgrRangesFormat.GetVertAlign: TvgrRangeVertAlign;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.VertAlign;
end;

function TvgrRangesFormat.GetAngle: Word;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.Angle;
end;

function TvgrRangesFormat.GetWordWrap: Boolean;
begin
  InitTempObject;
  FTempObject.Operations := [vgrroGet];
  Execute;
  Result := FTempObject.WordWrap;
end;

procedure TvgrRangesFormat.SetFontName(Value: TFontName);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFontName];
    FontName := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontSize(Value: Integer);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFontSize];
    FontSize := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontStyle(Value: TFontStyles);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [
    vgrroSetFontStyleBold,
    vgrroSetFontStyleItalic,
    vgrroSetFontStyleUnderline,
    vgrroSetFontStyleStrikeout];
    FontStyle := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontStyleBold(Value: Boolean);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [
    vgrroSetFontStyleBold];
    if Value then
      Include(FontStyle, fsBold)
    else
      Exclude(FontStyle, fsBold);
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontStyleItalic(Value: Boolean);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [
    vgrroSetFontStyleItalic];
    if Value then
      Include(FontStyle, fsItalic)
    else
      Exclude(FontStyle, fsItalic);
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontStyleUnderline(Value: Boolean);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [
    vgrroSetFontStyleUnderline];
    if Value then
      Include(FontStyle, fsUnderline)
    else
      Exclude(FontStyle, fsUnderline);
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontStyleStrikeout(Value: Boolean);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [
    vgrroSetFontStyleStrikeout];
    if Value then
      Include(FontStyle, fsStrikeout)
    else
      Exclude(FontStyle, fsStrikeout);
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontCharset(Value: TFontCharset);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFontCharset];
    FontCharset := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFontColor(Value: TColor);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFontColor];
    FontColor := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFillBackColor(Value: TColor);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFillBackColor];
    FillBackColor := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFillForeColor(Value: TColor);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFillForeColor];
    FillForeColor := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetFillPattern(Value: TBrushStyle);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetFillPattern];
    FillPattern := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetDisplayFormat(Value: string);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetDisplayFormat];
    DisplayFormat := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetHorzAlign(Value: TvgrRangeHorzAlign);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetHorzAlign];
    HorzAlign := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetVertAlign(Value: TvgrRangeVertAlign);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetVertAlign];
    VertAlign := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetAngle(Value: Word);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetAngle];
    Angle := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SetWordWrap(Value: Boolean);
begin
  InitTempObject;
  with FTempObject do
  begin
    FTempObject.Operations := [vgrroSetWordWrap];
    WordWrap := Value;
  end;
  Execute;
end;

procedure TvgrRangesFormat.SavePropertyToList(AList: TStrings; const AName: string; const AValue: Variant);
begin
  AList.Add(Format('%s=%s',[AName, AValue]));
end;

function TvgrRangesFormat.LoadPropertyFromList(AList: TStrings; const AName: string; var AValue: Variant): Boolean;
begin
  AValue := AList.Values[AName];
  Result := (AValue <> '') or (AList.IndexOf(AName+'=') >= 0);
end;

procedure TvgrRangesFormat.SaveToStream(AStream: TStream);
var
  AStringList: TStrings;
  AHasFlags: TvgrRangesFlags;
begin
  AStringList := TStringList.Create;
  with AStringList do
  try
    AHasFlags := HasFlags;
    if vgrrfFontName in AHasFlags then
      SavePropertyToList(AStringList, 'FontName', FontName);
    if vgrrfFontSize in AHasFlags then
      SavePropertyToList(AStringList, 'FontSize', FontSize);
    if vgrrfFontStyleBold in AHasFlags then
      SavePropertyToList(AStringList, 'FontStyleBold', FontStyleBold);
    if vgrrfFontStyleItalic in AHasFlags then
      SavePropertyToList(AStringList, 'FontStyleItalic', FontStyleItalic);
    if vgrrfFontStyleUnderline in AHasFlags then
      SavePropertyToList(AStringList, 'FontStyleUnderline', FontStyleUnderline);
    if vgrrfFontStyleStrikeout in AHasFlags then
      SavePropertyToList(AStringList, 'FontStyleStrikeout', FontStyleStrikeout);
    if vgrrfFontCharset in AHasFlags then
      SavePropertyToList(AStringList, 'FontCharset', FontCharset);
    if vgrrfFontColor in AHasFlags then
      SavePropertyToList(AStringList, 'FontColor', FontColor);
    if vgrrfFillBackColor in AHasFlags then
      SavePropertyToList(AStringList, 'FillBackColor', FillBackColor);
    if vgrrfFillForeColor in AHasFlags then
      SavePropertyToList(AStringList, 'FillForeColor', FillForeColor);
    if vgrrfFillPattern in AHasFlags then
      SavePropertyToList(AStringList, 'FillPattern', FillPattern);
    if vgrrfDisplayFormat in AHasFlags then
      SavePropertyToList(AStringList, 'DisplayFormat', DisplayFormat);
    if vgrrfHorzAlign in AHasFlags then
      SavePropertyToList(AStringList, 'HorzAlign', HorzAlign);
    if vgrrfVertAlign in AHasFlags then
      SavePropertyToList(AStringList, 'VertAlign', VertAlign);
    if vgrrfAngle in AHasFlags then
      SavePropertyToList(AStringList, 'Angle', Angle);
    if vgrrfWordWrap in AHasFlags then
      SavePropertyToList(AStringList, 'WordWrap', WordWrap);
    FBorders.SaveToList(AStringList);
    SaveToStream(AStream);
  finally
    AStringList.Free;
  end;
end;

procedure TvgrRangesFormat.LoadFromStream(AStream: TStream);
var
  AStringList: TStrings;
  APropValue: Variant;
begin
  AStringList := TStringList.Create;
  AStringList.LoadFromStream(AStream);
  try
    if LoadPropertyFromList(AStringList, 'FontName', APropValue) then
      FontName := APropValue;
    if LoadPropertyFromList(AStringList, 'FontSize', APropValue) then
      FontSize := APropValue;
    if LoadPropertyFromList(AStringList, 'FontStyleBold', APropValue) then
      FontStyleBold := APropValue;
    if LoadPropertyFromList(AStringList, 'FontStyleItalic', APropValue) then
      FontStyleItalic := APropValue;
    if LoadPropertyFromList(AStringList, 'FontStyleUnderline', APropValue) then
      FontStyleUnderline := APropValue;
    if LoadPropertyFromList(AStringList, 'FontStyleStrikeout', APropValue) then
      FontStyleStrikeout := APropValue;
    if LoadPropertyFromList(AStringList, 'FontCharset', APropValue) then
      FontCharset := APropValue;
    if LoadPropertyFromList(AStringList, 'FontColor', APropValue) then
      FontColor := APropValue;
    if LoadPropertyFromList(AStringList, 'FillBackColor', APropValue) then
      FillBackColor := APropValue;
    if LoadPropertyFromList(AStringList, 'FillForeColor', APropValue) then
      FillForeColor := APropValue;
    if LoadPropertyFromList(AStringList, 'FillPattern', APropValue) then
      FillPattern := APropValue;
    if LoadPropertyFromList(AStringList, 'DisplayFormat', APropValue) then
      DisplayFormat := APropValue;
    if LoadPropertyFromList(AStringList, 'HorzAlign', APropValue) then
      HorzAlign := APropValue;
    if LoadPropertyFromList(AStringList, 'VertAlign', APropValue) then
      VertAlign := APropValue;
    if LoadPropertyFromList(AStringList, 'Angle', APropValue) then
      Angle := APropValue;
    if LoadPropertyFromList(AStringList, 'WordWrap', APropValue) then
      WordWrap := APropValue;
    FBorders.LoadFromList(AStringList);
  finally
    AStringList.Free;
  end;
end;

function TvgrRangesFormat.GetForceBeginUpdateEndUpdate: Boolean;
begin
  Result := FForceBeginUpdateEndUpdate; 
end;

procedure TvgrRangesFormat.SetForceBeginUpdateEndUpdate(Value: Boolean);
begin
  FForceBeginUpdateEndUpdate := Value; 
end;

function TvgrRangesFormat.GetBorders: TvgrRangesFormatBorders;
begin
  Result := FBorders;
end;

procedure TvgrRangesFormat.InitTempObject;
begin
//  ZeroMemory(@FTempObject, SizeOf(FTempObject));
  with FTempObject do
  begin
    Count:= 0;
    Operations := [];
    FontName := '';
    FontSize := 0;
    FontStyle := [];
    FontCharset := DEFAULT_CHARSET;
    FontColor := 0;
    FillBackColor := 0;
    FillForeColor := 0;
    FillPattern := TBrushStyle(0);
    DisplayFormat := '';
    HorzAlign := TvgrRangeHorzAlign(0);
    VertAlign := TvgrRangeVertAlign(0);
    Angle := 0;
    WordWrap := False;
    HasFlags := [];
  end;
end;

procedure TvgrRangesFormat.Execute;
var
  ARectIndex: Integer;
  ACellsCount: Integer;
  ARow, ACol: Integer;
  ARect: TRect;
  ADefRangesStyle: rvgrRangeStyle;
  ARange: IvgrRange;
  ABeginUpdateStarted: Boolean;
begin
  ABeginUpdateStarted := (FTempObject.Operations <> [vgrroGet]) and FForceBeginUpdateEndUpdate;
  if ABeginUpdateStarted then
    FWorksheet.Workbook.BeginUpdate;
  try
    ACellsCount := 0;
    for ARectIndex := 0 to Length(FRectArray)-1 do
    begin
      ARect := FRectArray[ARectIndex];
      if (ARect.Left <= ARect.Right) and (ARect.Top <= ARect.Bottom) then
      begin
        Inc(ACellsCount, (ARect.Bottom - ARect.Top + 1)*(ARect.Right - ARect.Left + 1));
        if FTempObject.Operations = [vgrroGet] then
          FWorksheet.RangesList.FindAndCallBack(ARect, ExecuteCallback, nil)
        else
        begin
          if FAutoCreate then
          begin
            for ACol := ARect.Left to ARect.Right do
              for ARow := ARect.Top to ARect.Bottom do
              begin
                ARange := FWorksheet.RangesList.FindAtCell(ACol, ARow);
                if (ARange = nil) or ((ARange.Left = ACol) and (ARange.Top = ARow)) then
                begin
                  if ARange = nil then
                    ARange := FWorksheet.RangesList.Items[ACol, ARow, ACol, ARow];
                  ExecuteCallbackSet(ARange, 0, nil);
                end;
              end;
          end
          else
            FWorksheet.RangesList.FindAndCallBack(ARect, ExecuteCallbackSet, nil)
        end;
      end;
    end;

    if (FTempObject.Operations = [vgrroGet]) and (FAutoCreate) and (ACellsCount > FTempObject.Count) then
    begin
      FWorksheet.RangesList.Styles.InitStyle(@ADefRangesStyle);
      with ADefRangesStyle do
      begin
        if FTempObject.Count = 0 then
        begin
          FTempObject.FontName := Font.Name;
          FTempObject.FontSize := Font.Size;
          FTempObject.FontStyle := Font.Style;
          FTempObject.FontCharset := Font.Charset;
          FTempObject.FontColor := Font.Color;
          FTempObject.FillBackColor := FillBackColor;
          FTempObject.FillForeColor := FillForeColor;
          FTempObject.FillPattern := FillPattern;
          FTempObject.DisplayFormat := '';
          FTempObject.HorzAlign := HorzAlign;
          FTempObject.VertAlign := VertAlign;
          FTempObject.Angle := Angle;
          FTempObject.WordWrap := False;
          FTempObject.HasFlags :=
             [vgrrfFontName,
              vgrrfFontSize,
              vgrrfFontStyleBold,
              vgrrfFontStyleItalic,
              vgrrfFontStyleUnderline,
              vgrrfFontStyleStrikeout,
              vgrrfFontColor,
              vgrrfFillBackColor,
              vgrrfFillForeColor,
              vgrrfFillPattern,
              vgrrfDisplayFormat,
              vgrrfHorzAlign,
              vgrrfVertAlign,
              vgrrfAngle,
              vgrrfWordWrap,
              vgrrfFontCharset];
        end
        else
        begin
          if FTempObject.FontName <> Font.Name then
            Exclude(FTempObject.HasFlags,vgrrfFontName);
          if FTempObject.FontSize <> Font.Size then
            Exclude(FTempObject.HasFlags,vgrrfFontSize);
          if (fsBold in FTempObject.FontStyle) <> (fsBold in Font.Style) then
            Exclude(FTempObject.HasFlags,vgrrfFontStyleBold);
          if (fsItalic in FTempObject.FontStyle) <> (fsItalic in Font.Style) then
            Exclude(FTempObject.HasFlags,vgrrfFontStyleItalic);
          if (fsUnderline in FTempObject.FontStyle) <> (fsUnderline in Font.Style) then
            Exclude(FTempObject.HasFlags,vgrrfFontStyleUnderline);
          if (fsStrikeout in FTempObject.FontStyle) <> (fsStrikeout in Font.Style) then
            Exclude(FTempObject.HasFlags,vgrrfFontStyleStrikeout);
          if FTempObject.FontCharset <> Font.Charset then
            Exclude(FTempObject.HasFlags, vgrrfFontCharset);
          if FTempObject.FontColor <> Font.Color then
            Exclude(FTempObject.HasFlags,vgrrfFontColor);
          if FTempObject.FillBackColor <> FillBackColor then
            Exclude(FTempObject.HasFlags,vgrrfFillBackColor);
          if FTempObject.FillForeColor <> FillForeColor then
            Exclude(FTempObject.HasFlags,vgrrfFillForeColor);
          if FTempObject.FillPattern <> FillPattern then
            Exclude(FTempObject.HasFlags,vgrrfFillPattern);
          if FTempObject.DisplayFormat <> '' then
            Exclude(FTempObject.HasFlags,vgrrfDisplayFormat);
          if FTempObject.HorzAlign <> HorzAlign then
            Exclude(FTempObject.HasFlags,vgrrfHorzAlign);
          if FTempObject.VertAlign <> VertAlign then
            Exclude(FTempObject.HasFlags,vgrrfVertAlign);
          if FTempObject.Angle <> Angle then
            Exclude(FTempObject.HasFlags,vgrrfAngle);
          if FTempObject.WordWrap <> False then
            Exclude(FTempObject.HasFlags,vgrrfWordWrap);
        end;
        Inc(FTempObject.Count);
      end;
    end;
  finally
    if ABeginUpdateStarted then
      FWorksheet.Workbook.EndUpdate;
  end;
end;

procedure TvgrRangesFormat.ExecuteCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrRange do
  begin
    if FTempObject.Count = 0 then
    begin
      FTempObject.FontName := Font.Name;
      FTempObject.FontSize := Font.Size;
      FTempObject.FontStyle := Font.Style;
      FTempObject.FontCharset := Font.Charset;
      FTempObject.FontColor := Font.Color;
      FTempObject.FillBackColor := FillBackColor;
      FTempObject.FillForeColor := FillForeColor;
      FTempObject.FillPattern := FillPattern;
      FTempObject.DisplayFormat := DisplayFormat;
      FTempObject.HorzAlign := HorzAlign;
      FTempObject.VertAlign := VertAlign;
      FTempObject.Angle := Angle;
      FTempObject.WordWrap := WordWrap;
      FTempObject.HasFlags :=
         [vgrrfFontName,
          vgrrfFontSize,
          vgrrfFontStyleBold,
          vgrrfFontStyleItalic,
          vgrrfFontStyleUnderline,
          vgrrfFontStyleStrikeout,
          vgrrfFontColor,
          vgrrfFillBackColor,
          vgrrfFillForeColor,
          vgrrfFillPattern,
          vgrrfDisplayFormat,
          vgrrfHorzAlign,
          vgrrfVertAlign,
          vgrrfAngle,
          vgrrfWordWrap,
          vgrrfFontCharset];
    end
    else
    begin
      if FTempObject.FontName <> Font.Name then
        Exclude(FTempObject.HasFlags,vgrrfFontName);
      if FTempObject.FontSize <> Font.Size then
        Exclude(FTempObject.HasFlags,vgrrfFontSize);
      if (fsBold in FTempObject.FontStyle) <> (fsBold in Font.Style) then
        Exclude(FTempObject.HasFlags,vgrrfFontStyleBold);
      if (fsItalic in FTempObject.FontStyle) <> (fsItalic in Font.Style) then
        Exclude(FTempObject.HasFlags,vgrrfFontStyleItalic);
      if (fsUnderline in FTempObject.FontStyle) <> (fsUnderline in Font.Style) then
        Exclude(FTempObject.HasFlags,vgrrfFontStyleUnderline);
      if (fsStrikeout in FTempObject.FontStyle) <> (fsStrikeout in Font.Style) then
        Exclude(FTempObject.HasFlags,vgrrfFontStyleStrikeout);
      if FTempObject.FontCharset <> Font.Charset then
        Exclude(FTempObject.HasFlags, vgrrfFontCharset);
      if FTempObject.FontColor <> Font.Color then
        Exclude(FTempObject.HasFlags,vgrrfFontColor);
      if FTempObject.FillBackColor <> FillBackColor then
        Exclude(FTempObject.HasFlags,vgrrfFillBackColor);
      if FTempObject.FillForeColor <> FillForeColor then
        Exclude(FTempObject.HasFlags,vgrrfFillForeColor);
      if FTempObject.FillPattern <> FillPattern then
        Exclude(FTempObject.HasFlags,vgrrfFillPattern);
      if FTempObject.DisplayFormat <> DisplayFormat then
        Exclude(FTempObject.HasFlags,vgrrfDisplayFormat);
      if FTempObject.HorzAlign <> HorzAlign then
        Exclude(FTempObject.HasFlags,vgrrfHorzAlign);
      if FTempObject.VertAlign <> VertAlign then
        Exclude(FTempObject.HasFlags,vgrrfVertAlign);
      if FTempObject.Angle <> Angle then
        Exclude(FTempObject.HasFlags,vgrrfAngle);
      if FTempObject.WordWrap <> WordWrap then
        Exclude(FTempObject.HasFlags,vgrrfWordWrap);
    end;
    Inc(FTempObject.Count, (Right - Left + 1) * (Bottom - Top + 1));
  end;
end;

procedure TvgrRangesFormat.ExecuteCallbackSet(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with (AItem as IvgrRange) do
  begin
    if vgrroSetFontName in FTempObject.Operations then
      Font.Name := FTempObject.FontName;
    if vgrroSetFontSize in FTempObject.Operations then
      Font.Size := FTempObject.FontSize;
    if vgrroSetFontStyleBold in FTempObject.Operations then
      if fsBold in FTempObject.FontStyle then
        Font.Style := Font.Style + [fsBold]
      else
        Font.Style := Font.Style - [fsBold];
    if vgrroSetFontStyleItalic in FTempObject.Operations then
      if fsItalic in FTempObject.FontStyle then
        Font.Style := Font.Style + [fsItalic]
      else
        Font.Style := Font.Style - [fsItalic];
    if vgrroSetFontStyleUnderline in FTempObject.Operations then
      if fsUnderline in FTempObject.FontStyle then
        Font.Style := Font.Style + [fsUnderline]
      else
        Font.Style := Font.Style - [fsUnderline];
    if vgrroSetFontStyleStrikeout in FTempObject.Operations then
      if fsStrikeout in FTempObject.FontStyle then
        Font.Style := Font.Style + [fsStrikeout]
      else
        Font.Style := Font.Style - [fsStrikeout];
    if vgrroSetFontCharset in FTempObject.Operations then
      Font.Charset := FTempObject.FontCharset;
    if vgrroSetFontColor in FTempObject.Operations then
      Font.Color := FTempObject.FontColor;
    if vgrroSetFillBackColor in FTempObject.Operations then
      FillBackColor := FTempObject.FillBackColor;
    if vgrroSetFillForeColor in FTempObject.Operations then
      FillForeColor := FTempObject.FillForeColor;
    if vgrroSetFillPattern in FTempObject.Operations then
      FillPattern := FTempObject.FillPattern;
    if vgrroSetDisplayFormat in FTempObject.Operations then
      DisplayFormat := FTempObject.DisplayFormat;
    if vgrroSetHorzAlign in FTempObject.Operations then
      HorzAlign := FTempObject.HorzAlign;
    if vgrroSetVertAlign in FTempObject.Operations then
      VertAlign := FTempObject.VertAlign;
    if vgrroSetAngle in FTempObject.Operations then
      Angle := FTempObject.Angle;
    if vgrroSetWordWrap in FTempObject.Operations then
      WordWrap := FTempObject.WordWrap;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrClipboardList
//
/////////////////////////////////////////////////
constructor TvgrClipboardList.Create(AItemSize: Integer; AConvertProc: TvgrConvertRecordProc);
begin
  inherited Create;
  FItemSize := AItemSize;
  FConvertProc := AConvertProc;
  FData := nil;
  FCount := 0;
end;

destructor TvgrClipboardList.Destroy;
begin
  FreeMem(FData);
  inherited;
end;

function TvgrClipboardList.AddItem(ASource: Pointer): Pointer;
begin
  FCount := FCount + 1;
  ReallocMem(FData, ItemSize * FCount);
  Result := PChar(FData) + ItemSize * (FCount - 1);
  System.Move(ASource^, Result^, ItemSize);
end;

procedure TvgrClipboardList.SaveToStream(AStream: TStream);
begin
  AStream.Write(FItemSize, 4);
  AStream.Write(FCount, 4);
  AStream.Write(FData^, FItemSize * FCount);
end;

procedure TvgrClipboardList.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
var
  I, AOldItemSize: Integer;
  ARecBuf: Pointer;
begin
  AStream.Read(AOldItemSize, 4);
  AStream.Read(FCount, 4);
  ReallocMem(FData, ItemSize * FCount);
  if ADataStorageVersion = vgrDataStorageVersion then
    AStream.Read(FData^, ItemSize * FCount)
  else
  begin
    FillChar(FData^, FCount * ItemSize, #0);
    GetMem(ARecBuf, AOldItemSize);
    try
      for I := 0 to FCount - 1 do
      begin
        AStream.Read(ARecBuf^, AOldItemSize);
        ConvertProc(ARecBuf, PChar(FData) + I * ItemSize, AOldItemSize, ItemSize, ADataStorageVersion);
      end;
    finally
      FreeMem(ARecBuf);
    end;
  end;
end;

function TvgrClipboardList.GetItem(Index: Integer): Pointer;
begin
  Result := PChar(FData) + Index * ItemSize;
end;

/////////////////////////////////////////////////
//
// TvgrClipboardObject
//
/////////////////////////////////////////////////
constructor TvgrClipboardObject.Create(AWorkbook: TvgrWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  FRanges := TvgrClipboardList.Create(SizeOf(rvgrRange), TvgrRanges.LoadFromStreamConvertRecord);
  FBorders := TvgrClipboardList.Create(SizeOf(rvgrBorder), TvgrBorders.LoadFromStreamConvertRecord);
  FCols := TvgrClipboardList.Create(SizeOf(rvgrCol), TvgrCols.LoadFromStreamConvertRecord);
  FRows := TvgrClipboardList.Create(SizeOf(rvgrRow), TvgrRows.LoadFromStreamConvertRecord);
  FWBStrings := TvgrWBStrings.Create;
  FRangeStyles := TvgrRangeStylesList.Create;
  FBorderStyles := TvgrBorderStylesList.Create;
  FFormulas := TvgrFormulasList.Create(FWorkbook);
  FFormulas.FWBStrings := FWBStrings;
end;

destructor TvgrClipboardObject.Destroy;
begin
  FFormulas.Free;
  FRanges.Free;
  FBorders.Free;
  FCols.Free;
  FRows.Free;
  FRangeStyles.Free;
  FBorderStyles.Free;
  FWBStrings.Free;
  inherited;
end;

function TvgrClipboardObject.GetRangeCount: Integer;
begin
  Result := FRanges.Count;
end;

function TvgrClipboardObject.GetRange(Index: Integer): pvgrRange;
begin
  Result := pvgrRange(FRanges[Index]);
end;

function TvgrClipboardObject.GetBorderCount: Integer;
begin
  Result := FBorders.Count;
end;

function TvgrClipboardObject.GetBorder(Index: Integer): pvgrBorder;
begin
  Result := pvgrBorder(FBorders[Index]);
end;

function TvgrClipboardObject.GetColCount: Integer;
begin
  Result := FCols.Count;
end;

function TvgrClipboardObject.GetCol(Index: Integer): pvgrCol;
begin
  Result := pvgrCol(FCols[Index]);
end;

function TvgrClipboardObject.GetRowCount: Integer;
begin
  Result := FRows.Count;
end;

function TvgrClipboardObject.GetRow(Index: Integer): pvgrRow;
begin
  Result := pvgrRow(FRows[Index]);
end;

procedure TvgrClipboardObject.AddRange(ARange: pvgrRange);
var
  ANewRange: pvgrRange;
  ANewStyle: rvgrRangeStyle;
begin
  ANewRange := FRanges.AddItem(ARange);

  ANewStyle := Workbook.RangeStyles[ANewRange.Style]^;
  if ANewStyle.DisplayFormat <> -1 then
    ANewStyle.DisplayFormat := WBStrings.FindOrAdd(Workbook.WBStrings[ANewStyle.DisplayFormat]);
  ANewRange.Style := RangeStyles.FindOrAdd(ANewStyle, -1);
  
  if ANewRange.Value.ValueType = rvtString then
    ANewRange.Value.vString := WBStrings.FindOrAdd(Workbook.WBStrings[ANewRange.Value.vString]);

  if ANewRange.Formula <> -1 then
  begin
    ANewRange.Formula := Formulas.FindOrAdd(Workbook.Formulas.Formulas[ANewRange.Formula],
                                            Workbook.Formulas.FormulasSize[ANewRange.Formula],
                                            -1,
                                            Workbook.WBStrings);
  end;
end;

procedure TvgrClipboardObject.AddBorder(ABorder: pvgrBorder);
var
  ANewBorder: pvgrBorder; 
begin
  ANewBorder := FBorders.AddItem(ABorder);
  ANewBorder.Style := BorderStyles.FindOrAdd(Workbook.BorderStyles[ANewBorder.Style]^, -1);
end;

procedure TvgrClipboardObject.AddCol(ACol: pvgrCol);
begin
  FCols.AddItem(ACol);
end;

procedure TvgrClipboardObject.AddRow(ARow: pvgrRow);
begin
  FRows.AddItem(ARow);
end;

procedure TvgrClipboardObject.SaveToStream(AStream: TStream);
var
  Buf: Integer;
begin
  Buf := vgrDataStorageVersion;
  AStream.Write(Buf, 4);
  AStream.Write(FCopyPoint, SizeOf(TPoint));
  AStream.Write(FCopySize, SizeOf(TSize));

  FRanges.SaveToStream(AStream);
  FBorders.SaveToStream(AStream);
  FCols.SaveToStream(AStream);
  FRows.SaveToStream(AStream);
  FWBStrings.SaveToStream(AStream);
  FRangeStyles.SaveToStream(AStream);
  FBorderStyles.SaveToStream(AStream);
  FFormulas.SaveToStream(AStream);
end;

procedure TvgrClipboardObject.LoadFromStream(AStream: TStream);
var
  ALoadedDataStorageVersion: Integer;
begin
  AStream.Read(ALoadedDataStorageVersion, 4);
  AStream.Read(FCopyPoint, SizeOf(TPoint));
  AStream.Read(FCopySize, SizeOf(TSize));

  FRanges.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FBorders.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FCols.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FRows.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FWBStrings.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FRangeStyles.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FBorderStyles.LoadFromStream(AStream, ALoadedDataStorageVersion);
  FFormulas.LoadFromStream(AStream, ALoadedDataStorageVersion);
end;

procedure TvgrClipboardObject.CopyToClipboard;
var
  AStream: TMemoryStream;
begin
  AStream := TMemoryStream.Create;
  try
    SaveToStream(AStream);
    StreamToClipboard(vgrCF_WORKSHEET_RECT, AStream);
  finally
    AStream.Free;
  end;
end;

procedure TvgrClipboardObject.PasteFromClipboard;
var
  AStream: TMemoryStream;
begin
  AStream := TMemoryStream.Create;
  try
    StreamFromClipboard(vgrCF_WORKSHEET_RECT, AStream);
    AStream.Seek(soFromBeginning, 0);
    LoadFromStream(AStream);
  finally
    AStream.Free;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrWorksheet
//
/////////////////////////////////////////////////
constructor TvgrWorksheet.Create(AOwner: TComponent);
begin
  inherited;
  FRanges := TvgrRanges.Create(Self);
  FCols := TvgrCols.Create(Self);
  FRows := TvgrRows.Create(Self);
  FBorders := TvgrBorders.Create(Self);
  FHorzSections := GetSectionsClass.Create(Self);
  FHorzSections.InitNotifyEvents(vgrwcNewHorzSection, vgrwcDeleteHorzSection, vgrwcChangeHorzSection);
  FVertSections := GetSectionsClass.Create(Self);
  FVertSections.InitNotifyEvents(vgrwcNewVertSection, vgrwcDeleteVertSection, vgrwcChangeVertSection);
  FMergeList := TInterfaceList.Create;
  FPageProperties := TvgrPageProperties.Create(OnPagePropertiesChange);
  FDimensions := Rect(0,0,0,0);
  FDimensionsValid := True;
  FDimensionsEmpty := True;
  ClearExportData;
end;

destructor TvgrWorksheet.Destroy;
var
  ChangeInfo: TvgrWorkbookChangeInfo;
begin
  if Workbook <> nil then
  begin
    ChangeInfo.ChangesType := vgrwcDeleteWorksheet;
    ChangeInfo.ChangedObject := Self;
    ChangeInfo.ChangedInterface := nil;
    Workbook.BeforeChange(ChangeInfo);
  end;

  if FWorksheets <> nil then
    FWorksheets.RemoveWorksheet(Self);
  FRanges.Free;
  FCols.Free;
  FRows.Free;
  FBorders.Free;
  FHorzSections.Free;
  FVertSections.Free;

  if Workbook <> nil then
  begin
    ChangeInfo.ChangedObject := nil;
    Workbook.AfterChange(ChangeInfo);
  end;

  FreeAndNil(FMergeList);
  FPageProperties.Free;

  if FCalculateCanvas <> nil then
  begin
    ReleaseDC(0, FCalculateCanvas.Handle);
    FCalculateCanvas.Free;
    FCalculateCanvas := nil;
  end; 

  inherited;
end;

// protected
function TvgrWorksheet.GetSectionsClass: TvgrSectionsClass;
begin
  Result := TvgrSections;
end;

function TvgrWorksheet.GetChildOwner: TComponent;
begin
  Result := Owner;
end;

procedure TvgrWorksheet.DefineProperties(Filer: TFiler);
begin
  inherited DefineProperties(Filer);
  Filer.DefineBinaryProperty('BordersData', ReadBordersData, WriteBordersData,
    (BordersCount > 0));
  Filer.DefineBinaryProperty('ColsData', ReadColsData, WriteColsData,
    (ColsCount > 0));
  Filer.DefineBinaryProperty('RowsData', ReadRowsData, WriteRowsData,
    (RowsCount > 0));
  Filer.DefineBinaryProperty('HorzSectionsData', ReadHorzSectionsData, WriteHorzSectionsData,
    (HorzSectionCount > 0));
  Filer.DefineBinaryProperty('VertSectionsData', ReadVertSectionsData, WriteVertSectionsData,
    (VertSectionCount > 0));
  Filer.DefineBinaryProperty('RangesData', ReadRangesData, WriteRangesData,
    (RangesCount > 0));
end;

function TvgrWorksheet.GetParentComponent: TComponent;
begin
  Result := Workbook;
end;

function TvgrWorksheet.HasParent: Boolean;
begin
  Result := True;
end;

procedure TvgrWorksheet.SetParentComponent(Value: TComponent);
begin
  if Value is TvgrWorkbook then
  begin
    FWorksheets := TvgrWorkbook(Value).FWorksheets;
    FWorksheets.AddInList(Self);
  end;
end;

function TvgrWorksheet.GetDispIdOfName(const AName : String) : Integer;
begin
  if AnsiCompareText(AName,'Cols') = 0 then
    Result := cs_TvgrWorksheet_Cols
  else
  if AnsiCompareText(AName,'Rows') = 0 then
    Result := cs_TvgrWorksheet_Rows
  else
  if AnsiCompareText(AName,'Borders') = 0 then
    Result := cs_TvgrWorksheet_Borders
  else
  if AnsiCompareText(AName,'Ranges') = 0 then
    Result := cs_TvgrWorksheet_Ranges
  else
  if AnsiCompareText(AName,'HorzSections') = 0 then
    Result := cs_TvgrWorksheet_HorzSections
  else
  if AnsiCompareText(AName,'VertSections') = 0 then
    Result := cs_TvgrWorksheet_VertSections
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrWorksheet.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrWorksheet, siTvgrWorksheetLength);
end;

function TvgrWorksheet.DoInvoke (DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrWorksheet_Cols:
      AResult := Cols[AParameters[0]] as IDispatch;
    cs_TvgrWorksheet_Rows:
      AResult := Rows[AParameters[0]] as IDispatch;
    cs_TvgrWorksheet_Borders:
      AResult := Borders[AParameters[0], AParameters[1], AParameters[2]] as IDispatch;
    cs_TvgrWorksheet_Ranges:
      AResult := Ranges[AParameters[0], AParameters[1], AParameters[2], AParameters[3]] as IDispatch;
    cs_TvgrWorksheet_HorzSections:
      AResult := HorzSections[AParameters[0], AParameters[1]] as IDispatch;
    cs_TvgrWorksheet_VertSections:
      AResult := HorzSections[AParameters[0], AParameters[1]] as IDispatch;
    else
      Result := inherited DoInvoke(DispID, Flags, AParameters, AResult);
  end;
end;

procedure TvgrWorksheet.OnPagePropertiesChange(Sender: TObject);
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := vgrwcChangeWorksheetPageProperties;
  AChangeInfo.ChangedObject := Self;
  AChangeInfo.ChangedInterface := nil;
  FWorksheets.AfterChange(AChangeInfo);
end;

procedure TvgrWorksheet.WriteHorzSectionsData(Stream: TStream);
begin
  FHorzSections.SaveToStream(Stream);
end;

procedure TvgrWorksheet.CheckMinDimensions;
begin
  if FDimensions.Left < 0 then
    FDimensions.Left := 0;
  if FDimensions.Right < 0 then
    FDimensions.Right := 0;
  if FDimensions.Top < 0 then
    FDimensions.Top := 0;
  if FDimensions.Bottom < 0 then
    FDimensions.Bottom := 0;
end;

procedure TvgrWorksheet.ValidateDimensions;
var
  I: Integer;
begin
  FDimensionsEmpty := True;
  FDimensionsValid := True;
  FDimensions := Rect(0, 0, 0, 0);
  for I := 0 to RangesList.Count - 1 do
    RangeCheckDimensions(RangesList.ByIndex[I].Place);
  for I := 0 to BordersList.Count - 1 do
    with BordersList.ByIndex[I] do
      BorderCheckDimensions(Left, Top, Orientation);
end;

procedure TvgrWorksheet.RangeCheckDimensions(const APlace: TRect);
begin
  if FDimensionsValid then
  begin
    if FDimensionsEmpty then
    begin
      FDimensionsEmpty := False;
      FDimensions := APlace;
    end
    else
    begin
      if FDimensions.Left > APlace.Left then
        FDimensions.Left := APlace.Left;
      if FDimensions.Top > APlace.Top then
        FDimensions.Top := APlace.Top;
      if FDimensions.Right < APlace.Right then
        FDimensions.Right := APlace.Right;
      if FDimensions.Bottom < APlace.Bottom then
        FDimensions.Bottom := APlace.Bottom;
     end;
  end;
end;

procedure TvgrWorksheet.BorderCheckDimensions(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation);
begin
  if FDimensionsValid then
  begin
    if FDimensionsEmpty then
    begin
      FDimensionsEmpty := False;
      FDimensions.Left := ALeft;
      FDimensions.Right := ALeft;
      FDimensions.Top := ATop;
      FDimensions.Bottom := ATop;
    end
    else
    begin
      if FDimensions.Left > ALeft then
        FDimensions.Left := ALeft;
      if FDimensions.Top > ATop then
        FDimensions.Top := ATop;
      if AOrientation = vgrboLeft then
      begin
        if FDimensions.Right < ALeft - 1 then
          FDimensions.Right := ALeft - 1;
        if FDimensions.Bottom < ATop then
          FDimensions.Bottom := ATop;
      end
      else
      begin
        if FDimensions.Right < ALeft then
          FDimensions.Right := ALeft;
        if FDimensions.Bottom < ATop - 1 then
          FDimensions.Bottom := ATop - 1;
      end;
    end;
  end;
end;

procedure TvgrWorksheet.RangeOrBorderDeleted;
begin
  FDimensionsValid := False;
end;

procedure TvgrWorksheet.NewRowCheckDimensions(AIndexBefore, ACount: Integer);
begin
  if (FDimensionsValid) then
  begin
    if (AIndexBefore <= FDimensions.Top) then
      Inc(FDimensions.Top, ACount);
    if (AIndexBefore <= FDimensions.Bottom) then
      Inc(FDimensions.Bottom, ACount);
  end;
end;

procedure TvgrWorksheet.NewColCheckDimensions(AIndexBefore, ACount: Integer);
begin
  if (FDimensionsValid) then
  begin
    if (AIndexBefore <= FDimensions.Left) then
      Inc(FDimensions.Left, ACount);
    if (AIndexBefore <= FDimensions.Right) then
      Inc(FDimensions.Right, ACount);
  end;
end;

procedure TvgrWorksheet.DelColsCheckDimensions(AStartPos, AEndPos: Integer);
begin
  if (FDimensionsValid) then
  begin
    if (AStartPos <= FDimensions.Left) then
      Dec(FDimensions.Left, Min(AEndPos-AStartPos+1, FDimensions.Left-AStartPos+1));
    if (AStartPos <= FDimensions.Right) then
      Dec(FDimensions.Right, Min(AEndPos-AStartPos+1, FDimensions.Right-AStartPos+1));
  end;
  CheckMinDimensions;
end;

procedure TvgrWorksheet.DelRowsCheckDimensions(AStartPos, AEndPos: Integer);
begin
  if (FDimensionsValid) then
  begin
    if (AStartPos <= FDimensions.Top) then
      Dec(FDimensions.Top, Min(AEndPos-AStartPos+1, FDimensions.Top-AStartPos+1));
    if (AStartPos <= FDimensions.Bottom) then
      Dec(FDimensions.Bottom, Min(AEndPos-AStartPos+1, FDimensions.Bottom-AStartPos+1));
  end;
  CheckMinDimensions;
end;

procedure TvgrWorksheet.ReadHorzSectionsData(Stream: TStream);
begin
  FHorzSections.LoadFromStream(Stream, Workbook.DataStorageVersion);
end;

procedure TvgrWorksheet.WriteVertSectionsData(Stream: TStream);
begin
  FVertSections.SaveToStream(Stream);
end;

procedure TvgrWorksheet.ReadVertSectionsData(Stream: TStream);
begin
  FVertSections.LoadFromStream(Stream, Workbook.DataStorageVersion);
end;

procedure TvgrWorksheet.ReadRowsData(Stream: TStream);
begin
  FRows.LoadFromStream(Stream, Workbook.DataStorageVersion);
end;

procedure TvgrWorksheet.WriteRowsData(Stream: TStream);
begin
  FRows.SaveToStream(Stream);
end;

procedure TvgrWorksheet.ReadColsData(Stream: TStream);
begin
  FCols.LoadFromStream(Stream, Workbook.DataStorageVersion);
end;

procedure TvgrWorksheet.WriteColsData(Stream: TStream);
begin
  FCols.SaveToStream(Stream);
end;

procedure TvgrWorksheet.ReadBordersData(Stream: TStream);
begin
  FBorders.LoadFromStream(Stream, Workbook.DataStorageVersion);
end;

procedure TvgrWorksheet.WriteBordersData(Stream: TStream);
begin
  FBorders.SaveToStream(Stream);
end;

procedure TvgrWorksheet.ReadRangesData(Stream: TStream);
begin
  FRanges.LoadFromStream(Stream, Workbook.DataStorageVersion);
end;

procedure TvgrWorksheet.WriteRangesData(Stream: TStream);
begin
  FRanges.SaveToStream(Stream);
end;

function TvgrWorksheet.GetWorkbook: TvgrWorkbook;
begin
  Result := FWorksheets.Workbook;
end;

procedure TvgrWorksheet.SetTitle(const Value : string);
begin
  if FTitle <> Value then
  begin
    BeforeChangeProperty;
    FTitle := Value;
    AfterChangeProperty;
  end;
end;

function TvgrWorksheet.GetCol(ColNumber : integer) : IvgrCol;
begin
  Result := FCols[ColNumber];
end;

function TvgrWorksheet.GetColByIndex(Index : integer) : IvgrCol;
begin
  Result := FCols.ByIndex[Index];
end;

function TvgrWorksheet.GetColsCount : integer;
begin
  Result := FCols.Count;
end;

function TvgrWorksheet.GetRow(RowNumber : integer) : IvgrRow;
begin
  Result := FRows[RowNumber];
end;

function TvgrWorksheet.GetRowByIndex(Index : integer) : IvgrRow;
begin
  Result := FRows.ByIndex[Index];
end;

function TvgrWorksheet.GetRowsCount : integer;
begin
  Result := FRows.Count;
end;

function TvgrWorksheet.GetRange(Left,Top,Right,Bottom : integer) : IvgrRange;
begin
  Result := FRanges[Left,Top,Right,Bottom];
end;

function TvgrWorksheet.GetRangeByIndex(Index : integer) : IvgrRange;
begin
  Result := FRanges.ByIndex[Index];
end;

function TvgrWorksheet.GetRangesCount : integer;
begin
  Result := FRanges.Count;
end;

procedure TvgrWorksheet.InsertRows(AIndexBefore, ACount: Integer);
begin
  BeginUpdate;
  try
    RowsList.Insert(AIndexBefore, ACount);
    HorzSectionsList.InternalInsertLines(AIndexBefore, ACount);
    BordersList.InternalInsertLines(AIndexBefore, ACount);
    RangesList.InternalInsertLines(AIndexBefore, ACount);
    NewRowCheckDimensions(AIndexBefore, ACount);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.InsertCols(AIndexBefore, ACount: Integer);
begin
  BeginUpdate;
  try
    ColsList.Insert(AIndexBefore, ACount);
    VertSectionsList.InternalInsertLines(AIndexBefore, ACount);
    BordersList.InternalInsertCols(AIndexBefore, ACount);
    RangesList.InternalInsertCols(AIndexBefore, ACount);
    NewColCheckDimensions(AIndexBefore, ACount);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.DeleteRows(AStartPos, AEndPos: Integer);
begin
  BeginUpdate;
  try
    RowsList.Delete(AStartPos, AEndPos);
    BordersList.DeleteRows(AStartPos, AEndPOs);
    RangesList.DeleteRows(AStartPos, AEndPos);
    HorzSectionsList.DeleteRows(AStartPos, AEndPos);
    DelRowsCheckDimensions(AStartPos, AEndPos);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.DeleteCols(AStartPos, AEndPos: Integer);
begin
  BeginUpdate;
  try
    ColsList.Delete(AStartPos, AEndPos);
    BordersList.DeleteCols(AStartPos, AEndPOs);
    RangesList.DeleteCols(AStartPos, AEndPos);
    VertSectionsList.DeleteRows(AStartPos, AEndPos);
    DelColsCheckDimensions(AStartPos, AEndPos);
  finally
    EndUpdate;
  end;
end;

function TvgrWorksheet.PerformDeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean): Boolean;
begin
  Result := RangesList.PerformDeleteCells(ARect, ACellShift, ABreakRanges);
end;

function TvgrWorksheet.DeleteCells(const ARect: TRect; ACellShift: TvgrCellShift; ABreakRanges: Boolean): TvgrCellShiftState;
begin
  BeginUpdate;
  try
    if RangesList.PerformDeleteCells(ARect, ACellShift, ABreakRanges) then
    begin
      BordersList.DeleteCells(ARect, ACellShift);
      RangesList.DeleteCells(ARect, ACellShift, ABreakRanges);
      Result := vgrcsSuccess;
    end
    else
      Result := vgrcsLargeRange;
  finally
    EndUpdate;
  end;
end;

function TvgrWorksheet.GetRangesFormat(ARects: TvgrRectArray; AAutoCreate: Boolean): IvgrRangesFormat;
begin
  Result := TvgrRangesFormat.Create(ARects, Self, AAutoCreate) as IvgrRangesFormat;
end;

procedure TvgrWorksheet.MergeFindMergedCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrRange do
    if (Left <> Right) or (Top <> Bottom) then
      FMergeList.Add(AItem as IvgrRange);
end;

procedure TvgrWorksheet.MergeFindAllCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  FMergeList.Add(AItem as IvgrRange);
end;

procedure TvgrWorksheet.Merge(const AMergeRect: TRect; AWarningProc: TvgrMergeWarningProc = nil);
var
  I: Integer;
  ARange, ATopLeftRange: IvgrRange;
  ARangesFormat: IvgrRangesFormat;
  ARects: TvgrRectArray;
  AFormatStream: TMemoryStream;
  AWarning: Boolean;
begin
  AFormatStream := TMemoryStream.Create;
  try
    FMergeList.Clear;
    RangesList.FindAndCallback(AMergeRect, MergeFindAllCallback, nil);

    AWarning := False;
    if FMergeList.Count > 1 then
    begin
      for I := 0 to FMergeList.Count - 1 do
      begin
        ARange := FMergeList[I] as IvgrRange;
        if ATopLeftRange = nil then
          ATopLeftRange := ARange
        else
        begin
          if VarIsNull(ATopLeftRange.Value) then
          begin
            if not VarIsNull(ARange.Value) or
               ((ATopLeftRange.Top > ARange.Top) or
                ((ATopLeftRange.Top = ARange.Top) and (ATopLeftRange.Left > ARange.Left))) then
              ATopLeftRange := ARange;
          end
          else
          begin
            if not VarIsNull(ARange.Value) then
            begin
              AWarning := True;
              if (ATopLeftRange.Top > ARange.Top) or
                 ((ATopLeftRange.Top = ARange.Top) and (ATopLeftRange.Left > ARange.Left)) then
                ATopLeftRange := ARange;
            end;
          end;
        end;
      end;
      if AWarning and Assigned(AWarningProc) then
        if not AWarningProc then
          exit;
    end
    else
      if FMergeList.Count = 1 then
        ATopLeftRange := FMergeList[0] as IvgrRange
      else
        ATopLeftRange := nil;

    if ATopLeftRange = nil then
      with AMergeRect do
        Ranges[Left, Top, Right, Bottom]
    else
    begin
      SetLength(ARects, 1);
      ARects[0] := AMergeRect;
      ARangesFormat := RangesFormat[ARects, False];
      ARangesFormat.SaveToStream(AFormatStream);
      with AMergeRect do
        Ranges[Left, Top, Right, Bottom].Value := ATopLeftRange.Value;
      AFormatStream.Seek(0, soFromBeginning);
      ARangesFormat.LoadFromStream(AFormatStream);
    end;

  finally
    AFormatStream.Free;
    FMergeList.Clear;
  end;
end;

procedure TvgrWorksheet.UnMerge(const AMergeRect: TRect);
var
  I: Integer;
  ARangeValue: Variant;
  ARects: TvgrRectArray;
  AFormatStream: TMemoryStream;
begin
  AFormatStream := TMemoryStream.Create;
  try
    SetLength(ARects, 1);
    FMergeList.Clear;
    RangesList.FindAndCallback(AMergeRect, MergeFindMergedCallback, nil);

    for I := 0 to FMergeList.Count - 1 do
    begin
      ARects[0] := (FMergeList[I] as IvgrRange).Place;
      ARangeValue := (FMergeList[I] as IvgrRange).Value;

      RangesFormat[ARects, False].SaveToStream(AFormatStream);
      with ARects[0] do
        Ranges[Left, Top, Left, Top].Value := ARangeValue;
      AFormatStream.Seek(0, soFromBeginning);
      RangesFormat[ARects, True].LoadFromStream(AFormatStream);
    end;
  finally
    AFormatStream.Free;
    FMergeList.Clear;
  end;
end;

procedure TvgrWorksheet.MergeUnMerge(const AMergeRect: TRect; AWarningProc: TvgrMergeWarningProc = nil);
begin
  if IsMergeInRect(AMergeRect) then
    UnMerge(AMergeRect)
  else
    Merge(AMergeRect, AWarningProc);
end;

function TvgrWorksheet.IsMergeInRect(const ARect: TRect): Boolean;
begin
  FMergeList.Clear;
  RangesList.FindAndCallBack(ARect, MergeFindMergedCallback, nil);
  Result := FMergeList.Count > 0;
  FMergeList.Clear;
end;

function TvgrWorksheet.GetCalculateCanvas: TCanvas;
begin
  if FCalculateCanvas = nil then
  begin
    FCalculateCanvas := TCanvas.Create;
    FCalculateCanvas.Handle := GetDC(0);
  end;
  Result := FCalculateCanvas;
end;

procedure TvgrWorksheet.GetRangePixelsRect(const ARangeRect: TRect; var ARangePixelsRect: TRect);
var
  I: Integer;
  ACol: IvgrCol;
  ARow: IvgrRow;
begin
  ARangePixelsRect := Rect(0, 0, 0, 0);
  with ARangeRect do
  begin
    for I := Left to Right do
    begin
      ACol := ColsList.Find(I);
      if ACol = nil then
        ARangePixelsRect.Right := ARangePixelsRect.Right + PageProperties.Defaults.ColWidth
      else
        ARangePixelsRect.Right := ARangePixelsRect.Right + ACol.Width;
    end;

    for I := Top to Bottom do
    begin
      ARow := RowsList.Find(I);
      if ARow = nil then
        ARangePixelsRect.Bottom := ARangePixelsRect.Bottom + PageProperties.Defaults.RowHeight
      else
        ARangePixelsRect.Bottom := ARangePixelsRect.Bottom + ARow.Height;
    end;
  end;

  ARangePixelsRect.Right := ConvertTwipsToPixelsX(ARangePixelsRect.Right);
  ARangePixelsRect.Bottom := ConvertTwipsToPixelsY(ARangePixelsRect.Bottom);
end;

procedure TvgrWorksheet.GetRangeInternalPixelsRect(const ARangeRect: TRect; var ARangePixelsRect: TRect);

  function GetBorderSize(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation): Integer;
  var
    ABorder: IvgrBorder;
  begin
    ABorder := BordersList.Find(ALeft, ATop, AOrientation);
    if ABorder <> nil then
    begin
      if AOrientation = vgrboLeft then
        Result := Max(1, ConvertTwipsToPixelsY(ABorder.Width))
      else
        Result := Max(1, ConvertTwipsToPixelsX(ABorder.Width));
    end
    else
      Result := 1;
  end;

  procedure CheckBorderSide(X1, Y1, X2, Y2: Integer; AOrientation: TvgrBorderOrientation; ASide: TvgrBorderSide);
  var
    X, Y, ASize, APartSize, AMaxSize: Integer;
  begin
    AMaxSize := GetBorderSize(X1, Y1, AOrientation);
    for X := X1 to X2 do
      for Y := Y1 to Y2 do
      begin
        ASize := GetBorderSize(X1, Y1, AOrientation);
        if AMaxSize <> ASize then
          AMaxSize := ASize;
      end;
    APartSize := GetBorderPartSize(AMaxSize, ASide);

    with ARangePixelsRect do
      case ASide of
        vgrbsLeft: Left := Left + APartSize;
        vgrbsTop: Top := Top + APartSize;
        vgrbsRight: Right := Right - APartSize;
        vgrbsBottom: Bottom := Bottom - APartSize;
      end;
  end;

begin
  GetRangePixelsRect(ARangeRect, ARangePixelsRect);
  with ARangeRect do
  begin
    // Left borders
    CheckBorderSide(Left, Top, Left, Bottom, vgrboLeft, vgrbsLeft);

    // Right borders
    CheckBorderSide(Right + 1, Top, Right + 1, Bottom, vgrboLeft, vgrbsRight);

    // Top borders
    CheckBorderSide(Left, Top, Right, Top, vgrboTop, vgrbsTop);

    // Bottom borders
    CheckBorderSide(Left, Bottom + 1, Right, Bottom + 1, vgrboTop, vgrbsBottom);
  end;
end;

type
  rvgrAutoSizeParams = record
    Size: Integer;
    AutoSizeProc: TvgrRangeCalcSizeProc;
  end;
  pvgrAutoSizeParams = ^rvgrAutoSizeParams;

procedure TvgrWorksheet.GetRowAutoHeightCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ARange: IvgrRange;
  ASize: TSize;
begin
  ARange := AItem as IvgrRange;
  if ARange.Place.Top = ARange.Place.Bottom then
  begin
    ASize := pvgrAutoSizeParams(AData).AutoSizeProc(ARange);
    if ASize.cy > pvgrAutoSizeParams(AData).Size then
      pvgrAutoSizeParams(AData).Size := ASize.cy;
  end;
end;

function TvgrWorksheet.GetRowAutoHeight(ARowNumber: Integer; ACalcRangeSizeProc: TvgrRangeCalcSizeProc): Integer;
var
  AParams: rvgrAutoSizeParams;
begin
  AParams.Size := 0;
  AParams.AutoSizeProc := ACalcRangeSizeProc;
  RangesList.FindAndCallBack(Rect(0, ARowNumber, MaxInt - 1, ARowNumber), GetRowAutoHeightCallback, @AParams);
  Result := AParams.Size;
end;

function TvgrWorksheet.DefaultGetRangeSizeProc(ARange: IvgrRange): TSize;
var
  ARect: TRect;
begin
  GetRangeInternalPixelsRect(ARange.Place, ARect);
  with ARange do
  begin
    Font.AssignTo(CalculateCanvas.Font);
    Result := vgrCalcText(CalculateCanvas,
                          DisplayText,
                          ARect,
                          WordWrap,
                          GetAutoAlignForRangeValue(HorzAlign, ValueType),
                          VertAlign,
                          Angle);
  end;
  Result.cx := ConvertPixelsToTwipsX(Result.cx + 1);
  Result.cy := ConvertPixelsToTwipsY(Result.cy + 1);
end;

function TvgrWorksheet.GetRowAutoHeight(ARowNumber: Integer): Integer;
begin
  Result := GetRowAutoHeight(ARowNumber, DefaultGetRangeSizeProc);
end;

procedure TvgrWorksheet.GetColAutoWidthCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ARange: IvgrRange;
  ASize: TSize;
begin
  ARange := AItem as IvgrRange;
  if ARange.Place.Left = ARange.Place.Right then
  begin
    ASize := pvgrAutoSizeParams(AData).AutoSizeProc(ARange);
    if ASize.cx > pvgrAutoSizeParams(AData).Size then
      pvgrAutoSizeParams(AData).Size := ASize.cx;
  end;
end;

function TvgrWorksheet.GetColAutoWidth(AColNumber: Integer; ACalcRangeSizeProc: TvgrRangeCalcSizeProc): Integer;
var
  AParams: rvgrAutoSizeParams;
begin
  AParams.Size := 0;
  AParams.AutoSizeProc := ACalcRangeSizeProc;
  RangesList.FindAndCallBack(Rect(AColNumber, 0, AColNumber, MaxInt - 1), GetColAutoWidthCallback, @AParams);
  Result := AParams.Size;
end;

function TvgrWorksheet.GetColAutoWidth(AColNumber: Integer): Integer;
begin
  Result := GetColAutoWidth(AColNumber, DefaultGetRangeSizeProc)
end;

function TvgrWorksheet.GetBorder(Left,Top : integer; Orientation : TvgrBorderOrientation) : IvgrBorder;
begin
  Result := FBorders.Items[Left,Top,Orientation];
end;

function TvgrWorksheet.GetBorderByIndex(Index : integer) : IvgrBorder;
begin
  Result := FBorders.ByIndex[Index];
end;

function TvgrWorksheet.GetBordersCount : integer;
begin
  Result := FBorders.Count;
end;

function TvgrWorksheet.GetHorzSection(StartPos, EndPos: Integer): IvgrSection;
begin
  Result := FHorzSections.Items[StartPos, EndPos];
end;

function TvgrWorksheet.GetHorzSectionByIndex(Index: Integer): IvgrSection;
begin
  Result := FHorzSections.ByIndex[Index];
end;

function TvgrWorksheet.GetHorzSectionCount: Integer;
begin
  Result := FHorzSections.Count;
end;

function TvgrWorksheet.GetVertSection(StartPos, EndPos: Integer): IvgrSection;
begin
  Result := FVertSections.Items[StartPos, EndPos];
end;

function TvgrWorksheet.GetVertSectionByIndex(Index: Integer): IvgrSection;
begin
  Result := FVertSections.ByIndex[Index];
end;

function TvgrWorksheet.GetVertSectionCount: Integer;
begin
  Result := FVertSections.Count;
end;

procedure TvgrWorksheet.SetPageProperties(Value: TvgrPageProperties);
begin
  FPageProperties.Assign(Value);
end;

function TvgrWorksheet.GetIndexInWorkbook : integer;
begin
  Result := FWorksheets.IndexOf(Self);
end;

procedure TvgrWorksheet.SetIndexInWorkbook(Value: Integer);
var
  I: Integer;
begin
  I := IndexInWorkbook;
  if I <> Value then
  begin
    BeforeChangeProperty;
    FWorksheets.Move(I, Value);
//    FWorksheets.Exchange(I, Value);
    AfterChangeProperty;
  end;
end;

procedure TvgrWorksheet.SetName(const NewName: TComponentName);
begin
  BeforeChangeProperty;
  inherited;
  AfterChangeProperty;
end;

function TvgrWorksheet.GetDimensions : TRect;
begin
  if not FDimensionsValid then
    ValidateDimensions;
  Result := FDimensions;
end;

procedure TvgrWorksheet.BeforeChangeProperty;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := vgrwcChangeWorksheet;
  AChangeInfo.ChangedObject := Self;
  AChangeInfo.ChangedInterface := nil;
  FWorksheets.BeforeChange(AChangeInfo);
end;

procedure TvgrWorksheet.AfterChangeProperty;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := vgrwcChangeWorksheet;
  AChangeInfo.ChangedObject := Self;
  AChangeInfo.ChangedInterface := nil;
  FWorksheets.AfterChange(AChangeInfo);
end;

procedure TvgrWorksheet.BeginUpdate;
begin
  Inc(Workbook.FLockCount);
end;

procedure TvgrWorksheet.EndUpdate;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  Dec(Workbook.FLockCount);
  if Workbook.FLockCount < 0 then
    Workbook.FLockCount := 0;
  if Workbook.FLockCount = 0 then
  begin
    AChangeInfo.ChangesType := vgrwcChangeWorksheetContent;
    AChangeInfo.ChangedObject := Self;
    AChangeInfo.ChangedInterface := nil;
    FWorksheets.AfterChange(AChangeInfo);
  end;
end;

procedure TvgrWorksheet.SetBorders(const ARect: TRect; ASettedBorders: TvgrSettedBorders; ABorderStyle: TvgrBorderStyle; ABorderWidth: Integer; ABorderColor: TColor);
var
  I, J: Integer;
begin
  BeginUpdate;
  try
    if vgrsbLeft in ASettedBorders then
      for I := ARect.Top to ARect.Bottom do
        Borders[ARect.Left, I, vgrboLeft].SetStyles(ABorderWidth, ABorderColor, ABorderStyle);

    if vgrsbRight in ASettedBorders then
      for I := ARect.Top to ARect.Bottom do
        Borders[ARect.Right + 1, I, vgrboLeft].SetStyles(ABorderWidth, ABorderColor, ABorderStyle);

    if vgrsbTop in ASettedBorders then
      for I := ARect.Left to ARect.Right do
        Borders[I, ARect.Top, vgrboTop].SetStyles(ABorderWidth, ABorderColor, ABorderStyle);

    if vgrsbBottom in ASettedBorders then
      for I := ARect.Left to ARect.Right do
        Borders[I, ARect.Bottom + 1, vgrboTop].SetStyles(ABorderWidth, ABorderColor, ABorderStyle);

    if vgrsbInside in ASettedBorders then
      for I := ARect.Left to ARect.Right do
        for J := ARect.Top to ARect.Bottom do
        begin
          if I <> ARect.Left then
            Borders[I, J, vgrboLeft].SetStyles(ABorderWidth, ABorderColor, ABorderStyle);
          if J <> ARect.Top then
            Borders[I, J, vgrboTop].SetStyles(ABorderWidth, ABorderColor, ABorderStyle);
        end;
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.CopyToClipboardRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  TvgrClipboardObject(AData).AddRange((AItem as IvgrRange).ItemData);
end;

procedure TvgrWorksheet.CopyToClipboardBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  TvgrClipboardObject(AData).AddBorder((AItem as IvgrBorder).ItemData);
end;

procedure TvgrWorksheet.CopyToClipboardColsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  TvgrClipboardObject(AData).AddCol((AItem as IvgrCol).ItemData);
end;

procedure TvgrWorksheet.CopyToClipboardRowsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  TvgrClipboardObject(AData).AddRow((AItem as IvgrRow).ItemData);
end;

procedure TvgrWorksheet.CutToClipboard(const ARects: TvgrRectArray);
begin
  BeginUpdate;
  try
    CopyToClipboard(ARects);
    ClearContents(ARects, [vgrccBorders, vgrccRanges]);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.CopyToClipboard(const ARect: TRect);
var
  AClipboardObject: TvgrClipboardObject;
begin
  AClipboardObject := TvgrClipboardObject.Create(Workbook);
  try
    AClipboardObject.CopyPoint := ARect.TopLeft;
    AClipboardObject.CopySize := Size(ARect.Right - ARect.Left, ARect.Bottom - ARect.Top);

    RangesList.FindAndCallBack(ARect, CopyToClipboardRangesCallback, AClipboardObject);
    BordersList.FindAtRectAndCallBack(ARect, CopyToClipboardBordersCallback, AClipboardObject);
    ColsList.FindAndCallBack(ARect.Left, ARect.Right, CopyToClipboardColsCallback, AClipboardObject);
    RowsList.FindAndCallBack(ARect.Top, ARect.Bottom, CopyToClipboardRowsCallback, AClipboardObject);

    AClipboardObject.CopyToClipboard;
  finally
    AClipboardObject.Free;
  end;
end;

procedure TvgrWorksheet.CopyToClipboard(const ARects: TvgrRectArray);
var
  I: Integer;
  ABounds: TRect;
  AClipboardObject: TvgrClipboardObject;
begin
  if Length(ARects) > 0 then
  begin
    AClipboardObject := TvgrClipboardObject.Create(Workbook);
    try
      ABounds := Rect(MaxInt, MaxInt, 0, 0);
      for I := 0 to High(ARects) do
        with ARects[I] do
        begin
          if ABounds.Left > Left then
            ABounds.Left := Left;
          if ABounds.Top > Top then
            ABounds.Top := Top;
          if Right > ABounds.Right then
            ABounds.Right := Right;
          if Bottom > ABounds.Bottom then
            ABounds.Bottom := Bottom;
        end;
      AClipboardObject.CopyPoint := ABounds.TopLeft;
      AClipboardObject.CopySize := Size(ABounds.Right - ABounds.Left, ABounds.Bottom - ABounds.Top);

      for I := 0 to High(ARects) do
      begin
        RangesList.FindAndCallBack(ARects[I], CopyToClipboardRangesCallback, AClipboardObject);
        BordersList.FindAtRectAndCallBack(ARects[I], CopyToClipboardBordersCallback, AClipboardObject);
        with ARects[I] do
        begin
          ColsList.FindAndCallBack(Left, Right, CopyToClipboardColsCallback, AClipboardObject);
          RowsList.FindAndCallBack(Top, Bottom, CopyToClipboardRowsCallback, AClipboardObject);
        end;
      end;

      AClipboardObject.CopyToClipboard;
    finally
      AClipboardObject.Free;
    end;
  end;
end;

procedure TvgrWorksheet.PasteFromPoint(AClipboardObject: TObject; const APastePoint: TPoint; APasteType: TvgrPasteTypeSet);
var
  I: Integer;

  function PasteFormula(AFormulaIndex: Integer): Integer;
  begin
    if AFormulaIndex = -1 then
      Result := -1
    else
      with TvgrClipboardObject(AClipboardObject).Formulas do
        Result := Workbook.Formulas.FindOrAdd(Formulas[AFormulaIndex], FormulasSize[AFormulaIndex], -1, TvgrClipboardObject(AClipboardObject).WBStrings);
  end;

  procedure PasteRange(AData: pvgrRange; APasteValue, APasteStyle: Boolean);
  var
    ARange: IvgrRange;
    ARangePlace: TRect;
    ANewStyle: rvgrRangeStyle;
  begin
    ARangePlace := AData.Place;
    OffsetRect(ARangePlace, APastePoint.X - TvgrClipboardObject(AClipboardObject).CopyPoint.X, APastePoint.Y - TvgrClipboardObject(AClipboardObject).CopyPoint.Y);
    with ARangePlace do
      ARange := Ranges[Left, Top, Right, Bottom];

    with pvgrRange(ARange.ItemData)^ do
    begin
      if APasteStyle then
      begin
        ANewStyle := TvgrClipboardObject(AClipboardObject).RangeStyles[AData.Style]^;
        if ANewStyle.DisplayFormat <> -1 then
          ANewStyle.DisplayFormat := Workbook.WBStrings.FindOrAdd(TvgrClipboardObject(AClipboardObject).WBStrings[ANewStyle.DisplayFormat]);
        Style := Workbook.RangeStyles.FindOrAdd(ANewStyle, Style);
      end;
      
      if APasteValue then
      begin
        if AData.Value.ValueType = rvtString then
        begin
          Value.ValueType := rvtString;
          Value.vString := Workbook.WBStrings.FindOrAdd(TvgrClipboardObject(AClipboardObject).WBStrings[AData.Value.vString]);
        end
        else
          Value := AData.Value;
        Formula := PasteFormula(AData.Formula);
      end;
      Flags := AData.Flags;
    end;
  end;

  procedure PasteBorder(AData: pvgrBorder);
  var
    ABorder: IvgrBorder;
    ALeft, ATop: Integer;
  begin
    ALeft := AData.Left + APastePoint.X - TvgrClipboardObject(AClipboardObject).CopyPoint.X;
    ATop := AData.Top + APastePoint.Y - TvgrClipboardObject(AClipboardObject).CopyPoint.Y;
    ABorder := Borders[ALeft, ATop, AData.Orientation];
    with pvgrBorder(ABorder.ItemData)^ do
      Style := Workbook.BorderStyles.FindOrAdd(TvgrClipboardObject(AClipboardObject).BorderStyles[AData.Style]^, Style);
  end;

  procedure PasteCol(AData: pvgrCol);
  begin
    with pvgrCol(Cols[AData.Number + APastePoint.X - TvgrClipboardObject(AClipboardObject).CopyPoint.X].ItemData)^ do
    begin
      Width := AData.Width;
      Flags := AData.Flags;
    end;
  end;

  procedure PasteRow(AData: pvgrRow);
  begin
    with pvgrRow(Rows[AData.Number + APastePoint.Y - TvgrClipboardObject(AClipboardObject).CopyPoint.Y].ItemData)^ do
    begin
      Height := AData.Height;
      Flags := AData.Flags;
    end;
  end;

begin
  with TvgrClipboardObject(AClipboardObject) do
  begin
    if (vgrptRangeValue in APasteType) or (vgrptRangeStyle in APasteType) then
      for I := 0 to RangeCount - 1 do
        PasteRange(Ranges[I], vgrptRangeValue in APasteType, vgrptRangeStyle in APasteType);
    if vgrptBorders in APasteType then
      for I := 0 to BorderCount - 1 do
        PasteBorder(Borders[I]);
    if vgrptCols in APasteType then
      for I := 0 to ColCount - 1 do
        PasteCol(Cols[I]);
    if vgrptRows in APasteType then
      for I := 0 to RowCount - 1 do
        PasteRow(Rows[I]);
  end;
end;

function TvgrWorksheet.ExportDataIsActive: Boolean;
begin
  Result := (Length(FExportData) > 0);
end;

procedure TvgrWorksheet.PasteFromClipboard(const ARect: TRect; APasteType: TvgrPasteTypeSet);
var
  I, J: Integer;
  AClipboardObject: TvgrClipboardObject;

  function CheckBound(ASide, ASize: Integer): Boolean;
  begin
    Result := (ASize = 0) or (ASide mod ASize = 0);
  end;

begin
  if Clipboard.HasFormat(vgrCF_WORKSHEET_RECT) then
  begin
    BeginUpdate;
    AClipboardObject := TvgrClipboardObject.Create(Workbook);
    try
      AClipboardObject.PasteFromClipboard;
      if CheckBound(ARect.Right - ARect.Left, AClipboardObject.CopySize.cx) and
         CheckBound(ARect.Bottom - ARect.Top, AClipboardObject.CopySize.cy) then
      begin
        I := ARect.Left;
        while I <= ARect.Right do
        begin
          J := ARect.Top;
          while J <= ARect.Bottom do
          begin
            PasteFromPoint(AClipboardObject, Point(I, J), APasteType);
            J := J + AClipboardObject.CopySize.cy + 1;
          end;
          I := I + AClipboardObject.CopySize.cx + 1;
        end;
      end
      else
        PasteFromPoint(AClipboardObject, ARect.TopLeft, APasteType);
    finally
      AClipboardObject.Free;
      EndUpdate;
    end;
  end;
end;

procedure TvgrWorksheet.ClearBorders(const ARect: TRect);
begin
  BeginUpdate;
  try
    BordersList.DeleteBorders(ARect);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.ClearBorders(const ARect: TRect; ABorderTypes: TvgrBorderTypes);
begin
  BeginUpdate;
  try
    BordersList.DeleteBorders(ARect, ABorderTypes);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.ClearRanges(const ARect: TRect);
begin
  BeginUpdate;
  try
    RangesList.InternalFindItemsInRect(ARect, nil, True, nil);
  finally
    EndUpdate;
  end;
end;

type
  rvgrClearRangesContents = record
    ClearContentFlags: TvgrClearContentFlags;
  end;
  pvgrClearRangesContents = ^rvgrClearRangesContents;
  
procedure TvgrWorksheet.ClearContentsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrRange do
  begin
    // vgrccDisplayFormat, vgrccValue, vgrccFormat
    with pvgrClearRangesContents(AData)^ do
    begin
      if vgrccFormat in ClearContentFlags then
        SetDefaultStyle
      else
        if vgrccDisplayFormat in ClearContentFlags then
          DisplayFormat := '';
      if vgrccValue in ClearContentFlags then
        Value := Null;
    end;
  end;
end;

procedure TvgrWorksheet.ClearRangesContents(const ARect: TRect; AClearFlags: TvgrClearContentFlags);
var
  ACallBackData: rvgrClearRangesContents;
begin
  BeginUpdate;
  try
    ACallBackData.ClearContentFlags := AClearFlags;
    RangesList.FindAndCallBack(ARect, ClearContentsCallback, @ACallBackData);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.ClearContents(const ARect: TRect; AClearFlags: TvgrClearContentFlags);
begin
  BeginUpdate;
  try
    if vgrccBorders in AClearFlags then
      ClearBorders(ARect);
    if vgrccRanges in AClearFlags then
      ClearRanges(ARect)
    else
      if [vgrccDisplayFormat, vgrccValue, vgrccFormat] * AClearFlags <> [] then
        ClearRangesContents(ARect, AClearFlags);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.ClearContents(const ARects: TvgrRectArray; AClearFlags: TvgrClearContentFlags);
var
  I: Integer;
begin
  BeginUpdate;
  try
    for I := 0 to High(ARects) do
      ClearContents(ARects[I], AClearFlags);
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.PrepareExportData;
begin
  SetLength(FExportData, SizeOf(Pointer) * FRanges.Count);
end;

procedure TvgrWorksheet.ClearExportData;
begin
  SetLength(FExportData, 0);
end;

function TvgrWorksheet.Copy(AInsertIndex: Integer): TvgrWorksheet;
var
  I: Integer;

  procedure CopySections(ASourceSections, ADestSections: TvgrSections);
  var
    I: Integer;
  begin
    for I := 0 to ASourceSections.Count - 1 do
      with ASourceSections.ByIndex[I] do
        ADestSections[StartPos, EndPos].Assign(ASourceSections.ByIndex[I]);
  end;

begin
  Workbook.BeginUpdate;
  try
    Result := FWorksheets.Insert(AInsertIndex);
    
    for I := 0 to RangesList.Count - 1 do
      with RangesList.ByIndex[I].Place do
        Result.Ranges[Left, Top, Right, Bottom].Assign(RangesList.ByIndex[I]);
    for I := 0 to BordersList.Count - 1 do
      with BordersList.ByIndex[I] do
        Result.Borders[Left, Top, Orientation].Assign(BordersList.ByIndex[I]);
    for I := 0 to ColsList.Count - 1 do
      with ColsList.ByIndex[I] do
        Result.Cols[Number].Assign(ColsList.ByIndex[I]);
    for I := 0 to RowsList.Count - 1 do
      with RowsList.ByIndex[I] do
        Result.Rows[Number].Assign(RowsList.ByIndex[I]);

    CopySections(VertSectionsList, Result.VertSectionsList);
    CopySections(HorzSectionsList, Result.HorzSectionsList);

    Result.Title := Title;
    Result.PageProperties.Assign(PageProperties);
  finally
    Workbook.EndUpdate;
  end;
end;

procedure TvgrWorksheet.ChangeValueTypeOfRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  (AItem as IvgrRange).ChangeValueType(TvgrRangeValueType(AData));
end;

procedure TvgrWorksheet.ChangeValueTypeOfRanges(const ARect: TRect; ANewValueType: TvgrRangeValueType);
begin
  BeginUpdate;
  try
    RangesList.FindAndCallBack(ARect, ChangeValueTypeOfRangesCallback, Pointer(ANewValueType));
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorksheet.ChangeValueTypeOfRanges(const ACellsRects: TvgrRectArray; ANewValueType: TvgrRangeValueType);
var
  I: Integer;
begin
  BeginUpdate;
  try
    for I := 0 to High(ACellsRects) do
      ChangeValueTypeOfRanges(ACellsRects[I], ANewValueType);
  finally
    EndUpdate;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrWorksheets
//
/////////////////////////////////////////////////
constructor TvgrWorksheets.Create(AWorkbook : TvgrWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
end;

destructor TvgrWorksheets.Destroy;
begin
  inherited;
end;

function TvgrWorksheets.Add : TvgrWorksheet;
var
  ChangeInfo : TvgrWorkbookChangeInfo;
begin
  ChangeInfo.ChangesType := vgrwcNewWorksheet;
  ChangeInfo.ChangedObject := nil;
  ChangeInfo.ChangedInterface := nil;
  BeforeChange(ChangeInfo);

  Result := Workbook.GetWorksheetClass.Create(Workbook.Owner);
  Result.FWorksheets := Self;
  Result.FTitle := sSheetTitlePrefix + IntToStr(Count + 1);
  Result.Name := GetUniqueComponentName(Result);
  inherited Add(Result);

  ChangeInfo.ChangedObject := Result;
  AfterChange(ChangeInfo);
end;

function TvgrWorksheets.Insert(Index: Integer): TvgrWorksheet;
var
  ChangeInfo : TvgrWorkbookChangeInfo;
begin
  ChangeInfo.ChangesType := vgrwcNewWorksheet;
  ChangeInfo.ChangedObject := nil;
  ChangeInfo.ChangedInterface := nil;
  BeforeChange(ChangeInfo);

  Result := Workbook.GetWorksheetClass.Create(Workbook.Owner);
  Result.FWorksheets := Self;
  Result.FTitle := sSheetTitlePrefix + IntToStr(Count + 1);
  Result.Name := GetUniqueComponentName(Result);
  inherited Insert(Index, Result);

  ChangeInfo.ChangedObject := Result;
  AfterChange(ChangeInfo);

  ChangeInfo.ChangesType := vgrwcUpdateAll;
  ChangeInfo.ChangedObject := nil;
  ChangeInfo.ChangedInterface := nil;
  BeforeChange(ChangeInfo);
  AfterChange(ChangeInfo);
end;

procedure TvgrWorksheets.Clear;
begin
  while Count > 0 do
    Items[0].Free;
  inherited;
end;

procedure TvgrWorksheets.RemoveWorksheet(AWorksheet: TvgrWorksheet);
begin
  inherited Remove(AWorksheet);
end;

procedure TvgrWorksheets.Delete(Index : Integer);
begin
  Items[Index].Free;
end;

// private
function TvgrWorksheets.GetItm(Index : integer) : TvgrWorksheet;
begin
  Result := TvgrWorksheet(inherited Items[Index]);
end;

procedure TvgrWorksheets.BeforeChange(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  FWorkbook.BeforeChange(ChangeInfo);
end;

procedure TvgrWorksheets.AfterChange(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  FWorkbook.AfterChange(ChangeInfo);
end;

procedure TvgrWorksheets.AddInList(AWorksheet: TvgrWorksheet);
begin
  inherited Add(AWorksheet);
end;

/////////////////////////////////////////////////
//
// TvgrFormulasList
//
/////////////////////////////////////////////////
constructor TvgrFormulasList.Create(AWorkbook: TvgrWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  FItems := TList.Create;
  FCompiler := TvteExcelFormulaCompiler.Create(Self);
  FCompiler.ShowExceptions := False;
  FCalculator := TvgrFormulaCalculator.Create(Self);
  FCalculatorStack := TvgrCalculatorStack.Create(FCalculator);
  FCalculator.Stack := FCalculatorStack;
  FCalculator.ExternalWBStrings := Workbook.WBStrings;
end;

destructor TvgrFormulasList.Destroy;
begin
  Clear;
  FCalculator.Free;
  FCalculatorStack.Free;
  FCompiler.Free;
  FItems.Free;
  inherited;
end;

procedure TvgrFormulasList.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
  begin
    InternalDelete(I);
    FreeMem(Items[I]);
  end;
  FItems.Clear;
  FCompiler.Clear;
  FCalculatorStack.Clear;
  FCalculator.ClearStrings;
end;

function TvgrFormulasList.GetItem(Index: Integer): pvgrFormulaHeader;
begin
  Result := pvgrFormulaHeader(FItems[Index]);
end;

function TvgrFormulasList.GetCount: Integer;
begin
  Result := FItems.Count;
end;

function TvgrFormulasList.GetFormula(Index: Integer): pvteFormula;
begin
  Result := pvteFormula(PChar(FItems[Index]) + SizeOf(rvgrFormulaHeader));
end;

function TvgrFormulasList.GetFormulaSize(Index: Integer): Integer;
begin
  Result := Items[Index].ItemCount;
end;

function TvgrFormulasList.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if GetInterface(IID, Obj) then
    Result := S_OK
  else
    Result := E_NOINTERFACE
end;

function TvgrFormulasList._AddRef: Integer;
begin
  Result := S_OK
end;

function TvgrFormulasList._Release: Integer;
begin
  Result := S_OK
end;

function TvgrFormulasList.GetStringIndex(const S: string): Integer;
begin
  Result := WBStrings.FindOrAdd(S)
end;

function TvgrFormulasList.GetString(AStringIndex: Integer): string;
begin
  Result := WBStrings[AStringIndex];
end;

function TvgrFormulasList.GetExternalSheetIndex(const AWorkbookName, ASheetName: string): Integer;
begin
  Result := Workbook.WorksheetIndexByTitle(ASheetName);
end;

function TvgrFormulasList.GetExternalSheetName(const AWorkbookName: string; AIndex: Integer): string;
begin
  if (AIndex >= 0) and (AIndex < Workbook.WorksheetsCount) then
    Result := Workbook.Worksheets[AIndex].Title
  else
    Result := '';
end;

function TvgrFormulasList.GetExternalWorkbookIndex(const AWorkbookName: string): Integer;
begin
  Result := -1;
end;

function TvgrFormulasList.GetExternalWorkbookName(AIndex: Integer): string;
begin
  Result := '';
end;

function TvgrFormulasList.GetFunctionName(AId: Integer): string;
begin
  Result := RegisteredFunctions.GetFunctionName(AId);
end;

function TvgrFormulasList.GetFunctionId(const AFunctionName: string): Integer;
begin
  Result := RegisteredFunctions.GetFunctionIndex(AFunctionName);
end;

function TvgrFormulasList.RangeValueToFormulaValue(ACalculator: TvgrFormulaCalculator; ARange: IvgrRange; ARangeSheet: Integer; var AValue: rvteFormulaValue): Boolean;
var
  ARangeData: pvgrRange;
  AVariantValue: Variant;
begin
  AValue.ValueType := vteNull;
  ARangeData := ARange.ItemData;
  if (ARangeData.Formula <> -1) and ((ARangeData.Flags and vgrmask_RangeFlagsNotCalced) <> 0) then
  begin
    if (ARangeData.Flags and vgrmask_RangeFlagsPassed) <> 0 then
    begin
      // circular cell reference
      Result := False;
      ARangeData.Flags := ARangeData.Flags and not vgrmask_RangeFlagsPassed;
      exit;
    end;
    ARangeData.Flags := ARangeData.Flags or vgrmask_RangeFlagsPassed;
    Result := FCalculator.Calculate(Formulas[ARangeData.Formula],
                                    FormulasSize[ARangeData.Formula],
                                    AVariantValue,
                                    ARangeSheet,
                                    ARangeData.Place.Left,
                                    ARangeData.Place.Top);
    ARangeData.Flags := ARangeData.Flags and not vgrmask_RangeFlagsPassed;
    if not Result then exit;
    VariantToRangeValue(ARangeData.Value, AVariantValue, WBStrings);
//    ARangeData.Flags := ARangeData.Flags and not vgrmask_RangeFlagsNotCalced;
  end
  else
    Result := True;

  AValue.ValueType := TvteFormulaValueType(ARangeData.Value.ValueType);
  case ARangeData.Value.ValueType of
    rvtInteger: AValue.vInteger := ARangeData.Value.vInteger;
    rvtExtended: AValue.vExtended := ARangeData.Value.vExtended;
    rvtString: AValue.vString := ACalculator.AddString(WBStrings[ARangeData.Value.vString]);
  end;
end;

function TvgrFormulasList.GetCellValue(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; ASheet, ACol, ARow: Integer): Boolean;
var
  ARange: IvgrRange;
begin
  if (ASheet >= 0) and (ASheet < Workbook.WorksheetsCount) then
  begin
    ARange := Workbook.Worksheets[ASheet].RangesList.FindAtCell(ACol, ARow);
    if ARange = nil then
    begin
      Result := True;
      AValue.ValueType := vteNull;
    end
    else
      Result := RangeValueToFormulaValue(ACalculator, ARange, ASheet, AValue);
  end
  else
  begin
    Result := True;
    AValue.ValueType := vteNull;
  end;
end;

procedure TvgrFormulasList.EnumRangeValuesCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  AValue: rvteFormulaValue;
  ARange: IvgrRange;
begin
  with pvgrCalculatorRec(AData)^ do
    if Result then
    begin
      ARange := AItem as IvgrRange;
      if PointInSelRect(ARange.Place.TopLeft, RangeRect) then
      begin
        Result := RangeValueToFormulaValue(FCalculator, ARange, RangeSheet, AValue);
        if Result then
          Callback(FCalculator, @AValue, EnumRec^);
      end;
    end;
end;

function TvgrFormulasList.EnumRangeValues(ACalculator: TvgrFormulaCalculator;
                                          var AEnumRec: rvgrEnumStackRec;
                                          ACallback: TvgrFormulaCalculatorEnumStackProc;
                                          ASheet: Integer;
                                          const ARangeRect: TRect): Boolean;
var
  ACalcRec: rvgrCalculatorRec;
begin
  ACalcRec.Calculator := ACalculator;
  ACalcRec.EnumRec := @AEnumRec;
  ACalcRec.Callback := ACallback;
  ACalcRec.RangeRect := ARangeRect;
  ACalcRec.RangeSheet := ASheet;
  ACalcRec.Result := True;
  Workbook.Worksheets[ASheet].RangesList.FindAndCallBack(ARangeRect, EnumRangeValuesCallback, @ACalcRec);
  Result := ACalcRec.Result;
end;

function TvgrFormulasList.GetFormulaText(Index: Integer; AFormulaCol, AFormulaRow: Integer): string;
begin
  Result := FCompiler.DecompileFormula(Formulas[Index], FormulasSize[Index], AFormulaCol, AFormulaRow);
end;

procedure TvgrFormulasList.SaveToStream(AStream: TStream);
var
  I, Buf: Integer;
begin
  // size of formula header
  Buf := SizeOf(rvgrFormulaHeader);
  AStream.Write(Buf, 4);
  // size of formula item
  Buf := SizeOf(rvteFormulaItem);
  AStream.Write(Buf, 4);
  // count of formulas
  Buf := Count;
  AStream.Write(Buf, 4);
  for I := 0 to Count - 1 do
  begin
    Buf := Items[I].ItemCount;
    AStream.Write(Buf, 4);
    AStream.Write(FItems[I]^, Buf * SizeOf(rvteFormulaItem) + SizeOf(rvgrFormulaHeader));
  end;
end;

procedure TvgrFormulasList.LoadFromStreamConvertRecord(AOldData, ANewData: Pointer;
                                                       AItemCount,
                                                       AOldHeaderSize, AOldItemSize,
                                                       ANewHeaderSize, ANewItemSize,
                                                       ADataStorageVersion: Integer);
begin
  if AOldHeaderSize + AOldItemSize * AItemCount > ANewHeaderSize + ANewItemSize * AItemCount then
    System.Move(AOldData^, ANewData^, ANewHeaderSize + ANewItemSize * AItemCount)
  else
    System.Move(AOldData^, ANewData^, AOldHeaderSize + AOldItemSize * AItemCount);
end;

procedure TvgrFormulasList.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
var
  I, Buf, ACount, AHeaderSize, AItemSize: Integer;
  AReadBuf, P: Pointer;
begin
  Clear;
  AStream.Read(AHeaderSize, 4);
  AStream.Read(AItemSize, 4);
  AStream.Read(ACount, 4);
  FItems.Count := ACount;
  if ADataStorageVersion = vgrDataStorageVersion then
    for I := 0 to ACount - 1 do
    begin
      AStream.Read(Buf, 4);
      GetMem(P, AHeaderSize + AItemSize * Buf);
      AStream.Read(P^, AHeaderSize + AItemSize * Buf);
      FItems[I] := P;
    end
  else
  begin
    AReadBuf := nil;
    try
      for I := 0 to ACount - 1 do
      begin
        AStream.Read(Buf, 4);
        ReallocMem(AReadBuf, AHeadersize + AItemSize * Buf);
        AStream.Read(AReadBuf^, AHeadersize + AItemSize * Buf);
        GetMem(P, SizeOf(rvgrFormulaHeader) + SizeOf(rvteFormulaItem) * Buf);
        LoadFromStreamConvertRecord(AReadBuf,
                                    P,
                                    Buf,
                                    AHeaderSize,
                                    AItemSize,
                                    SizeOf(rvgrFormulaHeader),
                                    SizeOf(rvteFormulaItem),
                                    ADataStorageVersion);
        FItems[I] := P;
      end;
    finally
      if AReadBuf <> nil then
        FreeMem(AReadBuf);
    end;
  end;
end;

procedure TvgrFormulasList.InternalDelete(AIndex: Integer);
var
  AFormula: pvteFormula;
  I, AFormulaSize: Integer;
begin
  AFormula := Formulas[AIndex];
  AFormulaSize := FormulasSize[AIndex];
  for I := 0 to AFormulaSize - 1 do
    if (AFormula[I].ItemType = vteitValue) and (AFormula[I].Value.ValueType = vteString) then
      WBStrings.Release(AFormula[I].Value.vString);
end;

function TvgrFormulasList.InternalAdd(var AHeader: pvgrFormulaHeader; AFormulaSize: Integer): Integer;
begin
  GetMem(AHeader, SizeOf(rvgrFormulaHeader) + AFormulaSize * SizeOf(rvteFormulaItem));
  Result := FItems.Add(AHeader);
end;

function TvgrFormulasList.InternalFindOrAdd(AFormula: pvteFormula; AFormulaSize: Integer): Integer;
var
  I: Integer;
  AHash: Integer;
  AHeader: pvgrFormulaHeader;
begin
  AHash := GetHashCode(AFormula^, AFormulaSize * SizeOf(rvteFormulaItem));
  AHeader := nil;
  Result := 0;
  for I := 0 to Count - 1 do
    with Items[I]^ do
    begin
      if (Hash = AHash) and
         (FormulasSize[I] = AFormulaSize) and
         CompareMem(Formulas[I], AFormula, AFormulaSize) then
      begin
        Result := I;
        Inc(RefCount);
        exit;
      end
      else
        if RefCount = 0 then
        begin
          if Items[I].ItemCount <> AFormulaSize then
          begin
            GetMem(AHeader, SizeOf(rvgrFormulaHeader) + AFormulaSize * SizeOf(rvteFormulaItem));
            FItems[I] := AHeader;
            Result := I;
          end
          else
          begin
            AHeader := Items[I];
            Result := I;
          end;
          break;
        end;
    end;
  if AHeader = nil then
    Result := InternalAdd(AHeader, AFormulaSize);

  AHeader.Hash := AHash;
  AHeader.RefCount := 1;
  AHeader.ItemCount := AFormulaSize;
  System.Move(AFormula^, Formulas[Result]^, AFormulaSize * sizeof(rvteFormulaItem));
end;

procedure TvgrFormulasList.Release(AIndex: Integer);
begin
  if AIndex >= 0 then
  begin
    Dec(Items[AIndex].RefCount);
    if Items[AIndex].RefCount = 0 then
      InternalDelete(AIndex);
  end;
end;

procedure TvgrFormulasList.AddRef(AIndex: Integer);
begin
  Inc(Items[AIndex].RefCount);
end;

function TvgrFormulasList.FindOrAdd(AFormula: pvteFormula; AFormulaSize, AOldIndex: Integer; ASourceWBStrings: TvgrWBStrings): Integer;
var
  ADestFormula: pvteFormula;
  AFormulaItem: pvteFormulaItem;
  I: Integer;
begin
  GetMem(ADestFormula,
         SizeOf(rvteFormulaItem) * AFormulaSize);
  MoveMemory(ADestFormula,
             AFormula,
             SizeOf(rvteFormulaItem) * AFormulaSize);
  try
    for I := 0 to AFormulaSize - 1 do
    begin
      AFormulaItem := @ADestFormula[I];
      if (AFormulaItem.ItemType = vteitValue) and (AFormulaItem.Value.ValueType = vteString) then
        AFormulaItem.Value.vString := WBStrings.FindOrAdd(ASourceWBStrings[AFormulaItem.Value.vString]);
    end;

    Result := FindOrAdd(ADestFormula, AFormulaSize, AOldIndex);
  finally
    FreeMem(ADestFormula);
  end;
end;

function TvgrFormulasList.FindOrAdd(AFormula: pvteFormula; AFormulaSize, AOldIndex: Integer): Integer;
begin
  Release(AOldIndex);
  Result := InternalFindOrAdd(AFormula, AFormulaSize);
end;

function TvgrFormulasList.FindOrAdd(const AFormulaText: string; AOldIndex: Integer; AFormulaCol, AFormulaRow: Integer): Integer;
var
  AFormula: pvteFormula;
  AFormulaSize: Integer;
begin
  Release(AOldIndex);
  if (Trim(AFormulaText) = '') or (not FCompiler.CompileFormula(AFormulaText,
                                                                AFormula,
                                                                AFormulaSize,
                                                                AFormulaCol,
                                                                AFormulaRow)) then
    Result := -1
  else
  begin
    Result := InternalFindOrAdd(AFormula, AFormulaSize);
    if AFormula <> nil then
      FreeMem(AFormula);
  end;
end;

{$IFDEF VGR_DEBUG}
function TvgrFormulasList.DebugInfo: TvgrDebugInfo;
var
  I, ARefCountZeroCount: Integer;
begin
  Result := TvgrDebugInfo.Create;
  Result.Add('SizeOf(rvgrFormulaHeader)', SizeOf(rvgrFormulaHeader), 'Размер заголовка для каждой записи о формуле');
  Result.Add('Count', Count, 'Количество формул');
  ARefCountZeroCount := 0;
  for I := 0 to Count - 1 do
    if Items[I].RefCount = 0 then
      Inc(ARefCountZeroCount);
  Result.Add('(RefCount = 0)', ARefCountZeroCount, 'Количество свободных элементов, у которых RefCount = 0');
end;
{$ENDIF}

procedure TvgrFormulasList.Calculate;
var
  I, J: Integer;
  ARange: pvgrRange;
  AValue: Variant;
begin
  FCalculator.ClearStrings;
  FCalculatorStack.Reset;
  for I := 0 to Workbook.WorksheetsCount - 1 do
    for J := 0 to Workbook.Worksheets[I].RangesCount - 1 do
    begin
      ARange := Workbook.Worksheets[I].RangesList.DataList[J];
      ARange.Flags := ARange.Flags or vgrmask_RangeFlagsNotCalced and not vgrmask_RangeFlagsPassed;
    end;
  for I := 0 to Workbook.WorksheetsCount - 1 do
    for J := 0 to Workbook.Worksheets[I].RangesCount - 1 do
    begin
      ARange := Workbook.Worksheets[I].RangesList.DataList[J];
      if (ARange.Formula <> -1) and
         ((ARange.Flags and vgrmask_RangeFlagsNotCalced) <> 0) then
      begin
        FCalculator.Calculate(Formulas[ARange.Formula], FormulasSize[ARange.Formula], AValue, I, ARange.Place.Left, ARange.Place.Top);
        VariantToRangeValue(ARange.Value, AValue, WBStrings);
        ARange.Flags := ARange.Flags and not vgrmask_RangeFlagsNotCalced;
      end;
    end;
end;

/////////////////////////////////////////////////
//
// TvgrWorkbook
//
/////////////////////////////////////////////////
constructor TvgrWorkbook.Create(AOwner : TComponent);
begin
  inherited;
  FDataStorageVersion := vgrDataStorageVersion;
  FWBStrings := TvgrWBStrings.Create;
  FFormulas := TvgrFormulasList.Create(Self);
  FFormulas.FWBStrings := FWBStrings;
  FRangeStyles := TvgrRangeStylesList.Create;
  FBorderStyles := TvgrBorderStylesList.Create;
  FWorksheets := TvgrWorksheets.Create(Self);
  FWorkbookHandlers := TInterfaceList.Create;
  {$IFDEF VTK_DSSAVE_DBG}
  with FWorksheets.Add do
  begin
    Title := 'Sheet1';
    Name := 'Sheet1';
  end;
  {$ENDIF}
end;

destructor TvgrWorkbook.Destroy;
begin
  FWorksheets.Clear;
  FreeAndNil(FWorksheets);
  FreeAndNil(FWorkbookHandlers);
  FreeAndNil(FRangeStyles);
  FreeAndNil(FBorderStyles);
  FreeAndNil(FFormulas);
  FreeAndNil(FWBStrings);
  inherited;
end;

procedure TvgrWorkbook.ReadFormulasData(Stream: TStream);
begin
  Formulas.LoadFromStream(Stream, DataStorageVersion);
end;

procedure TvgrWorkbook.WriteFormulasData(Stream: TStream);
begin
  Formulas.SaveToStream(Stream);
end;

procedure TvgrWorkbook.ReadRangeStylesData(Stream: TStream);
begin
  RangeStyles.LoadFromStream(Stream, DataStorageVersion);
end;

procedure TvgrWorkbook.WriteRangeStylesData(Stream: TStream);
begin
  RangeStyles.SaveToStream(Stream);
end;

procedure TvgrWorkbook.ReadBorderStylesData(Stream: TStream);
begin
  BorderStyles.LoadFromStream(Stream, DataStorageVersion);
end;

procedure TvgrWorkbook.WriteBorderStylesData(Stream: TStream);
begin
  BorderStyles.SaveToStream(Stream);
end;

procedure TvgrWorkbook.ReadStringsData(Stream: TStream);
begin
  WBStrings.LoadFromStream(Stream, DataStorageVersion);
end;

procedure TvgrWorkbook.WriteStringsData(Stream: TStream);
begin
  WBStrings.SaveToStream(Stream);
end;

procedure TvgrWorkbook.ReadDataStorageVersion(Reader: TReader);
begin
  FDataStorageVersion := Reader.ReadInteger;
end;

procedure TvgrWorkbook.WriteDataStorageVersion(Writer: TWriter);
begin
  Writer.WriteInteger(vgrDataStorageVersion);
end;

procedure TvgrWorkbook.ReadSystemInfo(Reader: TReader);
begin
  Reader.ReadListBegin;
  while not Reader.EndOfList do
    Reader.ReadString;
  Reader.ReadListEnd;
end;

procedure TvgrWorkbook.WriteSystemInfo(Writer: TWriter);
const
  PlatformIDs: Array [0..2] of string = ('WIN32s','WIN32_WINDOWS','WIN32_NT');
var
  S: string;
  si: SYSTEM_INFO;
  ovi: OSVERSIONINFO;
begin
  Writer.WriteListBegin;

  ZeroMemory(@ovi, sizeof(ovi));
  ovi.dwOSVersionInfoSize := sizeof(ovi);
  if GetVersionEx(ovi) then
  begin
    if ovi.dwPlatformId <= 2 then
      S := PlatformIDs[ovi.dwPlatformId]
    else
      S := 'Unknown';
    Writer.WriteString(Format('OS: %s %d.%d.%d %s',[S,
                                                    ovi.dwMajorVersion,
                                                    ovi.dwMinorVersion,
                                                    ovi.dwBuildNumber,
                                                    StrPas(ovi.szCSDVersion)]));
    Writer.WriteString('');
  end;

  GetSystemInfo(si);
  Writer.WriteString(Format('PageSize: %d', [si.dwPageSize]));
  Writer.WriteString(Format('ActiveProcessorMask: $%4.4x', [si.dwPageSize]));
  Writer.WriteString(Format('NumberOfProcessors: %d', [si.dwNumberOfProcessors]));
  Writer.WriteString(Format('ProcessorType: %d', [si.dwProcessorType]));
  Writer.WriteString('');

  {$IFDEF VTK_D4}
  S := 'Delphi4';
  {$ENDIF}

  {$IFDEF VTK_D5}
  {$IFDEF VTK_CB5}
  S := 'Builder5';
  {$ELSE}
  S := 'Delphi5';
  {$ENDIF}
  {$ENDIF}

  {$IFDEF VTK_D6}
  {$IFDEF VTK_CB6}
  S := 'Builder6';
  {$ELSE}
  S := 'Delphi6';
  {$ENDIF}
  {$ENDIF}

  {$IFDEF VTK_D7}
  S := 'Delphi7';
  {$ENDIF}

  Writer.WriteString('Compiler version: ' + S);
  Writer.WriteString('DataStorage version: ' + IntToStr(vgrDataStorageVersion));
  Writer.WriteString('GridReport version: ' + vgrGridReportVersion);

  Writer.WriteListEnd;
end;

procedure TvgrWorkbook.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('DataStorageVersion', ReadDataStorageVersion, WriteDataStorageVersion, True);
  Filer.DefineProperty('SystemInfo', ReadSystemInfo, WriteSystemInfo, True);
  Filer.DefineBinaryProperty('RangeStylesData', ReadRangeStylesData, WriteRangeStylesData,
    (RangeStyles.Count > 0));
  Filer.DefineBinaryProperty('BorderStylesData', ReadBorderStylesData, WriteBorderStylesData,
    (BorderStyles.Count > 0));
  Filer.DefineBinaryProperty('StringsData', ReadStringsData, WriteStringsData,
    (WBStrings.Count > 0));
  Filer.DefineBinaryProperty('FormulasData', ReadFormulasData, WriteFormulasData,
    (Formulas.Count > 0));
end;

procedure TvgrWorkbook.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;
//  if (Operation = opRemove) and (AComponent is TvgrWorksheet) and (FWorksheets <> nil) then
//    FWorksheets.RemoveWorksheetFromList(AComponent);
end;

function TvgrWorkbook.GetChildOwner: TComponent;
begin
  Result := Owner;//inherited GetChildOwner;
end;

procedure TvgrWorkbook.Clear;
begin
  BeginUpdate;
  try
    FWorksheets.Clear;
    FRangeStyles.Clear;
    FBorderStyles.Clear;
    FFormulas.Clear;
    FWBStrings.Clear;
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorkbook.DeleteWorksheet(Index : integer);
begin
  FWorksheets.Delete(Index);
end;

procedure TvgrWorkbook.RemoveWorksheet(Worksheet : TvgrWorksheet);
begin
  FWorksheets.Remove(Worksheet);
end;

function TvgrWorkbook.AddWorksheet : TvgrWorksheet;
begin
  Result := FWorksheets.Add;
end;

function TvgrWorkbook.InsertWorksheet(AInsertIndex: Integer): TvgrWorksheet;
begin
  Result := FWorksheets.Insert(AInsertIndex);
end;

function TvgrWorkbook.WorksheetIndexByTitle(const ATitle: string): Integer;
begin
  Result := 0;
  while (Result < WorksheetsCount) and (AnsiCompareText(Worksheets[Result].Title, ATitle) <> 0) do Inc(Result);
  if Result >= WorksheetsCount then
    Result := -1;
end;

procedure TvgrWorkbook.ConnectHandler(Value : IvgrWorkbookHandler);
begin
  if FWorkbookHandlers.IndexOf(Value) < 0 then
    FWorkbookHandlers.Add(Value);
end;

procedure TvgrWorkbook.DisconnectHandler(Value : IvgrWorkbookHandler);
begin
  if FWorkbookHandlers <> nil then
    FWorkbookHandlers.Remove(Value);
end;

procedure TvgrWorkbook.GetChildren(Proc: TGetChildProc; Root: TComponent);
var
  I: Integer;
begin
  for I :=0 to WorksheetsCount - 1 do
    Proc(Worksheets[I]);
end;

// protected
function TvgrWorkbook.GetWorksheetClass: TvgrWorksheetClass;
begin
  Result := TvgrWorksheet;
end;

function TvgrWorkbook.GetDispIdOfName(const AName: String) : Integer;
begin
  if AnsiCompareText(AName,'Worksheets') = 0 then
    Result := cs_TvgrWorkbook_Worksheets
  else
  if AnsiCompareText(AName,'WorksheetsCount') = 0 then
    Result := cs_TvgrWorkbook_WorksheetsCount
  else
  if AnsiCompareText(AName,'Clear') = 0 then
    Result := cs_TvgrWorkbook_Clear
  else
  if AnsiCompareText(AName,'AddWorksheet') = 0 then
    Result := cs_TvgrWorkbook_AddWorksheet
  else
  if AnsiCompareText(AName,'SaveToFile') = 0 then
    Result := cs_TvgrWorkbook_SaveToFile
  else
  if AnsiCompareText(AName,'LoadFromFile') = 0 then
    Result := cs_TvgrWorkbook_LoadFromFile
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrWorkbook.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrWorkbook, siTvgrWorkbookLength);
end;

function TvgrWorkbook.DoInvoke (DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrWorkbook_Worksheets:
      AResult := Worksheets[AParameters[0]] as IDispatch;
    cs_TvgrWorkbook_WorksheetsCount:
      AResult := WorksheetsCount;
    cs_TvgrWorkbook_Clear:
      Clear;
    cs_TvgrWorkbook_AddWorksheet:
      AResult := AddWorksheet as IDispatch;
    cs_TvgrWorkbook_SaveToFile:
      SaveToFile(AParameters[0]);
    cs_TvgrWorkbook_LoadFromFile:
      LoadFromFile(AParameters[0]);
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

procedure TvgrWorkbook.DoChanged;
begin
  if not (csDestroying in ComponentState) then  
  begin
    if Assigned(FOnChanged) then
      FOnChanged(Self);
  end;    
end;

// private
function TvgrWorkbook.GetWorksheetsCount : integer;
begin
  Result := FWorksheets.Count;
end;

function TvgrWorkbook.GetWorksheet(Index : integer) : TvgrWorksheet;
begin
  Result := FWorksheets[Index];
end;

procedure TvgrWorkbook.BeforeChange(ChangeInfo : TvgrWorkbookChangeInfo);
var
  i : Integer;
begin
  if FLockCount = 0 then
    for i := 0 to FWorkbookHandlers.Count - 1 do
      IvgrWorkbookHandler(FWorkbookHandlers.Items[i]).BeforeChangeWorkbook(ChangeInfo);
end;

procedure TvgrWorkbook.AfterChange(ChangeInfo : TvgrWorkbookChangeInfo);
var
  i : Integer;
begin
  if FLockCount = 0 then
  begin
    if ChangeInfo.ChangesType in [{vgrwcNewCol, vgrwcNewRow,
                                  vgrwcDeleteCol, vgrwcDeleteRow,}
                                  vgrwcChangeWorksheetContent, vgrwcNewRange, vgrwcDeleteRange, vgrwcChangeRange, vgrwcUpdateAll] then
      Formulas.Calculate;

    for i := 0 to FWorkbookHandlers.Count - 1 do
      IvgrWorkbookHandler(FWorkbookHandlers.Items[i]).AfterChangeWorkbook(ChangeInfo);

    DoChanged;
  end;
  FModified := True;
end;

procedure TvgrWorkbook.SetModified(Value: Boolean);
begin
  FModified := Value;
end;

procedure TvgrWorkbook.BeginUpdate;
begin
  Inc(FLockCount);
end;

procedure TvgrWorkbook.EndUpdate;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  Dec(FLockCount);
  if FLockCount < 0 then
    FLockCount := 0;
  if FLockCount = 0 then
  begin
    AChangeInfo.ChangesType := vgrwcUpdateAll;
    AChangeInfo.ChangedObject := nil;
    AChangeInfo.ChangedInterface := nil;
    AfterChange(AChangeInfo);
  end;
end;

procedure TvgrWorkbook.LoadFromStream(AStream: TStream);
var
  AName: string;
begin
  BeginUpdate;
  try
    FWorksheets.Clear;
    AName := Name;
    AStream.ReadComponent(Self);
    Name := AName;
  finally
    EndUpdate;
  end;
end;

procedure TvgrWorkbook.LoadFromFile(const AFileName: string);
var
  AFileStream: TFileStream;
begin
  AFileStream := TFileStream.Create(AFileName, fmOpenRead);
  try
    LoadFromStream(AFileStream);
  finally
    AFileStream.Free;
  end;
end;

procedure TvgrWorkbook.SaveToStream(AStream: TStream);
begin
  AStream.WriteComponent(Self);
end;

procedure TvgrWorkbook.SaveToFile(const AFileName: string);
var
  AFileStream: TFileStream;
begin
  AFileStream := TFileStream.Create(AFileName, fmCreate);
  try
    SaveToStream(AFileStream);
  finally
    AFileStream.Free;
  end;
end;

initialization

  Classes.RegisterClass(TvgrWorksheet);

end.

