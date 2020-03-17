{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains TvgrWorkbookPainter class, which provide functionality for drawing a page of the worksheet.
Instance of this class is created by the TvgrPrintEngine class for drawing the pages on the
printer canvas and by the TvgrWorkbookPreviewer class for drawing pages in the preview window.
To divide worksheet on the separate pages you can use TvgrPageMaker class.
See also:
  TvgrWorkbookPainter, TvgrPrintPage, TvgrPageMaker}
unit vgr_WorkbookPainter;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  SysUtils, Windows, {$IFDEF VTK_D6_OR_D7} Types, {$ENDIF} Classes, Messages, Graphics, math, typinfo,

  vgr_CommonClasses, vgr_PageMaker, vgr_DataStorage, vgr_DataStorageRecords, vgr_DataStorageTypes,
  vgr_GUIFunctions, vgr_ReportGUIFunctions, vgr_Functions, vgr_PageProperties;

//{$DEFINE VGR_PAINTERDEBUG}

type
  /////////////////////////////////////////////////
  //
  // IvgrWorkbookPainterOwner
  //
  /////////////////////////////////////////////////
{This interface should be implemented by the class, which use the TvgrWorkbookPainter.}
  IvgrWorkbookPainterOwner = interface
{Returns value of the DefaultColumnWidth property.
See also:
  DefaultColumnWidth}
    function GetDefaultColumnWidth: Integer;
{Returns value of the DefaultRowHeight property.
See also:
  DefaultRowHeight}
    function GetDefaultRowHeight: Integer;
{Returns value of the PixelsPerInchX property.
See also:
  PixelsPerInchX}
    function GetPixelsPerInchX: Integer;
{Returns value of the PixelsPerInchY property.
See also:
  PixelsPerInchY}
    function GetPixelsPerInchY: Integer;
{Returns value of the BackgroundColor property.
See also:
  BackgroundColor}
    function GetBackgroundColor: TColor;
{Returns value of the DefaultRangeBackgroundColor property.
See also:
  DefaultRangeBackgroundColor}
    function GetDefaultRangeBackgroundColor: TColor;

{Specifies in twips the default column width, this value is used if width of column on page was not hardcoded.}
    property DefaultColumnWidth: Integer read GetDefaultColumnWidth;
{Specifies in twips the default row height, this value is used if height of row on page was not hardcoded.}
    property DefaultRowHeight: Integer read GetDefaultRowHeight;
{Specifies the number of pixels per logical inch along the width.}
    property PixelsPerInchX: Integer read GetPixelsPerInchX;
{Specifies the number of pixels per logical inch along the height.}
    property PixelsPerInchY: Integer read GetPixelsPerInchY;
{Specifies color, which is used to fill the area that is not occupied by the ranges.}
    property BackgroundColor: TColor read GetBackgroundColor;
{Specifies the default background color of the ranges. If the background color of
the range equals to this color the range is not fills.}
    property DefaultRangeBackgroundColor: TColor read GetDefaultRangeBackgroundColor;
  end;

  /////////////////////////////////////////////////
  //
  // rvgrPainterSize
  //
  /////////////////////////////////////////////////
{Internal structure, used for caching the sizes of the columns and rows.
Syntax:
  rvgrPainterSize = packed record
    Number: Integer;
    Start: Integer;
    Size: Integer;
  end;}
  rvgrPainterSize = packed record
    Number: Integer;
    Start: Integer;
    Size: Integer;
  end;
{Pointer to the rvgrPainterSize structure.
Syntax:
  pvgrPainterSize = ^rvgrPainterSize;}
  pvgrPainterSize = ^rvgrPainterSize;

(*$NODEFINE TvgrBorderCacheInfo*)
(*$HPPEMIT 'class PASCALIMPLEMENTATION TvgrBorderCacheInfo : public System::TObject'*)
(*$HPPEMIT '{'*)
(*$HPPEMIT '	typedef System::TObject inherited;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT 'protected:'*)
(*$HPPEMIT '	int Coord;'*)
(*$HPPEMIT '	#pragma pack(push, 1)'*)
(*$HPPEMIT '	TRect Bounds;'*)
(*$HPPEMIT '	#pragma pack(pop)'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '	Vgr_datastoragerecords::rvgrBorderStyle *Style;'*)
(*$HPPEMIT '	TvgrFastRegion* Region;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT 'public:'*)
(*$HPPEMIT '	__fastcall TvgrBorderCacheInfo(void);'*)
(*$HPPEMIT '	__fastcall virtual ~TvgrBorderCacheInfo(void);'*)
(*$HPPEMIT '};'*)

  /////////////////////////////////////////////////
  //
  // TvgrBorderCacheInfo
  //
  /////////////////////////////////////////////////
{Internal class, used for caching information about borders.}
  TvgrBorderCacheInfo = class(TObject)
  protected
    Coord: Integer;
    Bounds: TRect;
    Style: pvgrBorderStyle;
    Region: TvgrFastRegion;
  public
{Creates a instance of the TvgrBorderCacheInfo class.}
    constructor Create;
{Frees a instance of the TvgrBorderCacheInfo class.}
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBorderCacheList
  //
  /////////////////////////////////////////////////
{Internal class, contains list of the TvgrBorderCacheInfo objects.}
  TvgrBorderCacheList = class(TvgrObjectList)
  private
    function GetItm(Index: Integer): TvgrBorderCacheInfo;
  public
{Finds the TvgrBorderCacheInfo object by its border color.
Parameters:
  AColor - the color of the border.
Return value:
  Returns index of the found object or -1 if object are not found.}
    function IndexBySolidColor(AColor: TColor): Integer;
{Finds the TvgrBorderCacheInfo object by the border style and border position.
Parameters:
  ABorderStyle - pointer to the rvgrBorderStyle structure, that describes the border style.
  ACoord - position of the border.
Return value:
  Returns index of the found object or -1 if object are not found.}
    function IndexByStyleAndCoord(ABorderStyle: pvgrBorderStyle; ACoord: Integer): Integer;
{Creates and adds the new empty TvgrBorderCacheInfo object.
Return value:
  Returns the created TvgrBorderCacheInfo object.}
    function Add: TvgrBorderCacheInfo;

{Enumerates TvgrBorderCacheInfo objects, first object has index 0.
Parameters:
  Index - index of the object.}
    property Items[Index: Integer]: TvgrBorderCacheInfo read GetItm; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWorkbookPainter
  //
  /////////////////////////////////////////////////
{TvgrWorkbookPainter class provide functionality for drawing a page of the worksheet.
Instance of this class is created by the TvgrPrintEngine class for drawing the pages on the
printer canvas and by the TvgrWorkbookPreviewer class for drawing pages in the preview window.
To divide worksheet on the separate pages you can use TvgrPageMaker class.
See also:
  TvgrPrintPage, TvgrPrintEngine, TvgrWorkbookPreviewer}
  TvgrWorkbookPainter = class(TObject)
  private
    FOwner: IvgrWorkbookPainterOwner;
    FPage: TvgrPrintPage;
    FWorksheet: TvgrWorksheet;
    FCanvas: TCanvas;
    FRowSizes: Array of rvgrPainterSize;
    FColSizes: Array of rvgrPainterSize;
    FScale: Double;
//    FRangesRegion: HRGN;
    FBrushOrg: TPoint;
    FCurRow: Integer;
    FCurCol: Integer;
    FLeftOffs: Integer;
    FTopOffs: Integer;
    FWidth: Integer;
    FHeight: Integer;
    FSolidBorders: TvgrBorderCacheList;
    FHorzBorders: TvgrBorderCacheList;
    FVertBorders: TvgrBorderCacheList;
    FHiddenBorders: TList;

    function GetDefaultColWidth: Integer;
    function GetDefaultRowHeight: Integer;
    procedure SetWorksheet(Value: TvgrWorksheet);

    function GetNumberIndex(ANumber: Integer): Integer;
    procedure HideBorder(ABorder: IvgrWBListItem); overload;
    procedure HideBorder(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation); overload;
    function IsBorderHidden(ABorder: IvgrWBListItem): Boolean;
    function FindBorder(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation): IvgrBorder;
    function GetVerticalBorderPartSize(ALeft, ATop: Integer; ASide: TvgrBorderSide): Integer;
    function IsCellBusy(ALeft, ATop: Integer): Boolean;
    function IsRangeMayBeLong(ARange: IvgrRange): Boolean;
    procedure CheckSideCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure CalcRange(const ARangeRect, ADrawingRect: TRect;
                        var ADrawRect: TRect;
                        var ADrawPixelRect: TRect;
                        var AFullPixelRect: TRect;
                        var ADrawBorderPixelRect: TRect;
                        var AFullPixelBorderRect: TRect;
                        var ABorderExcludeRegion: HRGN);
    procedure FindBordersWithinLongRanges(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure PaintRange(ACanvas: TCanvas; ARange: IvgrRange; const ADrawingRect: TRect);
    procedure PaintNormalRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure PaintLongRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    procedure CalcBorder(ABorder: IvgrBorder; var ABorderPixelSize: Integer; var ABorderRect: TRect);
    procedure PrecisionCalcBorder(ABorder: IvgrBorder; var ABorderPixelSize: Integer; var ABorderRect: TRect);
    procedure PaintBorderLine(const ARect: TRect;
                              AIsHorizontal: Boolean;
                              AForeColor: TColor;
                              ABorderStyle: TvgrBorderStyle;
                              const AOrigin: TPoint;
                              AClipRegion: HRGN);
    procedure PaintBorders(ABordersList: TvgrBorderCacheList; AIsHorizontal: Boolean);
    procedure PaintSolidBorders(ABordersList: TvgrBorderCacheList);
    procedure PaintBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    procedure FindAndCallbackRanges(ACallbackProc: TvgrFindListItemCallbackProc);
    procedure FindAndCallbackBorders(ACallbackProc: TvgrFindListItemCallbackProc);

    function FindCol(AColNumber: Integer): pvgrPainterSize;
    function FindRow(ARowNumber: Integer): pvgrPainterSize;

    function GetColLeft(AColNumber: Integer): Integer;
    function GetColRight(AColNumber: Integer): Integer;
    function GetColWidth(AColNumber: Integer): Integer;
    function GetRowTop(ARowNumber: Integer): Integer;
    function GetRowBottom(ARowNumber: Integer): Integer;
    function GetRowHeight(ARowNumber: Integer): Integer;

    function CalculateRowHeight(ANumber: Integer): Integer;
    function CalculateColWidth(ANumber: Integer): Integer;
    procedure Cache(AScale: Double);
  protected
    function TwipsToPixelsX(ATwips: Integer): Integer;
    function TwipsToPixelsY(ATwips: Integer): Integer;
    property Canvas: TCanvas read FCanvas;

{Returns TvgrPrintPage object that are drawn.}
    property Page: TvgrPrintPage read FPage;
//    property Scale: Double read FScale;
    property BrushOrg: TPoint read FBrushOrg;
    property LeftOffs: Integer read FLeftOffs;
    property TopOffs: Integer read FTopOffs;

    procedure DrawHeaderFooter(ACanvas: TCanvas;
                                               AHeaderFooter: TvgrPageHeaderFooter;
                                               const ARect: TRect);
  public
{Creates a instance of the TvgrWorkbookPainter class.
Parameters:
  AOwner - IvgrWorkbookPainterOwner interace, which should be implemented by the parent object.
See also:
  IvgrWorkbookPainterOwner}
    constructor Create(AOwner: IvgrWorkbookPainterOwner);
{Frees a instance of the TvgrWorkbookPainter class.}
    destructor Destroy; override;

{Draws the page on the specified canvas.
Parameters:
  X - x coordinate of the upper-left corner of the drawn page.
  Y - y coordinate of the upper-left corner of the drawn page.
  ACanvas - TCanvas object, on which page must be drawn.
  APage - TvgrPrintPage object for drawn.
  AScale - the scale factor.
  ABrushOrg - not used, can contains any value}
    procedure DrawPage(X, Y: Integer; ACanvas: TCanvas; APage: TvgrPrintPage; AScale: Double; const ABrushOrg: TPoint); virtual;
{Returns width of the page in pixels.
Parameters:
  APage - TvgrPrintPage object.
  AScale - the scale factor.
Return value:
  The width of the page in pixels.}
    function GetPagePixelWidth(APage: TvgrPrintPage; AScale: Double): Integer;
{Returns height of the page in pixels.
Parameters:
  APage - TvgrPrintPage object.
  AScale - the scale factor.
Return value:
  The height of the page in pixels.}
    function GetPagePixelHeight(APage: TvgrPrintPage; AScale: Double): Integer;

{Returns the IvgrWorkbookPainterOwner passed to the constructor.}
    property Owner: IvgrWorkbookPainterOwner read FOwner;
{Sets or gets TvgrWorksheet object, that currently drawn.}
    property Worksheet: TvgrWorksheet read FWorksheet write SetWorksheet;
{Specifies in twips the default column width, this value is used if width of column on page was not hardcoded.}
    property DefaultColWidth: Integer read GetDefaultColWidth;
{Specifies in twips the default row height, this value is used if height of row on page was not hardcoded.}
    property DefaultRowHeight: Integer read GetDefaultRowHeight;
  end;

implementation

uses
  vgr_Consts, vgr_ReportFunctions;

/////////////////////////////////////////////////
//
// TvgrBorderCacheInfo
//
/////////////////////////////////////////////////
constructor TvgrBorderCacheInfo.Create;
begin
  inherited Create;
  Region := TvgrFastRegion.Create;
end;

destructor TvgrBorderCacheInfo.Destroy;
begin
  Region.Free;
  inherited;
end;

/////////////////////////////////////////////////
//
// TvgrBorderCacheList
//
/////////////////////////////////////////////////
function TvgrBorderCacheList.GetItm(Index: Integer): TvgrBorderCacheInfo;
begin
  Result := TvgrBorderCacheInfo(inherited Items[Index]);
end;

function TvgrBorderCacheList.IndexBySolidColor(AColor: TColor): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].Style.Color <> AColor) do
    Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrBorderCacheList.IndexByStyleAndCoord(ABorderStyle: pvgrBorderStyle; ACoord: Integer): Integer;
begin
  Result := 0;
  while (Result < Count) and
        ((Items[Result].Style.Color <> ABorderStyle.Color) or
         (Items[Result].Style.Pattern <> ABorderStyle.Pattern) or
         (Items[Result].Coord <> ACoord)) do
    Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrBorderCacheList.Add: TvgrBorderCacheInfo;
begin
  Result := TvgrBorderCacheInfo.Create;
  inherited Add(Result);
end;

/////////////////////////////////////////////////
//
// TvgrWorkbookPainter
//
/////////////////////////////////////////////////
constructor TvgrWorkbookPainter.Create(AOwner: IvgrWorkbookPainterOwner);
begin
  inherited Create;
  FOwner := AOwner;
  FSolidBorders := TvgrBorderCacheList.Create;
  FHorzBorders := TvgrBorderCacheList.Create;
  FVertBorders := TvgrBorderCacheList.Create;
  FHiddenBorders := TList.Create;
end;

destructor TvgrWorkbookPainter.Destroy;
begin
  FSolidBorders.Free;
  FHorzBorders.Free;
  FVertBorders.Free;
  FHiddenBorders.Free;
  inherited;
end;

function TvgrWorkbookPainter.GetDefaultColWidth: Integer;
begin
  Result := FOwner.DefaultColumnWidth;
end;

function TvgrWorkbookPainter.GetDefaultRowHeight: Integer;
begin
  Result := FOwner.DefaultRowHeight;
end;

procedure TvgrWorkbookPainter.SetWorksheet(Value: TvgrWorksheet);
begin
  if FWorksheet <> Value then
  begin
    FWorksheet := Value;
  end;
end;

function TvgrWorkbookPainter.FindCol(AColNumber: Integer): pvgrPainterSize;
var
  I: Integer;
begin
  for I := FCurCol to High(FColSizes) do
    if FColSizes[I].Number = AColNumber then
    begin
      Result := @FColSizes[I];
      exit;
    end;
  Result := nil;
end;

function TvgrWorkbookPainter.FindRow(ARowNumber: Integer): pvgrPainterSize;
var
  I: Integer;
begin
  for I := FCurRow to High(FRowSizes) do
    if FRowSizes[I].Number = ARowNumber then
    begin
      Result := @FRowSizes[I];
      exit;
    end;
  Result := nil;
end;

function TvgrWorkbookPainter.TwipsToPixelsX(ATwips: Integer): Integer;
begin
  Result := Round(Owner.PixelsPerInchX / TwipsPerInch * ATwips * FScale);
end;

function TvgrWorkbookPainter.TwipsToPixelsY(ATwips: Integer): Integer;
begin
  Result := Round(Owner.PixelsPerInchY / TwipsPerInch * ATwips * FScale);
end;

procedure TvgrWorkbookPainter.HideBorder(ABorder: IvgrWBListItem);
begin
  FHiddenBorders.Add(Pointer(ABorder.ItemIndex));
end;

procedure TvgrWorkbookPainter.HideBorder(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation);
var
  ABorder: IvgrBorder;
begin
  ABorder := Worksheet.BordersList.Find(ALeft, ATop, AOrientation);
  if ABorder <> nil then
    HideBorder(ABorder);
end;

function TvgrWorkbookPainter.IsBorderHidden(ABorder: IvgrWBListItem): Boolean;
begin
  Result := FHiddenBorders.IndexOf(Pointer(ABorder.ItemIndex)) <> -1;
end;

function TvgrWorkbookPainter.FindBorder(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation): IvgrBorder;
begin
  Result := Worksheet.BordersList.Find(ALeft, ATop, AOrientation);
  if (Result <> nil) and IsBorderHidden(Result) then
    Result := nil;
end;

function TvgrWorkbookPainter.GetVerticalBorderPartSize(ALeft, ATop: Integer; ASide: TvgrBorderSide): Integer;
var
  ABorder: IvgrBorder;
begin
  ABorder := FindBorder(ALeft, ATop, vgrboLeft);
  if ABorder <> nil then
    Result := vgr_Functions.GetBorderPartSize(ConvertTwipsToPixelsX(ABorder.Width), ASide)
  else
    Result := 0;
end;

function TvgrWorkbookPainter.GetNumberIndex(ANumber: Integer): Integer;
begin
  Result := 0;
  while (Result <= High(FColSizes)) and (FColSizes[Result].Number <> ANumber) do
    Inc(Result);
  if Result > High(FColSizes) then
    Result := -1;
end;

function TvgrWorkbookPainter.IsCellBusy(ALeft, ATop: Integer): Boolean;
var
  ARange: IvgrRange;
begin
  ARange := Worksheet.RangesList.FindAtCell(ALeft, ATop);
  Result := (ARange <> nil) and (ARange.DisplayText <> '');
end;

function TvgrWorkbookPainter.IsRangeMayBeLong(ARange: IvgrRange): Boolean;
begin
  with ARange do
    Result := (Place.Left = Place.Right) and
              (Place.Top = Place.Bottom) and
              ((Flags and vgrmask_RangeFlagsWordWrap) = 0) and
              (Angle = 0) and
              (DisplayText <> '');
end;

procedure TvgrWorkbookPainter.PaintNormalRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  if not IsRangeMayBeLong(AItem as IvgrRange) then
    PaintRange(Canvas, AItem as IvgrRange, PRect(AData)^);
end;

procedure TvgrWorkbookPainter.PaintLongRangesCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  if IsRangeMayBeLong(AItem as IvgrRange) then
    PaintRange(Canvas, AItem as IvgrRange, PRect(AData)^);
end;

procedure TvgrWorkbookPainter.FindAndCallbackRanges{(ACallbackProc: TvgrFindListItemCallbackProc)};
var
  I, J, N: Integer;
  ARect: TRect;
begin
  I := 0;
  FCurRow := I;
  while I < Page.RowsCount do
  begin
    N := Page.Row[I];
    ARect.Top := N;
    repeat
      Inc(N);
      Inc(I);
    until (I >= Page.RowsCount) or (N <> Page.Row[I]);
    ARect.Bottom := N - 1;

    J := 0;
    FCurCol := J;
    while J < Page.ColsCount do
    begin
      N := Page.Col[J];
      ARect.Left := N;
      repeat
        Inc(N);
        Inc(J);
      until (J >= Page.ColsCount) or (N <> Page.Col[J]);
      ARect.Right := N - 1;

      Worksheet.RangesList.FindAndCallback(ARect, ACallbackProc, @ARect);
      FCurCol := J;
    end;
    FCurRow := I;
  end;
end;

procedure TvgrWorkbookPainter.CalcBorder(ABorder: IvgrBorder; var ABorderPixelSize: Integer; var ABorderRect: TRect);
begin
  if ABorder.Orientation = vgrboLeft then
  begin
    ABorderPixelSize := TwipsToPixelsX(ABorder.Width);

    ABorderRect.Left := GetColLeft(ABorder.Left) - 1 - ABorderPixelSize div 2;
    ABorderRect.Right := ABorderRect.Left + ABorderPixelSize;
    ABorderRect.Top := GetRowTop(ABorder.Top);
    ABorderRect.Bottom := ABorderRect.Top + GetRowHeight(ABorder.Top);
  end
  else
  begin
    ABorderPixelSize := TwipsToPixelsY(ABorder.Width);

    ABorderRect.Left := GetColLeft(ABorder.Left);
    ABorderRect.Right := ABorderRect.Left + GetColWidth(ABorder.Left);
    ABorderRect.Top := GetRowTop(ABorder.Top) - 1 - ABorderPixelSize div 2;
    ABorderRect.Bottom := ABorderRect.Top + ABorderPixelSize;
  end;
end;

procedure TvgrWorkbookPainter.PrecisionCalcBorder(ABorder: IvgrBorder; var ABorderPixelSize: Integer; var ABorderRect: TRect);
var
  ABorder1, ABorder2: IvgrBorder;
  ABorder1Size: Integer;
  ABorder1Rect: TRect;
  ABorder2Size: Integer;
  ABorder2Rect: TRect;
begin
  CalcBorder(ABorder, ABorderPixelSize, ABorderRect);
  if ABorder.Orientation = vgrboTop then
  begin
    ABorder1 := FindBorder(ABorder.Left, ABorder.Top, vgrboLeft);
    if ABorder1 <> nil then
      CalcBorder(ABorder1, ABorder1Size, ABorder1Rect)
    else
    begin
      ABorder1Size := 0;
      ABorder1Rect.Left := MaxInt;
    end;
    ABorder2 := FindBorder(ABorder.Left, ABorder.Top - 1, vgrboLeft);
    if ABorder2 <> nil then
      CalcBorder(ABorder2, ABorder2Size, ABorder2Rect)
    else
    begin
      ABorder2Size := 0;
      ABorder2Rect.Left := MaxInt;
    end;
    ABorderRect.Left := Min(ABorderRect.Left, Min(ABorder1Rect.Left, ABorder2Rect.Left)); 

    ABorder1 := FindBorder(ABorder.Left + 1, ABorder.Top, vgrboLeft);
    if ABorder1 <> nil then
      CalcBorder(ABorder1, ABorder1Size, ABorder1Rect)
    else
    begin
      ABorder1Size := 0;
      ABorder1Rect.Right := -MaxInt;
    end;
    ABorder2 := FindBorder(ABorder.Left + 1, ABorder.Top - 1, vgrboLeft);
    if ABorder2 <> nil then
      CalcBorder(ABorder2, ABorder2Size, ABorder2Rect)
    else
    begin
      ABorder2Size := 0;
      ABorder2Rect.Right := -MaxInt;
    end;
    ABorderRect.Right := Max(ABorderRect.Right, Max(ABorder1Rect.Right, ABorder2Rect.Right));
  end;
end;

var
  APens: array [TvgrBorderStyle] of TDynIntegerArray;

procedure TvgrWorkbookPainter.PaintBorderLine(const ARect: TRect;
                                              AIsHorizontal: Boolean;
                                              AForeColor: TColor;
                                              ABorderStyle: TvgrBorderStyle;
                                              const AOrigin: TPoint;
                                              AClipRegion: HRGN);
var
  I, ALen, P: Integer;
  AForeRegion: TvgrFastRegion;
  AParts: TDynIntegerArray;
  AFore: Boolean;
  APartRect: TRect;

  procedure FillRegion(AColor: TColor; ARegion: TvgrFastRegion);
  var
    ABrush: HBRUSH;
    AHRGN: HRGN;
  begin
    ABrush := CreateSolidBrush(GetRGBColor(AColor));
    AHRGN := ARegion.BuildRegion;
    if AClipRegion <> 0 then
      CombineRegionAndRegionNoDelete(AHRGN, AClipRegion, RGN_AND);
    FillRgn(FCanvas.Handle, AHRGN, ABrush);
    DeleteObject(ABrush);
    DeleteObject(AHRGN);
  end;

begin
  AForeRegion := TvgrFastRegion.Create;
  try
    AParts := APens[ABorderStyle];
    ALen := Length(AParts);
    AFore := True;
    if AIsHorizontal then
    begin
      P := ARect.Left;
      APartRect.Top := ARect.Top;
      APartRect.Bottom := ARect.Bottom;
      while P < ARect.Right do
      begin
        I := 0;
        while (P < ARect.Right) and (I < ALen) do
        begin
          if AFore then
          begin
            APartRect.Left := P;
            APartRect.Right := Min(APartRect.Left + AParts[I], ARect.Right);
            AForeRegion.AddRect(APartRect);
          end;
          P := P + AParts[I];
          AFore := not AFore;
          Inc(I);
        end;
      end;
    end
    else
    begin
      P := ARect.Top;
      APartRect.Left := ARect.Left;
      APartRect.Right := ARect.Right;
      while P < ARect.Bottom do
      begin
        I := 0;
        while (P < ARect.Bottom) and (I < ALen) do
        begin
          if AFore then
          begin
            APartRect.Top := P;
            APartRect.Bottom := Min(APartRect.Top + AParts[I], ARect.Bottom);
            AForeRegion.AddRect(APartRect);
          end;
          P := P + AParts[I];
          AFore := not AFore;
          Inc(I);
        end;
      end;
    end;

    FillRegion(AForeColor, AForeRegion);
  finally
    AForeRegion.Free;
  end;
end;

procedure TvgrWorkbookPainter.PaintBorders(ABordersList: TvgrBorderCacheList; AIsHorizontal: Boolean);
var
  I: Integer;
  ARegion: HRGN;
{
  ARegion, AOldRegion: HRGN;
  APen, AOldPen: HPEN;
  ALogBrush: tagLOGBRUSH;
}
begin
  for I := 0 to ABordersList.Count - 1 do
    with ABordersList[I] do
    begin
      ARegion := Region.BuildRegion;
      PaintBorderLine(Bounds,
                      AIsHorizontal,
                      Style.Color,
                      Style.Pattern,
                      Point(0, 0),
                      ARegion);
      DeleteObject(ARegion);
(*      
{
      if Owner.BackgroundColor <> clNone then
      begin
        ABrush := CreateSolidBrush(GetRGBColor(Owner.BackgroundColor));
        FillRect(ACanvas.Handle, ABorderRect, ABrush);
        DeleteObject(ABrush);
      end;
}
      AOldRegion := GetClipRegion(FCanvas);
      ARegion := Region.BuildRegion;
      SetClipRegion(FCanvas, ARegion);

      // create pen
      ALogBrush.lbStyle := BS_SOLID;
      ALogBrush.lbColor := GetRGBColor(Style.Color);
      APen := ExtCreatePen(PS_GEOMETRIC or PS_ENDCAP_SQUARE or APens[Style.Pattern], Size, ALogBrush, 0, nil);
      AOldPen := SelectObject(FCanvas.Handle, APen);

      // draw border
      if AIsHorizontal then
      begin
        MoveToEx(FCanvas.Handle, FLeftOffs, Pos, nil);
        LineTo(FCanvas.Handle, FLeftOffs + FWidth, Pos);
      end
      else
      begin
        MoveToEx(FCanvas.Handle, Pos, FTopOffs, nil);
        LineTo(FCanvas.Handle, Pos, FTopOffs + FHeight);
      end;

      // delete pen
      SelectObject(FCanvas.Handle, AOldPen);
      DeleteObject(APen);

      // restore clip region
      SetClipRegion(FCanvas, AOldRegion);
      ExcludeClipRegion(FCanvas, ARegion);

      DeleteObject(ARegion);
      DeleteObject(AOldRegion);
*)      
    end;
end;

procedure TvgrWorkbookPainter.PaintSolidBorders(ABordersList: TvgrBorderCacheList);
var
  I: Integer;
  ABrush: HBRUSH;
  ARegion: HRGN;
begin
  for I := 0 to ABordersList.Count - 1 do
    with ABordersList[I] do
    begin
      ABrush := CreateSolidBrush(GetRGBColor(Style.Color));
      ARegion := Region.BuildRegion;
      FillRgn(FCanvas.Handle, ARegion, ABrush);
      DeleteObject(ABrush);
      DeleteObject(ARegion);
      ExcludeClipRegion(FCanvas, ARegion);
    end;
end;

procedure TvgrWorkbookPainter.PaintBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ABorder: pvgrBorder;
  ABorderStyle: pvgrBorderStyle;
  I, ACoord, ABorderPixelSize: Integer;
  ABorderRect: TRect;
  ABorderInfo: TvgrBorderCacheInfo;
  AList: TvgrBorderCacheList;
begin
  if IsBorderHidden(AItem) then
    exit;
    
  PrecisionCalcBorder(AItem as IvgrBorder, ABorderPixelSize, ABorderRect);
//  CalcBorder(AItem as IvgrBorder, ABorderPixelSize, ABorderRect);
  with AItem as IvgrBorder do
  begin
    ABorder := pvgrBorder(ItemData);
    ABorderStyle := pvgrBorderStyle(StyleData);
  end;

  if ABorderStyle.Pattern = vgrbsSolid then
  begin
    I := FSolidBorders.IndexBySolidColor(ABorderStyle.Color);
    if I = -1 then
    begin
      ABorderInfo := FSolidBorders.Add;
      with ABorderInfo do
      begin
        Coord := -1;
        Style := ABorderStyle;
        Region.AddRect(ABorderRect);
      end;
    end
    else
    begin
      ABorderInfo := FSolidBorders[I];
      ABorderInfo.Region.AddRect(ABorderRect);
    end;
  end
  else
  begin
    if ABorder.Orientation = vgrboLeft then
    begin
      AList := FVertBorders;
      ACoord := ABorder.Left;
    end
    else
    begin
      AList := FHorzBorders;
      ACoord := ABorder.Top;
    end;

    I := AList.IndexByStyleAndCoord(ABorderStyle, ACoord);
    if I = -1 then
    begin
      ABorderInfo := AList.Add;
      with ABorderInfo do
      begin
        Coord := ACoord;
        Bounds := ABorderRect;
        Style := ABorderStyle;
        Region.AddRect(ABorderRect);
      end;
    end
    else
    begin
      ABorderInfo := AList[I];
      with ABorderInfo do
      begin
        if Bounds.Left > ABorderRect.Left then
          Bounds.Left := ABorderRect.Left;
        if Bounds.Right < ABorderRect.Right then
          Bounds.Right := ABorderRect.Right;
        if Bounds.Top > ABorderRect.Top then
          Bounds.Top := ABorderRect.Top;
        if Bounds.Bottom < ABorderRect.Bottom then
          Bounds.Bottom := ABorderRect.Bottom;
        Region.AddRect(ABorderRect);
      end;
    end;
  end;
end;

procedure TvgrWorkbookPainter.FindAndCallbackBorders{(ACallbackProc: TvgrFindListItemCallbackProc)};
var
  I, J, N: Integer;
  ARect: TRect;
begin
  I := 0;
  FCurRow := I;
  while I < Page.RowsCount do
  begin
    N := Page.Row[I];
    ARect.Top := N;
    repeat
      Inc(N);
      Inc(I);
    until (I >= Page.RowsCount) or (N <> Page.Row[I]);
    ARect.Bottom := N - 1;

    J := 0;
    FCurCol := J;
    while J < Page.ColsCount do
    begin
      N := Page.Col[J];
      ARect.Left := N;
      repeat
        Inc(N);
        Inc(J);
      until (J >= Page.ColsCount) or (N <> Page.Col[J]);
      ARect.Right := N - 1;

      Worksheet.BordersList.FindAtRectAndCallBack(ARect, ACallbackProc, @ARect);

      FCurCol := J;
    end;
    FCurRow := I;
  end;
end;

function TvgrWorkbookPainter.GetColLeft(AColNumber: Integer): Integer;
var
  ASize: pvgrPainterSize;
begin
  ASize := FindCol(AColNumber);
  if ASize = nil then
  begin
    ASize := FindCol(AColNumber - 1);
    if ASize = nil then
      Result := 0
    else
      Result := ASize.Start + ASize.Size;
  end
  else
    Result := ASize.Start;
end;

function TvgrWorkbookPainter.GetColRight(AColNumber: Integer): Integer;
var
  ASize: pvgrPainterSize;
begin
  ASize := FindCol(AColNumber);
  if ASize = nil then
  begin
    ASize := FindCol(AColNumber - 1);
    if ASize = nil then
      Result := 0
    else
      Result := ASize.Start;
  end
  else
    Result := ASize.Start + ASize.Size;
end;

function TvgrWorkbookPainter.GetColWidth(AColNumber: Integer): Integer;
var
  ASize: pvgrPainterSize;
begin
  ASize := FindCol(AColNumber);
  if ASize = nil then
    Result := CalculateColWidth(AColNumber)
  else
    Result := ASize.Size;
end;

function TvgrWorkbookPainter.GetRowTop(ARowNumber: Integer): Integer;
var
  ASize: pvgrPainterSize;
begin
  ASize := FindRow(ARowNumber);
  if ASize = nil then
  begin
    ASize := FindRow(ARowNumber - 1);
    if ASize = nil then
      Result := 0
    else
      Result := ASize.Start + ASize.Size;
  end
  else
    Result := ASize.Start;
end;

function TvgrWorkbookPainter.GetRowBottom(ARowNumber: Integer): Integer;
var
  ASize: pvgrPainterSize;
begin
  ASize := FindRow(ARowNumber);
  if ASize = nil then
  begin
    ASize := FindRow(ARowNumber + 1);
    if ASize <> nil then
      Result := ASize.Start
    else
      Result := 0
  end
  else
    Result := ASize.Start + ASize.Size;
end;

function TvgrWorkbookPainter.GetRowHeight(ARowNumber: Integer): Integer;
var
  ASize: pvgrPainterSize;
begin
  ASize := FindRow(ARowNumber);
  if ASize = nil then
    Result := CalculateRowHeight(ARowNumber)
  else
    Result := ASize.Size;
end;

function TvgrWorkbookPainter.CalculateRowHeight(ANumber: Integer): Integer;
var
  ARow: IvgrRow;
begin
  ARow := Worksheet.RowsList.Find(ANumber);
  if ARow = nil then
    Result := TwipsToPixelsY(DefaultRowHeight)
  else
    if ARow.Visible then
      Result := TwipsToPixelsY(ARow.Height)
    else
      Result := 0;
end;

function TvgrWorkbookPainter.CalculateColWidth(ANumber: Integer): Integer;
var
  ACol: IvgrCol;
begin
  ACol := Worksheet.ColsList.Find(ANumber);
  if ACol = nil then
    Result := TwipsToPixelsX(DefaultColWidth)
  else
    if ACol.Visible then
      Result := TwipsToPixelsX(ACol.Width)
    else
      Result := 0;
end;

procedure TvgrWorkbookPainter.Cache(AScale: Double);
var
  I: Integer;
begin
  SetLength(FRowSizes, Page.RowsCount);
  for I := 0 to Page.RowsCount - 1 do
    with FRowSizes[I] do
    begin
      Number := Page.Row[I];
      Size := CalculateRowHeight(Number);
      if I = 0 then
        Start := TopOffs
      else
        Start := FRowSizes[I - 1].Start + FRowSizes[I - 1].Size;
    end;

  SetLength(FColSizes, Page.ColsCount);
  for I := 0 to Page.ColsCount - 1 do
    with FColSizes[I] do
    begin
      Number := Page.Col[I];
      Size := CalculateColWidth(Number);
      if I = 0 then
        Start := LeftOffs
      else
        Start := FColSizes[I - 1].Start + FColSizes[I - 1].Size;
    end;
end;

procedure TvgrWorkbookPainter.FindBordersWithinLongRanges(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ARange: IvgrRange;
  ADrawRect, ADrawClipRect,
  ADrawPixelRect, AFullPixelRect, ADrawBorderPixelRect, AFullBorderPixelRect: TRect;
  ABorderExcludeRegion: HRGN;
  ATextWidth, ATextAreaWidth: Integer;
  ARealAlign: TvgrRangeHorzAlign;

  procedure GetLeftLimit(AXCoord, ATextPart: Integer);
  var
    I: Integer;
    ABorder: IvgrBorder;
  begin
    I := GetNumberIndex(AXCoord);
    while ATextPart > 0 do
    begin
      Dec(I);
      if (I >= 0) and IsCellBusy(FColSizes[I].Number, ARange.Top) then
        break;

      ABorder := Worksheet.BordersList.Find(FColSizes[I + 1].Number, ARange.Top, vgrboLeft);
      if ABorder <> nil then
        HideBorder(ABorder);
      if I >= 0 then
      begin
        ABorder := Worksheet.BordersList.Find(FColSizes[I].Number + 1, ARange.Top, vgrboLeft);
        if ABorder <> nil then
          HideBorder(ABorder);
      end;

      if I < 0 then break;
      ATextPart := ATextPart - (FColSizes[I].Size - GetVerticalBorderPartSize(FColSizes[I].Number, ARange.Top, vgrbsLeft));
    end;
  end;

  procedure GetRightLimit(AXCoord, ATextPart: Integer);
  var
    I: Integer;
    ABorder: IvgrBorder;
  begin
    I := GetNumberIndex(AXCoord);
    while ATextPart > 0 do
    begin
      Inc(I);
      if (I <= High(FColSizes)) and IsCellBusy(FColSizes[I].Number, ARange.Top) then break;

      ABorder := Worksheet.BordersList.Find(FColSizes[I - 1].Number + 1, ARange.Top, vgrboLeft);
      if ABorder <> nil then
        HideBorder(ABorder);
      if I <= High(FColSizes) then
      begin
        ABorder := Worksheet.BordersList.Find(FColSizes[I].Number, ARange.Top, vgrboLeft);
        if ABorder <> nil then
          HideBorder(ABorder);
      end;

      if I > High(FColSizes) then break;
      ATextPart := ATextPart - (FColSizes[I].Size - GetVerticalBorderPartSize(FColSizes[I].Number + 1, ARange.Top, vgrbsRight));
    end;
  end;

begin
  ARange := AItem as IvgrRange;
  if IsRangeMayBeLong(ARange) then
  begin
    CalcRange(ARange.Place, PRect(AData)^,
              ADrawRect,
              ADrawPixelRect,
              AFullPixelRect,
              ADrawBorderPixelRect,
              AFullBorderPixelRect,
              ABorderExcludeRegion);
    if ABorderExcludeRegion <> 0 then
      DeleteObject(ABorderExcludeRegion);

    ATextWidth := FCanvas.TextWidth(ARange.DisplayText);
    ATextAreaWidth := AFullBorderPixelRect.Right - AFullBorderPixelRect.Left;
    // calculate clip rect for long range
    ADrawClipRect := ADrawBorderPixelRect; {AFullBorderPixelRect eqaul to ADrawBorderPixelRect}
    ARealAlign := GetAutoAlignForRangeValue(ARange.HorzAlign, ARange.ValueType);
    case ARealAlign of
      vgrhaLeft:
        begin
          GetRightLimit(ARange.Place.Right, ATextWidth - ATextAreaWidth);
        end;
      vgrhaCenter:
        begin
          GetLeftLimit(ARange.Place.Left, (ATextWidth - ATextAreaWidth) div 2);
          GetRightLimit(ARange.Place.Right, (ATextWidth - ATextAreaWidth) div 2);
        end;
      vgrhaRight:
        begin
          GetLeftLimit(ARange.Place.Left, ATextWidth - ATextAreaWidth);
        end;
    end;
  end;
end;

type
  rvgrCheckSizeCallbackData = record
    MaxPartBorderSize: Integer;
    Region: HRGN;
    Side: TvgrBorderSide;
  end;
  pvgrCheckSizeCallbackData = ^rvgrCheckSizeCallbackData;

procedure TvgrWorkbookPainter.CheckSideCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ASize, ATempBorderSize: Integer;
  ABorderRect: TRect;
begin
  with AItem as IvgrBorder, pvgrCheckSizeCallbackData(AData)^ do
  begin
    if (not IsBorderHidden(AItem)) and
       (((Orientation = vgrboLeft) and (Side in [vgrbsLeft, vgrbsRight])) or
        ((Orientation = vgrboTop) and (Side in [vgrbsTop, vgrbsBottom]))) then
    begin
      if Orientation = vgrboLeft then
        ASize := vgr_Functions.GetBorderPartSize(TwipsToPixelsX(Width), Side)
      else
        ASize := vgr_Functions.GetBorderPartSize(TwipsToPixelsY(Width), Side);

      if ASize > 0 then
      begin
        if ASize > MaxPartBorderSize then
          MaxPartBorderSize := ASize;

        CalcBorder(AItem as IvgrBorder, ATempBorderSize, ABorderRect);
        CombineRectAndRegion(Region, ABorderRect, RGN_DIFF);
      end;
    end;
  end;
end;

procedure TvgrWorkbookPainter.CalcRange(const ARangeRect, ADrawingRect: TRect;
                                        var ADrawRect: TRect;
                                        var ADrawPixelRect: TRect;
                                        var AFullPixelRect: TRect;
                                        var ADrawBorderPixelRect: TRect;
                                        var AFullPixelBorderRect: TRect;
                                        var ABorderExcludeRegion: HRGN);
var
  I: Integer;
  ACallbackData: rvgrCheckSizeCallbackData;

  procedure CheckSide(ALeft, ATop, ARight, ABottom: Integer; AOrientation: TvgrBorderOrientation; ASide: TvgrBorderSide; var AResult: Integer);
  begin
    ACallbackData.MaxPartBorderSize := 0;
    ACallbackData.Side := ASide;
    Worksheet.BordersList.FindAndCallBack(Rect(ALeft, ATop, ARight, ABottom), CheckSideCallback, @ACallbackData);
    if ASide in [vgrbsRight, vgrbsBottom] then
      AResult := AResult - ACallbackData.MaxPartBorderSize
    else
      AResult := AResult + ACallbackData.MaxPartBorderSize;
  end;

begin
  ADrawRect.Left := Max(ADrawingRect.Left, ARangeRect.Left);
  ADrawRect.Top := Max(ADrawingRect.Top, ARangeRect.Top);
  ADrawRect.Right := Min(ADrawingRect.Right, ARangeRect.Right);
  ADrawRect.Bottom := Min(ADrawingRect.Bottom, ARangeRect.Bottom);

  ADrawPixelRect.Left := GetColLeft(ADrawRect.Left);
  ADrawPixelRect.Top := GetRowTop(ADrawRect.Top);
  ADrawPixelRect.Right := GetColRight(ADrawRect.Right);
  ADrawPixelRect.Bottom := GetRowBottom(ADrawRect.Bottom);

  AFullPixelRect := ADrawPixelRect;
  for I := ADrawRect.Left - 1 downto ARangeRect.Left do
    AFullPixelRect.Left := AFullPixelRect.Left - GetColWidth(I);
  for I := ADrawRect.Top - 1 downto ARangeRect.Top do
    AFullPixelRect.Top := AFullPixelRect.Top - GetRowHeight(I);
  for I := ADrawRect.Right + 1 to ARangeRect.Right do
    AFullPixelRect.Right := AFullPixelRect.Right + GetColWidth(I);
  for I := ADrawRect.Bottom + 1 to ARangeRect.Bottom do
    AFullPixelRect.Bottom := AFullPixelRect.Bottom + GetRowHeight(I);

  // check borders
  AFullPixelBorderRect := AFullPixelRect;
  ABorderExcludeRegion := CreateRectRgnIndirect(AFullPixelBorderRect);
//  ABorderExcludeRegion := CreateRectRgnIndirect(ADrawPixelRect);
  ACallbackData.Region := ABorderExcludeRegion;
  with ARangeRect do
    CheckSide(Left, Top, Left, Bottom, vgrboLeft, vgrbsLeft, AFullPixelBorderRect.Left);
  with ARangeRect do
    CheckSide(Left, Top, Right, Top, vgrboTop, vgrbsTop, AFullPixelBorderRect.Top);
  with ARangeRect do
    CheckSide(Right + 1, Top, Right + 1, Bottom, vgrboLeft, vgrbsRight, AFullPixelBorderRect.Right);
  with ARangeRect do
    CheckSide(Left, Bottom + 1, Right, Bottom + 1, vgrboTop, vgrbsBottom, AFullPixelBorderRect.Bottom);
  CombineRectAndRegion(ABorderExcludeRegion, ADrawPixelRect, RGN_AND);

  ADrawBorderPixelRect := ADrawPixelRect;
  if ADrawRect.Left = ARangeRect.Left then
    ADrawBorderPixelRect.Left := AFullPixelBorderRect.Left;
  if ADrawRect.Top = ARangeRect.Top then
    ADrawBorderPixelRect.Top := AFullPixelBorderRect.Top;
  if ADrawRect.Right = ARangeRect.Right then
    ADrawBorderPixelRect.Right := AFullPixelBorderRect.Right;
  if ADrawRect.Bottom = ARangeRect.Bottom then
    ADrawBorderPixelRect.Bottom := AFullPixelBorderRect.Bottom;
end;

procedure TvgrWorkbookPainter.PaintRange(ACanvas: TCanvas; ARange: IvgrRange; const ADrawingRect: TRect);
var
  X, Y: Integer;
  ADrawRect, ADrawClipRect,
  ADrawPixelRect, AFullPixelRect, ADrawBorderPixelRect, AFullBorderPixelRect: TRect;
  ABorderExcludeRegion: HRGN;
  APreviosClipRegion: HRGN;
  AOldTransparentMode: Boolean;
  ADisplayText: string;
  ATextAreaWidth, ATextWidth: Integer;
  ARealAlign: TvgrRangeHorzAlign;
  AOldBrushOrg: TPoint;

  procedure GetLeftLimit(AXCoord, ATextPart: Integer; var ALimit: Integer);
  var
    I, APartSize: Integer;
  begin
    I := GetNumberIndex(AXCoord);
    while (I > 0) and
          (ATextPart > 0) and
          not IsCellBusy(FColSizes[I - 1].Number, ARange.Top) do
    begin
      Dec(I);
      APartSize := GetVerticalBorderPartSize(FColSizes[I].Number, ARange.Top, vgrbsLeft);
      ATextPart := ATextPart - (FColSizes[I].Size - APartSize);
//      if ATextPart > 0 then
        ALimit := ALimit - (FColSizes[I].Size - APartSize);
    end;
  end;

  procedure GetRightLimit(AXCoord, ATextPart: Integer; var ALimit: Integer);
  var
    I, APartSize: Integer;
  begin
    I := GetNumberIndex(AXCoord);
    while (I < High(FColSizes)) and
          (ATextPart > 0) and
          not IsCellBusy(FColSizes[I + 1].Number, ARange.Top) do
    begin
      Inc(I);
      APartSize := GetVerticalBorderPartSize(FColSizes[I].Number + 1, ARange.Top, vgrbsRight);
      ATextPart := ATextPart - (FColSizes[I].Size - APartSize);
//      if ATextPart > 0 then
        ALimit := ALimit + (FColSizes[I].Size - APartSize);
    end;
    if I >= High(FColSizes) then
    begin
      ALimit := ADrawPixelRect.Right + ATextWidth;
    end;
  end;

begin
  ADisplayText := ARange.DisplayText;

  CalcRange(ARange.Place, ADrawingRect,
            ADrawRect,
            ADrawPixelRect,
            AFullPixelRect,
            ADrawBorderPixelRect,
            AFullBorderPixelRect,
            ABorderExcludeRegion);

  // assign font
  with ARange.Font do
  begin
    ACanvas.Font.Name := Name;
    ACanvas.Font.Height := MulDiv(-Size, Owner.PixelsPerInchX, 72);
    ACanvas.Font.Pitch := Pitch;
    ACanvas.Font.Style := Style;
    ACanvas.Font.Charset := Charset;
    ACanvas.Font.Color := Color;
  end;

  // fill background
  if not ((ARange.FillPattern = bsSolid) and
          ((GetRGBColor(ARange.FillBackColor) = GetRGBColor(Owner.DefaultRangeBackgroundColor)) or
           (ARange.FillBackColor = clNone))) then
  begin
    SetBrushOrgEx(ACanvas.Handle, BrushOrg.X, BrushOrg.Y, @AOldBrushOrg);
    ACanvas.Brush.Style := ARange.FillPattern;
    if ACanvas.Brush.Style = bsSolid then
      ACanvas.Brush.Color := ARange.FillBackColor
    else
    begin
      ACanvas.Brush.Color := ARange.FillForeColor;
      SetBkColor(ACanvas.Handle, GetRGBColor(ARange.FillBackColor));
    end;
    AOldTransparentMode := SetCanvasTransparentMode(ACanvas, False);
    FillRegion(ACanvas, ABorderExcludeRegion);
    SetCanvasTransparentMode(ACanvas, AOldTransparentMode);
    SetBrushOrgEx(ACanvas.Handle, AOldBrushOrg.X, AOldBrushOrg.Y, nil);
  end;

  AOldTransparentMode := SetCanvasTransparentMode(ACanvas, True);

  if IsRangeMayBeLong(ARange) then
  begin
    ATextWidth := ACanvas.TextWidth(ADisplayText);
    ATextAreaWidth := AFullBorderPixelRect.Right - AFullBorderPixelRect.Left;
    // calculate clip rect for long range
    ADrawClipRect := ADrawBorderPixelRect; {AFullBorderPixelRect eqaul to ADrawBorderPixelRect}
    ARealAlign := GetAutoAlignForRangeValue(ARange.HorzAlign, ARange.ValueType);
    case ARealAlign of
      vgrhaLeft:
        begin
          GetRightLimit(ARange.Place.Right, ATextWidth - ATextAreaWidth, ADrawClipRect.Right);
        end;
      vgrhaCenter:
        begin
          GetLeftLimit(ARange.Place.Left, ATextWidth - ATextAreaWidth div 2, ADrawClipRect.Left);
          GetRightLimit(ARange.Place.Right, ATextWidth - ATextAreaWidth div 2, ADrawClipRect.Right);
        end;
      vgrhaRight:
        begin
          GetLeftLimit(ARange.Place.Left, ATextWidth - ATextAreaWidth, ADrawClipRect.Left);
        end;
    end;

    with AFullBorderPixelRect do
    begin
      case ARange.VertAlign of
        vgrvaCenter:
          Y := Top + (Bottom - Top - Canvas.TextHeight(ADisplayText)) div 2;
        vgrvaBottom:
          Y := Bottom - Canvas.TextHeight(ADisplayText);
      else
        Y := Top;
      end;
      case ARealAlign of
        vgrhaCenter:
          X := HorzCenterOfRect(AFullBorderPixelRect) - ATextWidth div 2;
        vgrhaRight:
          X := Right - ATextWidth;
      else
        X := Left;
      end;
    end;

    ExtTextOut(Canvas.Handle, X, Y,
               ETO_CLIPPED,
               @ADrawClipRect,
               PChar(ADisplayText),
               Length(ADisplayText),
               nil);
    CombineRectAndRegion(ABorderExcludeRegion, ADrawClipRect, RGN_OR);
  end
  else
  begin
    // hide borders within range area
    for X := ADrawRect.Left to ADrawRect.Right do
      for Y := ADrawRect.Top to ADrawRect.Bottom do
      begin
        if X <> ADrawRect.Left then
          HideBorder(X, Y, vgrboLeft);
        if Y <> ADrawRect.Top then
          HideBorder(X, Y, vgrboTop);
      end;
    
    // draw text
    APreviosClipRegion := GetClipRegion(ACanvas);
    SetClipRect(ACanvas, ADrawBorderPixelRect);
    vgrDrawText(ACanvas, ADisplayText, AFullBorderPixelRect, ARange.WordWrap, GetAutoAlignForRangeValue(ARange.HorzAlign, ARange.ValueType), ARange.VertAlign, ARange.Angle);
    SetClipRegion(ACanvas, APreviosClipRegion);
    if APreviosClipRegion <> 0 then
      DeleteObject(APreviosClipRegion);
  end;

  SetCanvasTransparentMode(ACanvas, AOldTransparentMode);

//  // Add range region
//  CombineRegionAndRegion(FRangesRegion, ABorderExcludeRegion, RGN_OR);
  DeleteObject(ABorderExcludeRegion);
end;

function TvgrWorkbookPainter.GetPagePixelWidth(APage: TvgrPrintPage; AScale: Double): Integer;
begin
  FScale := AScale;
  with Worksheet.PageProperties do
    Result := Round(TwipsToPixelsX(Width - Margins.Right - Margins.Left) * AScale);
end;

function TvgrWorkbookPainter.GetPagePixelHeight(APage: TvgrPrintPage; AScale: Double): Integer;
begin
  FScale := AScale;
  with Worksheet.PageProperties do
    Result := Round(TwipsToPixelsY(Height - Margins.Bottom - Margins.Top) * AScale);
end;

procedure TvgrWorkbookPainter.DrawHeaderFooter(ACanvas: TCanvas;
                                               AHeaderFooter: TvgrPageHeaderFooter;
                                               const ARect: TRect);

  function GetHeaderFooterText(const S: string): string;
  begin
    Result := ParseHeaderFooterText(S, Page.PageNumber + 1, Page.PageCount, Worksheet.Title)
  end;

begin
  if (ARect.Right <= ARect.Left) or (ARect.Bottom <= ARect.Top) then
    exit; // is invisible

  //
  if (AHeaderFooter.BackColor <> clNone) and
     (GetRGBColor(AHeaderFooter.BackColor) <> GetRGBColor(Owner.DefaultRangeBackgroundColor)) then
  begin
    ACanvas.Brush.Style := bsSolid;
    ACanvas.Brush.Color := AHeaderFooter.BackColor;
    ACanvas.FillRect(ARect);
  end;

  if AHeaderFooter.LeftSection <> '' then
  begin
    ACanvas.Font.Assign(AHeaderFooter.LeftFont);
    vgrDrawText(ACanvas, GetHeaderFooterText(AHeaderFooter.LeftSection), ARect, true, vgrhaLeft, vgrvaTop, 0);
  end;
  if AHeaderFooter.CenterSection <> '' then
  begin
    ACanvas.Font.Assign(AHeaderFooter.CenterFont);
    vgrDrawText(ACanvas, GetHeaderFooterText(AHeaderFooter.CenterSection), ARect, true, vgrhaCenter, vgrvaTop, 0);
  end;
  if AHeaderFooter.RightSection <> '' then
  begin
    ACanvas.Font.Assign(AHeaderFooter.RightFont);
    vgrDrawText(ACanvas, GetHeaderFooterText(AHeaderFooter.RightSection), ARect, true, vgrhaRight, vgrvaTop, 0);
  end;
end;

procedure TvgrWorkbookPainter.DrawPage(X, Y: Integer; ACanvas: TCanvas; APage: TvgrPrintPage; AScale: Double; const ABrushOrg: TPoint);
//var
//  AClipRegion: HRGN;
//  AOldClipRegion: HRGN;
var
  ARect: TRect;
{$IFDEF VGR_PAINTERDEBUG}
  AFile: TextFile;
  I: Integer;
  S: string;
{$ENDIF}
begin
  // Prepare data
  FLeftOffs := X;
  FTopOffs := Y;
  FCurRow := 0;
  FCurCol := 0;
  FCanvas := ACanvas;
  FScale := AScale;
  FPage := APage;
  FBrushOrg := ABrushOrg;
  FWidth := GetPagePixelWidth(APage, AScale);
  FHeight := GetPagePixelHeight(APage, AScale);

{$IFDEF VGR_PAINTERDEBUG}
  AssignFile(AFile, 'c:\painterdebug.txt');
  Rewrite(AFile);
  WriteLn(AFile, Format('Page: %d, ColCount: %d, RowCount: %d', [APage.PageNumber, APage.ColsCount, APage.RowsCount]));
  S := '';
  for I := 0 to APage.ColsCount - 1 do
    S := S + '  ' + IntToStr(APage.Col[I]);
  WriteLn(AFile, 'Cols: ' + S);
  S := '';
  for I := 0 to APage.RowsCount - 1 do
    S := S + '  ' + IntToStr(APage.Row[I]);
  WriteLn(AFile, 'Rows: ' + S);
  CloseFile(AFile);
{$ENDIF}

  // Fill other area
  if Owner.BackgroundColor <> clNone then
  begin
    SetBrushOrgEx(ACanvas.Handle, 0, 0, nil);
    ACanvas.Brush.Style := bsSolid;
    ACanvas.Brush.Color := Owner.BackgroundColor;
    ACanvas.FillRect(Rect(LeftOffs, TopOffs, LeftOffs + FWidth, TopOffs + FHeight));
  end;

  // draw header
  ARect := Rect(X,
                Y,
                X + FWidth,
                Y + Round(ConvertTwipsToPixelsY(FWorksheet.PageProperties.Header.Height) * AScale));
  DrawHeaderFooter(FCanvas,
                   FWorksheet.PageProperties.Header,
                   ARect);
  Y := Y + (ARect.Bottom - ARect.Top);
  FHeight := FHeight - (ARect.Bottom - ARect.Top);
  FTopOffs := Y;

  // draw footer
  ARect := Rect(X,
                Y + FHeight - Round(ConvertTwipsToPixelsY(FWorksheet.PageProperties.Footer.Height) * AScale),
                X + FWidth,
                Y + FHeight);
  DrawHeaderFooter(FCanvas,
                   FWorksheet.PageProperties.Footer,
                   ARect);
  FHeight := FHeight - (ARect.Bottom - ARect.Top);

  // cache sizes of columns and rows
  Cache(AScale);

  try
    // Search and store all borders which are placed within "long" ranges
    FindAndCallbackRanges(FindBordersWithinLongRanges);

    // Draw normal ranges
    FindAndCallbackRanges(PaintNormalRangesCallback);

    // Draw long ranges
    FindAndCallbackRanges(PaintLongRangesCallback);

//    AOldClipRegion := GetClipRegion(ACanvas);
//    try
//      // Calculate and set new clip region
//      AClipRegion := CreateRectRgn(LeftOffs, TopOffs, LeftOffs + FWidth, TopOffs + FHeight);
//      CombineRegionAndRegion(AClipRegion, FRangesRegion, RGN_DIFF);
//      SetClipRegion(ACanvas, AClipRegion);
//      DeleteObject(AClipRegion);

      // Draw borders
      try
        // 1. Cache borders info
        FindAndCallbackBorders(PaintBordersCallback);
        // 2. Paint borders
        PaintSolidBorders(FSolidBorders);
        PaintBorders(FHorzBorders, True);
        PaintBorders(FVertBorders, False);
      finally
        FSolidBorders.Clear;
        FHorzBorders.Clear;
        FVertBorders.Clear;
      end;
//    finally
//      // restore old clip region
//      SetClipRegion(ACanvas, AOldClipRegion);
//      if AOldClipRegion <> 0 then
//        DeleteObject(AOldClipRegion);
//    end;
  finally
    FHiddenBorders.Clear;
  end;
end;

procedure InitIntegerArray(ABorderStyle: TvgrBorderStyle; AParts: array of Integer);
var
  I: Integer;
begin
  SetLength(APens[ABorderStyle], Length(AParts));
  for I := 0 to High(AParts) do
    APens[ABorderStyle][I] := AParts[I];
end;

initialization

  InitIntegerArray(vgrbsSolid, [-1]);
  InitIntegerArray(vgrbsDash, [18, 6]);
  InitIntegerArray(vgrbsDot, [3, 3]);
  InitIntegerArray(vgrbsDashDot, [9, 6, 3, 6]);
  InitIntegerArray(vgrbsDashDotDot, [9, 3, 3, 3, 3, 3]);

end.



