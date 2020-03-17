{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains the TvgrPageMaker class, which provide functionality for divide rows and cols in worksheet to pages for previewing and printing.
Instance of TvgrPageMaker class creates and used when user want to preview
or print worksheet.
  This do main method of this class - Prepare.
  As result of calling Prepare we can read property Pages.
See also:
  TvgrPageMaker, TvgrPrintPage}
unit vgr_PageMaker;

interface

{$I vtk.inc}

//{$DEFINE VGR_DBG_MAKE}
uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, {$IFDEF VTK_D6_OR_D7}Types,{$ENDIF} {$IFDEF VGR_DBG_MAKE} windows, sysutils,{$ENDIF}

  vgr_CommonClasses, vgr_DataStorage, vgr_PageProperties;

type
{Array for storaging page vectors (rows or cols)
Syntax:
  TvgrPageVectIndexes = array of Integer;}
  TvgrPageVectIndexes = array of Integer;

  TvgrPageMaker = class;
  TvgrPrintPage = class;

  /////////////////////////////////////////////////
  //
  // TvgrPagesList
  //
  ////////////////////////////////////////////////
{Used internally in GridReport as list of TvgrPrintPage instances.
See also:
  TvgrPageMaker, TvgrPrintPage}
  TvgrPagesList = class(TList)
  private
    FPageMaker : TvgrPageMaker;
    function GetItm(Index : integer) : TvgrPrintPage;
    procedure AddInList(APage: TvgrPrintPage);
  public
{Creates an instance of the TvgrPagesList class.
Parameters:
  APageMaker - PageMaker, which creates this list
See also:
  TvgrPageMaker}
    constructor Create(APageMaker: TvgrPageMaker);
{Frees an instance of the TvgrPagesList class.}
    destructor Destroy; override;
{Clears the list and frees all items.}
    procedure Clear; override;
{Returns TvgrPrintPage item by Index.
See also:
  TvgrPrintPage}
    property Items[Index: integer] : TvgrPrintPage read GetItm;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPrintPage
  //
  ////////////////////////////////////////////////
{This class contains properties for describing page for printing or previewing.
Used internally by printing and previewing engines.}
  TvgrPrintPage = class(TObject)
  private
    FPageNumber: Integer;
    FPageMaker: TvgrPageMaker;
    function GetPageNumber: Integer;
    function GetPageCount: Integer;
    function GetRow(Index: Integer): Integer;
    function GetCol(Index: Integer): Integer;
    function GetRowsCount: Integer;
    function GetColsCount: Integer;
  public
{Returns the number of page.}
    property PageNumber: Integer read GetPageNumber;
{Returns the count of pages.}
    property PageCount: Integer read GetPageCount;
{Returns the index of this row on worksheet by index of this row on current page.}
    property Row[Index: Integer]: Integer read GetRow;
{Returns the index of this column on worksheet by index of this col on current page.}
    property Col[Index: Integer]: Integer read GetCol;
{Returns the number of rows, placed on current page.}
    property RowsCount: Integer read GetRowsCount;
{Returns count of columns, placed on current page.}
    property ColsCount: Integer read GetColsCount;
  end;

{Order of printing or previewing pages of worksheet.
Syntax:
  TvgrPagesOrder = (vgrpoBottomRight, vgrpoRightBottom);
Items:
  vgrpoBottomRight  - directions is from top to bottom, then to right
  vgrpoRightBottom  - directions is from left to right, then to bottom
}
  TvgrPagesOrder = (vgrpoBottomRight, vgrpoRightBottom);

{Used for firing information about progress of making pages.
Parameters:
  Sender - Instance of TvgrPageMaker
  AStage - Stage of Making 0 or 1
  APosition - position of progress in the current stage
  AMax - total points in the current stage
See also:
  TvgrPageMaker}
  TvgrPageMakingProgress = procedure(Sender: TObject; AStage, APosition, AMax: Integer) of object;

  /////////////////////////////////////////////////
  //
  // TvgrPageMaker
  //
  ////////////////////////////////////////////////
{Realizes functions, which need for printing and previewing on page with given sizes.
Describes numer of pages, numbers and orders of the rows and cols of worksheet,
which be showed on theese pages.
Sizes and margins on pages defines by PageProperties property;
See also:
  TvgrWorksheet, TvgrPageMakingProgress, TvgrPagesOrder, TvgrPagesList,
  TvgrPrintPage, TvgrPageProperties}
  TvgrPageMaker = class(TPersistent)
  private
    FWorksheet: TvgrWorksheet;
    FOnProgress: TvgrPageMakingProgress;
    FOnComplete: TNotifyEvent;
    FOrder: TvgrPagesOrder;
    FPages: TvgrPagesList;
    FRows: TvgrPageVectIndexes;
    FCols: TvgrPageVectIndexes;
    FHorzRulers: TvgrPageVectIndexes;
    FVertRulers: TvgrPageVectIndexes;
    FCurrHorzIndex: Integer;
    FCurrVertIndex: Integer;
    FComplete: Boolean;

    procedure InternalPrepareDimension(var ARulers, AVectIndexes: TvgrPageVectIndexes; AVectors: TvgrVectors; ASections: TvgrSections; ADefaultSize, APageSize, AEdge: Integer; AStopOnNewPage: Boolean);
    procedure PrepareHorzDimension;
    procedure PrepareVertDimension;

    procedure SetComplete(Value: Boolean);
    procedure Init;
    procedure InternalPrepareNextPage;

    procedure AddIdxItem(var AVectIndexes: TvgrPageVectIndexes; Value: Integer);
    function GetBottomEdge: Integer;
    function GetRightEdge: Integer;
    function GetHorzDimLen: Integer;
    function GetVertDimLen: Integer;
    function GetRow(APage, AIndex: Integer): Integer;
    function GetCol(APage, AIndex: Integer): Integer;
    function GetPageRowsCount(APage: Integer): Integer;
    function GetPageColsCount(APage: Integer): Integer;
    function GetPageVectorsCount(const ARulers: TvgrPageVectIndexes; APage: Integer): Integer;
    function GetPageVector(const ARulers, AVectors: TvgrPageVectIndexes; APage, AIndex: Integer): Integer;
    function CurrHorzIndex(APageNumber: Integer): Integer;
    function CurrVertIndex(APageNumber: Integer): Integer;
    function GetPage(Index: Integer): TvgrPrintPage;
    function GetPagesCount: Integer;
    function GetPrepareComplete: Boolean;
    function GetPageProperties: TvgrPageProperties;
    function GetOrder: TvgrPagesOrder;
    procedure SetOrder(Value: TvgrPagesOrder);
  protected  
    procedure PrepareFirst;
    function PrepareNext: Integer;
    property PrepareComplete: Boolean read GetPrepareComplete;
  public
{Creates the instance of TvgrPageMaker
Parameters:
  AWorksheet - worksheet for deviding on pages}
    constructor Create(AWorksheet: TvgrWorksheet);
{Frees the TvgrPageMaker instance}
    destructor Destroy; override;
{Starts the procedure of defining pages. Call this method before read list of pages and
this properties.}
    procedure Prepare;
{Resulting list of pages, is generated after executing of Prepare method.
Syntax:
  property Pages[Index: Integer]: TvgrPrintPage read;
Parameters:
  Index - Number of a page}
    property Pages[Index: Integer]: TvgrPrintPage read GetPage;
{Count of a pages on worksheet, maked as result of Prepare method.
Syntax:
  property PagesCount: Integer read;}
    property PagesCount: Integer read GetPagesCount;
{Properties: size, marging and orientation of pages, on which need to place worksheet.
Syntax:
  property PageProperties: TvgrPageProperties read;
See also:
  TvgrPageProperties}
    property PageProperties: TvgrPageProperties read GetPageProperties;
{Worksheet, which deviding on pages.
Syntax:
  property Worksheet: TvgrWorksheet read;
See also:
  TvgrWorksheet}
    property Worksheet: TvgrWorksheet read FWorksheet;
{Order of printing or previewing pages of worksheet.
Syntax:
  property Order: TvgrPagesOrder read write;
See also:
  TvgrPagesOrder}
    property Order: TvgrPagesOrder read GetOrder write SetOrder;
{This event fired when new page is maked.
Intended for different progress bars, etc.
Syntax:
  property OnProgress: TvgrPageMakingProgress read write;
See also:
  TvgrPageMakingProgress}
    property OnProgress: TvgrPageMakingProgress read FOnProgress write FOnProgress;
{This event fired when prepare process is done.
Syntax:
  property OnComplete: TNotifyEvent read write;}
    property OnComplete: TNotifyEvent read FOnComplete write FOnComplete;
  end;

implementation

type

  /////////////////////////////////////////////////
  //
  // TvgrPMTopVector
  //
  ////////////////////////////////////////////////
  TvgrPMTopVector = class
    PrintOnTop: Boolean;
    PrintOnBottom: Boolean;
    LinkedWithPrevios: Boolean;
    LinkedWithNext: Boolean;
    StartPos: Integer;
    EndPos: Integer;
    StartArea: Integer;
    EndArea: Integer;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPMTopVectors
  //
  ////////////////////////////////////////////////
  TvgrPMTopVectors = class(TList)
  private
    FSections: TvgrSections;
    FCurrentSectionIndex: Integer;

    constructor Create(ASections: TvgrSections);
    destructor Destroy; override;
    function GetItm(Index : Integer) : TvgrPMTopVector;
    procedure AddVector(ASection: IvgrSection);
    procedure Init;

    property Items[Index : integer] : TvgrPMTopVector read GetItm;
  end;

/////////////////////////////////////////////////
//
// TvgrPMTopVectors
//
////////////////////////////////////////////////
constructor TvgrPMTopVectors.Create(ASections: TvgrSections);
begin
  FSections := ASections;
  Init;
  FCurrentSectionIndex := -1;
end;

destructor TvgrPMTopVectors.Destroy;
var
  I : Integer;
begin
  for I := 0 to Count-1 do
    TObject(Items[I]).Free;
  inherited;
end;

function TvgrPMTopVectors.GetItm(Index : Integer) : TvgrPMTopVector;
begin
  Result := TvgrPMTopVector(inherited Items[Index]);
end;

procedure TvgrPMTopVectors.AddVector(ASection: IvgrSection);
var
  AVector: TvgrPMTopVector;
  AOuter: IvgrSection;
begin
  AVector := TvgrPMTopVector.Create;
  with AVector do
  begin
    PrintOnTop := ASection.RepeatOnPageTop;
    PrintOnBottom := ASection.RepeatOnPageBottom;
    StartPos := ASection.StartPos;
    EndPos := ASection.EndPos;
    AOuter := ASection.Parent;
    if AOuter = nil then
    begin
      StartArea := 0;
      EndArea := MaxInt;
    end
    else
    begin
      StartArea := AOuter.StartPos;
      EndArea := AOuter.EndPos;
    end;
  end;
  Add(AVector);
end;

procedure TvgrPMTopVectors.Init;
var
  I: Integer;
  ASection: IvgrSection;
begin
  for I := 0 to FSections.Count - 1 do
  begin
    ASection := FSections.ByIndex[I];
    if ASection.RepeatOnPageTop or ASection.RepeatOnPageBottom then
      AddVector(ASection);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrPagesList
//
////////////////////////////////////////////////
constructor TvgrPagesList.Create(APageMaker : TvgrPageMaker);
begin
  inherited Create;
  FPageMaker := APageMaker;
end;

destructor TvgrPagesList.Destroy;
begin
  Clear;
  inherited;
end;

procedure TvgrPagesList.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  inherited;
end;

function TvgrPagesList.GetItm(Index : integer) : TvgrPrintPage;
begin
  Result := TvgrPrintPage(inherited Items[Index]);
end;

procedure TvgrPagesList.AddInList(APage: TvgrPrintPage);
begin
  Add(APage);
end;

/////////////////////////////////////////////////
//
// TvgrPrintPage
//
////////////////////////////////////////////////
function TvgrPrintPage.GetPageNumber: Integer;
begin
  Result := FPageNumber;
end;

function TvgrPrintPage.GetPageCount: Integer;
begin
  Result := FPageMaker.PagesCount;
end;

function TvgrPrintPage.GetRow(Index: Integer): Integer;
begin
  Result := FPageMaker.GetRow(PageNumber, Index);
end;

function TvgrPrintPage.GetCol(Index: Integer): Integer;
begin
  Result := FPageMaker.GetCol(PageNumber, Index);
end;

function TvgrPrintPage.GetRowsCount: Integer;
begin
  Result := FPageMaker.GetPageRowsCount(PageNumber);
end;

function TvgrPrintPage.GetColsCount: Integer;
begin
  Result := FPageMaker.GetPageColsCount(PageNumber);
end;


/////////////////////////////////////////////////
//
// TvgrPageMaker
//
////////////////////////////////////////////////
constructor TvgrPageMaker.Create(AWorksheet: TvgrWorksheet);
begin
  FWorksheet := AWorksheet;
  FPages := TvgrPagesList.Create(Self);
end;

destructor TvgrPageMaker.Destroy;
begin
  FPages.Free;
  inherited;
end;

procedure TvgrPageMaker.Prepare;
begin
  Init;
  while not PrepareComplete do
    PrepareNext;
end;

procedure TvgrPageMaker.PrepareFirst;
begin
  Init;
  InternalPrepareNextPage;
end;

function TvgrPageMaker.PrepareNext: Integer;
begin
  InternalPrepareNextPage;
  Result := PagesCount - 1;
end;

function TvgrPageMaker.GetPage(Index: Integer): TvgrPrintPage;
begin
  if Index < FPages.Count then
    Result := FPages.Items[Index]
  else
    Result := nil;
end;

function TvgrPageMaker.GetPagesCount: Integer;
begin
  Result := FPages.Count;
end;

function TvgrPageMaker.GetPrepareComplete: Boolean;
begin
  Result := FPages.Count = (GetHorzDimLen * GetVertDimLen)
end;

function TvgrPageMaker.GetPageProperties: TvgrPageProperties;
begin
  if FWorksheet <> nil then
    Result := FWorksheet.PageProperties
  else
    Result := nil;
end;

function TvgrPageMaker.GetOrder: TvgrPagesOrder;
begin
  Result := FOrder;
end;

procedure TvgrPageMaker.SetOrder(Value: TvgrPagesOrder);
begin
  FOrder := Value;
end;

procedure TvgrPageMaker.SetComplete(Value: Boolean);
begin
  FComplete := Value;
  if Value then
    if Assigned(FOnComplete) then
      FOnComplete(Self);
end;

procedure TvgrPageMaker.Init;
begin
  SetComplete(False);
  SetLength(FRows,0);
  SetLength(FCols,0);
  SetLength(FHorzRulers,0);
  SetLength(FVertRulers,0);
  FCurrHorzIndex := 0;
  FCurrVertIndex := 0;
  PrepareHorzDimension;
  PrepareVertDimension;
  FPages.Clear;
end;

procedure TvgrPageMaker.InternalPrepareNextPage;
var
  APage: TvgrPrintPage;
begin
  APage := TvgrPrintPage.Create;
  FPages.AddInList(APage);
  with APage do
  begin
    FPageMaker := Self;
    FPageNumber := PagesCount - 1;
  end;
end;

procedure TvgrPageMaker.AddIdxItem(var AVectIndexes: TvgrPageVectIndexes; Value: Integer);
var
  I: Integer;
begin
  I := Length(AVectIndexes);
  SetLength(AVectIndexes, I + 1);
  AVectIndexes[I] := Value;
end;

function TvgrPageMaker.GetBottomEdge: Integer;
begin
  Result := FWorksheet.Dimensions.Bottom;
end;

function TvgrPageMaker.GetRightEdge: Integer;
begin
  Result := FWorksheet.Dimensions.Right;
end;

function TvgrPageMaker.GetHorzDimLen: Integer;
begin
  Result := Length(FHorzRulers);
end;

function TvgrPageMaker.GetVertDimLen: Integer;
begin
  Result := Length(FVertRulers);
end;

function TvgrPageMaker.GetRow(APage, AIndex: Integer): Integer;
begin
   Result := GetPageVector(FHorzRulers, FRows, CurrVertIndex(APage), AIndex);
end;

function TvgrPageMaker.GetCol(APage, AIndex: Integer): Integer;
begin
  Result := GetPageVector(FVertRulers, FCols, CurrHorzIndex(APage), AIndex);
end;

function TvgrPageMaker.GetPageRowsCount(APage: Integer): Integer;
begin
  Result := GetPageVectorsCount(FHorzRulers, CurrVertIndex(APage));
end;

function TvgrPageMaker.GetPageColsCount(APage: Integer): Integer;
begin
  Result := GetPageVectorsCount(FVertRulers, CurrHorzIndex(APage));
end;

function TvgrPageMaker.GetPageVectorsCount(const ARulers: TvgrPageVectIndexes; APage: Integer): Integer;
begin
  if Length(ARulers) > (APage) then
  begin
    if APage = 0 then
      Result := ARulers[APage] + 1
    else
      Result := ARulers[APage] - ARulers[APage-1];
  end
  else
    Result := 0;
end;

function TvgrPageMaker.GetPageVector(const ARulers, AVectors: TvgrPageVectIndexes; APage, AIndex: Integer): Integer;
var
  I : Integer;
begin
  if Length(ARulers) >= APage then
  begin
    if APage = 0 then
      Result := AVectors[AIndex]
    else
    begin
      I := ARulers[APage - 1];
      Result := AVectors[I+AIndex+1];
    end;
  end
  else
    Result := -1;
end;

function TvgrPageMaker.CurrHorzIndex(APageNumber: Integer): Integer;
begin
  if Order = vgrpoBottomRight then
    Result := APageNumber div GetHorzDimLen
  else
    Result := APageNumber mod GetVertDimLen;
end;

function TvgrPageMaker.CurrVertIndex(APageNumber: Integer): Integer;
begin
  if Order = vgrpoBottomRight then
    Result := APageNumber mod GetHorzDimLen
  else
    Result := APageNumber div GetVertDimLen;
end;

procedure TvgrPageMaker.PrepareHorzDimension;
begin
  InternalPrepareDimension(FHorzRulers, FRows, FWorksheet.RowsList, FWorksheet.HorzSectionsList, PageProperties.Defaults.RowHeight, PageProperties.ClientHeight, GetBottomEdge, False);
end;

procedure TvgrPageMaker.PrepareVertDimension;
begin
  InternalPrepareDimension(FVertRulers, FCols, FWorksheet.ColsList, FWorksheet.VertSectionsList, PageProperties.Defaults.ColWidth, PageProperties.ClientWidth, GetRightEdge, False);
end;

procedure TvgrPageMaker.InternalPrepareDimension(var ARulers, AVectIndexes: TvgrPageVectIndexes; AVectors: TvgrVectors; ASections: TvgrSections; ADefaultSize, APageSize, AEdge: Integer; AStopOnNewPage: Boolean);
var
  I: Integer;
  ACurPageLen: Integer;
  ACurDimVectIdx: Integer;
  ATopVectors: TvgrPMTopVectors;

  function GetVLen(Index: Integer): Integer;
  var
    AVector: IvgrVector;
  begin
    AVector := AVectors.Find(Index);
    if AVector = nil then
      Result := ADefaultSize
    else
    if AVector.Visible then
      Result := AVector.Size
    else
      Result := 0;
  end;

  procedure BeginNewPage;
  begin
    ACurPageLen := 0;
  end;

  procedure AddTopVectors(AVectorIndex: Integer);
  var
    I, J: Integer;
    APMTopVector: TvgrPMTopVector;
  begin
    for I := 0 to ATopVectors.Count - 1 do
    begin
      APMTopVector := ATopVectors.Items[I];
      with APMTopVector do
        if (StartArea <= AVectorIndex) and (EndArea >= AVectorIndex) and PrintOnTop then
        begin
          for J:= StartPos to EndPos do
          begin
            AddIdxItem(AVectIndexes, J);
            Inc(ACurPageLen, GetVLen(J));
            Inc(ACurDimVectIdx);
          end;
        end;
    end;
  end;

  function BottomVectorsSize(AVectorIndex: Integer): Integer;
  var
    I, J: Integer;
    APMTopVector: TvgrPMTopVector;
  begin
    Result := 0;
    for I := 0 to ATopVectors.Count - 1 do
    begin
      APMTopVector := ATopVectors.Items[I];
      with APMTopVector do
        if (StartArea <= AVectorIndex) and (EndArea >= AVectorIndex) and PrintOnBottom then
        begin
          for J:= StartPos to EndPos do
          begin
            Inc(Result, GetVLen(J));
          end;
        end;
    end;
  end;

  function SeekLinkedNextSection(Index: Integer): IvgrSection;
  var
    ASectionIndex: Integer;
  begin
    ASectionIndex := ASections.SearchDirectOuter(Index, Index, -1);
    if ASectionIndex >= 0 then
    begin
      Result := ASections.ByIndex[ASectionIndex];
      if not Result.PrintWithNextSection then
        Result := nil
      else
        Result := ASections.ByIndex[ASectionIndex];
    end
    else
      Result := nil;
  end;

  function SeekNextSectionLinkedWithPrev(Index: Integer): IvgrSection;
  var
    ASectionIndex: Integer;
  begin
    ASectionIndex := ASections.SearchDirectOuter(Index + 1, Index + 1, -1);
    if ASectionIndex >= 0 then
    begin
      Result := ASections.ByIndex[ASectionIndex];
      if not Result.PrintWithPreviosSection then
        Result := nil;
    end
    else
      Result := nil;
  end;

  function GetVLinkedLen(Index: Integer): Integer;
  var
    ASection: IvgrSection;
    I : Integer;
  begin
    Result := GetVLen(Index);

    ASection := SeekLinkedNextSection(Index);
    if ASection <> nil then
    begin
      for I := (ASection.StartPos + 1) to ASection.EndPos do
        Inc(Result, GetVLen(I));
      Inc(Result, GetVLen(ASection.EndPos + 1));
    end;

    ASection := SeekNextSectionLinkedWithPrev(Index);
    if ASection <> nil then
    begin
      for I := ASection.StartPos to ASection.EndPos do
        Inc(Result, GetVLen(I));
    end;
  end;

  function PageBreakAfter(Index: Integer): Boolean;
  var
    AVector: IvgrPageVector;
  begin
    AVector := AVectors.Find(I) as IvgrPageVector;
    Result := (AVector <> nil) and (AVector.PageBreak);
  end;

  procedure AddMiddleVectors(var AVectorIndex: Integer);
  var
    AFlagBreak: Boolean;
    AFlagFirst: Boolean;
  begin
    AFlagBreak := False;
    AFlagFirst := True;
    while (I < AEdge) and (AFlagFirst or (ACurPageLen <= (APageSize - GetVLinkedLen(I + 1) - BottomVectorsSize(I + 1)))) and not AFlagBreak do
    begin
      AFlagFirst := False;
      Inc(I);
      Inc(ACurDimVectIdx);
      if PageBreakAfter(AVectorIndex) then
        AFlagBreak := True;
      AddIdxItem(AVectIndexes, I);
      Inc(ACurPageLen, GetVLen(I));
      {$IFDEF VGR_DBG_MAKE}
      OutputDebugString(PChar(IntToStr(ACurPageLen)));
      {$ENDIF}
    end;
  end;

  procedure AddBottomVectors(var AVectorIndex: Integer);
  var
    I, J: Integer;
    APMTopVector: TvgrPMTopVector;
  begin
    for I := 0 to ATopVectors.Count - 1 do
    begin
      APMTopVector := ATopVectors.Items[I];
      with APMTopVector do
        if (StartArea <= AVectorIndex) and (EndArea >= AVectorIndex) and PrintOnBottom then
        begin
          for J:= StartPos to EndPos do
          begin
            AddIdxItem(AVectIndexes, J);
            Inc(ACurPageLen, GetVLen(J));
            Inc(ACurDimVectIdx);
          end;
        end;
    end;
  end;

  procedure EndPage;
  begin
    AddIdxItem(ARulers, ACurDimVectIdx);
  end;

  procedure Init;
  begin
    I := -1;
    ACurDimVectIdx := -1;
    ATopVectors := TvgrPMTopVectors.Create(ASections);
  end;

  procedure Final;
  begin
    ATopVectors.Free;
  end;

begin
  Init;
  repeat
    BeginNewPage;
    AddTopVectors(I);
    AddMiddleVectors(I);
    AddBottomVectors(I);
    EndPage;
  until I = AEdge;
  Final;
end;

end.
