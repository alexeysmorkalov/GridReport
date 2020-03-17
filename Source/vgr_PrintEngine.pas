{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains TvgrPrintEngine non visual component.
TvgrPrintEngine provide functionality for printing contents of the workbook.
To divide rows and cols of worksheet to pages TvgrPageMaker class is used.
See also:
  TvgrWorksheet, TvgrPageMaker, TvgrWorkbook.}
unit vgr_PrintEngine;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, SysUtils, Graphics, Windows, winspool, math,

  vgr_DataStorage, vgr_Printer, vgr_PageMaker, vgr_PageProperties,
  vgr_WorkbookPainter, vgr_CommonClasses;

type

{Specifies the part of the worksheet to print.
Items:
  vgrpmAll - all pages are printed.
  vgrpmCurrent - current page are printed.
  vgrpmRange - the range of the pages are printed.
  vgrpmList - the specified pages are printed.}
  TvgrPrintPageMode = (vgrpmAll, vgrpmCurrent, vgrpmRange, vgrpmList);

{Specifies the part of the workbook to print.
Items:
  vgrwmAll - all worksheets are printed.
  vgrwmDefined - the specified(one) worksheet are printed.}
  TvgrPrintWorksheetMode = (vgrwmAll, vgrwmDefined);

  /////////////////////////////////////////////////
  //
  // TvgrPrintProperties
  //
  /////////////////////////////////////////////////
{Specifies information about how a workbook is printed.
Class specifies various properties for printing: range of page to print, count of properties, etc.
Instance of this class created by TvgrPrintEngine to store printing settings.
See also:
  TvgrPrintEngine}
  TvgrPrintProperties = class(TvgrPersistent)
  private
    FPrintPageMode: TvgrPrintPageMode;
    FFromPage: Integer;
    FToPage: Integer;
    FPrintPages: string;
    FPrintWorksheetMode: TvgrPrintWorksheetMode;
    FPrintWorksheet: TvgrWorksheet;
    FCopies: Integer;
    FCollate: Boolean;
    FPrintPagesList: TList;
    procedure SetPrintPages(const Value: string);
    procedure SetCopies(Value: Integer);
    function GetPrintPageCount: Integer;
    function GetPrintPage(Index: Integer): Integer;
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation);
  public
{Creates the instance of a TvgrPrintProperties
See also:
  TvgrPrintProperties}
    constructor Create; override;
{Frees instance of a TvgrPrintProperties}
    destructor Destroy; override;

    procedure Assign(Source: TPersistent); override;

{Returns count of pages, which are listed in PrintPages property, for example,
if PrintPages contains "1, 2, 5-7", PrintPageCount returns 5.
Syntax:
  property PrintPageCount: Integer read;
See also:
  PrintPages, PrintPage}
    property PrintPageCount: Integer read GetPrintPageCount;
{Use this property to interate numbers of pages are listed in PrintPages property,
for example, if PrintPages contains "1, 2, 5-7", PrintPage returns 1, 2, 5, 6, 7.
Syntax:
  property PrintPage[Index: Integer]: Integer read;
See also:
  PrintPages, PrintPageCount}
    property PrintPage[Index: Integer]: Integer read GetPrintPage;
  published
{Gets or sets the page numbers that the user has specified to be printed.
If PrintPageMode = vgrpmRange then pages from FromPage to ToPage are printed.
If PrintPageMode = vgrpmList then pages are listed in PrintPages property are printed.
Syntax:
  property PrintPageMode: TvgrPrintPageMode read write default vgrpmAll;
See also:
  TvgrPrintPageMode, FromPage, ToPage, PrintPages}
    property PrintPageMode: TvgrPrintPageMode read FPrintPageMode write FPrintPageMode default vgrpmAll;
{Gets or sets the page number of the first page to print.
Syntax:
  property FromPage: Integer read write default 1;}
    property FromPage: Integer read FFromPage write FFromPage default 1;
{Gets or sets the number of the last page to print.
Syntax:
  property ToPage: Integer read write default 1;}
    property ToPage: Integer read FToPage write FToPage default 1;
{Gets or sets the numbers of pages to print, for example:
"1, 2, 5-7", or "10-20"
Syntax:
  property PrintPages: string read write;
See also:
  PrintPage, PrintPageCount}
    property PrintPages: string read FPrintPages write SetPrintPages;
{Gets or sets worksheets of the workbook to print.
If PrintWorksheetMode = vgrwmDefined then worksheet defined in PrintWorksheet property are printed.
Syntax:
  property PrintWorksheetMode: TvgrPrintWorksheetMode read write default vgrwmAll;
See also:
  TvgrPrintWorksheetMode, PrintWorksheet}
    property PrintWorksheetMode: TvgrPrintWorksheetMode read FPrintWorksheetMode write FPrintWorksheetMode default vgrwmAll;
{Gets or sets worksheet to printing, this property is used if PrintWorksheetMode = vgrwmDefined.
Syntax:
  property PrintWorksheet: TvgrWorksheet read FPrintWorksheet write FPrintWorksheet;
See also:
  PrintWorksheetMode}
    property PrintWorksheet: TvgrWorksheet read FPrintWorksheet write FPrintWorksheet;
{Gets or sets the number of copies of the pages to print.
Syntax:
  property Copies: Integer read write default 1;}
    property Copies: Integer read FCopies write SetCopies default 1;
{Gets or sets a value indicating whether the printed document is collated.
Syntax:
  property Collate: Boolean read write default True}
    property Collate: Boolean read FCollate write FCollate default True;
  end;

{Used for firing information about start of printing.
Syntax:
  TvgrDocumentStartEvent = procedure (Sender: TObject; var ADocumentTitle: string) of object;
Parameters:
  Sender - instance of TvgrPrintEngine.
  ADocumentTitle - name of the started printing job, you can change it.
See also:
  TvgrPrintEngine}
  TvgrDocumentStartEvent = procedure (Sender: TObject; var ADocumentTitle: string) of object;
{Used to firing information about drawing of the page on canvas of the printer,
you can paint something on the provided canvas.
Parameters:
  Sender - instance of TvgrPrintEngine.
  ACanvas - canvas of the printer.
  APage - the printed page.
See also:
  TvgrPrintEngine, TvgrPrintPage}
  TvgrDrawPageEvent = procedure (Sender: TObject; ACanvas: TCanvas; APage: TvgrPrintPage) of object;
{Used to firing information about end of printing of the page.
Parameters:
  Sender - instance of TvgrPrintEngine.
  APage - the printed page.
See also:
  TvgrPrintEngine, TvgrPrintPage}
  TvgrPagePrintedEvent = procedure (Sender: TObject; APage: TvgrPrintPage) of object;
  /////////////////////////////////////////////////
  //
  // TvgrPrintEngine
  //
  /////////////////////////////////////////////////
{Provide functionality for printing contents of the workbook.
Main methods of this class:
  Print - print workbook or part of the workbook, specified by the Workbook property.<br>
  PrintWorksheet -  print the one worksheet of the workbook, specified by the Workbook property.<br>
  PrintPage - print one page of the worksheet.<br>
To divide rows and cols of worksheet to pages TvgrPageMaker class is used.
To get information about installed printers use Printers property.
See also:
  TvgrPageMaker}
  TvgrPrintEngine = class(TComponent, IvgrWorkbookPainterOwner)
  private
    FWorkbook: TvgrWorkbook;
    FPrintProperties: TvgrPrintProperties;
    FAutoSelectDefaultPrinter: Boolean;
    FPrinter: TvgrPrinter;
    FPainter: TvgrWorkbookPainter;
    FDocumentTitle: string;

    FLastPageProperties: TvgrPageProperties;
    FPrinterDC: HDC;
    FCanvas: TCanvas;

    FOnDocumentStart: TvgrDocumentStartEvent;
    FOnBeforeDrawPage: TvgrDrawPageEvent;
    FOnAfterDrawPage: TvgrDrawPageEvent;
    FOnPagePrinted: TvgrPagePrintedEvent;

    procedure SetPrintProperties(Value: TvgrPrintProperties);
    procedure SetAutoSelectDefaultPrinter(Value: Boolean);
    procedure SetPrinterName(const Value: string);
    function GetPrinterName: string;
    function GetPrintActive: Boolean;
  protected
    // IvgrWorkbookPainterOwner
    function GetDefaultColumnWidth: Integer;
    function GetDefaultRowHeight: Integer;
    function GetPixelsPerInchX: Integer;
    function GetPixelsPerInchY: Integer;
    function GetBackgroundColor: TColor;
    function GetDefaultRangeBackgroundColor: TColor;

    function IsEqualPageSizes(APageProperties1, APageProperties2: TvgrPageProperties): Boolean;

    procedure DoDocumentStart(var ADocumentTitle: string);
    procedure DoBeforeDrawPage(ACanvas: TCanvas; APrintPage: TvgrPrintPage);
    procedure DoAfterDrawPage(ACanvas: TCanvas; APrintPage: TvgrPrintPage);
    procedure DoPagePrinted(APrintPage: TvgrPrintPage);

    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
    procedure Loaded; override;

    procedure PrepareWorkbook;

    procedure CheckPrinter;

    property LastPageProperties: TvgrPageProperties read FLastPageProperties;
  public
{Creates the instance of a TvgrPrintEngine
Syntax:
  constructor Create(AOwner: TComponent); override;
Parameters:
  AOwner - owner component. 
See also:
  TvgrPrintProperties}
    constructor Create(AOwner: TComponent); override;
{Frees instance of a TvgrPrintEngine
Syntax:
  destructor Destroy; override;}
    destructor Destroy; override;

{Prints contents of the workbook, use PrintProperties property to specify which part of the workbook
should be printed
This method fires OnBeforeDrawPage, OnAfterDrawPage, OnPagePrinted and OnDocumentStart events.
Example:
  var
    APE: TvgrPrintEngine;
  begin
    APE := TvgrPrintEngine.Create(nil);
    APE.Workbook := vgrWorkbook;
    APE.Print;
    APE.Free;
  end;
See also:
  PrintPage, PrintWorksheet.}
    procedure Print;
{Prints one page of the worksheet.
This method are not using Workbook, PrintProperties properties,
fires OnBeforeDrawPage, OnAfterDrawPage and OnDocumentStart events.
Parameters:
  APrintPage - page to printing, instance of this class creates by the TvgrPageMaker.
  AWorksheet - worksheet page which are printed.
  APageProperties - page properties, used to prepare print job.
Example:
  var
    APageMaker: TvgrPageMaker;
    APrintEngine: TvgrPrintEngine;
    AWorksheet: TvgrWorksheet;
  begin
    APrintEngine := TvgrPrintEngine.Create(nil);
    AWorksheet := wb.Worksheets[0];
    APageMaker := TvgrPageMaker.Create(AWorksheet);
    try
      APrintEngine.Workbook := wb;
      APageMaker.Prepare;

      // print first page
      if APageMaker.PagesCount > 0 then
        APrintEngine.PrintPage(APageMaker.Pages[0], AWorksheet, AWorksheet.PageProperties);

      // print last page
      if APageMaker.PagesCount > 1 then
        APrintEngine.PrintPage(APageMaker.Pages[(APageMaker.PagesCount - 1], AWorksheet, AWorksheet.PageProperties);

      if APageMaker.PagesCount > 0 then
        APrintEngine.EndDocument;

    finally
      APageMaker.Free;
      APrintEngine.Free;
    end;
  end;
See also:
  TvgrPageProperties, TvgrPageMaker, TvgrWorksheet.}
    procedure PrintPage(APrintPage: TvgrPrintPage; AWorksheet: TvgrWorksheet; APageProperties: TvgrPageProperties);
{Prints contents of the one worksheet, use PrintProperties property to specify which part of the worksheet
should be printed
This method fires OnBeforeDrawPage, OnAfterDrawPage, OnPagePrinted and OnDocumentStart events.
Example:
  var
    APE: TvgrPrintEngine;
  begin
    APE := TvgrPrintEngine.Create(nil);
    APE.PrintProperties.PrintPageMode := vgrpmRange;
    APE.PrintProperties.FromPage := 1;
    APE.PrintProperties.ToPage := 1; // only first page

    // first worksheet
    if wb.WorksheetsCount > 0 then
      APE.PrintWorksheet(wb.Worksheets[0]);

    // last worksheet
    if wb.WorksheetsCount > 1 then
      APE.PrintWorksheet(wb.Worksheets[wb.WorksheetsCount - 1]);

    if wb.WorksheetsCount > 0 then
      APrintEngine.EndDocument;
      
    APE.Free;
  end;
See also:
  PrintPage, Print.}
    procedure PrintWorksheet(AWorksheet: TvgrWorksheet);
{You must call this method to close printing job after using methods PrintPage and PrintWorksheet.
See also:
  PrintPage, PrintWorksheet}
    procedure EndDocument;

{Use this method to setup PrinterName property on the default printer.
See also:
  Printers, PrinterName}
    procedure SetToDefaultPrinter;

{Gets or sets name of the printer that used for printing.
Syntax:
  property PrinterName: string read write;
See also:
  Printers,  UpdatePrinters, SetToDefaultPrinter}
    property PrinterName: string read GetPrinterName write SetPrinterName;
{Gets TvgrPrinter object from Printers list according to PrinterName property.
Syntax:
  property Printer: TvgrPrinter read;}
    property Printer: TvgrPrinter read FPrinter;
{Returns true if TvgrPrintEngine currently prints.
Syntax:
  property PrintActive: Boolean read;}
    property PrintActive: Boolean read GetPrintActive;
  published
{Gets or sets a value indicating whether after creation of the component
PrinterName property setup to default printer.
Syntax:
  property AutoSelectDefaultPrinter: Boolean read write default True;
See also:
  Printers, PrinterName}
    property AutoSelectDefaultPrinter: Boolean read FAutoSelectDefaultPrinter write SetAutoSelectDefaultPrinter default True;
{Gets or sets properties of printing (Copies, Collate, etc).
Syntax:
  property PrintProperties: TvgrPrintProperties read write;
See also:
  TvgrPrintProperties}
    property PrintProperties: TvgrPrintProperties read FPrintProperties write SetPrintProperties;
{Gets or sets workbook, used for printing in the Print method.
Syntax:
  property Workbook: TvgrWorkbook read write;
See also:
  Print}
    property Workbook: TvgrWorkbook read FWorkbook write FWorkbook;
{Gets or sets name of the printing job, you can change it in OnDocumentStart event.
Syntax:
  property DocumentTitle: string read write;
See also:
  OnDocumentStart, TvgrDocumentStartEvent}
    property DocumentTitle: string read FDocumentTitle write FDocumentTitle;

{This event fired when printing job started.
Syntax:
  property OnDocumentStart: TvgrDocumentStartEvent read write;
See also:
  TvgrDocumentStartEvent}
    property OnDocumentStart: TvgrDocumentStartEvent read FOnDocumentStart write FOnDocumentStart;
{This event fired when painting of the page are started.
Syntax:
  property OnBeforeDrawPage: TvgrDrawPageEvent read write;
See also:
  TvgrDrawPageEvent, OnAfterDrawPage}
    property OnBeforeDrawPage: TvgrDrawPageEvent read FOnBeforeDrawPage write FOnBeforeDrawPage;
{This event fired when painting of the page are ended.
Syntax:
  property OnAfterDrawPage: TvgrDrawPageEvent read write;
See also:
  TvgrDrawPageEvent, OnBeforeDrawPage}
    property OnAfterDrawPage: TvgrDrawPageEvent read FOnAfterDrawPage write FOnAfterDrawPage;
{This event fired when printing of the page are ended.
Syntax:
  property OnPagePrinted: TvgrPagePrintedEvent read write;
See also:
  TvgrPagePrintedEvent}
    property OnPagePrinted: TvgrPagePrintedEvent read FOnPagePrinted write FOnPagePrinted;
  end;

implementation

uses
  vgr_Functions;

/////////////////////////////////////////////////
//
// TvgrPrintProperties
//
/////////////////////////////////////////////////
constructor TvgrPrintProperties.Create;
begin
  inherited Create;
  FPrintPagesList := TList.Create;
  FCopies := 1;
  FCollate := True;
  FFromPage := 1;
  FToPage := 1;
end;

destructor TvgrPrintProperties.Destroy;
begin
  FPrintPagesList.Free;
  inherited;
end;

procedure TvgrPrintProperties.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  if (AOperation = opRemove) and (AComponent = FPrintWorksheet) then
    FPrintWorksheet := nil;
end;

procedure TvgrPrintProperties.Assign(Source: TPersistent);
begin
  if Source is TvgrPrintProperties then
    with TvgrPrintProperties(Source) do
    begin
      Self.FPrintPageMode := PrintPageMode;
      Self.FFromPage := FromPage;
      Self.FToPage := ToPage;
      Self.PrintPages := PrintPages;
      Self.FPrintWorksheetMode := PrintWorksheetMode;
      Self.FPrintWorksheet := PrintWorksheet;
      Self.FCopies := Copies;
      Self.FCollate := Collate;
    end;
end;

procedure TvgrPrintProperties.SetPrintPages(const Value: string);
begin
  if FPrintPages <> Value then
  begin
    if CheckPageList(Value) then
    begin
      FPrintPages := Value;
      TextToPageList(FPrintPages, FPrintPagesList)
    end;
  end;
end;

procedure TvgrPrintProperties.SetCopies(Value: Integer);
begin
  if Value <= 0 then
    Value := 1;

  if FCopies <> Value then
    FCopies := Value;
end;

function TvgrPrintProperties.GetPrintPageCount: Integer;
begin
  Result := FPrintPagesList.Count;
end;

function TvgrPrintProperties.GetPrintPage(Index: Integer): Integer;
begin
  Result := Integer(FPrintPagesList[Index]);
end;

/////////////////////////////////////////////////
//
// TvgrPrintEngine
//
/////////////////////////////////////////////////
constructor TvgrPrintEngine.Create(AOwner: TComponent);
begin
  inherited;
  FAutoSelectDefaultPrinter := True;
  FPrintProperties := TvgrPrintProperties.Create;
  FCanvas := nil;
  FPrinterDC := 0;
  FPainter := TvgrWorkbookPainter.Create(Self);
end;

destructor TvgrPrintEngine.Destroy;
begin
  EndDocument;
  FPrintProperties.Free;
  LastPageProperties.Free;
  FPainter.Free;
  inherited;
end;

procedure TvgrPrintEngine.Loaded;
begin
  inherited;
  if AutoSelectDefaultPrinter then
    SetToDefaultPrinter;
end;

function TvgrPrintEngine.GetDefaultColumnWidth: Integer;
begin
  Result := LastPageProperties.Defaults.ColWidth;
end;

function TvgrPrintEngine.GetDefaultRowHeight: Integer;
begin
  Result := LastPageProperties.Defaults.RowHeight;
end;

function TvgrPrintEngine.GetPixelsPerInchX: Integer;
begin
  Result := Printer.PixelsPerX;
end;

function TvgrPrintEngine.GetPixelsPerInchY: Integer;
begin
  Result := Printer.PixelsPerY;
end;

function TvgrPrintEngine.GetBackgroundColor: TColor;
begin
  Result := clNone;
end;

function TvgrPrintEngine.GetDefaultRangeBackgroundColor: TColor;
begin
  Result := clWhite;
end;

procedure TvgrPrintEngine.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if (AOperation = opRemove) and (AComponent = FWorkbook) then
    FWorkbook := nil;
  FPrintProperties.Notification(AComponent, AOperation);
end;

procedure TvgrPrintEngine.SetPrintProperties(Value: TvgrPrintProperties);
begin
  FPrintProperties.Assign(Value);
end;

procedure TvgrPrintEngine.SetAutoSelectDefaultPrinter(Value: Boolean);
begin
  if FAutoSelectDefaultPrinter <> Value then
  begin
    FAutoSelectDefaultPrinter := Value;
    if FAutoSelectDefaultPrinter then
      SetToDefaultPrinter;
  end;
end;

function TvgrPrintEngine.GetPrinterName: string;
begin
  if Printer = nil then
    Result := ''
  else
    Result := Printer.PrinterName;
end;

function TvgrPrintEngine.GetPrintActive: Boolean;
begin
  Result := FPrinterDC <> 0;
end;

procedure TvgrPrintEngine.SetPrinterName(const Value: string);
var
  I: Integer;
begin
  I := Printers.IndexOfPrinterName(Value);
  if I <> -1 then
  begin
    FPrinter := Printers[I];
    FreeAndNil(FLastPageProperties);
  end;
end;

procedure TvgrPrintEngine.SetToDefaultPrinter;
begin
  if Printer <> Printers.DefaultPrinter then
  begin
    FPrinter := Printers.DefaultPrinter;
    FreeAndNil(FLastPageProperties);
  end;
end;

procedure TvgrPrintEngine.PrepareWorkbook;
begin
end;

procedure TvgrPrintEngine.CheckPrinter;
begin
  if (Printer = nil) and AutoSelectDefaultPrinter then
    SetToDefaultPrinter;
  if Printer = nil then
    raise Exception.Create('Printer not selected');
end;

procedure TvgrPrintEngine.Print;
var
  I: Integer;
begin
  CheckPrinter;
  case PrintProperties.PrintWorksheetMode of
    vgrwmAll:
      begin
        for I := 0 to Workbook.WorksheetsCount - 1 do
          PrintWorksheet(Workbook.Worksheets[I]);
        EndDocument;
      end;
    vgrwmDefined:
      if PrintProperties.PrintWorksheet <> nil then
      begin
        PrintWorksheet(PrintProperties.PrintWorksheet);
        EndDocument;
      end;
  end;
end;

procedure TvgrPrintEngine.DoDocumentStart(var ADocumentTitle: string);
begin
  if Assigned(OnDocumentStart) then
    OnDocumentStart(Self, ADocumentTitle);
end;

procedure TvgrPrintEngine.DoBeforeDrawPage(ACanvas: TCanvas; APrintPage: TvgrPrintPage);
begin
  if Assigned(OnBeforeDrawPage) then
    OnBeforeDrawPage(Self, ACanvas, APrintPage);
end;

procedure TvgrPrintEngine.DoAfterDrawPage(ACanvas: TCanvas; APrintPage: TvgrPrintPage);
begin
  if Assigned(OnAfterDrawPage) then
    OnAfterDrawPage(Self, ACanvas, APrintPage);
end;

procedure TvgrPrintEngine.DoPagePrinted(APrintPage: TvgrPrintPage);
begin
  if Assigned(OnPagePrinted) then
    OnPagePrinted(Self, APrintPage);
end;

function TvgrPrintEngine.IsEqualPageSizes(APageProperties1, APageProperties2: TvgrPageProperties): Boolean;
begin
  Result := (APageProperties1.UnitsWidth[vgruMms] = APageProperties2.UnitsWidth[vgruMms]) and
            (APageProperties1.UnitsHeight[vgruMms] = APageProperties2.UnitsHeight[vgruMms]);
end;

procedure TvgrPrintEngine.PrintPage(APrintPage: TvgrPrintPage; AWorksheet: TvgrWorksheet; APageProperties: TvgrPageProperties);
{$IFNDEF VGR_DEMO}
var
  ADevMode: PDevMode;
  ADocumentTitle: string;
  ADocInfo: TDocInfo;
  ALeftOffs: Integer;
  ATopOffs: Integer;
  APaperSize, APaperIndex: Integer;
  APaperName: string;
  APaperOrientation: TvgrPageOrientation;
  APaperDimensions: TPoint;
{$ENDIF}
begin
{$IFNDEF VGR_DEMO}
  CheckPrinter;
  if not PrintActive or (LastPageProperties = nil) or not IsEqualPageSizes(LastPageProperties, APageProperties) then
  begin
    // end current document
    EndDocument;

    ADevMode := Printer.AllocNewDevMode;
    if ADevMode = nil then
      exit;

    try
      // save new PageProperties
      if LastPageProperties = nil then
        FLastPageProperties := TvgrPageProperties.Create;
      FLastPageProperties.Assign(APageProperties);

      Printer.FindPaper(APageProperties, APaperIndex, APaperSize, APaperName, APaperOrientation, APaperDimensions);
      if APaperIndex = -1 then
      begin
        ADevMode.dmFields := (ADevMode.dmFields or DM_PAPERLENGTH or DM_PAPERWIDTH) and not DM_PAPERSIZE;
        ADevMode.dmPaperLength := APaperDimensions.Y;
        ADevMode.dmPaperWidth := APaperDimensions.X;
      end
      else
      begin
        ADevMode.dmFields := (ADevMode.dmFields or DM_PAPERSIZE) and not DM_PAPERLENGTH and not DM_PAPERWIDTH;
        ADevMode.dmPaperSize := APaperSize;
      end;
      ADevMode.dmFields := ADevMode.dmFields or DM_ORIENTATION;
      if APaperOrientation = vgrpoPortrait then
        ADevMode.dmOrientation := DMORIENT_PORTRAIT
      else
        ADevMode.dmOrientation := DMORIENT_LANDSCAPE;

      // create printer DC
      with Printer do
        FPrinterDC := CreateDC(PChar(DriverName), PChar(DeviceName), PChar(PortName), ADevMode);

      if FPrinterDC = 0 then
        exit;

      // create a Canvas
      if FCanvas <> nil then
        FCanvas.Free;
      FCanvas := TCanvas.Create;
      FCanvas.Handle := FPrinterDC;

      // start document
      ADocumentTitle := DocumentTitle;
      DoDocumentStart(ADocumentTitle);

      FillChar(ADocInfo, SizeOf(TDocInfo), #0);
      ADocInfo.cbSize := SizeOf(TDocInfo);
      ADocInfo.lpszDocName := PChar(ADocumentTitle);
      ADocInfo.lpszOutput := nil;

      StartDoc(FPrinterDC, ADocInfo);
      StartPage(FPrinterDC);
    finally
      FreeMem(ADevMode);
    end;
  end
  else
  begin
    StartPage(FPrinterDC);
    FCanvas.Refresh;
  end;

  // draw page
  DoBeforeDrawPage(FCanvas, APrintPage);

  ALeftOffs := ConvertTwipsToPixelsX(APageProperties.Margins.Left - Printer.PageMargins.Left, Printer.PixelsPerX);
  ATopOffs := ConvertTwipsToPixelsY(APageProperties.Margins.Top - Printer.PageMargins.Top, Printer.PixelsPerY);

  FPainter.Worksheet := AWorksheet;
  FPainter.DrawPage(ALeftOffs,
                    ATopOffs,
                    FCanvas,
                    APrintPage, 1, Point(ALeftOffs, ATopOffs));

  DoAfterDrawPage(FCanvas, APrintPage);

  EndPage(FPrinterDC);
{$ENDIF}
end;

procedure TvgrPrintEngine.PrintWorksheet(AWorksheet: TvgrWorksheet);
var
  I, J, AFrom, ATo: Integer;
  APageMaker: TvgrPageMaker;
  APrintPages: array of Integer;
begin
  APageMaker := TvgrPageMaker.Create(AWorksheet);
  try
    APageMaker.Prepare;

    // build array of printed pages
    case PrintProperties.PrintPageMode of
      vgrpmAll:
        begin
          SetLength(APrintPages, APageMaker.PagesCount);
          for I := 0 to APageMaker.PagesCount - 1 do
            APrintPages[I] := I;
        end;
      vgrpmRange:
        begin
          AFrom := Max(PrintProperties.FromPage - 1, 0);
          ATo := Min(PrintProperties.ToPage, APageMaker.PagesCount) - 1;
          SetLength(APrintPages, ATo - AFrom + 1);
          for I := AFrom to ATo do
            APrintPages[I - AFrom] := I;
        end;
      vgrpmList:
        begin
          SetLength(APrintPages, PrintProperties.PrintPageCount);
          for I := 0 to PrintProperties.PrintPageCount - 1 do
            APrintPages[I] := PrintProperties.PrintPage[I] - 1;
        end;
    end;

    // print pages
    if PrintProperties.Collate then
    begin
      for I := 1 to PrintProperties.Copies do
        for J := 0 to High(APrintPages) do
        begin
          PrintPage(APageMaker.Pages[APrintPages[J]], AWorksheet, AWorksheet.PageProperties);
          DoPagePrinted(APageMaker.Pages[APrintPages[J]]);
        end;
    end
    else
    begin
      for J := 0 to High(APrintPages) do
        for I := 1 to PrintProperties.Copies do
        begin
          PrintPage(APageMaker.Pages[APrintPages[J]], AWorksheet, AWorksheet.PageProperties);
          DoPagePrinted(APageMaker.Pages[APrintPages[J]]);
        end;
    end;

  finally
    APageMaker.Free;
  end;
end;

procedure TvgrPrintEngine.EndDocument;
begin
  if FCanvas <> nil then
  begin
    FCanvas.Free;
    FCanvas := nil;
  end;
  if FPrinterDC <> 0 then
  begin
    EndDoc(FPrinterDC);
    DeleteDC(FPrinterDC);
    FPrinterDC := 0;
  end;
end;

end.

