{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains the auxiliary classes - TvgrPrinter and TvgrPrinters.
TvgrPrinters is a list which contains printers installed in the system,
use Update method to update list of installed printers.
Each printer is described by the TvgrPrinter object, that contains
properties: PrinterName, DriverName and so on.
See also:
  TvgrPrinter, TvgrPrinters}
unit vgr_Printer;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  SysUtils, Classes, Windows, WinSpool, CommDlg, Graphics,

  vgr_Functions, vgr_PageProperties, vgr_CommonClasses,
  vgr_Consts;

type
  TvgrPrinters = class;

  /////////////////////////////////////////////////
  //
  // TvgrPaperInfo
  //
  /////////////////////////////////////////////////
{Internal struct, describes type of the paper - code, name and sizes.
Syntax:
  TvgrPaperInfo = record
    Typ : Integer;
    Name : string;
    X,Y : Integer;
  end;}
  TvgrPaperInfo = record
    Typ : Integer;
    Name : string;
    X,Y : Integer;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPrinter
  //
  /////////////////////////////////////////////////
{Describes one of printers, installed on the system.
Instance of this class is created by the TvgrPrinters to store information
about individually printer.
See also:
  TvgrPrinters}
  TvgrPrinter = class(TObject)
  private
    FPrinters: TvgrPrinters;
    FPrinterName: string;
    FDriverName: string;
    FPortName: string;
    FDeviceName: string;
    FComment: string;
    FDevMode: PDevMode;
    FDevModeSize: Integer;
    FPaperSizes: PArrayWord;
    FPaperDimensions: PArrayTPoint;
    FPageMargins: TvgrPageMargins;
    FPixelsPerX: Integer;
    FPixelsPerY: Integer;
    FPaperNames: TStrings;
    FLocal: Boolean;
    FInitializated: Boolean;
    function GetPaperCount: Integer;
    function GetPaperName(Index: Integer): string;
    function GetPaperSize(Index: Integer): Integer;
    function GetPaperDimension(Index: Integer): TPoint;
    procedure SetPrinterName(const Value: string);
    function GetDefault: Boolean;
    function GetImage: TBitmap;
    function GetPaperNames: TStrings;
    function GetDriverName: string;
    function GetPortName: string;
    function GetDeviceName: string;
    function GetPixelsPerX: Integer;
    function GetPixelsPerY: Integer;
    function GetComment: string;
    function GetPageMargins: TvgrPageMargins;
    function GetStatus: string;
  protected
    procedure Initialize;
    procedure UpdateLocal;
    procedure ClearStructures;

    constructor Create(APrinters: TvgrPrinters; const APrinterName: string); overload;

    property Initializated: Boolean read FInitializated;
  public
{Creates the instance of a TvgrPrinter.
Parameters:
  APrinters - TvgrPrinters object, that creates this object, this parameter can be nil.
See also:
  TvgrPrinter}
    constructor Create(APrinters: TvgrPrinters); overload;
{Frees the TvgrPrinter instance.}
    destructor Destroy; override;

{Finds standart type of the paper, supported by the printer.
List on the paper types contains in PaperNames, PaperSizes and PaperDimensions properties,
use PaperCount property to get count of supported paper types.
Parameters:
  APageProperties - dimensions of the paper.
  APaperIndex - returns index of the paper in PaperNames, PaperSizes and PaperDimensions properties
or -1 if paper are not found.
  APaperSize - returns code of the paper (see constants DMPAPER_XX, DMPAPER_A4 for example)
or -1 if paper are not found.
  APaperName - name of the paper or empty string if paper are not found.
  APaperOrientation - orientation of the paper, that should be used.
  APaperDimensions - sizes of the paper in 1/10 of millimeter.
See also:
  PaperCount, PaperNames, PaperSizes, PaperDimensions}
    procedure FindPaper(APageProperties: TvgrPageProperties;
                        var APaperIndex: Integer;
                        var APaperSize: Integer;
                        var APaperName: string;
                        var APaperOrientation: TvgrPageOrientation;
                        var APaperDimensions: TPoint);

{The function presents the printer driver's Print Setup property sheet and then
changes the settings in the printer's DEVMODE data structure to those values
specified by the user.
Return value:
  Returns true if user press OK in the dialog window.}
    function ShowPropertiesDialog(AParentWindowHandle: THandle): Boolean;

{Create a copy of DEVMODE structure for printer.
Return value:
  Returns the pointer to the DEVMODE structore.}
    function AllocNewDevMode: PDEVMODE;

{Gets names of the paper types, supported by the printer}
    property PaperNamesList: TStrings read GetPaperNames;
{Gets and sets name of the printer, that are described by this object.}
    property PrinterName: string read FPrinterName write SetPrinterName;
{Gets the name of the printer driver.}
    property DriverName: string read GetDriverName;
{Gets the port(s) used to transmit data to the printer.
If a printer is connected to more than one port, the names of each port is separated by commas.
(for example, "LPT1:,LPT2:,LPT3:")}
    property PortName: string read GetPortName;
    property DeviceName: string read GetDeviceName;
{Gets a brief description of the printer.}
    property Comment: string read GetComment;
{Gets the string that describes the printer status.}
    property Status: string read GetStatus;
{Gets the boolean value, that indicates used this printer as default or not.}
    property Default: Boolean read GetDefault;
{Gets icon of the printer, two icons are exists, one for local printer, second for network printer.
See also:
  Local}
    property Image: TBitmap read GetImage;
{Gets the boolean value, that indicates is this printer local or not.
See also:
  Image}
    property Local: Boolean read FLocal write FLocal;

{Gets count of the paper types, supported by the printer.
See also:
  PaperNames, PaperSizes, PaperDimensions}
    property PaperCount: Integer read GetPaperCount;
{Gets the names of the paper types, supported by the printer.
The names are enumerated from 0, use PaperCount property to get count of the paper types. 
See also:
  PaperCount, PaperSizes, PaperDimensions}
    property PaperNames[Index: Integer]: string read GetPaperName;
{Gets the codes of the paper types, supported by the printer.
The codes are enumerated from 0, use PaperCount property to get count of the paper types. 
See also:
  PaperCount, PaperNames, PaperDimensions}
    property PaperSizes[Index: Integer]: Integer read GetPaperSize;
{Gets the dimensions of the paper types, supported by the printer.
The dimensions are enumerated from 0, use PaperCount property to get count of the paper types. 
See also:
  PaperCount, PaperNames, PaperSizes}
    property PaperDimensions[Index: Integer]: TPoint read GetPaperDimension;
{Gets the TvgrPageMargins object, that specifies the distance from the edges of the
physical page to the edges of the printable area.}
    property PageMargins: TvgrPageMargins read GetPageMargins;
{Number of pixels per logical inch(dpi) along the width.}
    property PixelsPerX: Integer read GetPixelsPerX;
{Number of pixels per logical inch(dpi) along the height.}
    property PixelsPerY: Integer read GetPixelsPerY;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPrinters
  //
  /////////////////////////////////////////////////
{Contains list of the printers installed in the system.
Each printer are represented by the TvgrPrinter object.
Use Update method to update information about installed printers.
See also:
  TvgrPrinter}
  TvgrPrinters = class(TObject)
  private
    FList: TList;
    FDefaultPrinterName: string;
    function GetItem(Index: Integer): TvgrPrinter;
    function GetCount: Integer;
    function GetDefaultPrinter: TvgrPrinter;
  protected
    procedure ClearPrintersList;
    procedure UpdateDefaultPrinterName;
    procedure UpdatePrintersList;
  public
{Creates the instance of a TvgrPrinters.
See also:
  TvgrPrinters}
    constructor Create;
{Frees the TvgrPrinters instance.}
    destructor Destroy; override;

{Updates list of installed printers.}
    procedure Update;
{Returns index of the printer in the list by its name.
If printer are not found -1 are returned.
Parameters:
  APrinterName - name of the printer.}
    function IndexOfPrinterName(const APrinterName: string): Integer;

{Returns sizes of the paper supported by the specified printer.
Parameters:
  APaperSize - code of the papert (one of DMPAPER_XX constants).
  APrinter - printer, that provides list of the paper types. This parameter can be nil,
in this case the paper are searched in the PaperInfo global variable.
Return value:
  Sizes of the paper.}
    function GetPaperDimensionsBySize(APaperSize: Integer; APrinter: TvgrPrinter): TPoint;
{Finds standart type of the paper, supported by the printer.
Parameters:
  APageProperties - dimensions of the paper.
  APrinter - printer, that provides list of the paper types. This parameter can be nil,
in this case the paper are searched in the PaperInfo global variable.
  APaperIndex - returns index of the paper in PaperNames, PaperSizes and PaperDimensions properties
or -1 if paper are not found.
  APaperSize - returns code of the paper (see constants DMPAPER_XX, DMPAPER_A4 for example)
or -1 if paper are not found.
  APaperName - name of the paper or empty string if paper are not found.
  APaperOrientation - orientation of the paper, that should be used.
  APaperDimensions - sizes of the paper in 1/10 of millimeter.
See also:
  TvgrPrinter, TvgrPrinter.FindPaper}
    procedure FindPaper(APageProperties: TvgrPageProperties;
                        APrinter: TvgrPrinter;
                        var APaperIndex: Integer;
                        var APaperSize: Integer;
                        var APaperName: string;
                        var APaperOrientation: TvgrPageOrientation;
                        var APaperDimensions: TPoint);

{Gets count of the printers in the list.}
    property Count: Integer read GetCount;
{Use Fields to obtain a pointer to a specific TvgrPrinter object.
The Index parameter indicates the index of the printer,
where 0 is the index of the first field, 1 is the index of the second field, and so on.
Parameters:
  Index - index of the printer in the list.}
    property Items[Index: Integer]: TvgrPrinter read GetItem; default;
{Gets the default printer.}
    property DefaultPrinter: TvgrPrinter read GetDefaultPrinter;
  end;

{Returns global TvgrPrinters object.
Return value:
  Returns TvgrPrinters object that represents list of installed printers.}
  function Printers: TvgrPrinters;

const
{Count of the standart paper types listed in the PaperInfo variable.
See also:
  PaperInfo}
  DEF_PAPERCOUNT = 66;
{Contains the standart paper types. Each type of the paper are described by the
TvgrPaperInfo record.
See also:
  TvgrPaperInfo, DEF_PAPERCOUNT}
  PaperInfo: Array[0..DEF_PAPERCOUNT - 1] of TvgrPaperInfo = (
    (Typ: 1;  Name: 'Letter, 8 1/2 x 11"'; X: 2159; Y: 2794),
    (Typ: 2;  Name: 'Letter small, 8 1/2 x 11"'; X: 2159; Y: 2794),
    (Typ: 3;  Name: 'Tabloid, 11 x 17"'; X: 2794; Y: 4318),
    (Typ: 4;  Name: 'Ledger, 17 x 11"'; X: 4318; Y: 2794),
    (Typ: 5;  Name: 'Legal, 8 1/2 x 14"'; X: 2159; Y: 3556),
    (Typ: 6;  Name: 'Statement, 5 1/2 x 8 1/2"'; X: 1397; Y: 2159),
    (Typ: 7;  Name: 'Executive, 7 1/4 x 10 1/2"'; X: 1842; Y: 2667),
    (Typ: 8;  Name: 'A3 297 x 420 μμ'; X: 2970; Y: 4200),
    (Typ: 9;  Name: 'A4 210 x 297 μμ'; X: 2100; Y: 2970),
    (Typ: 10; Name: 'A4 small sheet, 210 x 297 μμ'; X: 2100; Y: 2970),
    (Typ: 11; Name: 'A5 148 x 210 μμ'; X: 1480; Y: 2100),
    (Typ: 12; Name: 'B4 250 x 354 μμ'; X: 2500; Y: 3540),
    (Typ: 13; Name: 'B5 182 x 257 μμ'; X: 1820; Y: 2570),
    (Typ: 14; Name: 'Folio, 8 1/2 x 13"'; X: 2159; Y: 3302),
    (Typ: 15; Name: 'Quarto Sheet, 215 x 275 μμ'; X: 2150; Y: 2750),
    (Typ: 16; Name: '10 x 14"'; X: 2540; Y: 3556),
    (Typ: 17; Name: '11 x 17"'; X: 2794; Y: 4318),
    (Typ: 18; Name: 'Note, 8 1/2 x 11"'; X: 2159; Y: 2794),
    (Typ: 19; Name: '9 Envelope, 3 7/8 x 8 7/8"'; X: 984;  Y: 2254),
    (Typ: 20; Name: '#10 Envelope, 4 1/8  x 9 1/2"'; X: 1048; Y: 2413),
    (Typ: 21; Name: '#11 Envelope, 4 1/2 x 10 3/8"'; X: 1143; Y: 2635),
    (Typ: 22; Name: '#12 Envelope, 4 3/4 x 11"'; X: 1207; Y: 2794),
    (Typ: 23; Name: '#14 Envelope, 5 x 11 1/2"'; X: 1270; Y: 2921),
    (Typ: 24; Name: 'C Sheet, 17 x 22"'; X: 4318; Y: 5588),
    (Typ: 25; Name: 'D Sheet, 22 x 34"'; X: 5588; Y: 8636),
    (Typ: 26; Name: 'E Sheet, 34 x 44"'; X: 8636; Y: 11176),
    (Typ: 27; Name: 'DL Envelope, 110 x 220 μμ'; X: 1100; Y: 2200),
    (Typ: 28; Name: 'C5 Envelope, 162 x 229 μμ'; X: 1620; Y: 2290),
    (Typ: 29; Name: 'C3 Envelope,  324 x 458 μμ'; X: 3240; Y: 4580),
    (Typ: 30; Name: 'C4 Envelope,  229 x 324 μμ'; X: 2290; Y: 3240),
    (Typ: 31; Name: 'C6 Envelope,  114 x 162 μμ'; X: 1140; Y: 1620),
    (Typ: 32; Name: 'C65 Envelope, 114 x 229 μμ'; X: 1140; Y: 2290),
    (Typ: 33; Name: 'B4 Envelope,  250 x 353 μμ'; X: 2500; Y: 3530),
    (Typ: 34; Name: 'B5 Envelope,  176 x 250 μμ'; X: 1760; Y: 2500),
    (Typ: 35; Name: 'B6 Envelope,  176 x 125 μμ'; X: 1760; Y: 1250),
    (Typ: 36; Name: 'Italy Envelope, 110 x 230 μμ'; X: 1100; Y: 2300),
    (Typ: 37; Name: 'Monarch Envelope, 3 7/8 x 7 1/2"'; X:984; Y:1905),
    (Typ: 38; Name: '6 3/4 Envelope, 3 5/8 x 6 1/2"'; X: 920; Y: 1651),
    (Typ: 39; Name: 'US Std Fanfold, 14 7/8 x 11"'; X: 3778; Y: 2794),
    (Typ: 40; Name: 'German Std Fanfold, 8 1/2 x 12"'; X: 2159; Y: 3048),
    (Typ: 41; Name: 'German Legal Fanfold, 8 1/2 x 13"'; X: 2159; Y: 3302),
    (Typ: 42; Name: 'B4 (ISO) 250 x 353 μμ'; X: 2500; Y: 3530),
    (Typ: 43; Name: 'Japanese Postcard 100 x 148 μμ'; X: 1000; Y: 1480),
    (Typ: 44; Name: '9 x 11"'; X: 2286; Y: 2794),
    (Typ: 45; Name: '10 x 11"'; X: 2540; Y: 2794),
    (Typ: 46; Name: '15 x 11"'; X: 3810; Y: 2794),
    (Typ: 47; Name: 'Envelope Invite 220 x 220 μμ'; X: 2200; Y: 2200),
    (Typ: 50; Name: 'Letter Extra 9 \ 275 x 12"'; X: 2355; Y: 3048),
    (Typ: 51; Name: 'Legal Extra 9 \275 x 15"'; X: 2355; Y: 3810),
    (Typ: 52; Name: 'Tabloid Extra 11.69 x 18"'; X: 2969; Y: 4572),
    (Typ: 53; Name: 'A4 Extra 9.27 x 12.69"'; X: 2354; Y: 3223),
    (Typ: 54; Name: 'Letter Transverse 8 \275 x 11"'; X: 2101; Y: 2794),
    (Typ: 55; Name: 'A4 Transverse 210 x 297 μμ'; X: 2100; Y: 2970),
    (Typ: 56; Name: 'Letter Extra Transverse 9\275 x 12"'; X: 2355; Y: 3048),
    (Typ: 57; Name: 'SuperASuperAA4 227 x 356 μμ'; X: 2270; Y: 3560),
    (Typ: 58; Name: 'SuperBSuperBA3 305 x 487 μμ'; X: 3050; Y: 4870),
    (Typ: 59; Name: 'Letter Plus 8.5 x 12.69"'; X: 2159; Y: 3223),
    (Typ: 60; Name: 'A4 Plus 210 x 330 μμ'; X: 2100; Y: 3300),
    (Typ: 61; Name: 'A5 Transverse 148 x 210 μμ'; X: 1480; Y: 2100),
    (Typ: 62; Name: 'B5 (JIS) Transverse 182 x 257 μμ'; X: 1820; Y: 2570),
    (Typ: 63; Name: 'A3 Extra 322 x 445 μμ'; X: 3220; Y: 4450),
    (Typ: 64; Name: 'A5 Extra 174 x 235 μμ'; X: 1740; Y: 2350),
    (Typ: 65; Name: 'B5 (ISO) Extra 201 x 276 μμ'; X: 2010; Y: 2760),
    (Typ: 66; Name: 'A2 420 x 594 μμ'; X: 4200; Y: 5940),
    (Typ: 67; Name: 'A3 Transverse 297 x 420 μμ'; X: 2970; Y: 4200),
    (Typ: 68; Name: 'A3 Extra Transverse 322 x 445 μμ'; X: 3220; Y: 4450));

implementation

{$R vgrPrinter.res}

var
  LocalPrinterBitmap: TBitmap;
  NetworkPrinterBitmap: TBitmap;
  FPrinters: TvgrPrinters;

function Printers: TvgrPrinters;
begin
  Result := FPrinters;
end;

/////////////////////////////////////////////////
//
// TvgrPrinter
//
/////////////////////////////////////////////////
constructor TvgrPrinter.Create(APrinters: TvgrPrinters);
begin
  inherited Create;
  FLocal := True;
  FPrinters := APrinters;
  FPageMargins := TvgrPageMargins.Create;
  FPaperNames := TStringList.Create;
end;

constructor TvgrPrinter.Create(APrinters: TvgrPrinters; const APrinterName: string);
begin
  Create(APrinters);
  FPrinterName := APrinterName;
  UpdateLocal;
end;

destructor TvgrPrinter.Destroy;
begin
  ClearStructures;
  FPageMargins.Free;
  FPaperNames.Free;
  inherited;
end;

procedure TvgrPrinter.ClearStructures;
begin
  FDriverName := '';
  FDeviceName := '';
  FPortName := '';
  FComment := '';
  FPixelsPerX := -1;
  FPixelsPerY := -1;

  if FDevMode <> nil then
  begin
    FreeMem(FDevMode);
    FDevMode := nil;
    FDevModeSize := 0;
  end;
  if FPaperSizes <> nil then
  begin
    FreeMem(FPaperSizes);
    FPaperSizes := nil;
  end;
  if FPaperDimensions <> nil then
  begin
    FreeMem(FPaperDimensions);
    FPaperDimensions := nil;
  end;
  FPaperNames.Clear;
  FPageMargins.Left := 0;
  FPageMargins.Top := 0;
  FPageMargins.Right := 0;
  FPageMargins.Bottom := 0;
  FInitializated := False;
end;

procedure TvgrPrinter.UpdateLocal;
var
  APrinterHandle: THandle;
  ABuffer: Pointer;
  ANeeded: Cardinal;
begin
  FLocal := True;
  if OpenPrinter(PChar(PrinterName), APrinterHandle, nil) then
  begin
    GetPrinter(APrinterHandle, 5, nil, 0, @ANeeded);
    if ANeeded <> 0 then
    begin
      GetMem(ABuffer, ANeeded);
      GetPrinter(APrinterHandle, 5, ABuffer, ANeeded, @ANeeded);
      with PPrinterInfo5(ABuffer)^ do
        FLocal := (Copy(StrPas(pPortName), 1, 2) <> '\\') and
                  ((Attributes and PRINTER_ATTRIBUTE_SHARED) = 0);
      FreeMem(ABuffer);
    end;
    ClosePrinter(APrinterHandle);
  end;
end;

procedure TvgrPrinter.Initialize;
var
  APrinterHandle: THandle;
  APrinterInfo2: PPrinterInfo2;
  I, ANeeded, ACount: Integer;
  ADevMode: TDevMode;
  OldError: UINT;
  ABuffer: PChar;
  APrinterDC: HDC;

  procedure CopyString(var S: string; ASource: PChar);
  var
    ALen: Integer;
  begin
    if ASource = nil then
      S := ''
    else
    begin
      ALen := StrLen(ASource);
      SetLength(S, ALen);
      if ALen > 0 then
        MoveMemory(@S[1], ASource, ALen);
    end;
  end;
  
begin
  ClearStructures;

  APrinterHandle := 0;
  APrinterInfo2 := nil;
  if OpenPrinter(PChar(PrinterName), APrinterHandle, nil) then
  begin
    try
      GetPrinter(APrinterHandle, 2, nil, 0, @ANeeded);
      if ANeeded <> 0 then
      begin
        GetMem(APrinterInfo2, ANeeded);
        if GetPrinter(APrinterHandle, 2, APrinterInfo2, ANeeded, @ANeeded) then
        begin
          CopyString(FDeviceName, APrinterInfo2.pPrinterName);
          CopyString(FDriverName, APrinterInfo2.pDriverName);
          CopyString(FPortName, APrinterInfo2.pPortName);
          CopyString(FComment, APrinterInfo2.pComment);

          OldError := SetErrorMode(SEM_FAILCRITICALERRORS);
          try
            FDevModeSize := DocumentProperties(0,
                                               APrinterHandle,
                                               PChar(FDeviceName),
                                               ADevMode,
                                               ADevMode,
                                               0);
          finally
            SetErrorMode(OldError);
          end;

          if FDevModeSize<=0 then
            FDevMode := AllocMem(SizeOf(TDeviceMode))
          else
            FDevMode := AllocMem(FDevModeSize);

          FDevMode.dmSize := sizeof(TDeviceMode);
          if FDevModeSize <= 0 then
          begin
            FDevMode.dmFields := DM_ORIENTATION or DM_PAPERSIZE;
            FDevMode.dmOrientation := DMORIENT_PORTRAIT;
            FDevMode.dmPaperSize := DMPAPER_A4;
          end
          else
            if DocumentProperties(0, APrinterHandle, PChar(FDeviceName), FDevMode^, FDevMode^, DM_OUT_BUFFER) < 0 then
            begin
              FreeMem(FDevMode);
              FDevMode := nil;
              FDevModeSize := 0;
            end;

          //
          OldError := SetErrorMode(SEM_FAILCRITICALERRORS);
          try
            APrinterDC := CreateIC(PChar(FDriverName),
                                   PChar(FDeviceName),
                                   PChar(FPortName),
                                   FDevMode);
          finally
            SetErrorMode(OldError);
          end;

          if APrinterDC <> 0 then
          begin
            FPixelsPerX := GetDeviceCaps(APrinterDC, LOGPIXELSX);
            FPixelsPerY := GetDeviceCaps(APrinterDC, LOGPIXELSY);

            with FPageMargins do
            begin
              Left := ConvertPixelsToTwipsX(GetDeviceCaps(APrinterDC, PHYSICALOFFSETX),
                                            FPixelsPerX);
              Top := ConvertPixelsToTwipsY(GetDeviceCaps(APrinterDC, PHYSICALOFFSETY),
                                           FPixelsPerY);
              Right := ConvertPixelsToTwipsX(GetDeviceCaps(APrinterDC, PHYSICALWIDTH) -
                                             GetDeviceCaps(APrinterDC, HORZRES) -
                                             GetDeviceCaps(APrinterDC, PHYSICALOFFSETX),
                                             FPixelsPerX);
              Bottom := ConvertPixelsToTwipsY(GetDeviceCaps(APrinterDC, PHYSICALHEIGHT) -
                                              GetDeviceCaps(APrinterDC, VERTRES) -
                                              GetDeviceCaps(APrinterDC, PHYSICALOFFSETY),
                                              FPixelsPerY);
            end;

            OldError := SetErrorMode(SEM_FAILCRITICALERRORS);
            try
              // supported paper sizes
              ACount := DeviceCapabilities(PChar(FDeviceName),
                                           PChar(FPortName),
                                           DC_PAPERS,
                                           nil,
                                           FDevMode);
              if ACount > 0 then
              begin
                GetMem(FPaperSizes, ACount * 2);
                DeviceCapabilities(PChar(FDeviceName),
                                   PChar(FPortName),
                                   DC_PAPERS,
                                   PChar(FPaperSizes),
                                   FDevMode);
              end
              else
              begin
                GetMem(FPaperSizes,sizeof(Word));
                FPaperSizes[0] := DMPAPER_A4;
              end;

              // supported paper dimensions
              ACount := DeviceCapabilities(PChar(FDeviceName),
                                           PChar(FPortName),
                                           DC_PAPERS,
                                           nil,
                                           FDevMode);
              if ACount > 0 then
              begin
                GetMem(FPaperDimensions, ACount * SizeOf(TPoint));
                DeviceCapabilities(PChar(FDeviceName),
                                   PChar(FPortName),
                                   DC_PAPERSIZE,
                                   PChar(FPaperDimensions),
                                   FDevMode);
              end
              else
              begin
                GetMem(FPaperDimensions, SizeOf(TPoint));
                FPaperDimensions^[0].X := A4_PaperWidth;
                FPaperDimensions^[0].Y := A4_PaperHeight;
              end;

              // supported paper names
              ACount := DeviceCapabilities(PChar(FDeviceName),
                                           PChar(FPortName),
                                           DC_PAPERNAMES,
                                           nil,
                                           FDevMode);
              if ACount > 0 then
              begin
                GetMem(ABuffer, ACount * 64);
                try
                  DeviceCapabilities(PChar(FDeviceName),
                                     PChar(FPortName),
                                     DC_PAPERNAMES,
                                     ABuffer,
                                     FDevMode);
                  FPaperNames.Clear;
                  for I := 0 to ACount - 1 do
                    FPaperNames.Add(ABuffer + I * 64);
                finally
                  FreeMem(ABuffer);
                end;
              end
              else
              begin
                FPaperNames.Clear;
                FPaperNames.Add(PaperInfo[DMPAPER_A4].Name);
              end;
            finally
              SetErrorMode(OldError);
              DeleteDC(APrinterDC);
            end;
          end;
          FInitializated := True;
        end;
      end;
    finally
      if APrinterInfo2 <> nil then
        FreeMem(APrinterInfo2);
      if APrinterHandle <> 0 then
        ClosePrinter(APrinterHandle);
    end;
  end;
end;

function TvgrPrinter.GetPaperCount: Integer;
begin
  if not Initializated then
    Initialize;

  if Initializated then
    Result := FPaperNames.Count
  else
    Result := -1;
end;

function TvgrPrinter.GetPaperName(Index: Integer): string;
begin
  if not Initializated then
    Initialize;

  if Initializated then
    Result := FPaperNames[Index]
  else
    Result := '';
end;

function TvgrPrinter.GetPaperSize(Index: Integer): Integer;
begin
  if not Initializated then
    Initialize;

  if Initializated then
    Result := FPaperSizes^[Index]
  else
    Result := -1;
end;

function TvgrPrinter.GetPaperDimension(Index: Integer): TPoint;
begin
  if not Initializated then
    Initialize;

  if Initializated then
    Result := FPaperDimensions^[Index]
  else
  begin
    Result.X := -1;
    Result.Y := -1;
  end;
end;

procedure TvgrPrinter.SetPrinterName(const Value: string);
begin
  if FPrinterName <> Value then
  begin
    FPrinterName := Value;
    Initialize;
  end;
end;

function TvgrPrinter.GetDefault: Boolean;
begin
  Result := FPrinters.DefaultPrinter = Self;
end;

function TvgrPrinter.GetImage: TBitmap;
begin
  if Local then
    Result := LocalPrinterBitmap
  else
    Result := NetworkPrinterBitmap;
end;

function TvgrPrinter.GetPaperNames: TStrings;
begin
  if not Initializated then
    Initialize;

  if Initializated then
    Result := FPaperNames
  else
    Result := nil;
end;

function TvgrPrinter.GetDriverName: string;
begin
  if not Initializated then
    Initialize;

  Result := FDriverName;
end;

function TvgrPrinter.GetPortName: string;
begin
  if not Initializated then
    Initialize;

  Result := FPortName;
end;

function TvgrPrinter.GetDeviceName: string;
begin
  if not Initializated then
    Initialize;

  Result := FDeviceName;
end;

function TvgrPrinter.GetPixelsPerX: Integer;
begin
  if not Initializated then
    Initialize;

  Result := FPixelsPerX;
end;

function TvgrPrinter.GetPixelsPerY: Integer;
begin
  if not Initializated then
    Initialize;

  Result := FPixelsPerY;
end;

function TvgrPrinter.GetComment: string;
begin
  if not Initializated then
    Initialize;

  Result := FComment;
end;

function TvgrPrinter.GetPageMargins: TvgrPageMargins;
begin
  if not Initializated then
    Initialize;

  Result := FPageMargins;
end;

function TvgrPrinter.GetStatus: string;
begin
  Result := '';
end;

function TvgrPrinter.ShowPropertiesDialog(AParentWindowHandle: THandle): Boolean;
var
  hPrinter: THandle;
  ADevMode: PDevMode;
  ADevModeHandle: THandle;
begin
  Result := False;
  if not FInitializated then
    Initialize;
  if not OpenPrinter(PChar(PrinterName), hPrinter, nil) then
    exit;

  // copy DEVMODE structure
  ADevModeHandle := GlobalAlloc(GHND, FDevModeSize);
  ADevMode := GlobalLock(ADevModeHandle);
  try
    CopyMemory(ADevMode, FDevMode, FDevModeSize);

    if DocumentProperties(AParentWindowHandle, hPrinter, PChar(PrinterName), ADevMode^, ADevMode^,
                          DM_IN_BUFFER or DM_OUT_BUFFER or DM_IN_PROMPT) = IDOK then
    begin
      CopyMemory(FDevMode, ADevMode, FDevModeSize);
      Result := True;
    end;
  finally
    GlobalUnlock(ADevModeHandle);
    GlobalFree(ADevModeHandle);
  end;
end;

function TvgrPrinter.AllocNewDevMode: PDEVMODE;
begin
  if not Initializated then
    Initialize;

  if (FDevMode = nil) or (FDevModeSize = 0) then
    Result := nil
  else
  begin
    GetMem(Result, FDevModeSize);
    CopyMemory(Result, FDevMode, FDevModeSize);
  end;
end;

procedure TvgrPrinter.FindPaper(APageProperties: TvgrPageProperties;
                                var APaperIndex: Integer;
                                var APaperSize: Integer;
                                var APaperName: string;
                                var APaperOrientation: TvgrPageOrientation;
                                var APaperDimensions: TPoint);
var
  ASize: TPoint;
begin
  ASize := Point(APageProperties.TenthsMMWidth, APageProperties.TenthsMMHeight);
  APaperIndex := 0;
  while (APaperIndex < PaperCount) and
        ((PaperDimensions[APaperIndex].X <> ASize.X) or
         (PaperDimensions[APaperIndex].Y <> ASize.Y)) do
    Inc(APaperIndex);
  if APaperIndex >= PaperCount then
  begin
    APaperIndex := 0;
    while (APaperIndex < PaperCount) and
          ((PaperDimensions[APaperIndex].Y <> ASize.X) or
           (PaperDimensions[APaperIndex].X <> ASize.Y)) do
      Inc(APaperIndex);
    if APaperIndex >= PaperCount then
    begin
      APaperIndex := -1;
      APaperSize := -1;
      APaperName := '';
      APaperOrientation := vgrpoPortrait;
      APaperDimensions := ASize;
      exit;
    end
    else
    begin
      APaperOrientation := vgrpoLandscape;
      APaperDimensions := Point(PaperDimensions[APaperIndex].Y, PaperDimensions[APaperIndex].X);
    end
  end
  else
  begin
    APaperOrientation := vgrpoPortrait;
    APaperDimensions := Point(PaperDimensions[APaperIndex].X, PaperDimensions[APaperIndex].Y);
  end;
  APaperSize := PaperSizes[APaperIndex];
  APaperName := PaperNames[APaperIndex];
  APaperDimensions := PaperDimensions[APaperIndex];
end;

/////////////////////////////////////////////////
//
// TvgrPrinters
//
/////////////////////////////////////////////////
constructor TvgrPrinters.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

destructor TvgrPrinters.Destroy;
begin
  ClearPrintersList;
  FList.Free;
  inherited;
end;

function TvgrPrinters.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TvgrPrinters.GetItem(Index: Integer): TvgrPrinter;
begin
  Result := TvgrPrinter(FList[Index]);
end;

procedure TvgrPrinters.ClearPrintersList;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  FList.Clear;
end;

procedure TvgrPrinters.UpdateDefaultPrinterName;
var
  ABuffer: Pointer;
  ANeeded, ANumInfo: Cardinal;
begin
  FDefaultPrinterName := '';
  ABuffer := nil;
  ANeeded := 0;
  case Win32Platform of
    VER_PLATFORM_WIN32_WINDOWS:
      begin
        EnumPrinters(PRINTER_ENUM_DEFAULT, nil, 5, ABuffer, 0, ANeeded, ANumInfo);
        if ANeeded <> 0 then
        begin
          GetMem(ABuffer, ANeeded);
          try
            if EnumPrinters(PRINTER_ENUM_DEFAULT, nil, 5, ABuffer, ANeeded, ANeeded, ANumInfo) then
              if ANumInfo > 0 then
                FDefaultPrinterName := StrPas(PPrinterInfo5(ABuffer).pPrinterName);
          finally
            FreeMem(ABuffer);
          end;
        end;
      end;
    VER_PLATFORM_WIN32_NT:
      begin
{
        EnumPrinters(PRINTER_ENUM_DEFAULT, nil, 4, Buffer, 0, BytesNeeded,NumInfo);
        if BytesNeeded<>0 then
        begin
          GetMem(Buffer,BytesNeeded);
          try
            if EnumPrinters(PRINTER_ENUM_DEFAULT,nil,4,Buffer,BytesNeeded,BytesNeeded,NumInfo) then
              if NumInfo>0 then
                Result:=StrPas(PPrinterInfo4(Buffer).pPrinterName);
          finally
            FreeMem(Buffer);
          end
        end
}
      end;
  end;

  if FDefaultPrinterName = '' then
  begin
    SetLength(FDefaultPrinterName, 80);
    ANeeded := GetProfileString('windows', 'device', '', @FDefaultPrinterName[1], 79);
    SetLength(FDefaultPrinterName, ANeeded);
    ANeeded := pos(',', FDefaultPrinterName);
    if ANeeded <> 0 then
      FDefaultPrinterName := Copy(FDefaultPrinterName, 1, ANeeded - 1);
  end;
end;

procedure TvgrPrinters.UpdatePrintersList;
var
  I: integer;
  APrinter: TvgrPrinter;
  ATempBuffer, ABuffer: PChar;
  ALevel, AFlags, ANeeded, ANumInfo: Cardinal;
begin
  ClearPrintersList;

  if Win32Platform = VER_PLATFORM_WIN32_NT then
  begin
    AFlags := PRINTER_ENUM_CONNECTIONS or PRINTER_ENUM_LOCAL;
    ALevel := 4;
  end
  else
  begin
    AFlags := PRINTER_ENUM_LOCAL;
    ALevel := 5;
  end;

  ABuffer := nil;
  ANeeded := 0;
  EnumPrinters(AFlags, nil, ALevel, ABuffer, 0, ANeeded, ANumInfo);
  if ANeeded <> 0 then
  begin
    GetMem(ABuffer, ANeeded);
    ATempBuffer := ABuffer;
    try
      if EnumPrinters(AFlags, nil, ALevel, ABuffer, ANeeded, ANeeded, ANumInfo) then
        for I := 0 to ANumInfo - 1 do
        begin
          if ALevel = 4 then
          begin
            APrinter := TvgrPrinter.Create(Self, PPrinterInfo4(ABuffer)^.pPrinterName);
            Inc(ABuffer, SizeOf(PRINTER_INFO_4));
          end
          else
          begin
            APrinter := TvgrPrinter.Create(Self, PPrinterInfo5(ABuffer)^.pPrinterName);
            Inc(ABuffer, SizeOf(PRINTER_INFO_5));
          end;
          FList.Add(APrinter);
        end;
    finally
      FreeMem(ATempBuffer);
    end;
  end;
end;

procedure TvgrPrinters.Update;
begin
  UpdateDefaultPrinterName;
  UpdatePrintersList;
end;

function TvgrPrinters.IndexOfPrinterName(const APrinterName: string): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].PrinterName <> APrinterName) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrPrinters.GetDefaultPrinter: TvgrPrinter;
var
  I: Integer;
begin
  I := IndexOfPrinterName(FDefaultPrinterName);
  if (I < Count) and (I >= 0) then
    Result := Items[I]
  else
    Result := nil;
end;

function TvgrPrinters.GetPaperDimensionsBySize(APaperSize: Integer; APrinter: TvgrPrinter): TPoint;
var
  I: Integer;
begin
  if APrinter = nil then
  begin
    I := 0;
    while (I < DEF_PAPERCOUNT) and (PaperInfo[I].Typ <> APaperSize) do Inc(I);
    if I < DEF_PAPERCOUNT then
      Result := Point(PaperInfo[I].X, PaperInfo[I].Y);
  end
  else
  begin
    I := 0;
    while (I < APrinter.PaperCount) and (APrinter.PaperSizes[I] <> APaperSize) do Inc(I);
    if I < APrinter.PaperCount then
      Result := APrinter.PaperDimensions[I];
  end;
end;

procedure TvgrPrinters.FindPaper(APageProperties: TvgrPageProperties;
                                 APrinter: TvgrPrinter;
                                 var APaperIndex: Integer;
                                 var APaperSize: Integer;
                                 var APaperName: string;
                                 var APaperOrientation: TvgrPageOrientation;
                                 var APaperDimensions: TPoint);
var
  ASize: TPoint;
begin
  if APrinter = nil then
  begin
    ASize := Point(APageProperties.TenthsMMWidth, APageProperties.TenthsMMHeight);
    APaperIndex := 0;
    while (APaperIndex < DEF_PAPERCOUNT) and
          ((PaperInfo[APaperIndex].X <> ASize.X) or
           (PaperInfo[APaperIndex].Y <> ASize.Y)) do
      Inc(APaperIndex);
    if APaperIndex >= DEF_PAPERCOUNT then
    begin
      APaperIndex := 0;
      while (APaperIndex < DEF_PAPERCOUNT) and
            ((PaperInfo[APaperIndex].Y <> ASize.X) or
             (PaperInfo[APaperIndex].X <> ASize.Y)) do
        Inc(APaperIndex);
      if APaperIndex >= DEF_PAPERCOUNT then
      begin
        APaperIndex := -1;
        APaperSize := -1;
        APaperName := '';
        if ASize.X <= ASize.Y then
          APaperOrientation := vgrpoPortrait
        else
          APaperOrientation := vgrpoLandscape;
        APaperDimensions := ASize;
        exit;
      end
      else
      begin
        APaperOrientation := vgrpoLandscape;
        APaperDimensions := Point(PaperInfo[APaperIndex].Y, PaperInfo[APaperIndex].X); 
      end
    end
    else
    begin
      APaperOrientation := vgrpoPortrait;
      APaperDimensions := Point(PaperInfo[APaperIndex].X, PaperInfo[APaperIndex].Y);
    end;
    APaperSize := PaperInfo[APaperIndex].Typ;
    APaperName := PaperInfo[APaperIndex].Name;
    APaperDimensions := Point(PaperInfo[APaperIndex].X, PaperInfo[APaperIndex].Y);
  end
  else
    APrinter.FindPaper(APageProperties, APaperIndex, APaperSize, APaperName, APaperOrientation, APaperDimensions);
end;

initialization

  NetworkPrinterBitmap := TBitmap.Create;
  NetworkPrinterBitmap.LoadFromResourceName(hInstance, 'VGR_BMP_NETWORKPRINTER');
  NetworkPrinterBitmap.Transparent := True;

  LocalPrinterBitmap := TBitmap.Create;
  LocalPrinterBitmap.LoadFromResourceName(hInstance, 'VGR_BMP_LOCALPRINTER');
  LocalPrinterBitmap.Transparent := True;

  FPrinters := TvgrPrinters.Create;
  FPrinters.Update;

finalization

  FreeAndNil(FPrinters);
  FreeAndNil(LocalPrinterBitmap);
  FreeAndNil(NetworkPrinterBitmap);

end.
