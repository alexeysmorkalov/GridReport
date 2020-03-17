{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{ This module holds classes which implements controls used in TvgrWorkbookGrid and TvgrWorkbookPreview. }
unit vgr_Controls;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, SysUtils, Controls, ExtCtrls, Windows, graphics, stdctrls, forms, messages,
  math, ImgList,

  vgr_CommonClasses, vgr_DataStorage;

const
  svgrControlParentGUID: TGUID = '{9AD8DD22-1ACF-40ED-BE50-4F888B521B9C}';
  svgrScrollBarParentGUID: TGUID = '{5A372535-5249-4D79-9E87-38F21F08C755}';
  svgrSheetsPanelParentGUID: TGUID = '{709B9610-0259-454F-A93B-5628570898C4}';
  svgrSheetsPanelResizerParentGUID: TGUID = '{26957074-7E05-4B45-B859-C8BD161B2F0F}';
  svgrSizerParent: TGUID = '{4D1BA17F-48A9-4681-9B7A-33B8D13F7E4D}';
  svgrControlRectParent: TGUID = '{23DB2C3C-33EA-48E0-A83A-4A7AA48EA10B}';

  svgrSheetsPanelBitmaps = 'VGR_BMP_SHEETSPANELBITMAP';
  { Font name for sections panel }
  sDefSheetsPanelFontName = 'Tahoma';
  { Font size for sections panel }
  sDefSheetsPanelFontSize = 6;
  cSheetsPanelResizerWidth = 6;
  cDefaultSheetsCaptionWidth = 200;

type
  TvgrScrollBar = class;
  TvgrOptionsScrollBar = class;
  TvgrSheetsPanel = class;
  TvgrSheetsPanelRect = class;
  TvgrPointInfo = class;

  /////////////////////////////////////////////////
  //
  // IvgrControlParent
  //
  /////////////////////////////////////////////////
  { Base interface for all controls in this module }
  IvgrControlParent = interface
  ['{9AD8DD22-1ACF-40ED-BE50-4F888B521B9C}']
    { Called when the user moves the mouse pointer while the mouse pointer is over a control.
Parameters:
  Sender - TControl
  Shift - TShiftState
  X - Integer
  Y - Integer}
    procedure DoMouseMove(Sender: TControl; Shift: TShiftState; X, Y: Integer);
    { Called when the user presses a mouse button with the mouse pointer over a control.
Parameters:
  Sender - TControl
  Button - TMouseButton
  Shift - TShiftState
  X Integer
  Y - Integer
  AExecuteDefault - Boolean}
    procedure DoMouseDown(Sender: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
    { Called when the user releases a mouse button that was pressed with the mouse pointer over a component.
Parameters:
  Sender - TControl
  Button - TMouseButton
  Shift - TShiftState
  X - Integer
  Y - Integer
  AExecuteDefault - Boolean}
    procedure DoMouseUp(Sender: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
    { Called when the user double-clicks the left mouse button when the mouse pointer is over the control.
Parameters:
  Sender - TControl
  AExecuteDefault - Boolean}
    procedure DoDblClick(Sender: TControl; var AExecuteDefault: Boolean);
    { Called when the user drags an object over a control.
Parameters:
  AChild - TControl
  Source - TObject
  X - Integer
  Y - Integer
  State - TDragState
  Accept - Boolean}
    procedure DoDragOver(AChild: TControl; Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    { Called when the user drops an object being dragged.
Parameters:
  AChild - TControl
  Source - TObject
  X - Integer
  Y - Integer}
    procedure DoDragDrop(AChild: TControl; Source: TObject; X, Y: Integer);
  end;

  /////////////////////////////////////////////////
  //
  // IvgrScrollBarParent
  //
  /////////////////////////////////////////////////
  { The given interface should realize a class, if it is parent for TvgrScrollBar. }
  IvgrScrollBarParent = interface(IvgrControlParent)
  ['{5A372535-5249-4D79-9E87-38F21F08C755}']
    { Return a options for TvgrScrollBar.
Parameters:
  AScrollBar - TvgrScrollBar
Return value:
  TvgrOptionsScrollBar}
    function GetScrollBarOptions(AScrollBar: TvgrScrollBar): TvgrOptionsScrollBar;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrSheetsPanelParent
  //
  /////////////////////////////////////////////////
  { The given interface should realize a class, if it is parent for TvgrSheetsPanel. }
  IvgrSheetsPanelParent = interface(IvgrControlParent)
  ['{709B9610-0259-454F-A93B-5628570898C4}']
    { Returns a workbook which used by TvgrSheetsPanel
Return value:
  TvgrWorkbook}
    function GetWorkbook: TvgrWorkbook;
    { Returns a workbook which used by TvgrSheetsPanel
Return value:
  TvgrWorksheet}
    function GetActiveWorksheet: TvgrWorksheet;
    { Returns a index of active worksheet in workbook
Return value:
  Integer}
    function GetActiveWorksheetIndex: Integer;
    { Sets a index of active worksheet in workbook
Parameters:
  Value - Integer, index of active worksheet}
    procedure SetActiveWorksheetIndex(Value: Integer);
    { Show popup menu
Parameters:
  APoint - TPoint
  APointInfo - TvgrPointInfo}
    procedure ShowPopupMenu(const APoint: TPoint; APointInfo: TvgrPointInfo);
    { Workbook which used by TvgrSheetsPanel }
    property Workbook: TvgrWorkbook read GetWorkbook;
    { Active worksheet which used by TvgrSheetsPanel }
    property ActiveWorksheet: TvgrWorksheet read GetActiveWorksheet;
    { Index of active worksheet in workbook }
    property ActiveWorksheetIndex: Integer read GetActiveWorksheetIndex write SetActiveWorksheetIndex;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrSheetsPanelResizerParent
  //
  /////////////////////////////////////////////////
  { The given interface should realize a class, if it is parent for TvgrSheetsPanelResizer. }
  IvgrSheetsPanelResizerParent = interface(IvgrControlParent)
  ['{26957074-7E05-4B45-B859-C8BD161B2F0F}']
    { Called then user resize.
Parameters:
  AOffset - Integer}
    procedure DoResize(AOffset: Integer);
  end;

  /////////////////////////////////////////////////
  //
  // IvgrSizerParent
  //
  /////////////////////////////////////////////////
  { The given interface should realize a class, if it is parent for TvgrSizerParent. }
  IvgrSizerParent = interface(IvgrControlParent)
  ['{4D1BA17F-48A9-4681-9B7A-33B8D13F7E4D}']
  end;

  /////////////////////////////////////////////////
  //
  // TCustomControlAccess
  //
  /////////////////////////////////////////////////
  TCustomControlAccess = class(TCustomControl)
  end;

  /////////////////////////////////////////////////
  //
  // TComponentAccess
  //
  /////////////////////////////////////////////////
  TComponentAccess = class(TWinControl)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPointInfo
  //
  /////////////////////////////////////////////////
  { This class contains information about the specific control spot.
    Each control define Each control defines the class inherited from TvgrPointInfo. }
  TvgrPointInfo = class(TObject)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScrollBarPointInfo
  //
  /////////////////////////////////////////////////
  { Contains information the specific TvgrScrollBar spot. }
  TvgrScrollBarPointInfo = class(TvgrPointInfo)
  private
    FScrollBar: TvgrScrollBar;
  public
    {creates TvgrScrollBarPointInfo objects
Parameters:
  AScrollBar - TvgrScrollBar}
    constructor Create(AScrollBar: TvgrScrollBar);
    { Reference to scroolbar which create this object }
    property ScrollBar: TvgrScrollBar read FScrollBar;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrOptionsScrollBar
  //
  /////////////////////////////////////////////////
  { Options of TvgrScrollBar control. }
  TvgrOptionsScrollBar = class(TvgrPersistent)
  private
    FVisible: Boolean;
    procedure SetVisible(Value: Boolean);
  public
  {Creates TvgrOptionsScrollBar control}
    constructor Create; override;
    procedure Assign(Source: TPersistent); override;
  published
    { Use the Visible property to control the visibility of the control at runtime.
      If Visible is true, the control appears.
      If Visible is false, the control is not visible. }
    property Visible: Boolean read FVisible write SetVisible default True;
  end;

  TOnScrollMessage = procedure(Sender: TObject; var Msg: TWMScroll) of object;
  /////////////////////////////////////////////////
  //
  // TvgrScrollBar
  //
  /////////////////////////////////////////////////
  { TvgrScrollBar implements scroll bar control, this class very similar to TScrollBar, but have some addititional functionality. }
  TvgrScrollBar = class(TWinControl)
  private
    FKind: TScrollBarKind;
    FPosition: Integer;
    FOnScrollMessage: TOnScrollMessage;
    procedure FillScrollInfo(var AScrollInfo: TScrollInfo);
    procedure SetKind(Value: TScrollBarKind);
    function GetPosition: Integer;
    function GetPageSize: Integer;
    procedure SetPosition(Value : integer);
    procedure WMHScroll(var Msg: TWMHScroll); message CN_HSCROLL;
    procedure WMVScroll(var Msg: TWMVScroll); message CN_VSCROLL;
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CNCtlColorScrollBar(var Msg: TMessage); message CN_CTLCOLORSCROLLBAR;
    function GetScrollBarParent: IvgrScrollBarParent;
    function GetOptions: TvgrOptionsScrollBar;
    function GetMin: Integer;
    function GetMax: Integer;
  protected
    procedure CreateParams(var Params: TCreateParams); override;
    procedure CreateWnd; override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;

    property ScrollBarParent: IvgrScrollBarParent read GetScrollBarParent;
  public
{Creates instance of the TvgrScrollBar control}
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

{Defines all properties of TvgrScrollBar.
Parameters:
      Position - describe current position of the thumb tab.
      Min - define minimum value the Position property can take.
      Max - define maximum value the Position property can take.
The Max and Min properties define the total range over which Position can vary.
      PageSize - PageSize is the size of the thumb tab, measured in the same units as Position, Min, and Max (not pixels). }
    procedure SetParams(Position,Min,Max,PageSize: Integer);
    procedure OptionsChanged;

{Returns information about spot within scroll bar control.
Parameters:
  APoint - TPoint
  APointInfo - TvgrPointInfo}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);

    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;
    procedure DragDrop(Source: TObject; X, Y: Integer); override;

    { Set Kind to indicate the orientation of the scroll bar. These are the possible values:
       <br><b>sbHorizontal</b>	- Scroll bar is horizontal
       <br><b>sbVertical</b> - Scroll bar is vertical }
    property Kind: TScrollBarKind read FKind write SetKind;
    { Indicates the current position of the scroll bar. } 
    property PageSize: Integer read GetPageSize;
    { Specifies the minimum position represented by the scroll bar. }
    property Min: Integer read GetMin;
    { Specifies the maximum position represented by the scroll bar. }
    property Max: Integer read GetMax;
    { Indicates the current position of the scroll bar. }
    property Position: integer read GetPosition write SetPosition;
    { Specifies common options of scroll bar. }
    property Options: TvgrOptionsScrollbar read GetOptions;
    { Occurs when messages CN_HSCROLL or CN_VSCROLL received by control. }
    property OnScrollMessage: TOnScrollMessage read FOnScrollMessage write FOnScrollMessage;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrControlRectParent
  //
  /////////////////////////////////////////////////
  IvgrControlRectParent = interface
  ['{23DB2C3C-33EA-48E0-A83A-4A7AA48EA10B}']
    function GetButtonEnabled(AButtonIndex: Integer): Boolean;
    function GetCurRect: TvgrSheetsPanelRect;
    procedure SetCurRect(Value: TvgrSheetsPanelRect);

    property CurRect: TvgrSheetsPanelRect read GetCurRect write SetCurRect;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanelRect
  //
  /////////////////////////////////////////////////
  { Represents area on TvgrSheetsPanel. 
  This is a base abstract class and should not be directly instantiated.}
  TvgrSheetsPanelRect = class(TObject)
  private
    FParentControl: TCustomControl;
    FRect: TRect;
    FMouseOver: Boolean;
    procedure SetMouseOver(Value: Boolean);
    function GetParent: IvgrControlRectParent;
  public
    {Creates TvgrSheetsPanelRect object}
    constructor Create(AParentControl: TCustomControl);
    destructor Destroy; override;
    { Specifies bounds of area.
Parameters:
  ALeft - Integer
  ATop - Integer
  AWidth - Integer
  AHeight - Integer}
    procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
    { Paint area on specified canvas.
Parameters:
  ACanvas - TCanvas}
    procedure Paint(ACanvas: TCanvas); virtual; abstract;
    { Called when user pressing a left mouse button with rect. }
    procedure Click; virtual; abstract;
    property ParentControl: TCustomControl read FParentControl;
    property Parent: IvgrControlRectParent read GetParent;
    { returns True if mouse over this PanelRect. }
    property MouseOver: Boolean read FMouseOver write SetMouseOver;
    { returns Bounds of PanelRect. }
    property BoundsRect: TRect read FRect;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanelButtonRect
  //
  /////////////////////////////////////////////////
  { TvgrSheetsPanelButtonRect implements button on TvgrSheetsPanel. }
  TvgrSheetsPanelButtonRect = class(TvgrSheetsPanelRect)
  private
    FImageList: TImageList;
    FImageIndex: Integer;
    FOnClick: TNotifyEvent;
  public
  { Creates TvgrSheetsPanelButtonRect button
Parameters:
  AParentControl - TCustomControl
  AImageList - TImageList
  AImageIndex - Integer
  AOnClick - TNotifyEvent}
    constructor CreateButton(AParentControl: TCustomControl; AImageList: TImageList; AImageIndex: Integer; AOnClick: TNotifyEvent);
    { Draw button on specified canvas. }
    procedure Paint(ACanvas: TCanvas); override;
    procedure Click; override;
    { Returns image index in ImageList which used while painting button. }
    property ImageIndex: Integer read FImageIndex write FImageIndex;
    { Define images which used while painting button. }
    property ImageList: TImageList read FImageList write FImageList;
    { Occurs when user click mouse button within area of TvgrSheetsPanelButtonRect. }
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanelCaptionRect
  //
  /////////////////////////////////////////////////
  { TvgrSheetsPanelCaptionRect implements worksheet caption on TvgrSheetsPanel. }
  TvgrSheetsPanelCaptionRect = class(TvgrSheetsPanelRect)
  private
    FWorksheet: TvgrWorksheet;
    function GetTitle: string;
    function GetSheetsPanel: TvgrSheetsPanel;
  public
    procedure Paint(ACanvas: TCanvas); override;
    procedure Click; override;
    function GetCaptionWidth: Integer;
    property Worksheet: TvgrWorksheet read FWorksheet write FWorksheet;
    property Title: string read GetTitle;
    property SheetsPanel: TvgrSheetsPanel read GetSheetsPanel;
  end;

  { Specifies type of place within sheets panel:
Items:
    vgrsipNone        - within empty space
    vgrsipFirstButton - within first button
    vgrsipPriorButton - within prior button
    vgrsipNextButton  - within next button
    vgrsipLastButton  - within last button
    vgrSheetCaption   - within caption of on of worksheet }
  TvgrSheetsPanelPointInfoPlace = (vgrsipNone, vgrsipFirstButton, vgrsipPriorButton, vgrsipNextButton, vgrsipLastButton, vgrSheetCaption);
  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanelPointInfo
  //
  /////////////////////////////////////////////////
  { Provide information about spot sheets panel area.
    Object of this type returns by method TvgrWorkbookGrid.GetPointInfoAt, if mouse within sheets panel area.}
  TvgrSheetsPanelPointInfo = class(TvgrPointInfo)
  private
    FPlace: TvgrSheetsPanelPointInfoPlace;
    FSheetIndex: Integer;
  public
    constructor Create(APlace: TvgrSheetsPanelPointInfoPlace; ASheetIndex: Integer);
    { Type of place within header. }
    property Place: TvgrSheetsPanelPointInfoPlace read FPlace;
    { If spot within worksheet caption (Place = vgrSheetCaption)
      returns index of this worksheet in workbook, otherwise returns -1. }
    property SheetIndex: Integer read FSheetIndex;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanel
  //
  /////////////////////////////////////////////////
  { TvgrSheetsPanel represents control which allow to navigate between worksheets of workbook. }
  TvgrSheetsPanel = class(TCustomControl, IvgrControlRectParent)
  private
    FCurRect: TvgrSheetsPanelRect;
    FFirstButton: TvgrSheetsPanelButtonRect;
    FPriorButton: TvgrSheetsPanelButtonRect;
    FNextButton: TvgrSheetsPanelButtonRect;
    FLastButton: TvgrSheetsPanelButtonRect;
    FSheetCaptions: TList;
    FFirstSheetIndex: Integer;
    FFont: TFont;
    function GetSheetCaption(Index: Integer): TvgrSheetsPanelCaptionRect;
    function GetSheetCaptionCount: Integer;
    procedure AlignSheetCaptions;
    procedure OnButtonClick(Sender: TObject);
    procedure WMEraseBkgnd(var Msg : TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CMMouseLeave(var Msg: TMessage); message CM_MOUSELEAVE;
    function GetSheetsPanelParent: IvgrSheetsPanelParent;
    function GetWorkbook: TvgrWorkbook;
    function GetActiveWorksheet: TvgrWorksheet;
    function GetActiveWorksheetIndex: Integer;
    procedure SetActiveWorksheetIndex(Value: Integer);
  protected
    // IvgrControlRectParent
    function GetButtonEnabled(AButtonIndex: Integer): Boolean;
    function GetCurRect: TvgrSheetsPanelRect;
    procedure SetCurRect(Value: TvgrSheetsPanelRect);

    function GetPanelRectAt(X, Y: Integer): TvgrSheetsPanelRect;
    procedure Resize; override;
    procedure Paint; override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;

    property SheetsPanelParent: IvgrSheetsPanelParent read GetSheetsPanelParent;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    { Clear worksheets list. }
    procedure ClearSheetCaptions;
    { Update worksheets list, (from workbook). }
    procedure UpdateSheets;

    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;
    procedure DragDrop(Source: TObject; X, Y: Integer); override;

    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
    { Returns the object of the TvgrSheetsPanelCaptionRect class appropriate AWorksheet.
Return value:
  TvgrSheetsPanelCaptionRect}
    function FindCaptionByWorksheet(AWorksheet: TvgrWorksheet): TvgrSheetsPanelCaptionRect;

    procedure BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
    procedure AfterChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);

    { Returns workbook which linked to this TvgrSheetsPanel, this value must be provided by SheetsPanelParent. }
    property Workbook: TvgrWorkbook read GetWorkbook;
    { Returns current active worksheet, this value must be provided by SheetsPanelParent. }
    property ActiveWorksheet: TvgrWorksheet read GetActiveWorksheet;
    { Gets or set active worksheet index, this value must be provided by SheetsPanelParent. }
    property ActiveWorksheetIndex: Integer read GetActiveWorksheetIndex write SetActiveWorksheetIndex;
    { Use SheetCaptions to gain direct access to a particular TvgrSheetsPanelCaptionRect object
      in the TvgrSheetsPanel. Count of TvgrSheetsPanelCaptionRect objects equals
      to count of Worksheet objects in Workbook, each TvgrSheetsPanelCaptionRect object linked
      to worksheet object. }
    property SheetCaptions[Index: Integer]: TvgrSheetsPanelCaptionRect read GetSheetCaption;
    { Use SheetCaptionCount to determine the number of TvgrSheetsPanelCaptionRect objects
      listed by the SheetCaptions property.
      Count of TvgrSheetsPanelCaptionRect objects equals
      to count of Worksheet objects in Workbook, each TvgrSheetsPanelCaptionRect object linked
      to worksheet object. }
    property SheetCaptionCount: Integer read GetSheetCaptionCount;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanelResizerPointInfo
  //
  /////////////////////////////////////////////////
  { Provide information abount spot within sheets panel resizer area.
    Object of this type returns by method TvgrWorkbookGrid.GetPointInfoAt, if mouse within sheets panel resizer area. }
  TvgrSheetsPanelResizerPointInfo = class(TvgrPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetsPanelResizer
  //
  /////////////////////////////////////////////////
  { TvgrSheetsPanelResizer implements control, used to resize sheets panel. }
  TvgrSheetsPanelResizer = class(TGraphicControl)
  private
    FResizeMode: Boolean;
    FStartMousePoint: TPoint;
    function GetSheetsPanelResizerParent: IvgrSheetsPanelResizerParent;
  protected
    procedure Paint; override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;

    property SheetsPanelResizerParent: IvgrSheetsPanelResizerParent read GetSheetsPanelResizerParent;
  public
    constructor Create(AOwner: TComponent); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;
    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: tvgrPointInfo);   
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSizerPointInfo
  //
  /////////////////////////////////////////////////
  { Provide information abount spot within sizer area.
    Object of this type returns by method TvgrWorkbookGrid.GetPointInfoAt, if mouse within sizer area. }
  TvgrSizerPointInfo = class(TvgrPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSizer
  //
  /////////////////////////////////////////////////
  { TvgrSizer implements simple control which is on an
    intersection horizontal and vertical scroll bars. }  
  TvgrSizer = class(TGraphicControl)
  private
    function GetSizerParent: IvgrSizerParent;
  protected
    procedure Paint; override;
    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;
  public
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;
    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: tvgrPointInfo);   

    property SizerParent: IvgrSizerParent read GetSizerParent;
  end;

implementation

{$R vgrRuntime.res}
{$R vgr_Controls.res}

uses
  vgr_Functions, vgr_GUIFunctions;

const
  cSheetsPanelBottomOffset = 1;

var
  FSheetsPanelBitmaps: TImageList;

/////////////////////////////////////////////////
//
// TvgrScrollBarPointInfo
//
/////////////////////////////////////////////////
constructor TvgrScrollBarPointInfo.Create(AScrollBar: TvgrScrollBar);
begin
  inherited Create;
  FScrollBar := AScrollBar;
end;

/////////////////////////////////////////////////
//
// TvgrOptionsScrollBar
//
/////////////////////////////////////////////////
constructor TvgrOptionsScrollBar.Create;
begin
  inherited Create;
  FVisible := True;
end;

procedure TvgrOptionsScrollBar.SetVisible(Value: Boolean);
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsScrollBar.Assign(Source: TPersistent);
begin
  with TvgrOptionsScrollBar(Source) do
  begin
    Self.FVisible := Visible;
  end;
  DoChange;
end;

/////////////////////////////////////////////////
//
// TvgrScrollBar
//
/////////////////////////////////////////////////
constructor TvgrScrollBar.Create(AOwner : TComponent);
begin
  inherited;
  
  FKind := sbHorizontal;
  Height := GetSystemMetrics(SM_CYHSCROLL);
  ControlStyle := [csDoubleClicks, csOpaque];
end;

destructor TvgrScrollBar.Destroy;
begin
  inherited;
end;

function TvgrScrollBar.GetScrollBarParent: IvgrScrollBarParent;
begin
  if Owner = nil then
    Result := nil
  else
    TComponentAccess(Owner).QueryInterface(svgrScrollBarParentGUID, Result);
end;

function TvgrScrollBar.GetOptions: TvgrOptionsScrollbar;
begin
  if ScrollBarParent <> nil then
    Result := ScrollBarParent.GetScrollBarOptions(Self)
  else
    Result := nil;
end;

procedure TvgrScrollBar.CreateParams(var Params: TCreateParams);
const
  Kinds: array[TScrollBarKind] of DWORD = (SBS_HORZ, SBS_VERT);
begin
  inherited CreateParams(Params);
  CreateSubClass(Params, 'SCROLLBAR');
  Params.Style := Params.Style or Kinds[FKind] or SBS_LEFTALIGN;
end;

procedure TvgrScrollBar.CreateWnd;
begin
  inherited CreateWnd;
  Position := FPosition;
  OptionsChanged;
end;

procedure TvgrScrollBar.FillScrollInfo(var AScrollInfo: TScrollInfo);
begin
  AScrollInfo.cbSize := sizeof(ScrollInfo);
  AScrollInfo.fMask := SIF_PAGE or SIF_POS or SIF_RANGE;
  GetScrollInfo(Handle, SB_CTL, AScrollInfo);
end;

procedure TvgrScrollBar.SetKind(Value: TScrollBarKind);
begin
  if FKind <> Value then
  begin
    FKind := Value;
    SetBounds(Left, Top, Height, Width);
    RecreateWnd;
  end;
end;

function TvgrScrollBar.GetPageSize: Integer;
var
  ScrollInfo: TScrollInfo;
begin
  if HandleAllocated then
  begin
    FillScrollInfo(ScrollInfo);
    Result := ScrollInfo.nPage;
  end
  else
    Result := 1;
end;

function TvgrScrollBar.GetMin: Integer;
var
  ScrollInfo: TScrollInfo;
begin
  if HandleAllocated then
  begin
    FillScrollInfo(ScrollInfo);
    Result := ScrollInfo.nMin;
  end
  else
    Result := -1;
end;

function TvgrScrollBar.GetMax: Integer;
var
  ScrollInfo: TScrollInfo;
begin
  if HandleAllocated then
  begin
    FillScrollInfo(ScrollInfo);
    Result := ScrollInfo.nMax;
  end
  else
    Result := -1;
end;

function TvgrScrollBar.GetPosition : integer;
var
  ScrollInfo: TScrollInfo;
begin
  if HandleAllocated then
  begin
    FillScrollInfo(ScrollInfo);
    FPosition := ScrollInfo.nPos;
  end;
  Result := FPosition;
end;

procedure TvgrScrollBar.SetPosition(Value : integer);
var
  ScrollInfo: TScrollInfo;
begin
  FPosition := Value;
  if HandleAllocated then
  begin
    ScrollInfo.cbSize := sizeof(ScrollInfo);
    ScrollInfo.fMask := SIF_POS;
    ScrollInfo.nPos := Value;
    SetScrollInfo(Handle,SB_CTL,ScrollInfo,true);
  end;
end;

procedure TvgrScrollBar.SetParams(Position,Min,Max,PageSize : integer);
var
  ScrollInfo: TScrollInfo;
begin
  ScrollInfo.cbSize := SizeOf(ScrollInfo);
  ScrollInfo.nMin := Min;
  ScrollInfo.nMax := Max;
  ScrollInfo.nPos := Position;
  ScrollInfo.nPage := PageSize;
  ScrollInfo.fMask := SIF_PAGE or SIF_POS or SIF_RANGE;
  SetScrollInfo(Handle,SB_CTL,ScrollInfo,true);
end;

procedure TvgrScrollBar.OptionsChanged;
begin
  if (Options = nil) or Options.Visible then
    case FKind of
      sbHorizontal: Height := GetSystemMetrics(SM_CYHSCROLL);
      sbVertical: Width := GetSystemMetrics(SM_CXVSCROLL);
    end
  else
    case FKind of
      sbHorizontal: Height := 0;
      sbVertical: Width := 0;
    end;
end;

procedure TvgrScrollBar.WMEraseBkgnd(var Msg: TWMEraseBkgnd); 
begin
  DefaultHandler(Msg);
end;

procedure TvgrScrollBar.CNCtlColorScrollBar(var Msg: TMessage);
begin
  with Msg do
    CallWindowProc(DefWndProc, Handle, Msg, WParam, LParam);
end;

procedure TvgrScrollBar.WMHScroll(var Msg : TWMHScroll);
begin
  if Assigned(FOnScrollMessage) then
    FOnScrollMessage(Self,Msg);
end;

procedure TvgrScrollBar.WMVScroll(var Msg : TWMVScroll);
begin
  if Assigned(FOnScrollMessage) then
    FOnScrollMessage(Self,Msg);
end;

procedure TvgrScrollBar.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  if ScrollBarParent <> nil then
    ScrollBarParent.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrScrollBar.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
begin
  APointInfo := TvgrScrollBarPointInfo.Create(Self);
end;

procedure TvgrScrollBar.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); 
begin
  if ScrollBarParent <> nil then
    ScrollBarParent.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrScrollBar.DragDrop(Source: TObject; X, Y: Integer);
begin
  if ScrollBarParent <> nil then
    ScrollBarParent.DoDragDrop(Self, Source, X, Y);
end;

/////////////////////////////////////////////////
//
// TvgrSheetsPanelRect
//
/////////////////////////////////////////////////
constructor TvgrSheetsPanelRect.Create(AParentControl: TCustomControl);
begin
  inherited Create;
  FParentControl := AParentControl;
end;

destructor TvgrSheetsPanelRect.Destroy;
begin
  if (Parent <> nil) and (Parent.CurRect = Self) then
    Parent.CurRect := nil;
  inherited;
end;

function TvgrSheetsPanelRect.GetParent: IvgrControlRectParent;
begin
  if ParentControl <> nil then
    TCustomControlAccess(ParentControl).QueryInterface(svgrControlRectParent, Result)
  else
    Result := nil
end;

procedure TvgrSheetsPanelRect.SetMouseOver(Value: Boolean);
begin
  FMouseOver := Value;
  InvalidateRect(ParentControl.Handle, @FRect, False);
end;

procedure TvgrSheetsPanelRect.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
begin
  FRect := Rect(ALeft, ATop, ALeft + AWidth, ATop + AHeight);
end;

/////////////////////////////////////////////////
//
// TvgrSheetsPanelButtonRect
//
/////////////////////////////////////////////////
constructor TvgrSheetsPanelButtonRect.CreateButton(AParentControl: TCustomControl; AImageList: TImageList; AImageIndex: Integer; AOnClick: TNotifyEvent);
begin
  inherited Create(AParentControl);
  FImageList := AImageList;
  FImageIndex := AImageIndex;
  FOnClick := AOnClick;
end;

procedure TvgrSheetsPanelButtonRect.Paint(ACanvas: TCanvas);
var
  ARect: TRect;
begin
  ARect := FRect;
  if FMouseOver and Parent.GetButtonEnabled(ImageIndex) then
  begin
    ACanvas.Pen.Color := GetHighlightColor(clBtnShadow);
    ACanvas.Brush.Color := GetHighlightColor(clBtnFace);
  end
  else
  begin
    ACanvas.Pen.Color := clBtnFace;
    ACanvas.Brush.Color := clBtnFace;
  end;
  ACanvas.Rectangle(ARect);
  InflateRect(ARect, -1, -1);
  FImageList.Draw(ACanvas,
                  ARect.Left + (ARect.Right - ARect.Left - FSheetsPanelBitmaps.Width) div 2,
                  ARect.Top + (ARect.Bottom - ARect.Top - FSheetsPanelBitmaps.Height) div 2,
                  ImageIndex,
                  Parent.GetButtonEnabled(ImageIndex));
  ExcludeClipRect(ACanvas, FRect);
end;

procedure TvgrSheetsPanelButtonRect.Click;
begin
  if Assigned(FOnClick) then
    FOnClick(Self);
end;

/////////////////////////////////////////////////
//
// TvgrSheetsPanelCaptionRect
//
/////////////////////////////////////////////////
function TvgrSheetsPanelCaptionRect.GetSheetsPanel: TvgrSheetsPanel;
begin
  Result := TvgrSheetsPanel(ParentControl);
end;

function TvgrSheetsPanelCaptionRect.GetTitle: string;
begin
  if Worksheet = nil then
    Result := ''
  else
    Result := Worksheet.Title;
end;

procedure TvgrSheetsPanelCaptionRect.Paint(ACanvas: TCanvas);
var
  ARect: TRect;
  AOffs: Integer;
  ARgn: HRGN;
  ARgnPoints: Array [0..3] of TPoint;
begin
  // Top line
  if Worksheet <> SheetsPanel.ActiveWorksheet then
  begin
    ACanvas.Brush.Color := clBtnShadow;
    ACanvas.FillRect(Rect(FRect.Left, FRect.Top, FRect.Right, FRect.Top + 1));
    ExcludeClipRect(ACanvas, Rect(FRect.Left, FRect.Top, FRect.Right, FRect.Top + 1));
  end;

  ACanvas.Font.Assign(SheetsPanel.FFont);
  if Worksheet = SheetsPanel.ActiveWorksheet then
  begin
    ACanvas.Font.Style := [fsBold];
    ACanvas.Brush.Color := clBtnHighlight;
  end
  else
  begin
    ACanvas.Font.Style := [];
    ACanvas.Brush.Color := clBtnFace;
  end;

  // Draw Title and exclude them Rect
  with ARect, ACanvas.TextExtent(Title) do
  begin
    Left := FRect.Left + (FRect.Right - FRect.Left - cx) div 2;
    Top := FRect.Top + (FRect.Bottom - FRect.Top - cy) div 2;
    Right := Left + cx;
    Bottom := Top + cy;
  end;
  ACanvas.TextRect(ARect, ARect.Left, ARect.Top, Title);
  ExcludeClipRect(ACanvas, ARect);

  AOffs := (FRect.Bottom - FRect.Top) div 2;

  // Fill regions

  // 1. BtnFace or BtnHighlight
  ARgnPoints[0] := Point(FRect.Left + 2, FRect.Top);
  ARgnPoints[1] := Point(FRect.Left + AOffs + 2, FRect.Bottom - 1);
  ARgnPoints[2] := Point(FRect.Right - AOffs - 2, FRect.Bottom - 1);
  ARgnPoints[3] := Point(FRect.Right - 2, FRect.Top);
  ARgn := CreatePolygonRgn(ARgnPoints, 4, ALTERNATE);
  FillRegion(ACanvas, ARgn);
  ExcludeClipRegion(ACanvas, ARgn);
  DeleteObject(ARgn);

  // 2. clBtnHighlight
  ARgnPoints[0].X := ARgnPoints[0].X - 1;
  ARgnPoints[1].X := ARgnPoints[1].X - 1;
  ARgn := CreatePolygonRgn(ARgnPoints, 4, ALTERNATE);
  ACanvas.Brush.Color := clBtnHighlight;
  FillRegion(ACanvas, ARgn);
  ExcludeClipRegion(ACanvas, ARgn);
  DeleteObject(ARgn);

  // 3. clBtnShadow
  ARgnPoints[2].X := ARgnPoints[2].X + 1;
  ARgnPoints[3].X := ARgnPoints[3].X + 1;
  ARgn := CreatePolygonRgn(ARgnPoints, 4, ALTERNATE);
  ACanvas.Brush.Color := clBtnShadow;
  FillRegion(ACanvas, ARgn);
  ExcludeClipRegion(ACanvas, ARgn);
  DeleteObject(ARgn);

  // 4. clWindowFrame
  ARgnPoints[0].X := ARgnPoints[0].X - 1;
  ARgnPoints[1].X := ARgnPoints[1].X - 1;
  ARgnPoints[2].X := ARgnPoints[2].X + 1;
  ARgnPoints[3].X := ARgnPoints[3].X + 1;
  ARgn := CreatePolygonRgn(ARgnPoints, 4, ALTERNATE);
  ACanvas.Brush.Color := clWindowFrame;
  FillRegion(ACanvas, ARgn);
  ExcludeClipRegion(ACanvas, ARgn);
  DeleteObject(ARgn);

  // 5. Bottom line
  ACanvas.Brush.Color := clBtnShadow;
  ARect := Rect(FRect.Left + AOffs, FRect.Bottom - 1, FRect.Right - AOffs, FRect.Bottom);
  ACanvas.FillRect(ARect);
  ExcludeClipRect(ACanvas, ARect);
end;

function TvgrSheetsPanelCaptionRect.GetCaptionWidth: Integer;
begin
  SheetsPanel.FFont.Style := [fsBold];
  with SheetsPanel.BoundsRect do
    Result := GetStringWidth(Title, SheetsPanel.FFont) + (Bottom - Top) * 2;
end;

procedure TvgrSheetsPanelCaptionRect.Click;
begin
  SheetsPanel.ActiveWorksheetIndex := Worksheet.IndexInWorkbook;
end;

/////////////////////////////////////////////////
//
// TvgrSheetsPanelPointInfo
//
/////////////////////////////////////////////////
constructor TvgrSheetsPanelPointInfo.Create(APlace: TvgrSheetsPanelPointInfoPlace; ASheetIndex: Integer);
begin
  inherited Create;
  FPlace := APlace;
  FSheetIndex := ASheetIndex;
end;

/////////////////////////////////////////////////
//
// TvgrSheetsPanel
//
/////////////////////////////////////////////////
constructor TvgrSheetsPanel.Create(AOwner: TComponent);

  procedure CreateButton(var AButton: TvgrSheetsPanelButtonRect; AImageIndex: Integer);
  begin
    AButton := TvgrSheetsPanelButtonRect.CreateButton(Self, FSheetsPanelBitmaps, AImageIndex, OnButtonClick);
  end;
  
begin
  inherited;
  FFont := TFont.Create;
  FFont.Name := sDefSheetsPanelFontName;
  FFont.Size := sDefSheetsPanelFontSize;
  Color := clBtnFace;
  FFirstSheetIndex := 0;
  FSheetCaptions := TList.Create;
  CreateButton(FFirstButton, 0);
  CreateButton(FPriorButton, 1);
  CreateButton(FNextButton, 2);
  CreateButton(FLastButton, 3);
end;

destructor TvgrSheetsPanel.Destroy;
begin
  FFirstButton.Free;
  FPriorButton.Free;
  FNextButton.Free;
  FLastButton.Free;
  FSheetCaptions.Free;
  FFont.Free;
  inherited;
end;

procedure TvgrSheetsPanel.OnButtonClick(Sender: TObject);
begin
  if Workbook = nil then exit;
  if Sender = FFirstButton then
  begin
    FFirstSheetIndex := 0;
  end
  else
    if Sender = FPriorButton then
    begin
      if FFirstSheetIndex <= 0 then exit;
      Dec(FFirstSheetIndex);
    end
    else
      if Sender = FNextButton then
      begin
        if FFirstSheetIndex >= SheetCaptionCount - 1 then exit;
        Inc(FFirstSheetIndex);
      end
      else
        if Sender = FLastButton then
        begin
          FFirstSheetIndex := SheetCaptionCount - 1;
        end;
  UpdateSheets;
end;

procedure TvgrSheetsPanel.ClearSheetCaptions;
var
  I: Integer;
begin
  for I := 0 to SheetCaptionCount - 1 do
    SheetCaptions[I].Free;
  FSheetCaptions.Clear;
end;

function TvgrSheetsPanel.GetSheetsPanelParent: IvgrSheetsPanelParent;
begin
  if Owner = nil then
    Result := nil
  else
    TComponentAccess(Owner).QueryInterface(svgrSheetsPanelParentGUID, Result);
end;

function TvgrSheetsPanel.GetWorkbook: TvgrWorkbook;
begin
  if SheetsPanelParent = nil then
    Result := nil
  else
    Result := SheetsPanelParent.Workbook;
end;

function TvgrSheetsPanel.GetActiveWorksheet: TvgrWorksheet;
begin
  if SheetsPanelParent = nil then
    Result := nil
  else
    Result := SheetsPanelParent.ActiveWorksheet;
end;

function TvgrSheetsPanel.GetActiveWorksheetIndex: Integer;
begin
  if SheetsPanelParent = nil then
    Result := -1
  else
    Result := SheetsPanelParent.ActiveWorksheetIndex;
end;

procedure TvgrSheetsPanel.SetActiveWorksheetIndex(Value: Integer);
begin
  if SheetsPanelParent <> nil then
    SheetsPanelParent.ActiveWorksheetIndex := Value;
end;

function TvgrSheetsPanel.GetSheetCaption(Index: Integer): TvgrSheetsPanelCaptionRect;
begin
  Result := TvgrSheetsPanelCaptionRect(FSheetCaptions[Index]);
end;

function TvgrSheetsPanel.GetSheetCaptionCount: Integer;
begin
  Result := FSheetCaptions.Count;
end;

procedure TvgrSheetsPanel.Resize;
var
  ATop, AHeight: Integer;
begin
  ATop := ClientRect.Top + 1;
  AHeight := ClientHeight - 2;
  FFirstButton.SetBounds(ClientRect.Left, ATop, AHeight, AHeight);
  FPriorButton.SetBounds(FFirstButton.FRect.Right, ATop, AHeight, AHeight);
  FNextButton.SetBounds(FPriorButton.FRect.Right, ATop, AHeight, AHeight);
  FLastButton.SetBounds(FNextButton.FRect.Right, ATop, AHeight, AHeight);

  AlignSheetCaptions;
end;

procedure TvgrSheetsPanel.CMMouseLeave(var Msg: TMessage);
begin
  if FCurRect <> nil then
  begin
    FCurRect.MouseOver := False;
    FCurRect := nil;
  end;
end;

procedure TvgrSheetsPanel.WMEraseBkgnd(var Msg : TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrSheetsPanel.Paint;
var
  I: Integer;
  ACaption: TvgrSheetsPanelCaptionRect;
begin
  // 1. Draw buttons
  FFirstButton.Paint(Canvas);
  FPriorButton.Paint(Canvas);
  FNextButton.Paint(Canvas);
  FLastButton.Paint(Canvas);
  
  // 2. Fill back under buttons
  Canvas.Brush.Color := clBtnShadow;
  Canvas.FillRect(Rect(0, 0, ClientRect.Right, 1));
  Canvas.FillRect(Rect(FLastButton.FRect.Right, 1, FLastButton.FRect.Right + 1, ClientRect.Bottom));
  Canvas.Brush.Color := clBtnFace;
  Canvas.FillRect(Rect(0, 1, FLastButton.FRect.Right, ClientRect.Bottom));
  ExcludeClipRect(Canvas, Rect(0, 0, FLastButton.FRect.Right + 1, ClientRect.Bottom));

  // 3. draw sheet captions
  ACaption := FindCaptionByWorksheet(ActiveWorksheet);
  if ACaption <> nil then
    ACaption.Paint(Canvas);
  for I := 0 to SheetCaptionCount - 1 do
    if SheetCaptions[I] <> ACaption then
      SheetCaptions[I].Paint(Canvas);
      
  // 4. Fill back under sheet captions
  Canvas.Brush.Color := clBtnFace;
  Canvas.FillRect(Rect(FLastButton.FRect.Right + 1, 1, ClientRect.Right, ClientRect.Bottom));
end;

procedure TvgrSheetsPanel.UpdateSheets;
var
  I: Integer;
  ASheetCaption: TvgrSheetsPanelCaptionRect;
begin
  if Workbook = nil then
  begin
    ClearSheetCaptions;
  end
  else
  begin
    if Workbook.WorksheetsCount < FSheetCaptions.Count then
    begin
      for I := Workbook.WorksheetsCount to SheetCaptionCount - 1 do
        SheetCaptions[I].Free;
      FSheetCaptions.Count := Workbook.WorksheetsCount;
    end
    else
      if Workbook.WorksheetsCount > FSheetCaptions.Count then
        for I := SheetCaptionCount to Workbook.WorksheetsCount - 1 do
        begin
          ASheetCaption := TvgrSheetsPanelCaptionRect.Create(Self);
          FSheetCaptions.Add(ASheetCaption);
        end;
    for I := 0 to Workbook.WorksheetsCount - 1 do
      SheetCaptions[I].Worksheet := Workbook.Worksheets[I];
  end;
  AlignSheetCaptions;
  Repaint;
end;

procedure TvgrSheetsPanel.AlignSheetCaptions;
var
  I, ALeft, ACaptionWidth, AHeight: Integer;
begin
  AHeight := BoundsRect.Bottom - BoundsRect.Top - cSheetsPanelBottomOffset;

  ALeft := FLastButton.FRect.Right + 1;
  for I := Min(FFirstSheetIndex - 1, SheetCaptionCount - 1) downto 0 do
    with SheetCaptions[I] do
    begin
      ACaptionWidth := GetCaptionWidth;
      SetBounds(ALeft - ACaptionWidth + AHeight div 2, 0, ACaptionWidth, AHeight);
      ALeft := ALeft - ACaptionWidth + AHeight div 2;
    end;
  
  ALeft := FLastButton.FRect.Right + 1;
  for I := Max(FFirstSheetIndex, 0) to SheetCaptionCount - 1 do
    with SheetCaptions[I] do
    begin
      ACaptionWidth := GetCaptionWidth;
      SetBounds(ALeft, 0, ACaptionWidth, AHeight);
      ALeft := ALeft + ACaptionWidth - AHeight div 2;
    end;
end;

procedure TvgrSheetsPanel.BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
begin
end;

procedure TvgrSheetsPanel.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcNewWorksheet, vgrwcDeleteWorksheet, vgrwcChangeWorksheet:
      UpdateSheets;
    vgrwcUpdateAll:
      UpdateSheets;
  end;
end;

function TvgrSheetsPanel.FindCaptionByWorksheet(AWorksheet: TvgrWorksheet): TvgrSheetsPanelCaptionRect;
var
  I: Integer;
begin
  for I := 0 to SheetCaptionCount - 1 do
    if SheetCaptions[I].Worksheet = ActiveWorksheet then
    begin
      Result := SheetCaptions[I];
      exit;
    end;
  Result := nil;
end;

function TvgrSheetsPanel.GetButtonEnabled(AButtonIndex: Integer): Boolean;
begin
  Result := Workbook <> nil;
end;

function TvgrSheetsPanel.GetCurRect: TvgrSheetsPanelRect;
begin
  Result := FCurRect;
end;

procedure TvgrSheetsPanel.SetCurRect(Value: TvgrSheetsPanelRect);
begin
  FCurRect := Value;
end;

function TvgrSheetsPanel.GetPanelRectAt(X, Y: Integer): TvgrSheetsPanelRect;
var
  APoint: TPoint;
  I: Integer;
begin
  Result := nil;
  APoint := Point(X, Y);
  if PtInRect(FFirstButton.FRect, APoint) then
    Result := FFirstButton
  else
    if PtInRect(FPriorButton.FRect, APoint) then
      Result := FPriorButton
    else
      if PtInRect(FNextButton.FRect, APoint) then
        Result := FNextButton
      else
        if PtInRect(FLastButton.FRect, APoint) then
          Result := FLastButton
        else
        begin
          for I := Max(0, FFirstSheetIndex - 1) to SheetCaptionCount - 1 do
            if PtInRect(SheetCaptions[I].FRect, APoint) then
            begin
              Result := SheetCaptions[I];
              exit;
            end;
        end;
end;

procedure TvgrSheetsPanel.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APanelRect: TvgrSheetsPanelRect;
  AExecuteDefault: Boolean;
  APointInfo: TvgrPointInfo;
begin
  AExecuteDefault := True;
  if SheetsPanelParent <> nil then
    SheetsPanelParent.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    APanelRect := GetPanelRectAt(X, Y);
    if (APanelRect <> nil) and (Button = mbLeft) then
      APanelRect.Click
    else
      if (Button = mbRight) and ((APanelRect = nil) or (APanelRect is TvgrSheetsPanelCaptionRect)) then
      begin
        GetPointInfoAt(Point(X, Y), APointInfo);
        SheetsPanelParent.ShowPopupMenu(ClientToScreen(Point(X, Y)), APointInfo);
        APointInfo.Free;
      end;
  end;
end;

procedure TvgrSheetsPanel.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  APanelRect: TvgrSheetsPanelRect;
begin
  APanelRect := GetPanelRectAt(X, Y);
  if APanelRect <> FCurRect then
  begin
    if FCurRect <> nil then
      FCurRect.MouseOver := False;
    FCurRect := APanelRect;
    if FCurRect <> nil then
      FCurRect.MouseOver := True;
  end;
  if SheetsPanelParent <> nil then
    SheetsPanelParent.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrSheetsPanel.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  if SheetsPanelParent <> nil then
    SheetsPanelParent.DoMouseUp(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrSheetsPanel.DblClick;
var
  AExecuteDefault: Boolean;
begin
  if SheetsPanelParent <> nil then
    SheetsPanelParent.DoDblClick(Self, AExecuteDefault);
end;

procedure TvgrSheetsPanel.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
var
  ARect: TvgrSheetsPanelRect;
  APlace: TvgrSheetsPanelPointInfoPlace;
  ASheetIndex: Integer;
begin
  ARect := GetPanelRectAt(APoint.X, APoint.Y);
  ASheetIndex := -1;
  if ARect = FFirstButton then
    APlace := vgrsipFirstButton
  else
    if ARect = FPriorButton then
      APlace := vgrsipPriorButton
    else
      if ARect = FNextButton then
        APlace := vgrsipNextButton
      else
        if ARect = FLastButton then
          APlace := vgrsipLastButton
        else
          if ARect is TvgrSheetsPanelCaptionRect then
          begin
            ASheetIndex := TvgrSheetsPanelCaptionRect(ARect).Worksheet.IndexInWorkbook;
            APlace := vgrSheetCaption;
          end
          else
            APlace := vgrsipNone;
  APointInfo := TvgrSheetsPanelPointInfo.Create(APlace, ASheetIndex);
end;

procedure TvgrSheetsPanel.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  if SheetsPanelParent <> nil then
    SheetsPanelParent.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrSheetsPanel.DragDrop(Source: TObject; X, Y: Integer);
begin
  if SheetsPanelParent <> nil then
    SheetsPanelParent.DoDragDrop(Self, Source, X, Y);
end;

/////////////////////////////////////////////////
//
// TvgrSheetsPanelResizer
//
/////////////////////////////////////////////////
constructor TvgrSheetsPanelResizer.Create(AOwner: TComponent);
begin
  inherited;
  Cursor := crHSplit;
end;

function TvgrSheetsPanelResizer.GetSheetsPanelResizerParent: IvgrSheetsPanelResizerParent;
begin
  if Owner = nil then
    Result := nil
  else
    TComponentAccess(Owner).QueryInterface(svgrSheetsPanelResizerParentGUID, Result)
end;

procedure TvgrSheetsPanelResizer.Paint;
var
  ARect: TRect;
begin
  ARect := ClientRect;
  with ARect do
  begin
    Canvas.Brush.Color := clBtnFace;
    Canvas.FillRect(Rect(Left, Top, Left + 1, Bottom));
    Canvas.FillRect(Rect(Left, Top, Right, Top + 1));
    Canvas.Brush.Color := clWindowFrame;
    Canvas.FillRect(Rect(Right - 1, Top, Right, Bottom));
    Canvas.Brush.Color := clBtnShadow;
    Canvas.FillRect(Rect(Left, Bottom - 1, Right - 1, Bottom));
    InflateRect(ARect, -1, -1);
    Canvas.Brush.Color := clBtnShadow;
    Canvas.FillRect(Rect(Right - 1, Top, Right, Bottom));
    Canvas.Brush.Color := clBtnHighlight;
    Canvas.FillRect(Rect(Left, Top, Left + 1, Bottom));
    Canvas.FillRect(Rect(Left, Top, Right - 1, Top + 1));
    ARect.Left := ARect.Left + 1;
    ARect.Top := ARect.Top + 1;
    ARect.Right := ARect.Right - 1;
    Canvas.Brush.Color := clBtnFace;
    Canvas.FillRect(ARect);
  end;
end;

procedure TvgrSheetsPanelResizer.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  if SheetsPanelResizerParent <> nil then
  begin
    if FResizeMode then
      SheetsPanelResizerParent.DoResize(X - FStartMousePoint.X);
    SheetsPanelResizerParent.DoMouseMove(Self, Shift, X, Y);
  end;
end;

procedure TvgrSheetsPanelResizer.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  AExecuteDefault := True;
  if SheetsPanelResizerParent <> nil then
    SheetsPanelResizerParent.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    FResizeMode := True;
    FStartMousePoint := Point(X, Y);
  end;
end;

procedure TvgrSheetsPanelResizer.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  AExecuteDefault := True;
  if SheetsPanelResizerParent <> nil then
    SheetsPanelResizerParent.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
    FResizeMode := False;
end;

procedure TvgrSheetsPanelResizer.DblClick;
var
  AExecuteDefault: Boolean;
begin
  AExecuteDefault := True;
  if SheetsPanelResizerParent <> nil then
    SheetsPanelResizerParent.DoDblClick(Self, AExecuteDefault);
end;

procedure TvgrSheetsPanelResizer.GetPointInfoAt(const APoint: TPoint; var APointInfo: tvgrPointInfo);
begin
  APointInfo := TvgrSheetsPanelResizerPointInfo.Create;
end;

procedure TvgrSheetsPanelResizer.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  if SheetsPanelResizerParent <> nil then
    SheetsPanelResizerParent.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrSheetsPanelResizer.DragDrop(Source: TObject; X, Y: Integer);
begin
  if SheetsPanelResizerParent <> nil then
    SheetsPanelResizerParent.DoDragDrop(Self, Source, X, Y);
end;

/////////////////////////////////////////////////
//
// TvgrSizer
//
/////////////////////////////////////////////////
function TvgrSizer.GetSizerParent: IvgrSizerParent;
begin
  if Owner = nil then
    Result := nil
  else
    TComponentAccess(Owner).QueryInterface(svgrSizerParent, Result)
end;

procedure TvgrSizer.Paint;
begin
  Canvas.Brush.Color := clBtnFace;
  Canvas.FillRect(ClientRect);
{
  DrawFrameControl(Canvas.Handle,
                   ClientRect,
                   DFC_SCROLL,
                   DFCS_SCROLLSIZEGRIP);
}
end;

procedure TvgrSizer.MouseMove(Shift: TShiftState; X,Y: Integer);
begin
  if SizerParent <> nil then
    SizerParent.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrSizer.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  if SizerParent <> nil then
    SizerParent.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrSizer.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  if SizerParent <> nil then
    SizerParent.DoMouseUp(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrSizer.DblClick;
var
  AExecuteDefault: Boolean;
begin
  if SizerParent <> nil then
    SizerParent.DoDblClick(Self, AExecuteDefault);
end;

procedure TvgrSizer.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  if SizerParent <> nil then
    SizerParent.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrSizer.DragDrop(Source: TObject; X, Y: Integer);
begin
  if SizerParent <> nil then
    SizerParent.DoDragDrop(Self, Source, X, Y);
end;

procedure TvgrSizer.GetPointInfoAt(const APoint: TPoint; var APointInfo: tvgrPointInfo);
begin
  APointInfo := TvgrSizerPointInfo.Create;
end;

initialization

  FSheetsPanelBitmaps := TImageList.Create(nil);
  FSheetsPanelBitmaps.Width := 5;
  FSheetsPanelBitmaps.Height := 5;
  FSheetsPanelBitmaps.ResourceLoad(rtBitmap, svgrSheetsPanelBitmaps, clFuchsia);

finalization

  FSheetsPanelBitmaps.Free;

end.
