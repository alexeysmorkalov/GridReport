{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains TvgrWorkbookPreview control, that can be used for previewing the workbook contents before printing.}
unit vgr_WorkbookPreview;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  SysUtils, Classes, Controls, Forms, messages, {$IFDEF VTK_D6_OR_D7}Types, {$ENDIF} Windows,
  math, Graphics, stdctrls, ImgList, menus,

  vgr_Consts, vgr_DataStorage, vgr_CommonClasses, vgr_PageMaker, vgr_PageProperties, vgr_Functions,
  vgr_Controls, vgr_WorkbookPainter, vgr_Keyboard;

//{$DEFINE VGR_UPDATEIMAGE_TEST}

type

  TvgrWorkbookPreview = class;
  TvgrWPPreview = class;

{Specifies the type of the current user action
Items:
  vgrapaNone - Default (no action mode)
  vgrapaHand - Hand tool(pan page)
  vgrapaZoomIn - Zoom In Tool
  vgrapaZoomOut - Zoom Out Tool
}
    TvgrPreviewActionMode = (vgrpamNone, vgrpamHand,vgrpamZoomIn,vgrpamZoomOut);

{Specifies the type of the place within TvgrWPNavigator control:
Items:
  vgrnbNone - empty position.
  vgrnbFirstPage - within the "First page" button.
  vgrnbPriorPage - within the "Prior page" button.
  vgrnbNextPage - within the "Next page" button.
  vgrnbLastPage - within the "Last page" button.}
  TvgrWPNavigatorButton = (vgrnbNone, vgrnbFirstPage, vgrnbPriorPage, vgrnbNextPage, vgrnbLastPage);

  /////////////////////////////////////////////////
  //
  // TvgrWPImagesCache
  //
  /////////////////////////////////////////////////
{Provides cache for images of the pages.}
  TvgrWPImagesCache = class
  private
{Array for stored page's numbers}
    FNumbers:  array of Integer;
{Array for stored page's images}
    FImages: array of TMetafile;
{Stack Pointer}
    FPos: Integer;
{Clear all stored images in cache}
    procedure ClearCache;
{Gets page's image from cache.
Parameters:
  ANumber - page's number
Return value:
  return page's image if image exists, return nil
if image not exists.}
    function GetImage(ANumber: Integer):TMetafile;
{Sets image of page to a cache
Parameters:
  ANumber - page's number
  AImage - page's image
Return value:
  AImage parameter}
    function SetImage(ANumber: Integer;AImage: TMetafile):TMetafile;    
  public
{Creates an instance of the TvgrWPImagesCache class.}
    constructor Create;
{Frees an instance of the TvgrWPImagesCache class.}
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWPMarginsResizer
  //
  /////////////////////////////////////////////////
{Provides information and actions for wysiwyg margins resizing, hand tool, Zoom In and Zoom Out tools .}
  TvgrWPMarginsResizer = class
  private
    FActivePageRect: TRect;
{point's coordinate in which TvgrWPMarginsResizer begins resize margin(in client coordinates).
Also used in Zoom and hand tools}
    FBegPoint: TPoint;
{point's coordinate where mouse down occures.}
    FBegPointMoseDown: TPoint;
{point's coordinate in which TvgrWPMarginsResizer begins resize margin(in screen coordinates)}
    FBegPointToScreen: TPoint;
    FCautionStep: Real;
{Current action for margins resizer.
Items:
 -1 - first init
 0 - no action
 1 - select current margin, begin resizing margin, activate hint,
     continue resizing margin, change position of hint
 2 - set new margin size, hide hint.
 4 - hand tool
 5 - Zoom In tool
 6 - Zoom Out tool
}
    FCurrentAction: Integer;
{Current margin to resize.
Items:
  0 - left margin
  1 - right margin
  2 - top margin
  3 - bottom margin
  4 - header
  5 = footer}
    FCurrentMargin: Integer;
    FSpotWidth: Integer;
{Hint window}
    FHintWindow: THintWindow;
{New size of current margin in current measurement system}
    FMarginNewSize: Real;
{Size of current margin's offset in pixels}
    FMarginOffsetSizePix: Integer;
{Size of margins in current measurement system}
    FLeft,FRight,FTop,FBottom: Real;
{Size of header and footer in current measurement system}
    FHeader, FFooter: Real; 
{Coordinates for margins lines}
    FCLeft,FCRight,FCTop,FCBottom,FCBLeft,FCBRight,FCBTop,FCBBottom: Integer;
{Y coordinates for header and footer lines}
    FCHeader, FCFooter: Integer;
{mesurment system}
    FMeasurementUnits: string;
    FMesUnits: TvgrUnits;
{array of page's rects}
    FPagesRects: array of TRect;
{array of page's rects}
    FPagesRectsNumbers: array of Integer;
    FParentWPPreview: TvgrWPPreview;
{Previous rect of margin's line - need for clear. Also used in Zoom In ad Zoom out tools}
    FPrevMarginRect: TRect;
{Rects of margins's spots}
    FSpotLeftMargin, FSpotRightMargin, FSpotTopMargin, FSpotBottomMargin: TRect;
{Rects of header and footer spots}
    FSpotHeader, FSpotFooter: TRect;
{Indicate when need scroll to NeedElaborateScroll_Rect rectangle.}
    NeedElaborateScroll: Boolean;
{Rectangle to scroll}
    NeedElaborateScroll_Rect: TRect;
{Rectangle of active page}
    NeedElaborateScroll_ActivePageRect: TRect;
{Creates instance of TvgrWPMarginsResizer
Parameters:
  AParentWPPreview - parent control TvgrWPPreview(not nil!).}
    procedure SetCurrentCursor;
{Clears stored pages rects}
    procedure ClearPagesSpots;
{Draws margin's lines with current offset if needed}
    procedure DrawMarginsResizerLines;
    procedure DoMouseUp(APoint, APointToScreen: TPoint);
    procedure DoMouseDown(APoint, APointToScreen: TPoint);
    procedure DoMouseMove(APoint, APointToScreen: TPoint; Shift: TShiftState);
{Procedure shows hint window for margin(header, footer) AMargin with text AHint in point APoint}
    procedure ShowHintForMargin(AHint: String; AMargin: Integer; APoint: TPoint);
    procedure SetPageSpot(APageRect: TRect; APageNamber: Integer);
    procedure SetSpots(ALeft, ARight, ATop, ABottom, ABLeft,
                                        ABRight, ABTop, ABBottom, AHeader, AFooter: Integer);
    procedure GetMarginsSize;
  public
    constructor Create(AParentWPPreview: TvgrWPPreview);
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWPNavigatorPointInfo
  //
  /////////////////////////////////////////////////
{Provides information about spot within area of the TvgrWPNavigator.
Object of this type returns by method TvgrWorkbookPreview.GetPointInfoAt,
if mouse within area of the TvgrWPNavigator.}
  TvgrWPNavigatorPointInfo = class(TvgrPointInfo)
  private
    FButton: TvgrWPNavigatorButton;
  public
{Creates a instance of the TvgrWPNavigatorPointInfo class.
Parameters:
  AButton - specifies type of the place within TvgrWPNavigator.}
    constructor Create(AButton: TvgrWPNavigatorButton);
{Specifies type of the place within TvgrWPNavigator.}
    property Button: TvgrWPNavigatorButton read FButton;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWPNavigator
  //
  /////////////////////////////////////////////////
{TvgrWPNavigator describes, control which allow to navigate between pages.}
  TvgrWPNavigator = class(TCustomControl, IvgrControlRectParent)
  private
    FCurRect: TvgrSheetsPanelRect;
    FPriorPage: TvgrSheetsPanelButtonRect;
    FNextPage: TvgrSheetsPanelButtonRect;
    FFirstPage: TvgrSheetsPanelButtonRect;
    FLastPage: TvgrSheetsPanelButtonRect;
    procedure OnButtonClick(Sender: TObject);
    function GetWorkbookPreview: TvgrWorkbookPreview;
    function GetWorkbook: TvgrWorkbook;
    function GetActivePage: Integer;
    procedure SetActivePage(Value: Integer);
    function GetPageCount: Integer;
    procedure WMEraseBkgnd(var Msg : TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CMMouseLeave(var Msg: TMessage); message CM_MOUSELEAVE;
  protected
    // IvgrControlRectParent
    function GetButtonEnabled(AButtonIndex: Integer): Boolean;
    function GetCurRect: TvgrSheetsPanelRect;
    procedure SetCurRect(Value: TvgrSheetsPanelRect);

    procedure Resize; override;
    procedure Paint; override;
    function GetPanelRectAt(X, Y: Integer): TvgrSheetsPanelRect;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;
  public
{Creates a instance of the TvgrWPNavigator class.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees a instance of the TvgrWPNavigator class.}
    destructor Destroy; override;
{Returns information about spot within control area.
Parameters:
  APoint - a point within control, relative to top left corner of control.
  APointInfo - object of the class, inherited from TvgrPointInfo class, that describes a point within control.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
{Returns TvgrWorkbookPreview object, which owns this object.}
    property WorkbookPreview: TvgrWorkbookPreview read GetWorkbookPreview;
{Returns TvgrWorkbook object, which is linked to the TvgrWorkbookPreview object.}
    property Workbook: TvgrWorkbook read GetWorkbook;
{Sets or gets the index of the active page.}
    property ActivePage: Integer read GetActivePage write SetActivePage;
{Returns the amount of the pages.}
    property PageCount: Integer read GetPageCount;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWPPreviewPointInfo
  //
  /////////////////////////////////////////////////
{Provides information about spot within area of the TvgrWPPreview.
Object of this type are returned by the TvgrWorkbookPreview.GetPointInfoAt method,
if mouse within area of the TvgrWPPreview.}
  TvgrWPPreviewPointInfo = class(TvgrPointInfo)
  end;

{Specifies the mode of the page zooming.
Items:
  vgrpsmPercent - the page are scaled by specified percent.
  vgrpsmPageWidth - the page must be fitted in the available width.
  vgrpsmWholePage - the page should be shown entirely.}
  TvgrPreviewScaleMode = (vgrpsmPercent, vgrpsmPageWidth, vgrpsmWholePage, vgrpsmTwoPages, vgrpsmMultiPages);
  /////////////////////////////////////////////////
  //
  // TvgrWPPreview
  //
  /////////////////////////////////////////////////
{TvgrWPPreview control, represents the client area of TvgrWorkbookPreview control.}
  TvgrWPPreview = class(TCustomControl, IvgrWorkbookPainterOwner)
  private
    FActivePage: Integer;
{Cache for page's images}
    FCache: TvgrWPImagesCache;
    FFullPageRect: TRect;
{Margins resizer}
    FMarginsResizer: TvgrWPMarginsResizer;
    FPainter: TvgrWorkbookPainter;
{Space in pixels between pages on horizontal line of preview}
    FPageHSpacing: Integer;
{Space in pixels between pages on vertical line of preview}
    FPageVSpacing: Integer;
{Number of columns of pages in preview}
    FPageColCount: Integer;
{Number of rows of pages in preview}
    FPageRowCount: Integer;
    FPreviewAreaWidth: Integer;
    FPreviewAreaHeight: Integer;
    FPageRect: TRect;
    FScaleMode: TvgrPreviewScaleMode;
    FScalePercent: Double;


    function GetWorksheet: TvgrWorksheet;
    function GetPageMaker: TvgrPageMaker;
    function GetWorkbookPreview: TvgrWorkbookPreview;
    function GetHorzScrollBar: TvgrScrollBar;
    function GetVertScrollBar: TvgrScrollBar;
    function GetActivePage: Integer;
    procedure ScrollToActivePage;
{Sets active page for preview}
    procedure SetActivePage(Value: Integer);
    function GetPageCount: Integer;
    procedure SetScaleMode(Value: TvgrPreviewScaleMode);
    function GetScalePercent: Double;
    procedure SetScalePercent(Value: Double);
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMGetDlgCode(var Msg: TWMGetDlgCode); message WM_GETDLGCODE;
    procedure WMMouseWheel(var Msg: TWMMouseWheel); message WM_MOUSEWHEEL;
  protected
    // IvgrWorkbookPainterOwner
    function GetDefaultColumnWidth: Integer;
    function GetDefaultRowHeight: Integer;
    function GetPixelsPerInchX: Integer;
    function GetPixelsPerInchY: Integer;
    function GetBackgroundColor: TColor;
    function GetDefaultRangeBackgroundColor: TColor;
    function GetImageForPage(APageNumber: Integer):TMetafile;

    procedure BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
    procedure AfterChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);

    procedure Paint; override;
    procedure Resize; override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;

    procedure UpdatePages;
    procedure UpdateSizes;
    procedure UpdateSizesWithEvent;
    procedure UpdateScrollBars;
    procedure UpdateAll;
    procedure DoScroll(const OldPos, NewPos: TPoint);

    function KbdLeft: Boolean;
    function KbdRight: Boolean;
    function KbdUp: Boolean;
    function KbdDown: Boolean;
    function KbdPageUp: Boolean;
    function KbdPageDown: Boolean;
    function KbdHome: Boolean;
    function KbdEnd: Boolean;
    function KbdPrint: Boolean;
    function KbdEditPageParams: Boolean;
    function KbdPriorPage: Boolean;
    function KbdNextPage: Boolean;
    function KbdPriorWorksheet: Boolean;
    function KbdNextWorksheet: Boolean;
    function KbdZoomIn: Boolean;
    function KbdZoomOut: Boolean;
    function KbdHand: Boolean;    


    procedure KeyPress(var Key: Char); override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;

    property PageMaker: TvgrPageMaker read GetPageMaker;
    property Painter: TvgrWorkbookPainter read FPainter;

    property PreviewAreaWidth: Integer read FPreviewAreaWidth;
    property PreviewAreaHeight: Integer read FPreviewAreaHeight;
    property PageRect: TRect read FPageRect;
    property FullPageRect: TRect read FFullPageRect;
    property Worksheet: TvgrWorksheet read GetWorksheet;

    property ActivePage: Integer read GetActivePage write SetActivePage;
    property PageCount: Integer read GetPageCount;
{Number of columns of pages in preview}
    property PageColCount: Integer read FPageColCount write FPageColCount;
{Number of rows of pages in preview}
    property PageRowCount: Integer read FPageRowCount write FPageRowCount;

    property ScaleMode: TvgrPreviewScaleMode read FScaleMode write SetScaleMode;
    property ScalePercent: Double read GetScalePercent write SetScalePercent;

    property WorkbookPreview: TvgrWorkbookPreview read GetWorkbookPreview;
    property HorzScrollBar: TvgrScrollBar read GetHorzScrollBar;
    property VertScrollBar: TvgrScrollBar read GetVertScrollBar;
  public
{Creates a instance of the TvgrWPPreview class.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees a instance of the TvgrWPPreview class.}
    destructor Destroy; override;

{Returns information about spot within control area.
Parameters:
  APoint - a point within control, relative to top left corner of control.
  APointInfo - object of the class, inherited from TvgrPointInfo class, that describes a point within control.
See also:
  TvgrWPNavigatorPointInfo, TvgrWPPreviewPointInfo}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWPSheetParams
  //
  /////////////////////////////////////////////////
{Internal class, used to store information about parameters of worksheet at switching between worksheets.}
  TvgrWPSheetParams = class(TObject)
  protected
    Worksheet: TvgrWorksheet;
    ScrollPosition: TPoint;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWPSheetParamsList
  //
  /////////////////////////////////////////////////
{Internal class, represents a list of TvgrWPSheetParams objects.}
  TvgrWPSheetParamsList = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrWPSheetParams;
  protected
    function FindByWorksheet(AWorksheet: TvgrWorksheet): TvgrWPSheetParams;
  public
{Enumerates objects in the list.
Parameters:
  Index - the index of the object, starts with 0.}
    property Items[Index: Integer]: TvgrWPSheetParams read GetItem; default;
  end;

{Specifies operation which can be accomplished by the user in preview.
Items:
  vgrpaZoomPageWidth - change mode of the page zooming to vgrpsmPageWidth.
  vgrpaZoomWholePage - change mode of the page zooming to vgrpsmPageWidth.
  vgrpaZoom100Percent - change mode of the page zooming to 100 percent.
  vgrpaZoomPercent - change mode of the page zooming to the specified percent.
  vgrpaFirstPage - go to the first page.
  vgrpaPriorPage - go to the prior page.
  vgrpaNextPage - go to the next page.
  vgrpaLastPage - go to the last page.
  vgrpaPrint - print workbbok.
  vgrpaPageParams - open the dialog form for editing the page parameters.
  vgrpaPriorWorksheet - go to the prior worksheet.
  vgrpaNextWorksheet - go to the next worksheet.
See also:
  TvgrWorkbookPreview, TvgrWorkbookPreview.ScaleMode, TvgrWorkbookPreview.ScalePercent.}
  TvgrWorkbookPreviewAction = (vgrpaZoomPageWidth, vgrpaZoomWholePage, vgrpaZoomTwoPages, vgrpaZoom100Percent,
                               vgrpaZoomPercent, vgrpaFirstPage, vgrpaPriorPage,
                               vgrpaNextPage, vgrpaLastPage, vgrpaPrint, vgrpaPageParams,
                               vgrpaPriorWorksheet, vgrpaNextWorksheet);
  /////////////////////////////////////////////////
  //
  // TvgrWorkbookPreview
  //
  /////////////////////////////////////////////////
{Represents the visual control, that can be used for previewing the workbook content before printing.}
  TvgrWorkbookPreview = class(TCustomControl,
                              IvgrWorkbookHandler,
                              IvgrControlParent,
                              IvgrScrollBarParent,
                              IvgrSheetsPanelParent,
                              IvgrSheetsPanelResizerParent,
                              IvgrSizerParent)
  private
    FWorkbook: TvgrWorkbook;
    FActiveWorksheetIndex: Integer;
    FBorderStyle: TBorderStyle;
    FSheetsCaptionWidth: Integer;

    FOptionsVertScrollBar: TvgrOptionsScrollBar;
    FOptionsHorzScrollBar: TvgrOptionsScrollBar;

    FWorkbookBackColor: TColor;
    FBackColor: TColor;
    FMarginsColor: TColor;
    FDefaultPopupMenu: Boolean;

    FActionMode: TvgrPreviewActionMode;
    FActiveWorksheet: TvgrWorksheet;
    FSheetParamsList: TvgrWPSheetParamsList;
    FPageMakers: TList;

    FWPNavigator: TvgrWPNavigator;
    FWPPreview: TvgrWPPreview;
    FHorzScrollBar: TvgrScrollBar;
    FVertScrollBar: TvgrScrollBar;
    FSheetsPanel: TvgrSheetsPanel;
    FSheetsPanelResizer: TvgrSheetsPanelResizer;
    FSizer: TvgrSizer;
    FDisableAlign: Boolean;
    FDeleteActiveWorksheet: Boolean;
    FPopupMenu: TPopupMenu;
    FKeyboardFilter: TvgrKeyboardFilter;

    FOnNeedPrint: TNotifyEvent;
    FOnNeedPageParams: TNotifyEvent;
    FOnActivePageChanged: TNotifyEvent;
    FOnActiveWorksheetChanged: TNotifyEvent;
    FOnScalePercentChanged: TNotifyEvent;
    FOnActiveWorksheetPagesCountChanged: TNotifyEvent;

    procedure SetOptionsVertScrollBar(Value: TvgrOptionsScrollBar);
    procedure SetOptionsHorzScrollBar(Value: TvgrOptionsScrollBar);
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMSetFocus(var Msg: TWMSetFocus); message WM_SETFOCUS;
    procedure OnScrollMessage(Sender: TObject; var Msg: TWMScroll);
    procedure SetWorkbook(Value: TvgrWorkbook);
    procedure SetBorderStyle(Value: TBorderStyle);
    procedure SetActiveWorksheetIndex(Value: Integer);
    function GetPageMakerCount: Integer;
    function GetPageMaker(Index: Integer): TvgrPageMaker;
    function FindPageMaker(AWorksheet: TvgrWorksheet): TvgrPageMaker;
    procedure InternalUpdateActiveWorksheet;
    procedure SetSheetsCaptionWidth(Value: Integer);
    procedure OnOptionsChanged(Sender: TObject);
    procedure SetWorkbookBackColor(Value: TColor);
    procedure SetBackColor(Value: TColor);
    procedure SetMarginsColor(Value: TColor);
    function GetTopOffset: Integer;
    function GetLeftOffset: Integer;
    function GetOffset: TPoint;
    function GetScaleMode: TvgrPreviewScaleMode;
    procedure SetScaleMode(Value: TvgrPreviewScaleMode);
    function GetScalePercent: Double;
    procedure SetScalePercent(Value: Double);
    function IsScalePercentStored: Boolean;
    function GetActivePage: Integer;
    function GetActionMode: TvgrPreviewActionMode;
    function GetActivePrintPage: TvgrPrintPage;
    procedure SetActivePage(Value: Integer);
    procedure SetActionMode(Value: TvgrPreviewActionMode);
    function GetPageCount: Integer;
  protected
    // IvgrWorkbookHandler
    procedure BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo); virtual;
    procedure AfterChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo); virtual;
    procedure DeleteItemInterface(AItem: IvgrWBListItem); virtual;
    // IvgrControlParent
    procedure DoMouseMove(Sender: TControl; Shift: TShiftState; X, Y: Integer);
    procedure DoMouseDown(Sender: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
    procedure DoMouseUp(Sender: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
    procedure DoDblClick(Sender: TControl; var AExecuteDefault: Boolean);
    procedure DoDragOver(AChild: TControl; Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure DoDragDrop(AChild: TControl; Source: TObject; X, Y: Integer);
    // IvgrScrollBarParent
    function GetScrollBarOptions(AScrollBar: TvgrScrollBar): TvgrOptionsScrollBar;
    // IvgrSheetsPanelParent
    function SheetsPanelParentGetWorkbook: TvgrWorkbook;
    function SheetsPanelParentGetActiveWorksheet: TvgrWorksheet;
    function SheetsPanelParentGetActiveWorksheetIndex: Integer;
    procedure SheetsPanelParentSetActiveWorksheetIndex(Value: Integer);
    function IvgrSheetsPanelParent.GetWorkbook = SheetsPanelParentGetWorkbook;
    function IvgrSheetsPanelParent.GetActiveWorksheet = SheetsPanelParentGetActiveWorksheet;
    function IvgrSheetsPanelParent.GetActiveWorksheetIndex = SheetsPanelParentGetActiveWorksheetIndex;
    procedure IvgrSheetsPanelParent.SetActiveWorksheetIndex = SheetsPanelParentSetActiveWorksheetIndex;
    // IvgrSheetsPanelResizerParent
    procedure SheetsPanelResizerParentDoResize(AOffset: Integer);
    procedure IvgrSheetsPanelResizerParent.DoResize = SheetsPanelResizerParentDoResize;
    // IvgrSizerParent

    procedure SaveWPSheetParams; virtual;
    procedure RestoreWPSheetParams; virtual;

    procedure ClearPageMakers;
    procedure UpdatePageMakers;

    procedure KbdAddCommands;

    procedure AlignSheetCaptions;
    procedure AlignControls(AControl: TControl; var AAlignRect: TRect); override;
    procedure Notification(AComponent : TComponent; AOperation : TOperation); override;
    procedure Paint; override;

    procedure OnPopupMenuClick(Sender: TObject); virtual;
    procedure OnPopupMenuClickChangeMode(Sender: TObject); virtual;
    
    procedure InitPopupMenu(APopupMenu: TPopupMenu); virtual;

    procedure DoScroll(const OldPos, NewPos: TPoint);

    procedure DoNeedPrint;
    procedure DoNeedPageParams;
    procedure DoActivePageChanged;
    procedure DoActiveWorksheetChanged;
    procedure DoScalePercentChanged;
    procedure DoActiveWorksheetPagesCountChanged;

    property PageMakerCount: Integer read GetPageMakerCount;
    property PageMakers[Index: Integer]: TvgrPageMaker read GetPageMaker;

    property HorzScrollBar: TvgrScrollBar read FHorzScrollBar;
    property VertScrollBar: TvgrScrollBar read FVertScrollBar;

    property KeyboardFilter: TvgrKeyboardFilter read FKeyboardFilter;
  public
{Creates a instance of the TvgrWorkbookPreview class.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner : TComponent); override;
{Frees a instance of the TvgrWorkbookPreview class.}
    destructor Destroy; override;
{Procedure sets number of pages per X and per Y in preview}
    procedure SetMultiPagePreview(const X, Y: Integer);
{Shows a popup menu at the specified position.
Parameters:
  APoint - specifies the position of the top-left corner of the popup menu.}
    procedure ShowPopupMenu(const APoint: TPoint; APointInfo: TvgrPointInfo);

{Returns the true value if specified action are allowed.
Parameters:
  AAction - the action code.
Return value:
  Returns the true value if specified action are allowed.
See also:
  TvgrWorkbookPreviewAction, DoPreviewAction}
    function IsPreviewActionAllowed(AAction: TvgrWorkbookPreviewAction): Boolean;
{Executes the specified action.
Parameters:
  AAction - the action code.
See also:
  TvgrWorkbookPreviewAction, IsPreviewActionAllowed}
    procedure DoPreviewAction(AAction: TvgrWorkbookPreviewAction);
{Changes current preview action mode
See also:
  TvgrPreviewActionMode}    
    procedure DoChangeActionMode(AActionMode: TvgrPreviewActionMode);

{Prints the workbook. TvgrWorkbookPreview not contains the print engine, he generates the OnNeedPrint
event, you may process this event and print workbook with using of the TvgrReportEngine component.
See also:
  OnNeedPrint, TvgrReportEngine}
    procedure Print;
{Opens the dialog form for editing the page properties. TvgrWorkbookPreview not contains the built-in,
dialog form, he generates the OnNeedPageParams event,
you may process this event and edit page properties of the selected worksheet with using
of the TvgrPageSetupDialog component.
See also:
  OnNeedPageParams, TvgrPageSetupDialog}
    procedure EditPageParams;

{Returns the active worksheet.}
    property ActiveWorksheet: TvgrWorksheet read FActiveWorksheet;
{Returns the current position of the vertical scrollbar.}
    property TopOffset: Integer read GetTopOffset;
{Returns the current position of the horizontal scrollbar.}
    property LeftOffset: Integer read GetLeftOffset;
{Returns the point, which contains the current positions of the vertical and horizontal scrollbars.}
    property Offset: TPoint read GetOffset;
{Returns index of the active page.}
    property ActivePage: Integer read GetActivePage write SetActivePage;
{Returns TvgrPrintPage, that identifies active page.}
    property ActivePrintPage: TvgrPrintPage read GetActivePrintPage;
{Returns the amount of the pages for the active worksheet.}
    property PageCount: Integer read GetPageCount;
  published
    property Align;
    property Anchors;
    property Constraints;
    property DockSite;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Visible;
    property TabOrder;
    property TabStop;
    property PopupMenu;

{Specifies workbook that are previewed with this control.
See also:
  TvgrWorkbook}
    property Workbook: TvgrWorkbook read FWorkbook write SetWorkbook;
{Specifies the mode of the current page action(Hand Tool, Zoom In Tool, Zoom Out Tool).}
    property ActionMode: TvgrPreviewActionMode read GetActionMode write SetActionMode default vgrpamNone;
{Specifies the index of the active worksheet.
See also:
  ActiveWorksheet}
    property ActiveWorksheetIndex: Integer read FActiveWorksheetIndex write SetActiveWorksheetIndex default -1;
{Determines the style of the line drawn around the perimeter of the preview control.}
    property BorderStyle: TBorderStyle read FBorderStyle write SetBorderStyle default bsSingle;
{Specifies width of sheets caption area. Default value is determined by cDefaultSheetsCaptionWidth const.}
    property SheetsCaptionWidth: Integer read FSheetsCaptionWidth write SetSheetsCaptionWidth default cDefaultSheetsCaptionWidth;
{Specifies common options of vertical scrollbar.
See also:
  TvgrOptionsScrollBar}
    property OptionsVertScrollBar: TvgrOptionsScrollBar read FOptionsVertScrollBar write SetOptionsVertScrollBar;
{Specifies common options of horizontal scrollbar.
See also:
  TvgrOptionsScrollBar}
    property OptionsHorzScrollBar: TvgrOptionsScrollBar read FOptionsHorzScrollBar write SetOptionsHorzScrollBar;
{Specifies the mode of the page zooming.
See also:
  ScalePercent, TvgrPreviewScaleMode}
    property ScaleMode: TvgrPreviewScaleMode read GetScaleMode write SetScaleMode default vgrpsmPageWidth;
{Specifies the page zooming percent.
See also:
  ScaleMode}
    property ScalePercent: Double read GetScalePercent write SetScalePercent stored IsScalePercentStored;

{Specifies the background color.
See also:
  WorkbookBackColor}
    property BackColor: TColor read FBackColor write SetBackColor default clBtnFace;
{Specifies the color, that used to filling contents of the transparent ranges and area not occupied by ranges. 
See also:
  BackColor}
    property WorkbookBackColor: TColor read FWorkbookBackColor write SetWorkbookBackColor default clWhite;
{Specifies the color, that is used to filling the page margins.}
    property MarginsColor: TColor read FMarginsColor write SetMarginsColor default clSilver;

{Specifies the boolean value, that indicates should be used built-in popup menu or not.}
    property DefaultPopupMenu: Boolean read FDefaultPopupMenu write FDefaultPopupMenu default True;

{This event occurs, when user wants to print the workbook.}
    property OnNeedPrint: TNotifyEvent read FOnNeedPrint write FOnNeedPrint;
{This event occurs, when user wants to edit the page properties.}
    property OnNeedPageParams: TNotifyEvent read FOnNeedPageParams write FOnNeedPageParams;
{This event occurs, when active page is changed.
See also:
  ActivePage, ActivePageIndex}
    property OnActivePageChanged: TNotifyEvent read FOnActivePageChanged write FOnActivePageChanged;
{This event occurs, when active worksheet is changed.
See also:
  ActiveWorksheet, ActiveWorksheetIndex}
    property OnActiveWorksheetChanged: TNotifyEvent read FOnActiveWorksheetChanged write FOnActiveWorksheetChanged;
{This event occurs, when the zooming percent is changed.}
    property OnScalePercentChanged: TNotifyEvent read FOnScalePercentChanged write FOnScalePercentChanged;
{This event occurs, when the amount of the pages of the active worksheet is changed.}
    property OnActiveWorksheetPagesCountChanged: TNotifyEvent read FOnActiveWorksheetPagesCountChanged write FOnActiveWorksheetPagesCountChanged;

    property OnDblClick;
    property OnMouseDown;
    property OnMouseUp;
    property OnMouseMove;
  end;

{$IFDEF VGR_UPDATEIMAGE_TEST}
var
  FTestCanvas: TCanvas;
{$ENDIF}

implementation

{$R vgrWorkbookPreview.res}
{$R ..\res\vgr_WorkbookPreviewStrigns.res}

uses
  vgr_GUIFunctions, vgr_Dialogs, vgr_StringIDs, vgr_Localize;

const
  cMinScalePercent = 10;
  cMaxScalePercent = 600;
{Number of images in preview's cache}
  cNumImagesInPreviewCache = 150;
  cGraySize = 20;

  { Cursor ID }
  cHandToolCursorID = 6;
  { Cursor ID }
  cHandTool2CursorID = 7;
  { Cursor ID }
  cZoomInCursorID = 8;
  { Cursor ID }
  cZoomOutCursorID = 9;
  { Cursor resource name }
  cHandToolCursorResName = 'VGR_HAND_TOOL';
  { Cursor resource name }
  cHandTool2CursorResName = 'VGR_HAND_TOOL2';
  { Cursor resource name }
  cZoomInCursorResName = 'VGR_ZOOMIN_TOOL';
  { Cursor resource name }
  cZoomOutCursorResName = 'VGR_ZOOMOUT_TOOL';

  svgrPreviewNavigatorBitmaps = 'VGR_BMP_PREVIEWNAVIGATORBITMAP';


var
  FPreviewNavigatorBitmaps: TImageList;

  FHandToolCursor: TCursor;
  FHandTool2Cursor: TCursor;
  FZoomInCursor: TCursor;
  FZoomOutCursor: TCursor;

/////////////////////////////////////////////////
//
// TvgrWPImagesCache
//
/////////////////////////////////////////////////

constructor TvgrWPImagesCache.Create;
begin
  FPos := 0;
end;

destructor TvgrWPImagesCache.Destroy;
begin
  clearCache;
  inherited;
end;

function TvgrWPImagesCache.GetImage(ANumber: Integer):TMetafile;
var
  I: Integer;
begin
  Result := Nil;
  for I:=0 to Length(FNumbers)-1 do
  begin
    if ANumber = FNumbers [I] then
    begin
      Result := FImages[I];
    end;
  end;
end;

procedure TvgrWPImagesCache.ClearCache;
var
  I: Integer;
begin
  for I:=0 to Length(FImages)-1 do
  begin
    FreeAndNil(FImages[I]);
  end;
  SetLength(FImages,0);
  SetLength(FNumbers,0);
  FPos := 0;
end;

function TvgrWPImagesCache.setImage(ANumber: Integer;AImage: TMetafile):TMetafile;
var
  I: Integer;
begin
  I := Length(FImages);
  if I < cNumImagesInPreviewCache then
  begin
    SetLength(FImages,I+1);
    SetLength(FNumbers,I+1);
  end
  else
  begin
    FImages[FPos].Free;
    I := FPos;
    FPos := FPos + 1;
    if FPos >= cNumImagesInPreviewCache then
      FPos := 0;
  end;
  FImages[I] := AImage;
  FNumbers[I] := ANumber;
  Result := AImage;
end;

/////////////////////////////////////////////////
//
// TvgrWPMarginsResizer
//
/////////////////////////////////////////////////
constructor TvgrWPMarginsResizer.Create(AParentWPPreview: TvgrWPPreview);
begin
  FCautionStep := 0;
  FCurrentAction := -1;
  FCurrentMargin := -1;  
  FParentWPPreview := AParentWPPreview;
  FMarginOffsetSizePix := 0;
  NeedElaborateScroll := False;
  FSpotWidth := 4;
  FHintWindow := THintWindow.Create(AParentWPPreview);
  FHintWindow.Color := clInfoBk;
end;

destructor TvgrWPMarginsResizer.Destroy;
begin
   FHintWindow.Free;
   ClearPagesSpots;
   inherited;
end;

procedure TvgrWPMarginsResizer.ClearPagesSpots;
begin
  SetLength(FPagesRects, 0);
  SetLength(FPagesRectsNumbers, 0);
  FActivePageRect := Rect(0, 0, 0, 0);
end;

procedure TvgrWPMarginsResizer.DrawMarginsResizerLines;
var
  ALOffset, AROffset, ATOffset, ABOffset, AHOffset, AFOffset: Integer;
begin
  //Margins:
  ALOffset := 0;
  AROffset := 0;
  ATOffset := 0;
  ABOffset := 0;
  //Header:
  AHOffset := 0;
  //Footer
  AFOffset := 0;
  case FCurrentMargin of
    0: ALOffset := FMarginOffsetSizePix;
    1: AROffset := FMarginOffsetSizePix;
    //top margin    
    2:
      begin
        ATOffset := FMarginOffsetSizePix;
        AHOffset := FMarginOffsetSizePix;
      end;
    //bottom margin
    3:
      begin
        ABOffset := FMarginOffsetSizePix;
        AFOffset := FMarginOffsetSizePix;
      end;
    //header
    4: AHOffset := FMarginOffsetSizePix;
    //footer
    5: AFOffset := FMarginOffsetSizePix;
  end;
  FParentWPPreview.Canvas.Brush.Color := clWhite;
  FParentWPPreview.Canvas.Pen.Style := psDot;//psDash
  FParentWPPreview.Canvas.Pen.Color := clBlack;
  FParentWPPreview.Canvas.Pen.Width := 1;
  //Left margin:
  FParentWPPreview.Canvas.Rectangle(Rect(FCLeft-1+ALOffset, FCBTop, FCLeft-1+ALOffset+1, FCBBottom));
  FParentWPPreview.Canvas.MoveTo(FCLeft-1+ALOffset,FCBTop);FParentWPPreview.Canvas.LineTo(FCLeft-1+ALOffset,FCBBottom);
  ExcludeClipRect(FParentWPPreview.Canvas, Rect(FCLeft-1+ALOffset, FCBTop, FCLeft-1+ALOffset+1, FCBBottom));
  //Right margin:
  FParentWPPreview.Canvas.Rectangle(Rect(FCRight+1-AROffset, FCBTop, FCRight+1-AROffset+1, FCBBottom));
  FParentWPPreview.Canvas.MoveTo(FCRight+1-AROffset,FCBTop);FParentWPPreview.Canvas.LineTo(FCRight+1-AROffset,FCBBottom);
  ExcludeClipRect(FParentWPPreview.Canvas, Rect(FCRight+1-AROffset, FCBTop, FCRight+1-AROffset+1, FCBBottom));
  //Top margin:
  FParentWPPreview.Canvas.Rectangle(Rect(FCBLeft, FCTop-1+ATOffset, FCBRight, FCTop-1+ATOffset+1));
  FParentWPPreview.Canvas.MoveTo(FCBLeft,FCTop-1+ATOffset);FParentWPPreview.Canvas.LineTo(FCBRight,FCTop-1+ATOffset);
  ExcludeClipRect(FParentWPPreview.Canvas, Rect(FCBLeft, FCTop-1+ATOffset, FCBRight, FCTop-1+ATOffset+1));
  //Bottom margin:
  FParentWPPreview.Canvas.Rectangle(Rect(FCBLeft, FCBottom+1+ABOffset, FCBRight, FCBottom+1+ABOffset+1));
  FParentWPPreview.Canvas.MoveTo(FCBLeft,FCBottom+1+ABOffset);FParentWPPreview.Canvas.LineTo(FCBRight,FCBottom+1+ABOffset);
  ExcludeClipRect(FParentWPPreview.Canvas, Rect(FCBLeft, FCBottom+1+ABOffset, FCBRight, FCBottom+1+ABOffset+1));
  //Header:
  FParentWPPreview.Canvas.Rectangle(Rect(FCBLeft, FCHeader-1+AHOffset, FCBRight, FCHeader-1+AHOffset+1));
  FParentWPPreview.Canvas.MoveTo(FCBLeft,FCHeader-1+AHOffset);FParentWPPreview.Canvas.LineTo(FCBRight,FCHeader-1+AHOffset);
  ExcludeClipRect(FParentWPPreview.Canvas, Rect(FCBLeft, FCHeader-1+AHOffset, FCBRight, FCHeader-1+AHOffset+1));
  //Footer:
  FParentWPPreview.Canvas.Rectangle(Rect(FCBLeft, FCFooter+1+AFOffset, FCBRight, FCFooter+1+AFOffset+1));
  FParentWPPreview.Canvas.MoveTo(FCBLeft,FCFooter+1+AFOffset);FParentWPPreview.Canvas.LineTo(FCBRight,FCFooter+1+AFOffset);
  ExcludeClipRect(FParentWPPreview.Canvas, Rect(FCBLeft, FCFooter+1+AFOffset, FCBRight, FCFooter+1+AFOffset+1));
end;

procedure TvgrWPMarginsResizer.DoMouseUp(APoint, APointToScreen: TPoint);
var
  I: Integer;
begin
  //Set new active page if needed:
  if (Abs(FBegPointMoseDown.X-APoint.X)+Abs(FBegPointMoseDown.Y-APoint.Y)<20)then
    for I:=0 to Length(FPagesRects)-1 do
      if PointInSelRect(APoint,FPagesRects[I]) = true then
        if (FPagesRectsNumbers[I]<>FParentWPPreview.FActivePage) then
        begin
          FParentWPPreview.SetActivePage(FPagesRectsNumbers[I]);
          FCurrentAction := 0;
          SetCurrentCursor;
          exit;
        end;
        
  case FCurrentAction of
     1,2: // resize margin, header, footer action:
        begin
          FCurrentAction := 0;
          FMarginOffsetSizePix := 0;
          if FMarginNewSize<FCautionStep then
            FMarginNewSize := FCautionStep;
          InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
          case FCurrentMargin of
            0:  //left margin:
              begin
                if(FMarginNewSize>=FParentWPPreview.PageMaker.PageProperties.UnitsWidth[FMesUnits]
                                   -FParentWPPreview.PageMaker.PageProperties.Margins.UnitsRight[FMesUnits]) then
                   FMarginNewSize := FParentWPPreview.PageMaker.PageProperties.UnitsWidth[FMesUnits]
                                    -FParentWPPreview.PageMaker.PageProperties.Margins.UnitsRight[FMesUnits]-FCautionStep;
                FParentWPPreview.PageMaker.PageProperties.Margins.UnitsLeft[FMesUnits] := FMarginNewSize;
              end;
            1:  //right margin:
              begin
                if(FMarginNewSize>=FParentWPPreview.PageMaker.PageProperties.UnitsWidth[FMesUnits]
                                   -FParentWPPreview.PageMaker.PageProperties.Margins.UnitsLeft[FMesUnits])then
                   FMarginNewSize := FParentWPPreview.PageMaker.PageProperties.UnitsWidth[FMesUnits]
                                    -FParentWPPreview.PageMaker.PageProperties.Margins.UnitsLeft[FMesUnits]-FCautionStep;
                FParentWPPreview.PageMaker.PageProperties.Margins.UnitsRight[FMesUnits] := FMarginNewSize;
              end;
            2:  //top margin:
              begin
                if(FMarginNewSize >= FParentWPPreview.PageMaker.PageProperties.UnitsHeight[FMesUnits]-FParentWPPreview.PageMaker.PageProperties.Margins.UnitsBottom[FMesUnits]) then
                  FMarginNewSize := FParentWPPreview.PageMaker.PageProperties.UnitsHeight[FMesUnits]-FParentWPPreview.PageMaker.PageProperties.Margins.UnitsBottom[FMesUnits]-FCautionStep;
                FParentWPPreview.PageMaker.PageProperties.Margins.UnitsTop[FMesUnits] := FMarginNewSize;
              end;
            3:  //bottom margin:
              begin
                if(FMarginNewSize>=FParentWPPreview.PageMaker.PageProperties.UnitsHeight[FMesUnits]
                                   -FParentWPPreview.PageMaker.PageProperties.Margins.UnitsTop[FMesUnits]) then
                  FMarginNewSize := FParentWPPreview.PageMaker.PageProperties.UnitsHeight[FMesUnits]
                                    -FParentWPPreview.PageMaker.PageProperties.Margins.UnitsTop[FMesUnits]-FCautionStep;
                FParentWPPreview.PageMaker.PageProperties.Margins.UnitsBottom[FMesUnits] := FMarginNewSize;
              end;
            4:  //header:
              begin
                if(FMarginNewSize < 0) then
                  FMarginNewSize := 0;
                FParentWPPreview.PageMaker.PageProperties.Header.UnitsHeight[FMesUnits] := FMarginNewSize;
              end;
            5:  //footer:
              begin
                if(FMarginNewSize < 0) then
                  FMarginNewSize := 0;
                FParentWPPreview.PageMaker.PageProperties.Footer.UnitsHeight[FMesUnits] := FMarginNewSize;
              end;
          end;
          GetMarginsSize;
        end;
    5,6: // Zoom In Tool or Zoom Out Tool action:
      begin
        DrawFocusRect(FParentWPPreview.Canvas.Handle,FPrevMarginRect);
        //Attention! Do not use property FParentWPPreview.ScaleMode here! Use FParentWPPreview.FScaleMode!
        FParentWPPreview.FScaleMode := vgrpsmPercent;
//        FParentWPPreview.WorkbookPreview.DoScalePercentChanged;
        NeedElaborateScroll := False;
        NeedElaborateScroll_Rect := FPrevMarginRect;
        NeedElaborateScroll_ActivePageRect := FActivePageRect;
        //Zoom In:
        if(FCurrentAction = 5) then
        begin
          if (FPrevMarginRect.Right-FPrevMarginRect.Left) < 20 then
          begin
            NeedElaborateScroll := True;
            NeedElaborateScroll_Rect.Left := APoint.X;
            NeedElaborateScroll_Rect.Right := NeedElaborateScroll_Rect.Left + 1;
            NeedElaborateScroll_Rect.Top := APoint.Y;
            NeedElaborateScroll_Rect.Bottom := NeedElaborateScroll_Rect.Top + 1;            
            FParentWPPreview.ScalePercent := FParentWPPreview.ScalePercent + 10;
            NeedElaborateScroll := False;
          end
          else
          begin
            NeedElaborateScroll := True;
            if(
            FParentWPPreview.ClientWidth*FParentWPPreview.ScalePercent/(FPrevMarginRect.Right-FPrevMarginRect.Left+0.00000000000000001)
            )<(
            FParentWPPreview.ClientHeight*FParentWPPreview.ScalePercent/(FPrevMarginRect.Bottom-FPrevMarginRect.Top+0.00000000000000001)
            )
            then
              FParentWPPreview.ScalePercent := FParentWPPreview.ClientWidth*FParentWPPreview.ScalePercent/(FPrevMarginRect.Right-FPrevMarginRect.Left+0.00000000000000001)
            else
              FParentWPPreview.ScalePercent := FParentWPPreview.ClientHeight*FParentWPPreview.ScalePercent/(FPrevMarginRect.Bottom-FPrevMarginRect.Top+0.00000000000000001);
            NeedElaborateScroll := False;
          end;
        end;
        //Zoom Out:
        if(FCurrentAction = 6) then
        begin
          if (FPrevMarginRect.Right-FPrevMarginRect.Left) < 20 then
            FParentWPPreview.ScalePercent := FParentWPPreview.ScalePercent - 10
          else
            FParentWPPreview.ScalePercent := FParentWPPreview.ScalePercent-(FPrevMarginRect.Right-FPrevMarginRect.Left)*FParentWPPreview.ScalePercent/(FParentWPPreview.ClientWidth+0.00000000000000001);
        end;
      end;
  end;
  FPrevMarginRect := Rect(0,0,0,0);
  FCurrentAction := 0;
  SetCurrentCursor;
end;

procedure TvgrWPMarginsResizer.DoMouseDown(APoint, APointToScreen: TPoint);
begin
  FBegPointMoseDown := APoint;
  FCurrentAction := 0;
  FMarginOffsetSizePix := 0;
  if PointInSelRect(APoint,FSpotLeftMargin) = true then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    FCurrentMargin := 0;
    FCurrentAction := 1;
  end;
  if PointInSelRect(APoint,FSpotRightMargin) = true then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    FCurrentMargin := 1;
    FCurrentAction := 1;
  end;
  if PointInSelRect(APoint,FSpotTopMargin) = true then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    FCurrentMargin := 2;
    FCurrentAction := 1;
  end;
  if PointInSelRect(APoint,FSpotBottomMargin) = true then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    FCurrentMargin := 3;
    FCurrentAction := 1;
  end;
  if PointInSelRect(APoint,FSpotHeader) = true then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    FCurrentMargin := 4;
    FCurrentAction := 1;
  end;
  if PointInSelRect(APoint,FSpotFooter) = true then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    FCurrentMargin := 5;
    FCurrentAction := 1;
  end;
  if(FCurrentAction <1 ) then
  begin
    FBegPoint := APoint;
    FBegPointToScreen := APointToScreen;
    case FParentWPPreview.WorkbookPreview.ActionMode of
      vgrpamNone:
          begin
          end;
      vgrpamHand: //Hand Tool:
          begin
            FParentWPPreview.Cursor := FHandTool2Cursor;
            FCurrentAction := 4;
          end;
      vgrpamZoomIn: // Zoom In Tool:
          begin
            if(PointInSelRect(APoint,FActivePageRect) = true) then
            begin
              FCurrentAction := 5;
            end;
          end;
      vgrpamZoomOut: // Zoom Out Tool:
          begin
            if(PointInSelRect(APoint,FActivePageRect) = true) then
            begin
              FCurrentAction := 6;
            end;
          end;
    end;
  end;
end;

procedure TvgrWPMarginsResizer.DoMouseMove(APoint, APointToScreen: TPoint; Shift: TShiftState);
var
  AHint: string;
begin
  case FCurrentAction of
    -1,0: //no action:
      begin
        //left margin:
        if PointInSelRect(APoint,FSpotLeftMargin) = true then
        begin
          GetMarginsSize;
          AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_LeftMargin);
          AHint := AHint +' '+FloatToStr(Round(FLeft*10/FCautionStep)/(10/FCautionStep))+' '+FMeasurementUnits;
          ShowHintForMargin(AHint, 0, APointToScreen);
        end
        else
        begin
          //right margin:
          if PointInSelRect(APoint,FSpotRightMargin) = true then
          begin
            GetMarginsSize;
            AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_RightMargin);
            AHint := AHint + ' '+FloatToStr(Round(FRight*10/FCautionStep)/(10/FCautionStep))+' '+FMeasurementUnits;
            ShowHintForMargin(AHint, 1, APointToScreen);
          end
          else
          begin
            //top margin:
            if PointInSelRect(APoint,FSpotTopMargin) = true then
            begin
              GetMarginsSize;
              AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_TopMargin);
              AHint := AHint + ' '+FloatToStr(Round(FTop*10/FCautionStep)/(10/FCautionStep))+' '+FMeasurementUnits;
              ShowHintForMargin(AHint, 2, APointToScreen);
            end
            else
            begin
              //bottom margin:
              if PointInSelRect(APoint,FSpotBottomMargin) = true then
              begin
                GetMarginsSize;
                AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_BottomMargin);
                AHint := AHint + ' '+FloatToStr(Round(FBottom*10/FCautionStep)/(10/FCautionStep))+' '+FMeasurementUnits;
                ShowHintForMargin(AHint, 3, APointToScreen);
              end
              else
              begin
                //Header:
                if PointInSelRect(APoint,FSpotHeader) = true then
                begin
                  GetMarginsSize;
                  AHint := vgrLoadStr(svgrid__vgr_WorkbookPreview__Hint__Header);
                  AHint := AHint + ' '+FloatToStr(Round(FHeader*10/FCautionStep)/(10/FCautionStep))+' '+FMeasurementUnits;
                  ShowHintForMargin(AHint, 4, APointToScreen);
                end
                else
                begin
                  //Footer:
                  if PointInSelRect(APoint,FSpotFooter) = true then
                  begin
                    GetMarginsSize;
                    AHint := vgrLoadStr(svgrid__vgr_WorkbookPreview__Hint__Footer);
                    AHint := AHint + ' '+FloatToStr(Round(FFooter*10/FCautionStep)/(10/FCautionStep))+' '+FMeasurementUnits;
                    ShowHintForMargin(AHint, 5, APointToScreen);
                  end
                  else
                  begin
                    FHintWindow.ReleaseHandle;
                    setCurrentCursor;
                  end;
                end;
              end;
            end;
          end;
        end;

      end;
    1: //select current margin, begin resizing margin, activate hint,
       //continue resizing margin, change position of hint
      begin
        case FCurrentMargin of
          0:  //left margin:
          begin
            FPrevMarginRect.Left:= FCLeft-1+FMarginOffsetSizePix;
            FPrevMarginRect.Top := FCBTop;
            FPrevMarginRect.Right := FCLeft-1+FMarginOffsetSizePix+1;
            FPrevMarginRect.Bottom := FCBBottom;
            FMarginOffsetSizePix := APoint.X - FBegPoint.X;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            DrawMarginsResizerLines;
            if ssShift in Shift then
            begin
              FMarginNewSize := Round((FLeft*((FCLeft-FCBLeft)+FMarginOffsetSizePix)
                                         /(FCLeft-FCBLeft+0.0000000000000001))/FCautionStep)*FCautionStep;
            end
            else
            begin
              FMarginNewSize := Round((FLeft*((FCLeft-FCBLeft)+FMarginOffsetSizePix)
                                         /(FCLeft-FCBLeft+0.0000000000000001))*10/FCautionStep)/(10/FCautionStep);
            end;
            AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_LeftMargin);
            AHint := AHint + ' '+FloatToStr(FMarginNewSize)+' '+FMeasurementUnits;
            ShowHintForMargin(AHint, 0, FBegPointToScreen);
          end;
          1:  //right margin:
          begin
            FPrevMarginRect.Left:= FCRight+1-FMarginOffsetSizePix;
            FPrevMarginRect.Top := FCBTop;
            FPrevMarginRect.Right := FCRight+1-FMarginOffsetSizePix+1;
            FPrevMarginRect.Bottom := FCBBottom;
            FMarginOffsetSizePix := FBegPoint.X-APoint.X;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            DrawMarginsResizerLines;
            if ssShift in Shift then
            begin
              FMarginNewSize := Round((FRight*((FCBRight-FCRight)+FMarginOffsetSizePix)
                                         /(FCBRight-FCRight+0.00000000000001))/FCautionStep)*FCautionStep;
            end
            else
            begin
            FMarginNewSize := Round((FRight*((FCBRight-FCRight)+FMarginOffsetSizePix)
                                         /(FCBRight-FCRight+0.00000000000001))*10/FCautionStep)/(10/FCautionStep);
            end;
            AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_RightMargin);
            AHint := AHint + ' '+FloatToStr(FMarginNewSize)+' '+FMeasurementUnits;
            ShowHintForMargin(AHint, 1, FBegPointToScreen);
          end;
          2:  //top margin:
          begin
            //clear old header line:
            FPrevMarginRect.Left:= FCBLeft;
            FPrevMarginRect.Top := FCHeader-1+FMarginOffsetSizePix;
            FPrevMarginRect.Right := FCBRight;
            FPrevMarginRect.Bottom := FCHeader-1+FMarginOffsetSizePix+1;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            //clear old margin line:
            FPrevMarginRect.Left:= FCBLeft;
            FPrevMarginRect.Top := FCTop-1+FMarginOffsetSizePix;
            FPrevMarginRect.Right := FCBRight;
            FPrevMarginRect.Bottom := FCTop-1+FMarginOffsetSizePix+1;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            FMarginOffsetSizePix := APoint.Y-FBegPoint.Y;            
            DrawMarginsResizerLines;
            if ssShift in Shift then
            begin
              FMarginNewSize := Round((FTop * ((FCTop - FCBTop) + FMarginOffsetSizePix)
                                         /(FCTop - FCBTop + 0.00000000000001)) / FCautionStep) * FCautionStep;
            end
            else
            begin
              FMarginNewSize := Round((FTop * ((FCTop - FCBTop) + FMarginOffsetSizePix)
                                         /(FCTop - FCBTop + 0.00000000000001)) * 10 / FCautionStep) / (10 / FCautionStep);
            end;
            AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_TopMargin);
            AHint := AHint + ' ' + FloatToStr(FMarginNewSize) + ' ' +FMeasurementUnits;
            ShowHintForMargin(AHint, 2, FBegPointToScreen);
          end;
          3:  //bottom margin:
          begin
            //clear old footer line:
            FPrevMarginRect.Left:= FCBLeft;
            FPrevMarginRect.Top := FCFooter+1+FMarginOffsetSizePix;
            FPrevMarginRect.Right := FCBRight;
            FPrevMarginRect.Bottom := FCFooter+1+FMarginOffsetSizePix+1;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            //Clear old margin line:
            FPrevMarginRect.Left:= FCBLeft;
            FPrevMarginRect.Top := FCBottom+1+FMarginOffsetSizePix;
            FPrevMarginRect.Right := FCBRight;
            FPrevMarginRect.Bottom := FCBottom+1+FMarginOffsetSizePix+1;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            FMarginOffsetSizePix := APoint.Y-FBegPoint.Y;            
            DrawMarginsResizerLines;
            if ssShift in Shift then
            begin
              FMarginNewSize := Round((FBottom * ((FCBottom - FCBBottom) + FMarginOffsetSizePix)
                                          /(FCBottom - FCBBottom + 0.00000000000001)) / FCautionStep) * FCautionStep;
            end
            else
            begin
              FMarginNewSize := Round((FBottom * ((FCBottom - FCBBottom) + FMarginOffsetSizePix)
                                          / (FCBottom - FCBBottom + 0.00000000000001)) * 10 / FCautionStep) / (10 / FCautionStep);
            end;
            AHint := vgrLoadStr(svgrid_vgr_WorkbookPreview_Hint_BottomMargin);
            AHint := AHint + ' ' + FloatToStr(FMarginNewSize) + ' ' + FMeasurementUnits;
            ShowHintForMargin(AHint, 3, FBegPointToScreen);
          end;
          4:  //header:
          begin
            FPrevMarginRect.Left:= FCBLeft;
            FPrevMarginRect.Top := FCHeader - 1 + FMarginOffsetSizePix;
            FPrevMarginRect.Right := FCBRight;
            FPrevMarginRect.Bottom := FCHeader - 1 + FMarginOffsetSizePix+1;
            FMarginOffsetSizePix := APoint.Y-FBegPoint.Y;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            DrawMarginsResizerLines;
            if ssShift in Shift then
            begin
              FMarginNewSize := Round((FHeader * ((FCHeader - FCTop) + FMarginOffsetSizePix)
                                         /(FCHeader - FCTop + 0.00000000000001)) / FCautionStep) * FCautionStep;
            end
            else
            begin
              FMarginNewSize := Round((FHeader * ((FCHeader - FCTop) + FMarginOffsetSizePix)
                                         /(FCHeader - FCTop + 0.00000000000001)) * 10 / FCautionStep) / (10 / FCautionStep);
            end;
            AHint := vgrLoadStr(svgrid__vgr_WorkbookPreview__Hint__Header);
            AHint := AHint + ' ' + FloatToStr(FMarginNewSize) + ' ' +FMeasurementUnits;
            ShowHintForMargin(AHint, 4, FBegPointToScreen);
          end;
          5:  //footer:
          begin
            //Clear old footer line:
            FPrevMarginRect.Left:= FCBLeft;
            FPrevMarginRect.Top := FCFooter + 1 + FMarginOffsetSizePix;
            FPrevMarginRect.Right := FCBRight;
            FPrevMarginRect.Bottom := FCFooter + 1 + FMarginOffsetSizePix + 1;
            InvalidateRect(FParentWPPreview.Handle,@FPrevMarginRect,true);
            FMarginOffsetSizePix := APoint.Y-FBegPoint.Y;            
            DrawMarginsResizerLines;
            if ssShift in Shift then
            begin
              FMarginNewSize := Round((FFooter * ((FCFooter - FCBottom) + FMarginOffsetSizePix)
                                          /(FCFooter - FCBottom + 0.00000000000001)) / FCautionStep) * FCautionStep;
            end
            else
            begin
              FMarginNewSize := Round((FFooter * ((FCFooter - FCBottom) + FMarginOffsetSizePix)
                                          / (FCFooter - FCBottom + 0.00000000000001)) * 10 / FCautionStep) / (10 / FCautionStep);
            end;
            AHint := vgrLoadStr(svgrid__vgr_WorkbookPreview__Hint__Footer);
            AHint := AHint + ' ' + FloatToStr(FMarginNewSize) + ' ' + FMeasurementUnits;
            ShowHintForMargin(AHint, 5, FBegPointToScreen);
          end;

        end;
      end;
    4: // Hand Tool action:
      begin
        FParentWPPreview.Cursor := FHandTool2Cursor;
        FParentWPPreview.VertScrollBar.Position := FParentWPPreview.VertScrollBar.Position + FBegPoint.Y-APoint.Y;
        FParentWPPreview.HorzScrollBar.Position := FParentWPPreview.HorzScrollBar.Position + FBegPoint.X-APoint.X;
        FBegPoint := APoint;
        FBegPointToScreen := APointToScreen;
        FParentWPPreview.Repaint;
      end;
    5,6: // Zoom In Tool and Zoom Out Tool actions:
      begin
        DrawFocusRect(self.FParentWPPreview.Canvas.Handle,FPrevMarginRect);
        FPrevMarginRect := Rect(FBegPoint.X,FBegPoint.Y,APoint.X,APoint.Y);
        if(APoint.X<FBegPoint.X) then
        begin
          FPrevMarginRect.Left := APoint.X;
          FPrevMarginRect.Right := FBegPoint.X;
        end;
        if(APoint.Y<FBegPoint.Y) then
        begin
          FPrevMarginRect.Top := APoint.Y;
          FPrevMarginRect.Bottom := FBegPoint.Y;
        end;
        DrawFocusRect(self.FParentWPPreview.Canvas.Handle,FPrevMarginRect);
      end;
  end;
end;

procedure TvgrWPMarginsResizer.ShowHintForMargin(AHint: String; AMargin: Integer; APoint: TPoint);
var
  ARect: TRect;
  AXPlus, AYPlus: Integer;
  ACursor: TCursor;  
begin
  ACursor := crDefault;
  AXPlus := 0;
  AYPlus := 0;
  case AMargin of
    0:
      begin
        ACursor := crSizeWE;
        AXPlus := -128;
        AYPlus := -50;
      end;
    1:
      begin
        ACursor := crSizeWE;
        AXPlus := 8;
        AYPlus := -50;
      end;
    2:
      begin
        ACursor := crSizeNS;
        AXPlus := -55;
        AYPlus := -50;
      end;
    3:
      begin
        ACursor := crSizeNS;
        AXPlus := -70;
        AYPlus := 25;
      end;
    4://header:
      begin
        ACursor := crSizeNS;
        AXPlus := -55;
        AYPlus := -50;
      end;
    5://footer
      begin
        ACursor := crSizeNS;
        AXPlus := -55;
        AYPlus := 25;
      end;
  end;
  ARect := FHintWindow.CalcHintRect(500, AHint, nil);
  ARect.Left := ARect.Left + APoint.X + AXPlus;
  ARect.Right := ARect.Right + APoint.X + AXPlus;
  ARect.Top := ARect.Top + APoint.Y + AYPlus;
  ARect.Bottom := ARect.Bottom + APoint.Y + AYPlus;
  FHintWindow.ActivateHint(ARect, AHint);
  FParentWPPreview.Cursor := ACursor;
end;

procedure TvgrWPMarginsResizer.SetCurrentCursor;
var
  NewCursor: TCursor;
  APoint: TPoint;
begin
  APoint := FParentWPPreview.ScreenToClient(Mouse.CursorPos);
  NewCursor := FParentWPPreview.Cursor;
  case FParentWPPreview.WorkbookPreview.ActionMode of
    vgrpamHand:
    begin
      if FCurrentAction = 4 then
        NewCursor := FHandTool2Cursor
      else
        NewCursor := FHandToolCursor;
    end;
  vgrpamZoomIn, vgrpamZoomOut:
    begin
      if(PointInSelRect(APoint, FActivePageRect) = true) then
      begin
        case FParentWPPreview.WorkbookPreview.ActionMode of
        vgrpamZoomIn:
          NewCursor := FZoomInCursor;
        vgrpamZoomOut:
          NewCursor := FZoomOutCursor;
        end;
      end
      else
        NewCursor := crDefault;
    end;
  else
    NewCursor := crDefault;
  end;
  if FParentWPPreview.Cursor <> NewCursor then
    FParentWPPreview.Cursor := NewCursor;
end;

procedure TvgrWPMarginsResizer.SetPageSpot(APageRect: TRect; APageNamber: Integer);
var
  I: Integer;
begin
  I := Length(FPagesRects);
  SetLength(FPagesRects, I+1);
  SetLength(FPagesRectsNumbers, I+1);
  FPagesRects[I] := APageRect;
  FPagesRectsNumbers[I] := APageNamber;
  if APageNamber = FParentWPPreview.FActivePage then
  begin
    FActivePageRect := APageRect;
  end;
end;

procedure TvgrWPMarginsResizer.SetSpots(ALeft, ARight, ATop, ABottom, ABLeft,
                                        ABRight, ABTop, ABBottom, AHeader, AFooter: Integer);
begin
  FCLeft := ALeft;  FCRight := ARight;  FCTop := ATop;  FCBottom := ABottom;
  FCBLeft := ABLeft;  FCBRight := ABRight;  FCBTop := ABTop;  FCBBottom := ABBottom;
  FCHeader := ATop + AHeader; FCFooter := ABottom - AFooter;
  //left margin:
  FSpotLeftMargin.Left := ALeft - FSpotWidth div 2;
  FSpotLeftMargin.Right := ALeft + FSpotWidth div 2;
  FSpotLeftMargin.Top := ABTop;
  FSpotLeftMargin.Bottom := ABBottom;
  //right margin:
  FSpotRightMargin.Left := ARight - FSpotWidth div 2;
  FSpotRightMargin.Right := ARight + FSpotWidth div 2;
  FSpotRightMargin.Top := ABTop;
  FSpotRightMargin.Bottom := ABBottom;
  //top margin:
  FSpotTopMargin.Left := ABLeft;
  FSpotTopMargin.Right := ABRight;
  FSpotTopMargin.Top := ATop - FSpotWidth div 2;
  FSpotTopMargin.Bottom := ATop + FSpotWidth div 2;
  //bottom margin:
  FSpotBottomMargin.Left := ABLeft;
  FSpotBottomMargin.Right := ABRight;
  FSpotBottomMargin.Top := ABottom - FSpotWidth div 2;
  FSpotBottomMargin.Bottom := ABottom + FSpotWidth div 2;
  //Header:
  FSpotHeader.Left := ABLeft;
  FSpotHeader.Right := ABRight;
  FSpotHeader.Top := FCHeader - FSpotWidth div 2;
  FSpotHeader.Bottom := FCHeader + FSpotWidth div 2;
  //Footer:
  FSpotFooter.Left := ABLeft;
  FSpotFooter.Right := ABRight;
  FSpotFooter.Top := FCFooter - FSpotWidth div 2;
  FSpotFooter.Bottom := FCFooter + FSpotWidth div 2;
end;

procedure TvgrWPMarginsResizer.GetMarginsSize;
begin
  if FParentWPPreview.PageMaker.PageProperties.MeasurementSystem = vgrmsMetric then
  begin
    FMeasurementUnits := vgrLoadStr(svgrid_Common_Mm);
    FMesUnits := vgruMms;
    FCautionStep := 1;
  end
  else
  begin
    FMeasurementUnits := vgrLoadStr(svgrid_Common_Mm);
    FMesUnits := vgruInches;
    FCautionStep := 0.1;
  end;
  FLeft := Round(FParentWPPreview.PageMaker.PageProperties.Margins.UnitsLeft[FMesUnits]*10/FCautionStep)/(10/FCautionStep);
  FRight := Round(FParentWPPreview.PageMaker.PageProperties.Margins.UnitsRight[FMesUnits]*10/FCautionStep)/(10/FCautionStep);
  FTop := Round(FParentWPPreview.PageMaker.PageProperties.Margins.UnitsTop[FMesUnits]*10/FCautionStep)/(10/FCautionStep);
  FBottom := Round(FParentWPPreview.PageMaker.PageProperties.Margins.UnitsBottom[FMesUnits]*10/FCautionStep)/(10/FCautionStep);
  FHeader := Round(FParentWPPreview.PageMaker.PageProperties.Header.UnitsHeight[FMesUnits]*10/FCautionStep)/(10/FCautionStep);
  FFooter := Round(FParentWPPreview.PageMaker.PageProperties.Footer.UnitsHeight[FMesUnits]*10/FCautionStep)/(10/FCautionStep);
end;

/////////////////////////////////////////////////
//
// TvgrWPNavigatorPointInfo
//
/////////////////////////////////////////////////
constructor TvgrWPNavigatorPointInfo.Create(AButton: TvgrWPNavigatorButton);
begin
  inherited Create;
  FButton := AButton;
end;

/////////////////////////////////////////////////
//
// TvgrWPNavigator
//
/////////////////////////////////////////////////
constructor TvgrWPNavigator.Create(AOwner: TComponent);
begin
  inherited;
  FFirstPage := TvgrSheetsPanelButtonRect.CreateButton(Self, FPreviewNavigatorBitmaps, 0, OnButtonClick);
  FPriorPage := TvgrSheetsPanelButtonRect.CreateButton(Self, FPreviewNavigatorBitmaps, 1, OnButtonClick);
  FNextPage := TvgrSheetsPanelButtonRect.CreateButton(Self, FPreviewNavigatorBitmaps, 2, OnButtonClick);
  FLastPage := TvgrSheetsPanelButtonRect.CreateButton(Self, FPreviewNavigatorBitmaps, 3, OnButtonClick);
end;

destructor TvgrWPNavigator.Destroy;
begin
  FPriorPage.Free;
  FNextPage.Free;
  FFirstPage.Free;
  FLastPage.Free;
  inherited;
end;

function TvgrWPNavigator.GetButtonEnabled(AButtonIndex: Integer): Boolean;
begin
  Result := Workbook <> nil;
  if Result then
    case AButtonIndex of
      0: Result := ActivePage > 0; { First }
      1: Result := ActivePage > 0; { Prior }
      2: Result := ActivePage < PageCount - 1; { Next }
      3: Result := ActivePage < PageCount - 1; { Last }
    end;
end;

function TvgrWPNavigator.GetCurRect: TvgrSheetsPanelRect;
begin
  Result := FCurRect;
end;

procedure TvgrWPNavigator.SetCurRect(Value: TvgrSheetsPanelRect);
begin
  FCurRect := Value;
end;

function TvgrWPNavigator.GetWorkbookPreview: TvgrWorkbookPreview;
begin
  Result := TvgrWorkbookPreview(Owner);
end;

function TvgrWPNavigator.GetWorkbook: TvgrWorkbook;
begin
  Result := WorkbookPreview.Workbook;
end;

function TvgrWPNavigator.GetActivePage: Integer;
begin
  Result := WorkbookPreview.FWPPreview.ActivePage;
end;

procedure TvgrWPNavigator.SetActivePage(Value: Integer);
begin
  WorkbookPreview.FWPPreview.ActivePage := Value;
end;

function TvgrWPNavigator.GetPageCount: Integer;
begin
  Result := WorkbookPreview.FWPPreview.PageCount;
end;

procedure TvgrWPNavigator.OnButtonClick(Sender: TObject);
begin
  if GetButtonEnabled(TvgrSheetsPanelButtonRect(Sender).ImageIndex) then
  begin
    if Sender = FFirstPage then
      ActivePage := 0
    else
      if Sender = FPriorPage then
        ActivePage := ActivePage - 1
      else
        if Sender = FNextPage then
          ActivePage := ActivePage + 1
        else
          if Sender = FLastPage then
            ActivePage := PageCount - 1;
    Invalidate;
  end;
end;

procedure TvgrWPNavigator.Resize;
var
  ALeft, AWidth: Integer;
begin
  ALeft := ClientRect.Left + 1;
  AWidth := ClientWidth - 2;
  FFirstPage.SetBounds(ALeft, ClientRect.Top + 1, AWidth, AWidth);
  FPriorPage.SetBounds(ALeft, FFirstPage.BoundsRect.Bottom + 2, AWidth, AWidth);
  FNextPage.SetBounds(ALeft, FPriorPage.BoundsRect.Bottom + 2, AWidth, AWidth);
  FLastPage.SetBounds(ALeft, FNextPage.BoundsRect.Bottom + 2, AWidth, AWidth);
end;

procedure TvgrWPNavigator.CMMouseLeave(var Msg: TMessage);
begin
  if FCurRect <> nil then
  begin
    FCurRect.MouseOver := False;
    FCurRect := nil;
  end;
end;

procedure TvgrWPNavigator.WMEraseBkgnd(var Msg : TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWPNavigator.Paint;
begin
  FFirstPage.Paint(Canvas);
  FNextPage.Paint(Canvas);
  FLastPage.Paint(Canvas);
  FPriorPage.Paint(Canvas);

  Canvas.Brush.Style := bsSolid;
  Canvas.Brush.Color := clBtnFace; 
  with ClientRect do
    Canvas.FillRect(Rect(1, 1, Right - 1, Bottom - 1));
  DrawFrame(Canvas, ClientRect, clBtnShadow);
end;

function TvgrWPNavigator.GetPanelRectAt(X, Y: Integer): TvgrSheetsPanelRect;
var
  APoint: TPoint;
begin
  APoint := Point(X, Y);
  if PtInRect(FPriorPage.BoundsRect, APoint) then
    Result := FPriorPage
  else
    if PtInRect(FNextPage.BoundsRect, APoint) then
      Result := FNextPage
    else
      if PtInRect(FFirstPage.BoundsRect, APoint) then
        Result := FFirstPage
      else
        if PtInRect(FLastPage.BoundsRect, APoint) then
          Result := FLastPage
        else
          Result := nil;
end;

procedure TvgrWPNavigator.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
var
  APanelRect: TvgrSheetsPanelRect;
begin
  APanelRect := GetPanelRectAt(APoint.X, APoint.Y);
  if APanelRect = FFirstPage then
    APointInfo := TvgrWPNavigatorPointInfo.Create(vgrnbFirstPage)
  else
    if APanelRect = FPriorPage then
      APointInfo := TvgrWPNavigatorPointInfo.Create(vgrnbPriorPage)
    else
      if APanelRect = FNextPage then
        APointInfo := TvgrWPNavigatorPointInfo.Create(vgrnbNextPage)
      else
        if APanelRect = FLastPage then
          APointInfo := TvgrWPNavigatorPointInfo.Create(vgrnbLastPage)
        else
          APointInfo := TvgrWPNavigatorPointInfo.Create(vgrnbNone);
end;

procedure TvgrWPNavigator.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APanelRect: TvgrSheetsPanelRect;
  AExecuteDefault: Boolean;
begin
  WorkbookPreview.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    APanelRect := GetPanelRectAt(X, Y);
    if APanelRect <> nil then
      APanelRect.Click;
  end;
end;

procedure TvgrWPNavigator.MouseMove(Shift: TShiftState; X, Y: Integer);
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
  WorkbookPreview.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrWPNavigator.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  WorkbookPreview.DoMouseUp(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrWPNavigator.DblClick;
var
  AExecuteDefault: Boolean;
begin
  WorkbookPreview.DoDblClick(Self, AExecuteDefault);
end;

/////////////////////////////////////////////////
//
// TvgrWPPreview
//
/////////////////////////////////////////////////
constructor TvgrWPPreview.Create(AOwner: TComponent);
begin
  inherited;
  FActivePage := -1;
  FCache := TvgrWPImagesCache.Create;
  FMarginsResizer := TvgrWPMarginsResizer.Create(self);
  FPainter := TvgrWorkbookPainter.Create(Self);
  FPageVSpacing := 9;
  FPageHSpacing := 9;
  FScaleMode := vgrpsmPageWidth;
  FScalePercent := 100;
  PageColCount := 3;
  PageRowCount := 3;
end;

destructor TvgrWPPreview.Destroy;
begin
  FreeAndNil(FPainter);
  FCache.Free;
  FMarginsResizer.Free;
  inherited;
end;



function TvgrWPPreview.GetWorkbookPreview: TvgrWorkbookPreview;
begin
  Result := TvgrWorkbookPreview(Parent);
end;

function TvgrWPPreview.GetWorksheet: TvgrWorksheet;
begin
  Result := WorkbookPreview.ActiveWorksheet;
end;

function TvgrWPPreview.GetPageMaker: TvgrPageMaker;
begin
  if Worksheet = nil then
    Result := nil
  else
    Result := WorkbookPreview.FindPageMaker(Worksheet);
end;

function TvgrWPPreview.GetHorzScrollBar: TvgrScrollBar;
begin
  Result := WorkbookPreview.HorzScrollBar;
end;

function TvgrWPPreview.GetVertScrollBar: TvgrScrollBar;
begin
  Result := WorkbookPreview.VertScrollBar;
end;

function TvgrWPPreview.GetActivePage: Integer;
begin
  Result := FActivePage;
end;

procedure TvgrWPPreview.ScrollToActivePage;
var
  ANewPos, ATmpInteger,I,J: Integer;
begin
  J := 0;
  if(PageRowCount>1) then
  begin
    for I:=0 to Length(FMarginsResizer.FPagesRects)-1 do
    begin
      if (FMarginsResizer.FPagesRectsNumbers[I]=FActivePage) then
      begin
        if FMarginsResizer.FPagesRects[I].Bottom<(FMarginsResizer.FPagesRects[I].Bottom-FMarginsResizer.FPagesRects[I].Top)/2 then
        begin
          ATmpInteger := (FMarginsResizer.FPagesRects[I].Bottom-FMarginsResizer.FPagesRects[I].Top)-FMarginsResizer.FPagesRects[I].Bottom;
          VertScrollBar.SetParams(VertScrollBar.Position-ATmpInteger,VertScrollBar.Min,VertScrollBar.Max,VertScrollBar.PageSize);
        end
        else
        begin
          if FMarginsResizer.FPagesRects[I].Top>ClientHeight-(FMarginsResizer.FPagesRects[I].Bottom-FMarginsResizer.FPagesRects[I].Top)/2 then
          begin
            ATmpInteger := FMarginsResizer.FPagesRects[I].Top-(ClientHeight-(FMarginsResizer.FPagesRects[I].Bottom-FMarginsResizer.FPagesRects[I].Top));
            VertScrollBar.SetParams(VertScrollBar.Position+ATmpInteger,VertScrollBar.Min,VertScrollBar.Max,VertScrollBar.PageSize);
          end;
        end;
        J := 1;
      end;
    end;//for I:=0 to
  end;//if(PageRowCount>1) then

  if (J<>1) then
  begin
    //Calculate new scroll bar position:
    ATmpInteger := round(int(FActivePage/PageColCount));
    ANewPos :=(FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing)
               * ATmpInteger + FFullPageRect.Top-cGraySize
               -Round(PageRowCount/2)*(FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing);
    VertScrollBar.SetParams(ANewPos,VertScrollBar.Min,VertScrollBar.Max,VertScrollBar.PageSize);
  end;

end;

procedure TvgrWPPreview.SetActivePage(Value: Integer);
begin
  if FActivePage <> Value then
  begin
    FActivePage := Value;
    if FActivePage < 0 then
      FActivePage := 0;
    if FActivePage > PageCount then
      FActivePage := PageCount - 1;
    ScrollToActivePage;
    Repaint;
    WorkbookPreview.DoActivePageChanged;
  end;
end;

function TvgrWPPreview.GetPageCount: Integer;
begin
  if PageMaker = nil then
    Result := -1
  else
    Result := PageMaker.PagesCount;
end;

procedure TvgrWPPreview.SetScaleMode(Value: TvgrPreviewScaleMode);
begin
  if FScaleMode <> Value then
  begin
    FScaleMode := Value;
    UpdateSizesWithEvent;
    UpdateScrollBars;
    Repaint;
  end;
end;

function TvgrWPPreview.GetScalePercent: Double;
begin
  Result := FScalePercent;
end;

procedure TvgrWPPreview.SetScalePercent(Value: Double);
begin
  if Value < cMinScalePercent then
    Value := cMinScalePercent;
  if Value > cMaxScalePercent then
    Value := cMaxScalePercent;

  if not IsZero(FScalePercent - Value) then
  begin
    FScalePercent := Value;
    if FScaleMode = vgrpsmPercent then
    begin
      UpdateSizes;
      UpdateScrollBars;
      Repaint;
      WorkbookPreview.DoScalePercentChanged;
    end;
  end;
end;

procedure TvgrWPPreview.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWPPreview.WMGetDlgCode(var Msg: TWMGetDlgCode);
begin
  inherited;
  Msg.Result := Msg.Result or DLGC_WANTALLKEYS or DLGC_WANTARROWS;
end;

procedure TvgrWPPreview.WMMouseWheel(var Msg: TWMMouseWheel);
var
  wParam: Cardinal;
begin
  if (Msg.Keys and MK_SHIFT) <> 0 then
    begin
      if Msg.WheelDelta>0 then
        wParam := SB_PAGELEFT
      else
        wParam := SB_PAGERIGHT;
    end
  else
    begin
      if Msg.WheelDelta > 0 then
        wParam := SB_LINELEFT
      else
        wParam := SB_LINERIGHT;
    end;

  if (Msg.Keys and MK_CONTROL) <> 0 then
    PostMessage(HorzScrollBar.Handle, CN_HSCROLL, wParam, 0)
  else
    PostMessage(VertScrollBar.Handle, CN_VSCROLL, wParam, 0);

  inherited;
end;

function TvgrWPPreview.GetDefaultColumnWidth: Integer;
begin
  if Worksheet <> nil then
    Result := Worksheet.PageProperties.Defaults.ColWidth
  else
    Result := cDefaultColWidthTwips;
end;

function TvgrWPPreview.GetDefaultRowHeight: Integer;
begin
  if Worksheet <> nil then
    Result := Worksheet.PageProperties.Defaults.RowHeight
  else
    Result := cDefaultRowHeightTwips;
end;

function TvgrWPPreview.GetPixelsPerInchX: Integer;
begin
  Result := ScreenPixelsPerInchX;
end;

function TvgrWPPreview.GetPixelsPerInchY: Integer;
begin
  Result := ScreenPixelsPerInchY;
end;

function TvgrWPPreview.GetBackgroundColor: TColor;
begin
  Result := WorkbookPreview.WorkbookBackColor;
end;

function TvgrWPPreview.GetDefaultRangeBackgroundColor: TColor;
begin
  Result := WorkbookPreview.WorkbookBackColor;
end;

function TvgrWPPreview.GetImageForPage(APageNumber: Integer):TMetafile;
var
  ACanvas: TMetafileCanvas;
  FImage: TMetafile;
begin
  Painter.Worksheet := Worksheet;
  FImage := FCache.getImage(APageNumber);
  if FImage = nil then
  begin
    FImage := TMetafile.Create;
    FImage.Width := Painter.GetPagePixelWidth(PageMaker.Pages[APageNumber], 1);
    FImage.Height := Painter.GetPagePixelHeight(PageMaker.Pages[APageNumber], 1);
    ACanvas := TMetafileCanvas.Create(FImage, 0);
    try
      Painter.DrawPage(0, 0, ACanvas, PageMaker.Pages[APageNumber], 1, Point(-WorkbookPreview.LeftOffset, -WorkbookPreview.TopOffset));
    finally
      ACanvas.Free;
    end;
    FCache.setImage(APageNumber,FImage);
  end;
  Result := FImage;

{$IFDEF VGR_UPDATEIMAGE_TEST}
    // for test only
    FImage.SaveToFile('c:\test.emf');
{$ENDIF}
end;

procedure TvgrWPPreview.BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
begin
end;

procedure TvgrWPPreview.AfterChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcNewCol, vgrwcNewRow, vgrwcNewHorzSection, vgrwcNewVertSection, vgrwcDeleteCol,
    vgrwcDeleteRow, vgrwcChangeCol, vgrwcChangeRow, vgrwcChangeHorzSection,
    vgrwcChangeVertSection, vgrwcDeleteHorzSection, vgrwcDeleteVertSection,
    vgrwcChangeWorksheetContent:
    begin
      // number of pages may be changed
      UpdatePages;
      UpdateScrollBars;
    end;
    vgrwcChangeWorksheetPageProperties, vgrwcChangeWorksheet, vgrwcUpdateAll:
    begin
      // size of page may be changed and number of pages.
      UpdateSizesWithEvent;
      UpdatePages;
      UpdateScrollBars;
    end;
    vgrwcNewRange, vgrwcNewBorder, vgrwcDeleteRange, vgrwcDeleteBorder,
    vgrwcChangeRange, vgrwcChangeBorder:
    begin
      // content of the pages may be changed.
//      UpdateImage;
    end;
  end;
  FCache.ClearCache;
  Repaint;
end;

procedure TvgrWPPreview.Resize;
begin
  UpdateSizesWithEvent;
  UpdateScrollBars;
  Repaint;
end;

procedure TvgrWPPreview.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  SetFocus;
  if Button = mbLeft then
    FMarginsResizer.DoMouseDown(Point(X, Y),ClientToScreen(Point(X, Y)));
  WorkbookPreview.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrWPPreview.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APointInfo: TvgrPointInfo;
  AExecuteDefault: Boolean;
begin
  if Button = mbLeft then
    FMarginsResizer.DoMouseUp(Point(X, Y),ClientToScreen(Point(X, Y)));

  WorkbookPreview.DoMouseUp(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    if Button = mbRight then
    begin
      GetPointInfoAt(Point(X, Y), APointInfo);
      WorkbookPreview.ShowPopupMenu(ClientToScreen(Point(X, Y)), APointInfo);
      APointInfo.Free;
    end;
  end;
end;

procedure TvgrWPPreview.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  FMarginsResizer.DoMouseMove(Point(X, Y), ClientToScreen(Point(X, Y)), Shift);
  WorkbookPreview.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrWPPreview.DblClick;
var
  AExecuteDefault: Boolean;
begin
  WorkbookPreview.DoDblClick(Self, AExecuteDefault);
end;

procedure TvgrWPPreview.Paint;
var
  ARect: TRect;
  ARectForBorders: TRect;
  I, N: Integer;

  procedure DrawFrame(ACanvas: TCanvas; const ARect: TRect; AColor: TColor);
  begin
    vgr_GUIFunctions.DrawFrame(ACanvas, ARect, AColor);
    ExcludeClipRect(ACanvas, ARect);
  end;

begin
  if FActivePage = -1 then
  begin
    Canvas.Brush.Color := WorkbookPreview.BackColor;
    Canvas.FillRect(ClientRect);
  end
  else
  begin
    ARect := FPageRect;
    OffsetRect(ARect, -WorkbookPreview.LeftOffset, -WorkbookPreview.TopOffset);
    FMarginsResizer.ClearPagesSpots;
    // Draw pages
    N := 1;
    for I:=0 to PageMaker.PagesCount-1 do
    begin
      //Calculate rect for current page if this not first page:
      if I<>0 then
      begin
        {Step right}
        if N < FPageColCount then
        begin
          ARect.Left := ARect.Left + (FFullPageRect.Right - FFullPageRect.Left + FPageHSpacing);
          ARect.Right := ARect.Right + (FFullPageRect.Right - FFullPageRect.Left + FPageHSpacing);
          N := N + 1;
        end
        else
        begin
          {Enter to a new line}
          ARect.Top := ARect.Top + (FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing);
          ARect.Bottom := ARect.Bottom  + (FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing);
          if FPageColCount>1 then
          begin
            {move to left}
            ARect.Left := ARect.Left - (FFullPageRect.Right - FFullPageRect.Left + FPageHSpacing)* (N - 1);
            ARect.Right := ARect.Right - (FFullPageRect.Right - FFullPageRect.Left + FPageHSpacing)* (N - 1);
          end;
          N := 1;
        end; // else if N < FPageColCount then
     end; // if I<>0 then
     //calculate rect for borders of page:
     ARectForBorders := FFullPageRect;
     ARectForBorders.Left := ARect.Left - (FPageRect.Left - FFullPageRect.Left);
     ARectForBorders.Right := ARect.Right + (FFullPageRect.Right - FPageRect.Right);
     ARectForBorders.Top := ARect.Top - (FPageRect.Top - FFullPageRect.Top);
     ARectForBorders.Bottom := ARect.Bottom + (FFullPageRect.Bottom - FPageRect.Bottom);
     //Check it is necessary draw page
     if (ARectForBorders.Bottom > 0) and (ARectForBorders.Top < ClientRect.Bottom) and (ARectForBorders.Left<ClientRect.Right) and (ARectForBorders.Right > 0) then
     begin
       //Sets pages spots:
       FMarginsResizer.SetPageSpot(ARectForBorders,I);
       //Margins Resizer lines:
       if FActivePage = I then
       begin
         FMarginsResizer.SetSpots(ARect.Left,ARect.Right,ARect.Top,ARect.Bottom,
                                  ARectForBorders.Left,ARectForBorders.Right,ARectForBorders.Top,
                                  ARectForBorders.Bottom,
                                  Round((PageMaker.PageProperties.Header.UnitsHeight[vgruPixels] * ScalePercent) / 100),
                                  Round((PageMaker.PageProperties.Footer.UnitsHeight[vgruPixels] * ScalePercent) / 100));
         FMarginsResizer.DrawMarginsResizerLines;
       end; //if FActivePage = I then
        //Draw Page:
        with ARect do
          Canvas.StretchDraw(Rect(Left, Top , Right + 1, Bottom + 1), GetImageForPage(i)); // right + 1 and bottom + 1 !!!
        ExcludeClipRect(Canvas, ARect);
        // Draw margins color
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := clWhite;
        Canvas.FillRect(ARectForBorders);
        if FActivePage = I then
        begin
          DrawFrame(Canvas, ARectForBorders, clRed);
        end
        else
        begin
          DrawFrame(Canvas, ARectForBorders, clWindowFrame);
          // Draw shadow
          with ARectForBorders do
          begin
            Canvas.FillRect(Rect(Right, Top + 1, Right + 1, Bottom + 1));
            ExcludeClipRect(Canvas, Rect(Right, Top + 1, Right + 1, Bottom + 1));
            Canvas.FillRect(Rect(Left + 1, Bottom, Right + 1, Bottom + 1));
            ExcludeClipRect(Canvas, Rect(Left + 1, Bottom, Right + 1, Bottom + 1));
          end;
        end; //end of else if FActivePage = I
      end;//end of check it is necessary draw page
    end;
    // fill client rect
    Canvas.Brush.Color := WorkbookPreview.BackColor;
    Canvas.FillRect(ClientRect);
  end;
end;

procedure TvgrWPPreview.UpdatePages;
var
  AOldPagesCount: Integer;
begin
  if Worksheet <> nil then
  begin
    FCache.ClearCache;
    AOldPagesCount := PageMaker.PagesCount;
    PageMaker.Prepare;
    if PageMaker.PagesCount <> AOldPagesCount then
      WorkbookPreview.DoActiveWorksheetPagesCountChanged;
    if PageCount > 0 then
    begin
      if (FActivePage < 0) or (FActivePage >= PageCount) then
        FActivePage := 0;
    end
    else
      FActivePage := -1;
  end
  else
    FActivePage := -1;
//  UpdateImage;
end;

procedure TvgrWPPreview.UpdateSizesWithEvent;
var
  AOldScalePercent: Double;
begin
  AOldScalePercent := FScalePercent;
  UpdateSizes;
  if not IsZero(FScalePercent - AOldScalePercent) then
    WorkbookPreview.DoScalePercentChanged;
end;

procedure TvgrWPPreview.UpdateSizes;
var
  AWidth, AHeight: Integer;
  APageWidth, APageHeight: Integer;
  AKoef: Double;
begin
  if Worksheet <> nil then
  begin
    // calculate page rectangles
    APageWidth := ConvertTwipsToPixelsX(Worksheet.PageProperties.Width);
    APageHeight := ConvertTwipsToPixelsY(Worksheet.PageProperties.Height);
    case FScaleMode of
      vgrpsmPercent:
        begin
          AWidth := Round(APageWidth * ScalePercent / 100);
          AHeight := Round(APageHeight * ScalePercent / 100);
          AKoef := ScalePercent / 100;
        end;
      vgrpsmWholePage:
        begin
          if (ClientWidth - cGraySize - cGraySize) / APageWidth  <
             (ClientHeight - cGraySize - cGraySize) / APageHeight then
            AKoef := (ClientWidth - cGraySize - cGraySize) / APageWidth
          else
            AKoef := (ClientHeight - cGraySize - cGraySize) / APageHeight;
          AWidth := Round(APageWidth * AKoef);
          AHeight := Round(APageHeight * AKoef);
          FScalePercent := AKoef * 100;
        end;
      vgrpsmTwoPages:
        begin
          if (ClientWidth - cGraySize - cGraySize) / (APageWidth * 2)  <
             (ClientHeight - cGraySize - cGraySize) / APageHeight then
            AKoef := (ClientWidth - cGraySize - cGraySize) / APageWidth/2
          else
            AKoef := (ClientHeight - cGraySize - cGraySize) / APageHeight;
          AWidth := Round(APageWidth * AKoef);
          AHeight := Round(APageHeight * AKoef);
          FScalePercent := AKoef * 100;
        end;
      vgrpsmMultiPages:
        begin
          if (ClientWidth - cGraySize - cGraySize) / (APageWidth * FPageColCount+0.0000000000000001)  <
             (ClientHeight - cGraySize - cGraySize) / (APageHeight * FPageRowCount+0.0000000000000001) then
            AKoef := (ClientWidth - cGraySize - cGraySize) / (APageWidth * FPageColCount)
          else
            AKoef := (ClientHeight - cGraySize - cGraySize) / (APageHeight * FPageRowCount);
          AWidth := Round(APageWidth * AKoef);
          AHeight := Round(APageHeight * AKoef);
          FScalePercent := AKoef * 100;
        end;        
    else
      begin
        AKoef := (ClientWidth - cGraySize - cGraySize) / APageWidth;
        FScalePercent := AKoef * 100;
        AWidth := Round(APageWidth * AKoef);
        AHeight := Round(APageHeight * AKoef);
      end;
    end;
    FPreviewAreaWidth := AWidth + cGraySize + cGraySize;
    FPreviewAreaHeight := AHeight + cGraySize + cGraySize;

    if (FScaleMode <> vgrpsmPageWidth) and (FScaleMode <> vgrpsmWholePage) and (FScaleMode <> vgrpsmMultiPages) then
    begin
      //Calculate amount of the pages on one line:
      PageColCount := Round(ClientWidth / (AWidth + FPageHSpacing));
      if (ClientWidth / (AWidth + FPageHSpacing)) < PageColCount then
        PageColCount := PageColCount - 1;
      if PageColCount < 1 then
        PageColCount := 1;
      if FScaleMode = vgrpsmTwoPages then
        PageColCount := 2;
    end
    else
    begin
      if (FScaleMode <> vgrpsmMultiPages) then
        PageColCount := 1;
    end;

    if (FScaleMode <> vgrpsmMultiPages) then
      PageRowCount := Round(Int(ClientHeight / AHeight));

    with FFullPageRect do
    begin
      if ClientWidth >= FPreviewAreaWidth then
        Left := Round((ClientWidth - AWidth * PageColCount - FPageHSpacing * (PageColCount - 1)) / 2)
      else
        Left := cGraySize;
      if ClientHeight >= FPreviewAreaHeight then
      begin
        if self.PageCount > 1 then
          Top := cGraySize
        else
          Top := (ClientHeight - AHeight) div 2;
      end
      else
        Top := cGraySize;
      Right := Left + AWidth;
      Bottom := Top + AHeight;
    end;

    with FPageRect do
    begin
      Left := FFullPageRect.Left + Round(ConvertTwipsToPixelsX(Worksheet.PageProperties.Margins.Left) * AKoef);
      Top := FFullPageRect.Top + Round(ConvertTwipsToPixelsY(Worksheet.PageProperties.Margins.Top) * AKoef);
      Right := FFullPageRect.Right - Round(ConvertTwipsToPixelsX(Worksheet.PageProperties.Margins.Right) * AKoef);
      Bottom := FFullPageRect.Bottom - Round(ConvertTwipsToPixelsY(Worksheet.PageProperties.Margins.Bottom) * AKoef);
    end;
  end;
end;

procedure TvgrWPPreview.UpdateScrollBars;
var
  ATmpInteger, AOldMax, AOldPosition, ANewMax, ANewPosition, ANewPageSize: Integer;
begin
  if (Worksheet <> nil) and (ActivePage >= 0) then
  begin
    //* * * VERTICAL SCROLLBAR * * *//
    //Calculate new Max for vertical scroll bar:
    ATmpInteger := Round(PageMaker.PagesCount/PageColCount);
    if (PageMaker.PagesCount/PageColCount)>ATmpInteger then
      ATmpInteger := ATmpInteger + 1 ;
    ANewMax := ((FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing) * ATmpInteger + cGraySize*2 + FFullPageRect.Top);
    //New vertical page size:
    ANewPageSize := ClientHeight;
    //calculate new vertical position:
    if FMarginsResizer.NeedElaborateScroll = True then
    begin
      ATmpInteger := round(int(FActivePage/PageColCount));
      ANewPosition := (FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing)
                       * ATmpInteger + FFullPageRect.Top
                       - Round(PageRowCount / 2) * (FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing);//active page top
      ANewPosition := ANewPosition +  Round((FMarginsResizer.NeedElaborateScroll_Rect.Top
                                       - FMarginsResizer.NeedElaborateScroll_ActivePageRect.Top)
                                       * (FFullPageRect.Bottom - FFullPageRect.Top)
                                       / (FMarginsResizer.NeedElaborateScroll_ActivePageRect.Bottom
                                       - FMarginsResizer.NeedElaborateScroll_ActivePageRect.Top + 0.00000000000000000000000000001))
                                       - Round((ClientHeight - (FMarginsResizer.NeedElaborateScroll_Rect.Bottom - FMarginsResizer.NeedElaborateScroll_Rect.Top)
                                       * (FFullPageRect.Bottom - FFullPageRect.Top)
                                       / (FMarginsResizer.NeedElaborateScroll_ActivePageRect.Bottom
                                       - FMarginsResizer.NeedElaborateScroll_ActivePageRect.Top + 0.00000000000000000000000000001))/2) ;
    end
    else
    begin
      ATmpInteger := round(int(FActivePage/PageColCount));
      ANewPosition :=(FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing)
             * ATmpInteger + FFullPageRect.Top
             -Round(PageRowCount/2)*(FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing);
      if (FMarginsResizer.FCurrentAction = 5) or (FMarginsResizer.FCurrentAction = 6) then
      begin
        ANewPosition := ANewPosition + round(
        (FMarginsResizer.NeedElaborateScroll_Rect.Top - FMarginsResizer.NeedElaborateScroll_ActivePageRect.Top
          - FMarginsResizer.NeedElaborateScroll_Rect.Top)
        *
        (FFullPageRect.Bottom - FFullPageRect.Top)
        /
        (FMarginsResizer.NeedElaborateScroll_ActivePageRect.Bottom-FMarginsResizer.NeedElaborateScroll_ActivePageRect.Top+0.00000000000000000000000000001)
        );
        ANewPosition := ANewPosition + ATmpInteger;
      end
      else
      begin
        ATmpInteger := round(int(FActivePage / PageColCount));
        ANewPosition :=(FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing)
                        * ATmpInteger + FFullPageRect.Top - Round(PageRowCount / 2)
                        * (FFullPageRect.Bottom - FFullPageRect.Top + FPageVSpacing) - cGraySize;
      end;
    end;
    VertScrollBar.SetParams(ANewPosition,0,ANewMax,ANewPageSize);
    //* * * HORIZONTAL SCROLLBAR * * *//
    //Remember old parameters:
    AOldMax := HorzScrollBar.Max;
    if AOldMax = 0 then
      AOldMax := 1;
    AOldPosition := HorzScrollBar.Position;
    //New Max for horizontal scroll bar:    
    ANewMax := FPreviewAreaWidth;
    //New horizontal page size:
    ANewPageSize := ClientWidth;
    //calculate new horizontal position:
    if FMarginsResizer.NeedElaborateScroll = True then
    begin
      ANewPosition :=FFullPageRect.Left;
      ANewPosition := ANewPosition + Round(
      (FMarginsResizer.NeedElaborateScroll_Rect.Left - FMarginsResizer.NeedElaborateScroll_ActivePageRect.Left)
      *
      (FFullPageRect.Right - FFullPageRect.Left)
      /
      (FMarginsResizer.NeedElaborateScroll_ActivePageRect.Right-FMarginsResizer.NeedElaborateScroll_ActivePageRect.Left + 0.00000000000000000000000000001)
      )
      -
      Round((ClientWidth - (FMarginsResizer.NeedElaborateScroll_Rect.Right - FMarginsResizer.NeedElaborateScroll_Rect.Left)
      * (FFullPageRect.Right - FFullPageRect.Left)
      / (FMarginsResizer.NeedElaborateScroll_ActivePageRect.Right - FMarginsResizer.NeedElaborateScroll_ActivePageRect.Left + 0.00000000000000000000000000001)
      ) / 2);
    end
    else
    begin
      ANewPosition := Round(ANewMax * AOldPosition / AOldMax);
    end;
    HorzScrollBar.SetParams(ANewPosition,0,ANewMax,ANewPageSize);
  end
  else
  begin
    if HandleAllocated then
    begin
      VertScrollBar.SetParams(0, 0, 0, 0);
      HorzScrollBar.SetParams(0, 0, 0, 0);
    end;
  end;
end;

procedure TvgrWPPreview.UpdateAll;
begin
  UpdateSizesWithEvent;
  UpdatePages;
  UpdateScrollBars;
  Repaint;
end;

procedure TvgrWPPreview.DoScroll(const OldPos, NewPos: TPoint);
begin
  Invalidate;
end;

function TvgrWPPreview.KbdLeft: Boolean;
begin
  WorkbookPreview.DoScroll(WorkbookPreview.Offset, Point(WorkbookPreview.Offset.X - 1, WorkbookPreview.Offset.Y));
  Result := True;
end;

function TvgrWPPreview.KbdRight: Boolean;
begin
  WorkbookPreview.DoScroll(WorkbookPreview.Offset, Point(WorkbookPreview.Offset.X + 1, WorkbookPreview.Offset.Y));
  Result := True;
end;

function TvgrWPPreview.KbdUp: Boolean;
begin
  WorkbookPreview.DoScroll(WorkbookPreview.Offset, Point(WorkbookPreview.Offset.X, WorkbookPreview.Offset.Y - 1));
  Result := True;
end;

function TvgrWPPreview.KbdDown: Boolean;
begin
  WorkbookPreview.DoScroll(WorkbookPreview.Offset, Point(WorkbookPreview.Offset.X, WorkbookPreview.Offset.Y + 1));
  Result := True;
end;

function TvgrWPPreview.KbdPageUp: Boolean;
begin
  with WorkbookPreview do
    DoScroll(Offset, Point(Offset.X, Offset.Y - VertScrollBar.PageSize));
  Result := True;
end;

function TvgrWPPreview.KbdPageDown: Boolean;
begin
  with WorkbookPreview do
    DoScroll(Offset, Point(Offset.X, Offset.Y + VertScrollBar.PageSize));
  Result := True;
end;

function TvgrWPPreview.KbdHome: Boolean;
begin
  with WorkbookPreview do
    DoScroll(Offset, Point(0, 0));
  Result := True;
end;

function TvgrWPPreview.KbdEnd: Boolean;
begin
  with WorkbookPreview do
    DoScroll(Offset, Point(Offset.X - HorzScrollBar.PageSize, Offset.Y + VertScrollBar.PageSize));
  Result := True;
end;

function TvgrWPPreview.KbdPrint: Boolean;
begin
  WorkbookPreview.DoPreviewAction(vgrpaPrint);
  Result := True;
end;

function TvgrWPPreview.KbdEditPageParams: Boolean;
begin
  WorkbookPreview.DoPreviewAction(vgrpaPageParams);
  Result := True;
end;

function TvgrWPPreview.KbdPriorPage: Boolean;
begin
  WorkbookPreview.DoPreviewAction(vgrpaPriorPage);
  Result := True;
end;

function TvgrWPPreview.KbdNextPage: Boolean;
begin
  WorkbookPreview.DoPreviewAction(vgrpaNextPage);
  Result := True;
end;

function TvgrWPPreview.KbdPriorWorksheet: Boolean;
begin
  WorkbookPreview.DoPreviewAction(vgrpaPriorWorksheet);
  Result := True;
end;

function TvgrWPPreview.KbdNextWorksheet: Boolean;
begin
  WorkbookPreview.DoPreviewAction(vgrpaNextWorksheet);
  Result := True;
end;

function TvgrWPPreview.KbdZoomIn: Boolean;
begin
  WorkbookPreview.SetActionMode(vgrpamZoomIn);
  Result := True;
end;

function TvgrWPPreview.KbdZoomOut: Boolean;
begin
  WorkbookPreview.SetActionMode(vgrpamZoomOut);
  Result := True;
end;

function TvgrWPPreview.KbdHand: Boolean;
begin
  WorkbookPreview.SetActionMode(vgrpamHand);
  Result := True;
end;

procedure TvgrWPPreview.KeyDown(var Key: Word; Shift: TShiftState);
begin
  if Worksheet <> nil then
    WorkbookPreview.KeyboardFilter.KeyDown(Key, Shift);
end;

procedure TvgrWPPreview.KeyPress(var Key: Char);
begin
  if Worksheet <> nil then
    WorkbookPreview.KeyboardFilter.KeyPress(Key);
end;

procedure TvgrWPPreview.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
begin
  APointInfo := TvgrWPPreviewPointInfo.Create;
end;

/////////////////////////////////////////////////
//
// TvgrWPSheetParams
//
/////////////////////////////////////////////////

/////////////////////////////////////////////////
//
// TvgrWPSheetParamsList
//
/////////////////////////////////////////////////
function TvgrWPSheetParamsList.GetItem(Index: Integer): TvgrWPSheetParams;
begin
  Result := TvgrWPSheetParams(inherited Items[Index]);
end;

function TvgrWPSheetParamsList.FindByWorksheet(AWorksheet: TvgrWorksheet): TvgrWPSheetParams;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    if Items[I].Worksheet = AWorksheet then
    begin
      Result := Items[I];
      exit;
    end;
  Result := nil;
end;

/////////////////////////////////////////////////
//
// TvgrWorkbookPreview
//
/////////////////////////////////////////////////
constructor TvgrWorkbookPreview.Create(AOwner : TComponent);
begin
  inherited;
  FBorderStyle := bsSingle;
  FActiveWorksheetIndex := -1;
  FPageMakers := TList.Create;
  FSheetParamsList := TvgrWPSheetParamsList.Create;
  FSheetsCaptionWidth := cDefaultSheetsCaptionWidth;

  FWorkbookBackColor := clWindow;
  FBackColor := clBtnFace;
  FMarginsColor := clSilver;

  FOptionsVertScrollBar := TvgrOptionsScrollBar.Create;
  FOptionsVertScrollBar.OnChange := OnOptionsChanged;

  FOptionsHorzScrollBar := TvgrOptionsScrollBar.Create;
  FOptionsHorzScrollBar.OnChange := OnOptionsChanged;

  FSizer := TvgrSizer.Create(Self);
  FSizer.Parent := Self;

  FWPPreview := TvgrWPPreview.Create(Self);
  FWPPreview.Parent := Self;

  FHorzScrollBar := TvgrScrollBar.Create(Self);
  FHorzScrollBar.Kind := sbHorizontal;
  FHorzScrollBar.OnScrollMessage := OnScrollMessage;
  FHorzScrollBar.Parent := Self;

  FVertScrollBar := TvgrScrollBar.Create(Self);
  FVertScrollBar.Kind := sbVertical;
  FVertScrollBar.OnScrollMessage := OnScrollMessage;
  FVertScrollBar.Parent := Self;

  FSheetsPanel := TvgrSheetsPanel.Create(Self);
  FSheetsPanel.Parent := Self;

  FWPNavigator := TvgrWPNavigator.Create(Self);
  FWPNavigator.Parent := Self;

  FSheetsPanelResizer := TvgrSheetsPanelResizer.Create(Self);
  FSheetsPanelResizer.Parent := Self;

  FPopupMenu := TPopupMenu.Create(Self);
  FDefaultPopupMenu := True;

  FKeyboardFilter := TvgrKeyboardFilter.Create;
  FKeyboardFilter.OnKeyPress := nil;
  FKeyboardFilter.OnBeforeSelectionCommand := nil;
  FKeyboardFilter.OnAfterSelectionCommand := nil;
  KbdAddCommands;
  FActionMode := vgrpamNone;
end;

destructor TvgrWorkbookPreview.Destroy;
begin
  Workbook := nil;
  ClearPageMakers;
  FKeyboardFilter.Free;
  FreeAndNil(FPageMakers);
  FreeAndNil(FSheetParamsList);
  FOptionsVertScrollBar.Free;
  FOptionsHorzScrollBar.Free;
  FPopupMenu.Free;
  inherited;
end;

procedure TvgrWorkbookPreview.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if AComponent = FWorkbook then
      Workbook := nil;
  end;
end;

procedure TvgrWorkbookPreview.Paint;
var
  ARect: TRect;
  AWidth, AHeight: Integer;
begin
  DrawBorder(Self, Canvas, FBorderStyle, ARect, AWidth, AHeight);
end;

procedure TvgrWorkbookPreview.OnPopupMenuClick(Sender: TObject);
begin
  DoPreviewAction(TvgrWorkbookPreviewAction(TMenuItem(Sender).Tag));
end;
procedure TvgrWorkbookPreview.OnPopupMenuClickChangeMode(Sender: TObject);
begin
  DoChangeActionMode(TvgrPreviewActionMode(TMenuItem(Sender).Tag));
end;

procedure TvgrWorkbookPreview.InitPopupMenu(APopupMenu: TPopupMenu);
begin
  ClearPopupMenu(APopupMenu);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemNoneTool), OnPopupMenuClickChangeMode, Integer(vgrpamNone), 'VGR_NONE_TOOL', True, False);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemHandTool), OnPopupMenuClickChangeMode, Integer(vgrpamHand), 'VGR_HAND_TOOL', True, False);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemZoomInTool), OnPopupMenuClickChangeMode, Integer(vgrpamZoomIn), 'VGR_ZOOMIN_TOOL', True, False);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemZoomOutTool), OnPopupMenuClickChangeMode, Integer(vgrpamZoomOut), 'VGR_ZOOMOUT_TOOL', True, False);
  AddMenuSeparator(APopupMenu, nil);  
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemZoomPageWidth), OnPopupMenuClick, Integer(vgrpaZoomPageWidth), 'VGR_ZOOM_PAGEWIDTH', IsPreviewActionAllowed(vgrpaZoomPageWidth), ScaleMode = vgrpsmPageWidth);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemZoom100Percent), OnPopupMenuClick, Integer(vgrpaZoom100Percent), 'VGR_ZOOM_100PERCENT', IsPreviewActionAllowed(vgrpaZoom100Percent), (ScaleMode = vgrpsmPercent) and (ScalePercent = 100));
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemZoomWholePage), OnPopupMenuClick, Integer(vgrpaZoomWholePage), 'VGR_ZOOM_WHOLEPAGE', IsPreviewActionAllowed(vgrpaZoomWholePage), ScaleMode = vgrpsmWholePage);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemZoomTwoPages), OnPopupMenuClick, Integer(vgrpaZoomTwoPages), 'VGR_ZOOM_TWOPAGES', IsPreviewActionAllowed(vgrpaZoomTwoPages), ScaleMode = vgrpsmTwoPages);    
  AddMenuSeparator(APopupMenu, nil);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemPriorPage), OnPopupMenuClick, Integer(vgrpaPriorPage), 'VGR_PRIOR', IsPreviewActionAllowed(vgrpaPriorPage), False, 'Ctrl+Left');
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemNextPage), OnPopupMenuClick, Integer(vgrpaNextPage), 'VGR_NEXT', IsPreviewActionAllowed(vgrpaNextPage), False, 'Ctrl+Right');
  AddMenuSeparator(APopupMenu, nil);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemPageParams), OnPopupMenuClick, Integer(vgrpaPageParams), 'VGR_PAGEPARAMS', IsPreviewActionAllowed(vgrpaPageParams), False);
  AddMenuSeparator(APopupMenu, nil);
  AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookPreview_MenuItemPrint), OnPopupMenuClick, Integer(vgrpaPrint), 'VGR_PRINT', IsPreviewActionAllowed(vgrpaPrint), False, 'Ctrl+P');
end;

procedure TvgrWorkbookPreview.DoScroll(const OldPos, NewPos: TPoint);
var
  X, Y: Integer;

  procedure CheckCoord(var ACoord: Integer; AScrollBar: TvgrScrollBar);
  begin
    with AScrollBar do
    begin
      if ACoord < Min then
        ACoord := Min;

      if Max - Min > PageSize then
      begin
        if ACoord > Max - PageSize then
          ACoord := Max - PageSize;
      end
      else
      begin
        if ACoord > Max then
          ACoord := Max;
      end;
    end;
  end;

begin
  X := NewPos.X;
  CheckCoord(X, HorzScrollBar);
  Y := NewPos.Y;
  CheckCoord(Y, VertScrollBar);

  if OldPos.X <> X then
    HorzScrollBar.Position := X;
  if OldPos.Y <> Y then
    VertScrollBar.Position := Y;
  if (OldPos.X <> X) or (OldPos.Y <> Y) then
    FWPPreview.DoScroll(OldPos, Point(X, Y));
end;

procedure TvgrWorkbookPreview.KbdAddCommands;
begin
  with KeyboardFilter do
  begin
    AddCommand(VK_LEFT, [], FWPPreview.KbdLeft);
    AddCommand(VK_RIGHT, [], FWPPreview.KbdRight);
    AddCommand(VK_UP, [], FWPPreview.KbdUp);
    AddCommand(VK_DOWN, [], FWPPreview.KbdDown);
    AddCommand(VK_PRIOR, [], FWPPreview.KbdPageUp);
    AddCommand(VK_NEXT, [], FWPPreview.KbdPageDown);
    AddCommand(VK_HOME, [], FWPPreview.KbdHome);
    AddCommand(VK_END, [], FWPPreview.KbdEnd);

    AddCommand(VK_LEFT, [ssCtrl], FWPPreview.KbdPriorPage);
    AddCommand(VK_RIGHT, [ssCtrl], FWPPreview.KbdNextPage);
    AddCommand(VK_PRIOR, [ssCtrl], FWPPreview.KbdPriorWorksheet);
    AddCommand(VK_NEXT, [ssCtrl], FWPPreview.KbdNextWorksheet);

    AddCommand(byte('P'), [ssCtrl], FWPPreview.KbdPrint);
    AddCommand(byte('E'), [ssCtrl], FWPPreview.KbdEditPageParams);

    AddCommand(byte('Z'), [], FWPPreview.KbdZoomIn);
    AddCommand(byte('Z'), [ssShift], FWPPreview.KbdZoomOut);    
    AddCommand(byte('H'), [], FWPPreview.KbdHand);
    AddCommand(byte(VK_SPACE), [], FWPPreview.KbdHand);    
  end;
end;

procedure TvgrWorkbookPreview.ClearPageMakers;
var
  I: Integer;
begin
  for I := 0 to PageMakerCount - 1 do
    PageMakers[I].Free;
  FPageMakers.Clear;
end;

procedure TvgrWorkbookPreview.UpdatePageMakers;
var
  I: Integer;
begin
  ClearPageMakers;
  if Workbook <> nil then
    for I := 0 to Workbook.WorksheetsCount - 1 do
      FPageMakers.Add(TvgrPageMaker.Create(Workbook.Worksheets[I]));
end;

procedure TvgrWorkbookPreview.SetWorkbook(Value: TvgrWorkbook);
begin
  if Value <> FWorkbook then
  begin
    if FWorkbook <> nil then
      FWorkbook.DisconnectHandler(Self);
    FWorkbook := Value;
    if FWorkbook <> nil then
      Value.ConnectHandler(Self);
    UpdatePageMakers;
    FSheetParamsList.Clear;
    if (FWorkbook <> nil) and (FWorkbook.WorksheetsCount > 0) then
      ActiveWorksheetIndex := 0
    else
      ActiveWorksheetIndex := -1;
    FSheetsPanel.UpdateSheets;
    FWPNavigator.Repaint;
  end;
end;

procedure TvgrWorkbookPreview.OnOptionsChanged(Sender: TObject);
begin
  if Sender = OptionsVertScrollBar then
    VertScrollBar.OptionsChanged
  else
    if Sender = OptionsHorzScrollBar then
      HorzScrollBar.OptionsChanged;
end;

procedure TvgrWorkbookPreview.SetWorkbookBackColor(Value: TColor);
begin
  if FWorkbookBackColor <> Value then
  begin
    FWorkbookBackColor := Value;
    Repaint;
  end;
end;

procedure TvgrWorkbookPreview.SetBackColor(Value: TColor);
begin
  if FBackColor <> Value then
  begin
    FBackColor := Value;
    Repaint;
  end;
end;

procedure TvgrWorkbookPreview.SetMarginsColor(Value: TColor);
begin
  if FMarginsColor <> Value then
  begin
    FMarginsColor := Value;
    Repaint;
  end;
end;

function TvgrWorkbookPreview.GetTopOffset: Integer;
begin
  Result := VertScrollBar.Position;
end;

function TvgrWorkbookPreview.GetLeftOffset: Integer;
begin
  Result := HorzScrollBar.Position;
end;

function TvgrWorkbookPreview.GetOffset: TPoint;
begin
  Result := Point(LeftOffset, TopOffset);
end;

function TvgrWorkbookPreview.GetScaleMode: TvgrPreviewScaleMode;
begin
  Result := FWPPreview.ScaleMode;
end;

procedure TvgrWorkbookPreview.SetScaleMode(Value: TvgrPreviewScaleMode);
begin
  FWPPreview.ScaleMode := Value;
end;

function TvgrWorkbookPreview.GetScalePercent: Double;
begin
  Result := FWPPreview.ScalePercent;
end;

procedure TvgrWorkbookPreview.SetScalePercent(Value: Double);
begin
  FWPPreview.ScalePercent := Value;
end;

function TvgrWorkbookPreview.IsScalePercentStored: Boolean;
begin
  Result := not IsZero(ScalePercent - 100);
end;

function TvgrWorkbookPreview.GetActivePage: Integer;
begin
  Result := FWPPreview.ActivePage;
end;

function TvgrWorkbookPreview.GetActionMode: TvgrPreviewActionMode;
begin
  Result := FActionMode;
end;

function TvgrWorkbookPreview.GetActivePrintPage: TvgrPrintPage;
begin
  if ActivePage >= 0 then
    Result := FWPPreview.PageMaker.Pages[ActivePage]
  else
    Result := nil;
end;

procedure TvgrWorkbookPreview.SetActivePage(Value: Integer);
begin
  FWPPreview.ActivePage := Value;
end;

procedure TvgrWorkbookPreview.SetActionMode(Value: TvgrPreviewActionMode);
begin
  FActionMode := Value;
  FWPPreview.FMarginsResizer.SetCurrentCursor;
end;

function TvgrWorkbookPreview.GetPageCount: Integer;
begin
  Result := FWPPreview.PageCount;
end;

procedure TvgrWorkbookPreview.SaveWPSheetParams;
var
  AParams: TvgrWPSheetParams;
begin
  if ActiveWorksheet <> nil then
  begin
    AParams := FSheetParamsList.FindByWorksheet(ActiveWorksheet);
    if AParams = nil then
    begin
      AParams := TvgrWPSheetParams.Create;
      AParams.Worksheet := ActiveWorksheet;
      FSheetParamsList.Add(AParams);
    end;

    AParams.ScrollPosition.X := HorzScrollbar.Position;
    AParams.ScrollPosition.Y := VertScrollbar.Position;
  end;
end;

procedure TvgrWorkbookPreview.RestoreWPSheetParams;
var
  AParams: TvgrWPSheetParams;
begin
  if (ActiveWorksheet <> nil) and HandleAllocated then
  begin
    AParams := FSheetParamsList.FindByWorksheet(ActiveWorksheet);
    if AParams <> nil then
    begin
    end;
  end;
end;

procedure TvgrWorkbookPreview.InternalUpdateActiveWorksheet;
begin
  if Workbook = nil then
  begin
    FActiveWorksheetIndex := -1;
    FActiveWorksheet := nil;
  end
  else
  begin
    FActiveWorksheetIndex := Min(Max(0, FActiveWorksheetIndex), Workbook.WorksheetsCount - 1);
    if FActiveWorksheetIndex = -1 then
      FActiveWorksheet := nil
    else
      FActiveWorksheet := Workbook.Worksheets[FActiveWorksheetIndex];
  end;

  FSheetsPanel.UpdateSheets;

  FWPPreview.UpdateAll;

  if FActiveWorksheetIndex = -1 then
  begin
    if HandleAllocated then
    begin
      HorzScrollBar.SetParams(0, 0, 0, 0);
      VertScrollBar.SetParams(0, 0, 0, 0);
      Repaint;
    end
  end
  else
  begin
    RestoreWPSheetParams;
  end;
end;

procedure TvgrWorkbookPreview.SetActiveWorksheetIndex(Value: Integer);
var
  AOldPagesCount: Integer;
begin
  if FActiveWorksheetIndex <> Value then
  begin
    if (ActiveWorksheet <> nil) and (FWPPreview.PageMaker <> nil) then
      AOldPagesCount := FWPPreview.PageMaker.PagesCount
    else
      AOldPagesCount := -1;

    SaveWPSheetParams;
    FActiveWorksheetIndex := Value;
    InternalUpdateActiveWorksheet;

    if (ActiveWorksheet <> nil) and (FWPPreview.PageMaker.PagesCount <> AOldPagesCount) then
      DoActiveWorksheetPagesCountChanged;

    DoActiveWorksheetChanged;
  end;
end;

function TvgrWorkbookPreview.GetPageMakerCount: Integer;
begin
  Result := FPageMakers.Count;
end;

function TvgrWorkbookPreview.GetPageMaker(Index: Integer): TvgrPageMaker;
begin
  Result := TvgrPageMaker(FPageMakers[Index]);
end;

function TvgrWorkbookPreview.FindPageMaker(AWorksheet: TvgrWorksheet): TvgrPageMaker;
var
  I: Integer;
begin
  I := 0;
  while (I < PageMakerCount) and (PageMakers[I].Worksheet <> AWorksheet) do Inc(I);
  if I < PageMakerCount then
    Result := PageMakers[I]
  else
    Result := nil;
end;

procedure TvgrWorkbookPreview.SetOptionsVertScrollBar(Value: TvgrOptionsScrollBar);
begin
  FOptionsVertScrollBar.Assign(Value);
end;

procedure TvgrWorkbookPreview.SetOptionsHorzScrollBar(Value: TvgrOptionsScrollBar);
begin
  FOptionsHorzScrollBar.Assign(Value);
end;

procedure TvgrWorkbookPreview.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWorkbookPreview.WMSetFocus(var Msg: TWMSetFocus);
begin
  Windows.SetFocus(FWPPreview.Handle);
end;

procedure TvgrWorkbookPreview.OnScrollMessage(Sender: TObject; var Msg: TWMScroll);
var
  OldPos: Integer;
  ScrollInfo: TScrollInfo;
begin
  if ActiveWorksheet <> nil then
  begin
    ScrollInfo.cbSize := sizeof(ScrollInfo);
    ScrollInfo.fMask := SIF_ALL;
    GetScrollInfo(TvgrScrollBar(Sender).Handle,SB_CTL,ScrollInfo);
    OldPos := ScrollInfo.nPos;

    case TScrollCode(Msg.ScrollCode) of
      scLineUp:
        Dec(ScrollInfo.nPos, 70);
      scLineDown:
        Inc(ScrollInfo.nPos, 70);
      scPageUp:
        Dec(ScrollInfo.nPos, ScrollInfo.nPage);
      scPageDown:
        Inc(ScrollInfo.nPos, ScrollInfo.nPage);
      scPosition, scTrack:
        begin
          if ScrollInfo.nTrackPos = ScrollInfo.nPos then exit;
          ScrollInfo.nPos := ScrollInfo.nTrackPos;
        end;
      scTop:
        ScrollInfo.nPos := ScrollInfo.nMin;
      scBottom:
        ScrollInfo.nPos := ScrollInfo.nMax;
    end;
    ScrollInfo.nPos := Max(ScrollInfo.nMin,ScrollInfo.nPos);
    if Sender = HorzScrollBar then
      DoScroll(Point(OldPos, TopOffset), Point(ScrollInfo.nPos, TopOffset))
    else
      DoScroll(Point(LeftOffset, OldPos), Point(LeftOffset, ScrollInfo.nPos));
  end;
end;

procedure TvgrWorkbookPreview.AlignSheetCaptions;
var
  ARect: TRect;
  ATop, AVertScrollBarWidth, AHorzScrollBarHeight, AWidth, AHeight: Integer;
begin
  if FVertScrollBar.Visible then
    AVertScrollBarWidth := FVertScrollBar.Width
  else
    AVertScrollBarWidth := 0;
  if FHorzScrollBar.Visible then
    AHorzScrollBarHeight := FHorzScrollBar.Height
  else
    AHorzScrollBarHeight := 0;

  DrawBorder(Self, Canvas, FBorderStyle, ARect, AWidth, AHeight);

  ATop := FWPPreview.Top + FWPPreview.Height;
  FSheetsPanel.SetBounds(ARect.Left,
                         ATop,
                         Min(FSheetsCaptionWidth, AWidth - AVertScrollBarWidth - cSheetsPanelResizerWidth),
                         AHorzScrollBarHeight);
  FSheetsPanelResizer.SetBounds(FSheetsPanel.BoundsRect.Right,
                                ATop,
                                cSheetsPanelResizerWidth,
                                AHorzScrollBarHeight);
  if OptionsHorzScrollBar.Visible then
    FHorzScrollBar.SetBounds(FSheetsPanelResizer.BoundsRect.Right,
                             FSheetsPanel.Top,
                             AWidth - FSheetsPanel.Width - FSheetsPanelResizer.Width - AVertScrollBarWidth,
                             FHorzScrollBar.Height)
  else
    FHorzScrollBar.SetBounds(FSheetsPanelResizer.BoundsRect.Right,
                             Height + 1,
                             AWidth - FSheetsPanel.Width - FSheetsPanelResizer.Width - AVertScrollBarWidth,
                             FHorzScrollBar.Height);
end;

procedure TvgrWorkbookPreview.AlignControls(AControl: TControl; var AAlignRect: TRect);
var
  ARect: TRect;
  ANavigatorHeight, AWidth, AHeight: Integer;
  AHorzScrollBarHeight, AVertScrollBarWidth: Integer;
begin
  if not FDisableAlign then
  begin
    if OptionsHorzScrollBar.Visible then
      AHorzScrollBarHeight := FHorzScrollBar.Height
    else
      AHorzScrollBarHeight := 0;
    if OptionsVertScrollBar.Visible then
      AVertScrollBarWidth := FVertScrollBar.Width
    else
      AVertScrollBarWidth := 0;

    DrawBorder(Self, Canvas, FBorderStyle, ARect, AWidth, AHeight);

    FWPPreview.SetBounds(ARect.Left,
                         ARect.Top,
                         AWidth - AVertScrollBarWidth,
                         AHeight - AHorzScrollBarHeight);
    // !!! MINIMUM Size of scrollbar - 2 pixel
    ANavigatorHeight := FVertScrollBar.Width * 4;
    if OptionsVertScrollBar.Visible then
    begin
      FWPNavigator.SetBounds(FWPPreview.Left + FWPPreview.Width,
                             FWPPreview.BoundsRect.Bottom - ANavigatorHeight,
                             FVertScrollBar.Width,
                             ANavigatorHeight);
      FVertScrollBar.SetBounds(FWPPreview.Left + FWPPreview.Width,
                               ARect.Top,
                               FVertScrollBar.Width,
                               FWPPreview.Height - ANavigatorHeight);
    end
    else
    begin
      FWPNavigator.SetBounds(Width + 1,
                             FWPPreview.Height - ANavigatorHeight,
                             FVertScrollBar.Width,
                             ANavigatorHeight);
      FVertScrollBar.SetBounds(Width + 1,
                               ARect.Top,
                               FVertScrollBar.Width,
                               FWPPreview.Height);
    end;

    AlignSheetCaptions;
    FSizer.SetBounds(FVertScrollBar.Left,
                     FWPNavigator.BoundsRect.Bottom,
                     AVertScrollBarWidth,
                     AHorzScrollBarHeight);
    Invalidate;
  end;
end;

procedure TvgrWorkbookPreview.SetBorderStyle(Value: TBorderStyle);
begin
  if FBorderStyle <> Value then
  begin
    FBorderStyle := Value;
    Realign;
  end;
end;

procedure TvgrWorkbookPreview.SetSheetsCaptionWidth(Value: Integer);
begin
  if Value <> FSheetsCaptionWidth then
  begin
    FSheetsCaptionWidth := Value;
    FDisableAlign := True;
    try
      AlignSheetCaptions;
    finally
      FDisableAlign := False;
    end;
  end;
end;

procedure TvgrWorkbookPreview.DoMouseMove(Sender: TControl; Shift: TShiftState; X, Y: Integer);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(Sender.ClientToScreen(Point(X, Y)));
  MouseMove(Shift, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookPreview.DoMouseDown(Sender: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(Sender.ClientToScreen(Point(X, Y)));
  AExecuteDefault := True;
  MouseDown(Button, Shift, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookPreview.DoMouseUp(Sender: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(Sender.ClientToScreen(Point(X, Y)));
  AExecuteDefault := True;
  MouseUp(Button, Shift, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookPreview.DoDblClick(Sender: TControl; var AExecuteDefault: Boolean);
begin
  AExecuteDefault := True;
  DblClick;
end;

procedure TvgrWorkbookPreview.DoDragOver(AChild: TControl; Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  DragOver(Source, APoint.X, APoint.Y, State, Accept);
end;

procedure TvgrWorkbookPreview.DoDragDrop(AChild: TControl; Source: TObject; X, Y: Integer);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  DragDrop(Source, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookPreview.DoNeedPageParams;
begin
  if Assigned(FOnNeedPageParams) then
    FOnNeedPageParams(Self);
end;

procedure TvgrWorkbookPreview.DoNeedPrint;
begin
  if Assigned(FOnNeedPrint) then
    FOnNeedPrint(Self);
end;

procedure TvgrWorkbookPreview.DoActivePageChanged;
begin
  FWPNavigator.Repaint;
  if Assigned(FOnActivePageChanged) then
    FOnActivePageChanged(Self);
end;

procedure TvgrWorkbookPreview.DoActiveWorksheetChanged;
begin
  FWPNavigator.Repaint;
  if Assigned(FOnActiveWorksheetChanged) then
    FOnActiveWorksheetChanged(Self);
end;

procedure TvgrWorkbookPreview.DoScalePercentChanged;
begin
  if Assigned(FOnScalePercentChanged) then
    FOnScalePercentChanged(Self);
end;

procedure TvgrWorkbookPreview.DoActiveWorksheetPagesCountChanged;
begin
  FWPPreview.UpdateScrollBars;
  if FWPPreview.ActivePage>FWPPreview.PageCount then
  begin
    FWPPreview.ActivePage := FWPPreview.PageCount-1;
  end
  else
  begin
    FWPPreview.ScrollToActivePage;  
  end;

  FWPNavigator.Repaint;
  if Assigned(FOnActiveWorksheetPagesCountChanged) then
    FOnActiveWorksheetPagesCountChanged(Self);
end;

procedure TvgrWorkbookPreview.BeforeChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcDeleteWorksheet:
      FDeleteActiveWorksheet := ActiveWorksheet.IndexInWorkbook = TvgrWorksheet(ChangeInfo.ChangedObject).IndexInWorkbook;
  end;
end;

procedure TvgrWorkbookPreview.AfterChangeWorkbook(ChangeInfo: TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcNewWorksheet, vgrwcChangeWorksheet:
      begin
        FSheetsPanel.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcDeleteWorksheet:
      begin
        UpdatePageMakers;
        if FDeleteActiveWorksheet then
          InternalUpdateActiveWorksheet;
        FSheetsPanel.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcUpdateAll:
    begin
      UpdatePageMakers;
      SaveWPSheetParams;
      InternalUpdateActiveWorksheet;
    end;
  end;
  FWPPreview.AfterChangeWorkbook(ChangeInfo);
end;

procedure TvgrWorkbookPreview.DeleteItemInterface(AItem: IvgrWBListItem);
begin
end;

function TvgrWorkbookPreview.GetScrollBarOptions(AScrollBar: TvgrScrollBar): TvgrOptionsScrollBar;
begin
  if AScrollBar = HorzScrollBar then
    Result := OptionsHorzScrollBar
  else
    Result := OptionsVertScrollBar;
end;

function TvgrWorkbookPreview.SheetsPanelParentGetWorkbook: TvgrWorkbook;
begin
  Result := Workbook;
end;

function TvgrWorkbookPreview.SheetsPanelParentGetActiveWorksheet: TvgrWorksheet;
begin
  Result := ActiveWorksheet;
end;

function TvgrWorkbookPreview.SheetsPanelParentGetActiveWorksheetIndex: Integer;
begin
  Result := ActiveWorksheetIndex;
end;

procedure TvgrWorkbookPreview.SheetsPanelParentSetActiveWorksheetIndex(Value: Integer);
begin
  ActiveWorksheetIndex := Value;
end;

procedure TvgrWorkbookPreview.SheetsPanelResizerParentDoResize(AOffset: Integer);
begin
  SheetsCaptionWidth := SheetsCaptionWidth + AOffset;
end;

function TvgrWorkbookPreview.IsPreviewActionAllowed(AAction: TvgrWorkbookPreviewAction): Boolean;
begin
  case AAction of
    vgrpaZoomPageWidth, vgrpaZoomWholePage, vgrpaZoom100Percent, vgrpaZoomPercent: Result := True;
    vgrpaFirstPage, vgrpaPriorPage: Result := ActivePage > 0;
    vgrpaNextPage, vgrpaLastPage: Result := ActivePage < PageCount - 1;
    vgrpaZoomTwoPages: Result := 1 < PageCount;
    vgrpaPrint: Result := Assigned(OnNeedPrint) and (Workbook <> nil);
    vgrpaPageParams: Result := Assigned(OnNeedPageParams) and (Workbook <> nil);
    vgrpaPriorWorksheet: Result := (Workbook <> nil) and (ActiveWorksheetIndex > 0);
    vgrpaNextWorksheet: Result := (Workbook <> nil) and (ActiveWorksheetIndex < Workbook.WorksheetsCount - 1);
  else
    Result := False;
  end;
end;

procedure TvgrWorkbookPreview.DoPreviewAction(AAction: TvgrWorkbookPreviewAction);
begin
  if IsPreviewActionAllowed(AAction) then
    case AAction of
      vgrpaZoomPageWidth: ScaleMode := vgrpsmPageWidth;
      vgrpaZoomWholePage: ScaleMode := vgrpsmWholePage;
      vgrpaZoomTwoPages: ScaleMode := vgrpsmTwoPages;
      vgrpaZoom100Percent:
        begin
          ScaleMode := vgrpsmPercent;
          ScalePercent := 100;
        end;
      vgrpaZoomPercent: ScaleMode := vgrpsmPercent;
      vgrpaFirstPage: ActivePage := 0;
      vgrpaPriorPage: ActivePage := ActivePage - 1;
      vgrpaNextPage: ActivePage := ActivePage + 1;
      vgrpaLastPage: ActivePage := PageCount - 1;
      vgrpaPrint: Print;
      vgrpaPageParams: EditPageParams;
      vgrpaPriorWorksheet:
        begin
          if ActiveWorksheetIndex > 0 then
            ActiveWorksheetIndex := ActiveWorksheetIndex - 1;
        end;
      vgrpaNextWorksheet:
        begin
          if ActiveWorksheetIndex < Workbook.WorksheetsCount - 1 then
            ActiveWorksheetIndex := ActiveWorksheetIndex + 1;
        end;
    end;
end;
procedure TvgrWorkbookPreview.DoChangeActionMode(AActionMode: TvgrPreviewActionMode);
begin
    ActionMode := AActionMode;
end;


procedure TvgrWorkbookPreview.Print;
begin
  DoNeedPrint;
end;

procedure TvgrWorkbookPreview.EditPageParams;
begin
  DoNeedPageParams;
end;

procedure TvgrWorkbookPreview.SetMultiPagePreview(const X, Y: Integer);
begin
  FWPPreview.FPageColCount := X;
  FWPPreview.FPageRowCount := Y;
  ScaleMode := vgrpsmMultiPages;
  FWPPreview.UpdateAll;
//  FWPPreview.UpdateSizes;
//  FWPPreview.UpdateScrollBars;
end;

procedure TvgrWorkbookPreview.ShowPopupMenu(const APoint: TPoint; APointInfo: TvgrPointInfo);
begin
  if DefaultPopupMenu and (PopupMenu = nil) then
  begin
    InitPopupMenu(FPopupMenu);
    FPopupMenu.Popup(APoint.X, APoint.Y);
  end;
end;

initialization

  FPreviewNavigatorBitmaps := TImageList.Create(nil);
  FPreviewNavigatorBitmaps.Width := 5;
  FPreviewNavigatorBitmaps.Height := 5;
  FPreviewNavigatorBitmaps.ResourceLoad(rtBitmap, svgrPreviewNavigatorBitmaps, clFuchsia);

  // Load cursors
  LoadCursor(cHandToolCursorResName, FHandToolCursor, cHandToolCursorID, crDefault);
  LoadCursor(cHandTool2CursorResName, FHandTool2Cursor, cHandTool2CursorID, crDefault);
  LoadCursor(cZoomInCursorResName, FZoomInCursor, cZoomInCursorID, crDefault);
  LoadCursor(cZoomOutCursorResName, FZoomOutCursor, cZoomOutCursorID, crDefault);
  
finalization

  FPreviewNavigatorBitmaps.Free;

end.

