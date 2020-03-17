{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{ This module contains TvgrWorkbookGrid class and some auxiliary classes. }
unit vgr_WorkbookGrid;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  Classes, SysUtils, Controls, ExtCtrls, Windows, graphics, forms, stdctrls,
  messages, math, ImgList, ComCtrls, Menus, ClipBrd,
  {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}

  vgr_Functions, vgr_DataStorage, vgr_DataStorageRecords, vgr_Consts, vgr_CommonClasses,
  vgr_DataStorageTypes, vgr_Keyboard, vgr_GUIFunctions, vgr_ReportGUIFunctions, vgr_Controls;

const
{Default delay of scrolling by mouse of worksheet area.}
  cDefaultScrollTimerInterval = 50;

type
  TvgrWBInplaceEdit = class;
  TvgrWBSheet = class;
  TvgrWorkbookGrid = class;
  TvgrWBSheetParams = class;
  TvgrOptionsHeader = class;
  TvgrOptionsFormulaPanel = class;
  TvgrWBFormulaPanel = class;

{Internal structure used while painting of worksheet.}
  rvgrCoordCacheInfo = record
    Pos: Integer;
    Size: Integer;
  end;
  { Pointer to rvgrCoordCacheInfo structure. }
  pvgrCoordCacheInfo = ^rvgrCoordCacheInfo;

{Specifies the current mouse operation in grid control.
Items:
  vgrmoNone - No operation.
  vgrmoHeaderResize - Header resizing (row or column).
  vgrmoWBSheetSelectionDrag - Selection dragging in worksheet area.
  vgrWBSheetSelect - Changing of the current selection in worksheet area.
  vgrmoHeaderInside - Changing of the current selection by moving mouse on header.}
  TvgrMouseOperation = (vgrmoNone, vgrmoHeaderResize, vgrmoWBSheetSelectionDrag, vgrWBSheetSelect, vgrmoHeaderInside);
{Specifies the type of place within a header (row or column).
Items:
  vgrhpNone - Within empty space.
  vgrhpInside - Within header.
  vgrhpResize - Within resize area of header.
Syntax:
  TvgrHeaderPlace = (vgrhpNone, vgrhpInside, vgrhpResize);}
  TvgrHeaderPlace = (vgrhpNone, vgrhpInside, vgrhpResize);

  TvgrHeadersPanelPointInfoClass = class of TvgrHeadersPanelPointInfo;
  /////////////////////////////////////////////////
  //
  // TvgrHeadersPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about a spot within the header area (row header or column header).
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt, if mouse within  headers area.}
  TvgrHeadersPanelPointInfo = class(TvgrPointInfo)
  private
    FNumber: Integer;
    FPlace: TvgrHeaderPlace;
    FVectors: TvgrVectors;
  public
{Creates an instance of the TvgrHeadersPanelPointInfo class.}
    constructor Create(APlace: TvgrHeaderPlace; AVectors: TvgrVectors; ANumber: Integer);
{Specifies the type of place within header.}
    property Place: TvgrHeaderPlace read FPlace;
{Specifies the number of header (number of column or row).}
    property Number: Integer read FNumber;
{Determines the list containing the given header.}
    property Vectors: TvgrVectors read FVectors;
  end;

{Internal structure, is used while calculating position inside TvgrWBHeadersPanel.
Syntax:
  rvgrHeaderPointInfo = record
    Number: Integer;
    Place: TvgrHeaderPlace;
  end;}
  rvgrHeaderPointInfo = record
    Number: Integer;
    Place: TvgrHeaderPlace;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBHeadersPanel
  //
  /////////////////////////////////////////////////
{TvgrWBHeadersPanel implements the control for editing columns and rows of active worksheet.
This is a base class, it has two descendants - TvgrWBColsPanel (for editing columns)
and TvgrWBRowsPanel (for editing rows).
This control is used only inside of the grid control.}
  TvgrWBHeadersPanel = class(TCustomControl)
  private
    FInvalidateRect: TRect;
    FDownPointInfo: rvgrHeaderPointInfo;
    FMouseOperation: TvgrMouseOperation;
    FDownMousePos: TPoint;
    FPriorMousePos: TPoint;
    FScrollTimer: TTimer;
    FScroll: Integer;
    function GetWorkbookGrid: TvgrWorkbookGrid;
    function GetWBSheet: TvgrWBSheet;
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
  protected
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); virtual; 
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); virtual;

    procedure OnScrollTimer(Sender: TObject);
    procedure InternalPaint(FromHeader,ToHeader,FromPixel,ToPixel: Integer);
    procedure InternalGetPointInfoAt(x: integer; var PointInfo: rvgrHeaderPointInfo);
    procedure InternalSetSelection(const ASelectionRect: TRect; ACurRectSide: TvgrBorderSide);
    procedure UpdateMouseCursor(const PointInfo: rvgrHeaderPointInfo);
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

    procedure SetSelection(ANumber1, ANumber2: Integer; ACurRectToMin: Boolean); virtual; abstract;
    procedure GetPointInfoAt(x,y: integer; var PointInfo: rvgrHeaderPointInfo); overload; virtual; abstract;
    function GetWBScrollBar: TvgrScrollBar; virtual; abstract;
    function GetOptions: TvgrOptionsHeader; virtual; abstract;
    function GetHeaderSize(Number: Integer): Integer; virtual; abstract;
    function GetHeaderCaption(Number: Integer): string; virtual; abstract;
    function GetHeaderRect(From,Size: Integer): TRect; virtual; abstract;
    function GetHeaderHasSelection(Number: Integer): boolean; virtual; abstract;
    function GetVisibleHeadersCount: Integer; virtual; abstract;
    function GetResizeMouseCursor: TCursor; virtual; abstract;
    function GetInsideMouseCursor: TCursor; virtual; abstract;
    procedure GetRectCoords(const ARect: TRect; var AStart, AEnd: Integer); virtual; abstract;
    function GetCoord(X, Y: Integer): Integer; virtual; abstract;
    function GetXCoord(ACoord: Integer): Integer; virtual; abstract;
    function GetYCoord(ACoord: Integer): Integer; virtual; abstract;
    procedure PaintHeaderBorder(const HeaderRect: TRect); virtual;
    procedure OptionsChanged; virtual;
    procedure DoScroll(OldPos,NewPos: Integer);
    procedure DoHeaderResize(Number,DeltaX,DeltaY: Integer); virtual; abstract;
    procedure DoSelectionChanged; virtual; abstract;
    procedure CreateWnd; override;

    function GetVectors: TvgrVectors; virtual; abstract;
    function GetPointInfoClass: TvgrHeadersPanelPointInfoClass; virtual; abstract;

    property Options: TvgrOptionsHeader read GetOptions;
    property Vectors: TvgrVectors read GetVectors;
  public
{Creates an instance of the TvgrWBHeadersPanel class.}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of the TvgrWBHeadersPanel class.}
    destructor Destroy; override;

    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;

{Returns the information about a spot within control area.
Parameters:
  APoint - The point within control, relative to top-left corner of control.
  APointInfo - The object of the class, inherited from TvgrPointInfo class, that describes a point within control.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo); overload;

{Returns the reference to parent grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
{Returns the reference to TvgrWBSheet control.}
    property WBSheet: TvgrWBSheet read GetWBSheet;
{Returns the reference to TvgrScrollBar control linked to this header control,
for TvgrWBColsPanel - vertical scrollbar, for TvgrWBRowsPanel -  horizontal scrollbar.}
    property WBScrollBar: TvgrScrollBar read GetWBScrollBar;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrColsPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within columns header area.
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt,
if mouse within columns header area.}
  TvgrColsPanelPointInfo = class(TvgrHeadersPanelPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBColsPanel
  //
  /////////////////////////////////////////////////
{TvgrWBColsPanel implements the control for editing columns of active worksheet,
it is linked with TvgrWorksheet.ColsList property.
This control is used only inside of the grid control.}
  TvgrWBColsPanel = class(TvgrWBHeadersPanel)
  protected
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); override;
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); override;

    procedure SetSelection(ANumber1, ANumber2: Integer; ACurRectToMin: Boolean); override;
    procedure GetPointInfoAt(x,y: integer; var PointInfo: rvgrHeaderPointInfo); override;
    function GetOptions: TvgrOptionsHeader; override;
    function GetHeaderSize(Number: Integer): Integer; override;
    function GetHeaderCaption(Number: Integer): string; override;
    function GetHeaderRect(From,Size: Integer): TRect; override;
    function GetHeaderHasSelection(Number: Integer): boolean; override;
    function GetVisibleHeadersCount: Integer; override;
    function GetResizeMouseCursor: TCursor; override;
    function GetInsideMouseCursor: TCursor; override;
    function GetWBScrollBar : TvgrScrollBar; override;
    procedure GetRectCoords(const ARect: TRect; var AStart, AEnd: Integer); override;
    function GetCoord(X, Y: Integer): Integer; override;
    function GetXCoord(ACoord: Integer): Integer; override;
    function GetYCoord(ACoord: Integer): Integer; override;
    procedure PaintHeaderBorder(const HeaderRect: TRect); override;
    procedure OptionsChanged; override;
    procedure DoHeaderResize(Number,DeltaX,DeltaY: Integer); override;
    procedure DoSelectionChanged; override;

    function GetVectors: TvgrVectors; override;
    function GetPointInfoClass: TvgrHeadersPanelPointInfoClass; override;

    procedure Paint; override;
    procedure Resize; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRowsPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within rows header area.
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt,
if mouse within rows header area.}
  TvgrRowsPanelPointInfo = class(TvgrHeadersPanelPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBRowsPanel
  //
  /////////////////////////////////////////////////
{TvgrWBRowsPanel implements the control for editing rows of active worksheet,
it is linked with TvgrWorksheet.RowsList property.
This control is used only inside of the grid control.}
  TvgrWBRowsPanel = class(TvgrWBHeadersPanel)
  protected
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); override;
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); override;

    procedure SetSelection(ANumber1, ANumber2: Integer; ACurRectToMin: Boolean); override;
    procedure GetPointInfoAt(x,y: integer; var PointInfo: rvgrHeaderPointInfo); override;
    function GetOptions: TvgrOptionsHeader; override;
    function GetHeaderSize(Number: Integer): Integer; override;
    function GetHeaderCaption(Number: Integer): string; override;
    function GetHeaderRect(From,Size: Integer): TRect; override;
    function GetHeaderHasSelection(Number: Integer): boolean; override;
    function GetVisibleHeadersCount: Integer; override;
    function GetResizeMouseCursor: TCursor; override;
    function GetInsideMouseCursor: TCursor; override;
    function GetWBScrollBar : TvgrScrollBar; override;
    procedure GetRectCoords(const ARect: TRect; var AStart, AEnd: Integer); override;
    function GetCoord(X, Y: Integer): Integer; override;
    function GetXCoord(ACoord: Integer): Integer; override;
    function GetYCoord(ACoord: Integer): Integer; override;
    procedure PaintHeaderBorder(const HeaderRect: TRect); override;
    procedure OptionsChanged; override;
    procedure DoHeaderResize(Number,DeltaX,DeltaY: Integer); override;
    procedure DoSelectionChanged; override;

    function GetVectors: TvgrVectors; override;
    function GetPointInfoClass: TvgrHeadersPanelPointInfoClass; override;

    procedure Paint; override;
    procedure Resize; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBLevelButton
  //
  /////////////////////////////////////////////////
{Represents the level button within the sections area.}
  TvgrWBLevelButton = class(TObject)
  public
    BoundsRect: TRect;
    Level: Integer;
  end;

{Specifies the type of place within sections area (vertical or horizontal).
Items:
  vgrsspNone - Within empty space.
  vgrsspSection - Within section.
  vgrsspSectionClose - Within the delete button of section.
  vgrsspSectionHide - Within the hide button of section.
  vgrsspLevelButton - Within the level button.
Syntax:
  TvgrWBSectionsPlace = (vgrsspNone, vgrsspSection, vgrsspSectionClose, vgrsspSectionHide, vgrsspLevelButton);}
  TvgrWBSectionsPlace = (vgrsspNone, vgrsspSection, vgrsspSectionClose, vgrsspSectionHide, vgrsspLevelButton);

{Internal structure, is used while calculating position inside TvgrWBSectionsPanel.
Syntax:
  rvgrWBSectionsPointInfo = record
    Place: TvgrWBSectionsPlace;
    Section: pvgrSection;
    SectionIndex: Integer;
    LevelButton: TvgrWBLevelButton;
  end;}
  rvgrWBSectionsPointInfo = record
    Place: TvgrWBSectionsPlace;
    Section: pvgrSection;
    SectionIndex: Integer;
    LevelButton: TvgrWBLevelButton;
  end;

{Internal structure, is used while painting of TvgrWBSectionsPanel.
Syntax:
  rvgrSectionCacheInfo = record
    Section: pvgrSection;
    SectionIndex: Integer;
  end;}
  rvgrSectionCacheInfo = record
    Section: pvgrSection;
    SectionIndex: Integer;
  end;
{Pointer to the rvgrSectionCacheInfo structure.}
  pvgrSectionCacheInfo = ^rvgrSectionCacheInfo;

  /////////////////////////////////////////////////
  //
  // TvgrSectionsCacheList
  //
  /////////////////////////////////////////////////
{Internal class, used while painting of TvgrWBSectionsPanel.}
  TvgrSectionsCacheList = class(TList)
  private
    function GetItem(Index: Integer): pvgrSectionCacheInfo;
  public
{Frees an instnce of the TvgrSectionsCacheList class.}
    destructor Destroy; override;
{Clears the list.}
    procedure Clear; override;
    function Add(ASection: pvgrSection; ASectionIndex: Integer): pvgrSectionCacheInfo;
    procedure SortByLevel;
    property Items[Index: Integer]: pvgrSectionCacheInfo read GetItem; default;
  end;

  TvgrSectionsPanelPointInfoClass = class of TvgrSectionsPanelPointInfo;
  /////////////////////////////////////////////////
  //
  // TvgrSectionsPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within sections area (vertical or horizontal).
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt,
if mouse within sections area. }
  TvgrSectionsPanelPointInfo = class(TvgrPointInfo)
  private
    FPlace: TvgrWBSectionsPlace;
    FSectionIndex: Integer;
    FLevelButton: TvgrWBLevelButton;
    FSections: TvgrSections;
  public
{Creates an instance of the TvgrSectionsPanelPointInfo class.
Parmeters:
  APlace - TvgrWBSectionsPlace
  ASections - TvgrSections
  ASectionIndex - Integer
  ALevelButton - TvgrWBLevelButton}
    constructor Create(APlace: TvgrWBSectionsPlace; ASections: TvgrSections; ASectionIndex: Integer; ALevelButton: TvgrWBLevelButton);
{Type of place within sections area.}
    property Place: TvgrWBSectionsPlace read FPlace;
{Returns the index of section in the Sections list if spot within section
(Place in (vgrsspSection, vgrsspSectionClose, vgrsspSectionHide) ), otherwise returns -1.}
    property SectionIndex: Integer read FSectionIndex;
{Returns the sections list linked with sections area.}
    property Sections: TvgrSections read FSections;
{Returns the reference to the level button if spot within section level button (Place = vgrsspLevelButton),
otherwise returns nil.}
    property LevelButton: TvgrWBLevelButton read FLevelButton;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBSectionsPanel
  //
  /////////////////////////////////////////////////
{TvgrWBSectionsPanel implements the control for editing sections of active worksheet.
This is a base class, it has two descendants - TvgrWBHorzSectionsPanel (for editing horizontal sections)
and TvgrWBVertSectionsPanel (for editing vertical sections).
This control is used only inside grid control.}
  TvgrWBSectionsPanel = class(TGraphicControl)
  private
    FLevelButtons: TList;
    FLevelButtonsFont: TFont;
    // for GetLevelCount
    FMaxLevel: Integer;
    // for GetHiddenSectionCountAtLevel
    FCountAtLevel: Integer;
    FCallbackLevel: Integer;
    // for GetCellBounds
    FMinStartPos: Integer;
    FMaxEndPos: Integer;
    // cached while paint
    FPaintSections: TvgrSectionsCacheList;
    FCellPoints: Array of rvgrCoordCacheInfo;
    FLevelSize: Integer;
    // for mouse operations
    FMouseSections: TvgrSectionsCacheList;
    FMovePointInfo: rvgrWBSectionsPointInfo;
    FDownPointInfo: rvgrWBSectionsPointInfo;

    function GetWorkbookGrid: TvgrWorkbookGrid;
    function GetWBSheet: TvgrWBSheet;
    function GetWorksheet: TvgrWorksheet;
    procedure GetLevelCountCallbackProc(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure GetSectionCountAtLevelCallbackProc(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure DoLevelButtonClickCallbackProc(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    function GetLevelButton(Index: Integer): TvgrWBLevelButton;
    function GetLevelButtonCount: Integer;
    procedure ClearLevelButtons;
  protected
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);

    function GetWBScrollBar: TvgrScrollBar; virtual; abstract;
    function GetSections: TvgrSections; virtual; abstract;
    function GetVisibleCellsCount: Integer; virtual; abstract;
    function GetMinSize: Integer; virtual; abstract;
    procedure UpdateSize; virtual;
    procedure RealignLevelButtons; virtual; abstract;
    function GetSectionRect(ASection: pvgrSection): TRect; virtual; abstract;
    procedure GetSectionButtonRects(const ASectionRect: TRect; var ACloseButton, AHideButton: TRect); virtual; abstract;
    function IsSectionHidden(ASection: pvgrSection): Boolean; virtual; abstract;

    function GetLevelCount: Integer;
    function GetHiddenSectionCountAtLevel(ALevel: Integer): Integer;
    function GetPanelSize: Integer;

    procedure DoOnFontChange(Sender: TObject); virtual;

    function GetPaintCellStart(ANumber: Integer): Integer;
    function GetPaintCellEnd(ANumber: Integer): Integer;
    function GetLevelStart(ALevel: Integer): Integer;
    function GetLevelEnd(ALevel: Integer): Integer;

    procedure DoScroll(AOldPos, ANewPos: Integer); virtual;

    procedure CalculateCellPoints; virtual; abstract;
    procedure PaintUpdateCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure PaintUpdate;
    procedure PaintLevelButton(ACanvas: TCanvas; AButton: TvgrWBLevelButton);
    procedure PaintSectionButton(ACanvas: TCanvas; const ARect: TRect; AImageIndex: Integer; AHighlighted: Boolean);
    procedure PaintSection(ACanvas: TCanvas; ASection: pvgrSection; ASectionIndex: Integer); virtual;
    procedure PaintLine(ACanvas: TCanvas; ALeft, ATop, AWidth, AHeight: Integer);
    procedure PaintLines(ACanvas: TCanvas); virtual; abstract;
    procedure PaintButtonsArea(ACanvas: TCanvas); virtual; abstract;
    function PaintGetBorders: Cardinal; virtual; abstract;
    procedure Paint; override;

    procedure GetPointInfoAtCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure GetPointInfoAt(X, Y: Integer; var PointInfo: rvgrWBSectionsPointInfo); overload;

    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;
    procedure CMMouseLeave(var Msg: TMessage); message CM_MOUSELEAVE;

    procedure DoSectionHide(ASection: pvgrSection; ASectionIndex: Integer); virtual; abstract;
    procedure DoSectionDelete(ASection: pvgrSection; ASectionIndex: Integer); virtual; abstract;
    procedure DoLevelButtonClick(ALevelButton: TvgrWBLevelButton); virtual;
    procedure DoSectionShow(ASection: pvgrSection); virtual; abstract;

    function GetPointInfoClass: TvgrSectionsPanelPointInfoClass; virtual; abstract;

    property LevelButtons[Index: Integer]: TvgrWBLevelButton read GetLevelButton;
    property LevelButtonCount: Integer read GetLevelButtonCount;
  public
{Creates an instance of the TvgrWBSectionsPanel class.}
    constructor Create(AOwner : TComponent); override;
{Frees an instance of the TvgrWBSectionsPanel class.}
    destructor Destroy; override;

{Returns the information about spot within control area.
Parameters:
  APoint - The point within control, relative to top-left corner of control.
  APointInfo - Contains the information about spot.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo); overload;

    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;
    
{Returns the reference to the parent grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
{Returns the reference to the TWBSheet control.}
    property WBSheet : TvgrWBSheet read GetWBSheet;
{Returns the reference to the TvgrScrollBar object, corresponds to this control.}
    property WBScrollBar : TvgrScrollBar read GetWBScrollBar;
{Returns the reference to the current active worksheet.}
    property Worksheet: TvgrWorksheet read GetWorksheet;
{Returns the reference to edited sections list of active worksheet.}
    property Sections: TvgrSections read GetSections;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrHorzSectionsPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within horizontal sections area.
Object of this type is returned by the TvgrWorkbookGrid.GetPointInfoAt method,
if mouse within horizontal sections area.}
  TvgrHorzSectionsPanelPointInfo = class(TvgrSectionsPanelPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBHorzSectionsPanel
  //
  /////////////////////////////////////////////////
{TvgrWBHorzSectionsPanel implements the control for editing horizontal sections of active worksheet.
It is linked with TvgrWorksheet.HorzSectionsList property.
This control is used only inside grid control.}
  TvgrWBHorzSectionsPanel = class(TvgrWBSectionsPanel)
  protected
    function GetWBScrollBar: TvgrScrollBar; override;
    function GetSections: TvgrSections; override;
    function GetVisibleCellsCount: Integer; override;
    function GetMinSize: Integer; override;
    function GetSectionRect(ASection: pvgrSection): TRect; override;
    procedure GetSectionButtonRects(const ASectionRect: TRect; var ACloseButton, AHideButton: TRect); override;
    function IsSectionHidden(ASection: pvgrSection): Boolean; override;
    procedure UpdateSize; override;
    procedure RealignLevelButtons; override;
    procedure DoOnFontChange(Sender: TObject); override;
    procedure CalculateCellPoints; override;
    procedure PaintSection(ACanvas: TCanvas; ASection: pvgrSection; ASectionIndex: Integer); override;
    procedure PaintLines(ACanvas: TCanvas); override;
    procedure PaintButtonsArea(ACanvas: TCanvas); override;
    function PaintGetBorders: Cardinal; override;

    procedure DoSectionHide(ASection: pvgrSection; ASectionIndex: Integer); override;
    procedure DoSectionDelete(ASection: pvgrSection; ASectionIndex: Integer); override;
    procedure DoSectionShow(ASection: pvgrSection); override;

    function GetPointInfoClass: TvgrSectionsPanelPointInfoClass; override;
  public
{Creates an instance of the TvgrWBHorzSectionsPanel class.
Parameters:
  AOwner - TComponent, owner component}
    constructor Create(AOwner : TComponent); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrVertSectionsPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within vertical sections area.
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt,
if mouse within vertical sections area.}
  TvgrVertSectionsPanelPointInfo = class(TvgrSectionsPanelPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBVertSectionsPanel
  //
  /////////////////////////////////////////////////
{TvgrWBHorzSectionsPanel implements the control for editing horizontal sections of active worksheet.
It is linked with TvgrWorksheet.VertSectionsList property.
This control used only inside grid control.}
  TvgrWBVertSectionsPanel = class(TvgrWBSectionsPanel)
  protected
    function GetWBScrollBar: TvgrScrollBar; override;
    function GetSections: TvgrSections; override;
    function GetVisibleCellsCount: Integer; override;
    function GetMinSize: Integer; override;
    function GetSectionRect(ASection: pvgrSection): TRect; override;
    procedure GetSectionButtonRects(const ASectionRect: TRect; var ACloseButton, AHideButton: TRect); override;
    function IsSectionHidden(ASection: pvgrSection): Boolean; override;
    procedure UpdateSize; override;
    procedure RealignLevelButtons; override;
    procedure DoOnFontChange(Sender: TObject); override;
    procedure CalculateCellPoints; override;
    procedure PaintSection(ACanvas: TCanvas; ASection: pvgrSection; ASectionIndex: Integer); override;
    procedure PaintLines(ACanvas: TCanvas); override;
    procedure PaintButtonsArea(ACanvas: TCanvas); override;
    function PaintGetBorders: Cardinal; override;

    procedure DoSectionHide(ASection: pvgrSection; ASectionIndex: Integer); override;
    procedure DoSectionDelete(ASection: pvgrSection; ASectionIndex: Integer); override;
    procedure DoSectionShow(ASection: pvgrSection); override;

    function GetPointInfoClass: TvgrSectionsPanelPointInfoClass; override;
  public
{Creates an instance of the TvgrWBVertSectionsPanel class.
Parameters:
  AOwner - TComponent, owner component}
    constructor Create(AOwner : TComponent); override;
  end;

{Describes the state of button on the TvgrFormulaPanel control.
Items:
  vgrbbsUnderMouse - The mouse pointer within button area.
  vgrbbsPressed - The mouse button is pressed within button area.}
  TvgrWBButtonState = (vgrbbsUnderMouse, vgrbbsPressed);
{Describes the state of button on the TvgrFormulaPanel control.
See also:
  TvgrWBButtonState}
  TvgrWBButtonStates = set of TvgrWBButtonState;
  /////////////////////////////////////////////////
  //
  // TvgrWBButton
  //
  /////////////////////////////////////////////////
{TvgrWBButton represents the button on TvgrFormulaPanel.}
  TvgrWBButton = class(TGraphicControl)
  private
    FBitmap: TBitmap;
    FState: TvgrWBButtonStates;
    procedure CMMouseEnter(var Msg: TMessage); message CM_MOUSEENTER;
    procedure CMMouseLeave(var Msg: TMessage); message CM_MOUSELEAVE;
  protected
    procedure Paint; override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

    property Bitmap: TBitmap read FBitmap write FBitmap;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBFormulaEdit
  //
  /////////////////////////////////////////////////
{TvgrWBFormulaEdit implements the text box on TvgrFormulaPanel.}
  TvgrWBFormulaEdit = class(TEdit)
  private
    FDisableChange: Boolean;
    function GetFormulaPanel: TvgrWBFormulaPanel;
    function GetSheet: TvgrWBSheet;
    function GetInplaceEdit: TvgrWBInplaceEdit;
  protected
    procedure DoEnter; override;
    procedure KeyPress(var Key: Char); override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure WMKillFocus(var Msg: TWMKillFocus); message WM_KILLFOCUS;
    procedure Change; override;
    procedure SetEditText(const S: string);

    property Sheet: TvgrWBSheet read GetSheet;
    property InplaceEdit: TvgrWBInplaceEdit read GetInplaceEdit;
  public
{Returns the reference to formula panel.}
    property FormulaPanel: TvgrWBFormulaPanel read GetFormulaPanel;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormulaPanelPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within formula panel area.}
  TvgrFormulaPanelPointInfo = class(TvgrPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBFormulaPanel
  //
  /////////////////////////////////////////////////
{TvgrWBHorzSectionsPanel implements the control for editing the value of range.
This control used only inside grid control. }
  TvgrWBFormulaPanel = class(TCustomControl)
  private
    FOkButton: TvgrWBButton;
    FCancelButton: TvgrWBButton;
    FEdit: TvgrWBFormulaEdit;
    function GetWorkbookGrid: TvgrWorkbookGrid;
    function GetOptions: TvgrOptionsFormulaPanel;
    procedure OnOkClick(Sender: TObject);
    procedure OnCancelClick(Sender: TObject);
  protected
    procedure AlignControls(AControl: TControl; var AAlignRect: TRect); override;
    procedure CreateWnd; override;
    procedure DoSelectedTextChanged;
    procedure OptionsChanged;
    procedure Paint; override;

    procedure UpdateEnabled;

    property Edit: TvgrWBFormulaEdit read FEdit;
    property OkButton: TvgrWBButton read FOkButton;
    property CancelButton: TvgrWBButton read FCancelButton;
    property Options: TvgrOptionsFormulaPanel read GetOptions;
  public
{Creates an instance of the TvgrWBFormulaPanel class.
Parameters:
  AOwner - The owner component}
    constructor Create(AOwner : TComponent); override;
{Frees an instance of a TvgrWBFormulaPanel class.}
    destructor Destroy; override;

{Returns the information about spot within control area.
Parameters:
  APoint - The point within control, relative to top-left corner of control.
  APointInfo - Contains information about spot.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);

{Returns the reference to grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrTopLeftButtonPointInfo
  //
  /////////////////////////////////////////////////
{Provides an information about spot within TopLeftButton area.
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt,
if mouse within TopLeftButton area. }
  TvgrTopLeftButtonPointInfo = class(TvgrPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBTopLeftButton
  //
  /////////////////////////////////////////////////
{TvgrWBTopLeftButton represents the control on crossing columns header and rows header.
By clicking on it user can select all cells from (0, 0) to (DimensionsRight, DimensionsBottom).
This control used only inside grid control. }
  TvgrWBTopLeftButton = class(TCustomControl)
  private
    function GetWorkbookGrid : TvgrWorkbookGrid;
    procedure WMEraseBkgnd(var Msg : TWMEraseBkgnd); message WM_ERASEBKGND;
  protected
    procedure Paint; override;
    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;
    procedure OptionsChanged;
  public
{Creates an intance of the TvgrWBTopLeftButton class.
Parameters:
  AOwner - The owner component}
    constructor Create(AOwner : TComponent); override;

    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;

{Returns the information about spot within control area.
Parameters:
  APoint - The point within control, relative to top-left corner of control.
  APointInfo - Contains information about spot.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);

{Returns the reference to grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrTopLeftButtonSectionsPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within TopLeftSectionsButton area.
Object of this type is returned by method TvgrWorkbookGrid.GetPointInfoAt,
if mouse within TopLeftSectionsButton area.}
  TvgrTopLeftButtonSectionsPointInfo = class(TvgrPointInfo)
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBTopLeftSectionsButton
  //
  /////////////////////////////////////////////////
{TvgrWBTopLeftSectionsButton represents the control on crossing vertical sections and horizontal sections.
This control used only inside grid control.}
  TvgrWBTopLeftSectionsButton = class(TCustomControl)
  private
    function GetWorkbookGrid : TvgrWorkbookGrid;
    procedure WMEraseBkgnd(var Msg : TWMEraseBkgnd); message WM_ERASEBKGND;
  protected
    procedure Paint; override;
    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    procedure OptionsChanged;
  public
{Creates an instance of the TvgrWBTopLeftSectionsButton class.
Parameters:
  AOwner - The owner component.}
    constructor Create(AOwner : TComponent); override;

    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;

{Returns the information about spot within control area.
Parameters:
  APoint - The point within control, relative to top-left corner of control.
  APointInfo - Contains the information about spot.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);

{Returns the reference to grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBInplaceEdit
  //
  /////////////////////////////////////////////////
{TvgrWBInplaceEdit represents the text box for editing range value within grid.}
  TvgrWBInplaceEdit = class(TCustomMemo)
  private
    FLastRect: TRect;
    FDisableChange: Boolean;

    function GetWorkbookGrid: TvgrWorkbookGrid;
    function GetWBSheet: TvgrWBSheet;
  protected
    procedure CNCommand(var Msg: TWMCommand); message CN_COMMAND;
    procedure WMSysChar(var Msg: TWMSysChar); message WM_SYSCHAR;

    procedure UpdateInternalTextRect; 
    procedure UpdateLongRangeSize;
    procedure UpdatePosition;
    procedure Init;
    procedure KeyPress(var Key: Char); override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure Resize; override;
    procedure CreateWnd; override;
    procedure WMKillFocus(var Msg: TWMKillFocus); message WM_KILLFOCUS;
    procedure Change; override;
    procedure SetEditText(const S: string);

    property WBSheet: TvgrWBSheet read GetWBSheet;
  public
{Creates an instance of the TvgrWBInplaceEdit class.
Parameters:
  AOwner - The owner component.}
    constructor Create(AOwner: TComponent); override;
{Returns the reference to grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBSelectedRegion
  //
  /////////////////////////////////////////////////
{Internal class, used for store information about selection in TvgrWorkbookGrid.}
  TvgrWBSelectedRegion = class(TObject)
  private
    FRectsList: TList;
    function GetRects(Index: Integer): PRect;
    procedure SetRects(Index: Integer; ARect: PRect);
protected
    function GetRectsCount: Integer;

    procedure SetSelection(const SelectionRect: TRect);
    procedure AddSelection(const SelectionRect: TRect);
    procedure Clear;

    function RowHasSelection(RowNumber: integer): boolean;
    function ColHasSelection(ColNumber: integer): boolean;

    procedure Assign(ASource: TvgrWBSelectedRegion);

    property Rects[Index: Integer] : PRect read GetRects write SetRects;
    property RectsCount: Integer read GetRectsCount;
  public
{Creates TvgrWBSelectedRegion object}
    constructor Create;
{Destroys TvgrWBSelectedRegion object}
    destructor Destroy; override;
  end;

{Internal structure, is used while painting TvgrWBSheet.
Caches the information about range of workbook.
Syntax:
  rvgrRangeCacheInfo = record
    Range: pvgrRange;
    Style: pvgrRangeStyle;

    DisplayText: string;
    RealAlign: TvgrRangeHorzAlign;
    TextWidth: Integer;

    Long: Boolean;
    Visible: Boolean;

    FullRect: TRect;
    InternalRect: TRect;
    FillRect: TRect;
    FillBordersRegion: HRGN;
  end;}
  rvgrRangeCacheInfo = record
    Range: pvgrRange;
    Style: pvgrRangeStyle;

    DisplayText: string;
    RealAlign: TvgrRangeHorzAlign;
    TextWidth: Integer;

    Long: Boolean;
    Visible: Boolean;

    FullRect: TRect;
    InternalRect: TRect;
    FillRect: TRect;
    FillBordersRegion: HRGN;
  end;
{Pointer to the rvgrRangeCacheInfo structure.
Syntax:
  pvgrRangeCacheInfo = ^rvgrRangeCacheInfo;}
  pvgrRangeCacheInfo = ^rvgrRangeCacheInfo;

{Internal structure, is used while painting TvgrWBSheet.
Caches the information about border of workbook.
Syntax:
  rvgrBorderCacheInfo = record
    Size: Integer;
    Border: pvgrBorder;
    Style: pvgrBorderStyle;
    Pos: Integer;
    Rect: TRect;
    Painted: Boolean;
  end;}
  rvgrBorderCacheInfo = record
    Size: Integer;
    Border: pvgrBorder;
    Style: pvgrBorderStyle;
    Pos: Integer;
    Rect: TRect;
    Painted: Boolean;
  end;
{Pointer to the rvgrBorderCacheInfo structure.
Syntax:
  pvgrBorderCacheInfo = ^rvgrBorderCacheInfo;}
  pvgrBorderCacheInfo = ^rvgrBorderCacheInfo;

{Internal structure, is used while painting TvgrWBSheet.
Caches the information about cell of workbook.
Syntax:
  rvgrCellCacheInfo = record
    Range: pvgrRangeCacheInfo;
    Empty: Boolean;
    Highlighted: Boolean;
  end;}
  rvgrCellCacheInfo = record
    Range: pvgrRangeCacheInfo;
    Empty: Boolean;
    Highlighted: Boolean;
    WithinLongRange: Boolean;
  end;
{Pointer to the rvgrCellCacheInfo structure.
Syntax:
  pvgrCellCacheInfo = ^rvgrCellCacheInfo;}
  pvgrCellCacheInfo = ^rvgrCellCacheInfo;

  TvgrRangeCacheInfoArray = Array of rvgrRangeCacheInfo;
  /////////////////////////////////////////////////
  //
  // TvgrWBSheetPaintInfo
  //
  /////////////////////////////////////////////////
{Internal class is used while painting of TvgrWBSheet.}
  TvgrWBSheetPaintInfo = class(TObject)
  private
    FWBSheet: TvgrWBSheet;
    FPaintRect: TRect;
    FCellsRect: TRect;
    FBounds: TRect;
    FBoundsWidth: Integer;
    FBoundsHeight: Integer;
    FCellsRectWidth: Integer;
    FCellsRectHeight: Integer;
    FVertPoints: Array of rvgrCoordCacheInfo;
    FRangesCount: Integer;
    FRanges: TvgrRangeCacheInfoArray;
    FLeftLongRanges: TvgrRangeCacheInfoArray;
    FRightLongRanges: TvgrRangeCacheInfoArray;
    FBorders: Array of rvgrBorderCacheInfo;
    FCells: Array of rvgrCellCacheInfo;
    function GetBorderInfoIndex(X, Y: Integer; AOrientation: TvgrBorderOrientation): Integer;
    procedure ExpandCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure CalcBorderRect(ABorderInfo: pvgrBorderCacheInfo; ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation);
    procedure UpdateBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure ResetBorder(X, Y: Integer; AOrientation: TvgrBorderOrientation);
    procedure CalcRangeCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure CalcInternalRect(const ARangeRect: TRect; var AInternalRect: TRect; var ABordersRegion: HRGN);
    procedure CalcNormalRange(AInfo: pvgrRangeCacheInfo);
    procedure CalcLongRange(AInfo: pvgrRangeCacheInfo);
    procedure CalcBorderPoint(X, Y: Integer);
  protected
    procedure Update(const ACellsRect, APaintRect: TRect);

    function GetBorderInfo(X, Y: Integer; AOrientation: TvgrBorderOrientation): pvgrBorderCacheInfo;
    function GetBorderSize(X, Y: Integer; AOrientation: TvgrBorderOrientation): Integer;
    function GetCellInfo(X, Y: Integer): pvgrCellCacheInfo;

    function IsCellPainted(X, Y: Integer): Boolean;
    function IsCellBackgroundPainted(X, Y: Integer): Boolean;

    procedure GetRectScreenPixels(const rRegions: TRect; var rPixels: TRect);
    function GetCellScreenRect(X, Y: Integer): TRect;
    function ColPixelRight(ACol: Integer): Integer;
    function ColPixelLeft(ACol: Integer): Integer;
    function ColPixelWidth(ACol: Integer): Integer;
    function RowPixelTop(ARow: Integer): Integer;
    function RowPixelBottom(ARow: Integer): Integer;
    function RowPixelHeight(ARow: Integer): Integer;
    function CellLeftBorderSize(X, Y: Integer): Integer;
    function CellRightBorderSize(X, Y: Integer): Integer;

    procedure Clear;
    procedure AfterPaint;

    property Bounds: TRect read FBounds;
  public
{Creates an instance of the TvgrWBSheetPaintInfo class.
Parameters:
  AWBSheet - The TvgrWBSheet object representing a painting worksheet.}
    constructor Create(AWBSheet: TvgrWBSheet);
{Destroys the instance of the TvgrWBSheetPaintInfo class.}
    destructor Destroy; override;
  end;

{Specifies the type of place within worksheet area.
Items:
  vgrspNone - Within empty space.
  vgrspInsideCell - Within cell.
  vgrspSelectionBorder - Within selection border.
Syntax:
  TvgrWBSheetPlace = (vgrspNone,vgrspInsideCell,vgrspSelectionBorder);}
  TvgrWBSheetPlace = (vgrspNone,vgrspInsideCell,vgrspSelectionBorder);
{Internal structure, is used for calculating position within TvgrWBSheet.
Syntax:
  rvgrWBSheetPointInfo = record
    Cell: TPoint;
    RangeIndex: Integer;
    Place: TvgrWBSheetPlace;
    DownRect: TRect;
  end;}
  rvgrWBSheetPointInfo = record
    Cell: TPoint;
    RangeIndex: Integer;
    Place: TvgrWBSheetPlace;
    DownRect: TRect;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSheetPointInfo
  //
  /////////////////////////////////////////////////
{Provides the information about spot within worksheet area.
Object of this type is returned by TvgrWorkbookGrid.GetPointInfoAt method,
if mouse within worksheet area. }
  TvgrSheetPointInfo = class(TvgrPointInfo)
  private
    FPlace: TvgrWBSheetPlace;
    FCell: TPoint;
    FRangeIndex: Integer;
    FDownRect: TRect;
  public
{Creates an instance of the TvgrSheetPointInfo class.}
    constructor Create(APlace: TvgrWBSheetPlace; const ACell: TPoint; ARangeIndex: Integer; const ADownRect: TRect);
{Type of place within worksheet area.
See also:
  TvgrWBSheetPlace}
    property Place: TvgrWBSheetPlace read FPlace;
{Returns the coordinates of cell if Place in (vgrspInsideCell, vgrspSelectionBorder).}
    property Cell: TPoint read FCell;
{Returns the index of range if Place in (vgrspInsideCell, vgrspSelectionBorder)
and range exists in this cell. You can get interface to this range like this:
ActiveWorksheet.RangeByIndex[RangeIndex]. Returns -1 if range does not exist.}
    property RangeIndex: Integer read FRangeIndex;
{Returns the Cell rect if RangeIndex <> -1, otherwise returns Place of Range.}
    property DownRect: TRect read FDownRect;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBSheet
  //
  /////////////////////////////////////////////////
{TvgrWBSheet represents the control for editing content of worksheet.
This control is used only inside grid control. }
  TvgrWBSheet = class(TCustomControl)
  private
    FSelectedRegion: TvgrWBSelectedRegion;
    FCurRect: TRect;
    FSelStartRect: TRect;
    FDownPointInfo: rvgrWBSheetPointInfo;
    FMouseOperation: TvgrMouseOperation;

    FScrollTimer: TTimer;
    FScrollX: Integer;
    FScrollY: Integer;

    FTempSelectionBounds: TRect;

    // for keyboard
    FKbdLeftOffs: Integer;
    FKbdTopOffs: Integer;
    FKbdOffsetX: Integer;
    FKbdOffsetY: Integer;
    // for Inplace edit
    FInplaceEditRect: TRect;
    FInplaceEditPixelRect: TRect;
    // for painting
    FPaintInfo: TvgrWBSheetPaintInfo;
    FPaintClipRect: TRect;
    // cached cols width
    FHorzPoints: Array of rvgrCoordCacheInfo;

    function GetWorkbookGrid: TvgrWorkbookGrid;
    function GetWorksheet: TvgrWorksheet;
    function GetHorzScrollBar: TvgrScrollBar;
    function GetVertScrollBar: TvgrScrollBar;
    function GetColsPanel: TvgrWBColsPanel;
    function GetRowsPanel: TvgrWBRowsPanel;
    function GetDimensionsRight: Integer;
    function GetDimensionsBottom: Integer;
    function GetLeftCol: Integer;
    function GetTopRow: Integer;
    function GetVisibleColsCount: Integer;
    function GetVisibleRowsCount: Integer;
    function GetFullVisibleColsCount: Integer;
    function GetFullVisibleRowsCount: Integer;
    function GetKeyboardFilter: TvgrKeyboardFilter;
    function GetInplaceEdit: TvgrWBInplaceEdit;
    function GetIsInplaceEdit: Boolean;

    procedure CheckSelectionBoundsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure UpdateHorzPoints;

    procedure PaintRangeBackground(AInfo: pvgrRangeCacheInfo);
    procedure PaintNormalRange(AInfo: pvgrRangeCacheInfo);
    procedure PaintLongRange(AInfo: pvgrRangeCacheInfo);
    procedure PaintCellBackground(AInfo: pvgrCellCacheInfo; X, Y: Integer);

    procedure PaintHorizontalBorders(ASelectionBorderRegion: HRGN; AGridBrush: HBRUSH);
    procedure PaintVerticalBorders(ASelectionBorderRegion: HRGN; AGridBrush: HBRUSH);
    procedure PaintBorders(ASelectionBorderRegion: HRGN);

    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMGetDlgCode(var Msg: TWMGetDlgCode); message WM_GETDLGCODE;
    procedure WMMouseWheel(var Msg: TWMMouseWheel); message WM_MOUSEWHEEL;

    function GetSelRect(const ACurRect, ASelStartRect: TRect): TRect;
    function CheckCell(const ACell: TPoint): TRect;
    procedure OnScrollTimer(Sender: TObject);
    function GetFormulaPanel: TvgrWBFormulaPanel;
  protected
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);

    procedure DoScroll(const OldPos, NewPos : TPoint);
    function ColPixelWidthNoCache(ColNumber: integer) : integer;
    function ColPixelWidth(ColNumber: integer) : integer;
    function RowPixelHeight(RowNumber: integer) : integer;
    function ColPixelLeft(ColNumber: integer) : integer;
    function ColPixelRight(ColNumber: Integer): Integer;
    function ColScreenPixelLeft(ColNumber: Integer): Integer;
    function ColScreenPixelRight(ColNumber: Integer): Integer;
    function RowPixelTop(RowNumber: integer) : integer;
    function RowPixelBottom(RowNumber: Integer): Integer;
    function RowScreenPixelTop(RowNumber: Integer): Integer;
    function RowScreenPixelBottom(RowNumber: Integer): Integer;
    function ColVisible(ColNumber: Integer): Boolean;
    function RowVisible(RowNumber: Integer): Boolean;
    procedure GetRectPixels(const rRegions: TRect; var rPixels: TRect);
    procedure GetRectScreenPixels(const rRegions: TRect; var rPixels: TRect);
    procedure GetRectInternalScreenPixels(const rRegions: TRect; var rPixels: TRect);
    procedure GetBorderPixelRect(ABorder: IvgrBorder; var ABorderRect: TRect; var ABorderPixelSize: Integer; var ABorderPos: Integer);
    function GetInvalidateRectForRange(ARange: IvgrRange): TRect;
    function IsRangeSelected(const ARangeRect: TRect): Boolean;
    function CheckAddCol(AColNumber: Integer): IvgrCol;
    function CheckAddRow(ARowNumber: Integer): IvgrRow;

    function GetSelectionBounds(ASelectionRect: TRect): TRect;
    procedure OffsetFCurRect(AOffsetX, AOffsetY: Integer);
    procedure SetUserSelection(ASelectionIndex, AOffsetX, AOffsetY, ALeftOffs, ATopOffs: Integer);
    function MakeRectVisible(const r: TRect; LeftOffs,TopOffs: Integer): boolean;
    procedure Paint; override;
    procedure GetPointInfoAt(X,Y: Integer; var PointInfo: rvgrWBSheetPointInfo); overload;
    procedure UpdateMouseCursor(const PointInfo: rvgrWBSheetPointInfo);

    function KbdLeft: Boolean;
    function KbdLeftCell: Boolean;
    function KbdRight: Boolean;
    function KbdRightCell: Boolean;
    function KbdUp: Boolean;
    function KbdTopCell: Boolean;
    function KbdDown: Boolean;
    function KbdBottomCell: Boolean;
    function KbdPageUp: Boolean;
    function KbdPageDown: Boolean;
    function KbdLeftTopCell: Boolean;
    function KbdRightBottomCell: Boolean;

    function KbdInplaceEdit: Boolean;
    function KbdClearValue: Boolean;
    function KbdCellProperties: Boolean;
    function KbdSelectedRowsAutoHeight: Boolean;
    function KbdSelectedColumnsAutoHeight: Boolean;

    function KbdCut: Boolean;
    function KbdCopy: Boolean;
    function KbdPaste: Boolean;

    procedure KbdBeforeSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
    procedure KbdAfterSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
    procedure KbdKeyPress(Key: Char);
    procedure KeyPress(var Key: Char); override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;

    function GetInplaceRange: IvgrRange;
    function IsLongRangeEdited: Boolean;
    function GetInplaceRangeHorzAlignment: TvgrRangeHorzAlign;
    function GetInplaceRangeVertAlignment: TvgrRangeVertAlign;
    procedure UpdateInplaceEditPosition;
    procedure StartInplaceEdit(AStartKey: Char = #0; AActiveInplaceEdit: Boolean = True);
    procedure EndInplaceEdit(ACancel: Boolean);

    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure DblClick; override;

    procedure SetSelection(const ASelectionRect: TRect);
    function CheckSelections(var ARepaintRect: TRect): Boolean;
    procedure CheckSelStartRect;

    property KeyboardFilter: TvgrKeyboardFilter read GetKeyboardFilter;
    property InplaceEdit: TvgrWBInplaceEdit read GetInplaceEdit;
  public
{Creates an instance of the TvgrWBSheet class.
Parameters:
  AOwner - The owner component.}
    constructor Create(AOwner : TComponent); override;
{Destroys an instance of the TvgrWBSheet class.}
    destructor Destroy; override;

{Returns the information about spot within control area.
Parameters:
  APoint - The point within control, relative to top-left corner of control.
  APointInfo - Contains the information about spot.}
    procedure GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo); overload;

    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean); override;

{Returns the reference to grid control.}
    property WorkbookGrid: TvgrWorkbookGrid read GetWorkbookGrid;
{Returns the reference to edited worksheet.}
    property Worksheet: TvgrWorksheet read GetWorksheet;

{Returns the true if inplace-editing is started.}
    property IsInplaceEdit: Boolean read GetIsInplaceEdit;

{Returns the value of DimensionsRight property of active worksheet.
Returns 0 if ActiveWorksheet = nil.}
    property DimensionsRight: Integer read GetDimensionsRight;
{Returns the value of DimensionsBottom property of active worksheet.
Returns 0 if ActiveWorksheet = nil.}
    property DimensionsBottom: Integer read GetDimensionsBottom;
{Returns the reference to the horizontal scrollbar.}
    property HorzScrollBar: TvgrScrollBar read GetHorzScrollBar;
{Returns the reference to the vertical scrollbar.}
    property VertScrollBar: TvgrScrollBar read GetVertScrollBar;
{Returns the reference to the columns header.
See also:
  TvgrWBColsPanel}
    property ColsPanel: TvgrWBColsPanel read GetColsPanel;
{Returns the reference to the rows header.
See also:
  TvgrWBRowsPanel}
    property RowsPanel: TvgrWBRowsPanel read GetRowsPanel;
{Returns the index of left column in the visible area.}
    property LeftCol: Integer read GetLeftCol;
{Returns the index of top row in the visible area.}
    property TopRow: Integer read GetTopRow;
{Returns the amount of visible columns.}
    property VisibleColsCount: Integer read GetVisibleColsCount;
{Returns the amount of visible rows.}
    property VisibleRowsCount: Integer read GetVisibleRowsCount;
{Returns the amount of completely visible columns,
if last column is visible partially, it is not taken into account.}
    property FullVisibleColsCount: Integer read GetFullVisibleColsCount;
{Returns the amount of completely visible rows,
if last row is visible partially, it is not taken into account.}
    property FullVisibleRowsCount: Integer read GetFullVisibleRowsCount;
{Returns the reference to the formula editor.}
    property FormulaPanel: TvgrWBFormulaPanel read GetFormulaPanel;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBSheetParams
  //
  /////////////////////////////////////////////////
{Internal class, is used to store information about parameters of worksheet at switching between worksheets.}
  TvgrWBSheetParams = class(TObject)
  public
    Worksheet: TvgrWorksheet;
    LeftTop: TPoint;
    SelectedRegion: TvgrWBSelectedRegion;
    SelStartRect: TRect;
    CurRect: TRect;
    constructor Create;
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBSheetParamsList
  //
  /////////////////////////////////////////////////
{Internal class representing a list of TvgrWBSheetParams objects.}
  TvgrWBSheetParamsList = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrWBSheetParams;
  public
    function FindByWorksheet(AWorksheet: TvgrWorksheet): TvgrWBSheetParams;
    function Add(AWorksheet: TvgrWorksheet): TvgrWBSheetParams;

    property Items[Index: Integer]: TvgrWBSheetParams read GetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrOptionsHeader
  //
  /////////////////////////////////////////////////
{Describes the options of header control (TvgrWBColsPanel and TvgrWBRowsPanel).}
  TvgrOptionsHeader = class(TvgrPersistent)
  private
    FVisible: Boolean;
    FBackColor: TColor;
    FHighlightedBackColor: TColor;
    FGridColor: TColor;
    FFont: TFont;
    procedure SetVisible(Value: Boolean);
    procedure SetBackColor(Value: TColor);
    procedure SetHighlightedBackColor(Value: TColor);
    procedure SetGridColor(Value: TColor);
    procedure SetFont(Value: TFont);
    function IsStoredHighlightBackColor: Boolean;
    function IsStoredFont: Boolean;
    procedure OnFontChange(Sender: TObject);
  public
{Creates an intance of the TvgrOptionsHeader class.}
    constructor Create; override;
{Frees an intance of the TvgrOptionsHeader class.}
    destructor Destroy; override;
{Copies the properties of source object.}
    procedure Assign(Source: TPersistent); override;
  published
{Specifies the value indicating whether the header control is visible.}
    property Visible: boolean read FVisible write SetVisible default True;
{Specifies the background color of the header control.}
    property BackColor: TColor read FBackColor write SetBackColor default clBtnFace;
{Specifies the highlighted color of the header control. Default color calculate according BackColor property.}
    property HighlightedBackColor: TColor read FHighlightedBackColor write SetHighlightedBackColor stored IsStoredHighlightBackColor;
{Specifies the color of grid lines in header control. Default clBtnText. }
    property GridColor: TColor read FGridColor write SetGridColor default clBtnText;
{Specifies the font of the header control. }
    property Font: TFont read FFont write SetFont stored IsStoredFont;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrOptionsTopLeftButton
  //
  /////////////////////////////////////////////////
{Describes the options of the TvgrWBTopLeftButton and TvgrWBTopLeftSectionsButton controls.}
  TvgrOptionsTopLeftButton = class(TvgrPersistent)
  private
    FBackColor: TColor;
    FBorderColor: TColor;
    procedure SetBackColor(Value: TColor);
    procedure SetBorderColor(Value: TColor);
  public
    constructor Create; override;
    procedure Assign(Source: TPersistent); override;
  published
{Specifies the background color.}
    property BackColor: TColor read FBackColor write SetBackColor default clBtnFace;
{Specifies the border color.}
    property BorderColor: TColor read FBorderColor write SetBorderColor default clBtnShadow;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrOptionsFormulaPanel
  //
  /////////////////////////////////////////////////
{Describes the options of TvgrWBFormulaPanel control.}
  TvgrOptionsFormulaPanel = class(TvgrPersistent)
  private
    FVisible: Boolean;
    FBackColor: TColor;
    procedure SetVisible(Value: Boolean);
    procedure SetBackColor(Value: TColor);
  public
{Creates an instance of the TvgrOptionsFormulaPanel class.}
    constructor Create; override;
{Copies the properties of another control.}
    procedure Assign(Source: TPersistent); override;
  published
{Specifies the value indicating whether the formula panel is visible.}
    property Visible: Boolean read FVisible write SetVisible default True;
{Specifies the background color of the formula panel.}
    property BackColor: TColor read FBackColor write SetBackColor default clBtnFace;
  end;

{TvgrWorkbookGridSectionDblClickEvent is the type for TvgrWorkbookGrid.OnSectionDblClick event.
This event occurs when the user double-clicks the left mouse button when the mouse pointer is
over the section.
Parameters:
  Sender - The TvgrWorkbookGrid control.
  ASections - Specified the TvgrSections object.
  ASectionIndex - Specifies the section's index in the ASections list.}
  TvgrWorkbookGridSectionDblClickEvent = procedure(Sender: TObject; ASections: TvgrSections; ASectionIndex: Integer) of object;
{TvgrWorkbookGridSetInplaceEditValueEvent is the type for TvgrWorkbookGrid.OnSetInplaceEditValue event.
This event occurs when the user ends inplace-editing.
Parameters:
  Sender - The TvgrWorkbookGrid control.
  AValue - The text that is entered by user.
  ARange - The range that is edited by user.
  ANewRange - Indicates whether the range is just created or has existed before (before inplace editing).
Example:
  procedure TMyForm.SetInplaceEditValue(Sender: TObject; const AValue: String; ARange: IvgrRange);
  begin
    ARange.SimpleStringValue := AValue;
    if ANewRange then
      ARange.WordWrap := (pos(#13, AValue) <> 0) or (pos(#10, AValue) <> 0);
  end;}
  TvgrWorkbookGridSetInplaceEditValueEvent = procedure(Sender: TObject; const AValue: string; ARange: IvgrRange; ANewRange: Boolean) of object;
  /////////////////////////////////////////////////
  //
  // TvgrWorkbookGrid
  //
  /////////////////////////////////////////////////
{Represents a spreadsheet control.}
  TvgrWorkbookGrid = class(TCustomControl,
                           IvgrWorkbookHandler,
                           IvgrControlParent,
                           IvgrScrollBarParent,
                           IvgrSheetsPanelParent,
                           IvgrSheetsPanelResizerParent,
                           IvgrSizerParent)
  private
    FSheet: TvgrWBSheet;
    FHorzScrollBar: TvgrScrollBar;
    FVertScrollBar: TvgrScrollBar;
    FRowsPanel: TvgrWBRowsPanel;
    FColsPanel: TvgrWBColsPanel;
    FSheetsPanel: TvgrSheetsPanel;
    FSheetsPanelResizer: TvgrSheetsPanelResizer;
    FFormulaPanel: TvgrWBFormulaPanel;
    FTopLeftButton: TvgrWBTopLeftButton;
    FTopLeftSectionsButton: TvgrWBTopLeftSectionsButton;
    FSizer: TvgrSizer;
    FHorzSections: TvgrWBHorzSectionsPanel;
    FVertSections: TvgrWBVertSectionsPanel;
    FInplaceEdit: TvgrWBInplaceEdit;
    FKeyboardFilter: TvgrKeyboardFilter;

    FVisibleRowsCount : integer;
    FFullVisibleRowsCount : integer;
    FVisibleColsCount : integer;
    FFullVisibleColsCount : integer;
    FLeftTopPixelsOffset : TPoint;
    FSheetParamsList: TvgrWBSheetParamsList;
    FDisableAlign: Boolean;
    FDeleteActiveWorksheet: Boolean;

    FWorkbook : TvgrWorkbook;
    FActiveWorksheetIndex : integer;
    FActiveWorksheet: TvgrWorksheet;
    FBorderStyle: TBorderStyle;

    FBackgroundColor : TColor;
    FGridColor : TColor;

    FOptionsCols: TvgrOptionsHeader;
    FOptionsRows: TvgrOptionsHeader;
    FOptionsVertScrollBar: TvgrOptionsScrollBar;
    FOptionsHorzScrollBar: TvgrOptionsScrollBar;
    FOptionsTopLeftButton: TvgrOptionsTopLeftButton;
    FOptionsFormulaPanel: TvgrOptionsFormulaPanel;

    FSelectionBorderColor: TColor;
    FSheetsCaptionWidth: Integer;
    FDefaultPopupMenu: Boolean;
    FReadOnly: Boolean;

    FWorksheetDimensionsColor: TColor;
    FShowWorksheetDimensions: Boolean;

    FPopupMenu: TPopupMenu;

    FSelectionFormat: IvgrRangesFormat;

    FOnSelectionChange: TNotifyEvent;
    FOnSectionDblClick: TvgrWorkbookGridSectionDblClickEvent;
    FOnCellProperties: TNotifyEvent;
    FOnSelectedTextChanged: TNotifyEvent;
    FOnActiveWorksheetChanged: TNotifyEvent;
    FOnSetInplaceEditValue: TvgrWorkbookGridSetInplaceEditValueEvent;

    FTempSize: Integer; // GetSelectedRowsHeight, GetSelectedColsHeight
    FTempCount: Integer;

    FPopupSections: TvgrSections;
    FPopupSectionIndex: Integer;
    FPopupVectors: TvgrVectors;
    FPopupVectorNumber: Integer;
    FPopupVectorNewSize: Integer;
    FPopupSheetIndex: Integer;

    FPriorSelectedText: string;

    procedure SetReadOnly(Value: Boolean);
    procedure SetBorderStyle(Value: TBorderStyle);
    procedure SetOptionsCols(Value: TvgrOptionsHeader);
    procedure SetOptionsRows(Value: TvgrOptionsHeader);
    procedure SetOptionsVertScrollBar(Value: TvgrOptionsScrollBar);
    procedure SetOptionsHorzScrollBar(Value: TvgrOptionsScrollBar);
    procedure SetOptionsTopLeftButton(Value: TvgrOptionsTopLeftButton);
    procedure SetOptionsFormulaPanel(Value: TvgrOptionsFormulaPanel);
    procedure SetWorkbook(Value : TvgrWorkbook);
    procedure InternalUpdateActiveWorksheet;
    procedure SetActiveWorksheetIndex(Value : integer);
    procedure SetBackgroundColor(Value : TColor);
    procedure SetGridColor(Value : TColor);
    procedure SetSelectionBorderColor(Value : TColor);
    function GetSelectionCount: Integer;
    function GetSelection(Index: Integer): TRect;
    function GetSelectionRects: TvgrRectArray;
    function GetIsInplaceEdit: Boolean;
    procedure SetSheetsCaptionWidth(Value: Integer);
    function GetScrollInterval: Integer;
    procedure SetScrollInterval(Value: Integer);
    function GetLeftCol: Integer;
    function GetTopRow: Integer;
    function GetDimensionsRight: Integer;
    function GetDimensionsBottom: Integer;
    function GetLastSelection: TRect;

    procedure OnOptionsChanged(Sender: TObject);

    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMSetFocus(var Msg: TWMSetFocus); message WM_SETFOCUS;

    procedure OnScrollMessage(Sender: TObject; var Msg : TWMScroll);
    procedure UpdateScrollBar(ScrollBar : TvgrScrollBar; NewPos,VisibleCells,FullVisibleCells,MaxCells : integer);
    procedure GetPaintRects(PaintRect: TRect; var rRegions,rPixels: TRect; ALeftCol, ATopRow: Integer);
    procedure CalcVisibleRegion(ALeftCol, ATopRow: Integer);

    procedure GetSelectedSizeCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);

    procedure SetShowWorksheetDimensions(Value: Boolean);
    procedure SetWorksheetDimensionsColor(Value: TColor);

    function MergeWarningProc: Boolean;
//    function GetRangeSizeCallback(ARange: IvgrRange): TSize;
  protected
    // IvgrWorkbookHandler
    procedure BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); virtual;
    procedure AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); virtual;
    procedure DeleteItemInterface(AItem: IvgrWBListItem); virtual;
    // IvgrControlParent
    procedure DoMouseMove(AChild: TControl; Shift: TShiftState; X, Y: Integer);
    procedure DoMouseDown(AChild: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
    procedure DoMouseUp(AChild: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
    procedure DoDblClick(AChild: TControl; var AExecuteDefault: Boolean);
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

    procedure SaveWBSheetParams; virtual;
    procedure RestoreWBSheetParams; virtual;

    procedure KbdAddCommands;

    procedure DoSectionDblClick(ASections: TvgrSections; ASectionIndex: Integer);
    procedure DoSelectionChange;
    procedure DoSelectedTextChanged;
    procedure DoCellProperties;

    procedure CheckSelectedTextChanged;

    procedure OnPopupMenuClick(Sender: TObject); virtual;
    procedure InitPopupMenu(APopupMenu: TPopupMenu; APointInfo: TvgrPointInfo); virtual;

    procedure DoScroll(const OldPos,NewPos: TPoint);
    procedure DoScrollBy(AScrollX, AScrollY: Integer; AStartTimer: Boolean);
    procedure DoColWidthChanged(ColNumber : integer);
    procedure DoRowHeightChanged(RowNumber : integer);
    procedure AlignSheetCaptions;
    procedure AlignControls(AControl: TControl; var AAlignRect: TRect); override;
    procedure Notification(AComponent : TComponent; AOperation : TOperation); override;
    procedure Loaded; override;

    procedure DrawBorder(Canvas: TCanvas; var AInternalRect: TRect; var AWidth, AHeight: Integer);
    procedure Paint; override;

    procedure DoActiveWorksheetChanged;
    procedure DoSetInplaceEditValue(const AValue: string; ARange: IvgrRange; ANewRange: Boolean);

    procedure AssignInplaceEditValueToRange(const AValue: string; ARange: IvgrRange; ANewRange: Boolean); virtual;

    procedure ReadDefaultColWidth(Reader: TReader);
    procedure WriteDefaultColWidth(Writer: TWriter);
    procedure ReadDefaultRowHeight(Reader: TReader);
    procedure WriteDefaultRowHeight(Writer: TWriter);
    procedure DefineProperties(Filer: TFiler); override;

    function GetDefaultColWidth: Integer;
    function GetDefaultRowHeight: Integer;

    function GetSelectedText: string;

    property HorzScrollBar: TvgrScrollBar read FHorzScrollBar;
    property VertScrollBar: TvgrScrollBar read FVertScrollBar;

    property KeyboardFilter: TvgrKeyboardFilter read FKeyboardFilter;
  public
{Creates an instance of the TvgrWorkbookGrid class.
Parameters:
  AOwner - The owner component.}
    constructor Create(AOwner : TComponent); override;
{Destroys an instance of the TvgrWorkbookGrid class.}
    destructor Destroy; override;

{Starts the inplace-editing mode.}
    procedure StartInplaceEdit;
{Ends the inplace editing.
Parameters:
  ACancel - Specifies the value indicating whether the user's changes must be saved.}
    procedure EndInplaceEdit(ACancel: Boolean);

{Defines the current selection rect.
Parameters:
  ASelectionRect - Specifies the selection rect.}
    procedure SetSelection(const ASelectionRect: TRect);

{Deletes the selected rows. Warning! The rows for deleting are determined according
of Selections property. So, for example, if you have two areas with
coordinates (0 (left), 3 (top), 4 (right), 5(bottom)) and (10, 7, 12, 8),
then rows with numbers 3, 4, 5, 7, 8 will be deleted. }
    procedure DeleteRows;
{Deletes the selected columns. Warning! The columns for deleting are determined according
of Selections property. So, for example, if you have two areas with
coordinates (0 (left), 3 (top), 4 (right), 5(bottom)) and (10, 7, 12, 8),
then columns with numbers 0, 1, 2, 3, 4, 10, 11, 12 will be deleted.}
    procedure DeleteColumns;
{Inserts the row before top of last selection rect.}
    procedure InsertRow;
{Inserts the column before left of last selection rect. }
    procedure InsertColumn;
{Inserts the worksheet at specified position. 
Parameters:
  AInsertIndex - If AInsertIndex = -1 then, new worksheet inserted before ActiveWorksheet }
    procedure InsertWorksheet(AInsertIndex: Integer = -1);

{Can be used to checking the height of selected rows.
If height is identical for all rows, returns this height, if
no rows are selected, returns DefaultRowHeight.
If the rows have different height, returns -1.
Returning value measured in twips.}
    function GetSelectedRowsHeight: Integer;
{Defines the one height (in twips) for all selected rows.
Parameters:
  AHeight - The row's height in twips.}
    procedure SetSelectedRowsHeight(AHeight: Integer);
{Can be used to checking the width of selected columns.
If width is identical for all columns, returns this width, if
no columsn are selected, returns DefaultColWidth.
If the columns have different width, returns -1.
Returning value measured in twips.}
    function GetSelectedColsWidth: Integer;
{Define the one width (in twips) for all selected columns.
Parameters:
  AWidth - The column's width in twips.}
    procedure SetSelectedColsWidth(AWidth: Integer);

{Shows or hides the selected columns.
Parameters:
  AShow - If equals to true then columns are shown.}
    procedure ShowHideSelectedColumns(AShow: Boolean);
{Shows or hides the selected rows.
Parameters:
  AShow - If equals to true then rows are shown.}
    procedure ShowHideSelectedRows(AShow: Boolean);

{Sets the optimal width for all selected columns.
Width is calculated on the basis of content of cells in columns.
If column does not have cells, DefaultColWidth is used.}
    procedure SetAutoWidthForSelectedColumns;
{Sets the optimal height for all selected rows.
Height is calculated on the basis of content of cells in rows.
If row does not have cells, DefaultColHeight is used.}
    procedure SetAutoHeightForSelectedRows;

{Returns the true if the merged cells exist in selection.}
    function IsMergeSelected: Boolean;
{Unmerges the merged cells (if they exist in selection) or
merges the cells in the selection.}
    procedure MergeUnMergeSelection;

{Returns the true if "Cut to clipboard" action is enabled.}
    function IsCutAllowed: Boolean;
{Returns the true if "Copy to clipboard" action is enabled.}
    function IsCopyAllowed: Boolean;
{Returns the true if "paste from clipboard" action is enabled.}
    function IsPasteAllowed: Boolean;
{Cuts the selected cells, rows, columns to clipboard.}
    procedure CutToClipboard;
{Copies the selected cells, rows, columns to clipboard.}
    procedure CopyToClipboard;
{Pastes the clipboard content at the current cursor position.}
    procedure PasteFromClipboard;
{Clears the selection contents.
Parameters:
  AClearFlags - Specifies the elements to delete.
See also:
  TvgrClearContentFlags}
    procedure ClearSelection(AClearFlags: TvgrClearContentFlags);

{Returns the information about point within grid.
X and Y coordinates are calculated from top-left corner of grid.
If the point is inside header, class TvgrHeadersPointInfo returns.
Parameters:
  X, Y - Specifies the spot within control area.
  APointInfo - Contains the inforamation about spot.}
    procedure GetPointInfoAt(X, Y: Integer; var APointInfo: TvgrPointInfo);

{Displays the standard popup menu at specified point.
The content of the popup menu is determenated by the APointInfo parameter.
Parameters:
  APopupPoint - Specifies the popup point.
  APointInfo - Contains the information about spot.}
    procedure ShowPopupMenu(const APopupPoint: TPoint; APointInfo: TvgrPointInfo);

{Opens the built-in dialog for editing name and title of worksheet.
Parameters:
  AWorksheet - The edited TvgrWorksheet object.
Return value:
  Returns the true if user press OK.}
    function EditWorksheet(AWorksheet: TvgrWorksheet): Boolean;
{Opens the built-in dialog for editing worsheets' positions in workbook.
Parameters:
  AWorksheet - The TvgrWorksheet object, which position is edited.}
    function CopyMoveWorksheet(AWorksheet: TvgrWorksheet): Boolean;

{Returns the current active worksheet.}
    property ActiveWorksheet: TvgrWorksheet read FActiveWorksheet;

{Returns the index of left column in the visible area.}
    property LeftCol: Integer read GetLeftCol;
{Returns the index of top row in the visible area.}
    property TopRow: Integer read GetTopRow;
{Returns the offset of left col and top row from beginning of workbook in pixels.
If LeftCol = 0 and TopRow = 0, returns (0, 0),
if LeftCol = 2, TopRow = 1, returns
(Columns[0].Width + Columns[1].Width + Columns[2].Width, Rows[0].Height + Rows[1].Height)}
    property LeftTopPixelsOffset: TPoint read FLeftTopPixelsOffset;
{Returns the amount of visible columns.}
    property VisibleColsCount: Integer read FVisibleColsCount;
{Returns the amount of visible rows.}
    property VisibleRowsCount: Integer read FVisibleRowsCount;
{Returns the amount of completely visible columns,
if last column is visible partially, it is not taken into account.}
    property FullVisibleColsCount: Integer read FFullVisibleColsCount;
{Returns the amount of completely visible rows,
if last row is visible partially, it is not taken into account.}
    property FullVisibleRowsCount: Integer read FFullVisibleRowsCount;
{Returns the value of DimensionsRight property of active worksheet.
Returns 0 if ActiveWorksheet = nil.
See also:
  TvgrWorksheet.DimensionsRight}
    property DimensionsRight: Integer read GetDimensionsRight;
{Returns the value of DimensionsBottom property of active worksheet.
Returns 0 if ActiveWorksheet = nil.
See also:
  TvgrWorksheet.DimensionsBottom}
    property DimensionsBottom: Integer read GetDimensionsBottom;

{Returns the amount of selection's rects on active worksheet.
The selection can consist of several rectangular areas,
amount of these areas can be received with use of this property.}
    property SelectionCount: Integer read GetSelectionCount;
{Lists the selection's rects.
See also:
  SelectionCount}
    property Selections[Index: Integer]: TRect read GetSelection;
{Returns the selections rect as dynamic array of TRect structures.}
    property SelectionRects: TvgrRectArray read GetSelectionRects;
{Returns the last selection rect, like Selections[SelectionCount - 1],
if active worksheet has no selection, returns Rect(-1, -1, -1, -1).}
    property LastSelection: TRect read GetLastSelection;
{Returns the IvgrRangeFormat interface for current selection.
See also:
  IvgrRangesFormat}
    property SelectionFormat: IvgrRangesFormat read FSelectionFormat;

{Returns the reference to InplaceEdit object.}
    property InplaceEdit: TvgrWBInplaceEdit read FInplaceEdit;
{Returns the true if grid is in a inplaceedit mode.}
    property IsInplaceEdit: Boolean read GetIsInplaceEdit;
{Returns the default column width in twips.
See also:
  TvgrWorkbook.DefaultColWidth}
    property DefaultColWidth: Integer read GetDefaultColWidth;
{Returns the default row height in twips.
See also:
  TvgrWorkbook.DefaultRowHeight}
    property DefaultRowHeight: Integer read GetDefaultRowHeight;
{Returns the currently selected text, which is displayed into the formula edit.}
    property SelectedText: string read GetSelectedText;
  published
    property Align;
    property Constraints;
    property UseDockManager default True;
{
    property DockSite;
    property DragCursor;
    property DragKind;
    property DragMode;
}    
    property Enabled;
    property Visible;
    property Anchors;
    property TabOrder;
    property TabStop;
    property PopupMenu;

{Specifies the TvgrWorkbook object, which is edited with use given Grid.}
    property Workbook : TvgrWorkbook read FWorkbook write SetWorkbook;
{Specifies the active worksheet index.}
    property ActiveWorksheetIndex: Integer read FActiveWorksheetIndex write SetActiveWorksheetIndex default -1;

{Specifies the background color for active worksheet area.}
    property BackgroundColor : TColor read FBackgroundColor write SetBackgroundColor default clWindow;
{Specifies the color of grid lines.}
    property GridColor : TColor read FGridColor write SetGridColor default clActiveBorder;
{Specifies the color of selection's borders.}
    property SelectionBorderColor: TColor read FSelectionBorderColor write SetSelectionBorderColor default clBlack;
{Determines whether the workbookgrid control has a single line border around the client area.}
    property BorderStyle: TBorderStyle read FBorderStyle write SetBorderStyle default bsSingle;
{Indicates whether the dimensions of active worksheet must be displayed.
See also:
  WorksheetDimensionsColor, TvgrWorksheet.Dimensions}
    property ShowWorksheetDimensions: Boolean read FShowWorksheetDimensions write SetShowWorksheetDimensions default True;
{Specifies the color that is used to display worksheet dimensions.
See also:
  ShowWorksheetDimensions, TvgrWorksheet.Dimensions}
    property WorksheetDimensionsColor: TColor read FWorksheetDimensionsColor write SetWorksheetDimensionsColor default clRed;

{Specifies the options of columns header area.
See also:
  TvgrOptionsHeader}
    property OptionsCols: TvgrOptionsHeader read FOptionsCols write SetOptionsCols;
{Specifies the options of rows header area.
See also:
  TvgrOptionsHeader}
    property OptionsRows: TvgrOptionsHeader read FOptionsRows write SetOptionsRows;
{Specifies the options of vertical scrollbar.
See also:
  TvgrOptionsScrollBar}
    property OptionsVertScrollBar: TvgrOptionsScrollBar read FOptionsVertScrollBar write SetOptionsVertScrollBar;
{Specifies the options of horizontal scrollbar.
See also:
  TvgrOptionsScrollBar}
    property OptionsHorzScrollBar: TvgrOptionsScrollBar read FOptionsHorzScrollBar write SetOptionsHorzScrollBar;
{Specifies the options of TopLeftButton control, which is on crossing columns header and rows header.
See also:
  TvgrOptionsTopLeftButton}
    property OptionsTopLeftButton: TvgrOptionsTopLeftButton read FOptionsTopLeftButton write SetOptionsTopLeftButton;
{Specifies the options of formula panel.
See also:
  TvgrOptionsFormulaPanel}
    property OptionsFormulaPanel: TvgrOptionsFormulaPanel read FOptionsFormulaPanel write SetOptionsFormulaPanel;

{Specifies the width of sheets caption area. Default value is determined by cDefaultSheetsCaptionWidth constant.}
    property SheetsCaptionWidth: Integer read FSheetsCaptionWidth write SetSheetsCaptionWidth default cDefaultSheetsCaptionWidth;
{Delay of scrolling by mouse of worksheet area. Default value is determined by cDefaultScrollTimerInterval constant.}
    property ScrollInterval: Integer read GetScrollInterval write SetScrollInterval default cDefaultScrollTimerInterval;

{If DefaultPopupMenu = true, the built-in PopupMenu is used,
otherwise the PopupMenu specified in PopupMenu property is used.}
    property DefaultPopupMenu: Boolean read FDefaultPopupMenu write FDefaultPopupMenu default True;
{To restrict the grid control to display only, set the ReadOnly property to true.
Set ReadOnly to false to allow the contents of the grid control to be edited. }
    property ReadOnly: Boolean read FReadOnly write SetReadOnly default False;

{This event occurs when current selection is changed. 
Changes in selection can occur when the user changes the selection using
the mouse or keyboard, or when the selection is changed through the SetSelection method. }
    property OnSelectionChange: TNotifyEvent read FOnSelectionChange write FOnSelectionChange;
{This event occurs when the user double-clicks the left mouse button when the mouse pointer is
over the section.
See also:
  TvgrWorkbookGridSectionDblClickEvent}
    property OnSectionDblClick: TvgrWorkbookGridSectionDblClickEvent read FOnSectionDblClick write FOnSectionDblClick;
{This event occurs when the user calls the cell properties dialog,
through built-in popup menu or shortcut keys.
Example:
  procedure TMyForm.aFormatCellExecute(Sender: TObject);
  begin
    CellPropertiesDialog.Execute(ActiveWorksheet, WorkbookGrid.SelectionRects);
  end;}
    property OnCellProperties: TNotifyEvent read FOnCellProperties write FOnCellProperties;
{This event occurs when the user changes the cursor's position in grid.}
    property OnSelectedTextChanged: TNotifyEvent read FOnSelectedTextChanged write FOnSelectedTextChanged;

{This event occurs the active worksheet is changed.}
    property OnActiveWorksheetChanged: TNotifyEvent read FOnActiveWorksheetChanged write FOnActiveWorksheetChanged;  

{This event occurs when user ends inplace-editing.
See also:
  TvgrWorkbookGridSetInplaceEditValueEvent}
    property OnSetInplaceEditValue: TvgrWorkbookGridSetInplaceEditValueEvent read FOnSetInplaceEditValue write FOnSetInplaceEditValue;

    property OnMouseMove;
    property OnMouseDown;
    property OnMouseUp;
    property OnDblClick;

    property OnDragDrop;
    property OnDragOver;
{
    property OnStartDrag;
}
  end;

implementation

{$R vgrWorkbookGrid.res}
{$R ..\res\vgr_WorkbookGridStrings.res}

uses
  {$IFDEF VGR_D6_OR_D7} Types, DateUtils, {$ENDIF}
  vgr_Dialogs, vgr_ReportFunctions,
  vgr_StringIDs,
  vgr_SheetFormatDialog,
  vgr_CopyMoveSheetDialog, vgr_Localize, vgr_PageProperties;

const
  WM_REDIRECTED_KEYDOWN = WM_USER + 1;
  { Clipboard format for workbook }
  CF_VGRCELLS = CF_MAX + 1;
  { Default font size of TvgrWBHeadersPanel (header of columns and rows) }
  cHeaderFontSize = 10;
  { Drawing offset, used by TvgrWBColsPanel }
  cColsPanelHorzOffset = 4;
  { Drawing offset, used by TvgrWBColsPanel }
  cColsPanelVertOffset = 4;
  { Drawing offset, used by TvgrWBRowsPanel }
  cRowsPanelHorzOffset = 4;
  { Drawing offset, used by TvgrWBRowsPanel }
  cRowsPanelVertOffset = 4;
  { Drawing offset, used by TvgrWBColsPanel and TvgrWBRowsPanel }
  cColsRowsResizeOffset = 3;
  { Drawing offset, used by TvgrWBSectionsPanel }
  cSectionsPanelLevelOffset = 4;
  { Width of close and hide buttons in section }
  cSectionButtonWidth = 7;
  { Height of close and hide buttons in section }
  cSectionButtonHeight = 7;
  { Drawing offset, used by TvgrWBSectionsPanel }
  cSectionLevelButtonOffset = 2;
  { Name of resource bitmap of section close and hide buttons }
  svgrSectionBitmaps = 'VGR_BMP_SECTIONBITMAPS';
  { Font used, by TvgrWBSectionsPanel to draw level buttons }
  sSectionLevelButtonFontName = 'Tahoma';
  cCellCursorID = 3;
  cToRightCursorID = 4;
  cToBottomCursorID = 5;
  { Cursor resource name }
  cCellCursorResName = 'VGR_CUR_CELL';
  { Cursor resource name }
  cToRightCursorResName = 'VGR_TO_RIGHT';
  { Cursor resource name }
  cToBottomCursorResName = 'VGR_TO_BOTTOM';
  sDefFont = 'Arial';
  sDefStringForWidth = '0';
  cSectionsMaxEndPos = MaxInt div 2;

var
  FFormulaOkBitmap: TBitmap;
  FFormulaCancelBitmap: TBitmap;
  FSectionBitmaps: TImageList;
  FCellCursor: TCursor;
  FToBottomCursor: TCursor;
  FToRightCursor: TCursor;

procedure AssignFont(const ASource: rvgrFont; ADest: TFont);
begin
  with ADest do
  begin
    Name := ASource.Name;
    Size := ASource.Size;
    Pitch := ASource.Pitch;
    Style := ASource.Style;
    Charset := ASource.Charset;
    Color := ASource.Color;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrHeadersPanelPointInfo
//
/////////////////////////////////////////////////
constructor TvgrHeadersPanelPointInfo.Create(APlace: TvgrHeaderPlace; AVectors: TvgrVectors; ANumber: Integer);
begin
  inherited Create;
  FPlace := APlace;
  FNumber := ANumber;
  FVectors := AVectors;
end;

/////////////////////////////////////////////////
//
// TvgrWBHeadersPanel
//
/////////////////////////////////////////////////
constructor TvgrWBHeadersPanel.Create(AOwner : TComponent);
begin
  inherited;
  FScrollTimer := TTimer.Create(nil);
  FScrollTimer.Enabled := False;
  FScrollTimer.Interval := cDefaultScrollTimerInterval;
  FScrollTimer.OnTimer := OnScrollTimer;
end;

destructor TvgrWBHeadersPanel.Destroy;
begin
  FreeAndNil(FScrollTimer);
  inherited;
end;

procedure TvgrWBHeadersPanel.CreateWnd;
begin
  inherited;
  OptionsChanged;
end;

procedure TvgrWBHeadersPanel.OptionsChanged;
begin
  Repaint;
end;

procedure TvgrWBHeadersPanel.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

function TvgrWBHeadersPanel.GetWorkbookGrid: TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

function TvgrWBHeadersPanel.GetWBSheet: TvgrWBSheet;
begin
  Result := WorkbookGrid.FSheet;
end;

procedure TvgrWBHeadersPanel.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
end;

procedure TvgrWBHeadersPanel.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo); 
begin
end;

procedure TvgrWBHeadersPanel.PaintHeaderBorder;
begin
end;

procedure TvgrWBHeadersPanel.InternalPaint(FromHeader,ToHeader,FromPixel,ToPixel: Integer);
var
  HeaderCaption: string;
  i,j,HeaderSize: integer;
  HeaderRect: TRect;
begin
  Canvas.Font.Assign(Options.Font);
  Canvas.Pen.Style := psSolid;
  Canvas.Pen.Color := Options.GridColor;
  Canvas.Pen.Width := 1;
  j := FromPixel;
  for i:=FromHeader to ToHeader do
    begin
      HeaderSize := GetHeaderSize(i);
      HeaderCaption := GetHeaderCaption(i);
      HeaderRect := GetHeaderRect(j,HeaderSize);
      if GetHeaderHasSelection(i) then
        Canvas.Brush.Color := Options.HighlightedBackColor
      else
        Canvas.Brush.Color := Options.BackColor;
      Canvas.TextRect(HeaderRect,
                      HeaderRect.Left + (HeaderRect.Right - HeaderRect.Left - Canvas.TextWidth(HeaderCaption)) div 2,
                      HeaderRect.Top + (HeaderRect.Bottom - HeaderRect.Top - Canvas.TextHeight(HeaderCaption)) div 2,
                      HeaderCaption);
      j := j + HeaderSize;
      PaintHeaderBorder(HeaderRect);
    end;
end;

procedure TvgrWBHeadersPanel.InternalGetPointInfoAt(x: integer; var PointInfo: rvgrHeaderPointInfo);
var
  i,j,m: integer;
begin
  PointInfo.Number := -1;
  PointInfo.Place := vgrhpNone;
  if (x < 0) or (WorkbookGrid.ActiveWorksheet = nil) then exit;
  i := cColsRowsResizeOffset;
  j := WBScrollBar.Position;
  m := j + GetVisibleHeadersCount;
  repeat
    i := i + GetHeaderSize(j);
    Inc(j);
  until (j > m) or (i >= x);
  if j > m then exit;

  if x >= i - cColsRowsResizeOffset then
  begin
    while GetHeaderSize(j) = 0 do Inc(j);
    PointInfo.Place := vgrhpResize;
  end
  else
    PointInfo.Place := vgrhpInside;
  PointInfo.Number := j - 1;
end;

procedure TvgrWBHeadersPanel.InternalSetSelection(const ASelectionRect: TRect; ACurRectSide: TvgrBorderSide);
begin
  WorkbookGrid.FSheet.SetSelection(ASelectionRect);
  
  with ASelectionRect do
    case ACurRectSide of
      vgrbsLeft:
        WorkbookGrid.FSheet.FCurRect := Rect(Left, Top, Left, Bottom);
      vgrbsTop:
        WorkbookGrid.FSheet.FCurRect := Rect(Left, Top, Right, Top);
      vgrbsRight:
        WorkbookGrid.FSheet.FCurRect := Rect(Right, Top, Right, Bottom);
      vgrbsBottom:
        WorkbookGrid.FSheet.FCurRect := Rect(Left, Bottom, Right, Bottom);
    end;
end;

procedure TvgrWBHeadersPanel.DoScroll(OldPos,NewPos: Integer);
begin
  Invalidate;
end;

procedure TvgrWBHeadersPanel.UpdateMouseCursor(const PointInfo: rvgrHeaderPointInfo);
var
  NewCursor: TCursor;
begin
  NewCursor := Cursor;
  case PointInfo.Place of
    vgrhpInside: NewCursor := GetInsideMouseCursor;
    vgrhpResize: if not WorkbookGrid.ReadOnly then
                   NewCursor := GetResizeMouseCursor;
    else exit;
  end;
  if Cursor <> NewCursor then
    Cursor := NewCursor;
end;

procedure TvgrWBHeadersPanel.OnScrollTimer(Sender: TObject);
var
  AStart, AEnd: Integer;
begin
  GetRectCoords(WorkbookGrid.SelectionRects[0], AStart, AEnd);

  if FScroll = -1 then
  begin
    if AStart = 0 then
      FScroll := 0;
  end;

  if FScroll = 0 then
    FScrollTimer.Enabled := False
  else
  begin
    WorkbookGrid.FSheet.SetUserSelection(0, GetXCoord(FScroll), GetYCoord(FScroll), 0, 0);
  end;
end;

procedure TvgrWBHeadersPanel.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  PointInfo: rvgrHeaderPointInfo;
  AStart, AEnd, ACoord: Integer;
begin
  case FMouseOperation of
    vgrmoHeaderResize:
      begin
        if not WorkbookGrid.ReadOnly then
        begin
          DoHeaderResize(FDownPointInfo.Number, ConvertPixelsToTwipsX(X - FPriorMousePos.X), ConvertPixelsToTwipsY(Y - FPriorMousePos.Y));
          FPriorMousePos := Point(X, Y);
        end;
      end;
    vgrmoHeaderInside:
      begin
        GetRectCoords(ClientRect, AStart, AEnd);
        ACoord := GetCoord(X, Y);
        if ACoord < AStart then
          FScroll := -1
        else
          if ACoord > AEnd then
            FScroll := 1
          else
            FScroll := 0;

        if FScroll <> 0 then
          FScrollTimer.Enabled := True
        else
        begin
          GetPointInfoAt(X, Y, PointInfo);
          if FDownPointInfo.Number < PointInfo.Number then
            SetSelection(FDownPointInfo.Number, PointInfo.Number, False)
          else
            SetSelection(PointInfo.Number, FDownPointInfo.Number, True);
        end;
      end;
  else
  begin
    GetPointInfoAt(x,y,PointInfo);
    UpdateMouseCursor(PointInfo);
  end;
  end;
  WorkbookGrid.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrWBHeadersPanel.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APointInfo: TvgrPointInfo;
  AExecuteDefault: Boolean;
  AStart, AEnd: Integer;
begin
  WorkbookGrid.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);

  if AExecuteDefault then
  begin
    GetPointInfoAt(x, y, FDownPointInfo);
    if FDownPointInfo.Place = vgrhpNone then
    begin
    end
    else
      if (not WorkbookGrid.ReadOnly) and (FDownPointInfo.Place = vgrhpResize) and (Button = mbLeft) then
      begin
        FMouseOperation := vgrmoHeaderResize;
        FDownMousePos := Point(X,Y);
        FPriorMousePos := FDownMousePos;
      end
      else
      begin
        if ssShift in Shift then
        begin
          GetRectCoords(WorkbookGrid.FSheet.FSelStartRect, AStart, AEnd);
          if FDownPointInfo.Number < AStart then
            SetSelection(FDownPointInfo.Number, AEnd, True)
          else
            if FDownPointInfo.Number > AEnd then
              SetSelection(AStart, FDownPointInfo.Number, False)
            else
              SetSelection(AStart, AEnd, True);
        end
        else
        begin
          SetSelection(FDownPointInfo.Number, FDownPointInfo.Number, True);
          with WorkbookGrid.SelectionRects[0] do
          begin
            WorkbookGrid.FSheet.FSelStartRect := Rect(Left, Top, Left, Top);
            WorkbookGrid.FSheet.CheckSelStartRect;
          end;
        end;

        if Button = mbRight then
        begin
          GetPointInfoAt(Point(X, Y), APointInfo);
          WorkbookGrid.ShowPopupMenu(ClientToScreen(Point(X, Y)), APointInfo);
          APointInfo.Free;
        end
        else
          FMouseOperation := vgrmoHeaderInside;
      end;
  end;
end;

procedure TvgrWBHeadersPanel.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  FScrollTimer.Enabled := False;
  WorkbookGrid.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    FMouseOperation := vgrmoNone;
  end;
end;

procedure TvgrWBHeadersPanel.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  WorkbookGrid.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrWBHeadersPanel.DragDrop(Source: TObject; X, Y: Integer); 
begin
  WorkbookGrid.DoDragDrop(Self, Source, X, Y);
end;

procedure TvgrWBHeadersPanel.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
var
  PointInfo: rvgrHeaderPointInfo;
begin
  GetPointInfoAt(APoint.x, APoint.y, PointInfo);
  with PointInfo do
    APointInfo := GetPointInfoClass.Create(Place, Vectors, Number);
end;

/////////////////////////////////////////////////
//
// TvgrColsPanelPointInfo
//
/////////////////////////////////////////////////

/////////////////////////////////////////////////
//
// TvgrWBColsPanel
//
/////////////////////////////////////////////////
procedure TvgrWBColsPanel.OptionsChanged;
begin
  if Options.Visible then
    Height := cColsPanelVertOffset + GetFontHeight(Options.Font)
  else
    Height := 0;
  inherited;
end;

function TvgrWBColsPanel.GetWBScrollBar : TvgrScrollBar;
begin
  Result := WorkbookGrid.HorzScrollBar;
end;

procedure TvgrWBColsPanel.GetRectCoords(const ARect: TRect; var AStart, AEnd: Integer);
begin
  AStart := ARect.Left;
  AEnd := ARect.Right;
end;

function TvgrWBColsPanel.GetCoord(X, Y: Integer): Integer;
begin
  Result := X;
end;

function TvgrWBColsPanel.GetXCoord(ACoord: Integer): Integer;
begin
  Result := ACoord;
end;

function TvgrWBColsPanel.GetYCoord(ACoord: Integer): Integer;
begin
  Result := 0;
end;

procedure TvgrWBColsPanel.Paint;
var
  rClip, rRegions, rPixels: TRect;
begin
  rClip := Canvas.ClipRect;
  rClip.Top := 0;
  rClip.Bottom := ClientHeight;
  WorkbookGrid.GetPaintRects(rClip, rRegions, rPixels, WorkbookGrid.LeftCol, WorkbookGrid.TopRow);
  InternalPaint(rRegions.Left, rRegions.Right, rPixels.Left, rPixels.Right);

  DrawBorders(Canvas,
              Rect(0, 0, ClientWidth, ClientHeight),
              False,
              (WorkbookGrid.FVertSections.Height > 0) or (WorkbookGrid.FFormulaPanel.Height > 0),
              False,
              True,
              clBtnShadow);
end;

procedure TvgrWBColsPanel.Resize;
begin
  inherited;
  WorkbookGrid.FVertSections.RealignLevelButtons;
end;

procedure TvgrWBColsPanel.PaintHeaderBorder(const HeaderRect: TRect);
begin
  with HeaderRect do
  begin
    Canvas.MoveTo(HeaderRect.Right, HeaderRect.Top);
    Canvas.LineTo(HeaderRect.Right, HeaderRect.Bottom);
  end;
end;

procedure TvgrWBColsPanel.SetSelection(ANumber1, ANumber2: Integer; ACurRectToMin: Boolean);
var
  ASide: TvgrBorderSide;
begin
  if ACurRectToMin then
    ASide := vgrbsLeft
  else
    ASide := vgrbsRight;
  InternalSetSelection(Rect(ANumber1, 0,
                            ANumber2,
                            Max(WorkbookGrid.FSheet.DimensionsBottom,
                                WorkbookGrid.FullVisibleRowsCount - 1)), ASide);
end;

procedure TvgrWBColsPanel.GetPointInfoAt(x,y: integer; var PointInfo: rvgrHeaderPointInfo);
begin
  InternalGetPointInfoAt(x,PointInfo);
end;

function TvgrWBColsPanel.GetOptions: TvgrOptionsHeader;
begin
  Result := WorkbookGrid.OptionsCols;
end;

function TvgrWBColsPanel.GetHeaderSize(Number: Integer): Integer;
begin
  Result := WBSheet.ColPixelWidth(Number);
end;

function TvgrWBColsPanel.GetHeaderCaption(Number: Integer): string;
begin
  Result := GetWorksheetColCaption(Number);
end;

function TvgrWBColsPanel.GetHeaderRect(From,Size: Integer): TRect;
begin
  Result := Rect(From, 0, From + Size - 1, ClientHeight - 1);
end;

function TvgrWBColsPanel.GetHeaderHasSelection(Number: Integer): boolean;
begin
  Result := WBSheet.FSelectedRegion.ColHasSelection(Number);
end;

function TvgrWBColsPanel.GetVisibleHeadersCount: Integer;
begin
  Result := WorkbookGrid.FVisibleColsCount;
end;

function TvgrWBColsPanel.GetResizeMouseCursor: TCursor;
begin
  Result := crHSplit;
end;

function TvgrWBColsPanel.GetInsideMouseCursor: TCursor;
begin
  Result := FToBottomCursor;
end;

procedure TvgrWBColsPanel.DoHeaderResize(Number, DeltaX, DeltaY: Integer);
var
  ACol: IvgrCol;
  AColWidth: Integer;
begin
  ACol := WBSheet.CheckAddCol(Number);
  if ACol.Visible then
    AColWidth := ACol.Width
  else
  begin
    ACol.Visible := True;
    AColWidth := 0;
  end;
  ACol.Width := Max(0, AColWidth + DeltaX);
end;

procedure TvgrWBColsPanel.DoSelectionChanged;
begin
  Invalidate;
end;

procedure TvgrWBColsPanel.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcDeleteRow:
      FInvalidateRect := Rect(WBSheet.ColPixelLeft((ChangeInfo.ChangedInterface as IvgrCol).Number) - WorkbookGrid.LeftTopPixelsOffset.x,
                              0, ClientWidth, ClientHeight);
  end;
end;


procedure TvgrWBColsPanel.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
var
  r: TRect;
begin
  case ChangeInfo.ChangesType of
    vgrwcChangeCol:
      begin
        r := Rect(WBSheet.ColPixelLeft((ChangeInfo.ChangedInterface as IvgrCol).Number) - WorkbookGrid.LeftTopPixelsOffset.x,
                  0, ClientWidth, ClientHeight);
        InvalidateRect(Handle, @r, false);
      end;
    vgrwcDeleteCol:
      InvalidateRect(Handle, @FInvalidateRect, false);
    vgrwcUpdateAll, vgrwcChangeWorksheetContent:
      Invalidate;
  end;
end;

function TvgrWBColsPanel.GetVectors: TvgrVectors;
begin
  if WorkbookGrid.ActiveWorksheet <> nil then
    Result := WorkbookGrid.ActiveWorksheet.ColsList
  else
    Result := nil;
end;

function TvgrWBColsPanel.GetPointInfoClass: TvgrHeadersPanelPointInfoClass;
begin
  Result := TvgrColsPanelPointInfo;
end;

/////////////////////////////////////////////////
//
// TvgrRowsPanelPointInfo
//
/////////////////////////////////////////////////

/////////////////////////////////////////////////
//
// TvgrWBRowsPanel
//
/////////////////////////////////////////////////
procedure TvgrWBRowsPanel.OptionsChanged;
begin
  if Options.Visible then
    Width := cRowsPanelHorzOffset + GetStringWidth(sDefStringForWidth, Options.Font) * 6
  else
    Width := 0;
  inherited;
end;

procedure TvgrWBRowsPanel.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcDeleteRow:
      FInvalidateRect := Rect(0,
                              WBSheet.RowPixelTop((ChangeInfo.ChangedInterface as IvgrRow).Number) - WorkbookGrid.LeftTopPixelsOffset.y,
                              ClientWidth,
                              ClientHeight);
  end;
end;

procedure TvgrWBRowsPanel.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
var
  r: TRect;
begin
  case ChangeInfo.ChangesType of
    vgrwcChangeRow:
      begin
        r := Rect(0,
                  WBSheet.RowPixelTop((ChangeInfo.ChangedInterface as IvgrRow).Number) - WorkbookGrid.LeftTopPixelsOffset.y,
                  ClientWidth,
                  ClientHeight);
        InvalidateRect(Handle, @r, False);
      end;
    vgrwcDeleteRow:
      InvalidateRect(Handle, @FInvalidateRect, False);
    vgrwcUpdateAll, vgrwcChangeWorksheetContent:
      Invalidate;
  end;
end;

function TvgrWBRowsPanel.GetWBScrollBar : TvgrScrollBar;
begin
  Result := WorkbookGrid.VertScrollBar;
end;

procedure TvgrWBRowsPanel.GetRectCoords(const ARect: TRect; var AStart, AEnd: Integer);
begin
  AStart := ARect.Top;
  AEnd := ARect.Bottom;
end;

function TvgrWBRowsPanel.GetCoord(X, Y: Integer): Integer;
begin
  Result := Y;
end;

function TvgrWBRowsPanel.GetXCoord(ACoord: Integer): Integer;
begin
  Result := 0;
end;

function TvgrWBRowsPanel.GetYCoord(ACoord: Integer): Integer;
begin
  Result := ACoord;
end;

procedure TvgrWBRowsPanel.Paint;
var
  rClip, rRegions, rPixels: TRect;
begin
  rClip := Canvas.ClipRect;
  rClip := Canvas.ClipRect;
  rClip.Left := 0;
  rClip.Right := ClientWidth;
  WorkbookGrid.GetPaintRects(rClip, rRegions, rPixels, WorkbookGrid.LeftCol, WorkbookGrid.TopRow);
  InternalPaint(rRegions.Top,rRegions.Bottom,rPixels.Top,rPixels.Bottom);

  DrawBorders(Canvas,
              Rect(0, 0, ClientWidth, ClientHeight),
              WorkbookGrid.FHorzSections.Visible and (WorkbookGrid.FHorzSections.Width > 0),
              False,
              True,
              False,
              clBtnShadow);
end;

procedure TvgrWBRowsPanel.Resize;
begin
  inherited;
  WorkbookGrid.FHorzSections.RealignLevelButtons;
end;

procedure TvgrWBRowsPanel.PaintHeaderBorder(const HeaderRect: TRect);
begin
  with HeaderRect do
  begin
    Canvas.MoveTo(HeaderRect.Left, HeaderRect.Bottom);
    Canvas.LineTo(HeaderRect.Right, HeaderRect.Bottom);
  end;
end;

procedure TvgrWBRowsPanel.SetSelection(ANumber1, ANumber2: Integer; ACurRectToMin: Boolean);
var
  ASide: TvgrBorderSide;
begin
  if ACurRectToMin then
    ASide := vgrbsTop
  else
    ASide := vgrbsBottom;
  InternalSetSelection(Rect(0, ANumber1,
                            Max(WorkbookGrid.FSheet.DimensionsRight,
                                WorkbookGrid.FullVisibleColsCount - 1),
                            ANumber2), ASide);
end;

procedure TvgrWBRowsPanel.GetPointInfoAt(x,y: integer; var PointInfo: rvgrHeaderPointInfo);
begin
  InternalGetPointInfoAt(y,PointInfo);
end;

function TvgrWBRowsPanel.GetOptions: TvgrOptionsHeader;
begin
  Result := WorkbookGrid.OptionsRows;
end;

function TvgrWBRowsPanel.GetHeaderSize(Number: Integer): Integer;
begin
  Result := WBSheet.RowPixelHeight(Number);
end;

function TvgrWBRowsPanel.GetHeaderCaption(Number: Integer): string;
begin
  Result := GetWorksheetRowCaption(Number);
end;

function TvgrWBRowsPanel.GetHeaderRect(From,Size: Integer): TRect; 
begin
  Result := Rect(0, From, ClientWidth - 1, From + Size - 1);
end;

function TvgrWBRowsPanel.GetHeaderHasSelection(Number: Integer): boolean;
begin
  Result := WBSheet.FSelectedRegion.RowHasSelection(Number);
end;

function TvgrWBRowsPanel.GetVisibleHeadersCount: Integer;
begin
  Result := WorkbookGrid.FVisibleRowsCount;
end;

function TvgrWBRowsPanel.GetResizeMouseCursor: TCursor;
begin
  Result := crVSplit;
end;

function TvgrWBRowsPanel.GetInsideMouseCursor: TCursor;
begin
  Result := FToRightCursor;
end;

procedure TvgrWBRowsPanel.DoHeaderResize(Number, DeltaX, DeltaY: Integer);
var
  ARow: IvgrRow;
  ARowHeight: Integer;
begin
  ARow := WBSheet.CheckAddRow(Number);
  if ARow.Visible then
    ARowHeight := ARow.Height
  else
  begin
    ARow.Visible := True;
    ARowHeight := 0;
  end;
  ARow.Height := Max(0, ARowHeight + DeltaY);
end;

procedure TvgrWBRowsPanel.DoSelectionChanged;
begin
  Invalidate;
end;

function TvgrWBRowsPanel.GetVectors: TvgrVectors;
begin
  if WorkbookGrid.ActiveWorksheet <> nil then
    Result := WorkbookGrid.ActiveWorksheet.RowsList
  else
    Result := nil;
end;

function TvgrWBRowsPanel.GetPointInfoClass: TvgrHeadersPanelPointInfoClass;
begin
  Result := TvgrRowsPanelPointInfo;
end;

/////////////////////////////////////////////////
//
// TvgrSectionsCacheList
//
/////////////////////////////////////////////////
destructor TvgrSectionsCacheList.Destroy;
begin
  Clear;
  inherited;
end;

function TvgrSectionsCacheList.GetItem(Index: Integer): pvgrSectionCacheInfo;
begin
  Result := pvgrSectionCacheInfo(inherited Items[Index]);
end;

procedure TvgrSectionsCacheList.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    FreeMem(Items[I]);
  inherited;
end;

function TvgrSectionsCacheList.Add(ASection: pvgrSection; ASectionIndex: Integer): pvgrSectionCacheInfo;
begin
  GetMem(Result, SizeOf(rvgrSectionCacheInfo));
  Result.Section := ASection;
  Result.SectionIndex := ASectionIndex;
  inherited Add(Result);
end;

function _SortByLevel(Item1, Item2: Pointer): Integer;
begin
  Result := pvgrSectionCacheInfo(Item1).Section.Level - pvgrSectionCacheInfo(Item2).Section.Level; 
end;

procedure TvgrSectionsCacheList.SortByLevel;
begin
  Sort(_SortByLevel);
end;

/////////////////////////////////////////////////
//
// TvgrSectionsPanelPointInfo
//
/////////////////////////////////////////////////
constructor TvgrSectionsPanelPointInfo.Create(APlace: TvgrWBSectionsPlace; ASections: TvgrSections; ASectionIndex: Integer; ALevelButton: TvgrWBLevelButton);
begin
  inherited Create;
  FPlace := APlace;
  FSections := ASections;
  FSectionIndex := ASectionIndex;
  FLevelButton := ALevelButton;
end;

/////////////////////////////////////////////////
//
// TvgrWBSectionsPanel
//
/////////////////////////////////////////////////
constructor TvgrWBSectionsPanel.Create(AOwner: TComponent);
begin
  inherited;
  FLevelButtonsFont := TFont.Create;
  with FLevelButtonsFont do
  begin
    Name := sSectionLevelButtonFontName;
    Size := 7;
    Color := clBtnText;
  end;
  FLevelButtons := TList.Create;
  FPaintSections := TvgrSectionsCacheList.Create;
  FMouseSections := TvgrSectionsCacheList.Create;
  FMovePointInfo.SectionIndex := -1;
  FDownPointInfo.SectionIndex := -1;
  Font.OnChange := DoOnFontChange;
  DoOnFontchange(Font);
end;

destructor TvgrWBSectionsPanel.Destroy;
begin
  FPaintSections.Free;
  FMouseSections.Free;
  ClearLevelButtons;
  FLevelButtons.Free;
  FLevelButtonsFont.Free;
  inherited;
end;

procedure TvgrWBSectionsPanel.DoScroll(AOldPos, ANewPos: Integer);
begin
  Invalidate;
end;

function TvgrWBSectionsPanel.GetWorkbookGrid: TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

function TvgrWBSectionsPanel.GetWBSheet: TvgrWBSheet;
begin
  Result := WorkbookGrid.FSheet;
end;

function TvgrWBSectionsPanel.GetWorksheet: TvgrWorksheet;
begin
  Result := WBSheet.Worksheet;
end;

function TvgrWBSectionsPanel.GetLevelButton(Index: Integer): TvgrWBLevelButton;
begin
  Result := TvgrWBLevelButton(FLevelButtons[Index]);
end;

function TvgrWBSectionsPanel.GetLevelButtonCount: Integer;
begin
  Result := FLevelButtons.Count;
end;

procedure TvgrWBSectionsPanel.ClearLevelButtons;
var
  I: Integer;
begin
  for I := 0 to FLevelButtons.Count - 1 do
    LevelButtons[I].Free;
  FLevelButtons.Clear;
end;                      

procedure TvgrWBSectionsPanel.GetLevelCountCallbackProc(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
    if Level > FMaxLevel then
      FMaxLevel := Level;
end;

function TvgrWBSectionsPanel.GetLevelCount: Integer;
begin
  FMaxLevel := -1;
  if Sections <> nil then
    Sections.FindAndCallBack(0, cSectionsMaxEndPos, GetLevelCountCallbackProc, nil);
  Result := FMaxLevel + 1;
end;

procedure TvgrWBSectionsPanel.GetSectionCountAtLevelCallbackProc(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
    if (Level = FCallbackLevel) and IsSectionHidden(ItemData) then
      Inc(FCountAtLevel);
end;

function TvgrWBSectionsPanel.GetHiddenSectionCountAtLevel(ALevel: Integer): Integer;
begin
  FCountAtLevel := 0;
  FCallbackLevel := ALevel;
  if Sections <> nil then
    Sections.FindAndCallBack(0, cSectionsMaxEndPos, GetSectionCountAtLevelCallbackProc, nil);
  Result := FCountAtLevel;
end;

function TvgrWBSectionsPanel.GetPanelSize: Integer;
var
  ALevelCount: Integer;
begin
  ALevelCount := GetLevelCount;
  if ALevelCount = 0 then
    Result := 0
  else
    Result := ALevelCount * FLevelSize + 4;
end;

procedure TvgrWBSectionsPanel.DoOnFontChange(Sender: TObject);
begin
  FLevelSize := Max(GetFontHeight(Font) + cSectionsPanelLevelOffset,
                    cSectionButtonWidth * 2 + 1);
end;

procedure TvgrWBSectionsPanel.UpdateSize;
var
  I, ALevelCount: Integer;
begin
  // add / remove level buttons
  ALevelCount := GetLevelCount;
  if LevelButtonCount > ALevelCount then
  begin
    for I := ALevelCount to LevelButtonCount - 1 do
      LevelButtons[I].Free;
    FLevelButtons.Count := ALevelCount;
  end
  else
    if LevelButtonCount < ALevelCount then
      for I := 0 to ALevelCount - LevelButtonCount - 1 do
        FLevelButtons.Add(TvgrWBLevelButton.Create);
  for I := 0 to LevelButtonCount - 1 do
    LevelButtons[I].Level := I;
end;

function TvgrWBSectionsPanel.GetPaintCellStart(ANumber: Integer): Integer;
begin
  Result := FCellPoints[ANumber - FMinStartPos].Pos;
end;

function TvgrWBSectionsPanel.GetPaintCellEnd(ANumber: Integer): Integer;
begin
  with FCellPoints[ANumber - FMinStartPos] do
    Result := Pos + Size;
end;

function TvgrWBSectionsPanel.GetLevelStart(ALevel: Integer): Integer;
begin
  Result := GetMinSize - (ALevel + 1) * FLevelSize;
end;

function TvgrWBSectionsPanel.GetLevelEnd(ALevel: Integer): Integer;
begin
  Result := GetMinSize - ALevel * FLevelSize;
end;

procedure TvgrWBSectionsPanel.PaintUpdateCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
  begin
    if StartPos < FMinStartPos then
      FMinStartPos := StartPos;
    if EndPos > FMaxEndPos then
      FMaxEndPos := EndPos;
    // Add founded section in List
    FPaintSections.Add((AItem as IvgrSection).ItemData, AItemIndex);
  end;
end;

procedure TvgrWBSectionsPanel.PaintUpdate;
begin
  FMinStartPos := WBScrollBar.Position;
  FMaxEndPos := WBScrollBar.Position + GetVisibleCellsCount;

  FPaintSections.Clear;

  Sections.FindAndCallBack(WBScrollBar.Position, WBScrollBar.Position + GetVisibleCellsCount, PaintUpdateCallback, nil);

  SetLength(FCellPoints, FMaxEndPos - FMinStartPos + 2);
  CalculateCellPoints;
  FPaintSections.SortByLevel;
end;

procedure TvgrWBSectionsPanel.PaintSection(ACanvas: TCanvas; ASection: pvgrSection; ASectionIndex: Integer);
begin
end;

procedure TvgrWBSectionsPanel.PaintLevelButton(ACanvas: TCanvas; AButton: TvgrWBLevelButton);
var
  ARect: TRect;
  ATextSize: TSize;
begin
  ARect := AButton.BoundsRect;
  if (FMovePointInfo.Place = vgrsspLevelButton) and (AButton = FMovePointInfo.LevelButton) then
  begin
    ACanvas.Pen.Color := GetHighlightColor(clBtnShadow);
    ACanvas.Brush.Color := GetHighlightColor(clBtnFace);
  end
  else
  begin
    ACanvas.Pen.Color := clBtnShadow;
    ACanvas.Brush.Color := clBtnFace;
  end;
  Canvas.Rectangle(ARect);
  InflateRect(ARect, -1 ,-1);

  Canvas.Brush.Style := bsClear;
  Canvas.Font.Assign(FLevelButtonsFont);
  if GetHiddenSectionCountAtLevel(AButton.Level) > 0 then
    Canvas.Font.Color := clHighlightText
  else
    Canvas.Font.Color := clBtnText;
  ATextSize := Canvas.TextExtent(IntToStr(AButton.Level));
  Canvas.TextRect(ARect,
                  ARect.Left + (ARect.Right - ARect.Left - ATextSize.cx) div 2,
                  ARect.Top + (ARect.Bottom - ARect.Top - ATextSize.cy) div 2,
                  IntToStr(AButton.Level));
  Canvas.Brush.Style := bsSolid;
  ExcludeClipRect(ACanvas, AButton.BoundsRect);
end;

procedure TvgrWBSectionsPanel.PaintSectionButton(ACanvas: TCanvas; const ARect: TRect; AImageIndex: Integer; AHighlighted: Boolean);
begin
  if AHighlighted then
  begin
    ACanvas.Pen.Color := GetHighlightColor(clBtnShadow);
    ACanvas.Brush.Color := GetHighlightColor(clBtnFace);
  end
  else
  begin
    ACanvas.Pen.Color := clBtnShadow;
    ACanvas.Brush.Color := clBtnFace;
  end;
  ACanvas.Rectangle(ARect);
  FSectionBitmaps.Draw(ACanvas,
                       ARect.Left + 1,
                       ARect.Top + 1,
                       AImageIndex,
                       True);
  ExcludeClipRect(ACanvas, ARect);
end;

procedure TvgrWBSectionsPanel.PaintLine(ACanvas: TCanvas; ALeft, ATop, AWidth, AHeight: Integer);
var
  ARect: TRect;
begin
  ARect := Bounds(ALeft, ATop, AWidth, AHeight);
  ACanvas.FillRect(ARect);
  ExcludeClipRect(ACanvas, ARect);
end;

procedure TvgrWBSectionsPanel.Paint;
var
  I: Integer;
begin
  if Sections <> nil then
  begin
    // 1. Paint level buttons
    for I := 0 to LevelButtonCount - 1 do
      PaintLevelButton(Canvas, LevelButtons[I]);
    PaintButtonsArea(Canvas);

    // 2. Cache cells coords
    PaintUpdate;

    // 3. Enumerates all sections and Paint them
    for I := 0 to FPaintSections.Count - 1 do
      with FPaintSections[I]^ do
        PaintSection(Canvas, Section, SectionIndex);
  end;

  PaintLines(Canvas);
  // fill background and draw borders
  DrawBorders(Canvas, Rect(0, 0, ClientWidth, ClientHeight),
              PaintGetBorders,
              clBtnShadow,
              clBtnFace); // Fill background
end;

procedure TvgrWBSectionsPanel.GetPointInfoAtCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  FMouseSections.Add((AItem as IvgrSection).ItemData, AItemIndex);
end;

procedure TvgrWBSectionsPanel.GetPointInfoAt(X, Y: Integer; var PointInfo: rvgrWBSectionsPointInfo);
var
  I: Integer;
  APoint: TPoint;
  ARect, ACloseRect, AHideRect: TRect;
begin
  APoint := Point(X, Y);
  with PointInfo do
  begin
    Place := vgrsspNone;
    Section := nil;
    SectionIndex := -1;
    LevelButton := nil;
  end;

  // 1. check level buttons
  for I := 0 to LevelButtonCount - 1 do
    if PtInRect(LevelButtons[I].BoundsRect, APoint) then
    begin
      PointInfo.Place := vgrsspLevelButton;
      PointInfo.LevelButton := LevelButtons[I];
      exit;
    end;

  // 2. check sections
  FMouseSections.Clear;
  Sections.FindAndCallBack(WBScrollBar.Position, WBScrollBar.Position + GetVisibleCellsCount, GetPointInfoAtCallback, nil);
  FMouseSections.SortByLevel;
  for I := 0 to FMouseSections.Count - 1 do
  begin
    ARect := GetSectionRect(FMouseSections[I].Section);
    if PtInRect(ARect, APoint) then
    begin
      PointInfo.SectionIndex := FMouseSections[I].SectionIndex;
      PointInfo.Section := FMouseSections[I].Section;
      GetSectionButtonRects(ARect, ACloseRect, AHideRect);
      if PtInRect(ACloseRect, APoint) then
        PointInfo.Place := vgrsspSectionClose
      else
        if PtInRect(AHideRect, APoint) then
          PointInfo.Place := vgrsspSectionHide
        else
          PointInfo.Place := vgrsspSection;
      break;
    end
    else
      PointInfo.Place := vgrsspNone;
  end;
end;

procedure TvgrWBSectionsPanel.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APointInfo: TvgrPointInfo;
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);

  if AExecuteDefault then
  begin
    GetPointInfoAt(X, Y, FDownPointInfo);
    case FDownPointInfo.Place of
      vgrsspSectionHide:
        begin
          if (Button = mbLeft) and not WorkbookGrid.ReadOnly then
            DoSectionHide(FDownPointInfo.Section, FDownPointInfo.SectionIndex);
        end;
      vgrsspSectionClose:
        begin
          if (Button = mbLeft) and not WorkbookGrid.ReadOnly then
            DoSectionDelete(FDownPointInfo.Section, FDownPointInfo.SectionIndex);
        end;
      vgrsspLevelButton:
        begin
          if (Button = mbLeft) and not WorkbookGrid.ReadOnly then
            DoLevelButtonClick(FDownPointInfo.LevelButton);
        end;
      vgrsspSection:
        begin
          if (mbLeft = Button) and (ssDouble in Shift) then
            WorkbookGrid.DoSectionDblClick(GetSections, FDownPointInfo.SectionIndex)
          else
            if mbRight = Button then
            begin
              GetPointInfoAt(Point(X, Y), APointInfo);
              WorkbookGrid.ShowPopupMenu(ClientToScreen(Point(X, Y)), APointInfo);
              APointInfo.Free;
            end;
        end;
    end;
  end;
end;

procedure TvgrWBSectionsPanel.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoMouseUp(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrWBSectionsPanel.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  APointInfo: rvgrWBSectionsPointInfo;
begin
  GetPointInfoAt(X, Y, APointInfo);
  if not CompareMem(@APointInfo, @FMovePointInfo, SizeOf(rvgrWBSectionsPointInfo)) then
  begin
    Invalidate;
    FMovePointInfo := APointInfo;
  end;
  WorkbookGrid.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrWBSectionsPanel.DblClick;
var
  AExecuteDefault: Boolean;
begin
  AExecuteDefault := True;
  WorkbookGrid.DoDblClick(Self, AExecuteDefault);
end;

procedure TvgrWBSectionsPanel.CMMouseLeave(var Msg: TMessage);
begin
  FMovePointInfo.SectionIndex := -1;
  FMovePointInfo.Place := vgrsspNone;
  FMovePointInfo.Section := nil;
  Invalidate;
end;

procedure TvgrWBSectionsPanel.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
end;

procedure TvgrWBSectionsPanel.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcChangeHorzSection, vgrwcChangeVertSection, vgrwcNewHorzSection, vgrwcDeleteHorzSection,
    vgrwcNewVertSection, vgrwcDeleteVertSection, vgrwcUpdateAll,
    vgrwcChangeWorksheetContent:
    begin
      UpdateSize;
      Invalidate;
    end;
    vgrwcChangeCol, vgrwcChangeRow, vgrwcDeleteCol, vgrwcDeleteRow, vgrwcNewCol, vgrwcNewRow:
    begin
      Invalidate;
    end;
  end;
end;

procedure TvgrWBSectionsPanel.DoLevelButtonClickCallbackProc(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrSection do
    if Level = FCallbackLevel then
      DoSectionShow(ItemData);
end;

procedure TvgrWBSectionsPanel.DoLevelButtonClick(ALevelButton: TvgrWBLevelButton);
begin
  // make visible all hidden section at level ALevelButton.Level
  FCallbackLevel := ALevelButton.Level;
  Sections.FindAndCallBack(0, cSectionsMaxEndPos, DoLevelButtonClickCallbackProc, nil);
end;

procedure TvgrWBSectionsPanel.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
var
  PointInfo: rvgrWBSectionsPointInfo;
begin
  GetPointInfoAt(APoint.X, APoint.Y, PointInfo);
  with PointInfo do
    APointInfo := GetPointInfoClass.Create(Place, Sections, SectionIndex, LevelButton);
end;

procedure TvgrWBSectionsPanel.DragDrop(Source: TObject; X, Y: Integer);
begin
  WorkbookGrid.DoDragDrop(Self, Source, X, Y);
end;

procedure TvgrWBSectionsPanel.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  WorkbookGrid.DoDragOver(Self, Source, X, Y, State, Accept);
end;

/////////////////////////////////////////////////
//
// TvgrWBVertSectionsPanel
//
/////////////////////////////////////////////////
constructor TvgrWBVertSectionsPanel.Create(AOwner: TComponent);
begin
  inherited;
  Height := 0;
end;

function TvgrWBVertSectionsPanel.GetWBScrollBar: TvgrScrollBar;
begin
  Result := WBSheet.HorzScrollBar;
end;

function TvgrWBVertSectionsPanel.GetSections: TvgrSections;
begin
  if Worksheet = nil then
    Result := nil
  else
    Result := Worksheet.VertSectionsList;
end;

function TvgrWBVertSectionsPanel.GetVisibleCellsCount: Integer;
begin
  Result := WorkbookGrid.FVisibleColsCount; 
end;

function TvgrWBVertSectionsPanel.GetMinSize: Integer;
begin
  Result := Height;
end;

function TvgrWBVertSectionsPanel.GetSectionRect(ASection: pvgrSection): TRect;
begin
  Result.Left := WBSheet.ColScreenPixelLeft(ASection.StartPos) + WorkbookGrid.FRowsPanel.Width;
  Result.Top := GetLevelStart(ASection.Level);
  Result.Right := WBSheet.ColScreenPixelRight(ASection.EndPos) + WorkbookGrid.FRowsPanel.Width;
  Result.Bottom := Height;
end;

procedure TvgrWBVertSectionsPanel.GetSectionButtonRects(const ASectionRect: TRect; var ACloseButton, AHideButton: TRect);
begin
  with ASectionRect do
  begin
    ACloseButton := Rect(Right - cSectionButtonWidth - 4, Top + 4, Right - 4, Top + 4 + cSectionButtonHeight);
    if ACloseButton.Left < Left then
    begin
      ACloseButton.Left := Left;
      ACloseButton.Right := Left;
    end;
    AHideButton := Rect(ACloseButton.Left - cSectionButtonWidth - 1, Top + 4, ACloseButton.Left - 1, Top + 4 + cSectionButtonHeight);
    if AHideButton.Left < Left then
    begin
      AHideButton.Left := Left;
      AHideButton.Right := Left;
    end;
  end;
end;

function TvgrWBVertSectionsPanel.IsSectionHidden(ASection: pvgrSection): Boolean;
var
  I: Integer;
begin
  for I := ASection.StartPos to ASection.EndPos do
    if WBSheet.ColVisible(I) then
    begin
      Result := False;
      exit;
    end;
  Result := True;
end;

procedure TvgrWBVertSectionsPanel.DoOnFontChange(Sender: TObject);
begin
  inherited;
  Height := GetPanelSize;
end;

procedure TvgrWBVertSectionsPanel.CalculateCellPoints;
var
  I: Integer;
begin
  with FCellPoints[0] do
  begin
    Pos := WBSheet.ColPixelLeft(FMinStartPos) - WorkbookGrid.FLeftTopPixelsOffset.X + WorkbookGrid.FRowsPanel.Width;
    Size := WBSheet.ColPixelWidth(FMinStartPos);
  end;
  for I := 1 to FMaxEndPos - FMinStartPos + 1 do
    with FCellPoints[I] do
    begin
      Pos := FCellPoints[I - 1].Pos + FCellPoints[I - 1].Size;
      Size := WBSheet.ColPixelWidth(I + FMinStartPos);
    end;
end;

procedure TvgrWBVertSectionsPanel.PaintSection(ACanvas: TCanvas; ASection: pvgrSection; ASectionIndex: Integer);
var
  s: string;
  ASize: TSize;
  ASectionIntf: IvgrSection;
  ASectionExtIntf: IvgrSectionExt;
  AInternalRect, ARect, ACloseRect, AHideRect: TRect;
begin
  // draw border
  ARect.Left := GetPaintCellStart(ASection.StartPos);
  ARect.Right := GetPaintCellEnd(ASection.EndPos);
  if ARect.Left <> ARect.Right then
  begin
    ARect.Top := GetLevelStart(ASection.Level);
    ARect.Bottom := Height;

    AInternalRect := ARect;
    DrawEdge(ACanvas.Handle, AInternalRect, EDGE_RAISED{EDGE_SUNKEN}, BF_ADJUST or BF_LEFT or BF_TOP or BF_RIGHT);

    // draw buttons
    GetSectionButtonRects(ARect, ACloseRect, AHideRect);
    PaintSectionButton(ACanvas, ACloseRect, 1, (FMovePointInfo.SectionIndex = ASectionIndex) and
                                               (FMovePointInfo.Place = vgrsspSectionClose));
    PaintSectionButton(ACanvas, AHideRect, 0, (FMovePointInfo.SectionIndex = ASectionIndex) and
                                              (FMovePointInfo.Place = vgrsspSectionHide));

    if FMovePointInfo.Section = ASection then
    begin
      ACanvas.Brush.Color := clBtnHighlight;
      ACanvas.Font.Color := clBtnText;
    end
    else
    begin
      ACanvas.Brush.Color := clBtnFace;
      ACanvas.Font.Color := clBtnText;
    end;
    ACanvas.FillRect(AInternalRect);

    // Draw section text
    AInternalRect.Left := AInternalRect.Left + 2;
    AInternalRect.Top := AInternalRect.Top + 2;
    AInternalRect.Right := AInternalRect.Right - 2;

    ASectionIntf := Sections.ByIndex[ASectionIndex];
    ASectionIntf.QueryInterface(svgrSectionExtGUID, ASectionExtIntf);
    if ASectionExtIntf <> nil then
    begin
      s := ASectionExtIntf.Text;
      ASize.cx := AInternalRect.Right - AInternalRect.Left;
      ASize.cy := AInternalRect.Bottom - AInternalRect.Top; 
      DrawRotatedText(ACanvas, S,
                      AInternalRect,
                      0,
                      True,
                      vgrhaLeft,
                      vgrvaTop);
    end;

    ExcludeClipRect(ACanvas, ARect);
  end;
end;

function TvgrWBVertSectionsPanel.PaintGetBorders: Cardinal;
begin
  if WorkbookGrid.FFormulaPanel.Height > 0 then
    Result := BF_TOP
  else
    Result := 0;
end;

procedure TvgrWBVertSectionsPanel.PaintButtonsArea(ACanvas: TCanvas);
var
  ARect: TRect;
begin
  ARect := Rect(0, 0, WorkbookGrid.FRowsPanel.Width, Height);
  DrawBorders(ACanvas,
              ARect,
              True,
              WorkbookGrid.FFormulaPanel.Height > 0,
              True,
              False,
              clBtnShadow, clBtnHighlight);
  ExcludeClipRect(Canvas, ARect);
end;

procedure TvgrWBVertSectionsPanel.PaintLines(ACanvas: TCanvas);
var
  I: Integer;
begin
  Canvas.Brush.Color := clBtnShadow;
  for I := 0 to LevelButtonCount - 2 do
    PaintLine(ACanvas, 0, GetLevelStart(I) - 1, Width, 1);
end;

procedure TvgrWBVertSectionsPanel.UpdateSize;
begin
  inherited;
  Height := GetPanelSize;
  RealignLevelButtons;
end;

procedure TvgrWBVertSectionsPanel.RealignLevelButtons;
var
  I: Integer;
begin
  for I := 0 to LevelButtonCount - 1 do
    with LevelButtons[I].BoundsRect do
    begin
      Top := GetLevelStart(I) + cSectionLevelButtonOffset;
      Bottom := GetLevelEnd(I) - cSectionLevelButtonOffset;
      Left := (WorkbookGrid.FRowsPanel.Width - (Bottom - Top)) div 2;
      Right := Left + (Bottom - Top);
    end;
end;

procedure TvgrWBVertSectionsPanel.DoSectionHide(ASection: pvgrSection; ASectionIndex: Integer);
var
  I: Integer;
begin
  for I := ASection.StartPos to ASection.EndPos do
    WBSheet.CheckAddCol(I).Visible := False;
end;

procedure TvgrWBVertSectionsPanel.DoSectionDelete(ASection: pvgrSection; ASectionIndex: Integer);
begin
  Sections.Delete(ASection.StartPos, ASection.EndPos);
end;

procedure TvgrWBVertSectionsPanel.DoSectionShow(ASection: pvgrSection);
var
  I: Integer;
begin
  for I := ASection.StartPos to ASection.EndPos do
    WBSheet.CheckAddCol(I).Visible := True;
end;

function TvgrWBVertSectionsPanel.GetPointInfoClass: TvgrSectionsPanelPointInfoClass;
begin
  Result := TvgrVertSectionsPanelPointInfo; 
end;

/////////////////////////////////////////////////
//
// TvgrWBHorzSectionsPanel
//
/////////////////////////////////////////////////
constructor TvgrWBHorzSectionsPanel.Create(AOwner: TComponent);
begin
  inherited;
  Width := 0;
end;

function TvgrWBHorzSectionsPanel.GetWBScrollBar: TvgrScrollBar;
begin
  Result := WBSheet.VertScrollBar;
end;

function TvgrWBHorzSectionsPanel.GetSections: TvgrSections;
begin
  if Worksheet = nil then
    Result := nil
  else
    Result := Worksheet.HorzSectionsList;
end;

function TvgrWBHorzSectionsPanel.GetVisibleCellsCount: Integer;
begin
  Result := WorkbookGrid.FVisibleRowsCount; 
end;

function TvgrWBHorzSectionsPanel.GetMinSize: Integer;
begin
  Result := Width;
end;

procedure TvgrWBHorzSectionsPanel.DoOnFontChange(Sender: TObject);
begin
  inherited;
  Width := GetPanelSize;
end;

procedure TvgrWBHorzSectionsPanel.CalculateCellPoints;
var
  I: Integer;
begin
  with FCellPoints[0] do
  begin
    Pos := WBSheet.RowPixelTop(FMinStartPos) - WorkbookGrid.FLeftTopPixelsOffset.Y + WorkbookGrid.FColsPanel.Height;
    Size := WBSheet.RowPixelHeight(FMinStartPos);
  end;
  for I := 1 to FMaxEndPos - FMinStartPos + 1 do
    with FCellPoints[I] do
    begin
      Pos := FCellPoints[I - 1].Pos + FCellPoints[I - 1].Size;
      Size := WBSheet.RowPixelHeight(I + FMinStartPos);
    end;
end;

function TvgrWBHorzSectionsPanel.IsSectionHidden(ASection: pvgrSection): Boolean;
var
  I: Integer;
begin
  for I := ASection.StartPos to ASection.EndPos do
    if WBSheet.RowVisible(I) then
    begin
      Result := False;
      exit;
    end;
  Result := True;
end;

function TvgrWBHorzSectionsPanel.GetSectionRect(ASection: pvgrSection): TRect;
begin
  Result.Left := GetLevelStart(ASection.Level);
  Result.Top := WBSheet.RowScreenPixelTop(ASection.StartPos) + WorkbookGrid.FColsPanel.Height;
  Result.Right := Width;
  Result.Bottom := WBSheet.RowScreenPixelBottom(ASection.EndPos) + WorkbookGrid.FColsPanel.Height;
end;

procedure TvgrWBHorzSectionsPanel.GetSectionButtonRects(const ASectionRect: TRect; var ACloseButton, AHideButton: TRect);
begin
  with ASectionRect do
  begin
    AHideButton := Rect(Left + 2, Top + 4, Left + 2 + cSectionButtonWidth, Top + 4 + cSectionButtonHeight);
    if AHideButton.Bottom > Bottom then
    begin
      AHideButton.Top := Top;
      AHideButton.Bottom := Top;
    end;
    ACloseButton := Rect(AHideButton.Right + 1, Top + 4, AHideButton.Right + 1 + cSectionButtonWidth, Top + 4 + cSectionButtonHeight);
    if ACloseButton.Bottom > Bottom then
    begin
      ACloseButton.Top := Top;
      ACloseButton.Bottom := Top;
    end;
  end;
end;

procedure TvgrWBHorzSectionsPanel.PaintSection(ACanvas: TCanvas; ASection: pvgrSection; ASectionIndex: Integer);
var
  s: string;
  ASize: TSize;
  ASectionIntf: IvgrSection;
  ASectionExtIntf: IvgrSectionExt;
  AInternalRect, ARect, ACloseRect, AHideRect: TRect;
begin
  // draw border
  ARect.Top := GetPaintCellStart(ASection.StartPos);
  ARect.Bottom := GetPaintCellEnd(ASection.EndPos);
  if ARect.Top <> ARect.Bottom then
  begin
    ARect.Left := GetLevelStart(ASection.Level);
    ARect.Right := Width;
    
    AInternalRect := ARect;
    DrawEdge(ACanvas.Handle, AInternalRect, EDGE_RAISED, BF_ADJUST or BF_LEFT or BF_TOP or BF_BOTTOM);

    // draw buttons
    GetSectionButtonRects(ARect, ACloseRect, AHideRect);
    PaintSectionButton(ACanvas, ACloseRect, 1, (FMovePointInfo.SectionIndex = ASectionIndex) and
                                               (FMovePointInfo.Place = vgrsspSectionClose));
    PaintSectionButton(ACanvas, AHideRect, 0, (FMovePointInfo.SectionIndex = ASectionIndex) and
                                              (FMovePointInfo.Place = vgrsspSectionHide));

    if FMovePointInfo.Section = ASection then
    begin
      ACanvas.Brush.Color := clBtnHighlight;
      ACanvas.Font.Color := clBtnText;
    end
    else
    begin
      ACanvas.Brush.Color := clBtnFace;
      ACanvas.Font.Color := clBtnText;
    end;
    ACanvas.FillRect(AInternalRect);

    // Draw section text
    AInternalRect.Left := AInternalRect.Left + 2;
    AInternalRect.Top := AInternalRect.Top + 2;
    AInternalRect.Bottom := AInternalRect.Bottom - 2;

    ASectionIntf := Sections.ByIndex[ASectionIndex];
    ASectionIntf.QueryInterface(svgrSectionExtGUID, ASectionExtIntf);
    if ASectionExtIntf <> nil then
    begin
      s := ASectionExtIntf.Text;
      ASize.cx := AInternalRect.Bottom - AInternalRect.Top;
      ASize.cy := AInternalRect.Right - AInternalRect.Left;
      DrawRotatedText(ACanvas, S,
                      AInternalRect,
                      90,
                      True,
                      vgrhaLeft,
                      vgrvaBottom{vgrvaTop});
    end;

    ExcludeClipRect(ACanvas, ARect);
  end;
end;

procedure TvgrWBHorzSectionsPanel.PaintLines(ACanvas: TCanvas);
var
  I: Integer;
begin
  Canvas.Brush.Color := clBtnShadow;
  for I := 0 to LevelButtonCount - 2 do
    PaintLine(ACanvas, GetLevelStart(I) - 1, 0, 1, Height);
end;

procedure TvgrWBHorzSectionsPanel.PaintButtonsArea(ACanvas: TCanvas);
var
  ARect: TRect;
begin
  ARect := Rect(0, 0, Width, WorkbookGrid.FColsPanel.Height);
  DrawBorders(ACanvas,
              ARect,
              False,
              True,
              False,
              True,
              clBtnShadow, clBtnHighlight);
  ExcludeClipRect(Canvas, ARect);
end;

function TvgrWBHorzSectionsPanel.PaintGetBorders: Cardinal;
begin
  Result := 0;
end;

procedure TvgrWBHorzSectionsPanel.UpdateSize;
begin
  inherited;
  Width := GetPanelSize;
  RealignLevelButtons;
end;

procedure TvgrWBHorzSectionsPanel.RealignLevelButtons;
var
  I: Integer;
begin
  for I := 0 to LevelButtonCount - 1 do
    with LevelButtons[I].BoundsRect do
    begin
      Left := GetLevelStart(I) + cSectionLevelButtonOffset;
      Right := GetLevelEnd(I) - cSectionLevelButtonOffset;
      Top := (WorkbookGrid.FColsPanel.Height - (Right - Left)) div 2;
      Bottom := Top + (Right - Left);
    end;
end;

procedure TvgrWBHorzSectionsPanel.DoSectionHide(ASection: pvgrSection; ASectionIndex: Integer);
var
  I: Integer;
begin
  for I := ASection.StartPos to ASection.EndPos do
    WBSheet.CheckAddRow(I).Visible := False;
end;

procedure TvgrWBHorzSectionsPanel.DoSectionDelete(ASection: pvgrSection; ASectionIndex: Integer);
begin
  Sections.Delete(ASection.StartPos, ASection.EndPos);
end;

procedure TvgrWBHorzSectionsPanel.DoSectionShow(ASection: pvgrSection);
var
  I: Integer;
begin
  for I := ASection.StartPos to ASection.EndPos do
    WBSheet.CheckAddRow(I).Visible := True;
end;

function TvgrWBHorzSectionsPanel.GetPointInfoClass: TvgrSectionsPanelPointInfoClass;
begin
  Result := TvgrHorzSectionsPanelPointInfo; 
end;

/////////////////////////////////////////////////
//
// TvgrWBTopLeftButton
//
/////////////////////////////////////////////////
constructor TvgrWBTopLeftButton.Create(AOwner : TComponent);
begin
  inherited;
end;

procedure TvgrWBTopLeftButton.WMEraseBkgnd(var Msg : TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWBTopLeftButton.Paint;
begin
  DrawBorders(Canvas,
              Rect(0, 0, ClientWidth, ClientHeight),
              WorkbookGrid.FHorzSections.Width > 0,
              (WorkbookGrid.FFormulaPanel.Height > 0) or
              (WorkbookGrid.FVertSections.Height > 0),
              True,
              True,
              WorkbookGrid.OptionsTopLeftButton.BorderColor,
              WorkbookGrid.OptionsTopLeftButton.BackColor);
end;

procedure TvgrWBTopLeftButton.MouseMove(Shift: TShiftState; X,Y: Integer);
begin
  WorkbookGrid.DoMouseMove(Self, Shift, X, Y);
end;

procedure TvgrWBTopLeftButton.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);

  if AExecuteDefault then
    WorkbookGrid.FSheet.SetSelection(Rect(0, 0,
                                          Max(WorkbookGrid.DimensionsRight,
                                              WorkbookGrid.FullVisibleColsCount - 1),
                                          Max(WorkbookGrid.DimensionsBottom,
                                              WorkbookGrid.FullVisibleRowsCount - 1)));
end;

procedure TvgrWBTopLeftButton.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoMouseUp(Self, Button, Shift, X, Y, AExecuteDefault);
end;

procedure TvgrWBTopLeftButton.DblClick;
var
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoDblClick(Self, AExecuteDefault);
end;

function TvgrWBTopLeftButton.GetWorkbookGrid : TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

procedure TvgrWBTopLeftButton.OptionsChanged;
begin
  Repaint;
end;

procedure TvgrWBTopLeftButton.DragDrop(Source: TObject; X, Y: Integer);
begin
  WorkbookGrid.DoDragDrop(Self, Source, X, Y);
end;

procedure TvgrWBTopLeftButton.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  WorkbookGrid.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrWBTopLeftButton.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
begin
  APointInfo := TvgrTopLeftButtonPointInfo.Create;
end;

/////////////////////////////////////////////////
//
// TvgrWBTopLeftSectionsButton
//
/////////////////////////////////////////////////
constructor TvgrWBTopLeftSectionsButton.Create(AOwner : TComponent);
begin
  inherited;
  Canvas.Brush.Color := WorkbookGrid.OptionsTopLeftButton.BackColor;
  Canvas.Pen.Color := WorkbookGrid.OptionsTopLeftButton.BorderColor;
end;

procedure TvgrWBTopLeftSectionsButton.WMEraseBkgnd(var Msg : TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWBTopLeftSectionsButton.Paint;
begin
  DrawBorders(Canvas,
              Rect(0, 0, ClientWidth, ClientHeight),
              False,
              WorkbookGrid.FFormulaPanel.Height > 0,
              False,
              False,
              WorkbookGrid.OptionsTopLeftButton.BorderColor,
              WorkbookGrid.OptionsTopLeftButton.BackColor);
end;

procedure TvgrWBTopLeftSectionsButton.MouseMove(Shift: TShiftState; X,Y: Integer);
begin
  WorkbookGrid.DoMouseMove(Self, Shift, X, Y);
end;

function TvgrWBTopLeftSectionsButton.GetWorkbookGrid : TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

procedure TvgrWBTopLeftSectionsButton.OptionsChanged;
begin
  Repaint;
end;

procedure TvgrWBTopLeftSectionsButton.DragDrop(Source: TObject; X, Y: Integer);
begin
  WorkbookGrid.DoDragDrop(Self, Source, X, Y);
end;

procedure TvgrWBTopLeftSectionsButton.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  WorkbookGrid.DoDragOver(Self, Source, X, Y, State, Accept);
end;

procedure TvgrWBTopLeftSectionsButton.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
begin
  APointInfo := TvgrTopLeftButtonSectionsPointInfo.Create;
end;

/////////////////////////////////////////////////
//
// TvgrWBInplaceEdit
//
/////////////////////////////////////////////////
constructor TvgrWBInplaceEdit.Create(AOwner: TComponent);
begin
  inherited;
  BorderStyle := bsNone;
end;

function TvgrWBInplaceEdit.GetWorkbookGrid: TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

function TvgrWBInplaceEdit.GetWBSheet: TvgrWBSheet;
begin
  Result := TvgrWBSheet(Parent);
end;

procedure TvgrWBInplaceEdit.CNCommand(var Msg: TWMCommand);
begin
  if (Msg.NotifyCode = EN_UPDATE) and WBSheet.IsLongRangeEdited then
    UpdateLongRangeSize;
  inherited;
end;

procedure TvgrWBInplaceEdit.WMSysChar(var Msg: TWMSysChar);
begin
end;

procedure TvgrWBInplaceEdit.Init;
begin
  FLastRect.Left := MaxInt;
  FLastRect.Right := 0;
end;

procedure TvgrWBInplaceEdit.UpdateLongRangeSize;
var
  ANewHeight, I, ACenter, APart, ALen, AHeight, AMaxLineWidth: Integer;
  ARect: TRect;
  AList: TStringList;
begin
  WorkbookGrid.FDisableAlign := True;

  AList := TStringList.Create;
  try
    AList.Text := Text;
    AMaxLineWidth := 0;
    for I := 0 to AList.Count - 1 do
    begin
      ALen := GetStringWidth(AList[I] + '0', Font);
      if ALen > AMaxLineWidth then
        AMaxLineWidth := ALen;
    end;
  finally
    AList.Free;
  end;

  AMaxLineWidth := Max(AMaxLineWidth, WBSheet.FInplaceEditPixelRect.Right - WBSheet.FInplaceEditPixelRect.Left);

  case WBSheet.GetInplaceRangeHorzAlignment of
    vgrhaCenter:
      begin
        with WBSheet.FInplaceEditPixelRect do
          ACenter := Left + (Right - Left) div 2;
        APart := Min(AMaxLineWidth div 2, Min(ACenter, WBSheet.ClientWidth - ACenter));
        ARect.Left := ACenter - APart;
        ARect.Right := ACenter + APart;
        FLastRect.Left := Min(FLastRect.Left, ARect.Left);
        FLastRect.Right := Max(FLastRect.Right, ARect.Right);
      end;
    vgrhaRight:
      begin
        ARect.Right := WBSheet.FInplaceEditPixelRect.Right;
        ARect.Left := Max(0, ARect.Right - AMaxLineWidth);
        FLastRect.Left := Min(FLastRect.Left, ARect.Left);
        FLastRect.Right := ARect.Right;
      end;
  else
    begin
      ARect.Left := WBSheet.FInplaceEditPixelRect.Left;
      ARect.Right := Min(WBSheet.ClientWidth, ARect.Left + AMaxLineWidth);
      FLastRect.Left := ARect.Left;
      FLastRect.Right := Max(FLastRect.Right, ARect.Right);
    end;
  end;

  with WBSheet.FInplaceEditPixelRect do
    if (Self.Left <> FLastRect.Left) or (Self.Width <> FLastRect.Right - FLastRect.Left) or
       (Self.Top <> Top) then
      SetBounds(FLastRect.Left, Top, FLastRect.Right - FLastRect.Left, Bottom - Top);

  // calculate height
  AHeight := GetFontHeight(Font) * Lines.Count;
  if Copy(Text, Length(Text) - 1, 2) = #13#10 then
    AHeight := AHeight + GetFontHeight(Font);
  ANewHeight := Max(AHeight, WBSheet.FInplaceEditPixelRect.Bottom - WBSheet.FInplaceEditPixelRect.Top);
  if (Height <> ANewHeight) and (Top + ANewHeight <= WBSheet.Height) then
    Height := ANewHeight;

  WorkbookGrid.FDisableAlign := False;
end;

procedure TvgrWBInplaceEdit.UpdatePosition;
begin
  if WBSheet.IsLongRangeEdited then
  begin
    Init;
    UpdateLongRangeSize;
  end
  else
  begin
    WorkbookGrid.FDisableAlign := True;
    with WBSheet.FInplaceEditPixelRect do
      SetBounds(Left, Top, Right - Left, Bottom - Top);
    WorkbookGrid.FDisableAlign := False;
  end;
end;

procedure TvgrWBInplaceEdit.KeyPress(var Key: Char);
begin
  inherited;
end;

procedure TvgrWBInplaceEdit.KeyDown(var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_RETURN) and (ssAlt in Shift) then
  begin
    SendMessage(Handle, WM_CHAR, 13, 0);
    Key := 0;
  end
  else
  if (Key = VK_RETURN) and (Shift = []) then
  begin
    WBSheet.EndInplaceEdit(False);
    Key := 0;
  end
  else
  if (Key = VK_ESCAPE) and (Shift = []) then
  begin
    WBSheet.EndInplaceEdit(True);
    Key := 0;
  end;

  inherited;
end;

procedure TvgrWBInplaceEdit.WMKillFocus(var Msg: TWMKillFocus);
begin
  inherited;
  if Msg.FocusedWnd <> WBSheet.FormulaPanel.Edit.Handle then
    WBSheet.EndInplaceEdit(False);
end;

procedure TvgrWBInplaceEdit.Change;
begin
  if not FDisableChange and WBSheet.WorkbookGrid.IsInplaceEdit then
    WBSheet.FormulaPanel.Edit.SetEditText(Lines.Text);
end;

procedure TvgrWBInplaceEdit.SetEditText(const S: string);
begin
  FDisableChange := True;
  Lines.Text := S;
  FDisableChange := False;
  UpdatePosition;
end;

procedure TvgrWBInplaceEdit.UpdateInternalTextRect;
var
  ARect: TRect;
begin
  if HandleAllocated then
  begin
    ARect := Rect(0, 0, Width, Height);
    SendMessage(Handle, EM_SETRECT, 0, Integer(@ARect));
  end;
end;

procedure TvgrWBInplaceEdit.CreateWnd;
begin
  inherited;
  UpdateInternalTextRect;
end;

procedure TvgrWBInplaceEdit.Resize;
begin
  inherited;
  UpdateInternalTextRect;
end;

/////////////////////////////////////////////////
//
// TvgrWBButton
//
/////////////////////////////////////////////////
procedure TvgrWBButton.CMMouseEnter(var Msg: TMessage);
begin
  Include(FState, vgrbbsUnderMouse);
  Invalidate;
end;

procedure TvgrWBButton.CMMouseLeave(var Msg: TMessage);
begin
  Exclude(FState, vgrbbsUnderMouse);
  Invalidate;
end;

procedure TvgrWBButton.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Include(FState, vgrbbsPressed);
  Invalidate;
end;

procedure TvgrWBButton.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Exclude(FState, vgrbbsPressed);
  Invalidate;
end;

procedure TvgrWBButton.Paint;
begin
  Canvas.Brush.Color := clBtnFace;
  if vgrbbsUnderMouse in FState then
  begin
    Canvas.Pen.Color := GetHighlightColor(clBtnShadow);
    Canvas.Brush.Color := GetHighlightColor(clBtnFace);
  end
  else
    Canvas.Pen.Color := clBtnFace;

  if vgrbbsPressed in FState then
    Canvas.Brush.Color := clBtnShadow;

  Canvas.Rectangle(ClientRect);

  with ClientRect do
    Canvas.Draw(Left + (Right - Left - Bitmap.Width) div 2,
                Top + (Bottom - Top - Bitmap.Height) div 2,
                Bitmap);
end;

/////////////////////////////////////////////////
//
// TvgrWBFormulaEdit
//
/////////////////////////////////////////////////
function TvgrWBFormulaEdit.GetFormulaPanel: TvgrWBFormulaPanel;
begin
  Result := TvgrWBFormulaPanel(Parent);
end;

function TvgrWBFormulaEdit.GetSheet: TvgrWBSheet;
begin
  Result := FormulaPanel.WorkbookGrid.FSheet;
end;

function TvgrWBFormulaEdit.GetInplaceEdit: TvgrWBInplaceEdit;
begin
  Result := FormulaPanel.WorkbookGrid.FSheet.InplaceEdit;
end;

procedure TvgrWBFormulaEdit.DoEnter;
begin
  if not FormulaPanel.WorkbookGrid.IsInplaceEdit then
    Sheet.StartInplaceEdit(#0, False);
end;

procedure TvgrWBFormulaEdit.KeyPress(var Key: Char);
begin
  if Key = #13 then
    Key := #0;
  inherited;
end;

procedure TvgrWBFormulaEdit.KeyDown(var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_RETURN) and (Shift = []) then
  begin
    Sheet.EndInplaceEdit(False);
    Key := 0;
  end
  else
  if (Key = VK_ESCAPE) and (Shift = []) then
  begin
    Sheet.EndInplaceEdit(True);
    Key := 0;
  end;

  inherited;
end;

procedure TvgrWBFormulaEdit.WMKillFocus(var Msg: TWMKillFocus);
begin
  inherited;
  if Msg.FocusedWnd <> Sheet.InplaceEdit.Handle then
    Sheet.EndInplaceEdit(False);
end;

procedure TvgrWBFormulaEdit.Change;
begin
  if not FDisableChange and FormulaPanel.WorkbookGrid.IsInplaceEdit then
    InplaceEdit.SetEditText(Text);
end;

procedure TvgrWBFormulaEdit.SetEditText(const S: string);
begin
  FDisableChange := True;
  Text := S;
  FDisableChange := False;
end;

/////////////////////////////////////////////////
//
// TvgrWBFormulaPanel
//
/////////////////////////////////////////////////
constructor TvgrWBFormulaPanel.Create(AOwner : TComponent);
begin
  inherited;
  FEdit := TvgrWBFormulaEdit.Create(nil);
  FEdit.BorderStyle := bsNone;
  FEdit.Parent := Self;
  FEdit.Enabled := False;
  FOkButton := TvgrWBButton.Create(nil);
  FOkButton.Visible := False;
  FOkButton.Bitmap := FFormulaOkBitmap;
  FOkButton.Parent := Self;
  FOkButton.OnClick := OnOkClick;
  FCancelButton := TvgrWBButton.Create(nil);
  FCancelButton.Bitmap := FFormulaCancelBitmap;
  FCancelButton.Visible := False;
  FCancelButton.Parent := Self;
  FCancelButton.OnClick := OnCancelClick;
  Height := 4 + GetFontHeight(FEdit.Font);
end;

destructor TvgrWBFormulaPanel.Destroy;
begin
  FreeAndNil(FEdit);
  FreeAndNil(FOkButton);
  FreeAndNil(FCancelButton);
  inherited;
end;

procedure TvgrWBFormulaPanel.DoSelectedTextChanged;
begin
  Edit.Text := WorkbookGrid.SelectedText;
end;

procedure TvgrWBFormulaPanel.CreateWnd;
begin
  inherited;
  OptionsChanged;
end;

procedure TvgrWBFormulaPanel.AlignControls(AControl: TControl; var AAlignRect: TRect);

  procedure AlignButton(ALeft: Integer; AButton: TvgrWBButton);
  begin
    AButton.SetBounds(ALeft,
                      (ClientHeight - (AButton.Bitmap.Height + 4)) div 2,
                      AButton.Bitmap.Width + 4,
                      AButton.Bitmap.Height + 4);
  end;

begin
  with BoundsRect do
  begin
    AlignButton(Left + 1, FOkButton);
    AlignButton(FOkButton.BoundsRect.Right + 1, FCancelButton);
    FEdit.SetBounds(FCancelButton.BoundsRect.Right + 1,
                    1,
                    ClientWidth - FCancelButton.BoundsRect.Right + 1,
                    Height - 2);
  end;
end;

function TvgrWBFormulaPanel.GetWorkbookGrid: TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

function TvgrWBFormulaPanel.GetOptions: TvgrOptionsFormulaPanel;
begin
  Result := WorkbookGrid.OptionsFormulaPanel;
end;

procedure TvgrWBFormulaPanel.OptionsChanged;
begin
  if Options.Visible then
  begin
    Height := 4 + GetFontHeight(FEdit.Font);
    Repaint;
  end
  else
    Height := 0;
end;

procedure TvgrWBFormulaPanel.Paint;
begin
  Canvas.Brush.Color := Options.BackColor;
  Canvas.FillRect(ClientRect);
end;

procedure TvgrWBFormulaPanel.UpdateEnabled;
begin
  FEdit.Enabled := not(csDesigning in WorkbookGrid.ComponentState) and (WorkbookGrid.Workbook <> nil) and (WorkbookGrid.ActiveWorksheet <> nil);
end;

procedure TvgrWBFormulaPanel.OnOkClick(Sender: TObject);
begin
  WorkbookGrid.EndInplaceEdit(False);
end;

procedure TvgrWBFormulaPanel.OnCancelClick(Sender: TObject);
begin
  WorkbookGrid.EndInplaceEdit(True);
end;

procedure TvgrWBFormulaPanel.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
begin
  APointInfo := TvgrFormulaPanelPointInfo.Create;
end;

/////////////////////////////////////////////////
//
// TvgrWBSelectedRegion
//
/////////////////////////////////////////////////
constructor TvgrWBSelectedRegion.Create;
begin
  inherited;
  FRectsList := TList.Create;
end;

destructor TvgrWBSelectedRegion.Destroy;
begin
  Clear;
  FRectsList.Free;
  inherited;
end;

function TvgrWBSelectedRegion.GetRects(Index: Integer): PRect;
begin
Result := PRect(FRectsList[Index]);
end;

procedure TvgrWBSelectedRegion.SetRects(Index: Integer; ARect: PRect);
begin
  PRect(FRectsList[Index])^ := ARect^;
end;

function TvgrWBSelectedRegion.GetRectsCount: integer;
begin
Result := FRectsList.Count;
end;

procedure TvgrWBSelectedRegion.Clear;
var
  i: integer;
begin
  for i:=0 to FRectsList.Count-1 do
    FreeMem(FRectsList[i]);
  FRectsList.Clear;
end;

procedure TvgrWBSelectedRegion.SetSelection(const SelectionRect: TRect);
begin
  Clear;
  AddSelection(SelectionRect);
end;

procedure TvgrWBSelectedRegion.AddSelection(const SelectionRect: TRect);
var
  r: PRect;
begin
  GetMem(r,sizeof(TRect));
  r^ := SelectionRect;
  FRectsList.Add(r);
end;

function TvgrWBSelectedRegion.RowHasSelection(RowNumber: integer): boolean;
var
  i: integer;
begin
  for i:=0 to RectsCount-1 do
    with Rects[i]^ do
      if (RowNumber >= Top) and (RowNumber <= Bottom) then
        begin
          Result := true;
          exit;
        end;
  Result := false;
end;

function TvgrWBSelectedRegion.ColHasSelection(ColNumber: integer): boolean;
var
  i: integer;
begin
  for i:=0 to RectsCount-1 do
    with Rects[i]^ do
      if (ColNumber >= Left) and (ColNumber <= Right) then
        begin
          Result := true;
          exit;
        end;
  Result := false;
end;

procedure TvgrWBSelectedRegion.Assign(ASource: TvgrWBSelectedRegion);
var
  I: Integer;
begin
  Clear;
  for I := 0 to ASource.RectsCount - 1 do
    AddSelection(ASource.Rects[I]^);
end;

/////////////////////////////////////////////////
//
// TvgrWBSheetPaintInfo
//
/////////////////////////////////////////////////
constructor TvgrWBSheetPaintInfo.Create(AWBSheet: TvgrWBSheet);
begin
  inherited Create;
  FWBSheet := AWBSheet;
end;

destructor TvgrWBSheetPaintInfo.Destroy;
begin
  Clear;
  inherited;
end;

procedure TvgrWBSheetPaintInfo.Clear;

  procedure DeleteRangeRegion(const ARanges: TvgrRangeCacheInfoArray);
  var
    I: Integer;
  begin
    for I := 0 to Length(ARanges) - 1 do
      with ARanges[I] do
        if FillBordersRegion <> 0 then
        begin
          DeleteObject(FillBordersRegion);
          FillBordersRegion := 0;
        end;
  end;
  
begin
  DeleteRangeRegion(FRanges);
  DeleteRangeRegion(FLeftLongRanges);
  DeleteRangeRegion(FRightLongRanges);

  FillChar(FRanges[0], SizeOf(rvgrRangeCacheInfo) * FRangesCount, #0);
  FillChar(FLeftLongRanges[0], SizeOf(rvgrRangeCacheInfo) * Length(FLeftLongRanges), #0);
  FillChar(FRightLongRanges[0], SizeOf(rvgrRangeCacheInfo) * Length(FRightLongRanges), #0);
  FillChar(FBorders[0], SizeOf(rvgrBorderCacheInfo) * Length(FBorders), #0);
  FillChar(FCells[0], SizeOf(rvgrCellCacheInfo) * Length(FCells), #0);
  FillChar(FVertPoints[0], SizeOf(rvgrCoordCacheInfo) * Length(FVertPoints), #0);
  FRangesCount := 0;
end;

procedure TvgrWBSheetPaintInfo.AfterPaint;
var
  I: Integer;
begin
  // 7. Finalize all strings in FRanges !!!
  for I := 0 to FRangesCount - 1 do
    Finalize(FRanges[I].DisplayText);
  for I := 0 to High(FLeftLongRanges) do
    if FLeftLongRanges[I].Style <> nil then
      Finalize(FLeftLongRanges[I].DisplayText);
  for I := 0 to High(FRightLongRanges) do
    if FRightLongRanges[I].Style <> nil then
      Finalize(FRightLongRanges[I].DisplayText);
end;

procedure TvgrWBSheetPaintInfo.ExpandCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrRange do
  begin
    if Left < FBounds.Left then
      FBounds.Left := Left;
    if Top < FBounds.Top then
      FBounds.Top := Top;
    if Right > FBounds.Right then
      FBounds.Right := Right;
    if Bottom > FBounds.Bottom then
      FBounds.Bottom := Bottom;
  end;
end;

function TvgrWBSheetPaintInfo.GetBorderInfoIndex(X, Y: Integer; AOrientation: TvgrBorderOrientation): Integer;
begin
  if (X < Bounds.Left) or (X > Bounds.Right + 1) or
     (Y < Bounds.Top) or (Y > Bounds.Bottom + 1) then
    Result := -1
  else
    Result := (Y - Bounds.Top) * (FBoundsWidth + 1) * 2 + (X - Bounds.Left) * 2 + Integer(AOrientation);
end;

procedure TvgrWBSheetPaintInfo.CalcBorderRect(ABorderInfo: pvgrBorderCacheInfo; ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation);
begin
  with ABorderInfo^ do
  begin
    if AOrientation = vgrboLeft then
    begin
      Pos := ColPixelLeft(ALeft) - 1;
      Rect.Left := Pos - (Size div 2);
      Rect.Right := Rect.Left + Size;
      Rect.Top := RowPixelTop(ATop);
      Rect.Bottom := RowPixelBottom(ATop);
    end
    else
    begin
      Rect.Left := ColPixelLeft(ALeft);
      Rect.Right := ColPixelRight(ALeft);
      Pos := RowPixelTop(ATop) - 1;
      Rect.Top := Pos - (Size div 2);
      Rect.Bottom := Rect.Top + Size;
    end;
  end;
end;

procedure TvgrWBSheetPaintInfo.UpdateBordersCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ABorder: pvgrBorder;
  ABorderStyle: pvgrBorderStyle;
  ASize: Integer;
  ABorderInfo: pvgrBorderCacheInfo;
begin
  with AItem as IvgrBorder do
  begin
    ABorder := pvgrBorder(ItemData);
    case ABorder.Orientation of
      vgrboLeft:
        if (ColPixelWidth(ABorder.Left - 1) <= 0) or
           (RowPixelHeight(ABorder.Top) <= 0) then exit;
      vgrboTop:
        if (ColPixelWidth(ABorder.Left) <= 0) or
           (RowPixelHeight(ABorder.Top - 1) <= 0) then exit;
    end;
    ABorderStyle := pvgrBorderStyle(StyleData);
  end;

  if ABorder.Orientation = vgrboLeft then
    ASize := ConvertTwipsToPixelsX(ABorderStyle.Width)
  else
    ASize := ConvertTwipsToPixelsY(ABorderStyle.Width);
  if ASize > 0 then
  begin
    // add border info
    ABorderInfo := @FBorders[GetBorderInfoIndex(ABorder.Left, ABorder.Top, ABorder.Orientation)];
    with ABorderInfo^ do
    begin
      Size := ASize;
      Border := ABorder;
      Style := ABorderStyle;
      CalcBorderRect(ABorderInfo, Border.Left, Border.Top, Border.Orientation);
    end;
  end;
end;

procedure TvgrWBSheetPaintInfo.CalcInternalRect(const ARangeRect: TRect; var AInternalRect: TRect; var ABordersRegion: HRGN);
var
  AStartRect: TRect;

  procedure CheckBorderSide(X1, Y1, X2, Y2: Integer; AOrientation: TvgrBorderOrientation; ASide: TvgrBorderSide);
  var
    X, Y, APartSize, AMaxSize, ASize: Integer;
    AFlag: Boolean;
  begin
    AMaxSize := GetBorderSize(X1, Y1, AOrientation);
    AFlag := False;
    for X := X1 to X2 do
      for Y := Y1 to Y2 do
      begin
        ASize := GetBorderSize(X1, Y1, AOrientation);
        if AMaxSize <> ASize then
        begin
          if AMaxSize < ASize then
            AMaxSize := ASize;
          AFlag := True;
        end
      end;
    APartSize := GetBorderPartSize(AMaxSize, ASide);

    if AFlag then
    begin
      // make region
      with AStartRect do
        case ASide of
          vgrbsLeft: CombineRectAndRegion(ABordersRegion, Rect(Left,
                                                Top,
                                                Left + APartSize,
                                                Bottom));
          vgrbsRight: CombineRectAndRegion(ABordersRegion, Rect(Right - APartSize,
                                                Top,
                                                Right,
                                                Bottom));
          vgrbsTop: CombineRectAndRegion(ABordersRegion, Rect(Left,
                                                Top,
                                                Right,
                                                Top + APartSize));
          vgrbsBottom: CombineRectAndRegion(ABordersRegion, Rect(Left,
                                                Bottom - APartSize,
                                                Right,
                                                Bottom));
        end;

      for X := X1 to X2 do
        for Y := Y1 to Y2 do
          CombineRectAndRegion(ABordersRegion, GetBorderInfo(X, Y, AOrientation).Rect, RGN_DIFF);
    end;

    with AInternalRect do
      case ASide of
        vgrbsLeft: Left := Left + APartSize;
        vgrbsTop: Top := Top + APartSize;
        vgrbsRight: Right := Right - APartSize;
        vgrbsBottom: Bottom := Bottom - APartSize;
      end;
  end;

begin
  AStartRect := AInternalRect;
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

procedure TvgrWBSheetPaintInfo.ResetBorder(X, Y: Integer; AOrientation: TvgrBorderOrientation);
var
  ABorderInfo: pvgrBorderCacheInfo;
begin
  ABorderInfo := GetBorderInfo(X, Y, AOrientation);
  if ABorderInfo <> nil then
    ABorderInfo.Size := 0;
end;

procedure TvgrWBSheetPaintInfo.CalcRangeCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ARange: pvgrRange;
  AStyle: pvgrRangeStyle;
  AText: string;
  AAlign: TvgrRangeHorzAlign;
  ARangeCacheInfo: pvgrRangeCacheInfo;
  ATextWidth: Integer;
  I, AStart, AEnd, ALeft, ARight: Integer;
  AFlag: Boolean;

  function IsMayBeLong: Boolean;
  begin
    Result := (ARange.Place.Left = ARange.Place.Right) and
              (ARange.Place.Top = ARange.Place.Bottom) and
              ((AStyle.Flags and vgrmask_RangeFlagsWordWrap) = 0) and
              (AStyle.Angle = 0) and
              (AText <> '');
  end;

  procedure CalcTextWidth;
  begin
    AssignFont(AStyle.Font, FWBSheet.Canvas.Font);
    ATextWidth := FWBSheet.Canvas.TextWidth(AText);
  end;

  procedure FillRangeInfo(ACacheInfo: pvgrRangeCacheInfo; ALong: Boolean; AVisible: Boolean; AInRect: Boolean);
  var
    X, Y: Integer;
    ACellInfo: pvgrCellCacheInfo;
  begin
    with ACacheInfo^ do
    begin
      DisplayText := AText;
      UniqueString(DisplayText);

      Range := ARange;
      Style := AStyle;
      TextWidth := ATextWidth;
      RealAlign := AAlign;
      Long := ALong;
      Visible := AVisible;
    end;

    if AInRect then
      for X := ARange.Place.Left to ARange.Place.Right do
        for Y := ARange.Place.Top to ARange.Place.Bottom do
        begin
          ACellInfo := GetCellInfo(X, Y);
          if ACellInfo <> nil then
            ACellInfo.Range := ACacheInfo;
        end;
  end;

begin
  with AItem as IvgrRange do
  begin
    ARange := ItemData;

    // check range (they can be invisible)
    AFlag := False;
    for I := ARange.Place.Left to ARange.Place.Right do
      AFlag := AFlag or (ColPixelWidth(I) <> 0);
    if not AFlag then exit;
    AFlag := False;
    for I := ARange.Place.Top to ARange.Place.Bottom do
      AFlag := AFlag or (RowPixelHeight(I) <> 0);
    if not AFlag then exit;

    AStyle := StyleData;
    AText := DisplayText;
    AAlign := GetAutoAlignForRangeValue(HorzAlign, ValueType);
  end;

  //
  if RectOverRect(ARange.Place, FCellsRect) then
  begin
    ARangeCacheInfo := @FRanges[FRangesCount];
    Inc(FRangesCount);
    if IsMayBeLong then
    begin
      CalcTextWidth;
      if ColPixelWidth(ARange.Place.Left) -
         CellLeftBorderSize(ARange.Place.Left, ARange.Place.Top) -
         CellRightBorderSize(ARange.Place.Left, ARange.Place.Bottom) < ATextWidth then
      begin
        FillRangeInfo(ARangeCacheInfo, True, True, True);
        exit;
      end;
    end;
    FillRangeInfo(ARangeCacheInfo, False, True, True);
    CalcNormalRange(ARangeCacheInfo);
  end
  else
  begin
    if IsMayBeLong then
    begin
      // this range my be overlapped with FCellsRect
      ARangeCacheInfo := nil;
      if ARange.Place.Left < FCellsRect.Left then
      begin
        with FLeftLongRanges[ARange.Place.Top - FCellsRect.Top] do
          if (Range = nil) or (Range.Place.Right < ARange.Place.Left) then
            ARangeCacheInfo := @FLeftLongRanges[ARange.Place.Top - FCellsRect.Top];
      end
      else
      begin
        with FRightLongRanges[ARange.Place.Top - FCellsRect.Top] do
          if (Range = nil) or (Range.Place.Left > ARange.Place.Right) then
            ARangeCacheInfo := @FRightLongRanges[ARange.Place.Top - FCellsRect.Top];
      end;
      if ARangeCacheInfo <> nil then
        with ARangeCacheInfo^ do
        begin
          Range := nil;
          CalcTextWidth;
          case AAlign of
            vgrhaCenter:
              ALeft := ColPixelLeft(ARange.Place.Left) - (ATextWidth - ColPixelWidth(ARange.Place.Left)) div 2;
            vgrhaRight:
              ALeft := ColPixelRight(ARange.Place.Left) - ATextWidth;
          else
            ALeft := ColPixelLeft(ARange.Place.Left);
          end;
          ARight := ALeft + ATextWidth;
          FillRangeInfo(ARangeCacheInfo,
                        True,
                        RectOverRect(FPaintRect, Rect(ALeft, FPaintRect.Top, ARight, FPaintRect.Bottom)),
                        False);
        end
    end
    else
    begin
      if AText <> '' then
      begin
        AStart := Max(0, ARange.Place.Top - FCellsRect.Top);
        AEnd := Min(Max(ARange.Place.Bottom - FCellsRect.Top, 0), FCellsRectHeight - 1);
        if ARange.Place.Left < FCellsRect.Left then
        begin
          // check FLeftLongRanges
          for I := AStart to AEnd do
            with FLeftLongRanges[I] do
              if (Range = nil) or (Range.Place.Right <= ARange.Place.Right) then
              begin
                FillRangeInfo(@FLeftLongRanges[I], False, False, False);
              end;
        end
        else
        begin
          // check FRightLongRanges
          for I := AStart to AEnd do
            with FRightLongRanges[I] do
              if (Range = nil) or (Range.Place.Left >= ARange.Place.Left) then
              begin
                FillRangeInfo(@FRightLongRanges[I], False, False, False);
              end;
        end;
      end;
    end;
  end;
end;

procedure TvgrWBSheetPaintInfo.CalcNormalRange(AInfo: pvgrRangeCacheInfo);
var
  X, Y: Integer;
begin
  with AInfo^ do
  begin
    GetRectScreenPixels(Range.Place, FullRect);
    InternalRect := FullRect;
    CalcInternalRect(Range.Place, InternalRect, FillBordersRegion);
    FillRect := InternalRect;
    // reset borders in this range
    for X := Range.Place.Left to Range.Place.Right do
      for Y := Range.Place.Top to Range.Place.Bottom do
      begin
        if X <> Range.Place.Left then
          ResetBorder(X, Y, vgrboLeft);
        if Y <> Range.Place.Top then
          ResetBorder(X, Y, vgrboTop);
      end;
  end;
end;

procedure TvgrWBSheetPaintInfo.CalcLongRange(AInfo: pvgrRangeCacheInfo);

  procedure ResetBorder(X, Y: Integer);
  var
    AInfo: pvgrCellCacheInfo;
  begin
    Self.ResetBorder(X, Y, vgrboLeft);
    AInfo := GetCellInfo(X, Y);
    if (AInfo <> nil) and (AInfo.Range <> nil) and not AInfo.Range.Long then
      CalcNormalRange(AInfo.Range);
    AInfo := GetCellInfo(X - 1, Y);
    if (AInfo <> nil) and (AInfo.Range <> nil) and not AInfo.Range.Long then
      CalcNormalRange(AInfo.Range);
  end;
  
  function GetLeftLimit(AXCoord: Integer; ATextPart: Integer; AFullLimit: Integer; var AFillLimit: Integer): Integer;
  begin
    repeat
      if IsCellPainted(AXCoord - 1, AInfo.Range.Place.Top) or
         (AXCoord = FCellsRect.Left) then
        break;

      AFillLimit := AFullLimit;

      ResetBorder(AXCoord, AInfo.Range.Place.Top);
      Dec(AXCoord);
      ATextPart := ATextPart - (ColPixelWidth(AXCoord) - CellLeftBorderSize(AXCoord, AInfo.Range.Place.Top));
    until ATextPart < 0;
    Result := ColPixelLeft(AXCoord) - CellLeftBorderSize(AXCoord, AInfo.Range.Place.Top);
// comment to remove BUG_0062  
//    if AXFillCoord = -1 then
//      AFillLimit := Result
//    else
//      AFillLimit := ColPixelLeft(AXFillCoord);
  end;

  function GetRightLimit(AXCoord: Integer; ATextPart: Integer; AFullLimit: Integer; var AFillLimit: Integer): Integer;
  begin
    repeat
      if IsCellPainted(AXCoord + 1, AInfo.Range.Place.Top) or
         (AXCoord = FCellsRect.Right) then
        break;

      AFillLimit := AFullLimit;

      ResetBorder(AXCoord + 1, AInfo.Range.Place.Top);
      Inc(AXCoord);
      ATextPart := ATextPart - (ColPixelWidth(AXCoord) - CellRightBorderSize(AXCoord, AInfo.Range.Place.Top));
    until ATextPart < 0;
    Result := ColPixelRight(AXCoord) - CellRightBorderSize(AXCoord, AInfo.Range.Place.Top);
// comment to remove BUG_0062
//    if AXFillCoord = -1 then
//      AFillLimit := Result
//    else
//      AFillLimit := ColPixelRight(AXFillCoord);
  end;
  
begin
  with AInfo^ do
  begin
    GetRectScreenPixels(Range.Place, FullRect);
    InternalRect := FullRect;
    CalcInternalRect(Range.Place, InternalRect, FillBordersRegion);
    FillRect := InternalRect;
    case RealAlign of
      vgrhaLeft:
        begin
          InternalRect.Right := GetRightLimit(Range.Place.Right,
                                              TextWidth - (FullRect.Right - FullRect.Left - CellLeftBorderSize(Range.Place.Left, Range.Place.Top)),
                                              FullRect.Right,
                                              FillRect.Right);
//          FillRect.Right := InternalRect.Right;
//          FillRect.Right := FullRect.Right;
        end;
      vgrhaCenter:
        begin
          InternalRect.Left := GetLeftLimit(Range.Place.Left,
                                            (TextWidth - (FullRect.Right - FullRect.Left)) div 2 - CellLeftBorderSize(Range.Place.Left, Range.Place.Top),
                                            FullRect.Left,
                                            FillRect.Left);
          InternalRect.Right := GetRightLimit(Range.Place.Right,
                                              (TextWidth - (FullRect.Right - FullRect.Left)) div 2 - CellRightBorderSize(Range.Place.Left, Range.Place.Top),
                                              FullRect.Right,
                                              FillRect.Right);
//          FillRect.Left := InternalRect.Left;
//          FillRect.Right := InternalRect.Right;
//          FillRect.Left := FullRect.Left;
//          FillRect.Right := FullRect.Right;
        end;
      vgrhaRight:
        begin
          InternalRect.Left := GetLeftLimit(Range.Place.Left,
                                            TextWidth - (ColPixelWidth(Range.Place.Left) - CellRightBorderSize(Range.Place.Left, Range.Place.Top)),
                                            FullRect.Left,
                                            FillRect.Left);
//          FillRect.Left := InternalRect.Left;
//          FillRect.Left := FullRect.Left;
        end;
    end;
  end;
end;

procedure TvgrWBSheetPaintInfo.CalcBorderPoint(X, Y: Integer);
var
  ALeft, ATop, ARight, ABottom: pvgrBorderCacheInfo;

  function Exists(ABorderInfo: pvgrBorderCacheInfo): Boolean;
  begin
    Result := (ABorderInfo <> nil) and (ABorderInfo.Border <> nil) and (ABorderInfo.Size <> 0);
  end;

  function ExistsGrid(ABorderInfo: pvgrBorderCacheInfo): Boolean;
  begin
    Result := (ABorderInfo <> nil) and (ABorderInfo.Border = nil) and (ABorderInfo.Size = 1);
  end;

  function Size(ABorderInfo: pvgrBorderCacheInfo): Integer;
  begin
    if ABorderInfo <> nil then
      Result := ABorderInfo.Size
    else
      Result := 0;
  end;

  procedure CorrectLeftBorder;
  begin
    ALeft.Rect.Right := ALeft.Rect.Right + GetBorderPartSize(Max(Size(ATop), Size(ABottom)), vgrbsLeft{vgrbsRight});
    if ARight <> nil then
      ARight.Rect.Left := ALeft.Rect.Right;
    if ATop <> nil then
      ATop.Rect.Bottom := ALeft.Rect.Top;
    if ABottom <> nil then
      ABottom.Rect.Top := ALeft.Rect.Bottom;
  end;

  procedure CorrectTopBorder;
  begin
    ATop.Rect.Bottom := ATop.Rect.Bottom + GetBorderPartSize(Max(Size(ALeft), Size(ARight)), vgrbsTop{vgrbsBottom});
    if ALeft <> nil then
      ALeft.Rect.Right := ATop.Rect.Left;
    if ARight <> nil then
      ARight.Rect.Left := ATop.Rect.Right;
    if ABottom <> nil then
      ABottom.Rect.Top := ATop.Rect.Bottom;
  end;

  procedure CorrectRightBorder;
  begin
    ARight.Rect.Left := ARight.Rect.Left - GetBorderPartSize(Max(Size(ATop), Size(ABottom)), vgrbsRight{vgrbsLeft});
    if ALeft <> nil then
      ALeft.Rect.Right := ARight.Rect.Left;
    if ATop <> nil then
      ATop.Rect.Bottom := ARight.Rect.Top;
    if ABottom <> nil then
      ABottom.Rect.Top := ARight.Rect.Bottom;
  end;

  procedure CorrectBottomBorder;
  begin
    ABottom.Rect.Top := ABottom.Rect.Top - GetBorderPartSize(Max(Size(ALeft), Size(ARight)), vgrbsBottom{vgrbsTop});
    if ATop <> nil then
      ATop.Rect.Bottom := ABottom.Rect.Top;
    if ALeft <> nil then
      ALeft.Rect.Right := ABottom.Rect.Left;
    if ARight <> nil then
      ARight.Rect.Left := ABottom.Rect.Right;
  end;

begin
  ALeft := GetBorderInfo(X - 1, Y, vgrboTop);
  ATop := GetBorderInfo(X, Y - 1, vgrboLeft);
  ARight := GetBorderInfo(X, Y, vgrboTop);
  ABottom := GetBorderInfo(X, Y, vgrboLeft);

  if Exists(ALeft) and Exists(ARight) then
  begin
    if ALeft.Size > ARight.Size then
      CorrectLeftBorder
    else
      CorrectRightBorder;
  end
  else
    if Exists(ATop) and Exists(ABottom) then
    begin
      if ATop.Size > ABottom.Size then
        CorrectTopBorder
      else
        CorrectBottomBorder;
    end
    else
      if Exists(ALeft) then
        CorrectLeftBorder
      else
        if Exists(ARight) then
          CorrectRightBorder
        else
          if Exists(ATop) then
            CorrectTopBorder
          else
            if Exists(ABottom) then
              CorrectBottomBorder
            else
              if ExistsGrid(ARight) then
                CorrectRightBorder
              else
                if ExistsGrid(ALeft) then
                  CorrectLeftBorder
                else
                  if ExistsGrid(ATop) then
                    CorrectTopBorder
                  else
                    if ExistsGrid(ABottom) then
                      CorrectBottomBorder;
end;

procedure TvgrWBSheetPaintInfo.Update(const ACellsRect, APaintRect: TRect);
var
  I, X, Y: Integer;

  function CheckLeftBorder(ALeft: Integer): Boolean;
  begin
    Result := (ALeft <= FBounds.Left) or (ColPixelWidth(ALeft - 1) > 0);
  end;

  function CheckTopBorder(ATop: Integer): Boolean;
  begin
    Result := (ATop <= FBounds.Top) or (RowPixelHeight(ATop - 1) > 0);
  end;

begin
  Clear;

  // 1. expand ACellsRect and include all ranges overlapped with this rect
  //   (calculate bounds)
  FPaintRect := APaintRect;
  FCellsRect := ACellsRect;
  FBounds := ACellsRect;
  FWBSheet.Worksheet.RangesList.FindAndCallback(ACellsRect, ExpandCallback, nil);
  FBoundsWidth := FBounds.Right - FBounds.Left + 1;
  FBoundsHeight := FBounds.Bottom - FBounds.Top + 1;
  FCellsRectWidth := ACellsRect.Right - ACellsRect.Left + 1;
  FCellsRectHeight := ACellsRect.Bottom - ACellsRect.Top + 1;

  // 2. init arrays
  SetLength(FVertPoints, FBoundsHeight + 2);

  SetLength(FLeftLongRanges, FCellsRectHeight);
  FillChar(FLeftLongRanges[0], FCellsRectHeight * SizeOf(rvgrRangeCacheInfo), 0);

  SetLength(FRightLongRanges, FCellsRectHeight);
  FillChar(FRightLongRanges[0], FCellsRectHeight * SizeOf(rvgrRangeCacheInfo), 0);

  SetLength(FRanges, FCellsRectHeight * FCellsRectWidth);

  SetLength(FCells, FCellsRectHeight * FCellsRectWidth);
  SetLength(FBorders, (FBoundsHeight + 1) * (FBoundsWidth + 1) * 2);

  // 3. calc FVertPoints
  with FVertPoints[0] do
  begin
    Pos := FWBSheet.RowPixelTop(Bounds.Top) - FWBSheet.WorkbookGrid.LeftTopPixelsOffset.Y;
    Size := FWBSheet.RowPixelHeight(Bounds.Top);
  end;
  for I := 1 to FBoundsHeight + 1 do
    with FVertPoints[I] do
    begin
      Pos := FVertPoints[I - 1].Pos + FVertPoints[I - 1].Size;
      Size := FWBSheet.RowPixelHeight(Bounds.Top + I);
    end;

  // 3. Borders
  FWBSheet.Worksheet.BordersList.FindAndCallback(Rect(FBounds.Left, FBounds.Top, FBounds.Right + 1, FBounds.Bottom + 1), UpdateBordersCallback, nil);
  I := 0;
  for Y := FBounds.Top to FBounds.Bottom + 1 do
    for X := FBounds.Left to FBounds.Right + 1 do
    begin
      if CheckLeftBorder(X) then
      begin
        if FBorders[I].Size = 0 then
        begin
          FBorders[I].Size := 1;
          CalcBorderRect(@FBorders[I], X, Y, vgrboLeft);
        end;
      end
      else
        ResetBorder(X, Y, vgrboLeft);
      Inc(I);
      if CheckTopBorder(Y) then
      begin
        if FBorders[I].Size = 0 then
        begin
          FBorders[I].Size := 1;
          CalcBorderRect(@FBorders[I], X, Y, vgrboTop);
        end;
      end
      else
        ResetBorder(X, Y, vgrboTop);
      Inc(I);
    end;

  // 4. Calc ranges
  FWBSheet.Worksheet.RangesList.FindAndCallBack(Rect(0, ACellsRect.Top, MaxInt, ACellsRect.Bottom), CalcRangeCallback, nil);
  for I := 0 to Length(FLeftLongRanges) - 1 do
    with FLeftLongRanges[I] do
      if not Visible then
        Range := nil;
  for I := 0 to Length(FRightLongRanges) - 1 do
    with FRightLongRanges[I] do
      if not Visible then
        Range := nil;

  for I := 0 to Length(FLeftLongRanges) - 1 do
    if FLeftLongRanges[I].Range <> nil then
      CalcLongRange(@FLeftLongRanges[I]);
  for I := 0 to Length(FRightLongRanges) - 1 do
    if FRightLongRanges[I].Range <> nil then
      CalcLongRange(@FRightLongRanges[I]);
  for I := 0 to FRangesCount - 1 do
    if FRanges[I].Long then
      CalcLongRange(@FRanges[I]);

  // 5. set cells Highlighted flag
  with FWBSheet.FSelectedRegion do
  begin
    for I := 0 to RectsCount - 1 do
      with Rects[I]^ do
        for X := Max(0, Left - FCellsRect.Left) to Min(FCellsRect.Right, Right) - FCellsRect.Left do
          for Y := Max(0, Top - FCellsRect.Top) to Min(FCellsRect.Bottom, Bottom) - FCellsRect.Top do
            FCells[FCellsRectWidth * Y + X].Highlighted := True;
  end;
  for X := Max(0, FWBSheet.FSelStartRect.Left - FCellsRect.Left) to Min(FCellsRect.Right, FWBSheet.FSelStartRect.Right) - FCellsRect.Left do
    for Y := Max(0, FWBSheet.FSelStartRect.Top - FCellsRect.Top) to Min(FCellsRect.Bottom, FWBSheet.FSelStartRect.Bottom) - FCellsRect.Top do
      FCells[FCellsRectWidth * Y + X].Highlighted := False;

  // 6. Calc borders intersect points
  for Y := FCellsRect.Top to FCellsRect.Bottom do
    for X := FCellsRect.Left to FCellsRect.Right do
      CalcBorderPoint(X, Y);
end;

function TvgrWBSheetPaintInfo.GetBorderInfo(X, Y: Integer; AOrientation: TvgrBorderOrientation): pvgrBorderCacheInfo;
var
  AIndex: Integer;
begin
  AIndex := GetBorderInfoIndex(X, Y, AOrientation);
  if AIndex <> -1 then
    Result := @FBorders[AIndex]
  else
    Result := nil;
end;

function TvgrWBSheetPaintInfo.GetBorderSize(X, Y: Integer; AOrientation: TvgrBorderOrientation): Integer;
var
  AIndex: Integer;
begin
  AIndex := GetBorderInfoIndex(X, Y, AOrientation);
  if AIndex <> -1 then
    Result := FBorders[AIndex].Size
  else
    Result := 0;
end;

function TvgrWBSheetPaintInfo.GetCellInfo(X, Y: Integer): pvgrCellCacheInfo;
begin
  if (X >= FCellsRect.Left) and (X <= FCellsRect.Right) and
     (Y >= FCellsRect.Top) and (Y <= FCellsRect.Bottom) then
    Result := @FCells[(FCellsRect.Right - FCellsRect.Left + 1) * (Y - FCellsRect.Top) + (X - FCellsRect.Left)]
  else
    Result := nil;
end;

function TvgrWBSheetPaintInfo.ColPixelRight(ACol: Integer): Integer;
begin
  Result := FWBSheet.ColPixelRight(ACol) - FWBSheet.WorkbookGrid.FLeftTopPixelsOffset.X;
end;

function TvgrWBSheetPaintInfo.ColPixelLeft(ACol: Integer): Integer;
begin
  Result := FWBSheet.ColPixelLeft(ACol) - FWBSheet.WorkbookGrid.FLeftTopPixelsOffset.X;
end;

function TvgrWBSheetPaintInfo.ColPixelWidth(ACol: Integer): Integer;
begin
  Result := FWBSheet.ColPixelWidth(ACol);
end;

function TvgrWBSheetPaintInfo.RowPixelTop(ARow: Integer): Integer;
begin
  if (ARow >= Bounds.Top) and (ARow <= Bounds.Bottom + 1) then
    Result := FVertPoints[ARow - Bounds.Top].Pos
  else
    Result := FWBSheet.RowPixelTop(ARow) - FWBSheet.WorkbookGrid.FLeftTopPixelsOffset.Y;
end;

function TvgrWBSheetPaintInfo.RowPixelBottom(ARow: Integer): Integer;
begin
  if (ARow >= Bounds.Top) and (ARow <= Bounds.Bottom + 1) then
    Result := FVertPoints[ARow - Bounds.Top + 1].Pos
  else
    Result := FWBSheet.RowPixelBottom(ARow) - FWBSheet.WorkbookGrid.FLeftTopPixelsOffset.Y;
end;

function TvgrWBSheetPaintInfo.RowPixelHeight(ARow: Integer): Integer;
begin
  if (ARow >= Bounds.Top) and (ARow <= Bounds.Bottom) then
    Result := FVertPoints[ARow - Bounds.Top].Size
  else
    Result := FWBSheet.RowPixelHeight(ARow);
end;

function TvgrWBSheetPaintInfo.CellLeftBorderSize(X, Y: Integer): Integer;
begin
  Result := GetBorderPartSize(GetBorderSize(X, Y, vgrboLeft), vgrbsLeft);
end;

function TvgrWBSheetPaintInfo.CellRightBorderSize(X, Y: Integer): Integer;
begin
  Result := GetBorderPartSize(GetBorderSize(X, Y, vgrboLeft), vgrbsRight);
end;

procedure TvgrWBSheetPaintInfo.GetRectScreenPixels(const rRegions: TRect; var rPixels: TRect);
begin
  rPixels.Left := ColPixelLeft(rRegions.Left);
  rPixels.Right := ColPixelRight(rRegions.Right);
  rPixels.Top := RowPixelTop(rRegions.Top);
  rPixels.Bottom := RowPixelBottom(rRegions.Bottom);
end;

function TvgrWBSheetPaintInfo.GetCellScreenRect(X, Y: Integer): TRect;
begin
  Result.Left := ColPixelLeft(X) + GetBorderPartSize(GetBorderSize(X, Y, vgrboLeft), vgrbsLeft);
  Result.Right := ColPixelRight(X) - GetBorderPartSize(GetBorderSize(X + 1, Y, vgrboLeft), vgrbsRight);
  Result.Top := RowPixelTop(Y) + GetBorderPartSize(GetBorderSize(X, Y, vgrboTop), vgrbsTop);
  Result.Bottom := RowPixelBottom(Y) - GetBorderPartSize(GetBorderSize(X, Y + 1, vgrboTop), vgrbsBottom);
end;

function TvgrWBSheetPaintInfo.IsCellPainted(X, Y: Integer): Boolean;
var
  ACellInfo: pvgrCellCacheInfo;
begin
  ACellInfo := GetCellInfo(X, Y);
  Result := (ACellInfo <> nil) and (ACellInfo.Range <> nil) and (ACellInfo.Range.DisplayText <> '');
end;

function TvgrWBSheetPaintInfo.IsCellBackgroundPainted(X, Y: Integer): Boolean;
var
  ACellInfo: pvgrCellCacheInfo;
begin
  ACellInfo := GetCellInfo(X, Y);
  Result := (ACellInfo <> nil) and (ACellInfo.Range <> nil);
end;

/////////////////////////////////////////////////
//
// TvgrSheetPointInfo
//
/////////////////////////////////////////////////
constructor TvgrSheetPointInfo.Create(APlace: TvgrWBSheetPlace; const ACell: TPoint; ARangeIndex: Integer; const ADownRect: TRect);
begin
  inherited Create;
  FPlace := APlace;
  FCell := ACell;
  FRangeIndex := ARangeIndex;
  FDownRect := ADownRect;
end;

/////////////////////////////////////////////////
//
// TvgrWBSheet
//
/////////////////////////////////////////////////
constructor TvgrWBSheet.Create(AOwner : TComponent);
begin
  inherited;
  TabStop := true;

  FPaintInfo := TvgrWBSheetPaintInfo.Create(Self);
  FSelectedRegion := TvgrWBSelectedRegion.Create;
  FSelectedRegion.SetSelection(Rect(0, 0, 0, 0));
  FCurRect := Rect(0, 0, 0, 0);
  FSelStartRect := Rect(0, 0, 0, 0);

  FScrollTimer := TTimer.Create(nil);
  FScrollTimer.Enabled := False;
  FScrollTimer.Interval := cDefaultScrollTimerInterval;
  FScrollTimer.OnTimer := OnScrollTimer;
end;

destructor TvgrWBSheet.Destroy;
begin
  FreeAndNil(FScrollTimer);
  FSelectedRegion.Free;
  FPaintInfo.Free;
  inherited;
end;

function TvgrWBSheet.GetFormulaPanel: TvgrWBFormulaPanel;
begin
  Result := WorkbookGrid.FFormulaPanel;
end;

procedure TvgrWBSheet.OnScrollTimer(Sender: TObject);
var
  ACurRect: PRect;
begin
  ACurRect := FSelectedRegion.Rects[0];

  if FScrollX = -1 then
  begin
    if ACurRect.Left = 0 then
      FScrollX := 0;
  end;

  if FScrollY = -1 then
  begin
    if ACurRect.Top = 0 then
      FScrollY := 0;
  end;

  if (FScrollX <> 0) or (FScrollY <> 0) then
    SetUserSelection(0, FScrollX, FScrollY, 0, 0)
  else
    FScrollTimer.Enabled := False;
end;

function TvgrWBSheet.GetWorkbookGrid : TvgrWorkbookGrid;
begin
  Result := TvgrWorkbookGrid(Owner);
end;

function TvgrWBSheet.GetWorksheet : TvgrWorksheet;
begin
  Result := WorkbookGrid.ActiveWorksheet;
end;

function TvgrWBSheet.GetHorzScrollBar : TvgrScrollBar;
begin
  Result := WorkbookGrid.HorzScrollBar;
end;

function TvgrWBSheet.GetVertScrollBar : TvgrScrollBar;
begin
  Result := WorkbookGrid.VertScrollBar;
end;

function TvgrWBSheet.GetColsPanel: TvgrWBColsPanel;
begin
  Result := WorkbookGrid.FColsPanel;
end;

function TvgrWBSheet.GetRowsPanel: TvgrWBRowsPanel;
begin
  Result := WorkbookGrid.FRowsPanel;
end;

function TvgrWBSheet.GetDimensionsRight: Integer;
begin
  if Worksheet = nil then
    Result := 0
  else
    Result := Worksheet.Dimensions.Right;
end;

function TvgrWBSheet.GetDimensionsBottom: Integer;
begin
  if Worksheet = nil then
    Result := 0
  else
    Result := Worksheet.Dimensions.Bottom;
end;

function TvgrWBSheet.GetLeftCol: Integer;
begin
  Result := HorzScrollBar.Position;
end;

function TvgrWBSheet.GetTopRow: Integer;
begin
  Result := VertScrollBar.Position;
end;

function TvgrWBSheet.GetVisibleColsCount: Integer;
begin
  Result := WorkbookGrid.VisibleColsCount;
end;

function TvgrWBSheet.GetVisibleRowsCount: Integer;
begin
  Result := WorkbookGrid.VisibleRowsCount;
end;

function TvgrWBSheet.GetFullVisibleColsCount: Integer;
begin
  Result := WorkbookGrid.FullVisibleColsCount;
end;

function TvgrWBSheet.GetFullVisibleRowsCount: Integer;
begin
  Result := WorkbookGrid.FullVisibleRowsCount;
end;

function TvgrWBSheet.GetKeyboardFilter: TvgrKeyboardFilter;
begin
  Result := WorkbookGrid.KeyboardFilter;
end;

function TvgrWBSheet.GetInplaceEdit: TvgrWBInplaceEdit;
begin
  Result := WorkbookGrid.InplaceEdit;
end;

function TvgrWBSheet.GetIsInplaceEdit: Boolean;
begin
  Result := InplaceEdit.Visible;
end;

function TvgrWBSheet.ColPixelWidthNoCache(ColNumber : integer) : integer;
var
  Col: IvgrCol;
begin
  if Worksheet <> nil then
  begin
    Col := Worksheet.ColsList.Find(ColNumber);
    if Col <> nil then
    begin
      if Col.Visible then
        Result := ConvertTwipsToPixelsX(Col.Width)
      else
        Result := 0;
      exit;
    end;
  end;
  Result := ConvertTwipsToPixelsX(WorkbookGrid.DefaultColWidth)
end;

function TvgrWBSheet.ColPixelWidth(ColNumber : integer) : integer;
begin
  if (ColNumber >= 0) and (ColNumber < Length(FHorzPoints)) then
    Result := FHorzPoints[ColNumber].Size
  else
    Result := ColPixelWidthNoCache(ColNumber);
end;

function TvgrWBSheet.RowPixelHeight(RowNumber : integer) : integer;
var
  Row: IvgrRow;
begin
  if Worksheet <> nil then
  begin
    Row := Worksheet.RowsList.Find(RowNumber);
    if Row <> nil then
    begin
      if Row.Visible then
        Result := ConvertTwipsToPixelsY(Row.Height)
      else
        Result := 0;
      exit;
    end;
  end;
  Result := ConvertTwipsToPixelsY(WorkbookGrid.DefaultRowHeight)
end;

function TvgrWBSheet.ColPixelLeft(ColNumber: integer) : integer;
var
  I: integer;
begin
  if Length(FHorzPoints) > 0 then
    Result := FHorzPoints[Min(ColNumber, High(FHorzPoints))].Pos
  else
    Result := 0;
  for I := Min(ColNumber, High(FHorzPoints)) to ColNumber - 1 do
    Result := Result + ColPixelWidth(I);
end;

function TvgrWBSheet.ColPixelRight(ColNumber: Integer): Integer;
var
  I: integer;
begin
  if Length(FHorzPoints) > 0 then
    Result := FHorzPoints[Min(ColNumber + 1, High(FHorzPoints))].Pos
  else
    Result := 0;
  for I := Min(ColNumber + 1, High(FHorzPoints)) to ColNumber do
    Result := Result + ColPixelWidth(I);
end;

function TvgrWBSheet.ColScreenPixelLeft(ColNumber: Integer): Integer;
begin
  Result := ColPixelLeft(ColNumber) - WorkbookGrid.FLeftTopPixelsOffset.X;
end;

function TvgrWBSheet.ColScreenPixelRight(ColNumber: Integer): Integer;
begin
  Result := ColPixelRight(ColNumber) - WorkbookGrid.FLeftTopPixelsOffset.X;
end;

function TvgrWBSheet.RowPixelTop(RowNumber: integer) : integer;
var
  i : integer;
begin
  Result := 0;
  for i := 0 to RowNumber-1 do
    Result := Result + RowPixelHeight(i);
end;

function TvgrWBSheet.RowPixelBottom(RowNumber: Integer): Integer;
var
  i : integer;
begin
  Result := 0;
  for i := 0 to RowNumber do
    Result := Result + RowPixelHeight(i);
end;

function TvgrWBSheet.RowScreenPixelTop(RowNumber: Integer): Integer;
var
  I: Integer;
begin
  Result := 0;
  if RowNumber < WorkbookGrid.TopRow then
    for I := WorkbookGrid.TopRow - 1 downto RowNumber do
      Result := Result - RowPixelHeight(I)
  else
    if RowNumber > WorkbookGrid.TopRow then
      for I := WorkbookGrid.TopRow to RowNumber - 1 do
        Result := Result + RowPixelHeight(I);
end;

function TvgrWBSheet.RowScreenPixelBottom(RowNumber: Integer): Integer;
begin
  Result := RowScreenPixelTop(RowNumber + 1);
end;

function TvgrWBSheet.ColVisible(ColNumber: Integer): Boolean;
var
  ACol: IvgrCol;
begin
  ACol := Worksheet.ColsList.Find(ColNumber);
  Result := (ACol = nil) or ((ConvertTwipsToPixelsX(ACol.Width) > 0) and ACol.Visible);
end;

function TvgrWBSheet.RowVisible(RowNumber: Integer): Boolean;
var
  ARow: IvgrRow;
begin
  ARow := Worksheet.RowsList.Find(RowNumber);
  Result := (ARow = nil) or ((ConvertTwipsToPixelsY(ARow.Height) > 0) and ARow.Visible);
end;

procedure TvgrWBSheet.GetRectPixels(const rRegions: TRect; var rPixels: TRect);
var
  I: Integer;
begin
  rPixels.Left := ColPixelLeft(rRegions.Left);
  rPixels.Top := RowPixelTop(rRegions.Top);
  rPixels.Right := rPixels.Left;
  for I := rRegions.Left to rRegions.Right do
    rPixels.Right := rPixels.Right + ColPixelWidth(I);
  rPixels.Bottom := rPixels.Top;
  for I := rRegions.Top to rRegions.Bottom do
    rPixels.Bottom := rPixels.Bottom + RowPixelHeight(I);
end;

procedure TvgrWBSheet.GetRectScreenPixels(const rRegions: TRect; var rPixels: TRect);
begin
  GetRectPixels(rRegions, rPixels);
  OffsetRect(rPixels, -WorkbookGrid.LeftTopPixelsOffset.X, -WorkbookGrid.LeftTopPixelsOffset.Y);
end;

procedure TvgrWBSheet.GetRectInternalScreenPixels(const rRegions: TRect; var rPixels: TRect);

  function GetBorderSize(ALeft, ATop: Integer; AOrientation: TvgrBorderOrientation): Integer;
  var
    ABorder: IvgrBorder;
  begin
    ABorder := Worksheet.BordersList.Find(ALeft, ATop, AOrientation);
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

    with rPixels do
      case ASide of
        vgrbsLeft: Left := Left + APartSize;
        vgrbsTop: Top := Top + APartSize;
        vgrbsRight: Right := Right - APartSize;
        vgrbsBottom: Bottom := Bottom - APartSize;
      end;
  end;

begin
  GetRectScreenPixels(rRegions, rPixels);
  with rRegions do
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

function TvgrWBSheet.GetInvalidateRectForRange(ARange: IvgrRange): TRect;
begin
  GetRectScreenPixels(ARange.Place, Result);
  Result.Left := 0;
  Result.Right := ClientWidth;
{
  if (not ARange.WordWrap) and
     (ARange.Angle = 0) and
     (ARange.Left = ARange.Right) and
     (ARange.Top = ARange.Bottom) then
  begin
    // this range possibly LONG, add all line
    Result.Left := 0;
    Result.Right := ClientWidth;
  end;
}
end;

procedure TvgrWBSheet.WMGetDlgCode(var Msg : TWMGetDlgCode);
begin
  inherited;
  Msg.Result := Msg.Result or DLGC_WANTALLKEYS or DLGC_WANTARROWS;
end;

procedure TvgrWBSheet.WMEraseBkgnd(var Msg : TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWBSheet.WMMouseWheel(var Msg : TWMMouseWheel);
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

function TvgrWBSheet.IsRangeSelected(const ARangeRect: TRect): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to FSelectedRegion.RectsCount - 1 do
    if RectInSelRect(ARangeRect, FSelectedRegion.Rects[I]^) then
    begin
      Result := True;
      exit;
    end;
end;

function TvgrWBSheet.CheckAddCol(AColNumber: Integer): IvgrCol;
begin
  Result := Worksheet.ColsList.Find(AColNumber);
  if Result = nil then
  begin
    Result := Worksheet.Cols[AColNumber];
    Result.Width := cDefaultColWidthTwips;
  end;
end;

function TvgrWBSheet.CheckAddRow(ARowNumber: Integer): IvgrRow;
begin
  Result := Worksheet.RowsList.Find(ARowNumber);
  if Result = nil then
  begin
    Result := Worksheet.Rows[ARowNumber];
    Result.Height := cDefaultRowHeightTwips;
  end;
end;

procedure TvgrWBSheet.GetBorderPixelRect(ABorder: IvgrBorder; var ABorderRect: TRect; var ABorderPixelSize: Integer; var ABorderPos: Integer);
begin
  if ABorder.Orientation = vgrboLeft then
  begin
    ABorderPixelSize := ConvertTwipsToPixelsX(ABorder.Width);
    ABorderPos := ColPixelLeft(ABorder.Left) - 1 - WorkbookGrid.LeftTopPixelsOffset.X;
    ABorderRect.Left := ABorderPos - (ABorderPixelSize div 2);
    ABorderRect.Right := ABorderRect.Left + ABorderPixelSize;
    ABorderRect.Top := RowPixelTop(ABorder.Top) - WorkbookGrid.LeftTopPixelsOffset.Y;
    ABorderRect.Bottom := ABorderRect.Top + RowPixelHeight(ABorder.Top);
  end
  else
  begin
    ABorderPixelSize := ConvertTwipsToPixelsY(ABorder.Width);
    ABorderRect.Left := ColPixelLeft(ABorder.Left) - WorkbookGrid.LeftTopPixelsOffset.X;
    ABorderRect.Right := ABorderRect.Left + ColPixelWidth(ABorder.Left);
    ABorderPos := RowPixelTop(ABorder.Top) - 1 - WorkbookGrid.LeftTopPixelsOffset.Y;
    ABorderRect.Top := ABorderPos - (ABorderPixelSize div 2);
    ABorderRect.Bottom := ABorderRect.Top + ABorderPixelSize;
  end;
end;

procedure TvgrWBSheet.PaintRangeBackground(AInfo: pvgrRangeCacheInfo);

  function IsHighlighted: Boolean;
  begin
    Result := IsRangeSelected(AInfo.Range.Place) and not RectOverRect(AInfo.Range.Place, FSelStartRect);
  end;

begin
  if AInfo.Style.FillBackColor = clNone then
  begin
    // paint by default color
    if IsHighlighted then
      Canvas.Brush.Color := GetHighlightColor(WorkbookGrid.BackgroundColor)
    else
      Canvas.Brush.Color := WorkbookGrid.BackgroundColor;
    Canvas.Brush.Style := bsSolid;
  end
  else
  begin
    Canvas.Brush.Style := AInfo.Style.FillPattern;
    if Canvas.Brush.Style = bsSolid then
    begin
      if IsHighlighted then
        Canvas.Brush.Color := GetHighlightColor(AInfo.Style.FillBackColor)
      else
        Canvas.Brush.Color := AInfo.Style.FillBackColor;
    end
    else
    begin
      if IsHighlighted then
      begin
        Canvas.Brush.Color := GetHighlightColor(AInfo.Style.FillForeColor);
        SetBkColor(Canvas.Handle, GetHighlightColor(AInfo.Style.FillBackColor));
      end
      else
      begin
        Canvas.Brush.Color := AInfo.Style.FillForeColor;
        SetBkColor(Canvas.Handle, GetRGBColor(AInfo.Style.FillBackColor));
      end;
    end;
  end;

  Canvas.FillRect(AInfo.FillRect);
  if AInfo.FillBordersRegion <> 0 then
    FillRegion(Canvas, AInfo.FillBordersRegion);
end;

procedure TvgrWBSheet.PaintNormalRange(AInfo: pvgrRangeCacheInfo);
var
  AOldTransparentMode: Boolean;
begin
  PaintRangeBackground(AInfo);
  with AInfo^ do
  begin
    AssignFont(Style.Font, Canvas.Font);
    AOldTransparentMode := SetCanvasTransparentMode(Canvas, True);
    vgrDrawText(Canvas,
                DisplayText,
                InternalRect,
                (Style.Flags and vgrmask_RangeFlagsWordWrap) <> 0,
                RealAlign,
                Style.VertAlign,
                Style.Angle);
    SetCanvasTransparentMode(Canvas, AOldTransparentMode);
  end;
end;

procedure TvgrWBSheet.PaintLongRange(AInfo: pvgrRangeCacheInfo);
var
  X, Y: Integer;
  AOldTransparentMode: Boolean;
begin
  PaintRangeBackground(AInfo);
  AOldTransparentMode := SetCanvasTransparentMode(Canvas, True);
  with AInfo^ do
  begin
    AssignFont(Style.Font, Canvas.Font);
    case Style.VertAlign of
      vgrvaCenter:
        Y := InternalRect.Top + (InternalRect.Bottom - InternalRect.Top - Canvas.TextHeight(DisplayText)) div 2;
      vgrvaBottom:
        Y := InternalRect.Bottom - Canvas.TextHeight(DisplayText);
    else
      Y := InternalRect.Top;
    end;
    case RealAlign of
      vgrhaCenter:
        X := HorzCenterOfRect(FillRect) - TextWidth div 2;
      vgrhaRight:
        X := InternalRect.Right - TextWidth;
    else
      X := InternalRect.Left;
    end;
    ExtTextOut(Canvas.Handle, X, Y, ETO_CLIPPED, @InternalRect, PChar(DisplayText), Length(DisplayText), nil);
  end;
  SetCanvasTransparentMode(Canvas, AOldTransparentMode);
end;

procedure TvgrWBSheet.PaintCellBackground(AInfo: pvgrCellCacheInfo; X, Y: Integer);
begin
  if AInfo.Highlighted then
    Canvas.Brush.Color := GetHighlightColor(WorkbookGrid.BackgroundColor)
  else
    Canvas.Brush.Color := WorkbookGrid.BackgroundColor;
  Canvas.Brush.Style := bsSolid;
  Canvas.FillRect(FPaintInfo.GetCellScreenRect(X, Y));
end;

procedure TvgrWBSheet.PaintHorizontalBorders(ASelectionBorderRegion: HRGN; AGridBrush: HBRUSH);
var
  X, Y: Integer;
  ARegion: HRGN;
  ABorderInfo: pvgrBorderCacheInfo;
  ABrush: HBRUSH;

  procedure PaintNotSolidBorder(X, Y: Integer; ABorderInfo: pvgrBorderCacheInfo);
  var
    I: Integer;
    AInfo: pvgrBorderCacheInfo;
    AOldRegion: HRGN;
    APen, AOldPen: HPEN;
  begin
    ARegion := CreateRectRgnIndirect(ABorderInfo.Rect);
    try
      for I := X + 1 to FPaintInfo.FCellsRect.Right + 1 do
      begin
        AInfo := FPaintInfo.GetBorderInfo(I, Y, vgrboTop);
        if (not AInfo.Painted) and (ABorderInfo.Size = AInfo.Size) and (ABorderInfo.Style = AInfo.Style) then
        begin
          AInfo.Painted := True;
          CombineRectAndRegion(ARegion, AInfo.Rect);
        end;
      end;

      // paint this region
      // select border region
      AOldRegion := GetClipRegion(Canvas);
      CombineRegionAndRegionNoDelete(ARegion, ASelectionBorderRegion, RGN_DIFF);
      SetClipRegion(Canvas, ARegion);

      // create border pen
      APen := CreateBorderPen(ABorderInfo.Size, ABorderInfo.Style);
      AOldPen := SelectObject(Canvas.Handle, APen);

      // draw border
      for I := ABorderInfo.Rect.Top to ABorderInfo.Rect.Bottom - 1 do
      begin
        MoveToEx(Canvas.Handle, 0, I, nil);
        LineTo(Canvas.Handle, FPaintClipRect.Right, I);
      end;

      // restore pen
      SelectObject(Canvas.Handle, AOldPen);
      DeleteObject(APen);

      // restore clip region
      SetClipRegion(Canvas, AOldRegion);
      if AOldRegion <> 0 then
        DeleteObject(AOldRegion);
    finally
      DeleteObject(ARegion);
    end;
  end;

begin
  for Y := FPaintInfo.FCellsRect.Top to FPaintInfo.FCellsRect.Bottom + 1 do
    for X := FPaintInfo.FCellsRect.Left to FPaintInfo.FCellsRect.Right + 1 do
    begin
      ABorderInfo := FPaintInfo.GetBorderInfo(X, Y, vgrboTop);
      if (ABorderInfo <> nil) and (ABorderInfo.Size > 0) then
      begin
        if ABorderInfo.Style = nil then
          FillRect(Canvas.Handle, ABorderInfo.Rect, AGridBrush)
        else
          if ABorderInfo.Style.Pattern = vgrbsSolid then
          begin
            ABrush := CreateSolidBrush(GetRGBColor(ABorderInfo.Style.Color));
            FillRect(Canvas.Handle, ABorderInfo.Rect, ABrush);
            DeleteObject(ABrush);
          end
          else
            PaintNotSolidBorder(X, Y, ABorderInfo);
      end;
    end;
end;

procedure TvgrWBSheet.PaintVerticalBorders(ASelectionBorderRegion: HRGN; AGridBrush: HBRUSH);
var
  X, Y: Integer;
  ARegion: HRGN;
  ABorderInfo: pvgrBorderCacheInfo;
  ABrush: HBRUSH;

  procedure PaintNotSolidBorder(X, Y: Integer; ABorderInfo: pvgrBorderCacheInfo);
  var
    I: Integer;
    AInfo: pvgrBorderCacheInfo;
    AOldRegion: HRGN;
    APen, AOldPen: HPEN;
  begin
    ARegion := CreateRectRgnIndirect(ABorderInfo.Rect);
    try
      for I := Y + 1 to FPaintInfo.FCellsRect.Bottom + 1 do
      begin
        AInfo := FPaintInfo.GetBorderInfo(X, I, vgrboLeft);
        if (not AInfo.Painted) and (ABorderInfo.Size = AInfo.Size) and (ABorderInfo.Style = AInfo.Style) then
        begin
          AInfo.Painted := True;
          CombineRectAndRegion(ARegion, AInfo.Rect);
        end;
      end;

      // paint this region
      // select border region
      AOldRegion := GetClipRegion(Canvas);
      CombineRegionAndRegionNoDelete(ARegion, ASelectionBorderRegion, RGN_DIFF);
      SetClipRegion(Canvas, ARegion);

      // create border pen
      APen := CreateBorderPen(ABorderInfo.Size, ABorderInfo.Style);
      AOldPen := SelectObject(Canvas.Handle, APen);

      // draw border
      for I := ABorderInfo.Rect.Left to ABorderInfo.Rect.Right - 1 do
      begin
        MoveToEx(Canvas.Handle, I, 0, nil);
        LineTo(Canvas.Handle, I, FPaintClipRect.Bottom);
      end;

      // restore pen
      SelectObject(Canvas.Handle, AOldPen);
      DeleteObject(APen);

      // restore clip region
      SetClipRegion(Canvas, AOldRegion);
      if AOldRegion <> 0 then
        DeleteObject(AOldRegion);
    finally
      DeleteObject(ARegion);
    end;
  end;

begin
  for X := FPaintInfo.FCellsRect.Left to FPaintInfo.FCellsRect.Right + 1 do
    for Y := FPaintInfo.FCellsRect.Top to FPaintInfo.FCellsRect.Bottom + 1 do
    begin
      ABorderInfo := FPaintInfo.GetBorderInfo(X, Y, vgrboLeft);
      if (ABorderInfo <> nil) and (ABorderInfo.Size > 0) then
      begin
        if ABorderInfo.Style = nil then
          FillRect(Canvas.Handle, ABorderInfo.Rect, AGridBrush)
        else
          if ABorderInfo.Style.Pattern = vgrbsSolid then
          begin
            ABrush := CreateSolidBrush(GetRGBColor(ABorderInfo.Style.Color));
            FillRect(Canvas.Handle, ABorderInfo.Rect, ABrush);
            DeleteObject(ABrush);
          end
          else
            PaintNotSolidBorder(X, Y, ABorderInfo);
      end;
    end;
end;

procedure TvgrWBSheet.PaintBorders(ASelectionBorderRegion: HRGN);
var
  AOldBackColor: TColor;
  AGridBrush: HBRUSH;
begin
  AGridBrush := CreateSolidBrush(GetRGBColor(WorkbookGrid.GridColor));
  AOldBackColor := SetBkColor(Canvas.Handle, ColorToRGB(WorkbookGrid.BackgroundColor));
  try
    PaintHorizontalBorders(ASelectionBorderRegion, AGridBrush);
    PaintVerticalBorders(ASelectionBorderRegion, AGridBrush);
  finally
    DeleteObject(AGridBrush);
    SetBkColor(Canvas.Handle, AOldBackColor);
  end;
end;

procedure TvgrWBSheet.Paint;
var
  I, J: integer;
  rRegions, rPixels, ASelectionOuterRect, ASelectionInnerRect: TRect;
  ASelectionBorderRegion: HRGN;
  ACellInfo: pvgrCellCacheInfo;
  ADimensions, ADimensionsPixels: TRect;
begin
  // Save ClipRect
  FPaintClipRect := Canvas.ClipRect;
  if IsRectEmpty(FPaintClipRect) then exit;
  
  if Worksheet = nil then
  begin
    Canvas.Brush.Style := bsSolid;
    Canvas.Brush.Color := WorkbookGrid.BackgroundColor;
    Canvas.FillRect(FPaintClipRect);
  end
  else
  begin
    // 0. Calculate rectangle which must painted
    WorkbookGrid.GetPaintRects(FPaintClipRect, rRegions, rPixels, WorkbookGrid.LeftCol, WorkbookGrid.TopRow);

    // 1. Calculate paint info
    FPaintInfo.Update(rRegions, FPaintClipRect);

    // 2. Draw selection border
    ASelectionBorderRegion := 0;
    with FSelectedRegion do
      for I := 0 to RectsCount-1 do
      begin
        FPaintInfo.GetRectScreenPixels(Rects[i]^, ASelectionOuterRect);
        if RectOverRect(FPaintClipRect, ASelectionOuterRect) then
        begin
          ASelectionInnerRect := ASelectionOuterRect;
          with ASelectionOuterRect do
          begin
            Left := Left - 2;
            Top := Top - 2;
            Right := Right + 1;
            Bottom := Bottom + 1;
          end;
          with ASelectionInnerRect do
          begin
            Left := Left + 1;
            Top := Top + 1;
            Right := Right - 2;
            Bottom := Bottom - 2;
          end;
          CombineRectAndRegion(ASelectionBorderRegion, ASelectionOuterRect);
          CombineRectAndRegion(ASelectionBorderRegion, ASelectionInnerRect, RGN_DIFF);
        end;
      end;
    if ASelectionBorderRegion <> 0 then
    begin
      // Draw selection border
      Canvas.Brush.Style := bsSolid;
      Canvas.Brush.Color := WorkbookGrid.SelectionBorderColor;
      FillRegion(Canvas, ASelectionBorderRegion);
      // Exclude selection border from clip rect
      ExcludeClipRegion(Canvas, ASelectionBorderRegion);
    end;

    // 2.1 draw worksheet bounds
    if WorkbookGrid.ShowWorksheetDimensions then
    begin
      ADimensions := Worksheet.Dimensions;
//      if not IsRectEmpty(ADimensions) then
      begin
        FPaintInfo.GetRectScreenPixels(ADimensions, ADimensionsPixels);
        with ADimensionsPixels do
          ASelectionOuterRect := Rect(0, Bottom - 1, Right + 1, Bottom + 1);
        if RectOverRect(ASelectionOuterRect, FPaintClipRect) then
        begin
          Canvas.Brush.Style := bsSolid;
          Canvas.Brush.Color := WorkbookGrid.WorksheetDimensionsColor;
          Canvas.FillRect(ASelectionOuterRect);
          ExcludeClipRect(Canvas, ASelectionOuterRect);
        end;

        with ADimensionsPixels do
          ASelectionOuterRect := Rect(Right - 1, 0, Right + 1, Bottom + 1);
        if RectOverRect(ASelectionOuterRect, FPaintClipRect) then
        begin
          Canvas.Brush.Style := bsSolid;
          Canvas.Brush.Color := WorkbookGrid.WorksheetDimensionsColor;
          Canvas.FillRect(ASelectionOuterRect);
          ExcludeClipRect(Canvas, ASelectionOuterRect);
        end;
      end;
    end;

    // 3. draw normal ranges
    for I := 0 to FPaintInfo.FRangesCount - 1 do
      if not FPaintInfo.FRanges[I].Long then
        PaintNormalRange(@FPaintInfo.FRanges[I]);

    // 4. draw background
    for I := FPaintInfo.FCellsRect.Left to FPaintInfo.FCellsRect.Right do
      for J := FPaintInfo.FCellsRect.Top to FPaintInfo.FCellsRect.Bottom do
      begin
        ACellInfo := FPaintInfo.GetCellInfo(I, J);
        if (ACellInfo <> nil) and
           (ACellInfo.Range = nil) and
           (FPaintInfo.ColPixelWidth(I) <> 0) and
           (FPaintInfo.RowPixelHeight(J) <> 0) then
          PaintCellBackground(ACellInfo, I, J)
      end;

    // 5. Draw long ranges
    for I := 0 to Length(FPaintInfo.FLeftLongRanges) - 1 do
      if FPaintInfo.FLeftLongRanges[I].Range <> nil then
        PaintLongRange(@FPaintInfo.FLeftLongRanges[I]);
    for I := 0 to Length(FPaintInfo.FRightLongRanges) - 1 do
      if FPaintInfo.FRightLongRanges[I].Range <> nil then
        PaintLongRange(@FPaintInfo.FRightLongRanges[I]);
    for I := 0 to FPaintInfo.FRangesCount - 1 do
      if FPaintInfo.FRanges[I].Long then
        PaintLongRange(@FPaintInfo.FRanges[I]);

    // 6.Draw borders
    PaintBorders(ASelectionBorderRegion);

    // 7. Delete section region
    DeleteObject(ASelectionBorderRegion);
    
    // 8. Finalize PaintInfo
    FPaintInfo.AfterPaint;
  end;
end;

procedure TvgrWBSheet.DoScroll(const OldPos, NewPos : TPoint);
begin
  UpdateInplaceEditPosition;
  Invalidate;
end;

procedure TvgrWBSheet.KbdBeforeSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
begin
  FKbdLeftOffs := 0;
  FKbdTopOffs := 0;
  FKbdOffsetX := 0;
  FKbdOffsetY := 0;
end;

procedure TvgrWBSheet.KbdAfterSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
begin
  if AShiftPressed then
  begin
    SetUserSelection(0, FKbdOffsetX, FKbdOffsetY, FKbdLeftOffs, FKbdTopOffs);
  end
  else
  begin
    if ((FKbdOffsetX = -1) and not ColVisible(FCurRect.Left)) or
       ((FKbdOffsetX = 1) and not ColVisible(FCurRect.Right)) or
       ((FKbdOffsetY = -1) and not RowVisible(FCurRect.Top)) or
       ((FKbdOffsetY = 1) and not RowVisible(FCurRect.Bottom)) then
      OffsetFCurRect(FKbdOffsetX, FKbdOffsetY);

    FCurRect := CheckCell(Point(FCurRect.Left, FCurRect.Top));
    MakeRectVisible(FCurRect, FKbdLeftOffs, FKbdTopOffs);
    FSelStartRect := FCurRect;
    SetSelection(FCurRect);
  end;
end;

function TvgrWBSheet.KbdLeft: Boolean;
begin
  Result := FCurRect.Left <> 0;
  if Result then
  begin
    FKbdOffsetX := -1;
    FCurRect.Left := FCurRect.Left - 1;
    FCurRect.Right := FCurRect.Left;
  end;
end;

function TvgrWBSheet.KbdLeftCell: Boolean;
begin
  Result := FCurRect.Left <> 0;
  if Result then
  begin
    FKbdOffsetX := -1;
    FCurRect.Left := 0;
    FCurRect.Right := 0;
  end;
end;

function TvgrWBSheet.KbdRight: Boolean;
begin
  Result := True;
  FKbdOffsetX := 1;
  FCurRect.Right := FCurRect.Right + 1;
  FCurRect.Left := FCurRect.Right;
end;

function TvgrWBSheet.KbdRightCell: Boolean;
begin
  Result := (FCurRect.Right <> DimensionsRight);
  if Result then
  begin
    FKbdOffsetX := 1;
    FCurRect.Right := DimensionsRight;
    FCurRect.Left := FCurRect.Right;
  end;
end;

function TvgrWBSheet.KbdUp: Boolean;
begin
  Result := FCurRect.Top <> 0;
  if Result then
  begin
    FKbdOffsetY := -1;
    FCurRect.Top := FCurRect.Top - 1;
    FCurRect.Bottom := FCurRect.Top;
  end;
end;

function TvgrWBSheet.KbdTopCell: Boolean;
begin
  Result := FCurRect.Top <> 0;
  if Result then
  begin
    FKbdOffsetY := -1;
    FCurRect.Top := FCurRect.Top - 1;
    FCurRect.Bottom := FCurRect.Top;
  end;
end;

function TvgrWBSheet.KbdDown: Boolean;
begin
  Result := True;
  FKbdOffsetY := 1;
  FCurRect.Bottom := FCurRect.Bottom + 1;
  FCurRect.Top := FCurRect.Bottom;
end;

function TvgrWBSheet.KbdBottomCell: Boolean;
begin
  Result := FCurRect.Bottom <> DimensionsBottom;
  if Result then
  begin
    FKbdOffsetY := 1;
    FCurRect.Bottom := DimensionsBottom;
    FCurRect.Top := FCurRect.Bottom;
  end;
end;

function TvgrWBSheet.KbdPageUp: Boolean;
begin
  Result := FCurRect.Top <> 0;
  if Result then
  begin
    FKbdOffsetY := -1;
    FKbdTopOffs := -Min(TopRow, WorkbookGrid.FFullVisibleRowsCount);
    FCurRect.Top := FCurRect.Top + FKbdTopOffs;
    FCurRect.Bottom := FCurRect.Top;
  end;
end;

function TvgrWBSheet.KbdPageDown: Boolean;
begin
  Result := True;
  FKbdOffsetY := 1;
  FKbdTopOffs := WorkbookGrid.FFullVisibleRowsCount;
  FCurRect.Bottom := FCurRect.Bottom + FKbdTopOffs;
  FCurRect.Top := FCurRect.Bottom;
end;

function TvgrWBSheet.KbdLeftTopCell: Boolean;
begin
  Result := (FCurRect.Left <> 0) or (FCurRect.Top <> 0);
  if Result then
  begin
    FCurRect.Left := 0;
    FCurRect.Top := 0;
    FCurRect.Right := 0;
    FCurRect.Bottom := 0;
    FKbdOffsetX := -1;
    FKbdOffsetY := -1;
  end;
end;

function TvgrWBSheet.KbdRightBottomCell: Boolean;
begin
  Result := (FCurRect.Right <> DimensionsRight) or (FCurRect.Bottom <> DimensionsBottom);
  if Result then
  begin
    FCurRect.Left := DimensionsRight;
    FCurRect.Top := DimensionsBottom;
    FCurRect.Right := DimensionsRight;
    FcurRect.Bottom := DimensionsBottom;
    FKbdOffsetX := 1;
    FKbdOffsetY := 1;
  end;
end;

function TvgrWBSheet.KbdInplaceEdit: Boolean;
begin
  StartInplaceEdit;
  Result := True;
end;

function TvgrWBSheet.KbdClearValue: Boolean;
begin
  if (Worksheet <> nil) and not WorkbookGrid.ReadOnly then
    WorkbookGrid.ClearSelection([vgrccValue]);
  Result := True;
end;

function TvgrWBSheet.KbdCellProperties: Boolean;
begin
  WorkbookGrid.DoCellProperties;
  Result := True;
end;

function TvgrWBSheet.KbdSelectedRowsAutoHeight: Boolean;
begin
  if not IsInplaceEdit and not WorkbookGrid.ReadOnly then
    WorkbookGrid.SetAutoHeightForSelectedRows;
  Result := True;
end;

function TvgrWBSheet.KbdSelectedColumnsAutoHeight: Boolean;
begin
  if not IsInplaceEdit and not WorkbookGrid.ReadOnly then
    WorkbookGrid.SetAutoWidthForSelectedColumns;
  Result := True;
end;

function TvgrWBSheet.KbdCut: Boolean;
begin
  WorkbookGrid.CutToClipboard;
  Result := True;
end;

function TvgrWBSheet.KbdCopy: Boolean;
begin
  WorkbookGrid.CopyToClipboard;
  Result := True;
end;

function TvgrWBSheet.KbdPaste: Boolean;
begin
  WorkbookGrid.PasteFromClipboard;
  Result := True;
end;

procedure TvgrWBSheet.KbdKeyPress(Key: Char);
begin
  if not IsInplaceEdit then
  begin
    StartInplaceEdit(Key);
  end;
end;

procedure TvgrWBSheet.KeyDown(var Key: Word; Shift: TShiftState);
begin
  if Worksheet <> nil then
    KeyboardFilter.KeyDown(Key, Shift);
end;

procedure TvgrWBSheet.KeyPress(var Key: Char);
begin
  if Worksheet <> nil then
    KeyboardFilter.KeyPress(Key);
end;

function TvgrWBSheet.GetInplaceRange: IvgrRange;
begin
  with FInplaceEditRect do
    Result := Worksheet.RangesList.Find(Left, Top, Right, Bottom);
end;

function TvgrWBSheet.IsLongRangeEdited: Boolean;
var
  ARange: IvgrRange;
begin
  ARange := GetInplaceRange;
  Result := (ARange = nil) or ((ARange.Left = ARange.Right) and
                               (ARange.Top = ARange.Bottom) and
                               (not ARange.WordWrap) and
                               (ARange.Angle = 0));
end;

function TvgrWBSheet.GetInplaceRangeHorzAlignment: TvgrRangeHorzAlign;
var
  ARange: IvgrRange;
begin
  ARange := GetInplaceRange;
  if ARange = nil then
    Result := vgrhaLeft
  else
    Result := GetAutoAlignForRangeValue(ARange.HorzAlign, ARange.ValueType);
end;

function TvgrWBSheet.GetInplaceRangeVertAlignment: TvgrRangeVertAlign;
var
  ARange: IvgrRange;
begin
  ARange := GetInplaceRange;
  if ARange = nil then
    Result := vgrvaTop
  else
    Result := ARange.VertAlign;
end;

procedure TvgrWBSheet.UpdateInplaceEditPosition;
begin
  if IsInplaceEdit then
  begin
    GetRectInternalScreenPixels(FInplaceEditRect, FInplaceEditPixelRect);
    InplaceEdit.UpdatePosition;
  end;
end;

procedure TvgrWBSheet.StartInplaceEdit(AStartKey: Char = #0; AActiveInplaceEdit: Boolean = True);
const
  aAlignment: Array [TvgrRangeHorzAlign] of TAlignment = (taLeftJustify, taLeftJustify, taCenter, taRightJustify);
var
  ARange: IvgrRange;
begin
  if (Worksheet <> nil) and not IsInplaceEdit then
  begin
    FInplaceEditRect := FSelStartRect;
    GetRectInternalScreenPixels(FInplaceEditRect, FInplaceEditPixelRect);

    with FInplaceEditRect do
      ARange := Worksheet.RangesList.Find(Left, Top, Right, Bottom);
    if ARange = nil then
    begin
      InplaceEdit.Color := WorkbookGrid.BackgroundColor;
      InplaceEdit.Text := '';
      Worksheet.RangesList.GetDefaultFont(InplaceEdit.Font);
      InplaceEdit.Alignment := taLeftJustify;
    end
    else
    begin
      if (AStartKey <> #0) and not WorkbookGrid.ReadOnly then
        InplaceEdit.Clear
      else
        InplaceEdit.Text := ARange.StringValue;
      ARange.Font.AssignTo(InplaceEdit.Font);
      if ARange.FillBackColor = clNone then 
        InplaceEdit.Color := WorkbookGrid.BackgroundColor
      else
        InplaceEdit.Color := ARange.FillBackColor;

      InplaceEdit.Alignment := aAlignment[GetInplaceRangeHorzAlignment];
    end;

    InplaceEdit.Init;
    with FInplaceEditPixelRect do
      InplaceEdit.SetBounds(Left, Top, Right - Left, Bottom - Top);
    InplaceEdit.UpdatePosition;
    InplaceEdit.Visible := True;

    with FormulaPanel do
    begin
      OkButton.Visible := True;
      CancelButton.Visible := True;
    end;

    if AActiveInplaceEdit then
      InplaceEdit.SetFocus
    else
      FormulaPanel.Edit.SetFocus;

    if AStartKey <> #0 then
      PostMessage(InplaceEdit.Handle, WM_CHAR, Integer(AStartKey), 0);
  end;
end;

procedure TvgrWBSheet.EndInplaceEdit(ACancel: Boolean);
var
  ARange: IvgrRange;
  ANewRange: Boolean;
begin
  if (Worksheet <> nil) and IsInplaceEdit then
  begin
    InplaceEdit.Visible := False;
    SetFocus;
    if not ACancel and not WorkbookGrid.ReadOnly then
    begin
      with FInplaceEditRect do
      begin
        ARange := Worksheet.RangesList.Find(Left, Top, Right, Bottom);
        ANewRange := ARange = nil;
        if ARange = nil then
        begin
          ARange := Worksheet.Ranges[Left, Top, Right, Bottom];
          ARange.WordWrap := False;
        end;
      end;
      WorkbookGrid.AssignInplaceEditValueToRange(InplaceEdit.Text, ARange, ANewRange);
//      ARange.StringValue := InplaceEdit.Text;
    end;
    InplaceEdit.Clear;
    with FormulaPanel do
    begin
      OkButton.Visible := False;
      CancelButton.Visible := False;
    end;
    WorkbookGrid.CheckSelectedTextChanged;
  end;
end;

procedure TvgrWBSheet.SetSelection(const ASelectionRect: TRect);
var
  i: Integer;
  rRegions,rPixels: TRect;
begin
  // calculate invalidate region
  rRegions := GetSelectionBounds(ASelectionRect);
  // check new selection
  if (FSelectedRegion.RectsCount = 1) and
     EqualRect(FSelectedRegion.Rects[0]^, rRegions) then
    exit;

  for i:=0 to FSelectedRegion.RectsCount-1 do
    with FSelectedRegion.Rects[i]^ do
    begin
      rRegions.Left := Min(rRegions.Left,Left);
      rRegions.Top := Min(rRegions.Top,Top);
      rRegions.Right := Max(rRegions.Right,Right);
      rRegions.Bottom := Max(rRegions.Bottom,Bottom);
    end;
  GetRectScreenPixels(rRegions,rPixels);
  // store new selection
  FSelectedRegion.SetSelection(FTempSelectionBounds);
  // repaint windows
  InflateRect(rPixels, 2, 2);
  InvalidateRect(Handle,@rPixels,false);
  ColsPanel.DoSelectionChanged;
  RowsPanel.DoSelectionChanged;

  WorkbookGrid.DoSelectionChange;
end;

procedure TvgrWBSheet.CheckSelStartRect;
begin
  if (FSelStartRect.Left <> FSelStartRect.Right) or
     (FSelStartRect.Top <> FSelStartRect.Bottom) then
    with FSelStartRect do
      FSelStartRect := Rect(Left, Top, Left, Top);
  FSelStartRect := GetSelectionBounds(FSelStartRect);
end;

function TvgrWBSheet.CheckSelections(var ARepaintRect: TRect): Boolean;
var
  I: Integer;
  ARangeRect, ANewBounds: TRect;

  procedure CheckRepaintRect(const ARect: TRect);
  begin
    if ARangeRect.Left > ARect.Left then
      ARangeRect.Left := ARect.Left;
    if ARangeRect.Top > ARect.Top then
      ARangeRect.Top := ARect.Top;
    if ARangeRect.Right < ARect.Right then
      ARangeRect.Right := ARect.Right;
    if ARangeRect.Bottom < ARect.Bottom then
      ARangeRect.Bottom := ARect.Bottom;
  end;

begin
  Result := False;
  if Worksheet <> nil then
  begin
    ARangeRect := Rect(MaxInt, MaxInt, 0, 0);
    for I := 0 to FSelectedRegion.RectsCount - 1 do
    begin
      ANewBounds := GetSelectionBounds(FSelectedRegion.Rects[I]^);
      if not EqualRect(ANewBounds, FSelectedRegion.Rects[I]^) then
      begin
        CheckRepaintRect(FSelectedRegion.Rects[I]^);
        CheckRepaintRect(ANewBounds);
        FSelectedRegion.Rects[I] := @ANewBounds;
        Result := True;
      end;
    end;

    CheckSelStartRect;
    
    if Result then
    begin
      GetRectScreenPixels(ARangeRect, ARepaintRect);
      InflateRect(ARepaintRect, 2, 2);
    end;
  end;
end;

function TvgrWBSheet.MakeRectVisible(const r: TRect; LeftOffs,TopOffs: Integer): boolean;
var
  NewPos: TPoint;
begin
  if r.Right - r.Left > WorkbookGrid.FFullVisibleColsCount then
    NewPos.X := r.Left
  else
  begin
    if r.Left <= LeftCol + LeftOffs then
      NewPos.X := r.Left
    else
      if r.Right >= LeftCol + LeftOffs + WorkbookGrid.FFullVisibleColsCount then
      begin
        if WorkbookGrid.FFullVisibleColsCount > 0 then
          NewPos.X := r.Right - WorkbookGrid.FFullVisibleColsCount + 1
        else
          NewPos.X := LeftCol + LeftOffs;
      end
      else
        NewPos.X := LeftCol + LeftOffs;
  end;

  if r.Bottom - r.Top > WorkbookGrid.FFullVisibleRowsCount then
    NewPos.Y := r.Top
  else
  begin
    if r.Top <= TopRow + TopOffs then
      NewPos.Y := r.Top
    else
      if r.Bottom >= TopRow + TopOffs + WorkbookGrid.FFullVisibleRowsCount then
      begin
        if WorkbookGrid.FFullVisibleRowsCount > 0 then
          NewPos.Y := r.Bottom - WorkbookGrid.FFullVisibleRowsCount + 1
        else
          NewPos.Y := r.Top
      end
      else
        NewPos.Y := TopRow + TopOffs;
  end;

  Result := (NewPos.X <> LeftCol) or (NewPos.Y <> TopRow);
  if Result then
  begin
    WorkbookGrid.DoScroll(Point(LeftCol, TopRow), NewPos);
  end;
end;

function TvgrWBSheet.GetSelRect(const ACurRect, ASelStartRect: TRect): TRect;
begin
  Result := NormalizeRect(Min(ACurRect.Left, ASelStartRect.Left),
                          Min(ACurRect.Top, ASelStartRect.Top),
                          Max(ACurRect.Right, ASelStartRect.Right),
                          Max(ACurRect.Bottom, ASelStartRect.Bottom));
end;

function TvgrWBSheet.CheckCell(const ACell: TPoint): TRect;
var
  ARange: IvgrRange;
begin
  ARange := Worksheet.RangesList.FindAtCell(ACell.X, ACell.Y);
  if ARange <> nil then
    Result := ARange.Place
  else
    Result := Rect(ACell.X, ACell.Y, ACell.X, ACell.Y);
end;

procedure TvgrWBSheet.CheckSelectionBoundsCallback(AItem: IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
begin
  with AItem as IvgrRange do
  begin
    if FTempSelectionBounds.Left > Place.Left then
      FTempSelectionBounds.Left := Place.Left;
    if FTempSelectionBounds.Right < Place.Right then
      FTempSelectionBounds.Right := Place.Right;
    if FTempSelectionBounds.Top > Place.Top then
      FTempSelectionBounds.Top := Place.Top;
    if FTempSelectionBounds.Bottom < Place.Bottom then
      FTempSelectionBounds.Bottom := Place.Bottom;
  end;
end;

procedure TvgrWBSheet.UpdateHorzPoints;
var
  I, ANewLength: Integer;
begin
  if Worksheet = nil then
    ANewLength := LeftCol + VisibleColsCount + 1
  else
    ANewLength := Max(LeftCol + VisibleColsCount + 1, DimensionsRight);
  SetLength(FHorzPoints, ANewLength);
  if Length(FHorzPoints) > 0 then
  begin
    with FHorzPoints[0] do
    begin
      Pos := 0;
      Size := ColPixelWidthNoCache(0);
    end;
    for I := 1 to High(FHorzPoints) do
    begin
      FHorzPoints[I].Pos := FHorzPoints[I - 1].Pos + FHorzPoints[I - 1].Size;
      FHorzPoints[I].Size := ColPixelWidthNoCache(I);
    end;
  end;
end;

function TvgrWBSheet.GetSelectionBounds(ASelectionRect: TRect): TRect;
var
  ATempRect: TRect;
begin
  if Worksheet = nil then
    Result := ASelectionRect
  else
  begin
    FTempSelectionBounds := ASelectionRect;
    ATempRect := ASelectionRect;
    Worksheet.RangesList.FindAndCallBack(ATempRect, CheckSelectionBoundsCallback, nil);
    if EqualRect(FTempSelectionBounds, ASelectionRect) then
      Result := FTempSelectionBounds
    else
    begin
      ATempRect := FTempSelectionBounds;
      Result := GetSelectionBounds(ATempRect);
    end;
  end;
end;

procedure TvgrWBSheet.OffsetFCurRect(AOffsetX, AOffsetY: Integer);
begin
  if AOffsetX = -1 then
  begin
    repeat
      FCurRect.Left := FCurRect.Left - 1;
      FCurRect.Right := FCurRect.Left;
    until (FCurRect.Left = 0) or ColVisible(FCurRect.Left);
  end
  else
    if AOffsetX = 1 then
    begin
      repeat
        FCurRect.Right := FCurRect.Right + 1;
        FCurRect.Left := FCurRect.Right;
      until ColVisible(FCurRect.Right)
    end;
  if AOffsetY = -1 then
  begin
    repeat
      FCurRect.Top := FCurRect.Top - 1;
      FCurRect.Bottom := FCurRect.Top;
    until (FCurRect.Top = 0) or RowVisible(FCurRect.Top);
  end
  else
    if AOffsetY = 1 then
    begin
      repeat
        FCurRect.Bottom := FCurRect.Bottom + 1;
        FCurRect.Top := FCurRect.Bottom;
      until RowVisible(FCurRect.Bottom);
    end;
end;

procedure TvgrWBSheet.SetUserSelection(ASelectionIndex, AOffsetX, AOffsetY, ALeftOffs, ATopOffs: Integer);
var
  ACurRect: PRect;
  ANewRect, APriorCurRect, ANewSelectionRect: TRect;
begin
  ACurRect := FSelectedRegion.Rects[ASelectionIndex];
  if ((AOffsetX = -1) and not ColVisible(FCurRect.Left)) or
     ((AOffsetX = 1) and not ColVisible(FCurRect.Right)) or
     ((AOffsetY = -1) and not RowVisible(FCurRect.Top)) or
     ((AOffsetY = 1) and not RowVisible(FCurRect.Bottom)) then
    OffsetFCurRect(AOffsetX, AOffsetY);

  repeat
    ANewSelectionRect := GetSelectionBounds(Self.GetSelRect(FCurRect, FSelStartRect));
    if not EqualRect(ANewSelectionRect, ACurRect^) then
      break;
    OffsetFCurRect(AOffsetX, AOffsetY);
  until (AOffsetX = 0) and (AOffsetY = 0);

  ANewRect := ANewSelectionRect;
  repeat
    APriorCurRect := FCurRect;
    OffsetFCurRect(AOffsetX, AOffsetY);
    ANewRect := GetSelectionBounds(Self.GetSelRect(FCurRect, FSelStartRect));
  until not EqualRect(ANewRect, ANewSelectionRect);
  FCurRect := APriorCurRect;

  MakeRectVisible(FCurRect, ALeftOffs, ATopOffs);
  SetSelection(ANewSelectionRect);
end;

procedure TvgrWBSheet.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
end;

procedure TvgrWBSheet.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcChangeWorksheet:
      begin
      end;
    vgrwcChangeCol, vgrwcChangeRow:
      begin
        Invalidate;
        UpdateInplaceEditPosition;
      end;
    vgrwcNewBorder, vgrwcChangeBorder:
      begin
      end;
    vgrwcNewRange, vgrwcChangeRange:
      begin
        CheckSelStartRect;
      end;
    vgrwcDeleteRange, vgrwcDeleteBorder, vgrwcDeleteCol, vgrwcDeleteRow:
      begin
        CheckSelStartRect;
      end;
    vgrwcChangeHorzSection, vgrwcChangeVertSection, vgrwcNewHorzSection, vgrwcDeleteHorzSection, vgrwcNewVertSection, vgrwcDeleteVertSection:
      begin
      end;
    vgrwcUpdateAll, vgrwcChangeWorksheetContent:
      begin
        UpdateInplaceEditPosition;
      end;
  end;
  Invalidate;
end;

procedure TvgrWBSheet.GetPointInfoAt(X,Y: Integer; var PointInfo: rvgrWBSheetPointInfo);
var
  i: Integer;
  r1,r2: TRect;
  ARange: IvgrRange;
begin
  PointInfo.Cell := Point(-1, -1);
  PointInfo.DownRect := Rect(-1, -1, -1, -1);
  PointInfo.RangeIndex := -1;
  PointInfo.Place := vgrspNone;
  if (X < 0) or (Y < 0) or (Worksheet = nil) then exit;
  // 1. Check selection
  with FSelectedRegion do
    for i:=0 to RectsCount-1 do
    begin
      GetRectScreenPixels(Rects[i]^, r1);
      r2 := r1;
      InflateRect(r1,2,2);
      InflateRect(r2,-1,-1);
      if PtInRect(r1,Point(X,Y)) and not PtInRect(r2,Point(X,Y)) then
        begin
          PointInfo.Place := vgrspSelectionBorder;
          exit;
        end;
    end;
  // Get cell
  WorkbookGrid.GetPaintRects(Rect(X,Y,X,Y),r1,r2, WorkbookGrid.LeftCol, WorkbookGrid.TopRow);
  PointInfo.Cell.X := r1.Left;
  PointInfo.Cell.Y := r1.Top;
  ARange := Worksheet.RangesList.FindAtCell(r1.Left, r1.Top);
  if ARange <> nil then
    PointInfo.RangeIndex := ARange.ItemIndex
  else
    PointInfo.RangeIndex := -1;
  PointInfo.Place := vgrspInsideCell;
  if PointInfo.RangeIndex = -1 then
    PointInfo.DownRect := Rect(PointInfo.Cell.X, PointInfo.Cell.Y, PointInfo.Cell.X, PointInfo.Cell.Y)
  else
    PointInfo.DownRect := ARange.Place;
end;

procedure TvgrWBSheet.UpdateMouseCursor(const PointInfo: rvgrWBSheetPointInfo);
var
  NewCursor: TCursor;
begin
  NewCursor := FCellCursor;
  case PointInfo.Place of
    vgrspSelectionBorder: NewCursor := crDrag;
  end;
  if NewCursor <> Cursor then
    Cursor := NewCursor;
end;

procedure TvgrWBSheet.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  PointInfo: rvgrWBSheetPointInfo;
begin
  if FMouseOperation <> vgrmoNone then
    case FMouseOperation of
      vgrWBSheetSelect:
        begin
          if X > Width then
            FScrollX := 1
          else
            if X < 0 then
              FScrollX := -1
            else
              FScrollX := 0;
          if Y > Height then
            FScrollY := 1
          else
            if Y < 0 then
              FScrollY := -1
            else
              FScrollY := 0;

          if (FScrollX <> 0) or (FScrollY <> 0) then
            FScrollTimer.Enabled := True
          else
          begin
            GetPointInfoAt(X, Y, PointInfo);
            if PointInfo.Place = vgrspInsideCell then
            begin
              FCurRect := PointInfo.DownRect;
              SetSelection(GetSelRect(FCurRect, FSelStartRect));
            end;
          end;
        end;
    end
  else
  begin
    GetPointInfoAt(X, Y, PointInfo);
    UpdateMouseCursor(PointInfo);
  end;
  WorkbookGrid.DoMouseMove(Self,Shift, X, Y);
end;

procedure TvgrWBSheet.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    SetFocus;
    GetPointInfoAt(X,Y,FDownPointInfo);
    if ssDouble in Shift then
      case FDownPointInfo.Place of
        vgrspInsideCell:
          begin
            if ssLeft in Shift then
              StartInplaceEdit;
          end;
      end
    else
      if Button = mbLeft then
        case FDownPointInfo.Place of
          vgrspInsideCell:
            begin
              EndInplaceEdit(False);
              if (Shift * [ssShift,ssCtrl]) = [] then
                begin
                  FSelStartRect := FDownPointInfo.DownRect;
                  FCurRect := FDownPointInfo.DownRect;
                  SetSelection(FSelStartRect);
                  FMouseOperation := vgrWBSheetSelect;
                end
              else
                if ssShift in Shift then
                  begin
                    FCurRect := FDownPointInfo.DownRect;
                    SetSelection(GetSelRect(FSelStartRect, FCurRect));
                    FMouseOperation := vgrWBSheetSelect;
                  end;
            end;
          vgrspSelectionBorder:
            begin
              EndInplaceEdit(False);
              FMouseOperation := vgrmoWBSheetSelectionDrag;
            end;
        end;
  end;
end;

procedure TvgrWBSheet.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APointInfo: TvgrPointInfo;
  AExecuteDefault: Boolean;
begin
  FScrollTimer.Enabled := False;
  WorkbookGrid.DoMouseDown(Self, Button, Shift, X, Y, AExecuteDefault);
  if AExecuteDefault then
  begin
    case FMouseOperation of
      vgrWBSheetSelect:
        begin
          FMouseOperation := vgrmoNone;
        end;
      vgrmoWBSheetSelectionDrag:
        begin
          FMouseOperation := vgrmoNone;
        end;
    end;
    if Button = mbRight then
    begin
      GetPointInfoAt(Point(X, Y), APointInfo);
      WorkbookGrid.ShowPopupMenu(ClientToScreen(Point(X, Y)), APointInfo);
      APointInfo.Free;
    end;
  end;
end;

procedure TvgrWBSheet.DblClick;
var
  AExecuteDefault: Boolean;
begin
  WorkbookGrid.DoDblClick(Self, AExecuteDefault);
end;

procedure TvgrWBSheet.GetPointInfoAt(const APoint: TPoint; var APointInfo: TvgrPointInfo);
var
  PointInfo: rvgrWBSheetPointInfo;
begin
  GetPointInfoAt(APoint.X, APoint.Y, PointInfo);
  with PointInfo do
    APointInfo := TvgrSheetPointInfo.Create(Place, Cell, RangeIndex, DownRect);
end;

procedure TvgrWBSheet.DragDrop(Source: TObject; X, Y: Integer);
begin
  WorkbookGrid.DoDragDrop(Self, Source, X, Y);
end;

procedure TvgrWBSheet.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  WorkbookGrid.DoDragOver(Self, Source, X, Y, State, Accept);
end;

/////////////////////////////////////////////////
//
// TvgrWBSheetParams
//
/////////////////////////////////////////////////
constructor TvgrWBSheetParams.Create;
begin
  inherited;
  SelectedRegion := TvgrWBSelectedRegion.Create;
end;

destructor TvgrWBSheetParams.Destroy;
begin
  SelectedRegion.Free;
  inherited;
end;

/////////////////////////////////////////////////
//
// TvgrWBSheetParamsList
//
/////////////////////////////////////////////////
function TvgrWBSheetParamsList.GetItem(Index: Integer): TvgrWBSheetParams;
begin
  Result := TvgrWBSheetParams(inherited Items[Index]);
end;

function TvgrWBSheetParamsList.Add(AWorksheet: TvgrWorksheet): TvgrWBSheetParams;
begin
  Result := TvgrWBSheetParams.Create;
  Result.Worksheet := AWorksheet;
  inherited Add(Result);
end;

function TvgrWBSheetParamsList.FindByWorksheet(AWorksheet: TvgrWorksheet): TvgrWBSheetParams;
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
// TvgrOptionsHeader
//
/////////////////////////////////////////////////
constructor TvgrOptionsHeader.Create;
begin
  inherited;
  FVisible := True;
  FBackColor := clBtnFace;
  FHighlightedBackColor := GetHighlightColor(FBackColor);
  FGridColor := clBtnShadow;
  FFont := TFont.Create;
  with FFont do
  begin
    Name := sDefFont;
    Size := cHeaderFontSize;
    Color := clBtnText;
    Charset := DEFAULT_CHARSET;
  end;
  FFont.OnChange := OnFontChange;
end;

destructor TvgrOptionsHeader.Destroy;
begin
  FFont.Free;
  inherited;
end;

procedure TvgrOptionsHeader.SetVisible(Value: Boolean);
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsHeader.SetBackColor(Value: TColor);
begin
  if FBackColor <> Value then
  begin
    if not IsStoredHighlightBackColor then
      FHighlightedBackColor := GetHighlightColor(Value);
    FBackColor := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsHeader.SetHighlightedBackColor(Value: TColor);
begin
  if FHighlightedBackColor <> Value then
  begin
    FHighlightedBackColor := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsHeader.SetGridColor(Value: TColor);
begin
  if FGridColor <> Value then
  begin
    FGridColor := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsHeader.SetFont(Value: TFont);
begin
  FFont.Assign(Value);
end;

function TvgrOptionsHeader.IsStoredHighlightBackColor: Boolean;
begin
  Result := FHighlightedBackColor <> GetHighlightColor(BackColor);
end;

function TvgrOptionsHeader.IsStoredFont: Boolean;
begin
  Result := (FFont.Name <> sDefFont) or
            (FFont.Size <> cHeaderFontSize) or
            (FFont.Style <> []) or
            (FFont.Color <> clBtnText) or
            (FFont.Charset <> DEFAULT_CHARSET);
end;

procedure TvgrOptionsHeader.OnFontChange(Sender: TObject);
begin
  DoChange;
end;

procedure TvgrOptionsHeader.Assign(Source: TPersistent);
begin
  with TvgrOptionsHeader(Source) do
  begin
    Self.FVisible := Visible;
    Self.FBackColor := BackColor;
    Self.FHighlightedBackColor := HighlightedBackColor;
    Self.FGridColor := GridColor;
    Self.FFont.Assign(Font);
  end;
  DoChange;
end;

/////////////////////////////////////////////////
//
// TvgrOptionsTopLeftButton
//
/////////////////////////////////////////////////
constructor TvgrOptionsTopLeftButton.Create;
begin
  inherited Create;
  FBackColor := clBtnFace;
  FBorderColor := clBtnShadow;
end;

procedure TvgrOptionsTopLeftButton.SetBackColor(Value: TColor);
begin
  if FBackColor <> Value then
  begin
    FBackColor := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsTopLeftButton.SetBorderColor(Value: TColor);
begin
  if FBorderColor <> Value then
  begin
    FBorderColor := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsTopLeftButton.Assign(Source: TPersistent);
begin
  with TvgrOptionsTopLeftButton(Source) do
  begin
    Self.FBackColor := BackColor;
    Self.FBorderColor := BorderColor;
  end;
  DoChange;
end;

/////////////////////////////////////////////////
//
// TvgrOptionsFormulaPanel
//
/////////////////////////////////////////////////
constructor TvgrOptionsFormulaPanel.Create;
begin
  inherited;
  FVisible := True;
  FBackColor := clBtnFace;
end;

procedure TvgrOptionsFormulaPanel.SetVisible(Value: Boolean);
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsFormulaPanel.SetBackColor(Value: TColor);
begin
  if FBackColor <> Value then
  begin
    FBackColor := Value;
    DoChange;
  end;
end;

procedure TvgrOptionsFormulaPanel.Assign(Source: TPersistent);
begin
  if Source is TvgrOptionsFormulaPanel then
  begin
    with TvgrOptionsFormulaPanel(Source) do
    begin
      Self.FVisible := Visible;
      Self.FBackColor := BackColor;
    end;
    DoChange;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrWorkbookGrid
//
/////////////////////////////////////////////////
constructor TvgrWorkbookGrid.Create(AOwner : TComponent);
begin
  inherited;

  FWorksheetDimensionsColor := clRed;
  FShowWorksheetDimensions := True;

  FBorderStyle := bsSingle;
  FSheetsCaptionWidth := cDefaultSheetsCaptionWidth;
  FSheetParamsList := TvgrWBSheetParamsList.Create;
  FActiveWorksheetIndex := -1;
  FBackgroundColor := clWindow;
  FGridColor := clActiveBorder;
  FSelectionBorderColor := clBlack;

  FOptionsCols := TvgrOptionsHeader.Create;
  FOptionsCols.OnChange := OnOptionsChanged;

  FOptionsRows := TvgrOptionsHeader.Create;
  FOptionsRows.OnChange := OnOptionsChanged;

  FOptionsVertScrollBar := TvgrOptionsScrollBar.Create;
  FOptionsVertScrollBar.OnChange := OnOptionsChanged;

  FOptionsHorzScrollBar := TvgrOptionsScrollBar.Create;
  FOptionsHorzScrollBar.OnChange := OnOptionsChanged;

  FOptionsTopLeftButton := TvgrOptionsTopLeftButton.Create;
  FOptionsTopLeftButton.OnChange := OnOptionsChanged;

  FOptionsFormulaPanel := TvgrOptionsFormulaPanel.Create;
  FOptionsFormulaPanel.OnChange := OnOptionsChanged;

  FSheet := TvgrWBSheet.Create(Self);
  FSheet.Parent := Self;

  FHorzScrollBar := TvgrScrollBar.Create(Self);
  FHorzScrollBar.Kind := sbHorizontal;
  FHorzScrollBar.OnScrollMessage := OnScrollMessage;
  FHorzScrollBar.Parent := Self;

  FVertScrollBar := TvgrScrollBar.Create(Self);
  FVertScrollBar.Kind := sbVertical;
  FVertScrollBar.OnScrollMessage := OnScrollMessage;
  FVertScrollBar.Parent := Self;

  FRowsPanel := TvgrWBRowsPanel.Create(Self);
  FRowsPanel.Visible := FOptionsRows.Visible;
  FRowsPanel.Parent := Self;

  FColsPanel := TvgrWBColsPanel.Create(Self);
  FColsPanel.Parent := Self;

  FSheetsPanel := TvgrSheetsPanel.Create(Self);
  FSheetsPanel.Parent := Self;
  
  FSheetsPanelResizer := TvgrSheetsPanelResizer.Create(Self);
  FSheetsPanelResizer.Parent := Self;

  FFormulaPanel := TvgrWBFormulaPanel.Create(Self);
  FFormulaPanel.OptionsChanged;
  FFormulaPanel.Parent := Self;

  FTopLeftButton := TvgrWBTopLeftButton.Create(Self);
  FTopLeftButton.Parent := Self;
  FTopLeftSectionsButton := TvgrWBTopLeftSectionsButton.Create(Self);
  FTopLeftSectionsButton.Parent := Self;

  FHorzSections := TvgrWBHorzSectionsPanel.Create(Self);
  FHorzSections.Parent := Self;
  FVertSections := TvgrWBVertSectionsPanel.Create(Self);
  FVertSections.Parent := Self;

  FSizer := TvgrSizer.Create(Self);
  FSizer.Parent := Self;
  
  FInplaceEdit := TvgrWBInplaceEdit.Create(Self);
  FInplaceEdit.Visible := False;
  FInplaceEdit.Parent := FSheet;

  FKeyboardFilter := TvgrKeyboardFilter.Create;
  FKeyboardFilter.OnKeyPress := FSheet.KbdKeyPress;
  FKeyboardFilter.OnBeforeSelectionCommand := FSheet.KbdBeforeSelectionCommand;
  FKeyboardFilter.OnAfterSelectionCommand := FSheet.KbdAfterSelectionCommand;
  KbdAddCommands;

  FPopupMenu := TPopupMenu.Create(Self);
  FDefaultPopupMenu := True;
end;

destructor TvgrWorkbookGrid.Destroy;
begin
  Workbook := nil;
  FreeAndNil(FPopupMenu);
  FSheetParamsList.Free;
  FKeyboardFilter.Free;
  FOptionsCols.Free;
  FOptionsRows.Free;
  FOptionsTopLeftButton.Free;
  FOptionsVertScrollBar.Free;
  FOptionsHorzScrollBar.Free;
  FOptionsFormulaPanel.Free;
  inherited;
end;

procedure TvgrWorkbookGrid.ReadDefaultColWidth(Reader: TReader);
begin
  Reader.ReadInteger;
end;

procedure TvgrWorkbookGrid.WriteDefaultColWidth(Writer: TWriter);
begin
end;

procedure TvgrWorkbookGrid.ReadDefaultRowHeight(Reader: TReader);
begin
  Reader.ReadInteger;
end;

procedure TvgrWorkbookGrid.WriteDefaultRowHeight(Writer: TWriter);
begin
end;

procedure TvgrWorkbookGrid.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('DefaultColWidth', ReadDefaultColWidth, WriteDefaultColWidth, False);
  Filer.DefineProperty('DefaultRowHeight', ReadDefaultRowHeight, WriteDefaultRowHeight, False);
end;

procedure TvgrWorkbookGrid.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if AComponent = FWorkbook then
      Workbook := nil;
  end;
end;

procedure TvgrWorkbookGrid.Loaded;
begin
  inherited;
end;

procedure TvgrWorkbookGrid.SetShowWorksheetDimensions(Value: Boolean);
begin
  if FShowWorksheetDimensions <> Value then
  begin
    FShowWorksheetDimensions := Value;
    FSheet.Repaint;
  end;
end;

procedure TvgrWorkbookGrid.SetWorksheetDimensionsColor(Value: TColor);
begin
  if FWorksheetDimensionsColor <> Value then
  begin
    FWorksheetDimensionsColor := Value;
    FSheet.Repaint;
  end;
end;

procedure TvgrWorkbookGrid.DrawBorder(Canvas: TCanvas; var AInternalRect: TRect; var AWidth, AHeight: Integer);
begin
  vgr_GUIFunctions.DrawBorder(Self, Canvas, FBorderStyle, AInternalRect, AWidth, AHeight);
end;

procedure TvgrWorkbookGrid.Paint;
var
  ARect: TRect;
  AWidth, AHeight: Integer;
begin
  DrawBorder(Canvas, ARect, AWidth, AHeight);
end;

procedure TvgrWorkbookGrid.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrWorkbookGrid.WMSetFocus(var Msg: TWMSetFocus);
begin
  Windows.SetFocus(FSheet.Handle);
end;

procedure TvgrWorkbookGrid.KbdAddCommands;
begin
  with KeyboardFilter do
  begin
    AddSelectionCommand(VK_LEFT, [], FSheet.KbdLeft);
    AddSelectionCommand(VK_LEFT, [ssCtrl], FSheet.KbdLeftCell);
    AddSelectionCommand(VK_RIGHT, [], FSheet.KbdRight);
    AddSelectionCommand(VK_RIGHT, [ssCtrl], FSheet.KbdRightCell);
    AddSelectionCommand(VK_UP, [], FSheet.KbdUp);
    AddSelectionCommand(VK_UP, [ssCtrl], FSheet.KbdTopCell);
    AddSelectionCommand(VK_DOWN, [], FSheet.KbdDown);
    AddSelectionCommand(VK_DOWN, [ssCtrl], FSheet.KbdBottomCell);
    AddSelectionCommand(VK_PRIOR, [], FSheet.KbdPageUp);
    AddSelectionCommand(VK_NEXT, [], FSheet.KbdPageDown);
    AddSelectionCommand(VK_HOME, [], FSheet.KbdLeftCell);
    AddSelectionCommand(VK_HOME, [ssCtrl], FSheet.KbdLeftTopCell);
    AddSelectionCommand(VK_END, [], FSheet.KbdRightCell);
    AddSelectionCommand(VK_END, [ssCtrl], FSheet.KbdRightBottomCell);

    AddCommand(VK_F2, [], FSheet.KbdInplaceEdit);
    AddCommand(VK_DELETE, [], FSheet.KbdClearValue);
    AddCommand(Word(49), [ssCtrl], FSheet.KbdCellProperties);

    AddCommand(Word('R'), [ssShift, ssCtrl], FSheet.KbdSelectedRowsAutoHeight);
    AddCommand(Word('C'), [ssShift, ssCtrl], FSheet.KbdSelectedColumnsAutoHeight);

    AddCommand(Word('C'), [ssCtrl], FSheet.KbdCopy);
    AddCommand(VK_INSERT, [ssCtrl], FSheet.KbdCopy);
    AddCommand(Word('X'), [ssCtrl], FSheet.KbdCut);
    AddCommand(VK_DELETE, [ssShift], FSheet.KbdCut);
    AddCommand(Word('V'), [ssCtrl], FSheet.KbdPaste);
    AddCommand(VK_INSERT, [ssShift], FSheet.KbdPaste);
  end;
end;

procedure TvgrWorkbookGrid.DoSectionDblClick(ASections: TvgrSections; ASectionIndex: Integer);
begin
  if Assigned(FOnSectionDblClick) then
    FOnSectionDblClick(Self, ASections, ASectionIndex);
end;

procedure TvgrWorkbookGrid.UpdateScrollBar(ScrollBar : TvgrScrollBar; NewPos,VisibleCells,FullVisibleCells,MaxCells : integer);
var
  RangeMax : integer;
begin
  if VisibleCells = FullVisibleCells then
    Inc(VisibleCells);
  if (MaxCells <= NewPos + VisibleCells - 1) and (NewPos > 0) then
    RangeMax := NewPos + VisibleCells - 2
  else
    RangeMax := Max(MaxCells, NewPos + VisibleCells - 1);
  with ScrollBar do
    SetParams(NewPos, 0, RangeMax, FullVisibleCells);
end;

procedure TvgrWorkbookGrid.OnScrollMessage(Sender: TObject; var Msg : TWMScroll);
var
  OldPos: integer;
  ScrollInfo: TScrollInfo;
begin
  if ActiveWorksheet = nil then exit;
  ScrollInfo.cbSize := sizeof(ScrollInfo);
  ScrollInfo.fMask := SIF_ALL;
  GetScrollInfo(TvgrScrollBar(Sender).Handle,SB_CTL,ScrollInfo);
  OldPos := ScrollInfo.nPos;
  case TScrollCode(Msg.ScrollCode) of
    scLineUp:
      Dec(ScrollInfo.nPos);
    scLineDown:
      Inc(ScrollInfo.nPos);
    scPageUp:
      Dec(ScrollInfo.nPos,ScrollInfo.nPage);
    scPageDown:
      Inc(ScrollInfo.nPos,ScrollInfo.nPage);
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
    DoScroll(Point(OldPos, TopRow),Point(ScrollInfo.nPos, TopRow))
  else
    DoScroll(Point(LeftCol, OldPos),Point(LeftCol, ScrollInfo.nPos));

  Msg.Result := 1;
end;

procedure TvgrWorkbookGrid.DoScroll(const OldPos,NewPos: TPoint);
begin
  CalcVisibleRegion(NewPos.X, NewPos.Y);
  if OldPos.X <> NewPos.X then
  begin
    UpdateScrollBar(FHorzScrollBar, NewPos.X, FVisibleColsCount, FFullVisibleColsCount, ActiveWorksheet.Dimensions.Right);
    FColsPanel.DoScroll(OldPos.X, NewPos.X);
    FVertSections.DoScroll(OldPos.X, NewPos.X);
  end;
  if OldPos.Y <> NewPos.Y then
  begin
    UpdateScrollBar(FVertScrollBar, NewPos.Y, FVisibleRowsCount, FFullVisibleRowsCount, ActiveWorksheet.Dimensions.Bottom);
    FRowsPanel.DoScroll(OldPos.Y, NewPos.Y);
    FHorzSections.DoScroll(OldPos.Y, NewPos.Y);
  end;
  FSheet.DoScroll(OldPos,NewPos);
end;

procedure TvgrWorkbookGrid.DoScrollBy(AScrollX, AScrollY: Integer; AStartTimer: Boolean);
var
  AOldPos, ANewPos: TPoint;
begin
  if (AScrollX <> 0) or (AScrollY <> 0) then
  begin
    AOldPos := Point(LeftCol, TopRow);
    ANewPos := Point(AOldPos.X + AScrollX, AOldPos.Y + AScrollY);
    DoScroll(AOldPos, ANewPos);
  end;
end;

procedure TvgrWorkbookGrid.DoColWidthChanged(ColNumber : integer);
begin
end;

procedure TvgrWorkbookGrid.DoRowHeightChanged(RowNumber : integer);
begin
end;

procedure TvgrWorkbookGrid.GetPaintRects(PaintRect: TRect; var rRegions,rPixels: TRect; ALeftCol, ATopRow: Integer);
var
  j: integer;
begin
  OffsetRect(PaintRect, FLeftTopPixelsOffset.x, FLeftTopPixelsOffset.y);
  rRegions.Left := ALeftCol;
  rPixels.Left := FLeftTopPixelsOffset.x;
  j := 0;
  while rPixels.Left<PaintRect.Left do
    begin
      j := FSheet.ColPixelWidth(rRegions.Left);
      rPixels.Left := rPixels.Left+j;
      Inc(rRegions.Left);
    end;
  rPixels.Right := rPixels.Left;
  rRegions.Right := rRegions.Left;
  if rPixels.Left>PaintRect.Left then
    begin
      Dec(rRegions.Left);
      rPixels.Left := rPixels.Left-j;
    end;

  while rPixels.Right<PaintRect.Right do
    begin
      j := FSheet.ColPixelWidth(rRegions.Right);
      rPixels.Right := rPixels.Right+j;
      Inc(rRegions.Right);
    end;
  if rPixels.Right>=PaintRect.Right then
    Dec(rRegions.Right);

  rRegions.Top := ATopRow;
  rPixels.Top := FLeftTopPixelsOffset.y;
  while rPixels.Top < PaintRect.Top do
    begin
      j := FSheet.RowPixelHeight(rRegions.Top);
      rPixels.Top := rPixels.Top+j;
      Inc(rRegions.Top);
    end;
  rPixels.Bottom := rPixels.Top;
  rRegions.Bottom := rRegions.Top;
  if rPixels.Top>PaintRect.Top then
    begin
      Dec(rRegions.Top);
      rPixels.Top := rPixels.Top-j;
    end;
  
  while rPixels.Bottom<PaintRect.Bottom do
    begin
      j := FSheet.RowPixelHeight(rRegions.Bottom);
      rPixels.Bottom := rPixels.Bottom+j;
      Inc(rRegions.Bottom);
    end;
  if rPixels.Bottom>=PaintRect.Bottom then
    Dec(rRegions.Bottom);
  OffsetRect(rPixels,-FLeftTopPixelsOffset.x,-FLeftTopPixelsOffset.y);
end;

procedure TvgrWorkbookGrid.CalcVisibleRegion(ALeftCol, ATopRow: Integer);
var
  rRegions,rPixels: TRect;
begin
  FSheet.UpdateHorzPoints;
  if ActiveWorksheet = nil then
  begin
    FVisibleColsCount := 0;
    FFullVisibleColsCount := 0;
    FVisibleRowsCount := 0;
    FFullVisibleRowsCount := 0;
    FLeftTopPixelsOffset := Point(0, 0);
  end
  else
  begin
    FLeftTopPixelsOffset.X := FSheet.ColPixelLeft(ALeftCol);
    FLeftTopPixelsOffset.Y := FSheet.RowPixelTop(ATopRow);
    GetPaintRects(FSheet.ClientRect, rRegions, rPixels, ALeftCol, ATopRow);
    FVisibleColsCount := rRegions.Right-rRegions.Left+1;
    FFullVisibleColsCount := FVisibleColsCount;
    if rPixels.Right > FSheet.ClientRect.Right then
      Dec(FFullVisibleColsCount);
    FVisibleRowsCount := rRegions.Bottom-rRegions.Top+1;
    FFullVisibleRowsCount := FVisibleRowsCount;
    if rPixels.Bottom > FSheet.ClientRect.Bottom then
      Dec(FFullVisibleRowsCount);
  end;
end;

procedure TvgrWorkbookGrid.AlignSheetCaptions;
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

  DrawBorder(Canvas, ARect, AWidth, AHeight);

  ATop := FSheet.Top + FSheet.Height;
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

procedure TvgrWorkbookGrid.AlignControls(AControl: TControl; var AAlignRect: TRect);
var
  ARect: TRect;
  ARowsWidth, AFormulaHeight, AColsHeight, AHorzScrollBarHeight, AVertScrollBarWidth, AWidth, AHeight: Integer;
begin
  if not FDisableAlign then
  begin
    if OptionsRows.Visible then
      ARowsWidth := FRowsPanel.Width
    else
      ARowsWidth := 0;
    if OptionsCols.Visible then
      AColsHeight := FColsPanel.Height
    else
      AColsHeight := 0;
    if OptionsHorzScrollBar.Visible then
      AHorzScrollBarHeight := FHorzScrollBar.Height
    else
      AHorzScrollBarHeight := 0;
    if OptionsVertScrollBar.Visible then
      AVertScrollBarWidth := FVertScrollBar.Width
    else
      AVertScrollBarWidth := 0;
    if OptionsFormulaPanel.Visible then
      AFormulaHeight := FFormulaPanel.Height
    else
      AFormulaHeight := 0;

    DrawBorder(Canvas, ARect, AWidth, AHeight);

    FSheet.SetBounds(ARect.Left + ARowsWidth + FHorzSections.Width,
                     ARect.Top + AFormulaHeight + AColsHeight + FVertSections.Height,
                     AWidth - ARowsWidth - AVertScrollBarWidth - FHorzSections.Width,
                     AHeight - AFormulaHeight - AColsHeight - AHorzScrollBarHeight - FVertSections.Height);
    FFormulaPanel.SetBounds(ARect.Left, ARect.Top, AWidth, FFormulaPanel.Height);
    FVertSections.SetBounds(ARect.Left + FHorzSections.Width,
                            ARect.Top + AFormulaHeight,
                            AWidth - FHorzSections.Width - AVertScrollBarWidth,
                            FVertSections.Height);
    FColsPanel.SetBounds(ARect.Left + ARowsWidth + FHorzSections.Width,
                         ARect.Top + AFormulaHeight + FVertSections.Height,
                         AWidth - FHorzSections.Width - ARowsWidth - AVertScrollBarWidth,
                         FColsPanel.Height);
    FHorzSections.SetBounds(ARect.Left,
                            ARect.Top + AFormulaHeight + FVertSections.Height,
                            FHorzSections.Width,
                            AHeight - AFormulaHeight - AHorzScrollBarHeight - FVertSections.Height);
    FRowsPanel.SetBounds(ARect.Left + FHorzSections.Width,
                         ARect.Top + AFormulaHeight + AColsHeight + FVertSections.Height,
                         FRowsPanel.Width,
                         AHeight - AFormulaHeight - AColsHeight - AHorzScrollBarHeight - FVertSections.Height);
    // !!! MINIMUM Size of scrollbar - 2 pixel                      
    if OptionsVertScrollBar.Visible then
      FVertScrollBar.SetBounds(FSheet.Left + FSheet.Width,
                               ARect.Top + AFormulaHeight,
                               FVertScrollBar.Width,
                               FSheet.Height + AColsHeight + FVertSections.Height)
    else
      FVertScrollBar.SetBounds(Width + 1,
                               ARect.Top + AFormulaHeight,
                               FVertScrollBar.Width,
                               FSheet.Height + AColsHeight + FVertSections.Height);

    FTopLeftButton.SetBounds(ARect.Left + FHorzSections.Width,
                             ARect.Top + AFormulaHeight + FVertSections.Height,
                             ARowsWidth,
                             AColsHeight);
    FTopLeftSectionsButton.SetBounds(ARect.Left,
                                     ARect.Top + AFormulaHeight,
                                     FHorzSections.Width,
                                     FVertSections.Height);
    AlignSheetCaptions;
    FSizer.SetBounds(FVertScrollBar.Left,
                     FVertScrollBar.BoundsRect.Bottom,
                     AVertScrollBarWidth,
                     AHorzScrollBarHeight);

    CalcVisibleRegion(LeftCol, TopRow);
    UpdateScrollBar(HorzScrollBar, LeftCol, VisibleColsCount, FullVisibleColsCount, DimensionsRight);
    UpdateScrollBar(VertScrollBar, TopRow, VisibleRowsCount, FullVisibleRowsCount, DimensionsBottom);

    Invalidate;
  end;
end;

procedure TvgrWorkbookGrid.SetReadOnly(Value: Boolean);
begin
  if FReadOnly <> Value then
  begin
    FReadOnly := Value;
    FFormulaPanel.Edit.ReadOnly := Value;
    FSheet.InplaceEdit.ReadOnly := Value; 
  end;
end;

procedure TvgrWorkbookGrid.SetBorderStyle(Value: TBorderStyle);
begin
  if FBorderStyle <> Value then
  begin
    FBorderStyle := Value;
    Realign;
  end;
end;

procedure TvgrWorkbookGrid.SetOptionsCols(Value: TvgrOptionsHeader);
begin
  FOptionsCols.Assign(Value);
end;

procedure TvgrWorkbookGrid.SetOptionsRows(Value: TvgrOptionsHeader);
begin
  FOptionsRows.Assign(Value);
end;

procedure TvgrWorkbookGrid.SetOptionsVertScrollBar(Value: TvgrOptionsScrollBar);
begin
  FOptionsVertScrollBar.Assign(Value);
end;

procedure TvgrWorkbookGrid.SetOptionsHorzScrollBar(Value: TvgrOptionsScrollBar);
begin
  FOptionsHorzScrollBar.Assign(Value);
end;

procedure TvgrWorkbookGrid.SetOptionsTopLeftButton(Value: TvgrOptionsTopLeftButton);
begin
  FOptionsTopLeftButton.Assign(Value);
end;

procedure TvgrWorkbookGrid.SetOptionsFormulaPanel(Value: TvgrOptionsFormulaPanel);
begin
  FOptionsFormulaPanel.Assign(Value);
end;

procedure TvgrWorkbookGrid.SetWorkbook(Value: TvgrWorkbook);
begin
  if Value <> FWorkbook then
  begin
    FSelectionFormat := nil;
    if FWorkbook <> nil then
      FWorkbook.DisconnectHandler(Self);
    FWorkbook := Value;
    if FWorkbook <> nil then
      Value.ConnectHandler(Self);
    FSheetParamsList.Clear;

    if (FWorkbook <> nil) and (FWorkbook.WorksheetsCount > 0) then
      FActiveWorksheetIndex := 0
    else
      FActiveWorksheetIndex := -1;
      
    InternalUpdateActiveWorksheet;
    DoActiveWorksheetChanged;

    FSheetsPanel.UpdateSheets;
    FFormulaPanel.UpdateEnabled;
  end;
end;

procedure TvgrWorkbookGrid.BeforeChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcDeleteRange, vgrwcDeleteBorder:
      FSheet.BeforeChangeWorkbook(ChangeInfo);
    vgrwcChangeBorder:
      FSheet.BeforeChangeWorkbook(ChangeInfo);
    vgrwcDeleteWorksheet:
      FDeleteActiveWorksheet := (ActiveWorksheet <> nil) and
                                (ActiveWorksheet.IndexInWorkbook = TvgrWorksheet(ChangeInfo.ChangedObject).IndexInWorkbook);
  end;
end;

procedure TvgrWorkbookGrid.AfterChangeWorkbook(ChangeInfo : TvgrWorkbookChangeInfo);
begin
  case ChangeInfo.ChangesType of
    vgrwcChangeCol, vgrwcDeleteCol:
      begin
        CalcVisibleRegion(LeftCol, TopRow);
        UpdateScrollBar(HorzScrollBar,
                        LeftCol,
                        VisibleColsCount,
                        FullVisibleColsCount,
                        DimensionsRight);
        FSheet.AfterChangeWorkbook(ChangeInfo);
        FColsPanel.AfterChangeWorkbook(ChangeInfo);
        FVertSections.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcChangeRow, vgrwcDeleteRow:
      begin
        CalcVisibleRegion(LeftCol, TopRow);
        UpdateScrollBar(VertScrollBar,
                        TopRow,
                        VisibleRowsCount,
                        FullVisibleRowsCount,
                        DimensionsBottom);
        FSheet.AfterChangeWorkbook(ChangeInfo);
        FRowsPanel.AfterChangeWorkbook(ChangeInfo);
        FHorzSections.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcNewRange, vgrwcDeleteRange, vgrwcNewBorder, vgrwcDeleteBorder, vgrwcChangeRange, vgrwcChangeBorder:
      begin
        if ChangeInfo.ChangesType in [vgrwcNewRange, vgrwcDeleteRange, vgrwcChangeRange] then
          CheckSelectedTextChanged;
        if ChangeInfo.ChangesType in [vgrwcNewRange, vgrwcDeleteRange] then
        begin
          UpdateScrollBar(HorzScrollBar, LeftCol, VisibleColsCount, FullVisibleColsCount, DimensionsRight);
          UpdateScrollBar(VertScrollBar, TopRow, VisibleRowsCount, FullVisibleRowsCount, DimensionsBottom);
        end;
        FSheet.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcChangeHorzSection, vgrwcChangeVertSection, vgrwcNewHorzSection, vgrwcDeleteHorzSection, vgrwcNewVertSection, vgrwcDeleteVertSection:
      begin
        FHorzSections.AfterChangeWorkbook(ChangeInfo);
        FVertSections.AfterChangeWorkbook(ChangeInfo);
        FSheet.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcChangeWorksheetContent:
      begin
        CalcVisibleRegion(LeftCol, TopRow);
        UpdateScrollBar(HorzScrollBar, LeftCol, VisibleColsCount, FullVisibleColsCount, DimensionsRight);
        UpdateScrollBar(VertScrollBar, TopRow, VisibleRowsCount, FullVisibleRowsCount, DimensionsBottom);

        FSheet.AfterChangeWorkbook(ChangeInfo);
        FHorzSections.AfterChangeWorkbook(ChangeInfo);
        FVertSections.AfterChangeWorkbook(ChangeInfo);
        FRowsPanel.AfterChangeWorkbook(ChangeInfo);
        FColsPanel.AfterChangeWorkbook(ChangeInfo);
        CheckSelectedTextChanged;
      end;
    vgrwcNewWorksheet, vgrwcChangeWorksheet:
      begin
        FSheetsPanel.AfterChangeWorkbook(ChangeInfo);
        if (FActiveWorksheetIndex >= 0) and (FActiveWorksheet <> Workbook.Worksheets[FActiveWorksheetIndex]) then
        begin
          SaveWBSheetParams;
          InternalUpdateActiveWorksheet;
          DoActiveWorksheetChanged;
        end;
        FSheet.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcDeleteWorksheet:
      begin
        if FDeleteActiveWorksheet then
          InternalUpdateActiveWorksheet;
        FSheetsPanel.AfterChangeWorkbook(ChangeInfo);
      end;
    vgrwcUpdateAll, vgrwcChangeWorksheetPageProperties:
      begin
        SaveWBSheetParams;
        InternalUpdateActiveWorksheet;
      end;
  end;
end;

procedure TvgrWorkbookGrid.DeleteItemInterface(AItem: IvgrWBListItem);
begin
end;

function TvgrWorkbookGrid.GetScrollBarOptions(AScrollBar: TvgrScrollBar): TvgrOptionsScrollBar;
begin
  if AScrollBar = HorzScrollBar then
    Result := OptionsHorzScrollBar
  else
    Result := OptionsVertScrollBar;
end;

function TvgrWorkbookGrid.SheetsPanelParentGetWorkbook: TvgrWorkbook;
begin
  Result := Workbook;
end;

function TvgrWorkbookGrid.SheetsPanelParentGetActiveWorksheet: TvgrWorksheet;
begin
  Result := ActiveWorksheet;
end;

function TvgrWorkbookGrid.SheetsPanelParentGetActiveWorksheetIndex: Integer;
begin
  Result := ActiveWorksheetIndex;
end;

procedure TvgrWorkbookGrid.SheetsPanelParentSetActiveWorksheetIndex(Value: Integer);
begin
  if not IsInplaceEdit then
    ActiveWorksheetIndex := Value;
end;

procedure TvgrWorkbookGrid.SheetsPanelResizerParentDoResize(AOffset: Integer);
begin
  SheetsCaptionWidth := SheetsCaptionWidth + AOffset;
end;

function TvgrWorkbookGrid.GetDefaultColWidth: Integer;
begin
  if ActiveWorksheet <> nil then
    Result := ActiveWorksheet.PageProperties.Defaults.ColWidth
  else
    Result := cDefaultColWidthTwips;
end;

function TvgrWorkbookGrid.GetDefaultRowHeight: Integer;
begin
  if ActiveWorksheet <> nil then
    Result := ActiveWorksheet.PageProperties.Defaults.RowHeight
  else
    Result := cDefaultRowHeightTwips;
end;

procedure TvgrWorkbookGrid.SaveWBSheetParams;
var
  AParams: TvgrWBSheetParams;
begin
  if ActiveWorksheet = nil then exit;
  AParams := FSheetParamsList.FindByWorksheet(ActiveWorksheet);
  if AParams = nil then
    AParams := FSheetParamsList.Add(ActiveWorksheet);

  AParams.LeftTop.X := LeftCol;
  AParams.LeftTop.Y := TopRow;
  AParams.SelectedRegion.Assign(FSheet.FSelectedRegion);
  AParams.SelStartRect := FSheet.FSelStartRect;
  AParams.CurRect := FSheet.FCurRect;
end;

procedure TvgrWorkbookGrid.RestoreWBSheetParams;
var
  AParams: TvgrWBSheetParams;
  ARect: TRect;
begin
  if ActiveWorksheet = nil then exit;
  AParams := FSheetParamsList.FindByWorksheet(ActiveWorksheet);

  if AParams = nil then
  begin
    if HandleAllocated then
    begin
      VertScrollBar.Position := 0;
      HorzScrollBar.Position := 0;
    end;
    FSheet.FSelectedRegion.SetSelection(Rect(0, 0, 0, 0));
    FSheet.FSelStartRect := Rect(0, 0, 0, 0);
    FSheet.FCurRect := Rect(0, 0, 0, 0);
  end
  else
  begin
    if HandleAllocated then
    begin
      VertScrollBar.SetParams(AParams.LeftTop.Y, AParams.LeftTop.Y, AParams.LeftTop.Y, 1);
      HorzScrollBar.SetParams(AParams.LeftTop.X, AParams.LeftTop.X, AParams.LeftTop.X, 1);
    end;
    FSheet.FSelectedRegion.Assign(AParams.SelectedRegion);
    FSheet.FSelStartRect := AParams.SelStartRect;
    FSheet.CheckSelStartRect;
    FSheet.FCurRect := AParams.CurRect;
  end;
  FSheet.CheckSelections(ARect);

  if HandleAllocated then
  begin
    CalcVisibleRegion(LeftCol, TopRow);

    UpdateScrollBar(HorzScrollBar, LeftCol, VisibleColsCount, FullVisibleColsCount, DimensionsRight);
    UpdateScrollBar(VertScrollBar, TopRow, VisibleRowsCount, FullVisibleRowsCount, DimensionsBottom);
    Repaint;
  end;
end;

procedure TvgrWorkbookGrid.SetActiveWorksheetIndex(Value : integer);
begin
  if FActiveWorksheetIndex = Value then exit;

  SaveWBSheetParams;

  FActiveWorksheetIndex := Value;

  InternalUpdateActiveWorksheet;

  DoActiveWorksheetChanged;
end;

procedure TvgrWorkbookGrid.InternalUpdateActiveWorksheet;
begin
  if Workbook = nil then
  begin
    FActiveWorksheetIndex := -1;
    FActiveWorksheet := nil
  end
  else
  begin
    FActiveWorksheetIndex := Min(Max(0, FActiveWorksheetIndex), Workbook.WorksheetsCount - 1);
    if FActiveWorksheetIndex = -1 then
      FActiveWorksheet := nil
    else
      FActiveWorksheet := Workbook.Worksheets[FActiveWorksheetIndex];
  end;

  FVertSections.UpdateSize;
  FHorzSections.UpdateSize;
  FSheetsPanel.UpdateSheets;

  if FActiveWorksheetIndex = -1 then
  begin
    FVisibleRowsCount := 0;
    FFullVisibleRowsCount := 0;
    FVisibleColsCount := 0;
    FFullVisibleColsCount := 0;
    FLeftTopPixelsOffset := Point(0, 0);
    FSheet.FSelectedRegion.SetSelection(Rect(0, 0, 0, 0));
    FSheet.FCurRect := Rect(0, 0, 0, 0);
    FSheet.FSelStartRect := Rect(0, 0, 0, 0);
    FFormulaPanel.Enabled := False;
    CalcVisibleRegion(0, 0);
    if HandleAllocated then
    begin
      HorzScrollBar.SetParams(0, 0, 0, 0);
      VertScrollBar.SetParams(0, 0, 0, 0);
      Repaint;
    end
  end
  else
  begin
    FFormulaPanel.Enabled := True;
    RestoreWBSheetParams;
  end;

  FFormulaPanel.UpdateEnabled;
  DoSelectionChange;
end;

procedure TvgrWorkbookGrid.SetBackgroundColor(Value : TColor);
begin
  if FBackgroundColor <> Value then
  begin
    FBackgroundColor := Value;
    FSheet.Repaint;
  end;
end;

procedure TvgrWorkbookGrid.SetGridColor(Value : TColor);
begin
  if FGridColor <> Value then
  begin
    FGridColor := Value;
    FSheet.Repaint;
  end;
end;

procedure TvgrWorkbookGrid.SetSelectionBorderColor(Value : TColor);
begin
  if FSelectionBorderColor <> Value then
  begin
    FSelectionBorderColor := Value;
    FSheet.Repaint;
  end;
end;

function TvgrWorkbookGrid.GetSelectionCount: Integer;
begin
  if ActiveWorksheet <> nil then
    Result := FSheet.FSelectedRegion.RectsCount
  else
    Result := -1;
end;

function TvgrWorkbookGrid.GetSelection(Index: Integer): TRect;
begin
  if ActiveWorksheet <> nil then
    Result := FSheet.FSelectedRegion.Rects[Index]^
  else
    Result := Rect(0, 0, 0, 0);
end;

function TvgrWorkbookGrid.GetSelectionRects: TvgrRectArray;
var
  I: Integer;
begin
  SetLength(Result, SelectionCount);
  if ActiveWorksheet <> nil then
    for I := 0 to High(Result) do
      Result[0] := Selections[I];
end;

function TvgrWorkbookGrid.GetIsInplaceEdit: Boolean;
begin
  Result := FSheet.IsInplaceEdit;
end;

procedure TvgrWorkbookGrid.SetSheetsCaptionWidth(Value: Integer);
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

function TvgrWorkbookGrid.GetScrollInterval: Integer;
begin
  Result := FSheet.FScrollTimer.Interval;
end;

procedure TvgrWorkbookGrid.SetScrollInterval(Value: Integer);
begin
  FSheet.FScrollTimer.Interval := Value;
  FColsPanel.FScrollTimer.Interval := Value;
  FRowsPanel.FScrollTimer.Interval := Value;
end;

function TvgrWorkbookGrid.GetLeftCol: Integer;
begin
  Result := HorzScrollbar.Position;
end;

function TvgrWorkbookGrid.GetTopRow: Integer;
begin
  Result := VertScrollbar.Position;
end;

function TvgrWorkbookGrid.GetDimensionsRight: Integer;
begin
  Result := FSheet.DimensionsRight;
end;

function TvgrWorkbookGrid.GetDimensionsBottom: Integer;
begin
  Result := FSheet.DimensionsBottom;
end;

function TvgrWorkbookGrid.GetLastSelection: TRect;
begin
  if SelectionCount = 0 then
    Result := Rect(-1, -1, -1, -1)
  else
    Result := Selections[SelectionCount - 1];
end;

procedure TvgrWorkbookGrid.OnOptionsChanged(Sender: TObject);
begin
  if Sender = OptionsCols then
    FColsPanel.OptionsChanged
  else
    if Sender = OptionsRows then
      FRowsPanel.OptionsChanged
    else
      if Sender = OptionsVertScrollBar then
        VertScrollBar.OptionsChanged
      else
        if Sender = OptionsHorzScrollBar then
          HorzScrollBar.OptionsChanged
        else
          if Sender = OptionsTopLeftButton then
          begin
            FTopLeftButton.OptionsChanged;
            FTopLeftSectionsButton.OptionsChanged;
          end
          else
            if Sender = OptionsFormulaPanel then
              FFormulaPanel.OptionsChanged;
end;

procedure TvgrWorkbookGrid.AssignInplaceEditValueToRange(const AValue: string; ARange: IvgrRange; ANewRange: Boolean);
begin
  DoSetInplaceEditValue(AValue, ARange, ANewRange);
end;

procedure TvgrWorkbookGrid.StartInplaceEdit;
begin
  FSheet.StartInplaceEdit;
end;

procedure TvgrWorkbookGrid.EndInplaceEdit(ACancel: Boolean);
begin
  FSheet.EndInplaceEdit(ACancel);
end;

procedure TvgrWorkbookGrid.SetSelection(const ASelectionRect: TRect);
begin
  FSheet.SetSelection(ASelectionRect);
end;

procedure TvgrWorkbookGrid.DoSelectionChange;
begin
  if not (csDestroying in ComponentState) then  
  begin
    if ActiveWorksheet = nil then
      FSelectionFormat := nil
    else
      FSelectionFormat := ActiveWorksheet.RangesFormat[SelectionRects, True];
    CheckSelectedTextChanged;
    if Assigned(FOnSelectionChange) then
      FOnSelectionChange(Self);
  end;
end;

function TvgrWorkbookGrid.GetSelectedText: string;
var
  ARange: IvgrRange;
begin
  Result := '';
  if ActiveWorksheet <> nil then
  begin
    with FSheet.FCurRect do
      ARange := ActiveWorksheet.RangesList.Find(Left, Top, Right, Bottom);
    if ARange <> nil then
      Result := ARange.StringValue;
  end;
end;

procedure TvgrWorkbookGrid.CheckSelectedTextChanged;
var
  S: string;
begin
  S := SelectedText;
  if FPriorSelectedText <> S then
  begin
    FPriorSelectedText := S;
    DoSelectedTextChanged;
  end;
end;

procedure TvgrWorkbookGrid.DoSelectedTextChanged;
begin
  FFormulaPanel.DoSelectedTextChanged;
  if Assigned(FOnSelectedTextChanged) then
    FOnSelectedTextChanged(Self);
end;

procedure TvgrWorkbookGrid.DoCellProperties;
begin
  if Assigned(FOnCellProperties) then
    FOnCellProperties(Self);
end;

procedure TvgrWorkbookGrid.OnPopupMenuClick(Sender: TObject);
var
  AVector: IvgrPageVector;
begin
  case TMenuItem(Sender).Tag of
    0: CutToClipboard;
    1: CopyToClipboard;
    2: PasteFromClipboard;
    3: DeleteRows;
    4: DeleteColumns;
    5: DoCellProperties;
    6: FPopupSections.ByIndex[FPopupSectionIndex].RepeatOnPageTop := not FPopupSections.ByIndex[FPopupSectionIndex].RepeatOnPageTop;
    7: FPopupSections.ByIndex[FPopupSectionIndex].RepeatOnPageBottom := not FPopupSections.ByIndex[FPopupSectionIndex].RepeatOnPageBottom;
    8: FPopupSections.ByIndex[FPopupSectionIndex].PrintWithNextSection := not FPopupSections.ByIndex[FPopupSectionIndex].PrintWithNextSection;
    9: FPopupSections.ByIndex[FPopupSectionIndex].PrintWithPreviosSection := not FPopupSections.ByIndex[FPopupSectionIndex].PrintWithPreviosSection;
    10: begin
          AVector := FPopupVectors.Find(FPopupVectorNumber) as IvgrPageVector;
          if AVector = nil then
          begin
            AVector := FPopupVectors[FPopupVectorNumber] as IvgrPageVector;
            AVector.PageBreak := True;
            AVector.Size := FPopupVectorNewSize;
          end
          else
            AVector.PageBreak := not AVector.PageBreak;
        end;
    11: begin
          CopyMoveWorksheet(ActiveWorksheet);
        end;
    12: begin
          if MBoxFmt(vgrLoadStr(svgrid_vgr_WorkbookGrid_DeleteWorksheetQuestion), MB_YESNO or MB_ICONQUESTION, [Workbook.Worksheets[FPopupSheetIndex].Title]) = IDYES then
            Workbook.DeleteWorksheet(FPopupSheetIndex);
        end;
    13: SetAutoHeightForSelectedRows;
    14: SetAutoWidthForSelectedColumns;
    15: begin
          Workbook.AddWorksheet;
          ActiveWorksheetIndex := Workbook.WorksheetsCount - 1;
        end;  
    16: begin
          Workbook.InsertWorksheet(FPopupSheetIndex); 
        end;
    17: begin
          EditWorksheet(ActiveWorksheet);
        end;
  end;
end;

procedure TvgrWorkbookGrid.InitPopupMenu(APopupMenu: TPopupMenu; APointInfo: TvgrPointInfo);
var
   AVector: IvgrPageVector;
begin
  ClearPopupMenu(APopupMenu);
  if (APointInfo is TvgrSheetPointInfo) or (APointInfo is TvgrHeadersPanelPointInfo) then
  begin
    AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_Common_MenuItemCut), OnPopupMenuClick, 0, 'VGR_CUT', IsCutAllowed, False, 'Ctrl+X');
    AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_Common_MenuItemCopy), OnPopupMenuClick, 1, 'VGR_COPY', IsCopyAllowed, False, 'Ctrl+C');
    AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_Common_MenuItemPaste), OnPopupMenuClick, 2, 'VGR_PASTE', IsPasteAllowed, False, 'Ctrl+V');
    if not IsInplaceEdit then
    begin
      AddMenuSeparator(APopupMenu, nil);
      AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemDeleteRows), OnPopupMenuClick, 3, '',  not ReadOnly and (ActiveWorksheet <> nil), False);
      AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemDeleteColumns), OnPopupMenuClick, 4, '', not ReadOnly and (ActiveWorksheet <> nil), False);
      AddMenuSeparator(APopupMenu, nil);
      AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemRowsAutoHeight), OnPopupMenuClick, 13, '',  not ReadOnly and (ActiveWorksheet <> nil), False, 'Shift+Ctrl+R');
      AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemColumnsAutoWidth), OnPopupMenuClick, 14, '', not ReadOnly and (ActiveWorksheet <> nil), False, 'Shift+Ctrl+C');
      AddMenuSeparator(APopupMenu, nil);
      AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemCellProperties), OnPopupMenuClick, 5, 'VGR_CELLPROPERTIES', not ReadOnly and (ActiveWorksheet <> nil) and Assigned(FOnCellProperties), False, 'Ctrl+1');
    end;

    if APointInfo is TvgrHeadersPanelPointInfo then
    begin
      AddMenuSeparator(APopupMenu, nil);
      with TvgrHeadersPanelPointInfo(APointInfo) do
      begin
        FPopupVectors := Vectors;
        FPopupVectorNumber := Number;
      end;
      if APointInfo is TvgrRowsPanelPointInfo then
        FPopupVectorNewSize := cDefaultRowHeightTwips
      else
        if APointInfo is TvgrColsPanelPointInfo then
          FPopupVectorNewSize := cDefaultColWidthTwips
        else
          FPopupVectorNewSize := 0;
      AVector := FPopupVectors.Find(FPopupVectorNumber) as IvgrPageVector;
      AddMenuItem(APopupMenu,
                  nil,
                  vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemBreakPageAfter),
                  OnPopupMenuClick, 10, '',
                  not ReadOnly,
                  (AVector <> nil) and AVector.PageBreak);
    end;
  end
  else
    if APointInfo is TvgrSectionsPanelPointInfo then
    begin
      with TvgrSectionsPanelPointInfo(APointInfo) do
      begin
        FPopupSections := Sections;
        FPopupSectionIndex := SectionIndex;
      end;

      with FPopupSections.ByIndex[FPopupSectionIndex] do
      begin
        AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemPrintAsPageHeader),
                    OnPopupMenuClick,
                    6, '', not ReadOnly, RepeatOnPageTop);
        AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemPrintAsPageFooter),
                    OnPopupMenuClick,
                    7, '', not ReadOnly, RepeatOnPageBottom);
        AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemPrintWithNextSection),
                    OnPopupMenuClick,
                    8, '', not ReadOnly, PrintWithNextSection);
        AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemPrintWithPreviosSection),
                    OnPopupMenuClick,
                    9, '', not ReadOnly, PrintWithPreviosSection);
      end;
    end
    else
      if APointInfo is TvgrSheetsPanelPointInfo then
      begin
        AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemAddWorksheet), OnPopupMenuClick, 15,
                    '', not ReadOnly and (Workbook <> nil), False);
        if TvgrSheetsPanelPointInfo(APointInfo).SheetIndex >= 0 then
        begin
          FPopupSheetIndex := TvgrSheetsPanelPointInfo(APointInfo).SheetIndex;
          AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemInsertWorksheet), OnPopupMenuClick, 16,
                      '', not ReadOnly, False);
          AddMenuSeparator(APopupMenu, nil);
          AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemCopyMoveWorksheet), OnPopupMenuClick, 11,
                      '', not ReadOnly, False);
          AddMenuSeparator(APopupMenu, nil);
          AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemRenameWorksheet), OnPopupMenuClick, 17,
                      '', not ReadOnly, False);
          AddMenuSeparator(APopupMenu, nil);
          AddMenuItem(APopupMenu, nil, vgrLoadStr(svgrid_vgr_WorkbookGrid_MenuItemDeleteWorksheet), OnPopupMenuClick, 12,
                      '', not ReadOnly, False);
        end;
      end;
end;

procedure TvgrWorkbookGrid.DeleteRows;
begin
  if ActiveWorksheet <> nil then
    with LastSelection do
      ActiveWorksheet.DeleteRows(Top, Bottom);
end;

procedure TvgrWorkbookGrid.DeleteColumns;
begin
  if ActiveWorksheet <> nil then
    with LastSelection do
      ActiveWorksheet.DeleteCols(Left, Right);
end;

procedure TvgrWorkbookGrid.InsertRow;
begin
  if ActiveWorksheet <> nil then
    with LastSelection do
      ActiveWorksheet.InsertRows(Top, 1);
end;

procedure TvgrWorkbookGrid.InsertColumn;
begin
  if ActiveWorksheet <> nil then
    with LastSelection do
      ActiveWorksheet.InsertCols(Left, 1);
end;

procedure TvgrWorkbookGrid.InsertWorksheet(AInsertIndex: Integer = -1);
begin
  if Workbook <> nil then
  begin
    if AInsertIndex = -1 then
      AInsertIndex := ActiveWorksheetIndex;
    ActiveWorksheetIndex := -1;
    Workbook.InsertWorksheet(AInsertIndex);
    ActiveWorksheetIndex := AInsertIndex;
  end;
end;

procedure TvgrWorkbookGrid.GetSelectedSizeCallback(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  AInt: IvgrWBListItem;
begin
  if Supports(AItem, svgrRowGUID, AInt) then
    with AItem as IvgrRow do
    begin
      if FTempSize = -2 then
        FTempSize := Height
      else
        if FTempSize <> Height then
          FTempSize := -1;
      Inc(FTempCount);
    end
  else
    if Supports(AItem, svgrColGUID, AInt) then
      with AItem as IvgrCol do
      begin
        if FTempSize = -2 then
          FTempSize := Width
        else
          if FTempSize <> Width then
            FTempSize := -1;
        Inc(FTempCount);
      end;
end;

function TvgrWorkbookGrid.GetSelectedRowsHeight: Integer;
var
  I, ACount: Integer;
begin
  if ActiveWorksheet <> nil then
  begin
    FTempCount := 0;
    FTempSize := -2;
    ACount := 0;
    for I := 0 to SelectionCount - 1 do
      with Selections[I] do
      begin
        ActiveWorksheet.RowsList.FindAndCallBack(Top, Bottom, GetSelectedSizeCallback, nil);
        ACount := ACount + Bottom - Top + 1;
      end;
    if FTempCount = 0 then
    begin
      if ACount > 0 then
        Result := DefaultRowHeight
      else
        Result := -1;
    end
    else
    begin
      if (FTempCount <> ACount) and (FTempSize <> DefaultRowHeight) then
        Result := -1
      else
        Result := FTempSize;
    end
  end
  else
    Result := -1;
end;

procedure TvgrWorkbookGrid.SetSelectedRowsHeight(AHeight: Integer);
var
  I, J: Integer;
begin
  if ActiveWorksheet <> nil then
  begin
    Workbook.BeginUpdate;
    try
      for I := 0 to SelectionCount - 1 do
        with Selections[I] do
          for J := Top to Bottom do
            ActiveWorksheet.Rows[J].Height := AHeight;
    finally
      Workbook.EndUpdate;
    end;
  end;
end;

function TvgrWorkbookGrid.GetSelectedColsWidth: Integer;
var
  I, ACount: Integer;
begin
  if ActiveWorksheet <> nil then
  begin
    FTempCount := 0;
    FTempSize := -2;
    ACount := 0;
    for I := 0 to SelectionCount - 1 do
      with Selections[I] do
      begin
        ActiveWorksheet.ColsList.FindAndCallBack(Left, Right, GetSelectedSizeCallback, nil);
        ACount := ACount + Right - Left + 1;
      end;
    if FTempCount = 0 then
    begin
      if ACount > 0 then
        Result := DefaultColWidth
      else
        Result := -1;
    end
    else
    begin
      if (FTempCount <> ACount) and (FTempSize <> DefaultColWidth) then
        Result := -1
      else
        Result := FTempSize;
    end
  end
  else
    Result := -1;
end;

procedure TvgrWorkbookGrid.SetSelectedColsWidth(AWidth: Integer);
var
  I, J: Integer;
begin
  if ActiveWorksheet <> nil then
  begin
    Workbook.BeginUpdate;
    try
      for I := 0 to SelectionCount - 1 do
        with Selections[I] do
          for J := Left to Right do
            ActiveWorksheet.Cols[J].Width := AWidth;
    finally
      Workbook.EndUpdate;
    end;
  end;
end;

procedure TvgrWorkbookGrid.ShowHideSelectedColumns(AShow: Boolean);
var
  I, J: Integer;
  ACol: IvgrCol;
begin
  if ActiveWorksheet <> nil then
  begin
    Workbook.BeginUpdate;
    try
      for I := 0 to SelectionCount - 1 do
        with Selections[I] do
          for J := Left to Right do
            if AShow then
            begin
              ACol := ActiveWorksheet.ColsList.Find(J);
              if (ACol <> nil) and not ACol.Visible then
              begin
                ACol.Visible := True;
                if ConvertTwipsToPixelsX(ACol.Width) < 2 then
                  ACol.Width := DefaultColWidth;
              end
            end
            else
              with ActiveWorksheet.Cols[J] do
              begin
                Visible := False;
//                Width := DefaultColWidth;
              end;
    finally
      Workbook.EndUpdate;
    end;
  end;
end;

procedure TvgrWorkbookGrid.ShowHideSelectedRows(AShow: Boolean);
var
  I, J: Integer;
  ARow: IvgrRow;
begin
  if ActiveWorksheet <> nil then
  begin
    Workbook.BeginUpdate;
    try
      for I := 0 to SelectionCount - 1 do
        with Selections[I] do
          for J := Top to Bottom do
            if AShow then
            begin
              ARow := ActiveWorksheet.RowsList.Find(J);
              if (ARow <> nil) and not ARow.Visible then
              begin
                ARow.Visible := True;
                if ConvertTwipsToPixelsY(ARow.Height) < 2 then
                  ARow.Height := DefaultRowHeight;
              end
            end
            else
              with ActiveWorksheet.Rows[J] do
              begin
                Visible := False;
//                Height := DefaultRowHeight;
              end;
    finally
      Workbook.EndUpdate;
    end;
  end;
end;

{
function TvgrWorkbookGrid.GetRangeSizeCallback(ARange: IvgrRange): TSize;
var
  ARect: TRect;
begin
  FSheet.GetRectInternalScreenPixels(ARange.Place, ARect);
  with ARange do
  begin
    AssignFont(pvgrRangeStyle(StyleData).Font, Canvas.Font);
    Result := vgrCalcText(Canvas,
                          DisplayText,
                          ARect,
                          WordWrap,
                          GetAutoAlignForRangeValue(HorzAlign, ValueType),
                          VertAlign,
                          Angle);
    Result.cx := ConvertPixelsToTwipsX(Result.cx);
    Result.cy := ConvertPixelsToTwipsY(Result.cy);
  end;
end;
}

procedure TvgrWorkbookGrid.SetAutoWidthForSelectedColumns;
var
  I, J, AAutoWidth: Integer;
  ACol: IvgrCol;
begin
  if ActiveWorksheet <> nil then
  begin
    Workbook.BeginUpdate;
    try
      for I := 0 to SelectionCount - 1 do
        with Selections[I] do
          for J := Left to Right do
          begin
            AAutoWidth := ActiveWorksheet.GetColAutoWidth(J{, GetRangeSizeCallback});
            if AAutoWidth = 0 then
              AAutoWidth := DefaultColWidth;
            ACol := ActiveWorksheet.ColsList.Find(J);
            if ACol = nil then
            begin
              if AAutoWidth <> DefaultColWidth then
                ActiveWorksheet.Cols[J].Width := AAutoWidth
            end
            else
            begin
              if (ACol.Width <> AAutoWidth) or not ACol.Visible then
              begin
                ACol.Width := AAutoWidth;
                ACol.Visible := True;
              end;
            end;
          end;
    finally
      Workbook.EndUpdate;
    end;
  end;
end;

procedure TvgrWorkbookGrid.SetAutoHeightForSelectedRows;
var
  I, J, AAutoHeight: Integer;
  ARow: IvgrRow;
begin
  if ActiveWorksheet <> nil then
  begin
    Workbook.BeginUpdate;
    try
      for I := 0 to SelectionCount - 1 do
        with Selections[I] do
          for J := Top to Bottom do
          begin
            AAutoHeight := ActiveWorksheet.GetRowAutoHeight(J{, GetRangeSizeCallback});
            if AAutoHeight = 0 then
              AAutoHeight := DefaultRowHeight;
            ARow := ActiveWorksheet.RowsList.Find(J);
            if ARow = nil then
            begin
              if AAutoHeight <> DefaultRowHeight then
                ActiveWorksheet.Rows[J].Height := AAutoHeight
            end
            else
            begin
              if (ARow.Height <> AAutoHeight) or not ARow.Visible then
              begin
                ARow.Height := AAutoHeight;
                ARow.Visible := True;
              end;
            end;
          end;
    finally
      Workbook.EndUpdate;
    end;
  end;
end;

function TvgrWorkbookGrid.IsMergeSelected: Boolean;
var
  I: Integer;
begin
  Result := ActiveWorksheet <> nil;
  if Result then
  begin
    for I := 0 to SelectionCount - 1 do
      if ActiveWorksheet.IsMergeInRect(Selections[I]) then
        exit;
    Result := False;
  end;
end;

function TvgrWorkbookGrid.MergeWarningProc: Boolean;
begin
  Result := MBox(vgrLoadStr(svgrid_vgr_WorkbookGrid_MergeWarning), MB_OKCANCEL or MB_ICONEXCLAMATION) = IDOK;
end;

procedure TvgrWorkbookGrid.MergeUnMergeSelection;
var
  I: Integer;
begin
  if ActiveWorksheet <> nil then
  begin
    for I := 0 to SelectionCount - 1 do
      ActiveWorksheet.MergeUnMerge(Selections[I], MergeWarningProc);
  end;
end;

function TvgrWorkbookGrid.IsCutAllowed: Boolean;
begin
  Result := not ReadOnly and (ActiveWorksheet <> nil) and (SelectionCount > 0);
end;

function TvgrWorkbookGrid.IsCopyAllowed: Boolean;
begin
  Result := (ActiveWorksheet <> nil) and (SelectionCount > 0);
end;

function TvgrWorkbookGrid.IsPasteAllowed: Boolean;
begin
  Result := not ReadOnly and Clipboard.HasFormat(CF_VGRCELLS);
end;

procedure TvgrWorkbookGrid.CutToClipboard;
begin
  if ActiveWorksheet <> nil then
    ActiveWorksheet.CutToClipboard(SelectionRects);
end;

procedure TvgrWorkbookGrid.CopyToClipboard;
begin
  if ActiveWorksheet <> nil then
    ActiveWorksheet.CopyToClipboard(SelectionRects);
end;

procedure TvgrWorkbookGrid.PasteFromClipboard;
begin
  if ActiveWorksheet <> nil then
  begin
    ActiveWorksheet.PasteFromClipboard(Selections[0], [vgrptRangeValue, vgrptRangeStyle, vgrptBorders, vgrptCols, vgrptRows]);
  end;
end;

procedure TvgrWorkbookGrid.ClearSelection(AClearFlags: TvgrClearContentFlags);
begin
  if ActiveWorksheet <> nil then
    ActiveWorksheet.ClearContents(SelectionRects, AClearFlags);
end;

procedure TvgrWorkbookGrid.ShowPopupMenu(const APopupPoint: TPoint; APointInfo: TvgrPointInfo);
begin
  if DefaultPopupMenu and (PopupMenu = nil) then
  begin
    InitPopupMenu(FPopupMenu, APointInfo);
    FPopupMenu.Popup(APopupPoint.X, APopupPoint.Y);
  end;
end;

function TvgrWorkbookGrid.EditWorksheet(AWorksheet: TvgrWorksheet): Boolean;
begin
  Result := TvgrSheetFormatForm.Create(nil).Execute(AWorksheet);
end;

function TvgrWorkbookGrid.CopyMoveWorksheet(AWorksheet: TvgrWorksheet): Boolean;
begin
  Result := TvgrCopyMoveSheetDialogForm.Create(nil).Execute(AWorksheet);
end;

procedure TvgrWorkbookGrid.GetPointInfoAt(X, Y: Integer; var APointInfo: TvgrPointInfo);
var
  APoint: TPoint;

  function ConvertPoint(AChild: TControl): TPoint;
  begin
    Result := AChild.ScreenToClient(Self.ClientToScreen(APoint));
  end;
  
begin
  APoint := Point(X, Y);
  if FFormulaPanel.Visible and PtInRect(FFormulaPanel.BoundsRect, APoint) then
    FFormulaPanel.GetPointInfoAt(ConvertPoint(FFormulaPanel), APointInfo)
  else
    if FTopLeftSectionsButton.Visible and PtInRect(FTopLeftSectionsButton.BoundsRect, APoint) then
      FTopLeftSectionsButton.GetPointInfoAt(ConvertPoint(FTopLeftSectionsButton), APointInfo)
    else
      if FTopLeftButton.Visible and PtInRect(FTopLeftButton.BoundsRect, APoint) then
        FTopLeftButton.GetPointInfoAt(ConvertPoint(FTopLeftButton), APointInfo)
      else
        if FHorzSections.Visible and PtInRect(FHorzSections.BoundsRect, APoint) then
          FHorzSections.GetPointInfoAt(ConvertPoint(FHorzSections), APointInfo)
        else
          if FVertSections.Visible and PtInRect(FVertSections.BoundsRect, APoint) then
            FVertSections.GetPointInfoAt(ConvertPoint(FVertSections), APointInfo)
          else
            if FColsPanel.Visible and PtInRect(FColsPanel.BoundsRect, APoint) then
              FColsPanel.GetPointInfoAt(ConvertPoint(FColsPanel), APointInfo)
            else
              if FRowsPanel.Visible and PtInRect(FRowsPanel.BoundsRect, APoint) then
                FRowsPanel.GetPointInfoAt(ConvertPoint(FRowsPanel), APointInfo)
              else
                if FSheet.Visible and PtInRect(FSheet.BoundsRect, APoint) then
                  FSheet.GetPointInfoAt(ConvertPoint(FSheet), APointInfo)
                else
                  if FSheetsPanel.Visible and PtInRect(FSheetsPanel.BoundsRect, APoint) then
                    FSheetsPanel.GetPointInfoAt(ConvertPoint(FSheetsPanel), APointInfo)
                  else
                    if FSheetsPanelResizer.Visible and PtInRect(FSheetsPanelResizer.BoundsRect, APoint) then
                      FSheetsPanelResizer.GetPointInfoAt(ConvertPoint(FSheetsPanelResizer), APointInfo)
                    else
                      if FHorzScrollBar.Visible and PtInRect(FHorzScrollBar.BoundsRect, APoint) then
                        FHorzScrollBar.GetPointInfoAt(ConvertPoint(FHorzScrollBar), APointInfo)
                      else
                        if FVertScrollBar.Visible and PtInRect(FVertScrollBar.BoundsRect, APoint) then
                          FVertScrollBar.GetPointInfoAt(ConvertPoint(FVertScrollBar), APointInfo)
                        else
                          if FSizer.Visible and PtInRect(FSizer.BoundsRect, APoint) then
                            FSizer.GetPointInfoAt(ConvertPoint(FSizer), APointInfo)
                          else
                            APointInfo := nil;
end;

procedure TvgrWorkbookGrid.DoMouseMove(AChild: TControl; Shift: TShiftState; X, Y: Integer);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  MouseMove(Shift, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookGrid.DoMouseDown(AChild: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  AExecuteDefault := True;
  MouseDown(Button, Shift, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookGrid.DoMouseUp(AChild: TControl; Button: TMouseButton; Shift: TShiftState; X, Y: Integer; var AExecuteDefault: Boolean);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  AExecuteDefault := True;
  MouseUp(Button, Shift, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookGrid.DoDblClick(AChild: TControl; var AExecuteDefault: Boolean);
begin
  AExecuteDefault := True;
  DblClick;
end;

procedure TvgrWorkbookGrid.DoDragOver(AChild: TControl; Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  DragOver(Source, APoint.X, APoint.Y, State, Accept);
end;

procedure TvgrWorkbookGrid.DoDragDrop(AChild: TControl; Source: TObject; X, Y: Integer);
var
  APoint: TPoint;
begin
  APoint := ScreenToClient(AChild.ClientToScreen(Point(X, Y)));
  DragDrop(Source, APoint.X, APoint.Y);
end;

procedure TvgrWorkbookGrid.DoActiveWorksheetChanged;
begin
  if Assigned(FOnActiveWorksheetChanged) then
    FOnActiveWorksheetChanged(Self);
end;

procedure TvgrWorkbookGrid.DoSetInplaceEditValue(const AValue: string; ARange: IvgrRange; ANewRange: Boolean);
begin
  if Assigned(OnSetInplaceEditValue) then
    OnSetInplaceEditValue(Self, AValue, ARange, ANewRange)
  else
  begin
    ARange.StringValue := AValue;
    if ANewRange then
      ARange.WordWrap := (pos(#13, AValue) <> 0) or (pos(#10, AValue) <> 0);
  end;
end;

initialization

  FFormulaOkBitmap := TBitmap.Create;
  FFormulaOkBitmap.LoadFromResourceName(hInstance, 'VGR_BMP_FORMULAOK');
  FFormulaOkBitmap.Transparent := True;

  FFormulaCancelBitmap := TBitmap.Create;
  FFormulaCancelBitmap.LoadFromResourceName(hInstance, 'VGR_BMP_FORMULACANCEL');
  FFormulaCancelBitmap.Transparent := True;

  FSectionBitmaps := TImageList.Create(nil);
  FSectionBitmaps.Width := 5;
  FSectionBitmaps.Height := 5;
  FSectionBitmaps.ResourceLoad(rtBitmap, svgrSectionBitmaps, clFuchsia);

  // Load cursors
  LoadCursor(cCellCursorResName, FCellCursor, cCellCursorID, crDefault);
  LoadCursor(cToBottomCursorResName, FToBottomCursor, cToBottomCursorID, crDefault);
  LoadCursor(cToRightCursorResName, FToRightCursor, cToRightCursorID, crDefault);

finalization

  FFormulaOkBitmap.Free;
  FFormulaCancelBitmap.Free;
  FSectionBitmaps.Free;

end.

