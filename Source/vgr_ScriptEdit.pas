
{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains the TvgrScriptEdit control.
See also:
  TvgrScriptEdit}
unit vgr_ScriptEdit;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, SysUtils, Controls, ExtCtrls, Windows, graphics, forms, stdctrls,
  messages, math, clipbrd, Contnrs, {$IFDEF VGR_D6_OR_D7} Types, VarUtils, {$ENDIF} ImgList,

  vgr_ScriptControl, vgr_AXScript, vgr_Keyboard;

const
  { Max line length in ScriptEdit }
  cSEMaxLineLength = 1024;

  { For internal use }
  CM_MOUSECHECK = WM_USER + 1;

  { Resource name }
  sSyntaxCheckBitmap = 'VGR_BMP_SCRIPTSYNTAXCHECK';
  { Resource name }
  sSyntaxCheckBitmapSel = 'VGR_BMP_SCRIPTSYNTAXCHECK_SEL';

  { Code if I char is pressed }
  VK_CHAR_I = 73;
  { Code if U char is pressed }
  VK_CHAR_U = 85;
  { Code if C char is pressed }
  VK_CHAR_C = 67;
  { Code if V char is pressed }
  VK_CHAR_V = 86;
  { Code if X char is pressed }
  VK_CHAR_X = 88;
  { Code if A char is pressed }
  VK_CHAR_A = 65;

  VGR_MK_CONTROL = 4;
  VGR_MK_LBUTTON = 0;
  VGR_MK_RBUTTON = 0;
  VGR_MK_SHIFT = 1;

type

  TvgrCustomScriptEdit = class;
  TvgrSEMemo = class;

  { Control's edit mode
    possible values: vmmInsert, vmmOverwrite }
  TvgrSEInsertMode = (vmmInsert, vmmOverwrite);

  /////////////////////////////////////////////////
  //
  // TvgrSETextElement
  //
  /////////////////////////////////////////////////
  { TvgrSETextElement represents the font characteristics of a section of text in a script edit control. }
  TvgrSETextElement = class(TPersistent)
  private
    FFontStyle : TFontStyles;
    FColor : TColor;
    FDefaultColor : Boolean;

    procedure SetTextAttrs(
      AStyle: TFontStyles;
      AColor: TColor;
      ADefaultColor: Boolean);
    function GetColor : TColor;
  published
    { Determines whether the font is normal, italic, underlined, bold, or strikeout. }
    property FontStyle : TFontStyles read FFontStyle write FFontStyle;
    { Indicates the foreground color of the text. }
    property Color : TColor read GetColor write FColor;
    { Use the default color for text element }
    property DefaultColor : Boolean read FDefaultColor write FDefaultColor;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSETextColors
  //
  /////////////////////////////////////////////////
  { Represents different text attributes for different script elements }
  TvgrSETextColors = class(TPersistent)
  private
    FWhitespace : TvgrSETextElement;
    FKeyword    : TvgrSETextElement;
    FComment    : TvgrSETextElement;
    FNonsource  : TvgrSETextElement;
    FOperator   : TvgrSETextElement;
    FNumber     : TvgrSETextElement;
    FString     : TvgrSETextElement;
    FFunction   : TvgrSETextElement;
    FIdentifier : TvgrSETextElement;

    procedure SetDefaultSheme;
  public
    { Creates and initializes a TextColors object }
    constructor Create;
    { Destroys the TvgrSETextColors object and frees its memory }
    destructor Destroy; override;
  published
    { Text attributes for "whitespace" element in the script }
    property AttrWhitespace : TvgrSETextElement read FWhitespace write FWhitespace;
    { Text attributes for "Keyword" element in the script }
    property AttrKeyword    : TvgrSETextElement read FKeyword write FKeyword;
    { Text attributes for "Comment" element in the script }
    property AttrComment    : TvgrSETextElement read FComment write FComment;
    { Text attributes for "Nonsource" element in the script }
    property AttrNonsource  : TvgrSETextElement read FNonsource write FNonsource;
    { Text attributes for "Operator" element in the script }
    property AttrOperator   : TvgrSETextElement read FOperator write FOperator;
    { Text attributes for "Number" element in the script }
    property AttrNumber     : TvgrSETextElement read FNumber  write FNumber;
    { Text attributes for "String" element in the script }
    property AttrString     : TvgrSETextElement read FString  write FString;
    { Text attributes for "Function" element in the script }
    property AttrFunction   : TvgrSETextElement read FFunction write FFunction;
    { Text attributes for "Identifier" element in the script }
    property AttrIdentifier : TvgrSETextElement read FIdentifier write FIdentifier;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSETextOptions
  //
  /////////////////////////////////////////////////
  { Specifies common text options for script edit }
  TvgrSETextOptions = class(TPersistent)
  private
    FFontName: TFontName;
    FFontSize: Integer;
    FTextColors: TvgrSETextColors;
    FColor: TColor;
    FErrorBackColor : TColor;
    FErrorForeColor : TColor;
    FMemo: TvgrSEMemo;
    procedure SetFontName(Value: TFontName);
    procedure SetFontSize(Value: Integer);
    procedure SetErrorBackColor(Value: TColor);
    procedure SetErrorForeColor(Value: TColor);
    procedure Change;
  public
    { Creates and initializes a TextOptions object }
    constructor Create(Memo: TvgrSEMemo);
    { Destroys the TvgrSETextOptions object and frees its memory }
    destructor Destroy; override;
  published
    { Specifies the typeface of the font for script edit }
    property FontName : TFontName read FFontName write SetFontName;
    { Specifies the height of the font in points for script edit }
    property FontSize : Integer read FFontSize write SetFontSize;
    { common text options for script edit }
    property TextColors : TvgrSETextColors read FTextColors write FTextColors;
    { Determines the background color of the script edit control }
    property BkColor : TColor read FColor write FColor;
    { Determines the background color of the error line in script edit control }
    property ErrorBackColor : TColor read FErrorBackColor write SetErrorBackColor;
    { Determines the foreground color of the text in error line }
    property ErrorForeColor : TColor read FErrorForeColor write SetErrorForeColor;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrWBScrollBar
  //
  /////////////////////////////////////////////////
  { TvgrSEScrollBar is a TScrollBar, which is used to scroll the content of a script edit. }
  TvgrSEScrollBar = class(TScrollBar)
  private
    FScriptEdit : TvgrCustomScriptEdit;
    procedure SetParams(APosition, AMin, AMax, APageSize : Integer);
  protected
    { Initializes a window-creation parameter record passed in the Params parameter }
    procedure CreateParams(var Params: TCreateParams); override;
    { Creates a Windows control corresponding to the script edit control }
    procedure CreateWnd; override;
    { Scroll ScriptEdit contents }
    procedure Scroll(ScrollCode: TScrollCode; var ScrollPos: Integer); override;
  public
    { Creates and initializes an instance of TScrollBar }
    constructor Create(AOwner : TComponent); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSEGutterSwitcher
  //
  /////////////////////////////////////////////////
  { It is a panel, with the help of which the user can
    hide or show the inspected element }
  TvgrSEControlSwitcher = class(TCustomControl)
  private
    FCurrColor: TColor;
    FScriptEdit : TvgrCustomScriptEdit;
    FControl: TControl;
    FKind: TScrollBarKind;
    function GetControlVisible: Boolean;
    function CreatePimpochkaRgn: THandle;
    procedure SetControl(Value: TControl);
    procedure ClearBackground(AColor : TColor);
    procedure DrawPup(X,Y : Integer);
    procedure DrawArrow;
    procedure DrawEdge;
    procedure DrawPimpochka(ABackground : TColor; AlwaysDraw: Boolean);
    procedure SetControlVisible(Value: Boolean);
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
  protected
    { Draws the image of the Switcher control on the screen }
    procedure Paint; override;
    { Respond to mouse moving over control area }
    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    { Handles mouse button clicks }
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
  public
    { Specifies the size of the switcher in pixels }
    property Width;
    { inspected control }
    property Control: TControl read FControl write SetControl;
    { Determines whether the inspected control is visible }
    property ControlVisible: Boolean read GetControlVisible write SetControlVisible;
    { Specifies whether the scroll bar is horizontal or vertical
    These are the possible values: sbHorizontal, sbVertical	}
    property Kind: TScrollBarKind read FKind write FKind;
    { Creates and initializes an instance of TvgrSEControlSwitcher }
    constructor Create(AOwner : TComponent); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSEGutter
  //
  /////////////////////////////////////////////////
  { gutter - panel with Script edit line numbers }
  TvgrSEGutter = class(TCustomControl)
  private
    FNumerate : Boolean;
    FScriptEdit : TvgrCustomScriptEdit;
    procedure ApplyFontForLine(ALine: Integer);
    procedure DrawLineBackGround(const ALineRect: TRect);
    procedure PaintLine(ALine: Integer; APaintBackground: Boolean);
    procedure PaintLines(APaintBackground: Boolean);
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    function GetSEScrollBar: TvgrSEScrollBar; virtual;
    function GetScriptEdit : TvgrCustomScriptEdit;
{    function GetHeaderRect(From,Size: Integer): TRect; virtual;
    function GetVisibleRowsCount: Integer; virtual;
    procedure PaintHeaderBorder(const HeaderRect: TRect); virtual;
    procedure DoOnFontChange(Sender: TObject); virtual;
    procedure DoScroll(OldPos,NewPos: Integer);}
  protected
    { Handles mouse button clicks }
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    { Draws the image of the Gutter on the screen }
    procedure Paint; override;
  public
    { Script edit control }
    property ScriptEdit : TvgrCustomScriptEdit read GetScriptEdit;
    { Vertical scroll bar of the script edit }
    property SEScrollBar : TvgrSEScrollBar read GetSEScrollBar;
    { Determines whether the line numbers is visible }
    property Numerate: Boolean read FNumerate write FNumerate default True;
    { Determines the width of the gutter, default is 35 }
    property Width;
    { Determines whether the gutter is visible }
    property Visible;

    { Creates and initializes an instance of TvgrSEGutter }
    constructor Create(AOwner : TComponent); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSESyntaxPanel
  //
  /////////////////////////////////////////////////
  { Panel for syntax checking of script source }
  TvgrSESyntaxPanel = class(TCustomControl)
  private
    FScriptEdit : TvgrCustomScriptEdit;
    FMouseIsPressed: Boolean;
    FMouseInCheckIconRect: Boolean;
    FImageList: TImageList;

    procedure CheckSyntax;
    function GetSyntaxCheckIconRect: TRect;
    function MouseInCheckIconRect(X, Y: Integer; ALBPressed: Boolean): Boolean;
    procedure FillBackground;
    procedure PaintSyntaxIcon;
    procedure PaintErrorString;
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CMMouseLeave(var Message: TMessage); message CM_MOUSELEAVE;
  protected
    { Draws the image of the Syntax Panel on the screen }
    procedure Paint; override;
    { Respond to mouse moving over control area}
    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    { Handles mouse button clicks }
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    { Handles releasing of a mouse button }
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
  public
    { Creates and initializes an instance of TvgrSESyntaxPanel }
    constructor Create(AOwner : TComponent); override;
    { Destroys the TvgrSESyntaxPanel object and frees its memory }
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSESelection
  //
  /////////////////////////////////////////////////
  { class for management of selected text }
  TvgrSESelection = class
    FStartPoint : TPoint;
    FActive : Boolean;
    FBlockMode : Boolean;
    FMemo : TvgrSEMemo;
  private
    procedure Reset;
    function GetTopPoint : TPoint;
    function GetEndPoint : TPoint;

    property StartPoint : TPoint read FStartPoint write FStartPoint;
    property TopPoint   : TPoint read GetTopPoint;
    property EndPoint   : TPoint read GetEndPoint;
    property Active : Boolean read FActive write FActive;
    property BlockMode : Boolean read FBlockMode write FBlockMode;
  public
    { Creates and initializes an instance of Selection object }
    constructor Create(AMemo : TvgrSEMemo);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSEMemo
  //
  /////////////////////////////////////////////////
  { Part of script edit control, where is fulfilled text editing }
  TvgrSEMemo = class(TCustomControl)
  private
    FAbsoluteCaretPos : TPoint;
    FVisibleTextOffset : TPoint;
    FCharacterWidth: Integer;
    FLineTextHeight: Integer;
    FSelection : TvgrSESelection;
    FLMBDown : Boolean;
    FMMShift : TShiftState;
    FMMPosX, FMMPosY : Integer;
    FUpdateState: Boolean;
    FKeyboardFilter: TvgrKeyboardFilter;

    function GetLineTextWithoutTabs(ALine: Integer): string;
    procedure SelectNearestWord;
    function GetTabString: string;
    procedure AddBlankLineIfNeed;
    procedure DrawSyntaxErrorLine(ALine: Integer);
    procedure DeleteNextCharAbs(var AText: string; APosition: Integer);
    procedure DeletePrevCharAbs(var AText: string; APosition: Integer; out ACharIsTab: Boolean);
    function CopyAbs(const AText: string; AIndex, ALength: Integer): string;
    function LengthAbs(const Value: string): Integer;
    function RemoveSelectedText: Boolean;
    procedure RecalcFontMetrics;
    function GetSelectionBrushColor: TColor;
    function GetSelectionFontColor: TColor;
    function GetErrorLineBrushColor: TColor;
    function GetErrorLineFontColor: TColor;
    procedure StartSelection(ABlockMode: Boolean);
    function CreateSelectedRgn(ALine: Integer; const ALineRect: TRect): THandle;
    function SymbolIsDelimiter(Value : Char) : Boolean;
    function GetCharacterWidth : Integer;
    procedure SetCaretToPosition;
    procedure SetCaretToPositionWithoutScrolling;
    procedure ScrollToCaret;
    procedure SetCaretPosToAbsolute(Char, Line : Integer);
    procedure MoveCaretToClientRect;
    procedure CheckSelectionMode(AShiftPressed, AAltPressed: Boolean);
    procedure ChangeSelectedText(AShiftPressed, AAltPressed: Boolean);


    function KeyPressCR: Boolean;
    function KeyPressDelete: Boolean;
    function KeyPressBackspace: Boolean;
    function KeyPressTab: Boolean;
    function KeyPressIndent: Boolean;
    function KeyPressUnindent: Boolean;

    function MoveCaretPosToLeft: Boolean;
    function MoveCaretPosToRight: Boolean;
    function MoveCaretPosToUp: Boolean;
    function MoveCaretPosToDown: Boolean;
    function MoveCaretToNextPage: Boolean;
    function MoveCaretToPrevPage: Boolean;
    function MoveCaretToBeginOfLine: Boolean;
    function MoveCaretToEndOfLine: Boolean;
    function MoveToNextWord: Boolean;
    function MoveToPrevWord: Boolean;
    function MoveCaretToBottomOfPage: Boolean;
    function MoveCaretToTopOfPage: Boolean;
    function MoveCaretToBeginOfLines: Boolean;
    function MoveCaretToEndOfLines: Boolean;
    function ScrollPageLineUp: Boolean;
    function ScrollPageLineDown: Boolean;
    function WheelLineUp: Boolean;
    function WheelLineDown: Boolean;
    function WheelPageUp: Boolean;
    function WheelPageDown: Boolean;

    function MoveToNextWordInLine: Boolean;
    function MoveToPrevWordInLine: Boolean;

    procedure SetCaretPosNear(X,Y : Integer);
    function GetCharacterNear(X,Y : Integer) : Integer;
    function GetLineNumberNear(Y : Integer) : Integer;
    function GetLineHeight : Integer;
    function GetLineTextHeight : Integer;
    function GetInterlineOffset : Integer;
    function GetLeftMargin : Integer;
    function GetFirstVisibleLine : Integer;
    function GetFirstVisibleCharacter : Integer;
    function GetLastVisibleLine : Integer;
    function GetVisibledLinesCount : Integer;
    function GetLastVisibleCharacter : Integer;
    function GetVisibledCharactersCount : Integer;
    procedure ClearBkGroundInRect(ARect : TRect);
    function  LineSelectionPresent(ALine: Integer): Boolean;
    function FindLineAttr(ALine : Integer) : PWord;
    procedure ApplyFontByAttr(AAttrib : Word; AUseFontColor: Boolean);
    procedure DrawLineText(ATextPosY: Integer; APAttr : PWord; const ALineRect: TRect; const ATextToOut: string; ABrushColor,AFontColor: TColor; AUseFontColor : Boolean);
    procedure PaintTextLine(ALine : Integer; ATextPosY : Integer; const ALineRect : TRect);
    procedure PaintText;
    procedure PaintTextLines(ABegLine,AEndLine : Integer);
    procedure PaintTextBlock(ALeft, ATop, ARight, ABottom : Integer);
    procedure SetCaret;
    procedure ResetCaret;
    procedure Change;
    procedure ChangeText;
    procedure ChangeTextLines(ABegLine, AEndLine : Integer);
    procedure ChangeTextBlock(ALeft, ATop, ARight, ABottom : Integer);
    function GetScriptEdit: TvgrCustomScriptEdit;
    function GetLines : TStrings;
    function GetTextOptions : TvgrSETextOptions;
    procedure SetTextOptions(AValue: TvgrSETextOptions);
    procedure SetVisibleTextOffset(Value: TPoint);
    procedure SetVisibleTextOffsetWithoutChangingCaretPos(Value: TPoint);
    procedure SetAbsoluteCaretPos(Value: TPoint);
    function  GetEditorMode: TvgrSEInsertMode;
    procedure SetEditorMode(Value: TvgrSEInsertMode);
    function GetSelectedText: string;
    function GetTabIncrement: Integer;
    function GetReadOnly: Boolean;
    function GetReplaceTabsWithSpaces: Boolean;
    procedure BeforeSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
    procedure AfterSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);

    procedure InsertTextInLine(const ATextToInserting: string; var AText: string; APosition: Integer);
    procedure InitKeyboardFilter;
    procedure AddInstantCommands;
    procedure AddDefLayerCommands;
    procedure DoKeyPress(AKey: Char);

    function ClipBoardCopy: Boolean;
    function ClipBoardPaste: Boolean;
    function ClipBoardCut: Boolean;

    function SelectAll: Boolean;
    function DeletePreviosWord: Boolean;


    function IsUpdated: Boolean;

    procedure CMMousecheck(var Msg : TWMMouse); message CM_MOUSECHECK;

    procedure WMEraseBkgnd(var Msg : TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMGetDlgCode(var Msg : TWMGetDlgCode); message WM_GETDLGCODE;
    procedure WMSetFocus(var Message: TWMSetFocus); message WM_SETFOCUS;
    procedure WMKillFocus(var Message: TWMSetFocus); message WM_KILLFOCUS;
    procedure WMCopy(var Message: TWMCopy); message WM_COPY;
    procedure WMPaste(var Message: TWMPaste); message WM_PASTE;
    procedure WMCut(var Message: TWMCut); message WM_CUT;
    procedure WMLButtonDblClk(var Message: TWMLButtonDblClk); message WM_LBUTTONDBLCLK;

    property Lines : TStrings read GetLines;
    property TextOptions : TvgrSETextOptions read GetTextOptions write SetTextOptions;
    property Selection : TvgrSESelection read FSelection;
    property VisibleTextOffset: TPoint read FVisibleTextOffset write SetVisibleTextOffset;
    property AbsoluteCaretPos: TPoint read FAbsoluteCaretPos write SetAbsoluteCaretPos;
    property EditorMode: TvgrSEInsertMode read GetEditorMode write SetEditorMode;
    property TabIncrement: Integer read GetTabIncrement;
    property ReadOnly: Boolean read GetReadOnly;
    property ReplaceTabsWithSpaces: Boolean read GetReplaceTabsWithSpaces;
    property SelectedText: string read GetSelectedText;
    property ScriptEdit: TvgrCustomScriptEdit read GetScriptEdit;
  protected
    { Initializes a window-creation parameter record passed in the Params parameter }
    procedure CreateParams(var Params: TCreateParams); override;
    { Respond to control resize }
    procedure Resize; override;
    { Draws the image of the control on the screen }
    procedure Paint; override;
    { Responds when the user presses a key }
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    { Key up event dispatcher }
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
    { Responds when the user presses a key }
    procedure KeyPress(var Key: Char); override;

    { Handles mouse movements }
    procedure MouseMove(Shift: TShiftState; X,Y: Integer); override;
    { Handles mouse button clicks }
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    { Mouse release event dispatcher }
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure Click; override;
    procedure DblClick; override;

  public
    { Creates and initializes an instance of TvgrSEMemo }
    constructor Create(AOwner : TComponent); override;
    { Frees the memory associated with the Memo }
    destructor Destroy; override;
    { Dispatches messages received from a mouse wheel. }
    procedure MouseWheelHandler(var Message: TMessage); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSEStrings
  //
  /////////////////////////////////////////////////
  { A list of strings, for storaging script text }
  TvgrSEStrings = class(TStringList)
  private
    FScriptEdit : TvgrCustomScriptEdit;
    FUpdated : Boolean;
  protected
    { Updates Script Edit control }
    procedure Changed; override;
    { Sets Update state to the Updating value }
    procedure SetUpdateState(Updating: Boolean); override;
  public
    { Creates an instance of TvgrSEStrings}
    constructor Create(AScriptEdit : TvgrCustomScriptEdit);
    { Loads the text from a stream }
    procedure LoadFromStream(Stream: TStream); override;
  end;

  TvgrCaretMoveEvent = procedure(Sender: TObject; X,Y : Integer) of object;
  /////////////////////////////////////////////////
  //
  // TvgrCustomScriptEdit
  //
  /////////////////////////////////////////////////
{TCustomEdit is the base class from which script edit controls are derived.}
  TvgrCustomScriptEdit = class(TCustomControl)
  private
    FOnCaretMove: TvgrCaretMoveEvent;
    FMaxLineLength : Integer;
    FTextOptions : TvgrSETextOptions;
    FLines : TStrings;
    FMemo : TvgrSEMemo;
    FHorzScrollBar : TvgrSEScrollBar;
    FVertScrollBar : TvgrSEScrollBar;
    FGutter : TvgrSEGutter;
    FSyntaxCheck: TvgrSESyntaxPanel;
    FGutterSwitcher: TvgrSEControlSwitcher;
    FSyntaxCheckSwitcher: TvgrSEControlSwitcher;
    FTextAttributes : PWord;
    FBriefCursorShapes : Boolean;
    FEditorMode : TvgrSEInsertMode;
    FAutoEditorMode : Boolean;
    FTabIncrement: Integer;
    FBorderStyle : TBorderStyle;
    FOEMConvert: Boolean;
    FReadOnly: Boolean;
    FScrollBars: TScrollStyle;
    FWantTabs: Boolean;
    FWantReturns: Boolean;
    FReplaceTabsWithSpaces: Boolean;
    FSyntaxErrorPresent: Boolean;
    FSyntaxError: TvgrScriptError;
    FBottomRightRect : TRect;
    FOnChange: TNotifyEvent;
    FOnEditorModeChange: TNotifyEvent;
    FScript: TvgrScriptSyntax;

    procedure SetCaretToSyntaxError;
    procedure DrawSyntaxErrorLine;
    function GetBottomRightRect : TRect;
    function GetBorderWidth: Integer;
    procedure TextChanged;
    procedure GetTextAttributes;
    procedure InitHorzScrollBar;
    procedure UpdateScrollBars;
    procedure UpdateGutter;
    procedure UpdateScrollBar(ScrollBar : TvgrSEScrollBar; NewPos, VisiblePoints, MaxPoints : integer);
    procedure ChangeAbsoluteCaretPos(AOldValue, ANewValue: Integer);
    procedure ChangeVisibleYOffset;
    procedure SetBorderStyle(Value: TBorderStyle);
    procedure SetLines(Value: TStrings);
    procedure SetScrollBars(Value: TScrollStyle);
    function GetGutterVisible: Boolean;
    procedure SetGutterVisible(Value: Boolean);
    procedure SetTabIncrement(Value: Integer);
    function GetSelectionActive: Boolean;
    procedure SetEditorMode(Value: TvgrSEInsertMode);
    function GetCaretPos: TPoint;
    procedure SetScript(Value: TvgrScriptSyntax);
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMSetFocus(var Message: TWMSetFocus); message WM_SETFOCUS;
    procedure OnScriptChange(Sender: TObject);
    property Gutter: TvgrSEGutter read FGutter;
  protected
    { Renders the image of the script edit }
    procedure Paint; override;
    { Responds when the dimensions of the script edit change }
    procedure Resize; override;
    { Aligns any controls for which the control is the parent within a specified area of the control. }
    procedure AlignControls(AControl: TControl; var Rect: TRect); override;
    { Initializes a window-creation parameter record passed in the Params parameter }
    procedure CreateParams(var Params: TCreateParams); override;
    { Contains the individual lines of text in the script edit control }
    property Lines : TStrings read FLines write SetLines;
    { Specifies common text options for script edit }
    property TextOptions : TvgrSETextOptions read FTextOptions write FTextOptions;
    { Represents the horizontal scroll bar for the script edit control }
    property HorzScrollBar : TvgrSEScrollBar read FHorzScrollBar;
    { Represents the vertical scroll bar for the script edit control }
    property VertScrollBar : TvgrSEScrollBar read FVertScrollBar;
{Specifies the script options: language, etc.
See also:
  TvgrScriptSyntax}
    property Script: TvgrScriptSyntax read FScript write SetScript;
    { Use brief cursor shapes }
    property BriefCursorShapes : Boolean read FBriefCursorShapes write FBriefCursorShapes;
    { Specifies the maximum number of characters the user can enter into the single line of script edit. }
    property MaxLineLength : Integer read FMaxLineLength write FMaxLineLength;
    { Determines whether the edit in Insert or Overwrite mode
      TvgrSEInsertMode = (vmmInsert, vmmOverwrite) }
    property EditorMode : TvgrSEInsertMode read FEditorMode write SetEditorMode;
    { Specifies whether the edit mode can be switched by pressing on Insert key. }
    property AutoEditorMode : Boolean read FAutoEditorMode write FAutoEditorMode default True;
    { Specifies the number of spaces in TAB character. }
    property TabIncrement : Integer read FTabIncrement write SetTabIncrement default 4;
    { Determines whether the script edit control has a single line border around the client area. }
    property BorderStyle: TBorderStyle read FBorderStyle write SetBorderStyle default bsSingle;
    { Determines whether characters typed in the edit control are converted from ANSI to OEM and then back to ANSI. }
    property OEMConvert: Boolean read FOEMConvert write FOEMConvert;
    { Determines whether the user can change the text of the script edit control.}
    property ReadOnly: Boolean read FReadOnly write FReadOnly;
    { Determines whether the memo control has scroll bars.
      Possble values: ssNone, ssHorizontal, ssVertical, ssBoth }
    property ScrollBars: TScrollStyle read FScrollBars write SetScrollBars default ssBoth;
    { Determines whether the user can insert tab characters into the text. }
    property WantTabs: Boolean read FWantTabs write FWantTabs;
    { Determines whether the user can insert return characters into the text. }
    property WantReturns: Boolean read FWantReturns write FWantReturns;
    { Specifies whether the gutter is visible }
    property GutterVisible : Boolean read GetGutterVisible write SetGutterVisible;
    { TAB symbols replaces with spaces automatically }
    property ReplaceTabsWithSpaces: Boolean read FReplaceTabsWithSpaces write FReplaceTabsWithSpaces;
    { True, if selection is Active }
    property SelectionActive: Boolean read GetSelectionActive;
    { position of the caret }
    property CaretPos: TPoint read GetCaretPos;
    { occurs then caret is moving }
    property OnCaretMove: TvgrCaretMoveEvent read FOnCaretMove write FOnCaretMove;
    {Occurs when the text for the edit control may have changed. }
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    {Occurs when the editor mode for the edit control is switched. }
    property OnEditorModeChange: TNotifyEvent read FOnEditorModeChange write FOnEditorModeChange;
  public
    { Creates an instance of TvgrCustomScriptEdit. }
    constructor Create(AOwner : TComponent); override;
    { Destroys the instance of a script edit object}
    destructor Destroy; override;
    { Do a syntax checking in script edit }
    procedure CheckSyntax;
    { Use to generate OnCaretMove event }
    procedure CaretMove;
    procedure SetCaretPos(ACol, ARow: Integer);
    procedure ClipBoardCopy;
    procedure ClipBoardPaste;
    procedure ClipBoardCut;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptEdit
  //
  /////////////////////////////////////////////////
  { TvgrScriptEdit implements a multiline edit control for scripts editing. }
  TvgrScriptEdit = class(TvgrCustomScriptEdit)
  public
    { position of the caret }
    property CaretPos;
  published
    { Determines how the control aligns within its container (parent control). }
    property Align;
    { Specifies how the control is anchored to its parent. }
    property Anchors;
    { Specifies which edges of the control are beveled. }
    property BevelEdges;
    { Specifies the control’s bevel style. }
    property BevelKind;
    { Determines the style of the inner bevel of a script edit. }
    property BevelInner;
    { Determines the style of the outer bevel of a script edit. }
    property BevelOuter;
    { Determines whether the script edit control has a single line border around the client area. }
    property BorderStyle;
    { Use brief cursor shapes }
    property BriefCursorShapes;
    { Specifies the size constraints for the control. }
    property Constraints;
    { Determines whether a control has a three-dimensional (3-D) or two-dimensional look. }
    property Ctl3D;
    { Specifies the image used to represent the mouse pointer when it passes into the region covered by the control. }
    property Cursor;
    { Indicates the image used to represent the mouse pointer when the control is being dragged. }
    property DragCursor;
    { Specifies whether the control is being dragged normally or for docking. }
    property DragKind;
    { Determines how the control initiates drag-and-drop or drag-and-dock operations. }
    property DragMode;
    { Controls whether the control responds to mouse, keyboard, and timer events. }
    property Enabled;
    { Specifies whether the gutter is visible }
    property GutterVisible;
    { Contains the individual lines of text in the script edit control }
    property Lines;
    { Determines whether the user can change the text of the script edit control.}
    property ReadOnly;
    { TAB symbols replaces with spaces automatically }
    property ReplaceTabsWithSpaces;
    { Determines whether the memo control has scroll bars.
      Possble values: ssNone, ssHorizontal, ssVertical, ssBoth }
    property ScrollBars;
{Specifies the script options: language, etc.
See also:
  TvgrScriptSyntax}
    property Script;
    { Determines whether the control displays a Help Hint when the mouse pointer rests momentarily on the control. }
    property ShowHint;
    { Contains the text string that can appear when the user moves the mouse over the control. }
    property Hint;
    { Indicates the position of the control in its parent's tab order. }
    property TabOrder;
    { Determines if the user can tab to a control. }
    property TabStop;
    { Specifies the amount of spaces by which the caret moves when the user press a TAB key.}
    property TabIncrement;
    { Specifies common text options for script edit }
    property TextOptions;
    { Determines whether the user can insert tab characters into the text. }
    property WantTabs;
    { Determines whether the user can insert return characters into the text. }
    property WantReturns;
    { True, if selection is Active }
    property SelectionActive;
    { Determines whether the edit in Insert or Overwrite mode
      TvgrSEInsertMode = (vmmInsert, vmmOverwrite) }
    property EditorMode;
    { occurs then caret is moving }
    property OnCaretMove;
    {Occurs when the text for the edit control may have changed. }
    property OnChange;
    {Occurs when the editor mode for the edit control is switched. }
    property OnEditorModeChange;

    {Occurs when the user clicks the edit area of control.}
    property OnClick;
    { Occurs when the user right-clicks the control or otherwise invokes the popup menu (such as using the keyboard).}
    property OnContextPopup;
    { Occurs when the user double-clicks the left mouse button when the mouse pointer is over the edit area of control. }
    property OnDblClick;
    { Occurs when the user drops an object being dragged.  }
    property OnDragDrop;
    { Occurs when the user drags an object over a control. }
    property OnDragOver;
    { Occurs when the dragging of an object ends, either by docking the object or by canceling the dragging. }
    property OnEndDock;
    { Occurs when the dragging of an object ends, either by dropping the object or by canceling the dragging.}
    property OnEndDrag;
    { Occurs when the user presses a mouse button with the mouse pointer over a control. }
    property OnMouseDown;
    { Occurs when the user moves the mouse pointer while the mouse pointer is over a control. }
    property OnMouseMove;
    { Occurs when the user releases a mouse button that was pressed with the mouse pointer over a component. }
    property OnMouseUp;
    { Occurs when the user begins to drag a control with a DragKind of dkDock. }
    property OnStartDock;
    { Occurs when the user begins to drag the control or an object it contains by left-clicking on the control and holding the mouse button down. }
    property OnStartDrag;
    { Occurs when another control is docked to the control. }
    property OnDockDrop;
    { Occurs when another control is dragged over the control. }
    property OnDockOver;
    { Occurs when a control receives the input focus. }
    property OnEnter;
    { Occurs when the input focus shifts away from one control to another.  }
    property OnExit;
    { Occurs when a user presses any key while the control has focus. }
    property OnKeyDown;
    { Occurs when key pressed. }
    property OnKeyPress;
    { Occurs when the user releases a key that has been pressed. }
    property OnKeyUp;
    { Occurs when the application tries to undock a control that is docked to the windowed control. }
    property OnUnDock;
  end;

implementation

uses
  vgr_Functions;
  
{$R *.res}

/////////////////////////////////////////////////
//
// TvgrSETextElement
//
/////////////////////////////////////////////////

procedure TvgrSETextElement.SetTextAttrs(
  AStyle : TFontStyles;
  AColor: TColor;
  ADefaultColor: Boolean);
begin
  FFontStyle := AStyle;
  FColor := AColor;
  FDefaultColor := ADefaultColor;
end;

function TvgrSETextElement.GetColor : TColor;
begin
  if DefaultColor then
    Result := clWindowText
  else
    Result := FColor;
end;

/////////////////////////////////////////////////
//
// TvgrSETextColors
//
/////////////////////////////////////////////////
constructor TvgrSETextColors.Create;
begin
  FWhitespace := TvgrSETextElement.Create;
  FKeyword    := TvgrSETextElement.Create;
  FComment    := TvgrSETextElement.Create;
  FNonsource  := TvgrSETextElement.Create;
  FOperator   := TvgrSETextElement.Create;
  FNumber     := TvgrSETextElement.Create;
  FString     := TvgrSETextElement.Create;
  FFunction   := TvgrSETextElement.Create;
  FIdentifier := TvgrSETextElement.Create;
  SetDefaultSheme;
end;

destructor TvgrSETextColors.Destroy;
begin
  FWhitespace.Free;
  FKeyword.Free;
  FComment.Free;
  FNonsource.Free;
  FOperator.Free;
  FNumber.Free;
  FString.Free;
  FFunction.Free;
  FIdentifier.Free;
  inherited;
end;

procedure TvgrSETextColors.SetDefaultSheme;
begin
  FWhitespace.SetTextAttrs([],clBlack,True);
  FComment.SetTextAttrs([fsItalic],clBlue,False);
  FKeyword.SetTextAttrs([fsBold],clBlack,False);
  FNonsource.SetTextAttrs([],clBlack,True);
  FOperator.SetTextAttrs([],clGreen,False);
  FNumber.SetTextAttrs([],clMaroon,False);
  FString.SetTextAttrs([],clNavy,False);
  FFunction.SetTextAttrs([],clMaroon,False);
  FIdentifier.SetTextAttrs([],clBlack,False);
end;

/////////////////////////////////////////////////
//
// TvgrSETextOptions
//
/////////////////////////////////////////////////
constructor TvgrSETextOptions.Create(Memo : TvgrSEMemo);
begin
  FTextColors := TvgrSETextColors.Create;
  FMemo := Memo;
  FMemo.TextOptions := Self;
  FontName := 'Courier New';
  FontSize := 10;
  BkColor := clWindow;
  FErrorBackColor := clMaroon;
  FErrorForeColor := clWhite;
end;

destructor TvgrSETextOptions.Destroy;
begin
  FTextColors.Free;
  inherited
end;

procedure TvgrSETextOptions.SetFontName(Value : TFontName);
begin
  FFontName := Value;
  Change;
end;

procedure TvgrSETextOptions.SetErrorBackColor(Value: TColor);
begin
  FErrorBackColor := Value;
  Change;
end;

procedure TvgrSETextOptions.SetErrorForeColor(Value: TColor);
begin
  FErrorForeColor := Value;
  Change;
end;

procedure TvgrSETextOptions.SetFontSize(Value : Integer);
begin
  FFontSize := Value;
  Change;
end;

procedure TvgrSETextOptions.Change;
begin
  FMemo.Change;
end;
/////////////////////////////////////////////////
//
// TvgrSEScrollBar
//
/////////////////////////////////////////////////
constructor TvgrSEScrollBar.Create(AOwner : TComponent);
begin
  inherited;
  ControlStyle := [csDoubleClicks, csOpaque, csNoDesignVisible];
  FScriptEdit := TvgrCustomScriptEdit(AOwner);
  TabStop := False;
end;

procedure TvgrSEScrollBar.CreateParams(var Params: TCreateParams);
const
  Kinds: array[TScrollBarKind] of DWORD = (SBS_HORZ, SBS_VERT);
begin
  inherited CreateParams(Params);
  CreateSubClass(Params, 'SCROLLBAR');
  Params.Style := Params.Style or Kinds[Kind] or SBS_LEFTALIGN;
end;

procedure TvgrSEScrollBar.CreateWnd;
begin
  inherited CreateWnd;
end;

procedure TvgrSEScrollBar.SetParams(APosition, AMin, AMax, APageSize : Integer);
begin
  if PageSize > (AMax - 1) then
    PageSize := AMax - 1;
  inherited SetParams(APosition, AMin, AMax);
  if PageSize <> Math.Min(APageSize, AMax-1) then
    PageSize := Math.Min(APageSize, AMax-1);
end;

procedure TvgrSEScrollBar.Scroll(ScrollCode: TScrollCode; var ScrollPos: Integer);
begin
  if (Self = FScriptEdit.HorzScrollBar) and (ScrollPos <> FScriptEdit.HorzScrollBar.Position) then
    with FScriptEdit.FMemo do
      SetVisibleTextOffsetWithoutChangingCaretPos(Point(ScrollPos, VisibleTextOffset.Y));
  if (Self = FScriptEdit.VertScrollBar) and (ScrollPos <> FScriptEdit.VertScrollBar.Position) then
    with FScriptEdit.FMemo do
      SetVisibleTextOffsetWithoutChangingCaretPos(Point(VisibleTextOffset.X, ScrollPos));
  inherited;
end;

/////////////////////////////////////////////////
//
// TvgrSEControlSwitcher
//
/////////////////////////////////////////////////
constructor TvgrSEControlSwitcher.Create(AOwner : TComponent);
begin
  inherited;
  TabStop := False;
  Width := 8;
  Color := clBtnFace;
  FScriptEdit := TvgrCustomScriptEdit(AOwner);
  FKind := sbVertical;
  Parent := TWinControl(AOwner);
end;

procedure TvgrSEControlSwitcher.ClearBackground(AColor : TColor);
begin
  with Canvas do
  begin
    Brush.Color := AColor;
    FillRect(ClipRect);
  end;
end;

const
  cArr_l1 = 0;
  cArr_l2 = 4;
  cEdge_l1 = 30;
  cEdge_l2 = 3;

procedure TvgrSEControlSwitcher.DrawEdge;
var
  APos : Integer;
begin
  with Canvas, Canvas.Pen do
  begin
    Style := psSolid;
    Width := 1;
    Color := clWindowText;
    if Kind = sbVertical then
    begin
      APos := Self.Height div 2;
      MoveTo(0, APos - cEdge_l1);
      LineTo(cEdge_l2, APos - cEdge_l1);
      LineTo(cEdge_l2+cEdge_l2, APos - cEdge_l1 + cEdge_l2);
      LineTo(cEdge_l2+cEdge_l2, APos + cEdge_l1 - cEdge_l2);
      LineTo(cEdge_l2, APos + cEdge_l1);
      LineTo(0, APos + cEdge_l1);
      Color := clWindow;
      MoveTo(0, APos - cEdge_l1 + 1);
      LineTo(cEdge_l2 - 1, APos - cEdge_l1 + 1);
      LineTo(cEdge_l2+cEdge_l2 - 1, APos - cEdge_l1 + cEdge_l2 + 1);
//      MoveTo(cEdge_l2+cEdge_l2, APos + cEdge_l1 - cEdge_l2 + 1);
//      LineTo(cEdge_l2, APos + cEdge_l1 + 1);
//      LineTo(0, APos + cEdge_l1+1);
    end
    else
    begin
      APos := Self.Width div 2;

      Color := clWindow;
      MoveTo(APos - cEdge_l1 + 1, cEdge_l2 + cEdge_l2);
      LineTo(APos - cEdge_l1 + 1, cEdge_l2 + 1);
      LineTo(APos - cEdge_l1 + cEdge_l2 + 1, 1);
      LineTo(APos + cEdge_l1 - cEdge_l2 - 1, 1);
//      LineTo(APos + cEdge_l1 + 1, cEdge_l2 - 1);
//      LineTo(APos + cEdge_l1 + 1, cEdge_l2 + cEdge_l2 + 1);
      Color := clWindowText;
      MoveTo(APos - cEdge_l1, cEdge_l2 + cEdge_l2 + 1);
      LineTo(APos - cEdge_l1, cEdge_l2);
      LineTo(APos - cEdge_l1 + cEdge_l2, 0);
      LineTo(APos + cEdge_l1 - cEdge_l2, 0);
      LineTo(APos + cEdge_l1, cEdge_l2);
      LineTo(APos + cEdge_l1, cEdge_l2 + cEdge_l2 + 1);

    end;
  end;
end;

procedure TvgrSEControlSwitcher.DrawPup(X,Y : Integer);
begin
  with Canvas, Canvas.Pen do
  begin
    Width := 2;
    Color := clWindow;
    Rectangle(X, Y, X + 1, Y + 1);
//    Color := clInactiveCaption;
    Color := clWindowText;
    Rectangle(X + 1, Y + 1, X + 2, Y + 2);
  end;
end;

procedure TvgrSEControlSwitcher.DrawArrow;
var
  APos : Integer;
  APBeginX, APBeginY : Integer;
begin
  with Canvas, Canvas.Pen do
  begin
    Brush.Color := clWindow;
    Brush.Style := bsSolid;
    Color := clWindowText;
    Style := psSolid;
    Width := 1;
    if Kind = sbVertical then
    begin
      APos := Self.Height div 2;
      if ControlVisible then
      Polygon([Point(cArr_l1 + cArr_l2, APos - cArr_l2),
        Point(cArr_l1, APos),
        Point(cArr_l1 + cArr_l2, APos + cArr_l2)])
      else
      Polygon([Point(cArr_l1, APos - cArr_l2),
        Point(cArr_l1 + cArr_l2, APos),
        Point(cArr_l1, APos + cArr_l2)]);
      APBeginX := (cArr_l1 + cArr_l2) div 2;
      APBeginY := APos - (cEdge_l1 div 2);
      DrawPup(APBeginX, APBeginY);
      APBeginY := APos + (cEdge_l1 div 2);
      DrawPup(APBeginX, APBeginY);
    end
    else
    begin
      APos := Self.Width div 2;
      if not ControlVisible then
      Polygon([Point(APos - cArr_l2, cArr_l1 + cArr_l2 + 2),
        Point(APos, cArr_l1 + 2),
        Point(APos + cArr_l2, cArr_l1 + cArr_l2+ 2)])
      else
      Polygon([Point(APos - cArr_l2, cArr_l1 + 2),
        Point(APos, cArr_l1 + cArr_l2 + 2),
        Point(APos + cArr_l2, cArr_l1 + 2)]);
      APBeginY := (cArr_l1 + cArr_l2) div 2 + 1;
      APBeginX := APos - (cEdge_l1 div 2);
      DrawPup(APBeginX, APBeginY);
      APBeginX := APos + (cEdge_l1 div 2);
      DrawPup(APBeginX, APBeginY);
    end;
  end;
end;

function TvgrSEControlSwitcher.GetControlVisible: Boolean;
begin
  Result := FControl.Visible;
end;

procedure TvgrSEControlSwitcher.SetControl(Value: TControl);
begin
  FControl := Value;
  if Value <> nil then
    ControlVisible := False;
end;

function TvgrSEControlSwitcher.CreatePimpochkaRgn: THandle;
var
  APos : Integer;
  Points : array [0..5] of TPoint;
begin
  if Kind = sbVertical then
  begin
    APos := Height div 2;
    Points[0] := Point(0, APos - cEdge_l1);
    Points[1] := Point(cEdge_l2 + 1, APos - cEdge_l1 );
    Points[2] := Point(cEdge_l2+ cEdge_l2 + 1, APos - cEdge_l1 + cEdge_l2);
    Points[3] := Point(cEdge_l2+ cEdge_l2 + 1, APos + cEdge_l1 - cEdge_l2 + 2);
    Points[4] := Point(cEdge_l2 + 1, APos + cEdge_l1 + 2);
    Points[5] := Point(0, APos + cEdge_l1 + 2);
  end
  else
  begin
    APos := Width div 2;
    Points[0] := Point(APos - cEdge_l1 + cEdge_l2 + 1, 0);
    Points[1] := Point(APos - cEdge_l1, cEdge_l2 - 1);
    Points[2] := Point(APos - cEdge_l1, cEdge_l2+ cEdge_l2 + 2);
    Points[3] := Point(APos + cEdge_l1, cEdge_l2+ cEdge_l2 + 2);
    Points[4] := Point(APos + cEdge_l1 + 2, cEdge_l2 - 1 );
    Points[5] := Point(APos + cEdge_l1 - cEdge_l2 + 2, 0);
  end;
  Result := CreatePolygonRgn(Points,6,WINDING);
end;

procedure TvgrSEControlSwitcher.DrawPimpochka(ABackground : TColor; AlwaysDraw: Boolean);
var
  ARegion : THandle;
begin
  if (FCurrColor <> ABackground) or AlwaysDraw then
  begin
    ARegion := CreatePimpochkaRgn;
    SelectClipRgn(Canvas.Handle,ARegion);
    ClearBackground(ABackGround);
    DrawArrow;
    DrawEdge;
    SelectClipRgn(Canvas.Handle,0);
    DeleteObject(ARegion);
//    CloseHandle(ARegion);
    FCurrColor := ABackground;
  end;
end;

procedure TvgrSEControlSwitcher.Paint;
begin
  ClearBackground(Self.Color);
  DrawPimpochka(Self.Color, True);
end;

procedure TvgrSEControlSwitcher.MouseMove(Shift: TShiftState; X,Y: Integer);
var
  ARegion : THandle;
begin
  ARegion := CreatePimpochkaRgn;
  if PtInRegion(ARegion,X,Y) then
  begin
    DrawPimpochka(clHighlight, False);
    SetCapture(Handle);
  end
  else
  begin
    DrawPimpochka(Self.Color, False);
    ReleaseCapture;
  end;
  DeleteObject(ARegion)
end;

procedure TvgrSEControlSwitcher.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ControlVisible := not ControlVisible;
  Invalidate;
end;

procedure TvgrSEControlSwitcher.SetControlVisible(Value: Boolean);
begin
  FControl.Visible := Value;
//  FScriptEdit.Realign;
  FScriptEdit.Invalidate;
end;

procedure TvgrSEControlSwitcher.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

/////////////////////////////////////////////////
//
// TvgrSEGutter
//
/////////////////////////////////////////////////
constructor TvgrSEGutter.Create(AOwner : TComponent);
begin
  inherited;
  FNumerate := True;
  TabStop := False;
  Width := 35;
  Color := clBtnFace;
  ControlStyle := ControlStyle + [csNoDesignVisible];
  FScriptEdit := TvgrCustomScriptEdit(AOwner);
  Parent :=TvgrCustomScriptEdit(AOwner);
end;

// protected
procedure TvgrSEGutter.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
end;

function TvgrSEGutter.GetSEScrollBar: TvgrSEScrollBar;
begin
  Result := nil;
end;

function TvgrSEGutter.GetScriptEdit : TvgrCustomScriptEdit;
begin
  Result := FScriptEdit;
end;

{function TvgrSEGutter.GetHeaderRect(From,Size: Integer): TRect;
begin
end;

function TvgrSEGutter.GetVisibleRowsCount: Integer;
begin
  Result := 0;
end;

procedure TvgrSEGutter.PaintHeaderBorder(const HeaderRect: TRect);
begin
end;

procedure TvgrSEGutter.DoOnFontChange(Sender: TObject);
begin
end;

procedure TvgrSEGutter.DoScroll(OldPos,NewPos: Integer);
begin
end;
}
procedure TvgrSEGutter.Paint;
begin
  PaintLines(True);
end;

// private
procedure TvgrSEGutter.ApplyFontForLine(ALine: Integer);
begin
  with ScriptEdit.FMemo do
    with Self.Canvas.Font do
    begin
      Name := TextOptions.FontName;
      Size := TextOptions.FontSize;
      Pitch := fpFixed;
      if ALine = AbsoluteCaretPos.Y then
      begin
        Style := [fsBold];
        Color := clWindowText;
      end
      else
      begin
        Style := [];
        Color := clWindowText;
      end;
    end;
end;

procedure TvgrSEGutter.DrawLineBackGround(const ALineRect: TRect);
begin
  with Canvas, Canvas.Brush do
  begin
    Style := bsSolid;
    Color := Self.Color;
    FillRect(ALineRect);
  end;
  with Canvas, Canvas.Pen, ALineRect do
  begin
    Style := psSolid;
    Color := clWhite;
    MoveTo(Right-2,Top);
    LineTo(Right-2,Bottom);
    Color := clGray;
    MoveTo(Right-1,Top);
    LineTo(Right-1,Bottom);
  end;
end;

procedure TvgrSEGutter.PaintLine(ALine: Integer; APaintBackground: Boolean);
var
  AYPos, AXPos : Integer;
  ALineNumStr : string;
  ANumRect : Trect;
begin
  with ScriptEdit.FMemo do
  begin
    if (ALine >= GetFirstVisibleLine) and (ALine <= (GetLastVisibleLine + 1)) then
    begin
      AYPos := (ALine - GetFirstVisibleLine) * GetLineHeight;
      AXPos := 2;
      if APaintBackground then
        DrawLineBackGround(Rect(0, AYpos, Self.Width, AYpos + GetLineHeight));
      ApplyFontForLine(ALine);
      if Numerate then
      begin
        if ALine <= Lines.Count - 1 then
          ALineNumStr := '  ' +IntToStr(ALine + 1)
        else
          ALineNumStr := '    ';

        ANumRect := Rect(AXPos, AYpos, Self.Width - 4, AYpos + GetLineHeight);
        DrawText(Self.Canvas.Handle,PChar(ALineNumStr),Length(ALineNumStr), ANumRect, DT_RIGHT);
      end;
    end;
  end;
end;

procedure TvgrSEGutter.PaintLines(APaintBackground: Boolean);
var
  ALine : Integer;
  AYPos, AXPos : Integer;
  ALineNumStr : string;
  ANumRect : Trect;
begin
  with ScriptEdit.FMemo do
  begin
    AYPos := 0;
    AXPos := 2;
    for ALine := GetFirstVisibleLine to GetLastVisibleLine + 1 do
    begin
      if APaintBackground then
        DrawLineBackGround(Rect(0, AYpos, Self.Width, AYpos + GetLineHeight));
      ApplyFontForLine(ALine);
      if Numerate then
      begin
        if ALine <= Lines.Count - 1 then
          ALineNumStr := '  ' +IntToStr(ALine + 1)
        else
          ALineNumStr := '    ';

        ANumRect := Rect(AXPos, AYpos, Self.Width - 4, AYpos + GetLineHeight);
        DrawText(Self.Canvas.Handle,PChar(ALineNumStr),Length(ALineNumStr), ANumRect, DT_RIGHT);
//        Self.Canvas.TextOut(AXpos, AYPos, ALineNumStr);
      end;
      Inc(AYPos,GetLineHeight);
    end;
  end;
end;

procedure TvgrSEGutter.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

/////////////////////////////////////////////////
//
// TvgrSESyntaxPanel
//
/////////////////////////////////////////////////
constructor TvgrSESyntaxPanel.Create(AOwner : TComponent);
begin
  inherited;
  ControlStyle := ControlStyle + [csNoDesignVisible];
  Parent := TWinControl(AOwner);
  FScriptEdit := TvgrScriptEdit(AOwner);

  FImageList := TImageList.Create(nil);
  FImageList.Width := 16;
  FImageList.Height := 16;
  FImageList.ResourceLoad(rtBitmap, sSyntaxCheckBitmap, clWhite);
  FImageList.ResourceLoad(rtBitmap, sSyntaxCheckBitmapSel, clWhite);

  FMouseIsPressed := False;
end;

destructor TvgrSESyntaxPanel.Destroy;
begin
  FImageList.Free;
  inherited;
end;

procedure TvgrSESyntaxPanel.Paint;
begin
  FillBackground;
  PaintSyntaxIcon;
  PaintErrorString;
end;

procedure TvgrSESyntaxPanel.MouseMove(Shift: TShiftState; X,Y: Integer);
begin
  MouseInCheckIconRect(X, Y, FMouseIsPressed);
  if FMouseInCheckIconRect then
    SetCapture(Handle)
  else
    ReleaseCapture;
end;

procedure TvgrSESyntaxPanel.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if not FMouseIsPressed and (Button = mbLeft) then
  begin
    FMouseIsPressed := True;
    if MouseInCheckIconRect(X, Y, FMouseIsPressed) then
      CheckSyntax;
    Invalidate;
  end;
end;

procedure TvgrSESyntaxPanel.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FMouseIsPressed and (Button = mbLeft) then
  begin
    FMouseIsPressed := False;
    MouseInCheckIconRect(X, Y, FMouseIsPressed);
    Invalidate;
  end;
end;

const
  px1 = 20;
  px2 = 5;
  px3 = 15;
  py1 = 6;
  py2 = 11;
  py3 = 24;
  poff = 3;

procedure TvgrSESyntaxPanel.CheckSyntax;
begin
  FScriptEdit.CheckSyntax;
end;

function TvgrSESyntaxPanel.GetSyntaxCheckIconRect: TRect;
begin
  Result := Rect(px2 - poff, py1 - poff, px1 + poff + 1, py3 + poff + 1);
end;

function TvgrSESyntaxPanel.MouseInCheckIconRect(X, Y: Integer; ALBPressed: Boolean): Boolean;
var
  ARect : TRect;
begin
  ARect := GetSyntaxCheckIconRect;
  Result := PtInRect(ARect,Point(X, Y));
  if Result <> FMouseInCheckIconRect then
  begin
    FMouseInCheckIconRect := Result;
    InvalidateRect(Handle, @ARect,  False);
  end;
end;

procedure TvgrSESyntaxPanel.FillBackground;
begin
  with Canvas, Canvas.Brush do
  begin
    Style := bsSolid;
    Color := clBtnFace;
    FillRect(ClipRect);
  end;
end;

procedure TvgrSESyntaxPanel.PaintSyntaxIcon;
var
  APicOffset: Word;
  AEdgeRect: TRect;
begin
 APicOffset := Integer(FMouseIsPressed and FMouseInCheckIconRect);
 if FMouseInCheckIconRect then
 begin
   AEdgeRect := Rect(3, 6, FImageList.Width + 8, FImageList.Height + 12);
   if FMouseIsPressed then
     DrawEdge(Canvas.Handle, AEdgeRect, BDR_SUNKENOUTER, BF_ADJUST or BF_RECT)
   else
     DrawEdge(Canvas.Handle, AEdgeRect, BDR_RAISEDINNER, BF_ADJUST or BF_RECT);
 end;
 FImageList.Draw(Canvas, 7+APicOffset,10+APicOffset, Ord(FMouseInCheckIconRect), True);
end;

procedure TvgrSESyntaxPanel.PaintErrorString;
var
  AErrorString: string;
begin
  with FScriptEdit do
  if FSyntaxErrorPresent then
  begin
    AErrorString := Format('Error Line: %d Pos: %d %s',[FSyntaxError.Line + 1, FSyntaxError.Position + 1, FSyntaxError.ErrorInfo.bstrDescription]);
  end;
  with Canvas do
  begin
    Brush.Color := clBtnFace;

    Font.Color := clWindowText;
//    Font.Name := FScriptEdit.TextOptions.FontName;
//    Font.Size := FScriptEdit.TextOptions.FontSize;

    TextOut(30,4,AErrorString);
  end;
end;

procedure TvgrSESyntaxPanel.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrSESyntaxPanel.CMMouseLeave(var Message: TMessage);
begin
  inherited;
  FMouseIsPressed := False;
  FMouseInCheckIconRect := False;
  Invalidate;
end;



/////////////////////////////////////////////////
//
// TvgrSESelection
//
/////////////////////////////////////////////////
constructor TvgrSESelection.Create(AMemo : TvgrSEMemo);
begin
  FMemo := AMemo;
end;

procedure TvgrSESelection.Reset;
begin
  FActive     := False;
  FBlockMode  := False;
  FStartPoint := Point(0,0);
end;

function TvgrSESelection.GetTopPoint : TPoint;
begin
  if not BlockMode then
    if (FStartPoint.Y < FMemo.AbsoluteCaretPos.Y) or
      ((FStartPoint.Y = FMemo.AbsoluteCaretPos.Y) and (FStartPoint.X < FMemo.AbsoluteCaretPos.X)) then
      Result := FStartPoint
    else
      Result := FMemo.AbsoluteCaretPos
  else
    begin
      if (FStartPoint.X < FMemo.AbsoluteCaretPos.X) then
        Result.X := FStartPoint.X
      else
        Result.X := FMemo.AbsoluteCaretPos.X;
      if (FStartPoint.Y < FMemo.AbsoluteCaretPos.Y) then
        Result.Y := FStartPoint.Y
      else
        Result.Y := FMemo.AbsoluteCaretPos.Y;
    end;
end;

function TvgrSESelection.GetEndPoint : TPoint;
begin
  if not BlockMode then
    if (FStartPoint.Y > FMemo.AbsoluteCaretPos.Y) or
      ((FStartPoint.Y = FMemo.AbsoluteCaretPos.Y) and (FStartPoint.X > FMemo.AbsoluteCaretPos.X)) then
      Result := FStartPoint
    else
      Result := FMemo.AbsoluteCaretPos
  else
    begin
      if (FStartPoint.X > FMemo.AbsoluteCaretPos.X) then
        Result.X := FStartPoint.X
      else
        Result.X := FMemo.AbsoluteCaretPos.X;
      if (FStartPoint.Y > FMemo.AbsoluteCaretPos.Y) then
        Result.Y := FStartPoint.Y
      else
        Result.Y := FMemo.AbsoluteCaretPos.Y;
    end;
end;

/////////////////////////////////////////////////
//
// TvgrSEStrings
//
/////////////////////////////////////////////////
constructor TvgrSEStrings.Create(AScriptEdit : TvgrCustomScriptEdit);
begin
  FScriptEdit := AScriptEdit;
  FUpdated := False;
end;

procedure TvgrSEStrings.LoadFromStream(Stream: TStream);
begin
//  BeginUpdate;
  inherited;
//  EndUpdate;
  FScriptEdit.GetTextAttributes;
  FScriptEdit.Invalidate;
end;

procedure TvgrSEStrings.Changed;
begin
  inherited;
//  if not FUpdated then
    FScriptEdit.TextChanged;
end;

procedure TvgrSEStrings.SetUpdateState(Updating: Boolean);
begin
  FUpdated := Updating;
end;

/////////////////////////////////////////////////
//
// TvgrSEMemo
//
/////////////////////////////////////////////////
constructor TvgrSEMemo.Create(AOwner : TComponent);
begin
  inherited;
  Parent := TvgrCustomScriptEdit(AOwner);
  TabStop := true;
  Align := alClient;
  Cursor := -4;
  FAbsoluteCaretPos := Point(0,0);
  FVisibleTextOffset := Point(0,0);
  FSelection := TvgrSESelection.Create(Self);
  Selection.Reset;
  FLMBDown := False;
  FUpdateState := True;
  FKeyboardFilter := TvgrKeyboardFilter.Create;
  InitKeyboardFilter;
end;

destructor TvgrSEMemo.Destroy;
begin
  FSelection.Free;
  FKeyboardFilter.Free;
  inherited;
end;

// protected
procedure TvgrSEMemo.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
end;

procedure TvgrSEMemo.Resize;
begin
end;

function TvgrSEMemo.GetLineHeight : Integer;
begin
  Result := GetLineTextHeight + GetInterlineOffset;
end;

function TvgrSEMemo.GetLineTextHeight : Integer;
begin
  Result := FLineTextHeight;
end;

function TvgrSEMemo.GetInterlineOffset : Integer;
begin
  Result := 2;
end;

function TvgrSEMemo.GetLeftMargin : Integer;
begin
  Result := 5;
end;

function TvgrSEMemo.GetFirstVisibleLine : Integer;
begin
  Result := VisibleTextOffset.Y;
end;

function TvgrSEMemo.GetFirstVisibleCharacter : Integer;
begin
  Result := VisibleTextOffset.X;
end;

function TvgrSEMemo.GetLastVisibleLine : Integer;
begin
  Result := GetFirstVisibleLine + GetVisibledLinesCount - 1;
end;

function TvgrSEMemo.GetVisibledLinesCount : Integer;
begin
  Result := Height div GetLineHeight;
end;

function TvgrSEMemo.GetLastVisibleCharacter : Integer;
begin
  Result := GetFirstVisibleCharacter + GetVisibledCharactersCount -  1;
end;

function TvgrSEMemo.GetVisibledCharactersCount : Integer;
begin
  Result := (Width - GetLeftMargin) div GetCharacterWidth;
end;

procedure TvgrSEMemo.ClearBkGroundInRect(ARect : TRect);
begin
  with Canvas do
  begin
    Brush.Color := TextOptions.BkColor;
    Brush.Style := bsSolid;
    FillRect(ARect);
  end;
end;

function TvgrSEMemo.CreateSelectedRgn(ALine: Integer; const ALineRect: TRect) : THandle;
var
  ALeft, ARight : Integer;
begin
  ALeft := 0;
  ARight := 0;
  if (Selection.TopPoint.Y <= ALine) then
  begin
    if (Selection.TopPoint.Y < ALine) and not Selection.BlockMode then
      ALeft := 0
    else
      ALeft := Max((Selection.TopPoint.X - GetFirstVisibleCharacter) * GetCharacterWidth, 0);
  end;
  if (Selection.EndPoint.Y >= ALine) then
  begin
    if (Selection.EndPoint.Y > ALine) and not Selection.BlockMode then
      ARight := ALineRect.Right
    else
      ARight := Max((Selection.EndPoint.X - GetFirstVisibleCharacter) * GetCharacterWidth, 0);
  end;
  Result := CreateRectRgn(ALeft + GetLeftMargin, ALineRect.Top, ARight + GetLeftMargin, ALineRect.Bottom);
end;

function TvgrSEMemo.LineSelectionPresent(ALine: Integer): Boolean;
begin
  Result := Selection.Active and (Selection.TopPoint.Y <= ALine) and (Selection.EndPoint.Y >= ALine)
end;

function TvgrSEMemo.FindLineAttr(ALine : Integer) : PWord;
var
  I : Integer;
begin
  Result := ScriptEdit.FTextAttributes;
  if Result <> nil then
  begin
    for I := 0 to ALine - 1 do
      Inc(Result,Length(GetLineTextWithoutTabs(I)) + 2);
  end;
end;

procedure TvgrSEMemo.ApplyFontByAttr(AAttrib: Word; AUseFontColor: Boolean);
  procedure SetFont(TextElement : TvgrSETextElement; AUseFontColor: Boolean);
  begin
    with Canvas.Font do
    begin
      Style := TextElement.FontStyle;
      if AUseFontColor then
        Color := TextElement.Color;
    end;
  end;

begin
  with TextOptions.TextColors do
  case AAttrib of
    SOURCETEXT_ATTR_KEYWORD         :  SetFont(AttrKeyword, AUseFontColor);
    SOURCETEXT_ATTR_COMMENT	        :  SetFont(AttrComment, AUseFontColor);
    SOURCETEXT_ATTR_NONSOURCE       :  SetFont(AttrNonsource, AUseFontColor);
    SOURCETEXT_ATTR_OPERATOR	      :  SetFont(AttrOperator, AUseFontColor);
    SOURCETEXT_ATTR_NUMBER	        :  SetFont(AttrNumber, AUseFontColor);
    SOURCETEXT_ATTR_STRING	        :  SetFont(AttrString, AUseFontColor);
    SOURCETEXT_ATTR_FUNCTION_START	:  SetFont(AttrFunction, AUseFontColor);
    SOURCETEXT_ATTR_IDENTIFIER      :  SetFont(AttrIdentifier, AUseFontColor);
  else
    SetFont(AttrWhitespace, AUseFontColor);
  end;
end;

procedure TvgrSEMemo.DrawLineText(ATextPosY: Integer; APAttr : PWord; const ALineRect: TRect; const ATextToOut: string; ABrushColor,AFontColor: TColor; AUseFontColor : Boolean);
var
  ACharPosX : Integer;
  AText : string;
  I : Integer;
  AAttrib : Word;
begin
  with Canvas do
  begin
    Brush.Color := ABrushColor;
    Font.Color  := AFontColor;
    FillRect(ALineRect);
  end;
  if ATextToOut <> '' then
  begin
    ACharPosX := GetLeftMargin;
    AText := ATextToOut[1];
    if APAttr <> nil then
    begin
      AAttrib := APAttr^;
      for I := 2 to Length(ATextToOut) do
      begin
        Inc(APAttr);
        if APAttr^ = AAttrib then
        begin
          AText := AText + ATextToOut[I];
        end
        else
        with Canvas do
        begin
          ApplyFontByAttr(AAttrib, AUseFontColor);
          TextOut(ACharPosX,ATextPosY,AText);
          Inc(ACharPosX, Length(AText) * GetCharacterWidth);
          AText := ATextToOut[I];
          AAttrib := APAttr^;
        end;
      end;
      ApplyFontByAttr(AAttrib, AUseFontColor);
      Canvas.TextOut(ACharPosX,ATextPosY,AText);
    end
    else
      with Canvas do
      begin
        ApplyFontByAttr(SOURCETEXT_ATTR_NONSOURCE, AUseFontColor);
        AText := ATextToOut;
        TextOut(ACharPosX,ATextPosY,AText);
      end;
  end
end;

procedure TvgrSEMemo.PaintTextLine(ALine: Integer; ATextPosY: Integer; const ALineRect: TRect);
var
  ATextToOut: string;
  ALineRgn: THandle;
  ASelRgn : THandle;
  ASelPresent : Boolean;
  APAttr : PWord;
begin
    ASelRgn := 0;
    APAttr := FindLineAttr(ALine);
    if APAttr <> nil then
      Inc(APAttr, VisibleTextOffset.X);
    ALineRgn := CreateRectRgnIndirect(ALineRect);

    ASelPresent := LineSelectionPresent(ALine);
    if ASelPresent then
    begin
      ASelRgn := CreateSelectedRgn(ALine, ALineRect);
      CombineRgn(ALineRgn,ALineRgn,ASelRgn,RGN_DIFF);
    end;

    ATextToOut := Copy(StringReplace(Lines[ALine],#9,StringOfChar(' ',TabIncrement),[rfReplaceAll]),VisibleTextOffset.X + 1, GetVisibledCharactersCount + 2);

    SelectClipRgn(Canvas.Handle,ALineRgn);
    DrawLineText(ATextPosY, APattr, ALineRect, ATextToOut, TextOptions.BkColor, TextOptions.TextColors.AttrWhitespace.Color, True);

    if ASelPresent then
    begin
      SelectClipRgn(Canvas.Handle, ASelRgn);
      DrawLineText(ATextPosY, APAttr, ALineRect, ATextToOut, GetSelectionBrushColor, GetSelectionFontColor, False);
      DeleteObject(ASelRgn);
    end;
    SelectClipRgn(Canvas.Handle,0);
    DeleteObject(ALineRgn);
end;

procedure TvgrSEMemo.PaintText;
var
  Line, Y : Integer;
begin
  HideCaret(Handle);
  Canvas.Font.Name := TextOptions.FontName;
  Canvas.Font.Size := TextOptions.FontSize;
  Canvas.Font.Pitch := fpFixed;
  Line := GetFirstVisibleLine;
  Y := 0;
  while (Line < ScriptEdit.Lines.Count) and (Y < Height) do
  begin
    PaintTextLine(Line,Y,Rect(0,Y,Width,Y + GetLineHeight));
    Inc(Y,GetLineHeight);
    Inc(Line);
  end;
  if Y < Height then
    ClearBkGroundInRect(Rect(0,Y,Width,Height));
  ShowCaret(Handle);
end;

procedure TvgrSEMemo.PaintTextLines(ABegLine,AEndLine : Integer);
var
  Line, Y : Integer;
begin
  HideCaret(Handle);
  Canvas.Font.Name := TextOptions.FontName;
  Canvas.Font.Size := TextOptions.FontSize;
  Canvas.Font.Pitch := fpFixed;
  Line := GetFirstVisibleLine;
  Y := 0;
  while (Line < ScriptEdit.Lines.Count) and (Y < Height) do
  begin
    if (Line >= ABegLine) and (Line <= AEndLine) then
      PaintTextLine(Line,Y,Rect(0,Y,Width,Y + GetLineHeight));
    Inc(Y,GetLineHeight);
    Inc(Line);
  end;
  if Y < Height then
    ClearBkGroundInRect(Rect(0,Y,Width,Height));
  ShowCaret(Handle);
end;

procedure TvgrSEMemo.PaintTextBlock(ALeft, ATop, ARight, ABottom : Integer);
var
  Line, Y : Integer;
  ALeftEdge, ARightEdge : Integer;
begin
  ALeftEdge := (ALeft - GetFirstVisibleCharacter) * GetCharacterWidth + GetLeftMargin;
  ARightEdge := (ARight - GetFirstVisibleCharacter) * GetCharacterWidth + GetLeftMargin;
  HideCaret(Handle);
  Canvas.Font.Name := TextOptions.FontName;
  Canvas.Font.Size := TextOptions.FontSize;
  Canvas.Font.Pitch := fpFixed;
  Line := GetFirstVisibleLine;
  Y := 0;
  while (Line < ScriptEdit.Lines.Count) and (Y < Height) do
  begin
    if (Line >= ATop) and (Line <= ABottom) then
      PaintTextLine(Line,Y,Rect(ALeftEdge,Y,ARightEdge,Y + GetLineHeight));
    Inc(Y,GetLineHeight);
    Inc(Line);
  end;
  ShowCaret(Handle);
end;

procedure TvgrSEMemo.Paint;
begin
  PaintText;
end;

procedure TvgrSEMemo.KeyDown(var Key: Word; Shift: TShiftState);
begin
  if Assigned(ScriptEdit.OnKeyDown) then
    ScriptEdit.OnKeyDown(ScriptEdit, Key, Shift);
  if not ReadOnly then
    FKeyboardFilter.KeyDown(Key, Shift);
end;

procedure TvgrSEMemo.KeyUp(var Key: Word; Shift: TShiftState);
begin
  if Assigned(ScriptEdit.OnKeyUp) then
    ScriptEdit.OnKeyUp(ScriptEdit, Key, Shift);
  inherited;
end;

procedure TvgrSEMemo.KeyPress(var Key: Char);
begin
  if Assigned(ScriptEdit.OnKeyPress) then
    ScriptEdit.OnKeyPress(ScriptEdit, Key);
  inherited;
  if not ReadOnly then
    FKeyboardFilter.KeyPress(Key);
end;

procedure TvgrSEMemo.MouseMove(Shift: TShiftState; X,Y: Integer);
begin
  if Assigned(ScriptEdit.OnMouseMove) then
    ScriptEdit.OnMouseMove(ScriptEdit, Shift, X, Y);
  if FLMBDown then
  begin
    if not Selection.Active then
      StartSelection(ssAlt in Shift);
    SetCaretPosNear(X,Y);
    if (X < 0) or (X > Width) or (Y < 0) or (Y > Width) then
    begin
      FMMShift := Shift;
      FMMPosX := X;
      FMMPosY := Y;
      PostMessage(Handle,CM_MOUSECHECK,0,0);
    end;
  end;
end;

procedure TvgrSEMemo.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ABegLine, AEndLine: Integer;
  ASelectionWasActive: Boolean;
begin
  if Assigned(ScriptEdit.OnMouseDown) then
    ScriptEdit.OnMouseDown(ScriptEdit, Button, Shift, X, Y);
  if Button = mbLeft then
  begin
    SetFocus;
    ASelectionWasActive := Selection.Active;
    ABegLine := 0;
    AEndLine := 0;
    if ASelectionWasActive then
    begin
      ABegLine := Selection.TopPoint.Y;
      AEndLine := Selection.EndPoint.Y;
      Selection.Reset;
    end;
    SetCaretPosNear(X,Y);
    SetCaretToPosition;
    if ASelectionWasActive then
    begin
      ABegLine := Min(ABegLine, Selection.TopPoint.Y);
      AEndLine := Max(AEndLine, Selection.EndPoint.Y);
      ChangeTextLines(ABegLine, AEndLine);
    end;
    FLMBDown := True;
  end;
end;

procedure TvgrSEMemo.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FLMBDown := False;
  if Assigned(ScriptEdit.OnMouseUp) then
    ScriptEdit.OnMouseUp(ScriptEdit, Button, Shift, X, Y);
end;

procedure TvgrSEMemo.Click;
begin
  inherited;
  if Assigned(ScriptEdit.OnClick) then
    ScriptEdit.OnClick(ScriptEdit);
end;

procedure TvgrSEMemo.DblClick;
begin
  inherited;
  if Assigned(ScriptEdit.OnDblClick) then
    ScriptEdit.OnDblClick(ScriptEdit);
end;

//private
function TvgrSEMemo.GetSelectionBrushColor: TColor;
begin
  Result := clHighlight;
end;

function TvgrSEMemo.GetSelectionFontColor: TColor;
begin
  Result := clHighlightText;
end;

function TvgrSEMemo.GetErrorLineBrushColor: TColor;
begin
  Result := ScriptEdit.FTextOptions.FErrorBackColor;
end;

function TvgrSEMemo.GetErrorLineFontColor: TColor;
begin
  Result := ScriptEdit.FTextOptions.FErrorForeColor;
end;

procedure TvgrSEMemo.StartSelection(ABlockMode: Boolean);
begin
  with Selection do
  begin
    Active := True;
    StartPoint := AbsoluteCaretPos;
    BlockMode := ABlockMode;
  end;
end;

function TvgrSEMemo.SymbolIsDelimiter(Value : Char) : Boolean;
begin
  Result := not (Value in ['0'..'9','a'..'z','A'..'Z'])
end;

procedure TvgrSEMemo.SetCaretPosToAbsolute(Char, Line  : Integer);
var
  ACaretPos : TPoint;
begin
  ACaretPos := AbsoluteCaretPos;
  if Lines.Count = 0 then
    ACaretPos.Y := 0
  else
  if Line < Lines.Count  then
    ACaretPos.Y := Line + GetFirstVisibleLine
  else
    ACaretPos.Y := Lines.Count - 1;
  ACaretPos.X := Char + GetFirstVisibleCharacter;
  if ACaretPos.Y < 0 then
    ACaretPos.Y := 0;
  if ACaretPos.X < 0 then
    ACaretPos.X := 0;
  AbsoluteCaretPos := ACaretPos;
//  SetCaretToPosition;
end;

procedure TvgrSEMemo.MoveCaretToClientRect;
begin
  if AbsoluteCaretPos.Y < GetFirstVisibleLine then
    FAbsoluteCaretPos := Point(AbsoluteCaretPos.X, GetFirstVisibleLine);
  if AbsoluteCaretPos.Y > GetLastVisibleLine then
    FAbsoluteCaretPos := Point(AbsoluteCaretPos.X, GetLastVisibleLine);
end;

procedure TvgrSEMemo.CheckSelectionMode(AShiftPressed, AAltPressed: Boolean);
begin
  if (AShiftPressed) and (not Selection.Active) then
    StartSelection(AAltPressed)
  else
  if (not (AShiftPressed)) and (Selection.Active) then
  begin
    Selection.Reset;
    ChangeText;
  end;
end;

procedure TvgrSEMemo.ChangeSelectedText(AShiftPressed, AAltPressed: Boolean);
begin
  if Selection.Active then
    ChangeText;
end;

function TvgrSEMemo.GetCharacterWidth : Integer;
begin
  Result := FCharacterWidth;
end;

procedure TvgrSEMemo.SetCaretToPosition;
var
  AYOffset : Integer;
begin
  if IsUpdated then
    if Focused then
    begin
      if ScriptEdit.BriefCursorShapes then
        AYOffset := GetLineTextHeight
      else
        AYOffset := 0;
      SetCaretPos(
        (FAbsoluteCaretPos.X - GetFirstVisibleCharacter) * GetCharacterWidth + GetLeftMargin,
        (FAbsoluteCaretPos.Y - GetFirstVisibleLine) * GetLineHeight + AYOffset);
      ScriptEdit.CaretMove;
      ScrollToCaret;
    end;
end;

procedure TvgrSEMemo.SetCaretToPositionWithoutScrolling;
var
  AXpos, AYOffset : Integer;

begin
  if Focused then
  begin
    if ScriptEdit.BriefCursorShapes then
      AYOffset := GetLineTextHeight
    else
      AYOffset := 0;
    AXpos := (FAbsoluteCaretPos.X - GetFirstVisibleCharacter) * GetCharacterWidth + GetLeftMargin;
    if AXpos >= GetLeftMargin then
      SetCaretPos(AXPos, (FAbsoluteCaretPos.Y - GetFirstVisibleLine) * GetLineHeight + AYOffset)
    else
      SetCaretPos(-100, (FAbsoluteCaretPos.Y - GetFirstVisibleLine) * GetLineHeight + AYOffset);
    ScriptEdit.CaretMove;
  end;
end;

procedure TvgrSEMemo.ScrollToCaret;
begin
  if AbsoluteCaretPos.Y > GetLastVisibleLine then
    VisibleTextOffset := Point(VisibleTextOffset.X, FAbsoluteCaretPos.Y - GetVisibledLinesCount + 1);
  if AbsoluteCaretPos.Y < GetFirstVisibleLine then
    VisibleTextOffset := Point(VisibleTextOffset.X, FAbsoluteCaretPos.Y);
  if AbsoluteCaretPos.X > GetLastVisibleCharacter then
    VisibleTextOffset := Point(FAbsoluteCaretPos.X - GetVisibledCharactersCount + 1, VisibleTextOffset.Y) ;
  if AbsoluteCaretPos.X < GetFirstVisibleCharacter then
    VisibleTextOffset := Point(FAbsoluteCaretPos.X, VisibleTextOffset.Y);
end;

function TvgrSEMemo.MoveCaretPosToLeft: Boolean;
begin
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X - 1, AbsoluteCaretPos.Y);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretPosToRight: Boolean;
begin
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X + 1, AbsoluteCaretPos.Y);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretPosToUp: Boolean;
begin
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X, AbsoluteCaretPos.Y - 1);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretPosToDown: Boolean;
begin
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X, AbsoluteCaretPos.Y + 1);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToNextPage: Boolean;
var
  CaretOffset : Integer;
begin
  CaretOffset := GetVisibledLinesCount - 1;
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X, AbsoluteCaretPos.Y + CaretOffset);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToPrevPage: Boolean;
var
  CaretOffset : Integer;
begin
  CaretOffset := GetVisibledLinesCount - 1;
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X, AbsoluteCaretPos.Y - CaretOffset);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToBeginOfLine: Boolean;
begin
  AbsoluteCaretPos := Point(0,AbsoluteCaretPos.Y);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToEndOfLine: Boolean;
begin
  AbsoluteCaretPos := Point(LengthAbs(Lines[FAbsoluteCaretPos.Y]), AbsoluteCaretPos.Y);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToBeginOfLines: Boolean;
begin
  AbsoluteCaretPos := Point(0,0);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToEndOfLines: Boolean;
begin
  AbsoluteCaretPos := Point(LengthAbs(Lines[FAbsoluteCaretPos.Y]), Lines.Count - 1);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToTopOfPage: Boolean;
begin
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X, GetFirstVisibleLine);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveCaretToBottomOfPage: Boolean;
begin
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X, GetLastVisibleLine);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.ScrollPageLineDown: Boolean;
begin
  if GetFirstVisibleLine < (Lines.Count - 1) then
  begin
    Inc(FVisibleTextOffset.Y);
    MoveCaretToClientRect;
    PaintText;
    ScriptEdit.UpdateGutter;
    ScriptEdit.UpdateScrollBars;
  end;
  SetCaretToPositionWithoutScrolling;
  Result := True;
end;

function TvgrSEMemo.ScrollPageLineUp: Boolean;
begin
  if GetFirstVisibleLine > 0 then
  begin
    Dec(FVisibleTextOffset.Y);
    MoveCaretToClientRect;
    PaintText;
    ScriptEdit.UpdateGutter;
    ScriptEdit.UpdateScrollBars;
  end;
  SetCaretToPositionWithoutScrolling;
  Result := True;
end;

function TvgrSEMemo.WheelLineUp: Boolean;
begin
  if GetFirstVisibleLine > 0 then
  begin
    Dec(FVisibleTextOffset.Y);
    PaintText;
    ScriptEdit.UpdateGutter;
    ScriptEdit.UpdateScrollBars;
  end;
  SetCaretToPositionWithoutScrolling;
  Result := True;
end;

function TvgrSEMemo.WheelLineDown: Boolean;
begin
  if GetFirstVisibleLine < (Lines.Count - 1) then
  begin
    Inc(FVisibleTextOffset.Y);
    PaintText;
    ScriptEdit.UpdateGutter;
    ScriptEdit.UpdateScrollBars;
  end;
  SetCaretToPositionWithoutScrolling;
  Result := True;
end;

function TvgrSEMemo.WheelPageUp: Boolean;
begin
  if GetFirstVisibleLine > 0 then
  begin
    Dec(FVisibleTextOffset.Y, GetVisibledLinesCount-1);
    if FVisibleTextOffset.Y < 0 then
      FVisibleTextOffset.Y := 0;
    PaintText;
    ScriptEdit.UpdateGutter;
    ScriptEdit.UpdateScrollBars;
  end;
  SetCaretToPositionWithoutScrolling;
  Result := True;
end;

function TvgrSEMemo.WheelPageDown: Boolean;
begin
  if GetFirstVisibleLine < (Lines.Count - 1) then
  begin
    Inc(FVisibleTextOffset.Y, GetVisibledLinesCount-1);
    if FVisibleTextOffset.Y >= (Lines.Count - 1) then
      FVisibleTextOffset.Y := (Lines.Count - 2);
    PaintText;
    ScriptEdit.UpdateGutter;
    ScriptEdit.UpdateScrollBars;
  end;
  SetCaretToPositionWithoutScrolling;
  Result := True;
end;

function TvgrSEMemo.MoveToNextWordInLine: Boolean;
var
  WasDelimiter : Boolean;
begin
  if FAbsoluteCaretPos.X < LengthAbs(Lines[FAbsoluteCaretPos.Y]) then
  begin
    WasDelimiter := False;
    while (FAbsoluteCaretPos.X < LengthAbs(Lines[FAbsoluteCaretPos.Y]))
      and ((not WasDelimiter) or SymbolIsDelimiter(StringReplace(Lines[FAbsoluteCaretPos.Y],#9,StringOfChar(' ',TabIncrement),[rfReplaceAll])[FAbsoluteCaretPos.X + 1])) do
    begin
      WasDelimiter := WasDelimiter or SymbolIsDelimiter(StringReplace(Lines[FAbsoluteCaretPos.Y],#9,StringOfChar(' ',TabIncrement),[rfReplaceAll])[FAbsoluteCaretPos.X + 1]);
      Inc(FAbsoluteCaretPos.X);
    end;
  end;
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveToPrevWordInLine: Boolean;
var
  WasCharacters : Boolean;
  ALineText : string;
begin
  if FAbsoluteCaretPos.X > 0 then
  begin
    WasCharacters := False;
    ALineText := StringReplace(Lines[FAbsoluteCaretPos.Y],#9,StringOfChar(' ',TabIncrement),[rfReplaceAll]);
    while (FAbsoluteCaretPos.X > 0)
      and ((not SymbolIsDelimiter(ALineText[FAbsoluteCaretPos.X]) or (not WasCharacters))) do
    begin
      WasCharacters := WasCharacters or not SymbolIsDelimiter(ALineText[FAbsoluteCaretPos.X]);
      Dec(FAbsoluteCaretPos.X);
    end;
  end;
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveToNextWord: Boolean;
begin
  if (FAbsoluteCaretPos.X >=  LengthAbs(Lines[FAbsoluteCaretPos.Y])) and (FAbsoluteCaretPos.Y < (Lines.Count - 1))  then
  begin
    FAbsoluteCaretPos.X := 0;
    Inc(FAbsoluteCaretPos.Y);
  end
  else
    MoveToNextWordInLine;
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.MoveToPrevWord: Boolean;
begin
  if (FAbsoluteCaretPos.X = 0) and (FAbsoluteCaretPos.Y > 0) then
  begin
    Dec(FAbsoluteCaretPos.Y);
    FAbsoluteCaretPos.X := LengthAbs(Lines[FAbsoluteCaretPos.Y]);
  end
  else
    MoveToPrevWordInLine;
  SetCaretToPosition;
  Result := True;
end;

procedure TvgrSEMemo.SetCaretPosNear(X,Y : Integer);
begin
  SetCaretPosToAbsolute(GetCharacterNear(X,Y),GetLineNumberNear(Y));
end;

function TvgrSEMemo.GetLineNumberNear(Y : Integer) : Integer;
begin
  Result := Y div GetLineHeight;
end;

function TvgrSEMemo.GetCharacterNear(X,Y : Integer) : Integer;
var
  CharacterWidth : Integer;
begin
  CharacterWidth := Canvas.TextWidth('A');
  Result := (X - GetLeftMargin) div CharacterWidth;
end;

procedure TvgrSEMemo.SetCaret;
begin
  if ScriptEdit.BriefCursorShapes then
    CreateCaret(Handle,0,GetCharacterWidth,2)
  else
    CreateCaret(Handle,0,2,GetLineHeight);
  SetCaretToPositionWithoutScrolling;
  ShowCaret(Handle);
end;

procedure TvgrSEMemo.ResetCaret;
begin
  DestroyCaret;
end;

procedure TvgrSEMemo.DeleteNextCharAbs(var AText: string; APosition: Integer);
var
  APos: Integer;
  ACharPos: Integer;
  AIncrement: Integer;
begin
  APos := 1;
  ACharPos := 1;
  while (APos < APosition) and (ACharPos <= Length(AText)) do
  begin
    if AText[ACharPos] = Chr(VK_TAB) then
      AIncrement := TabIncrement
    else
      AIncrement := 1;
    Inc(APos, AIncrement);
    Inc(ACharPos);
  end;
  if (APos = APosition) then
    Delete(AText, ACharPos, 1)
  else
  begin
    Delete(AText, ACharPos-1, 1);
    Insert(StringOfChar(' ',TabIncrement), AText, ACharPos-1);
    Delete(AText, ACharPos, 1);
  end;
end;

procedure TvgrSEMemo.DeletePrevCharAbs(var AText: string; APosition: Integer; out ACharIsTab: Boolean);
var
  APos: Integer;
  ACharPos: Integer;
  AIncrement: Integer;
begin
  APos := 1;
  ACharPos := 1;
  ACharIsTab := False;
  while (APos < APosition) and (ACharPos <= Length(AText)) do
  begin
    if AText[ACharPos] = Chr(VK_TAB) then
    begin
      AIncrement := TabIncrement;
      ACharIsTab := True;
    end
    else
    begin
      AIncrement := 1;
      ACharIsTab := False;
    end;
    Inc(APos, AIncrement);
    Inc(ACharPos);
  end;
  if (APos = APosition) then
  begin
    Delete(AText, ACharPos, 1);
    ACharIsTab := False;
  end
  else
    if (APos > APosition) then
    begin
      Delete(AText, ACharPos-1, 1);
      if not ACharIsTab then
      begin
        Insert(StringOfChar(' ',TabIncrement), AText, ACharPos-1);
        Delete(AText, ACharPos, 1);
      end;
    end;
end;

function TvgrSEMemo.KeyPressCR: Boolean;
var
  ACurLine : string;
  ANewLine : string;
begin
  AddBlankLineIfNeed;
  ACurLine := Lines[AbsoluteCaretPos.Y];
  if EditorMode = vmmInsert then
  begin
    ANewLine := Copy(ACurLine, 1, AbsoluteCaretPos.X);
    Lines[AbsoluteCaretPos.Y] := Copy(ACurLine, AbsoluteCaretPos.X + 1, Length(ACurLine));
    Lines.Insert(AbsoluteCaretPos.Y,ANewLine);
    AbsoluteCaretPos := Point(0,AbsoluteCaretPos.Y + 1);
    ChangeTextLines(AbsoluteCaretPos.Y-1, GetLastVisibleLine + 1);
    ScriptEdit.Gutter.PaintLine(Lines.Count - 1, False);
  end
  else
  begin
    if AbsoluteCaretPos.Y = (Lines.Count - 1) then
    begin
      Lines.Add('');
    end;
    AbsoluteCaretPos:= Point(0,AbsoluteCaretPos.Y + 1);
    SetCaretToPosition;
  end;
  Result := True;
end;

function TvgrSEMemo.KeyPressDelete: Boolean;
var
  ACurLine : string;
  AddSpaces : Integer;
begin
  if not RemoveSelectedText then
  begin
    ACurLine := Lines[AbsoluteCaretPos.Y];
    if AbsoluteCaretPos.X >= LengthAbs(ACurLine) then
    begin
      if AbsoluteCaretPos.Y < (Lines.Count - 1) then
      begin
        AddSpaces := AbsoluteCaretPos.X - LengthAbs(ACurLine);
        Lines[AbsoluteCaretPos.Y] := ACurLine + StringOfChar(' ', AddSpaces) + Lines[AbsoluteCaretPos.Y + 1];
        Lines.Delete(AbsoluteCaretPos.Y + 1);
        ChangeTextLines(AbsoluteCaretPos.Y, GetLastVisibleLine + 1);
        ScriptEdit.UpdateGutter;
      end;
    end
    else
    begin
      DeleteNextCharAbs(ACurLine,AbsoluteCaretPos.X + 1);
      Lines[AbsoluteCaretPos.Y] := ACurLine;
      ChangeTextLines(AbsoluteCaretPos.Y, AbsoluteCaretPos.Y);
    end;
  end;
  Result := True;
end;

function TvgrSEMemo.KeyPressBackspace: Boolean;
var
  ACurLine: string;
  ANewX: Integer;
  ATabDelete: Boolean;
begin
  if not RemoveSelectedText then
  begin
    ACurLine := Lines[AbsoluteCaretPos.Y];
    if AbsoluteCaretPos.X = 0 then
    begin
      if AbsoluteCaretPos.Y > 0 then
      begin
        ANewX := LengthAbs(Lines[AbsoluteCaretPos.Y - 1]);
        Lines[AbsoluteCaretPos.Y - 1] := Lines[AbsoluteCaretPos.Y - 1] + ACurLine;
        Lines.Delete(AbsoluteCaretPos.Y);
        AbsoluteCaretPos := Point(ANewX,AbsoluteCaretPos.Y - 1);
        ChangeTextLines(AbsoluteCaretPos.Y-1, GetLastVisibleLine + 1);
        ScriptEdit.UpdateGutter;
      end;
    end
    else
    begin
      DeletePrevCharAbs(ACurLine,AbsoluteCaretPos.X, ATabDelete);
      if ATabDelete then
        ANewX := AbsoluteCaretPos.X - TabIncrement
      else
        ANewX := AbsoluteCaretPos.X - 1;
      Lines[AbsoluteCaretPos.Y] := ACurLine;
      AbsoluteCaretPos := Point(ANewX,AbsoluteCaretPos.Y);
      ChangeTextLines(AbsoluteCaretPos.Y,AbsoluteCaretPos.Y);
    end;
  end;
  Result := True;
end;

function TvgrSEMemo.KeyPressTab: Boolean;
var
  ATextToInserting: string;
  ALine : Integer;
  ALineText : string;
begin
  if not RemoveSelectedText then
    if EditorMode = vmmInsert then
    begin
      AddBlankLineIfNeed;
      ALine := AbsoluteCaretPos.Y;
      ALineText := Lines[ALine];
      ATextToInserting := GetTabString;
      InsertTextInLine(ATextToInserting,ALineText,AbsoluteCaretPos.X + 1);
      Lines[ALine] := ALineText;
      ChangeTextLines(ALine,ALine);
    end;
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X + TabIncrement, AbsoluteCaretPos.Y);
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.KeyPressIndent: Boolean;
var
  ALine : Integer;
  ALineText : string;
begin
  AddBlankLineIfNeed;
  if not Selection.Active then
  begin
    ALine := AbsoluteCaretPos.Y;
    ALineText := GetTabString + Lines[ALine];
    Lines[ALine] := ALineText;
    ChangeTextLines(ALine,ALine);
    AbsoluteCaretPos := Point(AbsoluteCaretPos.X + TabIncrement, AbsoluteCaretPos.Y);
  end
  else
  begin
    ALine := Selection.TopPoint.Y;
    while ALine <= Selection.EndPoint.Y do
    begin
      if (AbsoluteCaretPos.X > 1) or (Selection.EndPoint.Y = Selection.StartPoint.Y) or (ALine < Selection.EndPoint.Y) then
      begin
        ALineText := GetTabString + Lines[ALine];
        Lines[ALine] := ALineText;
        if ALine = AbsoluteCaretPos.Y then
          AbsoluteCaretPos := Point(AbsoluteCaretPos.X + TabIncrement, AbsoluteCaretPos.Y);
      end;
      Inc(ALine);
    end;
    Selection.StartPoint := Point(Selection.StartPoint.X + TabIncrement, Selection.StartPoint.Y);
    ChangeTextLines(Selection.TopPoint.Y, Selection.EndPoint.Y);
  end;
  SetCaretToPosition;
  Result := True;
end;

function TvgrSEMemo.KeyPressUnindent: Boolean;
var
  ALine : Integer;

  procedure UnindentLine(AIndex: Integer);
  var
    ALineText : string;
    ARemovedSpaces: Integer;
    ARemovedChar: Char;
  begin
    ALineText := Lines[AIndex];
    if Length(ALineText) > 0 then
    begin
      if ALineText[1] = Char(VK_TAB) then
      begin
        Delete(ALineText,1,1);
        ARemovedSpaces := TabIncrement;
      end
      else
      begin
        ARemovedSpaces := 0;
        while (ARemovedSpaces < TabIncrement)
          and ((ALineText[1] = ' ') or (ALineText[1] = #9)) do
        begin
          ARemovedChar := ALineText[1];
          Delete(ALineText,1,1);
          Inc(ARemovedSpaces);
          if ARemovedChar = #9 then
            ALineText := StringOfChar(' ', TabIncrement - 1) + ALineText;
        end;
      end;
      Lines[AIndex] := ALineText;
      ChangeTextLines(AIndex,AIndex);
      if AIndex = AbsoluteCaretPos.Y then
        AbsoluteCaretPos := Point(AbsoluteCaretPos.X - ARemovedSpaces, AbsoluteCaretPos.Y);
      SetCaretToPosition;
    end;
  end;

begin
  AddBlankLineIfNeed;
  if not Selection.Active then
    UnIndentLine(AbsoluteCaretPos.Y)
  else
  begin
    ALine := Selection.TopPoint.Y;
    while ALine <= Selection.EndPoint.Y do
    begin
      if not ((ALine = Selection.EndPoint.Y) and (AbsoluteCaretPos.X = 0) and (Selection.EndPoint.Y = AbsoluteCaretPos.Y)) then
        UnindentLine(ALine);
      Inc(ALine);
    end;
    Selection.StartPoint := Point(Selection.StartPoint.X - TabIncrement, Selection.StartPoint.Y);
    ChangeTextLines(Selection.TopPoint.Y, Selection.EndPoint.Y);
  end;
  Result := True;
end;

function TvgrSEMemo.GetLineTextWithoutTabs(ALine: Integer): string;
begin
  if (Lines.Count > 0) then
    Result := StringReplace(Lines[ALine],#9,StringOfChar(' ',TabIncrement),[rfReplaceAll])
  else
    Result := '';
end;

procedure TvgrSEMemo.SelectNearestWord;
type
  TvgrLineWord = record
    StartPos: Integer;
    EndPos: Integer;
    Offset: Integer;
  end;

var
  ALine: Integer;
  I: Integer;
  ALineText : string;
  AWordBegin: Integer;
  AWords: array of TvgrLineWord;
  AMinOffset: Integer;

  function CalcWordOffset(AWordBegin, AWordEnd, APosition: Integer): Integer;
  begin
    if (APosition <= (AWordEnd-2)) and (APosition >= AWordBegin) then
      Result := 0
    else
      Result := Min(Abs(APosition - AWordBegin),Abs(APosition - (AWordEnd-2)));
  end;

  function EndOfWord: Boolean;
  begin
    Result := (AWordBegin > 0) and ((I >= Length(ALineText)) or SymbolIsDelimiter(ALineText[I]));
  end;

  procedure FindWords;
  var
    AElement: Integer;
  begin
    I := 0;
    AWordBegin := 0;
    AMinOffset := MaxInt;
    repeat
      Inc(I);
      if (AWordBegin = 0) and (I <= Length(ALineText)) and not SymbolIsDelimiter(ALineText[I]) then
        AWordBegin := I;
      if EndOfWord then
      begin
        AElement := Length(AWords);
        SetLength(AWords, AElement+1);
        with AWords[AElement] do
        begin
          StartPos := AWordBegin;
          if (I = Length(ALineText)) and not SymbolIsDelimiter(ALineText[I]) then
            EndPos := I
          else
            EndPos := I - 1;

          Offset := CalcWordOffset(AWordBegin, I, AbsoluteCaretPos.X);
          AMinOffset := Min(AMinOffset, Offset);
          AWordBegin := 0;
        end;
      end;
    until I >= Length(ALineText);
  end;

begin
  ALine := AbsoluteCaretPos.Y;
  ALineText := GetLineTextWithoutTabs(ALine);
  FindWords;
  if Length(AWords) > 0 then
  begin
    I := 0;
    while (I < Length(AWords)) and (AMinOffset <> AWords[I].Offset) do
      Inc(I);
    Selection.Active := True;
    Selection.StartPoint := Point(AWords[I].StartPos-1, ALine);
    AbsoluteCaretPos := Point(AWords[I].EndPos, ALine);
  end;
end;

function TvgrSEMemo.GetTabString: string;
begin
  if ReplaceTabsWithSpaces then
    Result := StringOfChar(' ',TabIncrement)
  else
    Result := Chr(VK_TAB);
end;

procedure TvgrSEMemo.AddBlankLineIfNeed;
begin
  if Lines.Count = 0 then
    Lines.Add('');
end;

procedure TvgrSEMemo.DrawSyntaxErrorLine(ALine: Integer);
var
  ACurLine,
  AYPos: Integer;
  APAttr: PWORD;
begin
  HideCaret(Handle);
  Canvas.Font.Name := TextOptions.FontName;
  Canvas.Font.Size := TextOptions.FontSize;
  ACurLine := GetFirstVisibleLine;
  AYPos := 0;
  APAttr := FindLineAttr(ALine);

  while (ACurLine < ScriptEdit.Lines.Count) and (ACurLine <= ALine) and (AYPos < Height) do
  begin
    if (ACurLine = ALine) then
      DrawLineText(AYPos, APAttr, Rect(0,AYPos,Width,AYPos + GetLineHeight), Lines[ALine], GetErrorLineBrushColor, GetErrorLineFontColor, False);
    Inc(AYPos,GetLineHeight);
    Inc(ACurLine);
  end;
  ShowCaret(Handle);
end;

function TvgrSEMemo.CopyAbs(const AText: string; AIndex, ALength: Integer): string;
var
  ACharPos, APos: Integer;
  AMaxCharPos: Integer;
begin
  Result := '';
  ACharPos := 1;
  APos := 1;
  AMaxCharPos := AIndex+ALength;
  if AMaxCharPos < 0 then
    AMaxCharPos := MaxInt;
  while (APos < (AMaxCharPos)) and (ACharPos <= Length(Atext)) do
  begin
    if APos >= AIndex then
      Result := Result + AText[ACharPos];
    Inc(ACharPos);
    Inc(APos);
  end;
end;

function TvgrSEMemo.LengthAbs(const Value: string): Integer;
var
  ACharPos : Integer;
begin
  ACharPos := 0;
  Result := 0;
  while ACharPos < Length(Value) do
  begin
    Inc(ACharPos);
    if Value[ACharPos] = Chr(VK_TAB) then
      Inc(Result, TabIncrement)
    else
      Inc(Result);
  end;
end;

function TvgrSEMemo.RemoveSelectedText: Boolean;
var
  ANewCaretPos: TPoint;
  ANumOfLines: Integer;
  ALine: Integer;
  ALineText: string;
  I : Integer;
begin
  AddBlankLineIfNeed;
  Result := Selection.Active;
  if Selection.Active then
  begin
    ANewCaretPos := Selection.TopPoint;
    ANumOfLines := Selection.EndPoint.Y - Selection.TopPoint.Y;
    ALine := Selection.TopPoint.Y;
    if not Selection.BlockMode then
    begin
      ALineText := CopyAbs(Lines[ALine], 1, Selection.TopPoint.X);
      ALineText := ALineText + CopyAbs(Lines[ALine + ANumOfLines], Selection.EndPoint.X + 1, Length(Lines[ALine + ANumOfLines]));
      for I := 1 to ANumOfLines do
        Lines.Delete(ALine+1);
      Lines[ALine] := ALineText;
    end
    else
    begin
      for I := ALine to ALine + ANumOfLines do
      begin
        ALineText :=  Lines[I];
        Delete(ALineText,Selection.TopPoint.X + 1, Selection.EndPoint.X - Selection.TopPoint.X);
        Lines[I] := ALineText;
      end;
    end;
    Selection.Reset;
    AbsoluteCaretPos := ANewCaretPos;
    if ANumOfLines > 0 then
    begin
      ChangeTextLines(ALine, GetLastVisibleLine + 1);
      ScriptEdit.Gutter.PaintLines(True);
    end
    else
      ChangeTextLines(ALine, ALine);
  end;
end;

procedure TvgrSEMemo.RecalcFontMetrics;
var
  ADC : THandle;
  ATextMetric : TTextMetric;
  AFont : TFont;
  AEtalonFontName: TFontName;
  AEtalonFontSize: Integer;
begin
  if TextOptions = nil then
  begin
    AEtalonFontName := Font.Name;
    AEtalonFontSize := Font.Size;
  end
  else
  begin
    AEtalonFontName := TextOptions.FontName;
    AEtalonFontSize := TextOptions.FontSize;
  end;
  AFont := TFont.Create;
  with AFont do
  begin
    Name := AEtalonFontName;
    Size := AEtalonFontSize;
    Pitch := fpFixed;
    Style := [fsBold, fsItalic];
    ADC := GetDC(0);
    SelectObject(ADC,AFont.Handle);
  end;
  GetTextMetrics(ADC,ATextMetric);
  ReleaseDC(0,ADC);
  AFont.Free;

  FCharacterWidth := ATextMetric.tmAveCharWidth;
  FLineTextHeight := ATextMetric.tmHeight;
end;

procedure TvgrSEMemo.Change;
begin
  RecalcFontMetrics;
  Invalidate;
  SetCaretToPosition;
end;

procedure TvgrSEMemo.ChangeText;
begin
  PaintText;
  SetCaretToPosition;
  ScriptEdit.UpdateScrollBars;
end;

procedure TvgrSEMemo.ChangeTextLines(ABegLine, AEndLine : Integer);
begin
  if (ScriptEdit.FSyntaxErrorPresent) and ((ScriptEdit.FSyntaxError.Line < ABegLine) or (ScriptEdit.FSyntaxError.Line < AEndLine)) then
    PaintTextLines(ScriptEdit.FSyntaxError.Line, ScriptEdit.FSyntaxError.Line);
  PaintTextLines(ABegLine,AEndLine);
  SetCaretToPosition;
  if AEndLine <= GetLastVisibleLine then
    ScriptEdit.UpdateScrollBars;
end;

procedure TvgrSEMemo.ChangeTextBlock(ALeft, ATop, ARight, ABottom : Integer);
begin
  PaintTextBlock(ALeft, ATop, ARight + 1, ABottom);
  SetCaretToPosition;
  ScriptEdit.UpdateScrollBars;
end;

function TvgrSEMemo.GetScriptEdit: TvgrCustomScriptEdit;
begin
  Result := TvgrCustomScriptEdit(Parent);
end;

function TvgrSEMemo.GetLines : TStrings;
begin
  Result := ScriptEdit.Lines;
end;

function TvgrSEMemo.GetTextOptions : TvgrSETextOptions;
begin
  Result := ScriptEdit.TextOptions;
end;

procedure TvgrSEMemo.SetTextOptions(AValue: TvgrSETextOptions);
begin
  ScriptEdit.TextOptions := AValue;
end;

procedure TvgrSEMemo.SetVisibleTextOffset(Value: TPoint);
var
  OldValue : TPoint;
begin
  OldValue := FVisibleTextOffset;
  FVisibleTextOffset := Value;
  if not ((OldValue.X = Value.X) and (OldValue.Y = Value.Y)) then
  begin
    ChangeText;
  end;
  if not (OldValue.Y = Value.Y) then
    ScriptEdit.ChangeVisibleYOffset;
end;

procedure TvgrSEMemo.SetVisibleTextOffsetWithoutChangingCaretPos(Value: TPoint);
var
  OldValue : TPoint;
begin
  OldValue := FVisibleTextOffset;
  FVisibleTextOffset := Value;
  if not ((OldValue.X = Value.X) and (OldValue.Y = Value.Y)) then
  begin
    SetCaretToPositionWithoutScrolling;
    PaintText;
  end;
  if not (OldValue.Y = Value.Y) then
    ScriptEdit.ChangeVisibleYOffset;
end;

procedure TvgrSEMemo.SetAbsoluteCaretPos(Value: TPoint);
var
  OldValue : TPoint;
  OldSelectionTopPoint, OldSelectionEndPoint : TPoint;
  ABegLine, AEndLine, ABegChar, AEndChar : Integer;
begin
  OldValue := FAbsoluteCaretPos;
  OldSelectionTopPoint := Selection.TopPoint;
  OldSelectionEndPoint := Selection.EndPoint;
  FAbsoluteCaretPos := Value;
  if FAbsoluteCaretPos.X > ScriptEdit.MaxLineLength then
    FAbsoluteCaretPos.X := ScriptEdit.MaxLineLength;
  if FAbsoluteCaretPos.X < 0 then
    FAbsoluteCaretPos.X := 0;
  if FAbsoluteCaretPos.Y > (Lines.Count - 1) then
    FAbsoluteCaretPos.Y := Lines.Count - 1;
  if FAbsoluteCaretPos.Y < 0 then
    FAbsoluteCaretPos.Y := 0;

  if not ((OldValue.X = FAbsoluteCaretPos.X) and (OldValue.Y = FAbsoluteCaretPos.Y)) then
  begin
    if Selection.Active then
    begin
      if Selection.BlockMode then
      begin
        if Selection.EndPoint.Y <> OldSelectionEndPoint.Y then
        begin
          ABegLine := Min(Selection.EndPoint.Y, OldSelectionEndPoint.Y);
          AEndLine := Max(Selection.EndPoint.Y, OldSelectionEndPoint.Y);
          ABegChar := Min(Selection.TopPoint.X, OldSelectionTopPoint.X);
          AEndChar := Max(Selection.EndPoint.X, OldSelectionEndPoint.X);
          ChangeTextBlock(ABegChar, ABegLine, AEndChar, AEndLine);
        end;
        if Selection.TopPoint.Y <> OldSelectionTopPoint.Y then
        begin
          ABegLine := Min(Selection.TopPoint.Y, OldSelectionTopPoint.Y);
          AEndLine := Max(Selection.TopPoint.Y, OldSelectionTopPoint.Y);
          ABegChar := Min(Selection.TopPoint.X, OldSelectionTopPoint.X);
          AEndChar := Max(Selection.EndPoint.X, OldSelectionEndPoint.X);
          ChangeTextBlock(ABegChar, ABegLine, AEndChar, AEndLine);
        end;
        if Selection.EndPoint.X <> OldSelectionEndPoint.X then
        begin
          ABegLine := Selection.TopPoint.Y;
          AEndLine := Selection.EndPoint.Y;
          ABegChar := Min(Selection.EndPoint.X, OldSelectionEndPoint.X);
          AEndChar := Max(Selection.EndPoint.X, OldSelectionEndPoint.X);
          ChangeTextBlock(ABegChar, ABegLine, AEndChar, AEndLine);
        end;
        if Selection.TopPoint.X <> OldSelectionTopPoint.X then
        begin
          ABegLine := Selection.TopPoint.Y;
          AEndLine := Selection.EndPoint.Y;
          ABegChar := Min(Selection.TopPoint.X, OldSelectionTopPoint.X);
          AEndChar := Max(Selection.TopPoint.X, OldSelectionTopPoint.X);
          ChangeTextBlock(ABegChar, ABegLine, AEndChar, AEndLine);
        end;
      end
      else
      begin
        if (OldValue.Y > FAbsoluteCaretPos.Y) then
        begin
          ABegLine := FAbsoluteCaretPos.Y;
          AEndLine := OldValue.Y;
        end
        else
        begin
          ABegLine := OldValue.Y;
          AEndLine := FAbsoluteCaretPos.Y;
        end;
        ChangeTextLines(ABegLine,AEndLine);
      end;
    end;
    ScriptEdit.ChangeAbsoluteCaretPos(OldValue.Y, FAbsoluteCaretPos.Y);
  end;
end;

function  TvgrSEMemo.GetEditorMode: TvgrSEInsertMode;
begin
  Result := ScriptEdit.EditorMode;
end;

procedure TvgrSEMemo.SetEditorMode(Value: TvgrSEInsertMode);
begin
  ScriptEdit.EditorMode := Value;
end;

function TvgrSEMemo.GetSelectedText: string;
var
  I : Integer;
  AFirstLength: Integer;
  AEndLength: Integer;
begin
  with Selection do
  if Active then
  begin
    if TopPoint.Y = EndPoint.Y then
      Result := Copy(Lines[TopPoint.Y], TopPoint.X + 1, EndPoint.X - TopPoint.X)
    else
    begin
      if not BlockMode then
      begin
        AFirstLength  := LengthAbs(Lines[TopPoint.Y])  - TopPoint.X;
        AEndLength    := EndPoint.X;
        Result := Copy(Lines[TopPoint.Y], TopPoint.X + 1, AFirstLength)+#13#10;
        for I := TopPoint.Y + 1 to EndPoint.Y - 1 do
          Result := Result + Lines[I] + #13#10;
        Result := Result + Copy(Lines[EndPoint.Y], 1, AEndLength);
      end
      else
      begin
        AFirstLength  := EndPoint.X - TopPoint.X;
        AEndLength    := AFirstLength;
        Result := Copy(Lines[TopPoint.Y], TopPoint.X + 1, AFirstLength)+#13#10;
        for I := TopPoint.Y + 1 to EndPoint.Y - 1 do
        begin
          Result := Result + Lines[I] + #13#10;
        end;
        Result := Result + Copy(Lines[EndPoint.Y], 1, AEndLength);
      end;
    end;
  end
  else
    Result := '';
end;

function TvgrSEMemo.GetTabIncrement: Integer;
begin
  Result := ScriptEdit.TabIncrement;
end;

function TvgrSEMemo.GetReplaceTabsWithSpaces: Boolean;
begin
  Result := ScriptEdit.ReplaceTabsWithSpaces;
end;

procedure TvgrSEMemo.BeforeSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
begin
  AddBlankLineIfNeed;
  CheckSelectionMode(AShiftPressed, AAltPressed);
end;

procedure TvgrSEMemo.AfterSelectionCommand(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean);
begin
  ChangeSelectedText(AShiftPressed, AAltPressed);
end;

procedure TvgrSEMemo.InsertTextInLine(const ATextToInserting: string; var AText: string; APosition: Integer);
var
  APos: Integer;
  ACharPos: Integer;
  AIncrement: Integer;
begin
  APos := 1;
  ACharPos := 1;
  while (APos < APosition) and (ACharPos <= Length(AText)) do
  begin
    if AText[ACharPos] = Chr(VK_TAB) then
      AIncrement := TabIncrement
    else
      AIncrement := 1;
    Inc(APos, AIncrement);
    Inc(ACharPos);
  end;
  if (APos = APosition) then
    Insert(ATextToInserting, AText, ACharPos)
  else
  begin
    Delete(AText, ACharPos-1, 1);
    Insert(StringOfChar(' ',TabIncrement), AText, ACharPos-1);
    Insert(ATextToInserting, AText, ACharPos + APos - APosition);
  end;
end;

procedure TvgrSEMemo.InitKeyboardFilter;
begin
  with FKeyboardFilter do
  begin
    OnKeyPress := DoKeyPress;
    AddInstantCommands;
    AddDefLayerCommands;
    OnBeforeSelectionCommand := BeforeSelectionCommand;
    OnAfterSelectionCommand := AfterSelectionCommand;
  end;
end;

procedure TvgrSEMemo.AddInstantCommands;
begin
  with FKeyboardFilter do
  begin
    AddSelectionCommand(VK_LEFT, [],   MoveCaretPosToLeft);
    AddSelectionCommand(VK_RIGHT, [],   MoveCaretPosToRight);
    AddSelectionCommand(VK_UP, [],   MoveCaretPosToUp);
    AddSelectionCommand(VK_DOWN, [],   MoveCaretPosToDown);
    AddSelectionCommand(VK_NEXT, [], MoveCaretToNextPage);
    AddSelectionCommand(VK_PRIOR, [], MoveCaretToPrevPage);
    AddSelectionCommand(VK_HOME, [], MoveCaretToBeginOfLine);
    AddSelectionCommand(VK_END, [], MoveCaretToEndOfLine);

    AddSelectionCommand(VK_LEFT, [ssShift],   MoveCaretPosToLeft);
    AddSelectionCommand(VK_RIGHT, [ssShift],   MoveCaretPosToRight);
    AddSelectionCommand(VK_UP, [ssShift],   MoveCaretPosToUp);
    AddSelectionCommand(VK_DOWN, [ssShift],   MoveCaretPosToDown);
    AddSelectionCommand(VK_NEXT, [ssShift], MoveCaretToNextPage);
    AddSelectionCommand(VK_PRIOR, [ssShift], MoveCaretToPrevPage);
    AddSelectionCommand(VK_HOME, [ssShift], MoveCaretToBeginOfLine);
    AddSelectionCommand(VK_END, [ssShift], MoveCaretToEndOfLine);

    AddCommand(VK_RETURN, [], KeyPressCR);
    AddCommand(VK_DELETE, [], KeyPressDelete);
    AddCommand(VK_BACK, [], KeyPressBackspace);
    AddCommand(VK_TAB, [], KeyPressTab);
  end;
end;

procedure TvgrSEMemo.AddDefLayerCommands;
begin
  with FKeyboardFilter do
  begin
    AddSelectionCommand(VK_RETURN, [ssCtrl], KeyPressCR);
    AddSelectionCommand(VK_LEFT, [ssCtrl],   MoveToPrevWord);
    AddSelectionCommand(VK_RIGHT, [ssCtrl],   MoveToNextWord);
    AddSelectionCommand(VK_UP, [ssCtrl],   ScrollPageLineUp);
    AddSelectionCommand(VK_DOWN, [ssCtrl],   ScrollPageLineDown);
    AddSelectionCommand(VK_NEXT, [ssCtrl], MoveCaretToBottomOfPage);
    AddSelectionCommand(VK_PRIOR, [ssCtrl], MoveCaretToTopOfPage);
    AddSelectionCommand(VK_HOME, [ssCtrl], MoveCaretToBeginOfLines);
    AddSelectionCommand(VK_END, [ssCtrl], MoveCaretToEndOfLines);

    AddSelectionCommand(VK_RETURN, [ssCtrl, ssShift], KeyPressCR);
    AddSelectionCommand(VK_LEFT, [ssCtrl, ssShift],   MoveToPrevWord);
    AddSelectionCommand(VK_RIGHT, [ssCtrl, ssShift],   MoveToNextWord);
    AddSelectionCommand(VK_UP, [ssCtrl, ssShift],   ScrollPageLineUp);
    AddSelectionCommand(VK_DOWN, [ssCtrl, ssShift],   ScrollPageLineDown);
    AddSelectionCommand(VK_NEXT, [ssCtrl, ssShift], MoveCaretToBottomOfPage);
    AddSelectionCommand(VK_PRIOR, [ssCtrl, ssShift], MoveCaretToTopOfPage);
    AddSelectionCommand(VK_HOME, [ssCtrl, ssShift], MoveCaretToBeginOfLines);
    AddSelectionCommand(VK_END, [ssCtrl, ssShift], MoveCaretToEndOfLines);

    AddCommand(VK_CHAR_I,[ssCtrl], KeyPressIndent);
    AddCommand(VK_CHAR_U,[ssCtrl], KeyPressUnindent);
    AddCommand(VK_CHAR_V,[ssCtrl], ClipBoardPaste);
    AddCommand(VK_CHAR_C,[ssCtrl], ClipBoardCopy);
    AddCommand(VK_CHAR_X,[ssCtrl], ClipBoardCut);
    AddCommand(VK_INSERT,[ssShift], ClipBoardPaste);
    AddCommand(VK_INSERT,[ssCtrl], ClipBoardCopy);
    AddCommand(VK_DELETE,[ssCtrl], ClipBoardCut);
    AddCommand(VK_CHAR_A,[ssCtrl], SelectAll);
    AddCommand(VK_BACK, [ssCtrl], DeletePreviosWord);
  end;
end;

procedure TvgrSEMemo.DoKeyPress(AKey: Char);
var
  ALine : Integer;
  ALineText : string;
begin
  AddBlankLineIfNeed;
  RemoveSelectedText;
  ALine := AbsoluteCaretPos.Y;
  ALineText := Lines[ALine];
  if EditorMode = vmmOverwrite then
    DeleteNextCharAbs(ALineText,AbsoluteCaretPos.X + 1);
  if AbsoluteCaretPos.X > LengthAbs(ALineText) then
    ALineText := ALineText + StringOfChar(' ',AbsoluteCaretPos.X - LengthAbs(ALineText));
  InsertTextInLine(AKey,ALineText,AbsoluteCaretPos.X + 1);
  Lines[ALine] := ALineText;
  AbsoluteCaretPos := Point(AbsoluteCaretPos.X + 1, AbsoluteCaretPos.Y);
  ChangeTextLines(ALine,ALine);
end;

function TvgrSEMemo.ClipBoardCopy: Boolean;
begin
  Clipboard.AsText := SelectedText;
  Result := True;
end;

function TvgrSEMemo.ClipBoardPaste: Boolean;
var
  ALine, AChar, ASymbolFrom, ASymbolsCount: Integer;
  AFirstChangedLine: Integer;
  AFirstLine: Boolean;
  ALastLine: Boolean;
  ABeginLineText, AEndLineText: string;
  ATempText : string;
  AClipboardText: string;

  procedure FinalizeLine;
  begin
    if ASymbolFrom > 0 then
    begin
      ATempText := Copy(AClipboardText, ASymbolFrom, ASymbolsCount);
      if AFirstLine then
        ATempText := ABeginLineText + ATempText;
    end
    else
      ATempText := '';
    if ALastLine then
    begin
      AbsoluteCaretPos := Point(Length(ATempText), ALine);
      ATempText := ATempText + AEndLineText;
    end;
    Lines[ALine] := ATempText;
  end;

  procedure NewLine;
  begin
    Inc(ALine);
    Lines.Insert(ALine, '');
  end;

begin
  if Clipboard.HasFormat(CF_TEXT) then
  begin
    RemoveSelectedText;
    AddBlankLineIfNeed;
    ALine := AbsoluteCaretPos.Y;
    ABeginLineText := CopyAbs(Lines[ALine], 1, AbsoluteCaretPos.X);
    AEndLineText := CopyAbs(Lines[ALine], AbsoluteCaretPos.X+1, MaxInt);
    AClipboardText := Clipboard.AsText;
    ATempText := '';
    ASymbolFrom := 0;
    ASymbolsCount := 0;
    AFirstLine := True;
    ALastLine := False;
    AFirstChangedLine := 0;
    for AChar := 1 to Length(AClipboardText) do
    begin
      if AClipboardText[AChar] = #13 then
      begin
        FinalizeLine;
        if AFirstLine then
          AFirstChangedLine := ALine;
        NewLine;
        AFirstLine := False;
        ASymbolsCount := 0;
        ASymbolFrom := 0;
      end
      else
      if AClipboardText[AChar] <> #10 then
      begin
        if ASymbolFrom = 0 then
          ASymbolFrom := AChar;
        Inc(ASymbolsCount);
      end;
    end;
    ALastLine := True;
    FinalizeLine;
    if AFirstLine then
      ChangeTextLines(ALine, ALine)
    else
      ChangeTextLines(AFirstChangedLine, GetLastVisibleLine+1);
  end;
  ScriptEdit.UpdateScrollBars;
  Result := True;
end;

function TvgrSEMemo.ClipBoardCut: Boolean;
begin
  ClipBoardCopy;
  RemoveSelectedText;
  ScriptEdit.UpdateScrollBars;
  Result := True;
end;

function TvgrSEMemo.SelectAll: Boolean;
begin
  Selection.Active := True;
  Selection.BlockMode := False;
  Selection.StartPoint := Point(0,0);
  with Lines do
    AbsoluteCaretPos := Point(Length(Lines[Count-1]), Count - 1);
  Result :=True
end;

function TvgrSEMemo.DeletePreviosWord: Boolean;
var
  I : Integer;
  ALine: Integer;
  AFirstDeletedSymbol: Integer;
  ALastDeletedSymbol: Integer;
  ALineText: string;
  ABeginDeleteWord: Boolean;
  AFirstLineToInvalidate: Integer;
  ALastLineToInvalidate: Integer;
begin
  ALine := AbsoluteCaretPos.Y;
  ALineText := GetLineTextWithoutTabs(ALine);
  I := Min(Length(ALineText), AbsoluteCaretPos.X + 1);
  ABeginDeleteWord := False;
  AFirstDeletedSymbol := 0;
  ALastDeletedSymbol := I;
  while (I > 0) and not ((ABeginDeleteWord) and (SymbolIsDelimiter(ALineText[I]))) do
  begin
    if not ABeginDeleteWord then
    begin
      if not SymbolIsDelimiter(ALineText[I]) then
        ABeginDeleteWord := True;
    end;
    AFirstDeletedSymbol := I;
    Dec(I);
  end;

  if ALastDeletedSymbol > 0 then
  begin
    Delete(ALineText, AFirstDeletedSymbol, ALastDeletedSymbol - AFirstDeletedSymbol + 1);
    Lines[ALine] := ALineText;
    AbsoluteCaretPos := Point(AFirstDeletedSymbol - 1, ALine);
  end;

  if Selection.Active then
  begin
    AFirstLineToInvalidate := Selection.TopPoint.Y;
    ALastLineToInvalidate := Selection.EndPoint.Y;
  end
  else
  begin
    AFirstLineToInvalidate := ALine;
    ALastLineToInvalidate := ALine;
  end;
  Selection.Reset;
  ChangeTextLines(AFirstLineToInvalidate, ALastLineToInvalidate);

  Result := True;
end;

function TvgrSEMemo.IsUpdated: Boolean;
begin
  Result := FUpdateState;
end;

procedure TvgrSEMemo.MouseWheelHandler(var Message: TMessage);
var
  ADelta: Integer;
  AKeys: Word;
begin
  ADelta := Short(Message.WParamHi);
  AKeys := Short(Message.WParamLo);
  if (AKeys = 0) then
  begin
    if ADelta > 0 then
      WheelLineUp
    else
      WheelLineDown;
  end
  else
  if (AKeys = VGR_MK_SHIFT) then
  begin
    if ADelta > 0 then
      MoveCaretPosToUp
    else
      MoveCaretPosToDown;
  end
  else
  if (AKeys = VGR_MK_CONTROL) then
  begin
    if ADelta > 0 then
      WheelPageUp
    else
      WheelPageDown;
  end;
  Message.Result := 1;
end;

function TvgrSEMemo.GetReadOnly: Boolean;
begin
  Result := ScriptEdit.ReadOnly;
end;

procedure TvgrSEMemo.CMMousecheck(var Msg : TWMMouse);
var
  ACursorPos : TPoint;
begin
  if Selection.Active then
  begin
    GetCursorPos(ACursorPos);
    ACursorPos := ScreenToClient(ACursorPos);
    if (ACursorPos.X = FMMPosX) and (ACursorPos.Y = FMMPosY) then
      MouseMove(FMMShift, FMMPosX, FMMPosY);
  end;
  Msg.Result := 1;
end;

procedure TvgrSEMemo.WMEraseBkgnd(var Msg : TWMEraseBkgnd);
begin
//  ClearBkGround;
  Msg.Result := 1;
end;

procedure TvgrSEMemo.WMGetDlgCode(var Msg : TWMGetDlgCode);
begin
  Msg.Result := DLGC_WANTARROWS or DLGC_WANTCHARS;
  if ScriptEdit.WantTabs then
    Msg.Result := Msg.Result or DLGC_WANTTAB
  else
    Msg.Result := Msg.Result and not DLGC_WANTTAB;
  if ScriptEdit.WantReturns then
    Msg.Result := Msg.Result or DLGC_WANTALLKEYS
  else
    Msg.Result := Msg.Result and not DLGC_WANTALLKEYS;
end;

procedure TvgrSEMemo.WMSetFocus(var Message: TWMSetFocus);
begin
  inherited;
  SetFocus;
  SetCaret;
  Message.FocusedWnd := Handle;
  Message.Result := 0;
end;

procedure TvgrSEMemo.WMKillFocus(var Message: TWMSetFocus);
begin
  inherited;
  ResetCaret;
end;

procedure TvgrSEMemo.WMCopy(var Message: TWMCopy);
begin
end;

procedure TvgrSEMemo.WMPaste(var Message: TWMPaste);
begin
end;

procedure TvgrSEMemo.WMCut(var Message: TWMCut);
begin
end;

procedure TvgrSEMemo.WMLButtonDblClk(var Message: TWMLButtonDblClk);
begin
  DblClick;
  SelectNearestWord;

end;

/////////////////////////////////////////////////
//
// TvgrCustomScriptEdit
//
/////////////////////////////////////////////////
constructor TvgrCustomScriptEdit.Create(AOwner : TComponent);
begin
  inherited;

  FScript := TvgrScriptSyntax.Create;
  FScript.OnChange := OnScriptChange;

  FHorzScrollBar := TvgrSEScrollBar.Create(Self);
  InitHorzScrollBar;
  FVertScrollBar := TvgrSEScrollBar.Create(Self);
  FVertScrollBar.Kind := sbVertical;
  FVertScrollBar.Parent := Self;

  FTabIncrement := 4;
  FBorderStyle := bsSingle;
  FAutoEditorMode := True;
  FTextAttributes := nil;
  FMaxLineLength := cSEMaxLineLength;
  FScrollBars := ssBoth;

  FGutter := TvgrSEGutter.Create(Self);
  FGutter.Parent := Self;
  FGutterSwitcher := TvgrSEControlSwitcher.Create(Self);
  FGutterSwitcher.Control := FGutter;

  FSyntaxCheck := TvgrSESyntaxPanel.Create(Self);
  FSyntaxCheck.Height := 30;
  FSyntaxCheck.Visible := False;

  FSyntaxCheckSwitcher := TvgrSEControlSwitcher.Create(Self);
  FSyntaxCheckSwitcher.Control := FSyntaxCheck;
  FSyntaxCheckSwitcher.Height := 8;
  FSyntaxCheckSwitcher.Kind := sbHorizontal;

  FLines := TvgrSEStrings.Create(Self);

  FMemo := TvgrSEMemo.Create(Self);
  FMemo.Align := alClient;

  TabStop := False;
  FTextOptions := TvgrSETextOptions.Create(FMemo);
  FReplaceTabsWithSpaces := True;
end;

destructor TvgrCustomScriptEdit.Destroy;
begin
  FScript.Free;
  if FTextAttributes <> nil then
    FreeMem(FTextAttributes);
  FHorzScrollBar.Free;
  FVertScrollBar.Free;
  FMemo.Free;
  FLines.Free;
  FGutter.Free;
  FSyntaxCheck.Free;
  FSyntaxCheckSwitcher.Free;
  FTextOptions.Free;
  inherited;
end;

procedure TvgrCustomScriptEdit.CheckSyntax;
begin
  FSyntaxErrorPresent := not FScript.CheckSyntax(Lines.Text, FSyntaxError);
  if FSyntaxErrorPresent then
  begin
    SetCaretToSyntaxError;
    DrawSyntaxErrorLine;
  end;
end;

procedure TvgrCustomScriptEdit.CaretMove;
begin
  if Assigned(FOnCaretMove) then
    FOnCaretMove(Self, FMemo.FAbsoluteCaretPos.X, FMemo.FAbsoluteCaretPos.Y);
end;

procedure TvgrCustomScriptEdit.SetCaretPos(ACol, ARow: Integer);
begin
  FMemo.AbsoluteCaretPos := Point(ACol, ARow);
  FMemo.ScrollToCaret;
end;

procedure TvgrCustomScriptEdit.ClipBoardCopy;
begin
  FMemo.ClipBoardCopy;
end;

procedure TvgrCustomScriptEdit.ClipBoardPaste;
begin
  FMemo.ClipBoardPaste;
end;

procedure TvgrCustomScriptEdit.ClipBoardCut;
begin
  FMemo.ClipBoardCut;
end;

procedure TvgrCustomScriptEdit.Paint;

  procedure PaintSingleBorder;
  begin
    with Canvas, Canvas.Pen do
    begin
      Style := psSolid;
      Color := clBtnShadow;
      MoveTo(0,ClipRect.Bottom - GetBorderWidth);
      LineTo(0,0);
      LineTo(ClipRect.Right - GetBorderWidth,0);
      Color := clWindow;
      LineTo(ClipRect.Right - GetBorderWidth, ClipRect.Bottom - GetBorderWidth);
      LineTo(0,ClipRect.Bottom - GetBorderWidth);
    end;
  end;

begin
  with Canvas do
  begin
    with Brush do
    begin
      Style := bsSolid;
      Color := clBtnFace;
      FillRect(GetBottomRightRect);
    end;
    if BorderStyle  = bsSingle then
      PaintSingleBorder;
  end;
end;

procedure TvgrCustomScriptEdit.Resize;
begin
  inherited;
  UpdateScrollBars;
end;

procedure TvgrCustomScriptEdit.AlignControls(AControl: TControl; var Rect: TRect);
var
  ATop1, ATop2, ATop3, ATop4, ATop5 : Integer;
  ALeft1, ALeft2, ALeft3, ALeft4, ALeft5 : Integer;
begin
  ATop1 := GetBorderWidth;
  ATop5 := Rect.Bottom - GetBorderWidth;
  ATop4 := ATop5;
  if FSyntaxCheckSwitcher.Visible then
    Dec(ATop4, FSyntaxCheckSwitcher.Height);
  ATop3 := ATop4;
  if FSyntaxCheck.Visible then
    Dec(ATop3, FSyntaxCheck.Height);
  ATop2 := ATop3;
  if FHorzScrollBar.Visible then
    Dec(ATop2, FHorzScrollBar.Height);

  ALeft1 := GetBorderWidth;
  ALeft2 := ALeft1;
  if FGutterSwitcher.Visible then
    Inc(ALeft2, FGutterSwitcher.Width);
  ALeft3 := ALeft2;
  if FGutter.Visible then
    Inc(ALeft3, FGutter.Width);
  ALeft5 := Rect.Right - GetBorderWidth;
  ALeft4 := ALeft5;
  if FVertScrollBar.Visible then
    Dec(ALeft4, FVertScrollBar.Width);

  with FBottomRightRect do
  begin
    Left := ALeft4;
    Top :=  ATop2;
    Right := ALeft5;
    Bottom := ATop3;
  end;
  FGutterSwitcher.SetBounds(ALeft1,ATop1,
    ALeft2 - ALeft1, ATop2 - ATop1);
  FGutter.SetBounds(ALeft2, ATop1,
    FGutter.Width, ATop2 - ATop1);
  FMemo.SetBounds(ALeft3, ATop1,
    ALeft4 - ALeft3, ATop2 - ATop1);
  FVertScrollBar.SetBounds(ALeft4, ATop1,
    FVertScrollBar.Width,
    ATop2 - ATop1);
  FHorzScrollBar.SetBounds(ALeft1, ATop2,
    ALeft4 - ALeft1,
    FHorzScrollBar.Height);
  FSyntaxCheck.SetBounds(ALeft1, ATop3, ALeft5 - ALeft1, FSyntaxCheck.Height);
  FSyntaxCheckSwitcher.SetBounds(ALeft1, ATop4,
    ALeft5 - ALeft1,
    ATop5 - ATop4);
end;

procedure TvgrCustomScriptEdit.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
end;

// private
procedure TvgrCustomScriptEdit.SetCaretToSyntaxError;
begin
  with FMemo, FSyntaxError do
  begin
    AbsoluteCaretPos := Point(Position, Line);
    SetCaretToPosition;
    ScrollToCaret;
  end;
end;

procedure TvgrCustomScriptEdit.DrawSyntaxErrorLine;
begin
  FMemo.DrawSyntaxErrorLine(FSyntaxError.Line);
end;

function TvgrCustomScriptEdit.GetBottomRightRect : TRect;
begin
  with Result do
  begin
    Left := GetBorderWidth;
    Right := Width - GetBorderWidth;
    Top := GetBorderWidth;
    Bottom := Height - GetBorderWidth;
  end;
  Result := FBottomRightRect;
end;

function TvgrCustomScriptEdit.GetBorderWidth: Integer;
begin
  if BorderStyle = bsSingle then
    Result := 1
  else
    Result := 0;
end;

procedure TvgrCustomScriptEdit.TextChanged;
begin
  GetTextAttributes;
  if FSyntaxErrorPresent then
  begin
    FSyntaxErrorPresent := False;
  end;
  FMemo.ChangeTextLines(FSyntaxError.Line, FSyntaxError.Line);
  if HandleAllocated then
    Invalidate;
  FSyntaxCheck.Invalidate;
  if Assigned(FonChange) then
    FOnChange(Self);
end;

procedure TvgrCustomScriptEdit.GetTextAttributes;
var
  AText: string;
begin
  AText := Lines.Text;
  AText := StringReplace(AText, #9, StringOfChar(' ',TabIncrement), [rfReplaceAll]);

  Script.GetTextHighlightAttributes(AText, FTextAttributes);
end;

procedure TvgrCustomScriptEdit.InitHorzScrollBar;
begin
  with FHorzScrollBar do
  begin
    Kind := sbHorizontal;
    Parent := Self;
    Max := cSEMaxLineLength;
  end;
end;

procedure TvgrCustomScriptEdit.UpdateScrollBars;
begin
  UpdateScrollBar(HorzScrollBar, FMemo.VisibleTextOffset.X, FMemo.GetVisibledCharactersCount - 2, cSEMaxLineLength);
  UpdateScrollBar(VertScrollBar, FMemo.VisibleTextOffset.Y, FMemo.GetVisibledLinesCount, Lines.Count);
end;

procedure TvgrCustomScriptEdit.UpdateGutter;
begin
  Gutter.PaintLines(False);
end;

procedure TvgrCustomScriptEdit.UpdateScrollBar(ScrollBar : TvgrSEScrollBar; NewPos, VisiblePoints, MaxPoints : integer);
begin
  with ScrollBar do
    SetParams(NewPos, 0, MaxPoints, VisiblePoints);
end;

procedure TvgrCustomScriptEdit.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrCustomScriptEdit.WMSetFocus(var Message: TWMSetFocus);
begin
  FMemo.SetFocus;
  Message.Result := 1;
end;

procedure TvgrCustomScriptEdit.ChangeAbsoluteCaretPos(AOldValue, ANewValue: Integer);
begin
  with Gutter do
  begin
    PaintLine(AOldValue, False);
    PaintLine(ANewValue, False);
  end;
end;

procedure TvgrCustomScriptEdit.ChangeVisibleYOffset;
begin
  Gutter.PaintLines(False);
end;

procedure TvgrCustomScriptEdit.SetBorderStyle(Value: TBorderStyle);
begin
  FBorderStyle := Value;
  Realign;
  Invalidate;
end;

procedure TvgrCustomScriptEdit.SetLines(Value: TStrings);
begin
  FLines.Assign(Value);
  Invalidate;
end;

procedure TvgrCustomScriptEdit.SetScrollBars(Value: TScrollStyle);
begin
  FScrollBars := Value;
  FHorzScrollBar.Visible := Value in [ssBoth, ssHorizontal];
  FVertScrollBar.Visible := Value in [ssBoth, ssVertical];
  Realign;
end;

function TvgrCustomScriptEdit.GetGutterVisible: Boolean;
begin
  Result := FGutterSwitcher.ControlVisible;
end;

procedure TvgrCustomScriptEdit.SetGutterVisible(Value: Boolean);
begin
  FGutterSwitcher.ControlVisible := Value;
end;

procedure TvgrCustomScriptEdit.SetTabIncrement(Value: Integer);
begin
  FTabIncrement := Max(Value, 2);
end;

procedure TvgrCustomScriptEdit.OnScriptChange(Sender: TObject);
begin
  if HandleAllocated then
  begin
    TextChanged;
    Invalidate;
  end;
end;

function TvgrCustomScriptEdit.GetSelectionActive: Boolean;
begin
  Result := FMemo.Selection.Active;
end;

procedure TvgrCustomScriptEdit.SetEditorMode(Value: TvgrSEInsertMode);
begin
  FEditorMode := Value;
  if Assigned(FOnEditorModeChange) then
    FOnEditorModeChange(Self);
end;

function TvgrCustomScriptEdit.GetCaretPos: TPoint;
begin
  Result := FMemo.AbsoluteCaretPos;
end;

procedure TvgrCustomScriptEdit.SetScript(Value: TvgrScriptSyntax);
begin
  FScript.Assign(Value);
end;

end.
