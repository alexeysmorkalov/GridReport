{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes which realize a calculation of formulas in TvgrWorkbook.
TvgrFormulaCalculator is main class, he implements a Calcule method used to
calculate value of the formula.}
unit vgr_FormulaCalculator;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, typinfo, SysUtils, math,
  {$IFDEF VTK_D6_OR_D7} Types, Variants, {$ELSE} Windows, {$ENDIF}

  vgr_CommonClasses, vgr_Functions, vgr_ExcelFormula;

type
  TvgrFormulaCalculator = class;

{Specifies type of operand in expression.
Itmes:
  vgrotInteger - Integer
  vgrotExtended - Extended
  vgrotString - String}
  TvgrOperandTypes = (vgrotInteger, vgrotExtended, vgrotString);

  /////////////////////////////////////////////////
  //
  // rvgrEnumStackRec
  //
  /////////////////////////////////////////////////
{Internal structure.}
  rvgrEnumStackRec = record
    Value: rvteFormulaValue;
    First: Boolean;
    Res1: Integer;
    Res2: Integer;
  end;
{Pointer to a rvgrEnumStackRec structure.}
  pvgrEnumStackRec = ^rvgrEnumStackRec;
  TvgrFormulaCalculatorEnumStackProc = function(ACalculator: TvgrFormulaCalculator; AValue: pvteFormulaValue; var AEnumStackRec: rvgrEnumStackRec): Boolean;

  /////////////////////////////////////////////////
  //
  // TvgrWBStrings
  //
  /////////////////////////////////////////////////
{Class for keeping string values of the ranges.}
  TvgrWBStrings = class(TObject)
  private
    FStrings: TStringList;
    FLastEmptyStringIndex: Integer;
    function GetItem(Index: Integer): string;
    function GetCount: Integer;
  public
{Creates an instance of the TvgrWBStrings class.}
    constructor Create;
{Disposes of an TvgrWBStrings instance.}
    destructor Destroy; override;
{Clear the list.}
    procedure Clear;
{Saves strings in the stream. Strings are saved with current position of the stream.}
    procedure SaveToStream(AStream: TStream);
{Restores strings from the stream.
Restoring of string starts with current position of the stream.}
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
{$IFDEF VGR_DS_DEBUG}
    function DebugInfo: TvgrDebugInfo;
{$ENDIF}
{FindOrAdd returns the 0-based index of the string.
Parameters:
  s - String to search.
Return value:
  Thus, if S matches the first string in the list,
  FindOrAdd returns 0, if S is the second string,
  FindOrAdd returns 1, and so on. If the string is not in the string list,
  FindOrAdd adds it to list and returns it index. }
    function FindOrAdd(const s: string): Integer;
{Decrements the reference count for string with this index.
Parameters:
  AIndex - Index of string.}
    procedure Release(AIndex: Integer);
{References the strings in the list by their positions.}
    property Items[Index: Integer]: string read GetItem; default;
{Count of the strings in the list.}
    property Count: Integer read GetCount;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrFormulaCalculatorOwner
  //
  /////////////////////////////////////////////////
{Specifies interface, which must be realizated by the object which uses TvgrFormulaCalculator object.}
  IvgrFormulaCalculatorOwner = Interface
{Calculates a value of cell.
Parameters:
  ACalculator - reference to the TvgrFormulaCalculator object.
  AValue - value of the cell.
  ASheet - index of sheet of the cell.
  ACol - column of the cell.
  ARow - row of the cell.
Return value:
  Returns true if value are successfully calculated.}
    function GetCellValue(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; ASheet, ACol, ARow: Integer): Boolean;
{This function must enumerate all ranges within ARangeRect on sheet with index ASheet.
For all found range ACallback procedure must be called.
Parameters:
  ACalculator - reference to the TvgrFormulaCalculator object.
  AEnumRec - this structure must be passed to a ACallback procedure.
  ACallback - this procedure must be called for all value within ARangeRect.
  ASheet - index of sheet, inside which the values are searched.
  ARangeRect - rectangle, inside which the values are searched. }
    function EnumRangeValues(ACalculator: TvgrFormulaCalculator;
                             var AEnumRec: rvgrEnumStackRec;
                             ACallback: TvgrFormulaCalculatorEnumStackProc;
                             ASheet: Integer;
                             const ARangeRect: TRect): Boolean;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrCalculatorStack
  //
  /////////////////////////////////////////////////
{Internal class, implements a stack of the operands of the formula.}
  TvgrCalculatorStack = class(TObject)
  private
    FList: TList;
    FCurPos: Integer;
    FCalculator: TvgrFormulaCalculator;
    function GetItem(Index: Integer): pvteFormulaItem;
    function PushItem: pvteFormulaItem;
  protected
    procedure Push(AItem: pvteFormulaItem);
    function Pop(var AItem: pvteFormulaItem): Boolean; overload;
    function Pop(AItemsToPop: Integer): Boolean; overload;


    property Calculator: TvgrFormulaCalculator read FCalculator;
    property Items[Index: Integer]: pvteFormulaItem read GetItem;
  public
    constructor Create(ACalculator: TvgrFormulaCalculator);
    destructor Destroy; override;
{Clears the content of the object.}
    procedure Clear;
{Resets the state of the object.}
    procedure Reset;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormulaCalculator
  //
  /////////////////////////////////////////////////
{Calculates the value of formula.}
  TvgrFormulaCalculator = class(TObject)
  private
    FOwner: IvgrFormulaCalculatorOwner;
    FStack: TvgrCalculatorStack;
    FStrings: TStringList;
    FExternalWBStrings: TvgrWBStrings;
    FErrorMessage: string;
    function GetStringCount: Integer;
    function GetString(I: Integer): string;
  protected
    procedure SetError(const AErrorMessage: string);

    function GetCell(const ACellRef: rvteFormulaCellRef;
                     AFormulaSheet, AFormulaCol, AFormulaRow: Integer;
                     var ACellX, ACellY, ACellSheet: Integer): Boolean;
    function GetRange(const ARangeRef: rvteFormulaRangeRef;
                      AFormulaSheet, AFormulaCol, AFormulaRow: Integer;
                      var ARangeLeft, ARangeTop, ARangeRight, ARangeBottom, ARangeSheet: Integer): Boolean;

    function CalculateCellValue(const ACellRef: rvteFormulaCellRef; var AValue: rvteFormulaValue; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
    function CalculateOperator(AItem: pvteFormulaItem; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
    function CalculateFunction(AItem: pvteFormulaItem; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

    procedure FromInteger(var AFormulaValue: rvteFormulaValue; AValue: Integer);
    procedure FromExtended(var AFormulaValue: rvteFormulaValue; AValue: Extended);
    procedure FromDateTime(var AFormulaValue: rvteFormulaValue; AValue: TDateTime);
    procedure FromString(var AFormulaValue: rvteFormulaValue; const AValue: string);
    procedure FromNull(var AFormulaValue: rvteFormulaValue);

    function ToInteger(const AFormulaValue: rvteFormulaValue; var AValue: Integer): Boolean;
    function ToExtended(const AFormulaValue: rvteFormulaValue; var AValue: Extended): Boolean;
    function ToDateTime(const AFormulaValue: rvteFormulaValue; var AValue: TDateTime): Boolean;
    function ToString(const AFormulaValue: rvteFormulaValue; var AValue: string): Boolean;
    function CopyValue(const ASourceValue: rvteFormulaValue; var ADestValue: rvteFormulaValue): Boolean;

    function EnumStack(ACount: Integer; var AEnumRec: rvgrEnumStackRec; ACallback: TvgrFormulaCalculatorEnumStackProc; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
    function Pop(var AItem: pvteFormulaItem; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean; overload;
    function Pop(var AValue: pvteFormulaValue; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean; overload;
    function Pop(var AInt: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean; overload;
    function Pop(var AStr: string; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean; overload;
    function Pop(var AExtended: Extended; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean; overload;
    function Pop(var ADateTime: TDateTime; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean; overload;
  public
    constructor Create(AOwner: IvgrFormulaCalculatorOwner);
    destructor Destroy; override;

{Calculates the value of formula. Before calculation formula must be compiled with using
TvteExcelFormulaCompiler object.
If error has occur during calculating, ErrorMessage property contains description of error.
Parameters:
  AFormula - The compiled formula. To convert string representation of the formula
to its compiled representation use TvgrExcelFormulaCompiler.
  AFormulaSize - count of elements in the formula.
  AValue - in this parameter value is returned.
  AFormulaSheet - Describes position of the formula within workbook.
  AFormulaCol - Describes position of the formula within workbook.
  AFormulaRow - Describes position of the formula within workbook.
Return value:
  Function returns true if value of the formula is calculated successfully.}
    function Calculate(AFormula: pvteFormula; AFormulaSize: Integer; var AValue: Variant; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

    function AddString(const S: string): Integer;
    procedure ClearStrings;

{Owner of the formula calculator, can not be nil.}
    property Owner: IvgrFormulaCalculatorOwner read FOwner;
{Refrence to TvgrCalculatorStack object, taht represents a stack, used while calculate of the formula,
can not be nil, must be initializated.}
    property Stack: TvgrCalculatorStack read FStack write FStack;
{Reference to TvgrWBStrings object, which stores strings of the workbook.}
    property ExternalWBStrings: TvgrWBStrings read FExternalWBStrings write FExternalWBStrings;
{Contains error description or empty string if no error.}
    property ErrorMessage: string read FErrorMessage;

    property Strings[I: Integer]: string read GetString;
    property StringCount: Integer read GetStringCount;
  end;

  TvgrFormulaFunctionProc = function(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredFunction
  //
  /////////////////////////////////////////////////
{TvgrRegisteredFunction contains information about registered function.
Each function used by TvgrFormulaCalculator must be registered before it can be used.
See also:
  TvgrRegisteredFunctions}
  TvgrRegisteredFunction = class(TObject)
  protected
    FName: string;
    FProc: TvgrFormulaFunctionProc;
    FMinParams: Integer;
    FMaxParams: Integer;
  public
    constructor Create(const AName: string; AProc: TvgrFormulaFunctionProc; AMinParams, AMaxParams: Integer);
{Name of function.}
    property Name: string read FName;
{Procedure which calculates value of the function.}
    property Proc: TvgrFormulaFunctionProc read FProc;
{Minimal number of in parameters.}
    property MinParams: Integer read FMinParams;
{Maximum number of in parameters.}
    property MaxParams: Integer read FMaxParams;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredFunctions
  //
  /////////////////////////////////////////////////
{Holds the list of registered functions.
See also:
  TvgrRegisteredFunction}
  TvgrRegisteredFunctions = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrRegisteredFunction;
  public
{Registers the new function. If function with name AFuncName not found - creates a new
TvgrRegisteredFunction object and add it to list.
Parameters:
  AFuncName - name of function.
  AFuncProc - procedure which calculates value of the function.
  AMinParams - minimal number of in parameters.
  AMaxParams - maximum number of in parameters.}
    procedure RegisterFunction(const AFuncName: string; AFuncProc: Pointer; AMinParams, AMaxParams: Integer);
{Returns index of function by its name. If function with name AFuncName not found returns -1.
Parameters:
  AFuncName - name of function
Return value:
  Return the index of function in the list.}
    function GetFunctionIndex(const AFuncName: string): Integer;
{Returns name of function by its index.
Parameters:
  AFuncIndex - Index of function
Return value:
  Name of function}
    function GetFunctionName(AFuncIndex: Integer): string;

{Use Items to access objects in the list.
Items is a zero-based array: The first object is indexed as 0, the second object is indexed as 1,
and so forth.<br>
You can read or change the value at a specific index, or use Items with the Count property
to iterate through the list.}
    property Items[Index: Integer]: TvgrRegisteredFunction read GetItem; default;
  end;

  TvgrFormulaOperatorIntProc = function(ACalculator: TvgrFormulaCalculator;
                                        var AValue: rvteFormulaValue;
                                        Op1, Op2: Integer): Boolean;
  TvgrFormulaOperatorExtProc = function(ACalculator: TvgrFormulaCalculator;
                                        var AValue: rvteFormulaValue;
                                        Op1, Op2: Extended): Boolean;
  TvgrFormulaOperatorStrProc = function(ACalculator: TvgrFormulaCalculator;
                                        var AValue: rvteFormulaValue;
                                        const Op1, Op2: string): Boolean;
  
/////////////////////////////////////////////////
//
// Functions
//
/////////////////////////////////////////////////
//function vgrAsInteger(AItem: pvteFormulaValue; AWBStrings: TvgrWBStrings; var AValue: Integer): Boolean;
//function vgrAsExtended(AItem: pvteFormulaValue; AWBStrings: TvgrWBStrings; var AValue: Extended): Boolean;
//function vgrAsString(AItem: pvteFormulaValue; AWBStrings: TvgrWBStrings; var AValue: string): Boolean;
//procedure vgrCopy(const ASource: rvteFormulaValue; var ADest: rvteFormulaValue);

{Registers new function which can be used in formulas.
Parameters:
  AFuncName - name of function.
  AFuncProc - procedure which calculates value of the function.
  AMinParams - minimal number of in parameters.
  AMaxParams - maximum number of in parameters. }
procedure RegisterFunction(const AFuncName: string; AFuncProc: Pointer; AMinParams, AMaxParams: Integer);

/////////////////////////////////////////////////
//
// Operators
//
/////////////////////////////////////////////////
function OpAddInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpSubInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpMulInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpDivInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpPowerInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpLTInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpLEInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpEQInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpGEInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpGTInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpNEInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpPlusInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpMinusInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
function OpPercentInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;

function OpAddExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpSubExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpMulExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpDivExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpPowerExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpLTExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpLEExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpEQExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpGEExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpGTExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpNEExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpPlusExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpMinusExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
function OpPercentExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;

function OpConcatStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
function OpLTStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
function OpLEStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
function OpEQStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
function OpGEStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
function OpGTStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
function OpNEStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;

/////////////////////////////////////////////////
//
// Functions
//
/////////////////////////////////////////////////
function FnCount(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnIf(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnSum(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnAverage(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnMin(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnMax(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

// date functions
function FnNow(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnDateValue(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnDate(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnDay(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnHour(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnMonth(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnMinute(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnSecond(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnTime(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnTimeValue(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnToday(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnWeekDay(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnYear(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

// math functions
function FnAbs(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnRound(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnSign(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

// lookup and reference functions
function FnColumn(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnRow(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnIndirect(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnRows(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnColumns(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

// text functions
function FnChar(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnCode(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnExact(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnLeft(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnLen(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnLower(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnMid(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnRight(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
function FnUpper(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;

var
  { List of the registered functions, which can be used by TvgrFormulaCalculator in formulas calculation. }
  RegisteredFunctions: TvgrRegisteredFunctions;

  AOperatorsIntProcs: Array [vteMinOperator_ptg..vteMaxOperator_ptg] of TvgrFormulaOperatorIntProc =
  (
    OpAddInt,
    OpSubInt,
    OpMulInt,
    OpDivInt,
    OpPowerInt,
    nil,
    OpLTInt,
    OpLEInt,
    OpEQInt,
    OpGEInt,
    OpGTInt,
    OpNEInt,
    nil,
    nil,
    nil,
    OpPlusInt,
    OpMinusInt,
    OpPercentInt
  );
  AOperatorsExtProcs: Array [vteMinOperator_ptg..vteMaxOperator_ptg] of TvgrFormulaOperatorExtProc =
  (
    OpAddExt,
    OpSubExt,
    OpMulExt,
    OpDivExt,
    OpPowerExt,
    nil,
    OpLTExt,
    OpLEExt,
    OpEQExt,
    OpGEExt,
    OpGTExt,
    OpNEExt,
    nil,
    nil,
    nil,
    OpPlusExt,
    OpMinusExt,
    OpPercentExt
  );
  AOperatorsStrProcs: Array [vteMinOperator_ptg..vteMaxOperator_ptg] of TvgrFormulaOperatorStrProc =
  (
    nil,
    nil,
    nil,
    nil,
    nil,
    OpConcatStr,
    OpLTStr,
    OpLEStr,
    OpEQStr,
    OpGEStr,
    OpGTStr,
    OpNEStr,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil
  );

implementation

uses
  vgr_StringIDs, vgr_Localize;

{$R ..\res\vgr_FormulaCalculatorStrings.res}

procedure RegisterFunction(const AFuncName: string; AFuncProc: Pointer; AMinParams, AMaxParams: Integer);
begin
  RegisteredFunctions.RegisterFunction(AFuncName, AFuncProc, AMinParams, AMaxParams);
end;

(*
procedure VariantToFormulaValue(var AValue: rvteFormulaValue; const AVariant: Variant; AWBStrings: TvgrWBStrings);
var
  AVarType: Integer;   
begin
  if AValue.ValueType = vteString then
    AWBStrings.Release(AValue.vString);
    
  AVarType := VarType(AVariant);
  case AVarType of
    {$IFDEF VGR_D6_OR_D7}varShortInt, varWord, varLongWord, varInt64, {$ENDIF} varByte, varSmallint, varInteger:
      begin
        AValue.ValueType := vteInteger;
        AValue.vInteger := AVariant;
      end;
    varCurrency, varSingle, varDouble, varDate:
      begin
        AValue.ValueType := vteExtended;
        AValue.vExtended := AVariant;
      end;
    varOleStr, varStrArg, varString:
      begin
        AValue.ValueType := vteString;
        AValue.vString := AWBStrings.FindOrAdd(VarToStr(AVariant));
      end;
    varBoolean:
      begin
        AValue.ValueType := vteInteger;
        AValue.vInteger := AVariant;
      end;
  else
    AValue.ValueType := vteNull;
  end;
end;
*)

function OpAddInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := AOp1 + AOp2;
  Result := True;
end;

function OpSubInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := AOp1 - AOp2;
  Result := True;
end;

function OpMulInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.ValueType := vteExtended;
  AValue.vExtended := AOp1 * Aop2;
  Result := True;
end;

function OpDivInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.ValueType := vteExtended;
  Result := AOp2 <> 0;
  if not Result then
    ACalculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_DivideByZero))
  else
    AValue.vExtended := AOp1 / AOp2;
end;

function OpPowerInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.ValueType := vteExtended;
  AValue.vExtended := Power(AOp1, AOp2);
  Result := True;
end;

function OpLTInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := Integer(AOp1 < Aop2);
  Result := True;
end;

function OpLEInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := Integer(AOp1 <= Aop2);
  Result := True;
end;

function OpEQInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := Integer(AOp1 = Aop2);
  Result := True;
end;

function OpGEInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := Integer(AOp1 >= Aop2);
  Result := True;
end;

function OpGTInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := Integer(AOp1 > Aop2);
  Result := True;
end;

function OpNEInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := Integer(AOp1 <> Aop2);
  Result := True;
end;

function OpPlusInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := AOp1;
  Result := True;
end;

function OpMinusInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.vInteger := -AOp1;
  Result := True;
end;

function OpPercentInt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Integer): Boolean;
begin
  AValue.ValueType := vteExtended;
  AValue.vExtended := AOp1 / 100;
  Result := True;
end;

function OpAddExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := AOp1 + AOp2;
  Result := True;
end;

function OpSubExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := AOp1 - AOp2;
  Result := True;
end;

function OpMulExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := AOp1 * AOp2;
  Result := True;
end;

function OpDivExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  Result := not IsZero(AOp2);
  if Result then
    AValue.vExtended := AOp1 / AOp2
  else
    ACalculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_DivideByZero));
end;

function OpPowerExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := Power(AOp1, AOp2);
  Result := True;
end;

function OpLTExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 < Aop2);
  Result := True;
end;

function OpLEExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 <= Aop2);
  Result := True;
end;

function OpEQExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 = Aop2);
  Result := True;
end;

function OpGEExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 >= Aop2);
  Result := True;
end;

function OpGTExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 > Aop2);
  Result := True;
end;

function OpNEExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 <> Aop2);
  Result := True;
end;

function OpPlusExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := AOp1;
  Result := True;
end;

function OpMinusExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := -AOp1;
  Result := True;
end;

function OpPercentExt(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; AOp1, AOp2: Extended): Boolean;
begin
  AValue.vExtended := AOp1 / 100;
  Result := True;
end;

function OpConcatStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteString;
  AValue.vString := ACalculator.AddString(AOp1 + AOp2);
  Result := True;
end;

function OpLTStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 < Aop2);
  Result := True;
end;

function OpLEStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 <= Aop2);
  Result := True;
end;

function OpEQStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 = Aop2);
  Result := True;
end;

function OpGEStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 >= Aop2);
  Result := True;
end;

function OpGTStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 > Aop2);
  Result := True;
end;

function OpNEStr(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaValue; const AOp1, AOp2: string): Boolean;
begin
  AValue.ValueType := vteInteger;
  AValue.vInteger := Integer(AOp1 <> Aop2);
  Result := True;
end;

/////////////////////////////////////////////////
//
// Functions
//
/////////////////////////////////////////////////
function FnCountCallback(ACalculator: TvgrFormulaCalculator; AValue: pvteFormulaValue; var AEnumStackRec: rvgrEnumStackRec): Boolean;
begin
  if (AValue <> nil) and (AValue.ValueType <> vteNull) then
    AEnumStackRec.Value.vInteger := AEnumStackRec.Value.vInteger + 1;
  Result := True;
end;

function FnCount(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AEnumStackRec: rvgrEnumStackRec;
begin
  with AEnumStackRec do
  begin
    Value.ValueType := vteInteger;
    Value.vInteger := 0;
    First := True;
  end;
  Result := ACalculator.EnumStack(AParamCount, AEnumStackRec, FnCountCallback, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
    Result := ACalculator.CopyValue(AEnumStackRec.Value, AValue.Value);
end;

function FnIf(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AOp1: Integer;
  AOp2, AOp3: pvteFormulaValue;
begin
  Result := ACalculator.Pop(AOp3, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AOp2, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AOp1, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    if Boolean(AOp1) then
      Result := ACalculator.CopyValue(AOp2^, AValue.Value)
    else
      Result := ACalculator.CopyValue(AOp3^, AValue.Value)
  end;
end;

function FnSumCallback(ACalculator: TvgrFormulaCalculator; AValue: pvteFormulaValue; var AEnumStackRec: rvgrEnumStackRec): Boolean;
var
  AType: TvteFormulaValueType;
  AInt: Integer;
  AExt: Extended;
begin
  if AValue <> nil then
  begin
    case AValue.ValueType of
      vteInteger:
        begin
          AType := vteInteger;
          AInt := AValue.vInteger;
        end;
      vteExtended:
        begin
          AType := vteExtended;
          AExt := AValue.vExtended;
        end;
    else
      begin
        if ACalculator.ToInteger(AValue^, AInt) then
          AType := vteInteger
        else
          if ACalculator.ToExtended(AValue^, AExt) then
            AType := vteExtended
          else
          begin
            Result := False;
            exit;
          end;
      end;
    end;
    if AType = vteInteger then
    begin
      if AEnumStackRec.Value.ValueType = vteInteger then
        AEnumStackRec.Value.vInteger := AEnumStackRec.Value.vInteger + AInt
      else
        AEnumStackRec.Value.vExtended := AEnumStackRec.Value.vExtended + AInt
    end
    else
    begin
      if AEnumStackRec.Value.ValueType = vteInteger then
      begin
        AEnumStackRec.Value.vExtended := AEnumStackRec.Value.vInteger + AExt;
        AEnumStackRec.Value.ValueType := vteExtended;
      end
      else
        AEnumStackRec.Value.vExtended := AEnumStackRec.Value.vExtended + AExt
    end;
  end;
  Result := True;
end;

function FnSum(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AEnumStackRec: rvgrEnumStackRec;
begin
  with AEnumStackRec do
  begin
    Value.ValueType := vteInteger;
    Value.vInteger := 0;
    First := True;
  end;
  Result := ACalculator.EnumStack(AParamCount, AEnumStackRec, FnSumCallback, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
    ACalculator.CopyValue(AEnumStackRec.Value, AValue.Value);
end;

function FnAverageCallback(ACalculator: TvgrFormulaCalculator; AValue: pvteFormulaValue; var AEnumStackRec: rvgrEnumStackRec): Boolean;
begin
  Result := FnSumCallback(ACalculator, AValue, AEnumStackRec);
  if Result and (AValue <> nil) and (AValue.ValueType <> vteNull) then
    AEnumStackRec.Res1 := AEnumStackRec.Res1 + 1;
end;

function FnAverage(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AEnumStackRec: rvgrEnumStackRec;
  AExt: Extended;
begin
  with AEnumStackRec do
  begin
    Value.ValueType := vteInteger;
    Value.vInteger := 0;
    Res1 := 0;
    First := True;
  end;
  
  Result := ACalculator.EnumStack(AParamCount, AEnumStackRec, FnAverageCallback, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    if AEnumStackRec.Res1 = 0 then
      AValue.Value.ValueType := vteNull
    else
      if ACalculator.ToExtended(AEnumStackRec.Value, AExt) then
      begin
        AValue.Value.ValueType := vteExtended;
        AValue.Value.vExtended := AExt / AEnumStackRec.Res1;
      end;
  end;
end;

function FnMinCallback(ACalculator: TvgrFormulaCalculator; AValue: pvteFormulaValue; var AEnumStackRec: rvgrEnumStackRec): Boolean;
var
  AInt: Integer;
  AExt: Extended;
  AStr: string;
begin
  if AEnumStackRec.First and (AValue.ValueType <> vteNull) then
  begin
    ACalculator.CopyValue(AValue^, AEnumStackRec.Value);
    AEnumStackRec.First := False;
    Result := True;
  end
  else
    case AEnumStackRec.Value.ValueType of
      vteInteger:
        begin
          Result := ACalculator.ToInteger(AValue^, AInt);
          if Result then
            if AInt < AEnumStackRec.Value.vInteger then
              AEnumStackRec.Value.vInteger := AInt;
        end;
      vteExtended:
        begin
          Result := ACalculator.ToExtended(AValue^, AExt);
          if Result then
            if AExt < AEnumStackRec.Value.vExtended then
              AEnumStackRec.Value.vExtended := AExt;
        end;
      vteString:
        begin
          Result := ACalculator.ToString(AValue^, AStr);
          if Result then
            if AStr < ACalculator.Strings[AEnumStackRec.Value.vString] then
              AEnumStackRec.Value.vString := ACalculator.AddString(AStr);
        end
    else
      Result := True;
    end;
end;

function FnMin(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AEnumStackRec: rvgrEnumStackRec;
begin
  with AEnumStackRec do
  begin
    Value.ValueType := vteNull;
    First := True;
  end;
  Result := ACalculator.EnumStack(AParamCount, AEnumStackRec, FnMinCallback, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
    Result := ACalculator.CopyValue(AEnumStackRec.Value, AValue.Value);
end;

function FnMaxCallback(ACalculator: TvgrFormulaCalculator; AValue: pvteFormulaValue; var AEnumStackRec: rvgrEnumStackRec): Boolean;
var
  AInt: Integer;
  AExt: Extended;
  AStr: string;
begin
  if AEnumStackRec.First and (AValue.ValueType <> vteNull) then
  begin
    ACalculator.CopyValue(AValue^, AEnumStackRec.Value);
    AEnumStackRec.First := False;
    Result := True;
  end
  else
    case AEnumStackRec.Value.ValueType of
      vteInteger:
        begin
          Result := ACalculator.ToInteger(AValue^, AInt);
          if Result then
            if AInt > AEnumStackRec.Value.vInteger then
              AEnumStackRec.Value.vInteger := AInt;
        end;
      vteExtended:
        begin
          Result := ACalculator.ToExtended(AValue^, AExt);
          if Result then
            if AExt > AEnumStackRec.Value.vExtended then
              AEnumStackRec.Value.vExtended := AExt;
        end;
      vteString:
        begin
          Result := ACalculator.ToString(AValue^, AStr);
          if Result then
            if AStr > ACalculator.Strings[AEnumStackRec.Value.vString] then
              AEnumStackRec.Value.vString := ACalculator.AddString(AStr);
        end
    else
      Result := True;
    end;
end;

function FnMax(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AEnumStackRec: rvgrEnumStackRec;
begin
  with AEnumStackRec do
  begin
    Value.ValueType := vteNull;
    First := True;
  end;
  Result := ACalculator.EnumStack(AParamCount, AEnumStackRec, FnMaxCallback, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
    Result := ACalculator.CopyValue(AEnumStackRec.Value, AValue.Value);
end;

function FnRow(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaItem;
begin
  if AParamCount = 0 then
  begin
    AValue.Value.ValueType := vteInteger;
    AValue.Value.vInteger := AFormulaRow + 1;
    Result := True;
  end
  else
  begin
    Result := ACalculator.Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow) and (AItem.ItemType = vteitCellRef);
    if Result then
    begin
      AValue.Value.ValueType := vteInteger;
      if (AItem.CellRef.Flags and rvteCellRefRowAbsolute) <> 0 then
        AValue.Value.vInteger := AItem.CellRef.Row + 1
      else
        AValue.Value.vInteger := AItem.CellRef.Row + AFormulaRow + 1;
    end
    else
      ACalculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidOperands));
  end;
end;

function FnColumn(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaItem;
begin
  if AParamCount = 0 then
  begin
    AValue.Value.ValueType := vteInteger;
    AValue.Value.vInteger := AFormulaCol + 1;
    Result := True;
  end
  else
  begin
    Result := ACalculator.Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow) and (AItem.ItemType = vteitCellRef);
    if Result then
    begin
      AValue.Value.ValueType := vteInteger;
      if (AItem.CellRef.Flags and rvteCellRefColAbsolute) <> 0 then
        AValue.Value.vInteger := AItem.CellRef.Col + 1
      else
        AValue.Value.vInteger := AItem.CellRef.Col + AFormulaCol + 1;
    end
    else
      ACalculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidOperands));
  end;
end;

function FnIndirect(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := CompileCellRef(S,
                             AValue.CellRef.Row,
                             AValue.CellRef.Col,
                             AValue.CellRef.Flags,
                             rvteCellRefColAbsolute,
                             rvteCellRefRowAbsolute,
                             AFormulaRow,
                             AFormulaCol);
    if Result then
    begin
      AValue.CellRef.Sheet := AFormulaSheet;
      AValue.CellRef.Workbook := 0;
      AValue.ItemType := vteitCellRef;
    end
    else
      ACalculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidOperands));
  end;
end;

/////////////////////////////////////////////////
//
// TvgrWBStrings
//
/////////////////////////////////////////////////
constructor TvgrWBStrings.Create;
begin
  inherited Create;
  FLastEmptyStringIndex := -1;
  FStrings := TStringList.Create;
  FStrings.AddObject('', Pointer(1))
end;

destructor TvgrWBStrings.Destroy;
begin
  FStrings.Free;
  inherited;
end;

procedure TvgrWBStrings.Clear;
begin
  FLastEmptyStringIndex := -1;
  FStrings.Clear;
  FStrings.AddObject('', Pointer(1))
end;

{$IFDEF VGR_DS_DEBUG}
function TvgrWBStrings.DebugInfo: TvgrDebugInfo;
var
  I, ARefCountZeroCount: Integer;
begin
  Result := TvgrDebugInfo.Create;
  Result.Add('Count', FStrings.Count, 'Количество уникальных строк');
  ARefCountZeroCount := 0;
  for I := 0 to FStrings.Count - 1 do
    if Integer(FStrings.Objects[I]) = 0 then
      Inc(ARefCountZeroCount);
  Result.Add('(RefCount = 0)', ARefCountZeroCount, 'Количество свободных строк, у которых RefCount = 0');
end;
{$ENDIF}

function TvgrWBStrings.GetItem(Index: Integer): string;
begin
  Result := FStrings[Index];
end;

function TvgrWBStrings.GetCount: Integer;
begin
  Result := FStrings.Count;
end;

procedure TvgrWBStrings.SaveToStream(AStream: TStream);
var
  I, ABuf, ACount: Integer;
begin
  ACount := FStrings.Count;
  AStream.Write(ACount, 4);
  for I := 0 to ACount - 1 do
  begin
    ABuf := Integer(FStrings.Objects[I]);
    AStream.Write(ABuf, 4);
    ABuf := Length(FStrings[I]);
    AStream.Write(ABuf, 4);
    if ABuf > 0 then
      AStream.Write((@FStrings[I][1])^, ABuf);
  end;
end;

procedure TvgrWBStrings.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
var
  S: string;
  I, ACount, ABuf, AObject: Integer;
begin
  FLastEmptyStringIndex := -1;
  FStrings.Clear;
  AStream.Read(ACount, 4);
  for I := 0 to ACount - 1 do
  begin
    AStream.Read(AObject, 4);
    AStream.Read(ABuf, 4);
    if ABuf > 0 then
    begin
      SetLength(S, ABuf);
      AStream.Read((@S[1])^, ABuf);
    end
    else
      S := '';
    FStrings.AddObject(S, Pointer(AObject));
  end;
end;

function TvgrWBStrings.FindOrAdd(const s: string): Integer;
begin
  if FLastEmptyStringIndex = -1 then
  begin
    Result := FStrings.AddObject(s, Pointer(1));
  end
  else
  begin
    Result := FLastEmptyStringIndex;
    FStrings.Objects[Result] := Pointer(Integer(FStrings.Objects[Result]) + 1);
    FStrings[Result] := s;
    FLastEmptyStringIndex := -1;
  end;
end;

procedure TvgrWBStrings.Release(AIndex: Integer);
begin
  if AIndex > 0 then
  begin
    FStrings.Objects[AIndex] := Pointer(Integer(FStrings.Objects[AIndex]) - 1);
    if Integer(FStrings.Objects[AIndex]) = 0 then
    begin
      // need to delete string
      FStrings[AIndex] := '';
      FLastEmptyStringIndex := AIndex;
    end;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrCalculatorStack
//
/////////////////////////////////////////////////
constructor TvgrCalculatorStack.Create(ACalculator: TvgrFormulaCalculator);
begin
  inherited Create;
  FCalculator := ACalculator;
  FList := TList.Create;
  FCurPos := -1;
end;

destructor TvgrCalculatorStack.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TvgrCalculatorStack.GetItem(Index: Integer): pvteFormulaItem;
begin
  Result := pvteFormulaItem(FList[Index]);
end;

procedure TvgrCalculatorStack.Clear;
var
  I: integer;
begin
  for I := 0 to FList.Count - 1 do
    FreeMem(FList[I]);
  FList.Clear;
  FCurPos := -1;
end;

procedure TvgrCalculatorStack.Reset;
begin
  FCurPos := -1;
end;

function TvgrCalculatorStack.PushItem: pvteFormulaItem;
begin
  Inc(FCurPos);
  if FCurPos = FList.Count then
  begin
    GetMem(Result, sizeof(rvteFormulaItem));
    FList.Add(Result);
  end
  else
    Result := FList[FCurPos];
end;

procedure TvgrCalculatorStack.Push(AItem: pvteFormulaItem);
var
  ANew: pvteFormulaItem;
begin
  ANew := PushItem;
  System.Move(AItem^, ANew^, sizeof(rvteFormulaItem));
end;

function TvgrCalculatorStack.Pop(var AItem: pvteFormulaItem): Boolean;
begin
  Result := FCurPos >= 0;
  if Result then
  begin
    AItem := FList[FCurPos];
    Dec(FCurPos);
  end
  else
    Calculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_StackPopError));
end;

function TvgrCalculatorStack.Pop(AItemsToPop: Integer): Boolean;
begin
  Result := FCurPos + 1 >= AItemsToPop;
  if Result then
    FCurPos := FCurPos - AItemsToPop
  else
    Calculator.SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_StackPopError));
end;

/////////////////////////////////////////////////
//
// TvgrRegisteredFunction
//
/////////////////////////////////////////////////
constructor TvgrRegisteredFunction.Create(const AName: string; AProc: TvgrFormulaFunctionProc; AMinParams, AMaxParams: Integer);
begin
  inherited Create;
  FName := AName;
  FProc := AProc;
  FMinParams := AMinParams;
  FMaxParams := AMaxParams;
end;

/////////////////////////////////////////////////
//
// TvgrRegisteredFunctions
//
/////////////////////////////////////////////////
function TvgrRegisteredFunctions.GetItem(Index: Integer): TvgrRegisteredFunction;
begin
  Result := TvgrRegisteredFunction(inherited Items[Index]);
end;

function TvgrRegisteredFunctions.GetFunctionIndex(const AFuncName: string): Integer;
begin
  Result := 0;
  while (Result < Count) and
        (AnsiCompareText(Items[Result].Name, AFuncName) <> 0) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrRegisteredFunctions.GetFunctionName(AFuncIndex: Integer): string;
begin
  Result := Items[AFuncIndex].Name;
end;

procedure TvgrRegisteredFunctions.RegisterFunction(const AFuncName: string; AFuncProc: Pointer; AMinParams, AMaxParams: Integer);
begin
  if GetFunctionIndex(AFuncName) = -1 then
    Add(TvgrRegisteredFunction.Create(AFuncName, AFuncProc, AMinParams, AMaxParams));
end;

/////////////////////////////////////////////////
//
// TvgrFormulaCalculator
//
/////////////////////////////////////////////////
constructor TvgrFormulaCalculator.Create(AOwner: IvgrFormulaCalculatorOwner);
begin
  inherited Create;
  FOwner := AOwner;
  FStrings := TStringList.Create;
end;

destructor TvgrFormulaCalculator.Destroy;
begin
  FStrings.Free;
  inherited;
end;

function TvgrFormulaCalculator.GetString(I: Integer): string;
begin
  Result := FStrings[I];
end;

function TvgrFormulaCalculator.GetStringCount: Integer;
begin
  Result := FStrings.Count;
end;

function TvgrFormulaCalculator.AddString(const S: string): Integer;
begin
  FStrings.Add(S);
  Result := FStrings.Count - 1;
end;

procedure TvgrFormulaCalculator.ClearStrings;
begin
  FStrings.Clear;
end;

procedure TvgrFormulaCalculator.FromInteger(var AFormulaValue: rvteFormulaValue; AValue: Integer);
begin
  AFormulaValue.ValueType := vteInteger;
  AFormulaValue.vInteger := AValue;
end;

procedure TvgrFormulaCalculator.FromExtended(var AFormulaValue: rvteFormulaValue; AValue: Extended);
begin
  AFormulaValue.ValueType := vteExtended;
  AFormulaValue.vExtended := AValue;
end;

procedure TvgrFormulaCalculator.FromDateTime(var AFormulaValue: rvteFormulaValue; AValue: TDateTime);
begin
  AFormulaValue.ValueType := vteExtended;
  AFormulaValue.vExtended := AValue;
end;

procedure TvgrFormulaCalculator.FromString(var AFormulaValue: rvteFormulaValue; const AValue: string);
begin
  AFormulaValue.ValueType := vteString;
  AFormulaValue.vString := AddString(AValue);
end;

procedure TvgrFormulaCalculator.FromNull(var AFormulaValue: rvteFormulaValue);
begin
  AFormulaValue.ValueType := vteNull;
end;

function TvgrFormulaCalculator.ToInteger(const AFormulaValue: rvteFormulaValue; var AValue: Integer): Boolean;
var
  ACode: Integer;
begin
  Result := True;
  with AFormulaValue do
    case ValueType of
      vteNull: AValue := 0;
      vteInteger: AValue := vInteger;
      vteExtended: AValue := Round(vExtended);
      vteString:
        begin
          Val(Strings[vString], AValue, ACode);
          Result := ACode = 0;
        end;
    end;
end;

function TvgrFormulaCalculator.ToExtended(const AFormulaValue: rvteFormulaValue; var AValue: Extended): Boolean;
begin
  Result := True;
  with AFormulaValue do
    case ValueType of
      vteNull: AValue := 0;
      vteInteger: AValue := vInteger;
      vteExtended: AValue := vExtended;
      vteString: Result := TextToFloat(PChar(Strings[vString]), AValue, fvExtended);
    end;
end;

function TvgrFormulaCalculator.ToDateTime(const AFormulaValue: rvteFormulaValue; var AValue: TDateTime): Boolean;
begin
  Result := True;
  with AFormulaValue do
    case ValueType of
      vteNull: AValue := 0;
      vteInteger: AValue := vInteger;
      vteExtended: AValue := vExtended;
      vteString: Result := TryStrToDateTime(Strings[vString], AValue);
    end;
end;

function TvgrFormulaCalculator.ToString(const AFormulaValue: rvteFormulaValue; var AValue: string): Boolean;
begin
  Result := True;
  with AFormulaValue do
    case ValueType of
      vteNull: AValue := '';
      vteInteger: AValue := IntToStr(vInteger);
      vteExtended: AValue := FloatToStr(vExtended);
      vteString: AValue := Strings[vString];
    end;
end;

function TvgrFormulaCalculator.CopyValue(const ASourceValue: rvteFormulaValue; var ADestValue: rvteFormulaValue): Boolean;
begin
  ADestValue := ASourceValue;
  Result := True;
end;

{
function TvgrFormulaCalculator.FormulaValueToVariant(const AValue: rvteFormulaValue): Variant;
begin
  case AValue.ValueType of
    vteInteger:
      Result := AValue.vInteger;
    vteExtended:
      Result := AValue.vExtended;
    vteString:
      Result := FStack.Strings[AValue.vString];
  else
    Result := Null;
  end;
end;
}

procedure TvgrFormulaCalculator.SetError(const AErrorMessage: string);
begin
  FErrorMessage := AErrorMessage;
end;

function TvgrFormulaCalculator.GetCell(const ACellRef: rvteFormulaCellRef;
                 AFormulaSheet, AFormulaCol, AFormulaRow: Integer;
                 var ACellX, ACellY, ACellSheet: Integer): Boolean;
begin
  if (ACellRef.Flags and rvteCellRefColAbsolute) <> 0 then
    ACellX := ACellRef.Col
  else
    ACellX := AFormulaCol + ACellRef.Col;
  if (ACellRef.Flags and rvteCellRefRowAbsolute) <> 0 then
    ACellY := ACellRef.Row
  else
    ACellY := AFormulaRow + ACellRef.Row;
  if ACellRef.Sheet = -1 then
    ACellSheet := AFormulaSheet
  else
    ACellSheet := ACellRef.Sheet;
  Result := (ACellX >= 0) and (ACellY >= 0);
  if not Result then
    SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidCellReference));
end;

function TvgrFormulaCalculator.GetRange(const ARangeRef: rvteFormulaRangeRef;
                  AFormulaSheet, AFormulaCol, AFormulaRow: Integer;
                  var ARangeLeft, ARangeTop, ARangeRight, ARangeBottom, ARangeSheet: Integer): Boolean;
begin
  if (ARangeRef.Flags and rvteRangeRefLeftColAbsolute) <> 0 then
    ARangeLeft := ARangeRef.LeftCol
  else
    ARangeLeft := AFormulaCol + ARangeRef.LeftCol;
  if (ARangeRef.Flags and rvteRangeRefTopRowAbsolute) <> 0 then
    ARangeTop := ARangeRef.TopRow
  else
    ARangeTop := AFormulaRow + ARangeRef.TopRow;

  if (ARangeRef.Flags and rvteRangeRefRightColAbsolute) <> 0 then
    ARangeRight := ARangeRef.RightCol
  else
    ARangeRight := AFormulaCol + ARangeRef.RightCol;
  if (ARangeRef.Flags and rvteRangeRefBottomRowAbsolute) <> 0 then
    ARangeBottom := ARangeRef.BottomRow
  else
    ARangeBottom := AFormulaRow + ARangeRef.BottomRow;

  if ARangeRef.Sheet = -1 then
    ARangeSheet := AFormulaSheet
  else
    ARangeSheet := ARangeRef.Sheet;

  Result := (ARangeLeft >= 0) and (ARangeTop >= 0) and
            (ARangeLeft <= ARangeRight) and (ARangeTop <= ARangeBottom);
  if not Result then
    SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidCellReference));
end;

function TvgrFormulaCalculator.CalculateCellValue(const ACellRef: rvteFormulaCellRef; var AValue: rvteFormulaValue; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ASheet, ACol, ARow: Integer;
begin
  Result := GetCell(ACellRef, AFormulaSheet, AFormulaCol, AFormulaRow,
                    ACol, ARow, ASheet);
  if Result then
    Result := Owner.GetCellValue(Self, AValue, ASheet, ACol, ARow)
end;

function TvgrFormulaCalculator.CalculateFunction(AItem: pvteFormulaItem; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ACalcItem: rvteFormulaItem;
begin
  Result := (AItem.Func.Id >= 0) and (AItem.Func.Id < RegisteredFunctions.Count);
  if Result then
    with RegisteredFunctions[AItem.Func.Id] do
    begin
      Result := ((MinParams = -1) or
                 (AItem.Func.ParamCount >= MinParams)) and
                ((MaxParams = -1) or
                 (AItem.Func.ParamCount <= MaxParams));
      if Result then
      begin
        ACalcItem.ItemType := vteitValue;
        Result := RegisteredFunctions[AItem.Func.Id].Proc(Self,
                                                          ACalcItem,//ACalcItem.Value,
                                                          AItem.Func.ParamCount,
                                                          AFormulaSheet,
                                                          AFormulaCol,
                                                          AFormulaRow);
        if Result then
          FStack.Push(@ACalcItem);
      end
      else
        SetError(Format(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidFunctionParamsCount),
                        [AItem.Func.ParamCount, MinParams, MaxParams]));
    end
  else
    SetError(Format(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidFunctionIndex),
                    [AItem.Func.Id]));
end;

function TvgrFormulaCalculator.CalculateOperator(AItem: pvteFormulaItem; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
type
  TvgropType = (tInt, tExt, tStr, tErr);
const
  AOperatorsTable : array [vteMinOperator_ptg..vteMaxOperator_ptg, TvteFormulaValueType, TvteFormulaValueType] of TvgropType =
  (
       {vtefoAdd}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tErr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tInt, tExt, tErr)),
      {vtefoSub}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tErr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tInt, tExt, tErr)),
      {vtefoMul}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tInt, tErr),
{Int} (tInt, tInt, tExt, tExt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tInt, tExt, tErr)),
      {vtefoDiv}
      {Nll} {Int} {Ext} {Str}
{Nll}((tErr, tInt, tInt, tErr),
{Int} (tErr, tExt, tExt, tExt),
{Ext} (tErr, tExt, tExt, tExt),
{Str} (tErr, tExt, tExt, tErr)),
      {vtefoPower}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tInt, tErr),
{Int} (tInt, tExt, tExt, tExt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tExt, tExt, tErr)),
      {vtefoConcat}
      {Nll} {Int} {Ext} {Str}
{Nll}((tStr, tStr, tStr, tStr),
{Int} (tStr, tStr, tStr, tStr),
{Ext} (tStr, tStr, tStr, tStr),
{Str} (tStr, tStr, tStr, tStr)),
      {vtefoLT}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tStr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tStr, tInt, tExt, tStr)),
      {vtefoLE}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tStr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tStr, tInt, tExt, tStr)),
      {vtefoEQ}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tStr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tStr, tInt, tExt, tStr)),
      {vtefoGE}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tStr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tStr, tInt, tExt, tStr)),
      {vtefoGT}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tStr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tStr, tInt, tExt, tStr)),
      {vtefoNE}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tExt, tStr),
{Int} (tInt, tInt, tExt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tStr, tInt, tExt, tStr)),

      {skipped}
      {Nll} {Int} {Ext} {Str}
{Nll}((tErr, tErr, tErr, tErr),
{Int} (tErr, tErr, tErr, tErr),
{Ext} (tErr, tErr, tErr, tErr),
{Str} (tErr, tErr, tErr, tErr)),
      {skipped}
      {Nll} {Int} {Ext} {Str}
{Nll}((tErr, tErr, tErr, tErr),
{Int} (tErr, tErr, tErr, tErr),
{Ext} (tErr, tErr, tErr, tErr),
{Str} (tErr, tErr, tErr, tErr)),
      {skipped}
      {Nll} {Int} {Ext} {Str}
{Nll}((tErr, tErr, tErr, tErr),
{Int} (tErr, tErr, tErr, tErr),
{Ext} (tErr, tErr, tErr, tErr),
{Str} (tErr, tErr, tErr, tErr)),

      {vtefoUPlus}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tInt, tInt),
{Int} (tInt, tInt, tInt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tErr, tErr, tErr)),

      {vtefoUMinus}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tInt, tInt),
{Int} (tInt, tInt, tInt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tErr, tErr, tErr)),
      {vtefoPercent}
      {Nll} {Int} {Ext} {Str}
{Nll}((tInt, tInt, tInt, tInt),
{Int} (tInt, tInt, tInt, tInt),
{Ext} (tExt, tExt, tExt, tExt),
{Str} (tErr, tErr, tErr, tErr))
);

var
  AItem1, AItem2: pvteFormulaValue;
  AOpInt1, AOpInt2: Integer;
  AOpExt1, AOpExt2: Extended;
  AOpStr1, AopStr2: string;
  AResultItem: rvteFormulaItem;
begin
  AResultItem.ItemType := vteitValue;
  if AItem.Operator.Id in vteFormulaUnaryOperatorsIds then
  begin
    AItem2 := nil;
    Result := Pop(AItem1, AFormulaSheet, AFormulaCol, AFormulaRow);
    if Result then
    begin
      case AOperatorsTable[AItem.Operator.Id, AItem1.ValueType, vteInteger] of
        tInt:
          begin
            AResultItem.Value.ValueType := vteInteger;
            Result := ToInteger(AItem1^, AOpInt1);
            if Result then
              Result := AOperatorsIntProcs[AItem.Operator.Id](Self, AResultItem.Value, AOpInt1, 0);
          end;
        tExt:
          begin
            AResultItem.Value.ValueType := vteExtended;
            Result := ToExtended(AItem1^, AOpExt1);
            if Result then
              Result := AOperatorsExtProcs[AItem.Operator.Id](Self, AResultItem.Value, AOpExt1, 0);
          end;
        tStr:
          begin
            AResultItem.Value.ValueType := vteString;
            Result := ToString(AItem1^, AOpStr1);
            if Result then
            begin
              Result := AOperatorsStrProcs[AItem.Operator.Id](Self, AResultItem.Value, AOpStr1, '');
            end;
          end;
        tErr:
          begin
            SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidOperands));
            Result := False;
          end;
      end;
    end;
  end
  else
  begin
    Result := Pop(AItem2, AFormulaSheet, AFormulaCol, AFormulaRow) and
              Pop(AItem1, AFormulaSheet, AFormulaCol, AFormulaRow);
    if Result then
    begin
      case AOperatorsTable[AItem.Operator.Id, AItem1.ValueType, AItem2.ValueType] of
        tInt:
          begin
            AResultItem.Value.ValueType := vteInteger;
            Result := ToInteger(AItem1^, AOpInt1) and
                      ToInteger(AItem2^, AOpInt2);
            if Result then
              Result := AOperatorsIntProcs[AItem.Operator.Id](Self, AResultItem.Value, AOpInt1, AOpInt2);
          end;
        tExt:
          begin
            AResultItem.Value.ValueType := vteExtended;
            Result := ToExtended(AItem1^, AOpExt1) and
                      ToExtended(AItem2^, AOpExt2);
            if Result then
              Result := AOperatorsExtProcs[AItem.Operator.Id](Self, AResultItem.Value, AOpExt1, AOpExt2);
          end;
        tStr:
          begin
            AResultItem.Value.ValueType := vteString;
            Result := ToString(AItem1^, AOpStr1) and
                      ToString(AItem2^, AOpStr2);
            if Result then
              Result := AOperatorsStrProcs[AItem.Operator.Id](Self, AResultItem.Value, AOpStr1, AOpStr2);
          end;
        tErr:
          begin
            SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidOperands));
            Result := False;
          end;
      end;
    end;
  end;

  if Result then
    FStack.Push(@AResultItem);
end;

function TvgrFormulaCalculator.EnumStack(ACount: Integer; var AEnumRec: rvgrEnumStackRec; ACallback: TvgrFormulaCalculatorEnumStackProc; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  I, ASheet: Integer;
  AItem: pvteFormulaItem;
  ARangeRect: TRect;
begin
  Result := FStack.FCurPos + 1 >= ACount;
  if Result then
  begin
    for I := 0 to ACount - 1 do
    begin
      AItem := FStack.Items[FStack.FCurPos - (ACount - I - 1)];
      case AItem.ItemType of
        vteitValue:
          Result := ACallback(Self, @AItem.Value, AEnumRec);
        vteitMissArg:
          Result := ACallback(Self, nil, AEnumRec);
        vteitRangeRef:
          begin
            Result := GetRange(AItem.RangeRef, AFormulaSheet, AFormulaCol, AFormulaRow,
                               ARangeRect.Left,
                               ARangeRect.Top,
                               ARangeRect.Right,
                               ARangeRect.Bottom, ASheet);
            if Result then
              Result := Owner.EnumRangeValues(Self, AEnumRec, ACallback, ASheet, ARangeRect);
          end;
      else
        begin
          SetError(vgrLoadStr(svgrid_vgr_FormulaCalculator_InvalidOperands));
          Result := False;
        end;
      end;
      if not Result then
        exit;
    end;
    FStack.Pop(ACount);
  end;
end;

function TvgrFormulaCalculator.Pop(var AItem: pvteFormulaItem; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
begin
  Result := FStack.Pop(AItem)
end;

function TvgrFormulaCalculator.Pop(var AValue: pvteFormulaValue; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ATempItem: pvteFormulaItem;
begin
  Result := FStack.Pop(ATempItem);
  if Result then
  begin
    case ATempItem.ItemType of
      vteitValue:
        begin
          AValue := @ATempItem.Value;
          Result := True;
        end;
      vteitCellRef:
        begin
          Result := CalculateCellValue(ATempItem.CellRef, ATempItem.Value, AFormulaSheet, AFormulaCol, AFormulaRow);
          if Result then
            AValue := @ATempItem.Value;
        end;
    else
      Result := False;
    end;
  end;
end;

function TvgrFormulaCalculator.Pop(var AInt: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaValue;
begin
  Result := Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ToInteger(AItem^, AInt);
end;

function TvgrFormulaCalculator.Pop(var AStr: string; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaValue;
begin
  Result := Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ToString(AItem^, AStr);
end;

function TvgrFormulaCalculator.Pop(var AExtended: Extended; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaValue;
begin
  Result := Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ToExtended(AItem^, AExtended);
end;

function TvgrFormulaCalculator.Pop(var ADateTime: TDateTime; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaValue;
begin
  Result := Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ToDateTime(AItem^, ADateTime);
end;

function TvgrFormulaCalculator.Calculate(AFormula: pvteFormula; AFormulaSize: Integer; var AValue: Variant; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  I: Integer;
  AItem: pvteFormulaItem;
  ATempValue: pvteFormulaValue;
  AStrItem: rvteFormulaItem;
begin
  Result := True;
  AStrItem.ItemType := vteitValue;
  AStrItem.Value.ValueType := vteString;
  for I := 0 to AFormulaSize - 1 do
  begin
    AItem := @AFormula[I];
    case AItem.ItemType of
      vteitCellRef, vteitRangeRef, vteitValue, vteitMissArg:
        begin
          if (AItem.ItemType = vteitValue) and (AItem.Value.ValueType = vteString) then
          begin
            AStrItem.Value.vString := AddString(ExternalWBStrings[AItem.Value.vString]);
            Stack.Push(@AStrItem);
          end
          else
            Stack.Push(AItem)
        end;
      vteitOperator:
        Result := CalculateOperator(AItem, AFormulaSheet, AFormulaCol, AFormulaRow);
      vteitFunc:
        Result := CalculateFunction(AItem, AFormulaSheet, AFormulaCol, AFormulaRow);
    end;
    if not Result then
    begin
      AValue := Null;
      exit;
    end;
  end;
  
  Result := Pop(ATempValue, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    case ATempValue.ValueType of
      vteInteger:
        AValue := ATempValue.vInteger;
      vteExtended:
        AValue := ATempValue.vExtended;
      vteString:
        AValue := Strings[ATempValue.vString];
      else
        AValue := Null;
    end;
  end;
end;

// date functions
function FnNow(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
begin
  ACalculator.FromDateTime(AValue.Value, Now);
  Result := True;
end;

function FnDate(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AYear, AMonth, ADay: Integer;
  ADate: TDateTime;
begin
  Result := ACalculator.Pop(ADay, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AMonth, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AYear, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := TryEncodeDate(AYear, AMonth, ADay, ADate);
    if Result then
      ACalculator.FromDateTime(AValue.Value, ADate);
  end;
end;

function FnDateValue(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
  ADate: TDateTime;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := TryStrToDateTime(S, ADate);
    if Result then
      ACalculator.FromDateTime(AValue.Value, ADate);
  end;
end;

function FnDay(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ADate: TDateTime;
  AYear, AMonth, ADay: Word;
begin
  Result := ACalculator.Pop(ADate, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    DecodeDate(ADate, AYear, AMonth, ADay);
    ACalculator.FromInteger(AValue.Value, ADay);
  end;
end;

function FnHour(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ATime: TDateTime;
  AHour, AMin, ASec, AMSec: Word;
begin
  Result := ACalculator.Pop(ATime, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    DecodeTime(ATime, AHour, AMin, ASec, AMSec);
    ACalculator.FromInteger(AValue.Value, AHour);
  end;
end;

function FnMonth(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ADate: TDateTime;
  AYear, AMonth, ADay: Word;
begin
  Result := ACalculator.Pop(ADate, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    DecodeDate(ADate, AYear, AMonth, ADay);
    ACalculator.FromInteger(AValue.Value, AMonth);
  end;
end;

function FnMinute(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ATime: TDateTime;
  AHour, AMin, ASec, AMSec: Word;
begin
  Result := ACalculator.Pop(ATime, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    DecodeTime(ATime, AHour, AMin, ASec, AMSec);
    ACalculator.FromInteger(AValue.Value, AMin);
  end;
end;

function FnSecond(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ATime: TDateTime;
  AHour, AMin, ASec, AMSec: Word;
begin
  Result := ACalculator.Pop(ATime, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    DecodeTime(ATime, AHour, AMin, ASec, AMSec);
    ACalculator.FromInteger(AValue.Value, ASec);
  end;
end;

function FnTime(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ASecond, AMinute, AHour: Integer;
  ATime: TDateTime;
begin
  Result := ACalculator.Pop(ASecond, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AMinute, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AHour, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := TryEncodeTime(AHour, AMinute, ASecond, 0, ATime);
    if Result then
      ACalculator.FromDateTime(AValue.Value, ATime);
  end;
end;

function FnTimeValue(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
  ADate: TDateTime;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := TryStrToTime(S, ADate);
    if Result then
      ACalculator.FromDateTime(AValue.Value, ADate);
  end;
end;

function FnToday(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
begin
  ACalculator.FromDateTime(AValue.Value, Date);
  Result := True;
end;

function FnWeekDay(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ADate: TDateTime;
begin
  Result := ACalculator.Pop(ADate, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromInteger(AValue.Value, DayOfWeek(ADate));
  end;
end;

function FnYear(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  ADate: TDateTime;
  AYear, AMonth, ADay: Word;
begin
  Result := ACalculator.Pop(ADate, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    DecodeDate(ADate, AYear, AMonth, ADay);
    ACalculator.FromInteger(AValue.Value, AYear);
  end;
end;

// math functions
function FnAbs(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AValueItem: pvteFormulaValue;
  AExt: Extended;
begin
  Result := ACalculator.Pop(AValueItem, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    case AValueItem.ValueType of
      vteNull: ACalculator.FromNull(AValue.Value);
      vteInteger: ACalculator.FromInteger(AValue.Value, Abs(AValueItem.vInteger));
      vteExtended: ACalculator.FromExtended(AValue.Value, Abs(AValueItem.vExtended));
      vteString:
        begin
          Result := ACalculator.ToExtended(AValueItem^, AExt);
          if Result then
            ACalculator.FromExtended(AValue.Value, AExt);
        end;
      else
        Result := False;
    end;
  end;
end;

function FnRound(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AExt: Extended;
  ANumDigits: Integer;
begin
  Result := ACalculator.Pop(ANumDigits, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AExt, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromExtended(AValue.Value, RoundEx(AExt, ANumDigits));
  end;
end;

function FnSign(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AExt: Extended;
  ARes: Integer;
begin
  Result := ACalculator.Pop(AExt, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    if IsZero(AExt) then
      ARes := 0
    else
    begin
      if AExt < 0 then
        ARes := -1
      else
        ARes := 1;
    end;
    ACalculator.FromInteger(AValue.Value, ARes)
  end;
end;

// lookup and reference functions
function FnRows(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaItem;
  ACol, ARow, ASheet: Integer;
  ARange: TRect;
begin
  Result := ACalculator.Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    case AItem.ItemType of
      vteitCellRef:
        begin
          Result := ACalculator.GetCell(AItem.CellRef, AFormulaSheet, AFormulaCol, AFormulaRow,
                                        ACol, ARow, ASheet);
          if Result then
            ACalculator.FromInteger(AValue.Value, 1);
        end;
      vteitRangeRef:
        begin
          Result := ACalculator.GetRange(AItem.RangeRef, AFormulaSheet, AFormulaCol, AFormulaRow,
                                         ARange.Left, ARange.Top, ARange.Right, ARange.Bottom, ASheet);
          if Result then
            ACalculator.FromInteger(AValue.Value, ARange.Bottom - ARange.Top + 1);
        end;
      else
        Result := False;
    end;
  end;
end;

function FnColumns(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AItem: pvteFormulaItem;
  ACol, ARow, ASheet: Integer;
  ARange: TRect;
begin
  Result := ACalculator.Pop(AItem, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    case AItem.ItemType of
      vteitCellRef:
        begin
          Result := ACalculator.GetCell(AItem.CellRef, AFormulaSheet, AFormulaCol, AFormulaRow,
                                        ACol, ARow, ASheet);
          if Result then
            ACalculator.FromInteger(AValue.Value, 1);
        end;
      vteitRangeRef:
        begin
          Result := ACalculator.GetRange(AItem.RangeRef, AFormulaSheet, AFormulaCol, AFormulaRow,
                                         ARange.Left, ARange.Top, ARange.Right, ARange.Bottom, ASheet);
          if Result then
            ACalculator.FromInteger(AValue.Value, ARange.Right - ARange.Left + 1);
        end;
      else
        Result := False;
    end;
  end;
end;

// text functions
function FnChar(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  AInt: Integer;
begin
  Result := ACalculator.Pop(AInt, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := AInt < 256;
    if Result then
      ACalculator.FromString(AValue.Value, char(AInt));
  end;
end;

function FnCode(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    Result := Length(S) > 0;
    if Result then
      ACalculator.FromInteger(AValue.Value, ord(S[1]));
  end;
end;

function FnExact(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S1, S2: string;
begin
  Result := ACalculator.Pop(S1, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(S2, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromInteger(AValue.Value, Integer(S1 = S2));
  end;
end;

function FnLeft(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
  ALeft: Integer;
begin
  if AParamCount = 1 then
  begin
    ALeft := 1;
    Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow)
  end
  else
    Result := ACalculator.Pop(ALeft, AFormulaSheet, AFormulaCol, AFormulaRow) and
              ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);

  if Result then
  begin
    ACalculator.FromString(AValue.Value, Copy(S, 1, ALeft));
  end;
end;

function FnLen(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromInteger(AValue.Value, Length(S));
  end;
end;

function FnLower(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromString(AValue.Value, AnsiLowerCase(S));
  end;
end;

function FnMid(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
  AChars, AStart: Integer;
begin
  Result := ACalculator.Pop(AChars, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(AStart, AFormulaSheet, AFormulaCol, AFormulaRow) and
            ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromString(AValue.Value, Copy(S, AStart, AChars));
  end;
end;

function FnRight(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
  ARight: Integer;
begin
  if AParamCount = 1 then
  begin
    ARight := 1;
    Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow)
  end
  else
    Result := ACalculator.Pop(ARight, AFormulaSheet, AFormulaCol, AFormulaRow) and
              ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);

  if Result then
  begin
    S := Copy(S, Max(1, Length(S) - ARight + 1), ARight);
    ACalculator.FromString(AValue.Value, S)
  end;
end;

function FnUpper(ACalculator: TvgrFormulaCalculator; var AValue: rvteFormulaItem; AParamCount: Integer; AFormulaSheet, AFormulaCol, AFormulaRow: Integer): Boolean;
var
  S: string;
begin
  Result := ACalculator.Pop(S, AFormulaSheet, AFormulaCol, AFormulaRow);
  if Result then
  begin
    ACalculator.FromString(AValue.Value, AnsiUpperCase(S));
  end;
end;

initialization

  RegisteredFunctions := TvgrRegisteredFunctions.Create;
  RegisterFunction('Count', @FnCount, 0, -1);
  RegisterFunction('If', @FnIf, 3, 3);
  RegisterFunction('Sum', @FnSum, 1, -1);
  RegisterFunction('Average', @FnAverage, 1, -1);
  RegisterFunction('Min', @FnMin, 1, -1);
  RegisterFunction('Max', @FnMax, 1, -1);

  RegisterFunction('Now', @FnNow, 0, 0);
  RegisterFunction('DateValue', @FnDateValue, 1, 1);
  RegisterFunction('Date', @FnDate, 3, 3);
  RegisterFunction('Day', @FnDay, 1, 1);
  RegisterFunction('Hour', @FnHour, 1, 1);
  RegisterFunction('Month', @FnMonth, 1, 1);
  RegisterFunction('Minute', @FnMinute, 1, 1);
  RegisterFunction('Second', @FnSecond, 1, 1);
  RegisterFunction('Time', @FnTime, 3, 3);
  RegisterFunction('TimeValue', @FnTimeValue, 1, 1);
  RegisterFunction('Today', @FnToday, 0, 0);
  RegisterFunction('WeekDay', @FnWeekDay, 1, 1);
  RegisterFunction('Year', @FnYear, 1, 1);

  RegisterFunction('Abs', @FnAbs, 1, 1);
  RegisterFunction('Round', @FnRound, 2, 2);
  RegisterFunction('Sign', @FnSign, 1, 1);

  RegisterFunction('Row', @FnRow, 0, 1);
  RegisterFunction('Column', @FnColumn, 0, 1);
  RegisterFunction('Indirect', @FnIndirect, 1, 1);
  RegisterFunction('Rows', @FnRows, 1, 1);
  RegisterFunction('Columns', @FnColumns, 1, 1);

  RegisterFunction('Char', @FnChar, 1, 1);
  RegisterFunction('Code', @FnCode, 1, 1);
  RegisterFunction('Exact', @FnExact, 2, 2);
  RegisterFunction('Left', @FnLeft, 1, 2);
  RegisterFunction('Len', @FnLen, 1, 1);
  RegisterFunction('Lower', @FnLower, 1, 1);
  RegisterFunction('Mid', @FnMid, 3, 3);
  RegisterFunction('Right', @FnRight, 1, 2);
  RegisterFunction('Upper', @FnUpper, 1, 1);

finalization

  FreeAndNil(RegisteredFunctions);

end.
