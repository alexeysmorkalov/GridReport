{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes which realize a compilation of formulas in TvgrWorkbook.
TvteExcelFormulaCompiler is main class, he implements a CompileFormula and
DecompileFormula methods which are used to convert of the string representation of the formula
to its binary representation and back.}
unit vgr_ExcelFormula;

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Classes, SysUtils;

{$I vtk.inc}

const
  SymbolA = 65;
  SymbolZ = 90;
  SymbolsAZ = SymbolZ - SymbolA + 1;

  ptgAdd = $03;
  ptgSub = $04;
  ptgMul = $05;
  ptgDiv = $06;
  ptgPower = $07;
  ptgConcat = $08;
  ptgLT = $09;
  ptgLE = $0A;
  ptgEQ = $0B;
  ptgGE = $0C;
  ptgGT = $0D;
  ptgNE = $0E;
  ptgUplus = $12;
  ptgUminus = $13;
  ptgPercent = $14;

  vtefoAdd = ptgAdd;
  vtefoSub = ptgSub;
  vtefoMul = ptgMul;
  vtefoDiv = ptgDiv;
  vtefoPower = ptgPower;
  vtefoConcat = ptgConcat;
  vtefoLT = ptgLT;
  vtefoLE = ptgLE;
  vtefoEQ = ptgEQ;
  vtefoGE = ptgGE;
  vtefoGT = ptgGT;
  vtefoNE = ptgNE;
  vtefoUplus = ptgUplus;
  vtefoUminus = ptgUminus;
  vtefoPercent = ptgPercent;

  vteMinOperator_ptg = ptgAdd;
  vteMaxOperator_ptg = ptgPercent;

  vteFormulaEndBracketChar = ')';
  vteFormulaStartBracketChar = '(';
  vteFormulaStringChar = '"';
  vteFormulaFuncParamsDelim = ';';
  vteFormulaPercentOperator = '%';
  vteFormulaUnaryPlusOperator = '+';
  vteFormulaUnaryMinusOperator = '-';

  vteFormulaUnaryOperatorsIds: set of Byte = [vtefoUplus, vtefoUminus, vtefoPercent];
  vteFormulaUnaryOperators : set of char = ['+', '-'];
  vteFormulaOperatorChars : set of char = ['>', '<', '=', '+', '-', '*', '/', '&', '%', '^'];
  vteFormulaStartIdentChars : set of char = ['a'..'z', 'A'..'Z', '''', '$', '['];
  vteFormulaIdentChars : set of char = ['a'..'z', 'A'..'Z', '_', '0'..'9', '.', '$', '@', '''', '!', ':', '[', ']'];

  vteOperatorsCount = 18;
  vteFormulaStartBracketOperatorIndex = 1;
  vteFormulaEndBracketOperatorIndex = 2;
  vteFormulaPercentOperatorIndex = 15;
  vteFormulaUnaryPlusOperatorIndex = 16;
  vteFormulaUnaryMinusOperatorIndex = 17;
  vteFormulaFunctionOperatorIndex = 18;

  vteFormulaFunctionPriority = 9;
  vteFormulaStartBracketPriority = 0;
  vteFormulaEndBracketPriority = 1;
  vteFormulaPercentOperatorPriority = 7; 

  svteFormulaCompileErrorInvalidBrackets = 'Invalid brackets';
  svteFormulaCompileErrorParameterWithoutFunction = 'Parameter without function';
  svteFormulaCompileErrorInvalidString = 'Invalid string';
  svteFormulaCompileErrorInvalidNumber = 'Invalid number [%s]';
  svteFormulaCompileErrorInvalidSymbol = 'Invalid symbol [%s]';
  svteFormulaCompileErrorUnknownOperator = 'Unknown operator [%s]';
  svteFormulaCompileErrorUnknownFunction = 'Unknown function [%s]';
  svteFormulaCompileErrorInvalidCellReference = 'Invalid cell reference [%s]';
  svteFormulaCompileErrorInvalidRangeReference = 'Invalid range reference [%s]';

  rvteCellRefColAbsolute = $01;
  rvteCellRefRowAbsolute = $02;

  rvteRangeRefLeftColAbsolute = $01;
  rvteRangeRefTopRowAbsolute = $02;
  rvteRangeRefRightColAbsolute = $04;
  rvteRangeRefBottomRowAbsolute = $08;

type

  /////////////////////////////////////////////////
  //
  // IvteFormulaCompilerOwner
  //
  /////////////////////////////////////////////////
{Specifies interface, which must be realized by the object which uses the TvteExcelFormulaCompiler object.}
  IvteFormulaCompilerOwner = interface
{Returns the index of the string.
Parameters:
  S - string}
    function GetStringIndex(const S: string): Integer;
{Returns the string by its index.
Parameters:
  AStringIndex - index of string.}
    function GetString(AStringIndex: Integer): string;
{Returns the index of workbook by its name.
Parameters:
  AWorkbookName - Name of workbook.}
    function GetExternalWorkbookIndex(const AWorkbookName: string): Integer;
{Returns the index of worksheet by its name and name of workbook which contains this worksheet.
Parameters:
  AWorkbookName - Name of workbook which contains worksheet.
  ASheetName - Name of worksheet.}
    function GetExternalSheetIndex(const AWorkbookName, ASheetName: string): Integer;
{Returns the name of the worksheet by its index.
Parameters:
  AWorkbookName - Name of workbook containing worksheet.
  AIndex - index of worksheet.}
    function GetExternalSheetName(const AWorkbookName: string; AIndex: Integer): string;
{Returns the index of the workbook by its index.
Parameters:
  AIndex - Index of workbook.}
    function GetExternalWorkbookName(AIndex: Integer): string;
{Returns the ID of function by its name.
Parameters:
  AFunctionName - Name of function.}
    function GetFunctionId(const AFunctionName: string): Integer;
{Returns the name of function by its ID.
Parameters:
  AId - ID if function.}
    function GetFunctionName(AId: Integer): string;
  end;

{Describes the type of formula's element.
Items:
  vteitOperator - operator, for example: +, -, <=.
  vteitFunc - function.
  vteitCellRef - reference to the cell.
  vteitRangeRef - reference to the range of the cells.
  vteitValue - value.
  vteitBracket - bracked - ( or ).
  vteitMissArg - missed argument, for example: Round(A1;). }
  TvteFormulaItemType = (vteitOperator, vteitFunc, vteitCellRef, vteitRangeRef, vteitValue, vteitBracket, vteitMissArg);
{Describes type of the formula's operand.
Syntax:
  TvteFormulaValueType = (vteNull, vteInteger, vteExtended, vteString);}
  TvteFormulaValueType = (vteNull, vteInteger, vteExtended, vteString);

  /////////////////////////////////////////////////
  //
  // rvteFormulaOperator
  //
  /////////////////////////////////////////////////
{Represents the formula's operator.
Syntax:
  rvteFormulaOperator = packed record
    Id: Integer;
  end;}
  rvteFormulaOperator = packed record
    Id: Integer;
  end;

  /////////////////////////////////////////////////
  //
  // rvteFormulaFunc
  //
  /////////////////////////////////////////////////
{Specifies the formula's function.
Syntax:
  rvteFormulaFunc = packed record
    Id: Integer;
    ParamCount: Integer;
  end;}
  rvteFormulaFunc = packed record
{ID of function.}
    Id: Integer;
{Number of the parameters passed to the function.}
    ParamCount: Integer;
  end;

  /////////////////////////////////////////////////
  //
  // rvteFormulaCellRef
  //
  /////////////////////////////////////////////////
{Specifies the formula's cell reference.
Syntax:
  rvteFormulaCellRef = packed record
    Flags: Byte;
    Workbook: Integer;
    Sheet: Integer;
    Col: Integer;
    Row: Integer;
  end;}
  rvteFormulaCellRef = packed record
{Internal flags.}
    Flags: Byte;
{Index of workbook used in the cell reference. }
    Workbook: Integer;
{Index of worksheet used in the cell reference. }
    Sheet: Integer;
{Column of the cell, zero based. }
    Col: Integer;
{Row of the cell, zero based. }
    Row: Integer;
  end;
  { Pointer to the rvteFormulaCellRef structure. }
  pvteFormulaCellRef = ^rvteFormulaCellRef;

  /////////////////////////////////////////////////
  //
  // rvteFormulaRangeRef
  //
  /////////////////////////////////////////////////
{Specifies the formula's range reference.
Syntax:
  rvteFormulaRangeRef = packed record
    Flags: Byte;
    Workbook: Integer;
    Sheet: Integer;
    LeftCol: Integer;
    TopRow: Integer;
    RightCol: Integer;
    BottomRow: Integer;
  end;}
  rvteFormulaRangeRef = packed record
    { Internal flags. }
    Flags: Byte;
    { Index of workbook used in the cell reference. }
    Workbook: Integer;
    { Index of worksheet used in the cell reference. }
    Sheet: Integer;
    { Left column of the range. }
    LeftCol: Integer;
    { Top row of the range. }
    TopRow: Integer;
    { Right column of the range. }
    RightCol: Integer;
    { Bottom row of the range. }
    BottomRow: Integer;
  end;
{Pointer to the rvteFormulaRangeRef structure.}
  pvteFormulaRangeRef = ^rvteFormulaRangeRef;

  /////////////////////////////////////////////////
  //
  // rvteFormulaValue
  //
  /////////////////////////////////////////////////
{Specifies the simple formula's value.
Syntax:
  rvteFormulaValue = packed record
    ValueType: TvteFormulaValueType;
    case TvteFormulaValueType of
      vteInteger: (vInteger: Integer);
      vteExtended: (vExtended: Extended);
      vteString: (vString: Integer);
  end;}
  rvteFormulaValue = packed record
    { Type of the value. }
    ValueType: TvteFormulaValueType;
    case TvteFormulaValueType of
      vteInteger: (vInteger: Integer);
      vteExtended: (vExtended: Extended);
      vteString: (vString: Integer);
  end;
  { Pointer to the rvteFormulaValue structure. }
  pvteFormulaValue = ^rvteFormulaValue;

  /////////////////////////////////////////////////
  //
  // rvteFormulaItem
  //
  /////////////////////////////////////////////////
{Specifies the formula's element, the compiled formula is the array of such elements.
Syntax:
  rvteFormulaItem = packed record
    ItemType: TvteFormulaItemType;
    case TvteFormulaItemType of
      vteitOperator: (Operator: rvteFormulaOperator);
      vteitFunc: (Func: rvteFormulaFunc);
      vteitCellRef: (CellRef: rvteFormulaCellRef);
      vteitRangeRef: (RangeRef: rvteFormulaRangeRef);
      vteitValue: (Value: rvteFormulaValue);
  end;}
  rvteFormulaItem = packed record
    ItemType: TvteFormulaItemType;
    case TvteFormulaItemType of
      vteitOperator: (Operator: rvteFormulaOperator);
      vteitFunc: (Func: rvteFormulaFunc);
      vteitCellRef: (CellRef: rvteFormulaCellRef);
      vteitRangeRef: (RangeRef: rvteFormulaRangeRef);
      vteitValue: (Value: rvteFormulaValue);
  end;
{Pointer to the rvteFormulaValue structure.}
  pvteFormulaItem = ^rvteFormulaItem;

  rvteFormulaItemArray = Array [0..MaxInt div SizeOf(rvteFormulaItem) div 2] of rvteFormulaItem;
{Specifies binary representation of the formula.
Binary representation of the formula is an array of the rvteFormulaItem records.}
  pvteFormula = ^rvteFormulaItemArray;

  /////////////////////////////////////////////////
  //
  // rvteOperatorInfo
  //
  /////////////////////////////////////////////////
{Describes the operator, which can be used in the formula.
Syntax:
  rvteOperatorInfo = packed record
    Name: string[2];
    Priority: Integer;
    ptg: Byte;
  end;}
  rvteOperatorInfo = packed record
    { Name of operator. }
    Name: string[2];
    { Priority of the operator. }
    Priority: Integer;
    { ID of the operator. }
    ptg: Byte;
  end;
{Pointer to the rvteOperatorInfo structure.}
  pvteOperatorInfo = ^rvteOperatorInfo;
  
  /////////////////////////////////////////////////
  //
  // rvteOperator
  //
  /////////////////////////////////////////////////
{This structure contains information about operator.
Structures of this type is being created temporarily during the compilation of formula.}
  rvteOperator = packed record
    OperatorInfo: pvteOperatorInfo;
    iftab: Word;
    ParCount: Integer;
    OperandExists: Boolean;
  end;
{Pointer to the rvteOperator structure.}
  pvteOperator = ^rvteOperator;
  
  /////////////////////////////////////////////////
  //
  // TvteCompileOpStack
  //
  /////////////////////////////////////////////////
{Internal class, implements a stack of the formula's elements used during the compilation of formula.}
  TvteCompileOpStack = class(TObject)
  private
    FList: TList;
    FCurPos: Integer;
    FLastFunction: pvteOperator;
    function GetItem(Index: Integer): pvteOperator;
    function GetCount: Integer;
  protected
    function Push: pvteOperator;
    function Pop: pvteOperator;
    function Last: pvteOperator;
    procedure Reset;
    procedure Clear;

    property Items[Index: Integer]: pvteOperator read GetItem; default;
    property Count: Integer read GetCount;
    property LastFunction: pvteOperator read FLastFunction write FLastFunction;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvteStringStack
  //
  /////////////////////////////////////////////////
{Internal class, implements a stack of the strings used in formula decompilation.}
  TvteStringStack = class(TObject)
  private
    FStrings: TStringList;
    procedure Push(const S: string);
    function Pop: string;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvteExcelFormulaCompiler
  //
  /////////////////////////////////////////////////
{Realizes formula compilation and decompilation.}
  TvteExcelFormulaCompiler = class(TObject)
  private
    FOwner: IvteFormulaCompilerOwner;
    FShowExceptions: Boolean;
    FCompileOpStack: TvteCompileOpStack;
    FDecompileStack: TvteStringStack;
    procedure SetError(const AErrorMessage : string);
  public
    constructor Create(AOwner: IvteFormulaCompilerOwner);
    destructor Destroy; override;

{Compiles a formula.
Parameters:
  S - The string containing the formula.
  AFormula - In this parameter the compiled formula is returned.
  AFormulaSize - Number of elements in the compiled formula.
  AFormulaCol - Specifies the position of the formula(column) within worksheet.
  AFormulaRow - Specifies the position of the formula(row) within worksheet.
Return value:
  Returns the true if no errors are occured.}
    function CompileFormula(const S: string; var AFormula: pvteFormula; var AFormulaSize: Integer; AFormulaCol, AFormulaRow: Integer): Boolean;
{Decompiles a formula.
Parameters:
  AFormula - The binary representation of the formula.
  AFormulaSize - Number of elements in the compiled formula.
  AFormulaCol - Specifies the position of the formula(column) within worksheet. 
  AFormulaRow - Specifies the position of the formula(row) within worksheet.
Return value:
  Returns the string representation of the formula.}
    function DecompileFormula(AFormula: pvteFormula; AFormulaSize: Integer; AFormulaCol, AFormulaRow: Integer): string;

{Clears the content of the object.}
    procedure Clear;

{Owner of the formula compiler, can not be nil.}
    property Owner: IvteFormulaCompilerOwner read FOwner;
{If true - generate exception when error occurs during the formula compilation.}
    property ShowExceptions: Boolean read FShowExceptions write FShowExceptions;
  end;

function CompileCellRef(S: string;
                        var ARow, ACol: Integer;
                        var AFlags: Byte;
                        AColAbsoluteFlag,
                        ARowAbsoluteFlag: Byte;
                        AFormulaRow, AFormulaCol: Integer): Boolean;

implementation

uses
  vgr_Functions, vgr_StringIDs, vgr_Localize;

const
  vteOperatorsInfos : array [1..vteOperatorsCount] of rvteOperatorInfo =
                        ((Name : '('; Priority : 0; ptg : 0),
                         (Name : ')'; Priority : 1; ptg : 0),
                         (Name : '>='; Priority : 2; ptg : ptgGE),
                         (Name : '<='; Priority : 2; ptg : ptgLE),
                         (Name : '<>'; Priority : 2; ptg : ptgNE),
                         (Name : '='; Priority : 2; ptg : ptgEQ),
                         (Name : '>'; Priority : 2; ptg : ptgGT),
                         (Name : '<'; Priority : 2; ptg : ptgLT),
                         (Name : '&'; Priority : 3; ptg : ptgConcat),
                         (Name : '+'; Priority : 4; ptg : ptgAdd),
                         (Name : '-'; Priority : 4; ptg : ptgSub),
                         (Name : '*'; Priority : 5; ptg : ptgMul),
                         (Name : '/'; Priority : 5; ptg : ptgDiv),
                         (Name : '^'; Priority : 6; ptg : ptgPower),
                         (Name : '%'; Priority : 7; ptg : ptgPercent),
                         (Name : '+'; Priority : 8; ptg : ptgUPlus),
                         (Name : '-'; Priority : 8; ptg : ptgUMinus),
                         (Name : ''; Priority : 9; ptg : $FF));

var
  vteFormulaNumberChars : set of char = ['0'..'9', '.'];

//{$R ..\res\vgr_FormulaCalculatorStrings.res}

function GetColNumber(const AColCaption: string): Integer;
var
  I, AMul: Integer;
begin
  AMul := 1;
  Result := 0;
  for I := Length(AColCaption) downto 1 do
  begin
    Result := Result + AMul * (Byte(AColCaption[I]) - SymbolA + 1);
    AMul := AMul * SymbolsAZ;
  end;
  Result := Result - 1;
end;

function CompileCellRef(S: string;
                        var ARow, ACol: Integer;
                        var AFlags: Byte;
                        AColAbsoluteFlag,
                        ARowAbsoluteFlag: Byte;
                        AFormulaRow, AFormulaCol: Integer): Boolean;
var
  I, J, L: integer;
begin
  Result := False;
  S := UpperCase(S);
  ARow := 0;
  ACol := 0;

  L := Length(S);
  I := L;
  while (I > 0) and (S[I] in ['0'..'9']) do Dec(I);
  if (I = 0) or (I = L) then exit;
  ARow := StrToInt(Copy(S, I + 1, L)) - 1;
  if S[I] = '$' then
  begin
    AFlags := AFlags or ARowAbsoluteFlag;
    Dec(I);
    if I = 0 then exit;
  end;

  J := I;
  while (I > 0) and (S[I] in ['A'..'Z']) do Dec(I);
  if J <= I then exit;
  if (I = 1) and (S[I] = '$') then
    AFlags := AFlags or AColAbsoluteFlag
  else
    if I <> 0 then
      exit;

  ACol := GetColNumber(Copy(S, I + 1, J - I));

  if (AFlags and ARowAbsoluteFlag) = 0 then
    ARow := ARow - AFormulaRow;
  if (AFlags and AColAbsoluteFlag) = 0 then
    ACol := ACol - AFormulaCol;

  Result := True;
end;

function GetOperatorInfo(const S: string): pvteOperatorInfo; overload;
var
  I: Integer;
begin
  I := 0;
  while (I < vteOperatorsCount) and (vteOperatorsInfos[I].Name <> S) do Inc(I);
  if I < vteOperatorsCount then
    Result := @vteOperatorsInfos[I]
  else
    Result := nil;
end;

function GetOperatorInfo(AId: Integer): pvteOperatorInfo; overload;
var
  I: Integer;
begin
  I := 0;
  while (I < vteOperatorsCount) and (vteOperatorsInfos[I].ptg <> AId) do Inc(I);
  if I < vteOperatorsCount then
    Result := @vteOperatorsInfos[I]
  else
    Result := nil;
end;

/////////////////////////////////////////////////
//
// TvteCompileOpStack
//
/////////////////////////////////////////////////
constructor TvteCompileOpStack.Create;
begin
  inherited;
  FList := TList.Create;
  FCurPos := -1;
end;

destructor TvteCompileOpStack.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TvteCompileOpStack.GetItem(Index: integer): pvteOperator;
begin
  Result := pvteOperator(FList[Index]);
end;

function TvteCompileOpStack.GetCount: Integer;
begin
  Result := FCurPos + 1;
end;

procedure TvteCompileOpStack.Clear;
var
  I: integer;
begin
  for I := 0 to FList.Count - 1 do
    FreeMem(Items[I]);
  FList.Clear;
  FCurPos := -1;
  FLastFunction := nil;
end;

procedure TvteCompileOpStack.Reset;
begin
  FCurPos := -1;
  FLastFunction := nil;
end;

function TvteCompileOpStack.Push: pvteOperator;
begin
  Inc(FCurPos);
  if FCurPos = FList.Count then
  begin
    GetMem(Result, sizeof(rvteOperator));
    FList.Add(Result);
  end
  else
    Result := Items[FCurPos];
end;

function TvteCompileOpStack.Pop: pvteOperator;
var
  I: integer;
begin
  Dec(FCurPos);
  Result := Last;
  I := Count - 1;
  while (I >= 0) and
        (Items[I].OperatorInfo.Priority <> vteFormulaFunctionPriority) do Dec(I);
  if I < 0 then
    FLastFunction := nil
  else
    FLastFunction := Items[I];
end;

function TvteCompileOpStack.Last : pvteOperator;
begin
  if FCurPos >= 0 then
    Result := Items[FCurPos]
  else
    Result := nil;
end;

/////////////////////////////////////////////////
//
// TvteStringStack
//
/////////////////////////////////////////////////
constructor TvteStringStack.Create;
begin
  inherited Create;
  FStrings := TStringList.Create;
end;

destructor TvteStringStack.Destroy;
begin
  FStrings.Free;
  inherited;
end;

procedure TvteStringStack.Push(const S: string);
begin
  FStrings.Add(S);
end;

function TvteStringStack.Pop: string;
begin
  Result := FStrings[FStrings.Count - 1];
  FStrings.Delete(FStrings.Count - 1);
end;

procedure TvteStringStack.Clear;
begin
  FStrings.Clear;
end;

/////////////////////////////////////////////////
//
// TvteExcelFormulaCompiler
//
/////////////////////////////////////////////////
constructor TvteExcelFormulaCompiler.Create(AOwner: IvteFormulaCompilerOwner);
begin
  inherited Create;
  FOwner := AOwner;
  FCompileOpStack := TvteCompileOpStack.Create;
  FDecompileStack := TvteStringStack.Create;
  FShowExceptions := True;
end;

destructor TvteExcelFormulaCompiler.Destroy;
begin
  FCompileOpStack.Free;
  FDecompileStack.Free;
  inherited;
end;

procedure TvteExcelFormulaCompiler.Clear;
begin
  FCompileOpStack.Clear;
end;

procedure TvteExcelFormulaCompiler.SetError(const AErrorMessage : string);
begin
  raise Exception.Create(AErrorMessage);
end;

function TvteExcelFormulaCompiler.CompileFormula(const S: string; var AFormula: pvteFormula; var AFormulaSize: Integer; AFormulaCol, AFormulaRow: Integer): Boolean;
var
  vd: extended;
  Last: pvteOperator;
  b1: string;
  i,j,l,vi,valCode: integer;

  procedure AddItem(AItemType: TvteFormulaItemType; AItem: Pointer; AItemSize: Integer);
  begin
    ReallocMem(AFormula, SizeOf(rvteFormulaItem) * (AFormulaSize + 1));
    FillChar(AFormula[AFormulaSize], SizeOf(rvteFormulaItem), #0);
    AFormula[AFormulaSize].ItemType := AItemType;
    if AItemSize > 0 then
      MoveMemory(@AFormula[AFormulaSize].Operator, AItem, AItemSize);
    Inc(AFormulaSize);
  end;

  procedure AddStr(const S: string);
  var
    AValue: rvteFormulaValue;
  begin
    AValue.ValueType := vteString;
    AValue.vString := Owner.GetStringIndex(S);
    AddItem(vteitValue, @AValue, SizeOf(rvteFormulaValue));
  end;

  procedure AddInt(ANumber: Integer);
  var
    AValue: rvteFormulaValue;
  begin
    AValue.ValueType := vteInteger;
    AValue.vInteger := ANumber;
    AddItem(vteitValue, @AValue, SizeOf(rvteFormulaValue));
  end;

  procedure AddNumber(ANumber: Extended);
  var
    AValue: rvteFormulaValue;
  begin
    AValue.ValueType := vteExtended;
    AValue.vExtended := ANumber;
    AddItem(vteitValue, @AValue, SizeOf(rvteFormulaValue));
  end;

  procedure AddOperator(o: pvteOperator);
  var
    AFunc: rvteFormulaFunc;
    AOperator: rvteFormulaOperator;
  begin
    if o.OperatorInfo.ptg = $FF then
    begin
      AFunc.Id := o.iftab;
      AFunc.ParamCount := o.ParCount;
      AddItem(vteitFunc, @AFunc, SizeOf(rvteFormulaFunc));
    end
    else
    begin
      AOperator.Id := o.OperatorInfo.ptg;
      AddItem(vteitOperator, @AOperator, SizeOf(rvteFormulaOperator));
    end;
  end;

  procedure AddIdentificator(const Ident: string);
  var
    I, AIdentLen, AExtRefLen: integer;
    ACellRef: rvteFormulaCellRef;
    ARangeRef: rvteFormulaRangeRef;
    ExtRef, CellRef, ExtBook, ExtSheet: string;

    procedure AddRefItem(AItemType: TvteFormulaItemType; AItem: pvteFormulaCellRef; AItemSize: Integer);
    begin
      if ExtBook <> '' then
        AItem.Workbook := Owner.GetExternalWorkbookIndex(ExtBook)
      else
        AItem.Workbook := -1;
      if ExtSheet <> '' then
        AItem.Sheet := Owner.GetExternalSheetIndex(ExtBook, ExtSheet)
      else
        AItem.Sheet := -1;
      AddItem(AItemType, AItem, AItemSize);
    end;

  begin
    // In the current version it can be only reference to a cell or range of cells of the same sheet
    AIdentLen := Length(Ident);
    I := pos('!', Ident);
    if I <> 0 then
    begin
      if Ident[1] = '''' then
      begin
        I := 2;
        while (I < AIdentLen) and (Ident[I] <> '''') do Inc(I);
        if (I >= AIdentLen) or (Ident[I + 1] <> '!') then
          SetError(Format(svteFormulaCompileErrorInvalidRangeReference, [Ident]));
        ExtRef := Copy(Ident, 2, I - 2);
        CellRef := Trim(Copy(Ident, I + 2, AIdentLen));
      end
      else
      begin
        ExtRef := Copy(Ident, 1, I - 1);
        CellRef := Trim(Copy(Ident, I + 1, AIdentLen));
      end;
      ExtRef := Trim(ExtRef);
      if ExtRef = '' then
        SetError(Format(svteFormulaCompileErrorInvalidRangeReference, [Ident]));

      // parse extref - it contains worksheet name or worksheet and workbook
      if ExtRef[1] = '[' then
      begin
        AExtRefLen := Length(ExtRef);
        I := 2;
        while (I <= AExtRefLen) and (ExtRef[I] <> ']') do Inc(I);
        if I > AExtRefLen then
          SetError(Format(svteFormulaCompileErrorInvalidRangeReference, [Ident]));
        ExtBook := Copy(ExtRef, 2, I - 2);
        ExtSheet := Copy(ExtRef, I + 1, AExtRefLen);
      end
      else
      begin
        ExtBook := '';
        ExtSheet := ExtRef;
      end;
    end
    else
    begin
      ExtBook := '';
      ExtSheet := '';
      CellRef := Ident;
    end;
      
    // parse cellref - reference to one cell or range of cells
    I := pos(':', CellRef);
    if I = 0 then
    begin
      // cell reference
      ACellRef.Flags := 0;
      if not CompileCellRef(CellRef,
                            ACellRef.Row,
                            ACellRef.Col,
                            ACellRef.Flags,
                            rvteCellRefColAbsolute,
                            rvteCellRefRowAbsolute,
                            AFormulaRow,
                            AFormulaCol) then
        SetError(Format(svteFormulaCompileErrorInvalidCellReference, [Ident]));
      AddRefItem(vteitCellRef, @ACellRef, SizeOf(rvteFormulaCellRef));
    end
    else
    begin
      // area reference
      ACellRef.Flags := 0;
      if not CompileCellRef(Copy(CellRef, 1, I - 1),
                            ARangeRef.TopRow,
                            ARangeRef.LeftCol,
                            ARangeRef.Flags,
                            rvteRangeRefLeftColAbsolute,
                            rvteRangeRefTopRowAbsolute,
                            AFormulaRow,
                            AFormulaCol) then
        SetError(Format(svteFormulaCompileErrorInvalidRangeReference, [Ident]))
      else
        if not CompileCellRef(Copy(CellRef, I + 1, AIdentLen),
                              ARangeRef.BottomRow,
                              ARangeRef.RightCol,
                              ARangeRef.Flags,
                              rvteRangeRefRightColAbsolute,
                              rvteRangeRefBottomRowAbsolute,
                              AFormulaRow,
                              AFormulaCol) then
          SetError(Format(svteFormulaCompileErrorInvalidRangeReference, [Ident]));
      AddRefItem(vteitRangeRef, @ARangeRef, SizeOf(rvteFormulaRangeRef));
    end;
  end;

  function GetOperatorInfo(const S: string): pvteOperatorInfo;
  begin
    Result := vgr_ExcelFormula.GetOperatorInfo(S);
    if Result = nil then
      SetError(Format(svteFormulaCompileErrorUnknownOperator, [S]));
  end;

  function GetFunction_iftab(const S: string): integer;
  begin
    Result := Owner.GetFunctionId(S);
    if Result = -1 then
      SetError(Format(svteFormulaCompileErrorUnknownFunction,[S]));
  end;

  function ProcessOperator(AInfo: pvteOperatorInfo): pvteOperator;
  var
    ALast: pvteOperator;
  begin
    Result := nil;

    ALast := FCompileOpStack.Last;
    if (ALast <> nil) and
       (AInfo.Priority <> 0) and
       (AInfo.Priority <= ALast.OperatorInfo.Priority) then
      while (ALast <> nil) and (ALast.OperatorInfo.Priority >= AInfo.Priority) do
      begin
        AddOperator(ALast);
        ALast := FCompileOpStack.Pop;
      end;

    if AInfo.Priority <> vteFormulaEndBracketPriority then
    begin
      Result := FCompileOpStack.Push;
      with Result^ do
      begin
        OperatorInfo := AInfo;
        ParCount := 0;
        OperandExists := False;
      end
    end
    else
    begin
      if ALast = nil then
        SetError(svteFormulaCompileErrorInvalidBrackets)
      else
      begin
        ALast := FCompileOpStack.Pop;
        if ALast <> nil then
        begin
          if ALast.OperatorInfo.Priority <> vteFormulaFunctionPriority then
            AddItem(vteitBracket, nil, 0)
          else
          begin
            if ALast.OperandExists then
              Inc(ALast.ParCount)
            else
              if ALast.ParCount > 0 then
              begin
                AddItem(vteitMissArg, nil, 0);
                Inc(ALast.ParCount)
              end;
          end;
        end
        else
          AddItem(vteitBracket, nil, 0)
      end;
    end;
  end;

begin
  AFormula := nil;
  AFormulaSize := 0;
  try
    FCompileOpStack.Reset;

    L := Length(s);
    I := 1;
    while I <= L do
    begin
      if S[I] in vteFormulaStartIdentChars then
      begin
        // identificator
        if FCompileOpStack.LastFunction <> nil then
          FCompileOpStack.LastFunction.OperandExists := True;

        J := I;
        while (I <= L) and (S[I] in vteFormulaIdentChars) do Inc(I);
        while (I <= L) and (S[I] <= #32) do Inc(I);
        b1 := Trim(Copy(S, J, I - J));
        if (I <= L) and (S[I] = vteFormulaStartBracketChar) then
        begin
          // this is function, find function iftab
          FCompileOpStack.LastFunction := ProcessOperator(@vteOperatorsInfos[vteFormulaFunctionOperatorIndex]);
          FCompileOpStack.LastFunction.iftab := GetFunction_iftab(b1)
        end
        else
          AddIdentificator(b1);
      end
      else
        if (S[I] = vteFormulaFuncParamsDelim) then
        begin
          // this is a function parameters delimeter
          if FCompileOpStack.LastFunction = nil then
            SetError(svteFormulaCompileErrorParameterWithoutFunction);
          if not FCompileOpStack.LastFunction.OperandExists then
            AddItem(vteitMissArg, nil, 0);

          // We should process all operators up to the first opening bracket
          Last := FCompileOpStack.Last;
          while Last.OperatorInfo.Priority <> vteFormulaStartBracketPriority do
          begin
            AddOperator(Last);
            Last := FCompileOpStack.Pop;
          end;
          with FCompileOpStack.LastFunction^ do
          begin
            OperandExists := false;
            Inc(ParCount);
          end;
          Inc(I);
        end
        else
          if S[I] = vteFormulaPercentOperator then
          begin
            ProcessOperator(@vteOperatorsInfos[vteFormulaPercentOperatorIndex]);
            Inc(I);
          end
          else
            if S[I] in vteFormulaOperatorChars then
            begin
              // operator
              J := I;
              while (I <= L) and (S[I] in vteFormulaOperatorChars) do Inc(I);
              b1 := Copy(S, J, I - J);
              vi := Length(b1);
              if (vi > 1) and (b1[vi] in vteFormulaUnaryOperators) then
              begin
                ProcessOperator(GetOperatorInfo(Copy(b1, 1, vi - 1)));
                if b1[vi] = vteFormulaUnaryPlusOperator then
                  ProcessOperator(@vteOperatorsInfos[vteFormulaUnaryPlusOperatorIndex])
                else
                  ProcessOperator(@vteOperatorsInfos[vteFormulaUnaryMinusOperatorIndex]);
              end
              else
                if (vi = 1) and (b1[1] in vteFormulaUnaryOperators) then
                begin
                  // Probably it is a unary operator
                  Dec(J);
                  while (J > 1) and (S[J] <= #32) do Dec(J);
                  if (J < 1) or (S[J] in [vteFormulaStartBracketChar,vteFormulaFuncParamsDelim]) then
                  begin
                    if b1[vi] = vteFormulaUnaryPlusOperator then
                      ProcessOperator(@vteOperatorsInfos[vteFormulaUnaryPlusOperatorIndex])
                    else
                      ProcessOperator(@vteOperatorsInfos[vteFormulaUnaryMinusOperatorIndex]);
                  end
                  else
                    ProcessOperator(GetOperatorInfo(b1));
                end
                else
                  ProcessOperator(GetOperatorInfo(b1));
            end
            else
              if S[I] = vteFormulaStartBracketChar then
              begin
                ProcessOperator(@vteOperatorsInfos[vteFormulaStartBracketOperatorIndex]);
                Inc(I);
              end
              else
                if S[I] = vteFormulaEndBracketChar then
                begin
                  ProcessOperator(@vteOperatorsInfos[vteFormulaEndBracketOperatorIndex]);
                  Inc(I);
                end
                else
                  if S[I] = vteFormulaStringChar then
                  begin
                    // text string
                    if FCompileOpStack.LastFunction <> nil then
                      FCompileOpStack.LastFunction.OperandExists := True;

                    Inc(I);
                    J := I;
                    while (I <= L) and (S[I] <> vteFormulaStringChar) do Inc(I);
                    if I > L then
                      SetError(svteFormulaCompileErrorInvalidString);
                   
                   AddStr(Copy(S, J, I - J));
                   Inc(i);
                 end
                 else
                   if S[I] in vteFormulaNumberChars then
                   begin
                     // number - integer or double
                     if FCompileOpStack.LastFunction <> nil then
                       FCompileOpStack.LastFunction.OperandExists := True;

                     J := I;
                     while (I <= L) and (S[I] in vteFormulaNumberChars) do Inc(I);
                     b1 := Copy(S, J, I - J);

                     val(b1, vi, valCode);
                     if valCode = 0 then
                       AddInt(vi)
                     else
                       if TextToFloat(PChar(b1), vd, fvExtended) then
                         AddNumber(vd)
                       else
                         SetError(Format(svteFormulaCompileErrorInvalidNumber, [b1]));
                   end
                   else
                     if S[I] <= #32 then
                       Inc(I)
                     else
                       SetError(Format(svteFormulaCompileErrorInvalidSymbol, [S[I]]));
    end;

    Last := FCompileOpStack.Last;
    while FCompileOpStack.Last <> nil do
    begin
      if Last.OperatorInfo.Priority = vteFormulaStartBracketPriority then
        SetError(svteFormulaCompileErrorInvalidBrackets);
      AddOperator(Last);
      Last := FCompileOpStack.Pop;
    end;
    Result := True;
    
  except
    if AFormula <> nil then
    begin
      FreeMem(AFormula);
      AFormula := nil;
    end;
    AFormulaSize := 0;
    Result := False;
    if ShowExceptions then
      raise;
  end;
end;

function TvteExcelFormulaCompiler.DecompileFormula(AFormula: pvteFormula; AFormulaSize: Integer; AFormulaCol, AFormulaRow: Integer): string;
var
  I, J: Integer;
  S, AName: string;
  AItem: pvteFormulaItem;
  AOperatorInfo: pvteOperatorInfo;

  function GetCellString(ACol, ARow: Integer; AColRel, ARowRel: Boolean): string;
  const
    ARel: Array [Boolean] of string = ('$', '');
  begin
    if AColRel then
      ACol := AFormulaCol + ACol;
    if ARowRel then
      ARow := AFormulaRow + ARow;
    if (ACol < 0) or (ARow < 0) then
      raise Exception.Create(vgrLoadStr(svgrid_vgr_ExcelFormula_InvalidCellReference))
    else
      Result := ARel[AColRel] + GetWorksheetColCaption(ACol) + ARel[ARowRel] + GetWorksheetRowCaption(ARow);
  end;

  function GetWorkbookAndSheetString(AWorkbook, ASheet: Integer): string;
  var
    ABookName: string;
  begin
    Result := '';
    if ASheet <> -1 then
    begin
      if AWorkbook = -1 then
        ABookName := ''
      else
      begin
        ABookName := Owner.GetExternalWorkbookName(AWorkbook);
        Result := '[' + ABookName + ']';
      end;
      Result := '''' + Result + Owner.GetExternalSheetName(ABookName, ASheet) + '''!';
    end;
  end;

begin
  try
    try
      for I := 0 to AFormulaSize - 1 do
      begin
        AItem := @AFormula[I];
        case AItem.ItemType of
          vteitFunc:
            begin
              AName := Owner.GetFunctionName(AItem.Func.Id);
              S := vteFormulaEndBracketChar;
              if AItem.Func.ParamCount > 0 then
              begin
                for J := 0 to AItem.Func.ParamCount - 1 do
                  S := vteFormulaFuncParamsDelim + FDecompileStack.Pop + S;
                Delete(S, 1, 1);
              end;
              S := AName + vteFormulaStartBracketChar + S;
              FDecompileStack.Push(S);
            end;
          vteitOperator:
            begin
              AOperatorInfo := GetOperatorInfo(AItem.Operator.Id);
              if AItem.Operator.Id in [vtefoUplus, vtefoUminus] then
                FDecompileStack.Push(AOperatorInfo.Name + FDecompileStack.Pop)
              else
              begin
                S := AOperatorInfo.Name + FDecompileStack.Pop;
                FDecompileStack.Push(FDecompileStack.Pop + S);
              end;
            end;
          vteitCellRef:
            begin
              with AItem.CellRef do
                FDecompileStack.Push(GetWorkbookAndSheetString(Workbook, Sheet) +
                                     GetCellString(Col,
                                                   Row,
                                                   (Flags and rvteCellRefColAbsolute) = 0,
                                                   (Flags and rvteCellRefRowAbsolute) = 0));
            end;
          vteitRangeRef:
            begin
              with AItem.RangeRef do
                FDecompileStack.Push(GetWorkbookAndSheetString(Workbook, Sheet) +
                                     GetCellString(LeftCol,
                                                   TopRow,
                                                   (Flags and rvteRangeRefLeftColAbsolute) = 0,
                                                   (Flags and rvteRangeRefTopRowAbsolute) = 0) + ':' +
                                     GetCellString(RightCol,
                                                   BottomRow,
                                                   (Flags and rvteRangeRefRightColAbsolute) = 0,
                                                   (Flags and rvteRangeRefBottomRowAbsolute) = 0));
            end;
          vteitValue:
            begin
              with AItem.Value do
                case ValueType of
                  vteNull: FDecompileStack.Push('');
                  vteInteger: FDecompileStack.Push(IntToStr(vInteger));
                  vteExtended: FDecompileStack.Push(FloatToStr(vExtended));
                  vteString: FDecompileStack.Push(vteFormulaStringChar + Owner.GetString(vString) + vteFormulaStringChar);
                end;
            end;
          vteitBracket:
            begin
              FDecompileStack.Push(vteFormulaStartBracketChar + FDecompileStack.Pop + vteFormulaEndBracketChar);
            end;
          vteitMissArg:
            begin
              FDecompileStack.Push('');
            end;
        end;
      end;
      Result := FDecompileStack.Pop;
    except
      on E: Exception do
      begin
        if ShowExceptions then
          raise
        else
          Result := E.Message;
      end;
    end;
  finally
    FDecompileStack.Clear;
  end;
end;

initialization

  Include(vteFormulaNumberChars, DecimalSeparator);

end.

