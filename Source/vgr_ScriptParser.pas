{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes, which provide functionality to parse the script and divide him on the separate procedures.
Each script engine are represented by the separate script parser class,
which are derived from TvgrScriptParser. As the result of the parsing the instance
of the TvgrScriptProgramInfo class are created.
See also:
  TvgrScriptParser, TvgrVBScriptParser, TvgrJScriptParser, TvgrScriptProgramInfo}
unit vgr_ScriptParser;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  SysUtils, Classes,

  vgr_Event, vgr_Functions;

type

  /////////////////////////////////////////////////
  //
  // TvgrScriptProcedureParameter
  //
  /////////////////////////////////////////////////
{TvgrScriptProcedureParameter represents the separate parameter of the scripting procedure.}
  TvgrScriptProcedureParameter = class(TObject)
  private
    FParamName: string;
    FIsVar: Boolean;
  public
{Creates a instance of the TvgrScriptProcedureParameter.
Parameters:
  AParamName - name of the parameter.
  AIsVar - boolean value that specifies is this parameter passed by reference of by value.
If true then parameter is passed by reference (var parameter).}
    constructor Create(const AParamName: string; AIsVar: Boolean);
{Returns name of the parameter.}
    property ParamName: string read FParamName;
{Returns boolean value that specifies is this parameter passed by reference of by value.
If true then parameter is passed by reference (var parameter).}
    property IsVar: Boolean read FIsVar;
  end;

{Specifies type of the script method.
Items:
  vgrsptProcedure - the method is a procedure.
  vgrsptFunction - the method is a function.}
  TvgrScriptProcedureType = (vgrsptProcedure, vgrsptFunction);
  /////////////////////////////////////////////////
  //
  // TvgrScriptProcedureInfo
  //
  /////////////////////////////////////////////////
{TvgrScriptProcedureInfo class describes the separate procedure of the script.}
  TvgrScriptProcedureInfo = class(TObject)
  private
    FName: string;
    FStartPos: Integer;
    FEndPos: Integer;
    FParameters: TList;
    FProcedureType: TvgrScriptProcedureType;
    function GetParameter(Index: Integer): TvgrScriptProcedureParameter;
    function GetParameterCount: Integer;
  public
{Creates a instance of the TvgrScriptProcedureInfo class.
Parameters:
  AName - name of the procedure.
  AProcedureType - type of the procedure (can be a function or procedure).
  AStartPos - specifies the position of the starting char within a script from wich procedure begins.
  AEndPos - specified the position of the last char within a script with which procedure ends.}
    constructor Create(const AName: string; AProcedureType: TvgrScriptProcedureType; AStartPos, AEndPos: Integer);
{Frees a instance of the TvgrScriptProcedureInfo class.}
    destructor Destroy; override;
{Returns name of the procedure.}
    property Name: string read FName;
{Returns the position of the starting char within a script from wich procedure begins.}
    property StartPos: Integer read FStartPos;
{Returns the position of the last char within a script with which procedure ends.}
    property EndPos: Integer read FEndPos;
{Returns the type of the procedure (can be a function or procedure).
See also:
  TvgrScriptProcedureType}
    property ProcedureType: TvgrScriptProcedureType read FProcedureType;
{Enumerates parameters of the procedure.
Parameters:
  Index - index of the parameter, starts from 0.
See also:
  ParameterCount}
    property Parameters[Index: Integer]: TvgrScriptProcedureParameter read GetParameter;
{Returns amount of the procedure parameters.
See also:
  Parameters}
    property ParameterCount: Integer read GetParameterCount;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredScriptEvent
  //
  /////////////////////////////////////////////////
{TvgrRegisteredScriptEvent class represents the event template based on which the script parser generates the empty procedure for this event.
See also:
  TvgrScriptParser}
  TvgrRegisteredScriptEvent = class(TObject)
  private
    FBaseClass: TClass;
    FEventsPropertyName: string;
    FEventName: string;
    FProcedureType: TvgrScriptProcedureType;
    FList: TList;
    function GetParameter(Index: Integer): TvgrScriptProcedureParameter;
    function GetParameterCount: Integer;
  public
{Creates a instance of the TvgrRegisteredScriptEvent class.}
    constructor Create; overload;
{Creates a instance of the TvgrRegisteredScriptEvent class.
Parameters:
  ABaseClass - class which contains a event.
  AEventsPropertyName - name of the property.
  AEventName - name of the event.
  AProcedureType - type of the event procedure.
  AParameters - describes parameters of the event.}
    constructor Create(ABaseClass: TClass;
                       const AEventsPropertyName: string;
                       const AEventName: string;
                       AProcedureType: TvgrScriptProcedureType;
                       const AParameters: string); overload;
{Frees instance of the TvgrRegisteredScriptEvent class.}
    destructor Destroy; override;

{Returns the class, which contains a event.}
    property BaseClass: TClass read FBaseClass;
{Returns the name of the property of the TvgrScriptEvents type.}
    property EventsPropertyName: string read FEventsPropertyName;
{Returns name of the property.}
    property EventName: string read FEventName;

{Returns type of the event procedure.}
    property ProcedureType: TvgrScriptProcedureType read FProcedureType;
{Enumerates parameters of the event.
Parameters:
  Index - index of the parameter, starts from 0.
See also:
  ParameterCount}
    property Parameters[Index: Integer]: TvgrScriptProcedureParameter read GetParameter;
{Returns amount of the event parameters.
See also:
  Parameters}
    property ParameterCount: Integer read GetParameterCount;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptProgramInfo
  //
  /////////////////////////////////////////////////
{TvgrScriptProgramInfo class describes the script.
Instance of this class created as result of working of the TvgrScriptParser.
See also:
  TvgrScriptParser}
  TvgrScriptProgramInfo = class(TObject)
  private
    FList: TList;
    function GetProcedure(Index: Integer): TvgrScriptProcedureInfo;
    function GetProcedureCount: Integer;
  public
{Creates a instance of the TvgrScriptProgramInfo class.}
    constructor Create;
{Frees a instance of the TvgrScriptProgramInfo class.}
    destructor Destroy; override;

{Adds the description of the procedure.
Parameters:
  AName - name of the procedure.
  AProcedureType - type of the procedure (can be procedure or function).
  AStartPos - specifies the position of the starting char within a script from wich procedure begins.
  AEndPos - specified the position of the last char within a script with which procedure ends.
Return value:
  Returns the created TvgrScriptProcedureInfo object.}
    function AddProcedure(const AName: string; AProcedureType: TvgrScriptProcedureType; AStartPos, AEndPos: Integer): TvgrScriptProcedureInfo;
{Returns the index of the procedure info by its name.
Parameters:
  AProcedureName - name of the procedure to find.
Return value:
  Returns the index of the TvgrScriptProcedureInfo object within internal list.}
    function IndexOfProcedure(const AProcedureName: string): Integer;

{Clear contents.}
    procedure Clear;

{Enumerates procedures of the script.
Parameters:
  Index - index of the procedure.
See also:
  ProcedureCount}
    property Procedures[Index: Integer]: TvgrScriptProcedureInfo read GetProcedure;
{Returns amount of the procedures.
See also:
  Procedures}
    property ProcedureCount: Integer read GetProcedureCount;
  end;

{TvgrScriptParserClass is a class type of a TvgrScriptParser descendant.
Syntax:
  TvgrScriptParserClass = class of TvgrScriptParser;}
  TvgrScriptParserClass = class of TvgrScriptParser;
  /////////////////////////////////////////////////
  //
  // TvgrScriptParser
  //
  /////////////////////////////////////////////////
{TvgrScriptParser class provides functionality to parse the script and divide him on the separate procedures.
Each script engine are represented by the separate script parser class,
which are derived from TvgrScriptParser. As the result of the parsing the instance
of the TvgrScriptProgramInfo class are created.
See also:
  TvgrVBScriptParser, TvgrJScriptParser}
  TvgrScriptParser = class(TObject)
  private
    FIgnoreCase: Boolean;
    FKeywordDelimiters: TSysCharSet;
    FStringDelimiters: TSysCharSet;
    FStringSpecialChar: Char;
    FKW_ProcedureStart: string;
    FKW_FunctionStart: string;
    FKW_StartBlock: string;
    FKW_EndBlock: string;
    FKW_VarParameterPrefix: string;
    FKW_ValParameterPrefix: string;
  protected
    function IsEQKW(const AKeyWord: string; const ATestedKeyWord: string): Boolean;
    function FindNextKeyword(const AText: string; var APos: Integer): string;

    property IgnoreCase: Boolean read FIgnoreCase;
    property KeywordDelimiters: TSysCharSet read FKeywordDelimiters;
    property StringDelimiters: TSysCharSet read FStringDelimiters;
    property StringSpecialChar: Char read FStringSpecialChar;
    property KW_ProcedureStart: string read FKW_ProcedureStart;
    property KW_FunctionStart: string read FKW_FunctionStart;
    property KW_StartBlock: string read FKW_StartBlock;
    property KW_EndBlock: string read FKW_EndBlock;
    property KW_VarParameterPrefix: string read FKW_VarParameterPrefix;
    property KW_ValParameterPrefix: string read FKW_ValParameterPrefix;
  public
{Creates a instance of the TvgrScriptParser class.}
    constructor Create; virtual;

{Generates the text of the event procedure.
Parameters:
  AEventInfo - describes the event.
  AProcedureName - name of the event procedure, if not specified the default name are generated.
  AText - receives the text of the event procedure.}
    procedure GenerateEventProcedureText(AEventInfo: TvgrRegisteredScriptEvent; const AProcedureName: string; var AText: string); virtual; abstract;
{Parse the script and fills TvgrScriptProgramInfo object.
Parameters:
  AProgram - the text of the script.
  AProgramInfo - the TvgrScriptProgramInfo object that receive information about the script.}
    procedure ParseProgram(const AProgram: string; AProgramInfo: TvgrScriptProgramInfo); virtual; abstract;
//    function FindProcedure(const AProcedureName: string; AProgram: TStrings): Boolean;
//    function FindProcedure(const AProcedureName: string; AProgram: TStrings; var AStartPos, AEndPos: Integer); Boolean;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrVBScriptParser
  //
  /////////////////////////////////////////////////
{TvgrVBScriptParser class intends to parsing the Visual Basic scripts.
See also:
  TvgrScriptParser, TvgrJScriptParser}
  TvgrVBScriptParser = class(TvgrScriptParser)
  public
{Creates a instance of the TvgrVBScriptParser class.}
    constructor Create; override;
{Generates the text of the event procedure on the VBScript language.
Parameters:
  AEventInfo - describes the event.
  AProcedureName - name of the event procedure, if not specified the default name are generated.
  AText - receives the text of the event procedure.}
    procedure GenerateEventProcedureText(AEventInfo: TvgrRegisteredScriptEvent; const AProcedureName: string; var AText: string); override;
{Parse the Visual Basic script and fills TvgrScriptProgramInfo object.
Parameters:
  AProgram - the text of the script.
  AProgramInfo - the TvgrScriptProgramInfo object that receive information about the script.}
    procedure ParseProgram(const AProgram: string; AProgramInfo: TvgrScriptProgramInfo); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrJScriptParser
  //
  /////////////////////////////////////////////////
{TvgrJScriptParser class intends to parsing the Java scripts.
See also:
  TvgrScriptParser, TvgrVBScriptParser}
  TvgrJScriptParser = class(TvgrScriptParser)
  public
{Creates a instance of the TvgrJScriptParser class.}
    constructor Create; override;
{Generates the text of the event procedure on the JavaScript language.
Parameters:
  AEventInfo - describes the event.
  AProcedureName - name of the event procedure, if not specified the default name are generated.
  AText - receives the text of the event procedure.}
    procedure GenerateEventProcedureText(AEventInfo: TvgrRegisteredScriptEvent; const AProcedureName: string; var AText: string); override;
{Parse the Java script and fills TvgrScriptProgramInfo object.
Parameters:
  AProgram - the text of the script.
  AProgramInfo - the TvgrScriptProgramInfo object that receive information about the script.}
    procedure ParseProgram(const AProgram: string; AProgramInfo: TvgrScriptProgramInfo); override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredScriptParser
  //
  /////////////////////////////////////////////////
{TvgrRegisteredScriptParser class represents the registered script parser.}
  TvgrRegisteredScriptParser = class(TObject)
  private
    FLanguage: string;
    FScriptParserClass: TvgrScriptParserClass;
  public
{Creates a instance of the TvgrRegisteredScriptParser class.
Parameters:
  ALanguage - the script language for which the parser are registered (VBScript, JScript and so on).
  AScriptParserClass - class of the script parser, must be derived from TvgrScriptParser.}
    constructor Create(const ALanguage: string; AScriptParserClass: TvgrScriptParserClass);
{Returns the script language for which the parser are registered (VBScript, JScript and so on).}
    property Language: string read FLanguage;
{Returns the class of the script parser, must be derived from TvgrScriptParser.}
    property ScriptParserClass: TvgrScriptParserClass read FScriptParserClass;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredScriptParsers
  //
  /////////////////////////////////////////////////
{Instance of the TvgrRegisteredScriptParsers class contains list of the registered script parsers and events.
Instance of this class is created at the start of the programm.
See also:
  RegisteredScriptParsers}
  TvgrRegisteredScriptParsers = class(TObject)
  private
    FList: TList;
    FEvents: TList;
    function GetItem(Index: Integer): TvgrRegisteredScriptParser;
    function GetCount: Integer;
    function GetEvent(Index: Integer): TvgrRegisteredScriptEvent;
    function GetEventCount: Integer;
  public
{Creates a instance of the TvgrRegisteredScriptParsers class.}
    constructor Create;
{Frees a instance of the TvgrRegisteredScriptParsers class.}
    destructor Destroy; override;

{Clears all lists.}
    procedure Clear;

{Registers the new script parser.
Parameters:
  ALanguage - the script language for which the parser are registered (VBScript, JScript and so on).
  AScriptParserClass - class of the script parser, must be derived from TvgrScriptParser.}
    procedure RegisterScriptParser(const ALanguage: string; AScriptParserClass: TvgrScriptParserClass);
{Returns index of the TvgrRegisteredScriptParser object by the Language.
Parameters:
  ALanguage - the script language for which the parser are registered (VBScript, JScript and so on).
Return value:
  Index of the found TvgrRegisteredScriptParser object.}
    function IndexByLanguage(const ALanguage: string): Integer;

{Registers the event.
Parameters:
  ABaseClass - class which contains a event.
  AEventsPropertyName - name of the property.
  AEventName - name of the event.
  AProcedureType - type of the event procedure.
  AParameters - describes parameters of the event.}
    procedure RegisterScriptEvent(ABaseClass: TClass;
                                  const AEventsPropertyName: string;
                                  const AEventName: string;
                                  AProcedureType: TvgrScriptProcedureType;
                                  const AParameters: string);
{Returns index of the TvgrRegisteredScriptEvent object.
Parameters:
  ABaseClass - class which contains a event.
  AEventsPropertyName - name of the property.
  AEventName - name of the event.
Return value:
  Index of the found TvgrRegisteredScriptEvent object.}
    function IndexOfEvent(ABaseClass: TClass;
                          const AEventsPropertyName: string;
                          const AEventName: string): Integer;

{Returns amount of the registered TvgrRegisteredScriptParser objects.}
    property Count: Integer read GetCount;
{Enumerates the registered TvgrRegisteredScriptParser objects.
Parameters:
  Index - index of the TvgrRegisteredScriptParser object, started with 0.}
    property Items[Index: Integer]: TvgrRegisteredScriptParser read GetItem; default;

{Returns amount of the registered TvgrRegisteredScriptEvent objects.}
    property EventCount: Integer read GetEventCount;
{Enumerates the registered TvgrRegisteredScriptEvent objects.
Parameters:
  Index - index of the TvgrRegisteredScriptEvent object, started with 0.}
    property Events[Index: Integer]: TvgrRegisteredScriptEvent read GetEvent;
  end;

{Returns the descendant of the TvgrScriptParser class, which are registered for the ALanguage.
Parameters:
  ALanguage - specifies the script language (VBScript or JScript).
Return value:
  Returns the descendant of the TvgrScriptParser class, which are registered for the ALanguage.}
  function GetScriptParserClass(const ALanguage: string): TvgrScriptParserClass;
{Registers the new script parser the global variable RegisteredScriptParsers is used for registering.
Parameters:
  ALanguage - the script language for which the parser are registered (VBScript, JScript and so on).
  AScriptParserClass - class of the script parser, must be derived from TvgrScriptParser.}
  procedure RegisterScriptParser(const ALanguage: string; AScriptParserClass: TvgrScriptParserClass);
{Registers the event, the global variable RegisteredScriptParsers is used for registering.
Parameters:
  ABaseClass - class which contains a event.
  AEventsPropertyName - name of the property.
  AEventName - name of the event.
  AProcedureType - type of the event procedure.
  AParameters - describes parameters of the event.}
  procedure RegisterScriptEvent(ABaseClass: TClass;
                                const AEventsPropertyName: string;
                                const AEventName: string;
                                AProcedureType: TvgrScriptProcedureType;
                                const AParameters: string);


var
{This variable contains instance of the TvgrRegisteredScriptParsers object,
which is created at the start of the programm.
See also:
  TvgrRegisteredScriptParsers}
  RegisteredScriptParsers: TvgrRegisteredScriptParsers;
  
implementation

uses
  Math;

function GetScriptParserClass(const ALanguage: string): TvgrScriptParserClass;
var
  I: Integer;
begin
  I := RegisteredScriptParsers.IndexByLanguage(ALanguage);
  if I <> -1 then
    Result := RegisteredScriptParsers[I].ScriptParserClass
  else
    Result := nil;
end;

procedure RegisterScriptParser(const ALanguage: string; AScriptParserClass: TvgrScriptParserClass);
begin
  RegisteredScriptParsers.RegisterScriptParser(ALanguage, AScriptParserClass);
end;

procedure RegisterScriptEvent(ABaseClass: TClass;
                              const AEventsPropertyName: string;
                              const AEventName: string;
                              AProcedureType: TvgrScriptProcedureType;
                              const AParameters: string);
begin
  RegisteredScriptParsers.RegisterScriptEvent(ABaseClass,
                                              AEventsPropertyName,
                                              AEventName,
                                              AProcedureType,
                                              AParameters);
end;


/////////////////////////////////////////////////
//
// TvgrRegisteredScriptEvent
//
/////////////////////////////////////////////////
constructor TvgrRegisteredScriptEvent.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

constructor TvgrRegisteredScriptEvent.Create(ABaseClass: TClass;
                                             const AEventsPropertyName: string;
                                             const AEventName: string;
                                             AProcedureType: TvgrScriptProcedureType;
                                             const AParameters: string);
var
  I: Integer;
  AIsVar: Boolean;
  AParamName: string;
begin
  Create;
  FBaseClass := ABaseClass;
  FEventsPropertyName := AEventsPropertyName;
  FEventName := AEventName;
  FProcedureType := AProcedureType;
  I := 1;
  while I <= Length(AParameters) do
  begin
    AParamName := Trim(ExtractSubStr(AParameters, I, [',']));
    AIsVar := AnsiCompareText(Trim(ExtractSubStr(AParameters, I, [','])), 'var') = 0;

    FList.Add(TvgrScriptProcedureParameter.Create(AParamName, AIsVar));
  end;
end;

destructor TvgrRegisteredScriptEvent.Destroy;
begin
  FreeList(FList);
  inherited;
end;

function TvgrRegisteredScriptEvent.GetParameter(Index: Integer): TvgrScriptProcedureParameter;
begin
  Result := TvgrScriptProcedureParameter(FList[Index]);
end;

function TvgrRegisteredScriptEvent.GetParameterCount: Integer;
begin
  Result := FList.Count;
end;

/////////////////////////////////////////////////
//
// TvgrRegisteredScriptParser
//
/////////////////////////////////////////////////
constructor TvgrRegisteredScriptParser.Create(const ALanguage: string; AScriptParserClass: TvgrScriptParserClass);
begin
  inherited Create;
  FLanguage := ALanguage;
  FScriptParserClass := AScriptParserClass;
end;

/////////////////////////////////////////////////
//
// TvgrRegisteredScriptParsers
//
/////////////////////////////////////////////////
constructor TvgrRegisteredScriptParsers.Create;
begin
  inherited Create;
  FList := TList.Create;
  FEvents := TList.Create;
end;

destructor TvgrRegisteredScriptParsers.Destroy;
begin
  Clear;
  FList.Free;
  FEvents.Free;
  inherited;
end;

function TvgrRegisteredScriptParsers.GetItem(Index: Integer): TvgrRegisteredScriptParser;
begin
  Result := TvgrRegisteredScriptParser(FList[Index]);
end;

function TvgrRegisteredScriptParsers.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TvgrRegisteredScriptParsers.GetEvent(Index: Integer): TvgrRegisteredScriptEvent;
begin
  Result := TvgrRegisteredScriptEvent(FEvents[Index]);
end;

function TvgrRegisteredScriptParsers.GetEventCount: Integer;
begin
  Result := FEvents.Count;
end;

procedure TvgrRegisteredScriptParsers.Clear;
begin
  FreeListItems(FList);
  FList.Clear;
  FreeListItems(FEvents);
  FEvents.Clear;
end;

function TvgrRegisteredScriptParsers.IndexByLanguage(const ALanguage: string): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].Language <> ALanguage) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

procedure TvgrRegisteredScriptParsers.RegisterScriptParser(const ALanguage: string; AScriptParserClass: TvgrScriptParserClass);
begin
  if IndexByLanguage(ALanguage) = -1 then
    FList.Add(TvgrRegisteredScriptParser.Create(ALanguage, AScriptParserClass));
end;


procedure TvgrRegisteredScriptParsers.RegisterScriptEvent(ABaseClass: TClass;
                                                          const AEventsPropertyName: string;
                                                          const AEventName: string;
                                                          AProcedureType: TvgrScriptProcedureType;
                                                          const AParameters: string);
begin
  if IndexOfEvent(ABaseClass, AEventsPropertyName, AEventName) = -1 then
    FEvents.Add(TvgrRegisteredScriptEvent.Create(ABaseClass,
                                                 AEventsPropertyName,
                                                 AEventName,
                                                 AProcedureType,
                                                 AParameters));
end;

function TvgrRegisteredScriptParsers.IndexOfEvent(ABaseClass: TClass;
                                                  const AEventsPropertyName: string;
                                                  const AEventName: string): Integer;
begin
  Result := 0;
  while (Result < EventCount) and
        not ((Events[Result].BaseClass = ABaseClass) and
             (AnsiCompareText(Events[Result].EventsPropertyName, AEventsPropertyName) = 0) and
             (AnsiCompareText(Events[Result].EventName, AEventName) = 0)) do Inc(Result);
  if Result >= EventCount then
    Result := -1;
end;

/////////////////////////////////////////////////
//
// TvgrScriptProcedureParameter
//
/////////////////////////////////////////////////
constructor TvgrScriptProcedureParameter.Create(const AParamName: string; AIsVar: Boolean);
begin
  inherited Create;
  FParamName := AParamName;
  FIsVar := AIsVar;
end;

/////////////////////////////////////////////////
//
// TvgrScriptProcedureInfo
//
/////////////////////////////////////////////////
constructor TvgrScriptProcedureInfo.Create(const AName: string; AProcedureType: TvgrScriptProcedureType; AStartPos, AEndPos: Integer);
begin
  inherited Create;
  FParameters := TList.Create;
  FName := AName;
  FStartPos := AStartPos;
  FEndPos := AEndPos;
  FProcedureType := AProcedureType;
end;

destructor TvgrScriptProcedureInfo.Destroy;
var
  I: Integer;
begin
  for I := 0 to ParameterCount - 1 do
    Parameters[I].Free;
  FParameters.Free;
  inherited;
end;

function TvgrScriptProcedureInfo.GetParameter(Index: Integer): TvgrScriptProcedureParameter;
begin
  Result := TvgrScriptProcedureParameter(FParameters[Index]);
end;

function TvgrScriptProcedureInfo.GetParameterCount: Integer;
begin
  Result := FParameters.Count;
end;

/////////////////////////////////////////////////
//
// TvgrScriptProgramInfo
//
/////////////////////////////////////////////////
constructor TvgrScriptProgramInfo.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

destructor TvgrScriptProgramInfo.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TvgrScriptProgramInfo.AddProcedure(const AName: string; AProcedureType: TvgrScriptProcedureType; AStartPos, AEndPos: Integer): TvgrScriptProcedureInfo;
begin
  Result := TvgrScriptProcedureInfo.Create(AName, AProcedureType, AStartPos, AEndPos);
  FList.Add(Result);
end;

function TvgrScriptProgramInfo.IndexOfProcedure(const AProcedureName: string): Integer;
begin
  Result := 0;
  while (Result < ProcedureCount) and (AnsiCompareText(Procedures[Result].Name, AProcedureName) <> 0) do Inc(Result);
  if Result >= ProcedureCount then
    Result := -1;
end;

function TvgrScriptProgramInfo.GetProcedure(Index: Integer): TvgrScriptProcedureInfo;
begin
  Result := TvgrScriptProcedureInfo(FList[Index]);
end;

function TvgrScriptProgramInfo.GetProcedureCount: Integer;
begin
  Result := FList.Count;
end;

procedure TvgrScriptProgramInfo.Clear;
begin
  FreeListItems(FList);
  FList.Clear;
end;

/////////////////////////////////////////////////
//
// TvgrScriptParser
//
/////////////////////////////////////////////////
constructor TvgrScriptParser.Create;
begin
  inherited Create;
  FIgnoreCase := True;
  FKeywordDelimiters := ['''', '"', ' ', #1 .. #31, '(', ')', '+', '-', '\', '*', '[', ']', '{', '}'];
  FStringDelimiters := ['"'];
  FStringSpecialChar := '\';
end;

function TvgrScriptParser.IsEQKW(const AKeyWord: string; const ATestedKeyWord: string): Boolean;
begin
  if IgnoreCase then
    Result := AnsiCompareText(ATestedKeyWord, AKeyWord) = 0
  else
    Result := ATestedKeyWord = AKeyWord;
end;

function TvgrScriptParser.FindNextKeyword(const AText: string; var APos: Integer): string;
var
  I, ALen: Integer;
  AStringStarted: Boolean;
  AStringStartChar: Char;
begin
  Result := '';
  ALen := Length(AText);
  AStringStarted := False;
  AStringStartChar := #0;

  while (APos <= ALen) and ((AText[APos] in KeywordDelimiters) or AStringStarted) do
  begin
    if AText[APos] in FStringDelimiters then
    begin
      if AStringStarted then
      begin
        if AText[APos] = AStringStartChar then
          AStringStarted := False
      end
      else
      begin
        AStringStarted := True;
        AStringStartChar := AText[APos];
      end;
    end
    else
      if (AText[APos] = FStringSpecialChar) and (AStringStarted) then
        Inc(APos);
    Inc(APos);
  end;
  
  if APos <= ALen then
  begin
    I := APos;
    while (APos <= ALen) and not (AText[APos] in KeywordDelimiters) do Inc(APos);
    Result := Copy(AText, I, APos - I);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrVBScriptParser
//
/////////////////////////////////////////////////
constructor TvgrVBScriptParser.Create;
begin
  inherited;
  FKW_ProcedureStart := 'sub';
  FKW_FunctionStart := 'function';
  FKW_EndBlock := 'end';
  FKW_VarParameterPrefix := 'ByRef';
  FKW_ValParameterPrefix := 'ByVal';
  FStringSpecialChar := #0;
end;

procedure TvgrVBScriptParser.ParseProgram(const AProgram: string; AProgramInfo: TvgrScriptProgramInfo);
var
  I, AStartPos, J: Integer;
  AProcedureType: TvgrScriptProcedureType;
  AKeyWord, APriorKeyWord, AProcName: string;
begin
  AProgramInfo.Clear;
  I := 1;
  J := 0;
  APriorKeyWord := '';
  AProcedureType :=  vgrsptProcedure;
  AStartPos := 1;
  repeat
    AKeyWord := FindNextKeyWord(AProgram, I);
    if AKeyWord <> '' then
    begin
      if IsEQKW(AKeyWord, KW_ProcedureStart) or
         IsEQKW(AKeyWord, KW_FunctionStart) then
      begin
        if IsEQKW(APriorKeyWord, KW_EndBlock) then
        begin
          Dec(J);
          if J = 0 then
          begin
            // new procedure found
            AProgramInfo.AddProcedure(AProcName, AProcedureType, AStartPos, I);
          end;
        end
        else
        begin
          Inc(J);
          if J = 1 then
          begin
            // new procedure found
            AStartPos := I - Length(AKeyWord);
            AProcName := FindNextKeyWord(AProgram, I);
            if IsEQKW(AKeyWord, KW_ProcedureStart) then
              AProcedureType := vgrsptProcedure
            else
              AProcedureType := vgrsptFunction;
          end;
        end;
      end;
      APriorKeyWord := AKeyWord;
    end;
  until AKeyWord = '';
end;

procedure TvgrVBScriptParser.GenerateEventProcedureText(AEventInfo: TvgrRegisteredScriptEvent; const AProcedureName: string; var AText: string);
var
  I: Integer;
  S: string;
begin
  if AEventInfo.ProcedureType = vgrsptProcedure then
    S := KW_ProcedureStart
  else
    S := KW_FunctionStart;

  AText := Format('%s %s', [S, AProcedureName]);
  if AEventInfo.ParameterCount > 0 then
    AText := AText + '('
  else
    AText := AText + ' ';
  for I := 0 to AEventInfo.ParameterCount - 1 do
  begin
    if AEventInfo.Parameters[I].IsVar then
      AText := AText + KW_VarParameterPrefix
    else
      AText := AText + KW_ValParameterPrefix;
    AText := AText + ' ' + AEventInfo.Parameters[I].ParamName + ', ';
  end;
  if AEventInfo.ParameterCount > 0 then
  begin
    Delete(AText, Length(AText) - 1, 2);
    AText := AText + ')';
  end;
  AText := AText + #13#10#13#10 + KW_EndBlock + ' ' + S + #13#10;
end;

/////////////////////////////////////////////////
//
// TvgrJScriptParser
//
/////////////////////////////////////////////////
constructor TvgrJScriptParser.Create;
begin
  inherited;
  FKW_ProcedureStart := 'function';
  FKW_FunctionStart := 'function';
  FKW_StartBlock := '{';
  FKW_EndBlock := '}';
  FKW_VarParameterPrefix := '';
  FKW_ValParameterPrefix := '';
  FKeywordDelimiters := FKeywordDelimiters - ['{', '}'];
end;

procedure TvgrJScriptParser.ParseProgram(const AProgram: string; AProgramInfo: TvgrScriptProgramInfo);
var
  I, AStartPos, ABlockCounter: Integer;
  AKeyWord, AProcName: string;
begin
  AProgramInfo.Clear;
  I := 1;
  AStartPos := 1;
  ABlockCounter := 0;
  repeat
    AKeyWord := FindNextKeyWord(AProgram, I);
    if AKeyWord <> '' then
    begin
      if IsEQKW(AKeyWord, KW_FunctionStart) then
      begin
        AStartPos := I - Length(AKeyWord);
        AProcName := FindNextKeyWord(AProgram, I);
        if AProcName = '' then
          break;
      end
      else
        if IsEQKW(AKeyWord, KW_StartBlock) then
          Inc(ABlockCounter)
        else
          if IsEQKW(AKeyWord, KW_EndBlock) then
          begin
            Dec(ABlockCounter);
            if ABlockCounter < 0 then
              break
            else
              if ABlockCounter = 0 then
              begin
                // new procedure found
                AProgramInfo.AddProcedure(AProcName, vgrsptFunction, AStartPos, I);
              end;
          end;
    end;
  until AKeyWord = '';
end;

procedure TvgrJScriptParser.GenerateEventProcedureText(AEventInfo: TvgrRegisteredScriptEvent; const AProcedureName: string; var AText: string);
var
  I: Integer;
begin
  AText := Format('%s %s(', [FKW_FunctionStart, AProcedureName]);
  for I := 0 to AEventInfo.ParameterCount - 1 do
    AText := AText + AEventInfo.Parameters[I].ParamName + ', ';
  if AEventInfo.ParameterCount > 0 then
    Delete(AText, Length(AText) - 1, 2);
  AText := AText + ')';
  AText := AText + #13#10 + KW_StartBlock + #13#10 + KW_EndBlock + #13#10;
end;

initialization

  RegisteredScriptParsers := TvgrRegisteredScriptParsers.Create;

  RegisterScriptParser('VBScript', TvgrVBScriptParser);
  RegisterScriptParser('JScript', TvgrJScriptParser);

finalization

  FreeAndNil(RegisteredScriptParsers);

end.

