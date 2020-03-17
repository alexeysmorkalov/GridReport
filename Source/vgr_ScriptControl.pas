{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains the TvgrScriptControl class.}
unit vgr_ScriptControl;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  classes, ActiveX, windows, SysUtils, typinfo, 
  {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}

  vgr_ScriptComponentsProvider, vgr_AXScript, vgr_AliasManager, vgr_StringIDs,
  vgr_CommonClasses, vgr_Localize;

type
  TvgrScript = class;
  
{Holds the propertry names, used for COM processing.
Syntax:
  TNamesArray = array[0..0] of PWideChar;}
  TNamesArray = array[0..0] of PWideChar;
{Pointer to the TNamesArray array.
Syntax:
  PNamesArray = ^TNamesArray;}
  PNamesArray = ^TNamesArray;

{Holds the DispIs.
Syntax:
  TDispIdsArray = array[0..0] of Integer;}
  TDispIdsArray = array[0..0] of Integer;
{Pointer to the TDispIdsArray array.
Syntax:
  PDispIdsArray = ^TDispIdsArray;}
  PDispIdsArray = ^TDispIdsArray;

  /////////////////////////////////////////////////
  //
  // TvgrScriptConstant
  //
  /////////////////////////////////////////////////
{Represents a constant that can be used in script.
See also:
  TvgrScriptConstantList}
  TvgrScriptConstant = class
  private
    FName: string;
    FValue: Variant;
  public
{Creates an instance of the TvgrScriptConstant class.}
    constructor Create(const AName: string; const AValue: Variant);
{Specifies the name of constant.}
    property Name: string read FName;
{Specifies the value of constant.}
    property Value: Variant read FValue;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptConstantList
  //
  /////////////////////////////////////////////////
{Represents the list of constants which can be used in the script(TvgrScriptConstantList object).
Each constant in the list is identified by its name.
See also:
  TvgrScriptConstant, TvgrScriptGlobalFunction}
  TvgrScriptConstants = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrScriptConstant;
  public
{Adds a constant to the list, if constant already exists - changes its value.
Parameters:
  AName - Name of constant.
  AValue - Value of constant.}
    procedure Add(const AName: string; const AValue: Variant);
{Searches a constant in the list.
Parameters:
  AName - Name of constant.
  AValue - Contains the value of the found constant.
  ADataType - Contains the type of the found constant.
Return value:
  Returns the index of constant in the list.}
    function FindConstant(const AName: string): Integer; overload;
{Lists the constants.}
    property Items[Index: Integer]: TvgrScriptConstant read GetItem; default;
  end;

{TvgrScriptGlobalFunctionProc is a type of procedure that executes the global function.
Parameters:
  AParameters - The parameters of function.
  AResult - Receives the result of function.}
  TvgrScriptGlobalFunctionProc = procedure(var AParameters: TvgrOleVariantDynArray;
                                           var AResult: OleVariant);
  /////////////////////////////////////////////////
  //
  // TvgrScriptGlobalFunction
  //
  /////////////////////////////////////////////////
{Represents the global function that can be used is script.
See also:
  TvgrScriptConstant, TvgrScriptGlobalFunctions}
  TvgrScriptGlobalFunction = class(TObject)
  private
    FProc: TvgrScriptGlobalFunctionProc;
    FName: string;
    FMinParameters: Integer;
    FMaxParameters: Integer;
  public
{Specifies the procedure that is called to execute a function.
See also:
  TvgrScriptGlobalFunctionProc}
    property Proc: TvgrScriptGlobalFunctionProc read FProc;
{Specifies the name of function.}
    property Name: string read FName;
{Specifies the minimal number of parameters, if procedure can accept any number
of parameters this property returns -1.}
    property MinParameters: Integer read FMinParameters;
{Specifies the maximal number of parameters, if procedure can accept any number
of parameters this property returns -1.}
    property MaxParameters: Integer read FMaxParameters;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptGlobalFunctions
  //
  /////////////////////////////////////////////////
{Represents the list of global functions which can be used in the script(TvgrScriptGlobalFunction object).
Each function in the list is identified by its name.
See also:
  TvgrScriptGlobalFunction}
  TvgrScriptGlobalFunctions = class(TvgrObjectList)
  private
    function GetItm(Index: Integer): TvgrScriptGlobalFunction;
  public
{Registers the new global function.
Parameters:
  AFunctionName - The name of function, must be unique.
  AMinParameters - The min number of input parameters, set this parameter to -1
if function can accept any number of parameters.
  AMaxParameters - The max number of input parameters, set this parameter to -1
if function can accept any number of parameters.
  AProc - The procedure which will be called to execute a function.
See also:
  TvgrScriptGlobalFunctionProc}
    procedure RegisterFunction(const AFunctionName: string;
                               AMinParameters: Integer;
                               AMaxParameters: Integer;
                               AProc: TvgrScriptGlobalFunctionProc);
{Returns the index of function by its name.
Parameters:
  AName - The name of function.
Return value:
  Returns the name of function or -1 if function is not found.}
    function IndexByName(const AName: string): Integer;
{Lists all functions.}
    property Items[Index: Integer]: TvgrScriptGlobalFunction read GetItm; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptError
  //
  /////////////////////////////////////////////////
{Represents the information about an error in script.
Syntax:
  TvgrScriptError = record
    Line: Integer;
    Position: Integer;
    ErrorInfo: TExcepInfo;
  end;}
  TvgrScriptError = record
    Line: Integer;
    Position: Integer;
    ErrorInfo: TExcepInfo;
  end;

  /////////////////////////////////////////////////
  //
  // IvgrScriptContext
  //
  /////////////////////////////////////////////////
{IvgrScriptContext is used for interaction between TvgrScriptObject and owner object.}
  IvgrScriptContext = interface
{Returns the code of identificator.
Parameters:
  AName - The name of identificator.}
    function GetIdOfName(const AName: string): Integer;
    
{Executes an action that is specified by DispId parameter.
Parameters:
  ADispID - Code of identificator.}
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;

{Returns the value of the AliasManager property.
See also:
  AliasManager}
    function GetAliasManager: TvgrAliasManager;

{Procedure is called when an unknown object is found in the name space of script.
Parameters:
  AObjectName - Name of object.
  AObject - Interface to object.}
    procedure DoOnGetObject(const AObjectName: string; var AObject: IDispatch);

{Procedure is called when an error occurs while expression is calculated by the
EvaluateExpression method.
Parameters:
  AExpression - Expression that is calculated.
  AErrorDescription - Description of error.
  AResult - Must contains value of expression.}
    procedure DoOnErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);

{Returns the value of ParentWindow property.}
    function GetParentWindow: HWND;

{Returns the TvgrAliasManager object that can be used by script.}
    property AliasManager: TvgrAliasManager read GetAliasManager;

{Returns the handle of parent window that is used as parent window
for windows which are displayed by script functions (MsgBox, InputLine and so on).}
    property ParentWindow: HWND read GetParentWindow;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptObject
  //
  /////////////////////////////////////////////////
{The base scripting class. This class iteracts with Windows Scripting Host (WSH).}
  TvgrScriptObject = class(TObject, IUnknown, IActiveScriptSite, IActiveScriptSiteDebug, IDispatch, IActiveScriptSiteWindow)
  private
    FContext: IvgrScriptContext;
    FLanguage: WideString;

    FActiveScript: IActiveScript;
    FActiveScriptError: IActiveScriptError;

    FExpressionEvent: THandle;

    FScriptEnterCounter: Integer;

    function GetActiveScript(Value: WideString): IActiveScript;
    function GetActiveScriptDebug: IActiveScriptDebug;
    function GetActiveScriptParse: IActiveScriptParse;
    function GetAliasManager: TvgrAliasManager;
  protected
    // IUnknown
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    // IActiveScriptSite
    function GetLCID(out Lcid: TLCID): HRESULT; stdcall;
    function GetItemInfo(const pstrName: POleStr; dwReturnMask: DWORD; out ppiunkItem: IUnknown; out Info: ITypeInfo): HRESULT; stdcall;
    function GetDocVersionString(out Version: TBSTR): HRESULT; stdcall;
    function OnScriptTerminate(const pvarResult: OleVariant; const pexcepinfo: TExcepInfo): HRESULT; stdcall;
    function OnStateChange(ScriptState: TScriptState): HRESULT; stdcall;
    function OnScriptError(const pscripterror: IActiveScriptError): HRESULT; stdcall;
    function OnEnterScript: HRESULT; stdcall;
    function OnLeaveScript: HRESULT; stdcall;

    // IActiveSriptSiteWindow
    function GetWindow(out phwnd: HWND): HResult; stdcall;
    function EnableModeless(fEnable: BOOL): HResult; stdcall;
    
    // IActiveScriptSiteDebug
    function  GetDocumentContextFromPosition {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_19:0}dwSourceContext: LongWord;
                                                                                  {VT_19:0}uCharacterOffset: LongWord;
                                                                                  {VT_19:0}uNumChars: LongWord;
                                                                                  {VT_29:2}out ppsc: IDebugDocumentContext): HResult; stdcall;

    function  GetApplication {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppda: IDebugApplication): HResult; stdcall;
    function  GetRootApplicationNode {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdanRoot: IDebugApplicationNode): HResult; stdcall;
    function  OnScriptErrorDebug {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_29:1}const pErrorDebug: IActiveScriptErrorDebug;
                                                                      {VT_3:1}out pfEnterDebugger: Integer;
                                                                      {VT_3:1}out pfCallOnScriptErrorWhenContinuing: Integer): HResult; stdcall;
    // IDispatch
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual;  stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; virtual;
    function GetDispIdOfName(const AName: String) : Integer; virtual;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; virtual;

    function ParseScriptText(const AScript: WideString;
                                          AFlags: Integer;
                                          var ARes: OleVariant;
                                          var AScriptError: TvgrScriptError): HResult;

    procedure AddObject(const Name: WideString);
    procedure DoErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);
    procedure DoGetObject(const AObjectName: string; AObject: IDispatch);
    function GetErrorDesc(AErrorCode: HRESULT): string;
    procedure CheckParserCall(AResult: HRESULT; const AScriptError: TvgrScriptError); overload;
    procedure CheckParserCall(AResult: HRESULT); overload;
  public
{Creates an instance of the TvgrScriptControl class.}
    constructor Create(AContext: IvgrScriptContext);
{Frees an instance of the TvgrScriptControl class.}
    destructor Destroy; override;

{Sets the script language. This method must be called before using of other methods.}
    procedure SetScriptLanguage(const ALanguage: WideString);

{Releases all interfaces.}
    procedure Close;

{Starts the script executing mode.}
    procedure SetScript(const AScript: WideString);

{Evaluates an expression and returns the result.
This method can be used only after StartScriptExecuting method.
See also:
  StartScriptExecuting}
    function EvaluateExpression(const AExpression: string; const ADefault: Variant): OleVariant;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
  AParameters - Parameters of procedure, all parameters are passed by value, no var parameters.
Return value:
  If script procedure returns a value then returns its value.}
    function RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray): Variant; overload;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
  AParameters - Parameters of procedure.
  AByRef - Specifies the type of parameters (byref or not), length of this array must be equal to AParameters.
Return value:
  If script procedure returns a value then returns its value.}
    function RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray): Variant; overload;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
Return value:
  If script procedure returns a value then returns its value.}
    function RunProcedure(const AProcName: string): Variant; overload;

{Checks the script syntax.
Parameters:
  AScripts - Script to check.
  AScriptError - Contains information about the found error, if it exists.
Return value:
  Returns the true value if errors do not exist in the script and false otherwise.}
    function CheckSyntax(const AScript: WideString; out AScriptError: TvgrScriptError): Boolean;

{Returns the text highlight attributes.
Parameters:
  AText - Text;
  AAttribytes - Pointer to memory block.
Return value:
  Returns the true if function have executed successfully and AAttributes parameter contains
the new text attributes. If function returns the false then AAttributes contains nil.}
    function GetTextHighlightAttributes(const AText: string; var AAttributes: PWord): Boolean;

{Returns the script language that is currently used.}
    property Language: WideString read FLanguage;

{Returns the used IActiveScript interface.}
    property ActiveScript : IActiveScript read FActiveScript;

{Returns the used IActiveScriptParse interface.}
    property ActiveScriptParse: IActiveScriptParse read GetActiveScriptParse;

{Returns the used IActiveScriptDebug interface.}
    property ActiveScriptDebug: IActiveScriptDebug read GetActiveScriptDebug;

{Returns the TvgrAliasManager object that is used by the TvgrScriptObject.
See also:
  TvgrAliasManager}
    property AliasManager: TvgrAliasManager read GetAliasManager;
  end;

{TvgrScriptGetObject is the type for OnScriptGetObject event.
This event occurs when an unknown object is found in the script.
Parameters:
  AName - The name of script object.
  AObject - Interface (IDispatch) that must be used for access to object with AName name. 
Syntax:
  TvgrScriptGetObject = procedure(Sender: TObject; AName: string; out AObject: IDispatch) of object;}
  TvgrScriptGetObject = procedure(Sender: TObject; AName: string; out AObject: IDispatch) of object;

{TvgrScriptErrorInExpression is the type for OnErrorInExpression event.
This event occurs when an error in calculating of expression occurs.
Parameters:
  Sender - The TvgrScriptControl object.
  AExpression - The expression that can not be calculated by script.
  ADescription - Description of error.
  Value - You may return value of expression in this parameter.
Syntax:
  TvgrScriptErrorInExpression = procedure(Sender: TObject; const AExpression: string;
                                          const ADescription: string; var Value: OleVariant)}
  TvgrScriptErrorInExpression = procedure(Sender: TObject; const AExpression: string; const ADescription: string; var Value: OleVariant) of object;

  /////////////////////////////////////////////////
  //
  // TvgrScriptControl
  //
  /////////////////////////////////////////////////
{TvgrScriptControl is a component that can interacts with Windows Scripting Host. 
It can be used for:<br>
<ul>
<li>- Scripts executing.</li>
<li>- Syntax highlighting.</li>
<li>- Scripts syntax checking.</li>
</ul>}
  TvgrScriptControl = class(TvgrComponent, IvgrScriptContext)
  private
    FScript: TStringList;
    FScriptObject: TvgrScriptObject;
    FAvailableComponentsProvider: TvgrAvailableComponentsProvider;
    FParentWindowHandle: HWND;
    FAliasManager: TvgrAliasManager;
    FLanguage: string;

    FOnScriptGetObject: TvgrScriptGetObject;
    FOnScriptErrorInExpression: TvgrScriptErrorInExpression;

    function GetOnGetAvailableComponents: TvgrGetAvailableComponentsEvent;
    procedure SetOnGetAvailableComponents(Value: TvgrGetAvailableComponentsEvent);

    function GetScript: TStrings;
    procedure SetScript(Value: TStrings);
    function GetBusy: Boolean;
    procedure SetAliasManager(Value: TvgrAliasManager);
    procedure SetLanguage(Value: string);

    procedure OnScriptChange(Sender: TObject);
    procedure CheckBusy;
    procedure CheckReady;
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;

    procedure DoScriptGetObject(const AObjectName: string; var AObject: IDispatch);
    procedure DoScriptErrorInExpression(const AExpression: string; const ADescription: string; var Value: OleVariant);
    
    // IvgrScriptContext
    function IvgrScriptContext_GetIdOfName(const AName: string): Integer;
    function IvgrScriptContext_DoInvoke(DispId: Integer;
                                        Flags: Integer;
                                        var AParameters: TvgrOleVariantDynArray;
                                        var AResult: OleVariant): HResult;
    function GetAliasManager: TvgrAliasManager;
    procedure DoOnGetObject(const AObjectName: string; var AObject: IDispatch);
    procedure DoOnErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);
    function GetParentWindow: HWND;
    function IvgrScriptContext.DoInvoke = IvgrScriptContext_DoInvoke;
    function IvgrScriptContext.GetIdOfName = IvgrScriptContext_GetIdOfName;
  public
{Creates an instance of the TvgrScriptControl class.}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of the TvgrScriptControl class.}
    destructor Destroy; override;

{Use this method to initialize script engine.
See also:
  EndScriptExecuting}
    procedure BeginScriptExecuting;

{Use this method to close script engine.
See also:
  BeginScriptExecuting}
    procedure EndScriptExecuting;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
  AParameters - Parameters of procedure.
  AByRef - Specifies the type of parameters (byref or not), length of this array must be equal to length of AParameters array.
Return value:
  If script procedure returns a value then returns its value.
See also:
  BeginScriptExecuting}
    function RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray): Variant; overload;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
  AParameters - Parameters of procedure all parameters are passed by value (no var parameters).
Return value:
  If script procedure returns a value then returns its value.
See also:
  BeginScriptExecuting}
    function RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray): Variant; overload;

{Runs the script procedure, this method can be used only after the BeginScriptExecuting method.
Parameters:
  AProcName - Name of procedure if procedure is not found in script then an exception will be raised.
Return value:
  If script procedure returns a value then returns its value.
See also:
  BeginScriptExecuting}
    function RunProcedure(const AProcName: string): Variant; overload;

{Calculates an expression and returns the result.
Parameters:
  AExpression - Expression for calculating.
  ADefault - The default value that will be returned if error occurs during the calculating.
Return value:
  Returns the value of expression.}
    function EvaluateExpression(const AExpression: string; const ADefault: Variant): Variant;

{Checks the script.
Parameters:
  AScript - Script to check.
  AError - If error is found then contains the error description.
Return value:
  Returns the true if no errors are found.}
    function CheckSyntax(const AScript: string; var AError: TvgrScriptError): Boolean;

{Returns the text highlight attributes.
Parameters:
  AText - Text;
  AAttribytes - Pointer to memory block.
Return value:
  Returns the true if function have executed successfully and AAttributes parameter contains
the new text attributes. If function returns the false then AAttributes contains nil.}
    function GetTextHighlightAttributes(const AText: string; var AAttributes: PWord): Boolean;

{Returns the value indicating whether the object executes the script and its properties
can not be changed in this time.}
    property Busy: Boolean read GetBusy;

{Returns the handle of parent window that is used as parent window
for windows which are displayed by script functions (MsgBox, InputLine and so on).}
    property ParentWindowHandle: HWND read FParentWindowHandle write FParentWindowHandle;
  published
{Alias manager component.}
    property AliasManager: TvgrAliasManager read FAliasManager write SetAliasManager;

{Specifies the language that is used in the script.}
    property Language: string read FLanguage write SetLanguage;

{Specifies the script.}
    property Script: TStrings read GetScript write SetScript;

{Occurs when an unknown object is found in the script.
See also:
  TvgrScriptGetObject}
    property OnScriptGetObject: TvgrScriptGetObject read FOnScriptGetObject write FOnScriptGetObject;
    
{Occurs when an error in calculating of value of expression occurs (in EvaluateExpression method).
See also:
  TvgrScriptErrorInExpression}
    property OnScriptErrorInExpression: TvgrScriptErrorInExpression read FOnScriptErrorInExpression write FOnScriptErrorInExpression;

{Occurs when list of available objects is formed
(in the Script.GetAvailableComponents method).
See also:
  Script.GetAvailableComponents, TvgrGetAvailableComponentsEvent}
    property OnGetAvailableComponents: TvgrGetAvailableComponentsEvent read GetOnGetAvailableComponents write SetOnGetAvailableComponents;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScript
  //
  /////////////////////////////////////////////////
{Represents the script options for GridReport objects that can execute the scripts.
This object holds script data and contains methods for executing the script procedures
and for evaluating the expressions.}
  TvgrScript = class(TvgrAvailableComponentsProvider, IvgrScriptContext)
  private
    FScriptObject: TvgrScriptObject;
    FLanguage: string;
    FScript: TStringList;
    FAliasManager: TvgrAliasManager;

    FOnScriptGetObject: TvgrScriptGetObject;
    FOnScriptErrorInExpression: TvgrScriptErrorInExpression;
    FParentWindowHandle: HWND;

    procedure SetLanguage(Value: string);
    function GetStrings: TStrings;
    procedure SetStrings(Value: TStrings);
    procedure OnScriptChange(Sender: TObject);
    procedure CheckBusy;
    procedure CheckReady;
    function GetBusy: Boolean;
  protected
    function GetOwner: TPersistent; override;

    procedure ReadScriptControl(Reader: TReader);
    procedure WriteScriptControl(Writer: TWriter);
    procedure DefineProperties(Filer: TFiler); override;

    procedure DoScriptGetObject(const AObjectName: string; var AObject: IDispatch);
    procedure DoScriptErrorInExpression(const AExpression: string; const ADescription: string; var Value: OleVariant);

    // IvgrScriptContext
    function IvgrScriptContext_GetIdOfName(const AName: string): Integer;
    function IvgrScriptContext_DoInvoke(DispId: Integer;
                                        Flags: Integer;
                                        var AParameters: TvgrOleVariantDynArray;
                                        var AResult: OleVariant): HResult;
    function GetAliasManager: TvgrAliasManager;
    procedure DoOnGetObject(const AObjectName: string; var AObject: IDispatch);
    procedure DoOnErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);
    function GetParentWindow: HWND;
    function IvgrScriptContext.DoInvoke = IvgrScriptContext_DoInvoke;
    function IvgrScriptContext.GetIdOfName = IvgrScriptContext_GetIdOfName;
  public
{Creates an instance of the TvgrScriptOptions class.}
    constructor Create; override; 

{Frees an instance of the TvgrScriptOptions class.}
    destructor Destroy; override;

{Copies the contents of another, similar object.
Parameters:
  Source - The source object.}
    procedure Assign(Source: TPersistent); override;

{Use this method to initialize script engine.
See also:
  EndScriptExecuting}
    procedure BeginScriptExecuting;
    
{Use this method to close script engine.
See also:
  BeginScriptExecuting}
    procedure EndScriptExecuting;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
  AParameters - Parameters of procedure.
  AByRef - Specifies the type of parameters (byref or not), length of this array must be equal to AParameters.
Return value:
  If script procedure returns a value then returns its value.
See also:
  BeginScriptExecuting}
    function RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray): Variant; overload;

{Executes the procedure that exists in the currently executing script.
This method can be used only after StartScriptExecuting method.
Parameters:
  AProcName - Name of procedure.
  AParameters - Parameters of procedure all parameters are passed by value (no var parameters).
Return value:
  If script procedure returns a value then returns its value.
See also:
  BeginScriptExecuting}
    function RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray): Variant; overload;

{Runs the script procedure, this method can be used only after the BeginScriptExecuting method.
Parameters:
  AProcName - Name of procedure if procedure is not found in script then an exception will be raised.
Return value:
  If script procedure returns a value then returns its value.
See also:
  BeginScriptExecuting}
    function RunProcedure(const AProcName: string): Variant; overload;

{Calculates an expression and returns the result.
Parameters:
  AExpression - Expression for calculating.
  ADefault - The default value that will be returned if error occurs during the calculating.
Return value:
  Returns the value of expression.}
    function EvaluateExpression(const AExpression: string; const ADefault: Variant): Variant;

{Returns the value indicating whether the object executes the script and its properties
can not be changed in this time.}
    property Busy: Boolean read GetBusy;

{Occurs when an unknown object is found in the script.
See also:
  TvgrScriptGetObject}
    property OnScriptGetObject: TvgrScriptGetObject read FOnScriptGetObject write FOnScriptGetObject;
{Occurs when an error in calculating of value of expression occurs (in EvaluateExpression method).
See also:
  TvgrScriptErrorInExpression}
    property OnScriptErrorInExpression: TvgrScriptErrorInExpression read FOnScriptErrorInExpression write FOnScriptErrorInExpression;

{Returns the handle of parent window that is used as parent window
for windows which are displayed by script functions (MsgBox, InputLine and so on).}
    property ParentWindowHandle: HWND read FParentWindowHandle write FParentWindowHandle;
  published
{Specifies the language that is used in the script.}
    property Language: string read FLanguage write SetLanguage;
{Specifies the script.}
    property Script: TStrings read GetStrings write SetStrings;
{Specifies the TvgrAliasManager object that is used in the script.}
    property AliasManager: TvgrAliasManager read FAliasManager write FAliasManager;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrScriptSyntax
  //
  /////////////////////////////////////////////////
{Represents the script options for GridReport objects which require the syntax checking or text highlighting.
This object holds script data and contains methods for syntax checking and text highlighting.}
  TvgrScriptSyntax = class(TvgrPersistent)
  private
    FScriptObject: TvgrScriptObject;

    function GetLanguage: string;
    procedure SetLanguage(Value: string);
  public
{Creates an instance of the TvgrScriptSyntax class.}
    constructor Create; override;
{Frees an instance of the TvgrScriptSyntax class.}
    destructor Destroy; override;

{Checks the script.
Parameters:
  AScript - Script to check.
  AError - If error is found then contains the error description.
Return value:
  Returns the true if no errors are found.}
    function CheckSyntax(const AScript: string; var AError: TvgrScriptError): Boolean;

{Returns the text highlight attributes.
Parameters:
  AText - Text;
  AAttribytes - Pointer to memory block.
Return value:
  Returns the true if function have executed successfully and AAttributes parameter contains
the new text attributes. If function returns the false then AAttributes contains nil.}
    function GetTextHighlightAttributes(const AText: string; var AAttributes: PWord): Boolean;
  published
{Specifies the script language.}
    property Language: string read GetLanguage write SetLanguage;
  end;

var
{Global script constants list.}
  vgrScriptConstants: TvgrScriptConstants;
{The list of global script functions.}
  vgrScriptFunctions: TvgrScriptGlobalFunctions;

implementation

{$R ..\Res\vgr_ScriptControlStrings.res}

uses ComObj, Math,
  vgr_Functions, vgr_ScriptGlobalFunctions, vgr_WBScriptConstants;

const
  sse_ProcedureNotFoundInScript = 'Procedure [%s] is not found in the script.';
  sse_ScriptLanguageNotSupported = 'Script language [%s] is not supported.';
  sse_DISP_E_EXCEPTION = 'DISP_E_EXCEPTION';
  sse_E_INVALIDARG  = 'E_INVALIDARG';
  sse_E_POINTER = 'E_POINTER';
  sse_E_NOTIMPL = 'E_NOTIMPL';
  sse_E_UNEXPECTED = 'E_UNEXPECTED';
  sse_E_FAIL = 'E_FAIL';
  sse_S_FALSE = 'S_FALSE';
  sse_ScriptError = 'Script error:'#13'%s'#13'Addititional information:'#13'%s';
  sse_None = 'None';
  sse_DISP_E_EXCEPTION_Info = 'Source: Line = %d, Position = %d'#13'Description: %s';
  sse_NotInitializated = 'Object is not initializated, use SetScriptLanguage method.';
  sse_ObjectNotFound = 'Object [%s] in script name space is not found';

  sse_ProcedureExecuting1 = 'Error executing procedure [%s]'#13'Error: %s'#13#13'Source:'#13'%s'#13'Description:'#13'%s';
  sse_ProcedureExecuting2 = 'Error executing procedure [%s]'#13'Error: %s'#13#13'No addititional information (may be an invalid list of parameters is specified).';

  sse_ScriptObjectIsBusy = 'The object is busy (probably executes a script) and its properties can not be changed.';
  sse_ScriptObjectIsNotInitialized = 'The object is not ready to execute a script.';

  sse_ParameterTypeNotSupported = 'Parameters of type [%s] is not supported.';
  sse_SourcePosition = 'Row: %d, Column: %d';

{Represents the default value that is returned when error in procedure occurs.}
  svgrSCError: WideString = 'ERROR: %s';

type
  TvgrAvailableComponentsProviderAccess = class(TvgrAvailableComponentsProvider)
  end;

/////////////////////////////////////////////////
//
// TvgrScriptGlobalFunctions
//
/////////////////////////////////////////////////
function TvgrScriptGlobalFunctions.GetItm(Index: Integer): TvgrScriptGlobalFunction;
begin
  Result := TvgrScriptGlobalFunction(inherited Items[Index]);
end;

function TvgrScriptGlobalFunctions.IndexByName(const AName: string): Integer;
begin
  Result := 0;
  while (Result < Count) and (AnsiCompareText(Items[Result].Name, AName) <> 0) do
    Inc(Result);
  if Result >= Count then
    Result := -1; 
end;

procedure TvgrScriptGlobalFunctions.RegisterFunction(const AFunctionName: string;
                                                     AMinParameters: Integer;
                                                     AMaxParameters: Integer;
                                                     AProc: TvgrScriptGlobalFunctionProc);
var
  I: Integer;
  AFunction: TvgrScriptGlobalFunction;
begin
  I := IndexByName(AFunctionName);
  if I <> -1 then
    raise Exception.CreateFmt('Global function with name [%s] is already registered.', [AFunctionName]);

  AFunction := TvgrScriptGlobalFunction.Create;
  with AFunction do
  begin
    FProc := AProc;
    FName := AFunctionName;
    FMinParameters := AMinParameters;
    FMaxParameters := AMaxParameters;
  end;
  inherited Add(AFunction);
end;

(*
type

TStringDesc = record
  BStr: PWideChar;
  PStr: PString;
end;
pStringDesc = ^TStringDesc;

TvgrStringDescArray = array [0..65535] of TStringDesc;
pvgrStringDescArray = ^TvgrStringDescArray;

/////////////////////////////////////////////////
//
// TvgrScriptParametersConverter
//
/////////////////////////////////////////////////
TvgrScriptParametersConverter = class(TObject)
private
//  FStrCount: Integer;
//  FStrings: pvgrStringDescArray;
  FStrings: TvgrPointerDynArray;
  FByRefParameters: TvgrPointerDynArray;

  FArgs: PVariantArgList;
  FArgsCount: Integer;

  procedure CheckVariantList(AValues: TvgrVariantDynArray);
public
  constructor Create(AValues: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray);
  destructor Destroy; override;

  property Args: PVariantArgList read FArgs;
  property ArgsCount: Integer read FArgsCount;
end;

constructor TvgrScriptParametersConverter.Create(AValues: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray);
var
  I: Integer;
  AVarData: PVarData;
  AArg: PVariantArg;
  AVarType: Word;
  AOleV: OleVariant;
begin
  inherited Create;

  CheckVariantList(AValues);

  FArgs := nil;
  FArgsCount := Length(AValues);

  if FArgsCount <= 0 then exit;

  SetLength(FStrings, FArgsCount);
  SetLength(FByRefParameters, FArgsCount);
  GetMem(FArgs, FArgsCount * SizeOf(TVariantArg));
  
  for I := 0 to FArgsCount - 1 do
  begin
    AArg := @(FArgs^[FArgsCount - I - 1]);
    AVarData := @AValues[I];

    AVarType := VarType(AValues[I]);
    AArg.vt := AVarType;
    AArg.wReserved1 := 0;
    AArg.wReserved2 := 0;
    AArg.wReserved3 := 0;

    if AByRef[I] then
    begin
      // parameter is passed by reference
      AArg.pvarVal := AllocMem(SizeOf(OleVariant));
      AOleV := AValues[I];
      AArg.pvarVal^ := AOleV;
      AArg.vt := VT_VARIANT or VT_BYREF;
      {
      case AVarType of
        varSmallint:
          begin
            GetMem(AArg.piVal, SizeOf(SmallInt));
            AArg.piVal^ := AVarData.VSmallInt;
            FByRefParameters[I] := AArg.piVal;
          end;
        varInteger:
          begin
            GetMem(AArg.plVal, SizeOf(Integer));
            AArg.plVal^ := AVarData.VInteger;
            FByRefParameters[I] := AArg.plVal;
          end;
        varSingle:
          begin
            GetMem(AArg.pfltVal, SizeOf(Single));
            AArg.pfltVal^ := AVarData.VSingle;
            FByRefParameters[I] := AArg.pfltVal;
          end;
        varDouble:
          begin
            GetMem(AArg.pdblVal, SizeOf(Double));
            AArg.pdblVal^ := AVarData.VDouble;
            FByRefParameters[I] := AArg.pdblVal;
          end;
        varCurrency:
          begin
            GetMem(AArg.pcyVal, SizeOf(Currency));
            AArg.pcyVal^ := AVarData.VCurrency;
            FByRefParameters[I] := AArg.pcyVal;
          end;
        varDate:
          begin
            GetMem(AArg.pdate, SizeOf(TOleDate));
            AArg.pdate^ := AVarData.VDate;
            FByRefParameters[I] := AArg.pdate;
          end;
        varOleStr:
          begin
            AArg.pbstrVal := @AVarData.VOleStr; // ???
            FByRefParameters[I] := AArg.pbstrVal;
          end;
        varBoolean:
          begin
            GetMem(AArg.pbool, sizeof(TOleBool));
            AArg.pbool^ := AVarData.VBoolean;
            FByRefParameters[I] := AArg.pbool;
          end;
        varUnknown:
          begin
            GetMem(AArg.punkVal, SizeOf(Pointer));
            AArg.punkVal := AVarData.VUnknown;
            FByRefParameters[I] := AArg.punkVal;
          end;
        varDispatch:
          begin
            GetMem(AArg.pdispVal, SizeOf(Pointer));
            AArg.pdispVal := AVarData.VDispatch;
            FByRefParameters[I] := AArg.pdispVal;
          end;
        varShortInt:
          begin
            GetMem(AArg.pcVal, SizeOf(char));
            AArg.pcVal^ := char(AVarData.VShortInt);
            FByRefParameters[I] := AArg.pcVal;
          end;
        varByte:
          begin
            GetMem(AArg.pbVal, SizeOf(byte));
            AArg.pbVal^ := AVarData.VByte;
            FByRefParameters[I] := AArg.pbVal;
          end;
        varWord:
          begin
            GetMem(AArg.puiVal, SizeOf(Word));
            AArg.puiVal^ := AVarData.VWord;
            FByRefParameters[I] := AArg.puiVal;
          end;
        varLongWord:
          begin
            GetMem(AArg.pulVal, SizeOf(LongWord));
            AArg.pulVal^ := AVarData.VLongWord;
            FByRefParameters[I] := AArg.pulVal;
          end;
        varString:
          begin
            FStrings[I] := StringToOleStr(string(AVarData.VString));
            AArg.vt := VT_BSTR;
            AArg.bstrVal := @FStrings[I];
          end;
        else AArg.vt := varNull;
      end;
      if AArg.vt <> varNull then
        AArg.vt := AArg.vt or VT_BYREF;
        }
    end
    else
    begin
      // parameter is passed by value
      case AVarType of
        varSmallint: AArg.iVal := AVarData.VSmallInt;
        varInteger: AArg.lVal := AVarData.VInteger;
        varSingle: AArg.fltVal := AVarData.VSingle;
        varDouble: AArg.dblVal := AVarData.VDouble;
        varCurrency: AArg.cyVal := AVarData.VCurrency;
        varDate: AArg.date := AVarData.VDate;
        varOleStr: AArg.bstrVal := AVarData.VOleStr;
        varBoolean: AArg.vbool := AVarData.VBoolean;
        varUnknown: AArg.unkVal := AVarData.VUnknown;
        varDispatch: AArg.dispVal := AVarData.VDispatch;
        varShortInt: AArg.cVal := char(AVarData.VShortInt);
        varByte: AArg.bVal := AVarData.VByte;
        varWord: AArg.uiVal := AVarData.VWord;
        varLongWord: AArg.ulVal := AVarData.VLongWord;
        varString:
          begin
            FStrings[I] := StringToOleStr(string(AVarData.VString));
            AArg.vt := VT_BSTR;
            AArg.bstrVal := FStrings[I];
          end;
        else AArg.vt := varNull;
      end;
    end;
  end;
end;

destructor TvgrScriptParametersConverter.Destroy;
var
  I: Integer;
begin
  for I := 0 to High(FStrings) do
    if FStrings[I] <> nil then
      SysFreeString(FStrings[I]);

  for I := 0 to High(FByRefParameters) do
    if FByRefParameters[I] <> nil then
      FreeMem(FByRefParameters[I]);
  inherited;
end;

{
constructor TvgrScriptParametersConverter.Create(AValues: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray);
type
  PVarArg = ^TVarArg;
  TVarArg = array[0..3] of DWORD;
var
  I, ArgType: Integer;
  AVarData: PVarData;
  AArg: PVariantArg;
  AVarType: Word;

  ParamPtr: ^Integer;
  ArgPtr, VarPtr: PVarArg;
  VarFlag: Boolean;
begin
  inherited Create;

  CheckVariantList(AValues);

  FArgs := nil;
  FArgsCount := Length(AValues);

  if FArgsCount <= 0 then exit;

  FArgs := AllocMem(FArgsCount * SizeOf(TVariantArg)); // allocate memory for dest parameters
  FStrings := AllocMem(SizeOf(TStringDesc) * SizeOf(TVariantArg)); // allocate memory for strings


  ParamPtr := @AValues[0]; // pointer to the current source parameter
  ArgPtr := @(FArgs^[FArgsCount]); // pointer to the current dest parameter
  I := 0;
  FStrCount := 0;
  repeat
    Dec(Integer(ArgPtr), SizeOf(TVarData));
    ArgType := VarType(AValues[I]) and varTypeMask;
    VarFlag := AByRef[I];

    if ArgType = varError then
    begin
      ArgPtr^[0] := varError;
      ArgPtr^[2] := DWORD(DISP_E_PARAMNOTFOUND);
    end
    else
    begin
      if ArgType = varStrArg then
      begin
        with FStrings^[FStrCount] do
          if VarFlag then
          begin
            BStr := StringToOleStr(PString(ParamPtr^)^);
            PStr := PString(ParamPtr^);
            ArgPtr^[0] := varOleStr or varByRef;
            ArgPtr^[2] := Integer(@BStr);
          end
          else
          begin
            BStr := StringToOleStr(PString(ParamPtr)^);
            PStr := nil;
            ArgPtr^[0] := varOleStr;
            ArgPtr^[2] := Integer(BStr);
          end;
        Inc(FStrCount);
      end

      else if VarFlag then
      begin
        if (ArgType = varVariant) and
           (PVarData(ParamPtr^)^.VType = varString) then
          VarCast(PVariant(ParamPtr^)^, PVariant(ParamPtr^)^, varOleStr);

        ArgPtr^[0] := ArgType or varByRef;
        ArgPtr^[2] := ParamPtr^;
      end

      else if ArgType = varVariant then
      begin
        if PVarData(ParamPtr)^.VType = varString then
        begin
          with FStrings[FStrCount] do
          begin
            BStr := StringToOleStr(string(PVarData(ParamPtr)^.VString));
            PStr := nil;
            ArgPtr^[0] := varOleStr;
            ArgPtr^[2] := Integer(BStr);
          end;
          Inc(FStrCount);
        end
        else
        begin
          VarPtr := PVarArg(ParamPtr);
          ArgPtr^[0] := VarPtr^[0];
          ArgPtr^[1] := VarPtr^[1];
          ArgPtr^[2] := VarPtr^[2];
          ArgPtr^[3] := VarPtr^[3];
          Inc(Integer(ParamPtr), 12);
        end;
      end

      else
      begin
        ArgPtr^[0] := ArgType;
        ArgPtr^[2] := ParamPtr^;
        if (ArgType >= varDouble) and (ArgType <= varDate) then
        begin
          Inc(Integer(ParamPtr), 4);
          ArgPtr^[3] := ParamPtr^;
        end;
      end;
      Inc(Integer(ParamPtr), 4);
    end;
    Inc(I);
  until I = FArgsCount;
end;

destructor TvgrScriptParametersConverter.Destroy;
var
  I: Integer;
begin
  if FArgs <> nil then
  begin
    FreeMem(FArgs);
    FArgs := 0;
  end;

  if FStrings <> nil then
  begin
    for I := 0 to FStrCount - 1 do
      if FStrings^[I].BStr <> nil then
        SysFreeString(FStrings^[I].BStr);
    FreeMem(FStrings);
    FStrings := nil;
  end;
  inherited;
end;
}

procedure TvgrScriptParametersConverter.CheckVariantList(AValues: TvgrVariantDynArray);
var
  I: Integer;
  AVarType: Word;
begin
  for I := 0 to High(AValues) do
  begin
    AVarType := VarType(AValues[I]);
    if (AVarType = varError) or
       (AVarType = varStrArg) or
       (AVarType = varAny) or
       (AVarType = varArray) or
       (AVarType = varInt64) then
      raise Exception.CreateFmt(sse_ParameterTypeNotSupported, [VarTypeAsText(AVarType)]);
  end;
end;
*)

/////////////////////////////////////////////////
//
// TvgrScriptConstant
//
/////////////////////////////////////////////////
//  public
constructor TvgrScriptConstant.Create(const AName: string; const AValue: Variant);
begin
  inherited Create;
  FName := AName;
  FValue := AValue;
end;

/////////////////////////////////////////////////
//
// TvgrScriptConstantList
//
/////////////////////////////////////////////////
function TvgrScriptConstants.GetItem(Index: Integer): TvgrScriptConstant;
begin
  Result := TvgrScriptConstant(inherited Items[Index]);
end;

procedure TvgrScriptConstants.Add(const AName: string; const AValue: Variant);
var
  I: Integer;
  AConstant: TvgrScriptConstant;
begin
  I := FindConstant(AName);
  if I >= 0 then
    Items[I].FValue := AValue
  else
  begin
    AConstant := TvgrScriptConstant.Create(AName, AValue);
    inherited Add(AConstant);
  end;
end;

function TvgrScriptConstants.FindConstant(const AName: string): Integer;
begin
  Result := 0;
  while (Result < Count) and not (AnsiCompareText(Items[Result].Name, AName) = 0) do
    Inc(Result);
  if Result >= Count then
    Result := -1;
end;

/////////////////////////////////////////////////
//
// TvgrScriptControl
//
/////////////////////////////////////////////////
constructor TvgrScriptControl.Create(AOwner: TComponent);
begin
  inherited;
  FScript := TStringList.Create;
  FScript.OnChange := OnScriptChange;
  FAvailableComponentsProvider := TvgrAvailableComponentsProvider.Create;
  FAvailableComponentsProvider.Root := Self;
end;

destructor TvgrScriptControl.Destroy;
begin
  FAvailableComponentsProvider.Free;
  FScript.Free;
  FreeAndNil(FScriptObject);
  inherited;
end;

procedure TvgrScriptControl.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if AComponent = FAliasManager then
      FAliasManager := nil;
  end;
end;

function TvgrScriptControl.GetOnGetAvailableComponents: TvgrGetAvailableComponentsEvent;
begin
  Result := FAvailableComponentsProvider.OnGetAvailableComponents;
end;

procedure TvgrScriptControl.SetOnGetAvailableComponents(Value: TvgrGetAvailableComponentsEvent);
begin
  FAvailableComponentsProvider.OnGetAvailableComponents := Value;
end;

function TvgrScriptControl.GetScript: TStrings;
begin
  Result := FScript;
end;

procedure TvgrScriptControl.SetScript(Value: TStrings);
begin
  CheckBusy;
  FScript.Assign(Value);
end;

function TvgrScriptControl.GetBusy: Boolean;
begin
  Result := FScriptObject <> nil;
end;

procedure TvgrScriptControl.SetAliasManager(Value: TvgrAliasManager);
begin
  CheckBusy;
  if FAliasManager <> Value then
  begin
    FAliasManager := Value;
  end;
end;

procedure TvgrScriptControl.SetLanguage(Value: string);
begin
  CheckBusy;
  if FLanguage <> Value then
    FLanguage := Value; 
end;

procedure TvgrScriptControl.OnScriptChange(Sender: TObject);
begin
  CheckBusy;
end;

procedure TvgrScriptControl.CheckBusy;
begin
  if Busy then
    raise Exception.Create(sse_ScriptObjectIsBusy);
end;

procedure TvgrScriptControl.CheckReady;
begin
  if not Busy then
    raise Exception.Create(sse_ScriptObjectIsNotInitialized);
end;

procedure TvgrScriptControl.DoScriptGetObject(const AObjectName: string; var AObject: IDispatch);
begin
  AObject := nil;
  if Assigned(FOnScriptGetObject) then
    FOnScriptGetObject(Self, AObjectName, AObject);
end;

procedure TvgrScriptControl.DoScriptErrorInExpression(const AExpression: string; const ADescription: string; var Value: OleVariant);
begin
  if Assigned(FOnScriptErrorInExpression) then
    FOnScriptErrorInExpression(Self, AExpression, ADescription, Value);
end;

function TvgrScriptControl.IvgrScriptContext_GetIdOfName(const AName: string): Integer;
begin
  Result := IndexInStrings(AName, TvgrAvailableComponentsProviderAccess(FAvailableComponentsProvider).FCache);
end;

function TvgrScriptControl.IvgrScriptContext_DoInvoke(DispId: Integer;
                                    Flags: Integer;
                                    var AParameters: TvgrOleVariantDynArray;
                                    var AResult: OleVariant): HResult;
var
  AInterface: IUnknown;
begin
  if DispId >= TvgrAvailableComponentsProviderAccess(FAvailableComponentsProvider).FCache.Count then
    Result := E_UNEXPECTED
  else
  begin
    AInterface := GetInterfaceToObject(TObject( TvgrAvailableComponentsProviderAccess(FAvailableComponentsProvider).FCache.Objects[DispId]));
    if AInterface = nil then
      Result := E_UNEXPECTED
    else
    begin
      AResult := AInterface;
      Result := S_OK;
    end;
  end;
end;

function TvgrScriptControl.GetAliasManager: TvgrAliasManager;
begin
  Result := FAliasManager;
end;

procedure TvgrScriptControl.DoOnGetObject(const AObjectName: string; var AObject: IDispatch);
begin
  DoScriptGetObject(AObjectName, AObject);
end;

procedure TvgrScriptControl.DoOnErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);
begin
  DoScriptErrorInExpression(AExpression, AErrorDescription, AResult);
end;

function TvgrScriptControl.GetParentWindow: HWND;
begin
  Result := FParentWindowHandle;
end;

procedure TvgrScriptControl.BeginScriptExecuting;
begin
  FScriptObject := TvgrScriptObject.Create(Self);
  FScriptObject.SetScriptLanguage(FLanguage);
  FScriptObject.SetScript(FScript.Text);
  FAvailableComponentsProvider.Cache;
end;

procedure TvgrScriptControl.EndScriptExecuting;
begin
  FAvailableComponentsProvider.ClearCache;
  FreeAndNil(FScriptObject);
end;

function TvgrScriptControl.RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray): Variant;
begin
  CheckReady;
  Result := FScriptObject.RunProcedure(AProcName, AParameters, AByRef);
end;

function TvgrScriptControl.RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray): Variant;
begin
  CheckReady;
  Result := FScriptObject.RunProcedure(AProcName, AParameters);
end;

function TvgrScriptControl.RunProcedure(const AProcName: string): Variant;
begin
  CheckReady;
  Result := FScriptObject.RunProcedure(AProcName);
end;

function TvgrScriptControl.EvaluateExpression(const AExpression: string; const ADefault: Variant): Variant;
begin
  CheckReady;
  Result := FScriptObject.EvaluateExpression(AExpression, ADefault);
end;

function TvgrScriptControl.CheckSyntax(const AScript: string; var AError: TvgrScriptError): Boolean;
var
  AScriptObject: TvgrScriptObject;
begin
  AScriptObject := TvgrScriptObject.Create(Self);
  try
    AScriptObject.SetScriptLanguage(FLanguage);
    Result := FScriptObject.CheckSyntax(AScript, AError);
  finally
    AScriptObject.Free;
  end;
end;

function TvgrScriptControl.GetTextHighlightAttributes(const AText: string; var AAttributes: PWord): Boolean;
var
  AScriptObject: TvgrScriptObject;
begin
  AScriptObject := TvgrScriptObject.Create(Self);
  try
    AScriptObject.SetScriptLanguage(FLanguage);
    Result := FScriptObject.GetTextHighlightAttributes(AText, AAttributes);
  finally
    AScriptObject.Free;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrScriptObject
//
/////////////////////////////////////////////////
constructor TvgrScriptObject.Create(AContext: IvgrScriptContext);
begin
  inherited Create;
  FContext := AContext;
  FLanguage := '';
  FScriptEnterCounter := 0;

  FExpressionEvent := CreateEvent(nil, True, False, nil);
end;

destructor TvgrScriptObject.Destroy;
begin
  Close;
  CloseHandle(FExpressionEvent);
  inherited;
end;

function TvgrScriptObject.GetActiveScript(Value: WideString): IActiveScript;
var
  CLASS_ScriptEngine: TGUID;
begin
  if CLSIDFromProgId(PWideChar(Value), CLASS_ScriptEngine) = S_OK then
    Result := CreateComObject(CLASS_ScriptEngine) as IActiveScript
  else
    Result := nil;
end;

procedure TvgrScriptObject.SetScriptLanguage(const ALanguage: WideString);
var
  AResult: HRESULT;
begin
  FActiveScript := GetActiveScript(ALanguage);
  if FActiveScript <> nil then
  begin
    FLanguage := ALanguage;
    AResult := FActiveScript.SetScriptSite(Self);
    CheckParserCall(AResult);

    AResult := ActiveScriptParse.InitNew;
    CheckParserCall(AResult);

    ActiveScript.AddNamedItem(PWideChar(WideString('_Self')),
                              SCRIPTITEM_ISPERSISTENT or SCRIPTITEM_ISVISIBLE or SCRIPTITEM_GLOBALMEMBERS);

    if AliasManager <> nil then
      FActiveScript.AddNamedItem(PWideChar(WideString(AliasManager.InternalScriptName)),
                                 SCRIPTITEM_ISPERSISTENT or SCRIPTITEM_ISVISIBLE or SCRIPTITEM_GLOBALMEMBERS);
  end
  else
    raise Exception.CreateFmt(sse_ScriptLanguageNotSupported, [ALanguage]);
end;

function TvgrScriptObject.GetErrorDesc(AErrorCode: HRESULT): string;
begin
  if AErrorCode = S_OK then
    Result := ''
  else
    case AErrorCode of
      DISP_E_EXCEPTION: Result := sse_DISP_E_EXCEPTION;
      E_INVALIDARG: Result := sse_E_INVALIDARG;
      E_POINTER: Result := sse_E_POINTER;
      E_NOTIMPL: Result := sse_E_NOTIMPL;
      E_UNEXPECTED: Result := sse_E_UNEXPECTED;
      E_FAIL: Result := sse_E_FAIL;
      S_FALSE: Result := sse_S_FALSE;
      else Result := Format('Code = %d', [Integer(AErrorCode)]);
    end;
end;

procedure TvgrScriptObject.CheckParserCall(AResult: HRESULT; const AScriptError: TvgrScriptError);
var
  S: string;
begin
  if AResult <> S_OK then
  begin
    if AScriptError.ErrorInfo.bstrDescription <> '' then
      S := Format(sse_DISP_E_EXCEPTION_Info, [AScriptError.Line + 1, AScriptError.Position + 1, AScriptError.ErrorInfo.bstrDescription])
    else
      S := sse_None;
    raise Exception.CreateFmt(sse_ScriptError, [GetErrorDesc(AResult), S]);
  end;
end;

procedure TvgrScriptObject.CheckParserCall(AResult: HRESULT);
begin
  if AResult <> S_OK then
    raise Exception.CreateFmt(sse_ScriptError, [GetErrorDesc(AResult), sse_None]);
end;

function TvgrScriptObject.ParseScriptText(const AScript: WideString;
                                          AFlags: Integer;
                                          var ARes: OleVariant;
                                          var AScriptError: TvgrScriptError): HResult;
var
  ASourceContext: DWORD;
  ALine: Cardinal;
  APosition: Integer;
begin
  Result := ActiveScriptParse.ParseScriptText(PWideChar(AScript),
                                              nil,
                                              nil,
                                              nil,
                                              0,
                                              0,
                                              AFlags,
                                              ARes,
                                              AScriptError.ErrorInfo);
                                              
  if Result <> S_OK then
  begin
    if FActiveScriptError <> nil then
    begin
      FActiveScriptError.GetExceptionInfo(AScriptError.ErrorInfo);
      FActiveScriptError.GetSourcePosition(ASourceContext, ALine, APosition);
      AScriptError.Line := ALine;
      AScriptError.Position := APosition;
    end
    else
    begin
      AScriptError.Line := 0;
      AScriptError.Position := 0;
    end;
  end;
end;

procedure TvgrScriptObject.SetScript(const AScript: WideString);
var
  Res: OleVariant;
  AResult: HRESULT;
  AScriptError: TvgrScriptError;
begin
  FActiveScriptError := nil;
  if FActiveScript <> nil then
  begin
    AResult := ParseScriptText(AScript, SCRIPTTEXT_ISVISIBLE or SCRIPTTEXT_ISPERSISTENT,
                               Res, AScriptError);
    CheckParserCall(AResult, AScriptError);

    AResult := FActiveScript.SetScriptState(SCRIPTSTATE_CONNECTED);
    CheckParserCall(AResult);
  end
  else
    raise Exception.Create(sse_NotInitializated);
end;

procedure TvgrScriptObject.Close;
begin
  if FActiveScript <> nil then
    FActiveScript.Close;
  FActiveScript := nil;
  FActiveScriptError := nil;
  FLanguage := '';
end;

procedure TvgrScriptObject.DoErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);
begin
  if FContext <> nil then
    FContext.DoOnErrorInExpression(AExpression, AErrorDescription, AResult);
end;

function TvgrScriptObject.EvaluateExpression(const AExpression: string; const ADefault: Variant): OleVariant;
var
  AResStr: WideString;
  ARes: OleVariant;
  ADescription: string;
  AResult: HRESULT;
  AScriptError: TvgrScriptError;
begin
  if ActiveScriptParse <> nil then
  begin
    Result := ADefault;

    if AliasManager = nil then
      AResStr := AExpression
    else
      AResStr := AliasManager.ProcessAliases(AExpression);

    AResult := ParseScriptText(AResStr, SCRIPTTEXT_ISEXPRESSION, ARes, AScriptError);
    if AResult = S_OK then
      Result := ARes
    else
    begin
      ADescription := AScriptError.ErrorInfo.bstrDescription;
      Result := Format(svgrSCError, [ADescription]);
      DoErrorInExpression(AExpression, ADescription, Result);
    end;
  end
  else
    raise Exception.Create(sse_NotInitializated);
end;

function TvgrScriptObject.RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray): Variant;
var
  I: Integer;
  ATemp: TvgrBooleanDynArray;
begin
  SetLength(ATemp, Length(AParameters));
  for I := 0 to High(ATemp) do
    ATemp[I] := False;
  Result := RunProcedure(AProcName, AParameters, ATemp);
end;

function TvgrScriptObject.RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray): Variant;
var
  AScriptProc: IDispatch;
  ANames: WideString;
  ADispId: Integer;
  ADispParams: TDispParams;
  ARes: OleVariant;
  AExceptionInfo: TExcepInfo;
  I, ASCode: Integer;
  ASourceLine: string;

  ASourceContext, ALineNumber: Cardinal;
  APosition: Integer;

  AByRefParams: TvgrPointerDynArray;
  AOleParams: TvgrOleVariantDynArray;
  AParametersCount: Integer;
  AVariantArg: PVariantArg;

  procedure OleVarToVar(var AVar: Variant; const AOleVar: OleVariant);
  begin
    if PVarData(@AOleVar).VType = varOleStr then
      VarCast(AVar, AOleVar, varString)
    else
      AVar := AOleVar;
  end;

begin
  OleCheck(ActiveScript.GetScriptDispatch('', AScriptProc));
  if AScriptProc <> nil then
  begin
    ANames := WideString(AProcName);
    AScriptProc.GetIDsOfNames(GUID_NULL, @ANames, 1, 0, @ADispId);
    if ADispId > 0 then
    begin
      AParametersCount := Length(AParameters);
      if AParametersCount > 0 then
      begin
        SetLength(AOleParams, AParametersCount);
        SetLength(AByRefParams, AParametersCount);
        for I := 0 to AParametersCount - 1 do
        begin
          if AByRef[I] then
          begin
            AByRefParams[AParametersCount - I - 1] := AllocMem(SizeOf(TVariantArg));
            POleVariant(AByRefParams[AParametersCount - I - 1])^ := AParameters[I];
            with PVariantArg(@AOleParams[AParametersCount - I - 1])^ do
            begin
              vt := VT_VARIANT or VT_BYREF;
              pvarVal := AByRefParams[AParametersCount - I - 1];
            end;
          end
          else
          begin
            AOleParams[AParametersCount - I - 1] := AParameters[I];
          end;
        end;
        ADispParams.rgvarg := @AOleParams[0];
        ADispParams.cArgs := AParametersCount;
      end
      else
      begin
        ADispParams.rgvarg := nil;
        ADispParams.cArgs := 0;
      end;
      ADispParams.rgdispidNamedArgs := nil;
      ADispParams.cNamedArgs := 0;

      try
        AScode := AScriptProc.Invoke(ADispId, GUID_NULL, 0, DISPATCH_METHOD or DISPATCH_PROPERTYGET, ADispParams, @ARes, @AExceptionInfo, nil);
        if AScode = S_OK then
        begin
          for I := 0 to AParametersCount - 1 do
          begin
            if AByRef[I] then
            begin
              AVariantArg := @AOleParams[AParametersCount - 1 - I];
              if (AVariantArg.vt and VT_BYREF) <> 0 then
              begin
                if (AVariantArg.vt and varTypeMask) = varVariant then
                begin
                  OleVarToVar(AParameters[I], AVariantArg.pvarVal^);
                end;
              end;
            end;
          end;
          OleVarToVar(Result, ARes);
        end
        else
        begin
          if FActiveScriptError <> nil then
          begin
            FActiveScriptError.GetExceptionInfo(AExceptionInfo);
            if FActiveScriptError.GetSourcePosition(ASourceContext, ALineNumber, APosition) = S_OK then
              ASourceLine := Format(sse_SourcePosition, [ALineNumber, APosition])
            else
              ASourceLine := sse_None;
            raise Exception.CreateFmt(sse_ProcedureExecuting1, [AProcName,
                                                                GetErrorDesc(AScode),
                                                                ASourceLine,
                                                                AExceptionInfo.bstrDescription]);
          end
          else
            raise Exception.CreateFmt(sse_ProcedureExecuting2, [AProcName, GetErrorDesc(AScode)]);
        end;
        
      finally
        for I := 0 to AParametersCount - 1 do
          if AByRefParams[I] <> nil then
          begin
            POleVariant(AByRefParams[I])^ := Unassigned;
            FreeMem(AByRefParams[I]);
          end;
        Finalize(AByRefParams);
        Finalize(AOleParams);
      end;
    end
    else
      raise Exception.CreateFmt(sse_ProcedureNotFoundInScript, [AProcName]);
  end
  else
    raise Exception.CreateFmt(sse_ProcedureNotFoundInScript, [AProcName]);
end;

function TvgrScriptObject.RunProcedure(const AProcName: string): Variant;
var
  ATemp1: TvgrVariantDynArray;
  ATemp2: TvgrBooleanDynArray;
begin
  SetLength(ATemp1, 0);
  SetLength(ATemp2, 0);
  Result := RunProcedure(AProcName, ATemp1, ATemp2);
end;

function TvgrScriptObject.CheckSyntax(const AScript: WideString; out AScriptError: TvgrScriptError): Boolean;
var
  Res: OleVariant;
  AResult: HRESULT;
begin
  FActiveScriptError := nil;
  if FActiveScript <> nil then
  begin
    AResult := ParseScriptText(AScript, SCRIPTTEXT_ISVISIBLE or SCRIPTTEXT_ISPERSISTENT, Res, AScriptError);
    Result := AResult = S_OK;
  end
  else
    raise Exception.Create(sse_NotInitializated);
end;

function TvgrScriptObject.GetTextHighlightAttributes(const AText: string; var AAttributes: PWord): Boolean;
var
  PWText: PWideChar;

  procedure FreeAttributes;
  begin
    if AAttributes <> nil then
    begin
      FreeMem(AAttributes);
      AAttributes := nil;
    end;
  end;

begin
  Result := False;
  if ActiveScriptDebug = nil then
    FreeAttributes
  else
  begin
    if Length(AText) = 0 then
    begin
      FreeAttributes;
      Result := True;
    end
    else
    begin
      GetMem(PWText,(Length(AText) + 1) * SizeOf(WideChar));
      try
        ReallocMem(AAttributes, Length(AText) * SizeOf(Word));
        try
          StringToWideChar(AText, PWText, Length(AText) * SizeOf(WideChar));
          Result := ActiveScriptDebug.GetScriptTextAttributes(PWText, Length(AText), nil, 0, AAttributes^) = S_OK;
          if not Result then
            FreeAttributes;
        except
          FreeAttributes;
          raise;
        end
      finally
        FreeMem(PWText);
      end;
    end;
  end;
end;

function TvgrScriptObject.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if GetInterface(IID, Obj) then
    Result := S_OK
  else
    Result := E_NOINTERFACE
end;

function TvgrScriptObject._AddRef: Integer;
begin
  Result := S_OK
end;

function TvgrScriptObject._Release: Integer;
begin
  Result := S_OK
end;

function TvgrScriptObject.GetWindow(out phwnd: HWND): HResult;
begin
  if FContext = nil then
    phwnd := 0
  else
    phwnd := FContext.ParentWindow;
  Result := S_OK;
end;

function TvgrScriptObject.EnableModeless(fEnable: BOOL): HResult; stdcall;
begin
  Result := S_OK;
end;

function TvgrScriptObject.GetLCID(out Lcid: TLCID): HRESULT; stdcall;
begin
  Result := E_NOTIMPL;
end;

procedure TvgrScriptObject.DoGetObject(const AObjectName: string; AObject: IDispatch);
begin
  AObject := nil;
  if FContext <> nil then
    FContext.DoOnGetObject(AObjectName, AObject);
end;

function TvgrScriptObject.GetItemInfo(const pstrName: POleStr; dwReturnMask: DWORD; out ppiunkItem: IUnknown; out Info: ITypeInfo): HRESULT; stdcall;
var
  AComponent: IDispatch;
begin
  Result := S_OK;
  if dwReturnMask = SCRIPTINFO_IUNKNOWN then
  begin
    if string(pstrName) = '_Self' then
      ppiunkItem := (Self as IUnknown)
    else
      if string(pstrName) = '_Evaluator' then
        ppiunkItem := FContext
      else
      if (AliasManager <> nil) and (string(pstrName) = AliasManager.InternalScriptName) then
        ppiunkItem := AliasManager
      else
      begin
        DoGetObject(pstrName, AComponent);
        if AComponent <> nil then
          ppiunkItem := (AComponent as IUnknown)
        else
          Result := E_NOINTERFACE;
      end
  end
  else
    Result := E_NOTIMPL;

  if (FScriptEnterCounter > 0) and (Result = E_NOINTERFACE) then
    raise Exception.CreateFmt(sse_ObjectNotFound,[string(pstrName)]);
end;

function TvgrScriptObject.GetDocVersionString(out Version: TBSTR): HRESULT; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TvgrScriptObject.OnScriptTerminate(const pvarResult: OleVariant; const pexcepinfo: TExcepInfo): HRESULT; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TvgrScriptObject.OnStateChange(ScriptState: TScriptState): HRESULT; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TvgrScriptObject.OnScriptError(const pscripterror: IActiveScriptError): HRESULT; stdcall;
begin
  FActiveScriptError := pscripterror;
  Result := S_OK;
end;

function TvgrScriptObject.OnEnterScript: HRESULT; stdcall;
begin
  Inc(FScriptEnterCounter);
  Result := S_OK;
end;

function TvgrScriptObject.OnLeaveScript: HRESULT; stdcall;
begin
  Dec(FScriptEnterCounter);
  Result := S_OK;
end;

function  TvgrScriptObject.GetDocumentContextFromPosition {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_19:0}dwSourceContext: LongWord;
                                                                              {VT_19:0}uCharacterOffset: LongWord;
                                                                              {VT_19:0}uNumChars: LongWord;
                                                                              {VT_29:2}out ppsc: IDebugDocumentContext): HResult; stdcall;

begin
  Result := E_NOTIMPL;
end;

function  TvgrScriptObject.GetApplication {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppda: IDebugApplication): HResult; stdcall;
begin
  Result := E_UNEXPECTED;
end;

function  TvgrScriptObject.GetRootApplicationNode {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdanRoot: IDebugApplicationNode): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function  TvgrScriptObject.OnScriptErrorDebug {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_29:1}const pErrorDebug: IActiveScriptErrorDebug;
                                                                  {VT_3:1}out pfEnterDebugger: Integer;
                                                                  {VT_3:1}out pfCallOnScriptErrorWhenContinuing: Integer): HResult; stdcall;
begin
  pfEnterDebugger := 0;
  Result := S_OK;
end;

function TvgrScriptObject.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := S_OK;
end;

function TvgrScriptObject.GetDispIdOfName(const AName: string): Integer;
begin
  Result := vgrScriptConstants.FindConstant(AName);
  if Result < 0 then
  begin
    Result := vgrScriptFunctions.IndexByName(AName);
    if Result >= 0 then
      Result := Result + vgrScriptConstants.Count
    else
    begin
      if FContext <> nil then
      begin
        Result := FContext.GetIdOfName(AName);
        if Result <> -1 then
          Inc(Result, vgrScriptConstants.Count + vgrScriptFunctions.Count);
      end;
    end;
  end;
end;

function TvgrScriptObject.DoInvoke(DispId: Integer;
                                   Flags: Integer;
                                   var AParameters: TvgrOleVariantDynArray;
                                   var AResult: OleVariant): HResult;
begin
  if DispId < vgrScriptConstants.Count then
  begin
    // global constant
    if (DISPATCH_PROPERTYGET and Flags) = 0 then
      Result := E_UNEXPECTED
    else
    begin
      AResult := vgrScriptConstants.Items[DispId].Value;
      Result := S_OK;
    end;
  end
  else
  begin
    if DispId - vgrScriptConstants.Count < vgrScriptFunctions.Count then
    begin
      // global function
      if (DISPATCH_PROPERTYGET and Flags) = 0 then
        Result := E_UNEXPECTED
      else
      begin
        vgrScriptFunctions[DispId - vgrScriptConstants.Count].Proc(AParameters, AResult);
        Result := S_OK;
      end;
    end
    else
    begin
      // must be processed by FContext
      if FContext <> nil then
        Result := FContext.DoInvoke(DispId - vgrScriptConstants.Count - vgrScriptFunctions.Count,
                                    Flags,
                                    AParameters,
                                    AResult)
      else
        Result := E_UNEXPECTED;
    end;
  end;
end;

// IDispatch
function TvgrScriptObject.GetTypeInfoCount(out Count: Integer): HResult; stdcall;
begin
  Count := 0;
  Result := S_OK;
end;

function TvgrScriptObject.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
begin
  Result := S_OK;
end;

function TvgrScriptObject.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;  stdcall;
begin
  Result := Common_GetIDsOfNames(IID, Names, NameCount, LocaleID, DispIDs, GetDispIdOfName);
end;

function TvgrScriptObject.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
begin
  Result := Common_Invoke(DispID,
                          IID,
                          LocaleID,
                          Flags,
                          Params,
                          VarResult,
                          ExcepInfo,
                          ArgErr,
                          ClassName,
                          DoInvoke,
                          DoCheckScriptInfo);
end;

function TvgrScriptObject.GetAliasManager: TvgrAliasManager;
begin
  if FContext = nil then
    Result := nil
  else
    Result := FContext.AliasManager;
end;

function TvgrScriptObject.GetActiveScriptDebug: IActiveScriptDebug;
var
  ED : IActiveScriptDebug;
begin
  if ActiveScript <> nil then
    if ActiveScript.QueryInterface(IID_IActiveScriptDebug,ED) = S_OK then
      Result := ED
    else
      Result := nil
  else
    Result := nil;
end;

function TvgrScriptObject.GetActiveScriptParse: IActiveScriptParse;
var
  AParse : IActiveScriptParse;
begin
  if ActiveScript <> nil then
    if ActiveScript.QueryInterface(IID_IActiveScriptParse, AParse) = S_OK then
      Result := AParse
    else
      Result := nil
  else
    Result := nil;
end;

procedure TvgrScriptObject.AddObject(const Name: WideString);
begin
  if FActiveScript <> nil then
    FActiveScript.AddNamedItem(PWideChar(Name), SCRIPTITEM_ISPERSISTENT or SCRIPTITEM_ISVISIBLE);
end;

/////////////////////////////////////////////////
//
// TvgrScript
//
/////////////////////////////////////////////////
constructor TvgrScript.Create;
begin
  inherited Create;
  FScript := TStringList.Create;
  FScript.OnChange := OnScriptChange;
end;

destructor TvgrScript.Destroy;
begin
  FreeAndNil(FScriptObject);
  FScript.Free;
  inherited;
end;

procedure TvgrScript.OnScriptChange(Sender: TObject);
begin
  CheckBusy;
  DoChange;
end;

function TvgrScript.GetBusy: Boolean;
begin
  Result := FScriptObject <> nil;
end;

procedure TvgrScript.CheckBusy;
begin
  if Busy then
    raise Exception.Create(sse_ScriptObjectIsBusy);
end;

procedure TvgrScript.CheckReady;
begin
  if not Busy then
    raise Exception.Create(sse_ScriptObjectIsNotInitialized);
end;

procedure TvgrScript.Assign(Source: TPersistent);
begin
  CheckBusy;
  if Source is TvgrScript then
    with TvgrScript(Source) do
    begin
      Self.FLanguage := Language;
      Self.FAliasManager := AliasManager;
      Self.FScript.Assign(Script);
    end;
end;

procedure TvgrScript.SetLanguage(Value: string);
begin
  CheckBusy;
  FLanguage := Value;
end;

function TvgrScript.GetStrings: TStrings;
begin
  Result := FScript;
end;

procedure TvgrScript.SetStrings(Value: TStrings);
begin
  CheckBusy;
  FScript.Assign(Value);
end;

procedure TvgrScript.BeginScriptExecuting;
begin
  FScriptObject := TvgrScriptObject.Create(Self);
  FScriptObject.SetScriptLanguage(FLanguage);
  FScriptObject.SetScript(FScript.Text);
  Cache;
end;

procedure TvgrScript.EndScriptExecuting;
begin
  ClearCache;
  FreeAndNil(FScriptObject);
end;

function TvgrScript.RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray; AByRef: TvgrBooleanDynArray): Variant;
begin
  CheckReady;
  Result := FScriptObject.RunProcedure(AProcName, AParameters, AByRef);
end;

function TvgrScript.RunProcedure(const AProcName: string; var AParameters: TvgrVariantDynArray): Variant;
begin
  CheckReady;
  Result := FScriptObject.RunProcedure(AProcName, AParameters)
end;

function TvgrScript.RunProcedure(const AProcName: string): Variant;
begin
  CheckReady;
  Result := FScriptObject.RunProcedure(AProcName);
end;

function TvgrScript.EvaluateExpression(const AExpression: string; const ADefault: Variant): Variant;
begin
  CheckReady;
  Result := FScriptObject.EvaluateExpression(AExpression, ADefault);
end;

procedure TvgrScript.DoScriptGetObject(const AObjectName: string; var AObject: IDispatch);
begin
  AObject := nil;
  if Assigned(FOnScriptGetObject) then
    FOnScriptGetObject(Root, AObjectName, AObject);
end;

procedure TvgrScript.DoScriptErrorInExpression(const AExpression: string; const ADescription: string; var Value: OleVariant);
begin
  if Assigned(FOnScriptErrorInExpression) then
    FOnScriptErrorInExpression(Root, AExpression, ADescription, Value);
end;

// IvgrScriptContext
function TvgrScript.IvgrScriptContext_GetIdOfName(const AName: string): Integer;
begin
  Result := IndexInStrings(AName, FCache);
end;

function TvgrScript.IvgrScriptContext_DoInvoke(DispId: Integer;
                                               Flags: Integer;
                                               var AParameters: TvgrOleVariantDynArray;
                                               var AResult: OleVariant): HResult;
var
  AInterface: IUnknown;
begin
  if DispId >= FCache.Count then
    Result := E_UNEXPECTED
  else
  begin
    AInterface := GetInterfaceToObject(TObject(FCache.Objects[DispId]));
    if AInterface = nil then
      Result := E_UNEXPECTED
    else
    begin
      AResult := AInterface;
      Result := S_OK;
    end;
  end;
end;

function TvgrScript.GetAliasManager: TvgrAliasManager;
begin
  Result := FAliasManager;
end;

procedure TvgrScript.DoOnGetObject(const AObjectName: string; var AObject: IDispatch);
begin
  DoScriptGetObject(AObjectName, AObject);
end;

procedure TvgrScript.DoOnErrorInExpression(const AExpression, AErrorDescription: string; var AResult: OleVariant);
begin
  DoScriptErrorInExpression(AExpression, AErrorDescription, AResult);
end;

function TvgrScript.GetParentWindow: HWND;
begin
  Result := FParentWindowHandle;
end;

procedure TvgrScript.ReadScriptControl(Reader: TReader);
begin
  Reader.ReadIdent;
end;

procedure TvgrScript.WriteScriptControl(Writer: TWriter);
begin
end;

procedure TvgrScript.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('ScriptControl', ReadScriptControl, WriteScriptControl, False);
end;

function TvgrScript.GetOwner: TPersistent;
begin
  Result := Root;
end;

/////////////////////////////////////////////////
//
// TvgrScriptSyntax
//
/////////////////////////////////////////////////
constructor TvgrScriptSyntax.Create;
begin
  inherited Create;
  FScriptObject := TvgrScriptObject.Create(nil);
end;

destructor TvgrScriptSyntax.Destroy;
begin
  FScriptObject.Free;
  inherited;
end;

function TvgrScriptSyntax.GetLanguage: string;
begin
  Result := FScriptObject.Language;
end;

procedure TvgrScriptSyntax.SetLanguage(Value: string);
begin
  if Value <> FScriptObject.Language then
  begin
    if Value = '' then
      FScriptObject.Close
    else
      FScriptObject.SetScriptLanguage(Value);
    DoChange;
  end;
end;

function TvgrScriptSyntax.CheckSyntax(const AScript: string; var AError: TvgrScriptError): Boolean;
begin
  Result := FScriptObject.CheckSyntax(AScript, AError);
end;

function TvgrScriptSyntax.GetTextHighlightAttributes(const AText: string; var AAttributes: PWord): Boolean;
begin
  Result := FScriptObject.GetTextHighlightAttributes(AText, AAttributes);
end;

initialization

  vgrScriptConstants := TvgrScriptConstants.Create;
  AddGlobalWbConstants;
  vgrScriptFunctions := TvgrScriptGlobalFunctions.Create;
  vgrScriptFunctions.RegisterFunction('ColumnCaption', 1, 1, ColumnCaption);

finalization

  vgrScriptConstants.Free;
  vgrScriptFunctions.Free;

end.
