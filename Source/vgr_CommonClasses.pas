
{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{ Common classes and types definitions unit. }
unit vgr_CommonClasses;

{$I vtk.inc}

//{$DEFINE LISTCHECK}
//{$DEFINE WB_DBG}
//{$DEFINE INRECT_DBG}
//{$DEFINE CheckOuterSearchRule_DBG}
//  {$DEFINE VTK_SCRIPT_DBG}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  DB, Classes, Windows, SysUtils, ComObj, ActiveX, TypInfo,
  vgr_ScriptDispIDs
  {$IFDEF VTK_D6_OR_D7} ,Variants {$ENDIF};

const

  IID_IvgrOwnedPersistentOwner: TGUID = '{53AC2500-8136-43A9-BE17-EE1F8332955E}';
  IID_IvgrObject: TGUID = '{8201E53E-C466-4BFA-9325-F5764673A3BD}';

type

{Dynamic array of integers.}
  TvgrIntegerDynArray = array of Integer;

{Dynamic array of variants.}
  TvgrVariantDynArray = array of Variant;

{Dynamic array of OLE variants.}
  TvgrOleVariantDynArray = array of OleVariant;

{Dynamic array of boolean.}
  TvgrBooleanDynArray = array of Boolean;

{Dynamic array of pointers.}
  TvgrPointerDynArray = array of Pointer;

  TCardinalSet = set of 0..SizeOf(Cardinal) * 8 - 1;
  
{$IFDEF VGR_DEBUG}
  /////////////////////////////////////////////////
  //
  // TvgrDebugInfoItem
  //
  /////////////////////////////////////////////////
  TvgrDebugInfoItem = class(TObject)
  private
    FDesc: string;
    FValue: Variant;
    FComment: string;
  public
    property Desc: string read FDesc write FDesc;
    property Value: Variant read FValue write FValue;
    property Comment: string read FComment write FComment;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrDebugInfo
  //
  /////////////////////////////////////////////////
  TvgrDebugInfo = class(TObject)
  private
    FItems: TList;
    function GetItem(Index: Integer): TvgrDebugInfoItem;
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Add(const ADesc: string; const AValue: Variant; const AComment: string = '');
    procedure Clear;
    property Items[Index: Integer]: TvgrDebugInfoItem read GetItem; default;
    property Count: Integer read GetCount;
  end;
{$ENDIF}

IvgrObject = interface
  ['{8201E53E-C466-4BFA-9325-F5764673A3BD}']
  function GetObjectReference: TObject;
end;

{Specifies the border orientation:
Items:
  vgrboLeft - a border is vertical.
  vgrboTop - a border is horizontal.}
  TvgrBorderOrientation = (vgrboLeft, vgrboTop);

{Specifies a type of change which occured in TvgrWorkbook
Items:
  vgrwcNone                          - there are no changes
  vgrwcNewWorksheet                  - new worksheet is added into workbook
  vgrwcNewCol                        - new column is added into worksheet
  vgrwcNewRow                        - new row is added into worksheet
  vgrwcNewRange                      - new range is added into worksheet
  vgrwcNewBorder                     - new border is added into worksheet
  vgrwcNewHorzSection                - new horizontal section is added into worksheet
  vgrwcNewVertSection                - new vertical section is added into worksheet
  vgrwcDeleteWorksheet               - worksheet is deleted
  vgrwcDeleteCol                     - column is deleted
  vgrwcDeleteRow                     - row is deleted
  vgrwcDeleteRange                   - range is deleted
  vgrwcDeleteBorder                  - border is deleted
  vgrwcDeleteHorzSection             - horizontal section is deleted
  vgrwcDeleteVertSection             - vertical section is deleted
  vgrwcChangeCol                     - column is changed
  vgrwcChangeRow                     - row is changed
  vgrwcChangeRange                   - range is changed
  vgrwcChangeBorder                  - border is changed
  vgrwcChangeWorksheet               - one of worksheet's properties is changed, also this type of change is used to notify
  vgrwcChangeWorksheetContent        - large changes in worksheet have occured: columns are deleted, rows are deleted and so on
  vgrwcChangeHorzSection             - horizontal section is changed
  vgrwcChangeVertSection             - vertical section is changed
  vgrwcChangeWorksheetPageProperties - page properties of worksheet is changed
  vgrwcUpdateAll                     - the occured changes concern set of objects }
  TvgrWorkbookChangesType =
    (vgrwcNone,
     vgrwcNewWorksheet,
     vgrwcNewCol,
     vgrwcNewRow,
     vgrwcNewRange,
     vgrwcNewBorder,
     vgrwcNewHorzSection,
     vgrwcNewVertSection,
     vgrwcDeleteWorksheet,
     vgrwcDeleteCol,
     vgrwcDeleteRow,
     vgrwcDeleteRange,
     vgrwcDeleteBorder,
     vgrwcDeleteHorzSection,
     vgrwcDeleteVertSection,
     vgrwcChangeCol,
     vgrwcChangeRow,
     vgrwcChangeRange,
     vgrwcChangeBorder,
     vgrwcChangeWorksheet,
     vgrwcChangeHorzSection,
     vgrwcChangeVertSection,
     vgrwcChangeWorksheetPageProperties,
     vgrwcChangeReportTemplate,
     vgrwcChangeWorksheetContent,
     vgrwcUpdateAll);

{Holds the information about change in workbook.}
  TvgrWorkbookChangeInfo = record
{Type of change.
See also:
  TvgrWorkbookChangesType}
    ChangesType: TvgrWorkbookChangesType;
{Changed object. This field is not nil if ChangesType in
(vgrwcNewWorksheet, vgrwcDeleteWorksheet, vgrwcChangeWorksheet, vgrwcChangeWorksheetPageProperties)}
    ChangedObject: TObject;
{Changed interface. This field is not nil if ChangesType in
(vgrwcNewCol, vgrwcNewRow, vgrwcNewRange, vgrwcNewBorder, vgrwcNewHorzSection, vgrwcNewVertSection,
vgrwcDeleteCol, vgrwcDeleteRow, vgrwcDeleteRange, vgrwcDeleteBorder, vgrwcDeleteHorzSection,
vgrwcDeleteVertSection, vgrwcChangeCol, vgrwcChangeRow, vgrwcChangeRange, vgrwcChangeBorder,
vgrwcChangeHorzSection, vgrwcChangeVertSection)}
    ChangedInterface: IUnknown;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrCustomScriptObject
  //
  /////////////////////////////////////////////////
{The internal abstract class, that allows to use in the scripts any objects.
Instance of this object contains the reference on the main object
and realizes the IDispatch interface for them.}
  TvgrCustomScriptObject = class(TInterfacedObject, IDispatch, IvgrObject)
  protected
    // IvgrObject
    function GetObjectReference: TObject; virtual;
    // IDispatch
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer;
                         out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID;
                           Names: Pointer;
                           NameCount, LocaleID: Integer;
                           DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer;
                    const IID: TGUID;
                    LocaleID: Integer;
                    Flags: Word;
                    var Params;
                    VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;


    function GetDispIdOfName(const AName: String) : Integer; virtual;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; virtual;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; virtual;
  public
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPersistentScriptObject
  //
  /////////////////////////////////////////////////
{The internal class, that implements the IDispatch interface for the any TPersistent object.
Due to this class you can use the published properties of the TPersistent classes in the scripts.}
  TvgrPersistentScriptObject = class(TvgrCustomScriptObject)
  private
    FObject: TPersistent;
  protected
    function GetObjectReference: TObject; override;
    function GetDispIdOfName(const AName: String): Integer; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;

    property MainObject: TPersistent read FObject;
  public
{Creates an instance of the TvgrPersistentScriptObject class.
Parameters:
  AObject - The main TPersistent object.}
    constructor Create(AObject: TPersistent);
  end;

  /////////////////////////////////////////////////
  //
  // TvgrDataSetScriptObject
  //
  /////////////////////////////////////////////////
{The internal class, that implements the IDispatch interface for the any TDataSet object.
Due to this class you can use the published properties of the TDataSet classes in the scripts.
See also:
  TvgrPersistentScriptObject}
  TvgrDataSetScriptObject = class(TvgrPersistentScriptObject)
  private
    function GetDataSet: TDataSet;
  protected
    function GetDispIdOfName(const AName: String) : Integer; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;

    property DataSet: TDataSet read GetDataSet;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPersistent
  //
  /////////////////////////////////////////////////
{Implements a persistent with OnChange event, also TvgrPersistent implements IUnknown and IDispatch interfaces.}
  TvgrPersistent = class(TPersistent, IUnknown, IDispatch, IvgrObject)
  private
    FOnChange: TNotifyEvent;
    FUpdateCount: Integer;
    function GetEnableUpdate: Boolean;
  protected
    // IvgrObject
    function GetObjectReference: TObject; virtual;

    // IUnknown
    function QueryInterface(const IID: TGUID; out Obj): HResult; virtual; stdcall;
    function _AddRef: Integer; virtual; stdcall;
    function _Release: Integer; virtual; stdcall;

    // IDispatch
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual;  stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    function GetDispIdOfName(const AName: String): Integer; virtual; 
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; virtual;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; virtual;

    // Support of OnChange event
    procedure DoChange;
    procedure BeginUpdate;
    procedure EndUpdate;
    property EnableUpdate: Boolean read GetEnableUpdate;
  public
{Creates an instance of the TvgrPersistent class.}
    constructor Create(AOnChange: TNotifyEvent); overload; 
{Frees an instance of the TvgrPersistent class.}
    constructor Create; overload; virtual;
{Occurs when any property is changed.}
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBaseComponent
  //
  /////////////////////////////////////////////////
  TvgrBaseComponent = class(TComponent, IvgrObject {$IFNDEF VTK_D6_OR_D7}, IUnknown {$ENDIF})
  protected
    // IvgrObject
    function GetObjectReference: TObject; virtual;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrComponent
  //
  /////////////////////////////////////////////////
{TvgrComponent implements an IDispath interface.
The majority of components in the GridReport library are derived from this class. }
  TvgrComponent = class(TvgrBaseComponent, IDispatch)
  protected
    // IDispatch
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual;  stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    function GetDispIdOfName(const AName: String) : Integer; virtual; 
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; virtual;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; virtual;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrObjectList
  //
  /////////////////////////////////////////////////
{Use TvgrObjectList to store and maintain a list of objects.
TvgrObjectList provides properties and methods to add, delete, access, and sort objects.
If the OwnsObjects property is set to true (the default),
TvgrObjectList controls the memory of its objects, freeing an object
when it is removed from the list with the Delete, Remove, or Clear method;
or when the TObjectList instance is itself destroyed.}
  TvgrObjectList = class(TObject)
  private
    FOwnsObjects: Boolean;
    FList: TList;
    function GetItem(Index: Integer): TObject;
    function GetCount: Integer;
  public
{Creates an instance of the TvgrObjectList class.
Parameters:
  AOwnsObjects - Allows TObjectList to free objects when they are deleted
from the list or the list is destroyed.}
    constructor Create(AOwnsObjects: Boolean = True);
{Frees an instance of the TvgrObjectList class.}
    destructor Destroy; override;
    { Call Clear to empty the Items array and set the Count to 0. }
    procedure Clear;
    { Call Add to insert an object at the end of the list.
      Add places the object after the last item,
      even if the array contains nil (Delphi) or NULL (C++) references,
Parameters:
  AObject - TObject
Return value:
  returns the index of the inserted object.
Comment:
  <br>(The first object in the list has an index of 0.) }
    function Add(AObject: TObject): Integer;
    { Call Remove to delete a specific object from the list when its index is unknown.
      <br>If OwnsObjects is true, Remove frees the object in addition to removing it from the list.
      After an object is deleted, all the objects that follow it are moved up in index position
      and Count is decremented.
      <br>If an object appears more than once on the list, Remove deletes only the first appearance.
      Hence, if OwnsObjects is true, removing an object that appears more than once results in
      empty object references later in the list.
      <br>To use an index position (rather than an object reference) to specify the object to be
      removed, call Delete.
Parameters:
  AObject - TObject}
    procedure Remove(AObject: TObject);
    { Call Delete to remove the item at a specific position from the list.
      Calling Delete moves up all items in the Items array that follow the deleted item,
      and reduces the Count.
Parameters:
      The index is zero-based, so the first item has an Index value of 0, the second item has an Index value of 1, and so on.}
    procedure Delete(Index: Integer);
    { Use Items to access objects in the list.
      <br>Items is a zero-based array: The first object is indexed as 0, the second object is indexed as 1,
      and so forth.
      <br>You can read or change the value at a specific index, or use Items with the Count property
      to iterate through the list. }
    property Items[Index: Integer]: TObject read GetItem; default;
    { Read Count to determine the number of entries in the Items array. }
    property Count: Integer read GetCount;
  end;
  
  TArrayWord = array [0..16383] of Word;
  {$EXTERNALSYM TArrayWord}
  PArrayWord = ^TArrayWord;
  {$EXTERNALSYM PArrayWord}

  TArrayTPoint = array [0..16383] of TPoint;
  {$EXTERNALSYM TArrayTPoint}
  PArrayTPoint = ^TArrayTPoint;
  {$EXTERNALSYM PArrayTPoint}

  /////////////////////////////////////////////////
  //
  // TvgrListItem
  //
  /////////////////////////////////////////////////
  { Internal class. }
  TvgrListItem = class(TvgrPersistent)
  protected
    function GetIndexField1 : Integer; virtual; abstract;
    procedure SetIndexField1(Value : Integer); virtual; abstract;
    function GetIndexField2 : Integer; virtual;
    procedure SetIndexField2(Value : Integer); virtual;
    function GetIndexField3 : Integer; virtual;
    procedure SetIndexField3(Value : Integer); virtual;
    function GetSizeField1 : Integer; virtual;
    function GetSizeField2 : Integer; virtual;
    procedure SetSizeField1(Value : Integer); virtual;
    procedure SetSizeField2(Value : Integer); virtual;
  public
    function InRect(ARect: TRect): Boolean;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrListFieldStatistic
  //
  /////////////////////////////////////////////////
  { Internal class. }
  TvgrListFieldStatistic = class
  private
    FStatisticArray : array of Integer;
    FMax : Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Add(ASize: Integer);
    procedure Delete(ASize: Integer);
    property Max: Integer read FMax;
  end;

(*
  TvgrFindListItemCallBackProc = procedure(AItem : TvgrListItem) of object;
  /////////////////////////////////////////////////
  //
  // TvgrList
  //
  /////////////////////////////////////////////////
  TvgrList = class(TList)
  private
    FListFieldStatistic1 : TvgrListFieldStatistic;
    FListFieldStatistic2 : TvgrListFieldStatistic;
    function RectFromItem(AItem : TvgrListItem) : TRect;
    procedure CalcRangeOfSearching(Idx1,Idx2,Size1,Size2 : Integer; out LowLev1,LowLev2,HiLev1,HiLev2 : Integer; DeleteIfCrossOver : Boolean);
    procedure SetLowRulerSpeedy(LowLev1,LowLev2 : Integer; var LowRuler : Integer);
    function CheckSearchRule(Index,LowLev1,LowLev2 : Integer) : boolean;
    function CheckOuterSearchRule(Index,HiLev1,HiLev2 : Integer) : boolean;
    procedure CorrectLowRuler(Item : TvgrListItem; ANumber1,ANumber2,CurPos : Integer; var LowRuler : Integer);
  protected
    procedure InternalDelete(Index : Integer);
    procedure InternalAdd(Item : TvgrListItem; LowRuler : Integer);
    function InternalFind(
      ANumber1, ANumber2, ANumber3 : Integer;
      ASizeField1, ASizeField2 : Integer;
      DeleteIfCrossOver : boolean;
      AutoCreateItem : boolean
      ) : TvgrListItem;
    procedure InternalFindItemsInRect(ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc; DeleteItem : boolean);
    function CreateNewItem(ANumber1, ANumber2, ANumber3 : Integer) : TvgrListItem; virtual;
    procedure BeforeAddItem; virtual;  abstract;
    procedure AfterAddItem(AItem : TvgrListItem); virtual; abstract;
    procedure BeforeDeleteItem(AItem : TvgrListItem); virtual;  abstract;
    procedure AfterDeleteItem; virtual; abstract;
    property ListFieldStatistic1 : TvgrListFieldStatistic read FListFieldStatistic1;
    property ListFieldStatistic2 : TvgrListFieldStatistic read FListFieldStatistic2;
  public
    constructor Create;
    procedure Clear; override;
    destructor Destroy; override;
    procedure Delete(Index : Integer);
    procedure FindAndCallback(const ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc);
    procedure DeleteItemsFromRect(ARect : TRect);
    procedure Remove(AItem : TvgrListItem);
  end;
*)

TvgrInvokeHandler = function(DispId: Integer; Flags: Integer; var AParameters: TvgrOleVariantDynArray; var AResult: OleVariant): HResult of object;
TvgrGetIDsOfNamesHandler = function(const AName: string): Integer of object;
TvgrCheckScriptInfoHandler = function (DispId: Integer; Flags: Integer; AParametersCount: Integer): HResult of object;

function Common_GetIDsOfNames(const IID: TGUID;
                              Names: Pointer;
                              NameCount, LocaleID: Integer;
                              DispIDs: Pointer;
                              Handler: TvgrGetIDsOfNamesHandler): HResult;

function CheckScriptInfo(ADispId: Integer;
                         Flags: Integer;
                         AParametersCount: Integer;
                         AScriptInfo: pvgrScriptInfo;
                         AScriptInfoLength: Integer): HResult;

function Common_Invoke(DispID: Integer;
                       const IID: TGUID;
                       LocaleID: Integer;
                       Flags: Word;
                       var Params; VarResult, ExcepInfo, ArgErr: Pointer;
                       const AClassName: string;
                       Handler: TvgrInvokeHandler;
                       CheckHandler: TvgrCheckScriptInfoHandler): HResult;

function Persistent_SetRTTIProperty(APersistent: TPersistent;
                                    APropInfo: PPropInfo;
                                    const AValue: OleVariant): HResult;

function Persistent_GetRTTIProperty(APersistent: TPersistent;
                                    APropInfo: PPropInfo;
                                    var AValue: OleVariant): HResult;

function Persistent_GetDispIdOfName(APersistent: TPersistent;
                                    const AName: String) : Integer;

function Persistent_DoInvoke(APersistent: TPersistent;
                             DispId: Integer;
                             Flags: Integer;
                             var AParameters: TvgrOleVariantDynArray;
                             var AResult: OleVariant): HResult;

function GetInterfaceToObject(AObject: TObject): IUnknown;

function OleVariantToObject(const AOleVariant: OleVariant): TObject;

function GetIntValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Integer;
function GetUIntValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Cardinal;
function GetStrValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): string;
function GetVarValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Variant;
function GetBoolValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Boolean;
function PointInCut(x, x1, x2: Integer) : boolean;

implementation

{$R ..\res\vgr_CommonStrings.res}

uses
  vgr_Functions;

function Common_GetIDsOfNames(const IID: TGUID;
                              Names: Pointer;
                              NameCount, LocaleID: Integer;
                              DispIDs: Pointer;
                              Handler: TvgrGetIDsOfNamesHandler): HResult;
type
  PNamesArray = ^TNamesArray;
  TNamesArray = array[0..0] of PWideChar;

  PDispIdsArray = ^TDispIdsArray;
  TDispIdsArray = array[0..0] of Integer;
var
  PropName: string;
  PropInfo: Integer;
begin
  PropName := PNamesArray(Names)[0];
  PropInfo := Handler(PropName);
  if PropInfo >= 0 then
  begin
    PDispIdsArray(DispIds)[0] := PropInfo;
    Result := S_OK;
  end
  else
    Result := DISP_E_UNKNOWNNAME;
end;

function CheckScriptInfo(ADispId: Integer;
                         Flags: Integer;
                         AParametersCount: Integer;
                         AScriptInfo: pvgrScriptInfo;
                         AScriptInfoLength: Integer): HResult;
var
  I: Integer;
begin
  for I := 0 to AScriptInfoLength - 1 do
    if AScriptInfo^[I].DispId = ADispId then
    begin
      with AScriptInfo^[I] do
      begin
        if ((GetMin = GetDisabled) or (GetMax = GetDisabled)) and
           ((Flags and DISPATCH_PROPERTYGET <> 0) or (Flags and DISPATCH_METHOD <> 0)) then
        begin
          Result := E_UNEXPECTED;
          exit;
        end;
        if ((SetMin = GetDisabled) or (SetMax = GetDisabled)) and
           (Flags and DISPATCH_PROPERTYPUT <> 0) then
        begin
          Result := E_UNEXPECTED;
          exit;
        end;
        if (Flags and DISPATCH_PROPERTYGET <> 0) or (Flags and DISPATCH_METHOD <> 0) then
        begin
          if (GetMin <> AnyParams) and (GetMin > AParametersCount) then
          begin
            Result := DISP_E_BADPARAMCOUNT;
            exit;
          end;
          if (GetMax <> AnyParams) and (GetMax < AParametersCount) then
          begin
            Result := DISP_E_BADPARAMCOUNT;
            exit;
          end;
        end;
        if Flags and DISPATCH_PROPERTYPUT <> 0 then
        begin
          if (SetMin <> AnyParams) and (SetMin > AParametersCount) then
          begin
            Result := DISP_E_BADPARAMCOUNT;
            exit;
          end;
          if (SetMax <> AnyParams) and (SetMax < AParametersCount) then
          begin
            Result := DISP_E_BADPARAMCOUNT;
            exit;
          end;
        end;
      end;
    end;
  Result := S_OK;
end;

function Common_Invoke(DispID: Integer;
                       const IID: TGUID;
                       LocaleID: Integer;
                       Flags: Word;
                       var Params; VarResult, ExcepInfo, ArgErr: Pointer;
                       const AClassName: string;
                       Handler: TvgrInvokeHandler;
                       CheckHandler: TvgrCheckScriptInfoHandler): HResult;
var
  dps: TDispParams absolute Params;
  WS: WideString;
  I: Integer;
  AParameters: TvgrOleVariantDynArray;
  APropertyGet: Boolean;
  AVarResult: OleVariant;
begin
  APropertyGet := (Flags and DISPATCH_PROPERTYGET) <> 0;
  if APropertyGet and (VarResult = nil) then
  begin
    Result := E_UNEXPECTED;
    exit; 
  end;

  Result := CheckHandler(DispId, Flags, dps.cArgs);
  if Result <> S_OK then
    exit;

  // Prepare params
  SetLength(AParameters, dps.cArgs);
  for I := 0 to dps.cArgs - 1 do
    AParameters[I] := OleVariant(dps.rgvarg^[dps.cArgs - I - 1]);

  try
    try
      Result := Handler(DispId, Flags, AParameters, AVarResult);
      if VarResult <> nil then
        POleVariant(VarResult)^ := AVarResult;

    except
      on E: Exception do
        begin
          if ExcepInfo <> nil then
          begin
            FillChar(ExcepInfo^, SizeOf(TExcepInfo), 0);
            TExcepInfo(ExcepInfo^).wCode := 1001;
            TExcepInfo(ExcepInfo^).BStrSource := AClassName;
            WS := E.Message;
            TExcepInfo(ExcepInfo^).bstrDescription := SysAllocString(PWideChar(WS));
          end;
          Result := DISP_E_EXCEPTION;
        end;
    end;
  finally
  end;
end;

function Persistent_GetRTTIProperty(APersistent: TPersistent;
                                    APropInfo: PPropInfo;
                                    var AValue: OleVariant): HResult;
var
  AObject: TPersistent;
begin
  Result := S_OK;
  case APropInfo.PropType^.Kind of
    tkInteger:
      AValue := GetOrdProp(APersistent, APropInfo);
    tkString, tkLString, tkWChar, tkWString:
      AValue := GetStrProp(APersistent, APropInfo);
    tkEnumeration:
      AValue := GetEnumName(APropInfo^.PropType^, GetOrdProp(APersistent, APropInfo));
{$IFDEF VTK_D6_OR_D7}
    tkInt64:
      AValue := GetInt64Prop(APersistent, APropInfo);
{$ENDIF}
    tkFloat:
      AValue := GetFloatProp(APersistent, APropInfo);
    tkVariant:
      AValue := GetVariantProp(APersistent, APropInfo);
    tkClass:
      begin
        AObject := TPersistent(GetOrdProp(APersistent, APropInfo));
        AValue := TvgrPersistentScriptObject.Create(AObject) as IUnknown;
      end;
    else
      begin
        AValue := UnAssigned;
        Result := E_UNEXPECTED;
      end;
  end;
end;

function Persistent_SetRTTIProperty(APersistent: TPersistent;
                                    APropInfo: PPropInfo;
                                    const AValue: OleVariant): HResult;
{$IFDEF VTK_D6_OR_D7}
var
  AInt64: Int64;
{$ENDIF}
begin
  Result := S_OK;
  case APropInfo.PropType^.Kind of
    tkInteger:
      SetOrdProp(APersistent, APropInfo, AValue);
    tkString, tkLString, tkWChar, tkWString:
      SetStrProp(APersistent, APropInfo, VarToStr(AValue));
    tkEnumeration:
      SetOrdProp(APersistent, APropInfo, GetEnumValue(APropInfo^.PropType^, AValue));
{$IFDEF VTK_D6_OR_D7}
    tkInt64:
      begin
        AInt64 := AValue;
        SetInt64Prop(APersistent, APropInfo, AInt64);
      end;
{$ENDIF}
    tkFloat:
      SetFloatProp(APersistent, APropInfo, AValue);
    tkVariant:
      SetVariantProp(APersistent, APropInfo, AValue);
    else
      begin
        Result := E_UNEXPECTED;
      end;
  end;
end;

function Persistent_GetDispIdOfName(APersistent: TPersistent;
                                    const AName: String): Integer;
var
  APropInfo: PPropInfo;
  AComponent: TComponent;
begin
  Result := -1;
  APropInfo := GetPropInfo(APersistent.ClassInfo, AName);
  if APropInfo = nil then
  begin
    if APersistent is TComponent then
    begin
      AComponent := TComponent(APersistent).FindComponent(AName);
      if AComponent <> nil then
        Result := AComponent.ComponentIndex;
    end;
  end
  else
    Result := Integer(APropInfo);
end;

function Persistent_DoInvoke(APersistent: TPersistent;
                             DispId: Integer;
                             Flags: Integer;
                             var AParameters: TvgrOleVariantDynArray;
                             var AResult: OleVariant): HResult;
var
  AComponent: TComponent;
begin
  Result := S_OK;
  if Flags and DISPATCH_PROPERTYPUT = 0then
  begin
    if APersistent is TComponent then
    begin
      AComponent := TComponent(APersistent);
      if DispId < AComponent.ComponentCount then
      begin
        AComponent := AComponent.Components[DispId];
        if AComponent is TDataset then
          AResult := TvgrDataSetScriptObject.Create(TDataset(AComponent)) as IUnknown
        else
          AResult := TvgrPersistentScriptObject.Create(AComponent) as IUnknown;
        exit;
      end;
    end;
    Result := Persistent_GetRTTIProperty(APersistent, PPropInfo(DispID), AResult);
  end
  else
  begin
    if APersistent is TComponent then
    begin
     AComponent := TComponent(APersistent);
     if DispId < AComponent.ComponentCount then
     begin
       Result := E_UNEXPECTED;
       exit;
     end;
    end;

    // set value of RTTI property
    Result := Persistent_SetRTTIProperty(APersistent, PPropInfo(DispId), AParameters[0]);
  end;
end;

function GetInterfaceToObject(AObject: TObject): IUnknown;
begin
  if AObject is TvgrComponent then
    Result := TvgrComponent(AObject) as IUnknown
  else
  if AObject is TvgrPersistent then
    Result := TvgrPersistent(AObject) as IUnknown
  else
  if AObject is TDataSet then
    Result := TvgrDataSetScriptObject.Create(TDataSet(AObject)) as IUnknown
  else
  if AObject is TPersistent then
    Result := TvgrPersistentScriptObject.Create(TPersistent(AObject)) as IUnknown
  else
  if AObject is TAutoObject then
    Result := TAutoObject(AObject) as IUnknown
  else
    Result := nil;
end;

function OleVariantToObject(const AOleVariant: OleVariant): TObject;
var
  AObjectInterface: IvgrObject;
begin
  with PVarData(@AOleVariant)^ do
  begin
    if (VType <> varDispatch) or not Supports(IDispatch(VDispatch), IID_IvgrObject, AObjectInterface) then
      if (VType <> varUnknown) or not Supports(IUnknown(VUnknown), IID_IvgrObject, AObjectInterface) then
        AObjectInterface := nil;

    if AObjectInterface <> nil then
      Result := AObjectInterface.GetObjectReference
    else
      raise Exception.Create('OlVariant can not be converted to object.');
  end;
end;

function GetIntValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Integer;
{
var
  Arg : TVariantArg;
  VT: Word;
  ByRef: Boolean;
}
begin
  Result := VarAsType(PVariant(@dps.rgvarg^[pDispIds^[Index]])^, varInteger);
{
  Arg := dps.rgvarg^[pDispIds^[Index]];
  VT := Arg.vt;
  ByRef := (VT and VT_BYREF) = VT_BYREF;

  if ByRef then
  begin
    VT := VT and (not VT_BYREF);
    case VT of
      VT_I1: Result := Arg.pbVal^;
      VT_I2: Result := Arg.piVal^;
      VT_I4: Result := Arg.piVal^;
      VT_VARIANT: Result := Arg.pvarVal^;
      else Result := 0;
    end;
  end
  else
    case VT of
      VT_I1: Result := Arg.bVal;
      VT_I2: Result := Arg.iVal;
      VT_I4: Result := Arg.iVal;
      else Result := 0;
    end;
}
end;

function GetBoolValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Boolean;
{
var
  Arg : TVariantArg;
  VT: Word;
  ByRef: Boolean;
}
begin
  Result := VarAsType(PVariant(@dps.rgvarg^[pDispIds^[Index]])^, varBoolean);
{
  Arg := dps.rgvarg^[pDispIds^[Index]];
  VT := Arg.vt;
  ByRef := (VT and VT_BYREF) = VT_BYREF;
  if ByRef then begin
   VT := VT and (not VT_BYREF);
   case VT of
     VT_I1: Result := Boolean(Arg.pbVal^);
     VT_I2: Result := Boolean(Arg.piVal^);
     VT_BOOL: Result := Boolean(Arg.piVal^);
     VT_VARIANT: Result := Boolean(Arg.pvarVal^);
   else
     Result := Boolean(Arg.plVal^);
   end;
  end else
  case VT of
   VT_I1: Result := Boolean(Arg.bVal);
   VT_I2: Result := Boolean(Arg.iVal);
  else
   Result := Boolean(Arg.lVal);
  end;
}
end;


function GetStrValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): string;
{
var
  Arg : TVariantArg;
  VT: Word;
  ByRef: Boolean;
}
begin
  Result := VarAsType(PVariant(@dps.rgvarg^[pDispIds^[Index]])^, varString);
{
  Arg := dps.rgvarg^[pDispIds^[Index]];
  VT := Arg.vt;
  ByRef := (VT and VT_BYREF) = VT_BYREF;
  if ByRef then begin
   VT := VT and (not VT_BYREF);
   case VT of
     VT_VARIANT: Result := Arg.pvarVal^;
     VT_BSTR:    Result := Arg.bstrVal^;
   end;
  end else
  case VT of
     VT_NULL : Result := '';
     VT_I2, VT_I4 : Result := IntToStr(Arg.lVal);
     VT_R4, VT_R8 : Result := FloatToStr(Arg.dblVal);
     VT_VARIANT: Result := Arg.bstrVal;
     VT_BSTR:    Result := Arg.bstrVal;
     VT_DISPATCH: Result := WideCharToString(Arg.bstrVal);
     VT_DATE : Result := DateTimeToStr(Arg.date);
  else
    Result := 'Unknown';
  end;
}
end;

function GetUIntValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Cardinal;
{
var
  Arg : TVariantArg;
  VT: Word;
  ByRef: Boolean;
}
begin
{$IFDEF VTK_D6_OR_D7}
  Result := VarAsType(PVariant(@dps.rgvarg^[pDispIds^[Index]])^, varLongWord);
{$ELSE}
  Result := VarAsType(PVariant(@dps.rgvarg^[pDispIds^[Index]])^, varInteger);
{$ENDIF}
{
  Arg := dps.rgvarg^[pDispIds^[Index]];
  VT := Arg.vt;
  ByRef := (VT and VT_BYREF) = VT_BYREF;
  if ByRef then begin
   VT := VT and (not VT_BYREF);
   case VT of
     VT_I1: Result := Arg.pbVal^;
     VT_I2: Result := Arg.puiVal^;
     VT_I4: Result := Arg.puiVal^;
     VT_VARIANT: Result := Arg.pvarVal^;
   else
     Result := 0;
   end;
  end else
  case VT of
   VT_I1: Result := Arg.bVal;
   VT_I2: Result := Arg.uiVal;
   VT_I4: Result := Arg.ulVal;
  else
   Result := 0;
  end;
}
end;

function GetVarValue(Index: Integer; const dps: TDispParams; pDispIds: PDispIdList): Variant;
{
var
  Arg: PVariantArg;
  VT: Word;
  ByRef: Boolean;
}
begin
//  Result := VarAsType(PVariant(@dps.rgvarg^[pDispIds^[Index]])^, varVariant);
  Result := PVariant(@dps.rgvarg^[pDispIds^[Index]])^;
  if (PVarData(@Result).VType and varTypeMask) = varOleStr then
    Result := VarAsType(Result, varString); 
{
  Arg := dps.rgvarg^[pDispIds^[Index]];
  VT := Arg.vt;
  ByRef := (VT and VT_BYREF) = VT_BYREF;
  if ByRef then begin
   VT := VT and (not VT_BYREF);
   case VT of
     VT_VARIANT: Result := Arg.pvarVal^;
     VT_BSTR:    Result := Arg.bstrVal^;
   end;
  end else
  case VT of
     VT_NULL : Result := Unassigned;
     VT_I2, VT_I4 : Result := Arg.lVal;
     VT_R4, VT_R8 : Result := Arg.dblVal;
     VT_VARIANT: Result := WideCharToString(Arg.bstrVal);
     VT_BSTR:    Result := WideCharToString(Arg.bstrVal);
     VT_DISPATCH: Result := WideCharToString(Arg.bstrVal);
     VT_DATE : Result := Arg.date;
  else
    Result := 'Unknown';
  end;
}
end;

function PointInCut(x,x1,x2 : Integer) : boolean;
begin
  Result := (x1 <= x) and (x <= x2);
end;

{$IFDEF VGR_DEBUG}
/////////////////////////////////////////////////
//
// TvgrDebugInfoItem
//
/////////////////////////////////////////////////

/////////////////////////////////////////////////
//
// TvgrDebugInfo
//
/////////////////////////////////////////////////
constructor TvgrDebugInfo.Create;
begin
  inherited Create;
  FItems := TList.Create;
end;

destructor TvgrDebugInfo.Destroy;
begin
  Clear;
  inherited;
end;

function TvgrDebugInfo.GetCount: Integer;
begin
  Result := FItems.Count;
end;

function TvgrDebugInfo.GetItem(Index: Integer): TvgrDebugInfoItem;
begin
  Result := TvgrDebugInfoItem(FItems[Index]);
end;

procedure TvgrDebugInfo.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  FItems.Clear;
end;

procedure TvgrDebugInfo.Add(const ADesc: string; const AValue: Variant; const AComment: string = '');
var
  AItem: TvgrDebugInfoItem;
begin
  AItem := TvgrDebugInfoItem.Create;
  with AItem do
  begin
    Desc := ADesc;
    Value := AValue;
    Comment := AComment;
  end;
  FItems.Add(AItem);
end;
{$ENDIF}

/////////////////////////////////////////////////
//
// TvgrCustomScriptObject
//
/////////////////////////////////////////////////
destructor TvgrCustomScriptObject.Destroy;
begin
  inherited;
end;

function TvgrCustomScriptObject.GetObjectReference: TObject;
begin
  Result := nil; 
end;

function TvgrCustomScriptObject.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Count := 0;
  Result := S_OK;
end;

function TvgrCustomScriptObject.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TvgrCustomScriptObject.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
begin
  Result := Common_GetIDsOfNames(IID, Names, NameCount, LocaleID, DispIDs, GetDispIdOfName);
end;

function TvgrCustomScriptObject.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;  stdcall;
{
var
  dps : TDispParams absolute Params;
  HasParams : boolean;
  pDispIds : PDispIdList;
  iDispIdsSize : integer;
  WS: WideString;
  I: Integer;
}
begin
  Result := Common_Invoke(DispId,
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
(*  
  {$IFDEF VTK_SCRIPT_DBG}
  OutputDebugString(PChar(Format('%s %d',[ClassName, DispId])));
  {$ENDIF}
  pDispIds := NIL;
  HasParams := (dps.cArgs > 0);
  if HasParams then
  begin
    iDispIdsSize := dps.cArgs * SizeOf(TDispId);
    GetMem(pDispIds, iDispIdsSize);
  end;
  try
    if HasParams then
      for I := 0 to dps.cArgs - 1 do
        pDispIds^[I] := dps.cArgs - 1 - I;
    try
      Result := DoInvoke(DispId, IID, LocaleID, Flags, dps, pDispIds, VarResult, ExcepInfo, ArgErr);
    except
      on E: EInvalidParamCount do Result := DISP_E_BADPARAMCOUNT;
      on E: EInvalidParamType do Result := DISP_E_BADVARTYPE;
      on E: Exception do
        begin
          if Assigned(ExcepInfo) then
          begin
            FillChar(ExcepInfo^, SizeOf(TExcepInfo), 0);
            TExcepInfo(ExcepInfo^).wCode := 1001;
            TExcepInfo(ExcepInfo^).BStrSource := ClassName;
            WS := E.Message;
            TExcepInfo(ExcepInfo^).bstrDescription := SysAllocString(PWideChar(WS));
          end;
          Result := DISP_E_EXCEPTION;
        end;
    end;
  finally
    if HasParams then
      FreeMem(pDispIds);
  end;
*)
end;

function TvgrCustomScriptObject.GetDispIdOfName(const AName : String) : Integer;
begin
  Result := -1;
end;

function TvgrCustomScriptObject.DoInvoke(DispId: Integer;
                  Flags: Integer;
                  var AParameters: TvgrOleVariantDynArray;
                  var AResult: OleVariant): HResult;
begin
  Result := S_OK;
end;

function TvgrCustomScriptObject.DoCheckScriptInfo(DispId: Integer;
                            Flags: Integer;
                            AParametersCount: Integer): HResult;
begin
  Result := S_OK;
end;

/////////////////////////////////////////////////
//
// TvgrPersistentScriptObject
//
/////////////////////////////////////////////////
constructor TvgrPersistentScriptObject.Create(AObject: TPersistent);
begin
  inherited Create;
  FObject := AObject;
end;

function TvgrPersistentScriptObject.GetObjectReference: TObject;
begin
  Result := FObject;
end;

function TvgrPersistentScriptObject.GetDispIdOfName(const AName: String) : Integer;
begin
  Result := Persistent_GetDispIdOfName(FObject, AName);
end;

function TvgrPersistentScriptObject.DoInvoke(DispId: Integer;
                                             Flags: Integer;
                                             var AParameters: TvgrOleVariantDynArray;
                                             var AResult: OleVariant): HResult;
begin
  Result := Persistent_DoInvoke(FObject,
                                DispId,
                                Flags,
                                AParameters,
                                AResult);
end;

function TvgrPersistentScriptObject.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := S_OK;
end;

/////////////////////////////////////////////////
//
// TvgrDataSetScriptObject
//
/////////////////////////////////////////////////
function TvgrDataSetScriptObject.GetDataSet: TDataSet;
begin
  Result := TDataSet(MainObject);
end;

function TvgrDataSetScriptObject.GetDispIdOfName(const AName: String) : Integer;
var
  AField: TField;
begin
  if AnsiCompareText(AName,'Bof') = 0 then
    Result := cs_DatasetBof
  else
  if AnsiCompareText(AName,'Eof') = 0 then
    Result := cs_DatasetEof
  else
  if AnsiCompareText(AName,'First') = 0 then
    Result := cs_DatasetFirst
  else
  if AnsiCompareText(AName,'Last') = 0 then
    Result := cs_DatasetLast
  else
  if AnsiCompareText(AName,'Next') = 0 then
    Result := cs_DatasetNext
  else
  if AnsiCompareText(AName,'Prior') = 0 then
    Result := cs_DatasetPrior
  else
  if AnsiCompareText(AName,'Field') = 0 then
    Result := cs_DatasetField
  else
  if AnsiCompareText(AName,'Locate') = 0 then
    Result := cs_DatasetLocate
  else
  if AnsiCompareText(AName, 'IifLocate') = 0 then
    Result := cs_DatasetIifLocate
  else
  if AnsiCompareText(AName, 'Open') = 0 then
    Result := cs_DatasetOpen
  else
  begin
    AField := DataSet.FindField(AName);
    if AField <> nil then
      Result := AField.Index + cs_DatasetFieldsStart
    else
      Result := inherited GetDispIdOfName(AName);
  end;
end;

function TvgrDataSetScriptObject.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siDataset, siDatasetLength);
end;

function TvgrDataSetScriptObject.DoInvoke(DispId: Integer;
                                          Flags: Integer;
                                          var AParameters: TvgrOleVariantDynArray;
                                          var AResult: OleVariant): HResult;
var
  ALocateOptions: TLocateOptions;
  ALocateValues: Variant;
  I: Integer;
  AValue1, AValue2: Variant;
  AStartEdit: Boolean;
  AField: TField;
begin
  if (DispId < cs_PersistentMax) or (DispId > cs_PersistentRTTIPropsStart) then
    Result := inherited DoInvoke(DispId, Flags, AParameters, AResult)
  else
  begin
    case DispId of
      cs_DatasetBof: AResult := DataSet.Bof;
      cs_DatasetEof: AResult := DataSet.Eof;
      cs_DatasetFirst: DataSet.First;
      cs_DatasetLast: DataSet.Last;
      cs_DatasetNext: DataSet.Next;
      cs_DatasetPrior: DataSet.Prior;
      cs_DatasetField:
        if Flags and DISPATCH_PROPERTYPUT = 0 then
          AResult := DataSet.FieldByName(AParameters[0]).AsVariant
        else
        begin
          AStartEdit := DataSet.State = dsBrowse;
          if AStartEdit then
            DataSet.Edit;
          DataSet.FieldByName(AParameters[0]).AsVariant := AParameters[1];
          if AStartEdit then
            DataSet.Post;
        end;
      cs_DatasetLocate:
        begin
          ALocateOptions := [];
          if AParameters[1] then
            Include(ALocateOptions, loCaseInsensitive);
          if AParameters[2] then
            Include(ALocateOptions, loPartialKey);
          ALocateValues := VarArrayCreate([0, Length(AParameters) - 4], varVariant);
          for I := 0 to Length(AParameters) - 4 do
            ALocateValues[I] := AParameters[I + 3];
          AResult := DataSet.Locate(AParameters[0], ALocateValues, ALocateOptions);
        end;
      cs_DatasetIifLocate:
        begin
          ALocateOptions := [];
          AValue1 := AParameters[0];
          AValue2 := AParameters[1];
          
          if AParameters[3] then
            Include(ALocateOptions, loCaseInsensitive);
            
          if AParameters[4] then
            Include(ALocateOptions, loPartialKey);

          ALocateValues := VarArrayCreate([0, Length(AParameters) - 6], varVariant);
          for I := 0 to Length(AParameters) - 6 do
            ALocateValues[I] := AParameters[I + 5];
          if DataSet.Locate(AParameters[2], ALocateValues, ALocateOptions) then
            AResult := AValue1
          else
            AResult := AValue2;
        end;
      cs_DatasetOpen: DataSet.Open;
      else
        begin
          if (DispId - cs_DatasetFieldsStart >= 0) and (DispId - cs_DatasetFieldsStart < DataSet.FieldCount) then
          begin
            AField := DataSet.Fields[DispId - cs_DatasetFieldsStart];
            if Flags and DISPATCH_PROPERTYPUT = 0 then
              AResult := AField.AsVariant
            else
            begin
              AStartEdit := DataSet.State = dsBrowse;
              if AStartEdit then
                DataSet.Edit;
              AField.AsVariant := AParameters[1];
              if AStartEdit then
                DataSet.Post;
            end;
          end;
        end;
    end;
    Result := S_OK;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrPersistent
//
/////////////////////////////////////////////////
constructor TvgrPersistent.Create(AOnChange: TNotifyEvent);
begin
  Create;
  FOnChange := AOnChange;
end;

constructor TvgrPersistent.Create;
begin
  inherited Create;
end;

function TvgrPersistent.GetObjectReference: TObject;
begin
  Result := Self;
end;

procedure TvgrPersistent.DoChange;
begin
  if Assigned(FOnChange) then
    FOnChange(Self);
end;

procedure TvgrPersistent.BeginUpdate;
begin
  Inc(FUpdateCount);
end;

procedure TvgrPersistent.EndUpdate;
begin
  Dec(FUpdateCount);
end;

function TvgrPersistent.GetEnableUpdate: Boolean;
begin
  Result := FUpdateCount = 0;
end;

// IUnknown
function TvgrPersistent.QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
begin
  if GetInterface(IID, Obj) then
    Result := S_OK
  else
    Result := E_NOINTERFACE
end;

function TvgrPersistent._AddRef: Integer; stdcall;
begin
  Result := S_OK
end;

function TvgrPersistent._Release: Integer; stdcall;
begin
  Result := S_OK
end;

// IDispatch
function TvgrPersistent.GetTypeInfoCount(out Count: Integer): HResult; stdcall;
begin
  Count := 0;
  Result := S_OK;
end;

function TvgrPersistent.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
begin
  Result := S_OK;
end;

function TvgrPersistent.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
{
type
  PNamesArray = ^TNamesArray;
  TNamesArray = array[0..0] of PWideChar;

  PDispIdsArray = ^TDispIdsArray;
  TDispIdsArray = array[0..0] of Integer;
var
  PropName: String;
  PropInfo: Integer;
}
begin
  Result := Common_GetIDsOfNames(IID, Names, NameCount, LocaleID, DispIDs, GetDispIdOfName);
{
  PropName := PNamesArray(Names)[0];
  PropInfo := GetDispIdOfName(PropName);
  if PropInfo > 0 then
  begin
    PDispIdsArray(DispIds)[0] := PropInfo;
    Result := S_OK;
  end
  else
    Result := DISP_E_UNKNOWNNAME;
}
end;

function TvgrPersistent.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
{
var
  dps : TDispParams absolute Params;
  HasParams : boolean;
  pDispIds : PDispIdList;
  iDispIdsSize : integer;
  WS: WideString;
  I: Integer;
}
begin
  Result := Common_Invoke(DispId,
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
(*
  pDispIds := NIL;
  HasParams := (dps.cArgs > 0);
  if HasParams then
  begin
    iDispIdsSize := dps.cArgs * SizeOf(TDispId);
    GetMem(pDispIds, iDispIdsSize);
  end;
  try
    if HasParams then
      for I := 0 to dps.cArgs - 1 do
        pDispIds^[I] := dps.cArgs - 1 - I;
    try
      Result := DoInvoke(DispId, IID, LocaleID, Flags, dps, pDispIds, VarResult, ExcepInfo, ArgErr);
    except
      on E: EInvalidParamCount do Result := DISP_E_BADPARAMCOUNT;
      on E: EInvalidParamType do  Result := DISP_E_BADVARTYPE;
      on E: Exception do begin
        if Assigned(ExcepInfo) then begin
          FillChar(ExcepInfo^, SizeOf(TExcepInfo), 0);
          TExcepInfo(ExcepInfo^).wCode := 1001;
          TExcepInfo(ExcepInfo^).BStrSource := ClassName;
          WS := E.Message;
          TExcepInfo(ExcepInfo^).bstrDescription := SysAllocString(PWideChar(WS));
        end;
        Result := DISP_E_EXCEPTION;
      end;
    end;
  finally
    if HasParams then
      FreeMem(pDispIds);
  end;
*)
end;

function TvgrPersistent.GetDispIdOfName(const AName: String): Integer;
begin
  Result := Persistent_GetDispIdOfName(Self, AName);
end;

function TvgrPersistent.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := S_OK;
end;

function TvgrPersistent.DoInvoke(DispId: Integer;
                                 Flags: Integer;
                                 var AParameters: TvgrOleVariantDynArray;
                                 var AResult: OleVariant): HResult;
begin
  Result := Persistent_DoInvoke(Self,
                                DispId,
                                Flags,
                                AParameters,
                                AResult);
end;

/////////////////////////////////////////////////
//
// TvgrBaseComponent
//
/////////////////////////////////////////////////
function TvgrBaseComponent.GetObjectReference: TObject;
begin
  Result := Self;
end;

/////////////////////////////////////////////////
//
// TvgrComponent
//
/////////////////////////////////////////////////
// IDispatch
function TvgrComponent.GetTypeInfoCount(out Count: Integer): HResult; stdcall;
begin
  Count := 0;
  Result := S_OK;
end;

function TvgrComponent.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
begin
  Result := S_OK;
end;

function TvgrComponent.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
{
type
  PNamesArray = ^TNamesArray;
  TNamesArray = array[0..0] of PWideChar;
  PDispIdsArray = ^TDispIdsArray;
  TDispIdsArray = array[0..0] of Integer;
var
  PropName: String;
  PropInfo: Integer;
}
begin
  Result := Common_GetIDsOfNames(IID,
                                 Names,
                                 NameCount,
                                 LocaleId,
                                 DispIDs,
                                 GetDispIdOfName);
{
  PropName := PNamesArray(Names)[0];
  PropInfo := GetDispIdOfName(PropName);
  if PropInfo > 0 then
  begin
    PDispIdsArray(DispIds)[0] := PropInfo;
    Result := S_OK;
  end
  else
    Result := DISP_E_UNKNOWNNAME;
}
end;

function TvgrComponent.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
{
var
  dps : TDispParams absolute Params;
  HasParams : boolean;
  pDispIds : PDispIdList;
  iDispIdsSize : integer;
  WS: WideString;
  I: Integer;
}
begin
  Result := Common_Invoke(DispId,
                          IID,
                          LocaleID,
                          Flags,
                          Params,
                          VarResult,
                          ExcepInfo,
                          ArgErr,
                          ClassName,
                          DoInvoke,
                          DoCheckScriptInfo)
(*
  {$IFDEF VTK_SCRIPT_DBG}
  OutputDebugString(PChar(Format('%s %d',[ClassName, DispId])));
  {$ENDIF}
  pDispIds := NIL;
  HasParams := (dps.cArgs > 0);
  if HasParams then
  begin
    iDispIdsSize := dps.cArgs * SizeOf(TDispId);
    GetMem(pDispIds, iDispIdsSize);
  end;
  try
    if HasParams then
      for I := 0 to dps.cArgs - 1 do
        pDispIds^[I] := dps.cArgs - 1 - I;
    try
      Result := DoInvoke(DispId, IID, LocaleID, Flags, dps, pDispIds, VarResult, ExcepInfo, ArgErr);
    except
      on E: EInvalidParamCount do Result := DISP_E_BADPARAMCOUNT;
      on E: EInvalidParamType do Result := DISP_E_BADVARTYPE;
      on E: Exception do
        begin
          if Assigned(ExcepInfo) then
          begin
            FillChar(ExcepInfo^, SizeOf(TExcepInfo), 0);
            TExcepInfo(ExcepInfo^).wCode := 1001;
            TExcepInfo(ExcepInfo^).BStrSource := ClassName;
            WS := E.Message;
            TExcepInfo(ExcepInfo^).bstrDescription := SysAllocString(PWideChar(WS));
          end;
          Result := DISP_E_EXCEPTION;
        end;
    end;
  finally
    if HasParams then
      FreeMem(pDispIds);
  end;
*)
end;

function TvgrComponent.GetDispIdOfName(const AName: String) : Integer;
begin
  Result := Persistent_GetDispIdOfName(Self, AName);
end;

function TvgrComponent.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := S_OK;
end;

function TvgrComponent.DoInvoke(DispId: Integer;
                                Flags: Integer;
                                var AParameters: TvgrOleVariantDynArray;
                                var AResult: OleVariant): HResult;
begin
  Result := Persistent_DoInvoke(Self, DispId, Flags, AParameters, AResult);
end;

/////////////////////////////////////////////////
//
// TvgrObjectList
//
/////////////////////////////////////////////////
constructor TvgrObjectList.Create(AOwnsObjects: Boolean = True);
begin
  inherited Create;
  FOwnsObjects := AOwnsObjects;
  FList := TList.Create;
end;

destructor TvgrObjectList.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TvgrObjectList.Clear;
begin
  if FOwnsObjects then
    FreeListItems(FList);
  FList.Clear;
end;

function TvgrObjectList.Add(AObject: TObject): Integer;
begin
  Result := FList.Add(AObject);
end;

procedure TvgrObjectList.Remove(AObject: TObject);
var
  I: Integer;
begin
  I := FList.IndexOf(AObject);
  if I <> -1 then
    Delete(I);
end;

procedure TvgrObjectList.Delete(Index: Integer);
begin
  if FOwnsObjects then
    Items[Index].Free;
  FList.Delete(Index);
end;

function TvgrObjectList.GetItem(Index: Integer): TObject;
begin
  Result := TObject(FList[Index]);
end;

function TvgrObjectList.GetCount: Integer;
begin
  Result := FList.Count;
end;

/////////////////////////////////////////////////
//
// TvgrListItem
//
/////////////////////////////////////////////////
function TvgrListItem.InRect(ARect: TRect): Boolean;
var
  RectItem : TRect;
  R1, R2 : TRect;
begin
  R1 := ARect;
  Inc(R1.Right);
  Inc(R1.Bottom);
  RectItem := Rect(GetIndexField2,GetIndexField1,GetIndexField2 + GetSizeField2 - 1, GetIndexField1 + GetSizeField1 - 1);
  R2 := RectItem;
  Inc(R2.Right);
  Inc(R2.Bottom);
  Result := (PointInCut(RectItem.Left,ARect.Left,ARect.Right)
    or PointInCut(RectItem.Right,ARect.Left,ARect.Right)
    or PointInCut(ARect.Left,RectItem.Left,RectItem.Right))
    and
    (PointInCut(RectItem.Top,ARect.Top,ARect.Bottom)
    or PointInCut(RectItem.Bottom,ARect.Top,ARect.Bottom)
    or PointInCut(ARect.Top,RectItem.Top,RectItem.Bottom));
end;

function TvgrListItem.GetIndexField2 : Integer;
begin
  Result := 0;
end;

procedure TvgrListItem.SetIndexField2(Value : Integer);
begin
end;

function TvgrListItem.GetIndexField3 : Integer;
begin
  Result := 0;
end;

procedure TvgrListItem.SetIndexField3(Value : Integer);
begin
end;

function TvgrListItem.GetSizeField1 : Integer;
begin
  Result := 1;
end;

function TvgrListItem.GetSizeField2 : Integer;
begin
  Result := 1;
end;

procedure TvgrListItem.SetSizeField1(Value : Integer);
begin
end;

procedure TvgrListItem.SetSizeField2(Value : Integer);
begin
end;

/////////////////////////////////////////////////
//
// TvgrListFieldStatistic
//
/////////////////////////////////////////////////
constructor TvgrListFieldStatistic.Create;
begin
  SetLength(FStatisticArray,1);
  FMax := 1;
end;

destructor TvgrListFieldStatistic.Destroy;
begin
  FStatisticArray := nil;
end;

procedure TvgrListFieldStatistic.Add(ASize : Integer);
begin
  if ASize > High(FStatisticArray) then
    SetLength(FStatisticArray,ASize);
  Inc(FStatisticArray[ASize-1]);
  if FMax < ASize then
    FMax := ASize;
end;

procedure TvgrListFieldStatistic.Delete(ASize : Integer);
var
  NewMax : Integer;
begin
  Dec(FStatisticArray[ASize - 1]);
  if (FStatisticArray[ASize - 1] = 0) and (FMax = ASize) then
  begin
    NewMax := (FMax - 1);
    while (NewMax > 1) and (FStatisticArray[NewMax - 1] = 0) do
      Dec(NewMax);
    if NewMax > 1 then
      FMax := NewMax
    else
      FMax := 1;
  end;
end;

(*
/////////////////////////////////////////////////
//
// TvgrList
//
/////////////////////////////////////////////////
constructor TvgrList.Create;
begin
  FListFieldStatistic1 := TvgrListFieldStatistic.Create;
  FListFieldStatistic2 := TvgrListFieldStatistic.Create;
end;

destructor TvgrList.Destroy;
begin
  Clear;
  FListFieldStatistic1.Free;
  FListFieldStatistic2.Free;
  inherited;
end;

procedure TvgrList.Delete(Index : Integer);
begin
  InternalDelete(Index);
end;

procedure TvgrList.FindAndCallback(const ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc);
begin
  InternalFindItemsInRect(ARect,CallBackProc,False);
end;

procedure TvgrList.Clear;
var
  i : Integer;
begin
  for i := 0 to Count - 1 do
  begin
    TObject(Items[i]).Free;
  end;
  inherited;
end;

// protected
procedure TvgrList.InternalDelete(Index : Integer);
begin
  BeforeDeleteItem(TvgrListItem(Items[Index]));
  FListFieldStatistic1.Delete(TvgrListItem(Items[Index]).GetSizeField1);
  FListFieldStatistic2.Delete(TvgrListItem(Items[Index]).GetSizeField2);
  TObject(Items[Index]).Free;
  inherited Delete(Index);
  AfterDeleteItem;
end;

procedure TvgrList.InternalAdd(Item : TvgrListItem; LowRuler : Integer);
begin
  FListFieldStatistic1.Add(Item.GetSizeField1);
  FListFieldStatistic2.Add(Item.GetSizeField2);
  if LowRuler >= Count then
    Add(Item)
  else
    Insert(LowRuler, Item);
end;

function TvgrList.InternalFind(
  ANumber1, ANumber2, ANumber3 : Integer;
  ASizeField1, ASizeField2 : Integer;
  DeleteIfCrossOver : boolean;
  AutoCreateItem : boolean
  ) : TvgrListItem;
var
  i : Integer;
  Item : TvgrListItem;
  FLowLev1 : Integer;
  FLowLev2 : Integer;
  FHiLev1 : Integer;
  FHiLev2 : Integer;
  LowRuler : Integer;
begin
  Result := nil;
  CalcRangeOfSearching(ANumber1,ANumber2,ASizeField1,ASizeField2,FLowLev1,FLowLev2,FHiLev1,FHiLev2, DeleteIfCrossOver);
  SetLowRulerSpeedy(FLowLev1,FLowLev2,LowRuler);
  i := LowRuler;
  while CheckOuterSearchRule(i,FHiLev1,FHiLev2) do
  begin
    Item := TvgrListItem(Items[i]);
    CorrectLowRuler(Item,ANumber1,ANumber2,i, LowRuler);
    if (ANumber1 = Item.GetIndexField1) and
       (ANumber2 = Item.GetIndexField2) and
       (ANumber3 = Item.GetIndexField3) and
       (ASizeField1 = Item.GetSizeField1) and
       (ASizeField2 = Item.GetSizeField2) then
    begin
      Result := Item;
      if not DeleteIfCrossOver then
        Break;
    end
    else
      if (DeleteIfCrossOver) and (ANumber3 = Item.GetIndexField3) and Item.InRect(Rect(ANumber2, ANumber1, ANumber2 + ASizeField2 - 1, ANumber1 + ASizeField1 - 1)) then
      begin
        InternalDelete(i);
        Dec(i);
      end;
    Inc(i);
  end;

  if AutoCreateItem and (Result = nil) then
  begin
    BeforeAddItem;
    Result := CreateNewItem(ANumber1,ANumber2,ANumber3);
    with Result do
    begin
      SetSizeField1(ASizeField1);
      SetSizeField2(ASizeField2);
    end;
    InternalAdd(Result,LowRuler);
    AfterAddItem(Result);
  end;
end;

procedure TvgrList.InternalFindItemsInRect(ARect : TRect; CallBackProc : TvgrFindListItemCallBackProc; DeleteItem : boolean);
var
  i : Integer;
  Item : TvgrListItem;
  FLowLev1 : Integer;
  FLowLev2 : Integer;
  FHiLev1 : Integer;
  FHiLev2 : Integer;
  LowRuler : Integer;
begin
  FLowLev1 := ARect.Top;
  FHiLev1  := ARect.Bottom;
  FLowLev2 := ARect.Left;
  FHiLev2  := ARect.Right;

  CalcRangeOfSearching(FLowLev1,FLowLev2,FHiLev1 - FLowLev1 + 1,FHiLev2 - FLowLev2 + 1,FLowLev1,FLowLev2,FHiLev1,FHiLev2, True);

  SetLowRulerSpeedy(FLowLev1,FLowLev2,LowRuler);
  i := LowRuler;
  while CheckOuterSearchRule(i,FHiLev1,FHiLev2) do
  begin
    Item := TvgrListItem(Items[i]);
    if Item.InRect(ARect) then
      if DeleteItem then
      begin
        InternalDelete(i);
        Dec(i);
      end
      else
        if Assigned(CallBackProc) then
          CallBackProc(Item);
    Inc(i);
  end;
end;

procedure TvgrList.DeleteItemsFromRect(ARect : TRect);
begin
  InternalFindItemsInRect(ARect,nil,True);
end;

procedure TvgrList.Remove(AItem : TvgrListItem);
begin
  InternalFindItemsInRect(RectFromItem(AItem),nil,True);
end;

function TvgrList.CreateNewItem(ANumber1, ANumber2, ANumber3 : Integer) : TvgrListItem;
begin
  Result := nil; // This method must be overrided;
end;

function TvgrList.RectFromItem(AItem : TvgrListItem) : TRect;
begin
  Result := Rect(AItem.GetIndexField2, AItem.GetIndexField1, AItem.GetIndexField2 + AItem.GetSizeField2 - 1, AItem.GetIndexField1 + AItem.GetSizeField1 - 1);
end;

procedure TvgrList.CalcRangeOfSearching(Idx1,Idx2,Size1,Size2 : Integer; out LowLev1,LowLev2,HiLev1,HiLev2 : Integer; DeleteIfCrossOver : Boolean);
begin
  LowLev1 := Idx1;
  LowLev2 := Idx2;
  HiLev1  := Idx1;
  HiLev2  := Idx2;
  if DeleteIfCrossOver then
  begin
    Dec(LowLev1,FListFieldStatistic1.Max + Size1 - 2);
    Dec(LowLev2,FListFieldStatistic2.Max + Size2 - 2);
    Inc(HiLev1,FListFieldStatistic1.Max + Size1 - 2);
    Inc(HiLev2,FListFieldStatistic2.Max + Size2 - 2);
  end;
end;

procedure TvgrList.SetLowRulerSpeedy(LowLev1,LowLev2 : Integer; var LowRuler : Integer);
var
  LowMarker, HiMarker, MiddleMarker : Integer;
begin
  if (Count = 0) then
    LowRuler := 0
  else
  if (Count = 1) then
  begin
    if  CheckSearchRule(0,LowLev1,LowLev2) then
      LowRuler := Count
    else
      LowRuler := 0;
  end
  else
  if CheckSearchRule(Count-1,LowLev1,LowLev2) then
      LowRuler := Count
  else
  begin
    LowMarker := 0;
    HiMarker := Count - 1;
    repeat
      MiddleMarker:=(LowMarker + HiMarker) div 2;
      if CheckSearchRule(MiddleMarker,LowLev1,LowLev2) then
        LowMarker := MiddleMarker + 1
      else
        HiMarker := MiddleMarker;
    until LowMarker = HiMarker;
    LowRuler := LowMarker;
  end;
end;

function TvgrList.CheckSearchRule(Index,LowLev1,LowLev2 : Integer) : boolean;
var
  Item : TvgrListItem;
begin
  Item := TvgrListItem(Items[Index]);
  Result := (Item.GetIndexField1 < LowLev1) or
            ((Item.GetIndexField1 = LowLev1) and
             (Item.GetIndexField2 < LowLev2));
end;

function TvgrList.CheckOuterSearchRule(Index,HiLev1,HiLev2 : Integer) : boolean;
begin
  if Index < Count then
    with TvgrListItem(Items[Index]) do
      Result := (GetIndexField1 < HiLev1) or
                ((GetIndexField1 = HiLev1) and
                 (GetIndexField2 <= HiLev2))
  else
    Result := False;
end;

procedure TvgrList.CorrectLowRuler(Item : TvgrListItem; ANumber1,ANumber2,CurPos : Integer; var LowRuler : Integer);
begin
  if (Item.GetIndexField1 < ANumber1) or
      ((Item.GetIndexField1 = ANumber1) and
       (Item.GetIndexField2 <= ANumber2)) then
    LowRuler := CurPos;
end;
*)

end.
