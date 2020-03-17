{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains definition of @link(TvgrFormLocalizer) class.
@link(TvgrFormLocalizer) is a non visual component, which can replace contents
of the components properties with values stored in resources.
See also:
  TvgrFormLocalizer, vgr_Localize}
unit vgr_FormLocalizer;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Classes, Forms, TypInfo, SysUtils,

  vgr_StringIDs;

type
  /////////////////////////////////////////////////
  //
  // TvgrFormLocalizerItem
  //
  /////////////////////////////////////////////////
{Describes link between the string in resources and component property.
Only string properties can be used.}
  TvgrFormLocalizerItem = class(TCollectionItem)
  private
    FComponent: TComponent;
    FPropName: string;
    FResStringID: Integer;
    procedure SetComponent(Value: TComponent);
  protected
    function GetPropInfo: PPropInfo;
    function IsStrProp: Boolean;
    procedure SetStrValue(const AValue: string);
  public
    constructor Create(ACollection: TCollection); override;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
  published
{Component, which property should be localized.}
    property Component: TComponent read FComponent write SetComponent;
{Name of the property of the component, which should be localized.}
    property PropName: string read FPropName write FPropName;
{ID of the resource string, which will be loaded, this property
defines offset of resource string from position defined in
TvgrFormLocalizer.BaseResIndex property. For example, if TvgrFormLocalizer.BaseResIndex = 40000
and ResStringID = 101 when actual ID of resource string is 40101.}
    property ResStringID: Integer read FResStringID write FResStringID default -1;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormLocalizerItems
  //
  /////////////////////////////////////////////////
{TvgrFormLocalizerItems holds the @link(TvgrFormLocalizerItem_ objects.
@link(TvgrFormLocalizerItem) object represents link between the string in
resources and component property.}
  TvgrFormLocalizerItems = class(TOwnedCollection)
  private
    function GetItem(Index: Integer): TvgrFormLocalizerItem;
    procedure SetItem(Index: Integer; Value: TvgrFormLocalizerItem);
  protected
    function IndexByResStringID(AResStringID: Integer): Integer;
    function IndexByComponent(AComponent: TComponent): Integer;
    function FindByComponent(AComponent: TComponent): TvgrFormLocalizerItem;
  public
{Call Add to create an TvgrFormLocalizerItem item in the collection.
The new item is placed at the end of the Items array.
Return value:
  Returns the new collection item.}
    function Add: TvgrFormLocalizerItem; reintroduce;
{Finds item by ResStringID property
Parameters:
  AResStringID - resource ID of string which should be found.
Return value:
  Returns a found TvgrFormLocalizerItem item.}
    function FindByResStringID(AResStringID: Integer): TvgrFormLocalizerItem;
{Use Items to access individual items in the collection.
It represents the position of the item in the collection.
Parameters:
  Index - The value of the Index parameter corresponds to the Index property of TvgrFormLocalizerItem.}
    property Items[Index: Integer]: TvgrFormLocalizerItem read GetItem write SetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormLocalizer
  //
  /////////////////////////////////////////////////
{TvgrFormLocalizer is a non visual component, which can replace contents of the components properties with values stored in resources.
TvgrFormLocalizer contains definitions of links
between the component properties and ID of resource strings.}
  TvgrFormLocalizer = class(TComponent)
  private
    FItems: TvgrFormLocalizerItems;
    FBaseResIndex: Integer;
    procedure SetItems(Value: TvgrFormLocalizerItems);
    function IsItemsStored: Boolean;
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
{Use DoLocalize to replace contents of all properties defined in Items property. If property
defined in item not found no exceptions or error messages are occures.
If resource string not exists value of the property not change.}
    procedure DoLocalize;
  published
{Lists the localize items, that represent a link between property of the component and
resource string.}
    property Items: TvgrFormLocalizerItems read FItems write SetItems stored IsItemsStored;
{Specifies the origin of the resources strings. For example, if BaseResIndex = 40000
and TvgrFormLocalizerItem.ResStringID = 101 than actual ID of resource string is 40101.}
    property BaseResIndex: Integer read FBaseResIndex write FBaseResIndex default svgrResStringsBase;
  end;

implementation

uses
  vgr_Functions, vgr_Localize;

/////////////////////////////////////////////////
//
// TvgrFormLocalizerItem
//
/////////////////////////////////////////////////
constructor TvgrFormLocalizerItem.Create(ACollection: TCollection);
begin
  inherited;
  FResStringID := -1;
end;

destructor TvgrFormLocalizerItem.Destroy;
begin
  inherited;
end;

procedure TvgrFormLocalizerItem.Assign(Source: TPersistent);
begin
  if Source is TvgrFormLocalizerItem then
    with TvgrFormLocalizerItem(Source) do
    begin
      Self.FComponent := FComponent;
      Self.FPropName := FPropName;
      Self.FResStringID := FResStringID;
    end;
end;

procedure TvgrFormLocalizerItem.SetComponent(Value: TComponent);
begin
  FComponent := Value;
end;

function TvgrFormLocalizerItem.GetPropInfo: PPropInfo;
begin
  if Component <> nil then
    Result := TypInfo.GetPropInfo(Component.ClassInfo, PropName)
  else
    Result := nil; 
end;

function TvgrFormLocalizerItem.IsStrProp: Boolean;
var
  APropInfo: PPropInfo;
begin
  APropInfo := GetPropInfo;
  Result := (APropInfo <> nil) and (APropInfo.PropType^.Kind in [tkString,tkLString]);
end;

procedure TvgrFormLocalizerItem.SetStrValue(const AValue: string);
var
  APropInfo: PPropInfo;
begin
  if Component <> nil then
  begin
    APropInfo := TypInfo.GetPropInfo(Component.ClassInfo, PropName);
    if APropInfo <> nil then
      SetStrProp(Component, APropInfo, AValue);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrFormLocalizerItems
//
/////////////////////////////////////////////////
function TvgrFormLocalizerItems.GetItem(Index: Integer): TvgrFormLocalizerItem;
begin
  Result := TvgrFormLocalizerItem(inherited Items[Index]);
end;

procedure TvgrFormLocalizerItems.SetItem(Index: Integer; Value: TvgrFormLocalizerItem);
begin
  inherited Items[Index] := Value;
end;

function TvgrFormLocalizerItems.IndexByResStringID(AResStringID: Integer): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].ResStringID <> AResStringID) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrFormLocalizerItems.Add: TvgrFormLocalizerItem;
begin
  Result := TvgrFormLocalizerItem(inherited Add);
end;

function TvgrFormLocalizerItems.FindByResStringID(AResStringID: Integer): TvgrFormLocalizerItem;
var
  I: Integer;
begin
  I := IndexByResStringID(AResStringID);
  if I = -1 then
    Result := nil
  else
    Result := Items[I];
end;

function TvgrFormLocalizerItems.IndexByComponent(AComponent: TComponent): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].Component <> AComponent) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrFormLocalizerItems.FindByComponent(AComponent: TComponent): TvgrFormLocalizerItem;
var
  I: Integer;
begin
  I := IndexByComponent(AComponent);
  if I = -1 then
    Result := nil
  else
    Result := Items[I];
end;

/////////////////////////////////////////////////
//
// TvgrFormLocalizer
//
/////////////////////////////////////////////////
constructor TvgrFormLocalizer.Create(AOwner: TComponent);
begin
  inherited;
  FItems := TvgrFormLocalizerItems.Create(Self, TvgrFormLocalizerItem);
  FBaseResIndex := svgrResStringsBase;
end;

destructor TvgrFormLocalizer.Destroy;
begin
  FreeAndNil(FItems);
  inherited;
end;

procedure TvgrFormLocalizer.Notification(AComponent: TComponent; AOperation: TOperation);
var
  AFormLocalizerItem: TvgrFormLocalizerItem;
begin
  inherited;
  if AOperation = opRemove then
  begin
    if (FItems <> nil) and not (csDestroying in ComponentState) then
    begin
      repeat
        AFormLocalizerItem := Items.FindByComponent(AComponent);
        if AFormLocalizerItem <> nil then
          AFormLocalizerItem.Free
        else
          break;
      until False;
    end;
  end;
end;

procedure TvgrFormLocalizer.SetItems(Value: TvgrFormLocalizerItems);
begin
  FItems.Assign(Value);
end;

function TvgrFormLocalizer.IsItemsStored: Boolean;
begin
  Result := not (csAncestor in ComponentState);
end;

procedure TvgrFormLocalizer.DoLocalize;
var
  I: Integer;
  S: string;
begin
  for I := 0 to Items.Count - 1 do
    with Items[I] do
      if IsStrProp then
        if Localizer.LoadStr(BaseResIndex, ResStringID, S) then
          SetStrValue(S);
end;

end.
