{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains the TvgrAvailableComponentsProvider class. This class provides the list of the components, which can be used in scripts.}
unit vgr_ScriptComponentsProvider;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  Classes, SysUtils, ActiveX, DB,

  vgr_CommonClasses, vgr_Functions;

type

{TvgrGetAvailableComponentsEvent is a type of event that can be used for
providing the custom list of accessible components.
Parameters:
  Sender - Instance of TvgrReportAvailableComponentsProvider.
  AList - Contains the default list of accessible components, you
can change this list. <b>AList.Items</b> contains names of the components.
<b>AList.Objects</b> contains references to the components.}
  TvgrGetAvailableComponentsEvent = procedure (Sender : TObject; AList : TStrings) of object;

  /////////////////////////////////////////////////
  //
  // TvgrAvailableComponentsProvider
  //
  /////////////////////////////////////////////////
{TvgrReportAvailableComponentsProvider provides the list of the accessible components.
By default this class are processing only objects derived from TComponent with not empty Name property.
The following algorithm is used to make the object list:<br>
1. Adds the root component (which specified by Root property).<br>
2. Adds all child components of the Root component.<br>
3. Adds the Root.Owner component, if it has no name then adds all its subcomponents.<br>
4. Adds the Root.Owner.Owner component and so on.<br>
For example, we have this hierarchy of the components:<br>
  Application (Owner for this component equals to NIL, name of the component not specified)<br>
    DataModule<br>
      Dataset1<br>
    Form1<br>
      vgrReportTemplate1 (assigned as Root to the Root property)<br>
      vgrWorksheet1<br>
      vgrWorksheet2<br>
      vgrDataBand1<br>
      LocalTable<br>
TvgrReportAvailableComponentsProvider builds the following list of the
available components:<br>
  Form1<br>
  vgrReportTemplate1<br>
  vgrWorksheet1<br>
  vgrWorksheet2<br>
  vgrDataBand1<br>
  LocalTable<br>
  DataModule<br>
<br>
You can use OnGetAvailableComponents event to override the default algorithm.
See also:
  OnGetAvailableComponents}
  TvgrAvailableComponentsProvider = class(TvgrPersistent)
  private
    FRoot: TComponent;
    FList: TStringList;
    FOnGetAvailableComponents: TvgrGetAvailableComponentsEvent;
    procedure SetRoot(Value: TComponent);
  protected
    FCache: TStringList;
  public
{Creates an instance of the TvgrAvailableComponentsProvider class.}
    constructor Create; override;
{Frees an instance of the TvgrAvailableComponentsProvider class.}
    destructor Destroy; override;

{Copies the contents of another, similar object.
Parameters:
  Source - The source object.}
    procedure Assign(Source: TPersistent); override;
{Caches the list of available objects.}
    procedure Cache;
{Clear the cache of available objects.}
    procedure ClearCache;
{Fills the list of available objects.
By default this procedure lists only objects derived from TComponent, you can
override this using events OnGetAvailableComponents
AList.Items contains names of the object.
AList.Objects contains references to the objects.
Parameters:
  AList - list of objects.
See also:
  GetAvailableDatasets, OnGetAvailableComponents, OnGetAvailableDatasets}
    procedure GetAvailableComponents(AList: TStrings); overload; virtual;
{Fills the list of available objects of specified type.
Parameters:
  AList - List of objects.
  AClass - Objects must be derived from this class.
See also:
  GetAvailableDatasets, OnGetAvailableComponents, OnGetAvailableDatasets}
    procedure GetAvailableComponents(AList: TStrings; AClass: TClass); overload; virtual;
{Fills the list of available datasets.
AList.Items contains names of the datasets.
AList.Objects contains references to the datasets.
Parameters:
  AList - list of datasets.
See also:
  GetAvailableComponents, OnGetAvailableComponents, OnGetAvailableDatasets}
    procedure GetAvailableDatasets(AList: TStrings); virtual;

{Sets or gets the parent component.}
    property Root: TComponent read FRoot write SetRoot;
    
{The cached components. This list contains valid data only after call the Cache method.
See also:
  Cache}
    property CachedComponents: TStringList read FCache;

{This event fired when list of available objects is formed
(in the GetAvailableComponents method).
See also:
  GetAvailableComponents, TvgrGetAvailableComponentsEvent}
    property OnGetAvailableComponents: TvgrGetAvailableComponentsEvent read FOnGetAvailableComponents write FOnGetAvailableComponents;   
  end;

implementation

const
  svgrDefaultRootName = 'Root';

/////////////////////////////////////////////////
//
// TvgrAvailableComponentsProvider
//
/////////////////////////////////////////////////
constructor TvgrAvailableComponentsProvider.Create;
begin
  inherited Create;
  FList := TStringList.Create;
  FCache := TStringList.Create;
end;

destructor TvgrAvailableComponentsProvider.Destroy;
begin
  FList.Free;
  FCache.Free;
  inherited;
end;

procedure TvgrAvailableComponentsProvider.SetRoot;
begin
  if FRoot <> Value then
  begin
    FRoot := Value;
    DoChange;
  end;
end;

procedure TvgrAvailableComponentsProvider.Assign(Source: TPersistent);
begin
  if Source is TvgrAvailableComponentsProvider then
    with TvgrAvailableComponentsProvider(Source) do
    begin
    end;
end;

procedure TvgrAvailableComponentsProvider.Cache;
begin
  GetAvailableComponents(FCache);
end;

procedure TvgrAvailableComponentsProvider.ClearCache;
begin
  FCache.Clear;
end;

procedure TvgrAvailableComponentsProvider.GetAvailableComponents(AList: TStrings; AClass: TClass);
var
  I: Integer;
begin
  GetAvailableComponents(AList);
  I := 0;
  while i < AList.Count do
    if TComponent(AList.Objects[I]).InheritsFrom(AClass) then
      Inc(I)
    else
      AList.Delete(I);
end;

procedure TvgrAvailableComponentsProvider.GetAvailableComponents(AList: TStrings);
var
  I: Integer;
  AOwner: TComponent;

  procedure AddComponent(AComponent: TComponent);
  begin
    if (AComponent.Name <> '') and
       (AList.IndexOfObject(AComponent) = -1) and
       (IndexInStrings(AComponent.Name, AList) = -1) then
      AList.AddObject(AComponent.Name, AComponent);
  end;

begin
  AList.Clear;
  if Root <> nil then
  begin
    if Root.Name = '' then
      AList.AddObject(svgrDefaultRootName, Root)
    else
      AList.AddObject(Root.Name, Root);

    AOwner := Root.Owner;
    while AOwner <> nil do
    begin
      AddComponent(AOwner);

      if (AOwner = Root.Owner) or (AOwner.Name = '') then
        for I := 0 to AOwner.ComponentCount - 1 do
          AddComponent(AOwner.Components[I]);

      AOwner := AOwner.Owner;
    end;
  end;
end;

procedure TvgrAvailableComponentsProvider.GetAvailableDatasets(AList: TStrings);
var
  I: Integer;
begin
  GetAvailableComponents(AList);
  I := 0;
  while i < AList.Count do
    if TComponent(AList.Objects[I]) is TDataSet then
      Inc(I)
    else
      AList.Delete(I);
end;

end.
