{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{TvgrAliasManager is a component, which allows to use aliases in the report template.}
unit vgr_AliasManager;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, SysUtils, ActiveX, Windows, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}

  vgr_CommonClasses, vgr_DataStorageTypes;

const
{This name using internally by script engine}
  sAliasManagerInternalScriptName = 'AliasManagerInternalNameA1E8065B04BA4B2CBCEF2FAB7BDC0F7E';
{Default node name mask}
  sNodeScriptName = 'Node%d';

type
  TvgrAliasNodes = class;

{Describes the types of node of Alias manager.
Items:
  vgranUnknown - TvgrAliasManager.OnUnknownVariable event will be fired when the value of node is needed.
  vgranCustom - The value of CustomValue property will be inserted in script instead of node name.
  vgranVariable - The Value property contains the value of node.
See also:
  TvgrAliasNode}
  TvgrAliasNodeType = (vgranUnknown, vgranCustom, vgranVariable);
  /////////////////////////////////////////////////
  //
  // TvgrAliasNode
  //
  /////////////////////////////////////////////////
{TvgrAliasNode Represents a separate node in the TvgrAliasManager component.}
  TvgrAliasNode = class(TCollectionItem)
  private
    FNodeType: TvgrAliasNodeType;
    FCustomValue: string;
    FParent: TvgrAliasNode;
    FName: string;
    FValue: Variant;
    FParentIndex: Integer; // used in load only

    function GetFullName: string;
    function GetHasChildren: Boolean;
    function GetNodes: TvgrAliasNodes;
    procedure SetParent(ANode: TvgrAliasNode);
    function GetNodeType: TvgrAliasNodeType;
    procedure SetNodeType(Value: TvgrAliasNodeType);
    function GetValueType: TvgrRangeValueType;
    function GetValue: OleVariant;
    procedure SetValue(const Value: Variant);
    function IsValueStored: Boolean;
  protected
    function CanBeParentFor(AAliasNode: TvgrAliasNode): Boolean;
    procedure DefineProperties(Filer: TFiler); override;
    procedure ReadParentIndex(Reader: TReader);
    procedure WriteParentIndex(Writer: TWriter);
  public
{Creates an instance of the AliasNode class.}
    constructor Create(Collection: TCollection); override;
{Frees an instance of the AliasNode class.}
    destructor Destroy; override;
{Copies the contents of another node to the current one.}
    procedure Assign(Source: TPersistent); override;
{Parent Node for current node. Nil if hode has no parent. }
    property Parent: TvgrAliasNode read FParent write SetParent;
{Returns the full name of node, you may use it in the script. This property take into account
all parent Nodes.}
    property FullName: string read GetFullName;
{True if node has children.}    
    property HasChildren: Boolean read GetHasChildren;
{Returns the type of a Value property.
See also:
  TvgrRangeValueType}
    property ValueType: TvgrRangeValueType read GetValueType;
{Tree of nodes in a AliasManager component.}
    property Nodes: TvgrAliasNodes read GetNodes;
  published
{Name of Alias, which you may use in script in a braces.}
    property Name: string read FName write FName;
{Defines the logic of the Alias calculating.
See also:
  TvgrAliasNodeType}
    property NodeType: TvgrAliasNodeType read GetNodeType write SetNodeType default vgranVariable;
(*If you want to use alias instead of, for example component name and property,
you may Add a node with NodeType vgranCustom and set CustomValue property to
ComponentName.Property.
Example:
  begin
    ...
    // When script will be executing, all {CountOfPages} will be replaced
    // with Edit1.Text
    ANode.Name := 'CountOfPages';
    ANode.NodeType := vgranCustom;
    ANode.CustomValue := 'Edit1.Text';
    ...
  end;*)
    property CustomValue: string read FCustomValue write FCustomValue;
{Value of Node, used if NodeType equals to vgranVariable.
See also:
  TvgrAliasNodeType}
    property Value: Variant read FValue write SetValue stored IsValueStored;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrAliasNodes
  //
  /////////////////////////////////////////////////
{TvgrAliasNodes maintains a tree of nodes in the TvgrAliasManager component.}
  TvgrAliasNodes = class(TOwnedCollection)
  private
    function GetItm(Index: Integer): TvgrAliasNode;
    procedure SetItm(Index: Integer; const Value: TvgrAliasNode);
{$IFNDEF VTK_D6_OR_D7}
    function _GetOwner: TPersistent;
{$ENDIF}
    procedure Loaded;
  public
{Creates a new TvgrAliasNode instance and adds it to the Items array.
Return value:
  TvgrAliasNode}
    function Add: TvgrAliasNode; reintroduce; overload;
{Adds the new node to a AliasManager. The node is added as a child of the
node specified by the AParentNode parameter.
Parameters:
  AParentNode - The parent node.
Return value:
  TvgrAliasNode}
    function Add(AParentNode: TvgrAliasNode): TvgrAliasNode; overload;
{Deletes all children nodes for node, specified by AParentNode parameter.
Parameters:
  AParentNode - The parent node.}
    procedure DeleteChildren(AParentNode: TvgrAliasNode);
{$IFNDEF VTK_D6_OR_D7}
    property Owner: TPersistent read _GetOwner;
{$ENDIF}
{Lists all nodes managed by the TvgrAliasNodes object.}
    property Items[Index: Integer]: TvgrAliasNode read GetItm  write SetItm; default;
  end;

{TvgrUnknownVariableProc is used for OnUnknownVariable event.
This event occurs when the value of node with NodeType = vgranUnknown is needed.
Parameters:
  Sender - The TvgrAliasManager object firing an event.
  ANode - The TvgrAliasNode object, value of which is requested.
  AValue - Must contains the value of node.}
  TvgrUnknownVariableProc = procedure(Sender: TObject; ANode: TvgrAliasNode; var AValue: OleVariant) of object;

  /////////////////////////////////////////////////
  //
  // TvgrAliasManager
  //
  /////////////////////////////////////////////////
{TvgrAliasManager represents a hierarchical list of user defined aliases for using in scripts and expressions.}
  TvgrAliasManager = class(TvgrComponent)
  private
    FNodes: TvgrAliasNodes;
    FOnUnknownVariable: TvgrUnknownVariableProc;
    function GetInternalScriptName: string;
    function GetNodeScriptName(Index: Integer): string;
  protected
    function GetDispIdOfName(const AName: string): Integer; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;
    procedure Loaded; override;
  public
{Creates an instance of the TvgrAliasManager class.}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of the TvgrAliasManager class.}
    destructor Destroy; override;

{Adds a node to the AliasManager. The node is added as the child of the node
specified by the AParentNode parameter
Parameters:
  AParentNode - The parent node for the created node.
Return value:
  The create node.}
    function AddNode(AParentNode: TvgrAliasNode): TvgrAliasNode;

{Finds and replaces the names of aliases with their values in string. 
Parameters:
  Value - String to replace.
Return value:
  The processed string.}
    function ProcessAliases(const Value: string): string;
{This name is used internally by script engine.}
    property InternalScriptName: string read GetInternalScriptName;
  published
{Nodes maintains a list of tree nodes in a AliasManager component.}
    property Nodes: TvgrAliasNodes read FNodes write FNodes;
{This event occurs when the value of node with NodeType = vgranUnknown is needed.
See also:
  TvgrUnknownVariableProc}
    property OnUnknownVariable: TvgrUnknownVariableProc read FOnUnknownVariable write FOnUnknownVariable;
  end;

implementation

uses
  vgr_Functions;

/////////////////////////////////////////////////
//
// TvgrAliasNode
//
/////////////////////////////////////////////////
constructor TvgrAliasNode.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FValue := Null;
  FNodeType := vgranVariable;
  FParentIndex := -1;
end;

destructor TvgrAliasNode.Destroy;
begin
  // Delete all children
  if Nodes <> nil then
    Nodes.DeleteChildren(Self);
  inherited;
end;

procedure TvgrAliasNode.Assign(Source: TPersistent);
begin
  if Source is TvgrAliasNode then
    with TvgrAliasNode(Source) do
    begin
      Self.FNodeType := FNodeType;
      Self.FCustomValue := FCustomValue;
      Self.FName := FName;
      Self.FValue := FValue;
      if (Parent = nil) or Parent.CanBeParentFor(Self) then
        Self.Parent := FParent;
    end;
end;

function TvgrAliasNode.CanBeParentFor(AAliasNode: TvgrAliasNode): Boolean;
var
  AParent: TvgrAliasNode;
begin
  AParent := Parent;
  while (AParent <> nil) and (AParent <> AAliasNode) do AParent := AParent.Parent;
  Result := AParent = nil;
end;

procedure TvgrAliasNode.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('ParentIndex', ReadParentIndex, WriteParentIndex, Parent <> nil);
end;

procedure TvgrAliasNode.ReadParentIndex(Reader: TReader);
begin
  FParentIndex := Reader.ReadInteger;
end;

procedure TvgrAliasNode.WriteParentIndex(Writer: TWriter);
begin
  Writer.WriteInteger(Parent.Index);
end;

function TvgrAliasNode.GetValueType: TvgrRangeValueType;
begin
  case VarType(FValue) of
    {$IFDEF VGR_D6_OR_D7}varShortInt, varWord, varLongWord, varInt64, {$ENDIF} varByte, varSmallint, varInteger:
      Result := rvtInteger;
    varCurrency, varSingle, varDouble:
      Result := rvtExtended;
    varDate:
      Result := rvtDateTime;
    varOleStr, varStrArg, varString:
      Result := rvtString;
    varBoolean:
      Result := rvtInteger;
  else
    Result := rvtNull;
  end;
end;

procedure TvgrAliasNode.SetValue(const Value: Variant);
begin
  FValue := Value;
  // TODO: add checks here (NO arrays and other)
end;

function TvgrAliasNode.IsValueStored: Boolean;
begin
  Result := not VarIsNull(Value); 
end;

function TvgrAliasNode.GetFullName: string;
var
  ANode: TvgrAliasNode;
begin
  Result := '';
  if not HasChildren then
  begin
    ANode := Self;
    repeat
      if Result <> '' then
        Result := '.' + Result;
      Result := ANode.Name + Result;
      ANode := ANode.Parent;
    until ANode = nil;
  end;
end;

function TvgrAliasNode.GetHasChildren: Boolean;
var
  I: Integer;
begin
  Result := False;
  I := 0;
  while (I < Nodes.Count) and not Result do
  begin
    Result := Result or (Nodes[I].Parent = Self);
    Inc(I);
  end;
end;

function TvgrAliasNode.GetNodes: TvgrAliasNodes;
begin
  Result := TvgrAliasNodes(Collection);
end;

procedure TvgrAliasNode.SetParent(ANode: TvgrAliasNode);
begin
  FParent := ANode;
end;

function TvgrAliasNode.GetNodeType: TvgrAliasNodeType;
begin
  Result := FNodeType;
end;

procedure TvgrAliasNode.SetNodeType(Value: TvgrAliasNodeType);
begin
  FNodeType := Value;
end;

function TvgrAliasNode.GetValue: OleVariant;
begin
  case NodeType of
    vgranCustom: Result := Unassigned;
    vgranUnknown:
      if Assigned(TvgrAliasManager(Nodes.Owner).OnUnknownVariable) then
        TvgrAliasManager(Nodes.Owner).OnUnknownVariable(TvgrAliasManager(Nodes.Owner),Self,Result);
    vgranVariable: Result := FValue;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrAliasNodes
//
/////////////////////////////////////////////////
{$IFNDEF VTK_D6_OR_D7}
function TvgrAliasNodes._GetOwner: TPersistent;
begin
  Result := GetOwner;
end;
{$ENDIF}

function TvgrAliasNodes.GetItm(Index: Integer): TvgrAliasNode;
begin
  Result := TvgrAliasNode(inherited Items[Index]);
end;

procedure TvgrAliasNodes.SetItm(Index: Integer; const Value: TvgrAliasNode);
begin
  inherited Items[Index] := Value;
end;

procedure TvgrAliasNodes.Loaded;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    if (Items[I].FParentIndex >= 0) and (Items[I].FParentIndex < Count) then
      Items[I].FParent := Items[Items[I].FParentIndex];
end;

function TvgrAliasNodes.Add: TvgrAliasNode;
begin
  Result := TvgrAliasNode(inherited Add);
end;

function TvgrAliasNodes.Add(AParentNode: TvgrAliasNode): TvgrAliasNode;
begin
  Result := Add;
  Result.Parent := AParentNode;
end;

procedure TvgrAliasNodes.DeleteChildren(AParentNode: TvgrAliasNode);
var
  I: Integer;
begin
  I := 0;
  while I < Count do
  begin
    if Items[i].Parent = AParentNode then
      Delete(I)
    else
      Inc(I);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrAliasManager
//
/////////////////////////////////////////////////
constructor TvgrAliasManager.Create(AOwner: TComponent);
begin
  inherited;
  FNodes := TvgrAliasNodes.Create(Self, TvgrAliasNode);
end;

destructor TvgrAliasManager.Destroy;
begin
  FNodes.Free;
  inherited;
end;

procedure TvgrAliasManager.Loaded;
begin
  inherited;
  FNodes.Loaded;
end;

{function TvgrAliasManager.GetString(const Value: string): string;
begin

end;}

function TvgrAliasManager.AddNode(AParentNode: TvgrAliasNode): TvgrAliasNode;
begin
  Result := FNodes.Add(AParentNode);
end;

function TvgrAliasManager.ProcessAliases(const Value: string): string;
var
  I: Integer;
begin
  Result := Value;
  for I := 0 to Nodes.Count - 1 do
  case Nodes[I].NodeType of
    vgranUnknown, vgranVariable:
      Result := StringReplace(Result, '{'+Nodes[I].FullName+'}', GetInternalScriptName+'.'+GetNodeScriptName(I), [rfReplaceAll, rfIgnoreCase]);
    vgranCustom:
      Result := StringReplace(Result, '{'+Nodes[I].FullName+'}', Nodes[I].CustomValue, [rfReplaceAll, rfIgnoreCase]);
  end;
end;

function TvgrAliasManager.GetDispIdOfName(const AName: string): Integer;
var
  I: Integer;
begin
  I := 0;
  while (I < Nodes.Count) and (AnsiCompareStr(AName, GetNodeScriptName(I)) <> 0) do
    Inc(I);
    
  if I < Nodes.Count then
    Result := I
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrAliasManager.DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  if DispId < Nodes.Count then
  begin
    AResult := Nodes[DispID].GetValue;
    Result := S_OK;
  end
  else
    Result := DoInvoke(DispId, Flags, AParameters, AResult);
end;

function TvgrAliasManager.GetInternalScriptName: string;
begin
  Result := sAliasManagerInternalScriptName;
end;

function TvgrAliasManager.GetNodeScriptName(Index: Integer): string;
begin
  Result := Format(sNodeScriptName,[Index]);
end;

end.
