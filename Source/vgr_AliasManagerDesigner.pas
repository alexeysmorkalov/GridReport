{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes for designing the TvgrAliasManager component.}
unit vgr_AliasManagerDesigner;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, Menus, ActnList, StdCtrls, ExtCtrls, ToolWin, ImgList, DB,

  vgr_DataStorageTypes, vgr_AliasManager, vgr_Functions, vgr_CommonClasses,
  vgr_ScriptComponentsProvider, vgr_Form, vgr_StringIDs, vgr_FormLocalizer;

type

  TvgrAliasManagerDesignerForm = class;

{Describes the various options of work of TvgrAliasManagerComponentsProvider class.
Items:
  vgrcpoDatasetsOnly - If specified only datasets (descendants of TDataset class)
can be used in the designer.
Syntax:
  TvgrAliasManagerComponentsProviderOptions = (vgrcpoDatasetsOnly);}
  TvgrAliasManagerComponentsProviderOptions = (vgrcpoDatasetsOnly);
{Describes the set of the various options of work of TvgrAliasManagerComponentsProvider class.
Syntax:
  TvgrAliasManagerComponentsProviderOptionsSet = set of TvgrAliasManagerComponentsProviderOptions;}
  TvgrAliasManagerComponentsProviderOptionsSet = set of TvgrAliasManagerComponentsProviderOptions;
  /////////////////////////////////////////////////
  //
  // TvgrAliasManagerComponentsProvider
  //
  /////////////////////////////////////////////////
{Provides the list of components which can be used in designer.}
  TvgrAliasManagerComponentsProvider = class(TvgrAvailableComponentsProvider)
  private
    FOptions: TvgrAliasManagerComponentsProviderOptionsSet;
    function GetOptions: TvgrAliasManagerComponentsProviderOptionsSet;
    procedure SetOptions(Value: TvgrAliasManagerComponentsProviderOptionsSet);
  public
{Creates an instance of the TvgrAliasManagerComponentsProvider class.}
    constructor Create; override;
  published
    property Root;
{Specifies the options of work.
See also:
  TvgrAliasManagerComponentsProviderOptionsSet}
    property Options: TvgrAliasManagerComponentsProviderOptionsSet read GetOptions write SetOptions default [vgrcpoDatasetsOnly];
  end;

  /////////////////////////////////////////////////
  //
  // TvgrAliasManagerImportWizard
  //
  /////////////////////////////////////////////////
  TvgrAliasManagerImportWizard = class(TObject)
  public
    class procedure AddNodes(AObject: TObject; const AFullComponentName: string; ADesignerForm: TvgrAliasManagerDesignerForm; ASelectedNode: TvgrAliasNode); virtual;
    class function GetClass: TClass; virtual;
  end;
  TvgrAliasManagerImportWizardClass = class of TvgrAliasManagerImportWizard;

  /////////////////////////////////////////////////
  //
  // TvgrAliasManagerImportWizards
  //
  /////////////////////////////////////////////////
  TvgrAliasManagerImportWizards = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrAliasManagerImportWizardClass;
  public
    procedure RegisterImportWizard(AWizard: TvgrAliasManagerImportWizardClass);
    function FindWizard(AClass: TClass): TvgrAliasManagerImportWizardClass;
    function IndexByClass(AClass: TClass): Integer;
    property Items[Index: Integer]: TvgrAliasManagerImportWizardClass read GetItem;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrDatasetImportWizard
  //
  /////////////////////////////////////////////////
  TvgrDatasetImportWizard = class(TvgrAliasManagerImportWizard)
  public
    class procedure AddNodes(AObject: TObject; const AFullComponentName: string; ADesignerForm: TvgrAliasManagerDesignerForm; ASelectedNode: TvgrAliasNode); override;
    class function GetClass: TClass; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrAliasManagerDesignerForm
  //
  /////////////////////////////////////////////////
  TvgrAliasManagerDesignerForm = class(TvgrDialogForm)
    ImageList: TImageList;
    Panel1: TPanel;
    Panel2: TPanel;
    TV: TTreeView;
    Panel3: TPanel;
    bAdd: TButton;
    bAddChild: TButton;
    bDelete: TButton;
    Panel4: TPanel;
    bClose: TButton;
    Label1: TLabel;
    edName: TEdit;
    rbUnknown: TRadioButton;
    rbCustom: TRadioButton;
    rbVariable: TRadioButton;
    edVariableValue: TEdit;
    edVariableType: TComboBox;
    Bevel1: TBevel;
    bCancel: TButton;
    bApply: TButton;
    Bevel2: TBevel;
    edCustom: TComboBox;
    bOptions: TButton;
    PMOptions: TPopupMenu;
    mShowAllComponents: TMenuItem;
    mShowDatasetsOnly: TMenuItem;
    N1: TMenuItem;
    mImport: TMenuItem;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure bAddClick(Sender: TObject);
    procedure bAddChildClick(Sender: TObject);
    procedure bDeleteClick(Sender: TObject);
    procedure bCloseClick(Sender: TObject);
    procedure TVChange(Sender: TObject; Node: TTreeNode);
    procedure FormCreate(Sender: TObject);
    procedure edNameChange(Sender: TObject);
    procedure rbUnknownClick(Sender: TObject);
    procedure rbCustomClick(Sender: TObject);
    procedure rbVariableClick(Sender: TObject);
    procedure edCustomChange(Sender: TObject);
    procedure edVariableValueChange(Sender: TObject);
    procedure bCancelClick(Sender: TObject);
    procedure bApplyClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure TVEdited(Sender: TObject; Node: TTreeNode; var S: String);
    procedure bOptionsClick(Sender: TObject);
    procedure PMOptionsPopup(Sender: TObject);
    procedure mShowAllComponentsClick(Sender: TObject);
    procedure mShowDatasetsOnlyClick(Sender: TObject);
    procedure mImportClick(Sender: TObject);
  private
    FUpdateCount: Integer;
    FAliasManager: TvgrAliasManager;
    FSelectedNodeChanged: Boolean;
    FAvailableComponentsProvider: TvgrAliasManagerComponentsProvider;

    procedure SetAliasManager(Value: TvgrAliasManager);
    procedure UpdateEnabled;
    procedure UpdateNode(AAliasNode: TvgrAliasNode; ANode: TTreeNode);
    function AddNode(AAliasNode: TvgrAliasNode): TTreeNode;
    procedure UpdateTV;
    procedure UpdateCustom;
    procedure CopyToControls(AAliasNode: TvgrAliasNode);
    procedure CopyFromControls(AAliasNode: TvgrAliasNode);
    procedure BeginUpdate;
    procedure EndUpdate;
    function GetEnableUpdate: Boolean;
    function GetAliasNodes: TvgrAliasNodes;
    function GetNodeCaption(AAliasNode: TvgrAliasNode): string;
    function GetSelectedAliasNode: TvgrAliasNode;
    function CheckUName(AAliasNode: TvgrAliasNode; const AName: string): Boolean;
    function GetUniqueNodeName(AAliasNode: TvgrAliasNode): string;
    procedure SetSelectedNodeChanged(Value: Boolean);
  protected
    property EnableUpdate: Boolean read GetEnableUpdate;
    property SelectedNodeChanged: Boolean read FSelectedNodeChanged write SetSelectedNodeChanged;
  public
    function AddAliasNode(AParentNode: TvgrAliasNode): TvgrAliasNode;
    procedure UpdateAliasNode(AAliasNode: TvgrAliasNode);
    procedure DeleteAliasNode(AAliasNode: TvgrAliasNode);
    procedure ClearChildNodes(AAliasNode: TvgrAliasNode);
    
    procedure Execute(AAliasManager: TvgrAliasManager; AAvailableComponentsProvider: TvgrAliasManagerComponentsProvider);

    property AliasManager: TvgrAliasManager read FAliasManager write SetAliasManager;
    property AliasNodes: TvgrAliasNodes read GetAliasNodes;
    property SelectedAliasNode: TvgrAliasNode read GetSelectedAliasNode;
    property AvailableComponentsProvider: TvgrAliasManagerComponentsProvider read FAvailableComponentsProvider;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrAliasManagerDesigner
  //
  /////////////////////////////////////////////////
{TvgrAliasManagerDesignerDialog displays a dialog for designing the TvgrAliasManager component.}
  TvgrAliasManagerDesignerDialog = class(TComponent)
  private
    FAliasManager: TvgrAliasManager;
    FForm: TvgrAliasManagerDesignerForm;
    FAvailableComponentsProvider: TvgrAliasManagerComponentsProvider;
    procedure SetAvailableComponentsProvider(Value: TvgrAliasManagerComponentsProvider);
    procedure OnProviderChange(Sender: TObject);
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
  public
{Creates an instance the TvgrAliasManagerDesignerDialog class.}
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    
    procedure Execute;
    property Form: TvgrAliasManagerDesignerForm read FForm;
  published
    property AliasManager: TvgrAliasManager read FAliasManager write FAliasManager;
    property AvailableComponentsProvider: TvgrAliasManagerComponentsProvider read FAvailableComponentsProvider write SetAvailableComponentsProvider;
  end;

var
  AMImportWizards: TvgrAliasManagerImportWizards;

implementation

uses Math, vgr_Dialogs, vgr_Localize;

{$R *.dfm}
{$R ..\res\vgr_AliasManagerDesignerStrings.res}

function FindTreeNodeByData(ATreeView: TTreeView; AData: Pointer): TTreeNode;
var
  I: Integer;
begin
  I := 0;
  while (I < ATreeView.Items.Count) and (ATreeView.Items[I].Data <> AData) do Inc(I);
  if I < ATreeView.Items.Count then
    Result := ATreeView.Items[I]
  else
    Result := nil;
end;

/////////////////////////////////////////////////
//
// TvgrDatasetImportWizard
//
/////////////////////////////////////////////////
class procedure TvgrDatasetImportWizard.AddNodes(AObject: TObject; const AFullComponentName: string; ADesignerForm: TvgrAliasManagerDesignerForm; ASelectedNode: TvgrAliasNode);
var
  I: Integer;
  AActive: Boolean;
  AFieldNode: TvgrAliasNode;
begin
  if AObject is TDataset then
    with TDataset(AObject) do
    begin
      AActive := Active;
      try
        try
          if not AActive then
            Open;

          ASelectedNode.Name := Name;
          ADesignerForm.UpdateAliasNode(ASelectedNode);
          for I := 0 to FieldCount - 1 do
          begin
            AFieldNode := ADesignerForm.AddAliasNode(ASelectedNode);
            AFieldNode.Name := Fields[I].FieldName;
            AFieldNode.NodeType := vgranCustom;
            AFieldNode.CustomValue := AFullComponentName + '.' + Fields[I].FieldName;
            ADesignerForm.UpdateAliasNode(AFieldNode);
          end;
        except
          on E: Exception do
            MBox(E.Message, MB_OK or MB_ICONEXCLAMATION);
        end;
      finally
        if not AActive and Active then
          Close; 
      end;
    end;
end;

class function TvgrDatasetImportWizard.GetClass: TClass;
begin
  Result := TDataset;
end;

/////////////////////////////////////////////////
//
// TvgrAliasManagerImportWizards
//
/////////////////////////////////////////////////
function TvgrAliasManagerImportWizards.GetItem(Index: Integer): TvgrAliasManagerImportWizardClass;
begin
  Result := TvgrAliasManagerImportWizardClass(inherited Items[Index]);
end;

procedure TvgrAliasManagerImportWizards.RegisterImportWizard(AWizard: TvgrAliasManagerImportWizardClass);
begin
  if IndexByClass(AWizard.GetClass) = -1 then
    Add(TObject(AWizard));
end;

function TvgrAliasManagerImportWizards.IndexByClass(AClass: TClass): Integer;
begin
  Result := 0;
  while (Result < Count) and (Items[Result].GetClass <> AClass) do
    Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrAliasManagerImportWizards.FindWizard(AClass: TClass): TvgrAliasManagerImportWizardClass;
var
  I: Integer;
begin
  I := 0;
  while (I < Count) and (not AClass.InheritsFrom(Items[I].GetClass)) do Inc(I);
  if I >= Count then
    Result := nil
  else
    Result := Items[I];
end;

/////////////////////////////////////////////////
//
// TvgrAliasManagerComponentsProvider
//
/////////////////////////////////////////////////
constructor TvgrAliasManagerComponentsProvider.Create;
begin
  inherited;
  FOptions := [vgrcpoDatasetsOnly];
end;

function TvgrAliasManagerComponentsProvider.GetOptions: TvgrAliasManagerComponentsProviderOptionsSet;
begin
  Result := FOptions;
end;

procedure TvgrAliasManagerComponentsProvider.SetOptions(Value: TvgrAliasManagerComponentsProviderOptionsSet);
begin
  if FOptions <> Value then
  begin
    FOptions := Value;
    DoChange;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrAliasManagerImportWizard
//
/////////////////////////////////////////////////
class procedure TvgrAliasManagerImportWizard.AddNodes(AObject: TObject; const AFullComponentName: string; ADesignerForm: TvgrAliasManagerDesignerForm; ASelectedNode: TvgrAliasNode);
begin
end;

class function TvgrAliasManagerImportWizard.GetClass: TClass;
begin
  Result := TObject;
end;

/////////////////////////////////////////////////
//
// TvgrAliasManagerDesignerDialog
//
/////////////////////////////////////////////////
constructor TvgrAliasManagerDesignerDialog.Create(AOwner: TComponent);
begin
  inherited;
  FAvailableComponentsProvider := TvgrAliasManagerComponentsProvider.Create;
  FAvailableComponentsProvider.OnChange := OnProviderChange;
end;

destructor TvgrAliasManagerDesignerDialog.Destroy;
begin
  FAvailableComponentsProvider.Free;
  inherited;
end;

procedure TvgrAliasManagerDesignerDialog.SetAvailableComponentsProvider(Value: TvgrAliasManagerComponentsProvider);
begin
  FAvailableComponentsProvider.Assign(Value);
end;

procedure TvgrAliasManagerDesignerDialog.OnProviderChange(Sender: TObject);
begin
  if Form <> nil then
    Form.UpdateCustom;
end;

procedure TvgrAliasManagerDesignerDialog.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if AComponent = FAliasManager then
      FAliasManager := nil;
    if AComponent = AvailableComponentsProvider.Root then
      AvailableComponentsProvider.Root := nil;
  end;
end;

procedure TvgrAliasManagerDesignerDialog.Execute;
begin
  FForm := TvgrAliasManagerDesignerForm.Create(Self);
  FForm.Execute(AliasManager, AvailableComponentsProvider);
  FForm := nil;
end;

/////////////////////////////////////////////////
//
// TvgrAliasManagerDesignerForm
//
/////////////////////////////////////////////////
procedure TvgrAliasManagerDesignerForm.Execute(AAliasManager: TvgrAliasManager; AAvailableComponentsProvider: TvgrAliasManagerComponentsProvider);
begin
  FAvailableComponentsProvider := AAvailableComponentsProvider;
  AliasManager := AAliasManager;
  ShowModal;
end;

procedure TvgrAliasManagerDesignerForm.SetAliasManager(Value: TvgrAliasManager);
begin
  FAliasManager := Value;
  UpdateTV;
  UpdateCustom;
end;

procedure TvgrAliasManagerDesignerForm.SetSelectedNodeChanged(Value: Boolean);
begin
  if FSelectedNodeChanged <> Value then
  begin
    FSelectedNodeChanged := Value;
    UpdateEnabled;
  end;
end;

procedure TvgrAliasManagerDesignerForm.BeginUpdate;
begin
  Inc(FUpdateCount);
end;

procedure TvgrAliasManagerDesignerForm.EndUpdate;
begin
  Dec(FUpdateCount);
end;

function TvgrAliasManagerDesignerForm.GetEnableUpdate: Boolean;
begin
  Result := FUpdateCount = 0;
end;

function TvgrAliasManagerDesignerForm.GetAliasNodes: TvgrAliasNodes;
begin
  if FAliasManager = nil then
    Result := nil
  else
    Result := FAliasManager.Nodes;
end;

function TvgrAliasManagerDesignerForm.GetNodeCaption(AAliasNode: TvgrAliasNode): string;
begin
  Result := AAliasNode.Name;
end;

function TvgrAliasManagerDesignerForm.GetSelectedAliasNode: TvgrAliasNode;
begin
  if TV.Selected = nil then
    Result := nil
  else
    Result := TvgrAliasNode(TV.Selected.Data);
end;

procedure TvgrAliasManagerDesignerForm.UpdateNode(AAliasNode: TvgrAliasNode; ANode: TTreeNode);
begin
  ANode.Text := GetNodeCaption(AAliasNode);
  ANode.Data := AAliasNode;
  if AAliasNode.HasChildren then
  begin
    ANode.ImageIndex := 0;
    ANode.SelectedIndex := 0;
  end
  else
  begin
    ANode.ImageIndex := Integer(AAliasNode.NodeType) + 1;
    ANode.SelectedIndex := Integer(AAliasNode.NodeType) + 1;
  end;
end;

function TvgrAliasManagerDesignerForm.AddNode(AAliasNode: TvgrAliasNode): TTreeNode;
var
  AParentNode: TTreeNode;
begin
  if AAliasNode.Parent = nil then
    AParentNode := nil
  else
  begin
    AParentNode := FindTreeNodeByData(TV, AAliasNode.Parent);
    if AParentNode = nil then
      AParentNode := AddNode(AAliasNode.Parent);
  end;
  Result := TV.Items.AddChildObject(AParentNode, GetNodeCaption(AAliasNode), AAliasNode);
  UpdateNode(AAliasNode, Result);
end;

procedure TvgrAliasManagerDesignerForm.UpdateTV;
var
  I: Integer;
begin
  BeginUpdate;
  TV.Items.BeginUpdate;
  try
    TV.Items.Clear;
    for I := 0 to AliasNodes.Count - 1 do
      AddNode(AliasNodes[I]);
    if TV.Items.Count > 0 then
      TV.Selected := TV.Items[0];
    UpdateEnabled;
  finally
    TV.Items.EndUpdate;
    EndUpdate;
  end;
end;

procedure TvgrAliasManagerDesignerForm.UpdateCustom;
begin
  BeginUpdate;
  edCustom.Items.BeginUpdate;
  try
    edCustom.Items.Clear;
    if vgrcpoDatasetsOnly in AvailableComponentsProvider.Options then
      AvailableComponentsProvider.GetAvailableDatasets(edCustom.Items)
    else
      AvailableComponentsProvider.GetAvailableComponents(edCustom.Items);
  finally
    edCustom.Items.EndUpdate;
    EndUpdate;
  end;
end;

procedure TvgrAliasManagerDesignerForm.CopyToControls(AAliasNode: TvgrAliasNode);
var
  AHasChildren: Boolean;
begin
  BeginUpdate;
  try
    edName.Text := AAliasNode.Name;

    AHasChildren := AAliasNode.HasChildren;
    rbUnknown.Enabled := not AHasChildren;
    rbCustom.Enabled := not AHasChildren;
    rbVariable.Enabled := not AHasChildren;
    edCustom.Enabled := not AHasChildren;
    edVariableValue.Enabled := not AHasChildren;
    edVariableType.Enabled := not AHasChildren;

    if AAliasNode.HasChildren then
    begin
      edCustom.Text := '';
      edVariableValue.Text := '';
      edVariableType.ItemIndex := -1;
    end
    else
      case AAliasNode.NodeType of
        vgranUnknown:
          begin
            rbUnknown.Checked := True;
            edCustom.Text := '';
            edVariableValue.Text := '';
            edVariableType.ItemIndex := -1;
          end;
        vgranCustom:
          begin
            rbCustom.Checked := True;
            edCustom.Text := AAliasNode.CustomValue;
            edVariableValue.Text := '';
            edVariableType.ItemIndex := -1;
          end;
        vgranVariable:
          begin
            rbVariable.Checked := True;
            edCustom.Text := '';
            edVariableValue.Text := VarToStr(AAliasNode.Value);
            edVariableType.ItemIndex := Integer(AAliasNode.ValueType);
          end;
      end;
  finally
    EndUpdate;
  end;
end;

procedure TvgrAliasManagerDesignerForm.CopyFromControls(AAliasNode: TvgrAliasNode);
begin
  AAliasNode.Name := edName.Text;
  if rbUnknown.Checked then
    AAliasNode.NodeType := vgranUnknown
  else
    if rbCustom.Checked then
    begin
      AAliasNode.NodeType := vgranCustom;
      AAliasNode.CustomValue := edCustom.Text;
    end
    else
      if rbVariable.Checked then
      begin
        AAliasNode.NodeType := vgranVariable;
        case TvgrRangeValueType(edVariableType.ItemIndex) of
          rvtNull: AAliasNode.Value := Null;
          rvtInteger: AAliasNode.Value := StrToInt(edVariableValue.Text);
          rvtExtended: AAliasNode.Value := StrToFloat(edVariableValue.Text);
          rvtDateTime: AAliasNode.Value := StrToDateTime(edVariableValue.Text);
          else
            AAliasNode.Value := edVariableValue.Text;
        end;
      end;
end;

procedure TvgrAliasManagerDesignerForm.UpdateEnabled;
begin
  bAdd.Enabled := AliasNodes <> nil;
  bAddChild.Enabled := SelectedAliasNode <> nil;
  bDelete.Enabled := SelectedAliasNode <> nil;

  bApply.Enabled := (SelectedAliasNode <> nil) and FSelectedNodeChanged;
  bCancel.Enabled := (SelectedAliasNode <> nil) and FSelectedNodeChanged;

  Label1.Enabled := SelectedAliasNode <> nil;
  edName.Enabled := SelectedAliasNode <> nil;
  rbUnknown.Enabled := SelectedAliasNode <> nil;
  rbCustom.Enabled := SelectedAliasNode <> nil;
  edCustom.Enabled := SelectedAliasNode <> nil;
  bOptions.Enabled := SelectedAliasNode <> nil;
  rbVariable.Enabled := SelectedAliasNode <> nil;
  edVariableValue.Enabled := SelectedAliasNode <> nil;
  edVariableType.Enabled := SelectedAliasNode <> nil;
end;

function TvgrAliasManagerDesignerForm.CheckUName(AAliasNode: TvgrAliasNode; const AName: string): Boolean;
var
  I: Integer;
begin
  Result := True;
  for I := 0 to AAliasNode.Nodes.Count - 1 do
    if (AAliasNode.Nodes[I].Parent = AAliasNode.Parent) and (AAliasNode.Nodes[I] <> AAliasNode) then
      Result := Result and (AnsiCompareText(AName, AAliasNode.Nodes[I].Name) <> 0);
end;

function TvgrAliasManagerDesignerForm.GetUniqueNodeName(AAliasNode: TvgrAliasNode): string;
var
  I: Integer;
begin
  I:= 1;
  repeat
    Result := 'Node' + IntToStr(I);
    Inc(I);
  until CheckUName(AAliasNode, Result);
end;

function TvgrAliasManagerDesignerForm.AddAliasNode(AParentNode: TvgrAliasNode): TvgrAliasNode;
var
  ANode: TTreeNode;
begin
  Result := AliasManager.AddNode(AParentNode);
  Result.Name := GetUniqueNodeName(Result);
  TV.Selected := AddNode(Result);
  if AParentNode <> nil then
  begin
    ANode := FindTreeNodeByData(TV, AParentNode);
    UpdateNode(AParentNode, ANode);
  end;
  UpdateEnabled;
end;

procedure TvgrAliasManagerDesignerForm.DeleteAliasNode(AAliasNode: TvgrAliasNode);
var
  ANode: TTreeNode;
  AParentAliasNode: TvgrAliasNode;
begin
  AParentAliasNode := AAliasNode.Parent;
  ANode := FindTreeNodeByData(TV, AAliasNode);
  TV.Items.Delete(ANode);
  AAliasNode.Free;
  if AParentAliasNode <> nil then
    UpdateNode(AParentAliasNode, FindTreeNodeByData(TV, AParentAliasNode));
  UpdateEnabled;
end;

procedure TvgrAliasManagerDesignerForm.ClearChildNodes(AAliasNode: TvgrAliasNode);
begin
  AAliasNode.Nodes.DeleteChildren(AAliasNode);
  UpdateTV;
end;

procedure TvgrAliasManagerDesignerForm.UpdateAliasNode(AAliasNode: TvgrAliasNode);
begin
  UpdateNode(AAliasNode, FindTreeNodeByData(TV, AAliasNode));
end;

procedure TvgrAliasManagerDesignerForm.bAddClick(Sender: TObject);
begin
  if SelectedAliasNode = nil then
    AddAliasNode(nil)
  else
    AddAliasNode(SelectedAliasNode.Parent);
  ActiveControl := edName;
end;

procedure TvgrAliasManagerDesignerForm.bAddChildClick(Sender: TObject);
begin
  if SelectedAliasNode <> nil then
  begin
    AddAliasNode(SelectedAliasNode);
    ActiveControl := edName;
  end;
end;

procedure TvgrAliasManagerDesignerForm.bDeleteClick(Sender: TObject);
begin
  DeleteAliasNode(SelectedAliasNode);
  ActiveControl := TV;
end;

procedure TvgrAliasManagerDesignerForm.bCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TvgrAliasManagerDesignerForm.TVChange(Sender: TObject;
  Node: TTreeNode);
begin
  if SelectedAliasNode <> nil then
    CopyToControls(SelectedAliasNode);
end;

procedure TvgrAliasManagerDesignerForm.FormCreate(Sender: TObject);
begin
  //
  edVariableType.Items.Add(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_NodeValueTypeNull));
  edVariableType.Items.Add(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_NodeValueTypeInteger));
  edVariableType.Items.Add(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_NodeValueTypeExtended));
  edVariableType.Items.Add(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_NodeValueTypeString));
  edVariableType.Items.Add(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_NodeValueTypeDateTime));

  UpdateEnabled;
end;

procedure TvgrAliasManagerDesignerForm.edNameChange(Sender: TObject);
begin
  if EnableUpdate then
    SelectedNodeChanged := True;
end;

procedure TvgrAliasManagerDesignerForm.rbUnknownClick(Sender: TObject);
begin
  if EnableUpdate then
    SelectedNodeChanged := True;
end;

procedure TvgrAliasManagerDesignerForm.rbCustomClick(Sender: TObject);
begin
  if EnableUpdate then
  begin
    SelectedNodeChanged := True;
    ActiveControl := edCustom;
  end;
end;

procedure TvgrAliasManagerDesignerForm.rbVariableClick(Sender: TObject);
begin
  if EnableUpdate then
  begin
    SelectedNodeChanged := True;
    ActiveControl := edVariableValue;
  end;
end;

procedure TvgrAliasManagerDesignerForm.edCustomChange(Sender: TObject);
begin
  if EnableUpdate then
  begin
    SelectedNodeChanged := True;
    rbCustom.Checked := True;
  end;
end;

{$HINTS OFF}
procedure TvgrAliasManagerDesignerForm.edVariableValueChange(
  Sender: TObject);
var
  AValue, ACode: Integer;
  AExtValue: Extended;
  AExtDateTime: TDateTime;
begin
  if EnableUpdate then
  begin
    SelectedNodeChanged := True;
    val(edVariableValue.Text, AValue, ACode);
    if ACode = 0 then
      edVariableType.ItemIndex := Integer(rvtInteger)
    else
      if TextToFloat(PChar(edVariableValue.Text), AExtValue, fvExtended) then
        edVariableType.ItemIndex := Integer(rvtExtended)
      else
        if TryStrToDateTime(edVariableValue.Text, AExtDateTime) then
          edVariableType.ItemIndex := Integer(rvtDateTime)
        else
          edVariableType.ItemIndex := Integer(rvtString);
  end;
end;
{$HINTS ON}

procedure TvgrAliasManagerDesignerForm.bCancelClick(Sender: TObject);
begin
  if EnableUpdate and (SelectedAliasNode <> nil) then
  begin
    CopyToControls(SelectedAliasNode);
    ActiveControl := TV;
  end;
end;

procedure TvgrAliasManagerDesignerForm.bApplyClick(Sender: TObject);
var
  S: string;
begin
  if SelectedAliasNode <> nil then
  begin
    // Check node name
    S := Trim(edName.Text);
    if S = '' then
    begin
      MBox(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_ErrorNodeNameEmpty), MB_OK or MB_ICONERROR);
      ActiveControl := edName;
      exit;
    end;

    if not CheckUName(SelectedAliasNode, S) then
    begin
      MBox(vgrLoadStr(svgrid_vgr_AliasManagerDesigner_ErrorNodeNameNotUnique), MB_OK or MB_ICONERROR);
      ActiveControl := edName;
      exit;
    end;

    CopyFromControls(SelectedAliasNode);
    UpdateNode(SelectedAliasNode, TV.Selected);
    SelectedNodeChanged := False;
    ActiveControl := TV;
  end;
end;

procedure TvgrAliasManagerDesignerForm.FormKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_INSERT) and (Shift = []) then
  begin
    bAddClick(nil);
    Key := 0;
  end
  else
  if (Key = VK_INSERT) and (Shift = [ssAlt]) then
  begin
    bAddChildClick(nil);
    Key := 0;
  end
  else
  if (ActiveControl = TV) and (Key = VK_DELETE) and (Shift = []) then
  begin
    bDeleteClick(nil);
    Key := 0;
  end;
end;

procedure TvgrAliasManagerDesignerForm.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13  then
  begin
    if (ActiveControl = TV) and (SelectedAliasNode <> nil) then
    begin
      ActiveControl := edName;
      Key := #0;
    end
    else
      if (ActiveControl = edName) and (SelectedAliasNode <> nil) and (not SelectedAliasNode.HasChildren) then
      begin
        if rbUnknown.Checked then
        begin
          bApplyClick(nil);
          Key := #0;
        end
        else
          if rbCustom.Checked then
          begin
            ActiveControl := edCustom;
            Key := #0;
          end
          else
            if rbVariable.Checked then
            begin
              ActiveControl := edVariableValue;
              Key := #0;
            end;
      end
      else
        if (ActiveControl = rbUnknown) or (ActiveControl = rbCustom) or
           (ActiveControl = edCustom) or (ActiveControl = rbVariable) or
           (ActiveControl = edVariableValue) or (ActiveControl = edVariableType) then
          begin
            if SelectedNodeChanged then
              bApplyClick(nil)
            else
              ActiveControl := TV;
            Key := #0;
          end;
  end;
end;

procedure TvgrAliasManagerDesignerForm.TVEdited(Sender: TObject;
  Node: TTreeNode; var S: String);
var
  AAliasNode: TvgrAliasNode;
begin
  AAliasNode := TvgrAliasNode(Node.Data);
  if (S = '') or not CheckUName(AAliasNode, S) then
    S := AAliasNode.Name
  else
  begin
    AAliasNode.Name := S;
    CopyToControls(AAliasNode);
  end;
end;

procedure TvgrAliasManagerDesignerForm.bOptionsClick(Sender: TObject);
begin
  with Panel1.ClientToScreen(Point(bOptions.Left, bOptions.Top + bOptions.Height)) do
    PMOptions.Popup(X, Y);
end;

procedure TvgrAliasManagerDesignerForm.PMOptionsPopup(Sender: TObject);
var
  I: Integer;
begin
  mShowAllComponents.Enabled := AvailableComponentsProvider <> nil;
  mShowDatasetsOnly.Enabled := AvailableComponentsProvider <> nil;
  I := edCustom.Items.IndexOf(edCustom.Text);
  mImport.Enabled := (I <> -1) and (AMImportWizards.FindWizard(TObject(edCustom.Items.Objects[I]).ClassType) <> nil);

  mShowAllComponents.Checked := not (vgrcpoDatasetsOnly in AvailableComponentsProvider.Options);
  mShowDatasetsOnly.Checked := vgrcpoDatasetsOnly in AvailableComponentsProvider.Options;
end;

procedure TvgrAliasManagerDesignerForm.mShowAllComponentsClick(
  Sender: TObject);
begin
  AvailableComponentsProvider.Options := AvailableComponentsProvider.Options - [vgrcpoDatasetsOnly];
  ActiveControl := edCustom;
end;

procedure TvgrAliasManagerDesignerForm.mShowDatasetsOnlyClick(
  Sender: TObject);
begin
  AvailableComponentsProvider.Options := AvailableComponentsProvider.Options + [vgrcpoDatasetsOnly];
  ActiveControl := edCustom;
end;

procedure TvgrAliasManagerDesignerForm.mImportClick(Sender: TObject);
var
  I: Integer;
  AWizard: TvgrAliasManagerImportWizardClass;
  AObject: TObject;
begin
  I := edCustom.Items.IndexOf(edCustom.Text);
  if I <> -1 then
  begin
    AObject := TObject(edCustom.Items.Objects[I]); 
    AWizard := AMImportWizards.FindWizard(AObject.ClassType);
    if AWizard <> nil then
    begin
      SelectedAliasNode.Nodes.DeleteChildren(SelectedAliasNode);
      AWizard.AddNodes(AObject, edCustom.Text, Self, SelectedAliasNode);
      SelectedNodeChanged := False;
      if SelectedAliasNode <> nil then
        CopyToControls(SelectedAliasNode);
    end;
  end;
end;

initialization

  AMImportWizards := TvgrAliasManagerImportWizards.Create(False);
  AMImportWizards.RegisterImportWizard(TvgrDatasetImportWizard);

finalization

  FreeAndNil(AMImportWizards);

end.





