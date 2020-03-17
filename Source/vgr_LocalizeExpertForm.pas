unit vgr_LocalizeExpertForm;

{$I vtk.inc}

interface

uses
//  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls,
  {$IFDEF VTK_D6_OR_D7} DesignIntf, {$ELSE} dsgnintf, {$ENDIF} ImgList, Buttons,
  ToolsApi, typinfo, menus,

  vgr_CommonClasses, vgr_Form, vgr_FormLocalizer;

resourcestring
  svgrid_vgr_LocalizeExpertForm_FormNotSelected = 'Form not selected';
  svgrid_vgr_LocalizeExpertForm_FileExists = 'The file exists, it will be updated';
  svgrid_vgr_LocalizeExpertForm_FileNotExists = 'The file not exists, it will be created';
  svgrid_vgr_LocalizeExpertForm_RCDirectoryNotExists = 'Directory of *.RC files not found';
  svgrid_vgr_LocalizeExpertForm_LocalizerNotExists = 'Component of class TvgrFormLocalizer not found, it will be created';
  svgrid_vgr_LocalizeExpertForm_LocalizerExists = 'Component of class TvgrFormLocalizer found, it will be updated';
  svgrid_vgr_LocalizeExpertForm_LocalizerNotOne = 'Please select TvgrFormLocalizer component, which will be updated';
  svgrid_vgr_LocalizeExpertForm_ErrorCreateLocalizer = 'Error while create TvgrFormLocalizer component.';
  svgrid_vgr_LocalizeExpertForm_NotExistsInLocalizerTree = 'This item not exists in localizer tree';

  svgrid_vgr_LocalizeExpertForm_ComponentNodeStatusBarText = 'Component: Name = {%s} Class = {%s}';
  svgrid_vgr_LocalizeExpertForm_ClassNodeStatusBarText = 'Class: %s';
  svgrid_vgr_LocalizeExpertForm_PropertyNodeStatusBarText = 'Property: %s, Value = {%s}';

const
  svgr_vgr_LocalizeExperForm_Caption = '%s - Localize expert';
  svgrStringProps = [tkString, tkLString, tkWString];

  sStateNotChecked = 1;
  sStateChecked = 2;
  sStateGrayed = 3;

  sIncFilePrefix = '{$I';
  sUnitPrefix = 'unit';
  sBaseIndexIdentName = 'svgrResStringsBase';
  sLastIndexIdentName = 'svgrResStringsLastIndex';

  sRCIncludeFilePrefix = '#include';
  svgrResStringPrefix = 'svgrid';

type
  /////////////////////////////////////////////////
  //
  // TvgrFormInterface
  //
  /////////////////////////////////////////////////
  TvgrFormInterface = class(TObject)
  private
    FFormEditor: IOTAFormEditor;
    FComponent: IOTAComponent;
    function GetUnitName: string;
  public
    constructor Create(AFormEditor: IOTAFormEditor; AComponent: IOTAComponent);
    property FormEditor: IOTAFormEditor read FFormEditor write FFormEditor;
    property Component: IOTAComponent read FComponent write FComponent;
    property UnitName: string read GetUnitName;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormInterfaceList
  //
  /////////////////////////////////////////////////
  TvgrFormInterfaceList = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrFormInterface;
  public
    property Items[Index: Integer]: TvgrFormInterface read GetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrLocalizeNode
  //
  /////////////////////////////////////////////////
  TvgrLocalizeNode = class(TObject)
  public
    function GetCaption: string; virtual; abstract;
    function GetStatusBarText: string; virtual; abstract;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrComponentNode
  //
  /////////////////////////////////////////////////
  TvgrComponentNode = class(TvgrLocalizeNode)
  private
    FComponent: IOTAComponent;
    function GetComponentName: string;
    function GetComponentHandle: TOTAHandle;
  public
    constructor Create(AComponent: IOTAComponent);
    function GetCaption: string; override;
    function GetStatusBarText: string; override;
    property Component: IOTAComponent read FComponent write FComponent;
    property ComponentName: string read GetComponentName;
    property ComponentHandle: TOTAHandle read GetComponentHandle;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrPropertyNode
  //
  /////////////////////////////////////////////////
  TvgrPropertyNode = class(TvgrLocalizeNode)
  private
    FPropName: string;
    FComponentNode: TvgrComponentNode;
    function GetPropValue: string;
    function GetComponent: IOTAComponent;
    function GetComponentHandle: TOTAHandle;
  public
    constructor Create(const APropName: string; AComponentNode: TvgrComponentNode);
    function GetCaption: string; override;
    function GetStatusBarText: string; override;
    property PropName: string read FPropName;
    property ComponentNode: TvgrComponentNode read FComponentNode;
    property Component: IOTAComponent read GetComponent;
    property PropValue: string read GetPropValue;
    property ComponentHandle: TOTAHandle read GetComponentHandle;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrClassNode
  //
  /////////////////////////////////////////////////
  TvgrClassNode = class(TvgrLocalizeNode)
  private
    FClassNameStr: string;
  public
    constructor Create(const AClassNameStr: string);
    function GetCaption: string; override;
    function GetStatusBarText: string; override;
    property ClassNameStr: string read FClassNameStr;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrStringIDFileItem
  //
  /////////////////////////////////////////////////
  TvgrStringIDFileItem = class(TObject)
  private
    FIdent: string;
    FValue: Integer;
    FValueIdent: string;
  public
    constructor Create(const AIdent: string; AValue: Integer; const AValueIdent: string);
    property Ident: string read FIdent;
    property Value: Integer read FValue;
    property ValueIdent: string read FValueIdent;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrStringIDFile
  //
  /////////////////////////////////////////////////
  TvgrStringIDFile = class(TvgrObjectList)
  private
    FBaseIndex: Integer;
    FLastIndex: Integer;
    FIncFileName: string;
    FUnitName: string;
    function GetItem(Index: Integer): TvgrStringIDFileItem;
  public
    procedure LoadFromStream(AStream: TStream);
    procedure SaveToStream(AStream: TStream);
    procedure LoadFromFile(const AFileName: string);
    procedure SaveToFile(const AFileName: string);
    function IndexByIdent(const AIdent: string): Integer;
    function FindByIdent(const AIdent: string): TvgrStringIDFileItem;
    function AddNewItem(const AIdent: string): TvgrStringIDFileItem;

    property Items[Index: Integer]: TvgrStringIDFileItem read GetItem; default;
    property BaseIndex: Integer read FBaseIndex write FBaseIndex;
    property LastIndex: Integer read FLastIndex write FLastIndex;
    property IncFileName: string read FIncFileName write FIncFileName;
    property UnitName: string read FUnitName write FUnitName;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRCFileItem
  //
  /////////////////////////////////////////////////
  TvgrRCFileItem = class(TObject)
  private
    FIdent: string;
    FValue: string;
    FStringID: TvgrStringIDFileItem;
  public
    constructor Create(const AIdent: string; const AValue: string; AStringID: TvgrStringIDFileItem);
    property Ident: string read FIdent write FIdent;
    property Value: string read FValue write FValue;
    property StringID: TvgrStringIDFileItem read FStringID write FStringID;  
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRCFile
  //
  /////////////////////////////////////////////////
  TvgrRCFile = class(TvgrObjectList)
  private
    FIncludeFile: string;
    function GetItem(Index: Integer): TvgrRCFileItem;
  public
    procedure LoadFromStream(AStream: TStream; AStringIDFile: TvgrStringIDFile);
    procedure SaveToStream(AStream: TStream);
    procedure LoadFromFile(const AFileName: string; AStringIDFile: TvgrStringIDFile);
    procedure SaveToFile(const AFileName: string);

    function IndexByIdent(const AIdent: string): Integer;
    function FindByIdent(const AIdent: string): TvgrRCFileItem;
    function AddNewItem(const AIdent, AValue: string; AStringID: TvgrStringIDFileItem): TvgrRCFileItem;

    property Items[Index: Integer]: TvgrRCFileItem read GetItem; default;
    property IncludeFile: string read FIncludeFile write FIncludeFile;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormLocalizerInterface
  //
  /////////////////////////////////////////////////
  TvgrFormLocalizerInterface = class(TObject)
  private
    FComponent: IOTAComponent;
    function GetName: string;
  public
    constructor Create(AComponent: IOTAComponent); 
    property Component: IOTAComponent read FComponent;
    property Name: string read GetName;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFormLocalizerList
  //
  /////////////////////////////////////////////////
  TvgrFormLocalizerList = class(TvgrObjectList)
  private
    function GetItem(Index: Integer): TvgrFormLocalizerInterface;
  public
    property Items[Index: Integer]: TvgrFormLocalizerInterface read GetItem; default;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrLocalizerDeleteItem
  //
  /////////////////////////////////////////////////
  TvgrLocalizerDeleteItem = class(TObject)
  private
    FDelete: Boolean;
    FDeleteReason: string;
    FDeleteObject: TObject;
  protected
    function GetDescription: string; virtual;
  public
    constructor Create(ADeleteObject: TObject; ADelete: Boolean; const ADeleteReason: string);
    property Description: string read GetDescription;
    property Delete: Boolean read FDelete write FDelete;
    property DeleteReason: string read FDeleteReason write FDeleteReason;
    property DeleteObject: TObject read FDeleteObject write FDeleteObject;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRCDeleteItem
  //
  /////////////////////////////////////////////////
  TvgrRCDeleteItem = class(TvgrLocalizerDeleteItem)
  private
    function GetRCItem: TvgrRCFileItem;
  protected
    function GetDescription: string; override;
  public
    property RCItem: TvgrRCFileItem read GetRCItem;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrFLDeleteItem
  //
  /////////////////////////////////////////////////
  TvgrFLDeleteItem = class(TvgrLocalizerDeleteItem)
  private
    function GetFLItem: TvgrFormLocalizerItem;
  protected
    function GetDescription: string; override;
  public
    property FLItem: TvgrFormLocalizerItem read GetFLItem;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrLocalizerBuildReport
  //
  /////////////////////////////////////////////////
  TvgrLocalizerBuildReport = class(TObject)
  private
    FRCItemsToDelete: TvgrObjectList;
    FFLItemsToDelete: TvgrObjectList;
    function GetRCToDeleteItem(Index: Integer): TvgrRCDeleteItem;
    function GetRCToDeleteItemCount: Integer;
    function GetFLToDeleteItem(Index: Integer): TvgrFLDeleteItem;
    function GetFLToDeleteItemCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;

    function AddRCItem(ARCItem: TvgrRCFileItem; ADelete: Boolean; const ADeleteReason: string): TvgrRCDeleteItem;
    function AddFLItem(AFLItem: TvgrFormLocalizerItem; ADelete: Boolean; const ADeleteReason: string): TvgrFLDeleteItem;

    property RCList: TvgrObjectList read FRCItemsToDelete;
    property FLList: TvgrObjectList read FFLItemsToDelete;
    property RCToDeleteItems[Index: Integer]: TvgrRCDeleteItem read GetRCToDeleteItem;
    property RCToDeleteItemCount: Integer read GetRCToDeleteItemCount; 
    property FLToDeleteItems[Index: Integer]: TvgrFLDeleteItem read GetFLToDeleteItem;
    property FLToDeleteItemCount: Integer read GetFLToDeleteItemCount; 
  end;

  /////////////////////////////////////////////////
  //
  // TvgrModule
  //
  /////////////////////////////////////////////////
  TvgrModule = class(TObject, IOTAModuleCreator, IOTAFile)
  private
    FFileName: string;
    FSourceData: TObject;
  protected
    { IUnknown }
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    { IOTAFile }
    function GetSource: string; virtual;  
    function GetAge: TDateTime;

    { IOTACreator }
    function GetCreatorType: string; virtual;
    function GetExisting: Boolean;
    function GetFileSystem: string;
    function GetOwner: IOTAModule;
    function GetUnnamed: Boolean;

    { IOTAModuleCreator }
    function GetAncestorName: string;
    function GetImplFileName: string;
    function GetIntfFileName: string;
    function GetFormName: string;
    function GetMainForm: Boolean;
    function GetShowForm: Boolean;
    function GetShowSource: Boolean;
    function NewFormFile(const FormIdent, AncestorIdent: string): IOTAFile;
    function NewImplSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile;
    function NewIntfSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile;
    procedure FormCreated(const FormEditor: IOTAFormEditor);

    function GetCurrentEditor(AModule: IOTAModule): IOTAEditor;
  public
    constructor Create(const AFileName: string; ASourceData: TObject);
    procedure ShowText;
    property FileName: string read FFileName;
    property SourceData: TObject read FSourceData;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRCFileModule
  //
  /////////////////////////////////////////////////
  TvgrRCFileModule = class(TvgrModule)
  private
    function GetRCFile: TvgrRCFile;
  protected
    function GetCreatorType: string; override;
    function GetSource: string; override;
  public
    property RCFile: TvgrRCFile read GetRCFile;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrStringIDFileModule
  //
  /////////////////////////////////////////////////
  TvgrStringIDFileModule = class(TvgrModule)
  private
    function GetStringIDFile: TvgrStringIDFile;
  protected
    function GetCreatorType: string; override;
    function GetSource: string; override;
  public
    property StringIDFile: TvgrStringIDFile read GetStringIDFile;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrLocalizeExpertForm
  //
  /////////////////////////////////////////////////
  TvgrLocalizeExpertForm = class(TvgrDialogForm)
    Panel1: TPanel;
    Splitter: TSplitter;
    Panel2: TPanel;
    TV: TTreeView;
    Label1: TLabel;
    edForm: TComboBox;
    Panel3: TPanel;
    Panel4: TPanel;
    PageControl: TPageControl;
    PMain: TTabSheet;
    Label2: TLabel;
    Label3: TLabel;
    edRCFileName: TEdit;
    edLocalizerComponent: TComboBox;
    POptions: TTabSheet;
    Label4: TLabel;
    Label5: TLabel;
    edDefaultRCFilesDir: TEdit;
    edStringIdentsFile: TEdit;
    PClasses: TTabSheet;
    rbProcessAnyClasses: TRadioButton;
    bBuild: TButton;
    bApply: TButton;
    rbProcessSpecifiedClasses: TRadioButton;
    mClasses: TMemo;
    PProperties: TTabSheet;
    rbProcessAnyProperties: TRadioButton;
    rbProcessSpecifiedProperties: TRadioButton;
    mProperties: TMemo;
    bStringIdentsFile: TSpeedButton;
    ImageList: TImageList;
    lRCFileNameDesc: TLabel;
    lLocalizerComponentDesc: TLabel;
    SaveDialog: TSaveDialog;
    rgAfterBuild: TRadioGroup;
    cbShowNotActions: TCheckBox;
    cbNotShowEmpty: TCheckBox;
    StatusBar: TStatusBar;
    cbNotShowSeparator: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure edFormClick(Sender: TObject);
    procedure TVMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TVDeletion(Sender: TObject; Node: TTreeNode);
    procedure edDefaultRCFilesDirChange(Sender: TObject);
    procedure edStringIdentsFileChange(Sender: TObject);
    procedure rbProcessAnyClassesClick(Sender: TObject);
    procedure mClassesChange(Sender: TObject);
    procedure rbProcessAnyPropertiesClick(Sender: TObject);
    procedure mPropertiesChange(Sender: TObject);
    procedure bApplyClick(Sender: TObject);
    procedure bStringIdentsFileClick(Sender: TObject);
    procedure bBuildClick(Sender: TObject);
    procedure cbShowNotActionsClick(Sender: TObject);
    procedure cbNotShowEmptyClick(Sender: TObject);
    procedure TVChange(Sender: TObject; Node: TTreeNode);
  private
    { Private declarations }
    FForms: TvgrFormInterfaceList;
    FLocalizers: TvgrFormLocalizerList;
    FRCFile: TvgrRCFile;
    FStringIDFile: TvgrStringIDFile;
    procedure UpdateForms;
    procedure UpdateTV(AForm: TvgrFormInterface);
    procedure UpdateRCFileName;
    procedure UpdatelLocalizerComponent;
    function GetSelectedForm: TvgrFormInterface;
    function GetSelectedLocalizer: TvgrFormLocalizerInterface;
    procedure ApplyOptions;
    procedure Build;
    function GetPropertyStringResourceID(APropertyNode: TvgrPropertyNode): string;
  protected
    procedure RestoreSettings; override;
    procedure SaveSettings; override;
  public
    { Public declarations }
    procedure Execute;

    property SelectedForm: TvgrFormInterface read GetSelectedForm;
    property SelectedLocalizer: TvgrFormLocalizerInterface read GetSelectedLocalizer;
  end;

implementation

uses
  vgr_Functions, vgr_GUIFunctions, vgr_Localize, vgr_LocalizerExpertReportForm;

{$R *.dfm}

/////////////////////////////////////////////////
//
// TvgrFormInterface
//
/////////////////////////////////////////////////
constructor TvgrFormInterface.Create(AFormEditor: IOTAFormEditor; AComponent: IOTAComponent);
begin
  inherited Create;
  FFormEditor := AFormEditor;
  FComponent := AComponent;
end;

function TvgrFormInterface.GetUnitName: string;
var
  I: Integer;
begin
  Result := ExtractFileName(FormEditor.FileName);
  I := Length(Result);
  while (I > 0) and (Result[I] <> '.') do Dec(I);
  if I > 0 then
    Delete(Result, I, Length(Result));
end;

/////////////////////////////////////////////////
//
// TvgrFormInterfaceList
//
/////////////////////////////////////////////////
function TvgrFormInterfaceList.GetItem(Index: Integer): TvgrFormInterface;
begin
  Result := TvgrFormInterface(inherited Items[Index]);
end;

/////////////////////////////////////////////////
//
// TvgrLocalizeNode
//
/////////////////////////////////////////////////

/////////////////////////////////////////////////
//
// TvgrComponentNode
//
/////////////////////////////////////////////////
constructor TvgrComponentNode.Create(AComponent: IOTAComponent);
begin
  inherited Create;
  FComponent := AComponent;
end;

function TvgrComponentNode.GetCaption: string;
var
  S: string;
begin
  if FComponent.GetPropValueByName('Name', S) then
    Result := S;
end;

function TvgrComponentNode.GetStatusBarText: string;
begin
  Result := Format(svgrid_vgr_LocalizeExpertForm_ComponentNodeStatusBarText, [ComponentName, Component.GetComponentType]);
end;

function TvgrComponentNode.GetComponentName: string;
begin
  if FComponent = nil then
    Result := ''
  else
  begin
    if not FComponent.GetPropValueByName('Name', Result) then
      Result := '';
  end;
end;

function TvgrComponentNode.GetComponentHandle: TOTAHandle;
begin
  if FComponent = nil then
    Result := nil
  else
    Result := FComponent.GetComponentHandle;
end;

/////////////////////////////////////////////////
//
// TvgrPropertyNode
//
/////////////////////////////////////////////////
constructor TvgrPropertyNode.Create(const APropName: string; AComponentNode: TvgrComponentNode);
begin
  inherited Create;
  FPropName := APropName;
  FComponentNode := AComponentNode;
end;

function TvgrPropertyNode.GetComponent: IOTAComponent;
begin
  Result := FComponentNode.Component;
end;

function TvgrPropertyNode.GetComponentHandle: TOTAHandle;
begin
  if FComponentNode = nil then
    Result := nil
  else
    Result := ComponentNode.ComponentHandle;
end;

function TvgrPropertyNode.GetPropValue: string;
begin
  if Component = nil then
    Result := ''
  else
  begin
    if not Component.GetPropValueByName(PropName, Result) then
      Result := '';
  end;
end;

function TvgrPropertyNode.GetCaption: string;
begin
  Result := FPropName;
end;

function TvgrPropertyNode.GetStatusBarText: string;
begin
  Result := Format(svgrid_vgr_LocalizeExpertForm_PropertyNodeStatusBarText, [PropName, PropValue]);
end;

/////////////////////////////////////////////////
//
// TvgrClassNode
//
/////////////////////////////////////////////////
constructor TvgrClassNode.Create(const AClassNameStr: string);
begin
  inherited Create;
  FClassNameStr := AClassNameStr;
end;

function TvgrClassNode.GetCaption: string;
begin
  Result := FClassNameStr;
end;

function TvgrClassNode.GetStatusBarText: string;
begin
  Result := Format(svgrid_vgr_LocalizeExpertForm_ClassNodeStatusBarText, [FClassNameStr]);
end;

/////////////////////////////////////////////////
//
// TvgrStringIDFileItem
//
/////////////////////////////////////////////////
constructor TvgrStringIDFileItem.Create(const AIdent: string; AValue: Integer; const AValueIdent: string);
begin
  inherited Create;
  FIdent := AIdent;
  FValue := AValue;
  FValueIdent := AValueIdent;
end;

/////////////////////////////////////////////////
//
// TvgrStringIDFile
//
/////////////////////////////////////////////////
function TvgrStringIDFile.GetItem(Index: Integer): TvgrStringIDFileItem;
begin
  Result := TvgrStringIDFileItem(inherited Items[Index]);
end;

procedure TvgrStringIDFile.LoadFromStream(AStream: TStream);
var
  ABuf: string;
  AIdent, AValueIdent, S: string;
  P: Integer;
  APos, AValue, AMaxValue, ACode: Integer;
begin
  Clear;
  FBaseIndex := 0;
  FLastIndex := 0;
  FIncFileName := '';
  FUnitName := '';

  SetLength(ABuf, AStream.Size - AStream.Position);
  AStream.Read(ABuf[1], Length(ABuf));

  AMaxValue := 0;
  APos := 1;
  while APos <= Length(ABuf) do
  begin
    S := Trim(ExtractLine(ABuf, APos));
    
    if AnsiCompareText(Copy(S, 1, Length(sIncFilePrefix)), sIncFilePrefix) = 0 then
      FIncFileName := Trim(Copy(S, Length(sIncFilePrefix) + 1, Length(S) - Length(sIncFilePrefix) - 1))
    else
    begin
      if AnsiCompareText(Copy(S, 1, Length(sUnitPrefix)), sUnitPrefix) = 0 then
        FUnitName := Trim(Copy(S, Length(sUnitPrefix) + 1, Length(S) - Length(sUnitPrefix) - 1))
      else
      begin
        P := pos('=', S);
        if P <> 0 then
        begin
          AIdent := Trim(Copy(S, 1, P - 1));
          S := Trim(Copy(S, P + 1, Length(S) - P - 1));
          val(S, AValue, ACode);
          if ACode <> 0 then
          begin
            AValueIdent := S;
            AValue := 0;
          end
          else
            AValueIdent := '';

          if AnsiCompareText(AIdent, sBaseIndexIdentName) = 0 then
            FBaseIndex := AValue
          else
            if AnsiCompareText(AIdent, sLastIndexIdentName) = 0 then
              FLastIndex := AValue
            else
            begin
              Add(TvgrStringIDFileItem.Create(AIdent, AValue, AValueIdent));
              if AMaxValue < AValue then
                AMaxValue := AValue;
            end;
        end;
      end;
    end;
  end;
  if FLastIndex < AMaxValue then
    FLastIndex := AMaxValue;
end;

procedure TvgrStringIDFile.SaveToStream(AStream: TStream);
var
  I: Integer;
begin
  WriteLnString(AStream, Format('%s %s;', [sUnitPrefix, UnitName]));
  WriteLnString(AStream, '');
  WriteLnString(AStream, Format('%s %s}', [sIncFilePrefix, IncFileName]));
  WriteLnString(AStream, '');
  WriteLnString(AStream, 'interface');
  WriteLnString(AStream, '');
  WriteLnString(AStream, 'const');
  WriteLnString(AStream, '');
  WriteLnString(AStream, Format('  %s = %d;', [sBaseIndexIdentName, BaseIndex]));
  WriteLnString(AStream, Format('  %s = %d;', [sLastIndexIdentName, LastIndex]));
  WriteLnString(AStream, '');
  WriteLnString(AStream, '');

  for I := 0 to Count - 1 do
    with Items[I] do
      if ValueIdent <> '' then
        WriteLnString(AStream, Format('  %s = %s;', [Ident, ValueIdent]))
      else
        WriteLnString(AStream, Format('  %s = %d;', [Ident, Value]));

  WriteLnString(AStream, '');
  WriteLnString(AStream, 'implementation');
  WriteLnString(AStream, '');
  WriteLnString(AStream, 'end.');
end;

procedure TvgrStringIDFile.LoadFromFile(const AFileName: string);
var
  AFileStream: TFileStream;
begin
  Clear;
  FBaseIndex := 0;
  FLastIndex := 0;
  FIncFileName := '';
  FUnitName := '';

  AFileStream := nil;
  try
    AFileStream := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
    LoadFromStream(AFileStream);
    AFileStream.Free;
  except
    if AFileStream <> nil then
      AFileStream.Free;
  end;
end;

procedure TvgrStringIDFile.SaveToFile(const AFileName: string);
var
  AFileStream: TFileStream;
begin
  AFileStream := nil;
  try
    AFileStream := TFileStream.Create(AFileName, fmCreate or fmShareDenyNone);
    SaveToStream(AFileStream);
    AFileStream.Free;
  except
    if AFileStream <> nil then
      AFileStream.Free;
  end;
end;

function TvgrStringIDFile.IndexByIdent(const AIdent: string): Integer;
begin
  Result := 0;
  while (Result < Count) and (AnsiCompareText(Items[Result].Ident, AIdent) <> 0) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrStringIDFile.FindByIdent(const AIdent: string): TvgrStringIDFileItem;
var
  I: Integer;
begin
  I := IndexByIdent(AIdent);
  if I = -1 then
    Result := nil
  else
    Result := Items[I];
end;

function TvgrStringIDFile.AddNewItem(const AIdent: string): TvgrStringIDFileItem;
begin
  Result := TvgrStringIDFileItem.Create(AIdent, FLastIndex + 1, '');
  FLastIndex := FLastIndex + 1;
  Add(Result);
end;

/////////////////////////////////////////////////
//
// TvgrRCFileItem
//
/////////////////////////////////////////////////
constructor TvgrRCFileItem.Create(const AIdent: string; const AValue: string; AStringID: TvgrStringIDFileItem);
begin
  inherited Create;
  FIdent := AIdent;
  FValue := AValue;
  FStringID := AStringID;
end;

/////////////////////////////////////////////////
//
// TvgrRCFile
//
/////////////////////////////////////////////////
function TvgrRCFile.GetItem(Index: Integer): TvgrRCFileItem;
begin
  Result := TvgrRCFileItem(inherited Items[Index]);
end;

procedure TvgrRCFile.LoadFromStream(AStream: TStream; AStringIDFile: TvgrStringIDFile);
var
  APos, P: Integer;
  S, ABuf, AValue, AIdent: string;
begin
  Clear;
  FIncludeFile := '';
  SetLength(ABuf, AStream.Size - AStream.Position);
  AStream.Read(ABuf[1], Length(ABuf));

  APos := 1;
  while APos <= Length(ABuf) do
  begin
    S := Trim(ExtractLine(ABuf, APos));

    if AnsiCompareText(Copy(S, 1, Length(sRCIncludeFilePrefix)), sRCIncludeFilePrefix) = 0 then
    begin
      P := Length(sRCIncludeFilePrefix) + 1;
      while (P <= Length(S)) and (S[P] <> '"') do Inc(P);
      Inc(P);
      FIncludeFile := Trim(Copy(S, P, Length(S) - P))
    end
    else
    begin
      P := pos(',', S);
      if P <> 0 then
      begin
        AIdent := Trim(Copy(S, 1, P - 1));
        AValue := Trim(Copy(S, P + 1, Length(S)));

        P := pos('+', AIdent);
        if P <> 0 then
          AIdent := Trim(Copy(AIdent, P + 1, Length(AIdent)));

        if (Length(AValue) > 0) and (AValue[1] = '"') then
          System.Delete(AValue, 1, 1);
        if (Length(AValue) > 0) and (AValue[Length(AValue)] = '"') then
          System.Delete(AValue, Length(AValue), 1);

        Add(TvgrRCFileItem.Create(AIdent, AValue, AStringIDFile.FindByIdent(AIdent)));
      end;
    end;
  end;
end;

procedure TvgrRCFile.SaveToStream(AStream: TStream);
var
  I: Integer;
begin
  WriteLnString(AStream, Format('%s "%s"', [sRCIncludeFilePrefix, IncludeFile]));
  WriteLnString(AStream, '');
  WriteLnString(AStream, 'STRINGTABLE');
  WriteLnString(AStream, '{');

  for I := 0 to Count - 1 do
    with Items[I] do
      WriteLnString(AStream, Format('  %s + %s, "%s"', [sBaseIndexIdentName, Ident, Value]));

  WriteLnString(AStream, '}');
end;

procedure TvgrRCFile.LoadFromFile(const AFileName: string; AStringIDFile: TvgrStringIDFile);
var
  AFileStream: TFileStream;
begin
  Clear;
  FIncludeFile := '';
  AFileStream := nil;
  try
    AFileStream := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
    LoadFromStream(AFileStream, AStringIDFile);
    AFileStream.Free;
  except
    if AFileStream <> nil then
      AFileStream.Free;
  end;
end;

procedure TvgrRCFile.SaveToFile(const AFileName: string);
var
  AFileStream: TFileStream;
begin
  AFileStream := nil;
  try
    AFileStream := TFileStream.Create(AFileName, fmCreate or fmShareDenyNone);
    SaveToStream(AFileStream);
    AFileStream.Free;
  except
    if AFileStream <> nil then
      AFileStream.Free;
  end;
end;

function TvgrRCFile.IndexByIdent(const AIdent: string): Integer;
begin
  Result := 0;
  while (Result < Count) and
        (AnsiCompareText(Items[Result].Ident, AIdent) <> 0) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

function TvgrRCFile.FindByIdent(const AIdent: string): TvgrRCFileItem;
var
  I: Integer;
begin
  I := IndexByIdent(AIdent);
  if I <> -1 then
    Result := Items[I]
  else
    Result := nil;
end;

function TvgrRCFile.AddNewItem(const AIdent, AValue: string; AStringID: TvgrStringIDFileItem): TvgrRCFileItem;
begin
  Result := TvgrRCFileItem.Create(AIdent, AValue, AStringID);
  Add(Result);
end;

/////////////////////////////////////////////////
//
// TvgrFormLocalizerInterface
//
/////////////////////////////////////////////////
constructor TvgrFormLocalizerInterface.Create(AComponent: IOTAComponent);
begin
  inherited Create;
  FComponent := AComponent;
end;

function TvgrFormLocalizerInterface.GetName: string;
begin
  if FComponent = nil then
    Result := ''
  else
    if not FComponent.GetPropValueByName('Name', Result) then
      Result := '';
end;

/////////////////////////////////////////////////
//
// TvgrFormLocalizerList
//
/////////////////////////////////////////////////
function TvgrFormLocalizerList.GetItem(Index: Integer): TvgrFormLocalizerInterface;
begin
  Result := TvgrFormLocalizerInterface(inherited Items[Index]);
end;

/////////////////////////////////////////////////
//
// TvgrLocalizerDeleteItem
//
/////////////////////////////////////////////////
constructor TvgrLocalizerDeleteItem.Create(ADeleteObject: TObject; ADelete: Boolean; const ADeleteReason: string);
begin
  inherited Create;
  FDeleteObject := ADeleteObject;
  FDelete := ADelete;
  FDeleteReason := ADeleteReason;
end;

function TvgrLocalizerDeleteItem.GetDescription: string;
begin
  Result := FDeleteObject.ClassName;
end;

/////////////////////////////////////////////////
//
// TvgrRCDeleteItem
//
/////////////////////////////////////////////////
function TvgrRCDeleteItem.GetRCItem: TvgrRCFileItem;
begin
  Result := TvgrRCFileItem(inherited DeleteObject);
end;

function TvgrRCDeleteItem.GetDescription: string;
begin
  Result := RCItem.Ident;
end;

/////////////////////////////////////////////////
//
// TvgrFLDeleteItem
//
/////////////////////////////////////////////////
function TvgrFLDeleteItem.GetFLItem: TvgrFormLocalizerItem;
begin
  Result := TvgrFormLocalizerItem(inherited DeleteObject);
end;

function TvgrFLDeleteItem.GetDescription: string;
begin
  with FLItem do
  begin
    if Component = nil then
      Result := Format('%s (%d)', [PropName, ResStringID])
    else
      Result := Format('%s.%s (%d)', [Component.ClassName, PropName, ResStringID]);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrLocalizerBuildReport
//
/////////////////////////////////////////////////
constructor TvgrLocalizerBuildReport.Create;
begin
  inherited Create;
  FRCItemsToDelete := TvgrObjectList.Create(True);
  FFLItemsToDelete := TvgrObjectList.Create(True);
end;

destructor TvgrLocalizerBuildReport.Destroy;
begin
  FRCItemsToDelete.Free;
  FFLItemsToDelete.Free;
end;

function TvgrLocalizerBuildReport.AddRCItem(ARCItem: TvgrRCFileItem; ADelete: Boolean; const ADeleteReason: string): TvgrRCDeleteItem;
begin
  Result := TvgrRCDeleteItem.Create(ARCItem, ADelete, ADeleteReason);
  FRCItemsToDelete.Add(Result);
end;

function TvgrLocalizerBuildReport.AddFLItem(AFLItem: TvgrFormLocalizerItem; ADelete: Boolean; const ADeleteReason: string): TvgrFLDeleteItem;
begin
  Result := TvgrFLDeleteItem.Create(AFLItem, ADelete, ADeleteReason);
  FFLItemsToDelete.Add(Result);
end;

function TvgrLocalizerBuildReport.GetRCToDeleteItem(Index: Integer): TvgrRCDeleteItem;
begin
  Result := TvgrRCDeleteItem(FRCItemsToDelete[Index]);
end;

function TvgrLocalizerBuildReport.GetRCToDeleteItemCount: Integer;
begin
  Result := FRCItemsToDelete.Count;
end;

function TvgrLocalizerBuildReport.GetFLToDeleteItem(Index: Integer): TvgrFLDeleteItem;
begin
  Result := TvgrFLDeleteItem(FFLItemsToDelete[Index]);
end;

function TvgrLocalizerBuildReport.GetFLToDeleteItemCount: Integer;
begin
  Result := FFLItemsToDelete.Count;
end;

/////////////////////////////////////////////////
//
// TvgrModule
//
/////////////////////////////////////////////////
constructor TvgrModule.Create(const AFileName: string; ASourceData: TObject);
begin
  inherited Create;
  FFileName := AFileName;
  FSourceData := ASourceData;
end;

function TvgrModule.GetCurrentEditor(AModule: IOTAModule): IOTAEditor;
{$IFNDEF VTK_D6_OR_D7}
var
  I: Integer;
{$ENDIF}
begin
{$IFNDEF VTK_D6_OR_D7}
  for I := 0 to AModule.GetModuleFileCount - 1 do
  begin
    Result := AModule.GetModuleFileEditor(I);
    if Result.QueryInterface(IOTASourceEditor, Result) = S_OK then
      exit;
  end;
  Result := nil;
{$ELSE}
  Result := AModule.GetCurrentEditor;
{$ENDIF}
end;

procedure TvgrModule.ShowText;
var
  AModuleServices: IOTAModuleServices;
  AModule: IOTAModule;
  AEditor: IOTAEditor;
  ASourceEditor: IOTASourceEditor;
  AEditWriter: IOTAEditWriter;
  APos: TOTACharPos;
begin
  if Supports(BorlandIDEServices, IOTAModuleServices, AModuleServices) then
  begin
    AModule := AModuleServices.FindModule(FileName);
    if AModule = nil then
    begin
      AModuleServices.CreateModule(IOTAModuleCreator(Self));
    end
    else
    begin
      // edit current module
      AEditor := GetCurrentEditor(AModule);
      if AEditor <> nil then
      begin
        ASourceEditor := AEditor as IOTASourceEditor;
        if ASourceEditor <> nil then
        begin
          APos.CharIndex := 0;
          APos.Line := 1;
          ASourceEditor.SetBlockStart(APos);
          AEditWriter := ASourceEditor.CreateWriter;
          AEditWriter.DeleteTo(MaxInt);
          AEditWriter.Insert(PChar(GetSource));
        end;
      end;
    end;
  end;
end;

function TvgrModule.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if GetInterface(IID, Obj) then
    Result := 0
  else
    Result := E_NOINTERFACE;
end;

function TvgrModule._AddRef: Integer;
begin
  Result := -1;
end;

function TvgrModule._Release: Integer;
begin
  Result := -1;
end;

function TvgrModule.GetSource: string;
begin
  Result := '';
end;

function TvgrModule.GetAge: TDateTime;
begin
  Result := -1; //Now;
end;

function TvgrModule.GetCreatorType: string;
begin
  Result := sUnit;
end;

function TvgrModule.GetExisting: Boolean;
begin
  Result := False;
end;

function TvgrModule.GetFileSystem: string;
begin
  Result := '';
end;

function TvgrModule.GetOwner: IOTAModule;
begin
  Result := nil;
end;

function TvgrModule.GetUnnamed: Boolean;
begin
  Result := False;  
end;

function TvgrModule.GetAncestorName: string;
begin
  Result := '';
end;

function TvgrModule.GetImplFileName: string;
begin
  Result := FileName;
end;

function TvgrModule.GetIntfFileName: string;
begin
  Result := '';
end;

function TvgrModule.GetFormName: string;
begin
  Result := '';
end;

function TvgrModule.GetMainForm: Boolean;
begin
  Result := False;
end;

function TvgrModule.GetShowForm: Boolean;
begin
  Result := False;
end;

function TvgrModule.GetShowSource: Boolean;
begin
  Result := True;
end;

function TvgrModule.NewFormFile(const FormIdent, AncestorIdent: string): IOTAFile;
begin
  Result := nil;
end;

function TvgrModule.NewImplSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile;
begin
  Result := Self;
end;

function TvgrModule.NewIntfSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile;
begin
  Result := nil;
end;

procedure TvgrModule.FormCreated(const FormEditor: IOTAFormEditor);
begin
end;

/////////////////////////////////////////////////
//
// TvgrRCFileModule
//
/////////////////////////////////////////////////
function TvgrRCFileModule.GetCreatorType: string;
begin
  Result := sText;
end;

function TvgrRCFileModule.GetSource: string;
var
  AMemoryStream: TMemoryStream;
begin
  AMemoryStream := TMemoryStream.Create;
  try
    RCFile.SaveToStream(AMemoryStream);
    SetLength(Result, AMemoryStream.Size);
    CopyMemory(@Result[1], AMemoryStream.Memory, Length(Result)); 
  finally
    AMemoryStream.Free;
  end;
end;

function TvgrRCFileModule.GetRCFile: TvgrRCFile;
begin
  Result := TvgrRCFile(inherited SourceData);
end;

/////////////////////////////////////////////////
//
// TvgrStringIDFileModule
//
/////////////////////////////////////////////////
function TvgrStringIDFileModule.GetCreatorType: string;
begin
  Result := sUnit;
end;

function TvgrStringIDFileModule.GetSource: string;
var
  AMemoryStream: TMemoryStream;
begin
  AMemoryStream := TMemoryStream.Create;
  try
    StringIDFile.SaveToStream(AMemoryStream);
    SetLength(Result, AMemoryStream.Size);
    CopyMemory(@Result[1], AMemoryStream.Memory, Length(Result)); 
  finally
    AMemoryStream.Free;
  end;
end;

function TvgrStringIDFileModule.GetStringIDFile: TvgrStringIDFile;
begin
  Result := TvgrStringIDFile(inherited SourceData);
end;

/////////////////////////////////////////////////
//
// TvgrLocalizeExpertForm
//
/////////////////////////////////////////////////
procedure TvgrLocalizeExpertForm.Execute;
begin
  ShowModal;
end;

procedure TvgrLocalizeExpertForm.UpdateForms;
var
  I, J: Integer;
  AModuleServices: IOTAModuleServices;
  AModule: IOTAModule;
  AFormEditor: IOTAFormEditor;
  AFormInterface: TvgrFormInterface;
begin
  FForms.Clear;
  edForm.Clear;
  
  if Supports(BorlandIDEServices, IOTAModuleServices, AModuleServices) then
  begin
    for I := 0 to AModuleServices.ModuleCount - 1 do
    begin
      AModule := AModuleServices.Modules[I];
      for J := 0 to AModule.GetModuleFileCount - 1 do
      begin
        if Supports(AModule.GetModuleFileEditor(J), IOTAFormEditor, AFormEditor) then
        begin
          AFormInterface := TvgrFormInterface.Create(AFormEditor, AFormEditor.GetRootComponent); 
          FForms.Add(AFormInterface);
          edForm.Items.AddObject(AFormInterface.Component.GetComponentType, AFormInterface);
        end;
      end;
    end;
  end;

  if edForm.Items.Count > 0 then
  begin
    edForm.ItemIndex := 0;
    UpdateTV(FForms[edForm.ItemIndex]);
  end;
end;

procedure TvgrLocalizeExpertForm.UpdateTV(AForm: TvgrFormInterface);
var
  I, J: Integer;
  ANode, AClassNode, AComponentNode: TTreeNode;
  AComponent: IOTAComponent;
  AComponentNodeData: TvgrComponentNode;
  AClassNodeData: TvgrClassNode;

  function IsAllowPropertyName(const APropName: string): Boolean;
  begin
    Result := (AnsiCompareText(APropName, 'Name') <> 0) and
              (rbProcessAnyProperties.Checked or
               (IndexInStrings(APropName, mProperties.Lines) <> -1));
  end;

  function IsAllowPropertyValue(AComponent: IOTAComponent; I: Integer): Boolean;
  var
    S: string;
  begin
    Result := AComponent.GetPropValue(I, S);
    if Result then
      Result := not cbNotShowEmpty.Checked or (S <> '');
  end;

  function IsAllowAddComponent(AComponent: IOTAComponent): Boolean;
  var
    I: Integer;
    ARealComponent: TComponent;
  begin
    Result := rbProcessAnyClasses.Checked or
              (IndexInStrings(AComponent.GetComponentType, mClasses.Lines) <> -1);
    if Result then
    begin
      ARealComponent := TComponent(AComponent.GetComponentHandle);
      Result := not (csAncestor in ARealComponent.ComponentState) and (not ((ARealComponent is TToolButton) and (TToolButton(ARealComponent).Style in [tbsSeparator, tbsDivider])));
      if Result then
      begin
        if cbShowNotActions.Checked then
        begin
          Result := not (((ARealComponent is TControl) and (TControl(ARealComponent).Action <> nil)) or
                         ((ARealComponent is TMenuItem) and (TMenuItem(ARealComponent).Action <> nil)));
        end;
        if Result then
        begin
          if cbNotShowSeparator.Checked then
          begin
            Result := not ((ARealComponent is TMenuItem) and (TMenuItem(ARealComponent).Caption = '-'));
          end;
          if Result then
          begin
            I := 0;
            while (I < AComponent.GetPropCount) and
                  not ((AComponent.GetPropType(I) in svgrStringProps) and
                       IsAllowPropertyName(AComponent.GetPropName(I)) and
                       IsAllowPropertyValue(AComponent, I)) do Inc(I);
            Result := I < AComponent.GetPropCount;
          end;
        end;
      end;
    end;
  end;
  
  procedure AddComponentProperties(AParentNode: TTreeNode; AComponentNodeData: TvgrComponentNode);
  var
    I: Integer;
    APropNodeData: TvgrPropertyNode;
  begin
    for I := 0 to AComponentNodeData.Component.GetPropCount - 1 do
      if (AComponentNodeData.Component.GetPropType(I) in svgrStringProps) and
         (IsAllowPropertyName(AComponentNodeData.Component.GetPropName(I)) and
          IsAllowPropertyValue(AComponentNodeData.Component, I)) then
      begin
        APropNodeData := TvgrPropertyNode.Create(AComponentNodeData.Component.GetPropName(I), AComponentNodeData);
        with TV.Items.AddChildObject(AParentNode, APropNodeData.GetCaption, APropNodeData) do
          StateIndex := sStateChecked;
      end;
  end;

begin
  TV.Items.BeginUpdate;
  try
    TV.Items.Clear;

    AComponentNodeData := TvgrComponentNode.Create(AForm.Component);
    ANode := TV.Items.AddObject(nil, AComponentNodeData.GetCaption, AComponentNodeData);
    ANode.StateIndex := sStateChecked;
    AddComponentProperties(ANode, AComponentNodeData);

    for I := 0 to AForm.Component.GetComponentCount - 1 do
    begin
      AComponent := AForm.Component.GetComponent(I);
      if IsAllowAddComponent(AComponent) then
      begin
        // find class node
        J := 0;
        while (J < TV.Items.Count) and
              not ((TObject(TV.Items[J].Data) is TvgrClassNode) and
                   (TvgrClassNode(TV.Items[J].Data).ClassNameStr = AComponent.GetComponentType)) do Inc(J);
        if J >= TV.Items.Count then
        begin
          AClassNodeData := TvgrClassNode.Create(AComponent.GetComponentType);
          AClassNode := TV.Items.AddChildObject(ANode, AClassNodeData.GetCaption, AClassNodeData);
          AClassNode.StateIndex := sStateChecked;
        end
        else
          AClassNode := TV.Items[J];
        AComponentNodeData := TvgrComponentNode.Create(AComponent);
        AComponentNode := TV.Items.AddChildObject(AClassNode, AComponentNodeData.GetCaption, AComponentNodeData);
        AComponentNode.StateIndex := sStateChecked;
        AddComponentProperties(AComponentNode, AComponentNodeData);
      end;
    end;

    ANode.Expand(False);
    TV.Selected := ANode;
  finally
    TV.Items.EndUpdate;
  end;
end;

procedure TvgrLocalizeExpertForm.FormCreate(Sender: TObject);
begin
  FForms := TvgrFormInterfaceList.Create;
  FLocalizers := TvgrFormLocalizerList.Create;
  FRCFile := TvgrRCFile.Create;
  FStringIDFile := TvgrStringIDFile.Create;

  UpdateForms;
  ApplyOptions;
end;

procedure TvgrLocalizeExpertForm.FormDestroy(Sender: TObject);
begin
  FForms.Free;
  FLocalizers.Free;
  FRCFile.Free;
  FStringIDFile.Free;
end;

procedure TvgrLocalizeExpertForm.edFormClick(Sender: TObject);
begin
  if edForm.ItemIndex >= 0 then
  begin
    UpdateTV(FForms[edForm.ItemIndex]);
    UpdateRCFileName;
    UpdatelLocalizerComponent;
    ActiveControl := TV;
  end;
end;

procedure TvgrLocalizeExpertForm.TVMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ANode: TTreeNode;
begin
  if (Button = mbLeft) and (Shift = [ssLeft]) and (htOnStateIcon in TV.GetHitTestInfoAt(X, Y)) then
  begin
    ANode := TV.GetNodeAt(X, Y);
    if ANode <> nil then
      ChangeStateOfTreeNode(ANode, sStateNotChecked, sStateChecked, sStateGrayed);
  end;
end;

procedure TvgrLocalizeExpertForm.UpdateRCFileName;
var
  AFileName: string;
begin
  AFileName := '';
  if SelectedForm = nil then
  begin
    edRCFileName.Text := '';
    lRCFileNameDesc.Caption := svgrid_vgr_LocalizeExpertForm_FormNotSelected;
  end
  else
  begin
    if DirectoryExists(edDefaultRCFilesDir.Text) then
    begin
      AFileName := Format('%s\%sStrings.rc',
                          [SysUtils.ExcludeTrailingBackslash(edDefaultRCFilesDir.Text),
                           SelectedForm.UnitName]);
      edRCFileName.Text := AFileName;
      if FileExists(AFileName) then
        lRCFileNameDesc.Caption := svgrid_vgr_LocalizeExpertForm_FileExists
      else
        lRCFileNameDesc.Caption := svgrid_vgr_LocalizeExpertForm_FileNotExists;
    end
    else
    begin
      edRCFileName.Text := '';
      lRCFileNameDesc.Caption := svgrid_vgr_LocalizeExpertForm_RCDirectoryNotExists;
    end;
  end;

  bBuild.Enabled := (SelectedForm <> nil) and (AFileName <> '');
end;

function TvgrLocalizeExpertForm.GetSelectedForm: TvgrFormInterface;
begin
  if (edForm.ItemIndex >= 0) and (edForm.ItemIndex < FForms.Count) then
    Result := FForms[edForm.ItemIndex]
  else
    Result := nil;
end;

function TvgrLocalizeExpertForm.GetSelectedLocalizer: TvgrFormLocalizerInterface;
begin
  if (edLocalizerComponent.ItemIndex >= 0) and (edLocalizerComponent.ItemIndex < edLocalizerComponent.Items.Count) then
    Result := FLocalizers[edLocalizerComponent.ItemIndex]
  else
    Result := nil;
end;

procedure TvgrLocalizeExpertForm.UpdatelLocalizerComponent;
var
  I: Integer;
  ALocalizerInterface: TvgrFormLocalizerInterface;
begin
  FLocalizers.Clear;
  edLocalizerComponent.Items.Clear;

  if SelectedForm = nil then
  begin
    edRCFileName.Text := '';
    lLocalizerComponentDesc.Caption := svgrid_vgr_LocalizeExpertForm_FormNotSelected;
  end
  else
  begin
    for I := 0 to SelectedForm.Component.GetComponentCount - 1 do
      if AnsiCompareText(SelectedForm.Component.GetComponent(I).GetComponentType, TvgrFormLocalizer.ClassName) = 0 then
      begin
        ALocalizerInterface := TvgrFormLocalizerInterface.Create(SelectedForm.Component.GetComponent(I));
        FLocalizers.Add(ALocalizerInterface);
        edLocalizerComponent.Items.AddObject(ALocalizerInterface.Name, ALocalizerInterface)
      end;
    if edLocalizerComponent.Items.Count > 0 then
    begin
      edLocalizerComponent.ItemIndex := 0;
      if edLocalizerComponent.Items.Count = 1 then
        lLocalizerComponentDesc.Caption := svgrid_vgr_LocalizeExpertForm_LocalizerExists
      else
        lLocalizerComponentDesc.Caption := svgrid_vgr_LocalizeExpertForm_LocalizerNotOne;
    end
    else
      lLocalizerComponentDesc.Caption := svgrid_vgr_LocalizeExpertForm_LocalizerNotExists;
  end;
end;

procedure TvgrLocalizeExpertForm.TVDeletion(Sender: TObject;
  Node: TTreeNode);
begin
  if Node.Data <> nil then
    TObject(Node.Data).Free;
end;

procedure TvgrLocalizeExpertForm.edDefaultRCFilesDirChange(
  Sender: TObject);
begin
  bApply.Enabled := True;
end;

procedure TvgrLocalizeExpertForm.edStringIdentsFileChange(Sender: TObject);
begin
  bApply.Enabled := True;
end;

procedure TvgrLocalizeExpertForm.rbProcessAnyClassesClick(Sender: TObject);
begin
  bApply.Enabled := True;
end;

procedure TvgrLocalizeExpertForm.mClassesChange(Sender: TObject);
begin
  bApply.Enabled := True;
  rbProcessSpecifiedClasses.Checked := True;
end;

procedure TvgrLocalizeExpertForm.rbProcessAnyPropertiesClick(
  Sender: TObject);
begin
  bApply.Enabled := True;
end;

procedure TvgrLocalizeExpertForm.mPropertiesChange(Sender: TObject);
begin
  bApply.Enabled := True;
  rbProcessSpecifiedProperties.Checked := True;
end;

procedure TvgrLocalizeExpertForm.bApplyClick(Sender: TObject);
begin
  ApplyOptions;
end;

procedure TvgrLocalizeExpertForm.bStringIdentsFileClick(Sender: TObject);
begin
  SaveDialog.FileName := edStringIdentsFile.Text;
  if SaveDialog.Execute then
  begin
    edStringIdentsFile.Text := SaveDialog.FileName;
    ActiveControl := edStringIdentsFile;
  end;
end;

procedure TvgrLocalizeExpertForm.ApplyOptions;
begin
  UpdateTV(SelectedForm);
  UpdateRCFileName;
  UpdatelLocalizerComponent;

  FStringIDFile.LoadFromFile(edStringIdentsFile.Text);
  FRCFile.LoadFromFile(edRCFileName.Text, FStringIDFile);
  if FRCFile.IncludeFile = '' then
  begin
    if edRCFileName.Text = '' then
      FRCFile.IncludeFile := edStringIdentsFile.Text
    else
    begin
      if edStringIdentsFile.Text <> '' then
        FRCFile.IncludeFile := GetRelativePathToFile(edStringIdentsFile.Text, ExtractFilePath(edDefaultRCFilesDir.Text))
    end;
  end;
  bApply.Enabled := False;
end;

procedure TvgrLocalizeExpertForm.Build;
var
  I, J: Integer;
  S: string;
  APropertyNode: TvgrPropertyNode;
  AStringIDItem: TvgrStringIDFileItem;
  ARCItem: TvgrRCFileItem;
  ALocalizer: IOTAComponent;
  ALocalizerInterface: TvgrFormLocalizerInterface;
  ARLoc: TvgrFormLocalizer;
  ARLocItem: TvgrFormLocalizerItem;
  AReport: TvgrLocalizerBuildReport;
  ARCModule: TvgrRCFileModule;
  AStringIDModule: TvgrStringIDFileModule;
begin
  // create localizer component
  if SelectedLocalizer <> nil then
    Alocalizer := SelectedLocalizer.Component
  else
  begin
    with SelectedForm do
      ALocalizer := FormEditor.CreateComponent(Component, 'TvgrFormLocalizer', 10, 10, 0, 0);
    if ALocalizer <> nil then
    begin
      ALocalizerInterface := TvgrFormLocalizerInterface.Create(ALocalizer);
      FLocalizers.Add(ALocalizerInterface);
      edLocalizerComponent.Items.AddObject(ALocalizerInterface.Name, ALocalizerInterface);
      edLocalizerComponent.ItemIndex := edLocalizerComponent.Items.Count - 1;
    end
    else
      raise Exception.Create(svgrid_vgr_LocalizeExpertForm_ErrorCreateLocalizer);
  end;
  ARLoc := TvgrFormLocalizer(ALocalizer.GetComponentHandle);

  for I := 0 to TV.Items.Count - 1 do
    if (TV.Items[I].StateIndex = sStateChecked) and (TObject(TV.Items[I].Data) is TvgrPropertyNode) then
    begin
      APropertyNode := TvgrPropertyNode(TV.Items[I].Data);
      S := GetPropertyStringResourceID(APropertyNode);
      if S <> '' then
      begin
        AStringIDItem := FStringIDFile.FindByIdent(S);
        if AStringIDItem = nil then
          AStringIDItem := FStringIDFile.AddNewItem(S);

        // find this property in RCFile
        ARCItem := FRCFile.FindByIdent(S);
        if ARCItem = nil then
          ARCItem := FRCFile.AddNewItem(S, APropertyNode.PropValue, AStringIDItem)
        else
        begin
          ARCItem.Value := APropertyNode.PropValue;
          ARCItem.StringID := AStringIDItem;
        end;

        // find this property in TvgrFormLocalizer (ARLoc)
        ARLocItem := ARLoc.Items.FindByResStringID(ARCItem.StringID.Value);
        if ARLocItem = nil then
        begin
          ARLocItem := ARLoc.Items.Add;
          ARLocItem.Component := TComponent(APropertyNode.ComponentHandle);
          ARLocItem.PropName := APropertyNode.PropName;
          ARLocItem.ResStringID := ARCItem.StringID.Value;
        end;
      end;
    end;

  // Build report
  AReport := TvgrLocalizerBuildReport.Create;
  try
    for I := 0 to FRCFile.Count - 1 do
    begin
      ARCItem := FRCFile[I];
      // find this item in TreeView
      J := 0;
      while (J < TV.Items.Count) and
            not ((TV.Items[J].StateIndex = sStateChecked) and
                 (TObject(TV.Items[J].Data) is TvgrPropertyNode) and
                 (AnsiCompareText(GetPropertyStringResourceID(TvgrPropertyNode(TV.Items[J].Data)), ARCItem.Ident) = 0)) do Inc(J);
      if J >= TV.Items.Count then
        AReport.AddRCItem(ARCItem, pos('__', ARCItem.Ident) <> 0, svgrid_vgr_LocalizeExpertForm_NotExistsInLocalizerTree);
    end;

    for I := 0 to ARLoc.Items.Count - 1 do
    begin
      ARLocItem := ARLoc.Items[I];
      // find this item in TreeView
      J := 0;
      while (J < TV.Items.Count) and
            not ((TV.Items[J].StateIndex = sStateChecked) and
                 (TObject(TV.Items[J].Data) is TvgrPropertyNode) and
                 (AnsiCompareText(TvgrPropertyNode(TV.Items[J].Data).PropName, ARLocItem.PropName) = 0) and
                 (TvgrPropertyNode(TV.Items[J].Data).ComponentHandle = ARLocItem.Component)) do Inc(J);
      if J >= TV.Items.Count then
        AReport.AddFLItem(ARLocItem, True, svgrid_vgr_LocalizeExpertForm_NotExistsInLocalizerTree);
    end;

    if (AReport.RCToDeleteItemCount > 0) or (AReport.FLToDeleteItemCount > 0) then
    begin
      if TvgrLocalizerExpertReportForm.Create(nil).Execute(AReport) then
      begin
        // delete items
        for I := 0 to AReport.RCToDeleteItemCount - 1 do
          if AReport.RCToDeleteItems[I].Delete then
            FRCFile.Remove(AReport.RCToDeleteItems[I].RCItem);

        for I := 0 to AReport.FLToDeleteItemCount - 1 do
          if AReport.FLToDeleteItems[I].Delete then
            AReport.FLToDeleteItems[I].FLItem.Free;
      end;
    end;

//    FRCFile.SaveToFile('c:\FRCFile.txt');
//    FStringIDFile.SaveToFile('c:\FStringIDFile.txt');

    if rgAfterBuild.ItemIndex = 0 then
    begin
      // open file in editor
      ARCModule := TvgrRCFileModule.Create(edRCFileName.Text, FRCFile);
      try
        ARCModule.ShowText;
      finally
        ARCModule.Free;
      end;

      AStringIDModule := TvgrStringIDFileModule.Create(edStringIdentsFile.Text, FStringIDFile);
      try
        AStringIDModule.ShowText;
      finally
        AStringIDModule.Free;
      end;
    end
    else
    begin
      FRCFile.SaveToFile(edRCFileName.Text);
      FStringIDFile.SaveToFile(edStringIdentsFile.Text);
    end;

  finally
    AReport.Free;
  end;
end;

function TvgrLocalizeExpertForm.GetPropertyStringResourceID(APropertyNode: TvgrPropertyNode): string;
begin
  with APropertyNode do
    if Component = nil then
      Result := ''
    else
      Result := Format('%s__%s__%s__%s',
                       [svgrResStringPrefix,
                       SelectedForm.UnitName,
                       ComponentNode.ComponentName,
                       PropName]);
end;

procedure TvgrLocalizeExpertForm.RestoreSettings;
begin
  inherited;
  Panel1.Width := SettingsStorage.ReadInteger(StorageSection, 'SplitterPosition', Panel1.Width);
  edDefaultRCFilesDir.Text := SettingsStorage.ReadString(StorageSection, 'DefaultRCFilesDir', '');
  edStringIdentsFile.Text := SettingsStorage.ReadString(StorageSection, 'StringIdentsFile', '');
  if SettingsStorage.ReadBool(StorageSection, 'ProcessAnyClasses', rbProcessAnyClasses.Checked) then
    rbProcessAnyClasses.Checked := True
  else
    rbProcessSpecifiedClasses.Checked := True;
  SettingsStorage.ReadStrings(StorageSection, 'Classes', mClasses.Lines);
  if SettingsStorage.ReadBool(StorageSection, 'ProcessAnyProperties', rbProcessAnyProperties.Checked) then
    rbProcessAnyProperties.Checked := True
  else
    rbProcessSpecifiedProperties.Checked := True;
  SettingsStorage.ReadStrings(StorageSection, 'Properties', mProperties.Lines);
  rgAfterBuild.ItemIndex := SettingsStorage.ReadInteger(StorageSection, 'AfterBuild', 0);

  cbShowNotActions.Checked := SettingsStorage.ReadBool(StorageSection, 'ShowNotAction', cbShowNotActions.Checked);
  cbNotShowEmpty.Checked := SettingsStorage.ReadBool(StorageSection, 'NotShowEmpty', cbNotShowEmpty.Checked);
  cbNotShowSeparator.Checked := SettingsStorage.ReadBool(StorageSection, 'NotShowSeparator', cbNotShowSeparator.Checked);
end;

procedure TvgrLocalizeExpertForm.SaveSettings;
begin
  SettingsStorage.WriteInteger(StorageSection, 'SplitterPosition', Panel1.Width);
  SettingsStorage.WriteString(StorageSection, 'DefaultRCFilesDir', edDefaultRCFilesDir.Text);
  SettingsStorage.WriteString(StorageSection, 'StringIdentsFile', edStringIdentsFile.Text);

  SettingsStorage.WriteBool(StorageSection, 'ProcessAnyClasses', rbProcessAnyClasses.Checked);
  SettingsStorage.WriteStrings(StorageSection, 'Classes', mClasses.Lines);

  SettingsStorage.WriteBool(StorageSection, 'ProcessAnyProperties', rbProcessAnyProperties.Checked);
  SettingsStorage.WriteStrings(StorageSection, 'Properties', mProperties.Lines);

  SettingsStorage.WriteInteger(StorageSection, 'AfterBuild', rgAfterBuild.ItemIndex);

  SettingsStorage.WriteBool(StorageSection, 'ShowNotAction', cbShowNotActions.Checked);
  SettingsStorage.WriteBool(StorageSection, 'NotShowEmpty', cbNotShowEmpty.Checked);
  SettingsStorage.WriteBool(StorageSection, 'NotShowSeparator', cbNotShowSeparator.Checked);
end;

procedure TvgrLocalizeExpertForm.bBuildClick(Sender: TObject);
begin
  Build;
end;

procedure TvgrLocalizeExpertForm.cbShowNotActionsClick(Sender: TObject);
begin
  bApply.Enabled := True;
end;

procedure TvgrLocalizeExpertForm.cbNotShowEmptyClick(Sender: TObject);
begin
  bApply.Enabled := True;
end;

procedure TvgrLocalizeExpertForm.TVChange(Sender: TObject;
  Node: TTreeNode);
begin
  if Node <> nil then
  begin
    StatusBar.SimpleText := TvgrLocalizeNode(Node.Data).GetStatusBarText;
  end;
end;

end.

