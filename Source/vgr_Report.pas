{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains TvgrReportTemplate and TvgrReportEngine classes, that realize the report generation.
TvgrReportTemplate represents the report template on the basis of which TvgrReportEngine
generates the workbook (TvgrWorkbook object).
Each report template can contain some TvgrBand objects.
Band represents part of the template, that can contain some rows or columns.
See also:
  TvgrReportTemplate, TvgrReportEngine, TvgrBand, TvgrDataBand, TvgrDetailBand, TvgrGroupBand.}
unit vgr_Report;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  {$IFDEF VTK_D6_OR_D7} Types, Variants, {$ENDIF} Windows,
  SysUtils, Classes, typinfo, Graphics, DB, ActiveX,

  vgr_CommonClasses, vgr_DataStorage, vgr_DataStorageRecords, vgr_ScriptComponentsProvider,
  vgr_DataStorageTypes, vgr_ScriptControl, vgr_Event, vgr_ScriptDispIDs;

{$R vgr_Report.res}

type
  TvgrReportTemplate = class;
  TvgrReportTemplateWorksheet = class;
  TvgrBands = class;
  TvgrBand =  class;
  TvgrDetailBand = class;
  TvgrDataBand = class;
  TvgrGroupBand = class;
  TvgrSectionsStack = class;
  TvgrTemplateTable = class;
  TvgrReportEngine = class;
  TvgrTemplateVector = class;
  TvgrTemplateRow = class;
  TvgrTemplateCol = class;
  TvgrBandClass = class of TvgrBand;

  /////////////////////////////////////////////////
  //
  // TvgrBandEvents
  //
  /////////////////////////////////////////////////
{Defines the script events that can be fired by TvgrBand object during the report generating.
See also:
  TvgrReportTemplate.OnBandBeforeGenerate, TvgrReportTemplate.OnBandAfterGenerate.}
  TvgrBandEvents = class(TvgrScriptEvents)
  private
    FBand: TvgrBand;
    FBeforeGenerateScriptProcName: string;
    FAfterGenerateScriptProcName: string;
    procedure SetBeforeGenerateScriptProcName(Value: string);
    procedure SetAfterGenerateScriptProcName(Value: string);
  protected
    function GetOwner: TPersistent; override;
    procedure BeforeChange;
    procedure AfterChange;
    property Band: TvgrBand read FBand;
  public
{Creates an instance of the TvgrBandEvents class.
Parameters:
  ABand - owner TvgrBand object.}
    constructor Create(ABand: TvgrBand);

    procedure Assign(Source: TPersistent); override;

{Executes BeforeGenerate event.
This event occurs when the generation of the band is started.
In the beginning the TvgrReportTemplate.OnBandBeforeGenerate event is executed,
then the script procedure.
See also:
  BeforeGenerateScriptProcName, TvgrReportTemplate.OnBandBeforeGenerate}
    procedure DoBeforeGenerate;

{Executes AfterGenerate event.
This event occurs when the generation of the band is finished.
In the beginning the TvgrReportTemplate.OnBandAfterGenerate event is executed,
then the script procedure.
See also:
  AfterGenerateScriptProcName, TvgrReportTemplate.OnBandAfterGenerate}
    procedure DoAfterGenerate;
  published
{Name of the script procedure that executed in the BeforeGenerate event.
See also:
  DoBeforeGenerate}
    property BeforeGenerateScriptProcName: string read FBeforeGenerateScriptProcName write SetBeforeGenerateScriptProcName;

{Name of the script procedure that executed in the AfterGenerate event.
See also:
  DoAfterGenerate}
    property AfterGenerateScriptProcName: string read FAfterGenerateScriptProcName write SetAfterGenerateScriptProcName;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBand
  //
  /////////////////////////////////////////////////
{Represents part of the report template, that can contain some rows or columns.
See also:
  TvgrDetailBand, TvgrDataBand, TvgrGroupBand}
  TvgrBand = class(TvgrComponent, IvgrSection, IvgrSectionExt)
  private
    FSkipGenerate: Boolean;
    FCreateSections: Boolean;
    FSection: rvgrSection;
    FBands: TvgrBands;
    FLoadWorksheet: TvgrReportTemplateWorksheet;
    FFirstVector: Integer;
    FEndVector: Integer;
    FTemplateVector: TvgrTemplateVector;
    FRangesForSecondPass: TInterfaceList;
    FSecondPass: Boolean;
    FRowColAutoSize: Boolean;

    FEvents: TvgrBandEvents;

    procedure SecondPass;
    procedure AddRangesForSecondPass(ASourceRange, ADestRange: IvgrRange);
    function GetTemplate: TvgrReportTemplate;
    function GetTemplateWorksheet: TvgrReportTemplateWorksheet;
    function GetSortIndex: Integer;
    function GetIndex: Integer;
    function GetHorizontal: Boolean;
    procedure RecalcLevel;
    procedure RecalcLevels;
    procedure SetStartPos(Value: Integer);
    procedure SetEndPos(Value: Integer);
    function GetParentBand: TvgrBand;
    function GetFirstVector: Integer;
    function GetEndVector: Integer;
    procedure SetStartEndVectors(AStartPos, AEndPos: Integer);
    procedure SetEvents(Value: TvgrBandEvents);
    procedure SetSkipGenerate(Value: Boolean);
    procedure SetRowColAutoSize(Value: Boolean);
    function GetTemplateTable: TvgrTemplateTable;
  protected
    function GetItemIndex: Integer;
    // IvgrWBListItem
    function GetWorksheet: TvgrWorksheet;
    function GetWorkbook: TvgrWorkbook;
    function GetItemData: Pointer;
    function GetStyleData: Pointer;
    function GetValid: Boolean;
    procedure SetInvalid;
    // IvgrSection
    function GetStartPos: Integer;
    function GetEndPos: Integer;
    function GetLevel: Integer;
    function GetParent: IvgrSection;
    function GetFlags(Index: Integer): Boolean;
    procedure SetFlags(Index: Integer; Value: Boolean);
    procedure SetCreateSections(Value: Boolean);
    procedure Assign(ASource: IvgrSection); reintroduce;
    // IvgrSectionExt
    function GetText: string;

    //
    function GetDispIdOfName(const AName: string): Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                               Flags: Integer;
                               AParametersCount: Integer): HResult; override;
    function DoInvoke(DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;

    procedure ReadVertical(Reader: TReader);
    procedure WriteVertical(Writer: TWriter);
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
    procedure SetParentComponent(Value: TComponent); override;

    // Generate
    procedure DoBeforeGenerate; virtual;
    procedure DoAfterGenerate; virtual;
    procedure InternalGenerate(ABandVector: TvgrTemplateVector); virtual;
    function GetCreateSection: Boolean; virtual;
    procedure Generate(ABandVector: TvgrTemplateVector); virtual;
    procedure DefineProperties(Filer: TFiler); override;

    procedure BeforeChange;
    procedure AfterChange;

    function GetRepeatOnPageTop: Boolean;
    procedure SetRepeatOnPageTop(Value: Boolean);
    function GetRepeatOnPageBottom: Boolean;
    procedure SetRepeatOnPageBottom(Value: Boolean);
    function GetPrintWithNextSection: Boolean;
    procedure SetPrintWithNextSection(Value: Boolean);
    function GetPrintWithPreviosSection: Boolean;
    procedure SetPrintWithPreviosSection(Value: Boolean);

    function CreateEvents: TvgrBandEvents; virtual;

    property SortIndex: Integer read GetSortIndex;
{Returns the TvgrTemplateVector object which corresonds to the given band.
This property is filled in the PrepareVector methods and is correct only
on the stage of building the report.}
    property TemplateVector: TvgrTemplateVector read FTemplateVector;
{Returns the TvgrTemplateTable object in which this band is used.
This property can be used only after calling of the TvgrReportEngine.Generate method.}
    property Table: TvgrTemplateTable read GetTemplateTable;
  public
{Creates an instance of the TvgrBand class.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of the TvgrBand class.}
    destructor Destroy; override;

    function HasParent: Boolean; override;
    function GetParentComponent: TComponent; override;

{Use this function to test is ABand a parent of this band.
Parameters:
  ABand - band to testing.
Return value:
  Returns true if ABand is a parent for this band.}
    function InBand(ABand: TvgrBand): Boolean;

{Use this function to get the string description of the band.
Return value:
  String that describes band.}
    class function GetCaption: string; virtual;
{Use this function to get the resource name of the band bitmap.
Return value:
  String that contains the resource name of the band bitmap.}
    class function GetBitmapResName: string; virtual;

{Returns the report template, that contains this band.}
    property Template: TvgrReportTemplate read GetTemplate;
{Returns worksheet of the report template, that contains this band.}
    property TemplateWorksheet: TvgrReportTemplateWorksheet read GetTemplateWorksheet;
{Returns list of the bands, that holds this band.
See also:
  Index}
    property Bands: TvgrBands read FBands;
{Returns index of the band in the Bands property.
See also:
  Bands}
    property Index: Integer read GetIndex;
{Returns true if the band is horizontal.}
    property Horizontal: Boolean read GetHorizontal;
{Returns level of the band.
The value of Level is 0 for bands that has no parent.
The value of Level is 1 for their children, and so on.
See also:
  ParentBand}
    property Level: Integer read GetLevel;
{Returns the parent band for this band.}
    property ParentBand: TvgrBand read GetParentBand;
{Returns an index of row (for horizontal bands) or column (for vertical)
with which the band begins in the generated workbook.
See also:
  GenEnd}
    property GenBegin: Integer read GetFirstVector;
{Returns an index of row (for horizontal bands) or column (for vertical)
with which the band ends in the generated workbook.
See also:
  GenBegin}
    property GenEnd: Integer read GetEndVector;
  published
{Sets or gets index of row (for horizontal bands) or column (for vertical)
with which the band begins in the report template.
See also:
  EndPos}
    property StartPos: Integer read GetStartPos write SetStartPos;
{Sets or gets index of row (for horizontal bands) or column (for vertical)
with which the band end in the report template.
See also:
  StartPos}
    property EndPos: Integer read GetEndPos write SetEndPos;
{Sets or gets the value indicating whether the band creates the sections.}
    property CreateSections: Boolean read FCreateSections write SetCreateSections default true;
{Sets or gets the value indicating whether band is used at creation of the report.}
    property SkipGenerate: Boolean read FSkipGenerate write SetSkipGenerate default False;
{Sets or gets the value indicating whether the size of rows (for horizontal bands) or
columns (for vertical) of band must be autocalculated at creation of the report.}
    property RowColAutoSize: Boolean read FRowColAutoSize write SetRowColAutoSize default False;
{Sets or gets value that indicates should band printed as page header or not.}
    property RepeatOnPageTop: Boolean Index 0 read GetFlags write SetFlags default false;
{Sets or gets value that indicates should band printed as page footer or not.}
    property RepeatOnPageBottom: Boolean Index 1 read GetFlags write SetFlags default false;
{Sets or gets value that indicates should band are linked to next band or not.
If two sections are connected with each other that at printing they will necessarily placed on one page.
See also:
  PrintWithPreviosSection}
    property PrintWithNextSection: Boolean Index 2 read GetFlags write SetFlags default false;
{Sets or gets value that indicates should band are linked to previous band or not.
If two sections are connected with each other that at printing they will necessarily placed on one page.
See also:
  PrintWithNextSection}
    property PrintWithPreviosSection: Boolean Index 3 read GetFlags write SetFlags default false;

{Contains the script events that can occur during the report generating.}
    property Events: TvgrBandEvents read FEvents write SetEvents;
  end;

{Used for firing information about preparing of the dataset of the detail band for report generating.
Syntax:
  TvgrDetailBandInitDataSetEvent = procedure (Sender: TObject; var AInitializated: Boolean) of object;
Parameters:
  Sender - instance of TvgrDetailBand.
  AInitializated - If you will return true in this parameter, GridReport there will be no additional operations on a dataset.
If false GridReport opens dataset if needed and go to first record of the dataset.
See also:
  TvgrDetailBand}
  TvgrDetailBandInitDataSetEvent = procedure (Sender: TObject; var AInitializated: Boolean) of object;

  /////////////////////////////////////////////////
  //
  // TvgrDetailBandEvents
  //
  /////////////////////////////////////////////////
{Defines the events that can be fired by TvgrDetailBand object during the report generating.
See also:
  TvgrReportTemplate.OnInitDetailBandDataset}
  TvgrDetailBandEvents = class(TvgrBandEvents)
  private
    FInitDatasetScriptProcName: string;
    procedure SetInitDatasetScriptProcName(Value: string);
  public
    procedure Assign(Source: TPersistent); override;
{Executes InitDataset event.
This event occurs when the generation of the detail band are started, after BeforeGenerate event.
In the beginning the TvgrReportTemplate.OnInitDetailBandDataSet event is executed,
then the script procedure.
Parameters:
  AInitializated - if true returned the GridReport will be no additional operations on a dataset.
If false GridReport opens dataset if needed and go to first record of the dataset.
See also:
  TvgrReportTemplate.OnInitDetailBandDataSet}
    procedure DoInitDataset(var AInitializated: Boolean);
  published
{Name of the script procedure that executed in the InitDataset event.
See also:
  DoInitDataset}
    property InitDatasetScriptProcName: string read FInitDatasetScriptProcName write SetInitDatasetScriptProcName;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrDetailBand
  //
  /////////////////////////////////////////////////
{Represents band that linked to the dataset.
Typically this band is a parent for the data band (TvgrDataBand) or group band (TvgrGroupBand).
The child data band repeates for each record of the dataset of detail band.
Also detail band can contains the group band (TvgrGroupBand) in this case
the records will be grouped.
See also:
  TvgrBand, TvgrDataBand, TvgrGroupBand}
  TvgrDetailBand = class(TvgrBand)
  private
    FDataSet: TDataSet;
    function GetEvents: TvgrDetailBandEvents;
    procedure SetEvents(Value: TvgrDetailBandEvents);
    procedure SetDataSet(Value: TDataSet);
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;

    procedure InitDataSet; virtual;
    procedure InternalGenerate(ABandVector: TvgrTemplateVector); override;
    function CreateEvents: TvgrBandEvents; override;
  public
{Use this function to get the string description of the band.
The TvgrDetailBand returns 'Detail band'.
Return value:
  String that describes band.}
    class function GetCaption: string; override;
{Use this function to get the resource name of the band bitmap.
The TvgrDetailBand returns 'VGR_DETAILBAND'.
Return value:
  String that contains the resource name of the band bitmap.}
    class function GetBitmapResName: string; override;
  published
{The TDataset object which linked to this band.}
    property DataSet: TDataSet read FDataSet write SetDataSet;
{Contains events that can occurs while generating of the detail band.
See also:
  TvgrDetailBandGenerateEvents }
    property Events: TvgrDetailBandEvents read GetEvents write SetEvents;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrGroupBand
  //
  /////////////////////////////////////////////////
{Represents the band grouping the data.
This band must be a child for detail band. The detail band provides a dataset the records of that are grouped.
Typically this band is a parent for the data band (TvgrDataBand).
See also:
  TvgrDetailBand, TvgrDataBand}
  TvgrGroupBand = class(TvgrBand)
  private
    FGenerateDataSet: TDataSet;
    FGroupExpression: string;

    function GetGroupValue: Variant;
    function GetDataSet: TDataSet;
    procedure SetGroupExpression(const Value: string);
  protected
    procedure ReadGroupFieldName(Reader: TReader);
    procedure WriteGroupFieldName(Writer: TWriter);
    procedure DefineProperties(Filer: TFiler); override;

    procedure InternalGenerate(ABandVector: TvgrTemplateVector); override;
    procedure Generate(ABandVector: TvgrTemplateVector); override;
    property GenerateDataSet: TDataSet read FGenerateDataSet;
  public
{Use this function to get the string description of the band.
The TvgrGroupBand returns 'Group band'.
Return value:
  String that describes band.}
    class function GetCaption: string; override;
{Use this function to get the resource name of the band bitmap.
The TvgrGroupBand returns 'VGR_GROUPBAND'.
Return value:
  String that contains the resource name of the band bitmap.}
    class function GetBitmapResName: string; override;

{Returns value of the current group.}
    property GroupValue: Variant read GetGroupValue;
{Returns TDataset object, this object is provided by the parent detail band.}
    property DataSet: TDataSet read GetDataSet;
  published
{Returns name of the grouping field of the dataset.}
    property GroupExpression: string read FGroupExpression write SetGroupExpression;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrDataBand
  //
  /////////////////////////////////////////////////
{Represents band which repeates for each record of the dataset.
This band must be a child of TvgrDetailBand or TvgrGroupBand.
See also:
  TvgrDetailBand, TvgrDataBand}
  TvgrDataBand = class(TvgrBand)
  private
    FGenerateSectionForEachRecord: Boolean;
    procedure SetGenerateSectionForEachRecord(Value: Boolean); 
  protected
    function GetCreateSection: Boolean; override;
    procedure Generate(ABandVector: TvgrTemplateVector); override;
  public
{Use this function to get the string description of the band.
The TvgrDataBand returns 'Data band'.
Return value:
  String that describes band.}
    class function GetCaption: string; override;
{Use this function to get the resource name of the band bitmap.
The TvgrDataBand returns 'VGR_DATABAND'.
Return value:
  String that contains the resource name of the band bitmap.}
    class function GetBitmapResName: string; override;
  published
{Sets or gets boolean value that indicates when the separate section for each record of a dataset will be generated.}
    property GenerateSectionForEachRecord: Boolean read FGenerateSectionForEachRecord write SetGenerateSectionForEachRecord default False;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrBands
  //
  /////////////////////////////////////////////////
{Contains list of TvgrBand objects.
Each TvgrReportTemplateWorksheet contains two TvgrBands list.
First list contains the horizontal bands, and second list contains the vertical bands.}
  TvgrBands = class(TvgrSections)
  private
    FList: TList;
    FSortedList: TList;
    FMassUpdateState: Boolean;

    function GetBand(StartPos, EndPos: Integer): TvgrBand;
    function GetBandByIndex(Index: Integer): TvgrBand;
    function GetTemplateWorksheet: TvgrReportTemplateWorksheet;
    function GetTemplate: TvgrReportTemplate;
    procedure BeginMassUpdate;
    procedure EndMassUpdate;
  protected
    procedure DeleteRows(AStartPos, AEndPos: Integer); override;
    procedure InternalInsertLines(AIndexBefore, ACount: Integer); override;
    function GetCount: Integer; override;
    function GetItem(StartPos, EndPos: Integer): IvgrSection; override;
    function GetByIndex(Index: Integer): IvgrSection; override;
    function GetMaxLevel: Integer; override;

    procedure SaveToStream(AStream: TStream); override;
    procedure LoadFromStream(AStream: TStream; ADataStorageVersion: Integer); override;

    function Add(ABandClass: TvgrBandClass): TvgrBand;
    procedure Remove(ABand: TvgrBand);

    procedure RecalcLevels;

    procedure Align;
  public
{Creates instance of the TvgrBands.
Parameters:
  AWorksheet - TvgrReportTemplateWorksheet object that creates this TvgrBands.}
    constructor Create(AWorksheet: TvgrWorksheet); override;
{Frees instance of TvgrBands.}
    destructor Destroy; override;

{Returns the index of the first occurrence in the list of a specified band.
Parameters:
  ABand - TvgrBand object to find.
Return value:
  Returns index of the first occurrence in the list of a specified band, returns -1 if band are not found.}
    function IndexOf(ABand: TvgrBand): Integer;
{Searches for all bands, which begin with a position AStartPos and come to an end on a position AEndPos.
For each found band the callback procedure is called.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  CallBackProc - procedure, that are called for each found band.
  AData - the arbitrary data, that will be passed to the callback procedure.}
    procedure FindAndCallBack(AStartPos, AEndPos: Integer; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer); override;
{Searches for first band, which are placed within interval from AStartPos to AEndPos.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
Return value:
  Returns IvgrSection interface the the band.
See also:
  FindBand, FindBandAt}
    function Find(AStartPos, AEndPos: Integer) : IvgrSection; override;
{Delete the band, which are placed within interval from AStartPos to AEndPos.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval in
which band must be.}
    procedure Delete(AStartPos, AEndPos: Integer); override;
{Searches for first band, which are placed within interval from AStartPos to AEndPos and at specified level.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  ALevel - specifies level of the band.
Return value:
  Returns index of the found band or -1.}
    function SearchCurrentOuter(AStartPos, AEndPos, ALevel: Integer): Integer; override;
{}
    function SearchDirectOuter(AStartPos, AEndPos, ALevel: Integer): Integer; override;

{Searches for first band, which are placed within interval from AStartPos to AEndPos.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
Return value:
  Returns the found band or nil.}
    function FindBand(AStartPos, AEndPos: Integer): TvgrBand;
{Searches for first band, which contains interval from AStartPos to AEndPos.
Band must begins from the AStartPos position or less and must ends with the AEndPos position or large.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval.
Return value:
  Returns the found band or nil.}
    function FindBandAt(AStartPos, AEndPos: Integer): TvgrBand;

{Gets band, which are placed within interval from AStartPos to AEndPos.
Parameters:
  AStartPos - specifies starting row (for horizontal bands) or column (for vertical) of the interval in
which band must be.
  AEndPos - specifies ending row (for horizontal bands) or column (for vertical) of the interval in
which band must be.}
    property Bands[StartPos, EndPos: Integer]: TvgrBand read GetBand; default;
{Gets band by its index.}
    property ByIndex[Index: Integer]: TvgrBand  read GetBandByIndex;
{Gets TvgrReportTemplate that contains this band.}
    property Template: TvgrReportTemplate read GetTemplate;
{Gets TvgrReportTemplateWorksheet that contains this band.}
    property TemplateWorksheet: TvgrReportTemplateWorksheet read GetTemplateWorksheet;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrReportTemplateWorksheetGenerateEvents
  //
  /////////////////////////////////////////////////
{Defines events that can be fired by TvgrReportTemplateWorksheet object during the report generating.
See also:
  TvgrReportTemplate.OnTemplateWorksheetBeforeGenerate, TvgrReportTemplate.OnTemplateWorksheetAfterGenerate,
TvgrReportTemplate.OnAddWorkbookWorksheet}
  TvgrReportTemplateWorksheetGenerateEvents = class(TvgrScriptEvents)
  private
    FWorksheet: TvgrReportTemplateWorksheet;
    FBeforeGenerateScriptProcName: string;
    FAfterGenerateScriptProcName: string;
    FAddWorkbookWorksheetScriptProcName: string;
    FCustomGenerateScriptProcName: string;
  protected
    function GetOwner: TPersistent; override;
    procedure ReadAfterAddWorksheet(Reader: TReader);
    procedure WriteAfterAddWorksheet(Writer: TWriter);
    procedure DefineProperties(Filer: TFiler); override;
  public
{Creates an instance of the TvgrReportTemplateWorksheetGenerateEvents class.
Parameters:
  AWorksheet - The TvgrReportTemplateWorksheet object.}
    constructor Create(AWorksheet: TvgrReportTemplateWorksheet);
    procedure Assign(Source: TPersistent); override;

{Executes BeforeGenerate event.
This event occurs when the generation of the template worksheet is started.
In the beginning the TvgrReportTemplate.OnTemplateWorksheetBeforeGenerate event is executed,
then the script procedure.
See also:
  TvgrReportTemplate.OnTemplateWorksheetBeforeGenerate}
    procedure DoBeforeGenerate;
    
{Executes AfterGenerate event.
This event occurs when the generation of the template worksheet is finished.
In the beginning the TvgrReportTemplate.OnTemplateWorksheetAfterGenerate event is executed,
then the script procedure.
See also:
  TvgrReportTemplate.OnTemplateWorksheetAfterGenerate}
    procedure DoAfterGenerate(AWorksheet: TvgrWorksheet);

{Executes AddWorkbookWorksheet event.
This event occurs when the generation of the template worksheet is started and the
resulting worksheet is added to the workbook (after BeforeGenerate event).
You can change properties of resulting worksheet in
this event (Title, PageSettings and so on).
In the beginning the TvgrReportTemplate.OnAddWorkbookWorksheet event is executed,
then the script procedure.
See also:
  TvgrReportTemplate.OnAddWorkbookWorksheet}
    procedure DoAfterAddWorksheet(AWorksheet: TvgrWorksheet);

{Executes CustomGenerate event.
Occurs after the creating of the worksheet of workbook
and before of running the default algorithm of building the report.
With using this event you can override the default algorithm of building the report.
See also:
  TvgrReportTemplate.OnCustomGenerate}
    procedure DoCustomGenerate(AWorksheet: TvgrWorksheet; var ADone: Boolean);
  published
{Name of the script procedure that executed in the BeforeGenerate event.
See also:
  DoBeforeGenerate}
    property BeforeGenerateScriptProcName: string read FBeforeGenerateScriptProcName write FBeforeGenerateScriptProcName;
{Name of the script procedure that executed in the AfterGenerate event.
See also:
  DoAfterGenerate}
    property AfterGenerateScriptProcName: string read FAfterGenerateScriptProcName write FAfterGenerateScriptProcName;
{Name of the script procedure that executed in the AfterAddWorksheet event.
See also:
  DoAfterAddWorksheet}
    property AddWorkbookWorksheetScriptProcName: string read FAddWorkbookWorksheetScriptProcName write FAddWorkbookWorksheetScriptProcName;
{Name of the script procedure that executed in the AfterAddWorksheet event.
See also:
  DoCustomGenerate}
    property CustomGenerateScriptProcName: string read FCustomGenerateScriptProcName write FCustomGenerateScriptProcName;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrReportTemplateWorksheet
  //
  /////////////////////////////////////////////////
{Represents a worksheet of the report template.
The report template can contains some worksheets, each worksheet of the template adds one
worksheet in the resulting book.
At generating of the report all worksheets are processed consistently since first.
At generating of the resulting workbook it Title and PageSettings are inherited
from an appropriate worksheet in the template. You can reassign these properties
in the AfterAddWorksheet event.
See also:
  TvgrReportTemplateWorksheetGenerateEvents}
  TvgrReportTemplateWorksheet = class(TvgrWorksheet)
  private
    FSkipGenerate: Boolean;
    FGenerateSize: TSize;
    FEvents: TvgrReportTemplateWorksheetGenerateEvents;
    FCurrentRangePlace: TRect;
    FTemplateTable: TvgrTemplateTable;
    FOnCustomGenerateEvent: Boolean;
    function GetHorzBands: TvgrBands;
    function GetVertBands: TvgrBands;
    function GetTemplate: TvgrReportTemplate;
    procedure SetEvents(Value: TvgrReportTemplateWorksheetGenerateEvents);
    procedure SetGenerateSize(Value: TSize);
    function IsGenerateSizeWidthStored: Boolean;
    function IsGenerateSizeHeightStored: Boolean;
    function GetGenerateSizeWidth: Integer;
    procedure SetGenerateSizeWidth(Value: Integer);
    function GetGenerateSizeHeight: Integer;
    procedure SetGenerateSizeHeight(Value: Integer);
    function GetWorkbookRow: Integer;
    function GetWorkbookColumn: Integer;
    function GetWorkbookColumnCaption: string;
  protected
    function GetDimensions: TRect; override;

    procedure DoBeforeGenerate;
    procedure DoAfterGenerate(AWorksheet: TvgrWorksheet);
    procedure DoAfterAddWorksheet(AWorksheet: TvgrWorksheet);
    procedure DoCustomGenerate(AWorksheet: TvgrWorksheet; var ADone: Boolean);

    function GetDispIdOfName(const AName : String) : Integer; override;
    function DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult; override;
    function DoInvoke (DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult; override;
    
    procedure AlignBands;
    function GetSectionsClass: TvgrSectionsClass; override;
    procedure GetChildren(Proc: TGetChildProc; Root: TComponent); override;
    procedure Loaded; override;

    procedure CheckInOnCustomGenerateEvent;
    procedure CheckBand(ABand: TvgrBand; AMustBeHorizontal: Boolean);
  public
{Creates an instance of the TvgrReportTemplateWorksheet class.
Parameters:
  AOwner - Owner component (TvgrReportTemplate).}
    constructor Create(AOwner: TComponent); override;
{Frees an instance of TvgrReportTemplateWorksheet class.}
    destructor Destroy; override;

{Creates and adds the horizontal band to the worksheet.
Parameters:
  ABandClass - class of the added band (TvgrBand, TvgrDetailBand and so on).
Return value:
  Returns the created band.
See also:
  HorzBands, CreateVertBand}
    function CreateHorzBand(ABandClass: TvgrBandClass): TvgrBand;
{Creates and adds the vertical band to the worksheet.
Parameters:
  ABandClass - class of the added band (TvgrBand, TvgrDetailBand and so on).
Return value:
  Returns the created band.
See also:
  VertBands, CreateHorzBand}
    function CreateVertBand(ABandClass: TvgrBandClass): TvgrBand;

{Searches for first vertical band, which contains interval from AStartPos to AEndPos.
Band must begins from the AStartPos column or less and must ends with the AEndPos column or large.
Parameters:
  AStartPos - specifies starting column of the interval.
  AEndPos - specifies ending column of the interval.
Return value:
  Returns the found band or nil.}
    function FindVertBandAt(AStartPos, AEndPos: Integer): TvgrBand;
{Searches for first horizontal band, which contains interval from AStartPos to AEndPos.
Band must begins from the AStartPos row or less and must ends with the AEndPos row or large.
Parameters:
  AStartPos - specifies starting row of the interval.
  AEndPos - specifies ending row of the interval.
Return value:
  Returns the found band or nil.}
    function FindHorzBandAt(AStartPos, AEndPos: Integer): TvgrBand;

{Copies the content of horizontal band into the workbook.
This method can be used only in the handler of the TvgrReportTemplate.OnCustomGenerate event.
Parameters:
  AHorizontalBand - The horizontal band to write.
Example:
  // This procedure displays the content of dataset
  procedure TForm1.SimpleWriteCustomGenerate(Sender: TObject;
    ATemplateWorksheet: TvgrReportTemplateWorksheet;
    AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
  begin
    // write report header
    ATemplateWorksheet.WriteLnBand(r1_Header);
  
    // open the dataset and go to first record
    GlobalDM.Customers.Open;
    GlobalDM.Customers.First;
  
    ADataStart := ATemplateWorksheet.WorkbookRow;
  
    while not GlobalDM.Customers.Eof do
    begin
      // write row of data
      ATemplateWorksheet.WriteLnBand(r1_Data);
  
      // go to next record
      GlobalDM.Customers.Next;
    end;

    // write footer
    ATemplateWorksheet.WriteLnBand(r1_Footer);
  end;}
    procedure WriteLnBand(AHorizontalBand: TvgrBand); 
{Copies the cell of the report template on the
intersection of the AHorizontalBand and AVerticalBand bands into the workbook.
Can be used only in the handler of the TvgrReportTemplate.OnCustomGenerate event.
Parameters:
  AHorizontalBand - The horizontal band to write.
  AVerticalBand - The vertical band to write.
Example:
  // this procedure displays the cross-tab table.
  //   DS - horizontal dataset
  //   KS - vertical dataset 
  procedure TWriteXXExamplesForm.CrossTabWriteCustomGenerate(Sender: TObject;
    ATemplateWorksheet: TvgrReportTemplateWorksheet;
    AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
  begin
    GlobalDm.DS.First;
    while not GlobalDm.DS.Eof do
    begin
      // row footer
      ATemplateWorksheet.WriteCell(r3_HorData, r3_VerHeader);
  
      // row data
      GlobalDm.KS.First;
      while not GlobalDm.KS.Eof do
      begin
        ATemplateWorksheet.WriteCell(r3_HorData, r3_VerData);
        GlobalDm.KS.Next;
      end;
  
      // summary on row
      ATemplateWorksheet.WriteCell(r3_HorData, r3_VerFooter);
  
      // go to next row
      ATemplateWorksheet.WriteLn;
      GlobalDm.DS.Next;
    end;
  end;}
    procedure WriteCell(AHorizontalBand: TvgrBand; AVerticalBand: TvgrBand);
{Begins the new row typically is used with WriteCell function,
to begin the new row.
If ARowCount >= 0 then skips the specified number of rows.
This method can be used only in the handler of the TvgrReportTemplate.OnCustomGenerate event.
Parameters:
  ARowCount - number of rows to skip.
See also:
  WriteCell}
    procedure WriteLn(ARowCount: Integer = -1); 
{Copies the cell of the report template on the
intersection of the AHorizontalBand and AVerticalBand bands into the workbook and begins new row.
This method can be used only in the handler of the TvgrReportTemplate.OnCustomGenerate event.
Parameters:
  AVerticalBand - The vertical band to write.
  AHorizontalBand - The horizontal band to write.
See also:
  WriteCell}
    procedure WriteLnCell(AHorizontalBand: TvgrBand; AVerticalBand: TvgrBand);

{Returns the list of the horizontal bands.}
    property HorzBands: TvgrBands read GetHorzBands;
{Returns the list of the vertical bands.}
    property VertBands: TvgrBands read GetVertBands;
{Returns the TvgrReportTemplate object, that contains this worksheet.}
    property Template: TvgrReportTemplate read GetTemplate;
{Specifies the area of worksheet of the report template that
will be used for report forming.}
    property GenerateSize: TSize read FGenerateSize write SetGenerateSize;

{Returns the number of row within generated worksheet from which the next
band will be placed, first row has index 0.}
    property WorkbookRow: Integer read GetWorkbookRow;

{Returns the number of column within generated worksheet from which the next
band will be placed, first column has index 0.}
    property WorkbookColumn: Integer read GetWorkbookColumn;

{Returns the caption of column within generated worksheet from which the next
band will be placed.}
    property WorkbookColumnCaption: string read GetWorkbookColumnCaption;
  published
{Sets or gets value that indicates should worksheet participate at creation of the report whether or not.}
    property SkipGenerate: Boolean read FSkipGenerate write FSkipGenerate default False;
{Contains events that can occurs while generating.}
    property Events: TvgrReportTemplateWorksheetGenerateEvents read FEvents write SetEvents;
    property GenerateSizeWidth: Integer read GetGenerateSizeWidth write SetGenerateSizeWidth stored IsGenerateSizeWidthStored;
    property GenerateSizeHeight: Integer read GetGenerateSizeHeight write SetGenerateSizeHeight stored IsGenerateSizeHeightStored;
  end;

{TvgrInitDetailBandDataSetEvent is the type for the OnInitDetailBandDataSet event.
Occurs when the dataset of TvgrDetailBand is prepared during the report generating.
Parameters:
  Sender - The TvgrReportTemplate object containing this band.
  AInitialized - If you will return true in this parameter, GridReport will be no additional operations on a dataset.
If false GridReport opens dataset if needed and go to first record of the dataset.
See also:
  TvgrDetailBand}
  TvgrInitDetailBandDataSetEvent = procedure (Sender: TObject; ABand: TvgrDetailBand; var AInitialized: Boolean) of object;
  
{TvgrScriptErrorInExpression is the type for the OnBandBeforeGenerate and OnBandAfterGenerate events.
These events occur before and after the band generating.
Parameters:
  Sender - The TvgrReportTemplate object containing this band.
  ABand - The TvgrBand object.
See also:
  TvgrReportTemplate.OnBandBeforeGenerate, TvgrReportTemplate.OnBandAfterGenerate}
  TvgrBandGenerateEvent = procedure (Sender: TObject; ABand: TvgrBand) of object;

{TvgrTemplateWorksheetGenerateEvent is the type for the OnTemplateWorksheetBeforeGenerate and OnTemplateWorksheetAfterGenerate events.
These events occur before and after generating of the report template worsheet.
Parameters:
  Sender - The TvgrReportTemplate object containing worksheet.
  ATemplateWorksheet - The TvgrReportTemplateWorksheet object.}
  TvgrTemplateWorksheetGenerateEvent = procedure (Sender: TObject; ATemplateWorksheet: TvgrReportTemplateWorksheet) of object;

{TvgrAddWorkbookWorksheetEvent is the type for the OnAddWorkbookWorksheet event.
Occurs when destination worksheet of workbook is created and added into workbook.
Parameters:
  Sender - The TvgrReportTemplate object containing the report template worksheet.
  ATemplateWorksheet - Sheet of the report template (TvgrReportTemplateWorksheet object) that adds worksheet.
  AResultWorksheet - The added worksheet.}
  TvgrAddWorkbookWorksheetEvent = procedure (Sender: TObject; ATemplateWorksheet: TvgrReportTemplateWorksheet; AResultWorksheet: TvgrWorksheet) of object;

{TvgrCustomGenerateEvent is the type for the OnCustomGenerate event.
Occurs after the creating of the worksheet of workbook
and before of running the default algorithm of building the report.
With using this event you can override the default processing.
Parameters:
  Sender - The TvgrReportTemplate object.
  ATemplateWorksheet - TvgrReportTemplateWorksheet object processing of which is started.
  AResultWorksheet - TvgrWorksheet object which is being generated.
  ADone - Return the true in this parameter to cancel the default algorith of building the report.}
  TvgrCustomGenerateEvent = procedure (Sender: TObject; ATemplateWorksheet: TvgrReportTemplateWorksheet; AResultWorksheet: TvgrWorksheet; var ADone: Boolean) of object;
  /////////////////////////////////////////////////
  //
  // TvgrReportTemplate
  //
  /////////////////////////////////////////////////
{Represents the report template.
The report template consists from the list of the worksheets (Worksheets property),
script that can executed while report are generated(Script property)
and the ComponentsProvider which provides list of the components, which available for use.
See also:
  TvgrTemplateScript, TvgrReportAvailableComponentsProvider}
  TvgrReportTemplate = class(TvgrWorkbook)
  private
    FScript: TvgrScript;

    FOnBandBeforeGenerate: TvgrBandGenerateEvent;
    FOnBandAfterGenerate: TvgrBandGenerateEvent;

    FOnInitDetailBandDataSet: TvgrInitDetailBandDataSetEvent;

    FOnTemplateWorksheetBeforeGenerate: TvgrTemplateWorksheetGenerateEvent;
    FOnTemplateWorksheetAfterGenerate: TvgrTemplateWorksheetGenerateEvent;
    FOnAddWorkbookWorksheet: TvgrAddWorkbookWorksheetEvent;
    FOnCustomGenerate: TvgrCustomGenerateEvent;

    function GetOnScriptGetObject: TvgrScriptGetObject;
    procedure SetOnScriptGetObject(Value: TvgrScriptGetObject);

    function GetOnScriptErrorInExpression: TvgrScriptErrorInExpression;
    procedure SetOnScriptErrorInExpression(Value: TvgrScriptErrorInExpression);

    function GetOnGetAvailableComponents: TvgrGetAvailableComponentsEvent;
    procedure SetOnGetAvailableComponents(Value: TvgrGetAvailableComponentsEvent);

    function GetWorksheet(Index: Integer): TvgrReportTemplateWorksheet;
    procedure SetScript(Value: TvgrScript);

    procedure OnScriptChange(Sender: TObject);
  protected
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
    function GetWorksheetClass: TvgrWorksheetClass; override;

    procedure BeforeChangeReportTemplate;
    procedure AfterChangeReportTemplate;

    procedure BeginGenerate;
    procedure EndGenerate;

    procedure DoBandBeforeGenerate(ABand: TvgrBand);
    procedure DoBandAfterGenerate(ABand: TvgrBand);
    procedure DoInitDetailBandDataSetEvent(ABand: TvgrDetailBand; var AInitializated: Boolean);
    procedure DoTemplateWorksheetBeforeGenerate(ATemplateWorksheet: TvgrReportTemplateWorksheet);
    procedure DoTemplateWorksheetAfterGenerate(ATemplateWorksheet: TvgrReportTemplateWorksheet);
    procedure DoAddWorkbookWorksheetEvent(ATemplateWorksheet: TvgrReportTemplateWorksheet; AResultWorksheet: TvgrWorksheet);
    procedure DoCustomGenerate(ATemplateWorksheet: TvgrReportTemplateWorksheet; AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
  public
{Creates instance of the TvgrReportTemplate.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees instance of a TvgrReportTemplate.}
    destructor Destroy; override;
{Clears the template.}
    procedure Clear; override;

{Lists the worksheets in the TvgrReportTemplate.
Parameters:
  Index - index of the worksheet.}
    property Worksheets[Index: Integer]: TvgrReportTemplateWorksheet read GetWorksheet;
  published
{Sets or gets script, that can executed while generating report.
See also:
  TvgrScriptOptions}
    property Script: TvgrScript read FScript write SetScript;

{Occurs when an unknown object is found in the script.
See also:
  TvgrScriptGetObject}
    property OnScriptGetObject: TvgrScriptGetObject read GetOnScriptGetObject write SetOnScriptGetObject;

{Occurs when an error in calculating of value of expression occurs.
All expressions in the report template are written in the square brackets.
See also:
  TvgrScriptErrorInExpression}
    property OnScriptErrorInExpression: TvgrScriptErrorInExpression read GetOnScriptErrorInExpression write SetOnScriptErrorInExpression;

{Occurs when list of available objects is forming
See also:
  Script.GetAvailableComponents, TvgrGetAvailableComponentsEvent}
    property OnGetAvailableComponents: TvgrGetAvailableComponentsEvent read GetOnGetAvailableComponents write SetOnGetAvailableComponents;

{Occurs when the processing of band of the report template is starting.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrBand.Event.BeforeGenerateScriptProcName property.
See also:
  TvgrBandGenerateEvent}
    property OnBandBeforeGenerate: TvgrBandGenerateEvent read FOnBandBeforeGenerate write FOnBandBeforeGenerate;

{Occurs when the processing of band of the report template is ended.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrBand.Events.AfterGenerateScriptProcName property.
See also:
  TvgrBandGenerateEvent}
    property OnBandAfterGenerate: TvgrBandGenerateEvent read FOnBandAfterGenerate write FOnBandAfterGenerate;

{Occurs when the dataset of TvgrDetailBand is preparing during the report generating.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrDetailBand.Events.InitDatasetScriptProcName property.
See also:
  TvgrInitDetailBandDataSetEvent}
    property OnInitDetailBandDataSet: TvgrInitDetailBandDataSetEvent read FOnInitDetailBandDataSet write FOnInitDetailBandDataSet;

{Occurs when processing of worksheet of the report template is starting.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrReportTemplateWorksheet.Events.BeforeGenerateScriptProcName.
See also:
  TvgrTemplateWorksheetGenerateEvent}
    property OnTemplateWorksheetBeforeGenerate: TvgrTemplateWorksheetGenerateEvent read FOnTemplateWorksheetBeforeGenerate write FOnTemplateWorksheetBeforeGenerate;

{Occurs when processing of worksheet of the report template is ended.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrReportTemplateWorksheet.Events.AfterGenerateScriptProcName.
See also:
  TvgrTemplateWorksheetGenerateEvent}
    property OnTemplateWorksheetAfterGenerate: TvgrTemplateWorksheetGenerateEvent read FOnTemplateWorksheetAfterGenerate write FOnTemplateWorksheetAfterGenerate;

{Occurs when the destination workbook worksheet is created and added into workbook.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrReportTemplateWorksheet.Events.AddWorkbookWorksheetScriptProcName.
See also:
  TvgrAddWorkbookWorksheetEvent}
    property OnAddWorkbookWorksheet: TvgrAddWorkbookWorksheetEvent read FOnAddWorkbookWorksheet write FOnAddWorkbookWorksheet;

{Occurs after the creating of the worksheet of workbook
and before of running the default algorithm of building the report.
With using this event you can override the default processing.
This event is being executed before the script procedure, name of script procedure
is defined in the TvgrReportTemplateWorksheet.Events.CustomGenerateScriptProcName.
See also:
  TvgrCustomGenerateEvent}
    property OnCustomGenerate: TvgrCustomGenerateEvent read FOnCustomGenerate write FOnCustomGenerate; 
  end;

{TvgrTemplateVectorClass is the class type of a TvgrTemplateVector descendant.
Syntax:
  TvgrTemplateVectorClass = class of TvgrTemplateVector;}
  TvgrTemplateVectorClass = class of TvgrTemplateVector;
  /////////////////////////////////////////////////
  //
  // TvgrTemplateVector
  //
  /////////////////////////////////////////////////
{Internal class, represents part of the template worksheet.
Instances of this class creates while generating report.}
  TvgrTemplateVector = class(TObject)
  private
    FStartPos: Integer;
    FSize: Integer;
    FSubVectors: TList;
    FParentVector: TvgrTemplateVector;
    FBand: TvgrBand;
    FTable: TvgrTemplateTable;
    function GetSubVectorCount: Integer;
    function GetSubVector(Index: Integer): TvgrTemplateVector;
    function GetEngine: TvgrReportEngine;
    function AddSubVector(AStartPos, ASize: Integer; ABand: TvgrBand; AParentVector: TvgrTemplateVector): TvgrTemplateVector;
    function GetRowColAutoSize: Boolean;
  protected
    function GetSelfClass: TvgrTemplateVectorClass; virtual; abstract;
    function GetStack: TvgrSectionsStack; virtual; abstract;
    function GetPos: Integer; virtual; abstract;
    function GetSections: TvgrSections; virtual; abstract;
    procedure StartSection;
    procedure EndSection;

    procedure Write; virtual; abstract;
    function FindBandVector(ABandClass: TvgrBandClass): TvgrTemplateVector;
    procedure Clear;

    property Table: TvgrTemplateTable read FTable;
    property Engine: TvgrReportEngine read GetEngine;
    property Band: TvgrBand read FBand;
    property Size: Integer read FSize;
    property StartPos: Integer read FStartPos;
    property SubVectorCount: Integer read GetSubVectorCount;
    property SubVectors[Index: Integer]: TvgrTemplateVector read GetSubVector;
    property RowColAutoSize: Boolean read GetRowColAutoSize;
    property ParentVector: TvgrTemplateVector read FParentVector;
  public
{Creates instance of the TvgrTemplateTable.
Parameters:
  ATable - TvgrTemplateTable that contains this TvgrTemplateVector object.}
    constructor Create(ATable: TvgrTemplateTable);
{Frees instance of a TvgrTemplateVector.}
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrTemplateRow
  //
  /////////////////////////////////////////////////
{Internal class, represents part of the template worksheet that contains some rows.
Instances of this class creates while generating report.}
  TvgrTemplateRow = class(TvgrTemplateVector)
  protected
    function GetSelfClass: TvgrTemplateVectorClass; override;
    function GetStack: TvgrSectionsStack; override;
    function GetPos: Integer; override;
    function GetSections: TvgrSections; override;

    procedure Write; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrTemplateCol
  //
  /////////////////////////////////////////////////
{Internal class, represents part of the template worksheet that contains some columns.
Instances of this class creates while generating report.}
  TvgrTemplateCol = class(TvgrTemplateVector)
  protected
    function GetSelfClass: TvgrTemplateVectorClass; override;
    function GetStack: TvgrSectionsStack; override;
    function GetPos: Integer; override;
    function GetSections: TvgrSections; override;

    procedure Write; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrSectionsStack
  //
  /////////////////////////////////////////////////
{Internal class.}
  TvgrSectionsStack = class(TObject)
  private
    FList: TList;
  protected
    procedure Push(APos: Integer);
    function Pop: Integer;
    procedure Clear;
  public
{Creates an instance of the TvgrSectionsStack class.}
    constructor Create;
{Frees an instance of the TvgrSectionsStack class.}
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrTemplateTable
  //
  /////////////////////////////////////////////////
{Internal class, represents the report template worksheet.
Instances of this class creates while generating report for each TvgrReportTemlateWorksheet.}
  TvgrTemplateTable = class(TObject)
  private
    FEngine: TvgrReportEngine;
    FTemplateWorksheet: TvgrReportTemplateWorksheet;
    FResultWorksheet: TvgrWorksheet;
    FRow: TvgrTemplateRow;
    FCol: TvgrTemplateCol;
    FCurrentRow: TvgrTemplateRow;
    FHorzStack: TvgrSectionsStack;
    FVertStack: TvgrSectionsStack;
    FStoreRangeForSecondPass: Boolean;
    FTempBand: TvgrBand;
    FWritedCols: TList;

    FWorkbookRow, FWorkbookCol: Integer;
    FWriteCellRow: TvgrTemplateRow;

    procedure StoreRangeForSecondPass(ABand: TvgrBand);
    function GetRowCount: Integer;
    function GetColCount: Integer;
    function GetRow(Index: Integer): TvgrTemplateRow;
    function GetCol(Index: Integer): TvgrTemplateCol;
    procedure PrepareVector(AVector: TvgrTemplateVector; ABands: TvgrBands; AMaxEndPos: Integer);
    procedure WriteRangesCallback1(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
    procedure WriteRangesCallback2(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
  protected
    function CalculateFormula(const AFormula: string): Variant;
    function ParseString(const S: string): string;
    procedure CalculateRangeValue(ASourceRange, ADestRange: IvgrRange);
    procedure WriteRanges(ALeft, ATop, AWidth, AHeight: Integer);
    procedure WriteRows(ATop, AHeight: Integer; AAutoSize: Boolean);
    procedure WriteCols(ALeft, AWidth: Integer);
    procedure WriteLn(ALineSize: Integer);

    procedure Clear;
    procedure Prepare;
    procedure Generate;

    property CurrentRow: TvgrTemplateRow read FCurrentRow write FCurrentRow;

    property Engine: TvgrReportEngine read FEngine;
    property TemplateWorksheet: TvgrReportTemplateWorksheet read FTemplateWorksheet;
    property ResultWorksheet: TvgrWorksheet read FResultWorksheet write FResultWorksheet;
    property RowCount: Integer read GetRowCount;
    property ColCount: Integer read GetColCount;
    property Rows[Index: Integer]: TvgrTemplateRow read GetRow;
    property Cols[Index: Integer]: TvgrTemplateCol read GetCol;
  public
{Creates instance of the TvgrTemplateTable.
Parameters:
  AEngine - TvgrReportEngine object that uses this object.
  ATemplateWorksheet - TvgrReportTemplateWorksheet object for which this TvgrTemplateTable are created.}
    constructor Create(AEngine: TvgrReportEngine; ATemplateWorksheet: TvgrReportTemplateWorksheet);
{Frees instance of a TvgrTemplateTable.}
    destructor Destroy; override;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrReportEngine
  //
  /////////////////////////////////////////////////
{Non visual component, that realizes the report engine.
Call the Generate method to generate report.
The Template property must be assigned to the TvgrReportTemplate object, that contains the report template.
The Workbook property must be assigned to the TvgrWorkbook object, that receives the results of generating.
See also:
  TvgrReportTemplate, TvgrWorkbook}
  TvgrReportEngine = class(TComponent)
  private
    FTemplate: TvgrReportTemplate;
    FWorkbook: TvgrWorkbook;
    FTables: TList;
    procedure SetTemplate(Value: TvgrReportTemplate);
    procedure SetWorkbook(Value: TvgrWorkbook);
    function GetTableCount: Integer;
    function GetTable(Index: Integer): TvgrTemplateTable;
  protected
    procedure ClearTables;
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;

    property TableCount: Integer read GetTableCount;
    property Tables[Index: Integer]: TvgrTemplateTable read GetTable;
  public
{Creates instance of the TvgrReportEngine.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees instance of a TvgrReportEngine.}
    destructor Destroy; override;

{Generates the report.
The Template property must be assigned to the TvgrReportTemplate object, that contains the report template.
The Workbook property must be assigned to the TvgrWorkbook object, that receives the results of generating.}
    procedure Generate;
  published
{Sets or gets TvgrReportTemplate object, that contains the report template.}
    property Template: TvgrReportTemplate read FTemplate write SetTemplate;
{Sets or gets TvgrWorkbook object, that receives the results of generating.}
    property Workbook: TvgrWorkbook read FWorkbook write SetWorkbook;
  end;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredBandClasses
  //
  /////////////////////////////////////////////////
{Contains the registered bands classes.
Only one instance of this class is created, at start of the
program and stored in the BandClasses global variable.
See also:
  BandClasses}
  TvgrRegisteredBandClasses = class(TObject)
  private
    FList: TList;
    FBitmapList: TList;
    function GetItem(Index: Integer): TvgrBandClass;
    function GetBitmap(Index: Integer): TBitmap;
    function GetCount: Integer;
  public
{Creates instance of the TvgrRegisteredBandClasses.}
    constructor Create;
{Frees instance of a TvgrRegisteredBandClasses.}
    destructor Destroy; override;
{Registers new band class.
After registering you can get the band bitmap with using Bitmaps property.
Parameters:
  ABandClass - class to registering.}
    procedure RegisterBandClass(ABandClass: TvgrBandClass);
{Gets index of the band class.
Parameters:
  ABand - the band object.
Return value:
  Returns index of the band class or -1 if the band class is not found.}
    function GetBandClassIndex(ABand: TvgrBand): Integer;
{Lists the registered band classes.
Parameters:
  Index - index of the band class.}
    property Items[Index: Integer]: TvgrBandClass read GetItem; default;
{Lists the bitmaps of the registered band classes.
Parameters:
  Index - index of the band class.}
    property Bitmaps[Index: Integer]: TBitmap read GetBitmap;
{Indicates the number of band classes in the list.}
    property Count: Integer read GetCount;
  end;

{Registers new band class.
Parameters:
  ABandClass - class to registering.}
procedure RegisterBandClass(ABandClass: TvgrBandClass);

var
{Instance of the TvgrRegisteredBandClasses class, which are created at the start of the programm.}
  BandClasses: TvgrRegisteredBandClasses;
  
implementation

uses
  vgr_ScriptParser, vgr_Functions, Math, {$IFDEF VTK_D6_OR_D7} DateUtils, {$ENDIF} ComObj;

const
  sse_Write_NotHorizontal = '[%s] must be a horizontal band.';
  sse_Write_NotVertical = '[%s] must be a vertical band.';
  sse_Write_NotOnCustomGenerateEvent = 'Method can be used only in TvgrReportTemplate.OnCustomGenerate event.';
  sse_Write_TemplateVectorIsNil = 'Internal error, TvgrBand.TemplateVector is nil for band [%s].';
  sse_Write_InvalidBandTemplateWorksheet = '[%s] band does not belong to [%s] worksheet.';
  sse_Write_HorizontalBandIsNotSpecified = 'Horizontal band is not specified.';
  sse_Write_VerticalBandIsNotSpecified = 'Vertical band is not specified.';

  sse_BandGenerate_InvalidSecondParameter = 'if TvgrBand.Generate used with parameter then it must be a vertical band.';

procedure RegisterBandClass(ABandClass: TvgrBandClass);
begin
  BandClasses.RegisterBandClass(ABandClass);
end;

/////////////////////////////////////////////////
//
// TvgrBandEvents
//
/////////////////////////////////////////////////
constructor TvgrBandEvents.Create(ABand: TvgrBand);
begin
  inherited Create;
  FBand := ABand;
end;

function TvgrBandEvents.GetOwner: TPersistent;
begin
  Result := FBand;
end;

procedure TvgrBandEvents.SetBeforeGenerateScriptProcName(Value: string);
begin
  if FBeforeGenerateScriptProcName <> Value then
  begin
    BeforeChange;
    FBeforeGenerateScriptProcName := Value;
    AfterChange;
  end;
end;

procedure TvgrBandEvents.SetAfterGenerateScriptProcName(Value: string);
begin
  if FAfterGenerateScriptProcName <> Value then
  begin
    BeforeChange;
    FAfterGenerateScriptProcName := Value;
    AfterChange;
  end;
end;

procedure TvgrBandEvents.BeforeChange;
begin
  Band.BeforeChange;
end;

procedure TvgrBandEvents.AfterChange;
begin
  Band.AfterChange;
  DoChange;
end;

procedure TvgrBandEvents.Assign(Source: TPersistent);
begin
  if Source is TvgrBandEvents then
    with TvgrBandEvents(Source) do
    begin
      Self.FBeforeGenerateScriptProcName := BeforeGenerateScriptProcName;
      Self.FAfterGenerateScriptProcName := AfterGenerateScriptProcName;
    end;
end;

procedure TvgrBandEvents.DoBeforeGenerate;
var
  AParameters: TvgrVariantDynArray;
begin
  Band.Template.DoBandBeforeGenerate(Band);

  if FBeforeGenerateScriptProcName <> '' then
  begin
    SetLength(AParameters, 1);
    AParameters[0] := Band as IDispatch;
    Band.Template.Script.RunProcedure(FBeforeGenerateScriptProcName, AParameters);
  end;
end;

procedure TvgrBandEvents.DoAfterGenerate;
var
  AParameters: TvgrVariantDynArray;
begin
  Band.Template.DoBandAfterGenerate(Band);

  if FAfterGenerateScriptProcName <> '' then
  begin
    SetLength(AParameters, 1);
    AParameters[0] := Band as IDispatch;
    Band.Template.Script.RunProcedure(FAfterGenerateScriptProcName, AParameters);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrDetailBandEvents
//
/////////////////////////////////////////////////
procedure TvgrDetailBandEvents.SetInitDatasetScriptProcName(Value: string);
begin
  if FInitDatasetScriptProcName <> Value then
  begin
    BeforeChange;
    FInitDatasetScriptProcName := Value;
    AfterChange;
  end;
end;

procedure TvgrDetailBandEvents.Assign(Source: TPersistent);
begin
  if Source is TvgrDetailBandEvents then
    with TvgrDetailBandEvents(Source) do
    begin
      Self.FInitDatasetScriptProcName := InitDatasetScriptProcName;
    end;
  inherited;
end;

procedure TvgrDetailBandEvents.DoInitDataset(var AInitializated: Boolean);
var
  AParameters: TvgrVariantDynArray;
  AByRef: TvgrBooleanDynArray;
begin
  AInitializated := False;
  Band.Template.DoInitDetailBandDataSetEvent(TvgrDetailBand(Band), AInitializated);

  if FInitDatasetScriptProcName <> '' then
  begin
    SetLength(AParameters, 2);
    AParameters[0] := Band as IDispatch;
    AParameters[1] := AInitializated;
    SetLength(AByRef, 1);
    AByRef[0] := False;
    AByRef[1] := True;
    Band.Template.Script.RunProcedure(FInitDatasetScriptProcName, AParameters, AByRef);
    AInitializated := AParameters[1];
  end;
end;

/////////////////////////////////////////////////
//
// TvgrBand
//
/////////////////////////////////////////////////
constructor TvgrBand.Create(AOwner: TComponent);
begin
  inherited;
  FCreateSections := True;
  FSection.StartPos := 0;
  FSection.EndPos := 0;
  FRangesForSecondPass := TInterfaceList.Create;
  FSecondPass := False;

  FEvents := CreateEvents;
end;

function TvgrBand.CreateEvents: TvgrBandEvents;
begin
  Result := TvgrBandEvents.Create(Self);
end;

destructor TvgrBand.Destroy;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  if Template <> nil then
  begin
    AChangeInfo.ChangesType := Bands.GetDeleteChangesType;
    AChangeInfo.ChangedObject := Self;
    AChangeInfo.ChangedInterface := Self;
    Template.BeforeChange(AChangeInfo);
  end;

  if FBands <> nil then
    FBands.Remove(Self);

  FRangesForSecondPass.Free;  
  inherited;

  if Template <> nil then
  begin
    AChangeInfo.ChangedObject := nil;
    AChangeInfo.ChangedInterface := nil;
    Template.AfterChange(AChangeInfo);
  end;
  FEvents.Free;
end;

procedure TvgrBand.ReadVertical(Reader: TReader);
var
  AVertical: Boolean;
begin
  AVertical := Reader.ReadBoolean;
  if FLoadWorksheet <> nil then
  begin
    if AVertical then
      FBands := FLoadWorksheet.VertBands
    else
      FBands := FLoadWorksheet.HorzBands;
    FBands.FList.Add(Self);
    FBands.FSortedList.Add(Self);
    RecalcLevels;
  end;
end;

procedure TvgrBand.WriteVertical(Writer: TWriter);
begin
  Writer.WriteBoolean(not Horizontal);
end;

procedure TvgrBand.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('Vertical', ReadVertical, WriteVertical, True);
end;

procedure TvgrBand.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
end;

procedure TvgrBand.SetParentComponent(Value : TComponent);
begin
  if Value is TvgrReportTemplateWorksheet then
  begin
    FLoadWorksheet := TvgrReportTemplateWorksheet(Value);
  end;
end;

function TvgrBand.HasParent: Boolean;
begin
  Result := True;
end;

function TvgrBand.GetParentComponent: TComponent;
begin
  Result := TemplateWorksheet
end;

function TvgrBand.InBand(ABand: TvgrBand): Boolean;
begin
  Result := (ABand.StartPos <= StartPos) and
            (ABand.EndPos >= EndPos) and
            (ABand.SortIndex < Self.SortIndex);
end;

class function TvgrBand.GetCaption: string;
begin
  Result := 'Simple band';
end;

class function TvgrBand.GetBitmapResName: string;
begin
  Result := 'VGR_BAND';
end;

procedure TvgrBand.SecondPass;
var
  I: Integer;
begin
  FSecondPass := True;
  for I := 0 to (FRangesForSecondPass.Count div 2) - 1 do
    Table.CalculateRangeValue(IvgrRange(FRangesForSecondPass[I*2]),IvgrRange(FRangesForSecondPass[I*2+1]));
  FRangesForSecondPass.Clear;
end;

procedure TvgrBand.AddRangesForSecondPass(ASourceRange, ADestRange: IvgrRange);
begin
  if not FSecondPass then
  begin
    FRangesForSecondPass.Add(ASourceRange);
    FRangesForSecondPass.Add(ADestRange);
  end;  
end;

function TvgrBand.GetTemplate: TvgrReportTemplate;
begin
  if Bands = nil then
    Result := nil
  else
    Result := Bands.Template;
end;

function TvgrBand.GetTemplateWorksheet: TvgrReportTemplateWorksheet;
begin
  if Bands = nil then
    Result := nil
  else
    Result := Bands.TemplateWorksheet;
end;

procedure TvgrBand.BeforeChange;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  if Bands <> nil then
  begin
    AChangeInfo.ChangesType := Bands.GetEditChangesType;
    AChangeInfo.ChangedObject := Self;
    AChangeInfo.ChangedInterface := Self;
    Template.BeforeChange(AChangeInfo);
  end;
end;

procedure TvgrBand.AfterChange;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  if Bands <> nil then
  begin
    AChangeInfo.ChangesType := Bands.GetEditChangesType;
    AChangeInfo.ChangedObject := Self;
    AChangeInfo.ChangedInterface := Self;
    Template.AfterChange(AChangeInfo);
  end;
end;

function TvgrBand.GetRepeatOnPageTop: Boolean;
begin
  Result := RepeatOnPageTop;
end;

procedure TvgrBand.SetRepeatOnPageTop(Value: Boolean);
begin
  RepeatOnPageTop := Value;
end;

function TvgrBand.GetRepeatOnPageBottom: Boolean;
begin
  Result := RepeatOnPageBottom;
end;

procedure TvgrBand.SetRepeatOnPageBottom(Value: Boolean);
begin
  RepeatOnPageBottom := Value;
end;

function TvgrBand.GetPrintWithNextSection: Boolean;
begin
  Result := PrintWithNextSection;
end;

procedure TvgrBand.SetPrintWithNextSection(Value: Boolean);
begin
  PrintWithNextSection := Value;
end;

function TvgrBand.GetPrintWithPreviosSection: Boolean;
begin
  Result := PrintWithPreviosSection;
end;

procedure TvgrBand.SetPrintWithPreviosSection(Value: Boolean);
begin
  PrintWithPreviosSection := Value;
end;

procedure TvgrBand.RecalcLevels;
begin
  if Bands <> nil then
    Bands.RecalcLevels;
end;

procedure TvgrBand.RecalcLevel;
var
  ABand, AParentBand: TvgrBand;
  I, ALevel, ASortIndex: Integer;

  function GetBandLevel(ABand: TvgrBand): Integer;
  begin
    Result := 0;
    while ABand.ParentBand <> nil do
    begin
      Inc(Result);
      ABand := ABand.ParentBand;
    end;
  end;

begin
  AParentBand := ParentBand;
  if AParentBand <> nil then
    FSection.Level := AParentBand.Level - 1
  else
  begin
    FSection.Level := 0;
    ASortIndex := SortIndex;
    I := ASortIndex + 1;
    while I < Bands.FSortedList.Count do
    begin
      ABand := TvgrBand(Bands.FSortedList[I]);
      if not ABand.InBand(Self) then
        break;
      ALevel := GetBandLevel(ABand);
      if ALevel > FSection.Level then
        FSection.Level := ALevel;
      Inc(I);
    end;
  end;
end;

procedure TvgrBand.SetStartPos(Value: Integer);
begin
  if (csLoading in ComponentState) or FBands.FMassUpdateState then
    FSection.StartPos := Value
  else
    if (StartPos <> Value) and (Value <= EndPos) then
    begin
      BeforeChange;
      FSection.StartPos := Value;
      RecalcLevels;
      AfterChange;
    end;
end;

procedure TvgrBand.SetEndPos(Value: Integer);
begin
  if (csLoading in ComponentState) or FBands.FMassUpdateState then
    FSection.EndPos := Value
  else
    if (EndPos <> Value) and (Value >= StartPos) then
    begin
      BeforeChange;
      FSection.EndPos := Value;
      RecalcLevels;
      AfterChange;
    end;
end;

function TvgrBand.GetParentBand: TvgrBand;
var
  I, ASortIndex: Integer;
begin
  ASortIndex := SortIndex;
  if ASortIndex = 0 then
    Result := nil
  else
  begin
    I := ASortIndex - 1;
    while (I >= 0) and not InBand(TvgrBand(Bands.FSortedList[I])) do
      Dec(I);
    if I >= 0 then
      Result := TvgrBand(Bands.FSortedList[I])
    else
      Result := nil;
  end;
end;

function TvgrBand.GetFirstVector: Integer;
begin
  Result := FFirstVector;
  Table.StoreRangeForSecondPass(Self);
end;

function TvgrBand.GetEndVector: Integer;
begin
  Result := FEndVector;
  Table.StoreRangeForSecondPass(Self);
end;

procedure TvgrBand.SetStartEndVectors(AStartPos, AEndPos: Integer);
begin
  FFirstVector := AStartPos;
  FEndVector := AEndPos;
end;

procedure TvgrBand.SetSkipGenerate(Value: Boolean);
begin
  if FSkipGenerate <> Value then
  begin
    BeforeChange;
    FSkipGenerate := Value;
    AfterChange;
  end;
end;

procedure TvgrBand.SetRowColAutoSize(Value: Boolean);
begin
  if FRowColAutoSize <> Value then
  begin
    BeforeChange;
    FRowColAutoSize := Value;
    AfterChange;
  end;
end;

function TvgrBand.GetTemplateTable: TvgrTemplateTable;
begin
  Result := TemplateWorksheet.FTemplateTable;
end;

procedure TvgrBand.SetEvents(Value: TvgrBandEvents);
begin
  BeforeChange;
  FEvents.Assign(Value);
  AfterChange;
end;

function TvgrBand.GetSortIndex: Integer;
begin
  Result := Bands.FSortedList.IndexOf(Self);
end;

function TvgrBand.GetIndex: Integer;
begin
  Result := Bands.IndexOf(Self);
end;

function TvgrBand.GetHorizontal: Boolean;
begin
  Result := Bands = TemplateWorksheet.HorzBands; 
end;

// IvgrWBListItem
function TvgrBand.GetWorksheet: TvgrWorksheet;
begin
  Result := TemplateWorksheet;
end;

function TvgrBand.GetItemIndex: Integer;
begin
  Result := FBands.IndexOf(Self);
end;

function TvgrBand.GetWorkbook: TvgrWorkbook;
begin
  Result := Template;
end;

function TvgrBand.GetItemData: Pointer;
begin
  Result := @FSection;
end;

function TvgrBand.GetStyleData: Pointer;
begin
  Result := nil;
end;

function TvgrBand.GetValid: Boolean;
begin
  Result := True;
end;

procedure TvgrBand.SetInvalid;
begin
end;

// IvgrSection
function TvgrBand.GetStartPos: Integer;
begin
  Result := FSection.StartPos;
end;

function TvgrBand.GetEndPos: Integer;
begin
  Result := FSection.EndPos;
end;

function TvgrBand.GetLevel: Integer;
begin
  Result := FSection.Level;
end;

function TvgrBand.GetParent: IvgrSection;
begin
  Result := ParentBand;
end;

function TvgrBand.GetText: string;
begin
  Result := Format('%s (%s)', [GetCaption, Name]);
end;

function TvgrBand.GetDispIdOfName(const AName: String): Integer;
begin
  // TvgrBand specific properties
  if AnsiCompareText(AName, 'GenBegin') = 0 then
    Result := cs_TvgrBand_GenBegin
  else
  if AnsiCompareText(AName, 'GenEnd') = 0 then
    Result := cs_TvgrBand_GenEnd
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrBand.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrBand, siTvgrBandLength);
end;

function TvgrBand.DoInvoke (DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrBand_GenBegin:
      AResult := GenBegin;
    cs_TvgrBand_GenEnd:
      AResult := GenEnd;
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

function TvgrBand.GetFlags(Index: Integer): Boolean;
begin
  Result := (FSection.Flags and (1 shl Index)) <> 0;
end;

procedure TvgrBand.SetCreateSections(Value: Boolean);
begin
  if Value <> FCreateSections then
  begin
    BeforeChange;
    FCreateSections := Value;
    AfterChange;
  end;
end;

procedure TvgrBand.SetFlags(Index: Integer; Value: Boolean);
var
  AMask: Word;
begin
  AMask := 1 shl Index;
  if ((FSection.Flags and AMask) <> 0) xor Value then
  begin
    BeforeChange;
    if Value then
      FSection.Flags := FSection.Flags or AMask
    else
      FSection.Flags := FSection.Flags and not AMask;
    AfterChange;
  end;
end;

procedure TvgrBand.Assign(ASource: IvgrSection);
begin
  BeforeChange;
  FSection.Flags := pvgrSection(ASource.ItemData).Flags;
  AfterChange;
end;

procedure TvgrBand.DoBeforeGenerate;
begin
  FEvents.DoBeforeGenerate;
end;

procedure TvgrBand.DoAfterGenerate;
begin
  FEvents.DoAfterGenerate;
end;

procedure TvgrBand.InternalGenerate(ABandVector: TvgrTemplateVector);
var
  I: Integer;
  AVector: TvgrTemplateVector;
begin
  for I := 0 to ABandVector.SubVectorCount - 1 do
  begin
    AVector := ABandVector.SubVectors[I];
    if AVector.Band = nil then
      AVector.Write
    else
      AVector.Band.Generate(AVector);
  end;
end;

function TvgrBand.GetCreateSection: Boolean;
begin
  Result := CreateSections;
end;

procedure TvgrBand.Generate(ABandVector: TvgrTemplateVector);
var
  ACreateSectionFlag: Boolean;
begin
  if not SkipGenerate then
  begin
    DoBeforeGenerate;
    if not SkipGenerate then
    begin
      ACreateSectionFlag := GetCreateSection;
      if ACreateSectionFlag then
        ABandVector.StartSection;

      InternalGenerate(ABandVector);

      if ACreateSectionFlag then
        ABandVector.EndSection;
      DoAfterGenerate;
    end;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrDetailBand
//
/////////////////////////////////////////////////
function TvgrDetailBand.CreateEvents: TvgrBandEvents;
begin
  Result := TvgrDetailBandEvents.Create(Self);
end;

function TvgrDetailBand.GetEvents: TvgrDetailBandEvents;
begin
  Result := TvgrDetailBandEvents(inherited Events);
end;

procedure TvgrDetailBand.SetEvents(Value: TvgrDetailBandEvents);
begin
  Events.Assign(Value);
end;

procedure TvgrDetailBand.SetDataSet(Value: TDataSet);
begin
  if FDataSet <> Value then
  begin
    BeforeChange;
    FDataSet := Value;
    AfterChange;
  end;
end;

procedure TvgrDetailBand.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if (AOperation = opRemove) and (FDataSet = AComponent) then
    FDataSet := nil;
end;

procedure TvgrDetailBand.InternalGenerate(ABandVector: TvgrTemplateVector);
var
  I: Integer;
  AVector, AGroupBandVector, ADataBandVector: TvgrTemplateVector;
begin
  AGroupBandVector := ABandVector.FindBandVector(TvgrGroupBand);
  if (AGroupBandVector <> nil) and (DataSet <> nil) then
  begin
    for I := 0 to ABandVector.SubVectorCount - 1 do
    begin
      AVector := ABandVector.SubVectors[I];
      if AVector.Band = nil then
        AVector.Write
      else
      begin
        if AVector = AGroupBandVector then
        begin
          InitDataSet;
          while not DataSet.Eof do
            AVector.Band.Generate(AVector);
        end
        else
          AVector.Band.Generate(AVector);
      end;
    end;
  end
  else
  begin
    ADataBandVector := ABandVector.FindBandVector(TvgrDataBand);
    if (ADataBandVector <> nil) and (DataSet <> nil) then
    begin
      for I := 0 to ABandVector.SubVectorCount - 1 do
      begin
        AVector := ABandVector.SubVectors[I];
        if AVector.Band = nil then
          AVector.Write
        else
        begin
          if AVector = ADataBandVector then
          begin
            InitDataSet;
            if not TvgrDataBand(ADataBandVector.Band).GenerateSectionForEachRecord then
              ADataBandVector.StartSection;
            while not DataSet.Eof do
            begin
              ADataBandVector.Band.Generate(ADataBandVector);
              DataSet.Next;
            end;
            if not TvgrDataBand(ADataBandVector.Band).GenerateSectionForEachRecord then
              ADataBandVector.EndSection;
//            ADataBandVector.Band.SecondPass;  
          end
          else
            AVector.Band.Generate(AVector);
        end;
      end;
    end
    else
      inherited;
  end;
end;

procedure TvgrDetailBand.InitDataSet;
var
  AInitializated: Boolean;
begin
  Events.DoInitDataSet(AInitializated);
  if not AInitializated and (FDataSet <> nil) then
  begin
    if FDataSet.Active then
      FDataSet.First
    else
      FDataSet.Open;
  end;
end;

class function TvgrDetailBand.GetCaption: string;
begin
  Result := 'Detail band';
end;

class function TvgrDetailBand.GetBitmapResName: string;
begin
  Result := 'VGR_DETAILBAND';
end;

/////////////////////////////////////////////////
//
// TvgrGroupBand
//
/////////////////////////////////////////////////
procedure TvgrGroupBand.ReadGroupFieldName(Reader: TReader);
begin
  FGroupExpression := Reader.ReadString;
end;

procedure TvgrGroupBand.WriteGroupFieldName(Writer: TWriter);
begin
end;

procedure TvgrGroupBand.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('GroupFieldName', ReadGroupFieldName, WriteGroupFieldName, False);
end;

function TvgrGroupBand.GetGroupValue: Variant;
begin
  if GenerateDataSet <> nil then
  begin
    Result := Template.Script.EvaluateExpression(GroupExpression, Null);
//    Result := GenerateDataSet.FieldByName(GroupFieldName).AsVariant
  end
  else
    Result := Null
end;

function TvgrGroupBand.GetDataSet: TDataSet;
var
  AParentBand: TvgrBand;
begin
  AParentBand := ParentBand;
  while (AParentBand <> nil) and not (AParentBand is TvgrDetailBand) do
    AParentBand := ParentBand;
  if AParentBand = nil then
    Result := nil
  else
    Result := TvgrDetailBand(AParentBand).DataSet;
end;

procedure TvgrGroupBand.SetGroupExpression(const Value: string);
begin
  if FGroupExpression <> Value then
  begin
    BeforeChange;
    FGroupExpression := Value;
    AfterChange;
  end;
end;

procedure TvgrGroupBand.InternalGenerate(ABandVector: TvgrTemplateVector);
var
  I: Integer;
  AValue: Variant;
  AVector, AGroupBandVector, ADataBandVector: TvgrTemplateVector;
  APriorFlag: Boolean;
begin
  ADataBandVector := ABandVector.FindBandVector(TvgrDataBand);
  if ADataBandVector <> nil then
  begin
    APriorFlag := False;
    for I := 0 to ABandVector.SubVectorCount - 1 do
    begin
      AVector := ABandVector.SubVectors[I];
      if AVector.Band = nil then
        AVector.Write
      else
      begin
        if AVector = ADataBandVector then
        begin
          AValue := GroupValue;
          if not TvgrDataBand(ADataBandVector.Band).GenerateSectionForEachRecord then
            ADataBandVector.StartSection;
          while not GenerateDataSet.Eof and (AValue = GroupValue) do
          begin
            ADataBandVector.Band.Generate(ADataBandVector);
            GenerateDataSet.Next;
          end;
          if not TvgrDataBand(ADataBandVector.Band).GenerateSectionForEachRecord then
            ADataBandVector.EndSection;
          if not GenerateDataSet.Eof then
          begin
            // The dataset is positioned on the next row, on the row of *NEXT* group
            GenerateDataSet.Prior;
            APriorFlag := True;
          end;
//          ADataBandVector.Band.SecondPass;
        end
        else
          AVector.Band.Generate(AVector);
      end;
    end;
    if APriorFlag then
      GenerateDataSet.Next;
  end
  else
  begin
    AGroupBandVector := ABandVector.FindBandVector(TvgrGroupBand);
    if AGroupBandVector <> nil then
    begin
      for I := 0 to ABandVector.SubVectorCount - 1 do
      begin
        AVector := ABandVector.SubVectors[I];
        if AVector.Band = nil then
          AVector.Write
        else
        begin
          if AVector = AGroupBandVector then
          begin
            AValue := GroupValue;
            while not DataSet.Eof and (AValue = GroupValue) do
              AVector.Band.Generate(AVector);
          end
          else
            AVector.Band.Generate(AVector);
        end;
      end;
    end
    else
      inherited;
  end;
end;

procedure TvgrGroupBand.Generate(ABandVector: TvgrTemplateVector);
begin
  FGenerateDataSet := DataSet;
  try
    inherited;
  finally
    FGenerateDataSet := nil;
  end;
end;

class function TvgrGroupBand.GetCaption: string;
begin
  Result := 'Group band';
end;

class function TvgrGroupBand.GetBitmapResName: string;
begin
  Result := 'VGR_GROUPBAND';
end;

/////////////////////////////////////////////////
//
// TvgrDataBand
//
/////////////////////////////////////////////////
procedure TvgrDataBand.SetGenerateSectionForEachRecord(Value: Boolean);
begin
  if FGenerateSectionForEachRecord <> Value then
  begin
    BeforeChange;
    FGenerateSectionForEachRecord := Value;
    AfterChange;
  end;
end;
 
function TvgrDataBand.GetCreateSection: Boolean;
begin
  Result := FGenerateSectionForEachRecord and CreateSections;
end;

procedure TvgrDataBand.Generate(ABandVector: TvgrTemplateVector);
begin
  inherited;
end;

class function TvgrDataBand.GetCaption: string;
begin
  Result := 'Data band';
end;

class function TvgrDataBand.GetBitmapResName: string;
begin
  Result := 'VGR_DATABAND';
end;

/////////////////////////////////////////////////
//
// TvgrBands
//
/////////////////////////////////////////////////
constructor TvgrBands.Create(AWorksheet: TvgrWorksheet);
begin
  inherited Create(AWorksheet);
  FList := TList.Create;
  FSortedList := TList.Create;
  EndMassUpdate;
end;

destructor TvgrBands.Destroy;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    ByIndex[I].FBands := nil;
  FreeAndNil(FList);
  FreeAndNil(FSortedList);
  inherited;
end;

procedure TvgrBands.Align;
begin
end;

procedure TvgrBands.DeleteRows(AStartPos, AEndPos: Integer);
var
  I, ADelta: Integer;
  ABand: TvgrBand;
begin
  I := 0;
  ADelta := AEndPos - AStartPos + 1;
  while I < Count do
  begin
    ABand := ByIndex[I];
    if (ABand.StartPos >= AStartPos) and (ABand.EndPos <= AEndPos) then
    begin
      ABand.Free;
      continue;
    end
    else
      if ABand.StartPos > AEndPos then
      begin
        ABand.StartPos := ABand.StartPos - ADelta;
        ABand.EndPos := ABand.EndPos - ADelta;
      end
      else
        if (ABand.StartPos <= AStartPos) and (ABand.EndPos >= AEndPos) then
        begin
          ABand.EndPos := ABand.EndPos - ADelta;
        end;
    Inc(I);
  end;
end;

procedure TvgrBands.InternalInsertLines(AIndexBefore, ACount: Integer);
var
  I: Integer;
  AItemStartPos, AItemEndPos: Integer;
  ABand: TvgrBand;
begin
  BeginMassUpdate;
  for I := 0 to Count - 1 do
  begin
    ABand := ByIndex[I];
    AItemStartPos := ABand.StartPos;
    AItemEndPos := ABand.EndPos;
    if (AItemEndPos >= AIndexBefore) then
      ABand.EndPos := ABand.EndPos + 1;
    if (AItemStartPos >= (AIndexBefore+1)) then
      ABand.StartPos := ABand.StartPos + 1;
  end;
  EndMassUpdate;
end;

function TvgrBands.IndexOf(ABand: TvgrBand): Integer;
begin
  Result := FList.IndexOf(ABand);
end;

function TvgrBands.GetBand(StartPos, EndPos: Integer): TvgrBand;
begin
  Result := TvgrBand(Find(StartPos, EndPos));
end;

function TvgrBands.GetBandByIndex(Index: Integer): TvgrBand;
begin
  Result := TvgrBand(FList[Index]);
end;

function TvgrBands.GetTemplateWorksheet: TvgrReportTemplateWorksheet;
begin
  Result := TvgrReportTemplateWorksheet(Worksheet);
end;

function TvgrBands.GetTemplate: TvgrReportTemplate;
begin
  Result := TvgrReportTemplate(TemplateWorksheet.Workbook);
end;

procedure TvgrBands.BeginMassUpdate;
begin
  FMassUpdateState := True;
end;

procedure TvgrBands.EndMassUpdate;
begin
  FMassUpdateState := False;
end;

function TvgrBands.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TvgrBands.GetItem(StartPos, EndPos: Integer): IvgrSection;
begin
  Result := Find(StartPos, EndPos);
end;

function TvgrBands.GetByIndex(Index: Integer): IvgrSection;
begin
  Result := TvgrBand(FList[Index]);
end;

function TvgrBands.GetMaxLevel: Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to Count - 1 do
    if Result < ByIndex[I].Level then
      Result := ByIndex[I].Level;
end;

procedure TvgrBands.FindAndCallBack(AStartPos, AEndPos: Integer; CallBackProc : TvgrFindListItemCallBackProc; AData: Pointer);
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    with ByIndex[I] do
      if RectOverRect(Rect(StartPos, StartPos, EndPos, EndPos), Rect(AStartPos, AStartPos, AEndPos, AEndPos)) then
        CallBackProc(ByIndex[I] as IvgrSection, I, AData);
end;

function TvgrBands.FindBand(AStartPos, AEndPos: Integer): TvgrBand;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    with ByIndex[I] do
      if (StartPos >= AStartPos) and (EndPos <= AEndPos) then
      begin
        Result := ByIndex[I];
        exit;
      end;
  Result := nil;
end;

function TvgrBands.FindBandAt(AStartPos, AEndPos: Integer): TvgrBand;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    with ByIndex[I] do
      if (StartPos <= AStartPos) and (EndPos >= AEndPos) then
      begin
        Result := ByIndex[I];
        exit;
      end;
  Result := nil;
end;

function TvgrBands.Find(AStartPos, AEndPos: Integer) : IvgrSection;
begin
  Result := FindBand(AStartPos, AEndPos) as IvgrSection;
end;

procedure TvgrBands.Delete(AStartPos, AEndPos: Integer);
var
  ABand: TvgrBand;
begin
  ABand := FindBand(AStartPos, AEndPos);
  if ABand <> nil then
    ABand.Free;
end;

function TvgrBands.SearchCurrentOuter(AStartPos, AEndPos, ALevel: Integer): Integer;
var
  I: Integer;
begin
  Result := -1;
  for I := 0 to Count - 1 do
    with ByIndex[I] do
      if (StartPos >= AStartPos) and (EndPos <= AEndPos) and (ALevel = Level) then
      begin
        Result := I;
        exit;
      end;
end;

function TvgrBands.SearchDirectOuter(AStartPos, AEndPos, ALevel: Integer): Integer;
var
  ABand: TvgrBand;
begin
  ABand := FindBand(AStartPos, AEndPos);
  if (ABand = nil) or (ABand.ParentBand = nil) then
    Result := -1
  else
    Result := ABand.ParentBand.Index;
end;

function LevelSortProc(Item1, Item2: Pointer): Integer;
begin
  Result :=  TvgrBand(Item1).StartPos - TvgrBand(Item2).StartPos;
  if Result = 0 then
  begin
    Result := TvgrBand(Item2).EndPos - TvgrBand(Item1).EndPos;
    if Result = 0 then
      Result := TvgrBand(Item1).Index - TvgrBand(Item2).Index;
  end;
end;

function TvgrBands.Add(ABandClass: TvgrBandClass): TvgrBand;
var
  ChangeInfo: TvgrWorkbookChangeInfo;
begin
  ChangeInfo.ChangesType := GetCreateChangesType;
  ChangeInfo.ChangedObject := nil;
  ChangeInfo.ChangedInterface := nil;
  Template.BeforeChange(ChangeInfo);

  Result := ABandClass.Create(Template.Owner);
  Result.Name := GetUniqueComponentName(Result);
  Result.FBands := Self;
  FList.Add(Result);
  FSortedList.Add(Result);
  FSortedList.Sort(LevelSortProc);
  RecalcLevels;

  ChangeInfo.ChangedObject := Result;
  ChangeInfo.ChangedInterface := Result;
  Template.AfterChange(ChangeInfo);
end;

procedure TvgrBands.Remove(ABand: TvgrBand);
begin
  if FList <> nil then
    FList.Remove(ABand);
  if FSortedList <> nil then
  begin
    FSortedList.Remove(ABand);
    RecalcLevels;
  end;
end;

procedure TvgrBands.RecalcLevels;
var
  I, AMaxParentLevel: Integer;
begin
  FSortedList.Sort(LevelSortProc);
  for I := 0 to FSortedList.Count - 1 do
    TvgrBand(FSortedList[I]).RecalcLevel;

  AMaxParentLevel := 0;
  for I := 0 to FSortedList.Count - 1 do
    if (TvgrBand(FSortedList[I]).ParentBand = nil) and
       (TvgrBand(FSortedList[I]).FSection.Level > AMaxParentLevel) then
      AMaxParentLevel := TvgrBand(FSortedList[I]).FSection.Level;

  for I := 0 to FSortedList.Count - 1 do
    if TvgrBand(FSortedList[I]).ParentBand = nil then
      TvgrBand(FSortedList[I]).FSection.Level := AMaxParentLevel;
end;

procedure TvgrBands.SaveToStream(AStream: TStream);
begin
end;

procedure TvgrBands.LoadFromStream(AStream: TStream; ADataStorageVersion: Integer);
begin
end;

/////////////////////////////////////////////////
//
// TvgrReportTemplateWorksheetGenerateEvents
//
/////////////////////////////////////////////////
constructor TvgrReportTemplateWorksheetGenerateEvents.Create(AWorksheet: TvgrReportTemplateWorksheet);
begin
  inherited Create;
  FWorksheet := AWorksheet;
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.Assign(Source: TPersistent);
begin
  if Source is TvgrReportTemplateWorksheetGenerateEvents then
    with TvgrReportTemplateWorksheetGenerateEvents(Source) do
    begin
      Self.FBeforeGenerateScriptProcName := BeforeGenerateScriptProcName;
      Self.FAfterGenerateScriptProcName := AfterGenerateScriptProcName;
      Self.FAddWorkbookWorksheetScriptProcName := AddWorkbookWorksheetScriptProcName;
    end;
end;

function TvgrReportTemplateWorksheetGenerateEvents.GetOwner: TPersistent;
begin
  Result := FWorksheet;
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.ReadAfterAddWorksheet(Reader: TReader);
begin
  Reader.ReadString;
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.WriteAfterAddWorksheet(Writer: TWriter);
begin
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.DefineProperties(Filer: TFiler);
begin
  inherited;
  Filer.DefineProperty('AfterAddWorksheetScriptProcName', ReadAfterAddWorksheet, WriteAfterAddWorksheet, False);
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.DoBeforeGenerate;
var
  AParameters: TvgrVariantDynArray;
begin
  FWorksheet.Template.DoTemplateWorksheetBeforeGenerate(FWorksheet);

  if FBeforeGenerateScriptProcName <> '' then
  begin
    SetLength(AParameters, 1);
    AParameters[0] := FWorksheet as IDispatch;
    FWorksheet.Template.Script.RunProcedure(FBeforeGenerateScriptProcName, AParameters);
  end;
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.DoAfterGenerate(AWorksheet: TvgrWorksheet);
var
  AParameters: TvgrVariantDynArray;
begin
  FWorksheet.Template.DoTemplateWorksheetAfterGenerate(FWorksheet);

  if FAfterGenerateScriptProcName <> '' then
  begin
    SetLength(AParameters, 2);
    AParameters[0] := FWorksheet as IDispatch;
    AParameters[1] := AWorksheet as IDispatch;
    FWorksheet.Template.Script.RunProcedure(FAfterGenerateScriptProcName, AParameters);
  end;
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.DoAfterAddWorksheet(AWorksheet: TvgrWorksheet);
var
  AParameters: TvgrVariantDynArray;
begin
  FWorksheet.Template.DoAddWorkbookWorksheetEvent(FWorksheet, AWorksheet);

  if FAddWorkbookWorksheetScriptProcName <> '' then
  begin
    SetLength(AParameters, 2);
    AParameters[0] := FWorksheet as IDispatch;
    AParameters[1] := AWorksheet as IDispatch;
    FWorksheet.Template.Script.RunProcedure(FAddWorkbookWorksheetScriptProcName, AParameters);
  end;
end;

procedure TvgrReportTemplateWorksheetGenerateEvents.DoCustomGenerate(AWorksheet: TvgrWorksheet; var ADone: Boolean);
var
  AParameters: TvgrVariantDynArray;
  AByRef: TvgrBooleanDynArray;
begin
  FWorksheet.Template.DoCustomGenerate(FWorksheet, AWorksheet, ADone);
  if not ADone and (FCustomGenerateScriptProcName <> '') then
  begin
    SetLength(AParameters, 3);
    AParameters[0] := FWorksheet as IDispatch;
    AParameters[1] := AWorksheet as IDispatch;
    AParameters[2] := ADone;
    SetLength(AByRef, 3);
    AByRef[0] := False;
    AByRef[1] := False;
    AByRef[2] := True;
    FWorksheet.Template.Script.RunProcedure(FCustomGenerateScriptProcName, AParameters, AByRef);
    ADone := AParameters[2];
  end;
end;

/////////////////////////////////////////////////
//
// TvgrReportTemplateWorksheet
//
/////////////////////////////////////////////////
constructor TvgrReportTemplateWorksheet.Create(AOwner: TComponent);
begin
  inherited;
  FGenerateSize.cx := -1;
  FGenerateSize.cy := -1;
  FCurrentRangePlace := Rect(-1, -1, -1, -1);
  FEvents := TvgrReportTemplateWorksheetGenerateEvents.Create(Self);
end;

destructor TvgrReportTemplateWorksheet.Destroy;
begin
  while HorzBands.Count > 0 do
    HorzBands.ByIndex[0].Free;
  while VertBands.Count > 0 do
    VertBands.ByIndex[0].Free;
  FEvents.Free;
  inherited;
end;

procedure TvgrReportTemplateWorksheet.Loaded;
begin
  inherited;
end;

procedure TvgrReportTemplateWorksheet.CheckInOnCustomGenerateEvent;
begin
  if not FOnCustomGenerateEvent then
    raise Exception.Create(sse_Write_NotOnCustomGenerateEvent);
end;

procedure TvgrReportTemplateWorksheet.CheckBand(ABand: TvgrBand; AMustBeHorizontal: Boolean);
begin
  if ABand = nil then
  begin
    if AMustBeHorizontal then
      raise Exception.Create(sse_Write_HorizontalBandIsNotSpecified)
    else
      raise Exception.Create(sse_Write_VerticalBandIsNotSpecified);
  end;
  if ABand.TemplateWorksheet <> Self then
    raise Exception.CreateFmt(sse_Write_InvalidBandTemplateWorksheet, [ABand.Name, Self.Name]);
    
  if AMustBeHorizontal then
  begin
    if not ABand.Horizontal then
      raise Exception.CreateFmt(sse_Write_NotHorizontal, [ABand.Name]);
  end
  else
  begin
    if ABand.Horizontal then
      raise Exception.CreateFmt(sse_Write_NotVertical, [ABand.Name]);
  end;

  if ABand.TemplateVector = nil then
    raise Exception.CreateFmt(sse_Write_TemplateVectorIsNil, [ABand.Name]);
end;

procedure TvgrReportTemplateWorksheet.WriteLnBand(AHorizontalBand: TvgrBand);
begin
  CheckInOnCustomGenerateEvent;
  CheckBand(AHorizontalBand, True);

  AHorizontalBand.Generate(AHorizontalBand.TemplateVector);
  FTemplateTable.FWriteCellRow := nil;
end;

procedure TvgrReportTemplateWorksheet.WriteCell(AHorizontalBand: TvgrBand; AVerticalBand: TvgrBand);
begin
  CheckInOnCustomGenerateEvent;
  CheckBand(AHorizontalBand, True);
  CheckBand(AVerticalBand, False);

  AHorizontalBand.Table.FCurrentRow := AHorizontalBand.TemplateVector as TvgrTemplateRow;
  AVerticalBand.TemplateVector.Write;

  if (FTemplateTable.FWriteCellRow = nil) or
     (FTemplateTable.FWriteCellRow.Size < AHorizontalBand.TemplateVector.Size) then
    FTemplateTable.FWriteCellRow := AHorizontalBand.TemplateVector as TvgrTemplateRow;
end;

procedure TvgrReportTemplateWorksheet.WriteLn(ARowCount: Integer = -1);
begin
  FTemplateTable.WriteLn(FTemplateTable.FWriteCellRow.Size + 1);
  if ARowCount > 0 then
    FTemplateTable.FWorkbookRow := FTemplateTable.FWorkbookRow + ARowCount;
end;

procedure TvgrReportTemplateWorksheet.WriteLnCell(AHorizontalBand: TvgrBand; AVerticalBand: TvgrBand);
begin
  WriteCell(AHorizontalBand, AVerticalBand);
  WriteLn;
end;

function TvgrReportTemplateWorksheet.IsGenerateSizeWidthStored: Boolean;
begin
  with FGenerateSize, Dimensions do
    Result := (cx >= 0) and (cx <> Right);
end;

function TvgrReportTemplateWorksheet.IsGenerateSizeHeightStored: Boolean;
begin
  with FGenerateSize, Dimensions do
    Result := (cy >= 0) and (cy <> Bottom);
end;

function TvgrReportTemplateWorksheet.GetGenerateSizeWidth: Integer;
begin
  Result := FGenerateSize.cx;
end;

procedure TvgrReportTemplateWorksheet.SetGenerateSizeWidth(Value: Integer);
begin
  if FGenerateSize.cx <> Value then
  begin
    BeforeChangeProperty;
    FGenerateSize.cx := Value;
    AfterChangeProperty;
  end;
end;

function TvgrReportTemplateWorksheet.GetGenerateSizeHeight: Integer;
begin
  Result := FGenerateSize.cy;
end;

procedure TvgrReportTemplateWorksheet.SetGenerateSizeHeight(Value: Integer);
begin
  if FGenerateSize.cy <> Value then
  begin
    BeforeChangeProperty;
    FGenerateSize.cy := Value;
    AfterChangeProperty;
  end;
end;

procedure TvgrReportTemplateWorksheet.SetGenerateSize(Value: TSize);
begin
  if (Value.cx <> FGenerateSize.cx) or (Value.cy <> FGenerateSize.cy) then
  begin
    BeforeChangeProperty;
    FGenerateSize := Value;
    AfterChangeProperty;
  end;
end;

procedure TvgrReportTemplateWorksheet.SetEvents(Value: TvgrReportTemplateWorksheetGenerateEvents);
begin
  FEvents.Assign(Value);
end;

procedure TvgrReportTemplateWorksheet.GetChildren(Proc: TGetChildProc; Root: TComponent);
var
  I: Integer;
begin
  for I := 0 to HorzBands.Count - 1 do
    Proc(HorzBands.ByIndex[I]);
  for I := 0 to VertBands.Count - 1 do
    Proc(VertBands.ByIndex[I]);
end;

function TvgrReportTemplateWorksheet.GetDimensions : TRect;
var
  I: Integer;
begin
  Result := inherited GetDimensions;

  Result.Left := 0;
  Result.Top := 0;
  for I := 0 to HorzBands.Count - 1 do
    if HorzBands.ByIndex[I].EndPos > Result.Bottom then
      Result.Bottom := HorzBands.ByIndex[I].EndPos;
  for I := 0 to VertBands.Count - 1 do
    if VertBands.ByIndex[I].EndPos > Result.Right then
      Result.Right := VertBands.ByIndex[I].EndPos;

  if (FGenerateSize.cx >= 0) and (FGenerateSize.cy >= 0) then
  begin
    Result.Right := FGenerateSize.cx - 1;
    Result.Bottom := FGenerateSize.cy - 1;
  end;
end;

procedure TvgrReportTemplateWorksheet.DoBeforeGenerate;
begin
  FEvents.DoBeforeGenerate;
end;

procedure TvgrReportTemplateWorksheet.DoAfterGenerate(AWorksheet: TvgrWorksheet);
begin
  FEvents.DoAfterGenerate(AWorksheet);
end;

procedure TvgrReportTemplateWorksheet.DoAfterAddWorksheet(AWorksheet: TvgrWorksheet);
begin
  FEvents.DoAfterAddWorksheet(AWorksheet);
end;

procedure TvgrReportTemplateWorksheet.DoCustomGenerate(AWorksheet: TvgrWorksheet; var ADone: Boolean);
begin
  FOnCustomGenerateEvent := True;
  try
    FEvents.DoCustomGenerate(AWorksheet, ADone);
  finally
    FOnCustomGenerateEvent := False;
  end;
end;

function TvgrReportTemplateWorksheet.GetDispIdOfName(const AName: String): Integer;
begin
  if AnsiCompareText(AName, 'CurrentRow') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_CurrentRow
  else
  if AnsiCompareText(AName,'CurrentColumn') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_CurrentColumn
  else
  if AnsiCompareText(AName,'CurrentColumnCaption') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_CurrentColumnCaption
  else
  if AnsiCompareText(AName, 'WriteLnBand') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WriteLnBand
  else
  if AnsiCompareText(AName, 'WriteCell') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WriteCell
  else
  if AnsiCompareText(AName, 'WriteLnCell') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WriteLnCell
  else
  if AnsiCompareText(AName, 'WriteLn') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WriteLn
  else
  if AnsiCompareText(AName, 'WorkbookRow') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WorkbookRow
  else
  if AnsiCompareText(AName, 'WorkbookColumn') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WorkbookColumn
  else
  if AnsiCompareText(AName, 'WorkbookColumnCaption') = 0 then
    Result := cs_TvgrReportTemplateWorksheet_WorkbookColumnCaption
  else
    Result := inherited GetDispIdOfName(AName);
end;

function TvgrReportTemplateWorksheet.DoInvoke (DispId: Integer;
                      Flags: Integer;
                      var AParameters: TvgrOleVariantDynArray;
                      var AResult: OleVariant): HResult;
begin
  Result := S_OK;
  case DispId of
    cs_TvgrReportTemplateWorksheet_CurrentRow:
      AResult := FCurrentRangePlace.Top + 1;
    cs_TvgrReportTemplateWorksheet_CurrentColumn:
      AResult := FCurrentRangePlace.Left + 1;
    cs_TvgrReportTemplateWorksheet_CurrentColumnCaption:
      AResult := GetWorksheetColCaption(FCurrentRangePlace.Left);
    cs_TvgrReportTemplateWorksheet_WriteLnBand:
      begin
        WriteLnBand(OleVariantToObject(AParameters[0]) as TvgrBand);
      end;
    cs_TvgrReportTemplateWorksheet_WriteCell:
      begin
        WriteCell(OleVariantToObject(AParameters[0]) as TvgrBand,
                  OleVariantToObject(AParameters[1]) as TvgrBand);
      end;
    cs_TvgrReportTemplateWorksheet_WriteLnCell:
      begin
        WriteLnCell(OleVariantToObject(AParameters[0]) as TvgrBand,
                    OleVariantToObject(AParameters[1]) as TvgrBand);
      end;
    cs_TvgrReportTemplateWorksheet_WriteLn:
      begin
        if Length(AParameters) > 0 then
          WriteLn(AParameters[0])
        else
          WriteLn;
      end;
    cs_TvgrReportTemplateWorksheet_WorkbookRow:
      begin
        AResult := WorkbookRow
      end;
    cs_TvgrReportTemplateWorksheet_WorkbookColumn:
      begin
        AResult := WorkbookColumn
      end;
    cs_TvgrReportTemplateWorksheet_WorkbookColumnCaption:
      begin
        AResult := WorkbookColumnCaption
      end;
    else
      Result := inherited DoInvoke(DispId, Flags, AParameters, AResult);
  end;
end;

function TvgrReportTemplateWorksheet.DoCheckScriptInfo(DispId: Integer;
                           Flags: Integer;
                           AParametersCount: Integer): HResult;
begin
  Result := inherited DoCheckScriptInfo(DispId, Flags, AParametersCount);
  if Result = S_OK then
    Result := CheckScriptInfo(DispId, Flags, AParametersCount, @siTvgrReportTemplateWorksheet, siTvgrReportTemplateWorksheetLength);
end;

function TvgrReportTemplateWorksheet.GetWorkbookRow: Integer;
begin
  Result := FTemplateTable.FWorkbookRow;
end;

function TvgrReportTemplateWorksheet.GetWorkbookColumn: Integer;
begin
  Result := FTemplateTable.FWorkbookCol;
end;

function TvgrReportTemplateWorksheet.GetWorkbookColumnCaption: string;
begin
  Result := GetWorksheetColCaption(FTemplateTable.FWorkbookCol);
end;

procedure TvgrReportTemplateWorksheet.AlignBands;
begin
end;

function TvgrReportTemplateWorksheet.GetSectionsClass: TvgrSectionsClass;
begin
  Result := TvgrBands;
end;

function TvgrReportTemplateWorksheet.GetHorzBands: TvgrBands;
begin
  Result := TvgrBands(HorzSectionsList);
end;

function TvgrReportTemplateWorksheet.GetVertBands: TvgrBands;
begin
  Result := TvgrBands(VertSectionsList);
end;

function TvgrReportTemplateWorksheet.GetTemplate: TvgrReportTemplate;
begin
  Result := TvgrReportTemplate(Workbook);
end;

function TvgrReportTemplateWorksheet.CreateHorzBand(ABandClass: TvgrBandClass): TvgrBand;
begin
  Result := HorzBands.Add(ABandClass);
end;

function TvgrReportTemplateWorksheet.CreateVertBand(ABandClass: TvgrBandClass): TvgrBand;
begin
  Result := VertBands.Add(ABandClass);
end;

function TvgrReportTemplateWorksheet.FindVertBandAt(AStartPos, AEndPos: Integer): TvgrBand;
begin
  Result := VertBands.FindBandAt(AStartPos, AEndPos)
end;

function TvgrReportTemplateWorksheet.FindHorzBandAt(AStartPos, AEndPos: Integer): TvgrBand;
begin
  Result := HorzBands.FindBandAt(AStartPos, AEndPos)
end;

/////////////////////////////////////////////////
//
// TvgrReportTemplate
//
/////////////////////////////////////////////////
constructor TvgrReportTemplate.Create(AOwner: TComponent);
begin
  inherited;
  FScript := TvgrScript.Create;
  FScript.OnChange := OnScriptChange;
  FScript.Root := Self;
end;

destructor TvgrReportTemplate.Destroy;
begin
  FreeAndNil(FScript);
  inherited;
end;

procedure TvgrReportTemplate.OnScriptChange(Sender: TObject);
begin
  AfterChangeReportTemplate;
end;

procedure TvgrReportTemplate.Clear;
begin
  BeginUpdate;
  try
    inherited;
    FScript.Script.Clear;
    FScript.Language := '';
  finally
    EndUpdate;
  end;
end;

function TvgrReportTemplate.GetOnScriptGetObject: TvgrScriptGetObject;
begin
  Result := FScript.OnScriptGetObject;
end;

procedure TvgrReportTemplate.SetOnScriptGetObject(Value: TvgrScriptGetObject);
begin
  FScript.OnScriptGetObject := Value;
end;

function TvgrReportTemplate.GetOnScriptErrorInExpression: TvgrScriptErrorInExpression;
begin
  Result := FScript.OnScriptErrorInExpression;
end;

procedure TvgrReportTemplate.SetOnScriptErrorInExpression(Value: TvgrScriptErrorInExpression);
begin
  FScript.OnScriptErrorInExpression := Value;
end;

function TvgrReportTemplate.GetOnGetAvailableComponents: TvgrGetAvailableComponentsEvent;
begin
  Result := FScript.OnGetAvailableComponents;
end;

procedure TvgrReportTemplate.SetOnGetAvailableComponents(Value: TvgrGetAvailableComponentsEvent);
begin
  FScript.OnGetAvailableComponents := Value;
end;

procedure TvgrReportTemplate.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if (Script <> nil) and (AComponent = Script.AliasManager) then
      Script.AliasManager := nil;
  end;
end;

function TvgrReportTemplate.GetWorksheetClass: TvgrWorksheetClass;
begin
  Result := TvgrReportTemplateWorksheet;
end;

procedure TvgrReportTemplate.BeforeChangeReportTemplate;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := vgrwcChangeReportTemplate;
  AChangeInfo.ChangedObject := Self;
  AChangeInfo.ChangedInterface := nil;
  BeforeChange(AChangeInfo);
end;

procedure TvgrReportTemplate.AfterChangeReportTemplate;
var
  AChangeInfo: TvgrWorkbookChangeInfo;
begin
  AChangeInfo.ChangesType := vgrwcChangeReportTemplate;
  AChangeInfo.ChangedObject := Self;
  AChangeInfo.ChangedInterface := nil;
  AfterChange(AChangeInfo);
end;

procedure TvgrReportTemplate.SetScript(Value: TvgrScript);
begin
  FScript.Assign(Value);
end;

function TvgrReportTemplate.GetWorksheet(Index: Integer): TvgrReportTemplateWorksheet;
begin
  Result := TvgrReportTemplateWorksheet(inherited Worksheets[Index]);
end;

procedure TvgrReportTemplate.BeginGenerate;
begin
  Script.BeginScriptExecuting;
end;

procedure TvgrReportTemplate.EndGenerate;
begin
  Script.EndScriptExecuting;
end;

procedure TvgrReportTemplate.DoBandBeforeGenerate(ABand: TvgrBand);
begin
  if Assigned(FOnBandBeforeGenerate) then
    FOnBandBeforeGenerate(Self, ABand);
end;

procedure TvgrReportTemplate.DoBandAfterGenerate(ABand: TvgrBand);
begin
  if Assigned(FOnBandAfterGenerate) then
    FOnBandAfterGenerate(Self, ABand);
end;

procedure TvgrReportTemplate.DoInitDetailBandDataSetEvent(ABand: TvgrDetailBand; var AInitializated: Boolean);
begin
  if Assigned(FOnInitDetailBandDataSet) then
    FOnInitDetailBandDataSet(Self, ABand, AInitializated);
end;

procedure TvgrReportTemplate.DoTemplateWorksheetBeforeGenerate(ATemplateWorksheet: TvgrReportTemplateWorksheet);
begin
  if Assigned(FOnTemplateWorksheetBeforeGenerate) then
    FOnTemplateWorksheetBeforeGenerate(Self, ATemplateWorksheet);
end;

procedure TvgrReportTemplate.DoTemplateWorksheetAfterGenerate(ATemplateWorksheet: TvgrReportTemplateWorksheet);
begin
  if Assigned(FOnTemplateWorksheetAfterGenerate) then
    FOnTemplateWorksheetAfterGenerate(Self, ATemplateWorksheet);
end;

procedure TvgrReportTemplate.DoAddWorkbookWorksheetEvent(ATemplateWorksheet: TvgrReportTemplateWorksheet; AResultWorksheet: TvgrWorksheet);
begin
  if Assigned(FOnAddWorkbookWorksheet) then
    FOnAddWorkbookWorksheet(Self, ATemplateWorksheet, AResultWorksheet);
end;

procedure TvgrReportTemplate.DoCustomGenerate(ATemplateWorksheet: TvgrReportTemplateWorksheet; AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
begin
  ADone := False;
  if Assigned(FOnCustomGenerate) then
    FOnCustomGenerate(Self, ATemplateWorksheet, AResultWorksheet, ADone);
end;

/////////////////////////////////////////////////
//
// TvgrTemplateVector
//
/////////////////////////////////////////////////
constructor TvgrTemplateVector.Create(ATable: TvgrTemplateTable);
begin
  inherited Create;
  FTable := ATable;
  FSubVectors := TList.Create;
end;

destructor TvgrTemplateVector.Destroy;
begin
  FreeList(FSubVectors);
  inherited;
end;

function TvgrTemplateVector.GetSubVectorCount: Integer;
begin
  Result := FSubVectors.Count;
end;

function TvgrTemplateVector.GetSubVector(Index: Integer): TvgrTemplateVector;
begin
  Result := TvgrTemplateVector(FSubVectors[Index]);
end;

function TvgrTemplateVector.GetEngine: TvgrReportEngine;
begin
  Result := Table.Engine;
end;

function TvgrTemplateVector.AddSubVector(AStartPos, ASize: Integer; ABand: TvgrBand; AParentVector: TvgrTemplateVector): TvgrTemplateVector;
begin
  Result := GetSelfClass.Create(Table);
  Result.FStartPos := AStartPos;
  Result.FSize := ASize;
  Result.FBand := ABand;
  Result.FParentVector := AParentVector;
{$IFDEF VGR_DEBUG}
  if ABand = nil then
    DbgStrFmt('Vector: StartPos = %d, Size = %d, ABand = NIL', [AStartPos, ASize, ABand])
  else
    DbgStrFmt('Vector: StartPos = %d, Size = %d, ABand = %s', [AStartPos, ASize, ABand.Name]);
{$ENDIF}    
  FSubVectors.Add(Result);
end;

function TvgrTemplateVector.GetRowColAutoSize: Boolean;
var
  AVector: TvgrTemplateVector;
begin
//  Result := True; exit;
  AVector := Self;
  while AVector <> nil do
  begin
    if (AVector.Band <> nil) and AVector.Band.RowColAutoSize then
    begin
      Result := AVector.Band.RowColAutoSize;
      exit;
    end;
    AVector := AVector.ParentVector;
  end;
  Result := False;
end;

function TvgrTemplateVector.FindBandVector(ABandClass: TvgrBandClass): TvgrTemplateVector;
var
  I: Integer;
begin
  for I := 0 to SubVectorCount - 1 do
    if SubVectors[I].Band is ABandClass then
    begin
      Result := SubVectors[I];
      exit;
    end;
  Result := nil;
end;

procedure TvgrTemplateVector.Clear;
var
  I: Integer;
begin
  for I := 0 to SubVectorCount - 1 do
    SubVectors[I].Free;
  FSubVectors.Clear;
end;

procedure TvgrTemplateVector.StartSection;
begin
  GetStack.Push(GetPos);
end;

procedure TvgrTemplateVector.EndSection;
var
  AOldPos, ANewPos: Integer;
  ASection: IvgrSection;
begin
  AOldPos := GetStack.Pop;
  ANewPos := GetPos;
  if FBand.CreateSections then
  begin
    ASection := GetSections[AOldPos, ANewPos - 1];
    with ASection do
    begin
      RepeatOnPageTop := Band.RepeatOnPageTop;
      RepeatOnPageBottom := Band.RepeatOnPageBottom;
      PrintWithNextSection := Band.PrintWithNextSection;
      PrintWithPreviosSection := Band.PrintWithPreviosSection;
    end;
  end;

  Band.SetStartEndVectors(AOldPos + 1, ANewPos);

  Band.SecondPass;
end;

/////////////////////////////////////////////////
//
// TvgrTemplateRow
//
/////////////////////////////////////////////////
function TvgrTemplateRow.GetSelfClass: TvgrTemplateVectorClass;
begin
  Result := TvgrTemplateRow;
end;

function TvgrTemplateRow.GetStack: TvgrSectionsStack;
begin
  Result := Table.FHorzStack;
end;

function TvgrTemplateRow.GetPos: Integer;
begin
  Result := Table.FWorkbookRow;
end;

function TvgrTemplateRow.GetSections: TvgrSections;
begin
  Result := Table.ResultWorksheet.HorzSectionsList;
end;

procedure TvgrTemplateRow.Write;
var
  I: Integer;
  ACol: TvgrTemplateCol;
begin
  Table.FWorkbookCol := 0;
  if Table.ColCount = 0 then
  begin
    Table.WriteRanges(0, StartPos, MaxInt - 1, Size);
    Table.WriteCols(0, Table.ResultWorksheet.Dimensions.Right);
  end
  else
  begin
    Table.CurrentRow := Self;
    for I := 0 to Table.ColCount - 1 do
    begin
      ACol := Table.Cols[I];
      if ACol.Band = nil then
        ACol.Write
      else
        ACol.Band.Generate(ACol);
    end;
  end;
  Table.WriteRows(StartPos, Size, RowColAutoSize);
  Table.WriteLn(Size + 1);
end;

/////////////////////////////////////////////////
//
// TvgrTemplateCol
//
/////////////////////////////////////////////////
function TvgrTemplateCol.GetSelfClass: TvgrTemplateVectorClass;
begin
  Result := TvgrTemplateCol;
end;

function TvgrTemplateCol.GetStack: TvgrSectionsStack;
begin
  Result := Table.FVertStack;
end;

function TvgrTemplateCol.GetPos: Integer;
begin
  Result := Table.FWorkbookCol;
end;

function TvgrTemplateCol.GetSections: TvgrSections;
begin
  Result := Table.ResultWorksheet.VertSectionsList;
end;

procedure TvgrTemplateCol.Write;
begin
  Table.WriteCols(StartPos, Size);
  Table.WriteRanges(StartPos, Table.CurrentRow.StartPos, Size, Table.CurrentRow.Size);
end;

/////////////////////////////////////////////////
//
// TvgrSectionsStack
//
/////////////////////////////////////////////////
constructor TvgrSectionsStack.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

destructor TvgrSectionsStack.Destroy;
begin
  FreeAndNil(FList);
  inherited;
end;

procedure TvgrSectionsStack.Push(APos: Integer);
begin
  FList.Add(Pointer(APos));
end;

function TvgrSectionsStack.Pop: Integer;
begin
  Result := Integer(FList[FList.Count - 1]);
  FList.Delete(FList.Count - 1);
end;

procedure TvgrSectionsStack.Clear;
begin
  FList.Clear;
end;

/////////////////////////////////////////////////
//
// TvgrTemplateTable
//
/////////////////////////////////////////////////
constructor TvgrTemplateTable.Create(AEngine: TvgrReportEngine; ATemplateWorksheet: TvgrReportTemplateWorksheet);
begin
  inherited Create;
  FEngine := AEngine;
  FTemplateWorksheet := ATemplateWorksheet;
  FRow := TvgrTemplateRow.Create(Self);
  FCol := TvgrTemplateCol.Create(Self);
  FHorzStack := TvgrSectionsStack.Create;
  FVertStack := TvgrSectionsStack.Create;
  FWritedCols := TList.Create;
end;

destructor TvgrTemplateTable.Destroy;
begin
  FreeAndNil(FHorzStack);
  FreeAndNil(FVertStack);
  FreeAndNil(FRow);
  FreeAndNil(FCol);
  FreeAndNil(FWritedCols);
  inherited;
end;

procedure TvgrTemplateTable.StoreRangeForSecondPass(ABand: TvgrBand);
begin
  FStoreRangeForSecondPass := True;
  FTempBand := ABand;
end;

function TvgrTemplateTable.GetRowCount: Integer;
begin
  Result := FRow.SubVectorCount;
end;

function TvgrTemplateTable.GetColCount: Integer;
begin
  Result := FCol.SubVectorCount;
end;

function TvgrTemplateTable.GetRow(Index: Integer): TvgrTemplateRow;
begin
  Result := TvgrTemplateRow(FRow.SubVectors[Index]);
end;

function TvgrTemplateTable.GetCol(Index: Integer): TvgrTemplateCol;
begin
  Result := TvgrTemplateCol(FCol.SubVectors[Index]);
end;

procedure TvgrTemplateTable.Clear;
begin
  FRow.Clear;
  FCol.Clear;
  FWritedCols.Clear;
  FHorzStack.Clear;
  FVertStack.Clear;
end;

function TvgrTemplateTable.CalculateFormula(const AFormula: string): Variant;
begin
  Result := TemplateWorksheet.Template.Script.EvaluateExpression(AFormula, Null);
end;

function TvgrTemplateTable.ParseString(const S: string): string;
const
  cStartExpressionBracket = '[';
  cEndExpressionBracket = ']';
  cStartFormatBracket = '<';
  cEndFormatBracket = '>';
  cSpecialChar = '\';
  cNumberChars = ['0'..'9'];
var
  I, J, LenS, ALenExpression: Integer;
  AExpression, AFormat, AText: string;
  AValue: Variant;
begin
  LenS := Length(S);
  I := 1;
  while I <= LenS do
  begin
    if (S[I] = cStartExpressionBracket) and (I < LenS) then
    begin
      Inc(I);
      J := I;
      while (I <= LenS) and (S[I] <> cEndExpressionBracket) do Inc(I);
      if I > LenS then
        Result := Result + Copy(S, J, LenS)
      else
      begin
        AExpression := Copy(S, J, I - J);
        ALenExpression := Length(AExpression);
        if (ALenExpression > 1) and (AExpression[1] = cStartFormatBracket) then
        begin
          J := 2;
          while (J <= ALenExpression) and (AExpression[J] <> cEndFormatBracket) do Inc(J);
          AFormat := Copy(AExpression, 2, J - 2);
          AExpression := Copy(AExpression, J + 1, ALenExpression);
        end
        else
          AFormat := '';

        AValue := CalculateFormula(AExpression);
        if AFormat <> '' then
          AText := FormatVariant(AFormat, AValue)
        else
          AText := VarToStr(AValue);
        Result := Result + AText;
        Inc(I);
      end
    end
    else
      if (S[I] = cSpecialChar) and (I < LenS) then
      begin
        Inc(I);
        if S[I] in cNumberChars then
        begin
          J := I;
          while (I <= LenS) and (S[I] in cNumberChars) do Inc(I);
          Result := Result + Chr(StrToInt(Copy(S, J, I - J)));
        end
        else
        begin
          Result := Result + S[I];
          Inc(I);
        end;
      end
      else
      begin
        Result := Result + S[I];
        Inc(I);
      end;
  end;

//  Result := S;
end;

procedure TvgrTemplateTable.CalculateRangeValue(ASourceRange, ADestRange: IvgrRange);
begin
  if ASourceRange.Formula <> '' then
  begin
    ADestRange.StringValue := '=' + ASourceRange.Formula;
  end
  else
  begin
    if ASourceRange.ValueType = rvtString then
      ADestRange.StringValue := ParseString(ASourceRange.SimpleStringValue)
    else
      ADestRange.Value := ASourceRange.Value;
  end;
{
    case ASourceRange.ValueType of
      rvtNull:
        ADestRange.ValueData.ValueType := rvtNull;
      rvtInteger:
        begin
          ADestRange.ValueData.ValueType := rvtInteger;
          ADestRange.ValueData.vInteger := ASourceRange.ValueData.vInteger
        end;
      rvtExtended:
        begin
          ADestRange.ValueData.ValueType := rvtExtended;
          ADestRange.ValueData.vExtended := ASourceRange.ValueData.vExtended;
        end;
      rvtDateTime:
        begin
          ADestRange.ValueData.ValueType := rvtDateTime;
          ADestRange.ValueData.vDateTime := ASourceRange.ValueData.vDateTime;
        end;
      else
        ADestRange.StringValue := ParseString(ASourceRange.StringValue);
    end;
}
end;

procedure TvgrTemplateTable.WriteRangesCallback1(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ADestRange, ASourceRange: IvgrRange;
  ARect: PRect;
begin
  ARect := PRect(AData);
  ASourceRange := AItem as IvgrRange;
  with ASourceRange.Place do
    TemplateWorksheet.FCurrentRangePlace := Rect(FWorkbookCol + Left - ARect.Left,
                                                 FWorkbookRow + Top - ARect.Top,
                                                 FWorkbookCol + Right - ARect.Left,
                                                 FWorkbookRow + Bottom - ARect.Top);
  try
    with TemplateWorksheet.FCurrentRangePlace do
      ADestRange := ResultWorksheet.Ranges[Left, Top, Right, Bottom];
    FStoreRangeForSecondPass := False;
    CalculateRangeValue(ASourceRange, ADestRange);
    ADestRange.AssignStyle(ASourceRange);
    if FStoreRangeForSecondPass then
      FTempBand.AddRangesForSecondPass(ASourceRange, ADestRange);
  finally
    TemplateWorksheet.FCurrentRangePlace := Rect(-1, -1, -1, -1);
  end;
end;

procedure TvgrTemplateTable.WriteRangesCallback2(AItem : IvgrWBListItem; AItemIndex: Integer; AData: Pointer);
var
  ASourceBorder: IvgrBorder;
  ARect: PRect;
begin
  ASourceBorder := AItem as IvgrBorder;
  ARect := PRect(AData);
  with ASourceBorder do
  begin
    if ((Left <= ARect.Right) or (Orientation = vgrboLeft)) and
       ((Top <= ARect.Bottom) or (Orientation = vgrboTop)) then
      ResultWorksheet.Borders[FWorkbookCol + Left - ARect.Left,
                              FWorkbookRow + Top - ARect.Top,
                              Orientation].Assign(ASourceBorder);
  end;
end;

procedure TvgrTemplateTable.WriteRanges(ALeft, ATop, AWidth, AHeight: Integer);
var
  AWriteRangesRect: TRect;
begin
  AWriteRangesRect := Bounds(ALeft, ATop, AWidth, AHeight);

  TemplateWorksheet.RangesList.FindAndCallBack(AWriteRangesRect, WriteRangesCallback1, @AWriteRangesRect);
  with AWriteRangesRect do
    TemplateWorksheet.BordersList.FindAndCallBack(Rect(Left, Top, Right + 1, Bottom + 1), WriteRangesCallback2, @AWriteRangesRect);

  FWorkbookCol := FWorkbookCol + AWidth + 1;
end;

procedure TvgrTemplateTable.WriteRows(ATop, AHeight: Integer; AAutoSize: Boolean);
var
  I, ASize: Integer;
  ASourceRow: IvgrRow;
begin
  for I := 0 to AHeight do
  begin
    ASourceRow := TemplateWorksheet.RowsList.Find(ATop + I);
    if ASourceRow <> nil then
      ResultWorksheet.Rows[FWorkbookRow + I].Assign(ASourceRow);
    if AAutoSize then
    begin
      ASize := ResultWorksheet.GetRowAutoHeight(FWorkbookRow + I);
      if ASize <> 0 then
        ResultWorksheet.Rows[FWorkbookRow + I].Height := ASize;
    end;
  end;
end;

procedure TvgrTemplateTable.WriteCols(ALeft, AWidth: Integer);
var
  I, ADestColNumber: Integer;
  ASourceCol: IvgrCol;
begin
  for I := 0 to AWidth do
  begin
    ADestColNumber := FWorkbookCol + I;
    if FWritedCols.IndexOf(Pointer(ADestColNumber)) = -1 then
    begin
      FWritedCols.Add(Pointer(ADestColNumber));
      ASourceCol := TemplateWorksheet.ColsList.Find(ALeft + I);
      if ASourceCol <> nil then
        ResultWorksheet.Cols[ADestColNumber].Assign(ASourceCol);
    end;
  end;
end;

procedure TvgrTemplateTable.WriteLn(ALineSize: Integer);
begin
  FWorkbookRow := FWorkbookRow + ALineSize;
  FWorkbookCol := 0;
end;

procedure TvgrTemplateTable.PrepareVector(AVector: TvgrTemplateVector; ABands: TvgrBands; AMaxEndPos: Integer);

  procedure PrepareRange(AStartPos, AEndPos: Integer; AParentBand: TvgrBand; AVector: TvgrTemplateVector);
  var
    I, ACurPos: Integer;
    ABand: TvgrBand;
    ASubVector: TvgrTemplateVector;
  begin
    ACurPos := AStartPos;
    for I := 0 to ABands.FSortedList.Count - 1 do
    begin
      ABand := TvgrBand(ABands.FSortedList[I]);
      if (ABand.StartPos >= AStartPos) and (ABand.EndPos <= AEndPos) then
      begin
        if ABand.ParentBand = AParentBand then
        begin
          if ABand.StartPos > ACurPos then
            AVector.AddSubVector(ACurPos, ABand.StartPos - ACurPos - 1, nil, AVector);
          ASubVector := AVector.AddSubVector(ABand.StartPos, ABand.EndPos - ABand.StartPos, ABand, AVector);
          ABand.FTemplateVector := ASubVector;
          PrepareRange(ABand.StartPos, ABand.EndPos, ABand, ASubVector);
          ACurPos := ABand.EndPos + 1;
        end;
      end;
    end;
    if ACurPos <= AEndPos then
      AVector.AddSubVector(ACurPos, AEndPos - ACurPos, nil, AVector);
  end;

begin
//  if ABands.Count > 0 then
  PrepareRange(0, AMaxEndPos, nil, AVector);
end;

procedure TvgrTemplateTable.Prepare;
begin
  with TemplateWorksheet.Dimensions do
  begin
    PrepareVector(FRow, TemplateWorksheet.HorzBands, Bottom);
    PrepareVector(FCol, TemplateWorksheet.VertBands, Right);
  end;
end;

procedure TvgrTemplateTable.Generate;
var
  I: Integer;
  ARow: TvgrTemplateRow;
  ADone: Boolean;
begin
  if (not TemplateWorksheet.SkipGenerate) and
     (TemplateWorksheet.Dimensions.Right >= 0) and
     (TemplateWorksheet.Dimensions.Bottom >= 0) then
  begin
    TemplateWorksheet.DoBeforeGenerate;
    
    ResultWorksheet := Engine.Workbook.AddWorksheet;
    ResultWorksheet.Title := TemplateWorksheet.Title;
    ResultWorksheet.PageProperties.Assign(TemplateWorksheet.PageProperties);
    TemplateWorksheet.DoAfterAddWorksheet(ResultWorksheet);

    FWorkbookRow := 0;
    TemplateWorksheet.DoCustomGenerate(ResultWorksheet, ADone);
    if not ADone then
    begin
      for I := 0 to RowCount - 1 do
      begin
        ARow := Rows[I];
        if ARow.Band = nil then
          ARow.Write
        else
          ARow.Band.Generate(ARow);
      end;
    end;

    TemplateWorksheet.DoAfterGenerate(ResultWorksheet);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrReportEngine
//
/////////////////////////////////////////////////
constructor TvgrReportEngine.Create(AOwner: TComponent);
begin
  inherited;
  FTables := TList.Create;
end;

destructor TvgrReportEngine.Destroy;
begin
  ClearTables;
  FreeAndNil(FTables);
  inherited;
end;

procedure TvgrReportEngine.SetTemplate(Value: TvgrReportTemplate);
begin
  if FTemplate <> Value then
  begin
    FTemplate := Value;
  end;
end;

procedure TvgrReportEngine.SetWorkbook(Value: TvgrWorkbook);
begin
  if FWorkbook <> Value then
  begin
    FWorkbook := Value;
  end;
end;

function TvgrReportEngine.GetTableCount: Integer;
begin
  Result := FTables.Count;
end;

function TvgrReportEngine.GetTable(Index: Integer): TvgrTemplateTable;
begin
  Result := TvgrTemplateTable(FTables[Index]);
end;

procedure TvgrReportEngine.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if FTemplate = AComponent then
      FTemplate := nil;
    if FWorkbook = AComponent then
      FWorkbook := nil;
  end;
end;

procedure TvgrReportEngine.ClearTables;
var
  I: Integer;
begin
  for I := 0 to TableCount - 1 do
    Tables[I].Free;
  FTables.Clear;
end;

procedure TvgrReportEngine.Generate;
var
  I: Integer;
  ATable: TvgrTemplateTable;
begin
  Workbook.BeginUpdate;
  FTemplate.BeginGenerate;

  try
    // 1. Prepare tables (calculate Horz and Vert vectors)
    for I := 0 to Template.WorksheetsCount - 1 do
    begin
      ATable := TvgrTemplateTable.Create(Self, Template.Worksheets[I]);
      FTables.Add(ATable);
      ATable.Prepare;
      Template.Worksheets[I].FTemplateTable := ATable;
    end;

    // 2. Clear Workbook
    Workbook.Clear;

    // 3. Main generate cycle
    for I := 0 to TableCount - 1 do
      Tables[I].Generate;
  finally
    FTemplate.EndGenerate;
    ClearTables;
    Workbook.EndUpdate;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrRegisteredBandClasses
//
/////////////////////////////////////////////////
constructor TvgrRegisteredBandClasses.Create;
begin
  inherited;
  FList := TList.Create;
  FBitmapList := TList.Create;
end;

destructor TvgrRegisteredBandClasses.Destroy;
begin
  FreeList(FBitmapList);
  FreeAndNil(FList);
  inherited;
end;

function TvgrRegisteredBandClasses.GetItem(Index: Integer): TvgrBandClass;
begin
  Result := TvgrBandClass(FList[Index]);
end;

function TvgrRegisteredBandClasses.GetBitmap(Index: Integer): TBitmap;
begin
  Result := TBitmap(FBitmapList[Index]);
end;

function TvgrRegisteredBandClasses.GetCount: Integer;
begin
  Result := FList.Count;
end;

procedure TvgrRegisteredBandClasses.RegisterBandClass(ABandClass: TvgrBandClass);
var
  ABitmap: TBitmap;
begin
  if FList.IndexOf(ABandClass) = -1 then
  begin
    FList.Add(ABandClass);
    ABitmap := TBitmap.Create;
    ABitmap.LoadFromResourceName(hInstance, ABandClass.GetBitmapResName);
    FBitmapList.Add(ABitmap);
  end;
end;

function TvgrRegisteredBandClasses.GetBandClassIndex(ABand: TvgrBand): Integer;
begin
  Result := 0;
  while (Result < Count) and not(ABand is Items[Result]) do Inc(Result);
  if Result >= Count then
    Result := -1;
end;

initialization

  BandClasses := TvgrRegisteredBandClasses.Create;

  Classes.RegisterClass(TvgrBand);
  Classes.RegisterClass(TvgrDetailBand);
  Classes.RegisterClass(TvgrGroupBand);
  Classes.RegisterClass(TvgrDataBand);
  Classes.RegisterClass(TvgrReportTemplateWorksheet);

  RegisterBandClass(TvgrBand);
  RegisterBandClass(TvgrDetailBand);
  RegisterBandClass(TvgrGroupBand);
  RegisterBandClass(TvgrDataBand);

finalization

  FreeAndNil(BandClasses);

end.