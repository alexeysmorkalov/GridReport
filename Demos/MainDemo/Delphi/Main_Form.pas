{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

unit Main_Form;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, Menus, ComCtrls, ExtCtrls, StdCtrls, vgr_WorkbookGrid,
  vgr_CommonClasses, vgr_DataStorage, vgr_ScriptControl, vgr_Report,
  vgr_AliasManager, vgr_WorkbookDesigner, vgr_ReportDesigner, vgr_Dialogs,
  vgr_WorkbookPreviewForm, vgr_Writers, ImgList, ToolWin, shellapi,
  vgr_AliasManagerDesigner, vgr_WBScriptConstants;

type
  TReportGenerateProc = procedure(ADestWorkbook: TvgrWorkbook) of object;
  TReportPrepareTemplateProc = procedure(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string) of object;

  /////////////////////////////////////////////////
  //
  // TvgrRegisteredReport
  //
  /////////////////////////////////////////////////
  TvgrRegisteredReport = class
  private
    FReportCategory: string;
    FReportCaption: string;
    FGenerateProc: TReportGenerateProc;
    FPrepareTemplateProc: TReportPrepareTemplateProc;
  public
    constructor Create(const AReportCategory, AReportCaption: string;
                       AGenerateProc: TReportGenerateProc;
                       APrepareTemplateProc: TReportPrepareTemplateProc);
    property ReportCategory: string read FReportCategory;
    property ReportCaption: string read FReportCaption;
    property GenerateProc: TReportGenerateProc read FGenerateProc;
    property PrepareTemplateProc: TReportPrepareTemplateProc read FPrepareTemplateProc;
  end;

  /////////////////////////////////////////////////
  //
  // TMainForm
  //
  /////////////////////////////////////////////////
  TMainForm = class(TForm)
    MainMenu: TMainMenu;
    ActionList: TActionList;
    File1: TMenuItem;
    aExit: TAction;
    Exit1: TMenuItem;
    TVReports: TTreeView;
    Splitter: TSplitter;
    Panel1: TPanel;
    Label1: TLabel;
    PageControl: TPageControl;
    PTemplate: TTabSheet;
    PResults: TTabSheet;
    TemplateGrid: TvgrWorkbookGrid;
    Label2: TLabel;
    Panel2: TPanel;
    Label3: TLabel;
    reTemplateDesc: TRichEdit;
    Splitter1: TSplitter;
    WorkbookGrid: TvgrWorkbookGrid;
    Workbook: TvgrWorkbook;
    ReportTemplate: TvgrReportTemplate;
    aEditTemplateInReportDesigner: TAction;
    aEditWorkbookInWorkbookDesigner: TAction;
    aGenerate: TAction;
    Edit1: TMenuItem;
    EdittemplateinTvgrReportDesigner1: TMenuItem;
    N1: TMenuItem;
    Generate1: TMenuItem;
    N2: TMenuItem;
    EditworkbookinTvgrWorkbookDesigner1: TMenuItem;
    aAutoGenerate: TAction;
    Options1: TMenuItem;
    Autogenerate1: TMenuItem;
    aPreviewReportTemplate: TAction;
    aPreviewWorkbook: TAction;
    Preview1: TMenuItem;
    Previewreporttemplate1: TMenuItem;
    Previewworkbook1: TMenuItem;
    PCategory: TTabSheet;
    Label4: TLabel;
    reCategoryDesc: TRichEdit;
    ReportEngine: TvgrReportEngine;
    ReportDesigner: TvgrReportDesigner;
    WorkbookDesigner: TvgrWorkbookDesigner;
    WorkbookPreviewer: TvgrWorkbookPreviewer;
    ImageList: TImageList;
    AliasManager: TvgrAliasManager;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    Export1: TMenuItem;
    aExportToExcel: TAction;
    aExportToHTML: TAction;
    ExporttoExcel1: TMenuItem;
    ExporttoHTML1: TMenuItem;
    SaveExportDialog: TSaveDialog;
    aEditAliases: TAction;
    Aliases1: TMenuItem;
    ReportTemplatevgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    procedure TVReportsChange(Sender: TObject; Node: TTreeNode);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure aEditTemplateInReportDesignerUpdate(Sender: TObject);
    procedure aEditTemplateInReportDesignerExecute(Sender: TObject);
    procedure PageControlChanging(Sender: TObject;
      var AllowChange: Boolean);
    procedure aEditWorkbookInWorkbookDesignerExecute(Sender: TObject);
    procedure aPreviewWorkbookExecute(Sender: TObject);
    procedure aExitExecute(Sender: TObject);
    procedure aGenerateExecute(Sender: TObject);
    procedure aPreviewReportTemplateExecute(Sender: TObject);
    procedure aAutoGenerateExecute(Sender: TObject);
    procedure TVReportsDblClick(Sender: TObject);
    procedure aExportToExcelExecute(Sender: TObject);
    procedure aExportToHTMLExecute(Sender: TObject);
    procedure aEditAliasesExecute(Sender: TObject);
  private
    { Private declarations }
    FReports: TList;
    FWorkbookGenerated: Boolean;
    procedure UpdateTVReports;
    function GetReport(Index: Integer): TvgrRegisteredReport;
    function GetReportCount: Integer;
    function GetSelectedReport: TvgrRegisteredReport;
    procedure SetWorkbookGenerated(Value: Boolean);
    function GetTemplateFile(const ATemplateName: string): string;
    function GetDescFile(const AObjectName, AExt: string): string;
    function GetTemplate(AReport: TvgrRegisteredReport): TvgrReportTemplate;
    procedure GetDesc(ARichEdit: TRichEdit; const AObjectName: string);
    procedure CheckWorkbookGenerated(AAlwaysGenerate: Boolean);
    procedure OnReportTemplateChanged(Sender: TObject);
  public
    { Public declarations }
    procedure RegisterReport(const AReportCategory, AReportCaption: string;
                             AGenerateProc: TReportGenerateProc;
                             APrepareTemplateProc: TReportPrepareTemplateProc);
    property Reports[Index: Integer]: TvgrRegisteredReport read GetReport;
    property ReportCount: Integer read GetReportCount;
    property SelectedReport: TvgrRegisteredReport read GetSelectedReport;
    property WorkbookGenerated: Boolean read FWorkbookGenerated write SetWorkbookGenerated;
  end;

var
  MainForm: TMainForm;

implementation

uses WriteXXExamples_Form;

{$R *.dfm}

/////////////////////////////////////////////////
//
// TvgrRegisteredReport
//
/////////////////////////////////////////////////
constructor TvgrRegisteredReport.Create(const AReportCategory, AReportCaption: string;
                   AGenerateProc: TReportGenerateProc;
                   APrepareTemplateProc: TReportPrepareTemplateProc);
begin
  inherited Create;
  FReportCategory := AReportCategory;
  FReportCaption := AReportCaption;
  FGenerateProc := AGenerateProc;
  FPrepareTemplateProc := APrepareTemplateProc;
end;

/////////////////////////////////////////////////
//
// TMainForm
//
/////////////////////////////////////////////////
procedure TMainForm.RegisterReport(const AReportCategory, AReportCaption: string;
                                   AGenerateProc: TReportGenerateProc;
                                   APrepareTemplateProc: TReportPrepareTemplateProc);
begin
  FReports.Add(TvgrRegisteredReport.Create(AReportCategory, AReportCaption, AGenerateProc, APrepareTemplateProc));
  UpdateTVReports;
  TVReportsChange(Self, nil);
end;

function TMainForm.GetReport(Index: Integer): TvgrRegisteredReport;
begin
  Result := TvgrRegisteredReport(FReports[Index]);
end;

function TMainForm.GetSelectedReport: TvgrRegisteredReport;
begin
  if (TVReports.Selected = nil) or (TVReports.Selected.Data = nil) then
    Result := nil
  else
    Result := TvgrRegisteredReport(TVReports.Selected.Data);
end;

procedure TMainForm.SetWorkbookGenerated(Value: Boolean);
begin
  if FWorkbookGenerated <> Value then
  begin
    if Value and (SelectedReport <> nil) then
    begin
      ReportEngine.Template := TemplateGrid.Workbook as TvgrReportTemplate;
      ReportEngine.Generate;
    end
    else
    begin
      Workbook.Clear;
    end;
    FWorkbookGenerated := Value;
  end;
end;

function TMainForm.GetReportCount: Integer;
begin
  Result := FReports.Count;
end;

procedure TMainForm.UpdateTVReports;
var
  I: Integer;
  ACategoryNode: TTreeNode;

  function FindOrAddCategory(const ACategory: string): TTreeNode;
  var
    I: Integer;
  begin
    I := 0;
    while (I < TVReports.Items.Count) and
          not ((TVReports.Items[I].Text = ACategory) and (TVReports.Items[I].Data = nil)) do Inc(I);
    if I = TVReports.Items.Count then
      Result := TVReports.Items.AddObject(nil, ACategory, nil)
    else
      Result := TVReports.Items[I];
  end;
  
begin
  TVReports.Items.BeginUpdate;
  try
    TVReports.Items.Clear;
    for I := 0 to ReportCount - 1 do
    begin
      ACategoryNode := FindOrAddCategory(Reports[I].ReportCategory);
      TVReports.Items.AddChildObject(ACategoryNode, Reports[I].ReportCaption, Reports[I]);
    end;
    TVReports.FullExpand;
  finally
    TVReports.Items.EndUpdate;
  end;
end;

function TMainForm.GetTemplateFile(const ATemplateName: string): string;
begin
  Result := ExtractFilePath(ParamStr(0)) + 'Templates\' + ATemplateName + '.grt';
end;

function TMainForm.GetDescFile(const AObjectName, AExt: string): string;
begin
  Result := ExtractFilePath(ParamStr(0)) + 'Descs\' + AObjectName + '.' + AExt;
end;

function TMainForm.GetTemplate(AReport: TvgrRegisteredReport): TvgrReportTemplate;
var
  ATemplateName: string;
begin
  Result := nil;
  if Assigned(AReport.PrepareTemplateProc) then
    AReport.PrepareTemplateProc(Result, ATemplateName)
  else
    ATemplateName := AReport.ReportCaption;
  if Result = nil then
  begin
    Result := ReportTemplate;
    Result.LoadFromFile(GetTemplateFile(ATemplateName));
  end;
end;

procedure TMainForm.GetDesc(ARichEdit: TRichEdit; const AObjectName: string);
begin
  if FileExists(GetDescFile(AObjectName, 'rtf')) then
    ARichEdit.Lines.LoadFromFile(GetDescFile(AObjectName, 'rtf'))
  else
    if FileExists(GetDescFile(AObjectName, 'txt')) then
      ARichEdit.Lines.LoadFromFile(GetDescFile(AObjectName, 'txt'))
    else
      ARichEdit.Lines.Text := 'No description available';
end;

procedure TMainForm.TVReportsChange(Sender: TObject; Node: TTreeNode);
var
  q: TvgrReportTemplate;
begin
  if SelectedReport <> nil then
  begin
    Label1.Caption := Format(' Report: %s \ %s', [SelectedReport.ReportCategory, SelectedReport.ReportCaption]);
    PageControl.Visible := True;
    PCategory.TabVisible := False;
    PTemplate.TabVisible := True;
    PResults.TabVisible := True;

    //
    WorkbookGenerated := False;
    // Prepare template
    q := GetTemplate(SelectedReport);
    TemplateGrid.Workbook := q;
    TemplateGrid.Workbook.OnChanged := OnReportTemplateChanged;
    // report description
    GetDesc(reTemplateDesc, SelectedReport.ReportCaption);
  end
  else
  begin
    if TemplateGrid.Workbook <> nil then
      TemplateGrid.Workbook.OnChanged := nil;
    TemplateGrid.Workbook := nil;
    if Node <> nil then
    begin
      Label1.Caption := Format(' Category: %s', [Node.Text]);
      PageControl.Visible := True;
      PCategory.TabVisible := True;
      PTemplate.TabVisible := False;
      PResults.TabVisible := False;
      GetDesc(reCategoryDesc, Node.Text);
    end
    else
    begin
      Label1.Caption := 'Select element in tree';
      PageControl.Visible := False;
    end;
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  FReports := TList.Create;

  RegisterReport('Simple reports', 'Simple customers list', nil, nil);
  RegisterReport('Simple reports', 'Simple events list', nil, nil);  
  RegisterReport('Simple reports', 'Simple biolife list with memo', nil, nil);
  RegisterReport('Simple reports', 'Simple multisheet workbook', nil, nil);
  RegisterReport('Groups', 'Grouped customers list', nil, nil);
  RegisterReport('Groups', 'Grouped biolife list', nil, nil);
  RegisterReport('CrossTab', 'Simple CrossTab', nil, nil);
  RegisterReport('CrossTab', 'Grouped CrossTab', nil, nil);  
  RegisterReport('Master-detail', 'Simple master-detail', nil, nil);
  RegisterReport('Aliases', 'Alias manager', nil, nil);
  RegisterReport('Script', 'Created by script executing', nil, nil);
  RegisterReport('Script', 'Highlighting rows by criterion', nil, nil);
  RegisterReport('Script', 'Changing report template from script', nil, nil);
  RegisterReport('Script', 'Defining events in script', nil, nil);
  RegisterReport('Script', 'Loading text files in script', nil, nil);
  RegisterReport('Script', 'InputBox example', nil, nil);
  RegisterReport('Formulas', 'All formulas', nil, nil);

  UpdateTVReports;
  TVReportsChange(Self, nil);
end;

procedure TMainForm.FormDestroy(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to ReportCount - 1 do
    Reports[I].Free;
  FReports.Free;
end;

procedure TMainForm.aEditTemplateInReportDesignerUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := TemplateGrid.Workbook <> nil;
end;

procedure TMainForm.aEditTemplateInReportDesignerExecute(Sender: TObject);
var
  ATemplate: TvgrReportTemplate;
  ATemplateName: string;
begin
  ReportDesigner.Template := TemplateGrid.Workbook as TvgrReportTemplate;

  if SelectedReport <> nil then
  begin
    if Assigned(SelectedReport.PrepareTemplateProc) then
    begin
      SelectedReport.PrepareTemplateProc(ATemplate, ATemplateName);
      if ATemplate = nil then
        ATemplateName := GetTemplateFile(ATemplateName)
      else
        ATemplateName := '';
    end
    else
      ATemplateName := GetTemplateFile(SelectedReport.ReportCaption);
    ReportDesigner.DocumentName := ATemplateName;
  end;

  ReportDesigner.Design(True);
end;

procedure TMainForm.CheckWorkbookGenerated(AAlwaysGenerate: Boolean);
begin
  if not WorkbookGenerated then
  begin
    if AAlwaysGenerate then
      WorkbookGenerated := True
    else
      MBox('Please generate workbook first', MB_OK or MB_ICONEXCLAMATION);
  end;
end;

procedure TMainForm.PageControlChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
  if PageControl.ActivePage = PTemplate then
  begin
    CheckWorkbookGenerated(aAutoGenerate.Checked);
    AllowChange := WorkbookGenerated;
  end;
end;

procedure TMainForm.aEditWorkbookInWorkbookDesignerExecute(
  Sender: TObject);
begin
  CheckWorkbookGenerated(aAutoGenerate.Checked);
  if WorkbookGenerated then
    WorkbookDesigner.Design(True);
end;

procedure TMainForm.aPreviewWorkbookExecute(Sender: TObject);
begin
  CheckWorkbookGenerated(aAutoGenerate.Checked);
  try
    if WorkbookGenerated then
    begin
      WorkbookPreviewer.Workbook := Workbook;
      WorkbookPreviewer.Preview(True);
    end;
  finally
    WorkbookPreviewer.Workbook := nil;
  end;
end;

procedure TMainForm.aExitExecute(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.aGenerateExecute(Sender: TObject);
begin
  CheckWorkbookGenerated(True);
end;

procedure TMainForm.aPreviewReportTemplateExecute(Sender: TObject);
begin
  try
    WorkbookPreviewer.Workbook := TemplateGrid.Workbook;
    WorkbookPreviewer.Preview(True);
  finally
    WorkbookPreviewer.Workbook := nil;
  end;
end;

procedure TMainForm.aAutoGenerateExecute(Sender: TObject);
begin
  TAction(Sender).Checked := not TAction(Sender).Checked;
end;

procedure TMainForm.OnReportTemplateChanged(Sender: TObject);
begin
  WorkbookGenerated := False;
end;

procedure TMainForm.TVReportsDblClick(Sender: TObject);
begin
  aEditTemplateInReportDesigner.Execute;
end;

procedure TMainForm.aExportToExcelExecute(Sender: TObject);
var
  AWriter: TvteExcelWriter;
begin
  CheckWorkbookGenerated(aAutoGenerate.Checked);
  if WorkbookGenerated then
  begin
    SaveExportDialog.FileName := ExtractFilePath(ParamStr(0)) + 'Export.xls';
    SaveExportDialog.DefaultExt := 'xls';
    SaveExportDialog.Filter := 'Excel files (*.xls)|*.xls|All files (*.*)|*.*';
    if SaveExportDialog.Execute then
    begin
      AWriter := TvteExcelWriter.Create;
      try
        AWriter.Save(Workbook, SaveExportDialog.FileName);
      finally
        AWriter.Free;
      end;
      ShellExecute(0, 'open', PChar(SaveExportDialog.FileName), '', '', SW_NORMAL);
    end;
  end;
end;

procedure TMainForm.aExportToHTMLExecute(Sender: TObject);
var
  AWriter: TvteHTMLWriter;
begin
  CheckWorkbookGenerated(aAutoGenerate.Checked);
  if WorkbookGenerated then
  begin
    SaveExportDialog.FileName := ExtractFilePath(ParamStr(0)) + 'Export.htm';
    SaveExportDialog.DefaultExt := 'htm';
    SaveExportDialog.Filter := 'HTML files (*.htm)|*.htm|All files (*.*)|*.*';
    if SaveExportDialog.Execute then
    begin
      AWriter := TvteHTMLWriter.Create;
      try
        AWriter.Save(Workbook, SaveExportDialog.FileName);
      finally
        AWriter.Free;
      end;
      ShellExecute(0, 'open', PChar(SaveExportDialog.FileName), '', '', SW_NORMAL);
    end;
  end;
end;

procedure TMainForm.aEditAliasesExecute(Sender: TObject);
var
  ADesigner: TvgrAliasManagerDesignerDialog;
begin
  ADesigner := TvgrAliasManagerDesignerDialog.Create(nil);
  try
    ADesigner.AliasManager := AliasManager;
    ADesigner.AvailableComponentsProvider.Root := ReportTemplate;
    ADesigner.Execute;
  finally
    ADesigner.Free;
  end;
end;

end.
