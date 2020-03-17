{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

unit WriteXXExamples_Form;

{$I vtk.inc}

interface

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, vgr_CommonClasses, vgr_DataStorage, vgr_Report, StdCtrls;

type
  TWriteXXExamplesForm = class(TForm)
    SimpleWrite: TvgrReportTemplate;
    SimpleWriteXXvgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    r1_Header: TvgrBand;
    r1_Data: TvgrBand;
    r1_Footer: TvgrBand;
    MasterDetailWrite: TvgrReportTemplate;
    vgrReportTemplate1vgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    r2_Header: TvgrBand;
    r2_CustomerHeader: TvgrBand;
    r2_OrdersDetail: TvgrDetailBand;
    r2_OrdersData: TvgrDataBand;
    CrossTabWrite: TvgrReportTemplate;
    CrossTabWriteSheet: TvgrReportTemplateWorksheet;
    r3_HorHeader: TvgrBand;
    r3_HorData: TvgrBand;
    r3_VerHeader: TvgrBand;
    r3_VerData: TvgrBand;
    r3_Header: TvgrBand;
    r3_VerFooter: TvgrBand;
    r3_HorFooter: TvgrBand;
    SimpleWriteScript: TvgrReportTemplate;
    vgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    r4_Header: TvgrBand;
    r4_Data: TvgrBand;
    r4_Footer: TvgrBand;
    MasterDetailWriteScript: TvgrReportTemplate;
    vgrReportTemplateWorksheet2: TvgrReportTemplateWorksheet;
    r5_Header: TvgrBand;
    r5_CustomerHeader: TvgrBand;
    r5_OrdersDetail: TvgrDetailBand;
    r5_OrdersData: TvgrDataBand;
    CrossTabWriteScript: TvgrReportTemplate;
    CrossTabWriteScriptSheet: TvgrReportTemplateWorksheet;
    r6_HorHeader: TvgrBand;
    r6_HorData: TvgrBand;
    r6_Header: TvgrBand;
    r6_HorFooter: TvgrBand;
    r6_VerHeader: TvgrBand;
    r6_VerData: TvgrBand;
    r6_VerFooter: TvgrBand;
    procedure SimpleWriteCustomGenerate(Sender: TObject;
      ATemplateWorksheet: TvgrReportTemplateWorksheet;
      AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure MasterDetailWriteCustomGenerate(Sender: TObject;
      ATemplateWorksheet: TvgrReportTemplateWorksheet;
      AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
    procedure CrossTabWriteCustomGenerate(Sender: TObject;
      ATemplateWorksheet: TvgrReportTemplateWorksheet;
      AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure SimpleWriteProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
    procedure SimpleWriteScriptProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
    procedure MasterDetailWriteProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
    procedure MasterDetailWriteScriptProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
    procedure CrossTabWriteProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
    procedure CrossTabWriteScriptProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
  end;

var
  WriteXXExamplesForm: TWriteXXExamplesForm;

implementation

uses Global_DM, Main_Form;

{$R *.dfm}

procedure TWriteXXExamplesForm.SimpleWriteProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
begin
  ATemplate := SimpleWrite;
  ATemplateFileName := '';
end;

procedure TWriteXXExamplesForm.MasterDetailWriteProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
begin
  ATemplate := MasterDetailWrite;
  ATemplateFileName := '';
end;

procedure TWriteXXExamplesForm.CrossTabWriteProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
begin
  ATemplate := CrossTabWrite;
  ATemplateFileName := '';
end;

procedure TWriteXXExamplesForm.SimpleWriteScriptProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
begin
  ATemplate := SimpleWriteScript;
  ATemplateFileName := '';
end;

procedure TWriteXXExamplesForm.MasterDetailWriteScriptProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
begin
  ATemplate := MasterDetailWriteScript;
  ATemplateFileName := '';
end;

procedure TWriteXXExamplesForm.CrossTabWriteScriptProc(var ATemplate: TvgrReportTemplate; var ATemplateFileName: string);
begin
  ATemplate := CrossTabWriteScript;
  ATemplateFileName := '';
end;

procedure TWriteXXExamplesForm.FormCreate(Sender: TObject);
begin
  MainForm.RegisterReport('Using of WriteXX functions', 'Simple report', nil, SimpleWriteProc);
  MainForm.RegisterReport('Using of WriteXX functions', 'Master-Detail report', nil, MasterDetailWriteProc);
  MainForm.RegisterReport('Using of WriteXX functions', 'CrossTab report', nil, CrossTabWriteProc);

  MainForm.RegisterReport('Using of WriteXX functions', 'Simple report (script)', nil, SimpleWriteScriptProc);
  MainForm.RegisterReport('Using of WriteXX functions', 'Master-Detail report (script)', nil, MasterDetailWriteScriptProc);
  MainForm.RegisterReport('Using of WriteXX functions', 'CrossTab report (script)', nil, CrossTabWriteScriptProc);
end;

procedure TWriteXXExamplesForm.SimpleWriteCustomGenerate(Sender: TObject;
  ATemplateWorksheet: TvgrReportTemplateWorksheet;
  AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
var
  ADataStart: Integer;
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

  // place formula which counts number of the processed records
  // WorkbookRow - returns the number of current row, first row has index 0
  AResultWorksheet.Ranges[3,
                          ATemplateWorksheet.WorkbookRow - 1,
                          4,
                          ATemplateWorksheet.WorkbookRow - 1].Formula :=
    Format('Count(A$%d:A%d)',
           [ADataStart + 1, ATemplateWorksheet.WorkbookRow - 1]);

  // cancel default processing
  ADone := True;
end;

procedure TWriteXXExamplesForm.MasterDetailWriteCustomGenerate(
  Sender: TObject; ATemplateWorksheet: TvgrReportTemplateWorksheet;
  AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
begin
  // write report header
  ATemplateWorksheet.WriteLnBand(r2_Header);

  // open the dataset and go to first record
  GlobalDM.Customers.Open;
  GlobalDM.Customers.First;

  while not GlobalDM.Customers.Eof do
  begin
    // write row of data
    ATemplateWorksheet.WriteLnBand(r2_CustomerHeader);

    // prepare Orders dataset
    GlobalDM.Orders.Filter := Format('[CustNo] = %d', [GlobalDm.Customers.FieldByName('CustNo').AsInteger]);
    GlobalDM.Orders.Filtered := True;
    GlobalDM.Orders.Open;

    // write Orders detail
    ATemplateWorksheet.WriteLnBand(r2_OrdersDetail);

    // go to next record
    GlobalDM.Customers.Next;
  end;

  // cancel default processing
  ADone := True;
end;

procedure TWriteXXExamplesForm.CrossTabWriteCustomGenerate(Sender: TObject;
  ATemplateWorksheet: TvgrReportTemplateWorksheet;
  AResultWorksheet: TvgrWorksheet; var ADone: Boolean);
begin
  // write report header
  ATemplateWorksheet.WriteLnBand(r3_Header);

  // write horizontal header
  ATemplateWorksheet.WriteCell(r3_HorHeader, r3_VerHeader);
  GlobalDm.KS.First;
  while not GlobalDm.KS.Eof do
  begin
    ATemplateWorksheet.WriteCell(r3_HorHeader, r3_VerData);
    GlobalDm.KS.Next;
  end;
  ATemplateWorksheet.WriteCell(r3_HorHeader, r3_VerFooter);
  ATemplateWorksheet.WriteLn;

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

  // write horizontal footer
  ATemplateWorksheet.WriteCell(r3_HorFooter, r3_VerHeader);
  GlobalDm.KS.First;
  while not GlobalDm.KS.Eof do
  begin
    ATemplateWorksheet.WriteCell(r3_HorFooter, r3_VerData);
    GlobalDm.KS.Next;
  end;
  ATemplateWorksheet.WriteCell(r3_HorFooter, r3_VerFooter);
  ATemplateWorksheet.WriteLn;

  // cancel default processing
  ADone := True;
end;

end.
 