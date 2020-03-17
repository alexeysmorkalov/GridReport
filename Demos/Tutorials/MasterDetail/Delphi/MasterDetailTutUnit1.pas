unit MasterDetailTutUnit1;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7}Variants, {$ENDIF}Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, DBTables, vgr_WorkbookDesigner, vgr_Report,
  vgr_CommonClasses, vgr_DataStorage;

type
  TForm1 = class(TForm)
    vgrReportTemplate1: TvgrReportTemplate;
    vgrWorkbook1: TvgrWorkbook;
    vgrReportEngine1: TvgrReportEngine;
    vgrWorkbookDesigner1: TvgrWorkbookDesigner;
    Table1: TTable;
    Table2: TTable;
    Button1: TButton;
    vgrReportTemplate1vgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    vgrDetailBand1: TvgrDetailBand;
    vgrDataBand1: TvgrDataBand;
    vgrDetailBand2: TvgrDetailBand;
    vgrBand1: TvgrBand;
    vgrDataBand2: TvgrDataBand;
    vgrBand2: TvgrBand;
    procedure Button1Click(Sender: TObject);
    procedure Table1AfterScroll(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  vgrReportEngine1.Generate;
  vgrWorkbookDesigner1.Design(True);
end;

procedure TForm1.Table1AfterScroll(DataSet: TDataSet);
begin
  Table2.Filter := Format('[CustNo] = %d', [Table1.FieldByName('CustNo').AsInteger]);
  Table2.Filtered := True;
end;

end.
