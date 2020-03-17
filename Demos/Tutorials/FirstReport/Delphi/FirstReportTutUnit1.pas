unit FirstReportTutUnit1;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}Classes, Graphics, Controls, Forms,
  Dialogs, vgr_Report, vgr_DataStorage, StdCtrls, DB, DBTables,
  vgr_WorkbookDesigner, vgr_CommonClasses;

type
  TForm1 = class(TForm)
    vgrReportTemplate1: TvgrReportTemplate;
    vgrWorkbook1: TvgrWorkbook;
    vgrReportEngine1: TvgrReportEngine;
    vgrWorkbookDesigner1: TvgrWorkbookDesigner;
    Table1: TTable;
    Button1: TButton;
    vgrReportTemplate1vgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    vgrDetailBand1: TvgrDetailBand;
    vgrDataBand1: TvgrDataBand;
    procedure Button1Click(Sender: TObject);
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
  //Parameter True of Design method shows that the designer's form will be modal
  vgrWorkbookDesigner1.Design(True);
end;
end.
