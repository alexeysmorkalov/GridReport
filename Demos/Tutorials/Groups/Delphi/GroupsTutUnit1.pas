unit GroupsTutUnit1;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, vgr_WorkbookDesigner, vgr_Report,
  vgr_CommonClasses, vgr_DataStorage, StdCtrls;

type
  TForm1 = class(TForm)
    vgrWorkbook1: TvgrWorkbook;
    vgrReportEngine1: TvgrReportEngine;
    vgrWorkbookDesigner1: TvgrWorkbookDesigner;
    Query1: TQuery;
    Button1: TButton;
    vgrReportTemplate1: TvgrReportTemplate;
    vgrReportTemplate1vgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    vgrDetailBand1: TvgrDetailBand;
    vgrBand1: TvgrBand;
    vgrGroupBand1: TvgrGroupBand;
    vgrBand2: TvgrBand;
    vgrDataBand1: TvgrDataBand;
    vgrBand3: TvgrBand;
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
  vgrWorkbookDesigner1.Design(True);
end;

end.
