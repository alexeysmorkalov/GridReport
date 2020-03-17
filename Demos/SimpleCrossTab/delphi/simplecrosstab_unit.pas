unit simplecrosstab_unit;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, vgr_WorkbookDesigner, vgr_Report, vgr_ScriptControl,
  {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}
  vgr_CommonClasses, vgr_DataStorage, StdCtrls, vgr_AliasManager;

type
  TForm1 = class(TForm)
    custoly: TTable;
    events: TTable;
    reservat: TTable;
    vgrWorkbook1: TvgrWorkbook;
    vgrScriptControl1: TvgrScriptControl;
    vgrReportEngine1: TvgrReportEngine;
    vgrWorkbookDesigner1: TvgrWorkbookDesigner;
    Button1: TButton;
    vgrWorkbook1vgrWorksheet1: TvgrWorksheet;
    vgrAliasManager1: TvgrAliasManager;
    vgrReportTemplate1: TvgrReportTemplate;
    vgrReportTemplate1vgrReportTemplateWorksheet1: TvgrReportTemplateWorksheet;
    vgrDetailBand1: TvgrDetailBand;
    vgrDataBand1: TvgrDataBand;
    vgrDetailBand2: TvgrDetailBand;
    vgrDataBand2: TvgrDataBand;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Button1Click(Sender: TObject);
begin
  vgrReportEngine1.Generate;
  vgrWorkbookDesigner1.Design(True);
end;





procedure TForm1.FormCreate(Sender: TObject);
begin
  custoly.Active := true;
  events.Active := true;  
  reservat.Active := true;
end;

end.
