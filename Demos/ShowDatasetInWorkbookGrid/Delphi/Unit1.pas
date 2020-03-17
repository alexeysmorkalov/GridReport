unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Math, StdCtrls, ExtCtrls, Db, DBTables, vgr_WorkbookGrid, vgr_CommonClasses,
  vgr_DataStorage, vgr_DataStorageTypes, vgr_Functions;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    EDDatabase: TEdit;
    Label2: TLabel;
    MQuery: TMemo;
    Database: TDatabase;
    Query: TQuery;
    Workbook: TvgrWorkbook;
    vgrWorkbookGrid1: TvgrWorkbookGrid;
    bExecute: TButton;
    procedure Panel1Resize(Sender: TObject);
    procedure bExecuteClick(Sender: TObject);
  private
    { Private declarations }
    procedure ShowQuery(Query: TQuery);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Panel1Resize(Sender: TObject);
begin
  bExecute.Left := Panel1.ClientWidth - bExecute.Width - 4;
  MQuery.Width := bExecute.Left - MQuery.Left - 4;
end;

procedure TForm1.ShowQuery(Query: TQuery);
var
  AWorksheet: TvgrWorksheet;
  I, J, ATableColCount: Integer;
  AOdd: Boolean;
begin
  AWorksheet := Workbook.AddWorksheet;
  AWorksheet.Title := IntToStr(Workbook.WorksheetsCount);

  // init variables
  ATableColCount := Max(4, Query.FieldCount); // size of table in columns

  // write DatabaseName
  with AWorksheet.Ranges[0, 0, 1, 0] do
  begin
    Value := 'Database:';
    Font.Style := [fsBold];
    VertAlign := vgrvaCenter;
  end;
  with AWorksheet.Ranges[2, 0, ATableColCount - 1, 0] do
  begin
    Value := Database.AliasName;
    Font.Name := 'Tahoma';
    Font.Size := 12;
    VertAlign := vgrvaCenter;
    FillBackColor := $99CCCC;
  end;
  AWorksheet.Rows[0].Height := ConvertUnitsToTwips(1, vgruCentimeters);

  // write SQL query
  with AWorksheet.Ranges[0, 1, 1, 1] do
  begin
    Value := 'Query:';
    Font.Style := [fsBold];
    VertAlign := vgrvaCenter;
  end;
  AWorksheet.Rows[1].Height := ConvertUnitsToTwips(1, vgruCentimeters);
  with AWorksheet.Ranges[0, 2, ATableColCount - 1, 2] do
  begin
    Value := Query.SQL.Text;
    Font.Name := 'Tahoma';
    Font.Size := 12;
    WordWrap := True;
    FillBackColor := $99CCCC;
  end;
  AWorksheet.Rows[2].Height := ConvertUnitsToTwips(1, vgruCentimeters);

  // write table header
  J := 5;
  for I := 0 to Query.FieldCount - 1 do
  begin
    with AWorksheet.Ranges[I, J, I, J] do
    begin
      Value := Query.Fields[I].DisplayName;
      FillBackColor := clSilver;
      HorzAlign := vgrhaCenter;
    end;
  end;

  // write table data
  J := J + 1;
  AOdd := true;
  Query.First;
  while not Query.Eof do
  begin
    for I := 0 to Query.FieldCount - 1 do
    begin
      with AWorksheet.Ranges[I, J, I, J] do
      begin
        Value := Query.Fields[I].AsVariant;
        if not AOdd then
          FillBackColor := $FFFFCC
      end;
    end;
    Query.Next;
    J := J + 1;
    AOdd := not AOdd;
  end;
end;

procedure TForm1.bExecuteClick(Sender: TObject);
begin
  try
    Query.Close;
    Database.Connected := False;
    Database.AliasName := EDDatabase.Text;
    Database.Connected := True;
    Query.SQL.Text := MQuery.Lines.Text;
    Query.Open;
  except
    on e: Exception do
    begin
      ShowMessage(Format('Error execute query:'#13#13'%s', [e.Message]));
      exit;
    end;
  end;

  Workbook.BeginUpdate; // must be called before global changes in workbook
  try
    ShowQuery(Query);
  finally
    Workbook.EndUpdate;
  end;
  vgrWorkbookGrid1.ActiveWorksheetIndex := Workbook.WorksheetsCount - 1;
end;

end.
