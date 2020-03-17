{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

unit Global_DM;

interface

{$I vtk.inc}

uses
  SysUtils, Classes, DB, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} DBTables {$IFNDEF VTK_D6_OR_D7},
  Forms {$ENDIF};

type
  TGlobalDM = class(TDataModule)
    Customers: TTable;
    CustomersByCountry: TQuery;
    Orders: TTable;
    DSCustomers: TDataSource;
    SimpleCrosstabEvents: TTable;
    SimpleCrosstabCustoly: TTable;
    Reservat: TTable;
    GroupCrosstabEventsByEvent_Date: TQuery;
    GroupCrosstabCustoly: TTable;
    Biolife: TTable;
    BiolifeByCategory: TQuery;
    Events: TQuery;
    SPOST: TQuery;
    DS: TQuery;
    KS: TQuery;
    procedure SimpleCrosstabEventsAfterScroll(DataSet: TDataSet);
    procedure SimpleCrosstabCustolyAfterScroll(DataSet: TDataSet);
    procedure SimpleCrossTabSync;
    procedure GroupCrossTabSync;
    procedure GroupCrosstabEventsByEvent_DateAfterScroll(
      DataSet: TDataSet);
    procedure GroupCrosstabCustolyAfterScroll(DataSet: TDataSet);
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GlobalDM: TGlobalDM;

implementation

uses Main_Form;

{$R *.dfm}

procedure TGlobalDM.SimpleCrosstabEventsAfterScroll(DataSet: TDataSet);
begin
  SimpleCrossTabSync;
end;

procedure TGlobalDM.SimpleCrosstabCustolyAfterScroll(DataSet: TDataSet);
begin
  SimpleCrossTabSync;
end;

procedure TGlobalDM.SimpleCrossTabSync;
begin
  if (reservat.Active = true)  and (SimpleCrosstabevents.Active = true) and (SimpleCrosstabcustoly.Active = true) then
   begin
     if reservat.Locate('EventNo;CustNo', VarArrayOf([SimpleCrosstabevents.FieldValues['EventNo'],SimpleCrosstabcustoly.FieldValues['CustNo']]), [loPartialKey]) = true then
     begin
       MainForm.AliasManager.Nodes.Items[0].Value := reservat.FieldValues['Amt_Paid'];
     end
     else
     begin
       MainForm.AliasManager.Nodes.Items[0].Value := '-';
     end;
   end;
end;

procedure TGlobalDM.GroupCrossTabSync;
begin
  if (reservat.Active = true)  and (GroupCrosstabEventsByEvent_Date.Active = true) and (GroupCrosstabCustOly.Active = true) then
   begin
     if reservat.Locate('EventNo;CustNo', VarArrayOf([GroupCrosstabEventsByEvent_Date.FieldValues['EventNo'],GroupCrossTabCustoly.FieldValues['CustNo']]), [loPartialKey]) = true then
     begin
       MainForm.AliasManager.Nodes.Items[0].Value := reservat.FieldValues['Amt_Paid'];
     end
     else
     begin
       MainForm.AliasManager.Nodes.Items[0].Value := '-';
     end;
   end;
end;

procedure TGlobalDM.GroupCrosstabEventsByEvent_DateAfterScroll(
  DataSet: TDataSet);
begin
  GroupCrossTabSync;
end;

procedure TGlobalDM.GroupCrosstabCustolyAfterScroll(DataSet: TDataSet);
begin
  GroupCrossTabSync;
end;

procedure TGlobalDM.DataModuleCreate(Sender: TObject);
begin
  DS.DatabaseName := ExtractFilePath(ParamStr(0));
  DS.SQL.Text := 'select distinct dks as dks from spost order by dks';
  DS.Open;

  KS.DatabaseName := ExtractFilePath(ParamStr(0));
  KS.SQL.Text := 'select distinct kks as kks from spost order by kks';
  KS.Open;

  SPOST.DatabaseName := ExtractFilePath(ParamStr(0));
  SPOST.SQL.Text := 'select dks, kks, sum(s) as s from spost group by dks, kks order by dks, kks';
  SPOST.Open;
end;

end.
