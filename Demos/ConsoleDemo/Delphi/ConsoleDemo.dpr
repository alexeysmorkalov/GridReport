{******************************************}
{                                          }
{           vtk GridReport library         }
{       Win32 Console Demo for Delphi      }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}
{ This program demonstrates how to use the GridReport in the console application.
  Usage: Start the program with parameter "Simple customers list.grt" .The program 
  will create the workbook with the same name but with extension .grw }
         
program ConsoleDemo;

{$APPTYPE CONSOLE}

uses
  SysUtils, vgr_WorkbookGrid, vgr_CommonClasses, vgr_DataStorage, vgr_ScriptControl,
  vgr_Report, vgr_AliasManager, vgr_WorkbookDesigner, vgr_ReportDesigner, vgr_Dialogs,
  vgr_WorkbookPreviewForm, vgr_Writers, ImgList, ToolWin, shellapi,
  vgr_AliasManagerDesigner, vgr_WBScriptConstants, DBTables, Classes,
  ActiveX;

var
  AReportTemplate: TvgrReportTemplate;
  AWorkbook: TvgrWorkbook;
  AReportEngine: TvgrReportEngine;
  ATable: TTable;
  APathToTemplate: string;
  ADetailBand: TvgrDetailBand;
  AParent: TComponent;
  S: string;  
begin

  { TODO -oUser -cConsole Main : Insert code here }
  
  CoInitializeEx(nil,COINIT_APARTMENTTHREADED);

  AParent := TComponent.Create(nil);

  APathToTemplate := ExtractFilePath(ParamStr(1)) + ExtractFileName(ParamStr(1));
  if not FileExists(APathToTemplate) then
    begin
      writeln ('Template file not found.');
      exit;
    end;
    
  writeln ('Creating components...');
  AReportTemplate := TvgrReportTemplate.Create(AParent);
  AWorkbook := TvgrWorkbook.Create(AParent);
  AReportEngine := TvgrReportEngine.Create(AParent);

  writeln ('Loading template file...');
  AReportTemplate.LoadFromFile(APathToTemplate);

  writeln ('Linking components...');
  AReportEngine.Template := AReportTemplate;
  AReportEngine.Workbook := AWorkbook;
  AReportTemplate.Script.Language := 'VBScript';

  writeln ('Creating dataset...');
  ATable := TTable.Create(AParent);
  ATable.DatabaseName := 'DBDemos';
  ATable.TableName := 'customer';
  ATable.Name := 'Customers';
  ATable.Active := true;

  ADetailBand := AParent.FindComponent('vgrDetailBand1') as TvgrDetailBand;
  ADetailBand.DataSet := ATable;

  writeln ('Generating report...');
  AReportEngine.Generate;
  S := Copy(APathToTemplate,0,(Pos(ExtractFileExt(APathToTemplate), APathToTemplate)-1));
  AReportEngine.Workbook.SaveToFile(S+'.grw');

  AReportTemplate.Free;
  AWorkbook.Free;
  AReportEngine.Free;
  ATable.Free;

end.
