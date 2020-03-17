/******************************************}
{                                          }
{           vtk GridReport library         }
{       Win32 Console Demo for CBuilder    }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}
{ This program demonstrates how to use the vtk GridReport in the console application.
  Usage: Start the program with parameter "Simple customers list.grt" .The program 
  will create the workbook with the same name but with extension .grw }
*/


#include <StdCtrls.hpp>
#include <vcl.h>
#pragma hdrstop

#include <excpt.h>
#include <iostream>
using std::cout;


#include <DBTables.hpp>

#include "vgr_CommonClasses.hpp"
#include "vgr_DataStorage.hpp"
#include "vgr_Report.hpp"
#include "vgr_ScriptControl.hpp"

#pragma link "vgr_CommonClasses"
#pragma link "vgr_DataStorage"
#pragma link "vgr_Report"
#pragma link "vgr_ScriptControl"

//---------------------------------------------------------------------------

#pragma argsused
int main(int argc, char* argv[])
{
  CoInitializeEx (NULL,COINIT_APARTMENTTHREADED);

  AnsiString ApathToTemplate;
  ApathToTemplate = ExtractFilePath(ParamStr(1)) + ExtractFileName(ParamStr(1));
  if(!FileExists(ApathToTemplate))
  {
    cout << "Template file not found.";
    return -1;
  }

  cout << "Creating components...\n";

  TComponent *AParent = new TComponent(NULL);
  TTable *ATable = new TTable(AParent);
  TvgrReportTemplate *AReportTemplate = new TvgrReportTemplate(AParent);
  TvgrWorkbook *AWorkbook = new TvgrWorkbook(AParent);
  TvgrReportEngine *AReportEngine = new TvgrReportEngine(AParent);
  TvgrScriptControl *AScriptControl = new TvgrScriptControl(AParent);
  TvgrDetailBand *ADetailBand = NULL;
  AnsiString S;

  try {
    cout << "Linking components...\n";
    AReportEngine->Template = AReportTemplate;
    AReportEngine->Workbook = AWorkbook;
    AReportTemplate->Script->ScriptControl = AScriptControl;
    AReportTemplate->Script->Language = "VBScript";

    cout << "Loading template...\n";
    AReportTemplate->LoadFromFile(ApathToTemplate);

    cout << "Creating dataset...\n";
    ATable->DatabaseName = "BCDEMOS";
    ATable->TableName = "customer";
    ATable->Name = "Customers";
    ATable->Active = true;

    ADetailBand = (TvgrDetailBand *)AParent->FindComponent("vgrDetailBand1");
    ADetailBand->DataSet = ATable;

    cout << "Generating report...\n";
    AReportEngine->Generate();
    S = ApathToTemplate.SubString(0,ApathToTemplate.Pos(ExtractFileExt(ApathToTemplate))-1);
    AReportEngine->Workbook->SaveToFile(S+".grw");
  }
  catch (...) {

  }
  delete(ATable);
  delete(AReportTemplate);
  delete(AWorkbook);
  delete(AReportEngine);
  delete(AScriptControl);
  delete(AParent);

  return 0;
}

