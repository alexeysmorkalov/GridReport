//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MasterDetailTutUnit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "vgr_CommonClasses"
#pragma link "vgr_DataStorage"
#pragma link "vgr_Report"
#pragma link "vgr_WorkbookDesigner"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{
  vgrReportEngine1->Generate();
  vgrWorkbookDesigner1->Design(TRUE);        
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Table1AfterScroll(TDataSet *DataSet)
{
  Table2->Filter = "[CustNo] = " + Table1->FieldByName("CustNo")->AsString;
  Table2->Filtered = true;         
}
//---------------------------------------------------------------------------

