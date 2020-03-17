//---------------------------------------------------------------------------

#ifndef MasterDetailTutUnit1H
#define MasterDetailTutUnit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "vgr_CommonClasses.hpp"
#include "vgr_DataStorage.hpp"
#include "vgr_Report.hpp"
#include "vgr_WorkbookDesigner.hpp"
#include <DB.hpp>
#include <DBTables.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TvgrReportTemplate *vgrReportTemplate1;
        TvgrWorkbook *vgrWorkbook1;
        TvgrReportEngine *vgrReportEngine1;
        TvgrWorkbookDesigner *vgrWorkbookDesigner1;
        TTable *Table1;
        TTable *Table2;
        TButton *Button1;
        TvgrReportTemplateWorksheet *vgrReportTemplate1vgrReportTemplateWorksheet1;
        TvgrDetailBand *vgrDetailBand1;
        TvgrDataBand *vgrDataBand1;
        TvgrDetailBand *vgrDetailBand2;
        TvgrBand *vgrBand1;
        TvgrDataBand *vgrDataBand2;
        TvgrBand *vgrBand2;
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Table1AfterScroll(TDataSet *DataSet);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
