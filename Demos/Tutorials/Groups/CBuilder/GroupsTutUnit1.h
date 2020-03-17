//---------------------------------------------------------------------------

#ifndef GroupsTutUnit1H
#define GroupsTutUnit1H
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
        TQuery *Query1;
        TButton *Button1;
        TvgrReportTemplateWorksheet *vgrReportTemplateWorksheet1;
        TvgrDetailBand *vgrDetailBand1;
        TvgrBand *vgrBand1;
        TvgrGroupBand *vgrGroupBand1;
        TvgrBand *vgrBand2;
        TvgrDataBand *vgrDataBand1;
        TvgrBand *vgrBand3;
        void __fastcall Button1Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
