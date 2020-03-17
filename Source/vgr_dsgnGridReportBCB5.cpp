//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USEPACKAGE("vcl50.bpi");
USEPACKAGE("dsnide50.bpi");
USEPACKAGE("vgr_GridReportBCB5.bpi");
USEFORMNS("vgr_LocalizeExpertForm.pas", Vgr_localizeexpertform, vgrLocalizeExpertForm);
USEFORMNS("vgr_LocalizerExpertReportForm.pas", Vgr_localizerexpertreportform, vgrLocalizerExpertReportForm);
USEUNIT("vgr_Reg.pas");
USEUNIT("vgr_IDEAddon.pas");
USEUNIT("vgr_PropertyEditors.pas");
USEPACKAGE("VGR_COMMONCONTROLSBCB5.bpi");
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------

//   Package source.
//---------------------------------------------------------------------------

#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void*)
{
        return 1;
}
//---------------------------------------------------------------------------
 