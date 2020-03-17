//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEUNIT("vgr_Writers.pas");
USEUNIT("vgr_AliasManager.pas");
USEFORMNS("vgr_AliasManagerDesigner.pas", Vgr_aliasmanagerdesigner, vgrAliasManagerDesignerForm);
USEUNIT("vgr_AXScript.pas");
USEFORMNS("vgr_BandFormatDialog.pas", Vgr_bandformatdialog, vgrBandFormatForm);
USEUNIT("vgr_BIFF8_Types.pas");
USEFORMNS("vgr_CellPropertiesDialog.pas", Vgr_cellpropertiesdialog, vgrCellPropertiesForm);
USEUNIT("vgr_CommonClasses.pas");
USEUNIT("vgr_Controls.pas");
USEFORMNS("vgr_CopyMoveSheetDialog.pas", Vgr_copymovesheetdialog, vgrCopyMoveSheetDialogForm);
USEUNIT("vgr_DataStorage.pas");
USEUNIT("vgr_DataStorageRecords.pas");
USEUNIT("vgr_DataStorageTypes.pas");
USEUNIT("vgr_Dialogs.pas");
USEUNIT("vgr_Event.pas");
USEUNIT("vgr_ExcelConsts.pas");
USEUNIT("vgr_ExcelFormula.pas");
USEUNIT("vgr_ExcelFormula_iftab.pas");
USEUNIT("vgr_ExcelTypes.pas");
USEUNIT("vgr_Form.pas");
USEUNIT("vgr_FormLocalizer.pas");
USEUNIT("vgr_FormulaCalculator.pas");
USEUNIT("vgr_Keyboard.pas");
USEUNIT("vgr_Localize.pas");
USEUNIT("vgr_PageMaker.pas");
USEUNIT("vgr_PageProperties.pas");
USEFORMNS("vgr_PageSetupDialog.pas", Vgr_pagesetupdialog, vgrPageSetupDialogForm);
USEUNIT("vgr_PrintEngine.pas");
USEUNIT("vgr_Printer.pas");
USEUNIT("vgr_PrinterComboBox.pas");
USEFORMNS("vgr_PrintSetupDialog.pas", Vgr_printsetupdialog, vgrPrintSetupDialogForm);
USEUNIT("vgr_Report.pas");
USEFORMNS("vgr_ReportDesigner.pas", Vgr_reportdesigner, vgrReportDesignerForm);
USEUNIT("vgr_ReportFunctions.pas");
USEUNIT("vgr_ReportGUIFunctions.pas");
USEFORMNS("vgr_ReportOptionsDialog.pas", Vgr_reportoptionsdialog, vgrReportOptionsDialogForm);
USEFORMNS("vgr_RowColSizeDialog.pas", Vgr_rowcolsizedialog, vgrRowColSizeForm);
USEUNIT("vgr_ScriptComponentsProvider.pas");
USEUNIT("vgr_ScriptControl.pas");
USEUNIT("vgr_ScriptEdit.pas");
USEFORMNS("vgr_ScriptEventEditDialog.pas", Vgr_scripteventeditdialog, vgrScriptEventEditForm);
USEUNIT("vgr_ScriptParser.pas");
USEUNIT("vgr_ScriptUtils.pas");
USEFORMNS("vgr_SheetFormatDialog.pas", Vgr_sheetformatdialog, vgrSheetFormatForm);
USEUNIT("vgr_StringIDs.pas");
USEUNIT("vgr_WBScriptConstants.pas");
USEFORMNS("vgr_WorkbookDesigner.pas", Vgr_workbookdesigner, vgrWorkbookDesignerForm);
USEUNIT("vgr_WorkbookGrid.pas");
USEUNIT("vgr_WorkbookPainter.pas");
USEUNIT("vgr_WorkbookPreview.pas");
USEFORMNS("vgr_WorkbookPreviewForm.pas", Vgr_workbookpreviewform, vgrWorkbookPreviewForm);
USEPACKAGE("vcl50.bpi");
USEPACKAGE("vgr_CommonControlsBCB5.bpi");
USEPACKAGE("VCLX50.bpi");
USEPACKAGE("VCLDB50.bpi");
USEFORMNS("vgr_ReportAdd.pas", Vgr_reportadd, vgrReportAddForm);
USEUNIT("vgr_ScriptDispIDs.pas");
USEUNIT("vgr_ScriptGlobalFunctions.pas");
USEFORMNS("vgr_HeaderFooterOptionsDialog.pas", Vgr_headerfooteroptionsdialog, vgrHeaderFooterOptionsDialogForm);
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
