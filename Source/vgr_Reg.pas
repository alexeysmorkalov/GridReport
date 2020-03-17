{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains Register procedure, that registers all visual and non visual components of the GridReport library.}
unit vgr_Reg;

{$I vtk.inc}

interface

uses
  Classes,
  {$IFDEF VTK_D6_OR_D7} DesignIntf, DesignEditors, DesignMenus,
  {$ELSE} dsgnintf, {$ENDIF} SysUtils,

  {$IFDEF VTK_D7} vgr_IDEAddon, ToolsAPI, {$ENDIF}
  vgr_DataStorage, vgr_ScriptEdit, vgr_ScriptControl,
  vgr_WorkbookGrid, vgr_Report,
  vgr_ReportDesigner, vgr_WorkbookDesigner, vgr_CellPropertiesDialog,
  vgr_WorkbookPreview,
  vgr_PrinterComboBox, vgr_AliasManager, vgr_PageSetupDialog,
  vgr_WorkbookPreviewForm, vgr_PrintSetupDialog, vgr_PrintEngine,
  vgr_PropertyEditors, vgr_FormLocalizer;

{Registers all visual and non visual components of the GridReport library.
See also:
  TvgrWorkbook, TvgrReportTemplate, TvgrWorkbookGrid, TvgrWorkbookPreview,
  TvgrScriptEdit, TvgrScriptControl, TvgrReportDesigner,
  TvgrReportEngine, TvgrWorkbookDesigner, TvgrCellPropertiesDialog,
  TvgrPrinterComboBox, TvgrAliasManager, TvgrPageSetupDialog,
  TvgrWorkbookPreviewer, TvgrPrintSetupDialog, TvgrPrintEngine, TvgrFormLocalizer}
procedure Register;

implementation

uses
  vgr_Functions;
  
{$R *.res}

{Registers all visual and non visual components of the GridReport library.
See also:
  TvgrWorkbook, TvgrReportTemplate, TvgrWorkbookGrid, TvgrWorkbookPreview,
  TvgrScriptEdit, TvgrScriptControl, TvgrReportDesigner,
  TvgrReportEngine, TvgrWorkbookDesigner, TvgrCellPropertiesDialog,
  TvgrPrinterComboBox, TvgrAliasManager, TvgrPageSetupDialog,
  TvgrWorkbookPreviewer, TvgrPrintSetupDialog, TvgrPrintEngine, TvgrFormLocalizer}
procedure Register;
begin
  RegisterComponents('vtkTools GridReport', [TvgrWorkbook]);
  RegisterComponents('vtkTools GridReport', [TvgrReportTemplate]);
  RegisterComponents('vtkTools GridReport', [TvgrReportEngine]);
  RegisterComponents('vtkTools GridReport', [TvgrAliasManager]);
  RegisterComponents('vtkTools GridReport', [TvgrWorkbookGrid]);
  RegisterComponents('vtkTools GridReport', [TvgrWorkbookPreview]);
  RegisterComponents('vtkTools GridReport', [TvgrScriptEdit]);
  RegisterComponents('vtkTools GridReport', [TvgrScriptControl]);
  RegisterComponents('vtkTools GridReport', [TvgrReportDesigner]);
  RegisterComponents('vtkTools GridReport', [TvgrWorkbookDesigner]);
  RegisterComponents('vtkTools GridReport', [TvgrCellPropertiesDialog]);
  RegisterComponents('vtkTools GridReport', [TvgrPrinterComboBox]);
  RegisterComponents('vtkTools GridReport', [TvgrPageSetupDialog]);
  RegisterComponents('vtkTools GridReport', [TvgrWorkbookPreviewer]);
  RegisterComponents('vtkTools GridReport', [TvgrPrintSetupDialog]);
  RegisterComponents('vtkTools GridReport', [TvgrPrintEngine]);
  RegisterComponents('vtkTools GridReport', [TvgrFormLocalizer]);

  RegisterNoIcon([TvgrWorksheet,
                  TvgrReportTemplateWorksheet,
                  TvgrBand,
                  TvgrDetailBand,
                  TvgrGroupBand,
                  TvgrDataBand]);

{$IFDEF VTK_D7}
  RegisterPackageWizard(TvgrLocalizeExpert.Create);
{$ENDIF}

  RegisterPropertyEditor(TypeInfo(string), TvgrCustomScriptEdit, 'FontName', TvgrFixedFontNameProperty);
  RegisterPropertyEditor(TypeInfo(TvgrAliasNodes), TvgrAliasManager, 'Nodes', TvgrAliasNodesProperty);

  RegisterComponentEditor(TvgrWorkbook, TvgrWorkbookEditor);
  RegisterComponentEditor(TvgrReportTemplate, TvgrReportEditor);
  RegisterComponentEditor(TvgrAliasManager, TvgrAliasEditor);
end;

end.
