{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains TvgrReportOptionsDialogForm class, which realize a dialog form for editing properties of the report template.
See also:
  TvgrReportOptionsDialogForm}
unit vgr_ReportOptionsDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7}Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls,

  vgr_Report, vgr_Form, vgr_AliasManager, vgr_FormLocalizer;

type
  /////////////////////////////////////////////////
  //
  // TvgrReportOptionsDialogForm
  //
  /////////////////////////////////////////////////
  {Implements a dialog form for editing properties of the report template.
Example:
  var
    AEditor: TvgrReportOptionsDialogForm;
  begin
  ...
  AEditor := TvgrReportOptionsDialogForm.Create(Application);
  if AEditor.Execute(MyReportTemplate) then
  begin
    do some actions
  end;
  ...
  end;}
  TvgrReportOptionsDialogForm = class(TvgrDialogForm)
    PageControl: TPageControl;
    PScript: TTabSheet;
    bOk: TButton;
    bCancel: TButton;
    Label1: TLabel;
    edScriptLanguage: TEdit;
    vgrFormLocalizer1: TvgrFormLocalizer;
    Label2: TLabel;
    edAliasManager: TComboBox;
  private
    { Private declarations }
  public
{Execute opens the report template editor dialog,
returning true when the user edit properties of the report template and clicks OK,
or false when the user cancels.
Parameters:
  AReportTemplate - the report template for editing.
Return value:
  Returns true if the user edit the report template properties and clicks OK.
See also:
  TvgrReportTemplate}
    function Execute(AReportTemplate: TvgrReportTemplate): Boolean;
  end;

implementation

{$R *.dfm}
{$R ..\res\vgr_ReportOptionsDialogStrings.res}

/////////////////////////////////////////////////
//
// TvgrReportOptionsDialogForm
//
/////////////////////////////////////////////////
function TvgrReportOptionsDialogForm.Execute(AReportTemplate: TvgrReportTemplate): Boolean;
begin
  edScriptLanguage.Text := AReportTemplate.Script.Language;
  BuildComponentsList(AReportTemplate, edAliasManager.Items, TvgrAliasManager);
  edAliasManager.Text := ComponentToText(AReportTemplate, AReportTemplate.Script.AliasManager, edAliasManager.Items);
  Result := ShowModal = mrOk;
  if Result then
  begin
    AReportTemplate.Script.Language := edScriptLanguage.Text;
    AReportTemplate.Script.AliasManager := TvgrAliasManager(TextToComponent(AReportTemplate, edAliasManager.Text, TvgrAliasManager, edAliasManager.Items));
  end;
end;

end.
