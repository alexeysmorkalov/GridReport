{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains TvgrScriptEventEditForm class, which realize a dialog form for selecting a name of the script procedure.
Instance of this class created by the TvgrReportDesigner.
See also:
  TvgrReportDesigner}
unit vgr_ScriptEventEditDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, typinfo, StdCtrls,

  vgr_ScriptParser, vgr_Form, vgr_FormLocalizer;

type

  /////////////////////////////////////////////////
  //
  // TvgrScriptEventEditForm
  //
  /////////////////////////////////////////////////
{Implements a dialog for for selecting a name of the script procedure.
The list of available procedures are provided by the instance of TvgrScriptProgramInfo class.
See also:
  TvgrScriptProgramInfo}
  TvgrScriptEventEditForm = class(TvgrDialogForm)
    Label1: TLabel;
    edProcName: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    bOK: TButton;
    bCancel: TButton;
    vgrFormLocalizer1: TvgrFormLocalizer;
  private
    { Private declarations }
  public
{Execute opens the selecting dialog,
returning true when the user select procedure and clicks OK,
or false when the user cancels.
Parameters:
  ADesc - short description.
  AScriptInfo - TvgrScriptProgramInfo object, that contains list of the available procedures.
  AProcedureName - receives the name of selected procedure.
Return value:
  Returns true if the user select procedure and clicks OK.}
    function Execute(const ADesc: string; AScriptInfo: TvgrScriptProgramInfo; var AProcedureName: string): Boolean;
  end;

implementation

{$R *.dfm}
{$R ..\res\vgr_ScriptEventEditDialogStrings.RES}

/////////////////////////////////////////////////
//
// TvgrScriptEventEditForm
//
/////////////////////////////////////////////////
function TvgrScriptEventEditForm.Execute(const ADesc: string; AScriptInfo: TvgrScriptProgramInfo; var AProcedureName: string): Boolean;
var
  I: Integer;
begin
  Label3.Caption := ADesc;
  for I := 0 to AScriptInfo.ProcedureCount - 1 do
    edProcName.Items.Add(AScriptInfo.Procedures[I].Name);
  edProcName.Text := AProcedureName;
  Result := ShowModal = mrOk;
  if Result then
    AProcedureName := edProcName.Text;
end;

end.
