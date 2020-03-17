{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains TvgrReportSheetFormatForm class, which realize a dialog form for editing properties of worksheet of the report template.
See also:
  TvgrSheetFormatForm}
unit vgr_ReportSheetFormatDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs,

  vgr_Form, vgr_Report, StdCtrls, vgr_FormLocalizer;

type
{Implements a dialog form for editing properties of worksheet of the report template.
Example:
  var
    AEditor: TvgrReportSheetFormatForm;
  begin
    ...
    AEditor := TvgrReportSheetFormatForm.Create(Application);
    if AEditor.Execute(WorkbookGrid.ActiveWorksheet as TvgrReportTemplateWorksheet) then
    begin
      // do some actions
    end;
    ...
  end;}
  TvgrReportSheetFormatForm = class(TvgrDialogForm)
    gbBounds: TGroupBox;
    cbAutoBounds: TCheckBox;
    Label1: TLabel;
    Label2: TLabel;
    edWidth: TEdit;
    edHeight: TEdit;
    cbSkipGenerate: TCheckBox;
    bOk: TButton;
    bCancel: TButton;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure bOkClick(Sender: TObject);
    procedure cbAutoBoundsClick(Sender: TObject);
  private
    { Private declarations }
    FWorksheet: TvgrReportTemplateWorksheet;
  public
{Execute opens the worksheet editor dialog,
returning true when the user edit worksheet properties and clicks OK,
or false when the user cancels.
Parameters:
  AWorksheet - the worksheet of the report template for editing.
Return value:
  Returns true if the user edit worksheet properties and clicks OK.
See also:
  TvgrReportTemplateWorksheet}
    function Execute(AWorksheet: TvgrReportTemplateWorksheet): Boolean;
  end;

implementation

{$R *.dfm}
{$R ..\res\vgr_ReportSheetFormatDialogStrings.res}

function TvgrReportSheetFormatForm.Execute(AWorksheet: TvgrReportTemplateWorksheet): Boolean;
begin
  FWorksheet := AWorksheet;
  cbSkipGenerate.Checked := FWorksheet.SkipGenerate;
  cbAutoBounds.Checked := (FWorksheet.GenerateSize.cx < 0) and (FWorksheet.GenerateSize.cy < 0);
  if not cbAutoBounds.Checked then
  begin
    edWidth.Text := IntToStr(AWorksheet.GenerateSize.cx);
    edHeight.Text := IntToStr(AWorksheet.GenerateSize.cy);
  end;
  cbAutoBoundsClick(nil);

  Result := ShowModal = mrOk;
end;

procedure TvgrReportSheetFormatForm.bOkClick(Sender: TObject);
var
  ASize: TSize;
begin
  if ((StrToIntDef(edWidth.Text, -1) <= 0) or
      (StrToIntDef(edHeight.Text, -1) <= 0)) and
     not cbAutoBounds.Checked then
  begin
    ActiveControl := edWidth;
    exit;
  end;

  if cbAutoBounds.Checked then
  begin
    ASize.cx := -1;
    ASize.cy := -1;
  end
  else
  begin
    ASize.cx := StrToIntDef(edWidth.Text, -1);
    ASize.cy := StrToIntDef(edHeight.Text, -1);
  end;
  FWorksheet.GenerateSize := ASize;
  FWorksheet.SkipGenerate := cbSkipGenerate.Checked;
  ModalResult := mrOk;
end;

procedure TvgrReportSheetFormatForm.cbAutoBoundsClick(Sender: TObject);
begin
  edWidth.Enabled := not cbAutoBounds.Checked;
  edHeight.Enabled := not cbAutoBounds.Checked;
  label1.Enabled := not cbAutoBounds.Checked;
  label2.Enabled := not cbAutoBounds.Checked;
end;

end.

