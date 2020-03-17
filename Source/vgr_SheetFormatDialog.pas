{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains TvgrSheetFormatForm class, which realize a dialog form for editing properties of the worksheet.
See also:
  TvgrSheetFormatForm}
unit vgr_SheetFormatDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,

  vgr_DataStorage, vgr_Form, vgr_FormLocalizer;

type
  /////////////////////////////////////////////////
  //
  // TvgrSheetFormatForm
  //
  /////////////////////////////////////////////////
{Implements a dialog form for editing properties of the worksheet.
Example:
  var
    AEditor: TvgrSheetFormatForm;
  begin
  ...
  AEditor := TvgrSheetFormatForm.Create(Application);
  if AEditor.Execute(WorkbookGrid.ActiveWorksheet) then
  begin
    // do some actions
  end;
  ...
  end;}
  TvgrSheetFormatForm = class(TvgrDialogForm)
    Label1: TLabel;
    Label2: TLabel;
    edSheetTitle: TEdit;
    edSheetName: TEdit;
    bOk: TButton;
    bCancel: TButton;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure bOkClick(Sender: TObject);
  private
    { Private declarations }
    FWorksheet: TvgrWorksheet;
  public
{Execute opens the worksheet editor dialog,
returning true when the user edit worksheet properties and clicks OK,
or false when the user cancels.
Parameters:
  AWorksheet - the worksheet for editing.
Return value:
  Returns true if the user edit worksheet properties and clicks OK.
See also:
  TvgrWorksheet}
    function Execute(AWorksheet: TvgrWorksheet): Boolean;
  end;

implementation

uses
  vgr_Dialogs, vgr_Functions, vgr_StringIDs, vgr_Localize;
  
{$R *.dfm}
{$R ..\res\vgr_SheetFormatDialogStrings.res}

function TvgrSheetFormatForm.Execute(AWorksheet: TvgrWorksheet): Boolean;
begin
  FWorksheet := AWorksheet;
  edSheetTitle.Text := AWorksheet.Title;
  edSheetName.Text := AWorksheet.Name;
  Result := ShowModal = mrOk;
end;

procedure TvgrSheetFormatForm.bOkClick(Sender: TObject);
begin
  if not CheckComponentName(FWorksheet, edSheetName.Text) then
  begin
    ActiveControl := edSheetName;
    MBoxFmt(vgrLoadStr(svgrid_vgr_SheetFormatDialog_InvalidSheetName), MB_OK or MB_ICONERROR, [edSheetName.Text]);
    exit;
  end;
  FWorksheet.Name := edSheetName.Text;
  FWorksheet.Title := edSheetTitle.Text;
  ModalResult := mrOk;
end;

end.
