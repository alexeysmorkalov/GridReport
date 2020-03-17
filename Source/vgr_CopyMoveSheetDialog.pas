{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{ Contains TvgrCopyMoveSheetDialogForm class, which realize a dialog form for copying and rearranging worksheet within workbook.}
unit vgr_CopyMoveSheetDialog;

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
  // TvgrCopyMoveSheetDialogForm
  //
  /////////////////////////////////////////////////
  { Implements a dialog form for copying and rearranging worksheet within workbook. }
  TvgrCopyMoveSheetDialogForm = class(TvgrDialogForm)
    Label1: TLabel;
    lbSheets: TListBox;
    cbCreateCopy: TCheckBox;
    bOk: TButton;
    bCancel: TButton;
    Label2: TLabel;
    lSelectedSheet: TLabel;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure lbSheetsClick(Sender: TObject);
  private
    FWorksheet: TvgrWorksheet;
    procedure UpdateEnabled;
  protected
    property Worksheet: TvgrWorksheet read FWorksheet;
  public
    { Execute opens the worksheet editor dialog.
Parameters:
  AWorksheet - worksheet for copy or rearraging.
Return value:
  Boolean, returning true when the user edit worksheets and clicks OK, or false when the user cancels.}
    function Execute(AWorksheet: TvgrWorksheet): Boolean;
  end;

implementation

{$R *.dfm}
{$R ..\res\vgr_CopyMoveSheetDialogStrings.RES}

/////////////////////////////////////////////////
//
// TvgrCopyMoveSheetDialogForm
//
/////////////////////////////////////////////////
procedure TvgrCopyMoveSheetDialogForm.UpdateEnabled;
begin
  with lbSheets do
    bOk.Enabled := (ItemIndex >= 0) and
                   (Items.Objects[ItemIndex] <> Worksheet);
end;

function TvgrCopyMoveSheetDialogForm.Execute(AWorksheet: TvgrWorksheet): Boolean;
var
  I: Integer;
begin
  FWorksheet := AWorksheet;
  lSelectedSheet.Caption := AWorksheet.Title;
  
  for I := 0 to AWorksheet.Workbook.WorksheetsCount - 1 do
    lbSheets.Items.AddObject(AWorksheet.Workbook.Worksheets[I].Title, AWorksheet.Workbook.Worksheets[I]);
  if lbSheets.Items.Count > 0 then
    lbSheets.ItemIndex := 0;
  Result := ShowModal = mrOk;
  if Result then
  begin
    if cbCreateCopy.Checked then
    begin
      AWorksheet.Copy(lbSheets.ItemIndex);
    end
    else
    begin
      if lbSheets.ItemIndex > 0 then
        AWorksheet.IndexInWorkbook := lbSheets.ItemIndex - 1
      else
        AWorksheet.IndexInWorkbook := lbSheets.ItemIndex;
    end;
  end;
end;

procedure TvgrCopyMoveSheetDialogForm.lbSheetsClick(Sender: TObject);
begin
  UpdateEnabled;
end;

end.
