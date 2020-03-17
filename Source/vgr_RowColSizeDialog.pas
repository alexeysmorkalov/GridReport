{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains the TvgrRowColSizeForm class, which realize a dialog form for editing sizes of the worsheet row or column.}
unit vgr_RowColSizeDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Math,

  vgr_Form, vgr_FormLocalizer;

type
{Implements a dialog form for editing sizes of the worsheet row or column.

Example:
  var
    AEditor: TvgrRowColSizeForm;
    ASize: Integer;
  begin
  ...
  AEditor := TvgrRowColSizeForm.Create(Application);
  ASize := vgrWorksheet1.Rows[0].Size;
  if AEditor.Execute('Editing height of the row', 'Height:', ASize) then
  begin
    // do some actions
    vgrWorksheet1.Rows[0].Size := ASize;
  end;
  ...
  end;}
  TvgrRowColSizeForm = class(TvgrDialogForm)
    Label1: TLabel;
    edSize: TEdit;
    edUnits: TComboBox;
    bOk: TButton;
    bCancel: TButton;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure bOkClick(Sender: TObject);
  private
    { Private declarations }
  protected
    procedure DoLocalize; override;
  public
{Execute opens the editor dialog,
returning true when the user edit band properties and clicks OK,
or false when the user cancels.
You must assign updated value to the destination.
Parameters:
  ACaption - the title of the window.
  AText - the short description.
  ASize - the initially size, this parameter contains new value, if user clicks OK.
Return value:
  Returns true if the user edit band properties and clicks OK.}
    function Execute(const ACaption, AText: string; var ASize: Integer): Boolean;
  end;

implementation

uses
  vgr_Functions, vgr_GUIFunctions, vgr_Localize, vgr_StringIDs;

{$R *.dfm}
{$R ..\res\vgr_RowColSizeDialogStrings.RES}

procedure TvgrRowColSizeForm.DoLocalize;
begin
  inherited;
  edUnits.Items.Clear;
  edUnits.Items.Add(vgrLoadStr(svgrid_Common_vgruMms));
  edUnits.Items.Add(vgrLoadStr(svgrid_Common_vgruCentimeters));
  edUnits.Items.Add(vgrLoadStr(svgrid_Common_vgruInches));
  edUnits.Items.Add(vgrLoadStr(svgrid_Common_vgruTenthsMMs));
  edUnits.Items.Add(vgrLoadStr(svgrid_Common_vgruPixels));
end;

function TvgrRowColSizeForm.Execute(const ACaption, AText: string; var ASize: Integer): Boolean;
var
  v: Extended;
  ACode: Integer;
begin
  Caption := ACaption;
  Label1.Caption := AText;
  edUnits.ItemIndex := 0;
  if ASize < 0 then
    edSize.Text := ''
  else
    edSize.Text := ConvertTwipsToUnitsStr(ASize, TvgrUnits(edUnits.ItemIndex));
  Result := ShowModal = mrOk;
  if Result then
  begin
    Val(edSize.Text, v, ACode);
    Result := (ACode = 0) and (v >= 0);
    if Result then
      ASize := ConvertUnitsToTwips(V, TvgrUnits(edUnits.ItemIndex));
  end;
end;

procedure TvgrRowColSizeForm.bOkClick(Sender: TObject);
var
  v: Extended;
  ACode: Integer;
begin
  Val(edSize.Text, v, ACode);
  if (ACode <> 0) or (v < 0) then
  begin
    ActiveControl := edSize;
    exit;
  end;
  ModalResult := mrOk;
end;

end.
