{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains the TvgrBandFormatForm class, that implements a dialog form for editing band properties.}
unit vgr_BandFormatDialog;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ComCtrls, DB,

  vgr_Report, vgr_Form, vgr_FormLocalizer;

type
{Implements a dialog form for editing the band of the report template.
 
Example:
  var
    AEditor: TvgrBandFormatForm;
  begin
  ...
  AEditor := TvgrBandFormatForm.Create(Application);
  if AEditor.Execute(MyBand) then
  begin
    // do some actions
  end;
  ...
  end;}
  TvgrBandFormatForm = class(TvgrDialogForm)
    Image: TImage;
    Label1: TLabel;
    Label2: TLabel;
    edBandName: TEdit;
    PageControl: TPageControl;
    PCommon: TTabSheet;
    PData: TTabSheet;
    cbSkipGenerate: TCheckBox;
    GroupBox1: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    edStartPos: TEdit;
    udStartPos: TUpDown;
    edEndPos: TEdit;
    udEndPos: TUpDown;
    bOk: TButton;
    bCancel: TButton;
    Label5: TLabel;
    Label6: TLabel;
    edDataset: TComboBox;
    edGroupExpression: TComboBox;
    vgrFormLocalizer1: TvgrFormLocalizer;
    cbRowColAutoSize: TCheckBox;
    cbCreateSections: TCheckBox;
    procedure bOkClick(Sender: TObject);
  private
    FBand: TvgrBand;
  public
{Execute opens the band editor dialog,
returning true when the user edit band properties and clicks OK,
or false when the user cancels.
Parameters:
  ABand - band for editing.
Return value:
  Returns true if the user edit band properties and clicks OK.}
    function Execute(ABand: TvgrBand): Boolean;
  end;

implementation

uses
  vgr_Functions, vgr_Dialogs, vgr_StringIDs, vgr_Localize;

{$R *.dfm}
{$R ..\res\vgr_BandFormatDialogStrings.res}

function TvgrBandFormatForm.Execute(ABand: TvgrBand): Boolean;
begin
  FBand := ABand;
  Label1.Caption := FBand.GetCaption;
  Image.Picture.Assign(BandClasses.Bitmaps[BandClasses.GetBandClassIndex(ABand)]);

  edBandName.Text := ABand.Name;
  cbSkipGenerate.Checked := ABand.SkipGenerate;
  cbRowColAutoSize.Checked := ABand.RowColAutoSize;
  cbCreateSections.Checked := ABand.CreateSections;
  udStartPos.Position := ABand.StartPos;
  udEndPos.Position := ABand.EndPos;
  PData.TabVisible := (ABand is TvgrDetailBand) or (ABand is TvgrGroupBand);
  Label5.Enabled := ABand is TvgrDetailBand;
  edDataSet.Enabled := ABand is TvgrDetailBand;
  Label6.Enabled := ABand is TvgrGroupBand;
  edGroupExpression.Enabled := ABand is TvgrGroupBand;

  if ABand is TvgrDetailBand then
    with TvgrDetailBand(ABand) do
    begin
      BuildComponentsList(Template, edDataset.Items, TDataset);
      if Dataset <> nil then
        edDataSet.Text := ComponentToText(Template, Dataset, edDataset.Items);
    end;

  if ABand is TvgrGroupBand then
    with TvgrGroupBand(ABand) do
    begin
      edGroupExpression.Text := GroupExpression;
    end;

  Result := ShowModal = mrOk;
end;

procedure TvgrBandFormatForm.bOkClick(Sender: TObject);
begin
  if not CheckComponentName(FBand, edBandName.Text) then
  begin
    ActiveControl := edBandName;
    MBoxFmt(vgrLoadStr(svgrid_vgr_BandFormatDialog_InvalidBandName), MB_OK or MB_ICONERROR, [edBandName.Text]);
    exit;
  end;

  if StrToIntDef(edStartPos.Text, -1) = -1 then
  begin
    ActiveControl := edStartPos;
    MBoxFmt(vgrLoadStr(svgrid_vgr_BandFormatDialog_InvalidBandStartPos), MB_OK or MB_ICONERROR, [edStartPos.Text]);
    exit;
  end;

  if StrToIntDef(edEndPos.Text, -1) = -1 then
  begin
    ActiveControl := edStartPos;
    MBoxFmt(vgrLoadStr(svgrid_vgr_BandFormatDialog_InvalidBandEndPos), MB_OK or MB_ICONERROR, [edEndPos.Text]);
    exit;
  end;

  if StrToInt(edStartPos.Text) > StrToInt(edEndPos.Text) then
  begin
    ActiveControl := edStartPos;
    MBox(vgrLoadStr(svgrid_vgr_BandFormatDialog_InvalidBandPosition), MB_OK or MB_ICONERROR);
    exit;
  end;

  FBand.Name := edBandName.Text;
  FBand.RowColAutoSize := cbRowColAutoSize.Checked;
  FBand.CreateSections := cbCreateSections.Checked;
  FBand.SkipGenerate := cbSkipGenerate.Checked;
  FBand.StartPos := StrToInt(edStartPos.Text);
  FBand.EndPos := StrToInt(edEndPos.Text);
  if FBand is TvgrDetailBand then
    with TvgrDetailBand(FBand) do
      Dataset := TDataset(TextToComponent(Template, edDataset.Text, TDataset, edDataset.Items));
  if FBand is TvgrGroupBand then
    with TvgrGroupBand(FBand) do
      GroupExpression := edGroupExpression.Text;

  ModalResult := mrOk;
end;

end.
