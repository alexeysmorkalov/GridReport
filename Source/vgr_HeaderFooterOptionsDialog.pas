{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

{Contains the TvgrHeaderFooterOptionsDialogForm form which can be used for editing properties of TvgrPageHeaderFooter class.
See also:
  TvgrPageHeaderFooter}
unit vgr_HeaderFooterOptionsDialog;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Classes, Graphics, Controls, Forms,
  Dialogs,

  vgr_Form, vgr_PageProperties, StdCtrls, ExtCtrls, Buttons, vgr_Button,
  vgr_ColorButton, vgr_FormLocalizer;

type
{Dialog form for editing properties of TvgrPageHeaderFooter class.
See also:
  TvgrPageHeaderFooter}
  TvgrHeaderFooterOptionsDialogForm = class(TvgrDialogForm)
    mLeft: TMemo;
    mCenter: TMemo;
    mRight: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    bOk: TButton;
    bCancel: TButton;
    Label5: TLabel;
    bBackColor: TvgrColorButton;
    bLeftFont: TSpeedButton;
    Bevel1: TBevel;
    bCenterFont: TSpeedButton;
    bRightFont: TSpeedButton;
    bPage: TSpeedButton;
    bPageCount: TSpeedButton;
    bDate: TSpeedButton;
    bTime: TSpeedButton;
    bTab: TSpeedButton;
    FontDialog: TFontDialog;
    vgrFormLocalizer1: TvgrFormLocalizer;
    procedure bLeftFontClick(Sender: TObject);
    procedure bCenterFontClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure bRightFontClick(Sender: TObject);
    procedure bPageClick(Sender: TObject);
  private
    FHeaderFooter: TvgrPageHeaderFooter;
    procedure EditFont(AFont: TFont);
  public
{Opens the dialog.
Parameters:
  AHeaderFooter - Specifies a TvgrPageHeaderFooter object, which should be edited.
Return value:
  Returns true when the user edit properties and clicks OK, or false when the user cancels.
See also:
  TvgrPageHeaderFooter}
    function Execute(AHeaderFooter: TvgrPageHeaderFooter): Boolean;
  end;

implementation

uses Math;

{$R *.dfm}
{$R ..\res\vgr_HeaderFooterOptionsDialogStrings.res}

procedure TvgrHeaderFooterOptionsDialogForm.FormCreate(Sender: TObject);
begin
  FHeaderFooter := TvgrPageHeaderFooter.Create;
end;

procedure TvgrHeaderFooterOptionsDialogForm.FormDestroy(Sender: TObject);
begin
  FHeaderFooter.Free;
end;

function TvgrHeaderFooterOptionsDialogForm.Execute(AHeaderFooter: TvgrPageHeaderFooter): Boolean;
begin
  FHeaderFooter.Assign(AHeaderFooter);
  mLeft.Lines.Text := FHeaderFooter.LeftSection;
  mCenter.Lines.Text := FHeaderFooter.CenterSection;
  mRight.Lines.Text := FHeaderFooter.RightSection;
  bBackColor.SelectedColor := FHeaderFooter.BackColor;

  Result := ShowModal = mrOk;
  if Result then
  begin
    with FHeaderFooter do
    begin
      BackColor := bBackColor.SelectedColor;
      LeftSection := mLeft.Lines.Text;
      CenterSection := mCenter.Lines.Text;
      RightSection := mRight.Lines.Text;
    end;
    AHeaderFooter.Assign(FHeaderFooter);
  end;
end;

procedure TvgrHeaderFooterOptionsDialogForm.EditFont(AFont: TFont);
begin
  FontDialog.Font.Assign(AFont);
  if FontDialog.Execute then
    AFont.Assign(FontDialog.Font);
end;

procedure TvgrHeaderFooterOptionsDialogForm.bLeftFontClick(
  Sender: TObject);
begin
  EditFont(FHeaderFooter.LeftFont);
end;

procedure TvgrHeaderFooterOptionsDialogForm.bCenterFontClick(
  Sender: TObject);
begin
  EditFont(FHeaderFooter.CenterFont);
end;

procedure TvgrHeaderFooterOptionsDialogForm.bRightFontClick(
  Sender: TObject);
begin
  EditFont(FHeaderFooter.RightFont);
end;

procedure TvgrHeaderFooterOptionsDialogForm.bPageClick(Sender: TObject);
var
  S: string;
  M: TMemo;
begin
  if ActiveControl = mLeft then
    M := mLeft
  else
    if ActiveControl = mCenter then
      M := mCenter
    else
      if ActiveControl = mRight then
        M := mRight
      else
        exit;
        
  if Sender = bPage then
    S := cHFKeyPage
  else
    if Sender = bPageCount then
      S := cHFKeyPages
    else
      if Sender = bDate then
        S := cHFKeyDate
      else
        if Sender = bTime then
          S := cHFKeyTime
        else
          if Sender = bTab then
            S := cHFKeyTab
          else
            exit;

  SendMessage(M.Handle, EM_REPLACESEL, 0, integer(PChar(S)));
end;

end.

