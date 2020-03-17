{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2004 by vtkTools      }
{                                          }
{******************************************}

unit vgr_ReportAdd;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}
  Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, ExtCtrls, shellapi;

const
  cSupportURL = 'mailto:support_gridreport@vtktools.com';
  cSalesURL = 'http://www.vtktools.com/bay/bay.html';
  cMainURL = 'http://www.vtktools.com';

type
  /////////////////////////////////////////////////
  //
  // TvgrImageButton
  //
  /////////////////////////////////////////////////
  TvgrImageButton = class(TCustomPanel)
  private
    FNormalImage: TImage;
    FHighlightImage: TImage;
    FUnderMouse: Boolean;

    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CMMouseEnter(var Msg: TMessage); message CM_MOUSEENTER;
    procedure CMMouseLeave(var Msg: TMessage); message CM_MOUSELEAVE;
    function GetActiveImage: TImage;
    procedure UpdateButton;
    procedure SetNormalImage(Value: TImage);
    procedure SetHighlightImage(Value: TImage);
  protected   
    procedure Paint; override;
    procedure Click; override;
    procedure Notification(AComponent: TComponent; AOperation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    property ActiveImage: TImage read GetActiveImage;

    property NormalImage: TImage read FNormalImage write SetNormalImage;
    property HighlightImage: TImage read FHighlightImage write SetHighlightImage;

    property OnClick;
  end;
  
  /////////////////////////////////////////////////
  //
  // TvgrReportAddForm
  //
  /////////////////////////////////////////////////
  TvgrReportAddForm = class(TForm)
    Image1: TImage;
    imgSupportN: TImage;
    imgSalesN: TImage;
    imgCloseN: TImage;
    imgSupportH: TImage;
    imgSalesH: TImage;
    imgCloseH: TImage;
    Image2: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Image1Click(Sender: TObject);
  private
    { Private declarations }
    FSupport: TvgrImageButton;
    FSales: TvgrImageButton;
    FClose: TvgrImageButton;

    procedure OnCloseClick(Sender: TObject);
    procedure OnSupportClick(Sender: TObject);
    procedure OnSalesClick(Sender: TObject);
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    function Execute: Boolean;
  end;

implementation

{$R *.dfm}

function IsDesignTime: Boolean;
var
  AExeName: string;
begin
  AExeName := ExtractFileName(ParamStr(0));
{$IFDEF BCB}
  Result := AnsiLowerCase(AExeName) = 'bcb.exe';
{$ELSE}
  Result := AnsiLowerCase(AExeName) = 'delphi32.exe';
{$ENDIF}
end;

/////////////////////////////////////////////////
//
// TvgrImageButton
//
/////////////////////////////////////////////////
constructor TvgrImageButton.Create(AOwner: TComponent);
begin
  inherited;
  Cursor := crHandPoint;
end;

procedure TvgrImageButton.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrImageButton.CMMouseEnter(var Msg: TMessage);
begin
  FUnderMouse := True;
  UpdateButton;
end;

procedure TvgrImageButton.CMMouseLeave(var Msg: TMessage);
begin
  FUnderMouse := False;
  UpdateButton;
end;

function TvgrImageButton.GetActiveImage: TImage;
begin
  if FUnderMouse then
    Result := FHighlightImage
  else
    Result := FNormalImage;
end;

procedure TvgrImageButton.UpdateButton;
begin
  if ActiveImage <> nil then
    SetBounds(Left, Top, ActiveImage.Width, ActiveImage.Height);
  Repaint;
end;

procedure TvgrImageButton.SetNormalImage(Value: TImage);
begin
  if FNormalImage <> Value then
  begin
    FNormalImage := Value;
    UpdateButton;
  end;
end;

procedure TvgrImageButton.SetHighlightImage(Value: TImage);
begin
  if FHighlightImage <> Value then
  begin
    FHighlightImage := Value;
    UpdateButton;
  end;
end;

procedure TvgrImageButton.Paint;
begin
  if ActiveImage <> nil then
    Canvas.Draw(0, 0, ActiveImage.Picture.Graphic)
  else
  begin
    Canvas.Brush.Color := clBtnFace;
    Canvas.FillRect(ClientRect);
  end;
end;

procedure TvgrImageButton.Click;
begin
  if Assigned(OnClick) then
    OnClick(Self);
end;

procedure TvgrImageButton.Notification(AComponent: TComponent; AOperation: TOperation);
begin
  inherited;
  if AOperation = opRemove then
  begin
    if AComponent = FNormalImage then
      NormalImage := nil;
    if AComponent = FHighlightImage then
      HighlightImage := nil;
  end;
end;

/////////////////////////////////////////////////
//
// TvgrReportAddForm
//
/////////////////////////////////////////////////
constructor TvgrReportAddForm.Create(AOwner: TComponent);
begin
  inherited;
  FSupport := TvgrImageButton.Create(Self);
  with FSupport do
  begin
    NormalImage := imgSupportN;
    HighlightImage := imgSupportH;
    OnClick := OnSupportClick;
    Left := 96;
    Top := 154;
    Parent := Self;
  end;

  FSales := TvgrImageButton.Create(Self);
  with FSales do
  begin
    NormalImage := imgSalesN;
    HighlightImage := imgSalesH;
    OnClick := OnSalesClick;
    Left := 96;
    Top := 171;
    Parent := Self;
  end;

  FClose := TvgrImageButton.Create(Self);
  with FClose do
  begin
    NormalImage := imgCloseN;
    HighlightImage := imgCloseH;
    OnClick := OnCloseClick;
    Left := 328;
    Top := 168;
    Parent := Self;
  end;

  if IsDesignTime then
  begin
    Image2.Visible := True;
    Image1.Visible := False;
  end
  else
  begin
    Image2.Visible := False;
    Image1.Visible := True;
  end;
end;

function TvgrReportAddForm.Execute: Boolean;
begin
  ShowModal;
  Result := True;
end;

procedure TvgrReportAddForm.OnCloseClick(Sender: TObject);
begin
  OnClose := nil;
  ModalResult := mrOk;
end;

procedure TvgrReportAddForm.OnSupportClick(Sender: TObject);
begin
  ShellExecute(0, nil, cSupportURL, nil, nil, SW_MAXIMIZE);
end;

procedure TvgrReportAddForm.OnSalesClick(Sender: TObject);
begin
  ShellExecute(0, nil, cSalesURL, nil, nil, SW_MAXIMIZE);
end;

procedure TvgrReportAddForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caNone;
end;

procedure TvgrReportAddForm.Image1Click(Sender: TObject);
begin
  ShellExecute(0, nil, cMainURL, nil, nil, SW_MAXIMIZE);
end;

initialization

  with TvgrReportAddForm.Create(nil) do
  begin
    Execute;
    Free;
  end;

end.

