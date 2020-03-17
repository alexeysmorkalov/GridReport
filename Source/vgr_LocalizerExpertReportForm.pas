unit vgr_LocalizerExpertReportForm;

{$I vtk.inc}

interface

uses
//  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, ImgList, Controls, StdCtrls, ExtCtrls,
  ComCtrls, Classes, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF} Graphics, Forms,
  Dialogs, 

  vgr_CommonClasses, vgr_Form, vgr_LocalizeExpertForm;

const
  sStateNotChecked = 1;
  sStateChecked = 2;

type
  TvgrLocalizerExpertReportForm = class(TvgrDialogForm)
    PageControl1: TPageControl;
    Panel1: TPanel;
    PResourceFile: TTabSheet;
    PFormLocalizer: TTabSheet;
    bOK: TButton;
    bCancel: TButton;
    Label1: TLabel;
    Label2: TLabel;
    LVResource: TListView;
    LVFormLocalizer: TListView;
    ImageList: TImageList;
    procedure LVResourceMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
    FReport: TvgrLocalizerBuildReport;
  public
    { Public declarations }
    function Execute(AReport: TvgrLocalizerBuildReport): Boolean;

    property Report: TvgrLocalizerBuildReport read FReport;
  end;

implementation

{$R *.dfm}

function TvgrLocalizerExpertReportForm.Execute(AReport: TvgrLocalizerBuildReport): Boolean;

  procedure CopyToList(ADest: TvgrObjectList; ASource: TListView);
  var
    I: Integer;
  begin
    for I := 0 to ADest.Count - 1 do
      TvgrLocalizerDeleteItem(ADest[I]).Delete := ASource.Items[I].StateIndex = sStateChecked;
  end;

  procedure CopyFromList(ADest: TListView; ASource: TvgrObjectList);
  const
    aState: array [Boolean] of Integer = (sStateNotChecked, sStateChecked);
  var
    I: Integer;
  begin
    ADest.Items.BeginUpdate;
    try
      ADest.Items.Clear;
      for I := 0 to ASource.Count - 1 do
        with ADest.Items.Add, TvgrLocalizerDeleteItem(ASource[I]) do
        begin
          Caption := Description;
          StateIndex := aState[Delete];
          SubItems.Add(DeleteReason);
        end;
      if ADest.Items.Count > 0 then
      begin
        ADest.ItemFocused := ADest.Items[0];
        ADest.ItemFocused.Selected := True;
      end;
    finally
      ADest.Items.EndUpdate;
    end;
  end;

begin
  FReport := AReport;
  CopyFromList(LVResource, Report.RCList);
  CopyFromList(LVFormLocalizer, Report.FLList);

  Result := ShowModal = mrOk;
  if Result then
  begin
    CopyToList(Report.RCList, LVResource);
    CopyToList(Report.FLList, LVFormLocalizer);
  end;
end;

procedure TvgrLocalizerExpertReportForm.LVResourceMouseDown(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
var
  AItem: TListItem;
begin
  if (Button = mbLeft) and (Shift = [ssLeft]) and (htOnStateIcon in TListView(Sender).GetHitTestInfoAt(X, Y)) then
  begin
    AItem := TListView(Sender).GetItemAt(X, Y);
    if AItem <> nil then
    begin
      if AItem.StateIndex = sStateChecked then
        AItem.StateIndex := sStateNotChecked
      else
        AItem.StateIndex := sStateChecked;
    end;
  end; 
end;

end.
