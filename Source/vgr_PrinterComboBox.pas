{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains TvgrPrinterComboBox visaul component.
TvgrPrinterComboBox represents a combo box that lets users select a printer.
You can gets or sets selected printer with using SelectedPrinterName property.
Use PrintersInfo property, to get list of installed printers.
See also:
  TvgrPrinterComboBox, TvgrPrinters, TvgrPrinter}
unit vgr_PrinterComboBox;

interface

{$I vtk.inc}

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  SysUtils, Messages, Windows, Classes, StdCtrls, graphics, Controls, math,

  vgr_Printer, vgr_Functions;

type
  /////////////////////////////////////////////////
  //
  // TvgrPrinterComboBox
  //
  /////////////////////////////////////////////////
  {TvgrPrinterComboBox represents a combo box that lets users select a printer.
You can gets or sets selected printer with using SelectedPrinterName property.
Use PrintersInfo property, to get list of installed printers.
Use SetToDefaultPrinter property to select default printer.
List of installed printers represents by the TvgrPrinters class.
See also:
  TvgrPrinter, TvgrPrinters}
  TvgrPrinterComboBox = class(TCustomComboBox)
  private
    FItemHeight: Integer;
    FAutoSelectDefaultPrinter: Boolean;
    FSelectedPrinterName: string;
    procedure UpdateItems(AUpdatePrinters: Boolean);
    function GetSelectedPrinter: TvgrPrinter;
    function GetSelectedPrinterName: string;
    procedure SetSelectedPrinterName(const Value: string);
    procedure DrawPrinterItem(ADC: HDC; APrinter: TvgrPrinter; const ARect: TRect; ASelected, ADisabled: Boolean);
    procedure SetAutoSelectDefaultPrinter(Value: Boolean);
  protected
    FIsFocused: Boolean;
    FDroppingDown: Boolean;
    FFocusChanged: Boolean;
    procedure AdjustDropDown; {$IFDEF VTK_D6_OR_D7} override; {$ENDIF}
    procedure WMEraseBkgnd(var Msg: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMPaint(var Msg: TWMPaint); message WM_PAINT;
    procedure WMDrawItem(var Msg: TWMDrawItem); message WM_DRAWITEM;
    procedure CBGetItemHeight(var Msg: TMessage); message CB_GETITEMHEIGHT;
    procedure CNCommand(var Msg: TWMCommand); message CN_COMMAND;
    procedure CreateWnd; override;
    procedure DoEnter; override;
    procedure DoExit; override;
  public
{Creates an instance of TvgrPrinterComboBox.
Parameters:
  AOwner - owner component.}
    constructor Create(AOwner: TComponent); override;
{Frees instance of TvgrPrinterComboBox.}
    destructor Destroy; override;

{Changes SelectedPrinterName property to default printer.}
    procedure SetToDefaultPrinter;

{Gets and sets name of the default printer.
See also:
  SelectedPrinter}
    property SelectedPrinterName: string read GetSelectedPrinterName write SetSelectedPrinterName;
{Gets currently selected printer as TvgrPrinter object.
See also:
  SelectedPrinterName}
    property SelectedPrinter: TvgrPrinter read GetSelectedPrinter;
  published
{Gets or sets a value indicating whether after creation of the component
SelectedPrinterName property setup to default printer.
See also:
  PrintersInfo, SelectedPrinterName, SelectedPrinter, SetToDefaultPrinter}
    property AutoSelectDefaultPrinter: Boolean read FAutoSelectDefaultPrinter write SetAutoSelectDefaultPrinter default True;
    property Anchors;
    property BiDiMode;
    property Color;
    property Constraints;
    property Ctl3D;
    property DragCursor;
    property DragKind;
    property DragMode;
    property DropDownCount;
    property Enabled;
    property Font;
    property ImeMode;
    property ImeName;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ShowHint;
    property Sorted;
    property TabOrder;
    property TabStop;
    property Visible;
    property OnChange;
    property OnClick;
{$IFDEF VTK_D5}
    property OnContextPopup;
{$ENDIF}
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnDropDown;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnStartDock;
    property OnStartDrag;
  end;

implementation

{$R ..\res\vgr_PrinterComboBox.res}

uses
  vgr_StringIDs, vgr_Localize;
  
const
  cPrinterImageWidth = 16;
  cPrinterImageHeight = 16;

/////////////////////////////////////////////////
//
// TvgrPrinterComboBox
//
/////////////////////////////////////////////////
constructor TvgrPrinterComboBox.Create(AOwner: TComponent);
begin
  inherited;
  Style := csOwnerDrawFixed;
  FAutoSelectDefaultPrinter := True;
end;

destructor TvgrPrinterComboBox.Destroy;
begin
  inherited;
end;

procedure TvgrPrinterComboBox.UpdateItems(AUpdatePrinters: Boolean);
var
  I: Integer;
  AMaxSize, ASize: TSize;
  ACanvas: TCanvas;
  ADC: HDC;
begin
  ADC := GetDC(0);
  ACanvas := TCanvas.Create;
  ACanvas.Handle := ADC;
  try
    if AUpdatePrinters then
      Printers.Update;

    Items.BeginUpdate;
    try
      Items.Clear;
      ACanvas.Font.Name := Font.Name;
      ACanvas.Font.Size := Font.Size;
      AMaxSize.cx := Width;
      AMaxSize.cy := ACanvas.TextHeight('Wg');
      for I := 0 to Printers.Count - 1 do
      begin
        Items.AddObject(Printers[I].PrinterName, Printers[I]);
        ASize := ACanvas.TextExtent(Printers[I].PrinterName);
        if ASize.cx > AMaxSize.cx then
          AMaxSize.cx := ASize.cx;
        if ASize.cy > AMaxSize.cy then
          AMaxSize.cy := ASize.cy;
      end;
      FItemHeight := Max(AMaxSize.cy, cPrinterImageHeight) + 2;
      SendMessage(Handle, CB_SETITEMHEIGHT, 0, FItemHeight);
      SendMessage(Handle, CB_SETDROPPEDWIDTH, AMaxSize.cx + cPrinterImageWidth, 0);

      //
      if AutoSelectDefaultPrinter then
        SetToDefaultPrinter;
    finally
      Items.EndUpdate;
    end;
  finally
    ReleaseDC(0, ADC);
    ACanvas.Free;
  end;
end;

procedure TvgrPrinterComboBox.SetToDefaultPrinter;
begin
  ItemIndex := Items.IndexOfObject(Printers.DefaultPrinter)
end;

function TvgrPrinterComboBox.GetSelectedPrinter: TvgrPrinter;
begin
  if (ItemIndex >= 0) and (ItemIndex < Items.Count) then
    Result := Printers[ItemIndex]
  else
    Result := nil;
end;

function TvgrPrinterComboBox.GetSelectedPrinterName: string;
begin
  if SelectedPrinter = nil then
    Result := ''
  else
    Result := SelectedPrinter.PrinterName;
end;

procedure TvgrPrinterComboBox.SetSelectedPrinterName(const Value: string);
begin
  if HandleAllocated then
    ItemIndex := Printers.IndexOfPrinterName(Value)
  else
    FSelectedPrinterName := Value;
end;

procedure TvgrPrinterComboBox.SetAutoSelectDefaultPrinter(Value: Boolean);
begin
  if FAutoSelectDefaultPrinter <> Value then
  begin
    FAutoSelectDefaultPrinter := Value;
    if FAutoSelectDefaultPrinter then
      SetToDefaultPrinter;
  end;
end;

procedure TvgrPrinterComboBox.CBGetItemHeight(var Msg: TMessage);
begin
  inherited;
end;

procedure TvgrPrinterComboBox.DrawPrinterItem(ADC: HDC; APrinter: TvgrPrinter; const ARect: TRect; ASelected, ADisabled: Boolean);
var
  S: string;
  ACanvas: TCanvas;
  AImage: TBitmap;
  ASize: TSize;
begin
  if APrinter = nil then
  begin
    S := '';
    AImage := nil;
  end
  else
  begin
    S := APrinter.PrinterName;
    AImage := APrinter.Image;
  end;

  ACanvas := TCanvas.Create;
  try
    ACanvas.Handle := ADC;
    ACanvas.Font.Size := Font.Size;
    ACanvas.Font.Name := Font.Name;

    if ASelected then
    begin
      ACanvas.Font.Color := clHighlightText;
      ACanvas.Brush.Color := clHighlight;
    end
    else
    begin
      ACanvas.Brush.Color := clWindow;
      if ADisabled then
        ACanvas.Font.Color := clGrayText
      else
        ACanvas.Font.Color := clWindowText;
    end;

    ACanvas.Brush.Style := bsSolid;
    ACanvas.FillRect(ARect);
    ACanvas.Brush.Style := bsClear;

    ASize := ACanvas.TextExtent(S);
    ACanvas.TextRect(ARect,
                     ARect.Left + cPrinterImageWidth + 2 + 2,
                     ARect.Top + (ARect.Bottom - ARect.Top - ASize.cy) div 2,
                     S);

    if AImage <> nil then
      ACanvas.Draw(ARect.Left,
                   ARect.Top + (ARect.Bottom - ARect.Top - cPrinterImageHeight) div 2,
                   AImage);
  finally
    ACanvas.Free;
  end;
end;

procedure TvgrPrinterComboBox.WMEraseBkgnd(var Msg: TWMEraseBkgnd);
begin
  Msg.Result := 1;
end;

procedure TvgrPrinterComboBox.WMPaint(var Msg: TWMPaint);
var
  ARect: TRect;
  AButtonRect: TRect;
  ACanvas: TControlCanvas;
  PS: TPaintStruct;
  S: string;
  ASize: TSize;
begin
  BeginPaint(Handle, PS);

  ARect := ClientRect;
  ACanvas := TControlCanvas.Create;
  try
    ACanvas.Control := Self;
    DrawEdge(ACanvas.Handle, ARect, EDGE_SUNKEN, BF_ADJUST or BF_RECT);

    AButtonRect := ARect;
    AButtonRect.Left := AButtonRect.Right - GetSystemMetrics(SM_CXHTHUMB) - 1;
    if DroppedDown then
      DrawFrameControl(ACanvas.Handle, AButtonRect, DFC_SCROLL, DFCS_FLAT or DFCS_SCROLLCOMBOBOX)
    else
      DrawFrameControl(ACanvas.Handle, AButtonRect, DFC_SCROLL, DFCS_SCROLLCOMBOBOX);

    ARect.Right := AButtonRect.Left;
    ACanvas.Brush.Color := clWindow;
    ACanvas.FrameRect(ARect);
    InflateRect(ARect, -1, -1);
    if FIsFocused then
    begin
      ACanvas.Brush.Color := clWindowFrame;
      ACanvas.FillRect(ARect);
      InflateRect(ARect, -1, -1);
    end;
    if SelectedPrinter = nil then
    begin
      if Printers.Count > 0 then
      begin
        ACanvas.Brush.Color := clWindow;
        ACanvas.FillRect(ARect);
      end
      else
      begin
        ACanvas.Brush.Color := clBtnFace;
        ACanvas.FillRect(ARect);

        ACanvas.Font.Color := clBtnText;
        S := vgrLoadStr(svgrid__vgr_PrinterComboBox__NoInstalledPrinters);
        ASize := ACanvas.TextExtent(S);
        ACanvas.TextRect(ARect,
                         ARect.Left + 2 + 2,
                         ARect.Top + (ARect.Bottom - ARect.Top - ASize.cy) div 2,
                         S);

      end;
    end
    else
      DrawPrinterItem(ACanvas.Handle, SelectedPrinter, ARect, FIsFocused, not Enabled);
  finally
    ACanvas.Free;
  end;

  EndPaint(Handle, PS);
  
  Msg.Result := 1;
end;

procedure TvgrPrinterComboBox.WMDrawItem(var Msg: TWMDrawItem);
var
  APrinter: TvgrPrinter;
  ADisabled, ASelected: Boolean;
begin
  Msg.Result := 1;

  with Msg.DrawItemStruct^ do
  begin
    if ((itemState and ODS_COMBOBOXEDIT) <> 0) or (ItemID = $FFFFFFFF) then
    begin
      if ItemIndex = -1 then
        APrinter := nil
      else
        APrinter := Printers[ItemIndex];
      ADisabled := (itemState and ODS_DISABLED) <> 0;
      ASelected := False;
    end
    else
    begin
      APrinter := Printers[ItemID];
      ASelected := (itemState and ODS_SELECTED) <> 0;
      ADisabled := False;
    end;

    DrawPrinterItem(hDC, APrinter, rcItem, ASelected, ADisabled);
  end;
end;

procedure TvgrPrinterComboBox.CreateWnd;
begin
  inherited;
  UpdateItems(True);
  if (FSelectedPrinterName <> '') and (Printers.IndexOfPrinterName(FSelectedPrinterName) <> -1) then
  begin
    SetSelectedPrinterName(FSelectedPrinterName);
    FSelectedPrinterName := '';
  end;
end;

procedure TvgrPrinterComboBox.DoEnter;
begin
  Invalidate;
  inherited;
end;

procedure TvgrPrinterComboBox.DoExit;
begin
  Invalidate;
  inherited;
end;

procedure TvgrPrinterComboBox.AdjustDropDown;
var
  ItemCount: Integer;
begin
  ItemCount := Items.Count;
  if ItemCount > DropDownCount then ItemCount := DropDownCount;
  if ItemCount < 1 then ItemCount := 1;
  FDroppingDown := True;
  try
    SetWindowPos(Handle, 0, 0, 0, Width, FItemHeight * ItemCount +
      Height + 2, SWP_NOMOVE + SWP_NOZORDER + SWP_NOACTIVATE + SWP_NOREDRAW +
      SWP_HIDEWINDOW);
  finally
    FDroppingDown := False;
  end;
  SetWindowPos(Handle, 0, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE +
    SWP_NOZORDER + SWP_NOACTIVATE + SWP_NOREDRAW + SWP_SHOWWINDOW);
end;

procedure TvgrPrinterComboBox.CNCommand(var Msg: TWMCommand);
begin
  case Msg.NotifyCode of
    CBN_DROPDOWN:
      begin
        FFocusChanged := False;
        DropDown;
        AdjustDropDown;
        if FFocusChanged then
        begin
          PostMessage(Handle, WM_CANCELMODE, 0, 0);
          if not FIsFocused then
            PostMessage(Handle, CB_SHOWDROPDOWN, 0, 0);
        end;
      end;
    CBN_CLOSEUP:
      begin
        Invalidate;
        inherited;
      end;
    CBN_SETFOCUS:
      begin
        FIsFocused := True;
        FFocusChanged := True;
        SetIme;
      end;
    CBN_KILLFOCUS:
      begin
        FIsFocused := False;
        FFocusChanged := True;
        ResetIme;
      end;
  else
    inherited;
  end;
end;

end.
