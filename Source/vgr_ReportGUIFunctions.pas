{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains the common GUI functions, used by GridReport (the text wrapping, the text measuring and so on).}
unit vgr_ReportGUIFunctions;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, Graphics, Windows, Math,

  vgr_Functions, vgr_GUIFunctions, vgr_DataStorageTypes, vgr_DataStorageRecords;

{Splits a string into multiple lines as its size approaches a specified rectangle.
Parameters:
  AText - text to wrap.
  ACanvas - TCanvas object, that is used to calculate sizes of the text.
  ATextRect - the rectangle, that specifies bounds of the text.
  AAngle - the rotation angle of the text.
Return value:
  Returns the wrapped text, lines is divided by CR and LF chars.}
  function vgrWrapRotatedText(const AText: string;
                              ACanvas: TCanvas;
                              ATextRect: TRect;
                              AAngle: Integer): string;
{The vgrDrawText procedure draws text in the specified rectangle.
Parameters:
  ACanvas - TCanvas object.
  AText - the string that specifies the text to be drawn.
  ARect - the TRect structure that contains the rectangle
(in logical coordinates) in which the text is to be formatted.
  AWordWrap - if true then breaks words. Lines are automatically broken between
words if a word would extend past the edge of the rectangle specified by the
ARect parameter. A carriage return-line feed sequence also breaks the line.
  AHorzAlign - the horizontally alignment of the text.
  AVertAlign - the vertically alignment of the text.
  ARotateAnlge - the rotation angle of the text.}
  procedure vgrDrawText(ACanvas: TCanvas;
                        const AText: string;
                        const ARect: TRect;
                        AWordWrap: Boolean;
                        AHorzAlign: TvgrRangeHorzAlign;
                        AVertAlign: TvgrRangeVertAlign;
                        ARotateAngle: Integer);
{The vgrCalcText function computes the width and height of the specified string of text.
The text can be rotated on any angle.
Parameters:
  ACanvas - TCanvas object, that is used to calculation of the text sizes.
  AText - the string that specifies the text to be calculated.
  ARect - the TRect structure that contains the rectangle
(in logical coordinates) in which the text is to be formatted.
  AWordWrap - if true then breaks words. Lines are automatically broken between
words if a word would extend past the edge of the rectangle specified by the
ARect parameter. A carriage return-line feed sequence also breaks the line.
  AHorzAlign - the horizontally alignment of the text.
  AVertAlign - the vertically alignment of the text.
  ARotateAnlge - the rotation angle of the text.
Return value:
  Returns the text sizes.}
  function vgrCalcText(ACanvas: TCanvas;
                       const AText: string;
                       const ARect: TRect;
                       AWordWrap: Boolean;
                       AHorzAlign: TvgrRangeHorzAlign;
                       AVertAlign: TvgrRangeVertAlign;
                       ARotateAngle: Integer): TSize;

{Creates the GDI pen object, that can used to drawing the border of specified style.
Syntax:
  function CreateBorderPen(ASize: Integer; AStyle: pvgrBorderStyle): HPEN;
See also:
  ASize - ignored.
  AStyle - specifies style of the pen (color, pattern and so on).
Return value:
  Returns handle of the created GDI pen.}
  function CreateBorderPen(ASize: Integer; AStyle: pvgrBorderStyle): HPEN;

{The DrawRotatedText procedure draws text in the specified rectangle.
Text can not be wrapped. 
Syntax:
  procedure vgrDrawText(ACanvas: TCanvas; const AText: string; const ARect: TRect; AWordWrap: Boolean; AHorzAlign: TvgrRangeHorzAlign; AVertAlign: TvgrRangeVertAlign; ARotateAngle: Integer);
Parameters:
  ACanvas - TCanvas object.
  AText - the string that specifies the text to be drawn.
  ARect - the TRect structure that contains the rectangle
(in logical coordinates) in which the text is to be formatted.
  ARotateAnlge - the rotation angle of the text.
  ABackTransparent - specifies should be background filled with current background color or not.
  AHorizontalAlignment - the horizontally alignment of the text.
  AVerticalAlignment - the vertically alignment of the text.
See also:
  vgrDrawText}
  procedure DrawRotatedText(ACanvas: TCanvas;
                            const AText: string;
                            const ARect: TRect;
                            ARotateAngle: Integer;
                            ABackTransparent: Boolean;
                            AHorizontalAlignment: TvgrRangeHorzAlign;
                            AVerticalAlignment: TvgrRangeVertAlign);

var
  APens: array [TvgrBorderStyle] of Integer = (PS_SOLID, PS_DASH, PS_DOT, PS_DASHDOT, PS_DASHDOTDOT);

implementation

procedure InternalCalcRotatedText(ACanvas: TCanvas;
                                  AText: string;
                                  ARotateAngle: Integer;
                                  AHorizontalAlignment: TvgrRangeHorzAlign;
                                  AVerticalAlignment: TvgrRangeVertAlign;
                                  ASelectRotatedFont: Boolean;
                                  ARestoreFont: Boolean;
                                  var ATextAlign: TvgrTextHorzAlign;
                                  var AStepX: Integer;
                                  var AStepY: Integer;
                                  var ATextSize: TSize;
                                  var ARotatedTextSize: TSize;
                                  var APos: TPoint;
                                  var AHalfLineSize: Integer;
                                  var ANewFont: HFONT;
                                  var AOldFont: HFONT);
const
  AHorzAlignConvert: array [TvgrRangeHorzAlign] of TvgrTextHorzAlign = (vgrthaLeft, vgrthaLeft, vgrthaCenter, vgrthaRight);
  AVertAlignConvert: array [TvgrRangeVertAlign] of TvgrTextHorzAlign = (vgrthaLeft, vgrthaCenter, vgrthaRight);
var
  ALineCount: Integer;
  ATextMetrics: TEXTMETRIC;
  ACos, ASin, ARadAngle: Double;
begin
  if ASelectRotatedFont then
  begin
    ANewFont := GetRotatedFontHandle(ACanvas.Font, ARotateAngle);
    AOldFont := SelectObject(ACanvas.Handle, ANewFont);
  end
  else
  begin
    AOldFont := 0;
    ANewFont := 0;
  end;

  ATextSize := vgrCalcStringSize(ACanvas, AText, ALineCount);
  ARotateAngle := ARotateAngle mod 360;
  GetTextMetrics(ACanvas.Handle, ATextMetrics);
  AHalfLineSize := ATextMetrics.tmHeight div 2;

  case Abs(ARotateAngle) of
    0, 180:
      begin
        ATextAlign := AHorzAlignConvert[AHorizontalAlignment];
        AStepX := 0;
        AStepY := ATextMetrics.tmHeight;

        ARotatedTextSize.cx := ATextSize.cx;
        ARotatedTextSize.cy := ATextSize.cy;

        APos.X := ATextSize.cx div 2;
        APos.Y := AHalfLineSize;
      end;
    90, 270:
      begin
        ATextAlign := AVertAlignConvert[AVerticalAlignment];
        AStepX := ATextMetrics.tmHeight;
        AStepY := 0;

        ARotatedTextSize.cx := ATextSize.cy;
        ARotatedTextSize.cy := ATextSize.cx;

        APos.X := AHalfLineSize;
        APos.Y := ATextSize.cx div 2;
      end;
    else
      begin
        ARadAngle := ToRadians(ARotateAngle);
        ACos := Cos(ARadAngle);
        ASin := Sin(ARadAngle);
        if Abs(ARotateAngle mod 90) < 45 then
        begin
          // by Y
          ATextAlign := AHorzAlignConvert[AHorizontalAlignment];
          AStepX := 0;
          AStepY := ATextMetrics.tmHeight;

          ARotatedTextSize.cx := Round(Abs(ATextSize.cx * ACos) + Abs(ATextMetrics.tmHeight * ASin));
          ARotatedTextSize.cy := Round(Abs(ATextSize.cx * ASin) + Abs(ATextMetrics.tmHeight{ATextMetrics.tmHeight / ACos} * ALineCount));

          APos.X := ARotatedTextSize.cx div 2;
          APos.Y := Round((Abs(ATextMetrics.tmHeight * ACos) + Abs(ATextSize.cx * ASin)) / 2);
        end
        else
        begin
          // by X
          ATextAlign := AVertAlignConvert[AVerticalAlignment];
          AStepX := ATextMetrics.tmHeight;
          AStepY := 0;

          ARotatedTextSize.cx := Round(Abs(ATextSize.cx * ACos) + Abs(ATextMetrics.tmHeight{ATextMetrics.tmHeight / ASin} * ALineCount));
          ARotatedTextSize.cy := Round(Abs(ATextSize.cx * ASin) + Abs(ATextMetrics.tmHeight * ACos));

          APos.X := Round((Abs(ATextMetrics.tmHeight * ASin) + Abs(ATextSize.cx * ACos)) / 2);
          APos.Y := ARotatedTextSize.cy div 2;
        end;
      end;
  end;

  if ASelectRotatedFont and ARestoreFont then
  begin
    SelectObject(ACanvas.Handle, AOldFont);
    DeleteObject(ANewFont);
    AOldFont := 0;
    ANewFont := 0;
  end;
end;

procedure DrawRotatedText(ACanvas: TCanvas;
                          const AText: string;
                          const ARect: TRect;
                          ARotateAngle: Integer;
                          ABackTransparent: Boolean;
                          AHorizontalAlignment: TvgrRangeHorzAlign;
                          AVerticalAlignment: TvgrRangeVertAlign);
var
  ALine: string;
  ALineInfo: TvgrTextLineInfo;
  AOldTransparent: Boolean;
  ARotatedTextSize, ATextSize, ALineSize: TSize;
  APos, ALinePos: TPoint;
  ATextHorzAlign: TvgrTextHorzAlign;
  ANewFont, AOldFont: HFONT;
  AHalfLineSize, AStepX, AStepY, ARectWidth, ARectHeight: Integer;
begin
  ARectWidth := RectWidth(ARect);
  ARectHeight := RectHeight(ARect);
  if (ARectWidth = 0) or (ARectHeight = 0) then exit;

  AOldTransparent := SetCanvasTransparentMode(ACanvas, ABackTransparent);

  InternalCalcRotatedText(ACanvas,
                          AText,
                          ARotateAngle,
                          AHorizontalAlignment,
                          AVerticalAlignment,
                          True,
                          False,
                          ATextHorzAlign,
                          AStepX,
                          AStepY,
                          ATextSize,
                          ARotatedTextSize,
                          APos,
                          AHalfLineSize,
                          ANewFont,
                          AOldFont);

  APos.X := APos.X + ARect.Left;
  APos.Y := APos.Y + ARect.Top;

  case AVerticalAlignment of
    vgrvaCenter: APos.Y := APos.Y + (ARectHeight - ARotatedTextSize.cy) div 2;
    vgrvaBottom: APos.Y := APos.Y + ARectHeight - ARotatedTextSize.cy;
    else ; // APos.Y - not changed
  end;

  case AHorizontalAlignment of
    vgrhaCenter: APos.X := APos.X + (ARectWidth - ARotatedTextSize.cx) div 2;
    vgrhaRight: APos.X := APos.X + ARectWidth - ARotatedTextSize.cx;
    else ; // APos.X - not changed
  end;

  BeginEnumLines(AText, ALineInfo);
  while EnumLines(AText, ALineInfo) do
  begin
    ALine := Copy(AText, ALineInfo.Start, ALineInfo.Length);

    GetTextExtentPoint32(ACanvas.Handle, PChar(ALine), ALineInfo.Length, ALineSize);
    case ATextHorzAlign of
      vgrthaCenter: ALinePos.X := - ALineSize.cx div 2 {- (ATextSize.cx - ALineSize.cx) div 2};
      vgrthaRight: ALinePos.X := (ATextSize.cx div 2) - ALineSize.cx;
      else ALinePos.X := -(ATextSize.cx div 2);
    end;
    ALinePos.Y := AHalfLineSize;

    ALinePos := RotatePoint(ALinePos, ARotateAngle);
    ALinePos.X := ALinePos.X + APos.X;
    ALinePos.Y := APos.Y - ALinePos.Y;

    ExtTextOut(ACanvas.Handle, ALinePos.X, ALinePos.Y,
               ETO_CLIPPED,
               @ARect,
               PChar(ALine),
               ALineInfo.Length,
               nil);

    APos.X := APos.X + AStepX;
    APos.Y := APos.Y + AStepY;
  end;

  SelectObject(ACanvas.Handle, AOldFont);
  DeleteObject(ANewFont);
  SetCanvasTransparentMode(ACanvas, AOldTransparent);
end;

{
procedure DrawRotatedStrings(Canvas: TCanvas;
                             const TextCenter: TPoint;
                             const TextSize: TSize;
                             const TextClipRect: TRect;
                             Strings: TStrings;
                             RotateAngle: Integer;
                             BackTransparent: Boolean;
                             HorizontalAlignment: TvgrRangeHorzAlign);
var
  NewFont, OldFont: HFONT;
  OldBkMode: Integer;
  ARotated: TPoint;
  I, Y: Integer;
  ALineSize: TSize;
  ALineRect: TRect;
  ALineRotatedCenter, ALineCenter: TPoint;
begin
  NewFont := GetRotatedFontHandle(Canvas.Font,RotateAngle);
  OldFont := SelectObject(Canvas.Handle,NewFont);

  if BackTransparent then
    OldBkMode := SetBkMode(Canvas.Handle,TRANSPARENT)
  else
    OldBkMode := SetBkMode(Canvas.Handle,OPAQUE);

  Y := TextCenter.Y - TextSize.cy div 2;
  for I := 0 to Strings.Count - 1 do
  begin
    if Strings[I] = '' then
    begin
      ALineSize.cx := 0;
      ALineSize.cy := Canvas.TextHeight('Wg')
    end
    else
      ALineSize := Canvas.TextExtent(Strings[I]);

    case HorizontalAlignment of
      vgrhaCenter:
        ALineRect.Left := TextCenter.X - TextSize.cx div 2 + (TextSize.cx - ALineSize.cx) div 2;
      vgrhaRight:
        ALineRect.Left := TextCenter.X + TextSize.cx div 2 - ALineSize.cx;
    else
      ALineRect.Left := TextCenter.X - TextSize.cx div 2;
    end;
    ALineRect.Right := ALineRect.Left + ALineSize.cx;
    ALineRect.Top := Y;
    ALineRect.Bottom := ALineRect.Top + ALineSize.cy;
    ALineCenter := RectCenter(ALineRect);
    ARotated := RotatePoint(Point(-ALineSize.cx div 2, ALineSize.cy div 2), RotateAngle);
    ALineRotatedCenter := RotatePoint(Point(ALineCenter.X - TextCenter.X, TextCenter.Y - ALineCenter.Y), RotateAngle);
    ALineRotatedCenter.X := ALineRotatedCenter.X + TextCenter.X;
    ALineRotatedCenter.Y := TextCenter.Y - ALineRotatedCenter.Y;
    ARotated.X := ARotated.X + ALineRotatedCenter.X;
    ARotated.Y := ALineRotatedCenter.Y - ARotated.Y;

    ExtTextOut(Canvas.Handle,
               ARotated.X,
               ARotated.Y,
               ETO_CLIPPED,
               @TextClipRect,
               PChar(Strings[I]),
               Length(Strings[I]),
               nil);

    Y := Y + ALineSize.cy;
  end;

  SetBkMode(Canvas.Handle,OldBkMode);

  SelectObject(Canvas.Handle,OldFont);
  DeleteObject(NewFont);
end;
}

(*
function vgrWrapRotatedText(const AText: string;
                            ACanvas: TCanvas;
                            ATextRect: TRect;
                            AAngle: Integer): string;
var
  ARadAngle, ATanAngle: Double;
  AWidth, ARectWidth, ARectHeight: Integer;
  ATextMetrics: TEXTMETRIC;
begin
  ARectWidth := RectWidth(ATextRect);
  ARectHeight := RectHeight(ATextRect);
  if ARectWidth = 0 then
    AWidth := ARectHeight
  else
  begin
    if ARectHeight = 0 then
      AWidth := ARectWidth
    else
    begin
      AAngle := AAngle mod 360;
      case Abs(AAngle) of
        0, 180: AWidth := RectWidth(ATextRect);
        90, 270: AWidth := RectHeight(ATextRect);
        else
          begin
            ARadAngle := ToRadians(AAngle);
            ATanAngle := Tan(ARadAngle);
            GetTextMetrics(ACanvas.Handle, ATextMetrics);
            if Abs(ATanAngle) < Abs(((ARectHeight div 2) - (ATextMetrics.tmHeight div 2)) / (ARectWidth div 2)) then
            begin
              // линии должны быть наклонены вдоль оси Y
              AWidth := Abs(Round(ARectWidth / Cos(ARadAngle) - ATextMetrics.tmHeight * ATanAngle));
            end
            else
            begin
              // линии должны быть наклонены вдоль оси X
              AWidth := Abs(Round(ARectHeight / Sin(ARadAngle) - ATextMetrics.tmHeight * ATanAngle));
            end;
          end;
      end;
    end;
  end;

  Result := WrapMemo(ACanvas, AText, AWidth);
end;
*)

function vgrWrapRotatedText(const AText: string;
                            ACanvas: TCanvas;
                            ATextRect: TRect;
                            AAngle: Integer): string;
var
  ARadAngle, ATanAngle: Double;
  AWidth, ARectWidth, ARectHeight: Integer;
  ATextMetrics: TEXTMETRIC;
begin
  ARectWidth := RectWidth(ATextRect);
  ARectHeight := RectHeight(ATextRect);
  if ARectWidth = 0 then
    AWidth := ARectHeight
  else
  begin
    if ARectHeight = 0 then
      AWidth := ARectWidth
    else
    begin
      AAngle := AAngle mod 360;
      case Abs(AAngle) of
        0, 180: AWidth := RectWidth(ATextRect);
        90, 270: AWidth := RectHeight(ATextRect);
        else
          begin
            ARadAngle := ToRadians(AAngle);
            ATanAngle := Tan(ARadAngle);
            GetTextMetrics(ACanvas.Handle, ATextMetrics);
            if Abs(AAngle mod 90) < 45 then
              // by Y
              AWidth := Round(Abs(ARectWidth / Cos(ARadAngle)) - Abs(ATextMetrics.tmHeight * ATanAngle))
            else
              // by X
              AWidth := Round(Abs(ARectHeight / Sin(ARadAngle)) - Abs(ATextMetrics.tmHeight / ATanAngle));
          end;
      end;
    end;
  end;

  Result := WrapMemo(ACanvas, AText, AWidth);
end;

procedure vgrDrawText(ACanvas: TCanvas;
                      const AText: string;
                      const ARect: TRect;
                      AWordWrap: Boolean;
                      AHorzAlign: TvgrRangeHorzAlign;
                      AVertAlign: TvgrRangeVertAlign;
                      ARotateAngle: Integer);
const
  AAlignConvert: Array [TvgrRangeHorzAlign] of TvgrTextHorzAlign = (vgrthaLeft, vgrthaLeft, vgrthaCenter, vgrthaRight);
var
  AWrappedText: string;
  ATextHeight, AVertOffset: Integer;
begin
  SetTextColor(ACanvas.Handle, GetRGBColor(ACanvas.Font.Color));

  // wrap text
  if AWordWrap then
    AWrappedText := vgrWrapRotatedText(AText, ACanvas, ARect, ARotateAngle)
  else
    AWrappedText := AText;

  // draw text
  if ARotateAngle <> 0 then
  begin
    DrawRotatedText(ACanvas, AWrappedText, ARect,
                    ARotateAngle, True, AHorzAlign, AVertAlign);
  end
  else
  begin
    ATextHeight := vgrCalcStringHeight(ACanvas, AWrappedText);

    case AVertAlign of
      vgrvaCenter:
        AVertOffset := (ARect.Bottom - ARect.Top - ATextHeight) div 2;
      vgrvaBottom:
        AVertOffset := ARect.Bottom - ARect.Top - ATextHeight;
      else
        AVertOffset := 0;
    end;
      
    vgrDrawString(ACanvas,
                  AWrappedText,
                  ARect,
                  AVertOffset,
                  AAlignConvert[AHorzAlign]);
  end;
end;

function CreateBorderPen(ASize: Integer; AStyle: pvgrBorderStyle): HPEN;
begin
  Result := CreatePen(APens[AStyle.Pattern], 1, ColorToRGB(AStyle.Color));
end;

function vgrCalcText(ACanvas: TCanvas;
                     const AText: string;
                     const ARect: TRect;
                     AWordWrap: Boolean;
                     AHorzAlign: TvgrRangeHorzAlign;
                     AVertAlign: TvgrRangeVertAlign;
                     ARotateAngle: Integer): TSize;
var
  AWrappedText: string;
  ATextSize: TSize;
  APos: TPoint;
  ATextHorzAlign: TvgrTextHorzAlign;
  ANewFont, AOldFont: HFONT;
  AHalfLineSize, AStepX, AStepY: Integer;
begin
  // wrap text
  if AWordWrap then
    AWrappedText := vgrWrapRotatedText(AText, ACanvas, ARect, ARotateAngle)
  else
    AWrappedText := AText;

  // calculate rectangle
  if ARotateAngle <> 0 then
  begin
    InternalCalcRotatedText(ACanvas,
                            AText,
                            ARotateAngle,
                            AHorzAlign,
                            AVertAlign,
                            True,
                            True,
                            ATextHorzAlign,
                            AStepX,
                            AStepY,
                            ATextSize,
                            Result,
                            APos,
                            AHalfLineSize,
                            ANewFont,
                            AOldFont);
  end
  else
  begin
    Result := vgrCalcStringSize(ACanvas, AWrappedText);
  end;
end;

(*
procedure DrawRotatedText(ACanvas: TCanvas;
                          const AText: string;
                          const ARect: TRect;
                          ARotateAngle: Integer;
                          ABackTransparent: Boolean;
                          AHorizontalAlignment: TvgrRangeHorzAlign;
                          AVerticalAlignment: TvgrRangeVertAlign);
const
  AHorzAlignConvert: array [TvgrRangeHorzAlign] of TvgrTextHorzAlign = (vgrthaLeft, vgrthaLeft, vgrthaCenter, vgrthaRight);
  AVertAlignConvert: array [TvgrRangeVertAlign] of TvgrTextHorzAlign = (vgrthaLeft, vgrthaCenter, vgrthaRight);
var
  ALine: string;
  ATan, ACos, ASin, ARadAngle: Double;
  ALineInfo: TvgrTextLineInfo;
  ATextMetrics: TEXTMETRIC;
  AOldTransparent: Boolean;
  ARotatedTextSize, ATextSize, ALineSize: TSize;
  APos, ALinePos: TPoint;
  ATextHorzAlign: TvgrTextHorzAlign;
  ANewFont, AOldFont: HFONT;
  ALineCount, AAdd, AStepX, AStepY, ARectWidth, ARectHeight: Integer;
begin
  ARectWidth := RectWidth(ARect);
  ARectHeight := RectHeight(ARect);
  if (ARectWidth = 0) or (ARectHeight = 0) then exit;

  ANewFont := GetRotatedFontHandle(ACanvas.Font, ARotateAngle);
  AOldFont := SelectObject(ACanvas.Handle, ANewFont);
  AOldTransparent := SetCanvasTransparentMode(ACanvas, ABackTransparent);

  ATextSize := vgrCalcStringSize(ACanvas, AText, ALineCount);
  ARotateAngle := ARotateAngle mod 360;
  GetTextMetrics(ACanvas.Handle, ATextMetrics);
  AAdd := ATextMetrics.tmHeight div 2;

  case ARotateAngle of
    0, 180:
      begin
        ATextHorzAlign := AHorzAlignConvert[AHorizontalAlignment];
        AStepX := 0;
        AStepY := ATextMetrics.tmHeight;

        ARotatedTextSize.cx := ATextSize.cx;
        ARotatedTextSize.cy := ATextSize.cy;

        APos.X := ATextSize.cx div 2;
        APos.Y := AAdd;
      end;
    90, 270:
      begin
        ATextHorzAlign := AVertAlignConvert[AVerticalAlignment];
        AStepX := ATextMetrics.tmHeight;
        AStepY := 0;

        ARotatedTextSize.cx := ATextSize.cy;
        ARotatedTextSize.cy := ATextSize.cx;

        APos.X := AAdd;
        APos.Y := ATextSize.cx div 2;
      end;
    else
      begin
        ARadAngle := ToRadians(ARotateAngle);
        ATan := Tan(ARadAngle);
        ACos := Cos(ARadAngle);
        ASin := Sin(ARadAngle);

        // линии должны быть наклонены вдоль оси X
        ATextHorzAlign := AVertAlignConvert[AVerticalAlignment];
        AStepX := Abs(Round(ATextMetrics.tmHeight / ASin)){ATextMetrics.tmHeight};
        AStepY := 0;

        ARotatedTextSize.cx := Round(Abs(ATextSize.cx * ACos) + Abs(ATextMetrics.tmHeight / ASin * ALineCount));
        ARotatedTextSize.cy := Round(Abs(ATextSize.cx * ASin) + Abs(ATextMetrics.tmHeight * ACos));

        APos.X := ARotatedTextSize.cy div 2;
        APos.Y := Round((Abs(ATextMetrics.tmHeight * ACos) + Abs(ATextSize.cx * ASin)) / 2);
      end;
  end;

  case AVerticalAlignment of
    vgrvaCenter: APos.Y := APos.Y + (ARectHeight - ARotatedTextSize.cy) div 2;
    vgrvaBottom: APos.Y := APos.Y + ARect.Bottom - ARotatedTextSize.cy;
    else ; // APos.Y - not changed
  end;

  case AHorizontalAlignment of
    vgrhaCenter: APos.X := APos.X + (ARectWidth - ARotatedTextSize.cx) div 2;
    vgrhaRight: APos.X := APos.X + ARect.Right - ARotatedTextSize.cx;
    else ; // APos.X - not changed
  end;

  BeginEnumLines(AText, ALineInfo);
  while EnumLines(AText, ALineInfo) do
  begin
    ALine := Copy(AText, ALineInfo.Start, ALineInfo.Length);

    GetTextExtentPoint32(ACanvas.Handle, PChar(ALine), ALineInfo.Length, ALineSize);
    case ATextHorzAlign of
      vgrthaCenter: ALinePos.X := -ALineSize.cx div 2 {- (ATextSize.cx - ALineSize.cx) div 2};
      vgrthaRight: ALinePos.X := (ATextSize.cx div 2) - ALineSize.cx;
      else ALinePos.X := -ATextSize.cx div 2;
    end;
    ALinePos.Y := AAdd;

    ALinePos := RotatePoint(ALinePos, ARotateAngle);
    ALinePos.X := ALinePos.X + APos.X;
    ALinePos.Y := APos.Y - ALinePos.Y;

    ExtTextOut(ACanvas.Handle, ALinePos.X, ALinePos.Y,
               ETO_CLIPPED,
               @ARect,
               PChar(ALine),
               ALineInfo.Length,
               nil);

    APos.X := APos.X + AStepX;
    APos.Y := APos.Y + AStepY;
  end;

  SelectObject(ACanvas.Handle, AOldFont);
  DeleteObject(ANewFont);
  SetCanvasTransparentMode(ACanvas, AOldTransparent);
end;
*)

end.
