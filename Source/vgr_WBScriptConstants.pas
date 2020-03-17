{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{    Copyright (c) 2003-2004 by vtkTools   }
{                                          }
{******************************************}

unit vgr_WBScriptConstants;

{$I vtk.inc}

interface

uses
  Graphics, TypInfo,
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  vgr_DataStorage, vgr_DataStorageTypes, vgr_ScriptControl, vgr_CommonClasses;

{$IFNDEF VTK_D6_OR_D7}
const
  clSystemColor = $FF000000;

  {$EXTERNALSYM COLOR_GRADIENTACTIVECAPTION}
  COLOR_GRADIENTACTIVECAPTION = 27;
  {$EXTERNALSYM COLOR_GRADIENTINACTIVECAPTION}
  COLOR_GRADIENTINACTIVECAPTION = 28;

  clMoneyGreen = TColor($C0DCC0);
  clSkyBlue = TColor($F0CAA6);
  clCream = TColor($F0FBFF);
  clMedGray = TColor($A4A0A0);
  clGradientActiveCaption = TColor(clSystemColor or COLOR_GRADIENTACTIVECAPTION);
  clGradientInactiveCaption = TColor(clSystemColor or COLOR_GRADIENTINACTIVECAPTION);
{$ENDIF}

procedure AddGlobalWbConstants;

implementation

procedure AddGlobalWbConstants;

  procedure AddEnumeration(AEnumTypeInfo: PTypeInfo);
  var
    I: Integer;
  begin
    for I := GetTypeData(AEnumTypeInfo).MinValue to GetTypeData(AEnumTypeInfo).MaxValue do
      vgrScriptConstants.Add(GetEnumName(AEnumTypeInfo, I), I);
  end;
  
begin
  AddEnumeration(TypeInfo(TBrushStyle));
  AddEnumeration(TypeInfo(TFontStyle));
  AddEnumeration(TypeInfo(TvgrRangeVertAlign));
  AddEnumeration(TypeInfo(TvgrRangeHorzAlign));
  AddEnumeration(TypeInfo(TvgrRangeValueType));
  AddEnumeration(TypeInfo(TvgrBorderStyle));
  AddEnumeration(TypeInfo(TvgrBorderOrientation));

  with vgrScriptConstants do
  begin
    // fonts charsets
    Add('ANSI_CHARSET', 0);
    Add('DEFAULT_CHARSET', 1);
    Add('SYMBOL_CHARSET', 2);
    Add('MAC_CHARSET', 77);
    Add('SHIFTJIS_CHARSET', 128);
    Add('HANGEUL_CHARSET', 129);
    Add('JOHAB_CHARSET', 130);

    Add('GB2312_CHARSET', 134);
    Add('CHINESEBIG5_CHARSET', 136);
    Add('GREEK_CHARSET', 161);
    Add('TURKISH_CHARSET', 162);
    Add('VIETNAMESE_CHARSET', 163);
    Add('HEBREW_CHARSET', 177);
    Add('ARABIC_CHARSET', 178);

    Add('BALTIC_CHARSET', 186);
    Add('RUSSIAN_CHARSET', 204);
    Add('THAI_CHARSET', 222);
    Add('EASTEUROPE_CHARSET', 238);
    Add('OEM_CHARSET', 255);

    // Colors
    Add('clNone', clNone);
    Add('clAqua', clAqua);
    Add('clBlack', clBlack);
    Add('clBlue', clBlue);
    Add('clDkGray', clDkGray);
    Add('clFuchsia', clFuchsia);
    Add('clGray', clGray);
    Add('clGreen', clGreen);
    Add('clLime', clLime);
    Add('clLtGray', clLtGray);
    Add('clMaroon', clMaroon);
    Add('clNavy', clNavy);
    Add('clOlive', clOlive);
    Add('clPurple', clPurple);
    Add('clRed', clRed);
    Add('clSilver', clSilver);
    Add('clTeal', clTeal);
    Add('clWhite', clWhite);
    Add('clYellow', clYellow);
    Add('clDefault', clDefault);
    Add('clWindowFrame', clWindowFrame);
    Add('clMenuText', clMenuText);
    Add('clWindowText', clWindowText);
    Add('clCaptionText', clCaptionText);
    Add('clActiveBorder', clActiveBorder);
    Add('clInactiveBorder', clInactiveBorder);
    Add('clAppWorkSpace', clAppWorkSpace);
    Add('clHighlight', clHighlight);
    Add('clBtnFace', clBtnFace);
    Add('clBtnShadow', clBtnShadow);
    Add('clGrayText', clGrayText);
    Add('clBtnText', clBtnText);
    Add('clInactiveCaptionText', clInactiveCaptionText);
    Add('cl3DDkShadow', cl3DDkShadow);
    Add('cl3DLight', cl3DLight);
    Add('clInfoText', clInfoText);
    Add('clInfoBk', clInfoBk);
    Add('clMenu', clMenu);
    Add('clWindow', clWindow);
    Add('clInactiveCaption', clInactiveCaption);
    Add('clBackground', clBackground);
    Add('clScrollBar', clScrollBar);
    Add('clCream', clCream);
    Add('clMedGray', clMedGray);
    Add('clMoneyGreen', clMoneyGreen);
    Add('clSkyBlue', clSkyBlue);
    Add('clGradientActiveCaption', clGradientActiveCaption);
    Add('clGradientInactiveCaption', clGradientInactiveCaption);
  end;
end;

end.
