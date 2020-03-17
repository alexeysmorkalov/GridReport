{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{This unit contais definition of some internal structures.}
unit vgr_DataStorageRecords;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Graphics, Windows, {$IFDEF VGR_D6_OR_D7} Types, {$ENDIF}

  vgr_CommonClasses, vgr_DataStorageTypes;

const
  vgrmask_RangeFlagsWordWrap = $0001;
  vgrmask_RangeFlagsNotCalced = $0002;
  vgrmask_RangeFlagsPassed = $0004;

type

  /////////////////////////////////////////////////
  //
  // rvgrFont
  //
  /////////////////////////////////////////////////
  rvgrFont = packed record
    Size: Integer;
    Pitch: TFontPitch;
    Style: TFontStylesBase;
    Charset: TFontCharset;
    Name: TFontDataName;
    Color: TColor;
  end;
  pvgrFont = ^rvgrFont;

  /////////////////////////////////////////////////
  //
  // rvgrRangeStyle
  //
  /////////////////////////////////////////////////
  rvgrRangeStyle = packed record
    Font: rvgrFont;
    FillBackColor: TColor;
    FillForeColor: TColor;
    FillPattern: TBrushStyle;
    DisplayFormat: TvgrStringIndex;
    HorzAlign: TvgrRangeHorzAlign;
    VertAlign: TvgrRangeVertAlign;
    Angle: Word;
    Flags: Word;
  end;
  pvgrRangeStyle = ^rvgrRangeStyle;

  /////////////////////////////////////////////////
  //
  // rvgrRangeValue
  //
  /////////////////////////////////////////////////
  rvgrRangeValue = packed record
    ValueType: TvgrRangeValueType;
    case TvgrRangeValueType of
      rvtInteger: (vInteger: Integer);
      rvtExtended: (vExtended: Extended);
      rvtString: (vString: TvgrStringIndex);
      rvtDateTime: (vDateTime: TDateTime);
  end;

(*$NODEFINE rvgrRangeValue*)
(*$HPPEMIT ' struct rvgrRangeValue'*)
(*$HPPEMIT '{'*)
(*$HPPEMIT '	Vgr_datastoragetypes::TvgrRangeValueType ValueType;'*)
(*$HPPEMIT '	union'*)
(*$HPPEMIT '	{'*)
(*$HPPEMIT '		struct a'*)
(*$HPPEMIT '		{'*)
(*$HPPEMIT '			System::TDateTime vDateTime;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '		};'*)
(*$HPPEMIT '		struct b'*)
(*$HPPEMIT '		{'*)
(*$HPPEMIT '			int vString;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '		};'*)
(*$HPPEMIT '		struct c'*)
(*$HPPEMIT '		{'*)
(*$HPPEMIT '			Extended vExtended;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '		};'*)
(*$HPPEMIT '		struct d'*)
(*$HPPEMIT '		{'*)
(*$HPPEMIT '			int vInteger;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '		};'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '	};'*)
(*$HPPEMIT '} ;'*)
  pvgrRangeValue = ^rvgrRangeValue;

  /////////////////////////////////////////////////
  //
  // rvgrRange
  //
  /////////////////////////////////////////////////
  rvgrRange = packed record
    Place: TRect;
    Style: Integer;
    Value: rvgrRangeValue;
    Formula: Integer;
    Flags: Byte;
  end;
  pvgrRange = ^rvgrRange;

  /////////////////////////////////////////////////
  //
  // rvgrBorderStyle
  //
  /////////////////////////////////////////////////
  rvgrBorderStyle = packed record
    Width: Integer;
    Pattern: TvgrBorderStyle;
    Color: TColor;
  end;
  pvgrBorderStyle = ^rvgrBorderStyle;

  /////////////////////////////////////////////////
  //
  // rvgrBorder
  //
  /////////////////////////////////////////////////
  rvgrBorder = packed record
    Left: Integer;
    Top: Integer;
    Orientation: TvgrBorderOrientation;
    Style: Integer;
  end;
  pvgrBorder = ^rvgrBorder;

  /////////////////////////////////////////////////
  //
  // rvgrVector
  //
  /////////////////////////////////////////////////
  rvgrVector = packed record
    Number: Integer;
    Size: Integer;
    Visible: Boolean;
  end;
  pvgrVector = ^rvgrVector;

  /////////////////////////////////////////////////
  //
  // rvgrPageVector
  //
  /////////////////////////////////////////////////
  rvgrPageVector = packed record
    Number: Integer;
    Size: Integer;
    Visible: Boolean;
    Flags: Word;
  end;
  pvgrPageVector = ^rvgrPageVector;

  /////////////////////////////////////////////////
  //
  // rvgrCol
  //
  /////////////////////////////////////////////////
  rvgrCol = packed record
    Number: Integer;
    Width: Integer;
    Visible: Boolean;
    Flags: Word;
  end;
  pvgrCol = ^rvgrCol;

  /////////////////////////////////////////////////
  //
  // rvgrRow
  //
  /////////////////////////////////////////////////
  rvgrRow = packed record
    Number: Integer;
    Height: Integer;
    Visible: Boolean;
    Flags: Word;
  end;
  pvgrRow = ^rvgrRow;

  /////////////////////////////////////////////////
  //
  // rvgrSection
  //
  /////////////////////////////////////////////////
  rvgrSection = packed record
    StartPos: Integer;
    EndPos: Integer;
    Level: Integer;
    Flags: Word;
  end;
  pvgrSection = ^rvgrSection;

  /////////////////////////////////////////////////
  //
  // rvgrStyleHeader
  //
  /////////////////////////////////////////////////
  rvgrStyleHeader = packed record
    Hash: Word;
    RefCount: Integer;
  end;
  pvgrStyleHeader = ^rvgrStyleHeader;

  /////////////////////////////////////////////////
  //
  // rvgrFormulaHeader
  //
  /////////////////////////////////////////////////
  rvgrFormulaHeader = packed record
    Hash: Word;
    RefCount: Integer;
    ItemCount: Integer; 
  end;
  pvgrFormulaHeader = ^rvgrFormulaHeader;

implementation

end.
