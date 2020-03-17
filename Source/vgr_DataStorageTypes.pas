{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{This unit contains some additional types.}
unit vgr_DataStorageTypes;

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  Classes;

type

{Defines a type for string index in the rvgrRangeStyle structure.
See also:
  rvgrRangeStyle}
  TvgrStringIndex = Integer;
{Describes a type of range value:
Items:
  rvtNull - The range has no value.
  rvtInteger - The range has an integer value.
  rvtExtended - The tange has an extended value.
  rvtString - The range has a string value.
  rvtDateTime - The range has a datetime value.
Syntax:
  TvgrRangeValueType = (rvtNull, rvtInteger, rvtExtended, rvtString, rvtDateTime)}
  TvgrRangeValueType = (rvtNull, rvtInteger, rvtExtended, rvtString, rvtDateTime);
{Describes a vertical text alignment within the range.
Items:
  vgrvaTop - The text appears at the top of the its range.
  vgrvaCenter - The text is vertically centered in the range.
  vgrvaBottom - The text appears along the bottom of the range.
Syntax:
  TvgrRangeVertAlign = (vgrvaTop, vgrvaCenter, vgrvaBottom)}
  TvgrRangeVertAlign = (vgrvaTop, vgrvaCenter, vgrvaBottom);
{Describes a horizontal text alignment within the range.
Items:
  vgrhaAuto - The horizontal text alignment depends from a type of range's value.
  vgrhaLeft - Text is left-justified: Lines all begin at the left edge of the range.
  vgrhaCenter - Text is centered in the range.
  vgrhaRight - Text is right-justified: Lines all end at the right edge of the control.
Syntax:
  TvgrRangeHorzAlign = (vgrhaAuto, vgrhaLeft, vgrhaCenter, vgrhaRight)}
  TvgrRangeHorzAlign = (vgrhaAuto, vgrhaLeft, vgrhaCenter, vgrhaRight);
{Describes the style of borders.
Items:
  vgrbsSolid - A solid line.
  vgrbsDash - A line made up of a series of dashes.
  vgrbsDot - A line made up of a series of dots.
  vgrbsDashDot - A line made up of alternating dashes and dots.
  vgrbsDashDotDot - A line made up of a series of dash-dot-dot combinations.
Syntax:
  TvgrBorderStyle = (vgrbsSolid, vgrbsDash, vgrbsDot, vgrbsDashDot, vgrbsDashDotDot)}
  TvgrBorderStyle = (vgrbsSolid, vgrbsDash, vgrbsDot, vgrbsDashDot, vgrbsDashDotDot);

implementation

end.
