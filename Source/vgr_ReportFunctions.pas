{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains common functions, used by GridReport.}
unit vgr_ReportFunctions;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  vgr_DataStorageTypes;

{Returns the default alignment of the text in the range for the specified type of the range value.
Syntax:
  function GetAutoAlignForRangeValue(AHorzAlign: TvgrRangeHorzAlign; AValueType: TvgrRangeValueType): TvgrRangeHorzAlign;
Parameters:
  AHorzAlign - the horizontally text alignment in the range.
  AValueType - the type of the value in the range.
Return value:
  If the AHorzAlign parameter is equals to vgrhaAuto, then returns
  the value based on AValueType parameter, otherwise returns AHorzAlign.}
  function GetAutoAlignForRangeValue(AHorzAlign: TvgrRangeHorzAlign; AValueType: TvgrRangeValueType): TvgrRangeHorzAlign;

implementation

function GetAutoAlignForRangeValue(AHorzAlign: TvgrRangeHorzAlign; AValueType: TvgrRangeValueType): TvgrRangeHorzAlign;
const
{
Syntax:
  AConvert : array [TvgrRangeValueType] of TvgrRangeHorzAlign =
             (vgrhaLeft, vgrhaRight, vgrhaRight, vgrhaLeft, vgrhaCenter);
Comment:
   rvtNull, rvtInteger, rvtDouble, rvtString, rvtDateTime <br>
   vgrhaLeft, vgrhaCenter, vgrhaRight }
  AConvert : array [TvgrRangeValueType] of TvgrRangeHorzAlign =
             (vgrhaLeft, vgrhaRight, vgrhaRight, vgrhaLeft, vgrhaCenter);
begin
  if AHorzAlign = vgrhaAuto then
    Result := AConvert[AValueType]
  else
    Result := AHorzAlign;
end;

end.
