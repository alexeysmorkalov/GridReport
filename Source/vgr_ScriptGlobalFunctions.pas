{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

{Contains the predefined global script functions.}
unit vgr_ScriptGlobalFunctions;

interface

{$I vtk.inc}
uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF}
  {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}Classes,

  vgr_CommonClasses, vgr_Functions;

  procedure ColumnCaption(var AParameters: TvgrOleVariantDynArray; var AResult: OleVariant);

implementation

procedure ColumnCaption(var AParameters: TvgrOleVariantDynArray; var AResult: OleVariant);
begin
  AResult := GetWorksheetColCaption(AParameters[0])
end;

end.
