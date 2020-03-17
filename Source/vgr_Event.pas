{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

unit vgr_Event;

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes,

  vgr_ScriptControl, vgr_CommonClasses;

type

  /////////////////////////////////////////////////
  //
  // TvgrScriptEvents
  //
  /////////////////////////////////////////////////
  TvgrScriptEvents = class(TvgrPersistent)
  end;

implementation

/////////////////////////////////////////////////
//                       
// TvgrScriptEvents
//
/////////////////////////////////////////////////

end.
