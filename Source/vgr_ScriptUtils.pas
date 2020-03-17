{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Functions for script engine.

See also:
  vgr_ScriptControl
}
unit vgr_ScriptUtils;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComObj, ActiveX;


{Determine what Active Scripting engines are installed on machine
Parameters:
  Container - TStrings object, which filled with names of installed script engine
Example:
  GetScriptEngines(Memo1.Lines);}
procedure GetScriptEngines(const Container : TStrings);

{vgrProgIDToClassID retrieves, CLSID for a given programmatic ID. 
Parameters:
  ProgID - specifies the programmatic ID for which the CLSID is requested. 
  CLASSID - the class ID (CLSID) TGUID that corresponds to the string specified as the ProgID parameter.
Return value:
  If vgrProgIDToClassID succeeds it returns true, if fails it returns false.
Example:
  if vgrProgIDToClassID('VBScript',VBSCLSID) then
    Obj := CreateComObject(VBSCLSID);}
function vgrProgIDToClassID(const ProgID: string; out CLASSID : TGUID): boolean;

implementation

uses
  vgr_AXScript;

{ Convert a programmatic ID to a class ID }
function vgrProgIDToClassID(const ProgID: string; out CLASSID : TGUID): boolean;
begin
  Result := CLSIDFromProgID(PWideChar(WideString(ProgID)), CLASSID) = S_OK;
end;

function CheckScriptEngine(const CLASSID : TGUID) : boolean;
var
  ActiveScript : IActiveScript;
begin
  Result :=  CoCreateInstance(ClassID, nil, CLSCTX_INPROC_SERVER or
    CLSCTX_LOCAL_SERVER, IUnknown, ActiveScript) = S_OK;
end;

procedure GetScriptEngines(const Container : TStrings);
var
  ourCatId : TCatId;
  CatInformation : ICatInformation;
  EnumGUID : IEnumGUID;
  ourCLSID : TGUID;
  lCount : cardinal;
  pString : PWideChar;
begin
  ourCatId := CATID_ActiveScript;
  OleCheck(CoCreateInstance(
            CLSID_StdComponentCategoriesMgr,
            nil,
            CLSCTX_ALL,
            IID_ICatInformation,
            CatInformation
            ));

  OleCheck(CatInformation.EnumClassesOfCategories(1,@ourCatId,Cardinal(-1),nil,EnumGUID));

  while true do
  begin
    lCount := 0;
    pString := nil;
    OleCheck(EnumGUID.Next(1,ourCLSID,lCount));
    if lCount = 0 then
      break;
    if CheckScriptEngine(ourCLSID) then
    begin
      if ProgIDFromCLSID(ourCLSID,pString) = S_OK then
        Container.Add(pString);
    end;
  end;

end;

end.
