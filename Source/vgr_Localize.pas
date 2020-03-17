{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes and methods working with language resources.
Main class - TvgrLocalizer implements methods for loading the  resource strings,
for loading and unloading the resource DLL`s.
You can get access to instance of TvgrLocalizer through a Localizer global variable.
The global functions (vgrLoadStr, vgrLoadResourceDll, vgrUnLoadResourceDll) use
methods of the TvgrLocalizer class, the reference on which is stored in a global
variable Localizer.
See also:
  TvgrLocalizer, vgrLoadStr, vgrLoadResourceDll, vgrUnLoadResourceDll}
unit vgr_Localize;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Classes, typinfo, vgr_CommonClasses;

type
  /////////////////////////////////////////////////
  //
  // TvgrLocalizer
  //
  /////////////////////////////////////////////////
{Implements methods for loading the resource strings, for loading and unloading the resource DLL`s.
In your program you can dynamically be switched between various resource DLL`s,
to make it use methods LoadDll and UnLoadDll.
To load a resource string use LoadStr method.
Instance of this class is created in the start of programm and is saved in Localizer
global variable.}
  TvgrLocalizer = class(TObject)
  private
    FDllModule: HMODULE;
    FDllFileName: string;
  public
  {Destroys TvgrLocalizer object}
    destructor Destroy; override;

{This method loads DLL with language resources.
After this, all pre-compiled resources are overlapped by this DLL resources.
Only one dll can be loaded in one time.
Parameters:
  AFileName - full name of the DLL file.
Return value:
  Returns true if the loading has passed successfully.}
    function LoadDll(const AFileName: string): Boolean;
{This method unloads the DLL, loaded by LoadDll call.
After this, the application will use pre-compiled language resources.}
    procedure UnLoadDll;
{This function loads and returns string from current resources.
The actual ID of the resource string is svgrResStringsBase + AStringID.
svgrResStringsBase constant defined in vgr_StringIDs module.
If string is not found the empty string returns.
Parameters:
  AStringID - offset of resource string ID relative to svgrResStringsBase constant.
Return value:
  Returns the loaded string or empty string if resource string is not found.}
    function LoadStr(AStringID: Integer): string; overload;
{This function loads and returns string from current resources.
The actual ID of the resource string is ABaseIndex + AStringID.
If string is not found the empty string returns.
Parameters:
  ABaseIndex - base index of resource string.
  AStringID - offset of resource string ID relative to ABaseIndex parameter.
Return value:
  Returns the loaded string or empty string if resource string is not found.}
    function LoadStr(ABaseIndex, AStringID: Integer): string; overload;
{This function loads and returns string from current resources.
The actual ID of the resource string is ABaseIndex + AStringID.
Parameters:
  ABaseIndex - base index of resource string.
  AStringID - offset of resource string ID relative to ABaseIndex parameter.
  S - in this parameter the loaded string is returned. if string is not found
the empty string is returned.
Return value:
  Returns true if resource string is found.}
    function LoadStr(ABaseIndex, AStringID: Integer; var S: string): Boolean; overload;
  end;

{Loads and returns string from current resources.
This function is equivalent to Localizer.LoadStr(AID).
Parameters:
  AID - offset of resource string ID relative to svgrResStringsBase constant.
Return value:
  Returns the loaded string or empty string if resource string is not found.}
function vgrLoadStr(AID: Integer): string; overload;
{Loads and returns string from current resources.
This function is equivalent to Localizer.LoadStr(ABaseIndex, AID).
Parameters:
  ABaseIndex - base index of resource string.
  AID - offset of resource string ID relative to ABaseIndex parameter.
Return value:
  Returns the loaded string or empty string if resource string is not found.}
function vgrLoadStr(ABaseIndex, AID: Integer): string; overload;
{This method loads DLL with language resources.
This function is equivalent to Localizer.LoadDll(AFileName).
Parameters:
  AFileName - full name of the DLL file.
Return value:
  Returns true if the loading has passed successfully.}
function vgrLoadResourceDll(const AFileName: string): Boolean;
{This method unloads the DLL, loaded by vgrLoadResourceDll call.
After this, the application will use pre-compiled language resources.
This function is equivalent to Localizer.UnLoadDll.}
procedure vgrUnLoadResourceDll;

var
  Localizer: TvgrLocalizer;

implementation

uses
  SysUtils, vgr_Functions, vgr_StringIDs;

const
  MAX_RESOURCESTRING_LENGTH = 1024;
  
/////////////////////////////////////////////////
//
// FUNCTIONS
//
/////////////////////////////////////////////////
function vgrLoadResourceDll(const AFileName: string): Boolean;
begin
  Result := Localizer.LoadDll(AFileName);
end;

procedure vgrUnLoadResourceDll;
begin
  Localizer.UnLoadDll;
end;

function vgrLoadStr(ABaseIndex, AID: Integer): string;
begin
  Result := Localizer.LoadStr(ABaseIndex, AID);
end;

function vgrLoadStr(AID: Integer): string;
begin
  Result := Localizer.LoadStr(AID);
end;

/////////////////////////////////////////////////
//
// TvgrLocalizer
//
/////////////////////////////////////////////////
destructor TvgrLocalizer.Destroy;
begin
  if FDllModule <> 0 then
    UnLoadDll;
  inherited;
end;

function TvgrLocalizer.LoadDll(const AFileName: string): Boolean;
begin
  if FDllFileName <> AFileName then
  begin
    if FDllModule <> 0 then
      UnLoadDll;
      
    FDllModule := LoadLibrary(PChar(AFileName));
    Result := FDllModule <> 0;
    if Result then
      FDllFileName := AFileName
    else
      FDllFileName := '';
  end
  else
    Result := True;
end;

procedure TvgrLocalizer.UnLoadDll;
begin
  FreeLibrary(FDllModule);
end;

function TvgrLocalizer.LoadStr(ABaseIndex, AStringID: Integer; var S: string): Boolean;
var
  AModule: HMODULE;
  ABuffer: array [0..MAX_RESOURCESTRING_LENGTH - 1] of char;
begin
  if FDllModule = 0 then
    AModule := HInstance
  else
    AModule := FDllModule;
  Result := LoadString(AModule, AStringID + ABaseIndex, ABuffer, MAX_RESOURCESTRING_LENGTH) <> 0;
  if Result then
    S := StrPas(ABuffer)
  else
    S := '';
end;

function TvgrLocalizer.LoadStr(AStringID: Integer): string;
begin
  Result := LoadStr(svgrResStringsBase, AStringID);
end;

function TvgrLocalizer.LoadStr(ABaseIndex, AStringID: Integer): string;
var
  S: string;
begin
  if LoadStr(ABaseIndex, AStringID, S) then
    Result := S
  else
  begin
    Result := SysUtils.LoadStr(ABaseIndex + AStringID);
  end;
end;

initialization

  Localizer := TvgrLocalizer.Create;

finalization

  FreeAndNil(Localizer);

end.

