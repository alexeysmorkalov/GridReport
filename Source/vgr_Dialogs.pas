{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{ Contains functions for creation and change menu, for display MessageBox`es, etc }
unit vgr_Dialogs;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes, SysUtils, Forms, Menus;

  {Use MBox to display a generic dialog box a message and one or more buttons.
   Caption is the caption of the dialog box and is optional.
   MBox is an encapsulation of the TApplication.MessageBox function.
Parameters:
   Text - the value of the Text parameter is the message, which can be longer than 255 characters if necessary.<br>
Long messages are automatically wrapped in the message box.
   Caption - the value of the Caption parameter is the caption that appears in the title bar of the dialog box.
Captions can be longer than 255 characters, but don't wrap. A long caption results in a wide message box.
   AFlags - the Flags parameter specifies what buttons appear on the message box and the behavior
(possible return values).<br>
The following table lists the possible values. These values can be combined to obtain the desired effect.
<br>MB_ABORTRETRYIGNORE - The message box contains three push buttons: Abort, Retry, and Ignore.
<br>MB_OK - The message box contains one push button: OK. This is the default.
<br>MB_OKCANCEL - The message box contains two push buttons: OK and Cancel.
<br>MB_RETRYCANCEL - The message box contains two push buttons: Retry and Cancel.
<br>MB_YESNO - The message box contains two push buttons: Yes and No.
<br>MB_YESNOCANCEL	- The message box contains three push buttons: Yes, No, and Cancel.
Return value:
   MBox returns 0 if there isn’t enough memory to create the message box.
   <br>Otherwise it returns one of the following values:
   <br>IDOK(1) - The user chose the OK button.
   <br>IDCANCEL(2) - The user chose the Cancel button.
   <br>IDABORT(3) - The user chose the Abort button.
   <br>IDRETRY(4) - The user chose the Retry button.
   <br>IDIGNORE(5) - The user chose the Ignore button.
   <br>IDYES(6) - The user chose the Yes button.
   <br>IDNO(7) - The user chose the No button. }
  function MBox(const Text, Caption: string; AFlags: Integer): Integer; overload;
  { This function is similar to the function,
    function MBox(const Text, Caption: string; AFlags: Integer): Integer;
    however as Title of the message window title of application is used (TApplication.Title property). }
  function MBox(const Text: string; AFlags: Integer): Integer; overload;
  { This function is similar to the function,
    function MBox(const Text, Caption: string; AFlags: Integer): Integer;
    however text of the message window is formed with use of the Format function,
    as parameters for it used TextMask and TextParams. }
  function MBoxFmt(const TextMask, Caption: string; AFlags: Integer; TextParams: Array of const): Integer; overload;
  { This function is similar to the function,
    function MBoxFmt(const TextMask, Caption: string; AFlags: Integer; TextParams: Array of const): Integer;
    however as Title of the message window title of application is used (TApplication.Title property). }
  function MBoxFmt(const TextMask: string; AFlags: Integer; TextParams: Array of const): Integer; overload;

  { Deletes all menu items from popup menu. 
Parameters:
  APopupMenu - TPopupMenu}
  procedure ClearPopupMenu(APopupMenu: TPopupMenu);
  { Deletes all sub items from menu item.
Parameters:
  AMenuItem - TMenuItem}
  procedure ClearMenuItem(AMenuItem: TMenuItem);
  { Add menu item to menu, returns added menu item.
Parameters:
    AMenu - menu in which is added menu item, used as owner of menu item.
    AParent - parent menu item for added item.
    ACaption - caption of menu item.
    AOnClick - procedure, which called when user click on menu item.
    ATag - tag of menu item.
    AImageResName - name of resource bitmap, loaded in bitmap property of menu item.
    AEnabled - true - menu item are enabled.
    AChecked - true - menu item are checked.
    AShortCut - shortcut of menu item.
Return value:
  TMenuItem}
  function AddMenuItem(AMenu: TMenu;
                   AParent: TMenuItem;
                   const ACaption: string;
                   AOnClick: TNotifyEvent = nil;
                   ATag: Integer = 0;
                   const AImageResName: string = '';
                   AEnabled: Boolean = True;
                   AChecked: Boolean = False;
                   const AShortCut: string = ''): TMenuItem;
  {  Add separator to menu, returns added menu item. 
Parameters:
     AMenu - menu in which is added separator, used as owner of separator.
     AParent - parent menu item for added separator.
Return value:
  TMenuItem}
  function AddMenuSeparator(AMenu: TMenu; AParent: TMenuItem): TMenuItem;

implementation

function MBox(const Text, Caption: string; AFlags: Integer): Integer;
begin
  Result := Application.MessageBox(PChar(Text), PChar(Caption), AFlags)
end;

function MBoxFmt(const TextMask, Caption: string; AFlags: Integer; TextParams: Array of const): Integer;
begin
  Result := Application.MessageBox(PChar(Format(TextMask, TextParams)), PChar(Caption), AFlags)
end;

function MBox(const Text: string; AFlags: Integer): Integer;
begin
  Result := MBox(Text, Application.Title, AFlags);
end;

function MBoxFmt(const TextMask: string; AFlags: Integer; TextParams: Array of const): Integer;
begin
  Result := MBoxFmt(TextMask, Application.Title, AFlags, TextParams);
end;

procedure ClearPopupMenu(APopupMenu: TPopupMenu);
begin
  while APopupMenu.Items.Count > 0 do
    APopupMenu.Items[0].Free;
end;

procedure ClearMenuItem(AMenuItem: TMenuItem);
begin
  while AMenuItem.Count > 0 do
    AMenuItem.Items[0].Free;
end;

function AddMenuItem(AMenu: TMenu;
                 AParent: TMenuItem;
                 const ACaption: string;
                 AOnClick: TNotifyEvent = nil;
                 ATag: Integer = 0;
                 const AImageResName: string = '';
                 AEnabled: Boolean = True;
                 AChecked: Boolean = False;
                 const AShortCut: string = ''): TMenuItem;
begin
  Result := TMenuItem.Create(AMenu);
  with Result do
  begin
    Caption := ACaption;
    OnClick := AOnClick;
    Tag := ATag;
    if AImageResName <> '' then
      Bitmap.LoadFromResourceName(hInstance, AImageResName);
    Enabled := AEnabled;
    Checked := AChecked;
    ShortCut := TextToShortCut(AShortCut);
  end;
  if AParent = nil then
    AMenu.Items.Add(Result)
  else
    AParent.Add(Result);
end;

function AddMenuSeparator(AMenu: TMenu; AParent: TMenuItem): TMenuItem;
begin
  Result := AddMenuItem(AMenu, AParent, '-');
end;

end.
