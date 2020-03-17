{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{      Copyright (c) 2003 by vtkTools      }
{                                          }
{******************************************}

{Contains classes used for keyboard handling.
Main class is a @link(TvgrKeyboardFilter) that contains mapping table
between keyboard shortcuts and procedures processing this commands.
All classes in this module are internal, please do not use them in
your programms.
See also:
  TvgrKeyboardFilter}
unit vgr_Keyboard;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, Classes;

type

{Specifies type of function, called when keys combination pressed.
Return value:
  Must returns true if addititional processing are required.}
  TvgrKeyboardCommandProc = function: Boolean of object;
  /////////////////////////////////////////////////
  //
  // TvgrKeyboardCommand
  //
  /////////////////////////////////////////////////
  {Describes link between combination of keys and function.}
  TvgrKeyboardCommand = class
  private
    FKey: Word;
    FShift: TShiftState;
    FCommand: TvgrKeyboardCommandProc;
  end;

{Specifies type of the callback procedure that called before and after
call of the keyboard command function.
Parameters:
  ACommand - keyboard command that occurs.
  AShiftPressed - true if the Shift key is held down.
  AAltPressed - true if the Alt key is hel down.}
  TvgrKeyboardCommandEvent = procedure(ACommand: TvgrKeyboardCommand; AShiftPressed, AAltPressed: Boolean) of object;
{TvgrKeyboardCommandList holds the keyboard commands (TvgrKeyboardCommand) objects that represent the
link between combination of keys and functions that must be called.}
  /////////////////////////////////////////////////
  //
  // TvgrKeyboardCommandList
  //
  /////////////////////////////////////////////////
  TvgrKeyboardCommandList = class(TList)
  private
    FOnBeforeCommand: TvgrKeyboardCommandEvent;
    FOnAfterCommand: TvgrKeyboardCommandEvent;
    procedure AddCommand(AKey: Word; AShift: TShiftState; ACommand: TvgrKeyboardCommandProc);
    procedure DoCommand(AKey: Word; AShift: TShiftState);
    function GetItem(Index: Integer): TvgrKeyboardCommand;
    function IsShiftPressed(AKey: Word; AShift: TShiftState): Boolean;
    function IsAltPressed(AKey: Word; AShift: TShiftState): Boolean;
  public
{Destroys TvgrKeyboardCommandList}
    destructor Destroy; override;
{Deletes all items from the list.}
    procedure Clear; override;
{Lists the object references. Use Items to obtain a pointer to a specific object in the array.
The Index parameter indicates the index of the object,
where 0 is the index of the first object, 1 is the index of the second object, and so on.
Syntax:
  property Items[Index: Integer]: TvgrKeyboardCommand read; default;
Parameters:
  Index - index of object in the array.}
    property Items[Index: Integer]: TvgrKeyboardCommand read GetItem; default;
{Occurs before processing of the keyboard command.
Syntax:
  property OnBeforeCommand: TvgrKeyboardCommandEvent read write;}
    property OnBeforeCommand: TvgrKeyboardCommandEvent read FOnBeforeCommand write FOnBeforeCommand;
{Occurs after processing of the keyboard command.
Syntax:
  property OnAfterCommand: TvgrKeyboardCommandEvent read write;}
    property OnAfterCommand: TvgrKeyboardCommandEvent read FOnAfterCommand write FOnAfterCommand;
  end;

  TvgrKeyboardFilterKeyPressEvent = procedure(Key: Char) of object;
  /////////////////////////////////////////////////
  //
  // TvgrKeyboardFilter
  //
  /////////////////////////////////////////////////
{TvgrKeyboardFilter contains mapping table
between keyboard shortcuts and procedures processing this commands.
Instance of this class created by the control, then the control adds
keyboard commands which can be processed.
In methods TWinControl.KeyDown, TWinControl.KeyPress the control invoke
appropriate methods of the TvgrKeyboardFilter object.
TvgrKeyboardFilter recognizes the keyboard command and execute them.
This is the internal class, please do not use them in your programms.}
  TvgrKeyboardFilter = class
  private
    FIgnoreKeyPress: Boolean;
    FOnKeyPress: TvgrKeyboardFilterKeyPressEvent;
    FSelectionCommandList: TvgrKeyboardCommandList;
    FCommandList: TvgrKeyboardCommandList;
    function IsChar(AKey: Char): Boolean;
    function GetOnBeforeSelectionCommand: TvgrKeyboardCommandEvent;
    procedure SetOnBeforeSelectionCommand(Value: TvgrKeyboardCommandEvent);
    function GetOnAfterSelectionCommand: TvgrKeyboardCommandEvent;
    procedure SetOnAfterSelectionCommand(Value: TvgrKeyboardCommandEvent);
    function GetOnBeforeCommand: TvgrKeyboardCommandEvent;
    procedure SetOnBeforeCommand(Value: TvgrKeyboardCommandEvent);
    function GetOnAfterCommand: TvgrKeyboardCommandEvent;
    procedure SetOnAfterCommand(Value: TvgrKeyboardCommandEvent);
  protected
    procedure DoCommand(AKey: Word; AShift: TShiftState);
  public
    {Creates TvgrKeyboardFilter object}
    constructor Create;
    {Destroys TvgrKeyboardFilter object}
    destructor Destroy; override;

{This method must be called when key is held down, in the TControl.KeyDown method.
Parameters:
  AKey - the key on the keyboard. For non-alphanumeric keys, use virtual key codes to determine the key pressed.
  AShift - indicates whether the Shift, Alt, or Ctrl keys are combined with the keystroke.}
    procedure KeyDown(AKey: Word; AShift: TShiftState);
{This method must be called when key is pressed, in the TControl.KeyPress method.
Parameters:
  AKey - the ASCII character of the key pressed.}
    procedure KeyPress(AKey: Char);
{Registers the new keyboard command. AKey and AShift parameters describes a combination of
keys for this command.
Parameters:
  AKey - the key on the keyboard.
  AShift - indicates whether the Shift, Alt, or Ctrl keys are combined with the keystroke.
  ACommand - The procedure, which should be called for this combination of keys.}
    procedure AddCommand(AKey: Word; AShift: TShiftState; ACommand: TvgrKeyboardCommandProc);
{Registers the new keyboard command.
This method adds at once two commands, one without Shift, second with Shift.
Parameters:
  AKey - the key on the keyboard.
  AShift - indicates whether the Shift, Alt, or Ctrl keys are combined with the keystroke.
  ACommand - The procedure, which should be called for this combination of keys.}
    procedure AddSelectionCommand(AKey: Word; AShift: TShiftState; ACommand: TvgrKeyboardCommandProc);

{Occurs when key pressed and Ctrl and Shift keys are not held down.
The Key parameter in the OnKeyPress event handler is of type Char;
therefore, the OnKeyPress event registers the ASCII character of the key pressed.}
    property OnKeyPress: TvgrKeyboardFilterKeyPressEvent read FOnKeyPress write FOnKeyPress;
{Occurs before executing the keyboard command.}
    property OnBeforeSelectionCommand: TvgrKeyboardCommandEvent read GetOnBeforeSelectionCommand write SetOnBeforeSelectionCommand;
{Occurs after executing the keyboard command.}
    property OnAfterSelectionCommand: TvgrKeyboardCommandEvent read GetOnAfterSelectionCommand write SetOnAfterSelectionCommand;
{Occurs before executing the keyboard command.}
    property OnBeforeCommand: TvgrKeyboardCommandEvent read GetOnBeforeCommand write SetOnBeforeCommand;
{Occurs after executing the keyboard command.}
    property OnAfterCommand: TvgrKeyboardCommandEvent read GetOnAfterCommand write SetOnAfterCommand;
  end;

implementation

/////////////////////////////////////////////////
//
// TvgrKeyboardCommandList
//
/////////////////////////////////////////////////

destructor TvgrKeyboardCommandList.Destroy;
begin
  Clear;
  inherited;
end;

function TvgrKeyboardCommandList.GetItem(Index: Integer): TvgrKeyboardCommand;
begin
  Result := TvgrKeyboardCommand(inherited Items[Index]);
end;

procedure TvgrKeyboardCommandList.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  inherited;
end;

function TvgrKeyboardCommandList.IsShiftPressed(AKey: Word; AShift: TShiftState): Boolean;
begin
  Result := ssShift in AShift;
end;

function TvgrKeyboardCommandList.IsAltPressed(AKey: Word; AShift: TShiftState): Boolean;
begin
  Result := ssAlt in AShift;
end;

procedure TvgrKeyboardCommandList.AddCommand(AKey: Word; AShift: TShiftState; ACommand: TvgrKeyboardCommandProc);
var
  AObject: TvgrKeyboardCommand;
begin
  AObject := TvgrKeyboardCommand.Create;
  with AObject do
  begin
    FKey := AKey;
    FShift := AShift;
    FCommand := ACommand;
  end;
  Add(AObject);
end;

procedure TvgrKeyboardCommandList.DoCommand(AKey: Word; AShift: TShiftState);
var
  I: Integer;
  AIsShiftPressed: Boolean;
  AIsAltPressed: Boolean;
begin
  I := 0;
  while (I < Count) and not ((AKey = Items[I].FKey) and (AShift = Items[I].FShift)) do
    Inc(I);
  if I < Count then
  begin
    AIsShiftPressed := IsShiftPressed(AKey, AShift);
    AIsAltPressed := IsAltPressed(AKey, AShift);
    if Assigned(FOnBeforeCommand) then
      FOnBeforeCommand(Items[I], AIsShiftPressed, AIsAltPressed);
    if Items[I].FCommand then
      if Assigned(FOnAfterCommand) then
        FOnAfterCommand(Items[I], AIsShiftPressed, AIsAltPressed);
  end;
end;

/////////////////////////////////////////////////
//
// TvgrKeyboardFilter
//
/////////////////////////////////////////////////
constructor TvgrKeyboardFilter.Create;
begin
  inherited Create;
  FSelectionCommandList := TvgrKeyboardCommandList.Create;
  FCommandList := TvgrKeyboardCommandList.Create;
end;

destructor TvgrKeyboardFilter.Destroy;
begin
  FSelectionCommandList.Free;
  FCommandList.Free;
  inherited;
end;

function TvgrKeyboardFilter.GetOnBeforeSelectionCommand: TvgrKeyboardCommandEvent;
begin
  Result := FSelectionCommandList.OnBeforeCommand;
end;

procedure TvgrKeyboardFilter.SetOnBeforeSelectionCommand(Value: TvgrKeyboardCommandEvent);
begin
  FSelectionCommandList.OnBeforeCommand := Value;
end;

function TvgrKeyboardFilter.GetOnAfterSelectionCommand: TvgrKeyboardCommandEvent;
begin
  Result := FSelectionCommandList.OnAfterCommand;
end;

procedure TvgrKeyboardFilter.SetOnAfterSelectionCommand(Value: TvgrKeyboardCommandEvent);
begin
  FSelectionCommandList.OnAfterCommand := Value;
end;

function TvgrKeyboardFilter.GetOnBeforeCommand: TvgrKeyboardCommandEvent;
begin
  Result := FCommandList.OnBeforeCommand;
end;

procedure TvgrKeyboardFilter.SetOnBeforeCommand(Value: TvgrKeyboardCommandEvent);
begin
  FCommandList.OnBeforeCommand := Value;
end;

function TvgrKeyboardFilter.GetOnAfterCommand: TvgrKeyboardCommandEvent;
begin
  Result := FCommandList.OnAfterCommand;
end;

procedure TvgrKeyboardFilter.SetOnAfterCommand(Value: TvgrKeyboardCommandEvent);
begin
  FCommandList.OnAfterCommand := Value;
end;

procedure TvgrKeyboardFilter.AddSelectionCommand(AKey: Word; AShift: TShiftState; ACommand: TvgrKeyboardCommandProc);
begin
  FSelectionCommandList.AddCommand(AKey, AShift, ACommand);
  FSelectionCommandList.AddCommand(AKey, AShift + [ssShift], ACommand);
end;

procedure TvgrKeyboardFilter.AddCommand(AKey: Word; AShift: TShiftState; ACommand: TvgrKeyboardCommandProc);
begin
  FCommandList.AddCommand(AKey, AShift, ACommand);
end;

procedure TvgrKeyboardFilter.DoCommand(AKey: Word; AShift: TShiftState);
begin
  FSelectionCommandList.DoCommand(AKey, AShift);
  FCommandList.DoCommand(AKey, AShift);
end;

function TvgrKeyboardFilter.IsChar(AKey: Char): Boolean;
begin
  Result := True;
end;

procedure TvgrKeyboardFilter.KeyDown(AKey: Word; AShift: TShiftState);
begin
  FIgnoreKeyPress := (ssCtrl in AShift) or (ssAlt in AShift) or (AKey in [VK_BACK, VK_RETURN, VK_ESCAPE, VK_TAB]);
  DoCommand(AKey, AShift);
end;

procedure TvgrKeyboardFilter.KeyPress(AKey: Char);
begin
  if (not FIgnoreKeyPress) and IsChar(AKey) then
    if Assigned(FOnKeyPress) then
      FOnKeyPress(AKey);
end;

end.

