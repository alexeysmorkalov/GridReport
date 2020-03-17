'Define global variables to store calculation data:
Public Operand1
Public Operand2
Public Operator

'fucnction clears all information about calculations:
Function Init
  Operand1 = "none"
  Operand2 = "none"
  Operator = "none"
  Init = "Calculator initialized. (VBScript)"
End Function

'function receives the command ACommandString and returns the string of the result for display of the calculator
Function Calc(byval ACommandString)
  Calc = ACommandString
  Select Case ACommandString
    'Adding new character to right of current operand
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ","
      Calc = EnterNumber(ACommandString)
    'Clear information about calculations
    Case "C"
      Init
      Calc = "0"
    'Assign new operator. If second operand not empty, then calculate
    Case "*","/","-","+"
      if Operand1 <> "none" then
        if Operand1 <> "none" and Operand2 = "none" then
          Operator = ACommandString
          Calc = Operand1
        end if
        if Operand1 <> "none" and Operand2 <> "none" then
          Calc = ApplyOperatorFunction(Operator,ACommandString)
        end if        
      else
        Calc = "0"      
      end if
    'Calculate
    Case "="
      if Operand1 = "none" then
        Calc = "0"
      end if    
      if Operand1 <> "none" and Operand2 = "none" then
        Calc = Operand1
      end if    
      if Operand1 <> "none" and Operand2 <> "none" then
        Calc = ApplyOperatorFunction(Operator, "none")
      end if
    Case Else      
      MsgBox "Can't read command, please edit script.txt"
  End Select  
End Function

'Add new character to first or to second operand at right
Function EnterNumber(byval ACommandString)
  'If operator empty, then adding new character to first operand
  if Operator = "none" then
    If Operand1 = "none" then
      Operand1 = ""
    End If
    if Operand1 = "" and ACommandString = "," then
      Operand1 = "0"
    end if    
    Operand1 = Operand1 & ACommandString
    EnterNumber = Operand1
  else
  'If operator not empty, then adding new character to second operand
    If Operand2 = "none" then
      Operand2 = ""
    End If
    if Operand2 = "" and ACommandString = "," then
      Operand2 = "0"
    end if
    Operand2 = Operand2 & ACommandString
    EnterNumber = Operand2
  End If
End Function

'Assign new operator. If second operand not empty, then calculate.
Function ApplyOperatorFunction(byval AOperator, byval ANewOperator)  
  Operand1 = CDbl(Operand1)
  Operand2 = CDbl(Operand2)
  Select Case AOperator
    Case "*"
      ApplyOperatorFunction = Operand1 * Operand2
    Case "/"
      if Operand2 = 0 then
        MsgBox "Division by zero!"
        ApplyOperatorFunction = Operand1
      else
        ApplyOperatorFunction = Operand1 / Operand2
      end if
    Case "+"
      ApplyOperatorFunction = Operand1 + Operand2
    Case "-"
      ApplyOperatorFunction = Operand1 - Operand2      
    Case Else
      MsgBox "Undefined operator"
  End Select
  if ANewOperator <> "none" then
    Operator = ANewOperator
    Operand1 = ApplyOperatorFunction    
    Operand2 = "none"
  else
    Init
    Operand1 = ApplyOperatorFunction    
  end if
End Function
