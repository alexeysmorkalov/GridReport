//Define global variables to store calculation data:
var Operand1;
var Operand2;
var Operator;

//fucnction clears all information about calculations:
function Init(){
  //Assign empty value to variables, because assigning null does not wipe previous value of the variable
  Operand1 = "";
  Operand2 = "";
  Operator = "";  
  //Assign null
  Operand1 = null;
  Operand2 = null;
  Operator = null;
  return "Calculator initialized. (JScript)"
}

//function receives the command ACommandString and returns the string of the result for display of the calculator:
function Calc(ACommandString){
  switch (ACommandString)
  {
    //Adding new character to right of current operand
    case "0": case "1": case "2": case "3": case "4": case "5": case "6": case "7":
    case "8": case "9": case "0": case ",":
      return EnterNumber(ACommandString);
    //Clear information about calculations
    case "C":
      Init();
      return "0";
    //Assign new operator. If second operand not empty, then calculate.
    case "*": case "/": case "-": case "+":
      if (Operand1){
        if((Operand1)&&(!Operand2)){
          Operator = ACommandString;
          return Operand1;
        }
        if ((Operand1)&&(Operand2))
          return ApplyOperatorFunction(Operator,ACommandString)
      }
      else
      {
        return "0";
      }
      break;
    //Calculate
    case "=":
      if(!Operand1)
        return "0";
      if((Operand1)&&(!Operand2))
        return Operand1;
      if((Operand1)&&(Operand2))
        return ApplyOperatorFunction(Operator, null);
      break;
    default:
      return "Can't read command";
  }
}

//Add new character to first or to second operand at right
function EnterNumber(ACommandString){
  //If operator null, then adding new character to first operand
  if (!Operator){
    if (!Operand1){
      Operand1 = "";
    }
    if ((Operand1=="")&&(ACommandString==",")){
      Operand1 = "0";
    }
    Operand1 = Operand1 + ACommandString;
    return Operand1
  }else{
  //If operator not null, then adding new character to second operand
    if (!Operand2){
      Operand2 = "";
    }
    if((Operand2=="")&&(ACommandString==",")){
      Operand2 = "0";
    }
    Operand2 = Operand2 + ACommandString;
    return Operand2;
  }
}

//Assign new operator. If second operand not empty, then calculate.
function ApplyOperatorFunction(AOperator,ANewOperator)
{  
  switch (AOperator){
    case "*":
      AResult = 1*Operand1 * 1*Operand2
      break;
    case "/":
      if(1*Operand2==0){
        AResult = 1*Operand1;
      }else{
        AResult = 1*Operand1/(1*Operand2);
      }
      break;
    case '+':
      AResult = 1*Operand1 + 1*Operand2;
      break;
    case "-":
      AResult = 1*Operand1 - 1*Operand2;
      break;
    default:
      AResult = AResult + "Undefined operator("+AOperator+")";
  }
  if(ANewOperator){
    Init();
    Operator = ANewOperator;
    Operand1 = AResult;
  }else{
    Init();
    Operand1 = AResult;
  }
  return AResult;
}
