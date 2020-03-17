
{******************************************}
{                                          }
{           vtk Export library             }
{                                          }
{      Copyright (c) 2002 by vtkTools      }
{                                          }
{******************************************}

{ Contains definition of formulas table used to convert from name of function to his MS Excel ID. }
unit vgr_ExcelFormula_iftab;

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Classes;

type

/////////////////////////////////////////////////
//
// rvteExcelFunctionInfo
//
/////////////////////////////////////////////////
rvteExcelFunctionInfo = record
  FuncName : string;
  iftab : word;
  VarParams : boolean;
end;

const
  vteExcelFunctionsCount = 340;

var
vteExcelFunctions : array [1..vteExcelFunctionsCount] of rvteExcelFunctionInfo =
((FuncName: 'Count'; iftab: 0; VarParams: false),
 (FuncName: 'If'; iftab: 1; VarParams: false),
 (FuncName: 'Isna'; iftab: 2; VarParams: false),
 (FuncName: 'Iserror'; iftab: 3; VarParams: false),
 (FuncName: 'Sum'; iftab: 4; VarParams: false),
 (FuncName: 'Average'; iftab: 5; VarParams: false),
 (FuncName: 'Min'; iftab: 6; VarParams: false),
 (FuncName: 'Max'; iftab: 7; VarParams: false),
 (FuncName: 'Row'; iftab: 8; VarParams: false),
 (FuncName: 'Column'; iftab: 9; VarParams: false),
 (FuncName: 'Na'; iftab: 10; VarParams: false),
 (FuncName: 'Npv'; iftab: 11; VarParams: false),
 (FuncName: 'Stdev'; iftab: 12; VarParams: false),
 (FuncName: 'Dollar'; iftab: 13; VarParams: false),
 (FuncName: 'Fixed'; iftab: 14; VarParams: false),
 (FuncName: 'Sin'; iftab: 15; VarParams: false),
 (FuncName: 'Cos'; iftab: 16; VarParams: false),
 (FuncName: 'Tan'; iftab: 17; VarParams: false),
 (FuncName: 'Atan'; iftab: 18; VarParams: false),
 (FuncName: 'Pi'; iftab: 19; VarParams: false),
 (FuncName: 'Sqrt'; iftab: 20; VarParams: false),
 (FuncName: 'Exp'; iftab: 21; VarParams: false),
 (FuncName: 'Ln'; iftab: 22; VarParams: false),
 (FuncName: 'Log10'; iftab: 23; VarParams: false),
 (FuncName: 'Abs'; iftab: 24; VarParams: false),
 (FuncName: 'Int'; iftab: 25; VarParams: false),
 (FuncName: 'Sign'; iftab: 26; VarParams: false),
 (FuncName: 'Round'; iftab: 27; VarParams: false),
 (FuncName: 'Lookup'; iftab: 28; VarParams: false),
 (FuncName: 'Index'; iftab: 29; VarParams: false),
 (FuncName: 'Rept'; iftab: 30; VarParams: false),
 (FuncName: 'Mid'; iftab: 31; VarParams: false),
 (FuncName: 'Len'; iftab: 32; VarParams: false),
 (FuncName: 'Value'; iftab: 33; VarParams: false),
 (FuncName: 'True'; iftab: 34; VarParams: false),
 (FuncName: 'False'; iftab: 35; VarParams: false),
 (FuncName: 'And'; iftab: 36; VarParams: false),
 (FuncName: 'Or'; iftab: 37; VarParams: false),
 (FuncName: 'Not'; iftab: 38; VarParams: false),
 (FuncName: 'Mod'; iftab: 39; VarParams: false),
 (FuncName: 'Dcount'; iftab: 40; VarParams: false),
 (FuncName: 'Dsum'; iftab: 41; VarParams: false),
 (FuncName: 'Daverage'; iftab: 42; VarParams: false),
 (FuncName: 'Dmin'; iftab: 43; VarParams: false),
 (FuncName: 'Dmax'; iftab: 44; VarParams: false),
 (FuncName: 'Dstdev'; iftab: 45; VarParams: false),
 (FuncName: 'Var'; iftab: 46; VarParams: false),
 (FuncName: 'Dvar'; iftab: 47; VarParams: false),
 (FuncName: 'Text'; iftab: 48; VarParams: false),
 (FuncName: 'Linest'; iftab: 49; VarParams: false),
 (FuncName: 'Trend'; iftab: 50; VarParams: false),
 (FuncName: 'Logest'; iftab: 51; VarParams: false),
 (FuncName: 'Growth'; iftab: 52; VarParams: false),
 (FuncName: 'Goto'; iftab: 53; VarParams: false),
 (FuncName: 'Halt'; iftab: 54; VarParams: false),
 (FuncName: 'Pv'; iftab: 56; VarParams: false),
 (FuncName: 'Fv'; iftab: 57; VarParams: false),
 (FuncName: 'Nper'; iftab: 58; VarParams: false),
 (FuncName: 'Pmt'; iftab: 59; VarParams: false),
 (FuncName: 'Rate'; iftab: 60; VarParams: false),
 (FuncName: 'Mirr'; iftab: 61; VarParams: false),
 (FuncName: 'Irr'; iftab: 62; VarParams: false),
 (FuncName: 'Rand'; iftab: 63; VarParams: false),
 (FuncName: 'Match'; iftab: 64; VarParams: false),
 (FuncName: 'Date'; iftab: 65; VarParams: false),
 (FuncName: 'Time'; iftab: 66; VarParams: false),
 (FuncName: 'Day'; iftab: 67; VarParams: false),
 (FuncName: 'Month'; iftab: 68; VarParams: false),
 (FuncName: 'Year'; iftab: 69; VarParams: false),
 (FuncName: 'Weekday'; iftab: 70; VarParams: false),
 (FuncName: 'Hour'; iftab: 71; VarParams: false),
 (FuncName: 'Minute'; iftab: 72; VarParams: false),
 (FuncName: 'Second'; iftab: 73; VarParams: false),
 (FuncName: 'Now'; iftab: 74; VarParams: false),
 (FuncName: 'Areas'; iftab: 75; VarParams: false),
 (FuncName: 'Rows'; iftab: 76; VarParams: false),
 (FuncName: 'Columns'; iftab: 77; VarParams: false),
 (FuncName: 'Offset'; iftab: 78; VarParams: false),
 (FuncName: 'Absref'; iftab: 79; VarParams: false),
 (FuncName: 'Relref'; iftab: 80; VarParams: false),
 (FuncName: 'Argument'; iftab: 81; VarParams: false),
 (FuncName: 'Search'; iftab: 82; VarParams: false),
 (FuncName: 'Transpose'; iftab: 83; VarParams: false),
 (FuncName: 'Error'; iftab: 84; VarParams: false),
 (FuncName: 'Step'; iftab: 85; VarParams: false),
 (FuncName: 'Type'; iftab: 86; VarParams: false),
 (FuncName: 'Echo'; iftab: 87; VarParams: false),
 (FuncName: 'SetName'; iftab: 88; VarParams: false),
 (FuncName: 'Caller'; iftab: 89; VarParams: false),
 (FuncName: 'Deref'; iftab: 90; VarParams: false),
 (FuncName: 'Windows'; iftab: 91; VarParams: false),
 (FuncName: 'Series'; iftab: 92; VarParams: false),
 (FuncName: 'Documents'; iftab: 93; VarParams: false),
 (FuncName: 'ActiveCell'; iftab: 94; VarParams: false),
 (FuncName: 'Selection'; iftab: 95; VarParams: false),
 (FuncName: 'Result'; iftab: 96; VarParams: false),
 (FuncName: 'Atan2'; iftab: 97; VarParams: false),
 (FuncName: 'Asin'; iftab: 98; VarParams: false),
 (FuncName: 'Acos'; iftab: 99; VarParams: false),
 (FuncName: 'Choose'; iftab: 100; VarParams: false),
 (FuncName: 'Hlookup'; iftab: 101; VarParams: false),
 (FuncName: 'Vlookup'; iftab: 102; VarParams: false),
 (FuncName: 'Links'; iftab: 103; VarParams: false),
 (FuncName: 'Input'; iftab: 104; VarParams: false),
 (FuncName: 'Isref'; iftab: 105; VarParams: false),
 (FuncName: 'GetFormula'; iftab: 106; VarParams: false),
 (FuncName: 'GetName'; iftab: 107; VarParams: false),
 (FuncName: 'SetValue'; iftab: 108; VarParams: false),
 (FuncName: 'Log'; iftab: 109; VarParams: false),
 (FuncName: 'Exec'; iftab: 110; VarParams: false),
 (FuncName: 'Char'; iftab: 111; VarParams: false),
 (FuncName: 'Lower'; iftab: 112; VarParams: false),
 (FuncName: 'Upper'; iftab: 113; VarParams: false),
 (FuncName: 'Proper'; iftab: 114; VarParams: false),
 (FuncName: 'Left'; iftab: 115; VarParams: false),
 (FuncName: 'Right'; iftab: 116; VarParams: false),
 (FuncName: 'Exact'; iftab: 117; VarParams: false),
 (FuncName: 'Trim'; iftab: 118; VarParams: false),
 (FuncName: 'Replace'; iftab: 119; VarParams: false),
 (FuncName: 'Substitute'; iftab: 120; VarParams: false),
 (FuncName: 'Code'; iftab: 121; VarParams: false),
 (FuncName: 'Names'; iftab: 122; VarParams: false),
 (FuncName: 'Directory'; iftab: 123; VarParams: false),
 (FuncName: 'Find'; iftab: 124; VarParams: false),
 (FuncName: 'Cell'; iftab: 125; VarParams: false),
 (FuncName: 'Iserr'; iftab: 126; VarParams: false),
 (FuncName: 'Istext'; iftab: 127; VarParams: false),
 (FuncName: 'Isnumber'; iftab: 128; VarParams: false),
 (FuncName: 'Isblank'; iftab: 129; VarParams: false),
 (FuncName: 'T'; iftab: 130; VarParams: false),
 (FuncName: 'N'; iftab: 131; VarParams: false),
 (FuncName: 'Fopen'; iftab: 132; VarParams: false),
 (FuncName: 'Fclose'; iftab: 133; VarParams: false),
 (FuncName: 'Fsize'; iftab: 134; VarParams: false),
 (FuncName: 'Freadln'; iftab: 135; VarParams: false),
 (FuncName: 'Fread'; iftab: 136; VarParams: false),
 (FuncName: 'Fwriteln'; iftab: 137; VarParams: false),
 (FuncName: 'Fwrite'; iftab: 138; VarParams: false),
 (FuncName: 'Fpos'; iftab: 139; VarParams: false),
 (FuncName: 'Datevalue'; iftab: 140; VarParams: false),
 (FuncName: 'Timevalue'; iftab: 141; VarParams: false),
 (FuncName: 'Sln'; iftab: 142; VarParams: false),
 (FuncName: 'Syd'; iftab: 143; VarParams: false),
 (FuncName: 'Ddb'; iftab: 144; VarParams: false),
 (FuncName: 'GetDef'; iftab: 145; VarParams: false),
 (FuncName: 'Reftext'; iftab: 146; VarParams: false),
 (FuncName: 'Textref'; iftab: 147; VarParams: false),
 (FuncName: 'Indirect'; iftab: 148; VarParams: false),
 (FuncName: 'Register'; iftab: 149; VarParams: false),
 (FuncName: 'Call'; iftab: 150; VarParams: false),
 (FuncName: 'AddBar'; iftab: 151; VarParams: false),
 (FuncName: 'AddMenu'; iftab: 152; VarParams: false),
 (FuncName: 'AddCommand'; iftab: 153; VarParams: false),
 (FuncName: 'EnableCommand'; iftab: 154; VarParams: false),
 (FuncName: 'CheckCommand'; iftab: 155; VarParams: false),
 (FuncName: 'RenameCommand'; iftab: 156; VarParams: false),
 (FuncName: 'ShowBar'; iftab: 157; VarParams: false),
 (FuncName: 'DeleteMenu'; iftab: 158; VarParams: false),
 (FuncName: 'DeleteCommand'; iftab: 159; VarParams: false),
 (FuncName: 'GetChartItem'; iftab: 160; VarParams: false),
 (FuncName: 'DialogBox'; iftab: 161; VarParams: false),
 (FuncName: 'Clean'; iftab: 162; VarParams: false),
 (FuncName: 'Mdeterm'; iftab: 163; VarParams: false),
 (FuncName: 'Minverse'; iftab: 164; VarParams: false),
 (FuncName: 'Mmult'; iftab: 165; VarParams: false),
 (FuncName: 'Files'; iftab: 166; VarParams: false),
 (FuncName: 'Ipmt'; iftab: 167; VarParams: false),
 (FuncName: 'Ppmt'; iftab: 168; VarParams: false),
 (FuncName: 'Counta'; iftab: 169; VarParams: false),
 (FuncName: 'CancelKey'; iftab: 170; VarParams: false),
 (FuncName: 'Initiate'; iftab: 175; VarParams: false),
 (FuncName: 'Request'; iftab: 176; VarParams: false),
 (FuncName: 'Poke'; iftab: 177; VarParams: false),
 (FuncName: 'Execute'; iftab: 178; VarParams: false),
 (FuncName: 'Terminate'; iftab: 179; VarParams: false),
 (FuncName: 'Restart'; iftab: 180; VarParams: false),
 (FuncName: 'Help'; iftab: 181; VarParams: false),
 (FuncName: 'GetBar'; iftab: 182; VarParams: false),
 (FuncName: 'Product'; iftab: 183; VarParams: false),
 (FuncName: 'Fact'; iftab: 184; VarParams: false),
 (FuncName: 'GetCell'; iftab: 185; VarParams: false),
 (FuncName: 'GetWorkspace'; iftab: 186; VarParams: false),
 (FuncName: 'GetWindow'; iftab: 187; VarParams: false),
 (FuncName: 'GetDocument'; iftab: 188; VarParams: false),
 (FuncName: 'Dproduct'; iftab: 189; VarParams: false),
 (FuncName: 'Isnontext'; iftab: 190; VarParams: false),
 (FuncName: 'GetNote'; iftab: 191; VarParams: false),
 (FuncName: 'Note'; iftab: 192; VarParams: false),
 (FuncName: 'Stdevp'; iftab: 193; VarParams: false),
 (FuncName: 'Varp'; iftab: 194; VarParams: false),
 (FuncName: 'Dstdevp'; iftab: 195; VarParams: false),
 (FuncName: 'Dvarp'; iftab: 196; VarParams: false),
 (FuncName: 'Trunc'; iftab: 197; VarParams: false),
 (FuncName: 'Islogical'; iftab: 198; VarParams: false),
 (FuncName: 'Dcounta'; iftab: 199; VarParams: false),
 (FuncName: 'DeleteBar'; iftab: 200; VarParams: false),
 (FuncName: 'Unregister'; iftab: 201; VarParams: false),
 (FuncName: 'Usdollar'; iftab: 204; VarParams: false),
 (FuncName: 'Findb'; iftab: 205; VarParams: false),
 (FuncName: 'Searchb'; iftab: 206; VarParams: false),
 (FuncName: 'Replaceb'; iftab: 207; VarParams: false),
 (FuncName: 'Leftb'; iftab: 208; VarParams: false),
 (FuncName: 'Rightb'; iftab: 209; VarParams: false),
 (FuncName: 'Midb'; iftab: 210; VarParams: false),
 (FuncName: 'Lenb'; iftab: 211; VarParams: false),
 (FuncName: 'Roundup'; iftab: 212; VarParams: false),
 (FuncName: 'Rounddown'; iftab: 213; VarParams: false),
 (FuncName: 'Asc'; iftab: 214; VarParams: false),
 (FuncName: 'Dbcs'; iftab: 215; VarParams: false),
 (FuncName: 'Rank'; iftab: 216; VarParams: false),
 (FuncName: 'Address'; iftab: 219; VarParams: false),
 (FuncName: 'Days360'; iftab: 220; VarParams: false),
 (FuncName: 'Today'; iftab: 221; VarParams: false),
 (FuncName: 'Vdb'; iftab: 222; VarParams: false),
 (FuncName: 'Median'; iftab: 227; VarParams: false),
 (FuncName: 'Sumproduct'; iftab: 228; VarParams: false),
 (FuncName: 'Sinh'; iftab: 229; VarParams: false),
 (FuncName: 'Cosh'; iftab: 230; VarParams: false),
 (FuncName: 'Tanh'; iftab: 231; VarParams: false),
 (FuncName: 'Asinh'; iftab: 232; VarParams: false),
 (FuncName: 'Acosh'; iftab: 233; VarParams: false),
 (FuncName: 'Atanh'; iftab: 234; VarParams: false),
 (FuncName: 'Dget'; iftab: 235; VarParams: false),
 (FuncName: 'CreateObject'; iftab: 236; VarParams: false),
 (FuncName: 'Volatile'; iftab: 237; VarParams: false),
 (FuncName: 'LastError'; iftab: 238; VarParams: false),
 (FuncName: 'CustomUndo'; iftab: 239; VarParams: false),
 (FuncName: 'CustomRepeat'; iftab: 240; VarParams: false),
 (FuncName: 'FormulaConvert'; iftab: 241; VarParams: false),
 (FuncName: 'GetLinkInfo'; iftab: 242; VarParams: false),
 (FuncName: 'TextBox'; iftab: 243; VarParams: false),
 (FuncName: 'Info'; iftab: 244; VarParams: false),
 (FuncName: 'Group'; iftab: 245; VarParams: false),
 (FuncName: 'GetObject'; iftab: 246; VarParams: false),
 (FuncName: 'Db'; iftab: 247; VarParams: false),
 (FuncName: 'Pause'; iftab: 248; VarParams: false),
 (FuncName: 'Resume'; iftab: 251; VarParams: false),
 (FuncName: 'Frequency'; iftab: 252; VarParams: false),
 (FuncName: 'AddToolbar'; iftab: 253; VarParams: false),
 (FuncName: 'DeleteToolbar'; iftab: 254; VarParams: false),
 (FuncName: 'ResetToolbar'; iftab: 256; VarParams: false),
 (FuncName: 'Evaluate'; iftab: 257; VarParams: false),
 (FuncName: 'GetToolbar'; iftab: 258; VarParams: false),
 (FuncName: 'GetTool'; iftab: 259; VarParams: false),
 (FuncName: 'SpellingCheck'; iftab: 260; VarParams: false),
 (FuncName: 'ErrorType'; iftab: 261; VarParams: false),
 (FuncName: 'AppTitle'; iftab: 262; VarParams: false),
 (FuncName: 'WindowTitle'; iftab: 263; VarParams: false),
 (FuncName: 'SaveToolbar'; iftab: 264; VarParams: false),
 (FuncName: 'EnableTool'; iftab: 265; VarParams: false),
 (FuncName: 'PressTool'; iftab: 266; VarParams: false),
 (FuncName: 'RegisterId'; iftab: 267; VarParams: false),
 (FuncName: 'GetWorkbook'; iftab: 268; VarParams: false),
 (FuncName: 'Avedev'; iftab: 269; VarParams: false),
 (FuncName: 'Betadist'; iftab: 270; VarParams: false),
 (FuncName: 'Gammaln'; iftab: 271; VarParams: false),
 (FuncName: 'Betainv'; iftab: 272; VarParams: false),
 (FuncName: 'Binomdist'; iftab: 273; VarParams: false),
 (FuncName: 'Chidist'; iftab: 274; VarParams: false),
 (FuncName: 'Chiinv'; iftab: 275; VarParams: false),
 (FuncName: 'Combin'; iftab: 276; VarParams: false),
 (FuncName: 'Confidence'; iftab: 277; VarParams: false),
 (FuncName: 'Critbinom'; iftab: 278; VarParams: false),
 (FuncName: 'Even'; iftab: 279; VarParams: false),
 (FuncName: 'Expondist'; iftab: 280; VarParams: false),
 (FuncName: 'Fdist'; iftab: 281; VarParams: false),
 (FuncName: 'Finv'; iftab: 282; VarParams: false),
 (FuncName: 'Fisher'; iftab: 283; VarParams: false),
 (FuncName: 'Fisherinv'; iftab: 284; VarParams: false),
 (FuncName: 'Floor'; iftab: 285; VarParams: false),
 (FuncName: 'Gammadist'; iftab: 286; VarParams: false),
 (FuncName: 'Gammainv'; iftab: 287; VarParams: false),
 (FuncName: 'Ceiling'; iftab: 288; VarParams: false),
 (FuncName: 'Hypgeomdist'; iftab: 289; VarParams: false),
 (FuncName: 'Lognormdist'; iftab: 290; VarParams: false),
 (FuncName: 'Loginv'; iftab: 291; VarParams: false),
 (FuncName: 'Negbinomdist'; iftab: 292; VarParams: false),
 (FuncName: 'Normdist'; iftab: 293; VarParams: false),
 (FuncName: 'Normsdist'; iftab: 294; VarParams: false),
 (FuncName: 'Norminv'; iftab: 295; VarParams: false),
 (FuncName: 'Normsinv'; iftab: 296; VarParams: false),
 (FuncName: 'Standardize'; iftab: 297; VarParams: false),
 (FuncName: 'Odd'; iftab: 298; VarParams: false),
 (FuncName: 'Permut'; iftab: 299; VarParams: false),
 (FuncName: 'Poisson'; iftab: 300; VarParams: false),
 (FuncName: 'Tdist'; iftab: 301; VarParams: false),
 (FuncName: 'Weibull'; iftab: 302; VarParams: false),
 (FuncName: 'Sumxmy2'; iftab: 303; VarParams: false),
 (FuncName: 'Sumx2my2'; iftab: 304; VarParams: false),
 (FuncName: 'Sumx2py2'; iftab: 305; VarParams: false),
 (FuncName: 'Chitest'; iftab: 306; VarParams: false),
 (FuncName: 'Correl'; iftab: 307; VarParams: false),
 (FuncName: 'Covar'; iftab: 308; VarParams: false),
 (FuncName: 'Forecast'; iftab: 309; VarParams: false),
 (FuncName: 'Ftest'; iftab: 310; VarParams: false),
 (FuncName: 'Intercept'; iftab: 311; VarParams: false),
 (FuncName: 'Pearson'; iftab: 312; VarParams: false),
 (FuncName: 'Rsq'; iftab: 313; VarParams: false),
 (FuncName: 'Steyx'; iftab: 314; VarParams: false),
 (FuncName: 'Slope'; iftab: 315; VarParams: false),
 (FuncName: 'Ttest'; iftab: 316; VarParams: false),
 (FuncName: 'Prob'; iftab: 317; VarParams: false),
 (FuncName: 'Devsq'; iftab: 318; VarParams: false),
 (FuncName: 'Geomean'; iftab: 319; VarParams: false),
 (FuncName: 'Harmean'; iftab: 320; VarParams: false),
 (FuncName: 'Sumsq'; iftab: 321; VarParams: false),
 (FuncName: 'Kurt'; iftab: 322; VarParams: false),
 (FuncName: 'Skew'; iftab: 323; VarParams: false),
 (FuncName: 'Ztest'; iftab: 324; VarParams: false),
 (FuncName: 'Large'; iftab: 325; VarParams: false),
 (FuncName: 'Small'; iftab: 326; VarParams: false),
 (FuncName: 'Quartile'; iftab: 327; VarParams: false),
 (FuncName: 'Percentile'; iftab: 328; VarParams: false),
 (FuncName: 'Percentrank'; iftab: 329; VarParams: false),
 (FuncName: 'Mode'; iftab: 330; VarParams: false),
 (FuncName: 'Trimmean'; iftab: 331; VarParams: false),
 (FuncName: 'Tinv'; iftab: 332; VarParams: false),
 (FuncName: 'MovieCommand'; iftab: 334; VarParams: false),
 (FuncName: 'GetMovie'; iftab: 335; VarParams: false),
 (FuncName: 'Concatenate'; iftab: 336; VarParams: false),
 (FuncName: 'Power'; iftab: 337; VarParams: false),
 (FuncName: 'PivotAddData'; iftab: 338; VarParams: false),
 (FuncName: 'GetPivotTable'; iftab: 339; VarParams: false),
 (FuncName: 'GetPivotField'; iftab: 340; VarParams: false),
 (FuncName: 'GetPivotItem'; iftab: 341; VarParams: false),
 (FuncName: 'Radians'; iftab: 342; VarParams: false),
 (FuncName: 'Degrees'; iftab: 343; VarParams: false),
 (FuncName: 'Subtotal'; iftab: 344; VarParams: false),
 (FuncName: 'Sumif'; iftab: 345; VarParams: false),
 (FuncName: 'Countif'; iftab: 346; VarParams: false),
 (FuncName: 'Countblank'; iftab: 347; VarParams: false),
 (FuncName: 'ScenarioGet'; iftab: 348; VarParams: false),
 (FuncName: 'OptionsListsGet'; iftab: 349; VarParams: false),
 (FuncName: 'Ispmt'; iftab: 350; VarParams: false),
 (FuncName: 'Datedif'; iftab: 351; VarParams: false),
 (FuncName: 'Datestring'; iftab: 352; VarParams: false),
 (FuncName: 'Numberstring'; iftab: 353; VarParams: false),
 (FuncName: 'Roman'; iftab: 354; VarParams: false),
 (FuncName: 'OpenDialog'; iftab: 355; VarParams: false),
 (FuncName: 'SaveDialog'; iftab: 356; VarParams: false));

implementation

end.
