rem *********************************************
rem Usage:
rem   CopyDemo.bat <директория источник>, <директория приемник>
rem *********************************************

set demosrc=%1
set demodst=%2

if exist %demosrc%\nul (
mkdir %demodst%

if exist %demosrc%\cbuilder\nul (
mkdir %demodst%\cbuilder
copy %demosrc%\cbuilder\*.cpp %demodst%\cbuilder
copy %demosrc%\cbuilder\*.dfm %demodst%\cbuilder
copy %demosrc%\cbuilder\*.res %demodst%\cbuilder
copy %demosrc%\cbuilder\*.h   %demodst%\cbuilder
copy %demosrc%\cbuilder\*.ini %demodst%\cbuilder
copy %demosrc%\cbuilder\*.bpr %demodst%\cbuilder
copy %demosrc%\cbuilder\*.grt %demodst%\cbuilder
copy %demosrc%\cbuilder\*.grw %demodst%\cbuilder
copy %demosrc%\cbuilder\*.dbf %demodst%\cbuilder
copy %demosrc%\cbuilder\readme.txt %demodst%\cbuilder
if exist %demosrc%\cbuilder\Templates\nul (
mkdir %demodst%\cbuilder\Templates
pause
copy %demosrc%\cbuilder\Templates\*.grt %demodst%\cbuilder\Templates))


if exist %demosrc%\delphi\nul (
mkdir %demodst%\delphi
copy %demosrc%\delphi\*.pas %demodst%\delphi
copy %demosrc%\delphi\*.dfm %demodst%\delphi
copy %demosrc%\delphi\*.res %demodst%\delphi
copy %demosrc%\delphi\*.ini %demodst%\delphi
copy %demosrc%\delphi\*.dpr %demodst%\delphi
copy %demosrc%\delphi\*.grt %demodst%\delphi
copy %demosrc%\delphi\*.grw %demodst%\delphi
copy %demosrc%\delphi\*.dbf %demodst%\delphi
copy %demosrc%\delphi\readme.txt %demodst%\delphi
if exist %demosrc%\delphi\Templates\nul (
mkdir %demodst%\delphi\Templates
copy %demosrc%\delphi\Templates\*.grt %demodst%\delphi\Templates)))

