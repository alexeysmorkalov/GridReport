rem *********************************************
rem Usage:
rem   BuildDelphi.bat <ЗначениеДляКлючаLU>
rem Должна быть установлена переменная dcat, на
rem каталог с нужной версией  Delphi
rem *********************************************

set lukey=%1
set define=%2

rem Directory with sources
set inc=%prcat%\Source

rem UNITS directories
set units=%dcat%\Lib;%dcat%\Projects\Bpl;%commoncontrols%;%prcat%\Source

rem Directory with resources
set res=%prcat%\RES

rem Directory with compiler
set cmp=%dcat%\Bin

rem Delete files
del %prcat%\Source\*.dcu
del %prcat%\Source\*.obj
del %prcat%\Source\*.hpp

for %%f in (%prcat%\Source\*.pas) do "%cmp%\dcc32.exe" -D%define% -LU%lukey%  -$D- -$L- -$Y- -B -R%res% -I"%inc%" -U"%units%" "%%f"
