@echo off

call Settings.bat

rem *********************************************
rem Recreate temporary directory
rem *********************************************
rmdir %dppr% /s /q
md %dppr%

rem *********************************************
rem Copy addititional HTML files
rem *********************************************
copy %prhtml%\*.* %dppr%

rem *********************************************
rem Build CHM
rem *********************************************
set curd=%CD%
cd %dppr%
start /wait %dipas% -OHtmlHelp -I%prcat%\Source %prcat%\Source\*.pas -Le -XHomePage -Mdefaultvisibility+ -Mprivate- -Mprotected- -NGridReport -F%prhtml%\footer -H%prhtml%\header -Askipempty
cd %curd%

rem *********************************************
rem Copy CHM to destination directory
rem *********************************************

@echo Подождите пока закроются все всплывающие окна и нажмите любую клавишу
pause

copy %dppr%\gridreport.chm %dpup%
move %dppr%\dipas_console_m.log %dpup%\GridReport_dipas_console_m.log

%dppr% /s /q
