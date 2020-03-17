@echo off

set prversion=%1
call Settings.bat

echo *********************************************************
echo
echo (bat file works only on NT/2000/XP):
echo
echo   %src%.zip - GridReport sources (Source and Res)
echo
echo *********************************************************
pause

md %dpup%
rmdir %dppr% /s /q
md %dppr%

rem *********************************************
rem Формируем каталог
rem *********************************************
md %dppr%\CommonControls
md %dppr%\CommonControls\Source
md %dppr%\GridReport
md %dppr%\GridReport\Source
md %dppr%\GridReport\RES
md %dppr%\GridReport\Demos

rem *********************************************
rem Sources
rem *********************************************
copy %prcat%\source\*.pas %dppr%\GridReport\Source
copy %prcat%\source\*.res %dppr%\GridReport\Source
copy %prcat%\source\*.bpk %dppr%\GridReport\Source
copy %prcat%\source\*.dfm %dppr%\GridReport\Source
copy %prcat%\source\*.inc %dppr%\GridReport\Source
copy %prcat%\source\*.dpk %dppr%\GridReport\Source
copy %prcat%\source\*.cpp %dppr%\GridReport\Source
copy %prcat%\source\*.inc %dppr%\GridReport\Source

rem *********************************************
rem Resources
rem *********************************************
call batch\CopyRes.bat

rem *********************************************
rem Demos
rem *********************************************
call batch\CopyDemos.bat

rem *********************************************
rem Addititional files
rem *********************************************
copy %prcat%\..\GridReport_documentation\readme.txt %dppr%
copy %prcat%\..\GridReport_documentation\GridReportEULA.txt %dppr%

rem *********************************************
rem CommonControls
rem *********************************************
copy %commoncontrols%\*.pas %dppr%\CommonControls\Source
copy %commoncontrols%\*.dfm %dppr%\CommonControls\Source
copy %commoncontrols%\*.dpk %dppr%\CommonControls\Source
copy %commoncontrols%\*.res %dppr%\CommonControls\Source
copy %commoncontrols%\*.bpk %dppr%\CommonControls\Source
copy %commoncontrols%\*.cpp %dppr%\CommonControls\Source
copy %commoncontrols%\*.inc %dppr%\CommonControls\Source

rem *********************************************
rem Создаем архив
rem *********************************************

%wz% -add -max -dir=relative %dpup%\%src%.zip %dppr%\*.*

rmdir %dppr% /q /s
