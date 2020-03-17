@echo off

set prversion=%1
call Settings.bat

echo *********************************************************
echo
echo (bat file works only on NT/2000/XP):
echo
echo   %d5%.zip - GridReport without sources
echo   %d6%.zip
echo   %d7%.zip
echo   %bcb5%.zip
echo   %bcb6%.zip
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
md %dppr%\GridReport\RES
md %dppr%\GridReport\Demos

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
rem Delphi
rem *********************************************

set dcat=%d5cat%
rmdir %dppr%\GridReport\Source /q /s
md %dppr%\GridReport\Source
call batch\BuildDelphi.bat vcl50 VTK_NOLIC
call batch\CopyDelphiDCU.bat %dppr%\GridReport\Source 5
%wz% -add -max -dir=relative %dpup%\%d5%.zip %dppr%\*.*

set dcat=%d6cat%
rmdir %dppr%\GridReport\Source /q /s
md %dppr%\GridReport\Source
call batch\BuildDelphi.bat DesignIde VTK_NOLIC
call batch\CopyDelphiDCU.bat %dppr%\GridReport\Source 6
%wz% -add -max -dir=relative %dpup%\%d6%.zip %dppr%\*.*

set dcat=%d7cat%
rmdir %dppr%\GridReport\Source /q /s
md %dppr%\GridReport\Source
call batch\BuildDelphi.bat DesignIde VTK_NOLIC
call batch\CopyDelphiDCU.bat %dppr%\GridReport\Source 7
%wz% -add -max -dir=relative %dpup%\%d7%.zip %dppr%\*.*

rem *********************************************
rem Builder
rem *********************************************

set dcat=%bcb5cat%
rmdir %dppr%\GridReport\Source /q /s
md %dppr%\GridReport\Source
call batch\BuildBuilder.bat vcl50 VTK_NOLIC
call batch\CopyBuilderDCU.bat %dppr%\GridReport\Source 5
%wz% -add -max -dir=relative %dpup%\%bcb5%.zip %dppr%\*.*

set dcat=%bcb6cat%
rmdir %dppr%\GridReport\Source /q /s
md %dppr%\GridReport\Source
call batch\BuildBuilder.bat DesignIde VTK_NOLIC
call batch\CopyBuilderDCU.bat %dppr%\GridReport\Source 6
%wz% -add -max -dir=relative %dpup%\%bcb6%.zip %dppr%\*.*

rmdir %dppr% /q /s
