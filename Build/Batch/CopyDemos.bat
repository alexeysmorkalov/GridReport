rem *********************************************
rem Usage:
rem   CopyDemos.bat
rem *********************************************

call batch\CopyDemo.bat %prcat%\Demos\ConsoleDemo %dppr%\GridReport\Demos\ConsoleDemo
call batch\CopyDemo.bat %prcat%\Demos\MainDemo %dppr%\GridReport\Demos\MainDemo
call batch\CopyDemo.bat %prcat%\Demos\SimpleCrossTab %dppr%\GridReport\Demos\SimpleCrossTab
call batch\CopyDemo.bat %prcat%\Demos\ShowDatasetInWorkbookGrid %dppr%\GridReport\Demos\ShowDatasetInWorkbookGrid
call batch\CopyDemo.bat %prcat%\Demos\GRView %dppr%\GridReport\Demos\GRView
call batch\CopyDemo.bat %prcat%\Demos\ScriptControlDemo %dppr%\GridReport\Demos\ScriptControlDemo
md %dppr%\GridReport\Demos\Tutorials
call batch\CopyDemo.bat %prcat%\Demos\Tutorials\FirstReport %dppr%\GridReport\Demos\Tutorials\FirstReport
call batch\CopyDemo.bat %prcat%\Demos\Tutorials\Groups %dppr%\GridReport\Demos\Tutorials\Groups
call batch\CopyDemo.bat %prcat%\Demos\Tutorials\MasterDetail %dppr%\GridReport\Demos\Tutorials\MasterDetail
