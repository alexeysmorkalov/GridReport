if not exist ..\BldHelp mkdir ..\BldHelp
del ..\BldHelp\*.chm
del ..\BldHelp\*.hhc
del ..\BldHelp\*.hhk
del ..\BldHelp\*.hhp
del ..\BldHelp\*.htm*
del ..\BldHelp\*.gif
del ..\BldHelp\*.log
dipas_console_m.exe -Mprotected- -OHtml -E..\BldHelp -I..\GridReport\Source -Le %1 
start ..\BldHelp\AllUnits.htm

