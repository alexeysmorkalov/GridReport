@echo on

for %%f in (*.rc) do brcc32 %%f

dcc32 vgr_RUS.dpr

move /y *.res ..\
