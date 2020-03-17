@echo on

for %%f in (*.rc) do brcc32 %%f

dcc32 vgr_ENG.dpr

move /y *.res ..\
