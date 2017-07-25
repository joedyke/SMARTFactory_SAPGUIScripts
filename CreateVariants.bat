
@echo off 
CMD /C Start /wait /d "%~dp0" CreateVariantCOOIS_Dlv_KitTimes.vbs
CMD /C Start /wait /d "%~dp0" CreateVariantCOOIS_Headers.vbs
CMD /C Start /wait /d "%~dp0" CreateVariantCOOIS_Operations.vbs
CMD /C Start /wait /d "%~dp0" CreateVariantMB52.vbs
CMD /c start /wait /d "%~dp0" CreateVariantZSC_ZPPWIP_SN.vbs
CMD /C Start /wait /d "%~dp0" CreateVariantZSD_SDEL.vbs
CMD /C Start /wait /d "%~dp0" CreateVariantZSD_SHP.vbs

