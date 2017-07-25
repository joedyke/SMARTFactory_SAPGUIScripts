@echo off 
	:LOOPSTART
		CMD /C Start /wait /d "%~dp0" COOIS_HEADERS_OPEN.vbs
		CMD /C Start /wait /d "%~dp0" COOIS_HEADERS_DLV.vbs
		CMD /C Start /wait /d "%~dp0" ZSC_ZPPWIP_SN.vbs
		CMD /c start /wait /d "%~dp0" ZSD_SDEL.vbs
		CMD /C Start /wait /d "%~dp0" ZSD_SHP.vbs
		CMD /C Start /wait /d "%~dp0" COOIS_KIT_DLV.vbs
		CMD /C Start /wait /d "%~dp0" MB52.vbs
		CMD /C Start /wait /d "%~dp0" COOIS_Operations.vbs

		Rem have the script pause to give the cpu a break (pause for 30 seconds)
		timeout /t 30 /nobreak
	GOTO LOOPSTART
:LOOPEND
