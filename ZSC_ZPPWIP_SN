If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(1)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzsc_zppwip_sn"
Session.findById("wnd[0]").sendVKey 0

'Get current working directory
Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
Set fso = Nothing

'Text file paths
exportpath = Replace(CurrentDirectory,"Scripts","Data")
PNPath = CurrentDirectory & "\PartNumbers.txt"


'Set export file name
exportfilename = "ZSC_ZPPWIP_SN.txt"

VarName = "SMARTSN"


'Gather PN list from text file
'read text file
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(PNPath)
'Extract PNs
PNs = objFileToRead.ReadAll()
'close .txt file
objFileToRead.Close
Set objFileToRead = Nothing
Wscript.Sleep 200

'Split PNs list on newline char into an array
PNsArray = Split(PNs,vbCrlf)

Session.findById("wnd[0]/usr/txtSP$00011-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00002-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00010-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00003-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00004-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00005-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00006-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00007-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00008-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxt%LAYOUT").Text = ""

session.findById("wnd[0]/usr/ctxt%LAYOUT").text = VarName
Session.findById("wnd[0]/usr/ctxtSP$00009-LOW").Text = "2475"

Session.findById("wnd[0]/usr/btn%_SP$00011_%_APP_%-VALU_PUSH").press


i = 0
j = 1
'Loop over every part number and send it individually to SAP
For each element in PNsArray
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1," & i & "]").text = element
	
	If i = 7 Then
		session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = j
		i = 0
	End if
	i = i + 1
	j = j+1
	
Next

session.findById("wnd[1]/tbar[0]/btn[8]").press


Session.findById("wnd[0]/usr/txtSP$00002-LOW").Text = ""
Session.findById("wnd[0]/usr/txt%_SP$00009_%_APP_%-TEXT").SetFocus
Session.findById("wnd[0]/usr/txt%_SP$00009_%_APP_%-TEXT").caretPosition = 5

Session.findById("wnd[0]/tbar[1]/btn[8]").press




Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = exportpath
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportfilename
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
Session.findById("wnd[1]/tbar[0]/btn[11]").press
