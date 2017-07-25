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

Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzsd_sdel"
Session.findById("wnd[0]").sendVKey 0

'Get current working directory
Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
Set fso = Nothing

'Text file paths
exportpath = Replace(CurrentDirectory,"Scripts","Data")
PNPath = CurrentDirectory & "\PartNumbers.txt"


'Set export file name
exportfilename = "ZSD_SDEL.txt"

'Layout name
VarName = "SMARTSDEL"


'Gather PN list from text file
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(PNPath)
'Extract PNs
PNs = objFileToRead.ReadAll()
'close .txt file
objFileToRead.Close
Set objFileToRead = Nothing
Wscript.Sleep 200

'Split PNs list on newline char into an array
PNsArray = Split(PNs,vbCrlf)

Session.findById("wnd[0]/usr/ctxtSP$00001-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00002-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00003-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00007-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00009-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00008-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00017-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00005-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00004-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00011-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00006-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00012-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00021-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00019-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00013-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00025-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00026-LOW").Text = ""
Session.findById("wnd[0]/usr/txtSP$00027-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxt%LAYOUT").Text = ""


Session.findById("wnd[0]/usr/btn%_SP$00006_%_APP_%-VALU_PUSH").press

i = 0
j = 1
'Loop over every part number and send it individually to SAP
For each element in PNsArray
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").text = element
	
	If i = 7 Then
		session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = j
		i = 0
	End if
	i = i + 1
	j = j+1
	
Next
session.findById("wnd[1]/tbar[0]/btn[8]").press



Session.findById("wnd[0]/usr/ctxtSP$00021-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtSP$00012-LOW").Text = "2475"


'set the start date to three months prior
If Month(Date) = "1" Then
	Mnth = "10"
	yr = CStr(Year(Date)-1)
ElseIf Month(Date) = "2" Then
	Mnth = "11"
	yr = CStr(Year(Date)-1)
ElseIf Month(Date) = "3" Then
	Mnth = "12"
	yr = CStr(Year(Date)-1)
Else 
	Mnth = Cstr(Month(Date) - 3)
	yr = CStr(Year(Date))
End If
Session.findById("wnd[0]/usr/ctxtSP$00019-LOW").Text = Mnth & "/01/" & yr

'End date is today
Session.findById("wnd[0]/usr/ctxtSP$00019-HIGH").Text = Cstr(Date)

'Enter layout
session.findById("wnd[0]/usr/ctxt%LAYOUT").text = VarName

Session.findById("wnd[0]/usr/ctxtSP$00019-LOW").SetFocus
Session.findById("wnd[0]/usr/ctxtSP$00019-LOW").caretPosition = 10
Session.findById("wnd[0]/tbar[1]/btn[8]").press



Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = exportpath
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportfilename
Session.findById("wnd[1]/tbar[0]/btn[11]").press
