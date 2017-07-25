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

Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZSD_SHP"
Session.findById("wnd[0]").sendVKey 0

'Get current working directory
Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
Set fso = Nothing

'Text file paths
exportpath = Replace(CurrentDirectory,"Scripts","Data")
PNPath = CurrentDirectory & "\PartNumbers.txt"

'Set export file name
exportfilename = "ZSD_SHP.txt"


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

Session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_VTWEG-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_VTWEG-HIGH").Text = ""
Session.findById("wnd[0]/usr/ctxtS_SPART-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_SPART-HIGH").Text = ""
Session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_CNTRCT-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_WADATE-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_TKNUM-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtS_CODE-LOW").Text = ""
Session.findById("wnd[0]/usr/txtS_TRACK-LOW").Text = ""

Session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").Text = "4475"
Session.findById("wnd[0]/usr/ctxtS_VTWEG-LOW").Text = "01"
Session.findById("wnd[0]/usr/ctxtS_VTWEG-HIGH").Text = "04"
Session.findById("wnd[0]/usr/ctxtS_SPART-LOW").Text = "10"
Session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = "ZOR"

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

Session.findById("wnd[0]/usr/ctxtS_WADATE-LOW").Text = Mnth & "/01/" & yr
Session.findById("wnd[0]/usr/ctxtS_WADATE-HIGH").Text = Cstr(Date)


Session.findById("wnd[0]/usr/ctxtS_WADATE-HIGH").SetFocus
Session.findById("wnd[0]/usr/ctxtS_WADATE-HIGH").caretPosition = 10
Session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press


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
Session.findById("wnd[0]/tbar[1]/btn[8]").press


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell


Session.findById("wnd[0]/tbar[1]/btn[45]").press
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
Session.findById("wnd[1]/tbar[0]/btn[0]").press

Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = exportpath
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportfilename

Session.findById("wnd[1]/usr/ctxtDY_FILENAME").SetFocus
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
Session.findById("wnd[1]/tbar[0]/btn[11]").press
