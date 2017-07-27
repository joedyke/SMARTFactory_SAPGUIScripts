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
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
Session.findById("wnd[0]/tbar[0]/btn[0]").press

'Get current working directory
Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
Set fso = Nothing

'Text file paths
exportpath = Replace(CurrentDirectory,"Scripts","Data")
PNPath = CurrentDirectory & "\PartNumbers.txt"


'layout name
VarName = "SMARTMB52"

'Set export file name
exportfilename = "MB52.txt"


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

'clear all fields (sometimes SAP preloads some values if this t-code was used recently)
Session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtMATART-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtMATKLA-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtEKGRUP-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtMFRPN-LOW").Text = ""

Session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "2475"
Session.findById("wnd[0]/usr/ctxtP_VARI").Text = ""
Session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtMATNR-LOW").caretPosition = 1
Session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press

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



Session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
Session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
Session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 4
Session.findById("wnd[0]/usr/chkPA_SOND").Selected = True
Session.findById("wnd[0]/usr/chkNEGATIV").Selected = False
Session.findById("wnd[0]/usr/chkNOVALUES").Selected = False
session.findById("wnd[0]/usr/chkXMCHB").selected = false
session.findById("wnd[0]/usr/chkNOZERO").selected = False
Session.findById("wnd[0]/usr/radPA_FLT").SetFocus
Session.findById("wnd[0]/usr/radPA_FLT").Select
session.findById("wnd[0]/usr/ctxtP_VARI").text = VarName

Session.findById("wnd[0]/tbar[1]/btn[8]").press


On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[45]").press
If Err.Number <> 0 Then
	Wscript.Quit
End If
On Error GoTo 0

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = exportpath
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportfilename
Session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus
Session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 0
Session.findById("wnd[1]/tbar[0]/btn[11]").press
