If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

VarName = "SMARTMB52"

session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb52"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/chkNOZERO").selected = false
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "30279-0301"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "2475"
session.findById("wnd[0]/usr/chkNOZERO").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/tbar[1]/btn[32]").press
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(3).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(4).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(5).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(6).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(9).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(10).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(11).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").verticalScrollbar.position = 12
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(12).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(13).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(14).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(15).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(16).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(17).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(19).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(20).selected = true
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").getAbsoluteRow(21).selected = true

session.findById("wnd[1]/usr/btnAPP_FL_SING").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[1]/usr/ctxtLTDX-VARIANT").text = VarName
session.findById("wnd[1]/usr/txtLTDXT-TEXT").text = VarName
session.findById("wnd[1]/usr/txtLTDXT-TEXT").setFocus
session.findById("wnd[1]/usr/txtLTDXT-TEXT").caretPosition = 9
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
