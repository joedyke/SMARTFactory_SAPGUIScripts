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

VarName = "SMARTSDEL"

session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_sdel"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSP$00006-LOW").text = "30279-0301"
session.findById("wnd[0]/usr/ctxtSP$00019-LOW").text = "5/1/2017"
session.findById("wnd[0]/usr/ctxtSP$00019-HIGH").text = "5/30/2017"
session.findById("wnd[0]/usr/ctxtSP$00019-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtSP$00019-HIGH").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Hide columns
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "VBELN"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "POSNR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZCUSTPO"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZPOITEM"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "LETZNR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "VKORG"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "VTWEG"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "SPART"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "CONS_ORDER"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "OBKNR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "OBZAE"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "KUNNR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "KUNAG"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ERDAT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZPARA"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZOCCUR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "HSDAT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "MATWA"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "VGBEL"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "VGPOS"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZSHIPNO"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZNETWR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZWAERK"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "NAME1"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZSOLD_TO_NAME"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "ZZNETPR"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn "WAERK"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell -1,"VBELN"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&COL_INV"




'Set order of columns
set coll1 = application.createGuiCollection
coll1.add "MATNR"
coll1.add "SERNR"
coll1.add "LFART"
coll1.add "LFIMG"
coll1.add "VRKME"
coll1.add "WERKS"
coll1.add "LGORT"
coll1.add "CHARG"
coll1.add "WADAT_IST"
coll1.add "ZZTRACK"
coll1.add "TNDR_TRKID"
coll1.add "ARKTX"

session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").columnOrder = coll1
set coll1 = nothing


'Save as variant
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&SAVE"


session.findById("wnd[1]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDX-VARIANT").text = VarName
session.findById("wnd[1]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").text = VarName
session.findById("wnd[1]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").setFocus
session.findById("wnd[1]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume Next
session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
