Attribute VB_Name = "script"
Sub GerarRemessas()

Dim datainicio, dataFinal, horaAtual
    
horaAtual = Format(Now(), "hh.mm")
datainicio = Format(Now() - 5, "yyyymmdd")
dataFinal = Format(Now() + 1, "yyyymmdd")

login = 0
If login = 1 Then
sap_login:
    Call SAP_Logon
End If

If Not IsObject(app) Then
    On Error GoTo sap_login
    Set SapGuiAuto = GetObject("SAPGUI")
    Set app = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
    Set connection = app.Children(0)
End If

If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If
    
'Inicio - Script
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0

' SP
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").caretPosition = 7
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = datainicio
session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").caretPosition = 2
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = dataFinal
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

'1027
If horaAtual < "16" Then
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
End If

'interior
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
    
'rj
If horaAtual < "17.45" Then
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
End If

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "6"
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

MsgBox ("Remessas geradas com sucesso!")

AtualizaLogs
    
End Sub

Sub AtualizaLogs()
    
Dim data
data = Format(Now(), "yyyymmdd")

login = 0
If login = 1 Then
sap_login:
    Call SAP_Logon
End If

If Not IsObject(app) Then
    On Error GoTo sap_login
    Set SapGuiAuto = GetObject("SAPGUI")
    Set app = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
    Set connection = app.Children(0)
End If

If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If

On Error Resume Next
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0
'Coleta as informações de pedido e peso
'SP
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press

session.findById("wnd[0]/usr/txtERNAM-LOW").Text = ""
session.findById("wnd[0]/usr/txtERNAM-LOW").SetFocus
session.findById("wnd[0]/usr/txtERNAM-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100F"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "100G"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "100I"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "150F"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "150G"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "150I"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "sp.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'Retira
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press

session.findById("wnd[0]/usr/txtERNAM-LOW").Text = ""
session.findById("wnd[0]/usr/txtERNAM-LOW").SetFocus
session.findById("wnd[0]/usr/txtERNAM-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = ""

session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100b"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "100c"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "150b"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "150c"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "retira.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'Sabado
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press

session.findById("wnd[0]/usr/txtERNAM-LOW").Text = ""
session.findById("wnd[0]/usr/txtERNAM-LOW").SetFocus
session.findById("wnd[0]/usr/txtERNAM-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = ""

session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100j"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "150j"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "sabado.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'Loja
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press

session.findById("wnd[0]/usr/txtERNAM-LOW").Text = ""
session.findById("wnd[0]/usr/txtERNAM-LOW").SetFocus
session.findById("wnd[0]/usr/txtERNAM-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = ""

session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100h"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "150h"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "loja.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'RJ
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10A"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press

session.findById("wnd[0]/usr/txtERNAM-LOW").Text = ""
session.findById("wnd[0]/usr/txtERNAM-LOW").SetFocus
session.findById("wnd[0]/usr/txtERNAM-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = ""
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = ""

session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100d"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "100e"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "150d"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "150e"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = data
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "rj.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
    
AjustaDados

End Sub

Sub AjustaDados()

'Reseta e importa os dados - sp
Sheets("previa-sp").Select
Cells.Select
Selection.ClearContents
Range("A1").Select
On Error Resume Next

With ActiveSheet.QueryTables.Add(connection:= _
    "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\sp.txt", Destination:= _
    Range("$A$1"))
    .Name = "previa_1"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:3").Select
Selection.Delete Shift:=xlUp
Rows("2:2").Select
Selection.Delete Shift:=xlUp
Range("A1").Select
Columns("D:D").NumberFormat = "0"

'retira
Sheets("previa-retira").Select
Cells.Select
Selection.ClearContents
Range("A1").Select

With ActiveSheet.QueryTables.Add(connection:= _
    "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\retira.txt", Destination:=Range("$A$1"))
    .Name = "previa_1"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:3").Select
Selection.Delete Shift:=xlUp
Rows("2:2").Select
Selection.Delete Shift:=xlUp
Range("A1").Select
Columns("D:D").NumberFormat = "0"

'sabado
Sheets("previa-sabado").Select
Cells.Select
Selection.ClearContents
Range("A1").Select

With ActiveSheet.QueryTables.Add(connection:= _
    "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\sabado.txt", Destination:=Range("$A$1"))
    .Name = "previa_1"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:3").Select
Selection.Delete Shift:=xlUp
Rows("2:2").Select
Selection.Delete Shift:=xlUp
Range("A1").Select
Columns("D:D").NumberFormat = "0"

'RJ
Sheets("previa-rj").Select
Cells.Select
Selection.ClearContents
Range("A1").Select

With ActiveSheet.QueryTables.Add(connection:= _
    "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\rj.txt", Destination:= _
    Range("$A$1"))
    .Name = "previa_1"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:3").Select
Selection.Delete Shift:=xlUp
Rows("2:2").Select
Selection.Delete Shift:=xlUp
Range("A1").Select
Columns("D:D").NumberFormat = "0"

'lj
Sheets("previa-loja").Select
Cells.Select
Selection.ClearContents
Range("A1").Select

With ActiveSheet.QueryTables.Add(connection:= _
    "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\loja.txt", Destination:= _
    Range("$A$1"))
    .Name = "previa_1"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Rows("1:3").Select
Selection.Delete Shift:=xlUp
Rows("2:2").Select
Selection.Delete Shift:=xlUp
Range("A1").Select
Columns("D:D").NumberFormat = "0"
    
'extrai os dados
Sheets("previa").Select
Range("L1").FormulaR1C1 = "=SUM('previa-sabado'!C[-8])"
Range("M1").FormulaR1C1 = "=SUM('previa-sabado'!C[-6])"
Range("L2").FormulaR1C1 = "=SUM('previa-sp'!C[-8])"
Range("M2").FormulaR1C1 = "=SUM('previa-sp'!C[-6])"
Range("L3").FormulaR1C1 = "=SUM('previa-retira'!C[-8])"
Range("M3").FormulaR1C1 = "=SUM('previa-retira'!C[-6])"
Range("L4").FormulaR1C1 = "=SUM('previa-loja'!C[-8])"
Range("M4").FormulaR1C1 = "=SUM('previa-loja'!C[-6])"
Range("L5").FormulaR1C1 = "=SUM('previa-rj'!C[-8])"
Range("M5").FormulaR1C1 = "=SUM('previa-rj'!C[-6])"

MsgBox ("Prévia Atualizada!")
    
End Sub

