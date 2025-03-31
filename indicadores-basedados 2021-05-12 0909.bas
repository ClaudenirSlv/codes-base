Attribute VB_Name = "BaseDados"
Option Explicit

Sub BuscaOrdensServicos(strStDt As String, strEndDt As String)
    
    Call AccessTcode("IW72")
    Session.findById("wnd[0]/usr/chkDY_HIS").Selected = True
    Session.findById("wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "ZSGI"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "ZSGE"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "ZSFE"
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/ctxtERDAT-LOW").Text = strStDt
    Session.findById("wnd[0]/usr/ctxtERDAT-HIGH").Text = strEndDt
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/btn%_BUKRS_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1001"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1007"
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press 'Aqui chegamos na lista.
    
    Session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "iw72.txt"
    Session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub

Sub BuscaNotasServicos()
    
    'Copiar ordens da planilha ordens.
    '---------------------------------
    ln = wsOS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    wsOS.Range("A2:A" & ln).Copy
    
    Call AccessTcode("IW59")
    Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    Session.findById("wnd[0]/usr/cmbDY_PARVW").Key = " "
    Session.findById("wnd[0]/usr/ctxtDY_PARNR").Text = ""
    Session.findById("wnd[0]/usr/ctxtVARIANT").Text = "/cs-war"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "iw59.txt"
    Session.findById("wnd[1]/tbar[0]/btn[11]").press
        
End Sub

Sub BuscaNotasQAs()
    
    'Copiar as notas de serviços.
    '----------------------------
    ln = wsNS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    wsNS.Range("D2:D" & ln).Copy
    
    Call AccessTcode("QM11")
    Session.findById("wnd[0]/usr/ctxtQMDAT-LOW").showContextMenu
    Session.findById("wnd[0]/usr").selectContextMenuItem "DELACTX"
    Session.findById("wnd[0]").sendVKey 16
    Session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_" & _
        "CONTAINER:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
        
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]").sendVKey 8
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "qm11.txt"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    Session.findById("wnd[1]/tbar[0]/btn[11]").press
    
End Sub

Sub BuscaLancamentos(strStDt As String, strEndDt As String)
    
    'Copia as ordens de serviços.
    '----------------------------
    ln = LastRowNumber(wsOS)
    wsOS.Range("A2:A" & ln).Copy
    
    Call AccessTcode("KOB1")
    Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").showContextMenu
    Session.findById("wnd[0]/usr").selectContextMenuItem "DELACTX"
    Session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/" & _
        "ctxtRSCSEL_255-SLOW_I[1,0]").Text = "411075004"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/" & _
        "ctxtRSCSEL_255-SLOW_I[1,1]").Text = "411075007"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/" & _
        "ctxtRSCSEL_255-SLOW_I[1,2]").Text = "411075008"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/" & _
        "ctxtRSCSEL_255-SLOW_I[1,3]").Text = "411075083"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/" & _
        "ctxtRSCSEL_255-SLOW_I[1,4]").Text = "411075117"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/" & _
        "ctxtRSCSEL_255-SLOW_I[1,5]").Text = "411075118"
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'No campo BUDAT-LOW é informado a data de início da pesquisa de lançamentos.
    'Será informado 11.04.2014 para os indicadores, porque todas as ordens anteriores estão encerradas.
    'Obs: não é necessário informar o limite superior porque o SAP já traz esta informação.
    Session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").Text = strStDt
    Session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").Text = strEndDt
    Session.findById("wnd[0]/usr/btnBUT1").press
    Session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").Text = "1048576"
    Session.findById("wnd[1]").sendVKey 0
    Session.findById("wnd[0]").sendVKey 8
    
    Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "kob1.txt"
    Session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub

Sub IndicadorPlanejamento()
    
    Dim Sap_Con As Boolean
    Dim obj As New DataObject
    Dim txt As String
          
    ln = wsNS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
        
    wsNS.Activate
    wsNS.Range("D2:D" & ln).Copy

    Call AccessTcode("IW66")
    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/chkDY_QMSM").Selected = False
    Session.findById("wnd[0]/usr/cmbDY_PARVW").Key = " "
    Session.findById("wnd[0]/usr/ctxtDY_PARNR").Text = ""
    Session.findById("wnd[0]/usr/ctxtERDAT-LOW").Text = "01.01.2020"
    Session.findById("wnd[0]/usr/ctxtERDAT-HIGH").Text = "31.12.2020"
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "iw66.txt"
    Session.findById("wnd[1]").sendVKey 4
    Session.findById("wnd[1]/tbar[0]/btn[11]").press
    Session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub

Sub OcupacaoEquipeCampo()

    Call AccessTcode("IW49N")
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_VARBPL_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "02050205"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "02050758"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "02050913"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "02051203"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "02051204"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "02051206"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = "02102013"
    Session.findById("wnd[1]").sendVKey 0
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").Text = "ELIM"
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5").Select
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_NTANF-LOW").showContextMenu
    Session.findById("wnd[0]/usr").selectContextMenuItem "&01415000000026230"
    Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 2, "TEXT"
    Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "2"
    Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_NTANF-LOW").Text = Format(DtFim, "dd.mm.yyyy")
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_NTEND-LOW").showContextMenu
    Session.findById("wnd[0]/usr").selectContextMenuItem "&01615000000026230"
    Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 1, "TEXT"
    Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
    Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_NTEND-LOW").Text = Format(DtInicio, "dd.mm.yyyy")
    Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_NTEND-LOW").caretPosition = 10
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    Session.findById("wnd[0]/usr/chkSP_HIS").SetFocus
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    Session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "iw49n.txt"
    Session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
    Session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    Session.findById("wnd[1]").sendVKey 4
    Session.findById("wnd[2]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[2]/usr/ctxtDY_PATH").SetFocus
    Session.findById("wnd[2]/tbar[0]/btn[11]").press
    Session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub

Sub OcpCalendCampo()

    Call AccessTcode("ZTCS001")
    Session.findById("wnd[0]/usr/ctxtP_FSAVD").Text = "01.10.2020"
    Session.findById("wnd[0]/usr/ctxtP_FSEDD").Text = "31.10.2020"
    Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "1200"
    Session.findById("wnd[0]/usr/ctxtS_NAME-LOW").Text = "02050009"
    Session.findById("wnd[0]/usr/ctxtS_NAME-LOW").SetFocus
    Session.findById("wnd[0]/usr/ctxtS_NAME-LOW").caretPosition = 8
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell[0]").pressToolbarContextButton "&MB_EXPORT"
    Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell[0]").selectContextMenuItem "&PC"
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
    Session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    Session.findById("wnd[1]").sendVKey 4
    Session.findById("wnd[2]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "ztcs001.txt"
    Session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 11
    Session.findById("wnd[2]/tbar[0]/btn[11]").press
    Session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub

Sub CabecalhoOrdens()

    wsOS.Cells(1, 3).Value2 = "Tp"
    wsOS.Cells(1, 4).Value2 = "Cen"
    wsOS.Cells(1, 5).Value2 = "Dt Entr"
    wsOS.Cells(1, 6).Value2 = "Dt Modif"
    wsOS.Cells(1, 7).Value2 = "Dt Refer"
    wsOS.Cells(1, 12).Value2 = "Total Real"

End Sub

Sub CabecalhoNotas()

    wsNS.Cells(1, 1).Value2 = "Ps"
    wsNS.Cells(1, 6).Value2 = "Secao"
    wsNS.Cells(1, 7).Value2 = "Responsavel"
    wsNS.Cells(1, 8).Value2 = "Stat Usuario"
    wsNS.Cells(1, 9).Value2 = "Cliente"
    wsNS.Cells(1, 10).Value2 = "Linha"
    wsNS.Cells(1, 12).Value2 = "Dt Criacao"

End Sub

Sub CabecalhoQAs()

    wsQA.Cells(1, 1).Value2 = "Tp"
    wsQA.Cells(1, 4).Value2 = "Dt Criacao"
    wsQA.Cells(1, 10).Value2 = "Qtd Recl"
    wsQA.Cells(1, 11).Value2 = "Cen"

End Sub

Sub CabecalhoLancamentos()

    wsVal.Cells(1, 5).Value2 = "Data lcto"
    wsVal.Cells(1, 6).Value2 = "Cl Custo"
    wsVal.Cells(1, 7).Value2 = "Denom Classe Custo"
    wsVal.Cells(1, 8).Value2 = "N doc"
    wsVal.Cells(1, 11).Value2 = "Obj Parceiro"
    wsVal.Cells(1, 12).Value2 = "Valor"
    wsVal.Cells(1, 13).Value2 = "TpL"

End Sub

Sub CreateDataBase()
    
    Call SetParameters
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define a Query
    '--------------
    str_Consulta = "SELECT T.[Data lcto], T.[TpP], T.[TpL], T.Tipo, T.Empresa, T.Unidade, T.Ordem, T.Nota, SUM(Total) / COUNT(*) AS Montante, T.Grupo, T.Status "
    str_Consulta = str_Consulta & "FROM" & vbCrLf
    str_Consulta = str_Consulta & "(SELECT CB.[Empr], CB.[Segmento], CB.[TpP], CB.[TpL], OS.[Cen], CB.[Data lcto], CB.[Cl Custo], OS.[Ordem], NS.[Nota], QM.[Nota] AS [Nota QA], SUM(CB.[Valor]) AS Total," & vbCrLf
    str_Consulta = str_Consulta & "OS.[Texto breve], OS.[Nome de lista], OS.[Status do sistema], QM.[Status da nota]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha]," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1200 OR OS.[Cen] = 1201 , 'WEN', IIF(OS.[Cen] = 1204 OR OS.[Cen] = 1206 , 'HISA', IIF(OS.[Cen] = 1220, "
    str_Consulta = str_Consulta & "'EOL', IIF(OS.[Cen] = 1211, 'TGM', 'OUTRA')))) AS Empresa," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1211, 'SZO', IIF(OS.[Cen] = 1201, 'SBC', IIF(OS.[Cen] = 1204, 'JOA', 'JGS'))) AS Unidade," & vbCrLf
    str_Consulta = str_Consulta & "IIF(CB.[Cl Custo] = 411075004, 'Viagens', IIF(CB.[Cl Custo] = 411075007, 'Alimentação', IIF(CB.[Cl Custo] = 411075008, 'Transporte', "
    str_Consulta = str_Consulta & "IIF(CB.[Cl Custo] = 411075083, 'Fretes', IIF(CB.[Cl Custo] = 411075117, 'Serviços', 'Materiais'))))) AS Tipo," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Custo apropriado', IIF(ISNULL(QM.[Nota]) OR QM.[Status da nota] LIKE '%PRDT%' OR QM.[Status da nota] LIKE '%PRNP%' "
    str_Consulta = str_Consulta & "OR QM.[Status da nota] LIKE '%MELH%', 'Com a Assistência Técnica', 'Com o Controle de Qualidade')) AS Grupo," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Área causadora definida', IIF(ISNULL(QM.[Nota]), 'Sem QA', IIF(QM.[Status da nota] LIKE '%PRDT%', 'Nota procedente', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%PRNP%', 'Nota não procedente', IIF(QM.[Status da nota] LIKE '%MELH%', 'Necessita inf. adicionais', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%NAAV%', 'Nota não avaliada', IIF(QM.[Status da nota] LIKE '%EMAV%', 'Em avaliação', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%AGDV%', 'Aguardando devolução')))))))) AS Status,"
    str_Consulta = str_Consulta & "YEAR(CB.[Data lcto]) AS Ano," & vbCrLf
    str_Consulta = str_Consulta & "MONTH(CB.[Data lcto]) AS Mês" & vbCrLf
    str_Consulta = str_Consulta & "FROM (([Ordens$] OS" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [LA$] CB ON OS.[Ordem]=CB.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [Notas$] NS ON OS.[Ordem]=NS.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [QAs$] QM ON NS.[Nota]=QM.[Nº modelo]" & vbCrLf
    str_Consulta = str_Consulta & "GROUP BY CB.[Empr], CB.[Segmento], CB.[TpP], CB.[TpL], OS.[Cen], CB.[Data lcto], CB.[Cl Custo], OS.[Ordem], NS.[Nota], QM.[Nota]," & vbCrLf
    str_Consulta = str_Consulta & "OS.[Texto breve], OS.[Nome de lista], OS.[Status do sistema], QM.[Status da nota]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha]" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY OS.[Ordem]) AS T" & vbCrLf
    str_Consulta = str_Consulta & "GROUP BY T.[Data lcto], T.[Empr], T.[Segmento], T.[TpP], T.[TpL], T.Tipo, T.Empresa, T.Unidade, T.Ordem, T.Nota, T.Grupo, T.Status" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY T.[Ordem] ASC"
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    If rs_Consulta.EOF = False Then
        
        rs_Consulta.MoveFirst
    
    Else
        
        MsgBox "Recordset vazio"
        Exit Sub
    
    End If
    
    i = 1
    For Each Fld In rs_Consulta.Fields
        
        wsDB.Cells(1, i).Value = Fld.Name
        i = i + 1
        
    Next
    
    'Cola Recordset na planilha
    wsDB.Range("A2").CopyFromRecordset rs_Consulta
    
    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing
    
    MsgBox "Banco de Dados Pronto!"
    
End Sub

Sub PlanCosts()
    
    Call SetParameters
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define a Query
    '--------------

    str_Consulta = "SELECT OS.[Ordem], NS.[Nota], QM.[Nota] AS [Nota QA], NS.[Stat Usuario], CB.[Data lcto], CB.[Cl Custo], CB.[Denom Classe Custo], SUM(CB.[Valor]) AS Total," & vbCrLf
    str_Consulta = str_Consulta & "OS.[Texto breve], OS.[Nome de lista]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], OS.[Status do sistema]," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1200 OR OS.[Cen] = 1201 , 'WEN', IIF(OS.[Cen] = 1204 OR OS.[Cen] = 1206 , 'HISA', IIF(OS.[Cen] = 1220, "
    str_Consulta = str_Consulta & "'EOL', IIF(OS.[Cen] = 1211, 'TGM', 'OUTRA')))) AS Empresa," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1211, 'SZO', IIF(OS.[Cen] = 1201, 'SBC', IIF(OS.[Cen] = 1204, 'JOA', 'JGS'))) AS Unidade," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Custo apropriado', IIF(ISNULL(QM.[Nota]) OR QM.[Status da nota] LIKE '%PRDT%' OR QM.[Status da nota] LIKE '%PRNP%' "
    str_Consulta = str_Consulta & "OR QM.[Status da nota] LIKE '%MELH%', 'Com a Assistência Técnica', 'Com o Controle de Qualidade')) AS Grupo," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Área causadora definida', IIF(ISNULL(QM.[Nota]), 'Sem QA', IIF(QM.[Status da nota] LIKE '%PRDT%', 'Nota procedente', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%PRNP%', 'Nota não procedente', IIF(QM.[Status da nota] LIKE '%MELH%', 'Necessita inf. adicionais', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%NAAV%', 'Nota não avaliada', IIF(QM.[Status da nota] LIKE '%EMAV%', 'Em avaliação', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%AGDV%', 'Aguardando devolução')))))))) AS Status,"
    str_Consulta = str_Consulta & "IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])) AS Ano," & vbCrLf
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto])) AS Mes" & vbCrLf
    str_Consulta = str_Consulta & "FROM (([OS$] OS" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [NS$] NS ON OS.[Ordem]=NS.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [LA$] CB ON OS.[Ordem]=CB.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [QA$] QM ON NS.[Nota]=QM.[Nº modelo]" & vbCrLf
    str_Consulta = str_Consulta & "WHERE YEAR(OS.[Dt Entr]) = 2021 AND ISNULL(TpP) AND NOT (OS.[Cen] = 1208 OR OS.[Cen] = 1210 OR OS.[Cen] = 1211)" & vbCrLf
    str_Consulta = str_Consulta & "GROUP BY CB.[TpP], CB.[TpL], CB.[Data lcto], CB.[Cl Custo], OS.[Ordem], NS.[Nota], QM.[Nota]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Stat Usuario], OS.[Texto breve], OS.[Nome de lista], OS.[Cen], OS.[Status do sistema], QM.[Status da nota], CB.[Denom Classe Custo]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])),"
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto]))" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY CB.[Data lcto], CB.[Cl Custo], OS.[Ordem]"
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    If rs_Consulta.EOF = False Then
        
        rs_Consulta.MoveFirst
    
    Else
        
        MsgBox "Recordset vazio"
        Exit Sub
    
    End If
        
    'Apaga todos os dados da planilha.
    '---------------------------------
    Call Clear_Entire_Worksheet(wsPOS)
    
    'Adiciona cabeçalho a planilha.
    '------------------------------
    i = 1
    For Each Fld In rs_Consulta.Fields
        
        wsPOS.Cells(1, i).Value = Fld.Name
        i = i + 1
        
    Next
    

    'Cola Recordset na planilha
    '--------------------------
    wsPOS.Range("A2").CopyFromRecordset rs_Consulta
    
    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing
    
    Call CorrigeValorLanc(wsPOS)
    wsPOS.Columns.AutoFit
    
    Call OptimizeVBA(False)
    
    MsgBox "Banco de Dados Pronto!"
    
End Sub

Sub OrdemServicoCustos(ws As Worksheet)
    
    Dim lngServOrder As Variant
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define da Query
    
    str_Consulta = "SELECT [Ordem], SUM([Valor]) AS 'Total'"
    str_Consulta = str_Consulta & "FROM [LA$] "
    str_Consulta = str_Consulta & "GROUP BY [Ordem] "
    str_Consulta = str_Consulta & "ORDER BY [Ordem]"
                    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    'Limpa a coluna Total.
    ln = ws.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    ws.Range("G2:G" & ln).Clear
    
    lngServOrder = Application.Transpose(ws.Range("A2:A" & ln).Value2)
    
    If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
    
    Do Until rs_Consulta.EOF = True
        
        If Not rs_Consulta("Ordem") = "" Then
            For i = 2 To ln
                'Debug.Print rs_Consulta("Ordem"), lngServOrder(i - 1)
                
                If rs_Consulta("Ordem") = lngServOrder(i - 1) Then
                    ws.Cells(i, 7).Value2 = rs_Consulta("'Total'")
                End If
            Next i
        
        End If
        rs_Consulta.MoveNext
        
    Loop

    ltr = FindColumnLetter("Valor", ws)
    ws.Columns(ltr & ":" & ltr).NumberFormat = "$ #,##0.00"

    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing

End Sub

Sub Cor_Values()
    
    Dim clSO        As Long
    Dim clVal       As Long
    Dim clConta     As Long
    Dim clDtLanc    As Long
    Dim counter     As Long

    Call SetParameters

    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    ln = wsDB.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    cl = wsDB.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 0).Column
    
    clVal = FindColumnNumber("Total", wsDB)
    clSO = FindColumnNumber("Ordem", wsDB)
    clConta = FindColumnNumber("Cl Custo", wsDB)
    clDtLanc = FindColumnNumber("Data lcto", wsDB)
    
    'Conecta ao arquivo Excel.
    '-------------------------
    Call ConnectXLFile(FilePath, FileName)
    
    'Define o RecordSet.
    '-------------------
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    For i = 2 To ln
    
        'Application.StatusBar = "Corrigindo valores: " & i & " de " & ln
        counter = 0
        
        str_Consulta = "SELECT COUNT([Ordem]) AS [Nº OS]" & vbCrLf
        str_Consulta = str_Consulta & "FROM [BD$]" & vbCrLf
        str_Consulta = str_Consulta & "WHERE [Ordem] = " & wsDB.Cells(i, clSO).Value & vbCrLf
        str_Consulta = str_Consulta & "GROUP BY [Empr], [Segmento], [Cen], [Data lcto], [Cl Custo]"
        
        'Abre Recordset.
        '---------------
        rs_Consulta.Open str_Consulta, ado_Conexao

        If rs_Consulta.EOF = False Then counter = CLng(rs_Consulta("Nº OS"))
        
        wsDB.Cells(i, clVal).Value = wsDB.Cells(i, clVal).Value / counter
        
        rs_Consulta.Close
        
    Next i

    'Fecha recordset e limpa a variável.
    '-----------------------------------
    Set rs_Consulta = Nothing
    ado_Conexao.Close
    Set ado_Conexao = Nothing

    Application.StatusBar = Empty
    
    MsgBox "Valores Corrigidos"
    
End Sub

Sub CorrigeValorLanc(ws As Worksheet)
    
    Dim cont        As Long
    Dim lnBase      As Long
    Dim i           As Long
    Dim j           As Long
    Dim lVal        As Long
    Dim vServOrder  As Variant
    Dim vDataLanc   As Variant
    Dim vTipoLanc   As Variant
    Dim vTotal      As Variant
    Dim sOS         As String
    Dim sDt         As String
    Dim sTL         As String
    Dim sTotal      As String
        
    lnBase = LastRowNumber(ws)
     
    sOS = FindColumnLetter("Ordem", ws)
    vServOrder = Application.Transpose(ws.Range(sOS & "2:" & sOS & lnBase).Value2)
    sDt = FindColumnLetter("Data lcto", ws)
    vDataLanc = Application.Transpose(ws.Range(sDt & "2:" & sDt & lnBase).Value2)
    sTL = FindColumnLetter("Cl Custo", ws)
    vTipoLanc = Application.Transpose(ws.Range(sTL & "2:" & sTL & lnBase).Value2)
    sTotal = FindColumnLetter("Total", ws)
    vTotal = Application.Transpose(ws.Range(sTotal & "2:" & sTotal & lnBase).Value2)
    
    lVal = FindColumnNumber("Total", ws)
    For i = 2 To lnBase
        cont = 0
        For j = 2 To lnBase
            If vServOrder(i - 1) = vServOrder(j - 1) And vDataLanc(i - 1) = vDataLanc(j - 1) And vTipoLanc(i - 1) = vTipoLanc(j - 1) And vTotal(i - 1) = vTotal(j - 1) Then
                cont = cont + 1
            End If
        Next j
        ws.Cells(i, lVal).Value = ws.Cells(i, lVal).Value / cont
    Next i

End Sub

Sub BasePassivo()

    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")

    'Define a Query
    
    str_Consulta = "SELECT [OS$].[Ordem], [OS$].[Nota], [OS$].[QA], [OS$].[Empresa], [OS$].[Descricao], "
    str_Consulta = str_Consulta & "[OS$].[Cliente], SUM ([LA$].[Valor]) AS 'Total', [OS$].[Status da ordem],  "
    str_Consulta = str_Consulta & "[OS$].[Status da QA], [OS$].[Secao], [OS$].[Responsavel], [OS$].[Linha], [OS$].[Data Criac], "
    str_Consulta = str_Consulta & "[OS$].[Data Conc], [OS$].[Grupo de status alocação], [OS$].[Status de alocação] "
    str_Consulta = str_Consulta & "FROM [OS$] "
    str_Consulta = str_Consulta & "LEFT JOIN [LA$] ON [OS$].[Ordem]=[LA$].[Ordem] "
    str_Consulta = str_Consulta & "WHERE [OS$].[Secao] = 'Secao Assistencia Tecnica' AND NOT [OS$].[Status da ordem] LIKE '%NOLQ%' "
    str_Consulta = str_Consulta & "AND [OS$].[Data Criac] < #01/01/2020# "
    str_Consulta = str_Consulta & "GROUP BY [OS$].[Ordem], [OS$].[Nota], [OS$].[QA], [OS$].[Empresa], [OS$].[Descricao], "
    str_Consulta = str_Consulta & "[OS$].[Cliente], [OS$].[Status da ordem],  "
    str_Consulta = str_Consulta & "[OS$].[Status da QA], [OS$].[Secao], [OS$].[Responsavel], [OS$].[Linha], [OS$].[Data Criac], "
    str_Consulta = str_Consulta & "[OS$].[Data Conc], [OS$].[Grupo de status alocação], [OS$].[Status de alocação] "
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
            
    'Apagar dados da planilha base
    Call ClearWorksheet(wsHist)
    
    'Cola Recordset na planilha
    wsHist.Range("A2").CopyFromRecordset rs_Consulta

    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing
    
    ltr = FindColumnLetter("Valor", wsHist)
    wsHist.Columns(ltr & ":" & ltr).NumberFormat = "$ #,##0.00"

End Sub

Sub CalculaDiasAber()

    Dim myrange() As Variant
    
    ln = wsDB.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    
    myrange = Application.Transpose(wsDB.Range("H2:H" & ln).Value)
    
    For i = 2 To ln
    
        If InStr(1, myrange(i - 1), "ENCE") Or InStr(1, myrange(i - 1), "ENTE") Then
            
            wsDB.Cells(i, 19).Formula = "=N" & i & "-M" & i
            
        End If
    
    Next i

End Sub

Sub CopiaCola(wsSrc As Worksheet, wsDest As Worksheet)
    
    wsDest.Activate
    wsDest.Cells.Clear
    wsSrc.Range("A1").CurrentRegion.Copy wsDest.Range("A1")
    wsDest.Columns.AutoFit

End Sub

Sub GastoAno()
    
    Call SetParameters
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define a Query
    '--------------
    str_Consulta = "SELECT OS.[Ordem], NS.[Nota], QM.[Nota] AS [Nota QA], NS.[Stat Usuario], CB.[Data lcto], CB.[Cl Custo], CB.[Denom Classe Custo], SUM(CB.[Valor]) AS Total," & vbCrLf
    str_Consulta = str_Consulta & "OS.[Texto breve], OS.[Nome de lista]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], OS.[Status do sistema]," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1200 OR OS.[Cen] = 1201 , 'WEN', IIF(OS.[Cen] = 1204 OR OS.[Cen] = 1206 , 'HISA', IIF(OS.[Cen] = 1220, "
    str_Consulta = str_Consulta & "'EOL', IIF(OS.[Cen] = 1211, 'TGM', 'OUTRA')))) AS Empresa," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1211, 'SZO', IIF(OS.[Cen] = 1201, 'SBC', IIF(OS.[Cen] = 1204, 'JOA', 'JGS'))) AS Unidade," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Custo apropriado', IIF(ISNULL(QM.[Nota]) OR QM.[Status da nota] LIKE '%PRDT%' OR QM.[Status da nota] LIKE '%PRNP%' "
    str_Consulta = str_Consulta & "OR QM.[Status da nota] LIKE '%MELH%', 'Com a Assistência Técnica', 'Com o Controle de Qualidade')) AS Grupo," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Área causadora definida', IIF(ISNULL(QM.[Nota]), 'Sem QA', IIF(QM.[Status da nota] LIKE '%PRDT%', 'Nota procedente', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%PRNP%', 'Nota não procedente', IIF(QM.[Status da nota] LIKE '%MELH%', 'Necessita inf. adicionais', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%NAAV%', 'Nota não avaliada', IIF((QM.[Status da nota] LIKE '%EMAV%' OR QM.[Status da nota] LIKE '%PEDT%'), 'Em avaliação', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%AGDV%', 'Aguardando devolução')))))))) AS Status,"
    str_Consulta = str_Consulta & "IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])) AS Ano," & vbCrLf
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto])) AS Mes" & vbCrLf
    str_Consulta = str_Consulta & "FROM (([OS$] OS" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [NS$] NS ON OS.[Ordem]=NS.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [LA$] CB ON OS.[Ordem]=CB.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [QA$] QM ON NS.[Nota]=QM.[Nº modelo]" & vbCrLf
    str_Consulta = str_Consulta & "WHERE YEAR(CB.[Data lcto]) > 2020 AND ISNULL(TpP) AND NOT (OS.[Cen] = 1211 OR OS.[Cen] = 1208 OR OS.[Cen] = 1210)" & vbCrLf
    str_Consulta = str_Consulta & "GROUP BY CB.[TpP], CB.[TpL], CB.[Data lcto], CB.[Cl Custo], OS.[Ordem], NS.[Nota], QM.[Nota]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Stat Usuario], OS.[Texto breve], OS.[Nome de lista], OS.[Cen], OS.[Status do sistema], QM.[Status da nota], CB.[Denom Classe Custo]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])),"
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto]))" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY CB.[Data lcto], CB.[Cl Custo], OS.[Ordem]"
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    If rs_Consulta.EOF = False Then
        
        rs_Consulta.MoveFirst
    
    Else
        
        MsgBox "Recordset vazio"
        Exit Sub
    
    End If
    
    i = 1
    For Each Fld In rs_Consulta.Fields
        
        wsGGA.Cells(1, i).Value = Fld.Name
        i = i + 1
        
    Next
    
    'Cola Recordset na planilha
    wsGGA.Range("A2").CopyFromRecordset rs_Consulta
      
    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing
    
    Call CorrigeValorLanc(wsGGA)
    
    MsgBox "Banco de Dados Pronto!"

End Sub

Sub ContaDiasNaoUteis()

    Dim DiadaSemana     As Long
    Dim DiasFolga       As Long
    Dim DiasFerias      As Long
    Dim DiasTreinamento As Long
    Dim DiasASO         As Long
    Dim Feriados        As Variant
    
    cl = wsOcpCalend.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 0).Column
    ln = wsOcpCalend.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    
    For j = 2 To cl
        
        DiasASO = 0
        DiasFerias = 0
        DiasFolga = 0
        DiasTreinamento = 0
        
        For i = 2 To ln
        
            DiadaSemana = WorksheetFunction.Weekday(wsOcpCalend.Cells(i, 1).Value)
            
            If Not (DiadaSemana = 1 Or DiadaSemana = 7) Then
            
                Select Case wsOcpCalend.Cells(i, j).Value
                    Case "Exames"
                        DiasASO = DiasASO + 1
                    Case "Férias"
                        DiasFerias = DiasFerias + 1
                    Case "Folga"
                        DiasFolga = DiasFolga + 1
                    Case "Treinamento"
                        DiasTreinamento = DiasTreinamento + 1
                End Select
            
            End If
        
        Next i
        
        wsOcpEqpCampo.Cells(j, 3).Value = DiasASO
        wsOcpEqpCampo.Cells(j, 4).Value = DiasFerias
        wsOcpEqpCampo.Cells(j, 5).Value = DiasFolga
        wsOcpEqpCampo.Cells(j, 6).Value = DiasTreinamento
        
        Feriados = Application.Transpose(wsFeriados.Range("A2:A13").Value)
        
        ln = wsOcpCalend.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
        wsOcpEqpCampo.Cells(j, 7).Value = WorksheetFunction.NetworkDays(wsOcpCalend.Cells(2, 1).Value, wsOcpCalend.Cells(ln, 1).Value, wsFeriados.Range("A2:A13")) - _
                                          DiasASO - DiasFerias - DiasFolga - DiasTreinamento
    
        ln = wsOcpOper.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
        wsOcpEqpCampo.Cells(j, 8).Value = WorksheetFunction.SumIf(wsOcpOper.Range("I2:I" & ln), _
                                          wsOcpEqpCampo.Range("A" & j), wsOcpOper.Range("N2:N" & ln))
        
    Next j
    

End Sub

Sub CorrigeDataInicioFim()

    ln = wsOcpOper.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    
    wsOcpOper.Cells(1, 14).Value2 = "DiasTrab."
    
    For i = 2 To ln
    
        If wsOcpOper.Cells(i, 11).Value < DtInicio Then
        
            wsOcpOper.Cells(i, 11).Value = DtInicio
        
        End If
    
        If wsOcpOper.Cells(i, 12).Value > DtFim Then
        
            wsOcpOper.Cells(i, 12).Value = DtFim
        
        End If
        
        wsOcpOper.Cells(i, 14).Value = wsOcpOper.Cells(i, 12).Value - wsOcpOper.Cells(i, 11).Value + 1
    
    Next i

End Sub

Sub Email_Responsible()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
    
    strResp = Resp
    
    For Each rp In strResp
        
        FilePath = wbIndicadores.Path & "\"
        FileName = wbIndicadores.Name
        
        Call ConnectXLFile(FilePath, FileName)
    
        'Seta RecordSet
        Set rs_Consulta = CreateObject("ADODB.Recordset")
    
        str_Consulta = "SELECT [Responsavel], [Status de alocação], [Cliente], [Ordem], SUM([Valor]) AS 'Total'" & _
                    "FROM [OS$] " & _
                    "WHERE NOT [Empresa] = 'WEN-SZO' AND NOT [Status da Ordem] LIKE '%EN_E%' AND [Responsavel] = '" & rp & "'" & _
                    "GROUP BY [Responsavel], [Status de alocação], [Cliente], [Ordem]" & _
                    "ORDER BY SUM([Valor]) DESC, [Status de alocação]"
                        
        'Abre Recordset
        rs_Consulta.Open str_Consulta, ado_Conexao
        
        HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;"
        HtmlContent = HtmlContent & "border-collapse: collapse;"
        HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;"
        HtmlContent = HtmlContent & "font-size: 12px;}"
        HtmlContent = HtmlContent & "th, td {padding: 5px;}"
        HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}</style>"
        HtmlContent = HtmlContent & "<p>Olá!</p>"
        HtmlContent = HtmlContent & "<p>Favor encerrar todas as ordens que podem ser encerradas, evitando deixar elas abertas por longos períodos.</p>"
        HtmlContent = HtmlContent & "<p>Segue abaixo lista com pendências de ordens:</p>"
        HtmlContent = HtmlContent & "<table>"
        HtmlContent = HtmlContent & "<thead>"
        HtmlContent = HtmlContent & "<tr><th>Status de alocação</th>"
        HtmlContent = HtmlContent & "<th>Cliente</th>"
        HtmlContent = HtmlContent & "<th>Ordem</th>"
        HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>"
        HtmlContent = HtmlContent & "</thead>"
    
        If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
        dblTotal = 0
        Do Until rs_Consulta.EOF = True
            HtmlContent = HtmlContent & "<tbody>"
            HtmlContent = HtmlContent & "<tr>"
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status de alocação") & "</td>"
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Cliente") & "</td>"
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>"
            HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("'Total'"), 2) & "</td>"
            HtmlContent = HtmlContent & "</tr>"
            HtmlContent = HtmlContent & "</tbody>"
            dblTotal = dblTotal + rs_Consulta("'Total'")
            rs_Consulta.MoveNext
        Loop
        HtmlContent = HtmlContent & "<tfoot>"
        HtmlContent = HtmlContent & "<tr>"
        HtmlContent = HtmlContent & "<td colspan=3><strong>Total</strong></td>"
        HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>"
        HtmlContent = HtmlContent & "</tfoot>"
        HtmlContent = HtmlContent & "</table>"
        HtmlContent = HtmlContent & "<p>Atenciosamente,<br>"
        HtmlContent = HtmlContent & "Claudenir da Silva</p>"
        
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        
        
        
        With OutMail
            If IsNull(rp) Or rp = "EDUARDO CAMPREGHER" Then
                .To = "claudenir@weg.net"
            ElseIf rp = "ANDRE FELIPE COPETTI" Then
                .To = "copetti@weg.net"
            Else
                .To = rp
            End If
            '.Cc = Range("B2").Value
            '.Bcc = Range("B3").Value
            .Subject = "PENDÊNCIAS DE ORDENS - " & rp & " - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
            .HTMLBody = HtmlContent
            .Display
        End With
        
        On Error GoTo 0
        
        Set OutMail = Nothing
        
    Next rp

End Sub

Sub EmailHISA()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
    
        
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)

    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")

    str_Consulta = "SELECT [Responsavel], [Status de alocação], [Cliente], [Ordem], [Status da Ordem], SUM([Valor]) AS 'Total'" & _
                "FROM [OS$] " & _
                "WHERE [Empresa] LIKE '%HISA%' AND NOT [Status da Ordem] LIKE '%EN_E%'" & _
                "GROUP BY [Responsavel], [Status de alocação], [Cliente], [Ordem], [Status da Ordem] " & _
                "ORDER BY SUM([Valor]) DESC, [Status de alocação]"
                    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;"
    HtmlContent = HtmlContent & "border-collapse: collapse;"
    HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;"
    HtmlContent = HtmlContent & "font-size: 12px;}"
    HtmlContent = HtmlContent & "th, td {padding: 5px;}"
    HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}</style>"
    HtmlContent = HtmlContent & "<p>Olá!</p>"
    HtmlContent = HtmlContent & "<p>Favor encerrar todas as ordens que podem ser encerradas, evitando deixar elas abertas por longos períodos.</p>"
    HtmlContent = HtmlContent & "<p>Segue abaixo lista com pendências de ordens:</p>"
    HtmlContent = HtmlContent & "<table>"
    HtmlContent = HtmlContent & "<thead>"
    HtmlContent = HtmlContent & "<tr><th>Status de alocação</th>"
    HtmlContent = HtmlContent & "<th>Cliente</th>"
    HtmlContent = HtmlContent & "<th>Ordem</th>"
    HtmlContent = HtmlContent & "<th>Status da Ordem</th>"
    HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>"
    HtmlContent = HtmlContent & "</thead>"

    If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
    dblTotal = 0
    Do Until rs_Consulta.EOF = True
        HtmlContent = HtmlContent & "<tbody>"
        HtmlContent = HtmlContent & "<tr>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status de alocação") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Cliente") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status da Ordem") & "</td>"
        HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("'Total'"), 2) & "</td>"
        HtmlContent = HtmlContent & "</tr>"
        HtmlContent = HtmlContent & "</tbody>"
        dblTotal = dblTotal + rs_Consulta("'Total'")
        rs_Consulta.MoveNext
    Loop
    HtmlContent = HtmlContent & "<tfoot>"
    HtmlContent = HtmlContent & "<tr>"
    HtmlContent = HtmlContent & "<td colspan=4><strong>Total</strong></td>"
    HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>"
    HtmlContent = HtmlContent & "</tfoot>"
    HtmlContent = HtmlContent & "</table>"
    HtmlContent = HtmlContent & "<p>Atenciosamente,<br>"
    HtmlContent = HtmlContent & "Claudenir da Silva</p>"
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    
    
    With OutMail
        .To = "Gustavo Henrique de Lima Bruno"
        '.Cc = Range("B2").Value
        '.Bcc = Range("B3").Value
        .Subject = "PENDÊNCIAS DE ORDENS - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
        .HTMLBody = HtmlContent
        .Display
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing

End Sub

Sub CheckPlan()

    Call SetParameters
    
    Sap_Con = ConectaSAP
    If Not Sap_Con Then
        Call OptimizeVBA(False)
        Exit Sub
    End If
    
    ln = wsNS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    wsNS.Range("D2:D" & ln).Copy

    Call AccessTcode("IW66")
    Session.findById("wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/chkDY_QMSM").Selected = False
    Session.findById("wnd[0]/usr/cmbDY_PARVW").Key = " "
    Session.findById("wnd[0]/usr/ctxtDY_PARNR").Text = ""
    Session.findById("wnd[0]/usr/chkDY_QMSM").SetFocus
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\claudenir\OneDrive - WEG EQUIPAMENTOS ELETRICOS S.A\GERENCIAMENTO DE ROTINA\2021\Base\"
    Session.findById("wnd[1]").sendVKey 4
    Session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "iw66.txt"
    Session.findById("wnd[2]/tbar[0]/btn[11]").press
    Session.findById("wnd[1]/tbar[0]/btn[11]").press
    
End Sub

Sub AddStatusPlan()
    
    Dim lnMOS As Long
    Dim lnPOS As Long

    lnMOS = wsMOS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    lnPOS = wsPOS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    cl = wsPOS.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).Column

    wsPOS.Cells(1, cl).Value = "Planejamento"

    For i = 2 To lnMOS
    
        If Not wsMOS.Cells(i, 1).Value = wsMOS.Cells(i - 1, 1).Value Then
        
            For j = 2 To lnPOS
            
                If wsMOS.Cells(i, 1).Value = wsPOS.Cells(j, 2).Value Then
                
                    wsPOS.Cells(j, cl).Value = "Planejada"
                
                End If
                
            Next j
        
        End If
    
    Next i
    
    For i = 2 To lnPOS
    
        If wsPOS.Cells(i, cl).Value = "" Then
        
            wsPOS.Cells(i, cl).Value = "Sem Planejamento"
        
        End If
    
    Next i

End Sub

Sub Email_PlanResp()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
    
    strResp = PlanResp
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
        
    Call ConnectXLFile(FilePath, FileName)
    
    For Each rp In strResp
           
        'Seta RecordSet
        Set rs_Consulta = CreateObject("ADODB.Recordset")
    
        str_Consulta = "SELECT [Planejamento], [Responsavel], [Status], [Nome de lista], [Ordem], SUM([Total]) AS Montante" & vbCrLf
        str_Consulta = str_Consulta & "FROM [POS$]" & vbCrLf
        str_Consulta = str_Consulta & "WHERE NOT [Status do sistema] LIKE '%EN_E%' AND [Responsavel] = '" & rp & "'" & vbCrLf
        str_Consulta = str_Consulta & "GROUP BY [Planejamento], [Responsavel], [Status], [Nome de lista], [Ordem]" & vbCrLf
        str_Consulta = str_Consulta & "ORDER BY SUM([Total]) DESC, [Status]"
                        
        'Abre Recordset
        rs_Consulta.Open str_Consulta, ado_Conexao
        
        HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;" & vbCrLf
        HtmlContent = HtmlContent & "border-collapse: collapse;" & vbCrLf
        HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;" & vbCrLf
        HtmlContent = HtmlContent & "font-size: 12px;}" & vbCrLf
        HtmlContent = HtmlContent & "th, td {padding: 5px;}" & vbCrLf
        HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}</style>" & vbCrLf
        HtmlContent = HtmlContent & "<p>Olá!</p>" & vbCrLf
        HtmlContent = HtmlContent & "<p>Favor providenciar o planejamento de custos das ordes listadas neste e-mail.</p>" & vbCrLf
        HtmlContent = HtmlContent & "<p>Segue abaixo lista com pendências de ordens:</p>" & vbCrLf
        HtmlContent = HtmlContent & "<table>" & vbCrLf
        HtmlContent = HtmlContent & "<thead>" & vbCrLf
        HtmlContent = HtmlContent & "<tr><th>Planejamento</th>" & vbCrLf
        HtmlContent = HtmlContent & "<th>Status de alocação</th>" & vbCrLf
        HtmlContent = HtmlContent & "<th>Cliente</th>" & vbCrLf
        HtmlContent = HtmlContent & "<th>Ordem</th>" & vbCrLf
        HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>" & vbCrLf
        HtmlContent = HtmlContent & "</thead>" & vbCrLf
    
        If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
        dblTotal = 0
        Do Until rs_Consulta.EOF = True
            HtmlContent = HtmlContent & "<tbody>" & vbCrLf
            HtmlContent = HtmlContent & "<tr>" & vbCrLf
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Planejamento") & "</td>" & vbCrLf
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status") & "</td>" & vbCrLf
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Nome de lista") & "</td>" & vbCrLf
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>" & vbCrLf
            HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("Montante"), 2) & "</td>" & vbCrLf
            HtmlContent = HtmlContent & "</tr>" & vbCrLf
            HtmlContent = HtmlContent & "</tbody>" & vbCrLf
            dblTotal = dblTotal + rs_Consulta("Montante") & vbCrLf
            rs_Consulta.MoveNext
        Loop
        HtmlContent = HtmlContent & "<tfoot>" & vbCrLf
        HtmlContent = HtmlContent & "<tr>" & vbCrLf
        HtmlContent = HtmlContent & "<td colspan=4><strong>Total</strong></td>" & vbCrLf
        HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>" & vbCrLf
        HtmlContent = HtmlContent & "</tfoot>" & vbCrLf
        HtmlContent = HtmlContent & "</table>" & vbCrLf
        HtmlContent = HtmlContent & "<p>Atenciosamente,<br>" & vbCrLf
        HtmlContent = HtmlContent & "Claudenir da Silva</p>"
        
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        
        
        
        With OutMail
            If IsNull(rp) Or rp = "EDUARDO CAMPREGHER" Then
                .To = "claudenir@weg.net"
            ElseIf rp = "ANDRE FELIPE COPETTI" Then
                .To = "copetti@weg.net"
            Else
                .To = rp
            End If
            '.Cc = Range("B2").Value
            '.Bcc = Range("B3").Value
            .Subject = "PLANEJAMENTO DE ORDENS - " & rp & " - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
            .HTMLBody = HtmlContent
            .Display
        End With
        
        On Error GoTo 0
        
        Set OutMail = Nothing
        
    Next rp

End Sub

Sub Email_PlanSuper()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
        
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")

    str_Consulta = "SELECT [Secao], [Planejamento], [Responsavel], [Status], [Nome de lista], [Ordem], [Stat Usuario], SUM([Total]) AS Montante" & vbCrLf
    str_Consulta = str_Consulta & "FROM [POS$]" & vbCrLf
    str_Consulta = str_Consulta & "WHERE NOT [Status do sistema] LIKE '%EN_E%' AND ([Secao] = 'Secao Implantacao e Assistencia Tec' OR [Secao] = 'Secao Assistencia Tecnica') AND " & vbCrLf
    str_Consulta = str_Consulta & "NOT ([Stat Usuario] = 'AnCC' OR [Stat Usuario] = 'CtCp')"
    str_Consulta = str_Consulta & "GROUP BY [Secao], [Planejamento], [Responsavel], [Status], [Nome de lista], [Ordem], [Stat Usuario]" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY [Secao], [Responsavel], SUM([Total]) DESC, [Status]"
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;" & vbCrLf
    HtmlContent = HtmlContent & "border-collapse: collapse;" & vbCrLf
    HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;" & vbCrLf
    HtmlContent = HtmlContent & "font-size: 12px;}" & vbCrLf
    HtmlContent = HtmlContent & "th, td {padding: 5px;}" & vbCrLf
    HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}</style>" & vbCrLf
    HtmlContent = HtmlContent & "<p>Olá!</p>" & vbCrLf
    HtmlContent = HtmlContent & "<p>Favor providenciar o planejamento de custos das ordes listadas neste e-mail.</p>" & vbCrLf
    HtmlContent = HtmlContent & "<p>Segue abaixo lista com pendências de ordens:</p>" & vbCrLf
    HtmlContent = HtmlContent & "<table>" & vbCrLf
    HtmlContent = HtmlContent & "<thead>" & vbCrLf
    HtmlContent = HtmlContent & "<tr><th>Seção</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Responsável</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Ordem</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Status da nota</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Planejamento</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Status de alocação</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Cliente</th>" & vbCrLf
    HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>" & vbCrLf
    HtmlContent = HtmlContent & "</thead>" & vbCrLf

    If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
    dblTotal = 0
    Do Until rs_Consulta.EOF = True
        HtmlContent = HtmlContent & "<tbody>" & vbCrLf
        HtmlContent = HtmlContent & "<tr>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Secao") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Responsavel") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Stat Usuario") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Planejamento") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Nome de lista") & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("Montante"), 2) & "</td>" & vbCrLf
        HtmlContent = HtmlContent & "</tr>" & vbCrLf
        HtmlContent = HtmlContent & "</tbody>" & vbCrLf
        dblTotal = dblTotal + rs_Consulta("Montante") & vbCrLf
        rs_Consulta.MoveNext
    Loop
    HtmlContent = HtmlContent & "<tfoot>" & vbCrLf
    HtmlContent = HtmlContent & "<tr>" & vbCrLf
    HtmlContent = HtmlContent & "<td colspan=7><strong>Total</strong></td>" & vbCrLf
    HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>" & vbCrLf
    HtmlContent = HtmlContent & "</tfoot>" & vbCrLf
    HtmlContent = HtmlContent & "</table>" & vbCrLf
    HtmlContent = HtmlContent & "<p>Atenciosamente,<br>" & vbCrLf
    HtmlContent = HtmlContent & "Claudenir da Silva</p>"
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    With OutMail
        If IsNull(rp) Or rp = "EDUARDO CAMPREGHER" Then
            .To = "claudenir@weg.net"
        ElseIf rp = "ANDRE FELIPE COPETTI" Then
            .To = "copetti@weg.net"
        Else
            .To = rp
        End If
        '.Cc = Range("B2").Value
        '.Bcc = Range("B3").Value
        .Subject = "PLANEJAMENTO DE ORDENS - POR SEÇÃO - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
        .HTMLBody = HtmlContent
        .Display
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing

End Sub


Sub OrderLeftBehind()
    
    Call SetParameters
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define a Query
    '--------------

    str_Consulta = "SELECT OS.[Ordem], NS.[Nota], QM.[Nota] AS [Nota QA], NS.[Stat Usuario], CB.[Data lcto], CB.[Cl Custo], CB.[Denom Classe Custo], SUM(CB.[Valor]) AS Total," & vbCrLf
    str_Consulta = str_Consulta & "OS.[Texto breve], OS.[Nome de lista]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], OS.[Status do sistema]," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1200 OR OS.[Cen] = 1201 , 'WEN', IIF(OS.[Cen] = 1204 OR OS.[Cen] = 1206 , 'HISA', IIF(OS.[Cen] = 1220, "
    str_Consulta = str_Consulta & "'EOL', IIF(OS.[Cen] = 1211, 'TGM', 'OUTRA')))) AS Empresa," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1211, 'SZO', IIF(OS.[Cen] = 1201, 'SBC', IIF(OS.[Cen] = 1204, 'JOA', 'JGS'))) AS Unidade," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Custo apropriado', IIF(ISNULL(QM.[Nota]) OR QM.[Status da nota] LIKE '%PRDT%' OR QM.[Status da nota] LIKE '%PRNP%' "
    str_Consulta = str_Consulta & "OR QM.[Status da nota] LIKE '%MELH%', 'Com a Assistência Técnica', 'Com o Controle de Qualidade')) AS Grupo," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Área causadora definida', IIF(ISNULL(QM.[Nota]), 'Sem QA', IIF(QM.[Status da nota] LIKE '%PRDT%', 'Nota procedente', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%PRNP%', 'Nota não procedente', IIF(QM.[Status da nota] LIKE '%MELH%', 'Necessita inf. adicionais', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%NAAV%', 'Nota não avaliada', IIF((QM.[Status da nota] LIKE '%EMAV%' OR QM.[Status da nota] LIKE '%PEDT%'), 'Em avaliação', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%AGDV%', 'Aguardando devolução')))))))) AS Status,"
    str_Consulta = str_Consulta & "IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])) AS Ano," & vbCrLf
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto])) AS Mes" & vbCrLf
    str_Consulta = str_Consulta & "FROM (([OS$] OS" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [NS$] NS ON OS.[Ordem]=NS.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [LA$] CB ON OS.[Ordem]=CB.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [QA$] QM ON NS.[Nota]=QM.[Nº modelo]" & vbCrLf
    str_Consulta = str_Consulta & "WHERE YEAR(CB.[Data lcto]) < 2021 AND ISNULL(TpP) AND YEAR(OS.[Dt Entr]) < 2021 AND NOT (OS.[Cen] = 1211 OR OS.[Cen] = 1208 OR OS.[Cen] = 1210) AND NOT (OS.[Status do sistema] LIKE '%ENCE%' OR OS.[Status do sistema] LIKE '%NOLQ%')" & vbCrLf
    str_Consulta = str_Consulta & "GROUP BY CB.[TpP], CB.[TpL], CB.[Data lcto], CB.[Cl Custo], OS.[Ordem], NS.[Nota], QM.[Nota]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Stat Usuario], OS.[Texto breve], OS.[Nome de lista], OS.[Cen], OS.[Status do sistema], QM.[Status da nota], CB.[Denom Classe Custo]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])),"
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto]))" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY CB.[Data lcto], CB.[Cl Custo], OS.[Ordem]"
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    If rs_Consulta.EOF = False Then
        
        rs_Consulta.MoveFirst
    
    Else
        
        MsgBox "Recordset vazio"
        Exit Sub
    
    End If
        
    'Apaga todos os dados da planilha.
    '---------------------------------
    Call Clear_Entire_Worksheet(wsPAS)
    
    'Adiciona cabeçalho a planilha.
    '------------------------------
    i = 1
    For Each Fld In rs_Consulta.Fields
        
        wsPAS.Cells(1, i).Value = Fld.Name
        i = i + 1
        
    Next
    

    'Cola Recordset na planilha
    '--------------------------
    wsPAS.Range("A2").CopyFromRecordset rs_Consulta
    
    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing
    
    Call CorrigeValorLanc(wsPAS)
    wsPAS.Columns.AutoFit
    
    Call OptimizeVBA(False)
    
    MsgBox "Banco de Dados Pronto!"
    
End Sub

Sub EmailPassivo_Responsaveis()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
    
    strResp = Responsavel
    
    For Each rp In strResp
        
        FilePath = wbIndicadores.Path & "\"
        FileName = wbIndicadores.Name
        
        Call ConnectXLFile(FilePath, FileName)
    
        'Seta RecordSet
        Set rs_Consulta = CreateObject("ADODB.Recordset")
    
        str_Consulta = "SELECT [Responsavel], [Status], [Nome de lista], [Ordem], SUM([Total]) AS 'Total'" & _
                    "FROM [PAS$] " & _
                    "WHERE NOT [Status do sistema] LIKE '%EN_E%' AND [Responsavel] = '" & rp & "'" & _
                    "GROUP BY [Responsavel], [Status], [Nome de lista], [Ordem]" & _
                    "ORDER BY SUM([Total]) DESC, [Status]"
                        
        'Abre Recordset
        rs_Consulta.Open str_Consulta, ado_Conexao
        
        HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;"
        HtmlContent = HtmlContent & "border-collapse: collapse;"
        HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;"
        HtmlContent = HtmlContent & "font-size: 12px;}"
        HtmlContent = HtmlContent & "th, td {padding: 5px;}"
        HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}</style>"
        HtmlContent = HtmlContent & "<p>Olá!</p>"
        HtmlContent = HtmlContent & "<p>Favor encerrar todas as ordens que podem ser encerradas, evitando deixar elas abertas por longos períodos.</p>"
        HtmlContent = HtmlContent & "<p>Segue abaixo lista com pendências de ordens:</p>"
        HtmlContent = HtmlContent & "<table>"
        HtmlContent = HtmlContent & "<thead>"
        HtmlContent = HtmlContent & "<tr><th>Status de alocação</th>"
        HtmlContent = HtmlContent & "<th>Cliente</th>"
        HtmlContent = HtmlContent & "<th>Ordem</th>"
        HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>"
        HtmlContent = HtmlContent & "</thead>"
    
        If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
        dblTotal = 0
        Do Until rs_Consulta.EOF = True
            HtmlContent = HtmlContent & "<tbody>"
            HtmlContent = HtmlContent & "<tr>"
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status") & "</td>"
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Nome de lista") & "</td>"
            HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>"
            HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("'Total'"), 2) & "</td>"
            HtmlContent = HtmlContent & "</tr>"
            HtmlContent = HtmlContent & "</tbody>"
            dblTotal = dblTotal + rs_Consulta("'Total'")
            rs_Consulta.MoveNext
        Loop
        HtmlContent = HtmlContent & "<tfoot>"
        HtmlContent = HtmlContent & "<tr>"
        HtmlContent = HtmlContent & "<td colspan=3><strong>Total</strong></td>"
        HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>"
        HtmlContent = HtmlContent & "</tfoot>"
        HtmlContent = HtmlContent & "</table>"
        HtmlContent = HtmlContent & "<p>Atenciosamente,<br>"
        HtmlContent = HtmlContent & "Claudenir da Silva</p>"
        
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        
        
        
        With OutMail
            If IsNull(rp) Or rp = "EDUARDO CAMPREGHER" Then
                .To = "claudenir@weg.net"
            ElseIf rp = "ANDRE FELIPE COPETTI" Then
                .To = "copetti@weg.net"
            Else
                .To = rp
            End If
            '.Cc = Range("B2").Value
            '.Bcc = Range("B3").Value
            .Subject = "PASSIVO DE ORDENS - " & rp & " - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
            .HTMLBody = HtmlContent
            .Display
        End With
        
        On Error GoTo 0
        
        Set OutMail = Nothing
        
    Next rp

End Sub

Sub EmailPassivo_Gestores()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
        
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)

    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")

    str_Consulta = "SELECT [Responsavel], [Secao], [Status], [Nome de lista], [Ordem], SUM([Total]) AS 'Total'" & _
                "FROM [PAS$] " & _
                "WHERE [Grupo] = 'Com a Assistência Técnica' AND NOT [Status do sistema] LIKE '%EN_E%'" & _
                "GROUP BY [Responsavel], [Secao], [Status], [Nome de lista], [Ordem]" & _
                "ORDER BY [Secao], SUM([Total]) DESC, [Status]"
                
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;"
    HtmlContent = HtmlContent & "border-collapse: collapse;"
    HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;"
    HtmlContent = HtmlContent & "font-size: 11px;}"
    HtmlContent = HtmlContent & "th, td {padding: 5px;}"
    HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}"
    HtmlContent = HtmlContent & "p {font-family: Arial, Helvetica, sans-serif;" & _
                                "font-size: 11pt;}</style>"
    HtmlContent = HtmlContent & "<p>Olá!</p>"
    HtmlContent = HtmlContent & "<p>Segue relatório de passivo de ordens sob a responsabilidade destas <strong>Seções</strong>.</p>"
    HtmlContent = HtmlContent & "<table>"
    HtmlContent = HtmlContent & "<thead>"
    HtmlContent = HtmlContent & "<tr><th>Seção</th>"
    HtmlContent = HtmlContent & "<th>Responsável</th>"
    HtmlContent = HtmlContent & "<th>Status de alocação</th>"
    HtmlContent = HtmlContent & "<th>Cliente</th>"
    HtmlContent = HtmlContent & "<th>Ordem</th>"
    HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>"
    HtmlContent = HtmlContent & "</thead>"

    If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
    dblTotal = 0
    Do Until rs_Consulta.EOF = True
        HtmlContent = HtmlContent & "<tbody>"
        HtmlContent = HtmlContent & "<tr>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Secao") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Responsavel") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Nome de lista") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>"
        HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("'Total'"), 2) & "</td>"
        HtmlContent = HtmlContent & "</tr>"
        HtmlContent = HtmlContent & "</tbody>"
        dblTotal = dblTotal + rs_Consulta("'Total'")
        rs_Consulta.MoveNext
    Loop
    HtmlContent = HtmlContent & "<tfoot>"
    HtmlContent = HtmlContent & "<tr>"
    HtmlContent = HtmlContent & "<td colspan=5><strong>Total</strong></td>"
    HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>"
    HtmlContent = HtmlContent & "</tfoot>"
    HtmlContent = HtmlContent & "</table>"
    HtmlContent = HtmlContent & "<p>Atenciosamente,<br>"
    HtmlContent = HtmlContent & "Claudenir da Silva</p>"
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    
    
    With OutMail
        .To = "clausw@weg.net; luizfernando@weg.net"
        .Cc = "glauco@weg.net"
        '.Bcc = Range("B3").Value
        .Subject = "PASSIVO DE ORDENS - " & rp & " - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
        .HTMLBody = HtmlContent
        .Display
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing

End Sub

Sub EmailPassivo_Qualidade()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim strResp As Variant
    Dim rp As Variant
    Dim dblTotal As Double
    
    Call SetParameters
        
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)

    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")

    str_Consulta = "SELECT [Empresa], [Responsavel], [Secao], [Status], [Nome de lista], [Ordem], [Nota QA], SUM([Total]) AS 'Total'" & _
                "FROM [PAS$] " & _
                "WHERE [Grupo] = 'Com o Controle de Qualidade' AND NOT [Status do sistema] LIKE '%EN_E%'" & _
                "GROUP BY [Empresa], [Responsavel], [Secao], [Status], [Nome de lista], [Ordem], [Nota QA]" & _
                "ORDER BY [Empresa], [Secao], SUM([Total]) DESC, [Status]"
                
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;"
    HtmlContent = HtmlContent & "border-collapse: collapse;"
    HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;"
    HtmlContent = HtmlContent & "font-size: 11px;}"
    HtmlContent = HtmlContent & "th, td {padding: 5px;}"
    HtmlContent = HtmlContent & "tfoot tr td {text-align: center;}"
    HtmlContent = HtmlContent & "p {font-family: Arial, Helvetica, sans-serif;" & _
                                "font-size: 11pt;}</style>"
    HtmlContent = HtmlContent & "<p>Olá!</p>"
    HtmlContent = HtmlContent & "<p>Segue relatório de passivo de ordens sob a responsabilidade do <strong>Controle de Qualidade</strong>.</p>"
    HtmlContent = HtmlContent & "<table>"
    HtmlContent = HtmlContent & "<thead>"
    HtmlContent = HtmlContent & "<tr><th>Empresa</th>"
    HtmlContent = HtmlContent & "<th>Seção</th>"
    HtmlContent = HtmlContent & "<th>Responsável</th>"
    HtmlContent = HtmlContent & "<th>Status de alocação</th>"
    HtmlContent = HtmlContent & "<th>Cliente</th>"
    HtmlContent = HtmlContent & "<th>Ordem</th>"
    HtmlContent = HtmlContent & "<th>QA</th>"
    HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>"
    HtmlContent = HtmlContent & "</thead>"

    If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
    dblTotal = 0
    Do Until rs_Consulta.EOF = True
        HtmlContent = HtmlContent & "<tbody>"
        HtmlContent = HtmlContent & "<tr>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Empresa") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Secao") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Responsavel") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Nome de lista") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Nota QA") & "</td>"
        HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("'Total'"), 2) & "</td>"
        HtmlContent = HtmlContent & "</tr>"
        HtmlContent = HtmlContent & "</tbody>"
        dblTotal = dblTotal + rs_Consulta("'Total'")
        rs_Consulta.MoveNext
    Loop
    HtmlContent = HtmlContent & "<tfoot>"
    HtmlContent = HtmlContent & "<tr>"
    HtmlContent = HtmlContent & "<td colspan=7><strong>Total</strong></td>"
    HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>"
    HtmlContent = HtmlContent & "</tfoot>"
    HtmlContent = HtmlContent & "</table>"
    HtmlContent = HtmlContent & "<p>Atenciosamente,<br>"
    HtmlContent = HtmlContent & "Claudenir da Silva</p>"
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    
    
    With OutMail
        .To = "dioneig@weg.net; okuno@weg.net; marcelohs@weg.net; marcelon@weg.net; mauriciov@weg.net"
        .Cc = "glauco@weg.net; clausw@weg.net; luizfernando@weg.net"
        '.Bcc = Range("B3").Value
        .Subject = "PASSIVO DE ORDENS - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
        .HTMLBody = HtmlContent
        .Send
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing
    
    MsgBox "E-mail enviado!"

End Sub

Sub OrdensAbertas()

    
    Call SetParameters
    
    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    Call ConnectXLFile(FilePath, FileName)
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define a Query
    '--------------

    str_Consulta = "SELECT OS.[Ordem], NS.[Nota], QM.[Nota] AS [Nota QA], NS.[Stat Usuario], CB.[Data lcto], CB.[Cl Custo], CB.[Denom Classe Custo], SUM(CB.[Valor]) AS Total," & vbCrLf
    str_Consulta = str_Consulta & "OS.[Texto breve], OS.[Nome de lista]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], OS.[Status do sistema]," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1200 OR OS.[Cen] = 1201 , 'WEN', IIF(OS.[Cen] = 1204 OR OS.[Cen] = 1206 , 'HISA', IIF(OS.[Cen] = 1220, "
    str_Consulta = str_Consulta & "'EOL', IIF(OS.[Cen] = 1211, 'TGM', 'OUTRA')))) AS Empresa," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Cen] = 1211, 'SZO', IIF(OS.[Cen] = 1201, 'SBC', IIF(OS.[Cen] = 1204, 'JOA', 'JGS'))) AS Unidade," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Custo apropriado', IIF(ISNULL(QM.[Nota]) OR QM.[Status da nota] LIKE '%PRDT%' OR QM.[Status da nota] LIKE '%PRNP%' "
    str_Consulta = str_Consulta & "OR QM.[Status da nota] LIKE '%MELH%', 'Com a Assistência Técnica', 'Com o Controle de Qualidade')) AS Grupo," & vbCrLf
    str_Consulta = str_Consulta & "IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Área causadora definida', IIF(ISNULL(QM.[Nota]), 'Sem QA', IIF(QM.[Status da nota] LIKE '%PRDT%', 'Nota procedente', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%PRNP%', 'Nota não procedente', IIF(QM.[Status da nota] LIKE '%MELH%', 'Necessita inf. adicionais', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%NAAV%', 'Nota não avaliada', IIF((QM.[Status da nota] LIKE '%EMAV%' OR QM.[Status da nota] LIKE '%PEDT%'), 'Em avaliação', "
    str_Consulta = str_Consulta & "IIF(QM.[Status da nota] LIKE '%AGDV%', 'Aguardando devolução')))))))) AS Status,"
    str_Consulta = str_Consulta & "IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])) AS Ano," & vbCrLf
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto])) AS Mes" & vbCrLf
    str_Consulta = str_Consulta & "FROM (([OS$] OS" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [NS$] NS ON OS.[Ordem]=NS.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [LA$] CB ON OS.[Ordem]=CB.[Ordem])" & vbCrLf
    str_Consulta = str_Consulta & "LEFT JOIN [QA$] QM ON NS.[Nota]=QM.[Nº modelo]" & vbCrLf
    str_Consulta = str_Consulta & "WHERE ISNULL(TpP) AND NOT (OS.[Cen] = 1211 OR OS.[Cen] = 1208 OR OS.[Cen] = 1210) AND NOT OS.[Status do sistema] LIKE '%EN%E%' AND" & vbCrLf
    str_Consulta = str_Consulta & "NOT IIF(OS.[Status do sistema] LIKE '%NOLQ%', 'Custo apropriado', IIF(ISNULL(QM.[Nota]) OR QM.[Status da nota] LIKE '%PRDT%' OR QM.[Status da nota] LIKE '%PRNP%' "
    str_Consulta = str_Consulta & "OR QM.[Status da nota] LIKE '%MELH%', 'Com a Assistência Técnica', 'Com o Controle de Qualidade')) LIKE 'Custo apropriado'" & vbCrLf
    str_Consulta = str_Consulta & "GROUP BY CB.[TpP], CB.[TpL], CB.[Data lcto], CB.[Cl Custo], OS.[Ordem], NS.[Nota], QM.[Nota]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Stat Usuario], OS.[Texto breve], OS.[Nome de lista], OS.[Cen], OS.[Status do sistema], QM.[Status da nota], CB.[Denom Classe Custo]," & vbCrLf
    str_Consulta = str_Consulta & "NS.[Secao], NS.[Responsavel], NS.[Linha], IIF(YEAR(CB.[Data lcto]) IS NULL, YEAR(OS.[Dt Entr]), YEAR(CB.[Data lcto])),"
    str_Consulta = str_Consulta & "IIF(MONTH(CB.[Data lcto]) IS NULL, MONTH(OS.[Dt Entr]), MONTH(CB.[Data lcto]))" & vbCrLf
    str_Consulta = str_Consulta & "ORDER BY CB.[Data lcto], CB.[Cl Custo], OS.[Ordem]"
    
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    If rs_Consulta.EOF = False Then
        
        rs_Consulta.MoveFirst
    
    Else
        
        MsgBox "Recordset vazio"
        Exit Sub
    
    End If
        
    'Apaga todos os dados da planilha.
    '---------------------------------
    Call Clear_Entire_Worksheet(wsOSA)
    
    'Adiciona cabeçalho a planilha.
    '------------------------------
    i = 1
    For Each Fld In rs_Consulta.Fields
        
        wsOSA.Cells(1, i).Value = Fld.Name
        i = i + 1
        
    Next
    

    'Cola Recordset na planilha
    '--------------------------
    wsOSA.Range("A2").CopyFromRecordset rs_Consulta
    
    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    
    ado_Conexao.Close
    Set ado_Conexao = Nothing
    
    Call CorrigeValorLanc(wsOSA)
    wsOSA.Columns.AutoFit
    
    Call OptimizeVBA(False)
    
    MsgBox "Banco de Dados Pronto!"

End Sub

Sub CorrigeValores()
    
    Dim clSO        As Long
    Dim clVal       As Long
    Dim clConta     As Long
    Dim clDtLanc    As Long
    Dim counter     As Long

    FilePath = wbIndicadores.Path & "\"
    FileName = wbIndicadores.Name
    
    ln = wsGGA.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    cl = wsGGA.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 0).Column
    
    clVal = FindColumnNumber("Total", wsGGA)
    clSO = FindColumnNumber("Ordem", wsGGA)
    clConta = FindColumnNumber("Cl Custo", wsGGA)
    clDtLanc = FindColumnNumber("Data lcto", wsGGA)
    
    'Conecta ao arquivo Excel.
    '-------------------------
    Call ConnectXLFile(FilePath, FileName)
    
    'Define o RecordSet.
    '-------------------
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    For i = 2 To ln
    
        Application.StatusBar = "Corrigindo valores: " & i & " de " & ln
        counter = 0
        
        str_Consulta = "SELECT COUNT([Ordem]) AS [Nº OS]" & vbCrLf
        str_Consulta = str_Consulta & "FROM [GGA$]" & vbCrLf
        str_Consulta = str_Consulta & "WHERE [Ordem] = " & wsGGA.Cells(i, clSO).Value & vbCrLf
        str_Consulta = str_Consulta & "GROUP BY [Data lcto], [Cl Custo]"
        
        'Abre Recordset.
        '---------------
        rs_Consulta.Open str_Consulta, ado_Conexao

        If rs_Consulta.EOF = False Then counter = CLng(rs_Consulta("Nº OS"))
        
        wsGGA.Cells(i, clVal).Value = wsGGA.Cells(i, clVal).Value / counter
        
        rs_Consulta.Close
        
    Next i

    'Fecha recordset e limpa a variável.
    '-----------------------------------
    Set rs_Consulta = Nothing
    ado_Conexao.Close
    Set ado_Conexao = Nothing

    Application.StatusBar = Empty
    
    MsgBox "Valores Corrigidos"
    
End Sub
