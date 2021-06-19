
Sub GetDataKOB1()
    Dim DtFim As String
    
    Application.StatusBar = "Importando dados da KOB1"
    
    Set wb = Workbooks.Open("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP\IW72.xlsx")
    Set ws = wb.Worksheets(1)
    
    ln = LastRowNumber(ws)
    ws.Range("A2:A" & ln).Copy
    Call AccessTcode("KOB1")

    Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").SetFocus
    Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").CaretPosition = 0
    Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").showContextMenu
    Session.findById("wnd[0]/usr").selectContextMenuItem "DELACTX"
    Session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "411075004"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "411075007"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "411075008"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "411075083"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "411075117"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "411075118"
    Session.findById("wnd[1]").sendVKey 0
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'No campo BUDAT-LOW é informado a data de início da pesquisa de lançamentos.
    'Será informado 01.01.2018 para os indicadores, porque todas as ordens anteriores estão encerradas.
    'Obs: não é necessário informar o limite superior porque o SAP já traz esta informação.
    DtFim = LastDay_CurrentMonth
    DtFim = Replace(DtFim, "/", ".")
    Session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").Text = "01.01.2018"
    Session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").Text = DtFim
    Session.findById("wnd[0]/usr/ctxtP_DISVAR").SetFocus
    Session.findById("wnd[0]/usr/ctxtP_DISVAR").CaretPosition = 9
    Session.findById("wnd[0]/usr/btnBUT1").press
    Session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").Text = "1048576"
    Session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").CaretPosition = 7
    Session.findById("wnd[1]").sendVKey 0
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]").sendVKey 43
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "KOB1.xlsx"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 9
    Session.findById("wnd[1]").sendVKey 11
    wb.Close
    
    Set ws = Nothing
    Set wb = Nothing

End Sub
