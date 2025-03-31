Sub GetDataIW72()

    Application.StatusBar = "Importando dados IW72"
    
    Call AccessTcode("IW72")
    Session.findById("wnd[0]/usr/chkDY_HIS").Selected = True
    Session.findById("wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "ZSGI"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "ZSGE"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "ZSFE"
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/usr/ctxtERDAT-LOW").Text = "01.01.2018"
    Session.findById("wnd[0]/usr/ctxtERDAT-HIGH").Text = Format(Now, "dd.mm.yyyy")
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/btn%_BUKRS_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1001"
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1007"
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press 'Aqui chegamos na lista.
    Session.findById("wnd[0]/tbar[1]/btn[16]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    'Definir caminho e nome do arquivos a serem salvos.
    FilePath = "Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP\"
    FileName = "IW72"
    Set wb = Workbooks("Planilha em Basis (1)")
    'Salvar o arquivo no caminho com nome especificado.
    wb.SaveAs FileName:=FilePath & FileName
    wb.Close
    
    Set wb = Nothing
    FilePath = Empty
    FileName = Empty
    
    Application.StatusBar = Empty

End Sub