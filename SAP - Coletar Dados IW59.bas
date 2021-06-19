Option Explicit

Sub GetDataIW59()
        
    Application.StatusBar = "Importando dados da IW59"
    
    Set wb = Workbooks.Open("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP\IW72.xlsx")
    Set ws = wb.Worksheets(1)
    
    ln = LastRowNumber(ws)
    ws.Range("A2:A" & ln).Copy
    
    Call AccessTcode("IW59")
    Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    Session.findById("wnd[0]/usr/cmbDY_PARVW").Key = " "
    Session.findById("wnd[0]/usr/ctxtDY_PARNR").Text = ""
    Session.findById("wnd[0]/usr/ctxtVARIANT").Text = "/cs-war"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    
    'Limpa as variáveis para usar elas novamente.
    wb.Close False
    Set ws = Nothing
    Set wb = Nothing
    
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/tbar[1]/btn[16]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    'Definir caminho e nome do arquivos a serem salvos.
    FilePath = "Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP\"
    FileName = "IW59"
    
    Set wb = Workbooks("Planilha em Basis (1)")
    'Salvar o arquivo no caminho com nome especificado.
    wb.SaveAs FileName:=FilePath & FileName
    wb.Close
    
    FilePath = Empty
    FileName = Empty
    
    Set wb = Nothing
    
    Application.StatusBar = Empty
        
End Sub
