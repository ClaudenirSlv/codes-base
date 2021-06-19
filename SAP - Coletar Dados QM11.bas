Option Explicit

Sub GetDataQM11()
    
    Application.StatusBar = "Importando dados QM11"
    
    Set wb = Workbooks.Open("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP\IW59.xlsx")
    Set ws = wb.Worksheets(1)
    
    ln = ws.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    
    ws.Range("D2:D" & ln).Copy
    
    Call AccessTcode("QM11")
    Session.findById("wnd[0]/usr/ctxtQMDAT-LOW").showContextMenu
    Session.findById("wnd[0]/usr").selectContextMenuItem "DELACTX"
    Session.findById("wnd[0]").sendVKey 16
    Session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]").sendVKey 8
    
    wb.Close False
    
    'Limpa as variáveis para usar elas novamente.
    Set ws = Nothing
    Set wb = Nothing
    
    Session.findById("wnd[0]").sendVKey 8 'Depois deste ponto estou na lista.
    Session.findById("wnd[0]/tbar[1]/btn[16]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    'Definir caminho e nome do arquivos a serem salvos.
    FilePath = "Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Relatórios de Assistência Técnica\Indicadores\Dados do SAP\"
    FileName = "QM11.xlsx"
    
    Set wb = Workbooks("Planilha em Basis (1)")
    
    'Salvar o arquivo no caminho com nome especificado.
    wb.SaveAs FileName:=FilePath & FileName
    wb.Close
    
    FilePath = Empty
    FileName = Empty
    
    Set wb = Nothing
    Application.StatusBar = Empty
    
End Sub
