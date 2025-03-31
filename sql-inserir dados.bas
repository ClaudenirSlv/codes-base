'Esta macro insere dados de uma planilha Excel em um banco de dados MDB.

Sub InserirDados()

    'Variáveis
    '---------
    Dim bd                   As DAO.Database
    Dim rst                  As DAO.Recordset
    Dim sql                  As String
    Dim i                    As Long
    Dim ln                   As Long
    Dim wb                   As Workbook
    Dim ws                   As Worksheet

    set wb = ThisWorkbook
    set ws = wb.Worksheets("Plan1")

    ln = ws.cells(Rows.Count,1).End(xlUp).Offset(0,0).Row

    Set bd = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Controle de Equipamentos\DataBaseEQC.0.0.MDB", False, False)

    For i=2 to ln

        sql = "INSERT INTO tblEquipments (Patrimonio, Num_Metrologia, Marca, Modelo, Descricao, StatusEquipamento)" & vbCrLf
        sql = sql & "VALUES ('" & ws.cells(i,2).Value & "', '" & ws.cells(i,3).Value & "', '" & ws.cells(i,4).Value & "', '" & ws.cells(i,5).Value & "', '" & ws.cells(i,6).Value & "'," & vbCrLf
        sql = sql & "'" & ws.cells(i,1).Value & "')"

        bd.Execute(sql)
        bd.Close

    Next i

    MsgBox "Concluído"

End Sub