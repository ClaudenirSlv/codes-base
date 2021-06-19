Sub SQL_Excel()

    'Conectar ao arquivo Excel.
    '**********************************************************************************
    FilePath = Fldr.Path & "\"
    FileName = Fl.Name
    
    Call ConnectXLFile(FilePath, FileName)
    '**********************************************************************************
    
    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")
    
    'Define da Query
    STR_consulta = "SELECT [Empresa], [Data de Lançamento], [Conta do Razão], [Nº documento], [Ordem], [Montante em moeda interna], " & _
                    "[Centro de lucro], [Centro custo], [Elemento PEP], [Segmento], [Referência], [Texto] " & _
                    "FROM [Sheet1$] " & _
                    "WHERE [Conta do Razão]='411075004' OR [Conta do Razão]='411075007' OR [Conta do Razão]='411075008' OR [Conta do Razão]='411075083' OR " & _
                        "[Conta do Razão]='411075117' OR [Conta do Razão]='411075118' " & _
                    "ORDER BY [Data de Lançamento] DESC"

    'Abre Recordset
    rs_Consulta.Open STR_consulta, ado_Conexao
            
    ln = wsFAG.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
    
    'Cola Recordset na planilha
    wsFAG.Range("A" & ln).CopyFromRecordset rs_Consulta
    
    'FechaConexão
    rs_Consulta.Close
    Set rs_Consulta = Nothing
    ado_Conexao.Close
    Set ado_Conexao = Nothing

End Sub

Sub ConnectXLFile(caminho As String, arquivo As String)

    'Define string de Conexão
    str_Conexao = _
        "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & _
                "DSN=TESTE_SQL;DBQ=" & caminho & arquivo & ";" _
                & "ReadOnly=0;DefaultDir=" & caminho & ";" _
                & "DriverId=1046;FIL=excel 12.0;MaxBufferSize=2048;PageTimeout=5;"
    
    'Seta ADODB
    Set ado_Conexao = CreateObject("ADODB.Connection")
    
    'Abre Conexão
    ado_Conexao.Open str_Conexao

End Sub