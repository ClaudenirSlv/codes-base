Sub Email_History_Managers()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range, cell As Range, HtmlContent As String
    Dim dblTotal As Double
    
    Call SetParameters
    
    FilePath = wbKpi.Path & "\"
    FileName = wbKpi.Name
    
    Call ConnectXLFile(FilePath, FileName)

    'Seta RecordSet
    Set rs_Consulta = CreateObject("ADODB.Recordset")

    str_Consulta = "SELECT [Empresa],[Responsável], [Status de alocação], [Cliente], [Ordem], [QA], SUM([Total]) AS 'Total'" & _
                "FROM [passivo$] " & _
                "WHERE [Total] > 0.0049 " & _
                "GROUP BY [Ordem], [Empresa], [Responsável], [Status de alocação], [Cliente], [QA]" & _
                "ORDER BY [Empresa] ASC, SUM([Total]) DESC, [Cliente] ASC"
                
    'Abre Recordset
    rs_Consulta.Open str_Consulta, ado_Conexao
    
    HtmlContent = "<style> table, th, td {text-align: center; border: 1px solid black;"
    HtmlContent = HtmlContent & "border-collapse: collapse;"
    HtmlContent = HtmlContent & "font-family: Arial, Helvetica, sans-serif;"
    HtmlContent = HtmlContent & "font-size: 12px;}"
    HtmlContent = HtmlContent & "th, td {padding: 5px;}"
    HtmlContent = HtmlContent & "tfoot tr td {text-align: right;}</style>"
    HtmlContent = HtmlContent & "<p>Olá!</p>"
    HtmlContent = HtmlContent & "<p>Segue lista de ordens do passivo:</p>"
    HtmlContent = HtmlContent & "<table>"
    HtmlContent = HtmlContent & "<thead>"
    HtmlContent = HtmlContent & "<tr><th>Empresa</th>"
    HtmlContent = HtmlContent & "<th>Responsável</th>"
    HtmlContent = HtmlContent & "<th>Status de Alocação</th>"
    HtmlContent = HtmlContent & "<th>Cliente</th>"
    HtmlContent = HtmlContent & "<th>Ordem</th>"
    HtmlContent = HtmlContent & "<th>QA</th>"
    HtmlContent = HtmlContent & "<th>Total da Ordem</th></tr>"
    HtmlContent = HtmlContent & "</thead>"

    If rs_Consulta.EOF = False Then rs_Consulta.MoveFirst
    
    Do Until rs_Consulta.EOF = True
        HtmlContent = HtmlContent & "<tbody>"
        HtmlContent = HtmlContent & "<tr>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Empresa") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Responsável") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Status de alocação") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Cliente") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("Ordem") & "</td>"
        HtmlContent = HtmlContent & "<td>" & rs_Consulta("QA") & "</td>"
        HtmlContent = HtmlContent & "<td>" & FormatCurrency(rs_Consulta("'Total'"), 2) & "</td>"
        HtmlContent = HtmlContent & "</tr>"
        HtmlContent = HtmlContent & "</tbody>"
        dblTotal = dblTotal + rs_Consulta("'Total'")
        rs_Consulta.MoveNext
    Loop
    HtmlContent = HtmlContent & "<tfoot>"
    HtmlContent = HtmlContent & "<tr>"
    HtmlContent = HtmlContent & "<td colspan=6><strong>Total</strong></td>"
    HtmlContent = HtmlContent & "<td><strong>" & FormatCurrency(dblTotal, 2) & "</strong></td>"
    HtmlContent = HtmlContent & "</tfoot>"
    HtmlContent = HtmlContent & "</table>"
    HtmlContent = HtmlContent & "<p>São consideradas ordens do passivo as ordens criadas em anos anteriores e "
    HtmlContent = HtmlContent & "lançamentos financeiros feitos até o último dia do ano anterior.</p>"
    HtmlContent = HtmlContent & "<p>Atenciosamente,<br>"
    HtmlContent = HtmlContent & "Claudenir da Silva</p>"
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    With OutMail
        .To = "clausw@weg.net; luizfernando@weg.net;miliane@weg.net"
        .Cc = "glauco@weg.net;joaotabalipa@weg.net"
        '.Bcc = Range("B3").Value
        .Subject = "PASSIVO DE ORDENS - " & UCase(Format(Now(), "DD/MMMM/YYYY"))
        .HTMLBody = HtmlContent
        .Display
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing

End Sub