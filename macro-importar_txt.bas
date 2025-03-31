Sub AbrirArquivo()

    Application.StatusBar = "Cabeção, to abrindo o arquivo!"

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Planilha1")
    
    Dim locArquivo As String
    Dim linhaArquivo As String
    Dim dados() As String
    Dim coluna As Long
    Dim linha As Long
    
    locArquivo = "C:\Users\aantunes\Desktop\VBA\Aula 8\ROL.txt"
    
    Open locArquivo For Input As #1
    
    Do Until EOF(1)
    
        Line Input #1, linhaArquivo
        dados = Split(linhaArquivo, vbTab)
        
        linha = linha + 1
        
        For coluna = 0 To UBound(dados)
            ws.Cells(linha, coluna + 1) = dados(coluna)
        Next

    Loop
    
    Close #1
    
    Application.StatusBar = ""
    
End Sub