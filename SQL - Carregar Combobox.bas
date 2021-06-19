Dim caixa As Explorer
Dim omail As MailItem
Dim i As Integer
Dim banco As DAO.Database
Dim tabela As DAO.Recordset
Dim sql As String
Dim indice As Integer

Function CarregaCombo()
    
    cBoxCliente.Clear
    cBoxPlataforma.Clear
    cBoxUnidade.Clear
    cBoxNotaServico.Clear
    cBoxOrdemServico.Clear
    cBoxProblema.Clear
        
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    
    Set tabela = banco.OpenRecordset("SELECT Cliente FROM tbCategorias GROUP BY Cliente")
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        If Not tabela("Cliente") = "" Then cBoxCliente.AddItem tabela("Cliente")
        tabela.MoveNext
    Loop
    tabela.Close
    
    Set tabela = banco.OpenRecordset("SELECT Plataforma FROM tbCategorias GROUP BY Plataforma")
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        If Not tabela("Plataforma") = "" Then cBoxPlataforma.AddItem tabela("Plataforma")
        tabela.MoveNext
    Loop
    tabela.Close
       
    Set tabela = banco.OpenRecordset("SELECT Unidade FROM tbCategorias GROUP BY Unidade")
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        If Not tabela("Unidade") = "" Then cBoxUnidade.AddItem tabela("Unidade")
        tabela.MoveNext
    Loop
    tabela.Close
    
    Set tabela = banco.OpenRecordset("SELECT NotaServico FROM tbCategorias GROUP BY NotaServico")
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        If Not tabela("NotaServico") = "" Then cBoxNotaServico.AddItem tabela("NotaServico")
        tabela.MoveNext
    Loop
    tabela.Close
    
    Set tabela = banco.OpenRecordset("SELECT OrdemServico FROM tbCategorias GROUP BY OrdemServico")
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        If Not tabela("OrdemServico") = "" Then cBoxOrdemServico.AddItem tabela("OrdemServico")
        tabela.MoveNext
    Loop
    tabela.Close
    
    Set tabela = banco.OpenRecordset("SELECT Problema FROM tbCategorias GROUP BY Problema")
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        If Not tabela("Problema") = "" Then cBoxProblema.AddItem tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close
    
    cBoxCliente.Text = ""
    cBoxPlataforma.Text = ""
    cBoxUnidade.Text = ""
    cBoxNotaServico.Text = ""
    cBoxOrdemServico.Text = ""
    cBoxProblema.Text = ""

End Function

Function Filtra_Cliente()

    sql = "SELECT * FROM tbCategorias WHERE Cliente = '" & cBoxCliente.Value & "' ORDER BY Plataforma, Unidade, NotaServico"
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset(sql)
    
    ListBox1.Clear
    
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        ListBox1.AddItem
        If Not tabela("Cliente") = "" Then ListBox1.List(ListBox1.ListCount - 1, 0) = tabela("Cliente")
        If Not tabela("Plataforma") = "" Then ListBox1.List(ListBox1.ListCount - 1, 1) = tabela("Plataforma")
        If Not tabela("Unidade") = "" Then ListBox1.List(ListBox1.ListCount - 1, 2) = tabela("Unidade")
        If Not tabela("NotaServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 3) = tabela("NotaServico")
        If Not tabela("OrdemServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 4) = tabela("OrdemServico")
        If Not tabela("Problema") = "" Then ListBox1.List(ListBox1.ListCount - 1, 5) = tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close

End Function

Function Filtra_OS()

    sql = "SELECT * FROM tbCategorias WHERE OrdemServico = '" & cBoxOrdemServico & "'"
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset(sql)
    
    ListBox1.Clear
    
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        ListBox1.AddItem
        If Not tabela("Cliente") = "" Then ListBox1.List(ListBox1.ListCount - 1, 0) = tabela("Cliente")
        If Not tabela("Plataforma") = "" Then ListBox1.List(ListBox1.ListCount - 1, 1) = tabela("Plataforma")
        If Not tabela("Unidade") = "" Then ListBox1.List(ListBox1.ListCount - 1, 2) = tabela("Unidade")
        If Not tabela("NotaServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 3) = tabela("NotaServico")
        If Not tabela("OrdemServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 4) = tabela("OrdemServico")
        If Not tabela("Problema") = "" Then ListBox1.List(ListBox1.ListCount - 1, 5) = tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close
    
End Function

Function Filtra_NS()

    
    sql = "SELECT * FROM tbCategorias WHERE NotaServico = '" & cBoxNotaServico & "'"
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset(sql)
    
    ListBox1.Clear
    
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        ListBox1.AddItem
        If Not tabela("Cliente") = "" Then ListBox1.List(ListBox1.ListCount - 1, 0) = tabela("Cliente")
        If Not tabela("Plataforma") = "" Then ListBox1.List(ListBox1.ListCount - 1, 1) = tabela("Plataforma")
        If Not tabela("Unidade") = "" Then ListBox1.List(ListBox1.ListCount - 1, 2) = tabela("Unidade")
        If Not tabela("NotaServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 3) = tabela("NotaServico")
        If Not tabela("OrdemServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 4) = tabela("OrdemServico")
        If Not tabela("Problema") = "" Then ListBox1.List(ListBox1.ListCount - 1, 5) = tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close


End Function

Function Filtra_Plataforma()

    sql = "SELECT * FROM tbCategorias WHERE Plataforma = '" & cBoxPlataforma.Value & "' ORDER BY Cliente, Unidade, NotaServico"
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset(sql)
    
    ListBox1.Clear
    
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        ListBox1.AddItem
        If Not tabela("Cliente") = "" Then ListBox1.List(ListBox1.ListCount - 1, 0) = tabela("Cliente")
        If Not tabela("Plataforma") = "" Then ListBox1.List(ListBox1.ListCount - 1, 1) = tabela("Plataforma")
        If Not tabela("Unidade") = "" Then ListBox1.List(ListBox1.ListCount - 1, 2) = tabela("Unidade")
        If Not tabela("NotaServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 3) = tabela("NotaServico")
        If Not tabela("OrdemServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 4) = tabela("OrdemServico")
        If Not tabela("Problema") = "" Then ListBox1.List(ListBox1.ListCount - 1, 5) = tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close

End Function

Function Filtra_Unidade()

    sql = "SELECT * FROM tbCategorias WHERE Unidade = '" & cBoxUnidade.Value & "' ORDER BY Cliente, Plataforma, NotaServico"
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset(sql)
    
    ListBox1.Clear
    
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        ListBox1.AddItem
        If Not tabela("Cliente") = "" Then ListBox1.List(ListBox1.ListCount - 1, 0) = tabela("Cliente")
        If Not tabela("Plataforma") = "" Then ListBox1.List(ListBox1.ListCount - 1, 1) = tabela("Plataforma")
        If Not tabela("Unidade") = "" Then ListBox1.List(ListBox1.ListCount - 1, 2) = tabela("Unidade")
        If Not tabela("NotaServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 3) = tabela("NotaServico")
        If Not tabela("OrdemServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 4) = tabela("OrdemServico")
        If Not tabela("Problema") = "" Then ListBox1.List(ListBox1.ListCount - 1, 5) = tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close

End Function

Function Filtra_Problema()

    sql = "SELECT * FROM tbCategorias WHERE Problema = '" & cBoxProblema.Value & "' ORDER BY Cliente, Plataforma, Unidade, NotaServico"
    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset(sql)
    
    ListBox1.Clear
    
    If tabela.EOF = False Then tabela.MoveFirst
    Do Until tabela.EOF = True
        ListBox1.AddItem
        If Not tabela("Cliente") = "" Then ListBox1.List(ListBox1.ListCount - 1, 0) = tabela("Cliente")
        If Not tabela("Plataforma") = "" Then ListBox1.List(ListBox1.ListCount - 1, 1) = tabela("Plataforma")
        If Not tabela("Unidade") = "" Then ListBox1.List(ListBox1.ListCount - 1, 2) = tabela("Unidade")
        If Not tabela("NotaServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 3) = tabela("NotaServico")
        If Not tabela("OrdemServico") = "" Then ListBox1.List(ListBox1.ListCount - 1, 4) = tabela("OrdemServico")
        If Not tabela("Problema") = "" Then ListBox1.List(ListBox1.ListCount - 1, 5) = tabela("Problema")
        tabela.MoveNext
    Loop
    tabela.Close

End Function

Private Sub cBoxCliente_Change()

    If Not cBoxCliente.Value = "" Then Call Filtra_Cliente

End Sub

Private Sub cBoxNotaServico_Change()
    
    If Not cBoxNotaServico.Value = "" Then Call Filtra_NS
    
End Sub

Private Sub cBoxOrdemServico_Change()

    If Not cBoxOrdemServico.Value = "" Then Call Filtra_OS

End Sub

Private Sub cBoxPlataforma_Change()
    
    If Not cBoxPlataforma.Value = "" Then Call Filtra_Plataforma
    
End Sub

Private Sub cBoxProblema_Change()

    If Not cBoxProblema.Value = "" Then Call Filtra_Problema

End Sub

Private Sub cBoxUnidade_Change()
    
    If Not cBoxUnidade.Value = "" Then Call Filtra_Unidade
    
End Sub

Private Sub cmdAdicionar_Click()

    Set banco = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Arquivamento_Email\dbEmail.mdb", False, False)
    Set tabela = banco.OpenRecordset("SELECT * FROM tbCategorias")
    tabela.AddNew
    If Not cBoxCliente.Value = "" Then tabela("Cliente") = cBoxCliente.Value
    If Not cBoxPlataforma.Value = "" Then tabela("Plataforma") = cBoxPlataforma.Value
    If Not cBoxUnidade.Value = "" Then tabela("Unidade") = cBoxUnidade.Value
    If Not cBoxNotaServico.Value = "" Then tabela("NotaServico") = cBoxNotaServico.Value
    If Not cBoxOrdemServico.Value = "" Then tabela("OrdemServico") = cBoxOrdemServico.Value
    If Not cBoxProblema.Value = "" Then tabela("Problema") = cBoxProblema.Value
    tabela.Update
    tabela.Close
    
    Call CarregaCombo
    
    
    
End Sub

Private Sub CommandButton1_Click()

MsgBox ListBox1.ListIndex

End Sub

Private Sub UserForm_Initialize()
    
    Call CarregaCombo
    
End Sub

Private Sub cmdGrava_Click()
    
    Set caixa = Application.ActiveExplorer
     
    If chkMover.Value = True Then Set oFolder = Outlook.Session.PickFolder
    If chkCopiar.Value = True Then Set oFolder = Outlook.Session.PickFolder

    For i = 1 To caixa.Selection.Count
        Set omail = caixa.Selection.Item(i)
        indice = ListBox1.ListIndex
        If indice > -1 Then
            If Not ListBox1.List(indice, 0) = "" Then omail.UserProperties.Add("Cliente", olText) = ListBox1.List(indice, 0)
            If Not ListBox1.List(indice, 1) = "" Then omail.UserProperties.Add("Plataforma", olText) = ListBox1.List(indice, 1)
            If Not ListBox1.List(indice, 2) = "" Then omail.UserProperties.Add("Unidade", olText) = ListBox1.List(indice, 2)
            If Not ListBox1.List(indice, 3) = "" Then omail.UserProperties.Add("Nota de Serviço", olText) = ListBox1.List(indice, 3)
            If Not ListBox1.List(indice, 4) = "" Then omail.UserProperties.Add("Ordem de Serviço", olText) = ListBox1.List(indice, 4)
            If Not ListBox1.List(indice, 5) = "" Then omail.UserProperties.Add("Problema", olText) = ListBox1.List(indice, 5)
            omail.Save
            If chkMover.Value = True Then omail.Move oFolder
            If chkCopiar.Value = True Then
                Set emailCopy = omail.Copy
                emailCopy.Move oFolder
            End If
        Else
            MsgBox "Selecione uma categoria"
        End If
    Next i
    
    'Unload Me
End Sub
