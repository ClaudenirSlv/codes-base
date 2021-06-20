Attribute VB_Name = "MyFunctions"
Option Explicit

Private Sub ChangeXLReference()
Attribute ChangeXLReference.VB_Description = "Altera referência das linhas e colunas do Excel de A1 para L1C1."
Attribute ChangeXLReference.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ' Altera referência das linhas e colunas do Excel de A1 para L1C1.
    '-----------------------------------------------------------------
        If Application.ReferenceStyle = xlA1 Then
        
        Application.ReferenceStyle = xlR1C1
    
    Else
    
        Application.ReferenceStyle = xlA1
    
    End If
    
End Sub

Private Sub DeleteEntireRow()

    'Substituir "A1" pela linha que deseja apagar.
    '---------------------------------------------
    Range("A1").EntireRow.Delete

End Sub

Private Sub EraseCellValues()

    'Erase the content of a range of cells.
    Range("E21:P21").ClearContents
    
End Sub
