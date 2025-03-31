Sub sbSaveExcelDialog()

    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = "Sample Output"
    sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=InitialName, fileFilter:="Excel Files (*.xlsm), *.xlsm")

    If sFileSaveName <> False Then
        ActiveWorkbook.SaveAs sFileSaveName
    End If

End Sub