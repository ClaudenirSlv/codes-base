Sub CopyFiles()

    Dim Fl As Object
    Dim Fldr As Object
    Dim FSO As Object
    Dim lnSrc As Long
    Dim ln As Long
    Dim wb As Workbook
    Dim wbSrc As Workbook
    Dim ws As Worksheet
    Dim wsSrc As Worksheet
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Fldr = FSO.GetFolder("C:\Users\claudenir\Documents\Macros\Copiar Arquivos\Copiar")
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)

    Application.ScreenUpdating = False
    Application.Calculate = xlManual

    'Look at each file in the folder
    For Each Fl In Fldr.Files
        If InStr(1, Right(Fl.Name, 5), ".csv") <> 0 Then
            Set wbSrc = Application.Workbooks.Open(Fl.Path)
            Set wsSrc = wbSrc.Worksheets(1) 
            lnSrc = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row        
            ln = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
            wsSrc.Range("A2:" & "A" & lnSrc).Select
            Selection.Copy ws.Cells(ln, 1)
            Set wsSrc = Nothing
            wbSrc.Close 0
            Set wbSrc = Nothing            
        End If  
    Next

    Application.ScreenUpdating = true
    Application.Calculate = xlAutomatic

    MsgBox "Arquivos copiados com sucesso!"

End Sub