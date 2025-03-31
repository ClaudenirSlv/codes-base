Sub WaitExporting
    
    Dim Directory As String, File As String, FindIt As String
    
    Directory = wb.Path & "\" 'path for the file
    File = Directory & "KOB1.XLSX" 'name of the file along with the path

    FindIt = Dir(File)
    While Len(FindIt) = 0: FindIt = Dir(File): Wend

End Sub