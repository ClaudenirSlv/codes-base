Attribute VB_Name = "MyFunctions"
Option Explicit

Function FindColumnLetter(label As String, sh As Worksheet)

    Dim col As Long
    
    Dim letra As String
    Dim lt() As String
    
    cl = LastColumnNumber(sh)
    
    For col = 1 To cl
    
        If label = sh.Cells(1, col) Then
            
                letra = sh.Cells(1, col).Address(True, False)
                lt = Split(letra, "$")
                Exit For
        
        End If
                
    Next col

    FindColumnLetter = lt(0)

End Function

Function FindColumnNumber(label As String, sh As Worksheet)

    Dim col As Long
    Dim cl As Integer
    Dim letra As String
    Dim num As String
    
    cl = sh.Cells(77, Columns.Count).End(xlToLeft).Offset(0, 0).Column
    
    For col = 1 To cl
    
        If label = sh.Cells(1, col) Then
            
                letra = sh.Cells(1, col).Address(True, False)
                num = col
                Exit For
        
        End If
                
    Next col

    FindColumnNumber = num

End Function

Function LastRowNumber(sh As Worksheet)

    Dim ln As Long
    
    ln = sh.Cells(Rows.Count, 2).End(xlUp).Offset(0, 0).Row
    
    LastRowNumber = ln
    
End Function

Function LastColumnNumber(sh As Worksheet)

    Dim cl As Long
    
    cl = sh.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 0).Column
    
    LastColumnNumber = cl
    
End Function

Sub ClearWorksheet(sh As Worksheet)

    ln = LastRowNumber(sh)
    cl = LastColumnNumber(sh)
    sh.Range(sh.Cells(2, 1), sh.Cells(ln, cl)).Clear
    
End Sub

Sub addSheet(ws As Worksheet)

    Worksheets.Add

End Sub

Sub DeleteWorksheet(ws As Worksheet)

    ws.Delete
    Set ws = Nothing
    
End Sub

Function LastDay_CurrentMonth()
    
    Dim strDate As String
    Dim strMonth As String
    Dim stryear As String

    strDate = DateValue(Now)
    strMonth = Month(strDate)
    stryear = Year(strDate)
    LastDay_CurrentMonth = DateSerial(stryear, strMonth + 1, 0)

End Function