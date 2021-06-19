Sub Convert_to_Number(ws As Worksheet, cl As String)

    Application.StatusBar = "Convertendo texto para número..."
    
    ws.Activate
    ws.Columns(cl & ":" & cl).Select
    ws.Range(cl & "200000").Activate
    Selection.TextToColumns Destination:=Range(cl & "1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

End Sub