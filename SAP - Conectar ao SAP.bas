Public SapGuiAuto       As Object
Public App              As Object
Public Connection       As Object
Public Session          As Object
Public Sap_Con          As Boolean

Function ConectaSAP()

    ConectaSAP = True
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set Session = Connection.Children(0)
    If Not Err.Number = 0 Then
        MsgBox "SAP não está aberto", vbCritical
        ConectaSAP = False
    End If
    On Error GoTo 0

End Function

'Macro principal

Sub Macro

    Sap_Con = ConectaSAP
    If Not Sap_Con Then Exit Sub

    'MACRO DO EXCEL

End Sub

Sub ClearSapConnection()

    Set SapGuiAuto = Nothing
    Set App = Nothing
    Set Connection = Nothing
    Set Session = Nothing

End Sub