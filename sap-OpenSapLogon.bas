Sub SapConn()

Dim Appl As Object
Dim Connection As Object
Dim session As Object
Dim WshShell As Object
Dim SapGui As Object

'Of course change for your file directory
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Do Until WshShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set WshShell = Nothing

Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine
Set Connection = Appl.Openconnection("paste name of module", _
    True)
Set session = Connection.Children(0)

'if You need to pass username and password
session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "900"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "user"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "password"
session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"

If session.Children.Count > 1 Then

    answer = MsgBox("You've got opened SAP already," & _
"please leave and try again", vbOKOnly, "Opened SAP")

    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT3").Select
    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT3").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    Exit Sub

End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]").sendVKey 0 'ENTER

'and there goes your code in SAP

End Sub