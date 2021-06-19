Attribute VB_Name = "AccessTransaction"
Option Explicit


Call AccessTcode("IW59")


Sub AccessTcode(Tcode As String)
    
    If Session.findById("wnd[0]").Text = "SAP Easy Access" Then
        Session.findById("wnd[0]/tbar[0]/okcd").Text = Tcode
    Else
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/tbar[0]/okcd").Text = Tcode
    End If
    Session.findById("wnd[0]").sendVKey 0
    
End Sub


