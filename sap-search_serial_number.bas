Sub SearchSerial()

    Call SetParameters
    Call ConectaSAP
    
    lnLM = wsLM.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    lnNS = wsNS.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row

    For i = 2 To lnLM
        Session.findById("wnd[0]").maximize
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/niq03"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/usr/ctxtRISA0-MATNR").Text = wsLM.Cells(i,1).Value
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/tbar[1]/btn[16]").press
        Session.findById("wnd[1]/tbar[0]/btn[0]").press
        Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
        Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus
        Session.findById("wnd[1]/tbar[0]/btn[0]").press
        Session.findById("wnd[1]/tbar[0]/btn[0]").press
        Set wb = ActiveWorkbook
    Next i

End Sub