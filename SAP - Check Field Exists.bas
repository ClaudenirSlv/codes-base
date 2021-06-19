If Not session.findById("wnd[0]/usr/subBLOCK1:SAPLKOBS:0200/ctxtIONRA-AUFNR", False) is Nothing Then

    text = session.findById("wnd[0]/usr/subBLOCK1:SAPLKOBS:0200/ctxtIONRA-AUFNR").Text

End If