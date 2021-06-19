Sub AbrirNavegador()

    Dim nav As Object
    Dim url As String
    
    Set nav = CreateObject("InternetExplorer.Application")
    
    nav.Visible = True
    
    'url = "https://www.google.com/search?q=cotacao usd"
    url = "https://info.flightmapper.net/flight/LATAM_Airlines_Group_LA_3014"
    
    
    nav.navigate url
    
    Do While nav.READYSTATE <> 4
        DoEvents
    Loop
    
    Dim data As String
    
    data = Format(Sheets(1).Range("A2").Value, "YYYY-MM-DD")
    
    nav.Document.querySelector("input[id='d_date_input']").Value = data
    nav.Document.forms(1).submit
    
    Dim x As String
    Dim y As String
    
    Application.Wait (Now + TimeValue("00:00:02"))
    
    'nav.Document.querySelector("input[id='knowledge-currency__src-input']").Value = 15
    x = nav.Document.querySelector("input[id='a61j6 vk_gy vk_sh Hg3mWc']").Value
    'y = nav.Document.getElementByID("knowledge-currency__tgt-input").Value
    
    Dim ln As Long
    
    
    'corrigir este Do
    Do
    
        ln = ln + 1
        If Cells(ln, 2) = "" Then
            Cells(ln, 1) = Now()
            Cells(ln, 2) = x
        Else
            Exit Do
        End If
    
    Loop

End Sub