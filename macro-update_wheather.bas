Sub Teste()

    Dim xmlhttp As New MSXML2.XMLHTTP60, myurl As String, xmlresponse As New DOMDocument60

    myurl = "http://api.openweathermap.org/data/2.5/weather?apikey=afadf582a41d7e0d82f8f6f3ad52b2ff&mode=xml&units=metric&q=" & Sheets(1).Range("A2").Value
    xmlhttp.Open "GET", myurl, False
    xmlhttp.send
    xmlresponse.LoadXML (xmlhttp.responseText)
    Range("B2").Value = xmlresponse.SelectNodes("//feed/entry/content/properties/simbolo")(0).Text
    Range("C2").Value = xmlresponse.SelectNodes("//current/temperature/@min")(0).Text
    Range("D2").Value = xmlresponse.SelectNodes("//current/temperature/@value")(0).Text
    Range("E2").Value = xmlresponse.SelectNodes("//current/humidity/@value")(0).Text
    'MsgBox (xmlresponse.getElementsByTagName("temperature")(0).Attributes(1).Text)  Alternate method to parse XML

End Sub