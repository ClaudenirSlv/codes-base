Sub GetSocial()

    ' Written by Philip Treacy
    ' https://www.myonlinetraininghub.com/web-scraping-with-vba
    '
    'To use HTMLDocument you need to set a reference to Tools -> References -> Microsoft HTML Object Library
    Dim HTML As New HTMLDocument
    Dim http As Object
    Dim links As Object
    Dim link As HTMLHtmlElement
    Dim counter As Long
    Dim website As Range
    Dim row As Long
    Dim continue As Boolean
    Dim respHead As String
    
    Application.ScreenUpdating = False
    
    ' The row where website addresses start
    row = 24
    continue = True
    
    ' XMLHTTP gives errors where ServerXMLHTTP does not
    ' even when using the same URL's
    'Set http = CreateObject("MSXML2.XMLHTTP")
        
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    Do While continue
    
        ' Could set this to first cell with URL then OFFSET columns to get next web site
        Set website = Range("A" & row)
        
        If Len(website.Value) < 1 Then
        
            continue = False
            Exit Sub
            
        End If
        
        If website Is Nothing Then
        
            continue = False
        
        End If
        
        'Debug.Print website
    
        With http
        
            On Error Resume Next
            .Open "GET", website.Value, False
            .send
            
            ' If Err.Num is not 0 then an error occurred accessing the website
            ' This checks for badly formatted URL's. The website can still return an error
            ' which should be checked in .Status
            
            'Debug.Print Err.Number
            
            ' Clear the row of any previous results
            Range("B" & row & ":e" & row).Clear
            
            ' If the website sent a valid response to our request
            If Err.Number = 0 Then
            
                ' I left this here for you to experiment with/check.
                ' responseText contains the HTML web page
                'Debug.Print Len(.responseText)
                'Debug.Print .responseText
                'Debug.Print Len(.responseXML.XML)
                'Debug.Print (.responseXML.XML)
                
                'respHead = http.getAllResponseHeaders
                'Debug.Print Len(respHead)
                'Debug.Print respHead

                If .Status = 200 Then
        
                    HTML.body.innerHTML = http.responseText
                            
                    Set links = HTML.getElementsByTagName("a")
    
                    For Each link In links
    
                        If InStr(UCase(link.outerHTML), "LINKEDIN") Then
                                
                            website.Offset(0, 1).Value = link.href
        
                        End If
        
        
                        If InStr(UCase(link.outerHTML), "FACEBOOK") Then
        
                            website.Offset(0, 2).Value = link.href
        
                        End If
                
                        If InStr(UCase(link.outerHTML), "TWITTER") Then
        
                            website.Offset(0, 3).Value = link.href
        
                        End If
        
                        If InStr(UCase(link.outerHTML), "YOUTUBE") Then
        
                            website.Offset(0, 4).Value = link.href
        
                        End If

                    Next
            
                End If
            
                Set website = Nothing
            
            Else
        
                    'Debug.Print "Error loading page"
                    website.Offset(0, 1).Value = "Error with website address"
            
            End If
                
                    On Error GoTo 0

        End With
        
        row = row + 1
    
    Loop

    Application.ScreenUpdating = True
    
End Sub