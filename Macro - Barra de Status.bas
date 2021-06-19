Option Explicit

Dim CurrentProgress As Double
Dim ProgressPercentage As Double
Dim BarWidth As Long
Dim ln As Long

Sub TesteForm()

    Call InitProgressBar
    ln = 100000
    
    For i = 1 To ln
    
        CurrentProgress = i / ln
        BarWidth = frmProgress.fraBorder.Width * CurrentProgress
        ProgressPercentage = Round(CurrentProgress * 100, 0)
        
        frmProgress.lblBar.Width = BarWidth
        frmProgress.lblText.Caption = ProgressPercentage & "% Complete"
        
        DoEvents
        
    Next i
    
    Unload frmProgress

End Sub

Sub InitProgressBar()

    With frmProgress
    
        .lblBar.Width = 0
        .lblText.Caption = "0% Complete"
        .Show vbModeless
        
    End With

End Sub
