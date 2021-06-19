Sub CalculateRunTime_Minutes()
    'PURPOSE: Determine how many minutes it took for code to completely run
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim StartTime As Double
    Dim MinutesElapsed As String

    'Remember time when macro starts
    StartTime = Timer

    '*****************************
    'Insert Your Code Here...
    '*****************************

    'Determine how many seconds code took to run
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

    'Notify user in seconds
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub