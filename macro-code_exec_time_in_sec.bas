Sub CalculateRunTime_Seconds()
    'PURPOSE: Determine how many seconds it took for code to completely run
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim StartTime As Double
    Dim SecondsElapsed As String

    'Remember time when macro starts
    StartTime = Timer

    '*****************************
    'Insert Your Code Here...
    '*****************************

    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub