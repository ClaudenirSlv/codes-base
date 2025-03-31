Option Explicit

Sub OptimizeVBA(isOn As Boolean)
    With Application
        .Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not(isOn)
        .ScreenUpdating = Not(isOn)
        .DisplayAlerts = Not(isOn)
    End With
    ActiveSheet.DisplayPageBreaks = Not(isOn)
End Sub

'Some macro
Sub ExampleMacro()
    OptimizeVBA True
    'Your code here
    OptimizeVBA False
End Sub