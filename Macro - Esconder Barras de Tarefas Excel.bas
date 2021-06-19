Option Explicit

Sub Normaliza_Janela()
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        ActiveWindow.DisplayHeadings = True
        Application.Caption = ""
        Application.WindowState = xlMaximized
    
    With ActiveWindow
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayWorkbookTabs = True
            .DisplayHeadings = True
            .DisplayZeros = True
            .DisplayGridlines = True
    End With
    
End Sub

Private Sub Workbook_Open()

    Dim maxWidth As Integer
    Dim maxHeight As Integer
    
        Sheets("INICIO").Activate
    
        Application.WindowState = xlMaximized
        maxWidth = Application.Width
        maxHeight = Application.Height
        Call CenterApp(maxWidth, maxHeight)
        
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        Application.Caption = "RELATÓRIO DE GARANTIAS"
         
    With ActiveWindow
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayWorkbookTabs = False
            .DisplayHeadings = False
            .DisplayZeros = False
            .DisplayGridlines = False
            .Height = 400
            .Width = 500
    End With
End Sub

Sub CenterApp(maxWidth As Integer, maxHeight As Integer)
    Dim appLeft As Integer
    Dim appTop As Integer
    Dim appWidth As Integer
    Dim appHeight As Integer
    Application.WindowState = xlNormal
    appLeft = maxWidth / 4
    appTop = maxHeight / 4
    appWidth = maxWidth / 2
    appHeight = maxHeight / 2
    Application.Left = appLeft + ((appWidth - 500) / 2)
    Application.Top = appTop + ((appHeight - 400) / 2)

End Sub