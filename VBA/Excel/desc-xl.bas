Private Sub Workbook_Activate() 
    Call Desliga 
End Sub 
 
Private Sub Workbook_BeforeClose(Cancel As Boolean) 
    Call Liga 
End Sub 
 
Private Sub Workbook_Deactivate() 
    Call Liga 
End Sub 
 
Private Sub Workbook_Open() 
    Call Desliga 
End Sub 
 
Sub Desliga() 
    Dim cbars As CommandBar 
    Application.ScreenUpdating = False 
    For Each cbars In Application.CommandBars 
        cbars.Enabled = False 
    Next 
    With ActiveWindow 
        .DisplayHeadings = False 
        .Zoom = 80 
    End With 
    With Application 
        .DisplayFormulaBar = False 
        .Caption = "Aplicação do Carlos"  'Substitua esse título pelo que desejar
        .ScreenUpdating = True 
    End With 
End Sub 
Sub Liga() 
    Dim cbars As CommandBar 
    Application.ScreenUpdating = False 
    ActiveWindow.Zoom = 100 
    For Each cbars In Application.CommandBars 
        cbars.Enabled = True 
    Next 
    ActiveWindow.DisplayHeadings = True 
    With Application 
        .DisplayFormulaBar = True 
        .Caption = "Microsoft Excel" 
        .ScreenUpdating = True 
    End With 
End Sub 
