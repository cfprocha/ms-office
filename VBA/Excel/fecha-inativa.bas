' ----- Isso vai em "Essa pasta de trabalho"
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
  Call Timer
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  Call Timer
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Call Limpa
End Sub

' ----- Isso vai no módulo
Public vartimer As Variant
  Sub Timer()
    Call Limpa
    vartimer = Format(Now + TimeSerial(0, 30, 0), "hh:mm:ss")
    If vartimer = "" Then Exit Sub
    Application.OnTime TimeValue(vartimer), "Fecha"
  End Sub
    
Sub Fecha()
  With Application
    .EnableEvents = False
    ActiveWorkbook.Save
    .Quit
  End With
End Sub
    
Sub Limpa()
  On Error Resume Next
  Application.OnTime earliesttime:=vartimer, _
  procedure:="Fecha", schedule:=False
  On Error GoTo 0
End Sub
