'Insira o código abaixo no módulo da pasta de trabalho (Workbook)
Private Sub Workbook_Open()
  Call Tempo
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Call Limpa
End Sub

'Insira o código abaixo em um módulo comum
Public vartimer As Variant
Const TimeOut = 5 'em minutos
Sub Salva()
  ActiveWorkbook.Save
  Call Tempo
End Sub

Sub Tempo()
  vartimer = Format(Now + TimeSerial(0, TimeOut, 0), "hh:mm:ss")
  If vartimer = "" Then Exit Sub
    Application.OnTime TimeValue(vartimer), "Salva"
End Sub
    
Sub Limpa()
  On Error Resume Next
  Application.OnTime earliesttime:=vartimer, _
  procedure:="Salva", schedule:=False
  On Error GoTo 0
End Sub
