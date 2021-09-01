'Insira esse código no Módulo 1 
Public gintRow As Integer 
 
In This_Workbook: 
Option Explicit 
 
Private Sub Workbook_Open() 
    gintRow = Range("Last").Row 
End Sub 
 
' Esse código deve ser inserido na Sheet1 (Plan1): 
Option Explicit 
 
Private Sub Worksheet_Calculate() 
    If Range("Last").Row < gintRow Then 
        MsgBox "The row is deleted!" 
    ElseIf Range("Last").Row > gintRow Then 
        MsgBox "The row is added!" 
    End If 
    gintRow = Range("Last").Row 
End Sub 
