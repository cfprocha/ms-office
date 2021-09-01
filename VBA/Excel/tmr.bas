Public Sub minhaMacro() 
    On Error Resume Next 'O uso de On Error Resume Next, não é uma boa prática para tratamento de erros
    MsgBox ("Tempo decorrido!") 'Aqui, onde está essa MsgBox, vocë insere o código que deseja ver sendo executado
    Application.OnTime EarliestTime:=Now + TimeValue("00:15:00"), Procedure:="minhaMacro" 'Definido para 15 minutos
End Sub 
