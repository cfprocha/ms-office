Public Sub minhaMacro() 
    On Error Resume Next 
    MsgBox ("Time elapsed!") ' Here you would insert your real code
    Application.OnTime EarliestTime:=Now + TimeValue("00:15:00"), Procedure:="minhaMacro" ' set to 15 minutes
End Sub 
