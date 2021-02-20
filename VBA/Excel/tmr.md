## Code

```
Public Sub minhaMacro() 
    On Error Resume Next 
    MsgBox ("Time elapsed!") ' Here you would insert your real code
    Application.OnTime EarliestTime:=Now + TimeValue("00:15:00"), Procedure:="minhaMacro" ' set to 15 minutes
End Sub 
```

## How to use

    1. Create a command button (cmdStart) for starting the counter, in the Main sheet
    2. Goto Excel --> Tools --> Macro --> Visual Basic Editor (Or Press Alt + F11)
    3. In VBE window, goto Insert --> Module
    4. Double Click the Module1 and paste the code shown above.