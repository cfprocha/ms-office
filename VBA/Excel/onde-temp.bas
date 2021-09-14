Declare Function GetTempPath _
  Lib "kernel32" Alias "GetTempPathA" _
  (ByVal nBufferLength As Long, _
  ByVal lpBuffer As String) As Long
  
Public Function fncGetTempPath() As String
  Dim PathLen As Long
  Dim WinTempDir As String
  Dim BufferLength As Long
  BufferLength = 260
  WinTempDir = Space(BufferLength)
  PathLen = GetTempPath(BufferLength, WinTempDir)
  If Not PathLen = 0 Then
    fncGetTempPath = Left(WinTempDir, PathLen)
  Else
    fncGetTempPath = CurDir()
  End If
End Function
  
Sub Test()
  MsgBox fncGetTempPath
End Sub
