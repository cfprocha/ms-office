' Se desejar usar em uma célula, basta inserir a seguinte fórmula nela: =IsEmailValid("nome@dominio.com")
' Caso o endereço seja válido, a fórmula irá retornar VERDADEIRO

Sub email()
  Dim txtEmail As String
  txtEmail = InputBox("Digite o endereço", "Endereço do e-mail")
  Dim Situacao As String
  ' Verifica a sintaxe
  If IsEmailValid(txtEmail) Then
    Situacao = "Sintaxe de e-mail válida!"
  Else
    Situacao = "Sintaxe de e-mail inválida!"
  End If
  ' Apresenta o resultado
  MsgBox Situacao
End Sub

Function IsEmailValid(strEmail)
  Dim strArray As Variant
  Dim strItem As Variant
  Dim i As Long, c As String, blnIsItValid As Boolean
  blnIsItValid = True
  i = Len(strEmail) - Len(Application.Substitute(strEmail, "@", ""))
  If i <> 1 Then IsEmailValid = False: Exit Function
    ReDim strArray(1 To 2)
    strArray(1) = Left(strEmail, InStr(1, strEmail, "@", 1) - 1)
    strArray(2) = Application.Substitute(Right(strEmail, Len(strEmail) - Len(strArray(1))), "@", "")
    For Each strItem In strArray
      If Len(strItem) <= 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
      End If
      For i = 1 To Len(strItem)
        c = LCase(Mid(strItem, i, 1))
        If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
          blnIsItValid = False
          IsEmailValid = blnIsItValid
          Exit Function
        End If
      Next i
      If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
      End If
    Next strItem
    If InStr(strArray(2), ".") <= 0 Then
      blnIsItValid = False
      IsEmailValid = blnIsItValid
      Exit Function
    End If
    i = Len(strArray(2)) - InStrRev(strArray(2), ".")
    If i <> 2 And i <> 3 Then
      blnIsItValid = False
      IsEmailValid = blnIsItValid
      Exit Function
    End If
    If InStr(strEmail, "..") > 0 Then
      blnIsItValid = False
      IsEmailValid = blnIsItValid
      Exit Function
    End If
    IsEmailValid = blnIsItValid
End Function
