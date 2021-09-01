Option Explicit 

' Primeiro acesse https://bitly.com, crie uma conta e gere um token de acesso genérico
' Então, acesse Tools (Ferramentas) -> References (Referências) e selecione "Microsoft WinHTTP Services, version 5.1"
' Insira o seu token na área abaixo

Sub EncurtarUmLink() 
    Dim Token, EndAPI, EndLongo, IniTexto, FimTexto As String 
    Token = "Insira_aqui_o_seu_token" 
    Dim HttpReq  As New WinHttpRequest 
    Dim response As String 
     
    EndLongo = ActiveCell.Value 
    EndAPI = "https://api-ssl.bitly.com/v3/shorten?access_token=" & Token & "&longUrl=" & EndLongo 
     
     
    On Error Resume Next 'Isso é para evitar URLs inválidos, ainda que não seja uma boa prática
     
    With HttpReq 
        .Open "GET", EndAPI, False 
        .Send 
    End With 
     
    response = HttpReq.ResponseText 
    HttpReq.WaitForResponse 
    IniTexto = InStr(response, "hash") 
    FimTexto = IniTexto + 15 
    resultado = Right(Mid(response, IniTexto, (FimTexto - IniTexto)), 7) 
    ActiveCell.Value = "http://bit.ly/" & resultado 
End Sub 
