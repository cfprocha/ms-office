Option Explicit 
 
Sub EncurtarUmLink() 
    Dim Token, EndAPI, EndLongo, IniTexto, FimTexto As String 
    Token = "In_order_to_work_you_need_to_get_a_token_from_Bitly_and_insert_it_here" 
    Dim HttpReq  As New WinHttpRequest 
    Dim response As String 
     
    EndLongo = ActiveCell.Value 
    EndAPI = "https://api-ssl.bitly.com/v3/shorten?access_token=" & Token & "&longUrl=" & EndLongo 
     
     
    On Error Resume Next 'This is to avoid errors on invalid URLs
     
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
