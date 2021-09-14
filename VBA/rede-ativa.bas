Private Const NETWORK_ALIVE_LAN = &H1 'conexão com placa de rede
Private Const NETWORK_ALIVE_WAN = &H2 'Conexão RAS
Private Const NETWORK_ALIVE_AOL = &H4 'AOL
Private Declare Function IsNetworkAlive Lib "Sensapi" _
(lpdwFlags As Long) As Long
Private Function IsNetConnectionAlive() As Boolean
  Dim lngAlive As Long
  IsNetConnectionAlive = IsNetworkAlive(lngAlive) = 1
End Function

Private Function IsNetConnectionLAN() As Boolean
  Dim lngLAN As Long
  If IsNetworkAlive(lngLAN) = 1 Then
    IsNetConnectionLAN = lngLAN = NETWORK_ALIVE_LAN 
  End If
End Function
  
Private Function IsNetConnectionRAS() As Boolean
  Dim lngRAS As Long
  If IsNetworkAlive(lngRAS) = 1 Then
    IsNetConnectionRAS = lngRAS = NETWORK_ALIVE_WAN
  End If
End Function
  
Private Function IsNetConnectionAOL() As Boolean
  Dim lngAOL As Long
  If IsNetworkAlive(tmp) = 1 Then
    IsNetConnectionAOL = lngAOL = NETWORK_ALIVE_AOL
  End If
End Function
  
Private Function GetNetConnectionType() As String
  Dim lngAlive As Long
    If IsNetworkAlive(lngAlive) = 1 Then
      Select Case lngAlive
        Case NETWORK_ALIVE_LAN:
          GetNetConnectionType = _
          "Mais de uma conexão LAN ativa."
        Case NETWORK_ALIVE_WAN:
          GetNetConnectionType = _
          "Mais de uma conexão RAS ativa."
        Case NETWORK_ALIVE_AOL:
          GetNetConnectionType = _
        "Há uma conexão via AOL."
        Case Else:
      End Select
    Else
      GetNetConnectionType = _
      "Posso obter outra conexão."
    End If
End Function
  
Sub IsConnection()
  MsgBox GetNetConnectionType
End Sub
