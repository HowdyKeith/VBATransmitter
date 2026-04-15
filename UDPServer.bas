Option Explicit

' ================================================================
' UDPServer.bas
' Version: 1.0 - Lightweight UDP server for IoT discovery
' ================================================================

Private udpSocket As LongPtr
Private udpRunning As Boolean
Private udpBuffer As String * 2048

Public Sub StartUDPServer(ByVal port As Long)
    If udpRunning Then Exit Sub
    If Not WinsockInit() Then Exit Sub
    
    udpSocket = CreateUDPSocket()   ' Uses WinsockHelper
    SetNonBlocking udpSocket
    
    If Not BindSocket(udpSocket, port) Then
        LogWinsockError "UDP Bind failed on port " & port
        Exit Sub
    End If
    
    udpRunning = True
    DebuggingLog.DebugLog "UDP Server started on port " & port, "INFO"
End Sub

Public Sub ProcessUDPServer()
    If Not udpRunning Then Exit Sub
    
    Dim bytes As Long
    bytes = recv(udpSocket, ByVal udpBuffer, 2048, 0)
    
    If bytes > 0 Then
        Dim data As String: data = Left$(udpBuffer, bytes)
        DebuggingLog.DebugLog "UDP Received: " & data, "INFO"
        
        ' Example echo / discovery response
        Dim response As String: response = "VBA-Transmitter-ACK|" & data
        Dim respBytes() As Byte: respBytes = StrConv(response, vbFromUnicode)
        send udpSocket, respBytes(0), UBound(respBytes) + 1, 0
    End If
End Sub

Public Sub StopUDPServer()
    If udpRunning Then
        closesocket udpSocket
        udpRunning = False
    End If
End Sub
