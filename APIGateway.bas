Option Explicit

' ================================================================
' APIGateway.bas
' Version: 1.1 - Uses WinsockHelper
' ================================================================

Private m_gatewaySocket As LongPtr
Private m_gatewayRunning As Boolean
Private m_gatewayClients() As LongPtr
Private m_gatewayCount As Long
Private m_gatewayBuffer As String * 8192

Private Sub LogGateway(message As String)
    Debug.Print "[GATEWAY] " & format(Now, "hh:nn:ss") & " " & message
End Sub

Public Sub StartGateway(portNum As Long, httpPort As Long, iotPort As Long, trafficPort As Long, _
                        appPort As Long, chatPort As Long, matterPort As Long, mqttPort As Long)
    
    If m_gatewayRunning Then Exit Sub
    If Not WinsockInit() Then Exit Sub
    
    m_gatewaySocket = CreateTCPSocket()
    If m_gatewaySocket = INVALID_SOCKET Then Exit Sub
    
    SetNonBlocking m_gatewaySocket
    If Not BindSocket(m_gatewaySocket, portNum) Then
        LogWinsockError "Gateway bind failed"
        Exit Sub
    End If
    If Not StartListening(m_gatewaySocket) Then Exit Sub
    
    ReDim m_gatewayClients(0 To 49)
    m_gatewayCount = 0
    m_gatewayRunning = True
    LogGateway "Gateway started on port " & portNum
End Sub

Public Sub ProcessGateway()
    If Not m_gatewayRunning Then Exit Sub
    
    Dim clientSocket As LongPtr
    clientSocket = AcceptConnection(m_gatewaySocket)
    If clientSocket <> INVALID_SOCKET Then
        If m_gatewayCount < 49 Then
            m_gatewayClients(m_gatewayCount) = clientSocket
            m_gatewayCount = m_gatewayCount + 1
        Else
            closesocket clientSocket
        End If
    End If
    
    Dim i As Long, bytes As Long
    For i = 0 To m_gatewayCount - 1
        bytes = recv(m_gatewayClients(i), ByVal m_gatewayBuffer, 8192, 0)
        If bytes > 0 Then
            HandleGatewayRequest m_gatewayClients(i), Left$(m_gatewayBuffer, bytes)
        ElseIf bytes = 0 Or (bytes = -1 And WSAGetLastError <> WSAEWOULDBLOCK) Then
            closesocket m_gatewayClients(i)
            If i < m_gatewayCount - 1 Then m_gatewayClients(i) = m_gatewayClients(m_gatewayCount - 1)
            m_gatewayCount = m_gatewayCount - 1
            i = i - 1
        End If
    Next i
End Sub

Public Sub ShutdownGateway()
    If m_gatewayRunning Then
        Dim i As Long
        For i = 0 To m_gatewayCount - 1
            closesocket m_gatewayClients(i)
        Next i
        closesocket m_gatewaySocket
        m_gatewayRunning = False
        LogGateway "Gateway shutdown"
    End If
End Sub

Private Sub HandleGatewayRequest(clientSocket As LongPtr, request As String)
    ' Add your routing logic here
    LogGateway "Received request"
End Sub

