Option Explicit

'***************************************************************
' HttpServer.bas
' Version: 1.3 - Dashboard route + improved status page
'***************************************************************

Private Type ClientConnection
    socket As LongPtr
    InUse As Boolean
    buffer As String
    IsWebSocket As Boolean
    LastActive As Double
End Type

Private Const MAX_CLIENTS As Long = 50
Private Const BUFFER_SIZE As Long = 32768

Private httpSocket As LongPtr
Private httpClients() As ClientConnection
Private httpRunning As Boolean
Private httpPort As Long

' ====================== UPDATED REQUEST HANDLING ======================
Private Sub ProcessCompleteRequest(ByVal idx As Long)
    Dim req As String: req = httpClients(idx).buffer
    If IsWebSocketUpgrade(req) Then
        PerformWebSocketHandshake idx, req
        Exit Sub
    End If
    
    Dim path As String: path = ParsePath(req)
    Dim response As String
    
    Select Case path
        Case "/", "/index", "/home"
            response = BuildHTTPResponse(200, "OK", GenerateHtmlPage(), "text/html")
        Case "/dashboard"
            response = BuildHTTPResponse(200, "OK", GenerateDashboardPage(), "text/html")
        Case "/api/status"
            response = BuildHTTPResponse(200, "OK", GetSystemStatusJSON(), "application/json")
 Case "/iot"
    response = BuildHTTPResponse(200, "OK", IoT_GetStatusHTML(), "text/html")
        
        Case Else
            response = BuildHTTPResponse(404, "Not Found", "<h1>404</h1>", "text/html")
            
            
            
    End Select
    
    SendHttpResponse idx, response
    CloseClient idx
End Sub

Private Function GenerateDashboardPage() As String
    GenerateDashboardPage = "<html><head><title>VBA Transmitter Dashboard</title></head>" & _
        "<body><h1>VBA Transmitter Dashboard</h1>" & _
        "<p><strong>HTTP:</strong> " & Config.httpPort & " | " & _
        "<strong>MQTT:</strong> " & Config.mqttPort & " | " & _
        "<strong>Gateway:</strong> " & Config.GatewayPort & "</p>" & _
        "<p><a href='/api/status'>JSON Status</a></p>" & _
        "<script>console.log('Dashboard loaded');</script></body></html>"
End Function


Private Sub LogHttp(message As String)
    Debug.Print "[HttpServer] " & format(Now, "hh:nn:ss") & " " & message
End Sub

' ====================== SERVER CONTROL ======================
Public Sub StartHttpServer(ByVal port As Long)
    If httpRunning Then Exit Sub
    If Not WinsockInit() Then Exit Sub
    
    httpPort = port
    httpSocket = CreateTCPSocket()
    If httpSocket = INVALID_SOCKET Then Exit Sub
    
    Dim opt As Long: opt = 1
    setsockopt httpSocket, SOL_SOCKET, SO_REUSEADDR, opt, 4
    
    SetNonBlocking httpSocket
    
    If Not BindSocket(httpSocket, port) Then
        LogWinsockError "Bind failed on port " & port
        closesocket httpSocket
        Exit Sub
    End If
    
    If Not StartListening(httpSocket) Then
        LogWinsockError "Listen failed"
        closesocket httpSocket
        Exit Sub
    End If
    
    ReDim httpClients(0 To MAX_CLIENTS - 1)
    httpRunning = True
    LogHttp "HTTP + WebSocket Server started on port " & port
End Sub

Public Sub ProcessHttpServer()
    If Not httpRunning Then Exit Sub
    
    ' Accept new clients
    Dim clientSock As LongPtr
    clientSock = AcceptConnection(httpSocket)
    If clientSock <> INVALID_SOCKET Then
        Dim i As Long
        For i = 0 To MAX_CLIENTS - 1
            If Not httpClients(i).InUse Then
                httpClients(i).socket = clientSock
                httpClients(i).InUse = True
                httpClients(i).buffer = ""
                httpClients(i).IsWebSocket = False
                httpClients(i).LastActive = Timer
                LogHttp "New client connected"
                Exit For
            End If
        Next i
        If i = MAX_CLIENTS Then closesocket clientSock
    End If
    
    ' Process clients
    Dim bytes As Long, i As Long
    For i = 0 To MAX_CLIENTS - 1
        If httpClients(i).InUse Then
            bytes = recv(httpClients(i).socket, ByVal httpClients(i).buffer, BUFFER_SIZE, 0)
            
            If bytes > 0 Then
                httpClients(i).buffer = Left$(httpClients(i).buffer, bytes)
                httpClients(i).LastActive = Timer
                
                If httpClients(i).IsWebSocket Then
                    ProcessWebSocketFrame i
                Else
                    ProcessCompleteRequest i
                End If
            ElseIf bytes = 0 Or (bytes = -1 And WSAGetLastError <> WSAEWOULDBLOCK) Then
                CloseClient i
            End If
        End If
    Next i
End Sub

Public Sub StopHttpServer()
    If Not httpRunning Then Exit Sub
    Dim i As Long
    For i = 0 To MAX_CLIENTS - 1
        If httpClients(i).InUse Then CloseClient i
    Next i
    closesocket httpSocket
    WinsockCleanup
    httpRunning = False
    LogHttp "HTTP Server stopped"
End Sub

' ====================== WEBSOCKET SUPPORT ======================
Public Sub BroadcastWebSocketMessage(message As String)
    Dim i As Long
    For i = 0 To MAX_CLIENTS - 1
        If httpClients(i).InUse And httpClients(i).IsWebSocket Then
            SendWebSocketMessage i, message
        End If
    Next i
End Sub

Public Sub NotifyNewMQTTMessage(topic As String, payload As String)
    BroadcastWebSocketMessage "MQTT|" & topic & "|" & payload
End Sub



Private Function IsWebSocketUpgrade(req As String) As Boolean
    IsWebSocketUpgrade = InStr(req, "Upgrade: websocket") > 0 And InStr(req, "Sec-WebSocket-Key") > 0
End Function

Private Sub PerformWebSocketHandshake(ByVal idx As Long, ByVal req As String)
    Dim key As String: key = GetHeaderValue(ParseHeaders(req), "Sec-WebSocket-Key")
    Dim acceptKey As String: acceptKey = ComputeWebSocketAccept(key)
    
    Dim hs As String
    hs = "HTTP/1.1 101 Switching Protocols" & vbCrLf & _
         "Upgrade: websocket" & vbCrLf & _
         "Connection: Upgrade" & vbCrLf & _
         "Sec-WebSocket-Accept: " & acceptKey & vbCrLf & vbCrLf
    
    Dim b() As Byte: b = StrConv(hs, vbFromUnicode)
    send httpClients(idx).socket, b(0), UBound(b) + 1, 0
    httpClients(idx).IsWebSocket = True
    LogHttp "WebSocket client connected"
End Sub

Private Sub ProcessWebSocketFrame(ByVal idx As Long)
    ' Basic implementation - expand as needed
    Dim data As String: data = httpClients(idx).buffer
    If Len(data) < 2 Then Exit Sub
    Dim opcode As Byte: opcode = AscB(LeftB(data, 1)) And &HF
    
    If opcode = 1 Then          ' Text
        BroadcastWebSocketMessage "Echo: " & Mid$(data, 7)
    ElseIf opcode = 8 Then      ' Close
        CloseClient idx
    End If
End Sub

Public Sub SendWebSocketMessage(ByVal idx As Long, message As String)
    If Not httpClients(idx).IsWebSocket Then Exit Sub
    Dim payload() As Byte: payload = StrConv(message, vbFromUnicode)
    Dim frame() As Byte
    ReDim frame(0 To 1 + UBound(payload))
    frame(0) = &H81
    frame(1) = UBound(payload) + 1
    Dim i As Long
    For i = 0 To UBound(payload)
        frame(2 + i) = payload(i)
    Next i
    send httpClients(idx).socket, frame(0), UBound(frame) + 1, 0
End Sub

Private Sub CloseClient(ByVal idx As Long)
    If httpClients(idx).InUse Then closesocket httpClients(idx).socket
    httpClients(idx).InUse = False
    httpClients(idx).IsWebSocket = False
End Sub

' ====================== HELPERS ======================
Private Function BuildHTTPResponse(code As Long, status As String, body As String, contentType As String) As String
    BuildHTTPResponse = "HTTP/1.1 " & code & " " & status & vbCrLf & _
                        "Content-Type: " & contentType & vbCrLf & _
                        "Content-Length: " & Len(body) & vbCrLf & _
                        "Connection: close" & vbCrLf & vbCrLf & body
End Function

Private Function GenerateHtmlPage() As String
    GenerateHtmlPage = "<h1>VBA Transmitter - HTTP + WebSocket Running</h1>"
End Function

' Add your ParsePath, ParseHeaders, GetHeaderValue, ComputeWebSocketAccept, etc. here if missing

