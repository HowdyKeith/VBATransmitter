'***************************************************************
' CHAT MODULE - Complete separated chat server implementation
' Save as: ChatModule.bas
'***************************************************************
Option Explicit

' --- Windows API Declarations (only if not in main module) ---
#If VBA7 Then
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
    Private Declare PtrSafe Function bind Lib "ws2_32.dll" (ByVal s As Long, addr As Any, ByVal namelen As Long) As Long
    Private Declare PtrSafe Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Private Declare PtrSafe Function accept Lib "ws2_32.dll" (ByVal s As Long, addr As Any, ByRef addrLen As Long) As Long
    Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function send Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Long
    Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
#Else
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
    Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, addr As Any, ByVal namelen As Long) As Long
    Private Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, addr As Any, ByRef addrLen As Long) As Long
    Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
    Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
#End If

' --- Constants (reference main module or declare locally) ---
Private Const AF_INET = 2
Private Const PF_INET = 2
Private Const SOCK_STREAM = 1
Private Const IPPROTO_TCP = 6
Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1
Private Const sockaddr_size = 16
Private Const FIONBIO = &H8004667E
Private Const WSAEWOULDBLOCK = 10035
Private Const SOL_SOCKET = &HFFFF
Private Const SO_REUSEADDR = &H4

' --- Type Definitions ---
Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

' --- Chat Module Variables ---
Private m_chatSocket As Long
Private m_chatRunning As Boolean
Private m_chatClients() As Long
Private m_chatCount As Long
Private m_chatBuffer(1 To 4096) As Byte
Private m_chatPort As Long
Private m_chatHistory As String

'***************************************************************
' PUBLIC INTERFACE FUNCTIONS (called from main server)
'***************************************************************

Public Function SetupChatServer(ByVal portNum As Long) As Boolean
    Dim addr As sockaddr
    Dim nonBlocking As Long
    Dim reuseAddr As Long
    
    On Error GoTo ErrorHandler
    
    m_chatPort = portNum
    m_chatSocket = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    If m_chatSocket = INVALID_SOCKET Then
        Debug.Print "Chat Module: Failed to create chat socket. Error: " & WSAGetLastError()
        SetupChatServer = False
        Exit Function
    End If
    
    ' Enable address reuse
    reuseAddr = 1
    setsockopt m_chatSocket, SOL_SOCKET, SO_REUSEADDR, reuseAddr, 4
    
    ' Set non-blocking mode
    nonBlocking = 1
    If ioctlsocket(m_chatSocket, FIONBIO, nonBlocking) = SOCKET_ERROR Then
        Debug.Print "Chat Module: Failed to set chat socket non-blocking. Error: " & WSAGetLastError()
        closesocket m_chatSocket
        SetupChatServer = False
        Exit Function
    End If
    
    ' Bind and listen
    addr.sin_family = AF_INET
    addr.sin_port = htons(m_chatPort)
    addr.sin_addr = inet_addr("127.0.0.1")
    
    If bind(m_chatSocket, addr, sockaddr_size) = SOCKET_ERROR Then
        Debug.Print "Chat Module: Failed to bind chat socket. Error: " & WSAGetLastError()
        closesocket m_chatSocket
        SetupChatServer = False
        Exit Function
    End If
    
    If listen(m_chatSocket, 5) = SOCKET_ERROR Then
        Debug.Print "Chat Module: Failed to listen on chat socket. Error: " & WSAGetLastError()
        closesocket m_chatSocket
        SetupChatServer = False
        Exit Function
    End If
    
    m_chatRunning = True
    m_chatCount = 0
    m_chatHistory = ""
    ReDim m_chatClients(0 To 0)
    m_chatClients(0) = INVALID_SOCKET
    
    Debug.Print "Chat Module: Chat server ready on port " & portNum
    SetupChatServer = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Chat Module: Error in SetupChatServer: " & Err.description
    SetupChatServer = False
End Function

Public Sub ProcessChatServer()
    On Error GoTo ErrorHandler
    
    If Not m_chatRunning Then Exit Sub
    
    AcceptChatClients
    ReceiveChatMessages
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chat Module: Error in ProcessChatServer: " & Err.description
End Sub

Private Function IsChatServerRunning() As Boolean
    IsChatServerRunning = m_chatRunning
End Function

Public Function GetChatServerStatus() As String
    Dim status As String
    status = "Chat Server Status:" & vbCrLf
    status = status & "• Port: " & m_chatPort & vbCrLf
    status = status & "• Running: " & IIf(m_chatRunning, "Yes", "No") & vbCrLf
    status = status & "• Connected Users: " & m_chatCount & vbCrLf
    status = status & "• Socket: " & m_chatSocket & vbCrLf
    GetChatServerStatus = status
End Function

Public Sub StopChatServer()
    Debug.Print "Chat Module: Stopping chat server"
    
    m_chatRunning = False
    
    On Error Resume Next
    
    ' Close all chat clients
    Dim i As Long
    For i = 0 To m_chatCount - 1
        If i <= UBound(m_chatClients) Then
            If m_chatClients(i) <> INVALID_SOCKET Then
                closesocket m_chatClients(i)
            End If
        End If
    Next i
    
    ' Close chat server socket
    If m_chatSocket <> INVALID_SOCKET Then
        closesocket m_chatSocket
        m_chatSocket = INVALID_SOCKET
    End If
    
    Debug.Print "Chat Module: Chat server stopped"
    On Error GoTo 0
End Sub

Public Sub HandleProxyRequest(ByVal request As String, ByVal clientSocket As Long)
    ' Handle requests forwarded from API Gateway
    On Error GoTo ErrorHandler
    
    Debug.Print "Chat Module: Handling proxy request from API Gateway"
    
    ' Parse the forwarded request and generate appropriate response
    If InStr(1, request, "GET /api/chat/status", vbTextCompare) > 0 Then
        SendChatApiStatus clientSocket
    ElseIf InStr(1, request, "GET /api/chat/history", vbTextCompare) > 0 Then
        SendChatHistory clientSocket
    ElseIf InStr(1, request, "POST /api/chat/message", vbTextCompare) > 0 Then
        HandleApiChatMessage request, clientSocket
    Else
        ' Default: return chat interface
        SendChatWebInterface clientSocket
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chat Module: Error handling proxy request: " & Err.description
    SendChatApiError clientSocket, 500, "Internal Server Error"
End Sub

'***************************************************************
' PRIVATE CHAT SERVER FUNCTIONS
'***************************************************************

Private Sub AcceptChatClients()
    Dim addr As sockaddr
    Dim addrLen As Long
    Dim newSocket As Long
    Dim errorCode As Long
    
    On Error GoTo ErrorHandler
    
    addrLen = sockaddr_size
    newSocket = accept(m_chatSocket, addr, addrLen)
    
    If newSocket = INVALID_SOCKET Then
        errorCode = WSAGetLastError()
        If errorCode <> WSAEWOULDBLOCK Then
            Debug.Print "Chat Module: Accept failed with error: " & errorCode
        End If
        Exit Sub
    End If
    
    ' Add new client
    If m_chatCount = 0 Then
        ReDim m_chatClients(0 To 0)
        m_chatClients(0) = newSocket
        m_chatCount = 1
    Else
        ReDim Preserve m_chatClients(0 To m_chatCount)
        m_chatClients(m_chatCount) = newSocket
        m_chatCount = m_chatCount + 1
    End If
    
    Debug.Print "Chat Module: Client connected: Socket=" & newSocket & ", Total clients: " & m_chatCount
    SendChatWelcome newSocket
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chat Module: Error in AcceptChatClients: " & Err.description
    If newSocket <> INVALID_SOCKET Then closesocket newSocket
End Sub

Private Sub ReceiveChatMessages()
    Dim i As Long
    Dim bytesRead As Long
    Dim msg As String
    Dim errorCode As Long
    
    On Error GoTo ErrorHandler
    
    If m_chatCount = 0 Then Exit Sub
    If Not IsArrayInitialized(m_chatClients) Then Exit Sub
    
    For i = 0 To m_chatCount - 1
        If i > UBound(m_chatClients) Then Exit For
        
        If m_chatClients(i) <> INVALID_SOCKET Then
            bytesRead = recv(m_chatClients(i), m_chatBuffer(1), 4096, 0)
            
            If bytesRead > 0 Then
                msg = Left$(StrConv(m_chatBuffer, vbUnicode), bytesRead)
                ProcessChatMessage msg, m_chatClients(i), i
            ElseIf bytesRead = 0 Then
                Debug.Print "Chat Module: Client disconnected: Socket=" & m_chatClients(i)
                closesocket m_chatClients(i)
                m_chatClients(i) = INVALID_SOCKET
            Else
                errorCode = WSAGetLastError()
                If errorCode <> WSAEWOULDBLOCK Then
                    closesocket m_chatClients(i)
                    m_chatClients(i) = INVALID_SOCKET
                End If
            End If
        End If
    Next i
    
    CleanupDisconnectedChatClients
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chat Module: Error in ReceiveChatMessages: " & Err.description
End Sub

Private Sub ProcessChatMessage(ByVal message As String, ByVal clientSocket As Long, ByVal clientIndex As Long)
    On Error GoTo ErrorHandler
    
    Debug.Print "Chat Module: Received message from socket " & clientSocket & ": " & Left(message, 100)
    
    ' Check if it's an HTTP request
    If InStr(1, message, "GET / HTTP/", vbTextCompare) > 0 Or _
       InStr(1, message, "GET /chat", vbTextCompare) > 0 Then
        SendChatWebInterface clientSocket
    ElseIf InStr(1, message, "POST /send HTTP/", vbTextCompare) > 0 Then
        HandleChatPost clientSocket, message
    ElseIf InStr(1, message, "GET /favicon.ico", vbTextCompare) > 0 Then
        SendFaviconResponse clientSocket
    ElseIf InStr(1, message, "GET ", vbTextCompare) > 0 And InStr(1, message, "HTTP/", vbTextCompare) > 0 Then
        SendHTTPRedirect clientSocket
    ElseIf Len(Trim(message)) > 0 And InStr(1, message, "HTTP/", vbTextCompare) = 0 Then
        ' Regular chat message (not HTTP)
        BroadcastChatMessage clientSocket, message, clientIndex
    Else
        ' Echo message for basic protocol
        Dim response As String
        response = "ECHO: " & Trim(message) & vbCrLf
        SendChatData clientSocket, response
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chat Module: Error processing chat message: " & Err.description
    On Error Resume Next
    closesocket clientSocket
End Sub

Private Sub BroadcastChatMessage(ByVal senderSocket As Long, ByVal message As String, ByVal senderIndex As Long)
    Dim i As Long
    Dim broadcastMsg As String
    Dim timestamp As String
    
    On Error Resume Next
    
    timestamp = format(Now, "hh:mm:ss")
    broadcastMsg = "User" & senderIndex & " [" & timestamp & "]: " & Trim(message) & vbCrLf
    
    ' Add to history
    m_chatHistory = m_chatHistory & broadcastMsg
    
    ' Keep history manageable
    If Len(m_chatHistory) > 10000 Then
        m_chatHistory = Right(m_chatHistory, 5000)
    End If
    
    ' Broadcast to all clients except sender
    For i = 0 To m_chatCount - 1
        If i <= UBound(m_chatClients) And i <> senderIndex Then
            If m_chatClients(i) <> INVALID_SOCKET Then
                SendChatData m_chatClients(i), broadcastMsg
            End If
        End If
    Next i
    
    Debug.Print "Chat Module: Broadcasted: " & Left(broadcastMsg, 50)
    
    On Error GoTo 0
End Sub

Private Sub SendChatWelcome(ByVal clientSocket As Long)
    Dim welcomeMsg As String
    welcomeMsg = "Welcome to VBA Chat Server!" & vbCrLf
    welcomeMsg = welcomeMsg & "Connected users: " & m_chatCount & vbCrLf
    welcomeMsg = welcomeMsg & "Type your messages and press Enter." & vbCrLf & vbCrLf
    
    ' Send recent history
    If Len(m_chatHistory) > 0 Then
        welcomeMsg = welcomeMsg & "Recent messages:" & vbCrLf & m_chatHistory
    End If
    
    SendChatData clientSocket, welcomeMsg
End Sub

Private Sub SendChatWebInterface(ByVal clientSocket As Long)
    Dim html As String
    Dim contentLength As Long
    Dim content As String
    
    ' Build the HTML content first
    content = "<!DOCTYPE html>" & vbCrLf
    content = content & "<html><head>" & vbCrLf
    content = content & "<meta charset='utf-8'>" & vbCrLf
    content = content & "<title>VBA Chat Server</title>" & vbCrLf
    content = content & "<style>" & vbCrLf
    content = content & "body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }" & vbCrLf
    content = content & "#messages { border: 1px solid #ccc; height: 400px; overflow-y: scroll; padding: 10px; background: white; margin-bottom: 10px; }" & vbCrLf
    content = content & "input[type='text'] { width: 300px; padding: 8px; border: 1px solid #ddd; }" & vbCrLf
    content = content & "input[type='submit'] { padding: 8px 15px; margin-left: 5px; background: #007acc; color: white; border: none; }" & vbCrLf
    content = content & ".status { background: #e7f3ff; padding: 10px; border-radius: 5px; margin-bottom: 10px; }" & vbCrLf
    content = content & "</style>" & vbCrLf
    content = content & "</head><body>" & vbCrLf
    content = content & "<h1>?? VBA Chat Server</h1>" & vbCrLf
    content = content & "<div class='status'>" & vbCrLf
    content = content & "<strong>Status:</strong> Active | <strong>Users:</strong> " & m_chatCount & " | <strong>Port:</strong> " & m_chatPort & vbCrLf
    content = content & "</div>" & vbCrLf
    content = content & "<div id='messages'>"
    
    ' Add chat history with proper HTML encoding
    If Len(m_chatHistory) > 0 Then
        content = content & Replace(Replace(m_chatHistory, vbCrLf, "<br>" & vbCrLf), "&", "&amp;")
    Else
        content = content & "<em>No messages yet. Start chatting!</em><br>" & vbCrLf
    End If
    
    content = content & "</div>" & vbCrLf
    content = content & "<form method='post' action='/send'>" & vbCrLf
    content = content & "<input type='text' name='message' placeholder='Type your message...' maxlength='500' required>" & vbCrLf
    content = content & "<input type='submit' value='Send Message'>" & vbCrLf
    content = content & "</form>" & vbCrLf
    content = content & "<p><small>Auto-refresh in 10 seconds | <a href='/'>Refresh Now</a></small></p>" & vbCrLf
    content = content & "<script>" & vbCrLf
    content = content & "setTimeout(function(){ location.reload(); }, 10000);" & vbCrLf
    content = content & "document.querySelector('input[name=\""message\""]').focus();" & vbCrLf
    content = content & "var messages = document.getElementById('messages');" & vbCrLf
    content = content & "messages.scrollTop = messages.scrollHeight;" & vbCrLf
    content = content & "</script>" & vbCrLf
    content = content & "</body></html>" & vbCrLf
    
    ' Calculate content length
    contentLength = Len(content)
    
    ' Build complete HTTP response
    html = "HTTP/1.1 200 OK" & vbCrLf
    html = html & "Content-Type: text/html; charset=utf-8" & vbCrLf
    html = html & "Content-Length: " & contentLength & vbCrLf
    html = html & "Cache-Control: no-cache" & vbCrLf
    html = html & "Connection: close" & vbCrLf
    html = html & "Server: VBA-Chat/1.0" & vbCrLf
    html = html & vbCrLf
    html = html & content
    
    SendChatData clientSocket, html
    closesocket clientSocket
End Sub

Private Sub HandleChatPost(ByVal clientSocket As Long, ByVal request As String)
    ' Extract message from POST data
    Dim messageStart As Long
    Dim messageEnd As Long
    Dim message As String
    Dim postDataStart As Long
    
    ' Find the start of POST data (after double CRLF)
    postDataStart = InStr(request, vbCrLf & vbCrLf)
    If postDataStart > 0 Then
        postDataStart = postDataStart + 4
        
        messageStart = InStr(postDataStart, request, "message=")
        If messageStart > 0 Then
            messageStart = messageStart + 8
            messageEnd = InStr(messageStart, request, "&")
            If messageEnd = 0 Then messageEnd = Len(request) + 1
            
            message = Mid(request, messageStart, messageEnd - messageStart)
            
            ' URL decode the message
            message = URLDecodeBasic(message)
            
            If Len(Trim(message)) > 0 Then
                BroadcastChatMessage clientSocket, message, -1
                Debug.Print "Chat Module: Web chat message: " & message
            End If
        End If
    End If
    
    ' Send redirect response
    Dim response As String
    response = "HTTP/1.1 302 Found" & vbCrLf
    response = response & "Location: /" & vbCrLf
    response = response & "Content-Length: 0" & vbCrLf
    response = response & "Connection: close" & vbCrLf
    response = response & vbCrLf
    
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

'***************************************************************
' API GATEWAY INTEGRATION FUNCTIONS
'***************************************************************

Private Sub SendChatApiStatus(ByVal clientSocket As Long)
    Dim response As String
    response = "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: application/json" & vbCrLf & "Connection: close" & vbCrLf & vbCrLf
    response = response & "{""chat"": {""running"": true, ""port"": " & m_chatPort & ", ""users"": " & m_chatCount & ", ""socket"": " & m_chatSocket & "}}"
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

Private Sub SendChatHistory(ByVal clientSocket As Long)
    Dim response As String
    Dim jsonHistory As String
    
    ' Convert history to JSON-safe format
    jsonHistory = Replace(Replace(m_chatHistory, """", "\"""), vbCrLf, "\\n")
    
    response = "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: application/json" & vbCrLf & "Connection: close" & vbCrLf & vbCrLf
    response = response & "{""history"": """ & jsonHistory & """, ""users"": " & m_chatCount & ", ""timestamp"": """ & Now() & """}"
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

Private Sub HandleApiChatMessage(ByVal request As String, ByVal clientSocket As Long)
    ' Handle POST message via API
    Dim message As String
    
    ' Extract message from request body (simplified)
    Dim postDataStart As Long
    postDataStart = InStr(request, vbCrLf & vbCrLf)
    If postDataStart > 0 Then
        message = Mid(request, postDataStart + 4)
        If Len(Trim(message)) > 0 Then
            BroadcastChatMessage -1, message, -1
        End If
    End If
    
    ' Send API response
    Dim response As String
    response = "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: application/json" & vbCrLf & "Connection: close" & vbCrLf & vbCrLf
    response = response & "{""status"": ""message_sent"", ""timestamp"": """ & Now() & """}"
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

Private Sub SendChatApiError(ByVal clientSocket As Long, ByVal errorCode As Long, ByVal errorMsg As String)
    Dim response As String
    response = "HTTP/1.1 " & errorCode & " " & errorMsg & vbCrLf
    response = response & "Content-Type: application/json" & vbCrLf
    response = response & "Connection: close" & vbCrLf & vbCrLf
    response = response & "{""error"": {""code"": " & errorCode & ", ""message"": """ & errorMsg & """}}"
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

'***************************************************************
' UTILITY FUNCTIONS
'***************************************************************

Private Function IsArrayInitialized(arr As Variant) As Boolean
    On Error GoTo NotInitialized
    Dim ub As Long
    ub = UBound(arr)
    IsArrayInitialized = True
    Exit Function
    
NotInitialized:
    IsArrayInitialized = False
End Function

Private Sub CleanupDisconnectedChatClients()
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim validClients() As Long
    Dim validCount As Long
    
    If m_chatCount = 0 Then Exit Sub
    If Not IsArrayInitialized(m_chatClients) Then Exit Sub
    
    ' Count valid clients
    validCount = 0
    For i = 0 To m_chatCount - 1
        If i <= UBound(m_chatClients) Then
            If m_chatClients(i) <> INVALID_SOCKET Then
                validCount = validCount + 1
            End If
        End If
    Next i
    
    ' Rebuild array with only valid clients
    If validCount > 0 Then
        ReDim validClients(0 To validCount - 1)
        j = 0
        For i = 0 To m_chatCount - 1
            If i <= UBound(m_chatClients) Then
                If m_chatClients(i) <> INVALID_SOCKET Then
                    validClients(j) = m_chatClients(i)
                    j = j + 1
                End If
            End If
        Next i
        m_chatClients = validClients
    Else
        ReDim m_chatClients(0 To 0)
        m_chatClients(0) = INVALID_SOCKET
    End If
    m_chatCount = validCount
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chat Module: Error in CleanupDisconnectedChatClients: " & Err.description
End Sub

Private Sub SendChatData(ByVal clientSocket As Long, ByVal data As String)
    On Error Resume Next
    Dim buffer() As Byte
    buffer = StrConv(data, vbFromUnicode)
    send clientSocket, buffer(0), UBound(buffer) + 1, 0
    On Error GoTo 0
End Sub

Private Sub SendFaviconResponse(ByVal clientSocket As Long)
    Dim response As String
    response = "HTTP/1.1 404 Not Found" & vbCrLf
    response = response & "Content-Length: 0" & vbCrLf
    response = response & "Connection: close" & vbCrLf & vbCrLf
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

Private Sub SendHTTPRedirect(ByVal clientSocket As Long)
    Dim response As String
    response = "HTTP/1.1 302 Found" & vbCrLf
    response = response & "Location: /" & vbCrLf
    response = response & "Connection: close" & vbCrLf & vbCrLf
    SendChatData clientSocket, response
    closesocket clientSocket
End Sub

Private Function URLDecodeBasic(ByVal encodedString As String) As String
    Dim result As String
    result = encodedString
    result = Replace(result, "+", " ")
    result = Replace(result, "%20", " ")
    result = Replace(result, "%21", "!")
    result = Replace(result, "%22", """")
    result = Replace(result, "%26", "&")
    result = Replace(result, "%3D", "=")
    result = Replace(result, "%3F", "?")
    URLDecodeBasic = result
End Function

'***************************************************************
' STANDALONE FUNCTIONS (for direct testing)
'***************************************************************

Public Sub StartStandaloneChatServer(Optional ByVal portNum As Long = 5000)
    Debug.Print "Chat Module: Starting standalone chat server on port " & portNum
    
    ' Initialize WSA (if needed)
    ' WSAStartup would be called from main module typically
    
    If SetupChatServer(portNum) Then
        Debug.Print "Chat Module: Standalone chat server started successfully"
        Debug.Print "Chat URL: http://localhost:" & portNum
        
        ' Simple monitoring loop for standalone mode
        Do While m_chatRunning
            DoEvents
            ProcessChatServer
            If GetAsyncKeyState(27) <> 0 Then Exit Do ' ESC key
            Sleep 50
        Loop
    Else
        Debug.Print "Chat Module: Failed to start standalone chat server"
    End If
End Sub

'***************************************************************
' COMPATIBILITY WRAPPERS FOR MAIN SERVER
'***************************************************************

' Returns True if chat server is running
Public Function GetChatRunning() As Boolean
    GetChatRunning = IsChatServerRunning()
End Function

' Starts the chat server on the specified port
Public Sub StartChatServer(ByVal portNum As Long)
    If Not IsChatServerRunning() Then
        SetupChatServer portNum
        Debug.Print "Chat Module: StartChatServer wrapper started server on port " & portNum
    Else
        Debug.Print "Chat Module: StartChatServer wrapper called but server already running"
    End If
End Sub


' Note: Sleep and GetAsyncKeyState would need to be declared if not in main module

