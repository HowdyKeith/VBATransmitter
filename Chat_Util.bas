Public Sub TestChatClient()
    On Error GoTo ErrorHandler
    
    Dim clientSocket As Long
    Dim addr As sockaddr
    Dim buffer As String * BUFFER_SIZE
    Dim bytesReceived As Long
    Dim message As String
    Dim ret As Long
    
    ' Initialize Winsock
    Dim wsa As WSADataType
    ret = WSAStartup(&H202, wsa)
    If ret <> 0 Then
        Debug.Print "WSAStartup failed: " & ret
        LogAppLauncherStdErr "WSAStartup failed in TestChatClient: " & ret
        Exit Sub
    End If
    
    ' Create socket
    clientSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If clientSocket = INVALID_SOCKET Then
        Debug.Print "Socket creation failed: " & WSAGetLastError
        LogAppLauncherStdErr "Socket creation failed in TestChatClient: " & WSAGetLastError
        WSACleanup
        Exit Sub
    End If
    
    ' Set up address
    addr.sin_family = AF_INET
    addr.sin_port = htons(5000)
    addr.sin_addr = inet_addr("127.0.0.1")
    
    ' Connect to server
    Debug.Print "Connecting to 127.0.0.1:5000..."
    LogAppLauncherStdOut "TestChatClient connecting to 127.0.0.1:5000"
    ret = connect(clientSocket, addr, sockaddr_size)
    If ret = SOCKET_ERROR Then
        Debug.Print "Connect failed: " & WSAGetLastError
        LogAppLauncherStdErr "Connect failed in TestChatClient: " & WSAGetLastError
        closesocket clientSocket
        WSACleanup
        Exit Sub
    End If
    Debug.Print "Connected to Chat server"
    LogAppLauncherStdOut "TestChatClient connected to Chat server"
    
    ' Send message
    message = "Hello from TestChatClient!"
    ret = send(clientSocket, message, Len(message), 0)
    If ret = SOCKET_ERROR Then
        Debug.Print "Send failed: " & WSAGetLastError
        LogAppLauncherStdErr "Send failed in TestChatClient: " & WSAGetLastError
        closesocket clientSocket
        WSACleanup
        Exit Sub
    End If
    Debug.Print "Sent: " & message
    LogAppLauncherStdOut "TestChatClient sent: " & message
    
    ' Receive response (wait up to 5 seconds)
    Dim startTime As Double
    startTime = Timer
    Do While Timer - startTime < 5
        bytesReceived = recv(clientSocket, buffer, BUFFER_SIZE, 0)
        If bytesReceived > 0 Then
            Dim response As String
            response = Left$(buffer, bytesReceived)
            Debug.Print "Received: " & response
            LogAppLauncherStdOut "TestChatClient received: " & response
            Exit Do
        ElseIf bytesReceived = SOCKET_ERROR Then
            Debug.Print "Receive failed: " & WSAGetLastError
            LogAppLauncherStdErr "Receive failed in TestChatClient: " & WSAGetLastError
            Exit Do
        End If
        DoEvents
        Sleep 100
    Loop
    
    ' Clean up
    closesocket clientSocket
    WSACleanup
    Debug.Print "TestChatClient connection closed"
    LogAppLauncherStdOut "TestChatClient connection closed"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in TestChatClient: " & Err.description
    LogAppLauncherStdErr "Error in TestChatClient: " & Err.description
    If clientSocket <> INVALID_SOCKET Then closesocket clientSocket
    WSACleanup
End Sub
