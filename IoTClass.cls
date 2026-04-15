Option Explicit

' ============================================================
' Class Module: IoTClass.cls
' Purpose: Event-driven IoT server using Windows Sockets API
' - Handles socket operations with proper error handling
' - Implements connection management for IoT devices
' - Integrates with bas
' - Date: August 24, 2025
' ============================================================

' --- Windows API Declarations ---
#If VBA7 Then
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
    Private Declare PtrSafe Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare PtrSafe Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Private Declare PtrSafe Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrLen As Long) As Long
    Private Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal Length As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal Length As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
    Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long
    Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal length As Long, ByVal flags As Long) As Long
    Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal length As Long, ByVal flags As Long) As Long
    Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' --- Constants ---
Private Const AF_INET = 2
Private Const SOCK_STREAM = 1
Private Const IPPROTO_TCP = 6
Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1
Private Const WSAEWOULDBLOCK = 10035
Private Const FIONBIO = &H8004667E
Private Const TRAFFIC_LOG_PATH = "C:\SmartTraffic\Logs\"
Private Const MAX_CLIENTS = 50

' --- Type Definitions ---
Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 127) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

' --- Class Variables ---
Private serverSocket As Long
Private clientSockets() As Long
Private clientCount As Long
Private isInitialized As Boolean
'Private IsServerRunning As Boolean
Private receiveBuf(0 To 8191) As Byte

' --- Class Initialize/Terminate ---
Private Sub Class_Initialize()
    serverSocket = INVALID_SOCKET
    clientCount = 0
    isInitialized = False
    IsServerRunning = False
    ReDim clientSockets(0 To MAX_CLIENTS - 1)
    Dim i As Long
    For i = 0 To MAX_CLIENTS - 1
        clientSockets(i) = INVALID_SOCKET
    Next i
    logMessage "IoTClass initialized"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    CloseAllConnections
    If isInitialized Then WSACleanup
    logMessage "IoTClass terminated"
End Sub

' --- Initialize Winsock ---
Private Function InitializeWinsock() As Boolean
    If isInitialized Then
        InitializeWinsock = True
        Exit Function
    End If
    
    Dim WSADATA As WSADataType
    Dim result As Long
    result = WSAStartup(&H202, WSADATA)
    
    If result = 0 Then
        isInitialized = True
        InitializeWinsock = True
        logMessage "Winsock initialized (Version: " & WSADATA.wVersion & ")"
    Else
        logMessage "WSAStartup failed with error: " & result
        InitializeWinsock = False
    End If
End Function

' --- Setup IoT Server ---
Public Function SetupIoTServer(ByVal portNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    If IsServerRunning Then
        logMessage "IoT server already running"
        SetupIoTServer = True
        Exit Function
    End If
    
    If Not InitializeWinsock() Then
        SetupIoTServer = False
        Exit Function
    End If
    
    ' Create server socket
    serverSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If serverSocket = INVALID_SOCKET Then
        logMessage "Failed to create server socket: " & WSAGetLastError()
        SetupIoTServer = False
        Exit Function
    End If
    
    ' Set non-blocking mode
    Dim nonBlocking As Long
    nonBlocking = 1
    If ioctlsocket(serverSocket, FIONBIO, nonBlocking) <> 0 Then
        logMessage "Failed to set non-blocking mode: " & WSAGetLastError()
        closesocket serverSocket
        serverSocket = INVALID_SOCKET
        SetupIoTServer = False
        Exit Function
    End If
    
    ' Bind to address
    Dim addr As SOCKADDR_IN
    addr.sin_family = AF_INET
    addr.sin_port = htons(portNum)
    addr.sin_addr = inet_addr("0.0.0.0")
    
    If bind(serverSocket, addr, LenB(addr)) <> 0 Then
        logMessage "Failed to bind server socket: " & WSAGetLastError()
        closesocket serverSocket
        serverSocket = INVALID_SOCKET
        SetupIoTServer = False
        Exit Function
    End If
    
    ' Start listening
    If listen(serverSocket, 5) <> 0 Then
        logMessage "Failed to listen on server socket: " & WSAGetLastError()
        closesocket serverSocket
        serverSocket = INVALID_SOCKET
        SetupIoTServer = False
        Exit Function
    End If
    
    IsServerRunning = True
    logMessage "IoT server started on port " & portNum
    SetupIoTServer = True
    Exit Function

ErrorHandler:
    logMessage "SetupIoTServer error: " & Err.description
    SetupIoTServer = False
End Function

' --- Wait for Connection/Activity ---
Public Function WaitForConnection(Optional ByVal timeoutMs As Long = 100) As Boolean
    If Not IsServerRunning Then
        WaitForConnection = False
        Exit Function
    End If
    
    Dim startTime As Long
    startTime = GetTickCount()
    
    Do While (GetTickCount() - startTime) < timeoutMs
        ' Accept new connections
        AcceptNewClients
        
        ' Process existing client messages
        ProcessClientMessages
        
        ' Check if we have active clients
        If HasActiveClients() Then
            WaitForConnection = True
            Exit Function
        End If
        
        DoEvents
        Sleep 10
    Loop
    
    WaitForConnection = HasActiveClients()
End Function

' --- Accept New Clients ---
Private Sub AcceptNewClients()
    If serverSocket = INVALID_SOCKET Then Exit Sub
    
    Dim addr As SOCKADDR_IN
    Dim addrLen As Long
    addrLen = LenB(addr)
    
    Do
        Dim newSocket As Long
        newSocket = accept(serverSocket, addr, addrLen)
        
        If newSocket = INVALID_SOCKET Then
            Dim lastError As Long
            lastError = WSAGetLastError()
            If lastError <> WSAEWOULDBLOCK Then
                logMessage "Accept error: " & lastError
            End If
            Exit Do
        End If
        
        ' Add client to array
        Dim i As Long
        For i = 0 To MAX_CLIENTS - 1
            If clientSockets(i) = INVALID_SOCKET Then
                clientSockets(i) = newSocket
                clientCount = clientCount + 1
                logMessage "Client connected: socket=" & newSocket
                Exit For
            End If
        Next i
        
        ' If no slot available, close the connection
        If i = MAX_CLIENTS Then
            logMessage "Max clients reached, closing connection"
            closesocket newSocket
        End If
    Loop
End Sub

' --- Process Client Messages ---
Private Sub ProcessClientMessages()
    Dim i As Long
    For i = 0 To MAX_CLIENTS - 1
        If clientSockets(i) <> INVALID_SOCKET Then
            ProcessClientSocket clientSockets(i), i
        End If
    Next i
End Sub

' --- Process Individual Client Socket ---
Private Sub ProcessClientSocket(ByVal clientSocket As Long, ByVal index As Long)
    Dim bytesReceived As Long
    bytesReceived = recv(clientSocket, receiveBuf(0), UBound(receiveBuf) + 1, 0)
    
    If bytesReceived > 0 Then
        ' Convert bytes to string
        Dim message As String
        message = StrConv(LeftB(receiveBuf, bytesReceived), vbUnicode)
        logMessage "Received from socket " & clientSocket & ": " & Left$(message, 100)
        
        ' Process the message (call back to IoTModule)
        ProcessIoTMessage message, clientSocket
        
    ElseIf bytesReceived = 0 Then
        ' Client disconnected gracefully
        logMessage "Client disconnected: socket=" & clientSocket
        closesocket clientSocket
        clientSockets(index) = INVALID_SOCKET
        clientCount = clientCount - 1
        
    Else
        ' Check for non-blocking error
        Dim lastError As Long
        lastError = WSAGetLastError()
        If lastError <> WSAEWOULDBLOCK Then
            logMessage "Recv error on socket " & clientSocket & ": " & lastError
            closesocket clientSocket
            clientSockets(index) = INVALID_SOCKET
            clientCount = clientCount - 1
        End If
    End If
End Sub

' --- Process IoT Message (calls back to IoTModule) ---
Private Sub ProcessIoTMessage(ByVal message As String, ByVal clientSocket As Long)
    ' This will call the public function in IoTModule
    ' Cannot use Declare statements to call VBA module functions
    ' Instead, we'll need IoTModule to call us, or use a different approach
    logMessage "Processing message: " & message & " from socket: " & clientSocket
    
    ' For now, send a simple ACK - the actual processing should be handled
    ' by having IoTModule call this class's methods directly
    SendToClient clientSocket, "ACK: " & Left$(message, 50) & vbCrLf
End Sub

' --- Send Data to Client ---
Public Sub SendToClient(ByVal clientSocket As Long, ByVal data As String)
    On Error GoTo ErrorHandler
    
    If clientSocket = INVALID_SOCKET Then Exit Sub
    
    Dim buffer() As Byte
    buffer = StrConv(data, vbFromUnicode)
    
    Dim result As Long
    result = send(clientSocket, buffer(0), UBound(buffer) + 1, 0)
    
    If result = SOCKET_ERROR Then
        logMessage "Send error to socket " & clientSocket & ": " & WSAGetLastError()
    Else
        logMessage "Sent to socket " & clientSocket & ": " & Left$(data, 100)
    End If
    Exit Sub

ErrorHandler:
    logMessage "SendToClient error: " & Err.description
End Sub

' --- Check if Has Active Clients ---
Private Function HasActiveClients() As Boolean
    HasActiveClients = (clientCount > 0)
End Function

' --- Close All Connections ---
Private Sub CloseAllConnections()
    On Error Resume Next
    
    If serverSocket <> INVALID_SOCKET Then
        closesocket serverSocket
        serverSocket = INVALID_SOCKET
    End If
    
    Dim i As Long
    For i = 0 To MAX_CLIENTS - 1
        If clientSockets(i) <> INVALID_SOCKET Then
            closesocket clientSockets(i)
            clientSockets(i) = INVALID_SOCKET
        End If
    Next i
    
    clientCount = 0
    IsServerRunning = False
    logMessage "All connections closed"
End Sub

' --- Get Client Count ---
Public Function GetClientCount() As Long
    GetClientCount = clientCount
End Function

' --- Get Server Status ---
Public Function IsServerRunning() As Boolean
    IsServerRunning = IsServerRunning
End Function

' --- Log Message ---
Private Sub logMessage(ByVal message As String)
    Debug.Print format(Now, "yyyy-mm-dd hh:mm:ss") & " - IoTClass: " & message
    
    ' Also write to file
    On Error Resume Next
    Dim fso As Object, logFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileName As String
    fileName = TRAFFIC_LOG_PATH & "iot_class_log_" & format(Date, "yyyy-mm-dd") & ".log"
    Set logFile = fso.OpenTextFile(fileName, 8, True)
    logFile.WriteLine format(Now, "yyyy-mm-dd hh:mm:ss") & " - IoTClass: " & message
    logFile.Close
    Set fso = Nothing
    On Error GoTo 0
End Sub
