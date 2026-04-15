Option Explicit

'***************************************************************
' TransmissionServer Module (Enhanced UDP with Heartbeat, Stats)
' Purpose: Unified server manager for UDP with advanced features
'***************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
    Private Declare PtrSafe Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare PtrSafe Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal Length As Long, ByVal flags As Long, ByRef from As SOCKADDR_IN, ByRef fromlen As Long) As Long
    Private Declare PtrSafe Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal Length As Long, ByVal flags As Long, ByRef to_addr As SOCKADDR_IN, ByVal tolen As Long) As Long
    Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
    Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Long
    Private Declare PtrSafe Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Long
    Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare PtrSafe Function inet_ntoa Lib "ws2_32.dll" (ByVal inaddr As Long) As LongPtr
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
    Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal Length As Long, ByVal flags As Long, ByRef from As SOCKADDR_IN, ByRef fromlen As Long) As Long
    Private Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal Length As Long, ByVal flags As Long, ByRef to_addr As SOCKADDR_IN, ByVal tolen As Long) As Long
    Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
    Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Long
    Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Long
    Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inaddr As Long) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
#End If

' --- Constants ---
Private Const AF_INET As Long = 2
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_UDP As Long = 17
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_REUSEADDR As Long = &H4
Private Const FIONBIO As Long = &H8004667E
Private Const MAX_BUFFER As Long = 65535
Private Const INVALID_SOCKET As Long = -1
Private Const SOCKET_ERROR As Long = -1
Private Const MAX_RECENT_MESSAGES As Long = 100
Private Const UDP_BUFFER_SIZE As Long = 65536
Private Const DEFAULT_HEARTBEAT As Long = 30000

' --- Types ---
Public Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Public Type UDPQueueItem
    msg As String
    remoteIP As String
    remotePort As Long
    msgType As Long
End Type

Public Type udpStats
    startTime As Long
    sentPackets As Long
    receivedPackets As Long
    lastHeartbeat As Long
End Type

' --- Unified UDP Message Types ---
Public Enum UDPMessageType
    UDP_DATA = 0            ' Normal data message
    UDP_PING = 1            ' Ping message
    UDP_PONG = 2            ' Pong reply
    UDP_BROADCAST = 3       ' Broadcast message
    UDP_MULTICAST = 4       ' Multicast message
    UDP_SECURE = 5          ' Encrypted message
    UDP_COMPRESSED = 6      ' Compressed message
    UDP_HEARTBEAT = 7       ' Heartbeat message
    UDP_COMMAND = 8         ' Command/control message
    UDP_RESPONSE = 9        ' Response to command
End Enum


' --- Module Variables ---
Private udpQueue As Collection
Private udpSocket As Long
Private UdpPort As Long
Private udpLoopEnabled As Boolean
Private udpNextRun As Date
Private packetSequence As Long
Private encryptionKey As String
Private compressionEnabled As Boolean
Private securityEnabled As Boolean
Private udpStatsData As udpStats
Private udpConnections As Collection
Private recentMessages As Collection
Private blockedIPs As Collection
Private allowedIPs As Collection
Private isWinsockInitialized As Boolean
Public heartbeatInterval As Long
Private udpReceiveSocket As Long
Private udpReceivePort As Long
Private udpReceiveEnabled As Boolean
' --- UDP Running Flag ---
'Private udpLoopEnabled As Boolean
' --- Initialize Queue ---
Private Sub InitUDPQueue()
    If udpQueue Is Nothing Then Set udpQueue = New Collection
End Sub

' --- Queue / Pop / Check ---
Public Sub UDPQueuePushMessage(ByVal msg As String, ByVal remoteIP As String, ByVal remotePort As Long, Optional ByVal msgType As Long = UDP_DATA)
    If udpQueue Is Nothing Then Set udpQueue = New Collection
    
    Dim newItem As cUDPQueueItem
    Set newItem = New cUDPQueueItem
    
    newItem.msg = msg
    newItem.remoteIP = remoteIP
    newItem.remotePort = remotePort
    newItem.msgType = msgType
    
    udpQueue.Add newItem
End Sub


Public Function UDPQueuePopMessage(ByRef ip As String, ByRef port As Long, ByRef msgType As UDPMessageType) As String
    InitUDPQueue
    If udpQueue.count = 0 Then
        UDPQueuePopMessage = ""
        ip = ""
        port = 0
        msgType = UDP_DATA
        Exit Function
    End If
    
    Dim item As cUDPQueueItem   ' <-- use the class, not the UDT
    Set item = udpQueue(1)
    
    UDPQueuePopMessage = item.msg
    ip = item.remoteIP
    port = item.remotePort
    msgType = item.msgType
    udpQueue.Remove 1
End Function


Public Function UDPMessageQueueIsEmpty() As Boolean
    InitUDPQueue
    UDPMessageQueueIsEmpty = (udpQueue.count = 0)
End Function

' --- Send / Process Queue ---
Public Sub ProcessUDPQueue()
    Dim ip As String, port As Long
    Dim msgType As UDPMessageType
    Dim msg As String
    
    ' Send heartbeat if due
    If GetTickCount() - udpStatsData.lastHeartbeat > heartbeatInterval Then
        UDPQueuePushMessage "HEARTBEAT", "255.255.255.255", UdpPort, UDP_HEARTBEAT
        udpStatsData.lastHeartbeat = GetTickCount()
    End If
    
    ' Process queue
    Do While Not UDPMessageQueueIsEmpty()
        msg = UDPQueuePopMessage(ip, port, msgType)
        If Len(msg) > 0 Then
            If compressionEnabled Then msg = SimpleCompress(msg)
            If securityEnabled Then msg = SimpleEncrypt(msg, encryptionKey)
            SendUDP msg, ip, port
            udpStatsData.sentPackets = udpStatsData.sentPackets + 1
        End If
    Loop
End Sub

' --- Send UDP ---
Public Sub SendUDP(ByVal msg As String, ByVal ip As String, ByVal port As Long)
    Dim addr As SOCKADDR_IN
    Dim buf() As Byte
    
    buf = StrConv(msg, vbFromUnicode)
    
    addr.sin_family = AF_INET
    addr.sin_port = htons(port)
    addr.sin_addr = inet_addr(ip)
    
    If udpSocket = 0 Then
        udpSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
        If udpSocket = INVALID_SOCKET Then Exit Sub
        Dim nonBlock As Long
        nonBlock = 1
        ioctlsocket udpSocket, FIONBIO, nonBlock
    End If
    
    sendto udpSocket, buf(0), UBound(buf) + 1, 0, addr, LenB(addr)
End Sub

' --- Compression ---
Private Function SimpleCompress(ByVal data As String) As String
    Dim result As String, i As Long, count As Long, currentChar As String
    result = ""
    i = 1
    Do While i <= Len(data)
        currentChar = Mid(data, i, 1)
        count = 1
        Do While i + count <= Len(data) And Mid(data, i + count, 1) = currentChar And count < 255
            count = count + 1
        Loop
        If count > 1 Then
            result = result & Chr(255) & Chr(count) & currentChar
            i = i + count
        Else
            result = result & currentChar
            i = i + 1
        End If
    Loop
    SimpleCompress = result
End Function

Private Function SimpleDecompress(ByVal data As String) As String
    Dim result As String, i As Long, count As Long, char As String
    result = ""
    i = 1
    Do While i <= Len(data)
        If Asc(Mid(data, i, 1)) = 255 And i + 2 <= Len(data) Then
            count = Asc(Mid(data, i + 1, 1))
            char = Mid(data, i + 2, 1)
            result = result & String(count, char)
            i = i + 3
        Else
            result = result & Mid(data, i, 1)
            i = i + 1
        End If
    Loop
    SimpleDecompress = result
End Function

' --- Encryption ---
Private Function SimpleEncrypt(ByVal data As String, ByVal key As String) As String
    Dim result As String, i As Long, keyPos As Long
    result = ""
    keyPos = 1
    For i = 1 To Len(data)
        result = result & Chr(Asc(Mid(data, i, 1)) Xor Asc(Mid(key, keyPos, 1)))
        keyPos = keyPos + 1
        If keyPos > Len(key) Then keyPos = 1
    Next i
    SimpleEncrypt = result
End Function

Private Function SimpleDecrypt(ByVal data As String, ByVal key As String) As String
    SimpleDecrypt = SimpleEncrypt(data, key)
End Function

' --- Winsock Init / Cleanup ---
Public Function InitializeWinsock() As Boolean
    Dim wsa As WSADATA
    If isWinsockInitialized Then
        InitializeWinsock = True
        Exit Function
    End If
    If WSAStartup(&H202, wsa) = 0 Then
        isWinsockInitialized = True
        Set udpConnections = New Collection
        Set recentMessages = New Collection
        Set blockedIPs = New Collection
        Set allowedIPs = New Collection
        udpStatsData.startTime = GetTickCount()
        packetSequence = 0
        encryptionKey = GenerateSecureKey(32)
        compressionEnabled = False
        securityEnabled = False
        heartbeatInterval = DEFAULT_HEARTBEAT
        InitializeWinsock = True
        Debug.Print "[TransmissionServer] Winsock initialized"
    Else
        Debug.Print "[TransmissionServer] Winsock initialization failed: " & WSAGetLastError
        InitializeWinsock = False
    End If
End Function

Private Sub CleanupWinsock()
    If isWinsockInitialized Then
        WSACleanup
        isWinsockInitialized = False
        Set udpConnections = Nothing
        Set recentMessages = Nothing
        Set blockedIPs = Nothing
        Set allowedIPs = Nothing
        Debug.Print "[TransmissionServer] Winsock cleaned up"
    End If
End Sub

' --- Generate Secure Key ---
Private Function GenerateSecureKey(ByVal keyLength As Long) As String
    Dim key As String
    Dim i As Long
    key = ""
    For i = 1 To keyLength
        key = key & Chr(Int(Rnd() * 256))
    Next i
    GenerateSecureKey = key
End Function
'***************************************************************
' UDP Receiver / Listener
'***************************************************************


' --- Initialize UDP Receiver on a given port ---
Public Function InitializeUDPReceiver(ByVal port As Long) As Boolean
    Dim addr As SOCKADDR_IN
    Dim result As Long
    
    If Not InitializeWinsock Then
        InitializeUDPReceiver = False
        Exit Function
    End If
    
    udpReceivePort = port
    udpReceiveEnabled = True
    
    ' Create UDP socket
    udpReceiveSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If udpReceiveSocket = INVALID_SOCKET Then
        DebuggingLog.DebugLog "[TransmissionServer] UDP Receiver socket creation failed: " & WSAGetLastError
        InitializeUDPReceiver = False
        Exit Function
    End If
    
    ' Allow address reuse
    Dim opt As Long
    opt = 1
    setsockopt udpReceiveSocket, SOL_SOCKET, SO_REUSEADDR, opt, LenB(opt)
    
    ' Bind to port
    addr.sin_family = AF_INET
    addr.sin_port = htons(port)
    addr.sin_addr = 0 ' 0 = INADDR_ANY
    result = bind(udpReceiveSocket, addr, LenB(addr))
    If result <> 0 Then
        DebuggingLog.DebugLog "[TransmissionServer] UDP Receiver bind failed: " & WSAGetLastError
        closesocket udpReceiveSocket
        InitializeUDPReceiver = False
        Exit Function
    End If
    
    ' Set non-blocking mode
    ioctlsocket udpReceiveSocket, FIONBIO, 1&
    
    DebuggingLog.DebugLog "[TransmissionServer] UDP Receiver initialized on port " & port
    InitializeUDPReceiver = True
End Function

' --- Shutdown UDP Receiver ---
Public Sub ShutdownUDPReceiver()
    If udpReceiveEnabled Then
        closesocket udpReceiveSocket
        udpReceiveEnabled = False
        DebuggingLog.DebugLog "[TransmissionServer] UDP Receiver shut down"
    End If
End Sub

' --- Poll / Process incoming UDP messages ---
Public Sub ProcessIncomingUDP()
    If Not udpReceiveEnabled Then Exit Sub
    
    Dim buf As String * UDP_BUFFER_SIZE
    Dim addr As SOCKADDR_IN
    Dim addrLen As Long
    Dim bytesReceived As Long
    Dim msg As String
    
    addrLen = LenB(addr)
    
    Do
        bytesReceived = recvfrom(udpReceiveSocket, ByVal buf, UDP_BUFFER_SIZE, 0, addr, addrLen)
        If bytesReceived > 0 Then
            msg = Left$(buf, bytesReceived)
            
            ' Convert IP to string
            Dim senderIP As String
            senderIP = LongToIP(addr.sin_addr)
            
            ' Enqueue the message
            UDPQueuePushMessage msg, senderIP, ntohs(addr.sin_port), 0 ' msgType = 0 for received
            
            ' Update stats
            udpStats.totalReceived = udpStats.totalReceived + 1
            udpStats.lastReceivedTime = GetTickCount()
        Else
            Exit Do ' No more messages in queue
        End If
    Loop
End Sub

' --- Starts the UDP server on a given port ---
Public Sub StartUDPServer(ByVal port As Long)
    If Not InitializeWinsock() Then Exit Sub
    
    UdpPort = port
    
    ' Create UDP socket
    udpSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If udpSocket = INVALID_SOCKET Then
        DebuggingLog.DebugLog "[UDP] Failed to create socket: " & WSAGetLastError()
        Exit Sub
    End If
    
    ' Enable non-blocking mode
    Dim arg As Long
    arg = 1
    ioctlsocket udpSocket, FIONBIO, arg
    
    ' Allow address reuse
    Dim optval As Long
    optval = 1
    setsockopt udpSocket, SOL_SOCKET, SO_REUSEADDR, optval, LenB(optval)
    
    ' Bind socket to local port
    Dim addr As SOCKADDR_IN
    addr.sin_family = AF_INET
    addr.sin_port = htons(port)
    addr.sin_addr = 0
    If bind(udpSocket, addr, LenB(addr)) = SOCKET_ERROR Then
        DebuggingLog.DebugLog "[UDP] Failed to bind socket: " & WSAGetLastError()
        closesocket udpSocket
        udpSocket = INVALID_SOCKET
        Exit Sub
    End If
    
    ' Initialize queue if needed
    If udpQueue Is Nothing Then Set udpQueue = New Collection
    
    ' Enable UDP loop
    udpLoopEnabled = True
    
    DebuggingLog.DebugLog "[UDP] Server started on port " & port
End Sub


' --- Stop UDP Server ---
Public Sub StopUDPServer()
    udpLoopEnabled = False
    If udpSocket <> INVALID_SOCKET Then
        closesocket udpSocket
        udpSocket = INVALID_SOCKET
    End If
    DebuggingLog.DebugLog "[TransmissionServer] UDP server stopped"
End Sub

' --- Helper: Convert Long IP to string ---
Private Function LongToIP(ByVal ipLong As Long) As String
    LongToIP = (ipLong And &HFF) & "." & ((ipLong \ 256) And &HFF) & "." & ((ipLong \ 65536) And &HFF) & "." & ((ipLong \ 16777216) And &HFF)
End Function

Public Function ReceiveUDPAdvanced(ByVal udpSock As Long, ByRef senderIP As String, ByRef senderPort As Long, Optional ByVal useDecompression As Boolean = False, Optional ByVal useDecryption As Boolean = False) As String
    Dim buffer() As Byte
    Dim recvLen As Long
    Dim fromaddr As SOCKADDR_IN
    Dim fromlen As Long
    Dim rawMsg As String

    ReDim buffer(0 To UDP_BUFFER_SIZE - 1)
    fromlen = Len(fromaddr)

    recvLen = recvfrom(udpSock, buffer(0), UDP_BUFFER_SIZE, 0, fromaddr, fromlen)
    If recvLen = SOCKET_ERROR Or recvLen = 0 Then
        ReceiveUDPAdvanced = ""
        senderIP = ""
        senderPort = 0
        Exit Function
    End If

    ' Convert byte array to string
    rawMsg = StrConv(buffer, vbUnicode)

    ' Trim to received length
    rawMsg = Left(rawMsg, recvLen)

    ' Apply optional decryption
    If useDecryption Then rawMsg = SimpleDecrypt(rawMsg, encryptionKey)

    ' Apply optional decompression
    If useDecompression Then rawMsg = SimpleDecompress(rawMsg)

    senderIP = inet_ntoa(fromaddr.sin_addr)
    senderPort = ntohs(fromaddr.sin_port)
    ReceiveUDPAdvanced = rawMsg
End Function
' --- Process UDP Server Tick ---

' --- UDP Tick (Process Queue + Incoming Packets) ---
Public Sub ProcessUDPServerTick()
    If udpSocket = INVALID_SOCKET Or Not udpLoopEnabled Then Exit Sub
    
    ' --- Outgoing queue ---
    Dim i As Long
    Dim item As cUDPQueueItem
    
    If Not udpQueue Is Nothing Then
        For i = udpQueue.count To 1 Step -1
            Set item = udpQueue(i)  ' <-- now valid, class object
            
            If SendUDPAdvanced(udpSocket, item.msg, item.remoteIP, item.remotePort, item.msgType, compressionEnabled, securityEnabled) <> SOCKET_ERROR Then
                Debug.Print "[UDP SENT] " & item.remoteIP & ":" & item.remotePort & " -> " & item.msg
            Else
                Debug.Print "[UDP SEND ERROR] " & item.remoteIP & ":" & item.remotePort
            End If
            
            udpQueue.Remove i
        Next i
    End If
    
    ' --- Incoming packets ---
    Dim buf() As Byte, fromaddr As SOCKADDR_IN, fromlen As Long
    ReDim buf(0 To UDP_BUFFER_SIZE - 1)
    fromlen = LenB(fromaddr)
    
    Dim ret As Long
    ret = recvfrom(udpSocket, buf(0), UDP_BUFFER_SIZE, 0, fromaddr, fromlen)
    
    Do While ret > 0
        Dim msg As String, ip As String, port As Long
        msg = StrConv(buf, vbUnicode)
        ip = inet_ntoa(fromaddr.sin_addr)
        port = ntohs(fromaddr.sin_port)
        
        UDPQueuePushMessage msg, ip, port, UDP_DATA
        
        ret = recvfrom(udpSocket, buf(0), UDP_BUFFER_SIZE, 0, fromaddr, fromlen)
    Loop
End Sub



' --- Returns True if UDP initialized and loop enabled ---
Public Function GetUDPRunning() As Boolean
    GetUDPRunning = (isWinsockInitialized And udpLoopEnabled)
End Function




' --- Optional: Get UDP Stats (JSON string for TrafficManager) ---
Public Function GetUDPStatsJSON() As String
    GetUDPStatsJSON = "{""UDP_QueueCount"":" & IIf(udpQueue Is Nothing, 0, udpQueue.count) & _
                      ",""UDP_Port"":" & UdpPort & "}"
End Function

Public Sub ProcessMQTTTick()
    ' Ping broker if keepalive interval passed
    PingMQTT
    
    ' Process outgoing MQTT queue
    Dim i As Long
    Dim topic As String, msg As String
    
    If mqttMessageLog Is Nothing Then Exit Sub
    
    ' Here you could integrate actual outgoing MQTT packets if you implement TCP
    ' For now, just simulate logging
    For i = 1 To mqttMessageLog.count
        Debug.Print "[MQTT LOG] " & mqttMessageLog.item(i)
    Next i
    
    ' Clear the log after processing to avoid duplicates
    Set mqttMessageLog = New Collection
End Sub

