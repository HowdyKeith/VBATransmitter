Option Explicit

'***************************************************************
' Advanced UDP Module with Enhanced Features
' Purpose: Enterprise-grade UDP functionality with advanced monitoring,
'          security, compression, encryption, and protocol handling
'***************************************************************


#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function inet_ntoa Lib "ws2_32.dll" (ByVal addr As Long) As LongPtr
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
        ByVal wVersionRequested As Long, _
        lpWSAData As WSADATA) As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function socket Lib "ws2_32.dll" ( _
        ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As LongPtr
    Private Declare PtrSafe Function bind Lib "ws2_32.dll" ( _
        ByVal s As LongPtr, ByRef addr As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare PtrSafe Function sendto Lib "ws2_32.dll" ( _
        ByVal s As LongPtr, ByVal buf As String, ByVal buflen As Long, _
        ByVal flags As Long, ByRef addrto As SOCKADDR_IN, ByVal tolen As Long) As Long
    Private Declare PtrSafe Function recvfrom Lib "ws2_32.dll" ( _
        ByVal s As LongPtr, ByVal buf As String, ByVal buflen As Long, _
        ByVal flags As Long, ByRef from As SOCKADDR_IN, ByRef fromlen As Long) As Long
    Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As LongPtr) As Long
    Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
    Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal level As Long, _
        ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal addr As Long) As Long
    Private Declare Function WSAStartup Lib "ws2_32.dll" ( _
        ByVal wVersionRequested As Long, lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function socket Lib "ws2_32.dll" ( _
        ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As Long
    Private Declare Function bind Lib "ws2_32.dll" ( _
        ByVal s As Long, ByRef addr As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare Function sendto Lib "ws2_32.dll" ( _
        ByVal s As Long, ByVal buf As String, ByVal buflen As Long, _
        ByVal flags As Long, ByRef addrto As SOCKADDR_IN, ByVal tolen As Long) As Long
    Private Declare Function recvfrom Lib "ws2_32.dll" ( _
        ByVal s As Long, ByVal buf As String, ByVal buflen As Long, _
        ByVal flags As Long, ByRef from As SOCKADDR_IN, ByRef fromlen As Long) As Long
    Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
    Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, _
        ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
#End If

' --- Types ---
Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 127) As Byte
    iMaxSockets As Long
    iMaxUdpDg As Long
    lpVendorInfo As LongPtr
End Type

Public Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Public Type ip_mreq
    imr_multiaddr As Long
    imr_interface As Long
End Type

Public Type UDPPacketHeader
    packetID As Long
    timestamp As Long
    packetType As Byte
    flags As Byte
    dataLength As Integer
    checksum As Long
End Type

Public Type UDPConnectionInfo
    remoteIP As String
    remotePort As Long
    messageCount As Long
    bytesSent As Long
    bytesReceived As Long
    PacketLoss As Double
    AverageLatency As Double
    securityLevel As Long
    IsBlocked As Boolean
    lastActivity As Long
End Type

Public Type udpStats
    TotalMessagesSent As Long
    TotalMessagesReceived As Long
    TotalBytesSent As Long
    TotalBytesReceived As Long
    ErrorCount As Long
    PacketsDropped As Long
    EncryptedPackets As Long
    CompressionRatio As Double
    startTime As Long
End Type

' --- Constants ---
Public Const AF_INET As Long = 2
Public Const SOCK_DGRAM As Long = 2
Public Const IPPROTO_UDP As Long = 17
Public Const WSA_VERSION As Long = &H202
Public Const SOCKET_ERROR As Long = -1
Public Const INVALID_SOCKET As Long = -1
Public Const SOL_SOCKET As Long = &HFFFF
Public Const SO_REUSEADDR As Long = &H4
Public Const SO_BROADCAST As Long = &H20
Public Const SO_RCVBUF As Long = &H1002
Public Const SO_SNDBUF As Long = &H1001
Public Const FIONBIO As Long = &H8004667E
Public Const IPPROTO_IP As Long = 0
Public Const IP_ADD_MEMBERSHIP As Long = 12
Public Const WSAEWOULDBLOCK As Long = 10035

' --- Module-level variables ---
Private isWinsockInitialized As Boolean
Private udpSockets As Collection
Private udpConnections As Collection
Private recentMessages As Collection
Private blockedIPs As Collection
Private allowedIPs As Collection
Private messageQueues As Collection
Private udpStats As udpStats
Private packetSequence As Long
Private securityEnabled As Boolean
Private compressionEnabled As Boolean
Private heartbeatInterval As Long
Private encryptionKey As String
Private Const MAX_RECENT_MESSAGES As Long = 200
Private Const MAX_PACKET_SIZE As Long = 65535
Private Const DEFAULT_HEARTBEAT As Long = 5000

' --- Helper functions ---
Private Function StringFromPtrA(ptr As LongPtr) As String
    Dim s As String
    s = Space$(255)
    CopyMemory ByVal StrPtr(s), ByVal ptr, 255
    StringFromPtrA = s
End Function


' --- Enhanced Initialization ---
Public Function InitializeUDP() As Boolean
    Dim WSADATA As WSADataType
    If isWinsockInitialized Then
        InitializeUDP = True
        Exit Function
    End If
    
    If WSAStartup(&H202, WSADATA) = 0 Then
        isWinsockInitialized = True
        
        ' Initialize collections and settings
        udpStats.startTime = GetTickCount()
        Set udpConnections = New Collection
        Set recentMessages = New Collection
        Set blockedIPs = New Collection
        Set allowedIPs = New Collection
        Set udpSockets = New Collection
        Set messageQueues = New Collection
        
        ' Default settings
        compressionEnabled = False
        securityEnabled = False
        heartbeatInterval = DEFAULT_HEARTBEAT
        maxPacketSize = MAX_PACKET_SIZE
        packetSequence = 0
        
        ' Generate encryption key
        encryptionKey = GenerateSecureKey(32)
        
        InitializeUDP = True
        DebuggingLog.DebugLog "Advanced UDP Module initialized"
    Else
        InitializeUDP = False
        DebuggingLog.DebugLog "UDP initialization failed: " & WSAGetLastError
    End If
End Function

' --- Security Functions ---
Private Function GenerateSecureKey(ByVal keyLength As Long) As String
    Dim hProv As Long
    Dim keyBuffer As String
    keyBuffer = Space$(keyLength)
    
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) Then
        If CryptGenRandom(hProv, keyLength, keyBuffer) Then
            GenerateSecureKey = keyBuffer
        End If
        CryptReleaseContext hProv, 0
    End If
    
    If Len(GenerateSecureKey) = 0 Then
        ' Fallback to time-based key
        GenerateSecureKey = String(keyLength, Chr((Timer * 1000) Mod 255))
    End If
End Function

Private Function SimpleEncrypt(ByVal data As String, ByVal key As String) As String
    Dim result As String
    Dim i As Long
    Dim keyPos As Long
    
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
    SimpleDecrypt = SimpleEncrypt(data, key) ' XOR is reversible
End Function

' --- Compression Functions ---
Public Function SimpleCompress(ByVal data As String) As String
    ' Basic run-length encoding
    Dim result As String
    Dim i As Long
    Dim currentChar As String
    Dim count As Long
    
    If Len(data) = 0 Then
        SimpleCompress = data
        Exit Function
    End If
    
    currentChar = Left(data, 1)
    count = 1
    result = ""
    
    For i = 2 To Len(data)
        If Mid(data, i, 1) = currentChar And count < 255 Then
            count = count + 1
        Else
            If count > 3 Then
                result = result & Chr(255) & Chr(count) & currentChar
            Else
                result = result & String(count, currentChar)
            End If
            currentChar = Mid(data, i, 1)
            count = 1
        End If
    Next i
    
    ' Handle last sequence
    If count > 3 Then
        result = result & Chr(255) & Chr(count) & currentChar
    Else
        result = result & String(count, currentChar)
    End If
    
    SimpleCompress = result
End Function

Private Function SimpleDecompress(ByVal data As String) As String
    Dim result As String
    Dim i As Long
    Dim count As Long
    Dim char As String
    
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

' --- Packet Management ---
Private Function CreatePacketHeader(ByVal msgType As UDPMessageType, ByVal dataLength As Long) As String
    Dim header As UDPPacketHeader
    header.packetID = packetSequence
    header.timestamp = GetTickCount()
    header.packetType = msgType
    header.flags = 0
    header.dataLength = dataLength
    header.checksum = CalculateChecksum(CStr(header.packetID) & CStr(header.timestamp) & CStr(dataLength))
    
    packetSequence = packetSequence + 1
    
    CreatePacketHeader = PacketHeaderToString(header)
End Function

Private Function PacketHeaderToString(ByRef header As UDPPacketHeader) As String
    Dim result As String
    result = String(16, Chr(0)) ' 16-byte header
    
    ' Pack header data (simplified binary format)
    Mid(result, 1, 4) = Chr(header.packetID And &HFF) & Chr((header.packetID \ 256) And &HFF) & Chr((header.packetID \ 65536) And &HFF) & Chr((header.packetID \ 16777216) And &HFF)
    Mid(result, 5, 4) = Chr(header.timestamp And &HFF) & Chr((header.timestamp \ 256) And &HFF) & Chr((header.timestamp \ 65536) And &HFF) & Chr((header.timestamp \ 16777216) And &HFF)
    Mid(result, 9, 1) = Chr(header.packetType)
    Mid(result, 10, 1) = Chr(header.flags)
    Mid(result, 11, 2) = Chr(header.dataLength And &HFF) & Chr((header.dataLength \ 256) And &HFF)
    Mid(result, 13, 4) = Chr(header.checksum And &HFF) & Chr((header.checksum \ 256) And &HFF) & Chr((header.checksum \ 65536) And &HFF) & Chr((header.checksum \ 16777216) And &HFF)
    
    PacketHeaderToString = result
End Function

Private Function CalculateChecksum(ByVal data As String) As Long
    Dim i As Long
    Dim checksum As Long
    
    For i = 1 To Len(data)
        checksum = checksum + Asc(Mid(data, i, 1))
    Next i
    
    CalculateChecksum = checksum Mod 65536
End Function

' --- Enhanced Send Functions ---
Public Function SendUDPAdvanced( _
    ByVal udpSocket As Long, _
    ByVal message As String, _
    ByVal destIP As String, _
    ByVal destPort As Long, _
    Optional ByVal msgType As UDPMessageType = UDP_DATA, _
    Optional ByVal useCompression As Boolean = False, _
    Optional ByVal useEncryption As Boolean = False) As Long

    Dim packet As String
    Dim finalMsg As String
    
    ' --- Apply compression if requested ---
    If useCompression Then
        packet = SimpleCompress(message)
    Else
        packet = message
    End If
    
    ' --- Apply encryption if requested ---
    If useEncryption Then
        finalMsg = SimpleEncrypt(packet, encryptionKey)
    Else
        finalMsg = packet
    End If
    
    ' --- Send UDP packet ---
    Dim destAddr As SOCKADDR_IN
    destAddr.sin_family = AF_INET
    destAddr.sin_port = htons(destPort)
    destAddr.sin_addr = inet_addr(destIP)
    
    SendUDPAdvanced = sendto(udpSocket, ByVal finalMsg, Len(finalMsg), 0, destAddr, Len(destAddr))
    
    ' --- Optional logging ---
    Debug.Print "[UDP] Sent message to " & destIP & ":" & destPort & " (Type=" & msgType & ")"

End Function


' --- Rate Limiting ---
Private Function CheckRateLimit(ByVal ip As String) As Boolean
    ' Simple rate limiting: max 100 messages per minute per IP
    Dim key As String
    Dim currentTime As Long
    Dim messageCount As Long
    
    key = "rate_" & ip
    currentTime = GetTickCount()
    
    On Error Resume Next
    messageCount = CLng(recentMessages.item(key))
    If Err.Number <> 0 Then messageCount = 0
    Err.Clear
    
    If messageCount > 100 Then
        CheckRateLimit = False
        DebuggingLog.DebugLog "Rate limit exceeded for IP: " & ip
    Else
        CheckRateLimit = True
        recentMessages.Remove key
        recentMessages.Add messageCount + 1, key
    End If
End Function

' --- Security Management ---
Public Sub BlockIP(ByVal ip As String)
    On Error Resume Next
    blockedIPs.Add ip, ip
    DebuggingLog.DebugLog "IP blocked: " & ip
End Sub

Public Sub UnblockIP(ByVal ip As String)
    On Error Resume Next
    blockedIPs.Remove ip
    DebuggingLog.DebugLog "IP unblocked: " & ip
End Sub

Public Sub AddAllowedIP(ByVal ip As String)
    On Error Resume Next
    allowedIPs.Add ip, ip
    DebuggingLog.DebugLog "IP whitelisted: " & ip
End Sub

Private Function IsIPBlocked(ByVal ip As String) As Boolean
    On Error Resume Next
    Dim temp As String
    temp = blockedIPs.item(ip)
    IsIPBlocked = (Err.Number = 0)
    Err.Clear
End Function

Private Function IsIPAllowed(ByVal ip As String) As Boolean
    If allowedIPs.count = 0 Then
        IsIPAllowed = True ' No whitelist = allow all
    Else
        On Error Resume Next
        Dim temp As String
        temp = allowedIPs.item(ip)
        IsIPAllowed = (Err.Number = 0)
        Err.Clear
    End If
End Function

' --- Enhanced Connection Tracking ---
Private Sub UpdateConnectionInfo(ByVal ip As String, ByVal port As Long, ByVal bytesSent As Long, ByVal bytesReceived As Long)
    On Error Resume Next
    Dim key As String
    key = ip & ":" & port
    
    Dim connInfo As UDPConnectionInfo
    If udpConnections.count > 0 Then
        connInfo = udpConnections.item(key)
    End If
    
    If Err.Number <> 0 Then
        ' New connection
        connInfo.remoteIP = ip
        connInfo.remotePort = port
        connInfo.messageCount = 0
        connInfo.bytesSent = 0
        connInfo.bytesReceived = 0
        connInfo.PacketLoss = 0
        connInfo.AverageLatency = 0
        connInfo.securityLevel = 0
        connInfo.IsBlocked = IsIPBlocked(ip)
        Err.Clear
    End If
    
    connInfo.lastActivity = GetTickCount()
    connInfo.messageCount = connInfo.messageCount + 1
    connInfo.bytesSent = connInfo.bytesSent + bytesSent
    connInfo.bytesReceived = connInfo.bytesReceived + bytesReceived
    
    udpConnections.Remove key
    udpConnections.Add connInfo, key
End Sub

' --- Enhanced Message Logging ---
Private Sub LogRecentMessage(ByVal direction As String, ByVal message As String, ByVal ip As String, ByVal port As Long, ByVal bytes As Long, Optional ByVal msgType As UDPMessageType = UDP_DATA)
    Dim logEntry As String
    Dim typeStr As String
    
    Select Case msgType
        Case UDP_PING: typeStr = "PING"
        Case UDP_PONG: typeStr = "PONG"
        Case UDP_BROADCAST: typeStr = "BCAST"
        Case UDP_MULTICAST: typeStr = "MCAST"
        Case UDP_SECURE: typeStr = "SECURE"
        Case UDP_COMPRESSED: typeStr = "COMP"
        Case UDP_HEARTBEAT: typeStr = "HEART"
        Case UDP_COMMAND: typeStr = "CMD"
        Case UDP_RESPONSE: typeStr = "RESP"
        Case Else: typeStr = "DATA"
    End Select
    
    logEntry = format(Now, "hh:mm:ss") & " [" & direction & "][" & typeStr & "] " & ip & ":" & port & " (" & bytes & "b) " & Left(message, 40)
    
    recentMessages.Add logEntry
    
    ' Keep only recent messages
    While recentMessages.count > MAX_RECENT_MESSAGES
        recentMessages.Remove 1
    Wend
End Sub

' --- Advanced Socket Management ---
Public Function CreateAdvancedUDPSocket(Optional ByVal bufferSize As Long = 65536) As Long
    Dim udpSocket As Long
    
    If Not InitializeUDP Then
        CreateAdvancedUDPSocket = INVALID_SOCKET
        Exit Function
    End If
    
    udpSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    
    If udpSocket <> INVALID_SOCKET Then
        ' Set socket options for better performance
        Dim optval As Long
        
        ' Increase buffer sizes
        optval = bufferSize
        setsockopt udpSocket, SOL_SOCKET, SO_RCVBUF, optval, 4
        setsockopt udpSocket, SOL_SOCKET, SO_SNDBUF, optval, 4
        
        ' Enable address reuse
        optval = 1
        setsockopt udpSocket, SOL_SOCKET, SO_REUSEADDR, optval, 4
        
        ' Add to socket collection
        udpSockets.Add udpSocket, CStr(udpSocket)
        
        DebuggingLog.DebugLog "Advanced UDP socket created with buffer size: " & bufferSize
    End If
    
    CreateAdvancedUDPSocket = udpSocket
End Function

' --- Heartbeat System ---
Public Sub SendHeartbeat(ByVal udpSocket As Long, ByVal destIP As String, ByVal destPort As Long)
    Dim heartbeatMsg As String
    heartbeatMsg = "HEARTBEAT:" & format(Now, "yyyy-mm-dd hh:mm:ss") & ":" & Environ("COMPUTERNAME")
    SendUDPAdvanced udpSocket, heartbeatMsg, destIP, destPort, UDP_HEARTBEAT
End Sub

Public Sub StartHeartbeatService(ByVal udpSocket As Long, ByVal destIP As String, ByVal destPort As Long)
    ' This would typically run in a separate thread
    ' For VBA, you'd call this periodically from your main loop
    Static lastHeartbeat As Long
    
    If GetTickCount() - lastHeartbeat > heartbeatInterval Then
        SendHeartbeat udpSocket, destIP, destPort
        lastHeartbeat = GetTickCount()
    End If
End Sub

' --- Enhanced HTML Dashboard ---
Public Function GetAdvancedUDPStatusHTML() As String
    Dim html As String
    html = "<!DOCTYPE html><html><head>" & vbCrLf
    html = html & "<title>Advanced UDP Server Dashboard</title>" & vbCrLf
    html = html & "<meta http-equiv='refresh' content='3'>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #333; }" & vbCrLf
    html = html & ".container { max-width: 1400px; margin: 0 auto; padding: 20px; }" & vbCrLf
    html = html & ".header { background: rgba(255,255,255,0.95); padding: 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 8px 32px rgba(0,0,0,0.1); backdrop-filter: blur(10px); }" & vbCrLf
    html = html & ".header h1 { margin: 0; color: #2c3e50; font-size: 2.5em; }" & vbCrLf
    html = html & ".header p { margin: 5px 0 0 0; color: #7f8c8d; }" & vbCrLf
    html = html & ".stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 25px; }" & vbCrLf
    html = html & ".stat-card { background: rgba(255,255,255,0.9); padding: 20px; border-radius: 12px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.1); transition: transform 0.3s; }" & vbCrLf
    html = html & ".stat-card:hover { transform: translateY(-5px); }" & vbCrLf
    html = html & ".stat-value { font-size: 2.2em; font-weight: bold; color: #2c3e50; margin-bottom: 5px; }" & vbCrLf
    html = html & ".stat-label { color: #7f8c8d; font-size: 0.9em; text-transform: uppercase; letter-spacing: 1px; }" & vbCrLf
    html = html & ".section { background: rgba(255,255,255,0.9); margin-bottom: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); overflow: hidden; }" & vbCrLf
    html = html & ".section-header { background: linear-gradient(135deg, #3498db, #2980b9); color: white; padding: 15px 20px; font-size: 1.3em; font-weight: bold; }" & vbCrLf
    html = html & ".section-content { padding: 20px; }" & vbCrLf
    html = html & "table { width: 100%; border-collapse: collapse; }" & vbCrLf
    html = html & "th, td { padding: 12px 8px; text-align: left; border-bottom: 1px solid #ecf0f1; }" & vbCrLf
    html = html & "th { background: linear-gradient(135deg, #34495e, #2c3e50); color: white; font-weight: bold; }" & vbCrLf
    html = html & "tr:hover { background-color: #f8f9fa; }" & vbCrLf
    html = html & ".message-log { background: #2c3e50; color: #ecf0f1; padding: 15px; border-radius: 8px; font-family: 'Consolas', monospace; max-height: 400px; overflow-y: auto; line-height: 1.4; }" & vbCrLf
    html = html & ".status-online { color: #27ae60; font-weight: bold; }" & vbCrLf
    html = html & ".status-blocked { color: #e74c3c; font-weight: bold; }" & vbCrLf
    html = html & ".security-high { color: #f39c12; }" & vbCrLf
    html = html & ".security-medium { color: #3498db; }" & vbCrLf
    html = html & ".security-low { color: #95a5a6; }" & vbCrLf
    html = html & ".controls { margin: 15px 0; }" & vbCrLf
    html = html & ".btn { display: inline-block; padding: 8px 16px; margin: 0 5px 5px 0; background: #3498db; color: white; text-decoration: none; border-radius: 5px; font-size: 0.9em; transition: background 0.3s; }" & vbCrLf
    html = html & ".btn:hover { background: #2980b9; }" & vbCrLf
    html = html & ".btn-danger { background: #e74c3c; }" & vbCrLf
    html = html & ".btn-danger:hover { background: #c0392b; }" & vbCrLf
    html = html & ".btn-success { background: #27ae60; }" & vbCrLf
    html = html & ".btn-success:hover { background: #229954; }" & vbCrLf
    html = html & "</style>" & vbCrLf
    html = html & "</head><body>" & vbCrLf
    
    html = html & "<div class='container'>" & vbCrLf
    html = html & "<div class='header'>" & vbCrLf
    html = html & "<h1>Advanced UDP Server Dashboard</h1>" & vbCrLf
    html = html & "<p>Enterprise-grade UDP monitoring with security, compression, and advanced analytics</p>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    ' Enhanced Statistics Cards
    html = html & "<div class='stats-grid'>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & udpStats.TotalMessagesReceived & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Messages Received</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & udpStats.TotalMessagesSent & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Messages Sent</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & FormatBytes(udpStats.TotalBytesReceived) & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Data Received</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & FormatBytes(udpStats.TotalBytesSent) & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Data Sent</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & udpConnections.count & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Active Connections</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & udpStats.EncryptedPackets & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Encrypted Packets</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & format(udpStats.CompressionRatio * 100, "0.0") & "%</div>" & vbCrLf
    html = html & "<div class='stat-label'>Compression Ratio</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='stat-card'>" & vbCrLf
    html = html & "<div class='stat-value'>" & blockedIPs.count & "</div>" & vbCrLf
    html = html & "<div class='stat-label'>Blocked IPs</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    ' Security Controls
    html = html & "<div class='section'>" & vbCrLf
    html = html & "<div class='section-header'>Security Controls</div>" & vbCrLf
    html = html & "<div class='section-content'>" & vbCrLf
    html = html & "<div class='controls'>" & vbCrLf
    html = html & "<a href='/api/udp/security/enable' class='btn btn-success'>Enable Security</a>" & vbCrLf
    html = html & "<a href='/api/udp/security/disable' class='btn btn-danger'>Disable Security</a>" & vbCrLf
    html = html & "<a href='/api/udp/compression/toggle' class='btn'>Toggle Compression</a>" & vbCrLf
    html = html & "<a href='/api/udp/stats/reset' class='btn'>Reset Stats</a>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<p><strong>Security Status:</strong> " & IIf(securityEnabled, "<span class='status-online'>Enabled</span>", "<span class='status-blocked'>Disabled</span>") & "</p>" & vbCrLf
    html = html & "<p><strong>Compression:</strong> " & IIf(compressionEnabled, "<span class='status-online'>Enabled</span>", "<span class='status-blocked'>Disabled</span>") & "</p>" & vbCrLf
    html = html & "<p><strong>Heartbeat Interval:</strong> " & heartbeatInterval / 1000 & " seconds</p>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    ' Enhanced Connections Table
    html = html & "<div class='section'>" & vbCrLf
    html = html & "<div class='section-header'>Active UDP Connections</div>" & vbCrLf
    html = html & "<div class='section-content'>" & vbCrLf
    html = html & "<table>" & vbCrLf
    html = html & "<tr><th>IP Address</th><th>Port</th><th>Messages</th><th>Data Sent</th><th>Data Received</th><th>Latency</th><th>Security</th><th>Status</th><th>Last Activity</th></tr>" & vbCrLf
    
    If udpConnections.count > 0 Then
        Dim i As Long
        For i = 1 To udpConnections.count
            Dim conn As UDPConnectionInfo
            conn = udpConnections.item(i)
            Dim lastActivity As String
            Dim securityLevel As String
            Dim status As String
            
            lastActivity = format(DateAdd("s", (GetTickCount() - conn.lastActivity) / 1000 * -1, Now), "hh:mm:ss")
            
            Select Case conn.securityLevel
                Case 0: securityLevel = "<span class='security-low'>Low</span>"
                Case 1: securityLevel = "<span class='security-medium'>Medium</span>"
                Case 2: securityLevel = "<span class='security-high'>High</span>"
                Case Else: securityLevel = "<span class='security-low'>Unknown</span>"
            End Select
            
            If conn.IsBlocked Then
                status = "<span class='status-blocked'>Blocked</span>"
            Else
                status = "<span class='status-online'>Active</span>"
            End If
            
            html = html & "<tr>" & vbCrLf
            html = html & "<td>" & conn.remoteIP & "</td>" & vbCrLf
            html = html & "<td>" & conn.remotePort & "</td>" & vbCrLf
            html = html & "<td>" & conn.messageCount & "</td>" & vbCrLf
            html = html & "<td>" & FormatBytes(conn.bytesSent) & "</td>" & vbCrLf
            html = html & "<td>" & FormatBytes(conn.bytesReceived) & "</td>" & vbCrLf
            html = html & "<td>" & conn.AverageLatency & "ms</td>" & vbCrLf
            html = html & "<td>" & securityLevel & "</td>" & vbCrLf
            html = html & "<td>" & status & "</td>" & vbCrLf
            html = html & "<td>" & lastActivity & "</td>" & vbCrLf
            html = html & "</tr>" & vbCrLf
        Next i
    Else
        html = html & "<tr><td colspan='9' style='text-align: center; color: #7f8c8d; padding: 30px;'>No active connections</td></tr>" & vbCrLf
    End If
    
    html = html & "</table>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    ' Enhanced Message Log
    html = html & "<div class='section'>" & vbCrLf
    html = html & "<div class='section-header'>Real-time Message Log</div>" & vbCrLf
    html = html & "<div class='section-content'>" & vbCrLf
    html = html & "<div class='message-log'>" & vbCrLf
    
    If recentMessages.count > 0 Then
        For i = recentMessages.count To 1 Step -1
            html = html & recentMessages.item(i) & "<br>" & vbCrLf
        Next i
    Else
        html = html & "Waiting for UDP messages..." & vbCrLf
    End If
    
    html = html & "</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    html = html & "</div>" & vbCrLf
    html = html & "</body></html>"
    
    GetAdvancedUDPStatusHTML = html
End Function

' --- Configuration Management ---
Public Sub EnableSecurity()
    securityEnabled = True
    DebuggingLog.DebugLog "UDP security enabled"
End Sub

Public Sub DisableSecurity()
    securityEnabled = False
    DebuggingLog.DebugLog "UDP security disabled"
End Sub

Public Sub EnableCompression()
    compressionEnabled = True
    DebuggingLog.DebugLog "UDP compression enabled"
End Sub

Public Sub DisableCompression()
    compressionEnabled = False
    DebuggingLog.DebugLog "UDP compression disabled"
End Sub

Public Sub SetHeartbeatInterval(ByVal intervalMs As Long)
    heartbeatInterval = intervalMs
    DebuggingLog.DebugLog "Heartbeat interval set to " & intervalMs & "ms"
End Sub

' --- API Endpoints for HTTP Integration ---
Public Function HandleUDPAPIRequest(ByVal path As String) As String
    Select Case path
        Case "/api/udp/stats"
            HandleUDPAPIRequest = GetUDPStatsJSON()
        Case "/api/udp/stats/reset"
            ResetUDPStats
            HandleUDPAPIRequest = "{""status"":""success"",""message"":""Statistics reset""}"
        Case "/api/udp/security/enable"
            EnableSecurity
            HandleUDPAPIRequest = "{""status"":""success"",""security"":""enabled""}"
        Case "/api/udp/security/disable"
            DisableSecurity
            HandleUDPAPIRequest = "{""status"":""success"",""security"":""disabled""}"
        Case "/api/udp/compression/toggle"
            compressionEnabled = Not compressionEnabled
            HandleUDPAPIRequest = "{""status"":""success"",""compression"":" & LCase(compressionEnabled) & "}"
        Case "/api/udp/connections"
            HandleUDPAPIRequest = GetConnectionsJSON()
        Case Else
            HandleUDPAPIRequest = "{""status"":""error"",""message"":""Unknown API endpoint""}"
    End Select
End Function

Private Function GetConnectionsJSON() As String
    Dim json As String
    json = "{""connections"":["
    
    If udpConnections.count > 0 Then
        Dim i As Long
        For i = 1 To udpConnections.count
            Dim conn As UDPConnectionInfo
            conn = udpConnections.item(i)
            
            If i > 1 Then json = json & ","
            json = json & "{"
            json = json & """ip"":""" & conn.remoteIP & ""","
            json = json & """port"":" & conn.remotePort & ","
            json = json & """messages"":" & conn.messageCount & ","
            json = json & """bytesSent"":" & conn.bytesSent & ","
            json = json & """bytesReceived"":" & conn.bytesReceived & ","
            json = json & """latency"":" & conn.AverageLatency & ","
            json = json & """blocked"":" & LCase(conn.IsBlocked)
            json = json & "}"
        Next i
    End If
    
    json = json & "]}"
    GetConnectionsJSON = json
End Function

' --- Enhanced Statistics ---
Public Function GetUDPStatsJSON() As String
    Dim json As String
    json = "{" & vbCrLf
    json = json & """totalMessagesSent"": " & udpStats.TotalMessagesSent & "," & vbCrLf
    json = json & """totalMessagesReceived"": " & udpStats.TotalMessagesReceived & "," & vbCrLf
    json = json & """totalBytesSent"": " & udpStats.TotalBytesSent & "," & vbCrLf
    json = json & """totalBytesReceived"": " & udpStats.TotalBytesReceived & "," & vbCrLf
    json = json & """errorCount"": " & udpStats.ErrorCount & "," & vbCrLf
    json = json & """packetsDropped"": " & udpStats.PacketsDropped & "," & vbCrLf
    json = json & """encryptedPackets"": " & udpStats.EncryptedPackets & "," & vbCrLf
    json = json & """compressionRatio"": " & udpStats.CompressionRatio & "," & vbCrLf
    json = json & """activeConnections"": " & udpConnections.count & "," & vbCrLf
    json = json & """blockedIPs"": " & blockedIPs.count & "," & vbCrLf
    json = json & """uptime"": " & (GetTickCount() - udpStats.startTime) & "," & vbCrLf
    json = json & """securityEnabled"": " & LCase(securityEnabled) & "," & vbCrLf
    json = json & """compressionEnabled"": " & LCase(compressionEnabled) & "," & vbCrLf
    json = json & """heartbeatInterval"": " & heartbeatInterval & vbCrLf
    json = json & "}"
    GetUDPStatsJSON = json
End Function

' --- Cleanup and Reset ---
Public Sub ResetUDPStats()
    udpStats.TotalMessagesSent = 0
    udpStats.TotalMessagesReceived = 0
    udpStats.TotalBytesSent = 0
    udpStats.TotalBytesReceived = 0
    udpStats.ErrorCount = 0
    udpStats.PacketsDropped = 0
    udpStats.EncryptedPackets = 0
    udpStats.CompressionRatio = 0
    udpStats.startTime = GetTickCount()
    Set udpConnections = New Collection
    Set recentMessages = New Collection
    DebuggingLog.DebugLog "Advanced UDP statistics reset"
End Sub

Public Sub CleanupUDP()
    If isWinsockInitialized Then
        WSACleanup
        isWinsockInitialized = False
        Set udpConnections = Nothing
        Set recentMessages = Nothing
        Set blockedIPs = Nothing
        Set allowedIPs = Nothing
        Set udpSockets = Nothing
        Set messageQueues = Nothing
        DebuggingLog.DebugLog "Advanced UDP Module cleaned up"
    End If
End Sub

' --- Helper Functions ---
Private Function FormatBytes(ByVal bytes As Long) As String
    If bytes < 1024 Then
        FormatBytes = bytes & " B"
    ElseIf bytes < 1048576 Then
        FormatBytes = format(bytes / 1024, "0.0") & " KB"
    ElseIf bytes < 1073741824 Then
        FormatBytes = format(bytes / 1048576, "0.0") & " MB"
    Else
        FormatBytes = format(bytes / 1073741824, "0.0") & " GB"
    End If
End Function

' --- Original Functions (Enhanced) ---
Public Function SendUDP(ByVal udpSocket As Long, ByVal message As String, ByVal destIP As String, ByVal destPort As Long) As Long
    SendUDP = SendUDPAdvanced(udpSocket, message, destIP, destPort, UDP_DATA, compressionEnabled, securityEnabled)
End Function

Public Function RecvUDP(ByVal udpSocket As Long, ByRef buffer() As Byte, ByRef fromIP As String, ByRef fromPort As Long) As Long
    ' Enhanced receive with packet processing
    Dim addr As SOCKADDR_IN
    Dim addrLen As Long
    addrLen = LenB(addr)
    
    RecvUDP = recvfrom(udpSocket, buffer(0), UBound(buffer) + 1, 0, addr, addrLen)
    
    If RecvUDP <> SOCKET_ERROR Then
        fromPort = ntohs(addr.sin_port)
        #If VBA7 Then
            Dim ptr As LongPtr
            ptr = inet_ntoa(addr.sin_addr)
            fromIP = StringFromPtrA(ptr)
        #Else
            fromIP = StringFromPtrA(inet_ntoa(addr.sin_addr))
        #End If
        
        ' Security and rate limiting checks
        If IsIPBlocked(fromIP) Then
            DebuggingLog.DebugLog "Blocked packet from: " & fromIP
            RecvUDP = -1
            Exit Function
        End If
        
        If Not CheckRateLimit(fromIP) Then
            udpStats.PacketsDropped = udpStats.PacketsDropped + 1
            RecvUDP = -1
            Exit Function
        End If
        
        ' Update statistics
        udpStats.TotalMessagesReceived = udpStats.TotalMessagesReceived + 1
        udpStats.TotalBytesReceived = udpStats.TotalBytesReceived + RecvUDP
        
        ' Track connection
        UpdateConnectionInfo fromIP, fromPort, 0, RecvUDP
        
        ' Log recent message
        Dim receivedMsg As String
        receivedMsg = Left(StrConv(buffer, vbUnicode), RecvUDP)
        LogRecentMessage "RECV", receivedMsg, fromIP, fromPort, RecvUDP
        
        DebuggingLog.DebugLog "UDP received " & RecvUDP & " bytes from " & fromIP & ":" & fromPort
    ElseIf WSAGetLastError <> WSAEWOULDBLOCK Then
        udpStats.ErrorCount = udpStats.ErrorCount + 1
    End If
End Function

' --- Standard Functions (maintained for compatibility) ---
Public Function CreateUDPSocket() As Long
    CreateUDPSocket = CreateAdvancedUDPSocket()
End Function

Public Function BindUDP(ByVal udpSocket As Long, ByVal port As Long, Optional ByVal bindAddr As String = "0.0.0.0") As Boolean
    Dim addr As SOCKADDR_IN
    addr.sin_family = AF_INET
    addr.sin_port = htons(port)
    addr.sin_addr = inet_addr(bindAddr)
    If bind(udpSocket, addr, LenB(addr)) = SOCKET_ERROR Then
        BindUDP = False
        udpStats.ErrorCount = udpStats.ErrorCount + 1
    Else
        BindUDP = True
        DebuggingLog.DebugLog "UDP socket bound to port " & port
    End If
End Function

Public Function SetUDPNonBlocking(ByVal udpSocket As Long) As Boolean
    Dim nonBlocking As Long
    nonBlocking = 1
    If ioctlsocket(udpSocket, FIONBIO, nonBlocking) = SOCKET_ERROR Then
        SetUDPNonBlocking = False
        udpStats.ErrorCount = udpStats.ErrorCount + 1
    Else
        SetUDPNonBlocking = True
    End If
End Function



Public Function JoinMulticastGroup(ByVal udpSocket As Long, ByVal multicastAddr As String, Optional ByVal localInterface As String = "0.0.0.0") As Boolean
    Dim mreq As ip_mreq
    mreq.imr_multiaddr = inet_addr(multicastAddr)
    mreq.imr_interface = inet_addr(localInterface)
    If setsockopt(udpSocket, IPPROTO_IP, IP_ADD_MEMBERSHIP, mreq, LenB(mreq)) = SOCKET_ERROR Then
        JoinMulticastGroup = False
        udpStats.ErrorCount = udpStats.ErrorCount + 1
    Else
        JoinMulticastGroup = True
        If Len(udpStats.MulticastGroups) > 0 Then
            udpStats.MulticastGroups = udpStats.MulticastGroups & "," & multicastAddr
        Else
            udpStats.MulticastGroups = multicastAddr
        End If
        DebuggingLog.DebugLog "Joined multicast group: " & multicastAddr
    End If
End Function

Public Function LeaveMulticastGroup(ByVal udpSocket As Long, ByVal multicastAddr As String, Optional ByVal localInterface As String = "0.0.0.0") As Boolean
    Dim mreq As ip_mreq
    mreq.imr_multiaddr = inet_addr(multicastAddr)
    mreq.imr_interface = inet_addr(localInterface)
    If setsockopt(udpSocket, IPPROTO_IP, IP_DROP_MEMBERSHIP, mreq, LenB(mreq)) = SOCKET_ERROR Then
        LeaveMulticastGroup = False
        udpStats.ErrorCount = udpStats.ErrorCount + 1
    Else
        LeaveMulticastGroup = True
        DebuggingLog.DebugLog "Left multicast group: " & multicastAddr
    End If
End Function

Public Function SetMulticastTTL(ByVal udpSocket As Long, ByVal ttl As Long) As Boolean
    If setsockopt(udpSocket, IPPROTO_IP, IP_MULTICAST_TTL, ttl, 4) = SOCKET_ERROR Then
        SetMulticastTTL = False
        udpStats.ErrorCount = udpStats.ErrorCount + 1
    Else
        SetMulticastTTL = True
        DebuggingLog.DebugLog "Multicast TTL set to: " & ttl
    End If
End Function

Public Sub CloseUDPSocket(ByVal udpSocket As Long)
    If udpSocket <> INVALID_SOCKET Then
        closesocket udpSocket
        DebuggingLog.DebugLog "UDP socket closed"
    End If
End Sub

Public Function GetUDPError() As Long
    GetUDPError = WSAGetLastError
End Function





' --- Non-blocking loop to receive messages ---
Public Sub UDPServerLoop()
    If Not udpServerRunning Then Exit Sub
    
    Dim buffer(1024) As Byte
    Dim fromaddr As SOCKADDR_IN
    Dim addrLen As Long
    addrLen = LenB(fromaddr)
    
    Dim bytesRead As Long
    bytesRead = recvfrom(udpServerSocket, buffer(0), 1024, 0, fromaddr, addrLen)
    
    If bytesRead > 0 Then
        ' Convert bytes to string and handle message
        Dim msg As String
        msg = StrConv(buffer, vbUnicode)
        Debug.Print "Received UDP: " & msg
        ' TODO: Call your message handler here
    End If
    
    ' Schedule next loop iteration (non-blocking)
    Application.OnTime Now + TimeValue("00:00:01"), "UDPServerLoop"
End Sub

Public Function EnableUDPBroadcast(ByVal udpSocket As Long) As Boolean
    Dim broadcast As Long
    broadcast = 1
    If setsockopt(udpSocket, SOL_SOCKET, SO_BROADCAST, broadcast, 4) = SOCKET_ERROR Then
        EnableUDPBroadcast = False
        udpStats.ErrorCount = udpStats.ErrorCount + 1
        DebuggingLog.DebugLog "Failed to enable UDP broadcast on socket"
    Else
        EnableUDPBroadcast = True
        DebuggingLog.DebugLog "UDP broadcast enabled on socket"
    End If
End Function



Private Function ntohs(ByVal netshort As Integer) As Integer
    ntohs = htons(netshort)
End Function


' --- Enable UDP queue processing ---
Public Sub StartUDPQueue()
    udpLoopEnabled = True
    InitUDPQueue
End Sub

' --- Disable UDP queue processing ---
Public Sub StopUDPQueue()
    udpLoopEnabled = False
    Set udpQueue = Nothing
End Sub





