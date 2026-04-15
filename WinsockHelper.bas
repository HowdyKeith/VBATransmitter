Option Explicit

' ================================================================
' WinsockHelper.bas
' Version: 1.2 - Centralized Winsock for all modules
' ================================================================
#If VBA7 Then
    Private Declare PtrSafe Function connect Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef addr As SOCKADDR_IN, ByVal namelen As Long) As Long
#Else
    Private Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByVal namelen As Long) As Long
#End If
#If VBA7 Then
    Public Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
    Public Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Public Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal sType As Long, ByVal protocol As Long) As LongPtr
    Public Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As LongPtr) As Long
    Public Declare PtrSafe Function bind Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef addr As Any, ByVal namelen As Long) As Long
    Public Declare PtrSafe Function listen Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal backlog As Long) As Long
    Public Declare PtrSafe Function accept Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef addr As Any, ByRef addrLen As Long) As LongPtr
    Public Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare PtrSafe Function send Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long
    Public Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
    Public Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Public Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
    Public Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
#Else
    Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
    Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal sType As Long, ByVal protocol As Long) As Long
    Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As Any, ByVal namelen As Long) As Long
    Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Public Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As Any, ByRef addrLen As Long) As Long
    Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
    Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
    Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
#End If

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
End Type

Public Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Public Const AF_INET As Long = 2
Public Const SOCK_STREAM As Long = 1
Public Const IPPROTO_TCP As Long = 6
Public Const SOL_SOCKET As Long = &HFFFF&
Public Const SO_REUSEADDR As Long = &H4&
Public Const INVALID_SOCKET As LongPtr = -1
Public Const FIONBIO As Long = &H8004667E
Public Const WSAEWOULDBLOCK As Long = 10035
Public Const SOMAXCONN As Long = 5

Public Function WinsockInit() As Boolean
    Dim wsa As WSADATA
    WinsockInit = (WSAStartup(&H202, wsa) = 0)
    If Not WinsockInit Then LogWinsockError "WSAStartup failed"
End Function

Public Sub WinsockCleanup()
    WSACleanup
End Sub

Public Function CreateTCPSocket() As LongPtr
    CreateTCPSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
End Function

Public Function SetNonBlocking(ByVal s As LongPtr) As Boolean
    Dim nb As Long: nb = 1
    SetNonBlocking = (ioctlsocket(s, FIONBIO, nb) = 0)
End Function

Public Function BindSocket(ByVal s As LongPtr, ByVal port As Long, Optional ip As String = "0.0.0.0") As Boolean
    Dim addr As SOCKADDR_IN
    With addr
        .sin_family = AF_INET
        .sin_port = htons(port)
        .sin_addr = inet_addr(ip)
    End With
    BindSocket = (bind(s, addr, LenB(addr)) = 0)
End Function

Public Function StartListening(ByVal s As LongPtr, Optional backlog As Long = SOMAXCONN) As Boolean
    StartListening = (listen(s, backlog) = 0)
End Function

Public Function AcceptConnection(ByVal serverSocket As LongPtr) As LongPtr
    Dim addr As SOCKADDR_IN, addrLen As Long
    addrLen = LenB(addr)
    AcceptConnection = accept(serverSocket, addr, addrLen)
End Function

Public Sub LogWinsock(message As String)
    Debug.Print "[Winsock] " & format(Now, "hh:nn:ss") & " " & message
End Sub

Public Sub LogWinsockError(message As String)
    Debug.Print "[Winsock ERROR] " & format(Now, "hh:nn:ss") & " " & message & " (Code: " & WSAGetLastError & ")"
End Sub
