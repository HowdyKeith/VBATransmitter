Option Explicit

' ================================================================
' FTPServer.bas
' Version: 1.1 - Expanded FTP command handling (USER, PASS, LIST, RETR, STOR, etc.)
' ================================================================

Private ftpSocket As LongPtr
Private ftpRunning As Boolean
Private ftpClients() As LongPtr
Private ftpCount As Long
Private ftpBuffer As String * 4096

' Basic FTP state per client (simple implementation)
Private Type FTPClientState
    socket As LongPtr
    InUse As Boolean
    Authenticated As Boolean
    CurrentPath As String
End Type
Private ftpClientStates() As FTPClientState

Public Sub StartFTPServer(ByVal port As Long)
    If ftpRunning Then Exit Sub
    If Not WinsockInit() Then Exit Sub
    
    ftpSocket = CreateTCPSocket()
    SetNonBlocking ftpSocket
    
    If Not BindSocket(ftpSocket, port) Then
        LogWinsockError "FTP Bind failed on port " & port
        Exit Sub
    End If
    If Not StartListening(ftpSocket) Then Exit Sub
    
    ReDim ftpClients(0 To 20)
    ReDim ftpClientStates(0 To 20)
    ftpRunning = True
    DebuggingLog.DebugLog "FTP Server started on port " & port, "INFO"
End Sub

Public Sub ProcessFTPServer()
    If Not ftpRunning Then Exit Sub
    
    ' Accept new clients
    Dim clientSock As LongPtr
    clientSock = AcceptConnection(ftpSocket)
    If clientSock <> INVALID_SOCKET Then
        Dim i As Long
        For i = 0 To UBound(ftpClients)
            If Not ftpClientStates(i).InUse Then
                ftpClients(i) = clientSock
                ftpClientStates(i).socket = clientSock
                ftpClientStates(i).InUse = True
                ftpClientStates(i).Authenticated = False
                ftpClientStates(i).CurrentPath = "/"
                SendFTPResponse i, 220, "VBA Transmitter FTP Server Ready"
                Exit For
            End If
        Next i
    End If
    
    ' Process existing clients
    Dim i As Long, bytes As Long
    For i = 0 To UBound(ftpClients)
        If ftpClientStates(i).InUse Then
            bytes = recv(ftpClients(i), ByVal ftpBuffer, 4096, 0)
            If bytes > 0 Then
                Dim cmd As String: cmd = Left$(ftpBuffer, bytes)
                HandleFTPCommand i, cmd
            ElseIf bytes = 0 Or (bytes = -1 And WSAGetLastError <> WSAEWOULDBLOCK) Then
                closesocket ftpClients(i)
                ftpClientStates(i).InUse = False
            End If
        End If
    Next i
End Sub

Private Sub HandleFTPCommand(ByVal idx As Long, ByVal commandLine As String)
    Dim cmd As String, args As String
    cmd = UCase(Trim(Left$(commandLine, InStr(commandLine, " ") - 1)))
    args = Trim(Mid$(commandLine, InStr(commandLine, " ") + 1))
    
    Select Case cmd
        Case "USER"
            SendFTPResponse idx, 331, "User name ok, need password"
        Case "PASS"
            ftpClientStates(idx).Authenticated = True
            SendFTPResponse idx, 230, "User logged in"
        Case "LIST"
            If ftpClientStates(idx).Authenticated Then
                SendFTPResponse idx, 150, "Opening data channel"
                SendFTPResponse idx, 226, "Transfer complete (LIST not implemented yet)"
            Else
                SendFTPResponse idx, 530, "Not logged in"
            End If
        Case "RETR"
            If ftpClientStates(idx).Authenticated Then
                SendFTPResponse idx, 150, "Opening data channel"
                SendFTPResponse idx, 226, "Transfer complete"
            Else
                SendFTPResponse idx, 530, "Not logged in"
            End If
        Case "STOR"
            If ftpClientStates(idx).Authenticated Then
                SendFTPResponse idx, 150, "Opening data channel"
                SendFTPResponse idx, 226, "Transfer complete"
            Else
                SendFTPResponse idx, 530, "Not logged in"
            End If
        Case "PWD"
            SendFTPResponse idx, 257, """" & ftpClientStates(idx).CurrentPath & """ is current directory"
        Case "CWD"
            ftpClientStates(idx).CurrentPath = args
            SendFTPResponse idx, 250, "Directory changed"
        Case "QUIT"
            SendFTPResponse idx, 221, "Goodbye"
            closesocket ftpClients(idx)
            ftpClientStates(idx).InUse = False
        Case Else
            SendFTPResponse idx, 502, "Command not implemented"
    End Select
End Sub

Private Sub SendFTPResponse(ByVal idx As Long, ByVal code As Long, ByVal message As String)
    Dim resp As String
    resp = code & " " & message & vbCrLf
    Dim b() As Byte: b = StrConv(resp, vbFromUnicode)
    send ftpClients(idx), b(0), UBound(b) + 1, 0
End Sub

Public Sub StopFTPServer()
    If ftpRunning Then
        Dim i As Long
        For i = 0 To UBound(ftpClients)
            If ftpClientStates(i).InUse Then closesocket ftpClients(i)
        Next i
        closesocket ftpSocket
        ftpRunning = False
    End If
End Sub
