Option Explicit
'***************************************************************
' TrafficManager Module (Clean, Tick-Based Live Loop)
' Purpose: Central loop/manager for HTTP, UDP, FTP, IoT, MQTT, Outlook
'***************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' --- Constants ---
Private Const VK_ESCAPE As Long = &H1B
Private Const SERVER_LOOP_SLEEP_MS As Long = 10 ' Small sleep to reduce CPU

' --- Manager state ---
Private g_escapePressed As Boolean
Private isRunning As Boolean
Private lastUpdate As Date
Private loopCount As Long

' --- MQTT ---
Private mqttEnabled As Boolean
Private lastMQTTHeartbeat As Long
' --- Subsystem status (add near other Private declarations) ---
Private statusHTTP As Boolean
Private statusUDP As Boolean
Private statusMQTT As Boolean
Private statusIoT As Boolean
Private statusFTP As Boolean

Private lastErrorHTTP As String
Private lastErrorUDP As String
Private lastErrorMQTT As String
Private lastErrorIoT As String
Private lastErrorFTP As String

'***************************************************************
' Initialize
'***************************************************************
Public Sub InitializeTrafficManager()
    On Error GoTo ErrHandler
    DebuggingLog.DebugLog "[TrafficManager] Initializing..."
    
    isRunning = False
    g_escapePressed = False
    lastUpdate = Now
    loopCount = 0
    lastMQTTHeartbeat = GetTickCount()
    
    DebuggingLog.DebugLog "[TrafficManager] Initialized"
    Exit Sub

ErrHandler:
    DebuggingLog.DebugLog "[TrafficManager] Initialize error: " & Err.description
End Sub

'***************************************************************
' Run All Servers Loop
'***************************************************************
Public Sub RunAllServers()
    On Error GoTo ErrorHandler
    DebuggingLog.DebugLog "[TrafficManager] Starting all servers..."
    
    If isRunning Then Exit Sub
    isRunning = True
    lastUpdate = Now
    loopCount = 0
    
    ' --- Start servers ---
    If Not HttpServer.isHTTPRunning() Then HttpServer.StartHttpServer 8080
    If Not TransmissionServer.GetUDPRunning() Then TransmissionServer.StartUDPServer 10004
    If Not IoTGateway.GetRunning() Then IoTGateway.StartServer
    If Not FTPServer.GetRunning() Then FTPServer.StartServer
    On Error Resume Next
    OutlookExternal.LaunchOutlookExternal
    On Error GoTo ErrorHandler
    
    DebuggingLog.DebugLog "[TrafficManager] All servers started, entering main loop..."
    
    ' --- Tick-based continuous loop ---
    Do While isRunning
        Dim tickStart As Long
        tickStart = GetTickCount()
        
        ProcessAllServers
        
        Sleep SERVER_LOOP_SLEEP_MS
        DoEvents
        
        If GetAsyncKeyState(VK_ESCAPE) <> 0 Then
            g_escapePressed = True
            DebuggingLog.DebugLog "[TrafficManager] Escape pressed, stopping servers..."
            StopAllServers
            Exit Do
        End If
        
        ' Keep loop at least SERVER_LOOP_SLEEP_MS
        Do While GetTickCount() - tickStart < SERVER_LOOP_SLEEP_MS
            DoEvents
        Loop
    Loop
    
    DebuggingLog.DebugLog "[TrafficManager] RunAllServers loop exited."
    Exit Sub

ErrorHandler:
    DebuggingLog.DebugLog "[TrafficManager] RunAllServers error: " & Err.description
    StopAllServers
End Sub

'***************************************************************
' Process All Servers Tick
' --- Helper to safely check object type ---
Private Function IsValidServer(obj As Variant, expectedClassName As String) As Boolean
    On Error Resume Next
    IsValidServer = False
    If Not IsEmpty(obj) Then
        If VarType(obj) = vbObject Then
            If TypeName(obj) = expectedClassName Then IsValidServer = True
        End If
    End If
    On Error GoTo 0
End Function



' --- Generic server tick processor ---
Private Sub ProcessServerTick(obj As Variant, expectedClassName As String, processRoutine As String, ByRef statusFlag As Boolean, ByRef lastError As String)
    statusFlag = False
    lastError = ""
    
    If IsValidServer(obj, expectedClassName) Then
        On Error Resume Next
        ' Call the routine dynamically via CallByName
        CallByName obj, processRoutine, VbMethod
        If Err.Number = 0 Then
            statusFlag = True
        Else
            lastError = processRoutine & " error: " & Err.description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        lastError = expectedClassName & " object invalid or missing"
    End If
End Sub

' --- Global tracking of last-known status ---
Private lastStatusHTTP As Boolean
Private lastStatusUDP As Boolean
Private lastStatusMQTT As Boolean
Private lastStatusIoT As Boolean
Private lastStatusFTP As Boolean

Public Sub ProcessAllServers()
    On Error GoTo ErrorHandler
    If Not isRunning Then Exit Sub
    
    Dim loopStart As Double
    loopStart = Timer
    
    ' -----------------------------
    ' HTTP Server
    ' -----------------------------
    Dim currentHTTP As Boolean
    On Error Resume Next
    currentHTTP = HttpServer.isHTTPRunning()
    If Err.Number <> 0 Then
        lastErrorHTTP = "HttpServer unavailable: " & Err.description
        currentHTTP = False
        Debug.Print "[TrafficManager] HTTP error: " & lastErrorHTTP
        Err.Clear
    Else
        If currentHTTP Then HttpServer.ProcessHttpServer
        lastErrorHTTP = ""
    End If
    statusHTTP = currentHTTP
    On Error GoTo ErrorHandler
    
    ' -----------------------------
    ' UDP Server
    ' -----------------------------
    Dim currentUDP As Boolean
    On Error Resume Next
    currentUDP = TransmissionServer.GetUDPRunning()
    If Err.Number <> 0 Then
        lastErrorUDP = "TransmissionServer unavailable: " & Err.description
        currentUDP = False
        Debug.Print "[TrafficManager] UDP error: " & lastErrorUDP
        Err.Clear
    Else
        If currentUDP Then TransmissionServer.ProcessUDPServerTick
        lastErrorUDP = ""
    End If
    statusUDP = currentUDP
    On Error GoTo ErrorHandler
    
    ' -----------------------------
    ' MQTT
    ' -----------------------------
    Dim currentMQTT As Boolean
    On Error Resume Next
    If mqttEnabled Then
        ProcessMQTTTickAdvanced
        If Err.Number <> 0 Then
            lastErrorMQTT = "ProcessMQTTTickAdvanced error: " & Err.description
            currentMQTT = False
            Debug.Print "[TrafficManager] MQTT error: " & lastErrorMQTT
            Err.Clear
        Else
            currentMQTT = True
            ' heartbeat
            If GetTickCount() - lastMQTTHeartbeat > 30000 Then
                PublishMQTTAdvanced "vba/heartbeat", "VBA TrafficManager alive at " & format(Now, "yyyy-mm-dd hh:nn:ss"), MQTT_QOS0
                PublishMQTTAdvanced "system/stats", GetEnhancedTrafficManagerStatus(), MQTT_QOS1
                lastMQTTHeartbeat = GetTickCount()
            End If
            lastErrorMQTT = ""
        End If
    Else
        lastErrorMQTT = "MQTT disabled"
        currentMQTT = False
    End If
    statusMQTT = currentMQTT
    On Error GoTo ErrorHandler
    
    ' -----------------------------
    ' IoT Gateway
    ' -----------------------------
    Dim currentIoT As Boolean
    On Error Resume Next
    currentIoT = IoTGateway.GetRunning()
    If Err.Number <> 0 Then
        lastErrorIoT = "IoTGateway unavailable: " & Err.description
        currentIoT = False
        Debug.Print "[TrafficManager] IoT error: " & lastErrorIoT
        Err.Clear
    Else
        If currentIoT Then IoTGateway.ProcessIoTServer
        lastErrorIoT = ""
    End If
    statusIoT = currentIoT
    On Error GoTo ErrorHandler
    
    ' -----------------------------
    ' FTP Server
    ' -----------------------------
    Dim currentFTP As Boolean
    On Error Resume Next
    currentFTP = FTPServer.GetRunning()
    If Err.Number <> 0 Then
        lastErrorFTP = "FTPServer unavailable: " & Err.description
        currentFTP = False
        Debug.Print "[TrafficManager] FTP error: " & lastErrorFTP
        Err.Clear
    Else
        If currentFTP Then FTPServer.ProcessFTPServer
        lastErrorFTP = ""
    End If
    statusFTP = currentFTP
    On Error GoTo ErrorHandler
    
    ' -----------------------------
    ' Heartbeat every 1000 loops
    ' -----------------------------
    loopCount = loopCount + 1
    If loopCount Mod 1000 = 0 Then
        Debug.Print "[TrafficManager] Heartbeat - Loop " & loopCount & _
                    ", HTTP: " & IIf(statusHTTP, "OK", "FAIL") & IIf(lastErrorHTTP <> "", " (" & lastErrorHTTP & ")", "") & _
                    ", UDP: " & IIf(statusUDP, "OK", "FAIL") & IIf(lastErrorUDP <> "", " (" & lastErrorUDP & ")", "") & _
                    ", MQTT: " & IIf(statusMQTT, "OK", "FAIL") & IIf(lastErrorMQTT <> "", " (" & lastErrorMQTT & ")", "") & _
                    ", IoT: " & IIf(statusIoT, "OK", "FAIL") & IIf(lastErrorIoT <> "", " (" & lastErrorIoT & ")", "") & _
                    ", FTP: " & IIf(statusFTP, "OK", "FAIL") & IIf(lastErrorFTP <> "", " (" & lastErrorFTP & ")", "")
    End If
    
    lastUpdate = Now
    
    ' -----------------------------
    ' Escape check
    ' -----------------------------
    If GetAsyncKeyState(VK_ESCAPE) <> 0 Then
        g_escapePressed = True
        Debug.Print "[TrafficManager] Escape pressed, stopping servers..."
        StopAllServers
        Exit Sub
    End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "[TrafficManager] ProcessAllServers fatal error: " & Err.description
    Err.Clear
End Sub






'***************************************************************
' Stop All Servers
'***************************************************************
Public Sub StopAllServers()
    On Error Resume Next
    DebuggingLog.DebugLog "[TrafficManager] Stopping all servers..."
    
    isRunning = False
    
    HttpServer.StopHttpServer
    TransmissionServer.StopUDPServer
    
    If mqttEnabled Then
        CleanupMQTTAdvanced
        mqttEnabled = False
        DebuggingLog.DebugLog "[TrafficManager] Enhanced MQTT stopped"
    End If
    
    IoTGateway.StopServer
    FTPServer.StopServer
    
    Application.OnTime Now + TimeValue("00:00:01"), "TrafficManager.ProcessAllServers", , False
    
    DebuggingLog.DebugLog "[TrafficManager] All servers stopped"
End Sub

'***************************************************************
' Convenience Launcher
'***************************************************************
Public Sub LaunchTrafficDemo()
    On Error GoTo ErrHandler
    InitializeTrafficManager
    RunAllServers
    DebuggingLog.DebugLog "[TrafficManager] LaunchTrafficDemo complete."
    Exit Sub

ErrHandler:
    DebuggingLog.DebugLog "[TrafficManager] LaunchTrafficDemo error: " & Err.description
    StopAllServers
End Sub

'***************************************************************
' Status Functions
'***************************************************************
Public Function GetTrafficManagerStatus() As String
    GetTrafficManagerStatus = "TrafficManager: " & IIf(isRunning, "ACTIVE", "INACTIVE") & _
                              ", Last Update: " & format(lastUpdate, "yyyy-mm-dd hh:mm:ss")
End Function

Public Function GetEnhancedTrafficManagerStatus() As String
    Dim status As String
    status = "TrafficManager: " & IIf(isRunning, "ACTIVE", "INACTIVE")
    status = status & ", Last Update: " & format(lastUpdate, "yyyy-mm-dd hh:nn:ss")
    status = status & ", Loop Count: " & loopCount
    status = status & ", HTTP: " & IIf(statusHTTP, "OK", "FAIL")
    If lastErrorHTTP <> "" Then status = status & " (" & lastErrorHTTP & ")"
    status = status & ", UDP: " & IIf(statusUDP, "OK", "FAIL")
    If lastErrorUDP <> "" Then status = status & " (" & lastErrorUDP & ")"
    status = status & ", MQTT: " & IIf(statusMQTT, "OK", "FAIL")
    If lastErrorMQTT <> "" Then status = status & " (" & lastErrorMQTT & ")"
    status = status & ", IoT: " & IIf(statusIoT, "OK", "FAIL")
    If lastErrorIoT <> "" Then status = status & " (" & lastErrorIoT & ")"
    status = status & ", FTP: " & IIf(statusFTP, "OK", "FAIL")
    If lastErrorFTP <> "" Then status = status & " (" & lastErrorFTP & ")"
    GetEnhancedTrafficManagerStatus = status
End Function


'***************************************************************
' Demo MQTT Utilities
'***************************************************************
Public Sub TestMQTTPublish()
    On Error GoTo ErrHandler
    If mqttEnabled And IsConnected() Then
        PublishMQTTAdvanced "vba/test/demo", "Hello from VBA TrafficManager! " & format(Now, "hh:mm:ss"), MQTT_QOS1
        PublishMQTTAdvanced "sensors/temperature", CStr(20 + Rnd() * 10), MQTT_QOS0
        PublishMQTTAdvanced "sensors/humidity", CStr(40 + Rnd() * 20), MQTT_QOS0
        PublishMQTTAdvanced "system/memory", "Available: " & CStr(Int(Rnd() * 1000)) & "MB", MQTT_QOS1
        DebuggingLog.DebugLog "[TrafficManager] MQTT test messages published"
    Else
        DebuggingLog.DebugLog "[TrafficManager] MQTT not available for testing"
    End If
    Exit Sub
ErrHandler:
    DebuggingLog.DebugLog "[TrafficManager] TestMQTTPublish error: " & Err.description
End Sub

Public Sub TestMQTTSubscribe()
    If mqttEnabled And IsConnected() Then
        SubscribeMQTTAdvanced "commands/vba/#", MQTT_QOS2
        SubscribeMQTTAdvanced "alerts/+/critical", MQTT_QOS1
        SubscribeMQTTAdvanced "data/realtime", MQTT_QOS0
        DebuggingLog.DebugLog "[TrafficManager] Additional MQTT subscriptions added"
    Else
        DebuggingLog.DebugLog "[TrafficManager] MQTT not available for subscription"
    End If
End Sub

'***************************************************************
' UDP Server Helpers
'***************************************************************
Public Sub StartUDPServer(ByVal port As Long)
    If Not InitializeWinsock() Then Exit Sub
    
    UdpPort = port
    udpSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If udpSocket = INVALID_SOCKET Then
        DebuggingLog.DebugLog "[TransmissionServer] UDP socket creation failed"
        Exit Sub
    End If
    
    Dim addr As SOCKADDR_IN
    addr.sin_family = AF_INET
    addr.sin_port = htons(UdpPort)
    addr.sin_addr = inet_addr("0.0.0.0")
    
    If bind(udpSocket, addr, LenB(addr)) = SOCKET_ERROR Then
        DebuggingLog.DebugLog "[TransmissionServer] UDP bind failed"
        Exit Sub
    End If
    
    Dim arg As Long: arg = 1
    ioctlsocket udpSocket, FIONBIO, arg
    
    udpLoopEnabled = True
    InitUDPQueue
    
    DebuggingLog.DebugLog "[TransmissionServer] UDP server started on port " & UdpPort
End Sub

Public Sub ClearSocketLockup()
    On Error Resume Next
    DebuggingLog.DebugLog "[TrafficManager] Clearing socket lockup..."
    TransmissionServer.ClearSocketLockup
    DebuggingLog.DebugLog "[TrafficManager] Socket lockup cleared"
End Sub


