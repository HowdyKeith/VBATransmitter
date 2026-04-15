Option Explicit

' ================================================================
' Main.bas - CENTRAL ORCHESTRATOR
' Version: 1.1 - Better stability, config, and polling
' ================================================================

Public Sub StartAllSystems()
    On Error GoTo ErrorHandler
    
    LoadDefaultConfig
    DebuggingLog.DebugLog "=== VBA TRANSMITTER STARTING ===", "INFO"
    
    If Not WinsockInit() Then
        MsgBox "Winsock failed to initialize!", vbCritical
        Exit Sub
    End If
    
    If Config.AutoStartAll Then
        StartGateway Config.GatewayPort, Config.httpPort, 0, Config.UdpPort, 0, Config.chatPort, 0, Config.mqttPort
        StartHttpServer Config.httpPort
        If Config.EnableMQTT Then MQTT_Connect "127.0.0.1", "VBAClient_" & format(Now, "hhnnss")
         If Config.EnableFTP Then StartFTPServer Config.FtpPort
         StartUDPServer Config.UdpPort
    End If
    
    StartPollingTimer
    DebuggingLog.DebugLog "All systems started successfully | " & GetFullStatus, "INFO"
    
    Exit Sub
ErrorHandler:
    DebuggingLog.DebugLog "StartAllSystems Error: " & Err.description, "ERROR"
End Sub

Public Sub ShutdownAllSystems()
    On Error Resume Next
    DebuggingLog.DebugLog "=== SHUTTING DOWN ===", "INFO"
    
    StopHttpServer
    ShutdownGateway
    ' ShutdownFTPServer
    WinsockCleanup
    
    DebuggingLog.DebugLog "Shutdown complete", "INFO"
End Sub

Private Sub StartPollingTimer()
    Application.OnTime Now + TimeSerial(0, 0, Config.PollIntervalSeconds), "PollAllServers"
End Sub

Public Sub PollAllServers()
    On Error Resume Next
    DoEvents                                      ' Prevent Excel freeze
    
    ProcessHttpServer
    ProcessGateway
    If Config.EnableMQTT Then ProcessMQTT
    
    StartPollingTimer                             ' Schedule next run
End Sub

Public Sub StartTrafficDemos()
    StartAllSystems
End Sub

Public Sub EmergencyStop()
    ShutdownAllSystems
End Sub
