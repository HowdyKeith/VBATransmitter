Option Explicit

' ================================================================
' Config.bas
' Version: 1.0 - Central configuration for the entire project
' ================================================================

Public Type ProjectConfig
    ' Server Ports
    GatewayPort As Long
    httpPort As Long
    mqttPort As Long
    UdpPort As Long
    FtpPort As Long
    chatPort As Long
    
    ' Settings
    DebugMode As Boolean
    AutoStartAll As Boolean
    EnableWebSocket As Boolean
    EnableMQTT As Boolean
    EnableFTP As Boolean
    
    ' Timeouts
    ClientTimeoutSeconds As Long
    PollIntervalSeconds As Long
End Type

Public Config As ProjectConfig

Public Sub LoadDefaultConfig()
    With Config
        .GatewayPort = 5000
        .httpPort = 8080
        .mqttPort = 1883
        .UdpPort = 9090
        .FtpPort = 2121          ' Use non-privileged port
        .chatPort = 8090
        
        .DebugMode = True
        .AutoStartAll = True
        .EnableWebSocket = True
        .EnableMQTT = True
        .EnableFTP = False
        
        .ClientTimeoutSeconds = 300
        .PollIntervalSeconds = 1
    End With
    DebuggingLog.DebugLog "Default configuration loaded", "INFO"
End Sub

Public Function GetFullStatus() As String
    Dim s As String
    s = "VBATransmitter Configuration:" & vbCrLf & _
        "HTTP: " & Config.httpPort & " | MQTT: " & Config.mqttPort & " | Gateway: " & Config.GatewayPort
    GetFullStatus = s
End Function

