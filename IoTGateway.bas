Option Explicit

'=========================================
' IoT Device Control Module (Full)
'=========================================

' Windows API Declarations
Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#If VBA7 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' Windows Socket Types
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
End Type

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Long
    sin_addr As Long
    sin_zero As String * 8
End Type

' Constants
Private Const AF_INET = 2
Private Const SOCK_STREAM = 1
Private Const IPPROTO_TCP = 6
Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1
Private Const INADDR_ANY As Long = 0
Private Const SOL_SOCKET = &HFFFF&
Private Const MAX_IOT_DEVICES = 100
Private Const TRAFFIC_LOG_PATH = "C:\TrafficLogs\"
Public Const IOT_INTERVAL_MS As Long = 500  ' Process every 0.5s

' Global Variables
Public iotSocket As Long
Public iotRunning As Boolean
Public iotClients() As Long
Public iotCount As Long
Public trafficSensors() As TrafficSensor
Public sensorCount As Long
Public totalVehicles As Long
Public lastTrafficUpdate As Date
Public serverStartTime As Date

' Enhanced IoT Device Type
Type IoTDevice
    DeviceID As String
    deviceType As String        ' "SENSOR", "ACTUATOR", "GATEWAY", "SMART_BULB"
    location As String
    ipAddress As String         ' For SMART_BULB, store the bulb's IP
    LastSeen As Date
    status As String            ' "ONLINE", "OFFLINE", "ERROR"
    batteryLevel As Integer
    FirmwareVersion As String
    socket As Long
    DataCount As Long
    ErrorCount As Long
    bulbState As String         ' "ON", "OFF", "UNKNOWN" for smart bulbs
    bulbBrightness As Integer   ' 0-100 for smart bulbs
    bulbColor As String         ' Optional: RGB or hue value
End Type

' IoT Message Queue Type
Type IoTMessage
    timestamp As Date
    DeviceID As String
    MessageType As String
    data As String
    priority As Integer        ' 1=High, 2=Normal, 3=Low
    Processed As Boolean
End Type

' Traffic Sensor Type
Type TrafficSensor
    SensorID As String
    vehicleCount As Long
    AverageSpeed As Double
    lastUpdate As Date
    status As String
End Type

' Enhanced Global Variables
Public IoTDevices() As IoTDevice
Public IoTDeviceCount As Long
Public iotMessageQueue() As IoTMessage
Public iotMessageQueueSize As Long
Public iotDataLogPath As String

Private iotLastTick As Long
Public iotLoopRunning As Boolean
Private nextIoTLoop As Date

'=========================================
' IoT Device Management
'=========================================
Public Sub RegisterIoTDevice(ByVal DeviceID As String, ByVal deviceType As String, ByVal location As String, ByVal port As Long, ByVal ipAddress As String)
    On Error GoTo ErrorHandler
    Dim i As Long
    ' Check if device already exists
    For i = 0 To IoTDeviceCount - 1
        If IoTDevices(i).DeviceID = DeviceID Then
            DebugLog "[IoT] Device already registered: " & DeviceID
            Exit Sub
        End If
    Next
    ' Resize array if needed
    If IoTDeviceCount = 0 Then
        ReDim IoTDevices(0 To MAX_IOT_DEVICES - 1)
    ElseIf IoTDeviceCount >= UBound(IoTDevices) + 1 Then
        ReDim Preserve IoTDevices(0 To UBound(IoTDevices) + MAX_IOT_DEVICES)
    End If
    ' Add new device
    With IoTDevices(IoTDeviceCount)
        .DeviceID = DeviceID
        .deviceType = UCase(deviceType)
        .location = location
        .ipAddress = ipAddress
        .LastSeen = Now
        .status = "OFFLINE"
        .batteryLevel = 100
        .FirmwareVersion = "1.0"
        .socket = INVALID_SOCKET
        .DataCount = 0
        .ErrorCount = 0
        .bulbState = "UNKNOWN"
        .bulbBrightness = 0
        .bulbColor = ""
    End With
    IoTDeviceCount = IoTDeviceCount + 1
    DebugLog "[IoT] Registered device: " & DeviceID & " (" & deviceType & ") at " & ipAddress
    Exit Sub
    
ErrorHandler:
    DebugLog "[IoT] Error in RegisterIoTDevice: " & Err.description
End Sub

Public Sub InitIoTDevices()
    On Error GoTo ErrorHandler
    ReDim IoTDevices(0 To MAX_IOT_DEVICES - 1)
    IoTDeviceCount = 0
    ReDim iotMessageQueue(0 To MAX_IOT_DEVICES - 1)
    iotMessageQueueSize = 0
    iotDataLogPath = TRAFFIC_LOG_PATH
    If Len(Dir(TRAFFIC_LOG_PATH, vbDirectory)) = 0 Then
        MkDir TRAFFIC_LOG_PATH
    End If
    DebugLog "[IoT] Initialized IoT devices"
    Exit Sub
    
ErrorHandler:
    DebugLog "[IoT] Error in InitIoTDevices: " & Err.description
End Sub

' Remove outdated Collection-based functions
' Public Sub AddIoTDevice(DeviceName As String, ip As String, Port As Long, Optional Protocol As String = "TCP")
' Public Sub RemoveIoTDevice(DeviceName As String)

'=========================================
' TCP Command Helper
'=========================================
Public Sub SendTCPCommand(deviceName As String, command() As Byte)
    On Error GoTo ErrorHandler
    Dim i As Long, dev As IoTDevice
    For i = 0 To IoTDeviceCount - 1
        If IoTDevices(i).DeviceID = deviceName Then
            dev = IoTDevices(i)
            Exit For
        End If
    Next
    If i = IoTDeviceCount Then
        DebugLog "[IoT] Device not found: " & deviceName
        Exit Sub
    End If
    
    Dim tcp As Object
    Set tcp = CreateObject("MSWinsock.Winsock")
    tcp.RemoteHost = dev.ipAddress
    tcp.remotePort = dev.socket
    
    tcp.connect
    Do While tcp.State <> 7 ' sckConnected
        DoEvents
        If DateDiff("s", Now, dev.LastSeen) > 10 Then
            DebugLog "[IoT] Connection timeout for " & deviceName
            Exit Sub
        End If
    Loop
    
    tcp.SendData command
    dev.LastSeen = Now
    dev.DataCount = dev.DataCount + 1
    dev.status = "ONLINE"
    
    tcp.Close
    DebugLog "[IoT] Sent TCP command to " & deviceName
    Exit Sub
    
ErrorHandler:
    dev.ErrorCount = dev.ErrorCount + 1
    dev.status = "ERROR"
    DebugLog "[IoT] Error sending TCP command to " & deviceName & ": " & Err.description
End Sub

'=========================================
' String / Byte Helpers
'=========================================
Public Function StringToBytesUTF8(str As String) As Byte()
    On Error GoTo ErrorHandler
    StringToBytesUTF8 = StrConv(str, vbFromUnicode)
    Exit Function
    
ErrorHandler:
    DebugLog "[IoT] Error in StringToBytesUTF8: " & Err.description
End Function

Public Function BytesToStringUTF8(ByRef data() As Byte) As String
    On Error GoTo ErrorHandler
    BytesToStringUTF8 = StrConv(data, vbUnicode)
    Exit Function
    
ErrorHandler:
    DebugLog "[IoT] Error in BytesToStringUTF8: " & Err.description
End Function

'=========================================
' Process messages in the IoT gateway queue
'=========================================
Public Sub IoTProcessGateway()
    On Error GoTo ErrorHandler
    Dim i As Long, devIndex As Long
    Dim msg As IoTMessage, dev As IoTDevice
    
    For i = LBound(iotMessageQueue) To UBound(iotMessageQueue)
        msg = iotMessageQueue(i)
        If Not msg.Processed Then
            devIndex = -1
            For devIndex = 0 To IoTDeviceCount - 1
                If IoTDevices(devIndex).DeviceID = msg.DeviceID Then Exit For
            Next
            If devIndex >= 0 And devIndex < IoTDeviceCount Then
                dev = IoTDevices(devIndex)
                Select Case msg.MessageType
                    Case "COMMAND"
                        If dev.deviceType = "SMART_BULB" And dev.socket <> INVALID_SOCKET Then
                            Call KasaSendCommand(dev, msg.data)
                        End If
                    Case "SENSOR_UPDATE"
                        dev.DataCount = dev.DataCount + 1
                        DebugLog "[IoT] Sensor " & dev.DeviceID & " updated with " & msg.data
                    Case Else
                        DebugLog "[IoT] Unhandled IoT message type: " & msg.MessageType
                End Select
                iotMessageQueue(i).Processed = True
            End If
        End If
    Next
    Exit Sub
    
ErrorHandler:
    DebugLog "[IoT] Error in IoTProcessGateway: " & Err.description
End Sub

'=========================================
' Kasa Smart Bulb Command (Placeholder)
'=========================================
Private Sub KasaSendCommand(dev As IoTDevice, ByVal command As String)
    On Error GoTo ErrorHandler
    Dim cmdBytes() As Byte
    cmdBytes = StringToBytesUTF8(command)
    Call SendTCPCommand(dev.DeviceID, cmdBytes)
    dev.LastSeen = Now
    dev.status = "ONLINE"
    DebugLog "[IoT] Sent Kasa command to " & dev.DeviceID & ": " & command
    Exit Sub
    
ErrorHandler:
    dev.ErrorCount = dev.ErrorCount + 1
    dev.status = "ERROR"
    DebugLog "[IoT] Error in KasaSendCommand: " & Err.description
End Sub

'***************************************************************
' COMPATIBILITY WRAPPERS FOR MAIN SERVER
'***************************************************************

' Returns True if IoT server is running
Public Function GetIoTRunning() As Boolean
    GetIoTRunning = iotRunning
End Function

' Starts the IoT server on the specified port
Public Sub StartIoTServer(Optional ByVal portNum As Long = 6000)
    If Not iotRunning Then
        ' Initialize WSA, devices, etc.
        InitIoTDevices
        iotRunning = True
        serverStartTime = Now
        DebugLog "[IoT] IoT server started on port " & portNum
        ' Optional: start a background loop or call ProcessIoTServer periodically
    Else
        DebugLog "[IoT] StartIoTServer called but IoT server is already running"
    End If
End Sub

' Stops the IoT server
Public Sub StopIoTServer()
    If iotRunning Then
        iotRunning = False
        ' Close all sockets
        Dim i As Long
        For i = LBound(iotClients) To UBound(iotClients)
            If iotClients(i) <> INVALID_SOCKET Then
                closesocket iotClients(i)
            End If
        Next i
        DebugLog "[IoT] IoT server stopped"
    End If
End Sub

Public Sub ProcessIoTServer()
    If Not iotRunning Then Exit Sub
    
    ' Process queued messages
    IoTProcessGateway
    
    ' Optional: update devices (simulate sensors)
    Dim i As Long
    For i = 0 To IoTDeviceCount - 1
        If IoTDevices(i).deviceType = "SENSOR" Then
            totalVehicles = totalVehicles + 1
        End If
    Next i
End Sub

' Compatibility Wrappers for TrafficManager
' Allows TrafficManager to call IoTGateway.GetRunning,
' IoTGateway.StartServer, and IoTGateway.StopServer

Public Function GetRunning() As Boolean
    GetRunning = GetIoTRunning()
End Function

Public Sub StartServer()
    StartIoTServer
End Sub

Public Sub StopServer()
    StopIoTServer
End Sub

' === IoTGateway: GetStatsJSON ===
Public Function GetStatsJSON() As String
    On Error GoTo ErrHandler
    
    If Not GetRunning() Then
        GetStatsJSON = "{}"
        Exit Function
    End If
    
    ' Example stats — you can extend these with real counters
    Dim connectedClients As Long
    Dim messagesPending As Long
    
    connectedClients = IoTGateway_GetConnectedClients()   ' implement in IoTGateway
    messagesPending = IoTGateway_GetPendingMessages()     ' implement in IoTGateway
    
    GetStatsJSON = "{""ConnectedClients"":" & connectedClients & _
                   ",""PendingMessages"":" & messagesPending & "}"
    Exit Function

ErrHandler:
    GetStatsJSON = "{}"
End Function


