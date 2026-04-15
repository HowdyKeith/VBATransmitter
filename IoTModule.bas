Option Explicit

' ================================================================
' IoTModule.bas
' Version: 1.0 - Kasa (TP-Link) + Govee integration
' ================================================================

Public Enum IoTDeviceType
    Kasa = 1
    govee = 2
    Generic = 99
End Enum

Public Type IoTDevice
    DeviceID As String
    deviceType As IoTDeviceType
    Name As String
    ip As String
    port As Long
    IsOn As Boolean
    Brightness As Integer          ' 0-100
    LastSeen As Date
    apiKey As String               ' For Govee / cloud devices
End Type

Public IoTDevices() As IoTDevice
Public IoTDeviceCount As Long

Private Sub LogIoT(message As String)
    DebuggingLog.DebugLog "[IoT] " & message, "INFO"
End Sub

' ====================== INITIALIZATION ======================
Public Sub IoT_Init()
    ReDim IoTDevices(0 To 49)
    IoTDeviceCount = 0
    LogIoT "IoT Module initialized"
End Sub

Public Sub IoT_AddDevice(deviceType As IoTDeviceType, deviceName As String, ipAddress As String, Optional apiKey As String = "")
    If IoTDeviceCount > UBound(IoTDevices) Then Exit Sub
    
    With IoTDevices(IoTDeviceCount)
        .DeviceID = "DEV" & format(Now, "yyyymmddhhnnss")
        .deviceType = deviceType
        .Name = deviceName
        .ip = ipAddress
        .port = IIf(deviceType = Kasa, 9999, 80)
        .IsOn = False
        .Brightness = 100
        .apiKey = apiKey
        .LastSeen = Now
    End With
    
    IoTDeviceCount = IoTDeviceCount + 1
    LogIoT "Added device: " & deviceName & " (" & ipAddress & ")"
End Sub

' ====================== KASA (TP-Link) CONTROL ======================
Public Sub Kasa_Toggle(deviceIndex As Long)
    If deviceIndex < 0 Or deviceIndex >= IoTDeviceCount Then Exit Sub
    If IoTDevices(deviceIndex).deviceType <> Kasa Then Exit Sub
    
    Dim cmd As String
    cmd = "{""system"":{""set_relay_state"":{""state"":" & IIf(IoTDevices(deviceIndex).IsOn, 0, 1) & "}}}"
    
    Dim encrypted As String
    encrypted = Kasa_Encrypt(cmd)
    
    Dim sock As LongPtr
    sock = CreateTCPSocket()
    SetNonBlocking sock
    
    Dim addr As SOCKADDR_IN
    With addr
        .sin_family = AF_INET
        .sin_port = htons(IoTDevices(deviceIndex).port)
        .sin_addr = inet_addr(IoTDevices(deviceIndex).ip)
    End With
    
    If connect(sock, addr, LenB(addr)) = 0 Or WSAGetLastError = WSAEWOULDBLOCK Then
        Dim b() As Byte: b = StrConv(encrypted, vbFromUnicode)
        send sock, b(0), UBound(b) + 1, 0
        IoTDevices(deviceIndex).IsOn = Not IoTDevices(deviceIndex).IsOn
        LogIoT "Toggled " & IoTDevices(deviceIndex).Name
    End If
    closesocket sock
End Sub

Private Function Kasa_Encrypt(cmd As String) As String
    Dim i As Long, result As String
    result = Chr(0) & Chr(0) & Chr(0) & Chr(Len(cmd))
    For i = 1 To Len(cmd)
        result = result & Chr(Asc(Mid(cmd, i, 1)) Xor &HAB)
    Next i
    Kasa_Encrypt = result
End Function

' ====================== GOVEE CONTROL (HTTP) ======================
Public Sub Govee_Toggle(deviceIndex As Long)
    If IoTDevices(deviceIndex).deviceType <> govee Then Exit Sub
    If IoTDevices(deviceIndex).apiKey = "" Then Exit Sub
    
    Dim url As String
    url = "https://developer-api.govee.com/v1/devices/control"
    
    Dim jsonBody As String
    jsonBody = "{""device"":""" & IoTDevices(deviceIndex).DeviceID & """,""model"":""H5083"",""cmd"":{""name"":""turn"",""value"":""" & IIf(IoTDevices(deviceIndex).IsOn, "off", "on") & """}}"
    
    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "POST", url, False
    req.SetRequestHeader "Govee-API-Key", IoTDevices(deviceIndex).apiKey
    req.SetRequestHeader "Content-Type", "application/json"
    req.send jsonBody
    
    IoTDevices(deviceIndex).IsOn = Not IoTDevices(deviceIndex).IsOn
    LogIoT "Govee toggled: " & IoTDevices(deviceIndex).Name
End Sub

' ====================== UNIFIED CONTROL ======================
Public Sub IoT_Toggle(deviceIndex As Long)
    Select Case IoTDevices(deviceIndex).deviceType
        Case Kasa:  Kasa_Toggle deviceIndex
        Case govee: Govee_Toggle deviceIndex
    End Select
End Sub

Public Sub IoT_SetBrightness(deviceIndex As Long, level As Integer)
    If level < 0 Then level = 0
    If level > 100 Then level = 100
    IoTDevices(deviceIndex).Brightness = level
    LogIoT "Brightness set to " & level & "% for " & IoTDevices(deviceIndex).Name
    ' Extend with actual Kasa/Govee commands as needed
End Sub

' ====================== DASHBOARD INTEGRATION ======================
Public Function IoT_GetStatusHTML() As String
    Dim html As String, i As Long
    html = "<h2>IoT Devices</h2><table border='1'><tr><th>Name</th><th>Type</th><th>IP</th><th>State</th></tr>"
    
    For i = 0 To IoTDeviceCount - 1
        html = html & "<tr><td>" & IoTDevices(i).Name & "</td>" & _
                      "<td>" & IIf(IoTDevices(i).deviceType = Kasa, "Kasa", "Govee") & "</td>" & _
                      "<td>" & IoTDevices(i).ip & "</td>" & _
                      "<td>" & IIf(IoTDevices(i).IsOn, "ON", "OFF") & "</td></tr>"
    Next i
    html = html & "</table>"
    IoT_GetStatusHTML = html
End Function

