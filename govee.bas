Option Explicit

'***************************************************************
' Enhanced Govee Module
' Purpose: Advanced Govee device discovery, control, and monitoring
' Features: Automatic discovery, scheduling, grouping, presets, monitoring
' Dependencies: UDP module, JSONParser, advanced web interface
'***************************************************************

Private deviceDict As Object  ' Dictionary of MAC -> Device Info Dictionary
Private Const MULTICAST_ADDR As String = "239.255.255.250"
Private Const DISCOVERY_PORT As Long = 4001
Private Const CONTROL_PORT As Long = 4003
Private Const LOCAL_BIND_PORT As Long = 4002
Private Const DISCOVERY_TIMEOUT_MS As Long = 5000
Private Const STATUS_TIMEOUT_MS As Long = 2000
Private Const SCAN_MESSAGE As String = "{""msg"":{""cmd"":""scan"",""data"":{""account_topic"":""reserve""}}}"
Private Const STATUS_MESSAGE As String = "{""msg"":{""cmd"":""status"",""data"":{}}}"
Private Const SHEET_NAME As String = "Govee"
Private Const PRESETS_SHEET As String = "GoveePresets"
Private Const GROUPS_SHEET As String = "GoveeGroups"
Private Const SCHEDULES_SHEET As String = "GoveeSchedules"
Private Const HEADER_ROW As Long = 1
Private Const AUTO_REFRESH_INTERVAL As Long = 300000 ' 5 minutes

' Enhanced column definitions
Private Enum GoveeColumns
    Col_MAC = 1
    Col_IP = 2
    Col_Model = 3
    Col_Name = 4
    Col_Room = 5
    Col_Group = 6
    Col_BleHardVer = 7
    Col_BleSoftVer = 8
    Col_WifiHardVer = 9
    Col_WifiSoftVer = 10
    Col_LastSeen = 11
    Col_Power = 12
    Col_Brightness = 13
    Col_Color = 14
    Col_Temperature = 15
    Col_Scene = 16
    Col_IsOnline = 17
    Col_ResponseTime = 18
End Enum

' --- Module Variables ---
Private goveeSocket As Long
Private presetDict As Object
Private groupDict As Object
Private scheduleDict As Object
Public isMonitoring As Boolean
Private lastAutoRefresh As Long
Private deviceStats As Object
' --- Enhanced Device Info Structure ---
' mac, ip, model, name, room, group, versions, lastSeen, power, brightness,
' color, temperature, scene, isOnline, responseTime, capabilities

' --- Get Routes ---
Public Function GetRoutes() As Collection
    Dim routes As New Collection
    ' Main Govee control page
    routes.Add Array("/govee", "GenerateGoveePage"), "/govee"
    ' Refresh devices
    routes.Add Array("/govee/refresh", "HandleGoveeRequest:refresh"), "/govee/refresh"
    ' Control device via command
    routes.Add Array("/govee/control", "HandleGoveeRequest:control"), "/govee/control"
    ' Set device name
    routes.Add Array("/govee/setname", "HandleGoveeRequest:setname"), "/govee/setname"
    ' Export devices
    routes.Add Array("/govee/export", "HandleGoveeRequest:export"), "/govee/export"
    ' Discover devices
    routes.Add Array("/govee/discover", "HandleGoveeRequest:discover"), "/govee/discover"
    ' API stats
    routes.Add Array("/api/govee/stats", "GetGoveeStatsJSON"), "/api/govee/stats"
    ' Device actions
    routes.Add Array("/govee/device", "HandleGoveeRequest:device"), "/govee/device"
    ' Group actions
    routes.Add Array("/govee/group", "HandleGoveeRequest:group"), "/govee/group"
    ' Preset actions
    routes.Add Array("/govee/preset", "HandleGoveeRequest:preset"), "/govee/preset"
    ' Schedule actions
    routes.Add Array("/govee/schedule", "HandleGoveeRequest:schedule"), "/govee/schedule"
    ' Export presets
    routes.Add Array("/govee/export/presets", "HandleEnhancedGoveeRequest:export/presets"), "/govee/export/presets"
    ' Export schedules
    routes.Add Array("/govee/export/schedules", "HandleEnhancedGoveeRequest:export/schedules"), "/govee/export/schedules"
    ' Reset all data
    routes.Add Array("/govee/reset", "HandleEnhancedGoveeRequest:reset"), "/govee/reset"
    Set GetRoutes = routes
End Function
' --- Handle App Request for AppLaunch Compatibility ---
Public Sub HandleAppRequest(ByVal method As String, ByVal path As String, ByVal body As String, ByRef response As String)
    On Error GoTo ErrorHandler
    
    Dim action As String
    Dim query As String
    ' Extract query from path if present
    If InStr(path, "?") > 0 Then
        action = Left(path, InStr(path, "?") - 1)
        query = Mid(path, InStr(path, "?") + 1)
    Else
        action = path
        query = ""
    End If
    
    Select Case LCase(action)
Case "GenerateStatusPage"
        ' If request comes from AJAX (home page), return snippet
        If InStr(path, "snippet") > 0 Then
            response = GenerateStatusSnippet
        Else
            response = GenerateStatusPage
        End If
        Case "/govee", "/govee/"
            response = GenerateGoveePage
        Case "/govee/refresh"
            DiscoverDevicesAdvanced
            response = GenerateGoveePage
        Case "/govee/control"
            If LCase(method) = "post" Then
                Dim params As Object
                Set params = ParseQueryOrBody(query, body)
                Dim identifier As String, command As String
                identifier = params("identifier")
                command = params("command")
                If SendCommand(identifier, command) Then
                    response = GenerateGoveePage
                Else
                    response = GenerateErrorPage("Failed to send command to " & identifier)
                End If
            Else
                response = GenerateErrorPage("Control requires POST method")
            End If
        Case "/govee/setname"
            If LCase(method) = "post" Then
                Dim setParams As Object
                Set setParams = ParseQueryOrBody(query, body)
                If UpdateDeviceName(setParams("identifier"), setParams("name")) Then
                    response = GenerateGoveePage
                Else
                    response = GenerateErrorPage("Failed to update device name")
                End If
            Else
                response = GenerateErrorPage("Setname requires POST method")
            End If
        Case "/govee/export"
            response = ExportDevices
        Case "/govee/discover"
            DiscoverDevicesAdvanced
            response = GenerateGoveePage
        Case "/api/govee/stats"
            response = GetGoveeStatsJSON
        Case "/govee/device"
            response = HandleEnhancedGoveeRequest("/govee/device", query)
        Case "/govee/group"
            response = HandleEnhancedGoveeRequest("/govee/group", query)
        Case "/govee/preset"
            response = HandleEnhancedGoveeRequest("/govee/preset", query)
        Case "/govee/schedule"
            response = HandleEnhancedGoveeRequest("/govee/schedule", query)
        Case Else
            response = GenerateErrorPage("Invalid Govee path: " & path)
    End Select
    
    DebuggingLog.DebugLog "Handled Govee request: " & path & " (method: " & method & ")"
    Exit Sub

ErrorHandler:
    DebuggingLog.DebugLog "Error in HandleAppRequest: " & Err.description
    response = GenerateErrorPage("Govee error: " & Err.description)
End Sub

' --- Wrapper for GenerateEnhancedGoveePage ---
Public Function GenerateGoveePage() As String
    GenerateGoveePage = GenerateEnhancedGoveePage
End Function

' --- Parse Query or Body ---
Private Function ParseQueryOrBody(ByVal query As String, ByVal body As String) As Object
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    
    ' Parse query string (if any)
    If query <> "" Then
        Dim pairs() As String, pair() As String, i As Long
        pairs = Split(query, "&")
        For i = LBound(pairs) To UBound(pairs)
            pair = Split(pairs(i), "=")
            If UBound(pair) >= 1 Then
                params(LCase(pair(0))) = DecodeURL(pair(1))
            End If
        Next i
    End If
    
    ' Parse body (for POST, assume URL-encoded key-value pairs)
    If body <> "" Then
        Dim bodyPairs() As String
        bodyPairs = Split(body, "&")
        For i = LBound(bodyPairs) To UBound(bodyPairs)
            pair = Split(bodyPairs(i), "=")
            If UBound(pair) >= 1 Then
                params(LCase(pair(0))) = DecodeURL(pair(1))
            End If
        Next i
    End If
    
    Set ParseQueryOrBody = params
End Function

' --- Decode URL ---
Private Function DecodeURL(ByVal str As String) As String
    str = Replace(str, "+", " ")
    Dim i As Long, hexVal As String
    i = 1
    Do While i <= Len(str)
        If Mid(str, i, 1) = "%" And i <= Len(str) - 2 Then
            hexVal = Mid(str, i + 1, 2)
            On Error Resume Next
            Mid(str, i, 3) = Chr(CLng("&H" & hexVal))
            On Error GoTo 0
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    DecodeURL = str
End Function

' --- Initialize Enhanced Govee Module ---
Public Sub InitializeGovee()
    On Error GoTo ErrorHandler
    
    ' Create or activate sheets
    CreateGoveeSheets
    
    ' Initialize dictionaries
    Set deviceDict = CreateObject("Scripting.Dictionary")
    Set presetDict = CreateObject("Scripting.Dictionary")
    Set groupDict = CreateObject("Scripting.Dictionary")
    Set scheduleDict = CreateObject("Scripting.Dictionary")
    Set deviceStats = CreateObject("Scripting.Dictionary")
    
    ' Load existing data
    LoadDataFromSheets
    
    ' Create UDP socket
    goveeSocket = UDP.CreateAdvancedUDPSocket(8192)
    If goveeSocket = UDP.INVALID_SOCKET Then
        DebuggingLog.DebugLog "Failed to create Govee UDP socket"
        Exit Sub
    End If
    
    ' Bind and configure socket
    If Not UDP.BindUDP(goveeSocket, LOCAL_BIND_PORT) Then
        DebuggingLog.DebugLog "Failed to bind Govee socket to port " & LOCAL_BIND_PORT
        UDP.CloseUDPSocket goveeSocket
        goveeSocket = UDP.INVALID_SOCKET
        Exit Sub
    End If
    
    UDP.SetUDPNonBlocking goveeSocket
    UDP.EnableUDPBroadcast goveeSocket
    
    ' Enable monitoring
    isMonitoring = True
    lastAutoRefresh = GetTickCount()
    
    DebuggingLog.DebugLog "Enhanced Govee module initialized with monitoring"
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error initializing Enhanced Govee: " & Err.description
End Sub

' --- Create Enhanced Govee Sheets ---
Private Sub CreateGoveeSheets()
    On Error Resume Next
    
    ' Main devices sheet
    Dim ws As Worksheet
    Set ws = CreateOrGetSheet(SHEET_NAME)
    SetupDeviceSheet ws
    
    ' Presets sheet
    Set ws = CreateOrGetSheet(PRESETS_SHEET)
    SetupPresetsSheet ws
    
    ' Groups sheet
    Set ws = CreateOrGetSheet(GROUPS_SHEET)
    SetupGroupsSheet ws
    
    ' Schedules sheet
    Set ws = CreateOrGetSheet(SCHEDULES_SHEET)
    SetupSchedulesSheet ws
End Sub

Private Function CreateOrGetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set CreateOrGetSheet = ThisWorkbook.Sheets(sheetName)
    If CreateOrGetSheet Is Nothing Then
        Set CreateOrGetSheet = ThisWorkbook.Sheets.Add
        CreateOrGetSheet.Name = sheetName
    End If
    On Error GoTo 0
End Function

Private Sub SetupDeviceSheet(ws As Worksheet)
    With ws
        .Cells(HEADER_ROW, Col_MAC) = "MAC"
        .Cells(HEADER_ROW, Col_IP) = "IP"
        .Cells(HEADER_ROW, Col_Model) = "Model"
        .Cells(HEADER_ROW, Col_Name) = "Name"
        .Cells(HEADER_ROW, Col_Room) = "Room"
        .Cells(HEADER_ROW, Col_Group) = "Group"
        .Cells(HEADER_ROW, Col_BleHardVer) = "BLE Hard Ver"
        .Cells(HEADER_ROW, Col_BleSoftVer) = "BLE Soft Ver"
        .Cells(HEADER_ROW, Col_WifiHardVer) = "WiFi Hard Ver"
        .Cells(HEADER_ROW, Col_WifiSoftVer) = "WiFi Soft Ver"
        .Cells(HEADER_ROW, Col_LastSeen) = "Last Seen"
        .Cells(HEADER_ROW, Col_Power) = "Power"
        .Cells(HEADER_ROW, Col_Brightness) = "Brightness"
        .Cells(HEADER_ROW, Col_Color) = "Color (RGB)"
        .Cells(HEADER_ROW, Col_Temperature) = "Color Temp"
        .Cells(HEADER_ROW, Col_Scene) = "Scene"
        .Cells(HEADER_ROW, Col_IsOnline) = "Online"
        .Cells(HEADER_ROW, Col_ResponseTime) = "Response (ms)"
        .Rows(HEADER_ROW).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

Private Sub SetupPresetsSheet(ws As Worksheet)
    With ws
        .Cells(1, 1) = "Preset Name"
        .Cells(1, 2) = "Description"
        .Cells(1, 3) = "Power"
        .Cells(1, 4) = "Brightness"
        .Cells(1, 5) = "Red"
        .Cells(1, 6) = "Green"
        .Cells(1, 7) = "Blue"
        .Cells(1, 8) = "Color Temperature"
        .Cells(1, 9) = "Scene Code"
        .Cells(1, 10) = "Created"
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

Private Sub SetupGroupsSheet(ws As Worksheet)
    With ws
        .Cells(1, 1) = "Group Name"
        .Cells(1, 2) = "Description"
        .Cells(1, 3) = "Device MACs (comma separated)"
        .Cells(1, 4) = "Created"
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

Private Sub SetupSchedulesSheet(ws As Worksheet)
    With ws
        .Cells(1, 1) = "Schedule Name"
        .Cells(1, 2) = "Device/Group"
        .Cells(1, 3) = "Action"
        .Cells(1, 4) = "Time"
        .Cells(1, 5) = "Days (Mon,Tue,Wed...)"
        .Cells(1, 6) = "Enabled"
        .Cells(1, 7) = "Last Run"
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' --- Enhanced Discovery with Capabilities Detection ---
Public Sub DiscoverDevicesAdvanced()
    On Error GoTo ErrorHandler
    
    DebuggingLog.DebugLog "Starting advanced Govee device discovery..."
    
    ' Clear old offline devices
    CleanupOfflineDevices
    
    ' Send enhanced discovery
    SendDiscoveryMessage
    
    ' Extended listening period with progress
    Dim startTime As Long
    startTime = GetTickCount()
    Dim devicesFound As Long
    devicesFound = deviceDict.count
    
    Do While GetTickCount() - startTime < DISCOVERY_TIMEOUT_MS
        ProcessDiscoveryResponses
        DoEvents
        Sleep 25
    Loop
    
    ' Probe each device for capabilities
    ProbeDeviceCapabilities
    
    ' Update all device statuses
    UpdateAllDeviceStatuses
    
    ' Save results
    SaveDataToSheets
    
    Dim newDevices As Long
    newDevices = deviceDict.count - devicesFound
    DebuggingLog.DebugLog "Discovery complete. Found " & newDevices & " new devices. Total: " & deviceDict.count
    
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error in DiscoverDevicesAdvanced: " & Err.description
End Sub

Private Sub SendDiscoveryMessage()
    ' Send to multicast
    UDP.SendUDP goveeSocket, SCAN_MESSAGE, MULTICAST_ADDR, DISCOVERY_PORT
    
    ' Also try broadcast
    UDP.BroadcastUDP goveeSocket, SCAN_MESSAGE, DISCOVERY_PORT
    
    ' Send to known device IPs
    Dim key As Variant
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        If device("ip") <> "" Then
            UDP.SendUDP goveeSocket, SCAN_MESSAGE, device("ip"), DISCOVERY_PORT
        End If
    Next key
End Sub

Private Sub ProcessDiscoveryResponses()
    Dim buffer(0 To 2047) As Byte
    Dim fromIP As String
    Dim fromPort As Long
    Dim bytesRecv As Long
    
    Do
        bytesRecv = UDP.RecvUDP(goveeSocket, buffer, fromIP, fromPort)
        If bytesRecv > 0 Then
            Dim response As String
            response = Left(StrConv(buffer, vbUnicode), bytesRecv)
            ProcessEnhancedDiscoveryResponse response, fromIP
        End If
    Loop While bytesRecv > 0
End Sub

Private Sub ProcessEnhancedDiscoveryResponse(ByVal response As String, ByVal fromIP As String)
    On Error GoTo ErrorHandler
    
    Dim parsed As Object
    Set parsed = JSONParser.ParseJSON(response)
    
    If Not parsed Is Nothing And parsed.exists("msg") Then
        Dim msg As Object
        Set msg = parsed("msg")
        
        If msg.exists("cmd") And msg("cmd") = "scan" And msg.exists("data") Then
            Dim data As Object
            Set data = msg("data")
            
            Dim mac As String
            mac = UCase(Replace(data("device"), ":", ""))
            
            Dim device As Object
            If deviceDict.exists(mac) Then
                Set device = deviceDict(mac)
            Else
                Set device = CreateObject("Scripting.Dictionary")
                device("mac") = mac
                device("name") = ""
                device("room") = ""
                device("group") = ""
                device("capabilities") = ""
            End If
            
            ' Update device info
            device("ip") = fromIP
            If data.exists("sku") Then device("model") = data("sku")
            If data.exists("bleVersionHard") Then device("bleVersionHard") = data("bleVersionHard")
            If data.exists("bleVersionSoft") Then device("bleVersionSoft") = data("bleVersionSoft")
            If data.exists("wifiVersionHard") Then device("wifiVersionHard") = data("wifiVersionHard")
            If data.exists("wifiVersionSoft") Then device("wifiVersionSoft") = data("wifiVersionSoft")
            device("lastSeen") = Now
            device("isOnline") = True
            
            ' Initialize missing fields
            If Not device.exists("power") Then device("power") = "Unknown"
            If Not device.exists("brightness") Then device("brightness") = "Unknown"
            If Not device.exists("color") Then device("color") = "Unknown"
            If Not device.exists("temperature") Then device("temperature") = "Unknown"
            If Not device.exists("scene") Then device("scene") = "Unknown"
            If Not device.exists("responseTime") Then device("responseTime") = 0
            
            deviceDict(mac) = device
            
            DebuggingLog.DebugLog "Enhanced discovery: " & mac & " at " & fromIP & " (" & device("model") & ")"
        End If
    End If
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error processing enhanced discovery response: " & Err.description
End Sub

' --- Capability Probing ---
Private Sub ProbeDeviceCapabilities()
    Dim key As Variant
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        
        If device("isOnline") Then
            ProbeDeviceCapability key
        End If
    Next key
End Sub

Private Sub ProbeDeviceCapability(ByVal mac As String)
    On Error GoTo ErrorHandler
    
    Dim device As Object
    Set device = deviceDict(mac)
    
    Dim capabilities As String
    capabilities = ""
    
    ' Test basic commands
    If TestCommand(device("ip"), "{""msg"":{""cmd"":""turn"",""data"":{""value"":1}}}") Then
        capabilities = capabilities & "power,"
    End If
    
    If TestCommand(device("ip"), "{""msg"":{""cmd"":""brightness"",""data"":{""value"":50}}}") Then
        capabilities = capabilities & "brightness,"
    End If
    
    If TestCommand(device("ip"), "{""msg"":{""cmd"":""color"",""data"":{""r"":255,""g"":255,""b"":255}}}") Then
        capabilities = capabilities & "color,"
    End If
    
    If TestCommand(device("ip"), "{""msg"":{""cmd"":""colorTem"",""data"":{""value"":3000}}}") Then
        capabilities = capabilities & "temperature,"
    End If
    
    If TestCommand(device("ip"), "{""msg"":{""cmd"":""mode"",""data"":{""value"":1}}}") Then
        capabilities = capabilities & "scenes,"
    End If
    
    device("capabilities") = Left(capabilities, Len(capabilities) - 1) ' Remove trailing comma
    
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error probing capabilities for " & mac & ": " & Err.description
End Sub

Private Function TestCommand(ByVal ip As String, ByVal command As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim startTime As Long
    startTime = GetTickCount()
    
    UDP.SendUDP goveeSocket, command, ip, CONTROL_PORT
    
    ' Wait for any response (don't care about content)
    Do While GetTickCount() - startTime < 1000 ' 1 second timeout
        Dim buffer(0 To 511) As Byte
        Dim fromIP As String
        Dim fromPort As Long
        Dim bytesRecv As Long
        
        bytesRecv = UDP.RecvUDP(goveeSocket, buffer, fromIP, fromPort)
        If bytesRecv > 0 And fromIP = ip Then
            TestCommand = True
            Exit Function
        End If
        DoEvents
        Sleep 10
    Loop
    
    TestCommand = False
    Exit Function
    
ErrorHandler:
    TestCommand = False
End Function

' --- Preset Management ---
Public Function CreatePreset(ByVal presetName As String, ByVal description As String, ByVal power As String, ByVal Brightness As Long, ByVal r As Long, ByVal g As Long, ByVal b As Long, Optional ByVal temperature As Long = 0, Optional ByVal scene As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    
    Dim preset As Object
    Set preset = CreateObject("Scripting.Dictionary")
    preset("name") = presetName
    preset("description") = description
    preset("power") = power
    preset("brightness") = Brightness
    preset("red") = r
    preset("green") = g
    preset("blue") = b
    preset("temperature") = temperature
    preset("scene") = scene
    preset("created") = Now
    
    presetDict(presetName) = preset
    SaveDataToSheets
    
    DebuggingLog.DebugLog "Created preset: " & presetName
    CreatePreset = True
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error creating preset: " & Err.description
    CreatePreset = False
End Function

Public Function ApplyPreset(ByVal identifier As String, ByVal presetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not presetDict.exists(presetName) Then
        DebuggingLog.DebugLog "Preset not found: " & presetName
        ApplyPreset = False
        Exit Function
    End If
    
    Dim preset As Object
    Set preset = presetDict(presetName)
    
    Dim success As Boolean
    success = True
    
    ' Apply preset settings
    If preset("power") = "On" Then
        success = success And TurnOn(identifier)
    ElseIf preset("power") = "Off" Then
        success = success And TurnOff(identifier)
    End If
    
    If preset("brightness") > 0 Then
        success = success And SetBrightness(identifier, preset("brightness"))
    End If
    
    If preset("red") >= 0 And preset("green") >= 0 And preset("blue") >= 0 Then
        success = success And SetColor(identifier, preset("red"), preset("green"), preset("blue"))
    End If
    
    If preset("temperature") > 0 Then
        success = success And SetColorTemperature(identifier, preset("temperature"))
    End If
    
    If preset("scene") > 0 Then
        success = success And SetScene(identifier, preset("scene"))
    End If
    
    ApplyPreset = success
    DebuggingLog.DebugLog "Applied preset " & presetName & " to " & identifier & ": " & success
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error applying preset: " & Err.description
    ApplyPreset = False
End Function

' --- Group Management ---
Public Function CreateGroup(ByVal groupName As String, ByVal description As String, ByVal deviceMACs As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim group As Object
    Set group = CreateObject("Scripting.Dictionary")
    group("name") = groupName
    group("description") = description
    group("devices") = deviceMACs
    group("created") = Now
    
    groupDict(groupName) = group
    SaveDataToSheets
    
    ' Update device group assignments
    Dim macs As Variant
    macs = Split(deviceMACs, ",")
    Dim i As Long
    For i = 0 To UBound(macs)
        Dim mac As String
        mac = Trim(UCase(macs(i)))
        If deviceDict.exists(mac) Then
            Dim device As Object
            Set device = deviceDict(mac)
            device("group") = groupName
        End If
    Next i
    
    DebuggingLog.DebugLog "Created group: " & groupName
    CreateGroup = True
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error creating group: " & Err.description
    CreateGroup = False
End Function

Public Function ControlGroup(ByVal groupName As String, ByVal action As String, Optional ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    If Not groupDict.exists(groupName) Then
        DebuggingLog.DebugLog "Group not found: " & groupName
        ControlGroup = False
        Exit Function
    End If
    
    Dim group As Object
    Set group = groupDict(groupName)
    
    Dim macs As Variant
    macs = Split(group("devices"), ",")
    
    Dim success As Boolean
    success = True
    
    Dim i As Long
    For i = 0 To UBound(macs)
        Dim mac As String
        mac = Trim(UCase(macs(i)))
        
        Select Case LCase(action)
            Case "on"
                success = success And TurnOn(mac)
            Case "off"
                success = success And TurnOff(mac)
            Case "brightness"
                success = success And SetBrightness(mac, value)
            Case "color"
                Dim rgb As Variant
                rgb = Split(value, ",")
                If UBound(rgb) >= 2 Then
                    success = success And SetColor(mac, rgb(0), rgb(1), rgb(2))
                End If
            Case "preset"
                success = success And ApplyPreset(mac, value)
        End Select
    Next i
    
    ControlGroup = success
    DebuggingLog.DebugLog "Group control " & action & " on " & groupName & ": " & success
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error controlling group: " & Err.description
    ControlGroup = False
End Function

' --- Schedule Management ---
Public Function CreateSchedule(ByVal scheduleName As String, ByVal deviceOrGroup As String, ByVal action As String, ByVal timeStr As String, ByVal days As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim schedule As Object
    Set schedule = CreateObject("Scripting.Dictionary")
    schedule("name") = scheduleName
    schedule("target") = deviceOrGroup
    schedule("action") = action
    schedule("time") = timeStr
    schedule("days") = days
    schedule("enabled") = True
    schedule("lastRun") = ""
    
    scheduleDict(scheduleName) = schedule
    SaveDataToSheets
    
    DebuggingLog.DebugLog "Created schedule: " & scheduleName
    CreateSchedule = True
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error creating schedule: " & Err.description
    CreateSchedule = False
End Function

Public Sub ProcessSchedules()
    On Error GoTo ErrorHandler
    
    Dim currentTime As String
    currentTime = format(Now, "hh:mm")
    Dim currentDay As String
    currentDay = format(Now, "ddd")
    
    Dim key As Variant
    For Each key In scheduleDict.Keys
        Dim schedule As Object
        Set schedule = scheduleDict(key)
        
        If schedule("enabled") And InStr(schedule("days"), currentDay) > 0 And schedule("time") = currentTime Then
            ' Check if already run today
            Dim lastRun As String
            lastRun = format(schedule("lastRun"), "yyyy-mm-dd")
            If lastRun <> format(Now, "yyyy-mm-dd") Then
                ExecuteSchedule schedule
                schedule("lastRun") = Now
            End If
        End If
    Next key
    
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error processing schedules: " & Err.description
End Sub

Private Sub ExecuteSchedule(ByRef schedule As Object)
    On Error GoTo ErrorHandler
    
    Dim target As String
    target = schedule("target")
    Dim action As String
    action = schedule("action")
    
    ' Determine if target is device or group
    If deviceDict.exists(target) Then
        ' It's a device MAC
        Select Case LCase(action)
            Case "on"
                TurnOn target
            Case "off"
                TurnOff target
            Case Else
                ' Try as preset
                ApplyPreset target, action
        End Select
    ElseIf groupDict.exists(target) Then
        ' It's a group
        If InStr(action, ":") > 0 Then
            Dim parts As Variant
            parts = Split(action, ":", 2)
            ControlGroup target, parts(0), parts(1)
        Else
            ControlGroup target, action
        End If
    End If
    
    DebuggingLog.DebugLog "Executed schedule: " & schedule("name") & " -> " & action & " on " & target
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error executing schedule: " & Err.description
End Sub

' --- Monitoring and Auto-refresh ---
Public Sub ProcessGoveeMonitoring()
    On Error GoTo ErrorHandler
    
    If Not isMonitoring Then Exit Sub
    
    ' Auto-refresh every 5 minutes
    If GetTickCount() - lastAutoRefresh > AUTO_REFRESH_INTERVAL Then
        DiscoverDevicesAdvanced
        lastAutoRefresh = GetTickCount()
    End If
    
    ' Process schedules every minute
    Static lastScheduleCheck As Long
    If GetTickCount() - lastScheduleCheck > 60000 Then ' 1 minute
        ProcessSchedules
        lastScheduleCheck = GetTickCount()
    End If
    
    ' Update device statistics
    UpdateDeviceStatistics
    
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "Error in Govee monitoring: " & Err.description
End Sub

Private Sub UpdateDeviceStatistics()
    Dim key As Variant
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        
        ' Update online status based on last seen
        Dim minutesSinceLastSeen As Long
        minutesSinceLastSeen = DateDiff("n", device("lastSeen"), Now)
        device("isOnline") = (minutesSinceLastSeen < 10) ' 10 minutes threshold
        
        ' Update stats
        If Not deviceStats.exists(key) Then
            Set deviceStats(key) = CreateObject("Scripting.Dictionary")
            deviceStats(key)("totalCommands") = 0
            deviceStats(key)("successfulCommands") = 0
            deviceStats(key)("lastCommand") = ""
        End If
    Next key
End Sub


' --- Enhanced Request Handler ---
Public Function HandleEnhancedGoveeRequest(ByVal path As String, ByVal query As String) As String
    On Error GoTo ErrorHandler
    
    Select Case True
        Case path = "/govee" Or path = "/govee/"
            HandleEnhancedGoveeRequest = GenerateEnhancedGoveePage()
        Case path = "/govee/discover"
            DiscoverDevicesAdvanced
            HandleEnhancedGoveeRequest = GenerateEnhancedGoveePage()
        Case InStr(path, "/govee/device") = 1
            HandleEnhancedGoveeRequest = HandleDeviceRequest(query)
        Case InStr(path, "/govee/group") = 1
            HandleEnhancedGoveeRequest = HandleGroupRequest(query)
        Case InStr(path, "/govee/preset") = 1
            HandleEnhancedGoveeRequest = HandlePresetRequest(query)
        Case InStr(path, "/govee/schedule") = 1
            HandleEnhancedGoveeRequest = HandleScheduleRequest(query)
        Case InStr(path, "/govee/export") = 1
            HandleEnhancedGoveeRequest = HandleExportRequest(path)
        Case path = "/govee/reset"
            ResetAllData
            HandleEnhancedGoveeRequest = GenerateEnhancedGoveePage()
        Case Else
            HandleEnhancedGoveeRequest = GenerateErrorPage("Invalid Govee path: " & path)
    End Select
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error handling enhanced Govee request: " & Err.description
    HandleEnhancedGoveeRequest = GenerateErrorPage("Error: " & Err.description)
End Function

Private Function HandleDeviceRequest(ByVal query As String) As String
    Dim params As Object
    Set params = ParseQueryString(query)
    
    If params.exists("mac") And params.exists("action") Then
        Dim mac As String, action As String
        mac = params("mac")
        action = params("action")
        
        Select Case LCase(action)
            Case "on": TurnOn mac
            Case "off": TurnOff mac
            Case "toggle": TogglePower mac
            Case "status": GetDeviceStatus mac
            Case "bright_up": AdjustBrightness mac, 10
            Case "bright_down": AdjustBrightness mac, -10
        End Select
    End If
    
    HandleDeviceRequest = GenerateEnhancedGoveePage()
End Function

Private Function HandleGroupRequest(ByVal query As String) As String
    Dim params As Object
    Set params = ParseQueryString(query)
    
    If params.exists("name") And params.exists("action") Then
        ControlGroup params("name"), params("action")
    End If
    
    HandleGroupRequest = GenerateEnhancedGoveePage()
End Function

Private Function HandlePresetRequest(ByVal query As String) As String
    Dim params As Object
    Set params = ParseQueryString(query)
    
    If params.exists("name") And params.exists("action") Then
        Dim presetName As String
        presetName = params("name")
        Select Case LCase(params("action"))
            Case "apply"
                If params.exists("target") Then
                    ApplyPreset params("target"), presetName
                End If
            Case "delete"
                If presetDict.exists(presetName) Then
                    presetDict.Remove presetName
                    SaveDataToSheets
                End If
        End Select
    End If
    
    HandlePresetRequest = GenerateEnhancedGoveePage()
End Function

Private Function HandleScheduleRequest(ByVal query As String) As String
    Dim params As Object
    Set params = ParseQueryString(query)
    
    If params.exists("name") And params.exists("action") Then
        Dim scheduleName As String
        scheduleName = params("name")
        Select Case LCase(params("action"))
            Case "toggle"
                If scheduleDict.exists(scheduleName) Then
                    Dim schedule As Object
                    Set schedule = scheduleDict(scheduleName)
                    schedule("enabled") = Not schedule("enabled")
                    SaveDataToSheets
                End If
            Case "delete"
                If scheduleDict.exists(scheduleName) Then
                    scheduleDict.Remove scheduleName
                    SaveDataToSheets
                End If
        End Select
    End If
    
    HandleScheduleRequest = GenerateEnhancedGoveePage()
End Function

Private Function HandleExportRequest(ByVal path As String) As String
    Dim exportType As String
    exportType = Mid(path, InStrRev(path, "/") + 1)
    
    Select Case LCase(exportType)
        Case "devices"
            HandleExportRequest = ExportDevices
        Case "presets"
            HandleExportRequest = ExportPresets
        Case "schedules"
            HandleExportRequest = ExportSchedules
        Case Else
            HandleExportRequest = GenerateErrorPage("Invalid export type: " & exportType)
    End Select
End Function

Private Function ParseQueryString(ByVal query As String) As Object
    Set ParseQueryString = CreateObject("Scripting.Dictionary")
    If Len(query) = 0 Then Exit Function
    
    Dim pairs As Variant
    pairs = Split(query, "&")
    Dim i As Long
    For i = 0 To UBound(pairs)
        Dim pair As Variant
        pair = Split(pairs(i), "=")
        If UBound(pair) >= 1 Then
            ParseQueryString(pair(0)) = DecodeURL(pair(1))
        End If
    Next i
End Function

Private Function GenerateErrorPage(ByVal message As String) As String
    GenerateErrorPage = "<html><body style='background:#000;color:#ff4444;font-family:monospace;padding:20px;'>"
    GenerateErrorPage = GenerateErrorPage & "<h1>Govee Control Error</h1>"
    GenerateErrorPage = GenerateErrorPage & "<p>" & message & "</p>"
    GenerateErrorPage = GenerateErrorPage & "<a href='/govee' style='color:#00ff88;'>Back to Govee Control</a>"
    GenerateErrorPage = GenerateErrorPage & "</body></html>"
End Function

' --- Data Management Functions ---
Private Sub LoadDataFromSheets()
    LoadDevicesFromSheet
    LoadPresetsFromSheet
    LoadGroupsFromSheet
    LoadSchedulesFromSheet
End Sub

Private Sub SaveDataToSheets()
    SaveDevicesToSheet
    SavePresetsToSheet
    SaveGroupsToSheet
    SaveSchedulesToSheet
End Sub

Private Sub LoadPresetsFromSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PRESETS_SHEET)
    If ws Is Nothing Then Exit Sub
    
    Dim row As Long
    row = 2 ' Skip header
    
    Do While ws.Cells(row, 1) <> ""
        Dim preset As Object
        Set preset = CreateObject("Scripting.Dictionary")
        preset("name") = ws.Cells(row, 1)
        preset("description") = ws.Cells(row, 2)
        preset("power") = ws.Cells(row, 3)
        preset("brightness") = ws.Cells(row, 4)
        preset("red") = ws.Cells(row, 5)
        preset("green") = ws.Cells(row, 6)
        preset("blue") = ws.Cells(row, 7)
        preset("temperature") = ws.Cells(row, 8)
        preset("scene") = ws.Cells(row, 9)
        preset("created") = ws.Cells(row, 10)
        
        presetDict(preset("name")) = preset
        row = row + 1
    Loop
End Sub

Private Sub SavePresetsToSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PRESETS_SHEET)
    If ws Is Nothing Then Exit Sub
    
    ' Clear existing data
    ws.Rows("2:" & ws.Rows.count).ClearContents
    
    Dim row As Long
    row = 2
    
    Dim key As Variant
    For Each key In presetDict.Keys
        Dim preset As Object
        Set preset = presetDict(key)
        
        ws.Cells(row, 1) = preset("name")
        ws.Cells(row, 2) = preset("description")
        ws.Cells(row, 3) = preset("power")
        ws.Cells(row, 4) = preset("brightness")
        ws.Cells(row, 5) = preset("red")
        ws.Cells(row, 6) = preset("green")
        ws.Cells(row, 7) = preset("blue")
        ws.Cells(row, 8) = preset("temperature")
        ws.Cells(row, 9) = preset("scene")
        ws.Cells(row, 10) = preset("created")
        
        row = row + 1
    Next key
    
    ws.Columns.AutoFit
End Sub

Private Sub LoadGroupsFromSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GROUPS_SHEET)
    If ws Is Nothing Then Exit Sub
    
    Dim row As Long
    row = 2 ' Skip header
    
    Do While ws.Cells(row, 1) <> ""
        Dim group As Object
        Set group = CreateObject("Scripting.Dictionary")
        group("name") = ws.Cells(row, 1)
        group("description") = ws.Cells(row, 2)
        group("devices") = ws.Cells(row, 3)
        group("created") = ws.Cells(row, 4)
        
        groupDict(group("name")) = group
        row = row + 1
    Loop
End Sub

Private Sub SaveGroupsToSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GROUPS_SHEET)
    If ws Is Nothing Then Exit Sub
    
    ' Clear existing data
    ws.Rows("2:" & ws.Rows.count).ClearContents
    
    Dim row As Long
    row = 2
    
    Dim key As Variant
    For Each key In groupDict.Keys
        Dim group As Object
        Set group = groupDict(key)
        
        ws.Cells(row, 1) = group("name")
        ws.Cells(row, 2) = group("description")
        ws.Cells(row, 3) = group("devices")
        ws.Cells(row, 4) = group("created")
        
        row = row + 1
    Next key
    
    ws.Columns.AutoFit
End Sub

Private Sub LoadSchedulesFromSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SCHEDULES_SHEET)
    If ws Is Nothing Then Exit Sub
    
    Dim row As Long
    row = 2 ' Skip header
    
    Do While ws.Cells(row, 1) <> ""
        Dim schedule As Object
        Set schedule = CreateObject("Scripting.Dictionary")
        schedule("name") = ws.Cells(row, 1)
        schedule("target") = ws.Cells(row, 2)
        schedule("action") = ws.Cells(row, 3)
        schedule("time") = ws.Cells(row, 4)
        schedule("days") = ws.Cells(row, 5)
        schedule("enabled") = ws.Cells(row, 6)
        schedule("lastRun") = ws.Cells(row, 7)
        
        scheduleDict(schedule("name")) = schedule
        row = row + 1
    Loop
End Sub

Private Sub SaveSchedulesToSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SCHEDULES_SHEET)
    If ws Is Nothing Then Exit Sub
    
    ' Clear existing data
    ws.Rows("2:" & ws.Rows.count).ClearContents
    
    Dim row As Long
    row = 2
    
    Dim key As Variant
    For Each key In scheduleDict.Keys
        Dim schedule As Object
        Set schedule = scheduleDict(key)
        
        ws.Cells(row, 1) = schedule("name")
        ws.Cells(row, 2) = schedule("target")
        ws.Cells(row, 3) = schedule("action")
        ws.Cells(row, 4) = schedule("time")
        ws.Cells(row, 5) = schedule("days")
        ws.Cells(row, 6) = schedule("enabled")
        ws.Cells(row, 7) = schedule("lastRun")
        
        row = row + 1
    Next key
    
    ws.Columns.AutoFit
End Sub

Private Sub CleanupOfflineDevices()
    ' Remove devices not seen for more than 24 hours
    Dim key As Variant
    Dim keysToRemove As New Collection
    
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        
        Dim hoursSinceLastSeen As Long
        hoursSinceLastSeen = DateDiff("h", device("lastSeen"), Now)
        If hoursSinceLastSeen > 24 Then
            keysToRemove.Add key
        End If
    Next key
    
    Dim i As Long
    For i = 1 To keysToRemove.count
        deviceDict.Remove keysToRemove(i)
        DebuggingLog.DebugLog "Removed offline device: " & keysToRemove(i)
    Next i
End Sub

Private Sub UpdateAllDeviceStatuses()
    Dim key As Variant
    For Each key In deviceDict.Keys
        GetDeviceStatus key
        Sleep 100 ' Small delay between status requests
    Next key
End Sub

Private Sub ResetAllData()
    Set deviceDict = CreateObject("Scripting.Dictionary")
    Set presetDict = CreateObject("Scripting.Dictionary")
    Set groupDict = CreateObject("Scripting.Dictionary")
    Set scheduleDict = CreateObject("Scripting.Dictionary")
    SaveDataToSheets
    DebuggingLog.DebugLog "All Govee data reset"
End Sub

' --- Keep all original functions for compatibility ---
Public Sub DiscoverDevices()
    DiscoverDevicesAdvanced
End Sub

Public Function TurnOn(ByVal identifier As String) As Boolean
    TurnOn = SendCommand(identifier, "{""msg"":{""cmd"":""turn"",""data"":{""value"":1}}}")
    If TurnOn Then
        Dim device As Object
        Set device = GetDeviceInfo(identifier)
        If Not device Is Nothing Then
            device("power") = "On"
            SaveDataToSheets
        End If
    End If
End Function

Public Function TurnOff(ByVal identifier As String) As Boolean
    TurnOff = SendCommand(identifier, "{""msg"":{""cmd"":""turn"",""data"":{""value"":0}}}")
    If TurnOff Then
        Dim device As Object
        Set device = GetDeviceInfo(identifier)
        If Not device Is Nothing Then
            device("power") = "Off"
            SaveDataToSheets
        End If
    End If
End Function

Public Function SetBrightness(ByVal identifier As String, ByVal value As Integer) As Boolean
    If value < 0 Or value > 100 Then
        SetBrightness = False
        Exit Function
    End If
    
    SetBrightness = SendCommand(identifier, "{""msg"":{""cmd"":""brightness"",""data"":{""value"":" & value & "}}}")
    If SetBrightness Then
        Dim device As Object
        Set device = GetDeviceInfo(identifier)
        If Not device Is Nothing Then
            device("brightness") = value
            SaveDataToSheets
        End If
    End If
End Function

Public Function SetColor(ByVal identifier As String, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Boolean
    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then
        SetColor = False
        Exit Function
    End If
    
    SetColor = SendCommand(identifier, "{""msg"":{""cmd"":""color"",""data"":{""r"":" & r & ",""g"":" & g & ",""b"":" & b & "}}}")
    If SetColor Then
        Dim device As Object
        Set device = GetDeviceInfo(identifier)
        If Not device Is Nothing Then
            device("color") = r & "," & g & "," & b
            SaveDataToSheets
        End If
    End If
End Function

Public Function SetColorTemperature(ByVal identifier As String, ByVal value As Integer) As Boolean
    If value < 2000 Or value > 9000 Then
        SetColorTemperature = False
        Exit Function
    End If
    
    SetColorTemperature = SendCommand(identifier, "{""msg"":{""cmd"":""colorTem"",""data"":{""value"":" & value & "}}}")
    If SetColorTemperature Then
        Dim device As Object
        Set device = GetDeviceInfo(identifier)
        If Not device Is Nothing Then
            device("temperature") = value
            SaveDataToSheets
        End If
    End If
End Function

Public Function SetScene(ByVal identifier As String, ByVal sceneCode As Integer) As Boolean
    SetScene = SendCommand(identifier, "{""msg"":{""cmd"":""mode"",""data"":{""value"":" & sceneCode & "}}}")
    If SetScene Then
        Dim device As Object
        Set device = GetDeviceInfo(identifier)
        If Not device Is Nothing Then
            device("scene") = sceneCode
            SaveDataToSheets
        End If
    End If
End Function

Public Function TogglePower(ByVal identifier As String) As Boolean
    Dim device As Object
    Set device = GetDeviceInfo(identifier)
    If device Is Nothing Then
        TogglePower = False
        Exit Function
    End If
    
    If device("power") = "On" Then
        TogglePower = TurnOff(identifier)
    Else
        TogglePower = TurnOn(identifier)
    End If
End Function

Public Function AdjustBrightness(ByVal identifier As String, ByVal delta As Integer) As Boolean
    Dim device As Object
    Set device = GetDeviceInfo(identifier)
    If device Is Nothing Then
        AdjustBrightness = False
        Exit Function
    End If
    
    Dim currentBrightness As Integer
    If IsNumeric(device("brightness")) Then
        currentBrightness = CInt(device("brightness"))
    Else
        currentBrightness = 50 ' Default if unknown
    End If
    
    Dim newBrightness As Integer
    newBrightness = currentBrightness + delta
    If newBrightness < 0 Then newBrightness = 0
    If newBrightness > 100 Then newBrightness = 100
    
    AdjustBrightness = SetBrightness(identifier, newBrightness)
End Function

Public Function GetDeviceStatus(ByVal identifier As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim device As Object
    Set device = GetDeviceInfo(identifier)
    If device Is Nothing Then
        GetDeviceStatus = False
        Exit Function
    End If
    
    Dim startTime As Long
    startTime = GetTickCount()
    
    UDP.SendUDP goveeSocket, STATUS_MESSAGE, device("ip"), CONTROL_PORT
    
    Dim buffer(0 To 2047) As Byte
    Dim fromIP As String
    Dim fromPort As Long
    Dim bytesRecv As Long
    
    Do While GetTickCount() - startTime < STATUS_TIMEOUT_MS
        bytesRecv = UDP.RecvUDP(goveeSocket, buffer, fromIP, fromPort)
        If bytesRecv > 0 And fromIP = device("ip") Then
            Dim response As String
            response = Left(StrConv(buffer, vbUnicode), bytesRecv)
            
            Dim parsed As Object
            Set parsed = JSONParser.ParseJSON(response)
            
            If Not parsed Is Nothing And parsed.exists("msg") Then
                Dim msg As Object
                Set msg = parsed("msg")
                
                If msg.exists("cmd") And msg("cmd") = "status" And msg.exists("data") Then
                    Dim data As Object
                    Set data = msg("data")
                    
                    If data.exists("powerState") Then
                        device("power") = IIf(data("powerState") = 1, "On", "Off")
                    End If
                    If data.exists("brightness") Then
                        device("brightness") = data("brightness")
                    End If
                    If data.exists("color") Then
                        Dim color As Object
                        Set color = data("color")
                        device("color") = color("r") & "," & color("g") & "," & color("b")
                    End If
                    If data.exists("colorTemInKelvin") Then
                        device("temperature") = data("colorTemInKelvin")
                    End If
                    If data.exists("mode") Then
                        device("scene") = data("mode")
                    End If
                    
                    device("isOnline") = True
                    device("lastSeen") = Now
                    device("responseTime") = GetTickCount() - startTime
                    
                    SaveDataToSheets
                    UpdateDeviceStats identifier, True
                    GetDeviceStatus = True
                    Exit Function
                End If
            End If
        End If
        DoEvents
        Sleep 10
    Loop
    
    device("isOnline") = False
    device("responseTime") = 0
    SaveDataToSheets
    UpdateDeviceStats identifier, False
    GetDeviceStatus = False
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error getting status for " & identifier & ": " & Err.description
    GetDeviceStatus = False
End Function

Private Sub UpdateDeviceStats(ByVal identifier As String, ByVal success As Boolean)
    If Not deviceStats.exists(identifier) Then
        Set deviceStats(identifier) = CreateObject("Scripting.Dictionary")
        deviceStats(identifier)("totalCommands") = 0
        deviceStats(identifier)("successfulCommands") = 0
        deviceStats(identifier)("lastCommand") = ""
    End If
    
    Dim stats As Object
    Set stats = deviceStats(identifier)
    stats("totalCommands") = stats("totalCommands") + 1
    If success Then
        stats("successfulCommands") = stats("successfulCommands") + 1
    End If
    stats("lastCommand") = Now
End Sub

Private Function SendCommand(ByVal identifier As String, ByVal command As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim device As Object
    Set device = GetDeviceInfo(identifier)
    If device Is Nothing Then
        DebuggingLog.DebugLog "Device not found: " & identifier
        SendCommand = False
        Exit Function
    End If
    
    Dim startTime As Long
    startTime = GetTickCount()
    
    UDP.SendUDP goveeSocket, command, device("ip"), CONTROL_PORT
    
    Dim buffer(0 To 511) As Byte
    Dim fromIP As String
    Dim fromPort As Long
    Dim bytesRecv As Long
    
    Do While GetTickCount() - startTime < 1000 ' 1 second timeout
        bytesRecv = UDP.RecvUDP(goveeSocket, buffer, fromIP, fromPort)
        If bytesRecv > 0 And fromIP = device("ip") Then
            device("responseTime") = GetTickCount() - startTime
            device("isOnline") = True
            device("lastSeen") = Now
            UpdateDeviceStats identifier, True
            SendCommand = True
            Exit Function
        End If
        DoEvents
        Sleep 10
    Loop
    
    device("isOnline") = False
    device("responseTime") = 0
    UpdateDeviceStats identifier, False
    SendCommand = False
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error sending command to " & identifier & ": " & Err.description
    SendCommand = False
End Function

Private Function GetDeviceInfo(ByVal identifier As String) As Object
    Dim mac As String
    mac = UCase(Replace(identifier, ":", ""))
    
    If deviceDict.exists(mac) Then
        Set GetDeviceInfo = deviceDict(mac)
    Else
        Set GetDeviceInfo = Nothing
    End If
End Function

Public Function UpdateDeviceName(ByVal identifier As String, ByVal newName As String) As Boolean
    Dim device As Object
    Set device = GetDeviceInfo(identifier)
    If device Is Nothing Then
        UpdateDeviceName = False
        Exit Function
    End If
    
    device("name") = newName
    SaveDataToSheets
    UpdateDeviceName = True
End Function

Private Sub LoadDevicesFromSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    
    Dim row As Long
    row = 2 ' Skip header
    
    Do While ws.Cells(row, Col_MAC) <> ""
        Dim device As Object
        Set device = CreateObject("Scripting.Dictionary")
        device("mac") = UCase(ws.Cells(row, Col_MAC))
        device("ip") = ws.Cells(row, Col_IP)
        device("model") = ws.Cells(row, Col_Model)
        device("name") = ws.Cells(row, Col_Name)
        device("room") = ws.Cells(row, Col_Room)
        device("group") = ws.Cells(row, Col_Group)
        device("bleVersionHard") = ws.Cells(row, Col_BleHardVer)
        device("bleVersionSoft") = ws.Cells(row, Col_BleSoftVer)
        device("wifiVersionHard") = ws.Cells(row, Col_WifiHardVer)
        device("wifiVersionSoft") = ws.Cells(row, Col_WifiSoftVer)
        device("lastSeen") = ws.Cells(row, Col_LastSeen)
        device("power") = ws.Cells(row, Col_Power)
        device("brightness") = ws.Cells(row, Col_Brightness)
        device("color") = ws.Cells(row, Col_Color)
        device("temperature") = ws.Cells(row, Col_Temperature)
        device("scene") = ws.Cells(row, Col_Scene)
        device("isOnline") = ws.Cells(row, Col_IsOnline)
        device("responseTime") = ws.Cells(row, Col_ResponseTime)
        device("capabilities") = ""
        
        deviceDict(device("mac")) = device
        row = row + 1
    Loop
End Sub

Private Sub SaveDevicesToSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    
    ' Clear existing data
    ws.Rows("2:" & ws.Rows.count).ClearContents
    
    Dim row As Long
    row = 2
    
    Dim key As Variant
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        
        ws.Cells(row, Col_MAC) = device("mac")
        ws.Cells(row, Col_IP) = device("ip")
        ws.Cells(row, Col_Model) = device("model")
        ws.Cells(row, Col_Name) = device("name")
        ws.Cells(row, Col_Room) = device("room")
        ws.Cells(row, Col_Group) = device("group")
        ws.Cells(row, Col_BleHardVer) = device("bleVersionHard")
        ws.Cells(row, Col_BleSoftVer) = device("bleVersionSoft")
        ws.Cells(row, Col_WifiHardVer) = device("wifiVersionHard")
        ws.Cells(row, Col_WifiSoftVer) = device("wifiVersionSoft")
        ws.Cells(row, Col_LastSeen) = device("lastSeen")
        ws.Cells(row, Col_Power) = device("power")
        ws.Cells(row, Col_Brightness) = device("brightness")
        ws.Cells(row, Col_Color) = device("color")
        ws.Cells(row, Col_Temperature) = device("temperature")
        ws.Cells(row, Col_Scene) = device("scene")
        ws.Cells(row, Col_IsOnline) = device("isOnline")
        ws.Cells(row, Col_ResponseTime) = device("responseTime")
        
        row = row + 1
    Next key
    
    ws.Columns.AutoFit
End Sub

Private Function ExportDevices() As String
    On Error GoTo ErrorHandler
    
    Dim csv As String
    csv = "MAC,IP,Model,Name,Room,Group,BLE Hard Ver,BLE Soft Ver,WiFi Hard Ver,WiFi Soft Ver,Last Seen,Power,Brightness,Color,Temperature,Scene,Online,Response Time" & vbCrLf
    
    Dim key As Variant
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        csv = csv & """" & device("mac") & """,""" & device("ip") & """,""" & device("model") & """,""" & device("name") & """,""" & device("room") & """,""" & device("group") & """,""" & device("bleVersionHard") & """,""" & device("bleVersionSoft") & """,""" & device("wifiVersionHard") & """,""" & device("wifiVersionSoft") & """,""" & device("lastSeen") & """,""" & device("power") & """,""" & device("brightness") & """,""" & device("color") & """,""" & device("temperature") & """,""" & device("scene") & """,""" & device("isOnline") & """,""" & device("responseTime") & """" & vbCrLf
    Next key
    
    ExportDevices = "<html><body><pre>" & csv & "</pre></body></html>"
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error exporting devices: " & Err.description
    ExportDevices = GenerateErrorPage("Error exporting devices: " & Err.description)
End Function

Private Function ExportPresets() As String
    On Error GoTo ErrorHandler
    
    Dim csv As String
    csv = "Name,Description,Power,Brightness,Red,Green,Blue,Temperature,Scene,Created" & vbCrLf
    
    Dim key As Variant
    For Each key In presetDict.Keys
        Dim preset As Object
        Set preset = presetDict(key)
        csv = csv & """" & preset("name") & """,""" & preset("description") & """,""" & preset("power") & """,""" & preset("brightness") & """,""" & preset("red") & """,""" & preset("green") & """,""" & preset("blue") & """,""" & preset("temperature") & """,""" & preset("scene") & """,""" & preset("created") & """" & vbCrLf
    Next key
    
    ExportPresets = "<html><body><pre>" & csv & "</pre></body></html>"
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error exporting presets: " & Err.description
    ExportPresets = GenerateErrorPage("Error exporting presets: " & Err.description)
End Function

Private Function ExportSchedules() As String
    On Error GoTo ErrorHandler
    
    Dim csv As String
    csv = "Name,Target,Action,Time,Days,Enabled,Last Run" & vbCrLf
    
    Dim key As Variant
    For Each key In scheduleDict.Keys
        Dim schedule As Object
        Set schedule = scheduleDict(key)
        csv = csv & """" & schedule("name") & """,""" & schedule("target") & """,""" & schedule("action") & """,""" & schedule("time") & """,""" & schedule("days") & """,""" & schedule("enabled") & """,""" & schedule("lastRun") & """" & vbCrLf
    Next key
    
    ExportSchedules = "<html><body><pre>" & csv & "</pre></body></html>"
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error exporting schedules: " & Err.description
    ExportSchedules = GenerateErrorPage("Error exporting schedules: " & Err.description)
End Function

Public Function GetGoveeStatsJSON() As String
    On Error GoTo ErrorHandler
    
    Dim stats As Object
    Set stats = CreateObject("Scripting.Dictionary")
    
    Dim onlineCount As Long
    Dim key As Variant
    
    stats("totalDevices") = deviceDict.count
    For Each key In deviceDict.Keys
        If deviceDict(key)("isOnline") Then onlineCount = onlineCount + 1
    Next key
    stats("onlineDevices") = onlineCount
    stats("groups") = groupDict.count
    stats("presets") = presetDict.count
    stats("schedules") = scheduleDict.count
    stats("monitoring") = isMonitoring
    stats("lastRefresh") = format(lastAutoRefresh, "yyyy-mm-dd hh:mm:ss")
    
    ' Device-specific stats
    Dim deviceStatsArray As Object
    Set deviceStatsArray = CreateObject("Scripting.Dictionary")
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        Dim deviceInfo As Object
        Set deviceInfo = CreateObject("Scripting.Dictionary")
        deviceInfo("name") = device("name")
        deviceInfo("ip") = device("ip")
        deviceInfo("model") = device("model")
        deviceInfo("isOnline") = device("isOnline")
        deviceInfo("responseTime") = device("responseTime")
        deviceInfo("power") = device("power")
        deviceInfo("brightness") = device("brightness")
        deviceInfo("color") = device("color")
        If deviceStats.exists(key) Then
            deviceInfo("totalCommands") = deviceStats(key)("totalCommands")
            deviceInfo("successfulCommands") = deviceStats(key)("successfulCommands")
            deviceInfo("lastCommand") = deviceStats(key)("lastCommand")
        End If
        deviceStatsArray(key) = deviceInfo
    Next key
    stats("devices") = deviceStatsArray
    
    GetGoveeStatsJSON = JSONParser.ConvertToJSON(stats)
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error generating stats JSON: " & Err.description
    GetGoveeStatsJSON = "{""error"":""" & Replace(Err.description, """", "\""") & """}"
End Function

Public Sub CleanupGovee()
    On Error Resume Next
    If goveeSocket <> UDP.INVALID_SOCKET Then
        UDP.CloseUDPSocket goveeSocket
        goveeSocket = UDP.INVALID_SOCKET
    End If
    isMonitoring = False
    DebuggingLog.DebugLog "Govee module cleaned up"
End Sub

Public Function HandleGoveeRequest(ByVal method As String, ByVal path As String, ByVal body As String, ByRef response As String)
    On Error GoTo ErrorHandler
    
    Dim action As String
    Dim query As String
    ' Extract query from path if present
    If InStr(path, "?") > 0 Then
        action = Left(path, InStr(path, "?") - 1)
        query = Mid(path, InStr(path, "?") + 1)
    Else
        action = path
        query = ""
    End If
    
    Select Case LCase(action)
        Case "/govee", "/govee/"
            response = GenerateGoveePage
        Case "/govee/refresh"
            DiscoverDevicesAdvanced
            response = GenerateGoveePage
        Case "/govee/control"
            If LCase(method) = "post" Then
                Dim params As Object
                Set params = ParseQueryOrBody(query, body)
                Dim identifier As String, command As String
                identifier = params("identifier")
                command = params("command")
                If SendCommand(identifier, command) Then
                    response = GenerateGoveePage
                Else
                    response = GenerateErrorPage("Failed to send command to " & identifier)
                End If
            Else
                response = GenerateErrorPage("Control requires POST method")
            End If
        Case "/govee/setname"
            If LCase(method) = "post" Then
                Dim setParams As Object
                Set setParams = ParseQueryOrBody(query, body)
                If UpdateDeviceName(setParams("identifier"), setParams("name")) Then
                    response = GenerateGoveePage
                Else
                    response = GenerateErrorPage("Failed to update device name")
                End If
            Else
                response = GenerateErrorPage("Setname requires POST method")
            End If
        Case "/govee/export"
            response = ExportDevices
        Case "/govee/discover"
            DiscoverDevicesAdvanced
            response = GenerateGoveePage
        Case "/api/govee/stats"
            response = GetGoveeStatsJSON
        Case "/govee/device"
            response = HandleEnhancedGoveeRequest("/govee/device", query)
        Case "/govee/group"
            response = HandleEnhancedGoveeRequest("/govee/group", query)
        Case "/govee/preset"
            response = HandleEnhancedGoveeRequest("/govee/preset", query)
        Case "/govee/schedule"
            response = HandleEnhancedGoveeRequest("/govee/schedule", query)
        Case Else
            response = GenerateErrorPage("Invalid Govee path: " & path)
    End Select
    
    DebuggingLog.DebugLog "Handled Govee request: " & path & " (method: " & method & ")"
    Exit Function

ErrorHandler:
    DebuggingLog.DebugLog "Error in HandleGoveeRequest: " & Err.description
    response = GenerateErrorPage("Govee error: " & Err.description)
End Function
