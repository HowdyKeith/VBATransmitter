Option Explicit

' ServerMonitor Module - REVISED VERSION
' Purpose: Creates and maintains a live LCARS-style dashboard worksheet
'          for monitoring server statuses with starfield animation
' Updates: Removed form-based code, consolidated sheet creation,
'          improved live updates, and integrated with existing modules

' --- Module Variables ---
Private monitorSheet As Worksheet
Private isLiveViewRunning As Boolean
Private Const UPDATE_INTERVAL_MS As Long = 1000 ' Update every 1 second
Private starParticles() As Particle
Private Const NUM_STARS As Long = 100
Private starfieldRunning As Boolean
Private starfieldParent As Worksheet

' --- Particle Type for Starfield ---
Private Type Particle
    Shape As Shape
    x As Double
    y As Double
    Speed As Double
    Size As Double
End Type

' --- Initialize Server Monitor ---
Public Sub InitializeServerMonitor()
    On Error GoTo ErrorHandler
    Debug.Print "[ServerMonitor] Initializing Server Monitor..."

    ' Check if sheet exists, delete if it does
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("ServerMonitor").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    ' Create new worksheet
    Set monitorSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    monitorSheet.Name = "ServerMonitor"
    Set starfieldParent = monitorSheet

    ' Set up LCARS-style layout
    With monitorSheet
        .Cells.Clear
        .Cells.Interior.color = rgb(0, 0, 32) ' Dark LCARS background
        .Cells.Font.Name = "Arial"
        .Cells.Font.color = rgb(255, 153, 102) ' LCARS orange text

        ' Title
        .Range("A1").value = "Smart Traffic LCARS Dashboard"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").Interior.color = rgb(255, 153, 51) ' LCARS orange panel
        .Range("A1").Borders.LineStyle = xlContinuous
        .Range("A1").Borders.color = rgb(102, 51, 153) ' Purple border
        .Range("A1:C1").Merge

        ' Last Updated
        .Range("A2").value = "Stardate: " & format(Now, "yyyy.mm.dd.hhmm")
        .Range("B2").value = "Last Updated: " & format(Now, "yyyy-mm-dd hh:mm:ss")
        .Range("A2:B2").Font.Size = 12
        .Range("A2:B2").Interior.color = rgb(51, 51, 102) ' Dark blue panel
        .Range("A2:B2").Borders.LineStyle = xlContinuous

        ' Server Table Headers
        .Range("A4:E4").value = Array("Server", "Status", "Port", "Clients", "Additional Info")
        .Range("A4:E4").Font.Bold = True
        .Range("A4:E4").Font.Size = 12
        .Range("A4:E4").Interior.color = rgb(102, 51, 153) ' Purple header
        .Range("A4:E4").Font.color = rgb(255, 255, 255)
        .Range("A4:E4").Borders.LineStyle = xlContinuous

        ' Server Names
        .Range("A5").value = "Chat Server"
        .Range("A6").value = "HTTP Server"
        .Range("A7").value = "IoT Server"
        .Range("A8").value = "Traffic Server"
        .Range("A9").value = "API Gateway"
        .Range("A10").value = "App Launcher"

        ' Initial Status
        .Range("B5:B10").value = "INACTIVE"
        .Range("B5:B10").Interior.color = rgb(255, 102, 102) ' Red for inactive
        .Range("C5:C10").value = "-"
        .Range("D5:D10").value = 0
        .Range("E5:E10").value = ""

        ' Formatting
        .Range("A5:E10").Borders.LineStyle = xlContinuous
        .Range("A5:E10").Borders.color = rgb(102, 51, 153)
        .Columns("A:E").AutoFit
        .Columns("A").ColumnWidth = 20
        .Columns("E").ColumnWidth = 30

        ' Add Control Buttons (Shapes for LCARS style)
        Dim shp As Shape
        Dim btnTop As Double: btnTop = 320
        Dim btnLeft As Double: btnLeft = 10
        Dim btnWidth As Double: btnWidth = 100
        Dim btnHeight As Double: btnHeight = 30

        ' Start Servers Button
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
        shp.Fill.ForeColor.rgb = rgb(0, 204, 0) ' Green
        shp.TextFrame2.TextRange.Text = "Start Servers"
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = True
        shp.OnAction = "ServerMonitor.StartAllServers"

        ' Stop Servers Button
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 120, btnTop, btnWidth, btnHeight)
        shp.Fill.ForeColor.rgb = rgb(204, 0, 0) ' Red
        shp.TextFrame2.TextRange.Text = "Stop Servers"
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = True
        shp.OnAction = "ServerMonitor.StopAllServers"

        ' Refresh Button
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 240, btnTop, btnWidth, btnHeight)
        shp.Fill.ForeColor.rgb = rgb(255, 204, 0) ' Yellow
        shp.TextFrame2.TextRange.Text = "Refresh"
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = True
        shp.OnAction = "ServerMonitor.RefreshMonitor"

        ' Clear Sockets Button
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 360, btnTop, btnWidth, btnHeight)
        shp.Fill.ForeColor.rgb = rgb(153, 51, 255) ' Purple
        shp.TextFrame2.TextRange.Text = "Clear Sockets"
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = True
        shp.OnAction = "ServerMonitor.ClearSockets"

        ' Instructions
        .Range("A12").value = "Instructions:"
        .Range("A13").value = "• Click 'Start Servers' to launch all servers"
        .Range("A14").value = "• Click 'Stop Servers' to stop all servers"
        .Range("A15").value = "• Click 'Refresh' for manual update"
        .Range("A16").value = "• Click 'Clear Sockets' if sockets are stuck"
        .Range("A12:A16").Font.Size = 10
        .Range("A12:A16").Font.color = rgb(255, 204, 153)
    End With

    ' Initialize Starfield
    InitStarfield monitorSheet, NUM_STARS

    Debug.Print "[ServerMonitor] Server Monitor initialized successfully"
    Exit Sub

ErrorHandler:
    Debug.Print "[ServerMonitor] Error in InitializeServerMonitor: " & Err.description
    Application.DisplayAlerts = True
End Sub

' --- Toggle Live View ---
Public Sub ToggleLiveView()
    On Error GoTo ErrorHandler
    isLiveViewRunning = Not isLiveViewRunning

    If isLiveViewRunning Then
        Debug.Print "[ServerMonitor] Starting live monitor updates..."
        UpdateServerMonitor
        starfieldRunning = True
        MoveStars
    Else
        Debug.Print "[ServerMonitor] Live monitor updates stopped"
        starfieldRunning = False
    End If

    ' Update status display
    If Not monitorSheet Is Nothing Then
        monitorSheet.Range("A2").value = "Stardate: " & format(Now, "yyyy.mm.dd.hhmm") & _
            IIf(isLiveViewRunning, " (LIVE)", " (STOPPED)")
        monitorSheet.Range("A2").Interior.color = IIf(isLiveViewRunning, rgb(51, 153, 51), rgb(153, 51, 51))
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[ServerMonitor] Error in ToggleLiveView: " & Err.description
End Sub

' --- Update Server Monitor with Live Data ---
Public Sub UpdateServerMonitor()
    On Error GoTo ErrorHandler
    If Not isLiveViewRunning Or monitorSheet Is Nothing Then Exit Sub

    ' Update timestamp and stardate
    With monitorSheet
        .Range("A2").value = "Stardate: " & format(Now, "yyyy.mm.dd.hhmm") & _
            IIf(isLiveViewRunning, " (LIVE)", " (STOPPED)")
        .Range("B2").value = format(Now, "yyyy-mm-dd hh:mm:ss")

        ' Chat Server
        Dim chatRunning As Boolean, chatPort As Long, chatCount As Long, chatMessages As Long
        On Error Resume Next
        chatRunning = TransmissionServer.GetChatRunning()
        chatPort = TransmissionServer.GetChatPort()
        chatCount = TransmissionServer.GetChatCount()
        chatMessages = TransmissionServer.GetChatMessageCount()
        On Error GoTo ErrorHandler
        .Range("B5").value = IIf(chatRunning, "ACTIVE", "INACTIVE")
        .Range("B5").Interior.color = IIf(chatRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C5").value = IIf(chatPort > 0, chatPort, "-")
        .Range("D5").value = chatCount
        .Range("E5").value = "Messages: " & chatMessages

        ' HTTP Server
        Dim httpRunning As Boolean, httpPort As Long, httpCount As Long, httpRequests As Long
        On Error Resume Next
        httpRunning = HttpServer.isHTTPRunning()
        httpPort = HttpServer.GetHTTPPort()
        httpCount = HttpServer.GetHTTPCount()
        httpRequests = HttpServer.GetHTTPRequestCount()
        On Error GoTo ErrorHandler
        .Range("B6").value = IIf(httpRunning, "ACTIVE", "INACTIVE")
        .Range("B6").Interior.color = IIf(httpRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C6").value = IIf(httpPort > 0, httpPort, "-")
        .Range("D6").value = httpCount
        .Range("E6").value = HttpServer.GetHTTPStats()

        ' IoT Server
        Dim iotRunning As Boolean, iotPort As Long, iotCount As Long, sensorCount As Long
        On Error Resume Next
        iotRunning = TransmissionServer.GetIoTRunning()
        iotPort = TransmissionServer.GetIoTPort()
        iotCount = TransmissionServer.GetIoTCount()
        sensorCount = TransmissionServer.GetSensorCount()
        On Error GoTo ErrorHandler
        .Range("B7").value = IIf(iotRunning, "ACTIVE", "INACTIVE")
        .Range("B7").Interior.color = IIf(iotRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C7").value = IIf(iotPort > 0, iotPort, "-")
        .Range("D7").value = iotCount
        .Range("E7").value = "Sensors: " & sensorCount

        ' Traffic Server
        Dim trafficRunning As Boolean, trafficPort As Long, trafficCount As Long
        On Error Resume Next
        trafficRunning = TransmissionServer.GetTrafficRunning()
        trafficPort = TransmissionServer.GetTrafficPort()
        trafficCount = TransmissionServer.GetTrafficCount()
        On Error GoTo ErrorHandler
        .Range("B8").value = IIf(trafficRunning, "ACTIVE", "INACTIVE")
        .Range("B8").Interior.color = IIf(trafficRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C8").value = IIf(trafficPort > 0, trafficPort, "-")
        .Range("D8").value = trafficCount
        .Range("E8").value = "Traffic Control"

        ' API Gateway
        Dim apiRunning As Boolean, apiPort As Long, apiCount As Long, apiCalls As Long
        On Error Resume Next
        apiRunning = TransmissionServer.GetApiGatewayRunning()
        apiPort = TransmissionServer.GetApiGatewayPort()
        apiCount = TransmissionServer.GetApiGatewayClientCount()
        apiCalls = TransmissionServer.GetApiGatewayCallCount()
        On Error GoTo ErrorHandler
        .Range("B9").value = IIf(apiRunning, "ACTIVE", "INACTIVE")
        .Range("B9").Interior.color = IIf(apiRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C9").value = IIf(apiPort > 0, apiPort, "-")
        .Range("D9").value = apiCount
        .Range("E9").value = "API Calls: " & apiCalls

        ' App Launcher
        Dim appRunning As Boolean, appPort As Long, appCount As Long
        On Error Resume Next
        appRunning = AppLaunch.GetAppLauncherStatus()
        appPort = AppLaunch.GetAppLauncherPort()
        appCount = AppLaunch.GetAppLauncherClientCount()
        On Error GoTo ErrorHandler
        .Range("B10").value = IIf(appRunning, "ACTIVE", "INACTIVE")
        .Range("B10").Interior.color = IIf(appRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C10").value = IIf(appPort > 0, appPort, "-")
        .Range("D10").value = appCount
        .Range("E10").value = AppLaunch.GetAppLauncherStats()

        ' UDP Server
        Dim udpRunning As Boolean, UdpPort As Long, udpCount As Long, udpPackets As Long
        On Error Resume Next
        udpRunning = TransmissionServer.GetUDPRunning()
        UdpPort = TransmissionServer.GetUDPPort()
        udpCount = TransmissionServer.GetUDPConnectionCount()
        udpPackets = TransmissionServer.GetUDPPacketCount()
        On Error GoTo ErrorHandler
        .Range("B11").value = IIf(udpRunning, "ACTIVE", "INACTIVE")
        .Range("B11").Interior.color = IIf(udpRunning, rgb(200, 255, 200), rgb(255, 102, 102))
        .Range("C11").value = IIf(UdpPort > 0, UdpPort, "-")
        .Range("D11").value = udpCount
        .Range("E11").value = "Packets: " & udpPackets

        .Columns("A:E").AutoFit
        .Columns("A").ColumnWidth = 20
        .Columns("E").ColumnWidth = 30
    End With

    ' Schedule next update
    If isLiveViewRunning Then
        Application.OnTime Now + TimeValue("00:00:01"), "ServerMonitor.UpdateServerMonitor"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[ServerMonitor] Error in UpdateServerMonitor: " & Err.description
    If isLiveViewRunning Then
        Application.OnTime Now + TimeValue("00:00:05"), "ServerMonitor.UpdateServerMonitor"
    End If
End Sub

' --- Start All Servers ---
Public Sub StartAllServers()
    On Error Resume Next
    Debug.Print "[ServerMonitor] Starting all servers..."
    TrafficManager.RunAllServers
End Sub

' --- Stop All Servers (Full Cleanup, Non-OnTime) ---
Public Sub StopAllServers()
    On Error Resume Next
    DebuggingLog.DebugLog "[TrafficManager] Stopping all servers..."

    ' Stop main loop
    isRunning = False

    ' Stop HTTP Server
    If HttpServer.isHTTPRunning() Then HttpServer.StopHttpServer

    ' Stop UDP Server
    If TransmissionServer.GetUDPRunning() Then TransmissionServer.StopUDPServer

    ' Stop IoT Gateway
    If IoTGateway.GetRunning() Then IoTGateway.StopServer

    ' Stop FTP Server
    If FTPServer.GetRunning() Then FTPServer.StopServer

    ' Cleanup MQTT
    If mqttEnabled Then
        CleanupMQTTAdvanced
        mqttEnabled = False
        DebuggingLog.DebugLog "[TrafficManager] MQTT client stopped"
    End If

    ' Clear any pending timers just in case
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "TrafficManager.ProcessAllServers", , False
    On Error GoTo 0

    DebuggingLog.DebugLog "[TrafficManager] All servers stopped"
End Sub


' --- Clear Sockets ---
Public Sub ClearSockets()
    On Error Resume Next
    Debug.Print "[ServerMonitor] Clearing socket lockup..."
    TrafficManager.ClearSocketLockup
End Sub

' --- Refresh Monitor ---
Public Sub RefreshMonitor()
    On Error Resume Next
    Debug.Print "[ServerMonitor] Manual monitor refresh..."
    UpdateServerMonitor
End Sub

' --- Get Monitor Status ---
Public Function GetMonitorStatus() As String
    GetMonitorStatus = "Monitor: " & IIf(isLiveViewRunning, "ACTIVE", "INACTIVE")
End Function

' --- Show Monitor Sheet ---
Public Sub ShowMonitorSheet()
    On Error Resume Next
    If monitorSheet Is Nothing Then
        InitializeServerMonitor
    End If
    monitorSheet.Activate
End Sub

' --- Emergency Stop ---
Public Sub EmergencyStop()
    On Error Resume Next
    Debug.Print "[ServerMonitor] EMERGENCY STOP - All monitoring and servers stopped"
    isLiveViewRunning = False
    starfieldRunning = False
    TrafficManager.StopAllServers
End Sub

' --- Force Cleanup ---
Public Sub ForceCleanup()
    On Error Resume Next
    isLiveViewRunning = False
    starfieldRunning = False
    Application.DisplayAlerts = False

    ' Delete starfield shapes
    StopStarfield

    ' Delete monitor sheet
    If Not monitorSheet Is Nothing Then
        monitorSheet.Delete
        Set monitorSheet = Nothing
    End If

    ' Clear pending timers
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "ServerMonitor.UpdateServerMonitor", schedule:=False
    Application.OnTime Now + TimeValue("00:00:05"), "ServerMonitor.UpdateServerMonitor", schedule:=False
    Application.OnTime Now + TimeValue("00:00:00.05"), "ServerMonitor.MoveStars", schedule:=False
    On Error GoTo 0

    Application.DisplayAlerts = True
    Debug.Print "[ServerMonitor] Monitor cleanup completed"
End Sub

' --- Initialize Starfield ---
Public Sub InitStarfield(ws As Worksheet, Optional numStars As Long = NUM_STARS)
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim p As Particle
    Dim x As Double, y As Double, sz As Double, spd As Double

    Set starfieldParent = ws
    starfieldRunning = False
    ReDim starParticles(1 To numStars)

    Randomize
    For i = 1 To numStars
        x = Rnd * ws.UsedRange.width
        y = Rnd * ws.UsedRange.height
        sz = 2 + Rnd * 2
        spd = 1 + Rnd * 3

        Set p.Shape = ws.Shapes.AddShape(msoShapeOval, x, y, sz, sz)
        p.Shape.Fill.ForeColor.rgb = rgb(255, 255, 255)
        p.Shape.Line.Visible = msoFalse
        p.x = x
        p.y = y
        p.Speed = spd
        p.Size = sz
        starParticles(i) = p
    Next i

    starfieldRunning = True
    MoveStars

    Exit Sub

ErrorHandler:
    Debug.Print "[ServerMonitor] Error in InitStarfield: " & Err.description
End Sub

' --- Animate Starfield ---
Public Sub MoveStars()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim p As Particle

    If Not starfieldRunning Or starfieldParent Is Nothing Then Exit Sub

    For i = LBound(starParticles) To UBound(starParticles)
        p = starParticles(i)
        p.y = p.y + p.Speed
        If p.y > starfieldParent.UsedRange.height Then
            p.y = 0
            p.x = Rnd * starfieldParent.UsedRange.width
        End If
        p.Shape.Left = p.x
        p.Shape.Top = p.y
        starParticles(i) = p
    Next i

    ' Schedule next animation frame
    If starfieldRunning Then
        Application.OnTime Now + TimeValue("00:00:00.05"), "ServerMonitor.MoveStars"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[ServerMonitor] Error in MoveStars: " & Err.description
End Sub

' --- Stop Starfield ---
Public Sub StopStarfield()
    On Error Resume Next
    Dim i As Long
    starfieldRunning = False
    For i = LBound(starParticles) To UBound(starParticles)
        If Not starParticles(i).Shape Is Nothing Then
            starParticles(i).Shape.Delete
        End If
    Next i
    ReDim starParticles(0 To 0) ' Clear array
End Sub

' --- Play Button Beep ---
Public Sub PlayButtonBeep()
    Beep
End Sub

