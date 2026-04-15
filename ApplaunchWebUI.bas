Option Explicit
' Module: AppLaunchWebUI
' Purpose: Unified LCARS-style dashboard for AppLauncher, Chat,
'          IoT Gateway, HTTP/UDP servers, Outlook, FTP, and APIs
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' --- Constants for Dashboard ---
Private Const REFRESH_INTERVAL As Double = 0.5 ' seconds

'Public Function GetChatClientCount() As Long: GetChatClientCount = 3: End Function
'Public Function GetIoTClientCount() As Long: GetIoTClientCount = 5: End Function
'Public Function GetHTTPRequestCount() As Long: GetHTTPRequestCount = 102: End Function
'Public Function IsUDPServerRunning() As Boolean: IsUDPServerRunning = True: End Function
'Public Function GetUDPMessageCount() As Long: GetUDPMessageCount = 256: End Function
Public Function IsAPIServerRunning() As Boolean: IsAPIServerRunning = True: End Function
Public Function GetAPICount() As Long: GetAPICount = 8: End Function




'***************************************************************
' --- Server Status Checks ---
'***************************************************************

Public Function IsChatServerRunning() As Boolean
    On Error Resume Next
    ' Replace with actual TrafficServer status check
    IsChatServerRunning = Not TrafficServer Is Nothing And TrafficServer.isRunning
End Function

Public Function IsMQTTBrokerRunning() As Boolean
    On Error Resume Next
    IsMQTTBrokerRunning = Not MQTTBroker Is Nothing And MQTTBroker.isRunning
End Function

Public Function IsHTTPServerRunning() As Boolean
    On Error Resume Next
    IsHTTPServerRunning = Not HttpServer Is Nothing And HttpServer.isRunning
End Function

Public Function IsIoTGatewayRunning() As Boolean
    On Error Resume Next
    IsIoTGatewayRunning = Not IoTGateway Is Nothing And IoTGateway.isRunning
End Function


Public Function GetActiveRuleCount() As Long
    On Error Resume Next
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If Not olApp Is Nothing Then
        GetActiveRuleCount = olApp.Session.DefaultStore.GetRules.count
    End If
    Set olApp = Nothing
End Function

'***************************************************************
' --- LCARS Dashboard Drawing ---
'***************************************************************

Public Sub DrawStatus(ByVal itemName As String, ByVal isRunning As Boolean)
    Dim status As String
    status = IIf(isRunning, "ONLINE", "OFFLINE")
    Debug.Print format(Now, "hh:nn:ss") & " | " & itemName & ": " & status
    ' Here you can replace Debug.Print with actual LCARS panel drawing code
End Sub

Public Sub UpdateServerStatus()
    DrawStatus "Chat Server", IsChatServerRunning
    DrawStatus "MQTT Broker", IsMQTTBrokerRunning
    DrawStatus "HTTP Server", IsHTTPServerRunning
    DrawStatus "IoT Gateway", IsIoTGatewayRunning
    DrawStatus "Outlook Rules", IsOutlookRunning
End Sub

'***************************************************************
' --- Starfield Animation (Simple Placeholder) ---
'***************************************************************

Public Sub DrawStarfield()
    ' Placeholder: replace with your starfield animation logic
    Debug.Print "? ? ? ? ?" ' just a simple star effect for console
End Sub

'***************************************************************
' --- Auto-Refresh Dashboard Loop ---
'***************************************************************

Public Sub StartDashboardLoop()
    Dim nextTick As Double
    nextTick = Timer + REFRESH_INTERVAL
    
    Do While True
        UpdateServerStatus
        DrawStarfield
        DoEvents
        
        ' Wait until next tick
        While Timer < nextTick
            DoEvents
        Wend
        
        ' Schedule next tick
        nextTick = Timer + REFRESH_INTERVAL
    Loop
End Sub

' --- Main Dashboard Page ---
Public Function GenerateLCARSDashboardPage() As String
    Dim html As String
    Dim panels As Collection
    Dim p As Variant
    
    ' Initialize HTML
    html = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>"
    html = html & "<title>AppLaunch Dashboard</title>"
    html = html & "<style>"
    html = html & "body{background-color:#000;color:#fff;font-family:Orbitron,sans-serif;margin:0;padding:0;}"
    html = html & "table{width:100%;border-collapse:collapse;}"
    html = html & "td{padding:15px;text-align:center;vertical-align:top;border-radius:15px;transition:0.3s;}"
    html = html & "td:hover{transform:scale(1.05);}"
    html = html & "h2{margin:5px;font-size:1.2em;}"
    html = html & "p{margin:5px;font-size:0.9em;}"
    html = html & "a{color:#0ff;text-decoration:none;font-weight:bold;}"
    html = html & ".traffic-status{max-height:100px;overflow-y:auto;text-align:left;font-size:0.8em;}"
    html = html & "</style></head><body>"
    
    ' Add canvas starfield
    html = html & "<canvas id='starfield' style='position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;'></canvas>"
    
    ' Initialize panels collection
    Set panels = New Collection
    
    ' --- Add TrafficManager status panel ---
    Dim tmRaw As String, tmHtml As String
    On Error Resume Next
    tmRaw = TrafficManager.GetEnhancedTrafficManagerStatus()
    If Err.Number <> 0 Then
        tmRaw = "TrafficManager: Status unavailable"
        Err.Clear
    End If
    On Error GoTo 0
    
    Dim lines() As String, startIdx As Long
    lines = Split(tmRaw, vbCrLf)
    startIdx = IIf(UBound(lines) - 19 > 0, UBound(lines) - 19, 0)
    
    tmHtml = "<div class='traffic-status'>"
    Dim i As Long, tickName As String, tickColor As String
    
    For i = startIdx To UBound(lines)
        If InStr(lines(i), ":") > 0 Then
            tickName = Trim(Mid(lines(i), InStrRev(lines(i), ":") + 1))
            
            ' Determine color: green if OK, red if FAIL
            If InStr(UCase(tickName), "OK") > 0 Then
                tickColor = "#0f0" ' green
            ElseIf InStr(UCase(tickName), "FAIL") > 0 Then
                tickColor = "#f00" ' red
            Else
                tickColor = "#ff0" ' yellow
            End If
            
            tmHtml = tmHtml & "<span style='color:" & tickColor & "'>" & HTMLEncode(lines(i)) & "</span><br>"
        End If
    Next i
    tmHtml = tmHtml & "</div>"
    
    panels.Add Array("Traffic Manager", _
                     "<span style='color:#0ff'>Active</span>", _
                     tmHtml, _
                     "/dashboard", _
                     "linear-gradient(135deg,#4444FF,#222266)")
    
    ' --- Add other panels ---
    panels.Add Array("Outlook", GetOutlookStatusHTML(), "Rules Checked: " & GetOutlookRuleCheckCount(), "/outlook", "linear-gradient(135deg,#FF9966,#CC6600)")
    panels.Add Array("Chat Server", IIf(IsChatServerRunning(), "<span style='color:#0f0'>Running</span>", "<span style='color:#f00'>Stopped</span>"), "Clients: " & GetChatClientCount(), "/chat", "linear-gradient(135deg,#66FFCC,#339966)")
    panels.Add Array("IoT Gateway", IIf(IsIoTGatewayRunning(), "<span style='color:#0f0'>Online</span>", "<span style='color:#f00'>Offline</span>"), "Connections: " & GetIoTClientCount(), "/iot", "linear-gradient(135deg,#FF66CC,#993366)")
    panels.Add Array("HTTP Server", IIf(IsHTTPServerRunning(), "<span style='color:#0f0'>Online</span>", "<span style='color:#f00'>Offline</span>"), "Requests: " & GetHTTPRequestCount(), "/http", "linear-gradient(135deg,#66CCFF,#336699)")
    panels.Add Array("UDP Server", IIf(IsUDPServerRunning(), "<span style='color:#0f0'>Listening</span>", "<span style='color:#f00'>Stopped</span>"), "Packets: " & GetUDPMessageCount(), "/udp", "linear-gradient(135deg,#FFFF66,#999933)")
    panels.Add Array("API Services", IIf(IsAPIServerRunning(), "<span style='color:#0f0'>Active</span>", "<span style='color:#f00'>Inactive</span>"), "Endpoints: " & GetAPICount(), "/api", "linear-gradient(135deg,#FF66FF,#993399)")
    
    ' --- Build panel table ---
    html = html & "<table><tr>"
    Dim colCount As Integer: colCount = 3
    i = 0
    
    For Each p In panels
        html = html & GenerateLCARSSquare(p(0), p(1), p(2), p(3), p(4))
        i = i + 1
        If i Mod colCount = 0 Then html = html & "</tr><tr>"
    Next
    html = html & "</tr></table>"
    
    ' --- Add Starfield JS ---
    html = html & "<script>"
    html = html & "var canvas=document.getElementById('starfield');"
    html = html & "var ctx=canvas.getContext('2d');"
    html = html & "canvas.width=window.innerWidth;canvas.height=window.innerHeight;"
    html = html & "var stars=[];for(var i=0;i<200;i++){stars.push({x:Math.random()*canvas.width,y:Math.random()*canvas.height,r:Math.random()*1.5+0.5,speed:Math.random()*0.5+0.2});}"
    html = html & "function animate(){ctx.clearRect(0,0,canvas.width,canvas.height);for(var i=0;i<stars.length;i++){var s=stars[i];s.y-=s.speed;if(s.y<0)s.y=canvas.height;ctx.beginPath();ctx.arc(s.x,s.y,s.r,0,Math.PI*2);ctx.fillStyle='#fff';ctx.fill();}requestAnimationFrame(animate);}"
    html = html & "animate();"
    html = html & "window.addEventListener('resize',function(){canvas.width=window.innerWidth;canvas.height=window.innerHeight;});"
    html = html & "</script>"
    
    html = html & "</body></html>"
    
    GenerateLCARSDashboardPage = html
End Function

' Helper function for HTML encoding
Public Function HTMLEncode(ByVal Text As String) As String
    HTMLEncode = Replace(Replace(Replace(Replace(Replace(Text, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;"), "'", "&#39;")
End Function

' Updated server status functions that connect to your actual modules
Public Function GetOutlookStatusHTML() As String
    On Error Resume Next
    If IsOutlookRunning() Then
        GetOutlookStatusHTML = "<span style='color:#0f0'>Connected</span>"
    Else
        GetOutlookStatusHTML = "<span style='color:#f00'>Disconnected</span>"
    End If
    If Err.Number <> 0 Then GetOutlookStatusHTML = "<span style='color:#f00'>Unknown</span>"
    On Error GoTo 0
End Function

Public Function IsChatServerRunningOther() As Boolean
    On Error Resume Next
    IsChatServerRunning = TransmissionServer.IsChatServerRunning()
    If Err.Number <> 0 Then IsChatServerRunning = False
    On Error GoTo 0
End Function

Public Function IsIoTGatewayRunningother() As Boolean
    On Error Resume Next
    IsIoTGatewayRunning = IoTGateway.GetRunning()
    If Err.Number <> 0 Then IsIoTGatewayRunning = False
    On Error GoTo 0
End Function

Public Function IsHTTPServerRunningOTHER() As Boolean
    On Error Resume Next
    IsHTTPServerRunning = HttpServer.isHTTPRunning()
    If Err.Number <> 0 Then IsHTTPServerRunning = False
    On Error GoTo 0
End Function

Public Function IsUDPServerRunning() As Boolean
    On Error Resume Next
    IsUDPServerRunning = TransmissionServer.GetUDPRunning()
    If Err.Number <> 0 Then IsUDPServerRunning = False
    On Error GoTo 0
End Function

Public Function GetHTTPRequestCount() As Long
    On Error Resume Next
    GetHTTPRequestCount = HttpServer.GetHTTPRequestCount()
    If Err.Number <> 0 Then GetHTTPRequestCount = 0
    On Error GoTo 0
End Function

Public Function GetUDPMessageCount() As Long
    On Error Resume Next
    GetUDPMessageCount = TransmissionServer.GetUDPMessageCount()
    If Err.Number <> 0 Then GetUDPMessageCount = 0
    On Error GoTo 0
End Function

Public Function GetChatClientCount() As Long
    On Error Resume Next
    GetChatClientCount = TransmissionServer.GetChatClientCount()
    If Err.Number <> 0 Then GetChatClientCount = 0
    On Error GoTo 0
End Function

Public Function GetIoTClientCount() As Long
    On Error Resume Next
    GetIoTClientCount = IoTGateway.GetClientCount()
    If Err.Number <> 0 Then GetIoTClientCount = 0
    On Error GoTo 0
End Function

' --- Generic LCARS Square Panel Generator ---
Public Function GenerateLCARSSquare(title As String, status As String, info As String, link As String, Optional colorGradient As String = "linear-gradient(135deg,#FF9966,#CC6600)") As String
    Dim html As String
    html = "<td style='border:1px solid #ccc;padding:15px;text-align:center;background:" & colorGradient & ";color:#fff;border-radius:15px;transition: all 0.3s ease;'>"
    html = html & "<h2>" & title & "</h2>"
    html = html & "<p>Status: " & status & "</p>"
    html = html & "<p>" & info & "</p>"
    html = html & "<a href='" & link & "'>View</a>"
    html = html & "</td>"
    GenerateLCARSSquare = html
End Function

