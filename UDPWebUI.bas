Option Explicit

'***************************************************************
' UDPWebUI Module - LCARS UDP Server Web Interface
' Purpose: Generate LCARS-style web pages, CSS, and JS endpoints
'***************************************************************

#Const HAS_APPLAUNCHER = True

' =============================================================
' Public Router: Handle UDP Web Requests
' =============================================================
Public Function HandleUDPWebRequest(ByVal path As String) As String
    Select Case LCase$(path)
        Case "/udp", "/udp.html"
            HandleUDPWebRequest = GenerateLCARSUDPLandingPage()
        Case "/udp.css"
            HandleUDPWebRequest = GenerateUDPCss()
        Case "/udp.js"
            HandleUDPWebRequest = GenerateUDPJs()
        Case Else
            HandleUDPWebRequest = "<h1>404 Not Found</h1><p>The requested UDP page was not found.</p>"
    End Select
End Function

' =============================================================
' Landing Page
' =============================================================
Public Function GenerateLCARSUDPLandingPage() As String
    Dim html As String, body As String
    Dim stardate As String, updateTime As String
    
    stardate = "2025." & format(Now, "ddd.dd.hh")
    updateTime = format(Now, "yyyy-mm-dd hh:mm:ss")
    
    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<h1 class='header'>LCARS - UDP Monitor</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; UDP Traffic Monitoring &bull; Stardate " & stardate & "</div>" & vbCrLf
    
    ' --- UDP Server Status ---
    body = body & "<div class='section'>" & vbCrLf
    body = body & "<div class='section-title'>UDP SERVER STATUS</div>" & vbCrLf
    body = body & "<p>Status: <span class='" & IIf(IsUDPServerRunning(), "status-active", "status-inactive") & "'>" & _
                  IIf(IsUDPServerRunning(), "ACTIVE", "INACTIVE") & "</span><span class='indicator'></span></p>" & vbCrLf
    body = body & "<p>Port: <strong>" & GetUDPPort() & "</strong></p>" & vbCrLf
    body = body & "<p>Total Messages Received: <strong>" & GetUDPMessageCount() & "</strong></p>" & vbCrLf
    body = body & "<p>Last Update: <strong>" & updateTime & "</strong></p>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/dashboard"">View System Dashboard</div>" & vbCrLf
    body = body & "</div>" & vbCrLf
    
    ' --- Device List ---
    body = body & "<div class='section'>" & vbCrLf
    body = body & "<div class='section-title'>DISCOVERED DEVICES</div>" & vbCrLf
    body = body & RenderUDPDevicesTable()
    body = body & "</div>" & vbCrLf
    
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<a href='/index.html' class='btn'>Return to Home</a>" & vbCrLf
    body = body & "</div>"
    
    ' --- Wrap in HTML ---
    html = "<html><head>" & vbCrLf
    html = html & "<title>LCARS UDP Monitor</title>" & vbCrLf
    html = html & "<link rel='stylesheet' href='/udp.css'>" & vbCrLf
    html = html & "<script src='/udp.js'></script>" & vbCrLf
    html = html & "</head><body>" & body & "</body></html>"
    
    GenerateLCARSUDPLandingPage = html
End Function

' =============================================================
' Render Device Table
' =============================================================
Private Function RenderUDPDevicesTable() As String
    Dim tbl As String, key As Variant
    On Error Resume Next
    
    tbl = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse; width:100%; color:#FF9966;'>" & vbCrLf
    tbl = tbl & "<tr style='background:#663399; color:#FFFF99;'>" & _
          "<th>Device</th><th>IP</th><th>Last Seen</th><th>Status</th></tr>" & vbCrLf
    
    If Not TransmissionServer.Devices Is Nothing Then
        For Each key In TransmissionServer.Devices.Keys
            Dim dev As Object
            Set dev = TransmissionServer.Devices(key)
            tbl = tbl & "<tr>" & _
                        "<td>" & dev("name") & "</td>" & _
                        "<td>" & dev("ip") & "</td>" & _
                        "<td>" & dev("lastSeen") & "</td>" & _
                        "<td><span class='" & IIf(dev("online"), "status-active", "status-inactive") & "'>" & _
                        IIf(dev("online"), "Online", "Offline") & "</span></td>" & _
                        "</tr>" & vbCrLf
        Next key
    Else
        tbl = tbl & "<tr><td colspan='4'>No devices discovered.</td></tr>" & vbCrLf
    End If
    
    tbl = tbl & "</table>" & vbCrLf
    
    RenderUDPDevicesTable = tbl
    On Error GoTo 0
End Function

' =============================================================
' Dynamic CSS
' =============================================================
Private Function GenerateUDPCss() As String
    Dim css As String
    css = "body {background:black; color:#FF9966; font-family:'OCR-A', Arial, sans-serif; margin:0; padding:20px;}" & vbCrLf
    css = css & ".container {max-width:1200px; margin:auto;}" & vbCrLf
    css = css & ".bar {height:40px; background:linear-gradient(to right,#663399,#CC6600); margin:10px 0; animation:flash 1.5s infinite alternate;}" & vbCrLf
    css = css & "@keyframes flash {from {opacity:0.6;} to {opacity:1;}}" & vbCrLf
    css = css & ".btn {padding:10px; background:#CC6600; color:black; font-weight:bold; border-radius:8px; cursor:pointer; border:2px solid #FFFF99;}" & vbCrLf
    css = css & ".btn:hover {background:#FF9966; border-color:#99CCFF;}" & vbCrLf
    css = css & ".header {font-size:36px; color:#99CCFF; text-shadow:0 0 10px #99CCFF; margin-bottom:20px;}" & vbCrLf
    css = css & ".subheader {font-size:18px; color:#FFFF99; margin:10px 0;}" & vbCrLf
    css = css & ".section {margin:15px 0; padding:10px; border:2px solid #663399; border-radius:10px; background:#1C2526;}" & vbCrLf
    css = css & ".section-title {font-size:24px; color:#99CCFF; text-transform:uppercase;}" & vbCrLf
    css = css & ".status-active {color:#33FF33;}" & vbCrLf
    css = css & ".status-inactive {color:#FF3333;}" & vbCrLf
    css = css & ".indicator {display:inline-block; width:10px; height:10px; background:#33FF33; border-radius:50%; margin-left:10px; animation:blink 1s infinite;}" & vbCrLf
    css = css & "@keyframes blink {0% {opacity:1;} 50% {opacity:0.3;} 100% {opacity:1;}}" & vbCrLf
    GenerateUDPCss = css
End Function

' =============================================================
' Dynamic JavaScript
' =============================================================
Private Function GenerateUDPJs() As String
    Dim js As String
    js = "function autoRefresh() { setTimeout(function() { location.reload(); }, 5000); }" & vbCrLf
    js = js & "window.onload = autoRefresh;" & vbCrLf
    GenerateUDPJs = js
End Function

' =============================================================
' Safe Wrappers
' =============================================================
Private Function IsUDPServerRunning() As Boolean
    On Error Resume Next
    IsUDPServerRunning = TransmissionServer.IsUDPServerRunning()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in IsUDPServerRunning: " & Err.description
        IsUDPServerRunning = False
    End If
    On Error GoTo 0
End Function

Private Function GetUDPPort() As String
    On Error Resume Next
    GetUDPPort = CStr(TransmissionServer.GetUDPPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetUDPPort: " & Err.description
        GetUDPPort = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function GetUDPMessageCount() As String
    On Error Resume Next
    GetUDPMessageCount = CStr(TransmissionServer.GetUDPMessageCount())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetUDPMessageCount: " & Err.description
        GetUDPMessageCount = "0"
    End If
    On Error GoTo 0
End Function

'store here
Sub TestUDPQueue()
    Dim udpQueue As clsUDPQueue
    Set udpQueue = New clsUDPQueue
    
    Dim item As clsUDPQueueItem
    
    ' Create first message
    Set item = New clsUDPQueueItem
    item.msg = "Hello"
    item.remoteIP = "192.168.1.1"
    item.remotePort = 5000
    item.timestamp = Timer
    udpQueue.Enqueue item
    
    ' Create second message
    Set item = New clsUDPQueueItem
    item.msg = "World"
    item.remoteIP = "10.0.0.2"
    item.remotePort = 5001
    item.timestamp = Timer
    udpQueue.Enqueue item
    
    ' Peek first item
    Set item = udpQueue.Peek()
    Debug.Print "Peek:", item.msg
    
    ' Dequeue all items
    Do While Not udpQueue.IsEmpty
        Set item = udpQueue.Dequeue()
        Debug.Print "Dequeued:", item.msg, item.remoteIP, item.remotePort
    Loop
End Sub

