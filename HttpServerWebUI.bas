Option Explicit

'***************************************************************
' HttpServerWebUI Module - Web Interface Generation
' Purpose: Generate LCARS-style web pages for the HTTP server
' Updated to include UDP server information panel and landing page
' VBA format for Excel compatibility
'***************************************************************

' Set to True as AppLauncher module exists
#Const HAS_APPLAUNCHER = True

Public Function GenerateLCARSAppLauncherPage() As String
    Dim html As String, style As String, body As String, stardate As String
    
    ' Calculate stardate for LCARS aesthetic
    stardate = "2025." & format(Now, "ddd.dd.hh")
    
    ' Build CSS styles for LCARS-inspired design
    style = "<style>" & vbCrLf
    style = style & "body {background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; margin: 0; padding: 20px;}" & vbCrLf
    style = style & ".container {max-width: 1200px; margin: auto;}" & vbCrLf
    style = style & ".bar {height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate;}" & vbCrLf
    style = style & "@keyframes flash {from {opacity: 0.6;} to {opacity: 1;}}" & vbCrLf
    style = style & ".btn {display: inline-block; padding: 15px 30px; margin: 10px; background: #CC6600; color: black; font-weight: bold; border-radius: 10px; cursor: pointer; transition: background 0.3s; border: 2px solid #FFFF99;}" & vbCrLf
    style = style & ".btn:hover {background: #FF9966; border-color: #99CCFF;}" & vbCrLf
    style = style & ".btn:active {background: #993300;}" & vbCrLf
    style = style & ".header {font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px;}" & vbCrLf
    style = style & ".subheader {font-size: 18px; color: #FFFF99; margin: 10px 0;}" & vbCrLf
    style = style & ".grid {display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;}" & vbCrLf
    style = style & "</style>"
    
    ' Build HTML body for main launcher
    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<h1 class='header'>LCARS - Application Launcher</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; Application Hub &bull; Stardate " & stardate & "</div>" & vbCrLf
    body = body & "<div class='grid'>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/outlook"">Outlook Scanner</div>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/reports"">Reports Dashboard</div>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/settings"">System Settings</div>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/data"">Data Analysis</div>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/udp"">UDP Monitor</div>" & vbCrLf
    body = body & "</div>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "</div>"
    
    ' Combine HTML components
    html = "<html><head><title>LCARS App Launcher</title>" & vbCrLf
    html = html & style & vbCrLf
    html = html & "</head><body>" & vbCrLf
    html = html & body & vbCrLf
    html = html & "</body></html>"
    
    GenerateLCARSAppLauncherPage = html
End Function


Public Function GenerateLCARSUDPLandingPage() As String
    Dim html As String, style As String, body As String, stardate As String
    Dim updateTime As String
    
    ' Calculate stardate and update time
    stardate = "2025." & format(Now, "ddd.dd.hh")
    updateTime = format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Build CSS styles
    style = "<style>" & vbCrLf
    style = style & "body {background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; margin: 0; padding: 20px;}" & vbCrLf
    style = style & ".container {max-width: 1200px; margin: auto;}" & vbCrLf
    style = style & ".bar {height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate;}" & vbCrLf
    style = style & "@keyframes flash {from {opacity: 0.6;} to {opacity: 1;}}" & vbCrLf
    style = style & ".btn {padding: 10px; background: #CC6600; color: black; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99;}" & vbCrLf
    style = style & ".btn:hover {background: #FF9966; border-color: #99CCFF;}" & vbCrLf
    style = style & ".header {font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px;}" & vbCrLf
    style = style & ".subheader {font-size: 18px; color: #FFFF99; margin: 10px 0;}" & vbCrLf
    style = style & ".section {margin: 15px 0; padding: 10px; border: 2px solid #663399; border-radius: 10px; background: #1C2526;}" & vbCrLf
    style = style & ".section-title {font-size: 24px; color: #99CCFF; text-transform: uppercase;}" & vbCrLf
    style = style & ".status-active {color: #33FF33;}" & vbCrLf
    style = style & ".status-inactive {color: #FF3333;}" & vbCrLf
    style = style & ".indicator {display: inline-block; width: 10px; height: 10px; background: #33FF33; border-radius: 50%; margin-left: 10px; animation: blink 1s infinite;}" & vbCrLf
    style = style & "@keyframes blink {0% {opacity: 1;} 50% {opacity: 0.3;} 100% {opacity: 1;}}" & vbCrLf
    style = style & "</style>"
    
    ' Build HTML body
    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<h1 class='header'>LCARS - UDP Monitor</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; UDP Traffic Monitoring &bull; Stardate " & stardate & "</div>" & vbCrLf
    body = body & "<div class='section'>" & vbCrLf
    body = body & "<div class='section-title'>UDP SERVER STATUS</div>" & vbCrLf
    body = body & "<p>Status: <span class='" & IIf(TransmissionServer.IsUDPServerRunning(), "status-active", "status-inactive") & "'>" & IIf(TransmissionServer.IsUDPServerRunning(), "ACTIVE", "INACTIVE") & "</span><span class='indicator'></span></p>" & vbCrLf
    body = body & "<p>Port: <strong>" & TransmissionServer.GetUDPPort() & "</strong></p>" & vbCrLf
    body = body & "<p>Total Messages Received: <strong>" & TransmissionServer.GetUDPMessageCount() & "</strong></p>" & vbCrLf
    body = body & "<p>Last Update: <strong>" & updateTime & "</strong></p>" & vbCrLf
    body = body & "<div class='btn' onclick='window.location.href=""/dashboard"">View System Dashboard</div>" & vbCrLf
    body = body & "</div>" & vbCrLf
    body = body & "<div class='section'>" & vbCrLf
    body = body & "<div class='section-title'>UDP TRAFFIC DATA</div>" & vbCrLf
    body = body & "<p>Traffic Data: <strong>Smart Traffic data processing is active. Detailed metrics are available in the Excel workbook.</strong></p>" & vbCrLf
    body = body & "<p>Note: Use the Traffic Control module for advanced data analysis.</p>" & vbCrLf
    body = body & "</div>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<a href='/index.html' class='btn'>Return to Home</a>" & vbCrLf
    body = body & "</div>"
    
    ' Combine HTML components
    html = "<html><head><title>LCARS UDP Monitor</title>" & vbCrLf
    html = html & style & vbCrLf
    html = html & "</head><body>" & vbCrLf
    html = html & body & vbCrLf
    html = html & "</body></html>"
    
    GenerateLCARSUDPLandingPage = html
End Function

Public Function GenerateLCARSReportsLandingPage() As String
    Dim html As String, style As String, body As String, stardate As String
    
    ' Calculate stardate
    stardate = "2025." & format(Now, "ddd.dd.hh")
    
    ' Build CSS styles
    style = "<style>" & vbCrLf
    style = style & "body {background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; margin: 0; padding: 20px;}" & vbCrLf
    style = style & ".container {max-width: 1200px; margin: auto;}" & vbCrLf
    style = style & ".bar {height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate;}" & vbCrLf
    style = style & "@keyframes flash {from {opacity: 0.6;} to {opacity: 1;}}" & vbCrLf
    style = style & ".btn {padding: 10px; background: #CC6600; color: black; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99;}" & vbCrLf
    style = style & ".btn:hover {background: #FF9966; border-color: #99CCFF;}" & vbCrLf
    style = style & ".header {font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px;}" & vbCrLf
    style = style & ".subheader {font-size: 18px; color: #FFFF99; margin: 10px 0;}" & vbCrLf
    style = style & "</style>"
    
    ' Build HTML body
    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<h1 class='header'>LCARS - Reports Dashboard</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; Data Reporting &bull; Stardate " & stardate & "</div>" & vbCrLf
    body = body & "<div class='btn'>View Reports</div><br><br>" & vbCrLf
    body = body & "<div style='color:#FF9966;'>Reports data will be displayed here.</div>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<a href='/index.html' class='btn'>Return to Home</a>" & vbCrLf
    body = body & "</div>"
    
    ' Combine HTML components
    html = "<html><head><title>LCARS Reports Dashboard</title>" & vbCrLf
    html = html & style & vbCrLf
    html = html & "</head><body>" & vbCrLf
    html = html & body & vbCrLf
    html = html & "</body></html>"
    
    GenerateLCARSReportsLandingPage = html
End Function

Public Function GenerateLCARSSettingsLandingPage() As String
    Dim html As String, style As String, body As String, stardate As String
    
    ' Calculate stardate
    stardate = "2025." & format(Now, "ddd.dd.hh")
    
    ' Build CSS styles
    style = "<style>" & vbCrLf
    style = style & "body {background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; margin: 0; padding: 20px;}" & vbCrLf
    style = style & ".container {max-width: 1200px; margin: auto;}" & vbCrLf
    style = style & ".bar {height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate;}" & vbCrLf
    style = style & "@keyframes flash {from {opacity: 0.6;} to {opacity: 1;}}" & vbCrLf
    style = style & ".btn {padding: 10px; background: #CC6600; color: black; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99;}" & vbCrLf
    style = style & ".btn:hover {background: #FF9966; border-color: #99CCFF;}" & vbCrLf
    style = style & ".header {font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px;}" & vbCrLf
    style = style & ".subheader {font-size: 18px; color: #FFFF99; margin: 10px 0;}" & vbCrLf
    style = style & "</style>"
    
    ' Build HTML body
    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<h1 class='header'>LCARS - System Settings</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; Configuration &bull; Stardate " & stardate & "</div>" & vbCrLf
    body = body & "<div class='btn'>Modify Settings</div><br><br>" & vbCrLf
    body = body & "<div style='color:#FF9966;'>System settings will be configurable here.</div>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<a href='/index.html' class='btn'>Return to Home</a>" & vbCrLf
    body = body & "</div>"
    
    ' Combine HTML components
    html = "<html><head><title>LCARS System Settings</title>" & vbCrLf
    html = html & style & vbCrLf
    html = html & "</head><body>" & vbCrLf
    html = html & body & vbCrLf
    html = html & "</body></html>"
    
    GenerateLCARSSettingsLandingPage = html
End Function

Public Function GenerateLCARSDataLandingPage() As String
    Dim html As String, style As String, body As String, stardate As String
    
    ' Calculate stardate
    stardate = "2025." & format(Now, "ddd.dd.hh")
    
    ' Build CSS styles
    style = "<style>" & vbCrLf
    style = style & "body {background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; margin: 0; padding: 20px;}" & vbCrLf
    style = style & ".container {max-width: 1200px; margin: auto;}" & vbCrLf
    style = style & ".bar {height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate;}" & vbCrLf
    style = style & "@keyframes flash {from {opacity: 0.6;} to {opacity: 1;}}" & vbCrLf
    style = style & ".btn {padding: 10px; background: #CC6600; color: black; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99;}" & vbCrLf
    style = style & ".btn:hover {background: #FF9966; border-color: #99CCFF;}" & vbCrLf
    style = style & ".header {font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px;}" & vbCrLf
    style = style & ".subheader {font-size: 18px; color: #FFFF99; margin: 10px 0;}" & vbCrLf
    style = style & ".section {margin: 15px 0; padding: 10px; border: 2px solid #663399; border-radius: 10px; background: #1C2526;}" & vbCrLf
    style = style & ".section-title {font-size: 24px; color: #99CCFF; text-transform: uppercase;}" & vbCrLf
    style = style & "</style>"
    
    ' Build HTML body
    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<h1 class='header'>LCARS - Data Analysis</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; Data Insights &bull; Stardate " & stardate & "</div>" & vbCrLf
    body = body & "<div class='section'>" & vbCrLf
    body = body & "<div class='section-title'>Data Overview</div>" & vbCrLf
    body = body & "<p>Data analysis and visualization tools will be available here.</p>" & vbCrLf
    body = body & "<div class='btn'>Run Analysis</div>" & vbCrLf
    body = body & "</div>" & vbCrLf
    body = body & "<div class='bar'></div>" & vbCrLf
    body = body & "<a href='/index.html' class='btn'>Return to Home</a>" & vbCrLf
    body = body & "</div>"
    
    ' Combine HTML components
    html = "<html><head><title>LCARS Data Analysis</title>" & vbCrLf
    html = html & style & vbCrLf
    html = html & "</head><body>" & vbCrLf
    html = html & body & vbCrLf
    html = html & "</body></html>"
    
    GenerateLCARSDataLandingPage = html
End Function

' --- Generate All LCARS Pages ---
Public Function GenerateAllLCARSPages() As String
    Dim pages As String
    pages = pages & "<page path='/index.html'>" & ApplaunchWebUI.GenerateIndexPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/outlook'>" & OutlookWWW.GenerateLCARSOutlookLandingPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/apps'>" & AppLaunch.GenerateAppsPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/dashboard'>" & GenerateLCARSDashboardPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/reports'>" & GenerateLCARSReportsLandingPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/settings'>" & GenerateLCARSSettingsLandingPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/data'>" & GenerateLCARSDataLandingPage() & "</page>" & vbCrLf
    pages = pages & "<page path='/udp'>" & GenerateLCARSUDPLandingPage() & "</page>" & vbCrLf
    GenerateAllLCARSPages = pages
End Function

' --- Get All LCARS Page Content ---
Public Function GetAllLCARSPageContent() As String
    Dim pages As String
    pages = pages & ApplaunchWebUI.GenerateIndexPage() & vbCrLf
    pages = pages & OutlookWWW.GenerateLCARSOutlookLandingPage() & vbCrLf
    pages = pages & AppLaunch.GenerateAppsPage() & vbCrLf
    pages = pages & GenerateLCARSDashboardPage() & vbCrLf
    pages = pages & GenerateLCARSReportsLandingPage() & vbCrLf
    pages = pages & GenerateLCARSSettingsLandingPage() & vbCrLf
    pages = pages & GenerateLCARSDataLandingPage() & vbCrLf
    pages = pages & GenerateLCARSUDPLandingPage() & vbCrLf
    GetAllLCARSPageContent = pages
End Function

' --- Safe Getter Functions ---
Private Function GetActivSocketCount() As Long
    On Error Resume Next
    GetActivSocketCount = TransmissionServer.GetActiveSocketCount()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetActivSocketCount: " & Err.description
        GetActivSocketCount = 0
    End If
    On Error GoTo 0
End Function

Private Function GetMemoryUsageSafe() As String
    On Error Resume Next
    GetMemoryUsageSafe = HttpServer.GetMemoryUsage()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetMemoryUsageSafe: " & Err.description
        GetMemoryUsageSafe = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function GetAppCountSafe() As String
    On Error Resume Next
    GetAppCountSafe = CStr(AppLaunch.m_appConfig.count)
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetAppCountSafe: " & Err.description
        GetAppCountSafe = "5"
    End If
    On Error GoTo 0
End Function

Private Function IsChatServerRunning() As Boolean
    On Error Resume Next
    IsChatServerRunning = TransmissionServer.IsChatServerRunning()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in IsChatServerRunning: " & Err.description
        IsChatServerRunning = False
    End If
    On Error GoTo 0
End Function

Public Function GetChatPort() As String
    On Error Resume Next
    GetChatPort = CStr(TransmissionServer.GetChatPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetChatPort: " & Err.description
        GetChatPort = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function GetChatClientCount() As String
    On Error Resume Next
    GetChatClientCount = CStr(TransmissionServer.GetChatClientCount())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetChatClientCount: " & Err.description
        GetChatClientCount = "0"
    End If
    On Error GoTo 0
End Function

Private Function IsHTTPServerRunning() As Boolean
    On Error Resume Next
    IsHTTPServerRunning = HttpServer.GetHTTPRunning()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in IsHttpServerRunning: " & Err.description
        IsHTTPServerRunning = False
    End If
    On Error GoTo 0
End Function

Private Function GetHTTPPort() As String
    On Error Resume Next
    GetHTTPPort = CStr(HttpServer.GetHTTPPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetHTTPPort: " & Err.description
        GetHTTPPort = "8080"
    End If
    On Error GoTo 0
End Function

Private Function GetHTTPRequestCount() As String
    On Error Resume Next
    GetHTTPRequestCount = CStr(AppLaunch.m_totalRequests)
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetHttpRequestCount: " & Err.description
        GetHTTPRequestCount = "0"
    End If
    On Error GoTo 0
End Function

Private Function IsIoTServerRunning() As Boolean
    On Error Resume Next
    IsIoTServerRunning = TransmissionServer.IsIoTServerRunning()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in IsIoTServerRunning: " & Err.description
        IsIoTServerRunning = False
    End If
    On Error GoTo 0
End Function

Private Function GetIoTPort() As String
    On Error Resume Next
    GetIoTPort = CStr(TransmissionServer.GetIoTPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetIoTPort: " & Err.description
        GetIoTPort = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function GetIoTClientCount() As String
    On Error Resume Next
    GetIoTClientCount = CStr(TransmissionServer.GetIoTClientCount())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetIoTClientCount: " & Err.description
        GetIoTClientCount = "0"
    End If
    On Error GoTo 0
End Function

Private Function IsTrafficServerRunning() As Boolean
    On Error Resume Next
    IsTrafficServerRunning = TransmissionServer.IsTrafficServerRunning()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in IsTrafficServerRunning: " & Err.description
        IsTrafficServerRunning = False
    End If
    On Error GoTo 0
End Function

Private Function GetTrafficPort() As String
    On Error Resume Next
    GetTrafficPort = CStr(TransmissionServer.GetTrafficPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetTrafficPort: " & Err.description
        GetTrafficPort = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function GetTrafficClientCount() As String
    On Error Resume Next
    GetTrafficClientCount = CStr(TransmissionServer.GetTrafficClientCount())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetTrafficClientCount: " & Err.description
        GetTrafficClientCount = "0"
    End If
    On Error GoTo 0
End Function

Private Function GetAppLauncherStatus() As Boolean
    On Error Resume Next
    GetAppLauncherStatus = AppLaunch.GetAppLauncherStatus()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetAppLauncherStatus: " & Err.description
        GetAppLauncherStatus = False
    End If
    On Error GoTo 0
End Function

Private Function GetAppLauncherPort() As String
    On Error Resume Next
    GetAppLauncherPort = CStr(AppLaunch.GetAppLauncherPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetAppLauncherPort: " & Err.description
        GetAppLauncherPort = "8080"
    End If
    On Error GoTo 0
End Function

Private Function GetAppLauncherClientCount() As String
    On Error Resume Next
    GetAppLauncherClientCount = CStr(AppLaunch.GetAppLauncherClientCount())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetAppLauncherClientCount: " & Err.description
        GetAppLauncherClientCount = "0"
    End If
    On Error GoTo 0
End Function

Private Function IsApiGatewayRunning() As Boolean
    On Error Resume Next
    IsApiGatewayRunning = TransmissionServer.IsApiGatewayRunning()
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in IsApiGatewayRunning: " & Err.description
        IsApiGatewayRunning = False
    End If
    On Error GoTo 0
End Function

Private Function GetApiGatewayPort() As String
    On Error Resume Next
    GetApiGatewayPort = CStr(TransmissionServer.GetApiGatewayPort())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetApiGatewayPort: " & Err.description
        GetApiGatewayPort = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function GetApiGatewayClientCount() As String
    On Error Resume Next
    GetApiGatewayClientCount = CStr(TransmissionServer.GetApiGatewayClientCount())
    If Err.Number <> 0 Then
        DebuggingLog.DebugLog "Error in GetApiGatewayClientCount: " & Err.description
        GetApiGatewayClientCount = "0"
    End If
    On Error GoTo 0
End Function

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

' --- Format Uptime ---
Private Function FormatUptime(ByVal seconds As Double) As String
    Dim hours As Long, minutes As Long, secs As Long
    If seconds < 0 Or seconds > 31536000 Then ' Cap at 1 year to prevent overflow
        FormatUptime = "Unknown"
        Exit Function
    End If
    hours = seconds \ 3600
    minutes = (seconds Mod 3600) \ 60
    secs = seconds Mod 60
    FormatUptime = format(hours, "00") & ":" & format(minutes, "00") & ":" & format(secs, "00")
End Function

Public Function GeneratePlaceholderPage(ByVal serviceName As String) As String
    Dim html As String
    html = GenerateHTMLHeader(serviceName)
    html = html & GenerateCss()
    html = html & "</head><body>"
    html = html & "<div class=""main-container"">"
    html = html & "<header class=""lcars-header""><h1>" & UCase(serviceName) & "</h1></header>"
    html = html & "<div class=""control-panel"">"
    html = html & "<p>" & serviceName & " page is not yet implemented.</p>"
    html = html & "<a href=""/"" class=""lcars-button"">Return Home</a>"
    html = html & "</div>"
    html = html & "</div></body></html>"
    GeneratePlaceholderPage = html
End Function


