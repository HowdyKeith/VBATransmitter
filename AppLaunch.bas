Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' --- Module Variables ---
Public m_totalRequests As Long
Public m_lastActivity As Date
Public m_serverStartTime As Date
Private LOG_DIR As String
Private LOG_PREFIX As String
Public m_appConfig As Object
Private Const CONFIG_DIR As String = "C:\SmartTraffic\"
Private Const CONFIG_FILE As String = "C:\SmartTraffic\app_config.json"
Public UseStaticFiles As Boolean
Public UseVirtualPages As Boolean
Public httpRunning As Boolean
Public isRunning As Boolean
Private httpPortNum As Long
Private routes As Collection
Private Const DEFAULT_PORT As Long = 8080

Dim dashboardForm As Object
Dim starShapes() As Object
Dim starX() As Single, starY() As Single, starSpeed() As Single
Dim starCount As Long
Dim animRunning As Boolean

' --- Initialize Module Variables ---
Private Sub InitializeModuleVariables()
    LOG_DIR = CONFIG_DIR
    LOG_PREFIX = "server_log_"
End Sub

' --- Initialize Routes ---
Public Sub InitializeAppLaunch()
    On Error GoTo ErrorHandler
    Set routes = New Collection
    
    ' Add core routes first
    routes.Add Array("/", "GenerateHomePage"), "/"
    routes.Add Array("/dashboard", "GenerateDashboardPage"), "/dashboard"
    routes.Add Array("/status", "GenerateStatusPage"), "/status"
    routes.Add Array("/config", "GenerateConfigPage"), "/config"
    
    ' Add module routes with error handling
    On Error Resume Next
    
    ' Try to get OutlookWebUI routes
    Dim moduleRoutes As Collection
    Set moduleRoutes = HttpServer.GetRoutes
    If Not moduleRoutes Is Nothing Then
        Dim route As Variant
        For Each route In moduleRoutes
            If IsArray(route) And UBound(route) >= 1 Then
                routes.Add route, route(0)
            End If
        Next route
        DebugLog "Added " & moduleRoutes.count & " OutlookWebUI routes"
    Else
        ' Add default Outlook routes if module not available
        DebugLog "Added default Outlook routes (module not available)"
    End If
    
    ' Try to get Govee routes
    Set moduleRoutes = Nothing
    Set moduleRoutes = govee.GetRoutes
    If Not moduleRoutes Is Nothing Then
        For Each route In moduleRoutes
            If IsArray(route) And UBound(route) >= 1 Then
                routes.Add route, route(0)
            End If
        Next route
        DebugLog "Added " & moduleRoutes.count & " Govee routes"
    Else
        ' Add default Govee routes if module not available
        routes.Add Array("/govee", "GenerateGoveePage"), "/govee"
        routes.Add Array("/govee/status", "GenerateGoveeStatusPage"), "/govee/status"
        DebugLog "Added default Govee routes (module not available)"
    End If
    
    On Error GoTo ErrorHandler
    
    ' Add launch routes for each app in m_appConfig
    If Not m_appConfig Is Nothing Then
        Dim appKey As Variant
        For Each appKey In m_appConfig.Keys
            routes.Add Array("/launch/" & appKey, "LaunchApplication:" & appKey), "/launch/" & appKey
        Next appKey
        DebugLog "Added " & m_appConfig.count & " launch routes from app config"
    End If
    
    DebugLog "AppLaunch initialized with " & routes.count & " routes"
    Exit Sub

ErrorHandler:
    DebugLog "Error initializing AppLaunch: " & Err.description
    ' Ensure we have at least basic routes
    If routes Is Nothing Then Set routes = New Collection
    If routes.count = 0 Then
        routes.Add Array("/", "GenerateHomePage"), "/"
        routes.Add Array("/dashboard", "GenerateDashboardPage"), "/dashboard"
    End If
End Sub

' --- Handle HTTP Requests ---
Public Sub HandleAppRequest(ByVal method As String, ByVal path As String, ByVal body As String, ByRef response As String)
    On Error GoTo ErrorHandler

    ' --- Update stats ---
    m_totalRequests = m_totalRequests + 1
    m_lastActivity = Now

    Dim query As String
    If InStr(path, "?") > 0 Then
        query = Mid(path, InStr(path, "?") + 1)
        path = Left(path, InStr(path, "?") - 1)
    End If

    ' --- Normalize path ---
    If Right(path, 1) = "/" And Len(path) > 1 Then path = Left(path, Len(path) - 1)

    ' --- Find route dynamically ---
    Dim route As Variant
    Dim handler As String
    Dim found As Boolean
    found = False

    For Each route In routes
        If LCase(route(0)) = LCase(path) Then
            handler = route(1)
            found = True
            Exit For
        End If
    Next route

    If Not found Then
        response = GenerateErrorPage("Invalid path: " & path, "Available paths: " & GetAvailablePaths())
        DebugLog "No route found for path: " & path
        Exit Sub
    End If

    ' --- Handle LaunchApplication routes dynamically ---
    If InStr(handler, ":") > 0 Then
        Dim parts() As String
        parts = Split(handler, ":")
        Select Case parts(0)
            Case "LaunchApplication"
                LaunchApplication parts(1)
                response = GenerateLaunchSuccessPage(parts(1))
            Case Else
                ' Delegate to module if available
                If IsModuleAvailable(parts(0)) Then
                    Select Case LCase(parts(0))
                        Case "outlookwebui": OutlookWebUI.HandleOutlookAction parts(1): response = OutlookWebUI.GenerateLCARSOutlookLandingPage
                        Case "govee": response = govee.HandleGoveeRequest(path, query, method, body)
                        Case Else: response = GenerateErrorPage("Unknown module: " & parts(0))
                    End Select
                Else
                    response = GenerateErrorPage("Module not available: " & parts(0))
                End If
        End Select
    Else
        ' --- Direct function call ---
        On Error Resume Next
        response = Application.Run(handler)
        If Err.Number <> 0 Then
            response = GenerateErrorPage("Handler error: " & handler & " - " & Err.description)
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If

    DebugLog "Handled request for path: " & path & " (method: " & method & ")"
    Exit Sub

ErrorHandler:
    DebugLog "Error in HandleAppRequest: " & Err.description
    response = GenerateErrorPage("Server error: " & Err.description)
End Sub

Private Function GenerateDashboardPage() As String
    On Error GoTo ErrorHandler

    ' --- Ensure objects exist ---
    If m_appConfig Is Nothing Then Set m_appConfig = LoadDefaultApps
    If routes Is Nothing Then InitializeAppLaunch

    Dim html As String
    Dim route As Variant
    Dim appKey As Variant

    html = GenerateHTMLHeader("SmartTraffic Dashboard")
    html = html & "<div class='container'>"
    html = html & "<h1 class='header'>System Dashboard</h1>"

    ' --- Server Status Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Server Status</div>"
    html = html & "<p>HTTP Server: " & IIf(httpRunning, "Running on port " & httpPortNum, "Stopped") & "</p>"
    html = html & "<p>Uptime: " & GetUptime() & "</p>"
    html = html & "<p>Total Requests: " & m_totalRequests & "</p>"
    html = html & "<p>Last Activity: " & IIf(IsDate(m_lastActivity), format(m_lastActivity, "yyyy-mm-dd hh:mm:ss"), "None") & "</p>"
    html = html & "</div>"

    ' --- Quick Launch Applications Section ---
    If m_appConfig.count > 0 Then
        html = html & "<div class='section'>"
        html = html & "<div class='section-title'>Quick Launch Applications</div>"
        For Each appKey In m_appConfig.Keys
            html = html & "<p><a href='/launch/" & appKey & "' class='btn'>" & UCase(Left(appKey, 1)) & Mid(appKey, 2) & "</a></p>"
        Next appKey
        html = html & "</div>"
    End If

    ' --- Available Routes Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Available Routes</div>"
    If routes.count > 0 Then
        For Each route In routes
            html = html & "<p><a href='" & route(0) & "' class='btn'>" & route(0) & "</a> ? " & route(1) & "</p>"
        Next route
    Else
        html = html & "<p>No routes available</p>"
    End If
    html = html & "</div>"

    html = html & "<div class='bar'></div>"
    html = html & "</div></body></html>"

    GenerateDashboardPage = html
    Exit Function

ErrorHandler:
    DebugLog "Error in GenerateDashboardPage: " & Err.description
    GenerateDashboardPage = GenerateErrorPage("Error generating dashboard: " & Err.description)
End Function

Private Function GenerateStatusPage() As String
    On Error GoTo ErrorHandler

    ' --- Ensure objects exist ---
    If routes Is Nothing Then InitializeAppLaunch
    If m_appConfig Is Nothing Then Set m_appConfig = LoadDefaultApps

    Dim html As String
    Dim route As Variant
    Dim appKey As Variant

    html = GenerateHTMLHeader("System Status")
    html = html & "<div class='container'>"
    html = html & "<h1 class='header'>System Status</h1>"

    ' --- HTTP Server Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>HTTP Server</div>"
    html = html & "<p>Status: " & IIf(httpRunning, "Running", "Stopped") & "</p>"
    html = html & "<p>Port: " & httpPortNum & "</p>"
    html = html & "<p>Start Time: " & IIf(IsDate(m_serverStartTime), format(m_serverStartTime, "yyyy-mm-dd hh:mm:ss"), "Unknown") & "</p>"
    html = html & "<p>Uptime: " & GetUptime() & "</p>"
    html = html & "</div>"

    ' --- Statistics Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Statistics</div>"
    html = html & "<p>Total Requests: " & m_totalRequests & "</p>"
    html = html & "<p>Routes Configured: " & IIf(routes Is Nothing, 0, routes.count) & "</p>"
    html = html & "<p>Apps Configured: " & IIf(m_appConfig Is Nothing, 0, m_appConfig.count) & "</p>"
    html = html & "</div>"

    ' --- Quick Launch Applications Section ---
    If Not m_appConfig Is Nothing And m_appConfig.count > 0 Then
        html = html & "<div class='section'>"
        html = html & "<div class='section-title'>Quick Launch Applications</div>"
        For Each appKey In m_appConfig.Keys
            html = html & "<p><a href='/launch/" & appKey & "' class='btn'>" & UCase(Left(appKey, 1)) & Mid(appKey, 2) & "</a></p>"
        Next appKey
        html = html & "</div>"
    End If

    ' --- Available Routes Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Available Routes</div>"
    If Not routes Is Nothing And routes.count > 0 Then
        For Each route In routes
            html = html & "<p><a href='" & route(0) & "' class='btn'>" & route(0) & "</a> ? " & route(1) & "</p>"
        Next route
    Else
        html = html & "<p>No routes available</p>"
    End If
    html = html & "</div>"

    html = html & "<div class='bar'></div>"
    html = html & "</div></body></html>"

    GenerateStatusPage = html
    Exit Function

ErrorHandler:
    DebugLog "Error in GenerateStatusPage: " & Err.description
    GenerateStatusPage = GenerateErrorPage("Error generating status page: " & Err.description)
End Function

Private Function GenerateConfigPage() As String
    On Error GoTo ErrorHandler

    ' --- Ensure objects exist ---
    If routes Is Nothing Then InitializeAppLaunch

    Dim html As String
    html = GenerateHTMLHeader("Configuration")
    html = html & "<div class='container'>"
    html = html & "<h1 class='header'>Configuration</h1>"

    ' --- Routes Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Available Routes</div>"
    If Not routes Is Nothing And routes.count > 0 Then
        Dim route As Variant
        For Each route In routes
            html = html & "<p><code>" & route(0) & "</code> ? " & route(1) & "</p>"
        Next route
    Else
        html = html & "<p>No routes configured</p>"
    End If
    html = html & "</div>"

    ' --- Navigation ---
    html = html & "<a href='/' class='btn'>Home</a>"
    html = html & "<a href='/dashboard' class='btn'>Dashboard</a>"
    html = html & "</div></body></html>"

    GenerateConfigPage = html
    Exit Function

ErrorHandler:
    DebugLog "Error in GenerateConfigPage: " & Err.description
    GenerateConfigPage = GenerateErrorPage("Error generating config page: " & Err.description)
End Function








Private Function GenerateHomePage() As String
    On Error GoTo ErrorHandler

    ' Ensure objects are initialized
    If routes Is Nothing Then
        Set routes = New Collection
        Call InitializeAppLaunch
    End If
    
    If m_appConfig Is Nothing Then
        Set m_appConfig = LoadDefaultApps()
    End If

    Dim html As String
    Dim stardate As String
    
    ' Calculate stardate for LCARS feel
    stardate = "2025." & format(Now, "ddd.dd.hh")
    
    html = GenerateLCARSHTMLHeader("SmartTraffic Control Center")
    html = html & "<div class='container'>"
    html = html & "<div class='bar'></div>"
    html = html & "<h1 class='header'>LCARS - SMARTTRAFFIC CONTROL</h1>"
    html = html & "<div class='subheader'>Starfleet Command &bull; System Control &bull; Stardate " & stardate & "</div>"

    ' Server Status Section (Live)
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>System Status</div>"
    html = html & "<div id='statusContainer'>"
    html = html & "<p>HTTP Server: " & IIf(httpRunning, "<span style='color:#99FF99'>ONLINE</span> - Port " & httpPortNum, "<span style='color:#FF9999'>OFFLINE</span>") & "</p>"
    html = html & "<p>Uptime: " & GetUptime() & "</p>"
    html = html & "<p>Total Requests: " & m_totalRequests & "</p>"
    html = html & "<p>Active Connections: " & IIf(httpRunning, HttpServer.GetHTTPCount(), 0) & "</p>"
    html = html & "</div>"
    html = html & "</div>"

    ' Main Control Panels Section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Control Panels</div>"
    html = html & "<div class='panel-grid'>"
    
    ' Outlook Panel
    html = html & "<div class='control-panel outlook-panel'>"
    html = html & "<div class='panel-header'>OUTLOOK CONTROL</div>"
    html = html & "<div class='panel-content'>"
    html = html & "<p>Email Management & Monitoring</p>"
    html = html & "<a href='/outlook' class='btn panel-btn'>ACCESS</a>"
    html = html & "</div>"
    html = html & "</div>"
    
    ' Dashboard Panel
    html = html & "<div class='control-panel dashboard-panel'>"
    html = html & "<div class='panel-header'>SYSTEM DASHBOARD</div>"
    html = html & "<div class='panel-content'>"
    html = html & "<p>Server Monitoring & Statistics</p>"
    html = html & "<a href='/dashboard' class='btn panel-btn'>ACCESS</a>"
    html = html & "</div>"
    html = html & "</div>"
    
    ' Applications Panel
    html = html & "<div class='control-panel apps-panel'>"
    html = html & "<div class='panel-header'>APPLICATION LAUNCHER</div>"
    html = html & "<div class='panel-content'>"
    html = html & "<p>Quick Launch Applications</p>"
    html = html & "<a href='/apps' class='btn panel-btn'>ACCESS</a>"
    html = html & "</div>"
    html = html & "</div>"
    
    ' Configuration Panel
    html = html & "<div class='control-panel config-panel'>"
    html = html & "<div class='panel-header'>SYSTEM CONFIG</div>"
    html = html & "<div class='panel-content'>"
    html = html & "<p>System Configuration & Settings</p>"
    html = html & "<a href='/config' class='btn panel-btn'>ACCESS</a>"
    html = html & "</div>"
    html = html & "</div>"
    
    html = html & "</div>" ' End panel-grid
    html = html & "</div>" ' End section

    ' Quick Launch Applications Section
    If Not m_appConfig Is Nothing And m_appConfig.count > 0 Then
        html = html & "<div class='section'>"
        html = html & "<div class='section-title'>Quick Launch</div>"
        html = html & "<div class='quick-launch-grid'>"
        
        Dim appKey As Variant
        Dim appCount As Long: appCount = 0
        For Each appKey In m_appConfig.Keys
            If appCount < 8 Then ' Limit to 8 for display
                html = html & "<a href='/launch/" & appKey & "' class='quick-launch-btn'>" & UCase(Left(appKey, 1)) & Mid(appKey, 2) & "</a>"
                appCount = appCount + 1
            End If
        Next appKey
        
        html = html & "</div>"
        html = html & "</div>"
    End If

    ' Available Routes Section (Collapsed by default)
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Available Endpoints <span class='toggle' onclick='toggleRoutes()'>?</span></div>"
    html = html & "<div id='routesContainer' style='display:none;'>"
    
    If Not routes Is Nothing And routes.count > 0 Then
        Dim route As Variant
        For Each route In routes
            html = html & "<div class='route-item'>"
            html = html & "<a href='" & route(0) & "' class='route-link'>" & route(0) & "</a>"
            html = html & "<span class='route-handler'>" & route(1) & "</span>"
            html = html & "</div>"
        Next route
    Else
        html = html & "<p>No routes available</p>"
    End If
    
    html = html & "</div>"
    html = html & "</div>"

    html = html & "<div class='bar'></div>"

    ' JavaScript for live updating and toggles
    html = html & "<script>"
    html = html & "function updateStatus() {"
    html = html & "  var xhr = new XMLHttpRequest();"
    html = html & "  xhr.onreadystatechange = function() {"
    html = html & "    if (xhr.readyState == 4 && xhr.status == 200) {"
    html = html & "      document.getElementById('statusContainer').innerHTML = xhr.responseText;"
    html = html & "    }"
    html = html & "  };"
    html = html & "  xhr.open('GET','/status?snippet=1',true);"
    html = html & "  xhr.send();"
    html = html & "}"
    html = html & "function toggleRoutes() {"
    html = html & "  var container = document.getElementById('routesContainer');"
    html = html & "  var toggle = document.querySelector('.toggle');"
    html = html & "  if (container.style.display === 'none') {"
    html = html & "    container.style.display = 'block';"
    html = html & "    toggle.innerHTML = '?';"
    html = html & "  } else {"
    html = html & "    container.style.display = 'none';"
    html = html & "    toggle.innerHTML = '?';"
    html = html & "  }"
    html = html & "}"
    html = html & "setInterval(updateStatus, 5000);" ' Update every 5 seconds
    html = html & "</script>"

    html = html & "</div></body></html>"

    GenerateHomePage = html
    Exit Function

ErrorHandler:
    DebugLog "Error in GenerateHomePage: " & Err.description & " (Line: " & Erl & ")"
    
    ' Create a simple fallback page
    Dim fallbackHtml As String
    fallbackHtml = GenerateLCARSHTMLHeader("SmartTraffic - Error Recovery")
    fallbackHtml = fallbackHtml & "<div class='container'>"
    fallbackHtml = fallbackHtml & "<h1 class='header'>SYSTEM RECOVERY MODE</h1>"
    fallbackHtml = fallbackHtml & "<div class='section'>"
    fallbackHtml = fallbackHtml & "<div class='section-title'>Error Details</div>"
    fallbackHtml = fallbackHtml & "<p>Error: " & Err.description & "</p>"
    fallbackHtml = fallbackHtml & "<p>Attempting to reinitialize system...</p>"
    fallbackHtml = fallbackHtml & "<a href='/dashboard' class='btn'>System Dashboard</a>"
    fallbackHtml = fallbackHtml & "<a href='/status' class='btn'>System Status</a>"
    fallbackHtml = fallbackHtml & "</div>"
    fallbackHtml = fallbackHtml & "</div></body></html>"
    
    GenerateHomePage = fallbackHtml
End Function

' Enhanced LCARS HTML Header with better styling
Private Function GenerateLCARSHTMLHeader(ByVal title As String) As String
    Dim html As String
    html = "<!DOCTYPE html><html><head><title>" & title & "</title>"
    html = html & "<meta name='viewport' content='width=device-width, initial-scale=1'>"
    html = html & "<style>"
    html = html & "body { background: #000; color: #FF9966; font-family: 'OCR A Extended', 'Courier New', monospace; padding: 20px; margin: 0; }"
    html = html & ".container { max-width: 1400px; margin: 0 auto; }"
    html = html & ".bar { height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate; border-radius: 20px 5px 20px 5px; }"
    html = html & "@keyframes flash { from { opacity: 0.6; } to { opacity: 1; } }"
    html = html & ".btn { padding: 12px 20px; background: #CC6600; color: #000; font-weight: bold; border-radius: 20px 5px 20px 5px; cursor: pointer; border: 2px solid #FFFF99; display: inline-block; margin: 5px; text-decoration: none; transition: all 0.3s; }"
    html = html & ".btn:hover { background: #FF9966; border-color: #99CCFF; transform: translateY(-2px); box-shadow: 0 4px 8px rgba(255,153,102,0.3); }"
    html = html & ".header { font-size: 42px; color: #99CCFF; text-shadow: 0 0 15px #99CCFF; margin-bottom: 10px; text-align: center; font-weight: bold; }"
    html = html & ".subheader { font-size: 16px; color: #FFFF99; margin: 10px 0; text-align: center; text-transform: uppercase; }"
    html = html & ".section { margin: 20px 0; padding: 20px; border: 3px solid #663399; border-radius: 20px 5px 20px 5px; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); box-shadow: inset 0 0 20px rgba(102,51,153,0.3); }"
    html = html & ".section-title { font-size: 24px; color: #99CCFF; text-transform: uppercase; margin-bottom: 15px; border-bottom: 2px solid #663399; padding-bottom: 10px; text-shadow: 0 0 5px #99CCFF; }"
    
    ' Panel grid styling
    html = html & ".panel-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-top: 20px; }"
    html = html & ".control-panel { border: 2px solid; border-radius: 15px 5px 15px 5px; padding: 20px; background: linear-gradient(145deg, #2a2a3e, #1e1e2e); transition: all 0.3s; }"
    html = html & ".control-panel:hover { transform: translateY(-5px); box-shadow: 0 8px 16px rgba(0,0,0,0.4); }"
    html = html & ".outlook-panel { border-color: #99CCFF; }"
    html = html & ".dashboard-panel { border-color: #99FF99; }"
    html = html & ".apps-panel { border-color: #FFFF99; }"
    html = html & ".config-panel { border-color: #FF9999; }"
    html = html & ".panel-header { font-size: 18px; font-weight: bold; color: inherit; text-align: center; margin-bottom: 15px; text-transform: uppercase; }"
    html = html & ".panel-content { text-align: center; }"
    html = html & ".panel-btn { width: 100%; margin-top: 10px; background: inherit; border-color: inherit; }"
    
    ' Quick launch styling
    html = html & ".quick-launch-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 10px; margin-top: 15px; }"
    html = html & ".quick-launch-btn { padding: 8px 12px; background: #663399; color: #FFFF99; border-radius: 10px; text-decoration: none; text-align: center; border: 1px solid #FFFF99; transition: all 0.2s; }"
    html = html & ".quick-launch-btn:hover { background: #7744AA; transform: scale(1.05); }"
    
    ' Route display styling
    html = html & ".route-item { margin: 5px 0; padding: 5px; border-left: 3px solid #663399; background: rgba(102,51,153,0.1); }"
    html = html & ".route-link { color: #99CCFF; text-decoration: none; margin-right: 20px; }"
    html = html & ".route-handler { color: #FFFF99; font-size: 12px; }"
    html = html & ".toggle { cursor: pointer; float: right; color: #FF9966; }"
    
    ' Status styling
    html = html & "#statusContainer p { margin: 8px 0; padding: 5px; background: rgba(0,0,0,0.3); border-radius: 5px; }"
    
    html = html & "code { background: #333; padding: 2px 5px; border-radius: 3px; color: #99CCFF; }"
    html = html & "p { line-height: 1.6; margin: 10px 0; }"
    html = html & "</style></head><body>"
    GenerateLCARSHTMLHeader = html
End Function



















' --- Generate HTML snippet for live server stats (used in home page) ---
Public Function GenerateStatusSnippet() As String
    On Error GoTo ErrorHandler
    
    Dim html As String
    html = "<p>HTTP Server: " & IIf(httpRunning, "Running on port " & httpPortNum, "Stopped") & "</p>"
    html = html & "<p>Uptime: " & GetUptime() & "</p>"
    html = html & "<p>Total Requests: " & m_totalRequests & "</p>"
    html = html & "<p>Last Activity: " & IIf(m_lastActivity > 0, format(m_lastActivity, "yyyy-mm-dd hh:mm:ss"), "None") & "</p>"
    
    GenerateStatusSnippet = html
    Exit Function
    
ErrorHandler:
    DebugLog "Error in GenerateStatusSnippet: " & Err.description
    GenerateStatusSnippet = "<p>Error fetching server status</p>"
End Function


' --- Generate HTML Header ---
Private Function GenerateHTMLHeader(ByVal title As String) As String
    Dim html As String
    html = "<!DOCTYPE html><html><head><title>" & title & "</title>"
    html = html & "<meta name='viewport' content='width=device-width, initial-scale=1'>"
    html = html & "<style>"
    html = html & "body { background: #000; color: #FF9966; font-family: 'Courier New', monospace; padding: 20px; margin: 0; }"
    html = html & ".container { max-width: 1200px; margin: 0 auto; }"
    html = html & ".bar { height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate; border-radius: 5px; }"
    html = html & "@keyframes flash { from { opacity: 0.6; } to { opacity: 1; } }"
    html = html & ".btn { padding: 12px 20px; background: #CC6600; color: #000; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99; display: inline-block; margin: 5px; text-decoration: none; transition: all 0.3s; }"
    html = html & ".btn:hover { background: #FF9966; border-color: #99CCFF; transform: translateY(-2px); }"
    html = html & ".header { font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 30px; text-align: center; }"
    html = html & ".section { margin: 20px 0; padding: 20px; border: 2px solid #663399; border-radius: 10px; background: linear-gradient(135deg, #1a1a2e, #16213e); }"
    html = html & ".section-title { font-size: 24px; color: #99CCFF; text-transform: uppercase; margin-bottom: 15px; border-bottom: 2px solid #663399; padding-bottom: 10px; }"
    html = html & "code { background: #333; padding: 2px 5px; border-radius: 3px; color: #99CCFF; }"
    html = html & "p { line-height: 1.6; margin: 10px 0; }"
    html = html & "</style></head><body>"
    GenerateHTMLHeader = html
End Function

' --- Helper Functions ---
Private Function GetAvailablePaths() As String
    If routes Is Nothing Then
        GetAvailablePaths = "No routes available"
        Exit Function
    End If
    
    Dim paths As String
    Dim route As Variant
    Dim count As Integer
    
    For Each route In routes
        If count > 0 Then paths = paths & ", "
        paths = paths & route(0)
        count = count + 1
        If count >= 10 Then
            paths = paths & "..."
            Exit For
        End If
    Next route
    
    GetAvailablePaths = paths
End Function

Private Function GetUptime() As String
    If m_serverStartTime = 0 Then
        GetUptime = "Unknown"
    Else
        Dim uptime As Double
        uptime = Now - m_serverStartTime
        GetUptime = format(uptime, "dd") & "d " & format(uptime, "hh:mm:ss")
    End If
End Function

Private Function IsModuleAvailable(ByVal moduleName As String) As Boolean
    On Error Resume Next
    Select Case LCase(moduleName)
        Case "outlookwebui"
            IsModuleAvailable = (TypeName(OutlookWebUI) <> "Nothing")
        Case "govee"
            IsModuleAvailable = (TypeName(govee) <> "Nothing")
        Case Else
            IsModuleAvailable = False
    End Select
    On Error GoTo 0
End Function

' --- Generate Error Page ---
Private Function GenerateErrorPage(ByVal message As String, Optional ByVal details As String = "") As String
    Dim html As String
    html = GenerateHTMLHeader("Error")
    html = html & "<div class='container'>"
    html = html & "<h1 class='header'>?? Error</h1>"
    html = html & "<div class='section'>"
    html = html & "<p><strong>Error:</strong> " & HTMLEncode(message) & "</p>"
    If details <> "" Then
        html = html & "<p><strong>Details:</strong> " & HTMLEncode(details) & "</p>"
    End If
    html = html & "<a href='/' class='btn'>?? Home</a>"
    html = html & "<a href='/dashboard' class='btn'>?? Dashboard</a>"
    html = html & "</div>"
    html = html & "</div></body></html>"
    GenerateErrorPage = html
End Function

' --- Generate Launch Success Page ---
Private Function GenerateLaunchSuccessPage(ByVal appName As String) As String
    Dim html As String
    html = GenerateHTMLHeader("Launch Success")
    html = html & "<div class='container'>"
    html = html & "<h1 class='header'>?? Launch Successful</h1>"
    html = html & "<div class='section'>"
    html = html & "<p><strong>Application:</strong> " & HTMLEncode(appName) & "</p>"
    html = html & "<p><strong>Launch Time:</strong> " & format(Now, "yyyy-mm-dd hh:mm:ss") & "</p>"
    html = html & "<a href='/' class='btn'>?? Home</a>"
    html = html & "<a href='/dashboard' class='btn'>?? Dashboard</a>"
    html = html & "</div>"
    html = html & "</div></body></html>"
    GenerateLaunchSuccessPage = html
End Function

' --- HTML Encode ---
Private Function HTMLEncode(ByVal Text As String) As String
    HTMLEncode = Replace(Replace(Replace(Replace(Text, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function

' --- Load App Config ---
Public Function LoadAppConfig() As Object
    On Error GoTo ErrorHandler
    
    Dim fso As Object, file As Object, jsonText As String
    Dim appsDict As Object
    Set appsDict = CreateObject("Scripting.Dictionary")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(CONFIG_FILE) Then
        Set file = fso.OpenTextFile(CONFIG_FILE, 1)
        jsonText = file.ReadAll
        file.Close
        
        ' Simple JSON parsing for app config
        ' This is a simplified version - you may need a proper JSON parser
        Set appsDict = LoadDefaultApps
        DebugLog "Config file found but using defaults (JSON parsing not implemented)"
    Else
        DebugLog "Config file not found, loading defaults"
        Set appsDict = LoadDefaultApps
    End If
    
    Set LoadAppConfig = appsDict
    Exit Function

ErrorHandler:
    DebugLog "Error loading app config: " & Err.description
    Set LoadAppConfig = LoadDefaultApps
End Function

' --- Load Default Apps ---
Public Function LoadDefaultApps() As Object
    Dim appsDict As Object
    Set appsDict = CreateObject("Scripting.Dictionary")
    appsDict.Add "notepad", "notepad.exe"
    appsDict.Add "calculator", "calc.exe"
    appsDict.Add "explorer", "explorer.exe"
    appsDict.Add "cmd", "cmd.exe"
    appsDict.Add "outlook", "outlook.exe"
    DebugLog "Loaded default apps: " & appsDict.count & " applications"
    Set LoadDefaultApps = appsDict
End Function

' --- Launch Application ---
Private Sub LaunchApplication(ByVal appName As String)
    On Error GoTo ErrorHandler
    
    Dim WShell As Object
    Set WShell = CreateObject("WScript.Shell")
    
    If m_appConfig.exists(appName) Then
        WShell.Run m_appConfig(appName)
        DebugLog "Launched application: " & appName
    Else
        DebugLog "Application not found: " & appName
    End If
    Exit Sub
    
ErrorHandler:
    DebugLog "Error launching application " & appName & ": " & Err.description
End Sub

' --- Setup AppLauncher ---
Public Function SetupAppLauncher(ByVal portNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    httpPortNum = portNum
    InitializeModuleVariables
    Set m_appConfig = LoadAppConfig
    
    ' Initialize routes
    InitializeAppLaunch
    
    ' Start HTTP server
    If HttpServer.StartHttpServer(httpPortNum) Then
        httpRunning = True
        isRunning = True
        m_serverStartTime = Now
        m_totalRequests = 0
        DebugLog "AppLauncher setup complete on port " & httpPortNum
        SetupAppLauncher = True
    Else
        httpRunning = False
        isRunning = False
        DebugLog "Failed to start HTTP server"
        SetupAppLauncher = False
    End If
    
    Exit Function

ErrorHandler:
    DebugLog "Error in SetupAppLauncher: " & Err.description
    SetupAppLauncher = False
    httpRunning = False
    isRunning = False
End Function

' --- Launch Browser ---
Public Function LaunchBrowser(ByVal url As String) As Boolean
    On Error GoTo ErrorHandler
    If Not isRunning Then
        DebugLog "Cannot launch browser: AppLauncher not running"
        LaunchBrowser = False
        Exit Function
    End If
    Call shell("cmd /c start " & url, vbHide)
    DebugLog "Browser launched to " & url
    LaunchBrowser = True
    Exit Function
ErrorHandler:
    DebugLog "Error in LaunchBrowser: " & Err.description
    LaunchBrowser = False
End Function

' --- Stop AppLauncher ---
Public Sub StopAppLauncher()
    On Error GoTo ErrorHandler
    
    HttpServer.StopHttpServer
    httpRunning = False
    isRunning = False
    m_totalRequests = 0
    m_lastActivity = Empty
    m_serverStartTime = Empty
    Set m_appConfig = Nothing
    Set routes = Nothing
    DebugLog "AppLauncher stopped and resources cleared"
    
    Exit Sub
    
ErrorHandler:
    DebugLog "Error stopping AppLauncher: " & Err.description
End Sub

' --- Process Pending Requests ---
Public Sub ProcessRequests()
    On Error GoTo ErrorHandler
    
    If httpRunning Then
        Call HttpServer.ProcessHttpServer
    End If
    
    Exit Sub
ErrorHandler:
    DebugLog "Error in ProcessRequests: " & Err.description
End Sub

' --- Add App ---
Public Sub AddApp(appName As String, Optional exePath As String = "")
    If m_appConfig Is Nothing Then
        Set m_appConfig = CreateObject("Scripting.Dictionary")
    End If
    
    If exePath = "" Then
        m_appConfig(appName) = appName
    Else
        m_appConfig(appName) = exePath
    End If
    DebugLog "Added app: " & appName & " -> " & IIf(exePath = "", appName, exePath)
End Sub

' --- Status Getters ---
Public Function GetAppLauncherStatus() As Boolean
    GetAppLauncherStatus = httpRunning And isRunning
End Function

Public Function GetAppLauncherPort() As Long
    GetAppLauncherPort = httpPortNum
End Function

Public Function GetAppLauncherClientCount() As Long
    GetAppLauncherClientCount = HttpServer.GetHTTPCount()
End Function

Public Function GetAppLauncherStats() As String
    GetAppLauncherStats = "Total Requests: " & m_totalRequests & ", Last Activity: " & format(m_lastActivity, "yyyy-mm-dd hh:mm:ss")
End Function

' --- Debug Logging ---
Private Sub DebugLog(ByVal message As String)
    Debug.Print "[" & format(Now, "yyyy-mm-dd hh:mm:ss") & "] AppLaunch: " & message
End Sub





Public Sub CreateLCARSDashboard()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim frmName As String
    Dim i As Long
    
    frmName = "frmLCARSDashboard"
    
    ' --- Get workbook VBProject ---
    Set vbProj = ThisWorkbook.VBProject
    
    ' --- Delete existing form if present ---
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(frmName)
    On Error GoTo 0
    
    ' --- Add new UserForm ---
    Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm = 3
    vbComp.Name = frmName
    
    ' --- Add basic initialization code ---
    With vbComp.CodeModule
        .InsertLines 1, "Option Explicit"
        .InsertLines 2, ""
        .InsertLines 3, "Private Sub UserForm_Initialize()"
        .InsertLines 4, "    Me.BackColor = RGB(0,0,0)"
        .InsertLines 5, "End Sub"
    End With
    
    ' --- Add STOP button dynamically ---
    Dim btn As Object
    Set btn = vbComp.Designer.Controls.Add("Forms.CommandButton.1", "btnStop")
    With btn
        .Caption = "STOP"
        .Left = 20
        .Top = 20
        .width = 60
        .height = 30
        .BackColor = rgb(255, 102, 0)
        .ForeColor = rgb(0, 0, 0)
    End With
    
    ' --- Show the form ---
    Set dashboardForm = VBA.UserForms.Add(frmName)
    dashboardForm.Show vbModeless
    
    ' --- Initialize starfield ---
    starCount = 100
    ReDim starShapes(1 To starCount)
    ReDim starX(1 To starCount)
    ReDim starY(1 To starCount)
    ReDim starSpeed(1 To starCount)
    
    For i = 1 To starCount
        starX(i) = Rnd * dashboardForm.width
        starY(i) = Rnd * dashboardForm.height
        starSpeed(i) = 1 + Rnd * 3
        
        Set starShapes(i) = dashboardForm.Controls.Add("Forms.Label.1", "star" & i)
        With starShapes(i)
            .Caption = ""
            .BackColor = rgb(255, 255, 255)
            .width = 2
            .height = 2
            .Left = starX(i)
            .Top = starY(i)
        End With
    Next i
    
    ' --- Start animation ---
    animRunning = True
    AnimateStarfield
End Sub

Private Sub AnimateStarfield()
    Dim i As Long
    Do While animRunning
        For i = 1 To starCount
            starY(i) = starY(i) + starSpeed(i)
            If starY(i) > dashboardForm.height Then
                starY(i) = 0
                starX(i) = Rnd * dashboardForm.width
            End If
            starShapes(i).Top = starY(i)
            starShapes(i).Left = starX(i)
        Next i
        DoEvents
        Application.Wait Now + TimeValue("0:00:00.02")
    Loop
End Sub

Public Sub StopDashboardAnimation()
    animRunning = False
    If Not dashboardForm Is Nothing Then
        Unload dashboardForm
    End If
End Sub

