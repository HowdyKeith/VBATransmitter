Option Explicit

'***************************************************************
' OutlookWebHandler Module (Updated - No OLE blocking)
' Purpose: Handle web requests for Outlook data using external
'          VBS scripts to prevent Excel from hanging
'***************************************************************

Public Function GetRoutes() As Collection
    Dim routes As New Collection
    routes.Add Array("/outlook", "GenerateOutlookDashboard"), "outlook"
    routes.Add Array("/outlook/unread", "GetUnreadCountAPI"), "unread"
    routes.Add Array("/outlook/rules", "GenerateRulesPage"), "rules"
    routes.Add Array("/outlook/run_rule", "ExecuteRuleAPI"), "run_rule"
    routes.Add Array("/outlook/status", "GetOutlookStatusAPI"), "status"
    routes.Add Array("/outlook/launch", "LaunchOutlookAPI"), "launch"
    Set GetRoutes = routes
End Function

'=======================================================
' MAIN DASHBOARD PAGE
'=======================================================
Public Function GenerateOutlookDashboard() As String
    Dim html As String
    Dim unreadInfo As String
    Dim isRunning As Boolean
    
    ' Check Outlook status without OLE
    isRunning = OutlookExternal.IsOutlookRunningExternal()
    
    html = "<!DOCTYPE html><html><head><title>Outlook Dashboard</title>"
    html = html & GetLCARSStyles()
    html = html & GetJavaScript()
    html = html & "</head><body>"
    
    html = html & "<div class='container'>"
    html = html & "<div class='bar'></div>"
    html = html & "<h1 class='header'>LCARS - Outlook Command Center</h1>"
    html = html & "<div class='subheader'>Starfleet Command &bull; Outlook Interface &bull; Stardate " & format(Now, "yyyy.ddd.hh") & "</div>"
    
    ' Status Section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>System Status</div>"
    html = html & "<div id='outlook-status' class='" & IIf(isRunning, "status-online", "status-offline") & "'>"
    html = html & "Outlook: " & IIf(isRunning, "ONLINE", "OFFLINE")
    html = html & "</div>"
    html = html & "</div>"
    
    ' Unread Count Section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Inbox Status</div>"
    html = html & "<div id='unread-count' class='data-display'>Loading...</div>"
    html = html & "<button class='btn' onclick='refreshUnreadCount()'>Refresh Count</button>"
    html = html & "</div>"
    
    ' Action Buttons
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Actions</div>"
    If Not isRunning Then
        html = html & "<button class='btn btn-launch' onclick='launchOutlook()'>Launch Outlook</button>"
    End If
    html = html & "<a href='/outlook/rules' class='btn'>Manage Rules</a>"
    html = html & "<button class='btn' onclick='refreshStatus()'>Refresh Status</button>"
    html = html & "</div>"
    
    ' Log Section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Activity Log</div>"
    html = html & "<div id='activity-log' class='log-display'></div>"
    html = html & "</div>"
    
    html = html & "<div class='bar'></div>"
    html = html & "<a href='/index.html' class='btn'>Return to Home</a>"
    html = html & "</div></body></html>"
    
    GenerateOutlookDashboard = html
End Function

'=======================================================
' API ENDPOINTS
'=======================================================

' API: Get unread count (JSON)
Public Function GetUnreadCountAPI() As String
    GetUnreadCountAPI = OutlookExternal.GetUnreadCountForWeb()
End Function

' API: Execute a rule (JSON)
Public Function ExecuteRuleAPI() As String
    ' This would be called with a rule name parameter
    ' For now, return generic response
    ExecuteRuleAPI = "{""status"":""error"",""message"":""Rule name required""}"
End Function

' API: Get Outlook status (JSON)
Public Function GetOutlookStatusAPI() As String
    Dim isRunning As Boolean
    isRunning = OutlookExternal.IsOutlookRunningExternal()
    
    GetOutlookStatusAPI = "{""status"":""success"",""outlook_running"":" & _
                         IIf(isRunning, "true", "false") & _
                         ",""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
End Function

' API: Launch Outlook (JSON)
Public Function LaunchOutlookAPI() As String
    On Error GoTo ErrorHandler
    
    OutlookExternal.LaunchOutlookExternal
    LaunchOutlookAPI = "{""status"":""success"",""message"":""Outlook launch initiated"",""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
    Exit Function
    
ErrorHandler:
    LaunchOutlookAPI = "{""status"":""error"",""message"":""" & Err.description & """,""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
End Function

'=======================================================
' RULES PAGE
'=======================================================
Public Function GenerateRulesPage() As String
    Dim html As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets("Outlook")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    html = "<!DOCTYPE html><html><head><title>Outlook Rules</title>"
    html = html & GetLCARSStyles()
    html = html & GetRulesJavaScript()
    html = html & "</head><body>"
    
    html = html & "<div class='container'>"
    html = html & "<div class='bar'></div>"
    html = html & "<h1 class='header'>LCARS - Outlook Rules</h1>"
    html = html & "<div class='subheader'>Rule Management Interface</div>"
    
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Available Rules</div>"
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).value = True Then ' Enabled rules
            Dim ruleName As String
            ruleName = ws.Cells(i, 1).value
            html = html & "<div class='rule-item'>"
            html = html & "<span class='rule-name'>" & ruleName & "</span>"
            html = html & "<button class='btn btn-small' onclick='executeRule(""" & ruleName & """)'>Execute</button>"
            html = html & "</div>"
        End If
    Next i
    
    html = html & "</div>"
    
    html = html & "<div id='rule-status' class='section' style='display:none;'>"
    html = html & "<div class='section-title'>Execution Status</div>"
    html = html & "<div id='rule-result'></div>"
    html = html & "</div>"
    
    html = html & "<div class='bar'></div>"
    html = html & "<a href='/outlook' class='btn'>Back to Dashboard</a>"
    html = html & "</div></body></html>"
    
    GenerateRulesPage = html
    Exit Function
    
ErrorHandler:
    GenerateRulesPage = "<html><body><h1>Error</h1><p>Error loading rules: " & Err.description & "</p></body></html>"
End Function

'=======================================================
' STYLES AND SCRIPTS
'=======================================================

Private Function GetLCARSStyles() As String
    GetLCARSStyles = "<style>" & _
        "body { background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; padding: 20px; margin: 0; }" & _
        ".container { max-width: 1200px; margin: auto; }" & _
        ".bar { height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate; }" & _
        "@keyframes flash { from { opacity: 0.6; } to { opacity: 1; } }" & _
        ".btn { padding: 10px 15px; background: #CC6600; color: black; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99; display: inline-block; margin: 5px; text-decoration: none; font-size: 14px; }" & _
        ".btn:hover { background: #FF9966; border-color: #99CCFF; }" & _
        ".btn-small { padding: 5px 10px; font-size: 12px; }" & _
        ".btn-launch { background: #009900; border-color: #00FF00; }" & _
        ".header { font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px; }" & _
        ".subheader { font-size: 18px; color: #FFFF99; margin: 10px 0; }" & _
        ".section { margin: 15px 0; padding: 15px; border: 2px solid #663399; border-radius: 10px; background: #1C2526; }" & _
        ".section-title { font-size: 24px; color: #99CCFF; text-transform: uppercase; margin-bottom: 10px; }" & _
        ".status-online { color: #00FF00; font-size: 20px; font-weight: bold; }" & _
        ".status-offline { color: #FF3300; font-size: 20px; font-weight: bold; }" & _
        ".data-display { font-size: 24px; color: #FFFF99; padding: 10px; background: #2A3132; border-radius: 5px; margin: 10px 0; }" & _
        ".log-display { background: #2A3132; padding: 10px; border-radius: 5px; height: 200px; overflow-y: auto; font-family: monospace; font-size: 12px; }" & _
        ".rule-item { display: flex; justify-content: space-between; align-items: center; padding: 8px; margin: 5px 0; background: #2A3132; border-radius: 5px; }" & _
        ".rule-name { flex-grow: 1; color: #FFFF99; }" & _
        "</style>"
End Function

Private Function GetJavaScript() As String
    Dim js As String
    
    js = "<script>"
    js = js & "function refreshUnreadCount() {"
    js = js & "  document.getElementById('unread-count').innerHTML = 'Loading...';"
    js = js & "  fetch('/outlook/unread')"
    js = js & "    .then(response => response.json())"
    js = js & "    .then(data => {"
    js = js & "      if (data.status === 'success') {"
    js = js & "        document.getElementById('unread-count').innerHTML = 'Unread Messages: ' + data.unread_count;"
    js = js & "        logActivity('Unread count refreshed: ' + data.unread_count);"
    js = js & "      } else {"
    js = js & "        document.getElementById('unread-count').innerHTML = 'Error: ' + data.message;"
    js = js & "        logActivity('Error getting unread count: ' + data.message);"
    js = js & "      }"
    js = js & "    })"
    js = js & "    .catch(err => {"
    js = js & "      document.getElementById('unread-count').innerHTML = 'Network Error';"
    js = js & "      logActivity('Network error: ' + err);"
    js = js & "    });"
    js = js & "}"
    
    js = js & "function refreshStatus() {"
    js = js & "  fetch('/outlook/status')"
    js = js & "    .then(response => response.json())"
    js = js & "    .then(data => {"
    
    GetJavaScript = js & GetJavaScriptPart2()
End Function

Private Function GetJavaScriptPart2() As String
    Dim js As String
    
    js = "      const statusEl = document.getElementById('outlook-status');"
    js = js & "      if (data.outlook_running) {"
    js = js & "        statusEl.className = 'status-online';"
    js = js & "        statusEl.innerHTML = 'Outlook: ONLINE';"
    js = js & "        logActivity('Outlook status: ONLINE');"
    js = js & "      } else {"
    js = js & "        statusEl.className = 'status-offline';"
    js = js & "        statusEl.innerHTML = 'Outlook: OFFLINE';"
    js = js & "        logActivity('Outlook status: OFFLINE');"
    js = js & "      }"
    js = js & "      location.reload();"
    js = js & "    });"
    js = js & "}"
    
    js = js & "function launchOutlook() {"
    js = js & "  logActivity('Launching Outlook...');"
    js = js & "  fetch('/outlook/launch')"
    js = js & "    .then(response => response.json())"
    js = js & "    .then(data => {"
    js = js & "      logActivity(data.message);"
    js = js & "      setTimeout(() => { refreshStatus(); }, 3000);"
    js = js & "    })"
    
    GetJavaScriptPart2 = js & GetJavaScriptPart3()
End Function

Private Function GetJavaScriptPart3() As String
    Dim js As String
    
    js = "    .catch(err => logActivity('Launch error: ' + err));"
    js = js & "}"
    js = js & "function logActivity(message) {"
    js = js & "  const log = document.getElementById('activity-log');"
    js = js & "  const timestamp = new Date().toLocaleTimeString();"
    js = js & "  log.innerHTML = '[' + timestamp + '] ' + message + '<br>' + log.innerHTML;"
    js = js & "}"
    js = js & "setInterval(refreshUnreadCount, 30000);"
    js = js & "window.onload = function() { refreshUnreadCount(); };"
    js = js & "</script>"
    
    GetJavaScriptPart3 = js
End Function

Private Function GetRulesJavaScript() As String
    Dim js As String
    
    js = "<script>"
    js = js & "function executeRule(ruleName) {"
    js = js & "  document.getElementById('rule-status').style.display = 'block';"
    js = js & "  document.getElementById('rule-result').innerHTML = 'Executing rule: ' + ruleName + '...';"
    js = js & "  fetch('/outlook/execute_rule?rule=' + encodeURIComponent(ruleName))"
    js = js & "    .then(response => response.json())"
    js = js & "    .then(data => {"
    js = js & "      const resultEl = document.getElementById('rule-result');"
    js = js & "      if (data.status === 'success') {"
    js = js & "        resultEl.innerHTML = '<span style=\"color: #00FF00\">SUCCESS:</span> ' + data.message;"
    js = js & "      } else {"
    js = js & "        resultEl.innerHTML = '<span style=\"color: #FF3300\">ERROR:</span> ' + data.message;"
    js = js & "      }"
    js = js & "    })"
    js = js & "    .catch(err => {"
    js = js & "      document.getElementById('rule-result').innerHTML = '<span style=\"color: #FF3300\">Network Error:</span> ' + err;"
    js = js & "    });"
    js = js & "}"
    js = js & "</script>"
    
    GetRulesJavaScript = js
End Function

Public Function HandleRuleExecution(ByVal requestPath As String) As String
    On Error GoTo ErrorHandler
    
    Dim ruleName As String
    ruleName = ExtractQueryParameter(requestPath, "rule")
    
    If ruleName = "" Then
        HandleRuleExecution = "{""status"":""error"",""message"":""No rule specified""}"
        Exit Function
    End If
    
    HandleRuleExecution = OutlookExternal.ExecuteRuleForWeb(ruleName)
    Exit Function
    
ErrorHandler:
    HandleRuleExecution = "{""status"":""error"",""message"":""" & Err.description & """}"
End Function

Private Function ExtractQueryParameter(ByVal url As String, ByVal paramName As String) As String
    On Error Resume Next
    
    If InStr(url, "?") = 0 Then Exit Function
    
    Dim queryString As String
    Dim pairs() As String
    Dim pair() As String
    Dim i As Long
    
    queryString = Mid(url, InStr(url, "?") + 1)
    pairs = Split(queryString, "&")
    
    For i = 0 To UBound(pairs)
        pair = Split(pairs(i), "=")
        If UBound(pair) >= 1 And LCase(pair(0)) = LCase(paramName) Then
            ExtractQueryParameter = DecodeURLString(pair(1))
            Exit Function
        End If
    Next i
End Function

Private Function DecodeURLString(ByVal str As String) As String
    On Error Resume Next
    str = Replace(str, "+", " ")
    str = Replace(str, "%20", " ")
    DecodeURLString = str
End Function
