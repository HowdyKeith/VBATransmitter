Option Explicit

'=======================================================
' Unified AppLaunchOutlook Module - Outlook Integration
' Handles requests, HTML UI, API JSON, and Outlook actions
'=======================================================

Private m_totalRequests As Long
Private m_lastActivity As Date

'=========================
' Main HTTP request handler
'=========================
Public Function HandleOutlookRequest(ByVal method As String, ByVal path As String, ByVal body As String) As String
    On Error GoTo ErrorHandler

    m_totalRequests = m_totalRequests + 1
    m_lastActivity = Now
    
    Dim response As String
    
    ' API vs UI routing
    Select Case path
        '--- UI Pages ---
        Case "/outlook"
            response = GenerateLCARSOutlookPage(False)
        Case "/outlook/enhanced"
            response = GenerateLCARSOutlookPage(True)
            
        '--- Standard actions ---
        Case "/outlook/run_rules"
            OutlookExtra.RunOutlookRules
            response = GenerateActionResponse("Outlook Rules Executed", "/outlook")
        Case "/outlook/force_check"
            Outlook.CheckOutlookBackground
            response = GenerateActionResponse("Forced Check - " & Outlook.GetUnreadMsgCount() & " unread", "/outlook")
        Case "/outlook/universal_check"
            If method = "POST" Then
                HandleUniversalEmailCheck body
                response = GenerateActionResponse("Universal Email Check Completed", "/outlook")
            Else
                response = GenerateActionResponse("Use POST method for universal check", "/outlook")
            End If
        Case "/outlook/create_vbs"
            Outlook.CreateOutlookVBSScript
            response = GenerateActionResponse("VBS Script Created", "/outlook")
        Case "/outlook/start_vbs"
            OutlookExtra.StartOutlookVBSChecker
            response = GenerateActionResponse("VBS Checker Started", "/outlook")
        Case "/outlook/stop_vbs"
            OutlookExtra.StopOutlookVBSChecker
            response = GenerateActionResponse("VBS Checker Stopped", "/outlook")
        Case "/outlook/enable_background"
            Outlook.UseBackgroundOutlookMode = True
            response = GenerateActionResponse("Background Mode Enabled", "/outlook")
        Case "/outlook/disable_background"
            Outlook.UseBackgroundOutlookMode = False
            response = GenerateActionResponse("Background Mode Disabled", "/outlook")
        Case "/outlook/stop_checking"
            Outlook.StopOutlookChecking
            OutlookExtra.StopOutlookVBSChecker
            response = GenerateActionResponse("All Outlook Processes Stopped", "/outlook")
        
        '--- API Endpoints ---
        Case "/outlook/api/status"
            response = GetOutlookStatusJSON()
        Case "/outlook/api/send"
            If method = "POST" Then
                response = SendOutlookMail(body)
            Else
                response = "{""error"":""POST required""}"
            End If
        Case Else
            response = GenerateActionResponse("Invalid Outlook Request: " & path, "/outlook")
    End Select
    
    DebugLog "Handled Outlook request: " & method & " " & path & " at " & format(Now, "yyyy-mm-dd hh:mm:ss")
    HandleOutlookRequest = response
    Exit Function

ErrorHandler:
    DebugLog "Error in HandleOutlookRequest: " & Err.description
    HandleOutlookRequest = GenerateErrorResponse("Error: " & Err.description, "/outlook")
End Function

'=========================
' Generate LCARS HTML
' enhanced=True for advanced UI
'=========================
Private Function GenerateLCARSOutlookPage(ByVal enhanced As Boolean) As String
    On Error GoTo ErrorHandler
    
    Dim html As String
    Dim statusJSON As String
    
    statusJSON = GetOutlookStatusJSON()
    
    html = "<!DOCTYPE html><html><head><title>Outlook Control</title>"
    html = html & "<meta http-equiv='refresh' content='30'>" ' auto-refresh
    html = html & "<style>"
    html = html & "body { background:#000;color:#FFA500;font-family:Arial,sans-serif; }"
    html = html & ".lcars-container{width:95%;margin:auto;border:2px solid #FFA500;border-radius:10px;padding:20px;}"
    html = html & ".lcars-bar{background:#FF4500;height:25px;margin-bottom:20px;border-radius:12px;}"
    html = html & ".lcars-button{background:#4169E1;color:#FFF;padding:12px 20px;margin:5px;border-radius:20px;text-decoration:none;display:inline-block;cursor:pointer;transition:all 0.3s;}"
    html = html & ".lcars-button:hover{background:#4682B4;transform:scale(1.05);}"
    html = html & ".lcars-button.success{background:#32CD32;}"
    html = html & ".lcars-button.danger{background:#DC143C;}"
    html = html & ".lcars-button.warning{background:#FF8C00;}"
    html = html & ".status-panel{background:#001122;border:2px solid #00BFFF;padding:15px;border-radius:8px;margin-bottom:10px;}"
    html = html & ".status-panel h3{color:#00BFFF;margin-top:0;}"
    html = html & ".status-value{color:#00FF00;font-weight:bold;}"
    html = html & ".control-section{margin:20px 0;padding:20px;background:#001100;border:1px solid #228B22;border-radius:8px;}"
    html = html & ".control-section h2{color:#32CD32;margin-top:0;}"
    html = html & "</style></head><body>"
    
    html = html & "<div class='lcars-container'>"
    html = html & "<div class='lcars-bar'></div>"
    html = html & "<h1>Outlook Control Interface</h1>"
    
    ' Status Panel
    html = html & "<div class='status-panel'>"
    html = html & "<h3>Email Status</h3>"
    html = html & "<p>Outlook Running: <span class='status-value'>" & IIf(Outlook.IsOutlookRunning(), "YES", "NO") & "</span></p>"
    html = html & "<p>Unread Emails: <span class='status-value'>" & Outlook.GetUnreadMsgCount() & "</span></p>"
    html = html & "<p>Background Mode: <span class='status-value'>" & IIf(Outlook.UseBackgroundOutlookMode, "ENABLED", "DISABLED") & "</span></p>"
    html = html & "<p>VBS Checker: <span class='status-value'>" & IIf(OutlookExtra.IsVBSCheckerRunning(), "RUNNING", "STOPPED") & "</span></p>"
    html = html & "<p>Total Requests: <span class='status-value'>" & m_totalRequests & "</span></p>"
    html = html & "<p>Last Activity: <span class='status-value'>" & format(m_lastActivity, "hh:mm:ss") & "</span></p>"
    html = html & "</div>"
    
    ' Control Buttons
    html = html & "<div class='control-section'>"
    html = html & "<h2>Email Operations</h2>"
    html = html & "<a href='/outlook/force_check' class='lcars-button success'>Force Check</a>"
    html = html & "<a href='/outlook/run_rules' class='lcars-button'>Run Rules</a>"
    html = html & "<a href='/outlook/universal_check' class='lcars-button warning'>Universal Check</a>"
    html = html & "</div>"
    
    html = html & "<div class='control-section'>"
    html = html & "<h2>System Control</h2>"
    html = html & "<a href='/outlook/create_vbs' class='lcars-button'>Create VBS</a>"
    html = html & "<a href='/outlook/start_vbs' class='lcars-button success'>Start VBS</a>"
    html = html & "<a href='/outlook/stop_vbs' class='lcars-button danger'>Stop VBS</a>"
    html = html & "<a href='/outlook/enable_background' class='lcars-button success'>Enable BG</a>"
    html = html & "<a href='/outlook/disable_background' class='lcars-button danger'>Disable BG</a>"
    html = html & "<a href='/outlook/stop_checking' class='lcars-button danger'>Stop All</a>"
    html = html & "</div>"
    
    html = html & "<p style='text-align:center;margin-top:30px;'>"
    html = html & "<a href='/outlook' class='lcars-button'>Standard View</a>"
    html = html & "<a href='/outlook/enhanced' class='lcars-button'>Enhanced View</a>"
    html = html & "<a href='/index.html' class='lcars-button'>Home</a>"
    html = html & "</p>"
    
    html = html & "</div></body></html>"
    
    GenerateLCARSOutlookPage = html
    Exit Function
ErrorHandler:
    GenerateLCARSOutlookPage = GenerateErrorResponse("Error generating page: " & Err.description, "/outlook")
End Function

'=========================
' Generate action response HTML
'=========================
Private Function GenerateActionResponse(ByVal message As String, ByVal backPath As String) As String
    GenerateActionResponse = "<html><body><h2>" & HTMLEncode(message) & "</h2>" & _
                             "<p><a href='" & backPath & "'>Back</a></p></body></html>"
End Function

'=========================
' Generate error response HTML
'=========================
Private Function GenerateErrorResponse(ByVal message As String, ByVal backPath As String) As String
    GenerateErrorResponse = "<html><body><h2>Error: " & HTMLEncode(message) & "</h2>" & _
                            "<p><a href='" & backPath & "'>Back</a></p></body></html>"
End Function

'=========================
' Outlook status JSON
'=========================
Public Function GetOutlookStatusJSON() As String
    On Error GoTo ErrorHandler
    
    Dim json As String
    json = "{"
    json = json & """timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & ""","
    json = json & """outlookRunning"":" & IIf(Outlook.IsOutlookRunning(), "true", "false") & ","
    json = json & """unreadCount"":" & Outlook.GetUnreadMsgCount() & ","
    json = json & """backgroundMode"":" & IIf(Outlook.UseBackgroundOutlookMode, "true", "false") & ","
    json = json & """vbsRunning"":" & IIf(OutlookExtra.IsVBSCheckerRunning(), "true", "false") & ","
    json = json & """totalRequests"":" & m_totalRequests & ","
    json = json & """lastActivity"":""" & format(m_lastActivity, "yyyy-mm-dd hh:mm:ss") & """}"
    
    GetOutlookStatusJSON = json
    Exit Function
ErrorHandler:
    GetOutlookStatusJSON = "{""error"":""" & Err.description & """}"
End Function

'=========================
' Send Outlook mail
'=========================
Private Function SendOutlookMail(ByVal jsonBody As String) As String
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olMail As Object
    Dim toaddr As String, subj As String, msg As String
    
    toaddr = ExtractJSONValue(jsonBody, "to")
    subj = ExtractJSONValue(jsonBody, "subject")
    msg = ExtractJSONValue(jsonBody, "body")
    
    If toaddr = "" Or InStr(toaddr, "@") = 0 Then
        SendOutlookMail = "{""error"":""Invalid 'to' address""}"
        Exit Function
    End If
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    With olMail
        .To = toaddr
        .Subject = subj
        .body = msg
        .send
    End With
    
    SendOutlookMail = "{""status"":""sent"",""to"":""" & toaddr & """,""subject"":""" & subj & """}"
    
    Set olMail = Nothing
    Set olApp = Nothing
    Exit Function
ErrorHandler:
    SendOutlookMail = "{""error"":""" & Err.description & """}"
End Function

'=========================
' Universal Email Check
'=========================
Private Sub HandleUniversalEmailCheck(ByVal body As String)
    On Error GoTo ErrorHandler
    
    Dim attachmentMode As String, useBG As Boolean, useColl As Boolean, forwardBody As Boolean, hoursBack As Long
    Dim hStr As String
    
    attachmentMode = ExtractJSONValue(body, "attachmentMode")
    If attachmentMode = "" Then attachmentMode = "Memory"
    
    useBG = (ExtractJSONValue(body, "background") = "true")
    useColl = (ExtractJSONValue(body, "useCollections") <> "false")
    forwardBody = (ExtractJSONValue(body, "forwardBody") <> "false")
    
    hStr = ExtractJSONValue(body, "hoursBack")
    If IsNumeric(hStr) Then hoursBack = CLng(hStr) Else hoursBack = 6
    
    OutlookExtra.UniversalEmailChecker attachmentMode, useBG, useColl, forwardBody, hoursBack
    Exit Sub
ErrorHandler:
    DebugLog "Error in HandleUniversalEmailCheck: " & Err.description
End Sub

'=========================
' Simple JSON value extractor
'=========================
Private Function ExtractJSONValue(ByVal json As String, ByVal key As String) As String
    On Error Resume Next
    
    Dim pattern As String, matches As Object
    pattern = """" & key & """:\s*""([^""]*)""|""" & key & """:\s*([^,}]*)"
    
    With CreateObject("VBScript.RegExp")
        .Global = False
        .IgnoreCase = True
        .pattern = pattern
        If .test(json) Then
            Set matches = .Execute(json)
            If matches(0).SubMatches(0) <> "" Then
                ExtractJSONValue = matches(0).SubMatches(0)
            Else
                ExtractJSONValue = Trim(matches(0).SubMatches(1))
            End If
        End If
    End With
End Function



