Option Explicit

'==============================================================
' Module: OutlookWebUI
' Purpose: Serve Outlook Web UI pages and handle all Outlook actions
'==============================================================

' -----------------------------
' Serve /outlook/status.json
' -----------------------------
Public Sub SendOutlookStatusJson(client As Object)
    On Error GoTo ErrorHandler
    
    Dim statusJSON As String
    statusJSON = "{""unreadCount"":""" & GetUnreadCountSafe() & """," & _
                 """status"":""" & GetOutlookStatusSafe() & """," & _
                 """background"":""" & GetBackgroundStatusSafe() & """," & _
                 """vbs"":""" & GetVBSStatusSafe() & """," & _
                 """recentEmails"":""" & Replace(GetOutlookDataSafe(), """", "\""") & """}"
    
    SendHttpResponse client, 200, "application/json", statusJSON
    Exit Sub

ErrorHandler:
    Debug.Print "Error sending Outlook status JSON: " & Err.description
    SendHttpResponse client, 500, "application/json", "{""error"":""" & HTMLEncode(Err.description) & """}"
End Sub

' -----------------------------
' HTTP Response Helpers
' -----------------------------
Public Sub SendHttpResponse(client As Object, statusCode As Long, contentType As String, content As String)
    Dim headers As String
    headers = "HTTP/1.1 " & statusCode & " OK" & vbCrLf & _
              "Content-Type: " & contentType & vbCrLf & _
              "Content-Length: " & LenB(content) & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    client.send headers & content
End Sub

Public Sub SendHttpResponseBytes(client As Object, statusCode As Long, contentType As String, content() As Byte)
    Dim headers As String
    headers = "HTTP/1.1 " & statusCode & " OK" & vbCrLf & _
              "Content-Type: " & contentType & vbCrLf & _
              "Content-Length: " & UBound(content) + 1 & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    client.send headers
    client.send content
End Sub

' -----------------------------
' Handle Web Requests
' -----------------------------
Public Sub HandleWebUIRequest(client As Object, requestPath As String)
    On Error GoTo ErrorHandler
    
    Select Case requestPath
        Case "/", "/index.html"
            SendHttpResponse client, 200, "text/html", GenerateLCARSOutlookLandingPage()
        Case "/outlook/status.json"
            SendOutlookStatusJson client
        Case "/outlook/force_check"
            ForceOutlookScanSafe
            SendHttpResponse client, 200, "application/json", "{""result"":""Force Scan triggered""}"
        Case "/outlook/run_rules"
            RunOutlookRulesSafe
            SendHttpResponse client, 200, "application/json", "{""result"":""Rules executed""}"
        Case "/outlook/enable_background"
            SetBackgroundModeSafe True
            SendHttpResponse client, 200, "application/json", "{""result"":""Background enabled""}"
        Case "/outlook/disable_background"
            SetBackgroundModeSafe False
            SendHttpResponse client, 200, "application/json", "{""result"":""Background disabled""}"
        Case "/style.css"
            SendHttpResponse client, 200, "text/css", GenerateCss()
        Case "/script.js"
            SendHttpResponse client, 200, "application/javascript", GenerateJs()
        Case "/favicon.ico"
            SendHttpResponseBytes client, 200, "image/png", GetFaviconBytes()
        Case Else
            SendHttpResponse client, 404, "text/plain", "404 Not Found"
    End Select
    Exit Sub

ErrorHandler:
    Debug.Print "Error in HandleWebUIRequest: " & Err.description
    SendHttpResponse client, 500, "text/plain", "Server Error: " & HTMLEncode(Err.description)
End Sub

' -----------------------------
' Handle Outlook Actions
' -----------------------------
Public Sub HandleOutlookAction(ByVal actionName As String)
    On Error GoTo ErrHandler
    Select Case LCase(actionName)
        Case "force_check": Call OutlookExtra.EnhancedOutlookHealthCheck
        Case "run_rules": Call OutlookExtra.RunOutlookRules
        Case "enable_background": Call OutlookExtra.EnableBackgroundMode
        Case "disable_background": Call OutlookExtra.DisableBackgroundMode
        Case "initialize": Call Outlook.InitializeOutlookChecking
        Case "stop": Call Outlook.StopOutlookChecking
        Case "diagnose": Call Outlook.DiagnoseOutlookConnection
        Case "test_connection": Call Outlook.TestBasicOutlookConnection
        Case "test_rules": Call Outlook.TestDirectOutlookRules
        Case Else: Debug.Print "Unknown Outlook action: " & actionName
    End Select
    Exit Sub
ErrHandler:
    Debug.Print "Error in HandleOutlookAction '" & actionName & "': " & Err.description
End Sub

' -----------------------------
' LCARS Landing Page
' -----------------------------
Public Function GenerateLCARSOutlookLandingPage() As String
    On Error GoTo ErrorHandler

    Dim html As String, style As String, body As String, stardate As String
    stardate = "2025." & format(Now, "ddd.dd.hh")
    
    style = "<style>" & vbCrLf
    style = style & "body{background:black;color:#FF9966;font-family:'OCR-A',Arial,sans-serif;margin:0;padding:20px;}" & vbCrLf
    style = style & ".container{max-width:1200px;margin:auto;}" & vbCrLf
    style = style & ".bar{height:40px;background:linear-gradient(to right,#663399,#CC6600);margin:10px 0;animation:flash 1.5s infinite alternate;}" & vbCrLf
    style = style & "@keyframes flash{from{opacity:0.6;}to{opacity:1;}}" & vbCrLf
    style = style & ".btn{padding:10px;background:#CC6600;color:black;font-weight:bold;border-radius:8px;cursor:pointer;border:2px solid #FFFF99;display:inline-block;margin:5px;text-decoration:none;}" & vbCrLf
    style = style & ".btn:hover{background:#FF9966;border-color:#99CCFF;}" & vbCrLf
    style = style & ".btn:disabled{opacity:0.5;cursor:default;}" & vbCrLf
    style = style & ".header{font-size:36px;color:#99CCFF;text-shadow:0 0 10px #99CCFF;margin-bottom:20px;}" & vbCrLf
    style = style & ".subheader{font-size:18px;color:#FFFF99;margin:10px 0;}" & vbCrLf
    style = style & ".section{margin:15px 0;padding:10px;border:2px solid #663399;border-radius:10px;background:#1C2526;}" & vbCrLf
    style = style & ".section-title{font-size:24px;color:#99CCFF;text-transform:uppercase;}" & vbCrLf
    style = style & ".status-active{color:#33FF33;}.status-inactive{color:#FF3333;}.status-warning{color:#FFFF00;}" & vbCrLf
    style = style & "</style>"

    body = "<div class='container'>" & vbCrLf
    body = body & "<div class='bar'></div><h1 class='header'>LCARS - Outlook Scanner</h1>" & vbCrLf
    body = body & "<div class='subheader'>Starfleet Command &bull; Email Monitoring &bull; Stardate " & stardate & "</div>" & vbCrLf

    ' Status Section
    body = body & "<div class='section'><div class='section-title'>Outlook Status</div>" & vbCrLf
    body = body & "<p>Unread Messages: <strong id='unreadCount'>--</strong></p>" & vbCrLf
    body = body & "<p>Status: <strong id='outlookStatus'>--</strong></p>" & vbCrLf
    body = body & "<p>Background Mode: <strong id='backgroundStatus'>--</strong></p>" & vbCrLf
    body = body & "<p>VBS Checker: <strong id='vbsStatus'>--</strong></p></div>" & vbCrLf

    ' Recent Emails
    body = body & "<div class='section'><div class='section-title'>Recent Emails</div>" & vbCrLf
    body = body & "<p id='recentEmails'>No recent emails available</p></div>" & vbCrLf

    ' Actions
    body = body & "<div class='section'><div class='section-title'>Actions</div>" & vbCrLf
    body = body & "<a href='/outlook/force_check' class='btn'>Force Scan</a>" & vbCrLf
    body = body & "<a href='/outlook/run_rules' class='btn'>Run Rules</a>" & vbCrLf
    body = body & "<a href='/outlook/enable_background' class='btn'>Enable Background</a>" & vbCrLf
    body = body & "<a href='/outlook/disable_background' class='btn'>Disable Background</a></div>" & vbCrLf

    body = body & "<div class='bar'></div><a href='/index.html' class='btn'>Return to Home</a></div>" & vbCrLf

    ' JS
    body = body & "<script>" & vbCrLf
    body = body & "function refreshStatus(){" & vbCrLf
    body = body & "fetch('/outlook/status.json').then(r=>r.json()).then(data=>{" & vbCrLf
    body = body & "document.getElementById('unreadCount').innerText=data.unreadCount;" & vbCrLf
    body = body & "document.getElementById('outlookStatus').innerText=data.status;" & vbCrLf
    body = body & "document.getElementById('backgroundStatus').innerText=data.background;" & vbCrLf
    body = body & "document.getElementById('vbsStatus').innerText=data.vbs;" & vbCrLf
    body = body & "document.getElementById('recentEmails').innerHTML=data.recentEmails.replace(/\\n/g,'<br>');}).catch(e=>console.error(e));}" & vbCrLf
    body = body & "setInterval(refreshStatus,5000);" & vbCrLf
    body = body & "document.addEventListener('DOMContentLoaded',refreshStatus);" & vbCrLf
    body = body & "document.querySelectorAll('.btn').forEach(function(btn){btn.addEventListener('click',function(e){e.preventDefault();" & vbCrLf
    body = body & "if(btn.disabled)return;btn.disabled=true;var url=btn.getAttribute('href');" & vbCrLf
    body = body & "fetch(url).then(r=>r.json()).then(data=>{console.log(data.result);refreshStatus();}).catch(err=>console.error(err)).finally(()=>{btn.disabled=false;});});});" & vbCrLf
    body = body & "</script>"

    html = "<html><head><title>LCARS Outlook Scanner</title>" & style & "</head><body>" & body & "</body></html>"

    GenerateLCARSOutlookLandingPage = html
    Exit Function
ErrorHandler:
    GenerateLCARSOutlookLandingPage = "<html><body><h2>Error generating page</h2><p>" & HTMLEncode(Err.description) & "</p></body></html>"
End Function

' -----------------------------
' CSS/JS for endpoints
' -----------------------------
Public Function GenerateCss() As String
    GenerateCss = "body{background:black;color:#FF9966;font-family:'OCR-A',Arial,sans-serif;margin:0;padding:20px;}" & _
                  ".btn{padding:10px;background:#CC6600;color:black;font-weight:bold;border-radius:8px;cursor:pointer;border:2px solid #FFFF99;}" & _
                  ".btn:hover{background:#FF9966;border-color:#99CCFF;}"
End Function

Public Function GenerateJs() As String
    GenerateJs = "function refreshStatus(){fetch('/outlook/status.json').then(r=>r.json()).then(data=>{console.log(data);}).catch(e=>console.error(e));}"
End Function

' -----------------------------
' Safe getters
' -----------------------------
Private Function GetUnreadCountSafe() As String
    On Error Resume Next
    GetUnreadCountSafe = CStr(Outlook.GetUnreadMsgCount())
    If Err.Number <> 0 Then GetUnreadCountSafe = "--"
    On Error GoTo 0
End Function

Private Function GetOutlookStatusSafe() As String
    On Error Resume Next
    If Outlook.IsOutlookRunning() Then GetOutlookStatusSafe = "Running" Else GetOutlookStatusSafe = "Stopped"
    If Err.Number <> 0 Then GetOutlookStatusSafe = "Unknown"
    On Error GoTo 0
End Function

Private Function GetBackgroundStatusSafe() As String
    On Error Resume Next
    If OutlookExtra.IsBackgroundModeEnabled() Then GetBackgroundStatusSafe = "Enabled" Else GetBackgroundStatusSafe = "Disabled"
    If Err.Number <> 0 Then GetBackgroundStatusSafe = "Unknown"
    On Error GoTo 0
End Function

Private Function GetVBSStatusSafe() As String
    On Error Resume Next
    If Outlook.IsVBSCheckerRunning() Then GetVBSStatusSafe = "Running" Else GetVBSStatusSafe = "Stopped"
    If Err.Number <> 0 Then GetVBSStatusSafe = "Unknown"
    On Error GoTo 0
End Function

Private Function GetOutlookDataSafe() As String
    On Error Resume Next
    GetOutlookDataSafe = OutlookExtra.SearchOutlookEmails(5)
    If Err.Number <> 0 Or GetOutlookDataSafe = "" Then GetOutlookDataSafe = "No recent emails available"
    On Error GoTo 0
End Function

' -----------------------------
' Favicon
' -----------------------------
Public Function GetFaviconBytes() As Byte()
    Dim b() As Byte
    b = Array(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 16, 0, 0, 0, 16, 8, 2, 0, 0, 0, 241, 216, 66, 174, 0, 0, 0, 12, 73, 68, 65, 84, 8, 153, 99, 248, 15, 4, 12, 0, 9, 251, 3, 253, 171, 179, 185, 103, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
    GetFaviconBytes = b
End Function


