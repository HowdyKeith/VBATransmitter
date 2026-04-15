Option Explicit


'==================================================
' Module: OutlookWebRules
' Purpose: Serve dynamic HTML page for Outlook rules
'          and execute specific rules safely via VBA
' Dependencies: OutlookExtra module for RunOutlookRule
'==================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const SHEET_NAME As String = "Outlook"

Private Const STATUS_TIMEOUT_MS As Long = 5000  ' Timeout for rule execution feedback

' OutlookLauncher Module (fixed)
' Purpose: Safely create and execute VBS script to launch Outlook
'          with proper error handling and path validation
'***************************************************************


' Improved VBS export with better error handling and path management
Public Sub ExportLaunchOutlookVBS(ByVal vbsPath As String)
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim f As Object
    Dim tempDir As String
    Dim actualPath As String
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' If no path specified or path doesn't exist, use temp directory
    If vbsPath = "" Or Not fso.FolderExists(fso.GetParentFolderName(vbsPath)) Then
        tempDir = GetTempDirectory()
        actualPath = tempDir & "launchoutlook.vbs"
        DebuggingLog.DebugLog "[ExportLaunchOutlookVBS] Using temp path: " & actualPath
    Else
        actualPath = vbsPath
    End If
    
    ' Delete existing file if it exists
    If fso.FileExists(actualPath) Then
        On Error Resume Next
        fso.DeleteFile actualPath, True
        On Error GoTo ErrorHandler
    End If
    
    ' Create the VBS file with enhanced error handling
    Set f = fso.CreateTextFile(actualPath, True, False) ' False = ASCII encoding
    
    ' Write improved VBS script
    f.WriteLine "On Error Resume Next"
    f.WriteLine "Set objShell = CreateObject(""WScript.Shell"")"
    f.WriteLine "Set objFSO = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine ""
    f.WriteLine "' Check if Outlook is already running"
    f.WriteLine "Set objWMIService = GetObject(""winmgmts:\\.\root\cimv2"")"
    f.WriteLine "Set colProcesses = objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE Name = 'OUTLOOK.EXE'"")"
    f.WriteLine "If colProcesses.Count > 0 Then"
    f.WriteLine "    WScript.Echo ""Outlook is already running"""
    f.WriteLine "    WScript.Quit"
    f.WriteLine "End If"
    f.WriteLine ""
    f.WriteLine "' Try to launch Outlook"
    f.WriteLine "Set olApp = CreateObject(""Outlook.Application"")"
    f.WriteLine "If Err.Number <> 0 Then"
    f.WriteLine "    WScript.Echo ""Error creating Outlook application: "" & Err.Description"
    f.WriteLine "    WScript.Quit"
    f.WriteLine "End If"
    f.WriteLine ""
    f.WriteLine "' Try to get session"
    f.WriteLine "Set olSession = olApp.Session"
    f.WriteLine "If Err.Number <> 0 Then"
    f.WriteLine "    WScript.Echo ""Error accessing Outlook session: "" & Err.Description"
    f.WriteLine "    WScript.Quit"
    f.WriteLine "End If"
    f.WriteLine ""
    f.WriteLine "WScript.Echo ""Outlook launched successfully"""
    
    f.Close
    Set f = Nothing
    Set fso = Nothing
    
    DebuggingLog.DebugLog "[ExportLaunchOutlookVBS] VBS file created: " & actualPath
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "[ExportLaunchOutlookVBS] Error: " & Err.description & " (Error " & Err.Number & ")"
    On Error Resume Next
    If Not f Is Nothing Then f.Close
    Set f = Nothing
    Set fso = Nothing
End Sub

' Safe Outlook launcher with multiple fallback methods
Public Sub LaunchOutlookSafely()
    On Error GoTo ErrorHandler
    
    Dim vbsPath As String
    Dim tempDir As String
    Dim result As Long
    
    ' Try Method 1: Direct VBA Outlook launch (fastest)
    If TryLaunchOutlookDirect() Then
        DebuggingLog.DebugLog "[LaunchOutlookSafely] Outlook launched directly via VBA"
        Exit Sub
    End If
    
    ' Try Method 2: VBS script launch
    tempDir = GetTempDirectory()
    vbsPath = tempDir & "launchoutlook.vbs"
    
    Call ExportLaunchOutlookVBS(vbsPath)
    
    ' Execute the VBS file with better error handling
    If LaunchVBSScript(vbsPath) Then
        DebuggingLog.DebugLog "[LaunchOutlookSafely] Outlook launched via VBS script"
    Else
        ' Try Method 3: Direct Shell command
        If TryLaunchOutlookShell() Then
            DebuggingLog.DebugLog "[LaunchOutlookSafely] Outlook launched via Shell command"
        Else
            DebuggingLog.DebugLog "[LaunchOutlookSafely] All launch methods failed"
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "[LaunchOutlookSafely] Error: " & Err.description
End Sub

' Method 1: Try direct VBA launch
Private Function TryLaunchOutlookDirect() As Boolean
    On Error Resume Next
    Dim olApp As Object
    
    ' Check if Outlook is already running
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number = 0 Then
        TryLaunchOutlookDirect = True
        Set olApp = Nothing
        Exit Function
    End If
    Err.Clear
    
    ' Try to create new Outlook instance
    Set olApp = CreateObject("Outlook.Application")
    If Err.Number = 0 Then
        ' Try to access session to ensure it's working
        Dim olSession As Object
        Set olSession = olApp.Session
        If Err.Number = 0 Then
            TryLaunchOutlookDirect = True
        End If
    End If
    
    Set olApp = Nothing
    Set olSession = Nothing
End Function

' Method 2: Launch VBS script safely
Private Function LaunchVBSScript(ByVal vbsPath As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Dim result As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verify file exists
    If Not fso.FileExists(vbsPath) Then
        LaunchVBSScript = False
        Exit Function
    End If
    
    ' Try different execution methods
    ' Method 2A: WScript
    result = shell("wscript.exe """ & vbsPath & """", vbHide)
    If Err.Number = 0 And result > 0 Then
        Sleep 2000 ' Wait for script to execute
        LaunchVBSScript = True
        Exit Function
    End If
    Err.Clear
    
    ' Method 2B: CScript (fallback)
    result = shell("cscript.exe """ & vbsPath & """ //NoLogo", vbHide)
    If Err.Number = 0 And result > 0 Then
        Sleep 2000
        LaunchVBSScript = True
        Exit Function
    End If
    
    LaunchVBSScript = False
    Set fso = Nothing
End Function

' Method 3: Direct Shell launch
Private Function TryLaunchOutlookShell() As Boolean
    On Error Resume Next
    Dim result As Long
    
    ' Try launching Outlook directly
    result = shell("outlook.exe", vbNormalFocus)
    If Err.Number = 0 And result > 0 Then
        TryLaunchOutlookShell = True
    Else
        TryLaunchOutlookShell = False
    End If
End Function

' Get system temp directory
Private Function GetTempDirectory() As String
    On Error Resume Next
    Dim tempPath As String
    Dim Length As Long
    
    tempPath = String(260, 0)
    Length = GetTempPath(260, tempPath)
    
    If Length > 0 Then
        GetTempDirectory = Left(tempPath, Length)
    Else
        ' Fallback to common temp locations
        GetTempDirectory = Environ("TEMP")
        If GetTempDirectory = "" Then GetTempDirectory = Environ("TMP")
        If GetTempDirectory = "" Then GetTempDirectory = "C:\Temp\"
    End If
    
    ' Ensure path ends with backslash
    If Right(GetTempDirectory, 1) <> "\" Then
        GetTempDirectory = GetTempDirectory & "\"
    End If
End Function


' Clean up temp VBS files
Public Sub CleanupTempFiles()
    On Error Resume Next
    Dim fso As Object
    Dim tempDir As String
    Dim vbsPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempDir = GetTempDirectory()
    vbsPath = tempDir & "launchoutlook.vbs"
    
    If fso.FileExists(vbsPath) Then
        fso.DeleteFile vbsPath, True
        DebuggingLog.DebugLog "[CleanupTempFiles] Deleted temp VBS file"
    End If
    
    Set fso = Nothing
End Sub






Public Function GetRoutes() As Collection
    Dim routes As New Collection
    routes.Add Array("/outlook", "GenerateLCARSOutlookLandingPage"), "outlook"
    routes.Add Array("/outlook/force_check", "HandleOutlookAction:force_check"), "force_check"
    routes.Add Array("/outlook/run_rules", "HandleOutlookAction:run_rules"), "run_rules"
    routes.Add Array("/outlook/enable_background", "HandleOutlookAction:enable_background"), "enable_background"
    routes.Add Array("/outlook/disable_background", "HandleOutlookAction:disable_background"), "disable_background"
    Set GetRoutes = routes
End Function

'==============================
' HTML PAGE GENERATION
'==============================
Public Function GenerateOutlookRulesPage() As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME) ' Sheet with rule list
    
    Dim html As String
    Dim lastRow As Long, i As Long
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' --- CSS Styles (LCARS-inspired, matching OutlookWebUI and Govee) ---
    html = "<!DOCTYPE html><html><head><title>Outlook Rules</title>"
    html = html & "<style>"
    html = html & "body { background: black; color: #FF9966; font-family: 'OCR-A', Arial, sans-serif; padding: 20px; }"
    html = html & ".container { max-width: 1200px; margin: auto; }"
    html = html & ".bar { height: 40px; background: linear-gradient(to right, #663399, #CC6600); margin: 10px 0; animation: flash 1.5s infinite alternate; }"
    html = html & "@keyframes flash { from { opacity: 0.6; } to { opacity: 1; } }"
    html = html & ".btn { padding: 10px; background: #CC6600; color: black; font-weight: bold; border-radius: 8px; cursor: pointer; border: 2px solid #FFFF99; display: inline-block; margin: 5px; text-decoration: none; }"
    html = html & ".btn:hover { background: #FF9966; border-color: #99CCFF; }"
    html = html & ".header { font-size: 36px; color: #99CCFF; text-shadow: 0 0 10px #99CCFF; margin-bottom: 20px; }"
    html = html & ".subheader { font-size: 18px; color: #FFFF99; margin: 10px 0; }"
    html = html & ".section { margin: 15px 0; padding: 10px; border: 2px solid #663399; border-radius: 10px; background: #1C2526; }"
    html = html & ".section-title { font-size: 24px; color: #99CCFF; text-transform: uppercase; }"
    html = html & "#status { margin-top: 15px; font-weight: bold; color: #FFA500; }"
    html = html & "</style></head><body>"
    
    ' --- HTML Body ---
    html = html & "<div class='container'>"
    html = html & "<div class='bar'></div>"
    html = html & "<h1 class='header'>LCARS - Outlook Rules</h1>"
    html = html & "<div class='subheader'>Starfleet Command &bull; Rule Execution &bull; Stardate " & format(Now, "yyyy.ddd.hh") & "</div>"
    
    ' --- Rules Section ---
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Available Rules</div>"
    For i = 2 To lastRow
        If ws.Cells(i, 2).value = True Then
            Dim ruleName As String
            ruleName = Replace(ws.Cells(i, 1).value, " ", "%20") ' URL-encode spaces
            html = html & "<a href='/outlook/run_rules?rule=" & ruleName & "' class='btn'>" & ws.Cells(i, 1).value & "</a>"
        End If
    Next i
    html = html & "</div>"
    
    html = html & "<div id='status' class='section'>Status: Ready</div>"
    html = html & "<div class='bar'></div>"
    html = html & "<a href='/outlook' class='btn'>Back to Outlook Dashboard</a>"
    html = html & "<a href='/index.html' class='btn'>Return to Home</a>"
    html = html & "</div></body></html>"
    
    GenerateOutlookRulesPage = html
    Exit Function
    
ErrorHandler:
    DebugLog "Error generating Outlook rules page: " & Err.description
    GenerateOutlookRulesPage = "<html><body><h1>Error</h1><p>Error generating page: " & Err.description & "</p></body></html>"
End Function

'==============================
' HTTP HANDLER FOR RULE EXECUTION
'==============================
Public Function HandleRunRulesRequest(ByVal requestURL As String) As String
    On Error GoTo ErrorHandler
    
    ' Parse rule name from query string
    Dim ruleName As String
    ruleName = ParseQueryParameter(requestURL, "rule")
    
    If ruleName = "" Then
        HandleRunRulesRequest = "<html><body><h1>Error</h1><p>No rule specified</p><a href='/outlook/rules'>Back to Rules</a></body></html>"
        DebugLog "No rule specified in request: " & requestURL
        Exit Function
    End If
    
    ' Execute the rule using OutlookExtra module
    Dim startTime As Long
    startTime = GetTickCount()
    
    OutlookExtra.RunOutlookRule ruleName
    
    ' Wait briefly for rule execution to complete
    Do While GetTickCount() - startTime < STATUS_TIMEOUT_MS
        DoEvents
        Sleep 10
    Loop
    
    ' Redirect back to rules page with status
    HandleRunRulesRequest = "<html><head><meta http-equiv='refresh' content='0;url=/outlook/rules'></head><body><p>Rule " & ruleName & " executed. Redirecting...</p></body></html>"
    DebugLog "Rule executed via HTTP: " & ruleName
    Exit Function
    
ErrorHandler:
    DebugLog "Error in HandleRunRulesRequest: " & Err.description
    HandleRunRulesRequest = "<html><body><h1>Error</h1><p>Error executing rule: " & Err.description & "</p><a href='/outlook/rules'>Back to Rules</a></body></html>"
End Function

'==============================
' QUERY STRING PARSER
'==============================
Public Function ParseQueryParameter(ByVal url As String, ByVal paramName As String) As String
    On Error GoTo ErrorHandler
    
    Dim parts() As String
    parts = Split(url, "?")
    If UBound(parts) < 1 Then Exit Function
    
    Dim query As String
    query = parts(1)
    
    Dim pairs() As String, pair() As String, i As Long
    pairs = Split(query, "&")
    
    For i = LBound(pairs) To UBound(pairs)
        pair = Split(pairs(i), "=")
        If UBound(pair) >= 1 And LCase(pair(0)) = LCase(paramName) Then
            ParseQueryParameter = DecodeURL(pair(1))
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    DebugLog "Error parsing query parameter: " & Err.description
End Function

'==============================
' URL Decode
'==============================
Public Function DecodeURL(ByVal str As String) As String
    On Error GoTo ErrorHandler
    
    str = Replace(str, "+", " ")
    Dim i As Long, hexVal As String
    i = 1
    Do While i <= Len(str)
        If Mid(str, i, 1) = "%" And i <= Len(str) - 2 Then
            hexVal = Mid(str, i + 1, 2)
            On Error Resume Next
            Mid(str, i, 3) = Chr(CLng("&H" & hexVal))
            On Error GoTo ErrorHandler
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    DecodeURL = str
    Exit Function
    
ErrorHandler:
    DebugLog "Error in DecodeURL: " & Err.description
    DecodeURL = str
End Function
Option Explicit

Public Sub ExportLaunchOutlookVBSOLD(ByVal vbsPath As String)
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(vbsPath, True)
    f.WriteLine "Set olApp = CreateObject(""Outlook.Application"")"
    f.WriteLine "olApp.Session.Logon"
    f.Close
End Sub


