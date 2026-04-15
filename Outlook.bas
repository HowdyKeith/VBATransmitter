Option Explicit

'***************************************************************
' Outlook Module - Core Outlook Functions with Background/Force Scan
'***************************************************************

Private m_backgroundEnabled As Boolean
Private m_lastScanTime As Date

Public Function GetRoutes() As Collection
    Dim routes As New Collection
    
    ' Main Outlook status/control page
    routes.Add Array("/outlook", "GenerateOutlookStatusPage"), "/outlook"
    
    ' Core Outlook actions
    routes.Add Array("/outlook/force_check", "HandleOutlookAction:force_check"), "/outlook/force_check"
    routes.Add Array("/outlook/enable_background", "HandleOutlookAction:enable_background"), "/outlook/enable_background"
    routes.Add Array("/outlook/disable_background", "HandleOutlookAction:disable_background"), "/outlook/disable_background"
    
    ' Initialization and control
    routes.Add Array("/outlook/initialize", "HandleOutlookAction:initialize"), "/outlook/initialize"
    routes.Add Array("/outlook/stop", "HandleOutlookAction:stop"), "/outlook/stop"
    
    ' Testing and diagnostics
    routes.Add Array("/outlook/diagnose", "HandleOutlookAction:diagnose"), "/outlook/diagnose"
    routes.Add Array("/outlook/test_connection", "HandleOutlookAction:test_connection"), "/outlook/test_connection"
    routes.Add Array("/outlook/test_rules", "HandleOutlookAction:test_rules"), "/outlook/test_rules"
    routes.Add Array("/outlook/run_rules", "HandleOutlookAction:run_rules"), "/outlook/run_rules"
    
    ' Testing modes
    routes.Add Array("/outlook/test_memory", "HandleOutlookAction:test_memory"), "/outlook/test_memory"
    routes.Add Array("/outlook/test_temp", "HandleOutlookAction:test_temp"), "/outlook/test_temp"
    routes.Add Array("/outlook/test_background", "HandleOutlookAction:test_background"), "/outlook/test_background"
    
    ' Status and info endpoints
    routes.Add Array("/outlook/status", "GenerateOutlookStatusJSON"), "/outlook/status"
    routes.Add Array("/outlook/unread_count", "GetOutlookUnreadCountJSON"), "/outlook/unread_count"
    routes.Add Array("/outlook/is_running", "GetOutlookRunningStatusJSON"), "/outlook/is_running"
    
    ' VBS checker management
    routes.Add Array("/outlook/start_checker", "HandleOutlookAction:start_checker"), "/outlook/start_checker"
    routes.Add Array("/outlook/stop_checker", "HandleOutlookAction:stop_checker"), "/outlook/stop_checker"
    routes.Add Array("/outlook/checker_status", "GetVBSCheckerStatusJSON"), "/outlook/checker_status"
    
    Set GetRoutes = routes
End Function

' Handle Outlook action requests (called by AppLaunch routing)
Public Sub HandleOutlookAction(ByVal action As String)
    On Error GoTo ErrorHandler
    
    Select Case LCase(action)
        Case "force_check"
            Call ForceScanOutlook
            DebugLog "Force check executed"
            
        Case "enable_background"
            Call EnableBackgroundMode
            Call StartOutlookVBSChecker
            DebugLog "Background mode enabled"
            
        Case "disable_background"
            Call DisableBackgroundMode
            Call StopOutlookVBSChecker
            DebugLog "Background mode disabled"
            
        Case "initialize"
            Call InitializeOutlookCheckingSimple
            DebugLog "Outlook checking initialized"
            
        Case "stop"
            Call StopOutlookCheckingSafe
            DebugLog "Outlook checking stopped"
            
        Case "diagnose"
            Call DiagnoseOutlookConnection
            DebugLog "Outlook diagnostics run"
            
        Case "test_connection"
            If IsOutlookRunning() Then
                DebugLog "Outlook connection test: PASSED"
            Else
                DebugLog "Outlook connection test: FAILED"
            End If
            
        Case "test_rules"
            Call TestOutlookRules
            DebugLog "Outlook rules test executed"
            
        Case "run_rules"
            Call RunOutlookRules
            DebugLog "Outlook rules executed"
            
        Case "test_memory"
            Call TestImmediateMemoryMode
            DebugLog "Memory mode test executed"
            
        Case "test_temp"
            Call TestImmediateTempMode
            DebugLog "Temp mode test executed"
            
        Case "test_background"
            Call TestBackgroundMode
            DebugLog "Background mode test executed"
            
        Case "start_checker"
            Call StartOutlookVBSChecker
            DebugLog "VBS checker started"
            
        Case "stop_checker"
            Call StopOutlookVBSChecker
            DebugLog "VBS checker stopped"
            
        Case Else
            DebugLog "Unknown Outlook action: " & action
    End Select
    
    Exit Sub
ErrorHandler:
    DebugLog "Error in HandleOutlookAction(" & action & "): " & Err.description
End Sub



' JSON status endpoints for API access
Public Function GenerateOutlookStatusJSON() As String
    On Error GoTo ErrorHandler
    
    Dim json As String
    json = "{""outlook_running"":" & IIf(IsOutlookRunning(), "true", "false")
    json = json & ",""unread_count"":" & GetUnreadMsgCount()
    json = json & ",""checker_running"":" & IIf(IsVBSCheckerRunning(), "true", "false")
    json = json & ",""background_enabled"":" & IIf(m_backgroundEnabled, "true", "false")
    json = json & ",""last_scan"":""" & IIf(m_lastScanTime > 0, format(m_lastScanTime, "yyyy-mm-dd hh:mm:ss"), "never") & """"
    json = json & "}"
    
    GenerateOutlookStatusJSON = json
    Exit Function
    
ErrorHandler:
    GenerateOutlookStatusJSON = "{""error"":""" & Err.description & """}"
End Function

Public Function GetOutlookUnreadCountJSON() As String
    On Error GoTo ErrorHandler
    GetOutlookUnreadCountJSON = "{""unread_count"":" & GetUnreadMsgCount() & "}"
    Exit Function
ErrorHandler:
    GetOutlookUnreadCountJSON = "{""error"":""" & Err.description & """}"
End Function

Public Function GetOutlookRunningStatusJSON() As String
    On Error GoTo ErrorHandler
    GetOutlookRunningStatusJSON = "{""is_running"":" & IIf(IsOutlookRunning(), "true", "false") & "}"
    Exit Function
ErrorHandler:
    GetOutlookRunningStatusJSON = "{""error"":""" & Err.description & """}"
End Function

Public Function GetVBSCheckerStatusJSON() As String
    On Error GoTo ErrorHandler
    GetVBSCheckerStatusJSON = "{""checker_running"":" & IIf(IsVBSCheckerRunning(), "true", "false") & "}"
    Exit Function
ErrorHandler:
    GetVBSCheckerStatusJSON = "{""error"":""" & Err.description & """}"
End Function

' --- Check if Outlook is running ---
Public Function IsOutlookRunning() As Boolean
    On Error Resume Next
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number = 0 Then
        IsOutlookRunning = True
    Else
        IsOutlookRunning = False
    End If
    Set olApp = Nothing
    On Error GoTo 0
End Function

' --- Get unread email count from Inbox ---
Public Function GetUnreadMsgCount() As Long
    On Error GoTo ErrorHandler
    Dim olApp As Object, olNS As Object, olFolder As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(6) ' Inbox
    GetUnreadMsgCount = olFolder.UnReadItemCount
    Exit Function

ErrorHandler:
    GetUnreadMsgCount = -1
End Function



' --- Force a Full Outlook Scan ---
Public Sub ForceScanOutlook()
    On Error GoTo ErrorHandler
    DebugLog "ForceScanOutlook triggered at " & format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Optional: scan inbox messages
    Dim olApp As Object, olNS As Object, olFolder As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(6) ' Inbox
    
    Dim olItem As Object
    For Each olItem In olFolder.items
        ' Mark as read temporarily or just touch items
    Next olItem
    
    Set olFolder = Nothing: Set olNS = Nothing: Set olApp = Nothing
    Exit Sub
ErrorHandler:
    DebugLog "Error in ForceScanOutlook: " & Err.description
End Sub

Private m_backgroundEnabled As Boolean

Public Sub EnableBackgroundMode()
    m_backgroundEnabled = True
    DebugLog "Outlook background mode enabled"
End Sub

Public Sub DisableBackgroundMode()
    m_backgroundEnabled = False
    DebugLog "Outlook background mode disabled"
End Sub

' --- Central Outlook Command Dispatcher ---
Public Sub InitializeOutlookCheckingSimple()
    On Error GoTo ErrorHandler
    
    ' --- Check if Outlook is running ---
    If Not OutlookResponsive() Then
        DebugLog "Outlook is not running. Attempting to start..."
        Dim olApp As Object
        Set olApp = CreateObject("Outlook.Application")
        If olApp Is Nothing Then
            DebugLog "Failed to start Outlook. Aborting initialization."
            Exit Sub
        End If
        Set olApp = Nothing
        DebugLog "Outlook started successfully."
    End If
    
    ' --- Ensure SmartTraffic folder exists ---
    If Dir("C:\SmartTraffic", vbDirectory) = "" Then
        MkDir "C:\SmartTraffic"
        DebugLog "Created folder: C:\SmartTraffic"
    End If
    
    ' --- Create VBS script if missing ---
    If Dir("C:\SmartTraffic\OutlookChecker.vbs") = "" Then
        Call CreateOutlookVBSScript
        DebugLog "Created OutlookChecker.vbs script"
    End If
    
    ' --- Start VBS checker only if not already running ---
    If Not IsVBSCheckerRunning() Then
        Call StartOutlookVBSChecker
        DebugLog "Started VBS Outlook checker (PID: " & m_vbsProcessID & ")"
    Else
        DebugLog "VBS Outlook checker already running"
    End If
    
    DebugLog "InitializeOutlookChecking completed"
    Exit Sub
    
ErrorHandler:
    DebugLog "Error in InitializeOutlookChecking: " & Err.description
End Sub

Public Sub InitializeOutlookCheckingSafe()
    On Error GoTo ErrorHandler
    
    ' --- Step 1: Create VBS script if missing ---
    If Dir("C:\SmartTraffic\OutlookChecker.vbs") = "" Then
        Call CreateOutlookVBSScript
    End If
    
    ' --- Step 2: Start VBS checker if not already running ---
    If Not IsVBSCheckerRunning() Then
        StartOutlookVBSChecker
    Else
        DebugLog "VBS Outlook checker already running"
    End If
    
    ' --- Step 3: Optional: run rules once immediately ---
    On Error Resume Next
    RunOutlookRules
    On Error GoTo ErrorHandler
    
    DebugLog "InitializeOutlookCheckingSafe completed"
    Exit Sub
    
ErrorHandler:
    DebugLog "Error in InitializeOutlookCheckingSafe: " & Err.description
End Sub


Public Sub StopOutlookCheckingSafe()
    On Error GoTo ErrorHandler
    
    ' --- Step 1: Stop the VBS checker ---
    StopOutlookVBSChecker
    
    ' --- Step 2: Update any internal status trackers / dashboard ---
    ' Example: Last scan time cleared
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OutlookDashboard") ' adjust sheet name if different
    ws.Range("B2").value = "Stopped"            ' Status
    ws.Range("B3").value = "-"                  ' Last Scan
    ws.Range("B4").value = 0                    ' Unread messages
    ws.Range("B5").value = "Disabled"          ' Background Mode
    On Error GoTo ErrorHandler
    
    DebugLog "StopOutlookCheckingSafe completed"
    Exit Sub
    
ErrorHandler:
    DebugLog "Error in StopOutlookCheckingSafe: " & Err.description
End Sub

Public Sub UpdateOutlookBackgroundModex(action As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OutlookDashboard") ' adjust if your sheet name differs
    
    Select Case LCase(action)
        Case "enable_background"
            ' Start the background VBS checker
            StartOutlookVBSChecker
            
            ' Update dashboard
            ws.Range("B2").value = "Running"           ' Status
            ws.Range("B5").value = "Enabled"          ' Background Mode
            ws.Range("B3").value = format(Now, "yyyy-mm-dd hh:nn:ss") ' Last scan placeholder
            
            DebugLog "Background mode enabled"
            
        Case "disable_background"
            ' Stop background checker safely
            StopOutlookCheckingSafe
            
            DebugLog "Background mode disabled"
            
        Case Else
            DebugLog "Unknown action for background mode: " & action
    End Select
    
    Exit Sub
ErrorHandler:
    DebugLog "Error in UpdateOutlookBackgroundMode: " & Err.description
End Sub

Public Sub UpdateOutlookBackgroundMode(action As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OutlookDashboard") ' adjust if your sheet name differs
    
    Select Case LCase(action)
        Case "enable_background"
            ' Start the background VBS checker
            StartOutlookVBSChecker
            
            ' Update dashboard
            ws.Range("B2").value = "Running"           ' Status
            ws.Range("B5").value = "Enabled"          ' Background Mode
            ws.Range("B3").value = format(Now, "yyyy-mm-dd hh:nn:ss") ' Last scan placeholder
            
            DebugLog "Background mode enabled"
            
        Case "disable_background"
            ' Stop background checker safely
            StopOutlookCheckingSafe
            
            DebugLog "Background mode disabled"
            
        Case Else
            DebugLog "Unknown action for background mode: " & action
    End Select
    
    Exit Sub
ErrorHandler:
    DebugLog "Error in UpdateOutlookBackgroundMode: " & Err.description
End Sub


'===============================================
' Ensure Outlook is running (non-blocking)
'===============================================
Public Function EnsureOutlookRunning() As Object
    On Error Resume Next
    Dim olApp As Object
    
    ' Try to get a running instance
    Set olApp = GetObject(, "Outlook.Application")
    
    ' If not running, start Outlook
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
        ' Optional: give Outlook a moment to initialize
        Dim t As Single: t = Timer
        Do While Timer - t < 2
            DoEvents
        Loop
    End If
    
    Set EnsureOutlookRunning = olApp
    On Error GoTo 0
End Function

'===============================================
' Updated TestImmediateMemoryMode
'===============================================
Public Sub TestImmediateMemoryMode()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olFolder As Object, olItems As Object, olItem As Object, i As Long
    
    ' Ensure Outlook is running
    Set olApp = EnsureOutlookRunning()
    Set olFolder = olApp.GetNamespace("MAPI").GetDefaultFolder(6) ' Inbox
    Set olItems = olFolder.items
    olItems.Sort "[ReceivedTime]", True
    
    For i = 1 To WorksheetFunction.Min(5, olItems.count)
        Set olItem = olItems.item(i)
        DebugLog "Memory mode: Email " & i & " from " & olItem.SenderName & ", Subject: " & olItem.Subject
    Next i
    
    Set olItems = Nothing: Set olFolder = Nothing: Set olApp = Nothing
    DebugLog "TestImmediateMemoryMode executed successfully"
    Exit Sub
ErrorHandler:
    DebugLog "Error in TestImmediateMemoryMode: " & Err.description
End Sub

'===============================================
' Updated TestImmediateTempMode
'===============================================
Public Sub TestImmediateTempMode()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olTempFolder As Object
    
    ' Ensure Outlook is running
    Set olApp = EnsureOutlookRunning()
    Set olTempFolder = olApp.GetNamespace("MAPI").GetDefaultFolder(3) ' Deleted Items
    DebugLog "Temp mode: Found " & olTempFolder.items.count & " items in Deleted Items"
    
    Set olTempFolder = Nothing: Set olApp = Nothing
    DebugLog "TestImmediateTempMode executed successfully"
    Exit Sub
ErrorHandler:
    DebugLog "Error in TestImmediateTempMode: " & Err.description
End Sub

'===============================================
' Updated TestBackgroundMode
'===============================================
Public Sub TestBackgroundMode()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olSession As Object
    
    ' Ensure Outlook is running
    Set olApp = EnsureOutlookRunning()
    Set olSession = olApp.GetNamespace("MAPI")
    olSession.SendAndReceive False
    DebugLog "Background mode: Initiated Outlook sync"
    
    Set olSession = Nothing: Set olApp = Nothing
    DebugLog "TestBackgroundMode executed successfully"
    Exit Sub
ErrorHandler:
    DebugLog "Error in TestBackgroundMode: " & Err.description
End Sub


'===============================================
' Stop the VBS Outlook Checker
'===============================================
Public Sub StopOutlookVBSChecker()
    On Error Resume Next
    
    If m_vbsProcessID > 0 Then
        shell "taskkill /F /PID " & m_vbsProcessID, vbHide
        m_vbsProcessID = 0
        DebugLog "Stopped VBS Outlook checker"
    Else
        ' Fallback: kill any running instance by window title
        shell "taskkill /F /FI ""WINDOWTITLE eq OutlookChecker.vbs*""", vbHide
        DebugLog "Attempted to stop any running VBS checker processes"
    End If
    
    On Error GoTo 0
End Sub

'===============================================
' Check if the VBS Checker is running
'===============================================
Public Function IsVBSCheckerRunning() As Boolean
    On Error Resume Next
    
    Dim wmi As Object, processes As Object, proc As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='cscript.exe'")
    
    For Each proc In processes
        If InStr(proc.commandLine, "OutlookChecker.vbs") > 0 Then
            IsVBSCheckerRunning = True
            Exit Function
        End If
    Next
    IsVBSCheckerRunning = False
End Function

'===============================================
' Create the VBS script for background checking
'===============================================
Public Sub CreateOutlookVBSScriptBackground(Optional ByVal vbsPath As String = "C:\SmartTraffic\OutlookChecker.vbs")
    On Error GoTo ErrorHandler
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure folder exists
    If Not fso.FolderExists(fso.GetParentFolderName(vbsPath)) Then
        fso.CreateFolder fso.GetParentFolderName(vbsPath)
    End If
    
    ' Create the VBS file
    Set ts = fso.CreateTextFile(vbsPath, True)
    ts.WriteLine "Do"
    ts.WriteLine "    On Error Resume Next"
    ts.WriteLine "    Set xlApp = GetObject(, ""Excel.Application"")"
    ts.WriteLine "    If xlApp Is Nothing Then Set xlApp = CreateObject(""Excel.Application"")"
    ts.WriteLine "    xlApp.Run ""OutlookExtra.InitializeOutlookChecking"""
    ts.WriteLine "    WScript.Sleep 60000 ' 60 seconds"
    ts.WriteLine "Loop"
    ts.Close
    
    DebugLog "VBS script written to " & vbsPath
    Exit Sub
ErrorHandler:
    DebugLog "Error creating VBS script: " & Err.description
End Sub

'===============================================
' Initialize Outlook Checking (modern, non-blocking)
'===============================================
Public Sub InitializeOutlookChecking()
    On Error GoTo ErrorHandler
    
    ' Ensure VBS script is running
    If Not IsVBSCheckerRunning() Then
        StartOutlookVBSChecker
    End If
    
    ' Run an initial background check immediately
    Call CheckOutlookBackground
    
    ' Optional: run rules once on startup
    If OutlookResponsive() Then
        RunOutlookRules
    End If
    
    DebugLog "InitializeOutlookChecking completed"
    Exit Sub
ErrorHandler:
    DebugLog "Error in InitializeOutlookChecking: " & Err.description
End Sub

Public Sub CreateOutlookVBSScript()
    On Error GoTo ErrorHandler
    
    Dim fso As Object, ts As Object
    Dim vbsPath As String
    Dim wbPath As String
    Dim vbsCode As String
    
    ' Path to the VBS file
    vbsPath = "C:\SmartTraffic\OutlookChecker.vbs"
    
    ' Path to your workbook
    wbPath = ThisWorkbook.FullName
    
    ' Create FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure folder exists
    If Not fso.FolderExists("C:\SmartTraffic") Then
        fso.CreateFolder "C:\SmartTraffic"
    End If
    
    ' VBS script content with stop flag support
    vbsCode = "'==================================================" & vbCrLf
    vbsCode = vbsCode & "' OutlookChecker.vbs - auto-generated with stop flag" & vbCrLf
    vbsCode = vbsCode & "'==================================================" & vbCrLf
    vbsCode = vbsCode & "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objExcel, objWorkbook, IntervalSeconds, stopFile" & vbCrLf
    vbsCode = vbsCode & "IntervalSeconds = 60 ' check every 60 seconds" & vbCrLf
    vbsCode = vbsCode & "stopFile = ""C:\SmartTraffic\OutlookChecker.stop""" & vbCrLf
    vbsCode = vbsCode & "Set objExcel = CreateObject(""Excel.Application"")" & vbCrLf
    vbsCode = vbsCode & "objExcel.Visible = False" & vbCrLf
    vbsCode = vbsCode & "Set objWorkbook = objExcel.Workbooks.Open(""" & wbPath & """)" & vbCrLf
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "    If CreateObject(""Scripting.FileSystemObject"").FileExists(stopFile) Then" & vbCrLf
    vbsCode = vbsCode & "        objWorkbook.Close False" & vbCrLf
    vbsCode = vbsCode & "        objExcel.Quit" & vbCrLf
    vbsCode = vbsCode & "        WScript.Quit" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    On Error Resume Next" & vbCrLf
    vbsCode = vbsCode & "    objExcel.Run ""OutlookExtra.InitializeOutlookChecking""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep IntervalSeconds * 1000" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    
    ' Write VBS file
    Set ts = fso.CreateTextFile(vbsPath, True)
    ts.Write vbsCode
    ts.Close
    
    DebugLog "OutlookChecker.vbs created at " & vbsPath
    Exit Sub
ErrorHandler:
    DebugLog "Error creating VBS script: " & Err.description
End Sub

Public Sub StopOutlookChecking()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile "C:\SmartTraffic\OutlookChecker.stop", True
    DebugLog "Outlook background check flagged to stop"
End Sub
'==================================================
' Process Outlook Rules in the main server loop
'==================================================
Public Sub ProcessOutlookRules()
    On Error GoTo ErrorHandler
    
    ' --- Ensure Outlook is running ---
    Dim olApp As Object
    Set olApp = EnsureOutlookRunning()
    If olApp Is Nothing Then Exit Sub
    
    ' --- Decide execution mode ---
    ' Background VBS checker handles periodic scans
    ' This loop call is "immediate" for live updates
    Call RunOutlookRules  ' Executes all enabled rules in the sheet
    
    ' --- Optional: log last scan time ---
    m_lastScanTime = Now
    DebugLog "[Outlook] ProcessOutlookRules executed at " & format(m_lastScanTime, "yyyy-mm-dd hh:nn:ss")
    
    ' --- Clean up objects ---
    Set olApp = Nothing
    Exit Sub

ErrorHandler:
    DebugLog "[Outlook] Error in ProcessOutlookRules: " & Err.description
End Sub

' Generate main Outlook status page
Public Function GenerateOutlookStatusPageINWEBUI() As String
    On Error GoTo ErrorHandler
    
    Dim html As String
    Dim isRunning As Boolean
    Dim unreadCount As Long
    Dim checkerRunning As Boolean
    
    ' Get current status
    isRunning = IsOutlookRunning()
    unreadCount = GetUnreadMsgCount()
    checkerRunning = IsVBSCheckerRunning()
    
    html = GenerateHTMLHeader("Outlook Control Panel")
    html = html & "<div class='container'>"
    html = html & "<h1 class='header'>Outlook Control Panel</h1>"
    
    ' Status section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Current Status</div>"
    html = html & "<p>Outlook Running: " & IIf(isRunning, "YES", "NO") & "</p>"
    html = html & "<p>Unread Messages: " & IIf(unreadCount >= 0, CStr(unreadCount), "Error") & "</p>"
    html = html & "<p>Background Checker: " & IIf(checkerRunning, "RUNNING", "STOPPED") & "</p>"
    html = html & "<p>Last Scan: " & IIf(m_lastScanTime > 0, format(m_lastScanTime, "yyyy-mm-dd hh:mm:ss"), "Never") & "</p>"
    html = html & "</div>"
    
    ' Control actions
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Control Actions</div>"
    html = html & "<p><a href='/outlook/force_check' class='btn'>Force Check Now</a></p>"
    html = html & "<p><a href='/outlook/enable_background' class='btn'>Enable Background Mode</a></p>"
    html = html & "<p><a href='/outlook/disable_background' class='btn'>Disable Background Mode</a></p>"
    html = html & "<p><a href='/outlook/initialize' class='btn'>Initialize Checking</a></p>"
    html = html & "<p><a href='/outlook/stop' class='btn'>Stop Checking</a></p>"
    html = html & "</div>"
    
    ' Testing section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Testing & Diagnostics</div>"
    html = html & "<p><a href='/outlook/diagnose' class='btn'>Run Diagnostics</a></p>"
    html = html & "<p><a href='/outlook/test_connection' class='btn'>Test Connection</a></p>"
    html = html & "<p><a href='/outlook/test_rules' class='btn'>Test Rules</a></p>"
    html = html & "<p><a href='/outlook/run_rules' class='btn'>Run Rules Now</a></p>"
    html = html & "</div>"
    
    ' Test modes section
    html = html & "<div class='section'>"
    html = html & "<div class='section-title'>Test Modes</div>"
    html = html & "<p><a href='/outlook/test_memory' class='btn'>Test Memory Mode</a></p>"
    html = html & "<p><a href='/outlook/test_temp' class='btn'>Test Temp Mode</a></p>"
    html = html & "<p><a href='/outlook/test_background' class='btn'>Test Background Mode</a></p>"
    html = html & "</div>"
    
    ' Navigation
    html = html & "<div class='section'>"
    html = html & "<a href='/' class='btn'>Home</a>"
    html = html & "<a href='/dashboard' class='btn'>Dashboard</a>"
    html = html & "</div>"
    
    html = html & "<div class='bar'></div>"
    html = html & "</div></body></html>"
    
    GenerateOutlookStatusPage = html
    Exit Function
    
ErrorHandler:
    DebugLog "Error in GenerateOutlookStatusPage: " & Err.description
    GenerateOutlookStatusPage = GenerateErrorPage("Error generating Outlook status page: " & Err.description)
End Function

' --- Safe wrapper for Force Scan ---
Public Sub ForceOutlookScanSafe()
    On Error GoTo ErrHandler
    ' Call your existing Force Scan routine here
    ForceOutlookScan
    Exit Sub
ErrHandler:
    DebugLog "Error in ForceOutlookScanSafe: " & Err.description
End Sub

' --- Safe wrapper for Run Rules ---
Public Sub RunOutlookRulesSafe()
    On Error GoTo ErrHandler
    ' Call your existing Run Rules routine here
    RunOutlookRules
    Exit Sub
ErrHandler:
    DebugLog "Error in RunOutlookRulesSafe: " & Err.description
End Sub

' --- Safe wrapper for enabling/disabling background mode ---
Public Sub SetBackgroundModeSafe(enable As Boolean)
    On Error GoTo ErrHandler
    ' Call your existing SetBackgroundMode routine here
    SetBackgroundMode enable
    Exit Sub
ErrHandler:
    DebugLog "Error in SetBackgroundModeSafe: " & Err.description
End Sub

' --- Safe wrapper to get unread count ---
Public Function GetUnreadCountSafe() As Long
    On Error GoTo ErrHandler
    GetUnreadCountSafe = GetUnreadCount
    Exit Function
ErrHandler:
    DebugLog "Error in GetUnreadCountSafe: " & Err.description
    GetUnreadCountSafe = -1
End Function

' --- Safe wrapper to get Outlook status ---
Public Function GetOutlookStatusSafe() As String
    On Error GoTo ErrHandler
    GetOutlookStatusSafe = GetOutlookStatus
    Exit Function
ErrHandler:
    DebugLog "Error in GetOutlookStatusSafe: " & Err.description
    GetOutlookStatusSafe = "Error"
End Function

' --- Safe wrapper to get background status ---
Public Function GetBackgroundStatusSafe() As String
    On Error GoTo ErrHandler
    If IsBackgroundEnabled Then
        GetBackgroundStatusSafe = "Enabled"
    Else
        GetBackgroundStatusSafe = "Disabled"
    End If
    Exit Function
ErrHandler:
    DebugLog "Error in GetBackgroundStatusSafe: " & Err.description
    GetBackgroundStatusSafe = "Unknown"
End Function

' --- Safe wrapper for VBS checker status ---
Public Function GetVBSStatusSafe() As String
    On Error GoTo ErrHandler
    GetVBSStatusSafe = CheckVBSSafe
    Exit Function
ErrHandler:
    DebugLog "Error in GetVBSStatusSafe: " & Err.description
    GetVBSStatusSafe = "Error"
End Function

' --- Safe wrapper to get recent Outlook emails ---
Public Function GetOutlookDataSafe() As String
    On Error GoTo ErrHandler
    GetOutlookDataSafe = GetRecentEmails
    Exit Function
ErrHandler:
    DebugLog "Error in GetOutlookDataSafe: " & Err.description
    GetOutlookDataSafe = "Unable to retrieve emails"
End Function

