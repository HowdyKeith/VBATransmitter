Option Explicit

'***************************************************************
' ExternalOutlookManager Module
' Purpose: Manage Outlook through external VBS scripts and .otm macros
'          to prevent OLE blocking Excel. Creates temp files for
'          communication and reads results back.
'***************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const TEMP_FOLDER_PREFIX As String = "ExcelOutlookBridge_"
Private Const MAX_WAIT_TIME As Long = 10000  ' 10 seconds max wait

' Get dedicated temp folder for our bridge files
Private Function GetBridgeTempPath() As String
    Dim tempPath As String
    Dim bridgePath As String
    Dim fso As Object
    
    tempPath = GetSystemTempPath()
    bridgePath = tempPath & TEMP_FOLDER_PREFIX & format(Now, "yyyymmdd") & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(bridgePath) Then
        fso.CreateFolder bridgePath
    End If
    
    GetBridgeTempPath = bridgePath
    Set fso = Nothing
End Function

Private Function GetSystemTempPath() As String
    Dim tempPath As String
    Dim Length As Long
    
    tempPath = String(260, 0)
    Length = GetTempPath(260, tempPath)
    
    If Length > 0 Then
        GetSystemTempPath = Left(tempPath, Length)
    Else
        GetSystemTempPath = Environ("TEMP") & "\"
    End If
End Function

'=======================================================
' VBS SCRIPT GENERATORS
'=======================================================

' Create VBS to launch Outlook without blocking Excel
Public Sub CreateOutlookLauncherVBS()
    Dim vbsPath As String
    Dim fso As Object
    Dim f As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    vbsPath = GetBridgeTempPath() & "LaunchOutlook.vbs"
    
    Set f = fso.CreateTextFile(vbsPath, True)
    f.WriteLine "' Outlook Launcher - Non-blocking"
    f.WriteLine "On Error Resume Next"
    f.WriteLine ""
    f.WriteLine "' Check if Outlook is already running"
    f.WriteLine "Set objWMI = GetObject(""winmgmts:\\.\root\cimv2"")"
    f.WriteLine "Set colProcesses = objWMI.ExecQuery(""SELECT * FROM Win32_Process WHERE Name = 'OUTLOOK.EXE'"")"
    f.WriteLine ""
    f.WriteLine "If colProcesses.Count > 0 Then"
    f.WriteLine "    WScript.Echo ""Outlook already running"""
    f.WriteLine "    WScript.Quit 0"
    f.WriteLine "End If"
    f.WriteLine ""
    f.WriteLine "' Launch Outlook"
    f.WriteLine "Set objShell = CreateObject(""WScript.Shell"")"
    f.WriteLine "objShell.Run ""outlook.exe"", 1, False"
    f.WriteLine "WScript.Echo ""Outlook launched"""
    f.WriteLine "WScript.Sleep 3000"
    f.Close
    
    DebuggingLog.DebugLog "[ExternalOutlookManager] Created launcher VBS: " & vbsPath
End Sub

' Create VBS to get unread count without blocking Excel
Public Sub CreateUnreadCountVBS()
    Dim vbsPath As String
    Dim resultPath As String
    Dim fso As Object
    Dim f As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    vbsPath = GetBridgeTempPath() & "GetUnreadCount.vbs"
    resultPath = GetBridgeTempPath() & "UnreadCount.txt"
    
    Set f = fso.CreateTextFile(vbsPath, True)
    f.WriteLine "' Get Outlook Unread Count - Non-blocking"
    f.WriteLine "On Error Resume Next"
    f.WriteLine ""
    f.WriteLine "Dim resultPath"
    f.WriteLine "resultPath = """ & resultPath & """"
    f.WriteLine ""
    f.WriteLine "' Try to connect to existing Outlook instance"
    f.WriteLine "Set olApp = GetObject(, ""Outlook.Application"")"
    f.WriteLine "If Err.Number <> 0 Then"
    f.WriteLine "    ' Write error to result file"
    f.WriteLine "    Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "    Set f = fso.CreateTextFile(resultPath, True)"
    f.WriteLine "    f.WriteLine ""ERROR: Outlook not running"""
    f.WriteLine "    f.Close"
    f.WriteLine "    WScript.Quit 1"
    f.WriteLine "End If"
    f.WriteLine ""
    f.WriteLine "' Get inbox and count unread items"
    f.WriteLine "Set olNS = olApp.GetNamespace(""MAPI"")"
    f.WriteLine "Set olInbox = olNS.GetDefaultFolder(6) ' olFolderInbox = 6"
    f.WriteLine "unreadCount = olInbox.UnreadItemCount"
    f.WriteLine ""
    f.WriteLine "' Write result to file"
    f.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "Set f = fso.CreateTextFile(resultPath, True)"
    f.WriteLine "f.WriteLine ""UNREAD_COUNT:"" & unreadCount"
    f.WriteLine "f.Close"
    f.WriteLine ""
    f.WriteLine "WScript.Echo ""Unread count: "" & unreadCount"
    f.Close
    
    DebuggingLog.DebugLog "[ExternalOutlookManager] Created unread count VBS: " & vbsPath
End Sub

' Create VBS to run Outlook rules without blocking Excel
Public Sub CreateRunRuleVBS(ByVal ruleName As String)
    Dim vbsPath As String
    Dim resultPath As String
    Dim fso As Object
    Dim f As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    vbsPath = GetBridgeTempPath() & "RunRule_" & CleanFileName(ruleName) & ".vbs"
    resultPath = GetBridgeTempPath() & "RuleResult_" & CleanFileName(ruleName) & ".txt"
    
    Set f = fso.CreateTextFile(vbsPath, True)
    f.WriteLine "' Run Outlook Rule - Non-blocking"
    f.WriteLine "On Error Resume Next"
    f.WriteLine ""
    f.WriteLine "Dim resultPath, ruleName"
    f.WriteLine "resultPath = """ & resultPath & """"
    f.WriteLine "ruleName = """ & ruleName & """"
    f.WriteLine ""
    f.WriteLine "' Connect to Outlook"
    f.WriteLine "Set olApp = GetObject(, ""Outlook.Application"")"
    f.WriteLine "If Err.Number <> 0 Then"
    f.WriteLine "    Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "    Set f = fso.CreateTextFile(resultPath, True)"
    f.WriteLine "    f.WriteLine ""ERROR: Outlook not running"""
    f.WriteLine "    f.Close"
    f.WriteLine "    WScript.Quit 1"
    f.WriteLine "End If"
    f.WriteLine ""
    f.WriteLine "Set olNS = olApp.GetNamespace(""MAPI"")"
    f.WriteLine "Set olRules = olNS.DefaultStore.GetRules()"
    f.WriteLine ""
    f.WriteLine "' Find and execute the rule"
    f.WriteLine "For Each olRule In olRules"
    f.WriteLine "    If olRule.Name = ruleName Then"
    f.WriteLine "        If olRule.Enabled Then"
    f.WriteLine "            ' Execute the rule on Inbox"
    f.WriteLine "            Set olInbox = olNS.GetDefaultFolder(6)"
    f.WriteLine "            olRule.Execute ShowProgress:=False, Folder:=olInbox"
    f.WriteLine "            Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "            Set f = fso.CreateTextFile(resultPath, True)"
    f.WriteLine "            f.WriteLine ""SUCCESS: Rule executed - "" & ruleName"
    f.WriteLine "            f.Close"
    f.WriteLine "            WScript.Echo ""Rule executed: "" & ruleName"
    f.WriteLine "            WScript.Quit 0"
    f.WriteLine "        Else"
    f.WriteLine "            Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "            Set f = fso.CreateTextFile(resultPath, True)"
    f.WriteLine "            f.WriteLine ""ERROR: Rule is disabled - "" & ruleName"
    f.WriteLine "            f.Close"
    f.WriteLine "            WScript.Quit 2"
    f.WriteLine "        End If"
    f.WriteLine "    End If"
    f.WriteLine "Next"
    f.WriteLine ""
    f.WriteLine "' Rule not found"
    f.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "Set f = fso.CreateTextFile(resultPath, True)"
    f.WriteLine "f.WriteLine ""ERROR: Rule not found - "" & ruleName"
    f.WriteLine "f.Close"
    f.Close
    
    DebuggingLog.DebugLog "[ExternalOutlookManager] Created rule execution VBS: " & vbsPath
End Sub

'=======================================================
' EXECUTION AND RESULT READING
'=======================================================

' Launch Outlook without blocking Excel
Public Sub LaunchOutlookExternal()
    On Error GoTo ErrorHandler
    
    Call CreateOutlookLauncherVBS
    Dim vbsPath As String
    vbsPath = GetBridgeTempPath() & "LaunchOutlook.vbs"
    
    ' Execute asynchronously
    shell "cscript.exe """ & vbsPath & """ //NoLogo", vbHide
    
    DebuggingLog.DebugLog "[ExternalOutlookManager] Outlook launch initiated"
    Exit Sub
    
ErrorHandler:
    DebuggingLog.DebugLog "[ExternalOutlookManager] Launch error: " & Err.description
End Sub

' Get unread count without blocking Excel
Public Function GetUnreadCountExternal() As String
    On Error GoTo ErrorHandler
    
    Call CreateUnreadCountVBS
    Dim vbsPath As String, resultPath As String
    Dim fso As Object, f As Object
    Dim result As String
    Dim waitTime As Long
    
    vbsPath = GetBridgeTempPath() & "GetUnreadCount.vbs"
    resultPath = GetBridgeTempPath() & "UnreadCount.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Delete old result file
    If fso.FileExists(resultPath) Then fso.DeleteFile resultPath
    
    ' Execute VBS asynchronously
    shell "cscript.exe """ & vbsPath & """ //NoLogo", vbHide
    
    ' Wait for result file (non-blocking wait)
    waitTime = 0
    Do While waitTime < MAX_WAIT_TIME And Not fso.FileExists(resultPath)
        Sleep 100
        waitTime = waitTime + 100
        DoEvents ' Allow Excel to remain responsive
    Loop
    
    ' Read result
    If fso.FileExists(resultPath) Then
        Set f = fso.OpenTextFile(resultPath, 1) ' ForReading
        result = f.ReadAll
        f.Close
        GetUnreadCountExternal = Trim(result)
    Else
        GetUnreadCountExternal = "TIMEOUT: No response from Outlook"
    End If
    
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    GetUnreadCountExternal = "ERROR: " & Err.description
    DebuggingLog.DebugLog "[ExternalOutlookManager] GetUnreadCount error: " & Err.description
End Function

' Run Outlook rule without blocking Excel
Public Function RunOutlookRuleExternal(ByVal ruleName As String) As String
    On Error GoTo ErrorHandler
    
    Call CreateRunRuleVBS(ruleName)
    Dim vbsPath As String, resultPath As String
    Dim fso As Object, f As Object
    Dim result As String
    Dim waitTime As Long
    
    vbsPath = GetBridgeTempPath() & "RunRule_" & CleanFileName(ruleName) & ".vbs"
    resultPath = GetBridgeTempPath() & "RuleResult_" & CleanFileName(ruleName) & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Delete old result file
    If fso.FileExists(resultPath) Then fso.DeleteFile resultPath
    
    ' Execute VBS asynchronously
    shell "cscript.exe """ & vbsPath & """ //NoLogo", vbHide
    
    ' Wait for result file (non-blocking wait)
    waitTime = 0
    Do While waitTime < MAX_WAIT_TIME And Not fso.FileExists(resultPath)
        Sleep 200
        waitTime = waitTime + 200
        DoEvents ' Keep Excel responsive
    Loop
    
    ' Read result
    If fso.FileExists(resultPath) Then
        Set f = fso.OpenTextFile(resultPath, 1)
        result = f.ReadAll
        f.Close
        RunOutlookRuleExternal = Trim(result)
    Else
        RunOutlookRuleExternal = "TIMEOUT: No response from Outlook"
    End If
    
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    RunOutlookRuleExternal = "ERROR: " & Err.description
    DebuggingLog.DebugLog "[ExternalOutlookManager] RunRule error: " & Err.description
End Function

'=======================================================
' WEB INTERFACE FUNCTIONS
'=======================================================

' Get unread count for web pages (JSON format)
Public Function GetUnreadCountForWeb() As String
    Dim result As String
    Dim unreadCount As Long
    
    result = GetUnreadCountExternal()
    
    If InStr(result, "UNREAD_COUNT:") > 0 Then
        unreadCount = CLng(Replace(result, "UNREAD_COUNT:", ""))
        GetUnreadCountForWeb = "{""status"":""success"",""unread_count"":" & unreadCount & ",""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
    Else
        GetUnreadCountForWeb = "{""status"":""error"",""message"":""" & Replace(result, """", """""") & """,""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
    End If
End Function

' Execute rule for web interface (JSON format)
Public Function ExecuteRuleForWeb(ByVal ruleName As String) As String
    Dim result As String
    
    result = RunOutlookRuleExternal(ruleName)
    
    If InStr(result, "SUCCESS:") > 0 Then
        ExecuteRuleForWeb = "{""status"":""success"",""message"":""" & Replace(result, """", """""") & """,""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
    Else
        ExecuteRuleForWeb = "{""status"":""error"",""message"":""" & Replace(result, """", """""") & """,""timestamp"":""" & format(Now, "yyyy-mm-dd hh:mm:ss") & """}"
    End If
End Function

'=======================================================
' UTILITY FUNCTIONS
'=======================================================

' Clean filename for safe file operations
Private Function CleanFileName(ByVal fileName As String) As String
    Dim cleaned As String
    cleaned = fileName
    cleaned = Replace(cleaned, " ", "_")
    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "\", "_")
    cleaned = Replace(cleaned, ":", "_")
    cleaned = Replace(cleaned, "*", "_")
    cleaned = Replace(cleaned, "?", "_")
    cleaned = Replace(cleaned, """", "_")
    cleaned = Replace(cleaned, "<", "_")
    cleaned = Replace(cleaned, ">", "_")
    cleaned = Replace(cleaned, "|", "_")
    CleanFileName = cleaned
End Function

' Cleanup temp files
Public Sub CleanupBridgeFiles()
    On Error Resume Next
    Dim fso As Object
    Dim bridgePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    bridgePath = GetBridgeTempPath()
    
    If fso.FolderExists(bridgePath) Then
        fso.DeleteFolder bridgePath, True
        DebuggingLog.DebugLog "[ExternalOutlookManager] Cleaned up bridge temp files"
    End If
    
    Set fso = Nothing
End Sub

' Check if Outlook is running (via process list, no OLE)
Public Function IsOutlookRunningExternal() As Boolean
    On Error Resume Next
    Dim objWMI As Object
    Dim colProcesses As Object
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'OUTLOOK.EXE'")
    
    IsOutlookRunningExternal = (colProcesses.count > 0)
    
    Set colProcesses = Nothing
    Set objWMI = Nothing
End Function

Private Sub ApplLinks()

'' Add this to your AppLaunch module's HandleAppRequest function
' to handle the new Outlook routes that don't block Excel

' In your HandleAppRequest subroutine, add these cases:

Case "/outlook"
    responseBody = OutlookWebHandler.GenerateOutlookDashboard()

Case "/outlook/unread"
    responseBody = OutlookWebHandler.GetUnreadCountAPI()

Case "/outlook/status"
    responseBody = OutlookWebHandler.GetOutlookStatusAPI()

Case "/outlook/launch"
    responseBody = OutlookWebHandler.LaunchOutlookAPI()

Case "/outlook/rules"
    responseBody = OutlookWebHandler.GenerateRulesPage()

Case Else
    ' Handle rule execution with parameters
    If InStr(requestPath, "/outlook/execute_rule") = 1 Then
        responseBody = OutlookWebHandler.HandleRuleExecution(requestPath)
    Else
        ' Your existing default handling
        responseBody = "<h1>404 Not Found</h1><p>Page not found: " & requestPath & "</p>"
    End If
End Sub

