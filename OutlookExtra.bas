Option Explicit

'***************************************************************
' OutlookExtra Module - Extended Outlook Functions
' Purpose: Advanced Outlook features, rules, background checks,
'          attachments handling, and complex operations
'***************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function IsHungAppWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function IsHungAppWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public m_vbsProcessID As Long

Private m_backgroundEnabled As Boolean
Private m_lastScanTime As Date

'===============================================
' Outlook Rule Management
'===============================================
Public Sub RunOutlookRules()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olSession As Object, olRules As Object, olRule As Object
    Dim olInbox As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olSession = olApp.GetNamespace("MAPI")
    Set olInbox = olSession.GetDefaultFolder(6) ' Inbox - restrict rules to here
    Set olRules = olSession.DefaultStore.GetRules
    
    DebugLog "Starting execution of " & olRules.count & " built-in Outlook rules"
    
    For Each olRule In olRules
        If olRule.Enabled Then  ' Only execute enabled rules
            On Error Resume Next
            olRule.Execute ShowProgress:=False, folder:=olInbox, IncludeSubfolders:=False
            If Err.Number = 0 Then
                DebugLog "Executed built-in Outlook rule: " & olRule.Name & " at " & format(Now, "yyyy-mm-dd hh:mm:ss")
            Else
                DebugLog "Failed to execute rule " & olRule.Name & ": " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Else
            DebugLog "Skipped disabled rule: " & olRule.Name
        End If
    Next
    
    Set olRules = Nothing: Set olInbox = Nothing: Set olSession = Nothing: Set olApp = Nothing
    DebugLog "All built-in rules execution complete"
    Exit Sub
ErrorHandler:
    DebugLog "Error in RunOutlookRules: " & Err.description
End Sub

Public Sub ExecuteOutlookRule(ByVal ruleName As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olNamespace As Object, olRules As Object, olRule As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olRules = olNamespace.DefaultStore.GetRules
    
    For Each olRule In olRules
        If olRule.Name = ruleName Then
            olRule.Execute
            DebugLog "Executed rule: " & ruleName
            Exit Sub
        End If
    Next
    
    DebugLog "Rule not found: " & ruleName
    Exit Sub
ErrorHandler:
    DebugLog "Error in ExecuteOutlookRule: " & Err.description
End Sub

'===============================================
' Universal Email Checker (with attachments & forwarding)
'===============================================
Public Sub UniversalEmailChecker(Optional ByVal attachmentMode As String = "Memory", _
                                 Optional ByVal background As Boolean = False, _
                                 Optional ByVal useCollections As Boolean = True, _
                                 Optional ByVal forwardBody As Boolean = True, _
                                 Optional ByVal hoursBack As Long = 6)
    On Error GoTo ErrorHandler
    
    If background Then
        Call CheckOutlookBackground  ' Existing background sync
    End If
    
    Dim olApp As Object, olNamespace As Object, olFolder As Object
    Dim xlWs As Worksheet, ruleData As Variant, numRules As Long, processedCount As Long
    Dim i As Long, filter As String
    
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(6) ' Inbox
    
    ' Load rules from sheet
    On Error Resume Next
    Set xlWs = ThisWorkbook.Worksheets("Outlook")
    If xlWs Is Nothing Then
        DebugLog "Outlook rules sheet not found - aborting"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    numRules = xlWs.Cells(xlWs.Rows.count, 1).End(xlUp).row - 1
    If numRules <= 0 Then
        DebugLog "No custom rules defined in sheet"
        Exit Sub
    End If
    ruleData = xlWs.Range("A2:I" & (numRules + 1)).value  ' A:RuleName, B:Enabled, C:SubjectContains, D:Action, E:ForwardTo, F:SenderContains, G:BodyContains, H:LastExecuted, I:? (unused)
    
    DebugLog "Processing " & numRules & " custom rules"
    processedCount = 0
    
    For i = 1 To numRules
        If UCase(ruleData(i, 2)) <> "TRUE" Then  ' Skip if not enabled (B)
            DebugLog "Skipped disabled rule: " & ruleData(i, 1)
            Continue For
        End If
        
        ' Build dynamic filter based on conditions
        filter = ""  ' Start with unread/recent filter
        If hoursBack > 0 Then filter = filter & "([ReceivedTime] > '" & format(DateAdd("h", -hoursBack, Now), "yyyy-mm-dd hh:mm") & "') AND "
        filter = filter & "([Unread] = True)"
        
        If ruleData(i, 3) <> "" Then  ' Subject Contains (C)
            filter = filter & " AND (urn:schemas:httpmail:subject LIKE '%" & Replace(ruleData(i, 3), "'", "''") & "%')"
        End If
        If ruleData(i, 6) <> "" Then  ' Sender Contains (F)
            filter = filter & " AND (urn:schemas:httpmail:fromname LIKE '%" & Replace(ruleData(i, 6), "'", "''") & "%')"
        End If
        If ruleData(i, 7) <> "" Then  ' Body Contains (G)
            filter = filter & " AND (urn:schemas:httpmail:textdescription LIKE '%" & Replace(ruleData(i, 7), "'", "''") & "%')"
        End If
        
        ' Apply filter
        Dim olItems As Object
        Set olItems = olFolder.items.Restrict("@SQL=" & filter)
        DebugLog "Rule '" & ruleData(i, 1) & "': Found " & olItems.count & " matching emails"
        
        Dim olMail As Object
        For Each olMail In olItems
            If olMail.Class = 43 Then  ' MailItem
                ' Perform action based on D:Action (expand cases as needed)
                Select Case UCase(ruleData(i, 4))  ' Action
                    Case "FORWARD"
                        ProcessEmailForwarding olMail, ruleData(i, 5), ruleData(i, 5), forwardBody, i, xlWs  ' E:ForwardTo
                    Case "SAVEATTACHMENTS"
                        ProcessEmailAttachments olMail, attachmentMode, useCollections, Nothing, processedCount, i, xlWs
                    Case "DELETE"
                        olMail.Delete
                        DebugLog "Deleted email: " & olMail.Subject
                    Case Else
                        DebugLog "Unknown action for rule " & ruleData(i, 1) & ": " & ruleData(i, 4)
                End Select
                
                ' Mark as read/processed
                olMail.unread = False
                olMail.Save
                processedCount = processedCount + 1
            End If
        Next olMail
        
        ' Update last executed (H)
        xlWs.Cells(i + 1, 8).value = Now
    Next i
    
    ThisWorkbook.Save
    DebugLog "UniversalEmailChecker processed " & processedCount & " emails across all rules"
    Exit Sub
ErrorHandler:
    DebugLog "Error in UniversalEmailChecker: " & Err.description
End Sub

' Helper function for logging (ensure this exists in OutlookExtra or a shared module)
Private Sub DebugLog(ByVal message As String)
    Debug.Print "[" & format(Now, "yyyy-mm-dd hh:mm:ss") & "] " & message
    On Error Resume Next
    Dim fso As Object, logFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile("C:\SmartTraffic\outlook_log.txt", 8, True)
    logFile.WriteLine "[" & format(Now, "yyyy-mm-dd hh:mm:ss") & "] " & message
    logFile.Close
    On Error GoTo 0
End Sub

'===============================================
' Process Email Attachments
'===============================================
Private Sub ProcessEmailAttachments(olMail As Object, attachmentMode As String, useCollections As Boolean, _
                                   attachmentCollection As Object, ByRef processedCount As Long, _
                                   ruleIndex As Long, xlWs As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim olAttachment As Object, fso As Object, tempFilePath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each olAttachment In olMail.Attachments
        tempFilePath = fso.GetSpecialFolder(2) & "\" & olAttachment.fileName
        olAttachment.SaveAsFile tempFilePath
        processedCount = processedCount + 1
        xlWs.Cells(ruleIndex + 1, 10).value = xlWs.Cells(ruleIndex + 1, 10).value & olAttachment.fileName & "; "
    Next
    
    Exit Sub
ErrorHandler:
    DebugLog "Error in ProcessEmailAttachments: " & Err.description
End Sub

'===============================================
' Process Email Forwarding
'===============================================
Private Sub ProcessEmailForwarding(olMail As Object, forwardTo As String, sanitizedForwardTo As String, _
                                  forwardBody As Boolean, ruleIndex As Long, xlWs As Worksheet)
    On Error GoTo ErrorHandler
    
    If forwardTo <> "" And forwardBody Then
        Dim forwardMail As Object
        Set forwardMail = olMail.Forward
        forwardMail.To = forwardTo
        forwardMail.body = olMail.body
        forwardMail.send
        xlWs.Cells(ruleIndex + 1, 8).value = Now
    End If
    Exit Sub
ErrorHandler:
    DebugLog "Error in ProcessEmailForwarding: " & Err.description
End Sub

'===============================================
' Background Email Checker
'===============================================
Public Sub CheckOutlookBackground()
    On Error Resume Next
    Dim olApp As Object, olSession As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olSession = olApp.GetNamespace("MAPI")
    olSession.SendAndReceive False
    olSession.Logoff
End Sub

'===============================================
' Outlook Safety / Status
'===============================================
Public Function OutlookHung() As Boolean
    Dim hWnd As LongPtr
    hWnd = FindWindow("rctrl_renwnd32", vbNullString)
    If hWnd <> 0 Then
        OutlookHung = (IsHungAppWindow(hWnd) <> 0)
    Else
        OutlookHung = True
    End If
End Function

Public Function OutlookResponsive() As Boolean
    On Error Resume Next
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    OutlookResponsive = Not olApp Is Nothing
End Function



'===============================================
' Setup Functions
'===============================================
Public Sub SetupOutlookSheetOLD()
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = "Outlook"
    
    xlSheet.Cells(1, 1).value = "Rule Name"
    xlSheet.Cells(1, 2).value = "Enabled"
    xlSheet.Cells(1, 3).value = "Subject Contains"
    xlSheet.Cells(1, 4).value = "Action"
    xlSheet.Cells(1, 5).value = "Forward To"
    xlSheet.Cells(1, 6).value = "Sender Contains"
    xlSheet.Cells(1, 7).value = "Body Contains"
    xlSheet.Cells(1, 8).value = "Last Executed"
    
    xlBook.SaveAs ThisWorkbook.path & "\OutlookRules.xlsx"
    xlBook.Close
    xlApp.Quit
    DebugLog "Outlook rules sheet created"
    Exit Sub
ErrorHandler:
    DebugLog "Error in SetupOutlookSheet: " & Err.description
End Sub

Public Sub SetupOutlookSheet()
    On Error GoTo ErrorHandler
    
    Dim xlSheet As Worksheet
    Dim sheetExists As Boolean
    Dim expectedHeaders As Variant
    Dim headersValid As Boolean
    Dim i As Long
    
    ' Define expected headers
    expectedHeaders = Array("Rule Name", "Enabled", "Subject Contains", "Action", "Forward To", _
                           "Sender Contains", "Body Contains", "Last Executed")
    
    ' Check if "Outlook" sheet exists in ThisWorkbook
    sheetExists = False
    For Each xlSheet In ThisWorkbook.Worksheets
        If xlSheet.Name = "Outlook" Then
            sheetExists = True
            Set xlSheet = xlSheet
            Exit For
        End If
    Next xlSheet
    
    ' Validate headers if sheet exists
    If sheetExists Then
        headersValid = True
        For i = 1 To 8
            If xlSheet.Cells(1, i).value <> expectedHeaders(i - 1) Then
                headersValid = False
                DebugLog "Invalid header in column " & i & ": Expected '" & expectedHeaders(i - 1) & "', Found '" & xlSheet.Cells(1, i).value & "'"
                Exit For
            End If
        Next i
        
        If headersValid Then
            DebugLog "Outlook sheet exists with valid headers"
        Else
            ' Option 1: Fix headers in existing sheet
            DebugLog "Fixing invalid headers in existing Outlook sheet"
            For i = 1 To 8
                xlSheet.Cells(1, i).value = expectedHeaders(i - 1)
            Next i
            ' Option 2: (Alternative) Delete and recreate sheet - uncomment if preferred
            ' Application.DisplayAlerts = False
            ' xlSheet.Delete
            ' Set xlSheet = ThisWorkbook.Worksheets.Add
            ' xlSheet.Name = "Outlook"
            ' For i = 1 To 8
            '     xlSheet.Cells(1, i).Value = expectedHeaders(i - 1)
            ' Next i
        End If
    Else
        ' Create new Outlook sheet
        Set xlSheet = ThisWorkbook.Worksheets.Add
        xlSheet.Name = "Outlook"
        For i = 1 To 8
            xlSheet.Cells(1, i).value = expectedHeaders(i - 1)
        Next i
        DebugLog "Created new Outlook sheet with headers"
    End If
    
    ' Add sample TEST rule if no rules exist
    If xlSheet.Cells(2, 1).value = "" Then
        With xlSheet
            .Cells(2, 1).value = "TestRule"
            .Cells(2, 2).value = "TRUE"
            .Cells(2, 3).value = "TEST"
            .Cells(2, 4).value = "Forward"
            .Cells(2, 5).value = "you@example.com" ' Replace with your email
            .Cells(2, 6).value = ""
            .Cells(2, 7).value = ""
            .Cells(2, 8).value = ""
        End With
        DebugLog "Added sample TEST rule to Outlook sheet"
    End If
    
    ThisWorkbook.Save
    DebugLog "Outlook rules sheet created/updated in ThisWorkbook"
    Exit Sub
ErrorHandler:
    DebugLog "Error in SetupOutlookSheet: " & Err.description
End Sub
'===============================================
' --- Test Functions for OutlookExtra ---
'===============================================

Public Sub TestImmediateMemoryMode()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olFolder As Object, olItems As Object, olItem As Object, i As Long
    Set olApp = CreateObject("Outlook.Application")
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

Public Sub TestImmediateTempMode()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olTempFolder As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olTempFolder = olApp.GetNamespace("MAPI").GetDefaultFolder(3) ' Deleted Items
    DebugLog "Temp mode: Found " & olTempFolder.items.count & " items in Deleted Items"
    
    Set olTempFolder = Nothing: Set olApp = Nothing
    DebugLog "TestImmediateTempMode executed successfully"
    Exit Sub
ErrorHandler:
    DebugLog "Error in TestImmediateTempMode: " & Err.description
End Sub

Public Sub TestBackgroundMode()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olSession As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olSession = olApp.GetNamespace("MAPI")
    olSession.Logon
    olSession.SendAndReceive False
    DebugLog "Background mode: Initiated Outlook sync"
    olSession.Logoff
    
    Set olSession = Nothing: Set olApp = Nothing
    DebugLog "TestBackgroundMode executed successfully"
    Exit Sub
ErrorHandler:
    DebugLog "Error in TestBackgroundMode: " & Err.description
End Sub

Public Sub TestOriginalAttachmentHandler()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object, olFolder As Object, olItems As Object
    Dim olItem As Object, olAttachment As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olFolder = olApp.GetNamespace("MAPI").GetDefaultFolder(6) ' Inbox
    Set olItems = olFolder.items
    olItems.Sort "[ReceivedTime]", True
    
    If olItems.count > 0 Then
        Set olItem = olItems.item(1)
        For Each olAttachment In olItem.Attachments
            DebugLog "Attachment handler: Found attachment " & olAttachment.fileName & " in email from " & olItem.SenderName
        Next
    End If
    
    Set olItems = Nothing: Set olFolder = Nothing: Set olApp = Nothing
    DebugLog "TestOriginalAttachmentHandler executed successfully"
    Exit Sub
ErrorHandler:
    DebugLog "Error in TestOriginalAttachmentHandler: " & Err.description
End Sub



' Outlook Hung Detection and Recovery System
' Add this to your OutlookExtra module

Public Sub CheckOutlookHealthAndRecoverOLD()
    On Error GoTo ErrorHandler
    
    Debug.Print "Checking Outlook health at " & format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Check if Outlook is responsive
    If Not OutlookResponsive() Then
        Debug.Print "Outlook not responsive - checking if hung..."
        
        If OutlookHung() Then
            Debug.Print "Outlook is hung - initiating recovery process"
            Call RecoverHungOutlook
        Else
            Debug.Print "Outlook not running - starting fresh instance"
            Call StartOutlookSafely
        End If
    Else
        Debug.Print "Outlook is responsive"
        
        ' Optional: Test COM access to make sure it's really working
        If Not TestOutlookCOMAccess() Then
            Debug.Print "Outlook running but COM access failed - recovering"
            Call RecoverHungOutlook
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "Error in CheckOutlookHealthAndRecover: " & Err.description
End Sub



Private Sub ForceTerminateOutlookOLD()
    On Error Resume Next
    
    ' Get all Outlook processes and terminate them
    Dim wmi As Object, processes As Object, proc As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='OUTLOOK.EXE'")
    
    For Each proc In processes
        Debug.Print "Force terminating Outlook PID: " & proc.ProcessId
        proc.Terminate
    Next
    
    ' Backup method using taskkill
    shell "taskkill /F /IM outlook.exe", vbHide
    
    On Error GoTo 0
End Sub



Public Sub TestBasicOutlookConnection()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Debug.Print "Outlook connected: " & Now
    
    Dim olNS As Object, olInbox As Object
    Set olNS = olApp.GetNamespace("MAPI")
    Set olInbox = olNS.GetDefaultFolder(6)
    Debug.Print "Unread emails: " & olInbox.UnReadItemCount
    
    Exit Sub
ErrorHandler:
    Debug.Print "Connection failed: " & Err.description
End Sub

Public Sub TestDirectOutlookRules()
    On Error GoTo ErrorHandler
    
    Debug.Print "Starting direct rules test..."
    Call OutlookExtra.RunOutlookRules
    Debug.Print "Rules completed"
    
    Exit Sub
ErrorHandler:
    Debug.Print "Rules failed: " & Err.description
End Sub

' Outlook System Call Failed - Diagnostic & Fix Module
' Place this in a new VBA module to diagnose and fix the connection issue

Option Explicit

Public Sub DiagnoseOutlookConnection()
    Debug.Print "=== OUTLOOK DIAGNOSTIC START ==="
    Debug.Print "Time: " & Now
    
    ' Step 1: Check if Outlook process is running
    If IsOutlookProcessRunning() Then
        Debug.Print "? Outlook process (OUTLOOK.EXE) is running"
    Else
        Debug.Print "? Outlook process is NOT running"
        Debug.Print "Attempting to start Outlook..."
        Call StartOutlookSafely
        Exit Sub
    End If
    
    ' Step 2: Try to get existing Outlook instance
    On Error Resume Next
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number = 0 Then
        Debug.Print "? Successfully connected to existing Outlook instance"
        Err.Clear
    Else
        Debug.Print "? Failed to connect to existing Outlook: " & Err.description
        Err.Clear
        
        ' Step 3: Try creating new instance
        Set olApp = CreateObject("Outlook.Application")
        If Err.Number = 0 Then
            Debug.Print "? Successfully created new Outlook instance"
            Err.Clear
        Else
            Debug.Print "? Failed to create Outlook instance: " & Err.description
            Debug.Print "This suggests Outlook installation or COM registration issues"
            Err.Clear
            Exit Sub
        End If
    End If
    On Error GoTo 0
    
    ' Step 4: Test MAPI namespace access
    On Error Resume Next
    Dim olNS As Object
    Set olNS = olApp.GetNamespace("MAPI")
    If Err.Number = 0 Then
        Debug.Print "? Successfully accessed MAPI namespace"
        Err.Clear
        
        ' Step 5: Test inbox access
        Dim olInbox As Object
        Set olInbox = olNS.GetDefaultFolder(6) ' Inbox = 6
        If Err.Number = 0 Then
            Debug.Print "? Successfully accessed Inbox folder"
            Debug.Print "Unread email count: " & olInbox.UnReadItemCount
            Debug.Print "Total items in inbox: " & olInbox.items.count
        Else
            Debug.Print "? Failed to access Inbox: " & Err.description
            Debug.Print "This suggests Outlook profile or mailbox issues"
        End If
    Else
        Debug.Print "? Failed to access MAPI namespace: " & Err.description
        Debug.Print "This suggests Outlook is not properly initialized"
    End If
    On Error GoTo 0
    
    Debug.Print "=== OUTLOOK DIAGNOSTIC END ==="
End Sub

Private Function IsOutlookProcessRunningOther() As Boolean
    On Error Resume Next
    Dim objWMI As Object, colProcesses As Object, objProcess As Object
    
    Set objWMI = GetObject("winmgmts:")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='OUTLOOK.EXE'")
    
    IsOutlookProcessRunning = (colProcesses.count > 0)
    
    If colProcesses.count > 0 Then
        For Each objProcess In colProcesses
            Debug.Print "Found Outlook process PID: " & objProcess.ProcessId
        Next
    End If
    
    On Error GoTo 0
End Function


Public Sub FixOutlookCOMRegistration()
    ' Run this if Outlook won't create COM objects
    ' This requires running Excel as Administrator
    
    Debug.Print "Attempting to re-register Outlook COM components..."
    
    On Error Resume Next
    Dim WShell As Object
    Set WShell = CreateObject("WScript.Shell")
    
    ' Re-register Outlook
    WShell.Run "regsvr32 /s ""C:\Program Files\Microsoft Office\root\Office16\OUTLCTL.DLL""", 0, True
    WShell.Run "regsvr32 /s ""C:\Program Files\Microsoft Office\root\Office16\MSOUTL.OLB""", 0, True
    
    Debug.Print "COM registration attempted. Restart Excel and try again."
    On Error GoTo 0
End Sub

Public Sub TestSimpleOutlookAccess()
    ' Simplified test without error handling to see exact error
    Debug.Print "Testing direct Outlook access..."
    
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Debug.Print "Outlook object created"
    
    Dim olNS As Object
    Set olNS = olApp.GetNamespace("MAPI")
    Debug.Print "MAPI namespace obtained"
    
    Dim olInbox As Object
    Set olInbox = olNS.GetDefaultFolder(6)
    Debug.Print "Inbox accessed successfully"
    Debug.Print "Unread count: " & olInbox.UnReadItemCount
End Sub

Public Sub CreateFixedOutlookChecker()
    ' Create a working version that handles the connection issues
    
    Dim fso As Object, ts As Object
    Dim vbsPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    vbsPath = "C:\SmartTraffic\FixedOutlookChecker.vbs"
    
    If Not fso.FolderExists("C:\SmartTraffic") Then
        fso.CreateFolder "C:\SmartTraffic"
    End If
    
    Set ts = fso.CreateTextFile(vbsPath, True)
    ts.WriteLine "Option Explicit"
    ts.WriteLine ""
    ts.WriteLine "Dim objExcel, objWorkbook, objOutlook, objNamespace, objInbox"
    ts.WriteLine "Dim stopFile, logFile, fso"
    ts.WriteLine ""
    ts.WriteLine "stopFile = ""C:\SmartTraffic\OutlookChecker.stop"""
    ts.WriteLine "logFile = ""C:\SmartTraffic\OutlookChecker.log"""
    ts.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    ts.WriteLine ""
    ts.WriteLine "Sub WriteLog(message)"
    ts.WriteLine "    Dim ts"
    ts.WriteLine "    Set ts = fso.OpenTextFile(logFile, 8, True)"
    ts.WriteLine "    ts.WriteLine Now & "" - "" & message"
    ts.WriteLine "    ts.Close"
    ts.WriteLine "End Sub"
    ts.WriteLine ""
    ts.WriteLine "WriteLog ""Starting Outlook checker..."""
    ts.WriteLine ""
    ts.WriteLine "' Try to connect to Outlook directly"
    ts.WriteLine "On Error Resume Next"
    ts.WriteLine "Set objOutlook = GetObject(, ""Outlook.Application"")"
    ts.WriteLine "If objOutlook Is Nothing Then"
    ts.WriteLine "    WriteLog ""Outlook not running, attempting to start..."""
    ts.WriteLine "    Set objOutlook = CreateObject(""Outlook.Application"")"
    ts.WriteLine "    WScript.Sleep 5000"
    ts.WriteLine "End If"
    ts.WriteLine ""
    ts.WriteLine "If objOutlook Is Nothing Then"
    ts.WriteLine "    WriteLog ""Failed to connect to Outlook"""
    ts.WriteLine "    WScript.Quit"
    ts.WriteLine "End If"
    ts.WriteLine ""
    ts.WriteLine "Set objNamespace = objOutlook.GetNamespace(""MAPI"")"
    ts.WriteLine "Set objInbox = objNamespace.GetDefaultFolder(6)"
    ts.WriteLine ""
    ts.WriteLine "WriteLog ""Connected to Outlook successfully"""
    ts.WriteLine ""
    ts.WriteLine "Do"
    ts.WriteLine "    If fso.FileExists(stopFile) Then"
    ts.WriteLine "        WriteLog ""Stop file detected, exiting..."""
    ts.WriteLine "        WScript.Quit"
    ts.WriteLine "    End If"
    ts.WriteLine ""
    ts.WriteLine "    On Error Resume Next"
    ts.WriteLine "    Dim unreadCount"
    ts.WriteLine "    unreadCount = objInbox.UnReadItemCount"
    ts.WriteLine "    WriteLog ""Unread emails: "" & unreadCount"
    ts.WriteLine ""
    ts.WriteLine "    ' Trigger send/receive"
    ts.WriteLine "    objNamespace.SendAndReceive False"
    ts.WriteLine ""
    ts.WriteLine "    WScript.Sleep 60000 ' Wait 1 minute"
    ts.WriteLine "Loop"
    ts.Close
    
    Debug.Print "Fixed VBS checker created at: " & vbsPath
    Debug.Print "To test: cscript """ & vbsPath & """"
End Sub

Public Sub ManualOutlookTest()
    ' Run this step by step to isolate the exact failure point
    Debug.Print "=== MANUAL OUTLOOK TEST ==="
    
    ' Step 1: Try to start Outlook if not running
    Debug.Print "Step 1: Checking if Outlook is running..."
    If Not IsOutlookProcessRunning() Then
        Debug.Print "Outlook not running. Starting manually..."
        shell "outlook.exe", vbNormalFocus
        
        ' Wait 10 seconds for startup
        Dim i As Integer
        For i = 1 To 10
            DoEvents
            Sleep 1000
            Debug.Print "Waiting for Outlook startup... " & i & "/10"
        Next i
    Else
        Debug.Print "Outlook process found"
    End If
    
    ' Step 2: Test GetObject (existing instance)
    Debug.Print "Step 2: Testing GetObject..."
    On Error Resume Next
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Debug.Print "GetObject failed: " & Err.description & " (Error " & Err.Number & ")"
        Err.Clear
        
        ' Step 3: Test CreateObject (new instance)
        Debug.Print "Step 3: Testing CreateObject..."
        Set olApp = CreateObject("Outlook.Application")
        If Err.Number <> 0 Then
            Debug.Print "CreateObject failed: " & Err.description & " (Error " & Err.Number & ")"
            Debug.Print "CRITICAL: Outlook COM interface is broken"
            Exit Sub
        Else
            Debug.Print "CreateObject succeeded"
        End If
    Else
        Debug.Print "GetObject succeeded"
    End If
    On Error GoTo 0
    
    ' Step 4: Test MAPI access
    Debug.Print "Step 4: Testing MAPI namespace..."
    On Error Resume Next
    Dim olNS As Object
    Set olNS = olApp.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        Debug.Print "MAPI access failed: " & Err.description
        Exit Sub
    Else
        Debug.Print "MAPI access succeeded"
    End If
    On Error GoTo 0
    
    ' Step 5: Test Inbox access
    Debug.Print "Step 5: Testing Inbox access..."
    On Error Resume Next
    Dim olInbox As Object
    Set olInbox = olNS.GetDefaultFolder(6)
    If Err.Number <> 0 Then
        Debug.Print "Inbox access failed: " & Err.description
        Exit Sub
    Else
        Debug.Print "Inbox access succeeded"
        Debug.Print "Unread emails: " & olInbox.UnReadItemCount
    End If
    On Error GoTo 0
    
    Debug.Print "=== ALL TESTS PASSED ==="
End Sub


' Outlook Hung Detection and Recovery System
' Add this to your OutlookExtra module

Public Sub CheckOutlookHealthAndRecover()
    On Error GoTo ErrorHandler
    
    Debug.Print "Checking Outlook health at " & format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Check if Outlook is responsive
    If Not OutlookResponsive() Then
        Debug.Print "Outlook not responsive - checking if hung..."
        
        If OutlookHung() Then
            Debug.Print "Outlook is hung - initiating recovery process"
            Call RecoverHungOutlook
        Else
            Debug.Print "Outlook not running - starting fresh instance"
            Call StartOutlookSafely
        End If
    Else
        Debug.Print "Outlook is responsive"
        
        ' Optional: Test COM access to make sure it's really working
        If Not TestOutlookCOMAccess() Then
            Debug.Print "Outlook running but COM access failed - recovering"
            Call RecoverHungOutlook
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "Error in CheckOutlookHealthAndRecover: " & Err.description
End Sub


Private Function GracefulOutlookShutdown() As Boolean
    On Error Resume Next
    
    ' Try to quit Outlook through COM interface
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    
    If Not olApp Is Nothing Then
        olApp.Quit
        Sleep 3000 ' Wait 3 seconds for graceful shutdown
        
        ' Check if it actually closed
        Set olApp = Nothing
        Set olApp = GetObject(, "Outlook.Application")
        If olApp Is Nothing Then
            GracefulOutlookShutdown = True
        Else
            GracefulOutlookShutdown = False
        End If
    Else
        GracefulOutlookShutdown = False
    End If
    
    On Error GoTo 0
End Function

Private Sub ForceTerminateOutlook()
    On Error Resume Next
    
    ' Get all Outlook processes and terminate them
    Dim wmi As Object, processes As Object, proc As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='OUTLOOK.EXE'")
    
    For Each proc In processes
        Debug.Print "Force terminating Outlook PID: " & proc.ProcessId
        proc.Terminate
    Next
    
    ' Backup method using taskkill
    shell "taskkill /F /IM outlook.exe", vbHide
    
    On Error GoTo 0
End Sub


Private Function FindOutlookExecutable() As String
    ' Try common Outlook installation paths
    Dim paths As Variant
    Dim i As Integer
    
    paths = Array( _
        "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE", _
        "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE", _
        "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE", _
        "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE", _
        "C:\Program Files\Microsoft Office\root\Office15\OUTLOOK.EXE", _
        "C:\Program Files (x86)\Microsoft Office\root\Office15\OUTLOOK.EXE", _
        "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE", _
        "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" _
    )
    
    For i = 0 To UBound(paths)
        If Dir(paths(i)) <> "" Then
            FindOutlookExecutable = paths(i)
            Exit Function
        End If
    Next i
    
    ' If not found in standard locations, try registry lookup
    FindOutlookExecutable = GetOutlookPathFromRegistry()
End Function

Private Function GetOutlookPathFromRegistry() As String
    On Error Resume Next
    
    Dim WShell As Object
    Set WShell = CreateObject("WScript.Shell")
    
    ' Try to get path from registry
    GetOutlookPathFromRegistry = WShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE\")
    
    If GetOutlookPathFromRegistry = "" Then
        ' Fallback - let Windows find it
        GetOutlookPathFromRegistry = "outlook.exe"
    End If
    
    On Error GoTo 0
End Function

Private Function TestOutlookCOMAccess() As Boolean
    On Error Resume Next
    
    Dim olApp As Object, olNS As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    
    ' Try to access inbox
    Dim olInbox As Object
    Set olInbox = olNS.GetDefaultFolder(6)
    
    If Err.Number = 0 And Not olInbox Is Nothing Then
        TestOutlookCOMAccess = True
    Else
        TestOutlookCOMAccess = False
    End If
    
    On Error GoTo 0
End Function


' Update your existing functions to use health checking
Public Sub InitializeOutlookCheckingWithRecovery()
    On Error GoTo ErrorHandler
    
    ' First check health and recover if needed
    Call CheckOutlookHealthAndRecover
    
    ' Then run your normal initialization
    If OutlookResponsive() Then
        Call InitializeOutlookChecking
    Else
        Debug.Print "Could not initialize Outlook - manual intervention may be required"
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "Error in InitializeOutlookCheckingWithRecovery: " & Err.description
End Sub

Private Function IsOutlookProcessRunning() As Boolean
    On Error Resume Next
    Dim wmi As Object, processes As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='OUTLOOK.EXE'")
    IsOutlookProcessRunning = (processes.count > 0)
    On Error GoTo 0
End Function



' --- Search and retrieve latest emails ---
Public Function SearchOutlookEmails(ByVal count As Long) As String
    On Error GoTo ErrorHandler
    Dim olApp As Object, olNS As Object, olFolder As Object
    Dim olMail As Object
    Dim i As Long
    Dim result As String
    
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(6) ' Inbox

    ' Sort items descending by received time
    Dim olItems As Object
    Set olItems = olFolder.items
    olItems.Sort "[ReceivedTime]", True
    
    result = ""
    For i = 1 To count
        If i > olItems.count Then Exit For
        Set olMail = olItems(i)
        If olMail.Class = 43 Then ' MailItem
            result = result & format(olMail.ReceivedTime, "yyyy-mm-dd hh:mm:ss") & " - " & HTMLEncode(olMail.Subject) & "<br>"
        End If
    Next i

    SearchOutlookEmails = result
    Exit Function

ErrorHandler:
    SearchOutlookEmails = "Unable to retrieve emails: " & Err.description
End Function

Public Function IsBackgroundModeEnabled() As Boolean
    IsBackgroundModeEnabled = m_backgroundEnabled
End Function

' --- UPDATED VBS Script Creation with Health Checking ---
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
    
    ' UPDATED VBS script content - doesn't keep Excel open
    vbsCode = "'==================================================" & vbCrLf
    vbsCode = vbsCode & "' OutlookChecker.vbs - Safe Outlook checker with recovery" & vbCrLf
    vbsCode = vbsCode & "'==================================================" & vbCrLf
    vbsCode = vbsCode & "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim fso, logFile, stopFile, workbookPath" & vbCrLf
    vbsCode = vbsCode & "Dim checkInterval, consecutiveFailures" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "checkInterval = 90000 ' 90 seconds" & vbCrLf
    vbsCode = vbsCode & "consecutiveFailures = 0" & vbCrLf
    vbsCode = vbsCode & "workbookPath = """ & wbPath & """" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "logFile = ""C:\SmartTraffic\OutlookChecker.log""" & vbCrLf
    vbsCode = vbsCode & "stopFile = ""C:\SmartTraffic\OutlookChecker.stop""" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "Sub WriteLog(message)" & vbCrLf
    vbsCode = vbsCode & "    On Error Resume Next" & vbCrLf
    vbsCode = vbsCode & "    Dim ts" & vbCrLf
    vbsCode = vbsCode & "    Set ts = fso.OpenTextFile(logFile, 8, True)" & vbCrLf
    vbsCode = vbsCode & "    ts.WriteLine Now & "" - "" & message" & vbCrLf
    vbsCode = vbsCode & "    ts.Close" & vbCrLf
    vbsCode = vbsCode & "End Sub" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "Function IsOutlookRunning()" & vbCrLf
    vbsCode = vbsCode & "    On Error Resume Next" & vbCrLf
    vbsCode = vbsCode & "    Dim wmi, processes" & vbCrLf
    vbsCode = vbsCode & "    Set wmi = GetObject(""winmgmts:"")" & vbCrLf
    vbsCode = vbsCode & "    Set processes = wmi.ExecQuery(""SELECT * FROM Win32_Process WHERE Name='OUTLOOK.EXE'"")" & vbCrLf
    vbsCode = vbsCode & "    IsOutlookRunning = (processes.Count > 0)" & vbCrLf
    vbsCode = vbsCode & "End Function" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "Function StartOutlook()" & vbCrLf
    vbsCode = vbsCode & "    On Error Resume Next" & vbCrLf
    vbsCode = vbsCode & "    Dim shell" & vbCrLf
    vbsCode = vbsCode & "    Set shell = CreateObject(""WScript.Shell"")" & vbCrLf
    vbsCode = vbsCode & "    shell.Run ""outlook.exe"", 1, False" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 30000 ' Wait 30 seconds" & vbCrLf
    vbsCode = vbsCode & "    StartOutlook = IsOutlookRunning()" & vbCrLf
    vbsCode = vbsCode & "End Function" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "Sub CallVBAFunctions()" & vbCrLf
    vbsCode = vbsCode & "    On Error Resume Next" & vbCrLf
    vbsCode = vbsCode & "    Dim xlApp, xlWorkbook" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    WriteLog ""Opening Excel to call VBA functions...""" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    Set xlApp = CreateObject(""Excel.Application"")" & vbCrLf
    vbsCode = vbsCode & "    If Err.Number <> 0 Then" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Failed to create Excel: "" & Err.Description" & vbCrLf
    vbsCode = vbsCode & "        Exit Sub" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    xlApp.Visible = False" & vbCrLf
    vbsCode = vbsCode & "    xlApp.DisplayAlerts = False" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    Set xlWorkbook = xlApp.Workbooks.Open(workbookPath)" & vbCrLf
    vbsCode = vbsCode & "    If Err.Number <> 0 Then" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Failed to open workbook: "" & Err.Description" & vbCrLf
    vbsCode = vbsCode & "        xlApp.Quit" & vbCrLf
    vbsCode = vbsCode & "        Exit Sub" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    ' Call the health check function from OutlookExtra" & vbCrLf
    vbsCode = vbsCode & "    WriteLog ""Calling OutlookExtra.EnhancedOutlookHealthCheck""" & vbCrLf
    vbsCode = vbsCode & "    xlApp.Run ""OutlookExtra.EnhancedOutlookHealthCheck""" & vbCrLf
    vbsCode = vbsCode & "    If Err.Number <> 0 Then" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Health check failed: "" & Err.Description" & vbCrLf
    vbsCode = vbsCode & "        Err.Clear" & vbCrLf
    vbsCode = vbsCode & "    Else" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Health check completed successfully""" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    ' CRITICAL: Close Excel immediately" & vbCrLf
    vbsCode = vbsCode & "    xlWorkbook.Close False" & vbCrLf
    vbsCode = vbsCode & "    xlApp.Quit" & vbCrLf
    vbsCode = vbsCode & "    Set xlWorkbook = Nothing" & vbCrLf
    vbsCode = vbsCode & "    Set xlApp = Nothing" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    WriteLog ""VBA functions completed, Excel closed""" & vbCrLf
    vbsCode = vbsCode & "End Sub" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "' Main execution loop" & vbCrLf
    vbsCode = vbsCode & "WriteLog ""OutlookChecker starting...""" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "    If fso.FileExists(stopFile) Then" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Stop file detected - exiting""" & vbCrLf
    vbsCode = vbsCode & "        Exit Do" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    ' Ensure Outlook is running" & vbCrLf
    vbsCode = vbsCode & "    If Not IsOutlookRunning() Then" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Outlook not running - starting it""" & vbCrLf
    vbsCode = vbsCode & "        If Not StartOutlook() Then" & vbCrLf
    vbsCode = vbsCode & "            WriteLog ""Failed to start Outlook""" & vbCrLf
    vbsCode = vbsCode & "            consecutiveFailures = consecutiveFailures + 1" & vbCrLf
    vbsCode = vbsCode & "        End If" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    ' Call VBA functions if Outlook is running" & vbCrLf
    vbsCode = vbsCode & "    If IsOutlookRunning() Then" & vbCrLf
    vbsCode = vbsCode & "        CallVBAFunctions" & vbCrLf
    vbsCode = vbsCode & "        consecutiveFailures = 0" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    ' Safety exit" & vbCrLf
    vbsCode = vbsCode & "    If consecutiveFailures > 5 Then" & vbCrLf
    vbsCode = vbsCode & "        WriteLog ""Too many failures - exiting for safety""" & vbCrLf
    vbsCode = vbsCode & "        Exit Do" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    vbsCode = vbsCode & "    " & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep checkInterval" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "" & vbCrLf
    vbsCode = vbsCode & "WriteLog ""OutlookChecker exiting""" & vbCrLf
    
    ' Write VBS file
    Set ts = fso.CreateTextFile(vbsPath, True)
    ts.Write vbsCode
    ts.Close
    
    DebugLog "Updated OutlookChecker.vbs created at " & vbsPath
    Exit Sub
ErrorHandler:
    DebugLog "Error creating VBS script: " & Err.description
End Sub

' --- UPDATED VBS Management Functions ---
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

Public Sub StartOutlookVBSChecker()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure VBS exists
    If Dir("C:\SmartTraffic\OutlookChecker.vbs") = "" Then
        Call CreateOutlookVBSScript
    End If
    
    ' Delete stop flag if it exists
    If fso.FileExists("C:\SmartTraffic\OutlookChecker.stop") Then
        fso.DeleteFile "C:\SmartTraffic\OutlookChecker.stop", True
    End If
    
    ' Check if already running
    If IsVBSCheckerRunning() Then
        DebugLog "VBS Outlook checker already running"
        Exit Sub
    End If
    
    ' Start the VBS script
    shell "cscript.exe ""C:\SmartTraffic\OutlookChecker.vbs"" //nologo", vbHide
    DebugLog "Started VBS Outlook checker"
    Exit Sub
ErrorHandler:
    DebugLog "Error in StartOutlookVBSChecker: " & Err.description
End Sub

Public Sub StopOutlookVBSChecker()
    On Error Resume Next
    ' Create stop file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile "C:\SmartTraffic\OutlookChecker.stop", True
    
    ' Also kill any running processes
    shell "taskkill /F /FI ""IMAGENAME eq cscript.exe"" /FI ""COMMANDLINE eq *OutlookChecker.vbs*""", vbHide
    DebugLog "Stopped VBS Outlook checker"
    On Error GoTo 0
End Sub

' --- Background Mode Control ---
Public Sub EnableBackgroundMode()
    m_backgroundEnabled = True
    Call StartOutlookVBSChecker
    DebugLog "Outlook background mode enabled"
End Sub

Public Sub DisableBackgroundMode()
    m_backgroundEnabled = False
    Call StopOutlookVBSChecker
    DebugLog "Outlook background mode disabled"
End Sub

' --- Initialize Outlook Checking ---
Public Sub InitializeOutlookChecking()
    On Error GoTo ErrorHandler
    
    ' Ensure SmartTraffic folder exists
    If Dir("C:\SmartTraffic", vbDirectory) = "" Then
        MkDir "C:\SmartTraffic"
    End If
    
    ' Create/update VBS script
    Call CreateOutlookVBSScript
    
    ' Start VBS checker if not already running
    If Not IsVBSCheckerRunning() Then
        Call StartOutlookVBSChecker
    End If
    
    DebugLog "InitializeOutlookChecking completed"
    Exit Sub
    
ErrorHandler:
    DebugLog "Error in InitializeOutlookChecking: " & Err.description
End Sub

Public Sub StopOutlookChecking()
    Call StopOutlookVBSChecker
    DebugLog "Outlook checking stopped"
End Sub

' --- HTML Encode Helper ---
Private Function HTMLEncode(ByVal Text As String) As String
    HTMLEncode = Replace(Replace(Replace(Text, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

' --- Keep your other existing functions unchanged ---
' (TestImmediateMemoryMode, TestImmediateTempMode, TestBackgroundMode, etc.)

'==================================================
' ADDITIONS TO OUTLOOKEXTRA MODULE
'==================================================

' Add this function to your OutlookExtra module:
Public Sub EnhancedOutlookHealthCheck()
    On Error GoTo ErrorHandler
    
    DebugLog "=== Enhanced Outlook Health Check Starting ==="
    
    ' Step 1: Check if Outlook is responsive
    If Not OutlookResponsive() Then
        DebugLog "Outlook not responsive - attempting recovery"
        
        If OutlookHung() Then
            DebugLog "Outlook detected as hung - restarting"
            Call RecoverHungOutlook
        Else
            DebugLog "Outlook not running - starting"
            Call StartOutlookSafely
        End If
        
        ' Wait and check again
        Sleep 5000
        If Not OutlookResponsive() Then
            DebugLog "Outlook still not responsive after recovery attempt"
            Exit Sub
        End If
    End If
    
    ' Step 2: Run normal operations if Outlook is responsive
    DebugLog "Outlook is responsive - running normal checks"
    
    On Error Resume Next
    
    ' Run rules
    Call RunOutlookRules
    If Err.Number <> 0 Then
        DebugLog "RunOutlookRules failed: " & Err.description
        Err.Clear
    End If
    
    ' Background check
    Call CheckOutlookBackground
    If Err.Number <> 0 Then
        DebugLog "CheckOutlookBackground failed: " & Err.description
        Err.Clear
    End If
    
    ' Universal email checker
    Call UniversalEmailChecker(background:=True)
    If Err.Number <> 0 Then
        DebugLog "UniversalEmailChecker failed: " & Err.description
        Err.Clear
    End If
    
    On Error GoTo ErrorHandler
    
    DebugLog "=== Enhanced Outlook Health Check Complete ==="
    Exit Sub
    
ErrorHandler:
    DebugLog "Error in EnhancedOutlookHealthCheck: " & Err.description
End Sub

' Add this recovery function to your OutlookExtra module:
Private Sub RecoverHungOutlook()
    On Error Resume Next
    
    DebugLog "Starting hung Outlook recovery"
    
    ' Try graceful shutdown first
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If Not olApp Is Nothing Then
        olApp.Quit
        Sleep 3000
    End If
    
    ' Force terminate if still running
    Dim wmi As Object, processes As Object, proc As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='OUTLOOK.EXE'")
    
    For Each proc In processes
        DebugLog "Force terminating Outlook PID: " & proc.ProcessId
        proc.Terminate
    Next
    
    Sleep 2000
    
    ' Start fresh instance
    shell "outlook.exe", vbNormalFocus
    Sleep 10000 ' Wait for startup
    
    DebugLog "Outlook recovery completed"
End Sub

Private Sub StartOutlookSafely()
    On Error Resume Next
    DebugLog "Starting Outlook safely"
    shell "outlook.exe", vbNormalFocus
    Sleep 10000 ' Wait for startup
    DebugLog "Outlook startup completed"
End Sub
