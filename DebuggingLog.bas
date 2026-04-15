Option Explicit

' ============================================================
' Subroutine: DebugLog v1.1
' Purpose: Writes timestamped debug messages to a daily log file
'          in C:\SmartTraffic\Logs\server_log_YYYY-MM-DD.log.
'          Creates the directory if it doesn't exist. Thread-safe
'          for non-blocking server operations.
' Parameters:
'   - msg (String): The message to log.
' Dependencies:
'   - None (uses standard VBA file operations).
' Usage:
'   - Called by MQTT module and other server components to log
'     events, errors, and status messages.
' Notes:
'   - Logs are appended to avoid overwriting.
'   - Directory is created if missing.
'   - Errors during logging are silently ignored to prevent
'     server disruption.
' ============================================================

Public Sub DebugLog(ByVal msg As String)
    On Error GoTo ErrorHandler
    Dim f As Integer
    Dim path As String
    Dim timestamp As String
    Dim folder As String
Debug.Print msg

    ' Format timestamp as YYYY-MM-DD HH:MM:SS
    timestamp = format(Now, "yyyy-mm-dd hh:nn:ss") & " [DEBUG] "
    
    ' Set log file path
    folder = "C:\SmartTraffic\Logs\"
    path = folder & "server_log_" & format(Date, "yyyy-mm-dd") & ".log"
    
    ' Create Logs directory if it doesn't exist
    If Dir(folder, vbDirectory) = "" Then
        MkDir folder
    End If
    
    ' Open file in append mode
    f = FreeFile
    Open path For Append As #f
    Print #f, timestamp & msg
    Close #f
    
    Exit Sub

ErrorHandler:
    ' Silently handle errors to avoid disrupting server
    On Error Resume Next
    Close #f
End Sub

Public Sub SafeCall(ProcName As String)
    On Error GoTo ErrHandler
    Application.Run ProcName
    Exit Sub
ErrHandler:
    DebuggingLog.DebugLog "SafeCall failed on " & ProcName & ": " & Err.description, "ERROR"
End Sub
