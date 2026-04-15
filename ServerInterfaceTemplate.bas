Option Explicit
' ================================================================
' ServerInterfaceTemplate
' Purpose: Defines a standard interface that every server module
'          must implement so TrafficManager can run them uniformly.
' ================================================================

' --- State flag ---
Private serverRunning As Boolean

' --- GetRunning: return TRUE if server is active ---
Public Function GetRunning() As Boolean
    GetRunning = serverRunning
End Function

' --- StartServer: launch server ---
Public Sub StartServer(Optional ByVal port As Long = 0)
    On Error GoTo ErrHandler
    If serverRunning Then Exit Sub

    ' TODO: Replace with actual server startup logic
    Debug.Print "[ServerTemplate] Starting on port " & port
    
    serverRunning = True
    Exit Sub

ErrHandler:
    Debug.Print "[ServerTemplate] Start error: " & Err.description
    serverRunning = False
End Sub

' --- ProcessServer: handle tick/poll loop ---
Public Sub ProcessServer()
    On Error Resume Next
    If Not serverRunning Then Exit Sub

    ' TODO: Replace with actual server processing logic
    Debug.Print "[ServerTemplate] Processing tick..."
End Sub

' --- StopServer: shut down cleanly ---
Public Sub StopServer()
    On Error Resume Next
    If Not serverRunning Then Exit Sub

    ' TODO: Replace with actual server shutdown logic
    Debug.Print "[ServerTemplate] Stopping..."
    
    serverRunning = False
End Sub

