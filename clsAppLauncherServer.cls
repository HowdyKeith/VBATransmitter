Option Explicit

Private mPort As Long
Private mIsRunning As Boolean

' --- Initialize server with port ---
Public Sub Init(ByVal portNum As Long)
    mPort = portNum
    mIsRunning = False
End Sub

' --- Start server ---
Public Function StartServer() As Boolean
    On Error GoTo ErrHandler
    
    If mIsRunning Then
        StartServer = True
        Exit Function
    End If
    
    ' TODO: Add your server start logic here (listening on mPort)
    ' For now, simulate server start
    mIsRunning = True
    Debug.Print "App Launcher server initialized on port " & mPort
    StartServer = True
    Exit Function

ErrHandler:
    Debug.Print "Error starting App Launcher server: " & Err.description
    StartServer = False
End Function

' --- Stop server ---
Public Sub StopServer()
    If mIsRunning Then
        ' TODO: Add cleanup logic here
        mIsRunning = False
        Debug.Print "App Launcher server stopped on port " & mPort
    End If
End Sub

' --- Properties ---
Public Property Get IsRunningX() As Boolean
    isRunning = mIsRunning
End Property

Public Property Get port() As Long
    port = mPort
End Property

