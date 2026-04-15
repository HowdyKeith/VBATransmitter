Public Function HandleBulbRequest(ByVal method As String, Optional ByVal body As String = "") As String
    On Error GoTo ErrHandler
    Dim urlPath As String
    
    ' For testing, we’ll just simulate extracting the path
    ' In your real server, you’d parse the HTTP request line
    urlPath = LCase$(Trim$(body)) ' replace with actual path extraction
    
    Select Case urlPath
        Case "", "/", "/index.html"
            HandleBulbRequest = GenerateLCARSAppLauncherPage()
        Case "/outlook"
            HandleBulbRequest = OutlookWWW.GenerateLCARSOutlookLandingPage()
        Case "/apps"
            HandleBulbRequest = GenerateLCARSAppLauncherPage()
        Case "/dashboard"
            HandleBulbRequest = GenerateLCARSDashboardPage()
        Case "/reports"
            HandleBulbRequest = GenerateLCARSReportsLandingPage()
        Case "/settings"
            HandleBulbRequest = GenerateLCARSSettingsLandingPage()
        Case "/data"
            HandleBulbRequest = GenerateLCARSDataLandingPage()
        Case Else
            HandleBulbRequest = "<html><body><h2>404 - Page Not Found</h2></body></html>"
    End Select
    
    Exit Function
ErrHandler:
    Debug.Print "Error in HandleBulbRequest: " & Err.description
    HandleBulbRequest = "<html><body><h2>500 - Internal Server Error</h2></body></html>"
End Function

