Option Explicit

Sub TestWICNoPopup()
    Dim factory As Object
    On Error GoTo ErrHandler
    
    ' Create WIC Imaging Factory (built-in Windows COM object)
    Set factory = CreateObject("WICImagingFactory")
    
    If Not factory Is Nothing Then
        MsgBox "WIC Imaging Factory created successfully!" & vbCrLf & _
               "No ActiveX warning appeared.", vbInformation
    Else
        MsgBox "Failed to create WIC Imaging Factory.", vbCritical
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.description, vbCritical
End Sub

