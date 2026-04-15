    ' Windows API for downloading files
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
         ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Sub DownloadNextdoorImages()
    Dim imageURLs As Variant
    Dim outputFolder As String
    Dim i As Integer
    Dim url As String
    Dim fileName As String
    Dim result As Long
    
    ' List of image URLs (add more as you find them)
    imageURLs = Array( _
        "https://us1-photo.nextdoor.com/post_photos/9d/29/9d299baf83101f01ef363706d68d1eb0.jpeg" _
    )
    
    ' Output folder (create it if it doesn't exist)
    outputFolder = "C:\Downloads\NextdoorImages\" ' Change to your desired path
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    
    
    For i = LBound(imageURLs) To UBound(imageURLs)
        url = imageURLs(i)
        fileName = outputFolder & "cat_image_" & (i + 1) & ".jpeg" ' Auto-name files
        result = URLDownloadToFile(0, url, fileName, 0, 0)
        If result = 0 Then
            Debug.Print "Downloaded: " & fileName
        Else
            Debug.Print "Failed to download: " & url
        End If
    Next i
    
    MsgBox "Download complete! Check " & outputFolder, vbInformation
End Sub

