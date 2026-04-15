Option Explicit

' ===========================================================
' VBA Image to RGB Array Module
' Supports: PNG, JPG, BMP, GIF, TIFF
' Sources: File or Clipboard
' ===========================================================

' ----------------- GDI+ API Declarations -------------------
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As LongPtr, ByRef inputbuf As Any, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" _
    (ByVal fileName As LongPtr, ByRef image As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" _
    (ByVal image As LongPtr, ByRef rect As Any, ByVal flags As Long, ByVal PixelFormat As Long, ByRef bitmapData As Any) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal image As LongPtr, ByRef bitmapData As Any) As Long

' ----------------- Clipboard API --------------------------
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" _
    (ByVal hDrop As LongPtr, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
Private Const CF_DIB As Long = 8
Private Const CF_HDROP As Long = 15

' ----------------- Memory Helpers -------------------------
Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" ( _
    ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As Long) As Long

' ----------------- Types -------------------------
Private gdiplusToken As LongPtr
Private Type GpRect
    x As Long
    y As Long
    width As Long
    height As Long
End Type

Private Type GpBitmapData
    width As Long
    height As Long
    Stride As Long
    scan0 As LongPtr
    PixelFormat As Long
End Type

Private Const ImageLockModeRead As Long = &H1
Private Const PixelFormat32bppARGB As Long = &H26200A

' ===========================================================
' 1) GDI+ Initialization
' ===========================================================
Public Sub InitGDIPlus()
    Dim input As Long
    If gdiplusToken = 0 Then GdiplusStartup gdiplusToken, input, 0
End Sub

Public Sub ShutdownGDIPlus()
    If gdiplusToken <> 0 Then
        GdiplusShutdown gdiplusToken
        gdiplusToken = 0
    End If
End Sub

' ===========================================================
' 2) Load image from file to RGB array
' ===========================================================
Public Function LoadImageToRGB(ByVal FilePath As String) As Variant
    Dim img As LongPtr
    Dim w As Long, h As Long
    Dim rgbArray() As Long
    Dim rect As GpRect
    Dim bitmapData As GpBitmapData
    Dim row As Long, col As Long
    Dim ptrRow As LongPtr
    Dim pixel As Long
    Dim r As Long, g As Long, b As Long
    
    If gdiplusToken = 0 Then InitGDIPlus
    
    ' Load image
    If GdipLoadImageFromFile(StrPtr(FilePath), img) <> 0 Then Exit Function
    GdipGetImageWidth img, w
    GdipGetImageHeight img, h
    ReDim rgbArray(1 To h, 1 To w, 1 To 3)
    
    ' Lock full bitmap
    rect.x = 0: rect.y = 0: rect.width = w: rect.height = h
    GdipBitmapLockBits img, rect, ImageLockModeRead, PixelFormat32bppARGB, bitmapData
    
    ' Loop through each pixel
    For row = 0 To h - 1
        ptrRow = bitmapData.scan0 + row * bitmapData.Stride
        For col = 0 To w - 1
            RtlMoveMemory pixel, ByVal (ptrRow + col * 4), 4
            ' pixel = ARGB
            b = pixel And &HFF
            g = (pixel \ 256) And &HFF
            r = (pixel \ 65536) And &HFF
            rgbArray(row + 1, col + 1, 1) = r
            rgbArray(row + 1, col + 1, 2) = g
            rgbArray(row + 1, col + 1, 3) = b
        Next col
    Next row
    
    ' Unlock and dispose
    GdipBitmapUnlockBits img, bitmapData
    GdipDisposeImage img
    
    LoadImageToRGB = rgbArray
End Function

' ===========================================================
' 3) Read image from clipboard (CF_DIB) to RGB array
' ===========================================================
Public Function ClipboardImageToRGB() As Variant
    Dim hClip As LongPtr, pBits As LongPtr
    Dim ptr As LongPtr
    Dim BITMAPINFOHEADER(0 To 39) As Byte
    Dim width As Long, height As Long, bpp As Long
    Dim row As Long, col As Long, i As Long
    Dim scan0 As LongPtr
    Dim rgbArray() As Long
    Dim r As Long, g As Long, b As Long
    
    If OpenClipboard(0) = 0 Then Exit Function
    
    hClip = GetClipboardData(CF_DIB)
    If hClip = 0 Then CloseClipboard: Exit Function
    
    pBits = GlobalLock(hClip)
    If pBits = 0 Then CloseClipboard: Exit Function
    
    For i = 0 To 39
        BITMAPINFOHEADER(i) = PeekByte(pBits + i)
    Next i
    
    width = GetLongFromBytes(BITMAPINFOHEADER, 4)
    height = GetLongFromBytes(BITMAPINFOHEADER, 8)
    bpp = GetLongFromBytes(BITMAPINFOHEADER, 14)
    
    If bpp <> 24 And bpp <> 32 Then
        GlobalUnlock hClip
        CloseClipboard
        Exit Function
    End If
    
    ReDim rgbArray(1 To height, 1 To width, 1 To 3)
    scan0 = pBits + 40
    Dim rowBytes As Long
    rowBytes = ((width * bpp + 31) \ 32) * 4
    
    For row = 0 To height - 1
        For col = 0 To width - 1
            ptr = scan0 + (height - 1 - row) * rowBytes + col * (bpp \ 8)
            b = PeekByte(ptr)
            g = PeekByte(ptr + 1)
            r = PeekByte(ptr + 2)
            rgbArray(row + 1, col + 1, 1) = r
            rgbArray(row + 1, col + 1, 2) = g
            rgbArray(row + 1, col + 1, 3) = b
        Next col
    Next row
    
    GlobalUnlock hClip
    CloseClipboard
    
    ClipboardImageToRGB = rgbArray
End Function

' ===========================================================
' 4) Read multiple image files from clipboard (CF_HDROP)
' ===========================================================
Public Function ClipboardFilesToRGB() As Collection
    Dim col As New Collection
    Dim hDrop As LongPtr
    Dim fileCount As Long, i As Long, fileName As String, Length As Long
    Dim RGBArr As Variant
    
    ' Check if CF_HDROP format is available
    If IsClipboardFormatAvailable(CF_HDROP) = 0 Then
        Set ClipboardFilesToRGB = col
        Exit Function
    End If
    
    If OpenClipboard(0) = 0 Then
        Set ClipboardFilesToRGB = col
        Exit Function
    End If
    
    hDrop = GetClipboardData(CF_HDROP)
    If hDrop = 0 Then CloseClipboard: Set ClipboardFilesToRGB = col: Exit Function
    
    ' Count files
    fileCount = DragQueryFile(hDrop, &HFFFFFFFF, vbNullString, 0)
    
    ' Loop through files
    For i = 0 To fileCount - 1
        Length = DragQueryFile(hDrop, i, vbNullString, 0) + 1
        fileName = String(Length, vbNullChar)
        DragQueryFile hDrop, i, fileName, Length
        fileName = Left(fileName, Length - 1) ' remove trailing null
        
        ' Load image file to RGB
        RGBArr = LoadImageToRGB(fileName)
        If Not IsEmpty(RGBArr) Then
            col.Add RGBArr, fileName
        End If
    Next i
    
    CloseClipboard
    Set ClipboardFilesToRGB = col
End Function

' ===========================================================
' 5) Helper Functions
' ===========================================================
Private Function PeekByte(ByVal ptr As LongPtr) As Byte
    Dim b As Byte
    RtlMoveMemory b, ByVal ptr, 1
    PeekByte = b
End Function

Private Function GetLongFromBytes(arr() As Byte, offset As Long) As Long
    GetLongFromBytes = arr(offset) + arr(offset + 1) * 256 + arr(offset + 2) * 65536 + arr(offset + 3) * 16777216
End Function

Sub usageexample()
' Load from file
Dim rgbArray As Variant
rgbArray = LoadImageToRGB("C:\Images\test.png")

' Load single image from clipboard
rgbArray = ClipboardImageToRGB()

' Load multiple image files from clipboard
Dim col As Collection, arr As Variant, key As Variant
Set col = ClipboardFilesToRGB()
For Each key In col
    arr = col(key)
    Debug.Print "File: " & key & " Height=" & UBound(arr, 1) & " Width=" & UBound(arr, 2)
Next key
End Sub
