Option Explicit

' ======================================================
' Module: modVideoPlayer
' Purpose: Dynamically create a UserForm with video player
'          using MCI (winmm.dll) — no ActiveX.
' Features:
'   - Auto-creates form only if missing
'   - Play / Pause / Resume / Stop buttons
'   - Auto-resizes video when form resized
'   - Can hardcode or prompt for video path
' ======================================================

#If VBA7 Then
    Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long

    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
         ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
#Else
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
         ByVal lpszClass As String, ByVal lpszWindow As String) As Long
#End If


' ======================================================
' Public Entry Point
' ======================================================
Public Sub ShowVideoForm()
    EnsureVideoForm
End Sub


' ======================================================
' Ensure form exists (create/import if missing)
' ======================================================
Private Sub EnsureVideoForm()
    Dim vbProj As Object, comp As Object
    Dim exists As Boolean
    Dim frmFile As String
    
    Set vbProj = ThisWorkbook.VBProject
    
    ' --- Check if form already exists ---
    For Each comp In vbProj.VBComponents
        If comp.Type = 3 Then ' vbext_ct_MSForm
            If comp.Name = "frmVideoPlayer" Then
                exists = True
                Exit For
            End If
        End If
    Next
    
    ' --- If not, create it ---
    If Not exists Then
        frmFile = GetSafeTempPath & "\frmVideoPlayer.frm"
        WriteVideoFormFile frmFile
        vbProj.VBComponents.Import frmFile
        On Error Resume Next
        Kill frmFile
        On Error GoTo 0
    End If
    
    ' --- Show the form ---
    VBA.UserForms.Add("frmVideoPlayer").Show
End Sub


' ======================================================
' Safe Temp Path helper
' ======================================================
Private Function GetSafeTempPath() As String
    Dim basePath As String
    basePath = ThisWorkbook.path
    
    ' --- Always prefer local path, but fall back if suspicious ---
    If Len(basePath) = 0 Then
        ' Workbook never saved
        GetSafeTempPath = Environ$("TEMP")
    ElseIf IsLikelyOnlinePath(basePath) Then
        GetSafeTempPath = Environ$("TEMP")
    Else
        GetSafeTempPath = basePath
    End If
End Function

Private Function IsLikelyOnlinePath(p As String) As Boolean
    Dim check As String
    check = LCase$(p)
    
    ' --- Detect common online/cloud keywords ---
    If InStr(check, "http://") > 0 _
       Or InStr(check, "https://") > 0 _
       Or InStr(check, "sharepoint") > 0 _
       Or InStr(check, "onedrive") > 0 _
       Or InStr(check, "teams") > 0 _
       Or InStr(check, "://") > 0 Then
        IsLikelyOnlinePath = True
    Else
        IsLikelyOnlinePath = False
    End If
End Function


' ======================================================
' Create the .frm source file dynamically
' ======================================================
Private Sub WriteVideoFormFile(ByVal path As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    
    Print #f, "VERSION 5.00"
    Print #f, "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVideoPlayer "
    Print #f, "   Caption         =   ""Video Player"""
    Print #f, "   ClientHeight    =   6000"
    Print #f, "   ClientWidth     =   9000"
    Print #f, "   StartUpPosition =   1  'CenterOwner"
    Print #f, "   Begin Frame Frame1 "
    Print #f, "      Caption         =   ""Video"""
    Print #f, "      Height          =   5000"
    Print #f, "      Width           =   8500"
    Print #f, "      Left            =   200"
    Print #f, "      Top             =   200"
    Print #f, "   End"
    Print #f, "   Begin CommandButton cmdPlay "
    Print #f, "      Caption         =   ""Play"""
    Print #f, "      Height          =   300"
    Print #f, "      Width           =   1000"
    Print #f, "      Left            =   200"
    Print #f, "      Top             =   5300"
    Print #f, "   End"
    Print #f, "   Begin CommandButton cmdPause "
    Print #f, "      Caption         =   ""Pause"""
    Print #f, "      Height          =   300"
    Print #f, "      Width           =   1000"
    Print #f, "      Left            =   1500"
    Print #f, "      Top             =   5300"
    Print #f, "   End"
    Print #f, "   Begin CommandButton cmdResume "
    Print #f, "      Caption         =   ""Resume"""
    Print #f, "      Height          =   300"
    Print #f, "      Width           =   1000"
    Print #f, "      Left            =   2800"
    Print #f, "      Top             =   5300"
    Print #f, "   End"
    Print #f, "   Begin CommandButton cmdStop "
    Print #f, "      Caption         =   ""Stop"""
    Print #f, "      Height          =   300"
    Print #f, "      Width           =   1000"
    Print #f, "      Left            =   4100"
    Print #f, "      Top             =   5300"
    Print #f, "   End"
    Print #f, "End"
    Print #f, "Attribute VB_Name = ""frmVideoPlayer"""
    Print #f, "Attribute VB_GlobalNameSpace = False"
    Print #f, "Attribute VB_Creatable = False"
    Print #f, "Attribute VB_PredeclaredId = True"
    Print #f, "Attribute VB_Exposed = False"
    
    ' --- Inject code behind the form ---
    Print #f, "Private Sub cmdPlay_Click()"
    Print #f, "    Dim f As String"
    Print #f, "    f = Application.GetOpenFilename(""Video Files (*.avi;*.mpg;*.wmv),*.avi;*.mpg;*.wmv"")"
    Print #f, "    If f <> ""False"" Then PlayVideoInFrame Me, ""Frame1"", f"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Private Sub cmdPause_Click()"
    Print #f, "    mciSendString ""pause MyVideo"", vbNullString, 0, 0"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Private Sub cmdResume_Click()"
    Print #f, "    mciSendString ""resume MyVideo"", vbNullString, 0, 0"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Private Sub cmdStop_Click()"
    Print #f, "    mciSendString ""stop MyVideo"", vbNullString, 0, 0"
    Print #f, "    mciSendString ""close MyVideo"", vbNullString, 0, 0"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Private Sub UserForm_Resize()"
    Print #f, "    ResizeVideo Me, ""Frame1"""
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Private Sub UserForm_Terminate()"
    Print #f, "    mciSendString ""close MyVideo"", vbNullString, 0, 0"
    Print #f, "End Sub"
    
    Close #f
End Sub


' ======================================================
' Video Helpers
' ======================================================
Public Sub PlayVideoInFrame(frm As Object, frameName As String, videoPath As String)
    Dim hFrame As LongPtr, cmd As String
    
    hFrame = GetFrameHwnd(frm, frameName)
    If hFrame = 0 Then
        MsgBox "Frame not found!", vbCritical
        Exit Sub
    End If
    
    mciSendString "close MyVideo", vbNullString, 0, 0
    
    cmd = "open """ & videoPath & """ type mpegvideo alias MyVideo parent " & CStr(hFrame)
    mciSendString cmd, vbNullString, 0, 0
    
    ResizeVideo frm, frameName
    
    mciSendString "play MyVideo", vbNullString, 0, 0
End Sub

Public Sub ResizeVideo(frm As Object, frameName As String)
    Dim cmd As String, w As Long, h As Long
    w = frm.Controls(frameName).width
    h = frm.Controls(frameName).height
    cmd = "put MyVideo window at 0 0 " & w & " " & h
    mciSendString cmd, vbNullString, 0, 0
End Sub

Private Function GetFrameHwnd(frm As Object, frameName As String) As LongPtr
    Dim hForm As LongPtr, hFrame As LongPtr
    hForm = FindWindow("ThunderDFrame", vbNullString)
    hFrame = FindWindowEx(hForm, 0&, "ThunderFrame", vbNullString)
    GetFrameHwnd = hFrame
End Function


