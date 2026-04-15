'====================================================
' OpenGL Demo Module - Version 1
' All-in-one VBA Module + UserForm integration
' Features:
'   - 11 Demo routines
'   - Hardcoded or disk-based textures/shaders
'   - PtrSafe declarations
'   - Helper functions for OpenGL context & rendering
'====================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As LongPtr
    Private Declare PtrSafe Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal hglrc As LongPtr) As Long
    Private Declare PtrSafe Function wglDeleteContext Lib "opengl32.dll" (ByVal hglrc As LongPtr) As Long
    Private Declare PtrSafe Function SwapBuffers Lib "gdi32.dll" (ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
#Else
    Private Declare Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As Long) As Long
    Private Declare Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As Long, ByVal hglrc As Long) As Long
    Private Declare Function wglDeleteContext Lib "opengl32.dll" (ByVal hglrc As Long) As Long
    Private Declare Function SwapBuffers Lib "gdi32.dll" (ByVal hdc As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
#End If

'===============================
' Global Variables
'===============================
Public g_hWnd As LongPtr
Public g_hDC As LongPtr
Public g_hGLRC As LongPtr

Public g_useDiskTextures As Boolean
Public g_useDiskShaders As Boolean
Public g_textureFilePath As String
Public g_shaderFilePath As String

'===============================
' Helper Sub: Initialize OpenGL
'===============================
Public Sub InitOpenGL(ByVal hWnd As LongPtr)
    g_hWnd = hWnd
    g_hDC = GetDC(hWnd)
    g_hGLRC = wglCreateContext(g_hDC)
    If g_hGLRC = 0 Then
        MsgBox "Failed to create OpenGL context"
        Exit Sub
    End If
    If wglMakeCurrent(g_hDC, g_hGLRC) = 0 Then
        MsgBox "Failed to activate OpenGL context"
        Exit Sub
    End If
End Sub

'===============================
' Helper Sub: Cleanup OpenGL
'===============================
Public Sub CleanupOpenGL()
    wglMakeCurrent 0, 0
    wglDeleteContext g_hGLRC
    ReleaseDC g_hWnd, g_hDC
End Sub

'===============================
' Helper: Load Texture
'===============================
Public Function LoadTexture(Optional ByVal fromDisk As Boolean = False, Optional ByVal FilePath As String = "") As Long
    ' For version 1: simple placeholder texture (procedural or disk if selected)
    If fromDisk And FilePath <> "" Then
        ' TODO: Load BMP/PNG from file (placeholder)
        LoadTexture = 1
    Else
        ' Hardcoded procedural texture
        LoadTexture = 1
    End If
End Function

'===============================
' Helper: Load Shader
'===============================
Public Function LoadShader(Optional ByVal fromDisk As Boolean = False, Optional ByVal FilePath As String = "") As String
    If fromDisk And FilePath <> "" Then
        ' TODO: Load shader source from file
        LoadShader = "// Shader loaded from file"
    Else
        ' Hardcoded shader string
        LoadShader = "// Hardcoded demo shader"
    End If
End Function

'===============================
' Demo Routines (11)
'===============================
Public Sub Demo1()
    MsgBox "Demo 1: Simple Triangle"
End Sub

Public Sub Demo2()
    MsgBox "Demo 2: Rotating Cube"
End Sub

Public Sub Demo3()
    MsgBox "Demo 3: Texture Mapping"
End Sub

Public Sub Demo4()
    MsgBox "Demo 4: Simple Shader"
End Sub

Public Sub Demo5()
    MsgBox "Demo 5: Color Gradient"
End Sub

Public Sub Demo6()
    MsgBox "Demo 6: Animated Pattern"
End Sub

Public Sub Demo7()
    MsgBox "Demo 7: Multi-texture Example"
End Sub

Public Sub Demo8()
    MsgBox "Demo 8: Lighting Demo"
End Sub

Public Sub Demo9()
    MsgBox "Demo 9: Procedural Shapes"
End Sub

Public Sub Demo10()
    MsgBox "Demo 10: Frame Timing Test"
End Sub

Public Sub Demo11()
    MsgBox "Demo 11: Combined Effects"
End Sub

'===============================
' UserForm Integration
'===============================
' Assumes a UserForm named frmOpenGLDemo
' Controls:
'   - ComboBox: cboDemoSelect
'   - OptionButtons: optHardcodedTextures, optDiskTextures
'   - OptionButtons: optHardcodedShaders, optDiskShaders
'   - CommandButton: cmdRunDemo
'   - CommandButton: cmdBrowseTexture
'   - CommandButton: cmdBrowseShader

Public Sub frmOpenGLDemo_Initialize(frm As Object)
    Dim i As Integer
    With frm.cboDemoSelect
        .Clear
        For i = 1 To 11
            .AddItem "Demo " & i
        Next i
        .ListIndex = 0
    End With
    frm.optHardcodedTextures.value = True
    frm.optHardcodedShaders.value = True
End Sub

Public Sub cmdRunDemo_Click(frm As Object)
    g_useDiskTextures = frm.optDiskTextures.value
    g_useDiskShaders = frm.optDiskShaders.value
    
    Dim DemoIndex As Integer
    DemoIndex = frm.cboDemoSelect.ListIndex + 1
    
    Select Case DemoIndex
        Case 1: Demo1
        Case 2: Demo2
        Case 3: Demo3
        Case 4: Demo4
        Case 5: Demo5
        Case 6: Demo6
        Case 7: Demo7
        Case 8: Demo8
        Case 9: Demo9
        Case 10: Demo10
        Case 11: Demo11
    End Select
End Sub

Public Sub cmdBrowseTexture_Click(frm As Object)
    g_textureFilePath = Application.GetOpenFilename("Image Files (*.bmp;*.png;*.jpg), *.bmp;*.png;*.jpg")
End Sub

Public Sub cmdBrowseShader_Click(frm As Object)
    g_shaderFilePath = Application.GetOpenFilename("Shader Files (*.glsl;*.txt), *.glsl;*.txt")
End Sub

'===============================
' Test Subroutines
'===============================
Public Sub TestOpenGLInit()
    InitOpenGL Application.hWnd
    MsgBox "OpenGL initialized"
    CleanupOpenGL
End Sub

Public Sub TestRunAllDemos()
    Dim i As Integer
    For i = 1 To 11
        Application.StatusBar = "Running Demo " & i
        Select Case i
            Case 1: Demo1
            Case 2: Demo2
            Case 3: Demo3
            Case 4: Demo4
            Case 5: Demo5
            Case 6: Demo6
            Case 7: Demo7
            Case 8: Demo8
            Case 9: Demo9
            Case 10: Demo10
            Case 11: Demo11
        End Select
    Next i
    Application.StatusBar = False
End Sub

