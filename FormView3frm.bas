' ============================================================
' OpenGL Demo Module with Dynamic UserForm Support (Late Binding)
' ============================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

' ============================================================
' Launch OpenGL demo inside dynamic form
' ============================================================
Public Sub ShowOpenGLDemoInForm(demoName As String, Optional DeleteFormAfterClose As Boolean = False)
    Dim frm As Object
    Dim hwndForm As LongPtr
    Dim formName As String
    Dim vbComp As Object
    Dim exportPath As String
    
    formName = "TempOpenGLForm"
    exportPath = Environ("TEMP") & "\" & formName & ".frm"
    
    ' --- Remove existing form if exists ---
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(formName)
    On Error GoTo 0
    
    ' --- Create dynamic UserForm ---
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    vbComp.Name = formName
    
    Set frm = vbComp.Designer
    With frm
        .Caption = demoName
        .width = 820
        .height = 640
    End With
    
    ' Add Close button
    With frm.Controls.Add("Forms.CommandButton.1")
        .Caption = "Close"
        .Left = 20
        .Top = 580
        .width = 100
        .height = 30
    End With
    
    ' Export form
    vbComp.Export exportPath
    
    ' Show form modeless
    VBA.UserForms.Add(formName).Show vbModeless
    
    ' --- Get HWND ---
    hwndForm = FindWindow(vbNullString, demoName)
    If hwndForm = 0 Then
        MsgBox "Could not get UserForm handle.", vbCritical
        Exit Sub
    End If
    
    ' --- Initialize OpenGL ---
    If InitializeOpenGLWithHWND(hwndForm) Then
        Select Case demoName
            Case "Array Similarity Visualization": VisualizeArrayComparisonDemo
            Case "3D Rotating Cube Demo": Demo3DRotatingCube
            Case "Simple 2D Game Demo": DemoSimple2DGame
            Case "Real-time Data Visualization": DemoDataVisualization
            Case "Particle System Demo": DemoParticleSystem
            Case "Mandelbrot Fractal Demo": DemoMandelbrot
            Case "Wireframe Sphere Demo": DemoWireframeSphere
            Case "Rotating Spiral Demo": DemoRotatingSpiral
            Case "3D Terrain Flyover Demo": DemoTerrainFlyover
            Case "Bouncing Balls Demo": DemoBouncingBalls
            Case "3D Torus Demo": DemoTorus
            Case Else
                MsgBox "Demo not recognized: " & demoName, vbExclamation
        End Select
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL in Form", vbCritical
    End If
    
    ' --- Optional cleanup ---
    If DeleteFormAfterClose Then
        DoEvents
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
        On Error GoTo 0
    End If
End Sub

' ============================================================
' ================= Test Subs ================================
' ============================================================
Public Sub TestArraySimilarityDemo()
    ShowOpenGLDemoInForm "Array Similarity Visualization", True
End Sub

Public Sub Test3DRotatingCubeDemo()
    ShowOpenGLDemoInForm "3D Rotating Cube Demo", True
End Sub

Public Sub TestSimple2DGameDemo()
    ShowOpenGLDemoInForm "Simple 2D Game Demo", True
End Sub

Public Sub TestDataVisualizationDemo()
    ShowOpenGLDemoInForm "Real-time Data Visualization", True
End Sub

Public Sub TestParticleSystemDemo()
    ShowOpenGLDemoInForm "Particle System Demo", True
End Sub

Public Sub TestMandelbrotDemo()
    ShowOpenGLDemoInForm "Mandelbrot Fractal Demo", True
End Sub

Public Sub TestWireframeSphereDemo()
    ShowOpenGLDemoInForm "Wireframe Sphere Demo", True
End Sub

Public Sub TestRotatingSpiralDemo()
    ShowOpenGLDemoInForm "Rotating Spiral Demo", True
End Sub

Public Sub Test3DTerrainFlyoverDemo()
    ShowOpenGLDemoInForm "3D Terrain Flyover Demo", True
End Sub

Public Sub TestBouncingBallsDemo()
    ShowOpenGLDemoInForm "Bouncing Balls Demo", True
End Sub

Public Sub Test3DTorusDemo()
    ShowOpenGLDemoInForm "3D Torus Demo", True
End Sub


