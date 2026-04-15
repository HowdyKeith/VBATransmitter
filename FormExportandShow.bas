'https://chatgpt.com/c/68b8ad64-7c48-8333-ac22-7f7d46151c3d

Sub CreateAndExportForm()
    Dim vbComp As Object
    Dim exportPath As String
    Dim frm As Object
    
    exportPath = "C:\Temp\MyDynamicForm.frm"
    
    ' Add a new UserForm to the project
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    vbComp.Name = "MyDynamicForm"
    
    ' Reference the UserForm object
    Set frm = vbComp.Designer
    
    ' Add a label
    With frm.Controls.Add("Forms.Label.1")
        .Caption = "Hello from dynamic form!"
        .Left = 20
        .Top = 20
        .width = 200
    End With
    
    ' Add a command button
    With frm.Controls.Add("Forms.CommandButton.1")
        .Caption = "Click Me"
        .Left = 20
        .Top = 60
        .width = 100
    End With
    
    ' Export the UserForm as .frm
    vbComp.Export exportPath
    
    MsgBox "Dynamic UserForm exported to: " & exportPath
End Sub


Sub CreateExportImportShowForm()
    Dim vbComp As Object
    Dim exportPath As String
    Dim frm As Object
    Dim formName As String
    
    formName = "DynamicForm1"
    exportPath = Environ("TEMP") & "\" & formName & ".frm" ' Example: C:\Users\YourUser\AppData\Local\Temp\DynamicForm1.frm
    
    ' --- Step 1: Add a new UserForm ---
    On Error Resume Next
    ' Delete existing form if it exists
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(formName)
    On Error GoTo 0
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    vbComp.Name = formName
    
    ' --- Step 2: Add controls to the form ---
    Set frm = vbComp.Designer
    
    ' Add Label
    With frm.Controls.Add("Forms.Label.1")
        .Caption = "Hello from dynamic form!"
        .Left = 20
        .Top = 20
        .width = 200
    End With
    
    ' Add CommandButton
    With frm.Controls.Add("Forms.CommandButton.1")
        .Caption = "Click Me"
        .Left = 20
        .Top = 60
        .width = 100
    End With
    
    ' --- Step 3: Export the form ---
    vbComp.Export exportPath
    MsgBox "Form exported to: " & exportPath
    
    ' --- Step 4: Remove and re-import the form (simulates full cycle) ---
    ThisWorkbook.VBProject.VBComponents.Remove vbComp
    ThisWorkbook.VBProject.VBComponents.Import exportPath
    
    ' --- Step 5: Show the imported form ---
    Application.VBE.MainWindow.Visible = False ' optional, hide VBE
    VBA.UserForms.Add(formName).Show vbModeless
End Sub

' ============================================================
' Launch OpenGL demo inside a dynamic UserForm
' Optional parameter: DeleteFormAfterClose
' ============================================================
Public Sub ShowOpenGLDemoInForm(demoName As String, Optional DeleteFormAfterClose As Boolean = False)
    Dim frm As Object ' UserForm
    Dim hwndForm As LongPtr
    Dim formName As String
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    
    formName = "TempOpenGLForm"
    exportPath = Environ("TEMP") & "\" & formName & ".frm"
    
    ' --- Step 1: Create a temporary UserForm dynamically ---
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(formName)
    On Error GoTo 0
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    vbComp.Name = formName
    
    Set frm = vbComp.Designer
    With frm
        .Caption = demoName
        .width = 820
        .height = 640
    End With
    
    ' Optional: Add a Close button
    With frm.Controls.Add("Forms.CommandButton.1")
        .Caption = "Close"
        .Left = 20
        .Top = 580
        .width = 100
        .height = 30
    End With
    
    ' Export form (optional, keeps file)
    vbComp.Export exportPath
    
    ' Show the form modeless
    VBA.UserForms.Add(formName).Show vbModeless
    
    ' --- Step 2: Get HWND of the form ---
    hwndForm = FindWindow(vbNullString, demoName)
    If hwndForm = 0 Then
        MsgBox "Could not get UserForm handle.", vbCritical
        Exit Sub
    End If
    
    ' --- Step 3: Initialize OpenGL ---
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
    
    ' --- Step 4: Optional cleanup after form closes ---
    If DeleteFormAfterClose Then
        ' Wait until form is unloaded, then remove VBComponent
        DoEvents
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
        On Error GoTo 0
    End If
End Sub

