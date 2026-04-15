
Option Explicit

' Module to dynamically create a WIA UserForm and integrate with WIAModule.bas
' Requires WIAModule.bas (artifact_id: 1c4c5fd7-9328-404f-acc7-ca0ddd734645) in the project
' Uses Collection to avoid VBA's 24-line continuation limit

' Constants for WIA formats (must match WIAModule.bas)
Private Const WiaFormat_JPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Private Const WiaFormat_PNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
Private Const WiaFormat_BMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
Private Const WiaFormat_TIFF As String = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"

' Sub to create and show the dynamic UserForm
Public Sub CreateAndShowWIAForm()
    Dim frm As Object
    Dim vbProj As Object
    Dim vbComp As Object
    Dim frmCodeColl As Collection
    Dim frmCode As String
    Dim ctl As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Initialize Collection for UserForm code
    Set frmCodeColl = New Collection
    
    ' Add UserForm code segments to Collection
    frmCodeColl.Add "Option Explicit"
    
    frmCodeColl.Add "Private Sub UserForm_Initialize()" & vbCrLf & _
                    "    PopulateDeviceList" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub PopulateDeviceList()" & vbCrLf & _
                    "    Dim wiaManager As Object" & vbCrLf & _
                    "    Dim deviceInfo As Object" & vbCrLf & _
                    "    Dim i As Long" & vbCrLf & _
                    "    On Error Resume Next" & vbCrLf & _
                    "    Set wiaManager = CreateObject(""WIA.DeviceManager"")" & vbCrLf & _
                    "    Me.cboDevice.Clear" & vbCrLf & _
                    "    For i = 1 To wiaManager.DeviceInfos.Count" & vbCrLf & _
                    "        Set deviceInfo = wiaManager.DeviceInfos(i)" & vbCrLf & _
                    "        Me.cboDevice.AddItem deviceInfo.Properties(""Name"").Value & "" ("" & GetDeviceTypeName(deviceInfo.Type) & "")""" & vbCrLf & _
                    "    Next i" & vbCrLf & _
                    "    If Me.cboDevice.ListCount > 0 Then Me.cboDevice.ListIndex = 0" & vbCrLf & _
                    "    Set deviceInfo = Nothing" & vbCrLf & _
                    "    Set wiaManager = Nothing" & vbCrLf & _
                    "    On Error GoTo 0" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Function GetDeviceTypeName(deviceType As Long) As String" & vbCrLf & _
                    "    Select Case deviceType" & vbCrLf & _
                    "        Case 1: GetDeviceTypeName = ""Scanner""" & vbCrLf & _
                    "        Case 2: GetDeviceTypeName = ""Camera""" & vbCrLf & _
                    "        Case 3: GetDeviceTypeName = ""Video""" & vbCrLf & _
                    "        Case Else: GetDeviceTypeName = ""Unknown""" & vbCrLf & _
                    "    End Select" & vbCrLf & _
                    "End Function"
    
    frmCodeColl.Add "Private Sub cmdBrowse_Click()" & vbCrLf & _
                    "    Dim shellApp As Object" & vbCrLf & _
                    "    Dim folder As Object" & vbCrLf & _
                    "    Set shellApp = CreateObject(""Shell.Application"")" & vbCrLf & _
                    "    Set folder = shellApp.BrowseForFolder(0, ""Select Save Folder"", 0, 0)" & vbCrLf & _
                    "    If Not folder Is Nothing Then" & vbCrLf & _
                    "        Me.txtSavePath.Text = folder.Self.Path & ""\ScannedImage"" & GetExtension(Me.cboFormat.Text)" & vbCrLf & _
                    "    End If" & vbCrLf & _
                    "    Set folder = Nothing" & vbCrLf & _
                    "    Set shellApp = Nothing" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Function GetExtension(format As String) As String" & vbCrLf & _
                    "    Select Case format" & vbCrLf & _
                    "        Case ""JPEG"": GetExtension = "".jpg""" & vbCrLf & _
                    "        Case ""PNG"": GetExtension = "".png""" & vbCrLf & _
                    "        Case ""BMP"": GetExtension = "".bmp""" & vbCrLf & _
                    "        Case ""TIFF"": GetExtension = "".tiff""" & vbCrLf & _
                    "        Case Else: GetExtension = "".jpg""" & vbCrLf & _
                    "    End Select" & vbCrLf & _
                    "End Function"
    
    frmCodeColl.Add "Private Sub cboFormat_Change()" & vbCrLf & _
                    "    Dim path As String" & vbCrLf & _
                    "    Dim ext As String" & vbCrLf & _
                    "    path = Me.txtSavePath.Text" & vbCrLf & _
                    "    ext = GetExtension(Me.cboFormat.Text)" & vbCrLf & _
                    "    If InStrRev(path, ""."") > 0 Then" & vbCrLf & _
                    "        path = Left(path, InStrRev(path, ""."") - 1)" & vbCrLf & _
                    "        Me.txtSavePath.Text = path & ext" & vbCrLf & _
                    "    End If" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdListDevices_Click()" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Listing devices...""" & vbCrLf & _
                    "    Call ListWIADepartments" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Ready""" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdScan_Click()" & vbCrLf & _
                    "    On Error GoTo ErrHandle" & vbCrLf & _
                    "    Dim format As String" & vbCrLf & _
                    "    Dim res As Long" & vbCrLf & _
                    "    Dim colorMode As Long" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Scanning...""" & vbCrLf & _
                    "    Select Case Me.cboFormat.Text" & vbCrLf & _
                    "        Case ""JPEG"": format = """ & WiaFormat_JPEG & """" & vbCrLf & _
                    "        Case ""PNG"": format = """ & WiaFormat_PNG & """" & vbCrLf & _
                    "        Case ""BMP"": format = """ & WiaFormat_BMP & """" & vbCrLf & _
                    "        Case ""TIFF"": format = """ & WiaFormat_TIFF & """" & vbCrLf & _
                    "    End Select" & vbCrLf & _
                    "    res = CLng(Me.txtResolution.Text)" & vbCrLf & _
                    "    Select Case Me.cboColorMode.Text" & vbCrLf & _
                    "        Case ""Color"": colorMode = 1" & vbCrLf & _
                    "        Case ""Grayscale"": colorMode = 2" & vbCrLf & _
                    "        Case ""Black/White"": colorMode = 4" & vbCrLf & _
                    "    End Select" & vbCrLf & _
                    "    Call ScanImage(format, res, colorMode)" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Scan complete.""" & vbCrLf & _
                    "    Exit Sub" & vbCrLf & _
                    "ErrHandle:" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Error: "" & Err.Description" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdCapture_Click()" & vbCrLf & _
                    "    On Error GoTo ErrHandle" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Capturing photo...""" & vbCrLf & _
                    "    Call CapturePhoto" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Photo captured.""" & vbCrLf & _
                    "    Exit Sub" & vbCrLf & _
                    "ErrHandle:" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Error: "" & Err.Description" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdResize_Click()" & vbCrLf & _
                    "    On Error GoTo ErrHandle" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Scanning and resizing...""" & vbCrLf & _
                    "    Call ScanAndResizeImage(0.5)" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Resize complete.""" & vbCrLf & _
                    "    Exit Sub" & vbCrLf & _
                    "ErrHandle:" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Error: "" & Err.Description" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdEmbed_Click()" & vbCrLf & _
                    "    On Error GoTo ErrHandle" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Scanning and embedding...""" & vbCrLf & _
                    "    Call ScanAndEmbedInExcel" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Image embedded.""" & vbCrLf & _
                    "    Exit Sub" & vbCrLf & _
                    "ErrHandle:" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Error: "" & Err.Description" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdLog_Click()" & vbCrLf & _
                    "    On Error GoTo ErrHandle" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Scanning and logging...""" & vbCrLf & _
                    "    Call ScanAndLogMetadata" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Metadata logged.""" & vbCrLf & _
                    "    Exit Sub" & vbCrLf & _
                    "ErrHandle:" & vbCrLf & _
                    "    Me.lblStatus.Caption = ""Error: "" & Err.Description" & vbCrLf & _
                    "End Sub"
    
    frmCodeColl.Add "Private Sub cmdCancel_Click()" & vbCrLf & _
                    "    Unload Me" & vbCrLf & _
                    "End Sub"
    
    ' Concatenate code segments
    For i = 1 To frmCodeColl.count
        frmCode = frmCode & frmCodeColl(i) & vbCrLf
    Next i
    
    ' Get VBProject
    Set vbProj = ThisWorkbook.VBProject
    
    ' Check if UserForm already exists and delete it
    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = "WIAControlForm" Then
            vbProj.VBComponents.Remove vbComp
            Exit For
        End If
    Next vbComp
    
    ' Create new UserForm
    Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm = 3
    vbComp.Name = "WIAControlForm"
    Set frm = vbComp.Designer
    
    ' Set UserForm properties
    With frm
        .Properties("Caption") = "WIA Control Panel"
        .Properties("Width") = 400
        .Properties("Height") = 300
    End With
    
    ' Add controls to UserForm
    ' Label: Device Selection
    Set ctl = frm.Controls.Add("Forms.Label.1")
    With ctl
        .Caption = "Select Device:"
        .Left = 10
        .Top = 10
        .width = 80
        .height = 15
    End With
    
    ' ComboBox: Device List
    Set ctl = frm.Controls.Add("Forms.ComboBox.1", "cboDevice")
    With ctl
        .Left = 100
        .Top = 10
        .width = 250
        .height = 20
    End With
    
    ' Label: Image Format
    Set ctl = frm.Controls.Add("Forms.Label.1")
    With ctl
        .Caption = "Image Format:"
        .Left = 10
        .Top = 40
        .width = 80
        .height = 15
    End With
    
    ' ComboBox: Image Format
    Set ctl = frm.Controls.Add("Forms.ComboBox.1", "cboFormat")
    With ctl
        .Left = 100
        .Top = 40
        .width = 100
        .height = 20
        .AddItem "JPEG"
        .AddItem "PNG"
        .AddItem "BMP"
        .AddItem "TIFF"
        .ListIndex = 0 ' Default to JPEG
    End With
    
    ' Label: Resolution
    Set ctl = frm.Controls.Add("Forms.Label.1")
    With ctl
        .Caption = "Resolution (DPI):"
        .Left = 10
        .Top = 70
        .width = 80
        .height = 15
    End With
    
    ' TextBox: Resolution
    Set ctl = frm.Controls.Add("Forms.TextBox.1", "txtResolution")
    With ctl
        .Left = 100
        .Top = 70
        .width = 100
        .height = 20
        .Text = "300"
    End With
    
    ' Label: Color Mode
    Set ctl = frm.Controls.Add("Forms.Label.1")
    With ctl
        .Caption = "Color Mode:"
        .Left = 10
        .Top = 100
        .width = 80
        .height = 15
    End With
    
    ' ComboBox: Color Mode
    Set ctl = frm.Controls.Add("Forms.ComboBox.1", "cboColorMode")
    With ctl
        .Left = 100
        .Top = 100
        .width = 100
        .height = 20
        .AddItem "Color"
        .AddItem "Grayscale"
        .AddItem "Black/White"
        .ListIndex = 0 ' Default to Color
    End With
    
    ' Label: Save Path
    Set ctl = frm.Controls.Add("Forms.Label.1")
    With ctl
        .Caption = "Save Path:"
        .Left = 10
        .Top = 130
        .width = 80
        .height = 15
    End With
    
    ' TextBox: Save Path
    Set ctl = frm.Controls.Add("Forms.TextBox.1", "txtSavePath")
    With ctl
        .Left = 100
        .Top = 130
        .width = 200
        .height = 20
        .Text = Environ("USERPROFILE") & "\Desktop\ScannedImage.jpg"
    End With
    
    ' Button: Browse
    Set ctl = frm.Controls.Add("Forms.CommandButton.1", "cmdBrowse")
    With ctl
        .Caption = "Browse..."
        .Left = 310
        .Top = 130
        .width = 60
        .height = 20
    End With
    
    ' Action Buttons
    Dim btnCaptions As Variant
    Dim btnNames As Variant
    btnCaptions = Array("List Devices", "Scan Image", "Capture Photo", "Scan and Resize", "Scan and Embed", "Scan and Log", "Cancel")
    btnNames = Array("cmdListDevices", "cmdScan", "cmdCapture", "cmdResize", "cmdEmbed", "cmdLog", "cmdCancel")
    
    For i = 0 To UBound(btnCaptions)
        Set ctl = frm.Controls.Add("Forms.CommandButton.1", btnNames(i))
        With ctl
            .Caption = btnCaptions(i)
            .Left = 10 + (i Mod 2) * 190
            .Top = 160 + (i \ 2) * 30
            .width = 180
            .height = 25
        End With
    Next i
    
    ' Label: Status
    Set ctl = frm.Controls.Add("Forms.Label.1", "lblStatus")
    With ctl
        .Caption = "Ready"
        .Left = 10
        .Top = 260
        .width = 350
        .height = 20
        .ForeColor = &H800000 ' Blue
    End With
    
    ' Add code to UserForm's code module
    vbComp.CodeModule.AddFromString frmCode
    
    ' Populate device list on form load
    Dim wiaManager As Object
    Set wiaManager = CreateObject("WIA.DeviceManager")
    With frm.cboDevice
        .Clear
        For i = 1 To wiaManager.DeviceInfos.count
            .AddItem wiaManager.DeviceInfos(i).Properties("Name").value & " (" & GetDeviceTypeName(wiaManager.DeviceInfos(i).Type) & ")"
        Next i
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    ' Show the form
    VBA.UserForms.Add("WIAControlForm").Show
    
    ' Clean up
    Set wiaManager = Nothing
    Set frm = Nothing
    Set vbComp = Nothing
    Set vbProj = Nothing
    Set frmCodeColl = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error creating UserForm: " & Err.description, vbCritical
    If Not frm Is Nothing Then Set frm = Nothing
    If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
    If Not vbProj Is Nothing Then Set vbProj = Nothing
    If Not frmCodeColl Is Nothing Then Set frmCodeColl = Nothing
End Sub

' Helper Function: Get device type name (used in form initialization)
Private Function GetDeviceTypeName(deviceType As Long) As String
    Select Case deviceType
        Case 1: GetDeviceTypeName = "Scanner"
        Case 2: GetDeviceTypeName = "Camera"
        Case 3: GetDeviceTypeName = "Video"
        Case Else: GetDeviceTypeName = "Unknown"
    End Select
End Function

