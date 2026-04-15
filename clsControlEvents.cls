' === Class Module Name: clsControlEvents ===
Option Explicit

Public WithEvents chkBox As MSForms.CheckBox
Public WithEvents optBtn As MSForms.OptionButton
Public WithEvents cmdBtn As MSForms.CommandButton
Public WithEvents cboBox As MSForms.ComboBox

Private Sub chkBox_Click()
    Select Case chkBox.Name
        Case "chkVSync": glVSyncEnabled = chkBox.value: SetVSync glVSyncEnabled
        Case "chkMultisample": glMSAAEnabled = chkBox.value: SetMSAA glSampleCount
        Case "chkUseAdvancedShaders": glAdvancedShaders = chkBox.value
    End Select
End Sub

Private Sub optBtn_Click()
    Select Case optBtn.Name
        Case "optTextureHardcoded": glTextureModeHardcoded = True
        Case "optTextureFromDisk": glTextureModeHardcoded = False
        Case "optShaderHardcoded": glShaderModeHardcoded = True
        Case "optShaderFromDisk": glShaderModeHardcoded = False
    End Select
End Sub

Private Sub cmdBtn_Click()
    Dim f As Variant
    Select Case cmdBtn.Name
        Case "btnLoadTexture"
            f = Application.GetOpenFilename("Images (*.png;*.bmp), *.png;*.bmp")
            If f <> False Then LoadTexture f
        Case "btnLoadShader"
            f = Application.GetOpenFilename("Shader Files (*.glsl), *.glsl")
            If f <> False Then LoadShader f
        Case "btnRunDemo"
            Demo1
    End Select
End Sub

Private Sub cboBox_Change()
    If cboBox.Name = "cboSamples" Then
        glSampleCount = val(cboBox.Text)
        SetMSAA glSampleCount
    End If
End Sub

