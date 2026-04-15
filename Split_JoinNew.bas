Option Explicit

' ================================================================
' VBA Code Export Utility - All In One Module
' ================================================================
' Features:
'   - Export text modules (.bas, .cls, .frm)
'   - Export binary contents
'   - Optional zip
'   - Optional split into AI-friendly chunks
'   - Front-End Picker
'   - Master Menu with Return-key default
'   - Export external workbook modules/classes in one consolidated text file
'   - Automatic subfolder for each export: WorkbookName - Export - DATE
' ================================================================

' ---------------------------
' Helpers: File & Path Utils
' ---------------------------
Private Function GetExportFolder() As String
    Dim fDialog As Object
    Set fDialog = Application.FileDialog(4) ' msoFileDialogFolderPicker
    With fDialog
        .title = "Select Export Folder"
        If .Show = -1 Then
            GetExportFolder = .SelectedItems(1)
        Else
            GetExportFolder = vbNullString
        End If
    End With
End Function

Private Function SafeFileName(ByVal Name As String) As String
    Dim badChars As Variant, c As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafeFileName = Name
    For Each c In badChars
        SafeFileName = Replace(SafeFileName, c, "_")
    Next c
End Function

Private Function AskYesNo(ByVal msg As String) As Boolean
    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbYesNo + vbQuestion)
    AskYesNo = (ans = vbYes)
End Function

' ================================================================
' Create Export Subfolder automatically
' ================================================================
Private Function CreateExportSubfolder(Optional ByVal wbName As String = "") As String
    Dim baseFolder As String, folderName As String, fso As Object
    baseFolder = GetExportFolder()
    If Len(baseFolder) = 0 Then Exit Function
    
    If wbName = "" Then wbName = ThisWorkbook.Name
    wbName = Replace(wbName, ".xlsm", "")
    wbName = Replace(wbName, ".xls", "")
    
    folderName = baseFolder & "\" & SafeFileName(wbName & " - Export - " & format(Now, "yyyymmdd_HHMMSS"))
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderName) Then fso.CreateFolder folderName
    
    CreateExportSubfolder = folderName
End Function

' ================================================================
' ===== Export as Text (.txt) =====
' ================================================================
Public Sub ExportModulesAsText()
    Dim vbProj As Object, vbComp As Object
    Dim exportFolder As String, fileName As String
    
    exportFolder = CreateExportSubfolder(ThisWorkbook.Name)
    If Len(exportFolder) = 0 Then Exit Sub
    
    Set vbProj = ThisWorkbook.VBProject
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type <> 3 Then
            fileName = exportFolder & "\" & SafeFileName(vbComp.Name) & ".txt"
            vbComp.Export fileName
        End If
    Next vbComp
    
    MsgBox "Modules exported as text to: " & exportFolder, vbInformation
    
    ' Optional zip
    If AskYesNo("Do you want to zip the exported folder?") Then Call ZipExportFolder(exportFolder)
    
    ' Optional split
    If AskYesNo("Do you want to split exported files into AI-friendly chunks?") Then
        Call SplitTextFiles(exportFolder, 20000)
    End If
End Sub

' ================================================================
' ===== Export with Forms (.frm) =====
' ================================================================
Public Sub ExportModulesWithForms()
    Dim vbProj As Object, vbComp As Object
    Dim exportFolder As String, fileName As String
    
    exportFolder = CreateExportSubfolder(ThisWorkbook.Name)
    If Len(exportFolder) = 0 Then Exit Sub
    
    Set vbProj = ThisWorkbook.VBProject
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1: fileName = exportFolder & "\" & SafeFileName(vbComp.Name) & ".bas"
            Case 2: fileName = exportFolder & "\" & SafeFileName(vbComp.Name) & ".cls"
            Case 3: fileName = exportFolder & "\" & SafeFileName(vbComp.Name) & ".frm"
            Case Else: fileName = exportFolder & "\" & SafeFileName(vbComp.Name) & ".txt"
        End Select
        vbComp.Export fileName
    Next vbComp
    
    MsgBox "Modules + Forms exported to: " & exportFolder, vbInformation
    
    ' Optional zip
    If AskYesNo("Do you want to zip the exported folder?") Then Call ZipExportFolder(exportFolder)
    
    ' Optional split
    If AskYesNo("Do you want to split exported files into AI-friendly chunks?") Then
        Call SplitTextFiles(exportFolder, 20000)
    End If
End Sub

' ================================================================
' ===== Export Binary Contents =====
' ================================================================
Public Sub ExportModulesBinary()
    Dim vbProj As Object, vbComp As Object
    Dim exportFolder As String, fileName As String
    
    exportFolder = CreateExportSubfolder(ThisWorkbook.Name)
    If Len(exportFolder) = 0 Then Exit Sub
    
    Set vbProj = ThisWorkbook.VBProject
    For Each vbComp In vbProj.VBComponents
        fileName = exportFolder & "\" & SafeFileName(vbComp.Name) & ".bin"
        SaveBinary vbComp, fileName
    Next vbComp
    
    MsgBox "Binary contents exported to: " & exportFolder, vbInformation
    
    ' Optional zip
    If AskYesNo("Do you want to zip the exported folder?") Then Call ZipExportFolder(exportFolder)
End Sub

Private Sub SaveBinary(vbComp As Object, fileName As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'Binary
    stream.Open
    vbComp.Export fileName
    stream.LoadFromFile fileName
    stream.SaveToFile fileName, 2
    stream.Close
End Sub

' ================================================================
' ===== Zip Export Folder =====
' ================================================================
Public Sub ZipExportFolder(Optional ByVal folderPath As String)
    Dim shellApp As Object, zipFile As String
    If Len(folderPath) = 0 Then
        folderPath = GetExportFolder()
        If Len(folderPath) = 0 Then Exit Sub
    End If
    
    zipFile = folderPath & "\Export_" & format(Now, "yyyymmdd_HHMMSS") & ".zip"
    Open zipFile For Output As #1: Close #1
    Set shellApp = CreateObject("Shell.Application")
    shellApp.Namespace(zipFile).CopyHere shellApp.Namespace(folderPath).items
    
    MsgBox "Exported files zipped to: " & zipFile, vbInformation
End Sub

' ================================================================
' ===== Split Text Files for AI =====
' ================================================================
Public Sub SplitTextFiles(exportFolder As String, Optional chunkSize As Long = 20000)
    Dim fso As Object, file As Object
    Dim textData As String, fileNum As Long
    Dim pos As Long, partNum As Long
    Dim baseName As String, ext As String, outFile As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each file In fso.GetFolder(exportFolder).Files
        ext = LCase(fso.GetExtensionName(file.path))
        If ext = "txt" Or ext = "bas" Or ext = "cls" Or ext = "frm" Then
            fileNum = FreeFile
            Open file.path For Input As #fileNum
            textData = Input$(LOF(fileNum), fileNum)
            Close #fileNum
            
            If Len(textData) > chunkSize Then
                baseName = fso.GetBaseName(file.path)
                pos = 1
                partNum = 1
                Do While pos <= Len(textData)
                    outFile = exportFolder & "\" & baseName & "_Part" & partNum & ".txt"
                    fileNum = FreeFile
                    Open outFile For Output As #fileNum
                    Print #fileNum, Mid$(textData, pos, chunkSize)
                    Close #fileNum
                    pos = pos + chunkSize
                    partNum = partNum + 1
                Loop
            End If
        End If
    Next
    MsgBox "Large files split into AI-friendly chunks in: " & exportFolder, vbInformation
End Sub

' ================================================================
' ===== Export External Workbook Modules =====
' ================================================================
Public Sub ExportExternalWorkbookModules()
    Dim wbPath As String, wb As Workbook
    Dim vbComp As Object
    Dim exportFolder As String, outFile As String
    Dim fNum As Long
    
    ' Pick external workbook
    wbPath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    If wbPath = "False" Then Exit Sub
    
    Set wb = Workbooks.Open(wbPath, ReadOnly:=True)
    
    ' Create export subfolder using external workbook name
    exportFolder = CreateExportSubfolder(wb.Name)
    If Len(exportFolder) = 0 Then
        wb.Close False
        Exit Sub
    End If
    
    ' Output file
    outFile = exportFolder & "\ConsolidatedModules.txt"
    
    fNum = FreeFile
    Open outFile For Output As #fNum
    
    For Each vbComp In wb.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2
                Print #fNum, "----- START MODULE/CLASS: " & vbComp.Name & " -----"
                Print #fNum, vbComp.CodeModule.lines(1, vbComp.CodeModule.CountOfLines)
                Print #fNum, "----- END MODULE/CLASS: " & vbComp.Name & " -----"
                Print #fNum, ""
            Case 3
                Print #fNum, "----- START FORM: " & vbComp.Name & " -----"
                Print #fNum, "Form export not supported in text. Use .frm export if needed."
                Print #fNum, "----- END FORM: " & vbComp.Name & " -----"
                Print #fNum, ""
        End Select
    Next vbComp
    
    Close #fNum
    wb.Close False
    MsgBox "External workbook modules exported to: " & exportFolder, vbInformation
End Sub

' ================================================================
' ===== Run Subs for Quick Access (Multi-Line) =====
' ================================================================
Public Sub Run_ExportModulesAsText()
    ExportModulesAsText
End Sub

Public Sub Run_ExportModulesWithForms()
    ExportModulesWithForms
End Sub

Public Sub Run_ExportModulesBinary()
    ExportModulesBinary
End Sub

Public Sub Run_ExportExternalWorkbookModules()
    ExportExternalWorkbookModules
End Sub

