Attribute VB_Name = "CodeExport"
'Attribute VB_Name = "CodeExport"
Sub ExportAllVBAComponents()
    Dim vbComponent As vbComponent
    Dim exportPath As String
    Dim gitPath As String
    Dim fileName As String
    Dim dateFolder As String
    Dim folderSuffix As Integer
    Dim folderPath As String
    Dim frxFile As String
    
    ' === ORIGINAL BACKUP EXPORT ===
    ' Set the export base path for backup
    exportPath = GetRelativePath("\code bak\")
    
    ' Ensure the export path ends with a backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Create a base date-stamped folder name in the format YYYYMMDD
    dateFolder = Format(Now(), "YYYYMMDD")
    folderPath = exportPath & dateFolder
    
    ' Check if the folder exists and find a unique folder name with suffix if needed
    folderSuffix = 2
    Do While Dir(folderPath, vbDirectory) <> ""
        folderPath = exportPath & dateFolder & "-" & folderSuffix
        folderSuffix = folderSuffix + 1
    Loop
    
    ' Create the directory
    MkDir folderPath
    
    ' Loop through each component and export to backup
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        Select Case vbComponent.Type
            Case vbext_ct_StdModule
                fileName = folderPath & "\" & vbComponent.name & ".bas"
            Case vbext_ct_ClassModule
                fileName = folderPath & "\" & vbComponent.name & ".cls"
            Case vbext_ct_MSForm
                fileName = folderPath & "\" & vbComponent.name & ".frm"
            Case vbext_ct_Document
                Debug.Print "Skipped: " & vbComponent.name & " (Document type)"
                fileName = ""
            Case Else
                Debug.Print "Skipped: " & vbComponent.name & " (Unrecognized or unsupported type)"
                fileName = ""
        End Select
        
        If fileName <> "" Then
            On Error Resume Next
            vbComponent.Export fileName
            If Err.Number = 0 Then
                Debug.Print "Backup exported: " & vbComponent.name & " to " & fileName
            Else
                Debug.Print "Backup failed: " & vbComponent.name
            End If
            On Error GoTo 0
        End If
    Next vbComponent
    
    ' === PARALLEL GIT EXPORT ===
    ' Set the export path to the Git repo folder
    gitPath = "C:\Users\03669\fp-gen\"
    If Right(gitPath, 1) <> "\" Then gitPath = gitPath & "\"
    
    ' Check if Git folder exists
    If Dir(gitPath, vbDirectory) = "" Then
        MsgBox "Git folder not found: " & gitPath & ". Skipping Git export.", vbExclamation
        GoTo SkipGit
    End If
    
    ' Loop through again and export to Git (direct, no subfolders, exclude .frx)
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        Select Case vbComponent.Type
            Case vbext_ct_StdModule
                fileName = gitPath & vbComponent.name & ".bas"
            Case vbext_ct_ClassModule
                fileName = gitPath & vbComponent.name & ".cls"
            Case vbext_ct_MSForm
                fileName = gitPath & vbComponent.name & ".frm"
            Case vbext_ct_Document
                fileName = ""
            Case Else
                fileName = ""
        End Select
        
        If fileName <> "" Then
            On Error Resume Next
            vbComponent.Export fileName
            If Err.Number = 0 Then
                Debug.Print "Git exported: " & vbComponent.name & " to " & fileName
                ' Exclude .frx for Git
                If vbComponent.Type = vbext_ct_MSForm Then
                    frxFile = gitPath & vbComponent.name & ".frx"
                    If Dir(frxFile) <> "" Then
                        Kill frxFile
                        Debug.Print "Deleted .frx for Git: " & frxFile
                    End If
                End If
            Else
                Debug.Print "Git export failed: " & vbComponent.name
            End If
            On Error GoTo 0
        End If
    Next vbComponent
    
SkipGit:
    MsgBox "Backup and Git export complete!"
End Sub
