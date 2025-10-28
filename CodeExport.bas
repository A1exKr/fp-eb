Attribute VB_Name = "CodeExport"
'Attribute VB_Name = "CodeExport"
Sub ExportAllVBAComponents()
    Dim vbComponent As vbComponent
    Dim exportPath As String
    Dim fileName As String
    Dim dateFolder As String
    Dim folderSuffix As Integer
    Dim folderPath As String
    
    ' Set the export base path
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
    
    ' Loop through each component in the VBA project
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        Select Case vbComponent.Type
            Case vbext_ct_StdModule
                fileName = folderPath & "\" & vbComponent.name & ".bas"
            Case vbext_ct_ClassModule
                fileName = folderPath & "\" & vbComponent.name & ".cls"
            Case vbext_ct_MSForm
                fileName = folderPath & "\" & vbComponent.name & ".frm"
            Case vbext_ct_Document
                ' Skip ThisWorkbook and sheet modules
                Debug.Print "Skipped: " & vbComponent.name & " (Document type)"
                fileName = ""
            Case Else
                ' Skip other unrecognized types
                Debug.Print "Skipped: " & vbComponent.name & " (Unrecognized or unsupported type)"
                fileName = ""
        End Select
        
        ' If fileName is not empty, export the component
        If fileName <> "" Then
            On Error Resume Next
            vbComponent.Export fileName
            If Err.Number = 0 Then
                Debug.Print "Exported: " & vbComponent.name & " to " & fileName
            Else
                Debug.Print "Failed to export: " & vbComponent.name
            End If
            On Error GoTo 0
        End If
    Next vbComponent
    
    MsgBox "Export complete!"
End Sub




