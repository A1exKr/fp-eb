Attribute VB_Name = "modReviewBridge"
'Attribute VB_Name = "modReviewBridge"
Option Explicit

Public Sub OpenReviewWithJson(ByVal rfpJson As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== OpenReviewWithJson START ==="
    Debug.Print "JSON length: " & Len(rfpJson)
    
    ' Map to canonical model
    Dim review As Object
    Set review = modRfpMapper.MapRfpJsonToReviewModel(rfpJson)
    Debug.Print "Mapped to review model"
    
    ' Publish to bus
    modDataBus.PublishRfpData review
    Debug.Print "Published to bus"
    
    ' HIDE (not Unload) the Fee Proposal form to allow modeless Review form
    On Error Resume Next
    frmFeeProposal.Hide
    Debug.Print "Hidden frmFeeProposal (still in memory)"
    On Error GoTo ErrorHandler
    
    ' Show the Review form
    Debug.Print "Creating Review form..."
    Dim frmReview As Object
    Set frmReview = VBA.UserForms.Add("frmReviewFinalize")
    
    Debug.Print "Showing Review form..."
    frmReview.Show vbModeless
    
    Debug.Print "=== OpenReviewWithJson END ==="
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR: " & Err.Description
    MsgBox "Failed to open Review & Finalize: " & Err.Description, vbCritical
End Sub


' Diagnostic: List all VBA components in the project
Public Sub DiagnoseVBAComponents()
    On Error Resume Next
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    If vbProj Is Nothing Then
        Debug.Print "ERROR: Cannot access VBProject. Check Trust Center settings:"
        Debug.Print "  File > Options > Trust Center > Trust Center Settings"
        Debug.Print "  > Macro Settings > Trust access to the VBA project object model"
        Exit Sub
    End If
    
    Debug.Print "=== VBA Components in Project ==="
    Debug.Print "Project Name: " & vbProj.name
    Debug.Print "Component Count: " & vbProj.VBComponents.Count
    Debug.Print ""
    
    Dim comp As Object
    Dim i As Long
    For i = 1 To vbProj.VBComponents.Count
        Set comp = vbProj.VBComponents(i)
        Debug.Print i & ". Name: " & comp.name & " | Type: " & GetComponentTypeName(comp.Type)
    Next i
    
    Debug.Print ""
    Debug.Print "=== Checking frmReviewFinalize specifically ==="
    
    Set comp = Nothing
    Set comp = vbProj.VBComponents("frmReviewFinalize")
    
    If comp Is Nothing Then
        Debug.Print "frmReviewFinalize: NOT FOUND"
    Else
        Debug.Print "frmReviewFinalize: FOUND"
        Debug.Print "  Type: " & GetComponentTypeName(comp.Type)
        Debug.Print "  Type Number: " & comp.Type
    End If
    
    On Error GoTo 0
End Sub

Private Function GetComponentTypeName(componentType As Long) As String
    Select Case componentType
        Case 1: GetComponentTypeName = "Module"
        Case 2: GetComponentTypeName = "Class"
        Case 3: GetComponentTypeName = "MSForm"
        Case 100: GetComponentTypeName = "Document"
        Case Else: GetComponentTypeName = "Unknown (" & componentType & ")"
    End Select
End Function

' Add this to modReviewBridge.bas for detailed debugging
Public Sub TestFormInstantiation()
    On Error Resume Next
    
    Debug.Print "=== Testing Form Instantiation ==="
    
    ' Test 1: Check if form component exists
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents("frmReviewFinalize")
    Debug.Print "1. Component exists: " & (Not comp Is Nothing)
    If Not comp Is Nothing Then
        Debug.Print "   Component Type: " & comp.Type
        Debug.Print "   Component Name: " & comp.name
    End If
    
    ' Test 2: Try to add via VBA.UserForms
    Debug.Print ""
    Debug.Print "2. Attempting VBA.UserForms.Add..."
    Err.Clear
    Dim frm1 As Object
    Set frm1 = VBA.UserForms.Add("frmReviewFinalize")
    If Err.Number <> 0 Then
        Debug.Print "   ERROR: " & Err.Number & " - " & Err.Description
        Debug.Print "   Source: " & Err.Source
    Else
        Debug.Print "   SUCCESS: Form instance created"
        If Not frm1 Is Nothing Then
            Debug.Print "   TypeName: " & TypeName(frm1)
            Unload frm1
        End If
    End If
    
    ' Test 3: Try alternative method with Application.VBE
    Debug.Print ""
    Debug.Print "3. Checking VBE access..."
    Err.Clear
    Dim vbe As Object
    Set vbe = Application.vbe
    If Err.Number <> 0 Then
        Debug.Print "   ERROR: Cannot access VBE - " & Err.Description
    Else
        Debug.Print "   VBE accessible: Yes"
    End If
    
    ' Test 4: Check form's code module
    Debug.Print ""
    Debug.Print "4. Checking form's code module..."
    If Not comp Is Nothing Then
        Err.Clear
        Dim codeLines As Long
        codeLines = comp.CodeModule.CountOfLines
        If Err.Number <> 0 Then
            Debug.Print "   ERROR reading code: " & Err.Description
        Else
            Debug.Print "   Code lines: " & codeLines
            
            ' Show first few lines
            If codeLines > 0 Then
                Debug.Print "   First line: " & comp.CodeModule.lines(1, 1)
            End If
        End If
    End If
    
    ' Test 5: Try creating instance using CreateObject style
    Debug.Print ""
    Debug.Print "5. Alternative instantiation methods..."
    Err.Clear
    Dim frm2 As Object
    ' This probably won't work but let's try
    Set frm2 = CreateObject("Forms.frmReviewFinalize.1")
    If Err.Number <> 0 Then
        Debug.Print "   CreateObject method: FAILED (" & Err.Number & ")"
    Else
        Debug.Print "   CreateObject method: SUCCESS"
    End If
    
    ' Test 6: Check if there are any existing instances
    Debug.Print ""
    Debug.Print "6. Checking existing UserForm instances..."
    Debug.Print "   UserForms.Count: " & VBA.UserForms.Count
    Dim i As Long
    For i = 0 To VBA.UserForms.Count - 1
        Debug.Print "   " & i & ". " & TypeName(VBA.UserForms(i))
    Next i
    
    On Error GoTo 0
    Debug.Print ""
    Debug.Print "=== Test Complete ==="
End Sub
