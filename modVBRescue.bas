Attribute VB_Name = "modVBRescue"
Option Explicit

' Requires: Trust access to the VBA project object model + VBIDE reference
' Purpose: help recover when Excel shows "repaired/removed content" on save.

Public Sub Rescue_ExportAll(Optional ByVal exportFolder As String = "")
    Dim vbProj As VBIDE.VBProject
    Dim comp As VBIDE.vbComponent
    Dim path As String
    
    If Len(exportFolder) = 0 Then
        exportFolder = ThisWorkbook.path
    End If
    
    If Len(exportFolder) = 0 Then
        MsgBox "Please save the workbook first so it has a folder to export to.", vbExclamation
        Exit Sub
    End If
    
    If Right$(exportFolder, 1) <> "\" Then exportFolder = exportFolder & "\"
    
    Set vbProj = Application.vbe.ActiveVBProject
    
    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                comp.Export exportFolder & comp.name & ".bas"
            Case vbext_ct_Document ' Sheet/Workbook
                ' skip
            Case vbext_ct_MSForm
                comp.Export exportFolder & comp.name & ".frm"
        End Select
    Next comp
    
    MsgBox "Exported modules/forms to: " & exportFolder, vbInformation
End Sub
' Unload every live UserForm instance, even if hidden/modeless or erroring.
Public Sub Rescue_UnloadAllUserForms(Optional ByVal retries As Long = 3)
    Dim i As Long, beforeCount As Long, afterCount As Long, r As Long
    
    On Error Resume Next
    For r = 1 To retries
        beforeCount = VBA.UserForms.Count
        
        ' Walk backwards so indexing stays valid while we unload.
        For i = VBA.UserForms.Count - 1 To 0 Step -1
            VBA.UserForms(i).Hide
            DoEvents
            Unload VBA.UserForms(i)
            DoEvents
        Next i
        
        ' One more pass, in case anything re-instantiated in event code.
        For i = VBA.UserForms.Count - 1 To 0 Step -1
            Unload VBA.UserForms(i)
        Next i
        
        afterCount = VBA.UserForms.Count
        If afterCount = 0 Then Exit For    ' all gone
        If afterCount >= beforeCount Then  ' nothing changed, try harder
            Call Rescue_KillDefaultInstances
        End If
    Next r
    On Error GoTo 0
End Sub

' Some projects keep default instances alive via module/global variables.
' This clears common references so the GC can release them.
Public Sub Rescue_KillDefaultInstances()
    On Error Resume Next
    ' Add any form/global refs you use here:
    ' Example:
    '   Unload frmReviewFinalize: Set frmReviewFinalize = Nothing
    '   Unload frmSearchForm:     Set frmSearchForm = Nothing
    '   Unload frmFeeProposal:    Set frmFeeProposal = Nothing
    
    ' If you do NOT have object variables for forms, this is harmless.
    On Error GoTo 0
End Sub

' Verify there are no live forms (for your sanity)
Public Sub Rescue_DumpLiveForms()
    Dim i As Long
    Debug.Print "Live UserForms: "; VBA.UserForms.Count
    For i = 0 To VBA.UserForms.Count - 1
        Debug.Print "  - "; TypeName(VBA.UserForms(i))
    Next i
End Sub

Public Sub Rescue_RemoveReviewForms()
    Dim vbProj As VBIDE.VBProject: Set vbProj = Application.vbe.ActiveVBProject
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("frmReviewFinalize")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmDiffViewer")
    On Error GoTo 0
    MsgBox "Removed frmReviewFinalize / frmDiffViewer. Save the workbook now.", vbInformation
End Sub

Public Sub Rescue_ReinstallReviewForms()
    On Error Resume Next
    InstallReviewFinalizeForm
    On Error GoTo 0
End Sub

Public Sub Rescue_FullCycle()
    ' 1) Unload any live forms
    Rescue_UnloadAllUserForms
    ' 2) Remove review forms
    Rescue_RemoveReviewForms
    ' 3) Prompt user to save manually
    MsgBox "Now save the workbook (Ctrl+S). After saving successfully, run 'Rescue_ReinstallReviewForms' to rebuild the forms.", vbInformation
End Sub
