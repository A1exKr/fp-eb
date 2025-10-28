Attribute VB_Name = "modReviewBridge"
Option Explicit

' Single public entry point: called by frmFeeProposal when AI parsing is done.
' Minimal change needed in existing code:
'     modReviewBridge.OpenReviewWithJson rfpAnalysisJson

Public Sub OpenReviewWithJson(ByVal rfpJson As String)
    Debug.Print "=== OpenReviewWithJson START ==="
    Debug.Print "JSON length: " & Len(rfpJson)
    
    ' Put raw JSON on the bus
    modDataBus.PublishRfpJson rfpJson
    Debug.Print "Published JSON to bus"

    ' Map to canonical model and publish
    Dim review As Object
    On Error GoTo MapError
    Set review = modRfpMapper.MapRfpJsonToReviewModel(rfpJson)
    Debug.Print "Mapped to review model, keys count: " & review.keys.Count
    On Error GoTo 0
    
    modDataBus.PublishRfpData review
    Debug.Print "Published review data"

    ' Hand-off to Review form
    On Error Resume Next
    Set gReview = review
    Debug.Print "Set gReview global"
    
    ' Install form if needed
    Debug.Print "Calling InstallReviewFinalizeForm..."
    InstallReviewFinalizeForm
    If Err.Number <> 0 Then
        Debug.Print "ERROR in InstallReviewFinalizeForm: " & Err.Description
        MsgBox "Failed to install form: " & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
    Debug.Print "Form installed OK"
    
    ' Show the form
    Debug.Print "Calling ShowReviewFinalize..."
    ShowReviewFinalize
    If Err.Number <> 0 Then
        Debug.Print "ERROR showing form: " & Err.Description
        MsgBox "Failed to show form: " & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
    Debug.Print "Form shown OK"
    
    On Error GoTo 0
    Debug.Print "=== OpenReviewWithJson END ==="
    Exit Sub

MapError:
    Debug.Print "ERROR mapping JSON: " & Err.Description
    MsgBox "Failed to map RFP data: " & Err.Description, vbCritical
End Sub
