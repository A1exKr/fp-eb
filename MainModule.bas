Attribute VB_Name = "MainModule"
'Attribute VB_Name = "MainModule"
' Main Module: MainModule

Option Explicit

Public parsedDataSet As Object 'data from the RFP
Public selectedFolderPath As String 'folder to analyse project sheets from
Public parsedSheetsData As Object 'project sheets parsed data

' =========================
' Deferred UI runner (modal-safe)
' =========================

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
#End If

#If VBA7 Then
    Private gDeferredTimerId As LongPtr
#Else
    Private gDeferredTimerId As Long
#End If

' Schedule a one-shot timer to run deferred work while a modal UserForm is open.
Public Sub ScheduleDeferredSearchFormWork(Optional ByVal delayMs As Long = 50)
    On Error Resume Next
    If delayMs < 1 Then delayMs = 1
    CancelDeferredSearchFormWork
    gDeferredTimerId = SetTimer(0, 0, delayMs, AddressOf DeferredSearchFormTimerProc)
End Sub

Public Sub CancelDeferredSearchFormWork()
    On Error Resume Next
    If gDeferredTimerId <> 0 Then
        KillTimer 0, gDeferredTimerId
        gDeferredTimerId = 0
    End If
End Sub

' Back-compat: legacy code paths may call this directly.
' Runs deferred UI work immediately (without scheduling a timer).
Public Sub RunDeferredSearchFormWork()
    Dim uf As Object
    On Error Resume Next

    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmSearchForm" Then
            uf.ProcessDeferredUIWork
            Exit For
        End If
    Next uf
End Sub

#If VBA7 Then
Public Sub DeferredSearchFormTimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub DeferredSearchFormTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    Dim uf As Object

    On Error Resume Next
    ' One-shot: stop the timer first to avoid re-entrancy.
    CancelDeferredSearchFormWork

    ' Find the already-loaded instance of frmSearchForm and run its deferred work.
    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmSearchForm" Then
            uf.ProcessDeferredUIWork
            Exit For
        End If
    Next uf
End Sub

Public Sub Main()
    ' Create and show the UserForm
    'Call frmFeeProposal.UserForm_Initialize
End Sub

Public Function GetRelativePath(relativePath As String) As String
    GetRelativePath = ThisWorkbook.path & "\" & relativePath
End Function


