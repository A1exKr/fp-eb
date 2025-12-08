Attribute VB_Name = "modDataBus"
Option Explicit

' Simple in-memory handoff bus between forms/modules.
' Carries the raw AI JSON string and the mapped review model (Dictionary).

Private mRfpJson As String
Private mRfpData As Object ' Scripting.Dictionary

Public Sub PublishRfpJson(ByVal json As String)
    mRfpJson = json
    Set mRfpData = Nothing
End Sub

Public Function ConsumeRfpJson() As String
    ConsumeRfpJson = mRfpJson
    mRfpJson = ""
End Function

Public Sub PublishRfpData(ByVal d As Object)
    Set mRfpData = d
End Sub

Public Function ConsumeRfpData() As Object
    Set ConsumeRfpData = mRfpData
    Set mRfpData = Nothing
End Function

Public Sub ClearBus()
    mRfpJson = ""
    Set mRfpData = Nothing
End Sub
