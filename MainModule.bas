Attribute VB_Name = "MainModule"
'Attribute VB_Name = "MainModule"
' Main Module: MainModule

Public parsedDataSet As Object 'data from the RFP
Public selectedFolderPath As String 'folder to analyse project sheets from
Public parsedSheetsData As Object 'project sheets parsed data

Public Sub Main()
    ' Create and show the UserForm
    'Call frmFeeProposal.UserForm_Initialize
End Sub

Public Function GetRelativePath(relativePath As String) As String
    GetRelativePath = ThisWorkbook.path & "\" & relativePath
End Function


