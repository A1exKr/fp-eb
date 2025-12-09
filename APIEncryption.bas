Attribute VB_Name = "APIEncryption"
' Module for API Key Encryption/Decryption
Option Explicit

' Encrypt a string using XOR (basic encryption for demonstration purposes)
Function Encrypt(key As String, data As String) As String
    Dim i As Integer
    Dim result As String
    result = ""
    
    For i = 1 To Len(data)
        result = result & Chr(Asc(Mid(data, i, 1)) Xor Asc(Mid(key, ((i - 1) Mod Len(key)) + 1)))
    Next i

    Encrypt = result
End Function

' Decrypt a string using XOR (symmetric with Encrypt function)
Function Decrypt(key As String, encryptedData As String) As String
    Decrypt = Encrypt(key, encryptedData)
End Function

' Open a dialog to select a TXT file with the API key, then create an encrypted file
Sub CreateEncryptedAPIKeyFile()
    Dim fd As fileDialog
    Dim txtFilePath As String
    Dim folderDialog As fileDialog
    Dim folderPath As String
    Dim apiKey As String
    Dim encryptionKey As String
    Dim encryptedData As String
    Dim fso As Object
    Dim file As Object
    
    ' Step 1: Ask user to select the TXT file containing the API key
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    fd.title = "Select the TXT file containing the API key"
    fd.Filters.Add "Text Files", "*.txt", 1
    
    If fd.Show = -1 Then
        txtFilePath = fd.selectedItems(1)
    Else
        MsgBox "No file selected. Operation canceled.", vbExclamation
        Exit Sub
    End If
    
    ' Step 2: Read the API key from the selected file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(txtFilePath, ForReading)
    apiKey = Trim(file.ReadAll) ' Trim to remove extra spaces/newlines
    file.Close
    
    ' Step 3: Ask user to enter an encryption key
    encryptionKey = InputBox("Enter an encryption key to secure your API key:", "Encryption Key")
    If encryptionKey = "" Then
        MsgBox "No encryption key entered. Operation canceled.", vbExclamation
        Exit Sub
    End If
    
    ' Step 4: Encrypt the API key
    encryptedData = Encrypt(encryptionKey, apiKey)
    
    ' Step 5: Save the encrypted data to a new file in the same folder as the workbook
    Dim encryptedFilePath As String
    encryptedFilePath = ThisWorkbook.Path & "\apikey.dat"
    
    Set file = fso.CreateTextFile(encryptedFilePath, True)
    file.Write encryptedData
    file.Close
    
    ' Step 7: Notify user of success
    MsgBox "Encrypted API key file saved to: " & encryptedFilePath, vbInformation
End Sub

