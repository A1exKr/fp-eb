Attribute VB_Name = "APIModule"
Option Explicit

'Attribute VB_Name = "APIModule"
Public Function CallAPIsyn(fileContent As String, promptFilePath As String) As String
    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim promptTxt As String
    Dim payload As String
    Dim requestText As String
    Dim responseText As String
    Dim apiKeyFilePath As String
    Dim encryptionKeyFilePath As String
    Dim encryptionKey As String

    On Error GoTo ErrorHandler

    ' Path to the encrypted API key file (relative path)
    apiKeyFilePath = GetRelativePath("apikey.dat")
    
    encryptionKeyFilePath = GetRelativePath("encryptionkey.dat")
    
    ' Debug message for API key file path
    Debug.Print "CallAPIsyn.apiKeyFilePath: " & apiKeyFilePath
    Debug.Print "CallAPIsyn.encryptionKeyFilePath: " & encryptionKeyFilePath

    ' Decrypt and read the API key
    encryptionKey = ReadFileWithEncoding(encryptionKeyFilePath, "UTF-8")
    apiKey = GetAPIKey(encryptionKey, apiKeyFilePath)
    If Len(apiKey) = 0 Then
        MsgBox "Error: API key file is empty or could not be read.", vbCritical
        Exit Function
    End If

    ' Validate the input file content
    If Len(fileContent) = 0 Then
        MsgBox "Error: Input text is empty or could not be read.", vbCritical
        Exit Function
    End If
    
    ' Read the prompt text from the file
    promptTxt = ReadFileWithEncoding(promptFilePath, "UTF-8")
    Debug.Print "CallAPIsyn.promptTxt Length: " & Len(promptTxt)
    If Len(promptTxt) = 0 Then
        MsgBox "Error: Prompt text file is empty or could not be read.", vbCritical
        Exit Function
    End If
    
    Debug.Print "fileContent Length: " & Len(fileContent)

    ' Compile request from the initial prompt and PDF retrieved text
    requestText = promptTxt & "File Path: " & fileContent
    Debug.Print "requestText Length: " & Len(requestText)
'Debug.Print "fileContent: " & fileContent
    
    ' Initialize HTTP object and set the request details
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://api.openai.com/v1/chat/completions"
    
    Debug.Print "Payload Length: " & Len(payload)

    ' Create valid JSON payload
    'payload = "{""model"": ""gpt-4o-mini-2024-07-18"", ""messages"": [{""role"": ""user"", ""content"": """ & EscapeJsonString(requestText) & """}], ""max_tokens"": 16000}"
    'payload = "{""model"": ""gpt-5-nano"", ""messages"": [{""role"": ""user"", ""content"": """ & EscapeJsonString(requestText) & """}], ""max_tokens"": 16000}"
    payload = "{""model"": ""gpt-4.1-nano"", ""messages"": [{""role"": ""user"", ""content"": """ & EscapeJsonString(requestText) & """}], ""max_tokens"": 16000}"

    Debug.Print "Payload: " & payload

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send payload

    responseText = http.responseText

    CallAPIsyn = responseText
    Exit Function

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    CallAPIsyn = ""
End Function

Public Function ReadFileWithEncoding(filePath As String, encoding As String) As String
    Const CHUNK_SIZE As Long = 1048576 ' 1 MB chunks
    
    Dim stream As Object
    Dim chunk As String
    Dim result As String
    
    ' Create ADODB Stream object for reading the file
    Set stream = CreateObject("ADODB.Stream")
    
    ' Configure the stream
    With stream
        .Type = 2 ' Text data
        .Charset = encoding ' Specify the encoding (e.g., "UTF-8", "ISO-8859-1")
        .Open
        .LoadFromFile filePath
    End With
    
    ' Read file in chunks
    Do Until stream.EOS
        chunk = stream.ReadText(CHUNK_SIZE)
        result = result & chunk
    Loop
    
    ' Close the stream
    stream.Close
    Set stream = Nothing
Debug.Print "%%% filePath: " & filePath
Debug.Print "%%% encoding: " & encoding
'Debug.Print "%%% result: " & result
    ' Return the read file contents
    ReadFileWithEncoding = result
End Function



' Function to escape JSON special characters
Public Function EscapeJsonString(ByVal str As String) As String
    Dim i As Long
    Dim char As String
    Dim asciiVal As Long
    Dim escapedStr As String

    escapedStr = ""
    Debug.Print ("Len(str): " & Len(str))
    
    ' Loop through each character of the string
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        asciiVal = AscW(char) ' Get ASCII/Unicode value of the character

        ' Escape special JSON characters
        Select Case char
            Case "\": escapedStr = escapedStr & "\\"
            Case """": escapedStr = escapedStr & "\"""
            Case "/": escapedStr = escapedStr & "\/"
            Case vbBack: escapedStr = escapedStr & "\b"
            Case vbFormFeed: ' Remove form feed characters entirely
            Case vbNewLine, vbCr
                ' Replace newlines or carriage returns with \n
                If Len(escapedStr) > 0 And Right$(escapedStr, 1) <> "\n" Then
                    escapedStr = escapedStr & "\n"
                End If
            Case vbTab: escapedStr = escapedStr & "\t"
            Case Else
                ' Keep all printable characters including Unicode
                If asciiVal >= 32 Then  ' Only remove control characters (ASCII < 32)
                    escapedStr = escapedStr & char
                    'Debug.Print "char: " & char
                    'Debug.Print "asciiVal: " & asciiVal
                    'Debug.Print "escapedStr: " & escapedStr
                End If
        End Select
    Next i

    ' Remove any final trailing newline escape sequence
    If Right$(escapedStr, 2) = "\n" Then
        escapedStr = Left$(escapedStr, Len(escapedStr) - 2)
    End If
    
    EscapeJsonString = Trim(escapedStr)
End Function


Function GetAPIKey(key As String, filePath As String) As String
    Dim encryptedKey As String
    Dim decryptedKey As String
    Dim fso As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.fileExists(filePath) Then
        MsgBox "API key file not found.", vbCritical
        Exit Function
    End If

    Set file = fso.OpenTextFile(filePath, ForReading)
    encryptedKey = file.ReadAll
    file.Close

    decryptedKey = Decrypt(key, encryptedKey)
    GetAPIKey = decryptedKey
End Function

