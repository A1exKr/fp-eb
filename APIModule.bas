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
    encryptionKey = Trim(ReadFileWithEncoding(encryptionKeyFilePath, "UTF-8"))
    Debug.Print "Encryption key length: " & Len(encryptionKey)
    
    If Len(encryptionKey) = 0 Then
        MsgBox "Error: Encryption key file is empty." & vbCrLf & _
               "File: " & encryptionKeyFilePath, vbCritical, "Encryption Key Error"
        Exit Function
    End If
    
    apiKey = Trim(GetAPIKey(encryptionKey, apiKeyFilePath))
    Debug.Print "Decrypted API key length: " & Len(apiKey)
    
    If Len(apiKey) = 0 Then
        MsgBox "Error: API key file is empty or could not be read." & vbCrLf & _
               "Encryption key file: " & encryptionKeyFilePath & vbCrLf & _
               "API key file: " & apiKeyFilePath, vbCritical, "API Key Error"
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
        MsgBox "Error: Prompt text file is empty or could not be read." & vbCrLf & _
               "Prompt file: " & promptFilePath, vbCritical, "Prompt File Error"
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
    
    ' Build separate system + user messages to reduce ambiguity and enforce strict JSON output
    Dim systemMsg As String
    Dim userMsg As String
    systemMsg = "You are a strict JSON extractor. Return only valid JSON as specified in the prompt file. Use ""n/a"" for missing fields. Do not include any explanation or markdown."
    userMsg = requestText

    ' Create valid JSON payload (use a deterministic temperature for extraction tasks)
    ' Use the model's supported parameter name for maximum completion tokens
    payload = "{""model"": ""gpt-5.1-2025-11-13"", ""messages"": [{""role"": ""system"", ""content"": """ & EscapeJsonString(systemMsg) & """}, {""role"": ""user"", ""content"": """ & EscapeJsonString(userMsg) & """}], ""max_completion_tokens"": 16000, ""temperature"": 0}"

    Debug.Print "Payload Length: " & Len(payload)
    Debug.Print "Payload: " & Left(payload, 1000)

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    
    Debug.Print "Sending HTTP request to: " & url
    http.send payload

    Debug.Print "HTTP Status: " & http.status
    Debug.Print "HTTP Response length: " & Len(http.responseText)
    
    If http.status <> 200 Then
        MsgBox "API Error: HTTP " & http.status & vbCrLf & _
               "Response: " & Left(http.responseText, 500), vbCritical, "API Call Failed"
        CallAPIsyn = ""
        Exit Function
    End If

    responseText = http.responseText
    
    ' Debug: Save raw API response to file for inspection
    Dim debugFilePath As String
    debugFilePath = GetRelativePath("\test output\api_response_debug.txt")
    Dim debugFile As Integer
    debugFile = FreeFile
    On Error Resume Next
    Open debugFilePath For Output As debugFile
    Print #debugFile, "=== Raw API Response ==="
    Print #debugFile, "Timestamp: " & Now
    Print #debugFile, "Response length: " & Len(responseText)
    Print #debugFile, "First 2000 chars:"
    Print #debugFile, Left(responseText, 2000)
    Print #debugFile, ""
    Print #debugFile, "Last 1000 chars:"
    Print #debugFile, Right(responseText, 1000)
    Close debugFile
    On Error GoTo 0
    Debug.Print "API response saved to: " & debugFilePath

    CallAPIsyn = responseText
    Exit Function

ErrorHandler:
    Debug.Print "ERROR in CallAPIsyn: " & Err.Number & " - " & Err.Description
    MsgBox "Error calling API: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Check:" & vbCrLf & _
           "- API key files exist and are readable" & vbCrLf & _
           "- Encryption key is correct" & vbCrLf & _
           "- Network connection is available" & vbCrLf & _
           "- Model name is valid", vbCritical, "API Error"
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
        MsgBox "API key file not found: " & filePath, vbCritical
        GetAPIKey = ""
        Exit Function
    End If

    Set file = fso.OpenTextFile(filePath, ForReading)
    encryptedKey = file.ReadAll
    file.Close
    
    Debug.Print "Encrypted key length: " & Len(encryptedKey)

    decryptedKey = Decrypt(key, encryptedKey)
    Debug.Print "After decryption length: " & Len(decryptedKey)
    
    GetAPIKey = decryptedKey
End Function



