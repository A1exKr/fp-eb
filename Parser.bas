Attribute VB_Name = "Parser"

Option Explicit

' -------------------------
' Module: Parser.bas
' -------------------------


'===========================================================
' Convert parsed dictionary to JSON for Review & Finalize
'===========================================================
Public Function ExportParsedRFPToJson(ByVal parsedData As Object) As String
    ' Convert the flattened dictionary back to nested JSON structure
    Dim json As String
    Dim k As Variant
    
    ' Start building JSON manually to preserve structure
    json = "{"
    
    ' Project block
    json = json & """project"":{"
    json = json & """name"":""" & CleanJsonValue(parsedData, "project.name") & ""","
    json = json & """location"":""" & CleanJsonValue(parsedData, "project.location") & ""","
    json = json & """type"":""" & CleanJsonValue(parsedData, "project.type") & ""","
    json = json & """siteArea"":""" & CleanJsonValue(parsedData, "project.siteArea") & """"
    json = json & "},"
    
    ' Client block
    json = json & """client"":{"
    json = json & """name"":""" & CleanJsonValue(parsedData, "client.name") & """"
    json = json & "},"
    
    ' Understanding block
    json = json & """understanding"":{"
    json = json & """understanding"":""" & CleanJsonValue(parsedData, "understanding.understanding") & """"
    json = json & "},"
    
    ' Methodology block
    json = json & """methodology"":{"
    json = json & """text"":""" & CleanJsonValue(parsedData, "methodology.text") & """"
    json = json & "},"
    
    ' Scope block
    json = json & """scope"":{"
    json = json & """scopeList"":" & ArrayToJson(parsedData, "scope.scopeList") & ","
    json = json & """deliverablesList"":" & ArrayToJson(parsedData, "scope.deliverablesList")
    json = json & "},"
    
    ' Schedule block
    json = json & """schedule"":{"
    json = json & """totalWeeks"":" & Nz(parsedData, "schedule.totalWeeks", "0") & ","
    json = json & """milestones"":" & ArrayToJson(parsedData, "schedule.milestones")
    json = json & "},"
    
    ' Team block
    json = json & """team"":{"
    json = json & """principal"":{""name"":""" & CleanJsonValue(parsedData, "team.principal.name") & """,""title"":""" & CleanJsonValue(parsedData, "team.principal.title") & """},"
    json = json & """pm"":{""name"":""" & CleanJsonValue(parsedData, "team.pm.name") & """,""title"":""" & CleanJsonValue(parsedData, "team.pm.title") & """}"
    json = json & "},"
    
    ' Fee block
    json = json & """fee"":{"
    json = json & """currency"":""" & CleanJsonValue(parsedData, "fee.currency") & ""","
    json = json & """rates"":{},"
    json = json & """effortByPhase"":{},"
    json = json & """overheadPct"":" & Nz(parsedData, "fee.overheadPct", "0") & ","
    json = json & """travel"":{},"
    json = json & """subconsultants"":[]"
    json = json & "},"
    
    ' Experience block (handle indexed items)
    json = json & """experience"":["
    Dim expJson As String, i As Long
    i = 1
    Do While parsedData.Exists("experience." & i & ".name")
        If i > 1 Then expJson = expJson & ","
        expJson = expJson & "{""name"":""" & CleanJsonValue(parsedData, "experience." & i & ".name") & ""","
        expJson = expJson & """location"":""" & CleanJsonValue(parsedData, "experience." & i & ".location") & ""","
        expJson = expJson & """summary"":""" & CleanJsonValue(parsedData, "experience." & i & ".summary") & """}"
        i = i + 1
    Loop
    json = json & expJson & "],"
    
    ' Assumptions block
    json = json & """assumptions"":{"
    json = json & """defaultText"":""" & CleanJsonValue(parsedData, "assumptions.defaultText") & """"
    json = json & "}"
    
    json = json & "}"
    
    ExportParsedRFPToJson = json
End Function

' Helper to convert multi-line string to JSON array
Private Function ArrayToJson(dict As Object, key As String) As String
    On Error Resume Next
    Dim val As String, items() As String
    val = dict(key)
    
    If Len(val) = 0 Then
        ArrayToJson = "[]"
        Exit Function
    End If
    
    ' Split by line breaks
    items = Split(val, vbCrLf)
    
    Dim i As Long, json As String
    json = "["
    For i = LBound(items) To UBound(items)
        If Len(Trim(items(i))) > 0 Then
            If i > LBound(items) Then json = json & ","
            json = json & """" & Replace(Trim(items(i)), """", "\""") & """"
        End If
    Next
    json = json & "]"
    
    ArrayToJson = json
End Function

' Helper to safely get value or default
Private Function Nz(dict As Object, key As String, Optional defaultVal As String = "") As String
    On Error Resume Next
    If dict.Exists(key) Then
        Nz = dict(key)
        If Len(Nz) = 0 Then Nz = defaultVal
    Else
        Nz = defaultVal
    End If
End Function

Private Function CleanJsonValue(ByVal dict As Object, ByVal key As String) As String
    On Error Resume Next
    If dict.Exists(key) Then
        CleanJsonValue = Replace(CStr(dict(key)), """", "\""")
    Else
        CleanJsonValue = ""
    End If
End Function



'===========================================================
' Direct bridge: parse text and open Review & Finalize screen
'===========================================================
Public Sub ParseAndOpenReview(ByVal rfpText As String)
    Dim chatGPTResponse As String
    Dim promptFilePath As String
    Dim rfpJson As String
    
    On Error GoTo ErrHandler
    
    Debug.Print "=== ParseAndOpenReview START ==="
    
    ' Path to the file containing prompt text
    promptFilePath = GetRelativePath("\prompts\Prompt_all-parse.txt")

    ' Making an API call to ChatGPT
    Debug.Print "Calling ChatGPT API..."
    chatGPTResponse = APIModule.CallAPIsyn(rfpText, promptFilePath)
    Debug.Print "Got API response"

    ' Extract the inner JSON from ChatGPT's response
    rfpJson = ExtractInnerJson(chatGPTResponse)
    Debug.Print "Extracted inner JSON, length: " & Len(rfpJson)
    
    If Len(rfpJson) = 0 Then
        MsgBox "Failed to extract JSON from API response.", vbExclamation
        Exit Sub
    End If
    
    ' Hand off to Review & Finalize
    Debug.Print "Opening Review & Finalize..."
    modReviewBridge.OpenReviewWithJson rfpJson
    
    Debug.Print "=== ParseAndOpenReview END ==="
    Exit Sub

ErrHandler:
    MsgBox "Error in ParseAndOpenReview: " & Err.Description, vbCritical
End Sub

' New helper function to extract the inner JSON from ChatGPT response
Private Function ExtractInnerJson(ByVal gptResponse As String) As String
    On Error GoTo ErrHandler
    
    Dim jsonResponse As Object
    Dim contentText As String
    
    ' Remove markdown formatting
    gptResponse = RemoveMarkdownFormatting(gptResponse)
    
    ' Parse the outer JSON response
    Set jsonResponse = JsonConverter.ParseJSON(gptResponse)
    
    ' Extract the 'content' field (the actual project JSON)
    contentText = jsonResponse("choices")(1)("message")("content")
    
    ExtractInnerJson = contentText
    Exit Function
    
ErrHandler:
    Debug.Print "Error extracting inner JSON: " & Err.Description
    ExtractInnerJson = ""
End Function


Public Function ParseRFPText(rfpText As String) As Object
    Dim chatGPTResponse As String
    Dim promptFilePath As String
    Dim parsedData As Object
    
    ' Initialize the dictionary to hold parsed data
    Set parsedData = CreateObject("Scripting.Dictionary")
    
    ' Path to the file containing prompt text
    promptFilePath = GetRelativePath("\prompts\Prompt_all-parse.txt")

    ' Making an API call to ChatGPT
    chatGPTResponse = APIModule.CallAPIsyn(rfpText, promptFilePath)

    'Debug.Print "chatGPTResponse: " & chatGPTResponse
    
    ' Dynamically extract items from the GPT response
    Call DynamicExtractItems(chatGPTResponse, parsedData)
Debug.Print "Cleaned GPT Response: " & chatGPTResponse
    
    ' Return the parsed data
    Set ParseRFPText = parsedData
End Function

' Subroutine to dynamically extract items from the GPT response
Sub DynamicExtractItems(gptResponse As String, ByRef parsedData As Object)
    On Error GoTo ErrorHandler
    Dim jsonResponse As Object
    Dim contentText As String
    Dim innerJson As Object

    ' Remove surrounding markdown formatting (```json\n...\n```)
    gptResponse = RemoveMarkdownFormatting(gptResponse)
Debug.Print "gptResponse = RemoveMarkdownFormatting(gptResponse)"

    ' Parse the JSON response using JsonConverter
    Set jsonResponse = JsonConverter.ParseJSON(gptResponse)
Debug.Print "Set jsonResponse"

    ' Extract the 'content' field
    contentText = jsonResponse("choices")(1)("message")("content")
Debug.Print "contentText " & contentText
    ' Parse the inner JSON content
    Set innerJson = JsonConverter.ParseJSON(contentText)
Debug.Print "5 innerJson "

    ' Initialize parsedData if not already initialized
    If parsedData Is Nothing Then
        Set parsedData = CreateObject("Scripting.Dictionary")
    End If
Debug.Print "6  "

    ' Process the innerJson dictionary
    ProcessDictionary innerJson, parsedData
    
    Debug.Print "=== Parsed Data Keys ==="
Dim debugKey As Variant
For Each debugKey In parsedData.keys
    Debug.Print debugKey & " = " & Left(parsedData(debugKey), 50)
Next

Debug.Print "7  "

    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    MsgBox errMsg, vbCritical, "Runtime Error"
End Sub

Private Function BuildNestedString__(ByVal dict As Variant, Optional indent As String = "") As String
    Dim result As String
    Dim key As Variant
    Dim value As Variant

    For Each key In dict.keys
        value = dict(key)
        result = result & indent & key & ": " & CStr(value) & vbCrLf
    Next key

    BuildNestedString__ = result
End Function



Sub SheetsDynamicExtractItems(gptResponse As String, ByRef parsedData As Object)
    On Error GoTo ErrorHandler
    Dim jsonResponse As Object
    Dim contentText As String
    Dim innerJson As Object
    Dim projectItem As Variant
    Dim attributeKey As Variant
    Dim childDict As Object
    Dim ProjectName As Variant
    Dim projectIndex As Integer
    Dim jsonString As String
    Dim fileContent As String
    Dim cacheFilePath As String

    cacheFilePath = GetRelativePath("\test output\cache.json")

    ' Step 1: Remove surrounding markdown formatting (```json\n...\n```)
    gptResponse = RemoveMarkdownFormatting(gptResponse)

    ' Step 2: Parse the JSON response using JsonConverter
    Set jsonResponse = JsonConverter.ParseJSON(gptResponse)
    'Debug.Print "00 gptResponse = " & gptResponse

    ' Step 3: Extract the 'content' field containing the inner JSON
    contentText = jsonResponse("choices")(1)("message")("content")
    'Debug.Print "0 contentText = " & contentText
    contentText = RemoveBSEscape(contentText)
    'Debug.Print "1 contentText = " & contentText

    ' Step 4: Parse the inner JSON content into a dictionary
    Set innerJson = JsonConverter.ParseJSON(contentText)

    ' Step 5: Prepare to update cache file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Read existing content if the cache file exists and is not empty
    If fso.fileExists(cacheFilePath) Then
        fileContent = ReadTextFile(cacheFilePath)
    Else
        fileContent = ""
    End If

    ' Initialize or modify the JSON array in the cache file
    If Trim(fileContent) = "" Then
        ' If the cache file is empty, start a new array
        'fileContent = "[" & vbCrLf & contentText & vbCrLf & "]"
        fileContent = contentText & vbCrLf
    Else
        ' If the cache file has content, remove the final bracket and append the new content with a comma
        'fileContent = Left(Trim(fileContent), Len(Trim(fileContent)) - 1) & "," & vbCrLf & contentText & vbCrLf & "]"
        ' Step 1: Locate and extract JSON objects within square brackets in contentText
        Dim startPos As Integer, endPos As Integer
        
        startPos = InStr(contentText, "[")   ' Find the first `[`
        endPos = InStrRev(contentText, "]")  ' Find the last `]`
        
        ' Only keep the content between the first `[` and the last `]`, if both are found
        If startPos > 0 And endPos > 0 And endPos > startPos Then
            contentText = Mid(contentText, startPos + 1, endPos - startPos - 1)
        End If
        
        ' Step 2: Locate and remove only the last closing `]` in fileContent
        Dim lastBracketPos As Integer
        lastBracketPos = InStrRev(fileContent, "]") ' Find the last `]`
        
        If lastBracketPos > 0 Then
            fileContent = Left(fileContent, lastBracketPos - 3) ' Remove the last `]` and line breakss
        End If
        
        ' Step 3: Append contentText to fileContent and close the JSON array with `]`
        fileContent = fileContent & "," & contentText & "]"
        
    End If

    ' Rewrite the entire cache file with updated content
    WriteTextFile cacheFilePath, fileContent
    Debug.Print "Parsed innerJson and updated cache file"

    ' Step 6: Reset parsedData to ensure no previous entries are retained
    Set parsedData = Nothing
    Set parsedData = CreateObject("Scripting.Dictionary")
    ' Step 7: Process each project in the collection
    projectIndex = 1
For Each projectItem In innerJson
    ' Handle different types of projectItem
    Select Case TypeName(projectItem)
    
        ' If it's a String, try to parse it as JSON
        Case "String"
            On Error Resume Next
            Set projectItem = JsonConverter.ParseJSON(projectItem)
            On Error GoTo 0
            
            ' If parsing fails, use the string directly
            If TypeName(projectItem) = "String" Then
                ProjectName = projectItem  ' Assign the string directly if it's valid
            ElseIf TypeName(projectItem) = "Dictionary" Then
                ' Proceed to extract project name if it successfully parsed as a Dictionary
                If projectItem.Exists("Project Name") Then
                    ProjectName = projectItem("Project Name")
                Else
                    ProjectName = "N/A"
                End If
            Else
                ProjectName = "N/A"
            End If
        
        ' If it's already a Dictionary, extract the project name directly
        Case "Dictionary"
            If projectItem.Exists("Project Name") Then
                ProjectName = projectItem("Project Name")
            Else
                ProjectName = "N/A"
            End If
        
        ' If it's a Collection or Array, try to extract the first item
        Case "Collection", "Variant()"
            If projectItem.Count > 0 Then
                If TypeName(projectItem(1)) = "Dictionary" Then
                    ' Extract project name from the first item
                    If projectItem(1).Exists("Project Name") Then
                        ProjectName = projectItem(1)("Project Name")
                    Else
                        ProjectName = "N/A"
                    End If
                End If
            Else
                ProjectName = "N/A"
            End If
        
        ' Default case if it's an unknown type
        Case Else
            ProjectName = "N/A"
    
    End Select

    Debug.Print "projectName: " & ProjectName
Next projectItem
Debug.Print "Next projectItem"

    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    MsgBox errMsg, vbCritical, "Runtime Error"
End Sub

' Helper function to read the content of a text file
Private Function ReadTextFile(filePath As String) As String
    Dim fso As Object, fileStream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.fileExists(filePath) Then
        Set fileStream = fso.OpenTextFile(filePath, 1) ' ForReading
        ReadTextFile = fileStream.ReadAll
        fileStream.Close
    Else
        ReadTextFile = ""
    End If
End Function

' Helper function to write content to a text file
Private Sub WriteTextFile(filePath As String, content As String)
    Dim fso As Object, fileStream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
Debug.Print "CreateObject"
    Set fileStream = fso.CreateTextFile(filePath, True)
Debug.Print "Set fileStream"
    fileStream.Write content
Debug.Print "Write content"
Debug.Print " filePath: " & filePath
Debug.Print " content: " & content
    fileStream.Close
End Sub

' Helper function to convert nested dictionaries to a formatted string
Private Function BuildNestedString(ByVal dict As Variant, Optional indent As String = "") As String
    Dim result As String
    Dim key As Variant
    Dim value As Variant
    For Each key In dict.keys
        value = dict(key)
        result = result & indent & key & ": " & CStr(value) & vbCrLf
    Next key
    BuildNestedString = result
End Function




Private Function OpenTextFileForAppend(fileName As String) As textStream
    Dim fso As Object
    Dim fileStream As textStream

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.fileExists(fileName) Then
        Set fileStream = fso.OpenTextFile(fileName, ForAppending, True)
    Else
        Set fileStream = fso.CreateTextFile(fileName, True)
    End If

    Set OpenTextFileForAppend = fileStream
End Function

' Helper function to convert nested dictionaries to a formatted string
Private Function BuildNestedString_(ByVal dict As Variant, Optional indent As String = "") As String
    Dim result As String
    Dim key As Variant
    Dim value As Variant
    For Each key In dict.keys
        value = dict(key)
        result = result & indent & key & ": " & CStr(value) & vbCrLf
    Next key
    BuildNestedString_ = result
End Function




Private Function RemoveMarkdownFormatting(ByVal response As String) As String
    ' Remove the ```json or ``` from the response
    response = Replace(response, "```json", "")
    'response = Replace(response, "[", "")
    'response = Replace(response, "]", "")
    'response = Replace(response, "\\", "\")
    
    ' Remove extra newlines
    'response = Replace(response, vbCrLf, "")
    'response = Replace(response, vbLf, "")
    
    ' Return the cleaned response
    RemoveMarkdownFormatting = response
End Function

Private Function RemoveRemnants(ByVal response As String) As String
    ' Remove the ```json or ``` from the response
    'response = Replace(response, "```json", "")
    response = Replace(response, "[", "")
    response = Replace(response, "]", "")
    response = Replace(response, "\\", "\")
    response = Replace(response, "},", "}" & vbCrLf)
    
    ' Remove extra newlines
    'response = Replace(response, vbCrLf, "")
    'response = Replace(response, vbLf, "")
    
    ' Return the cleaned response
    RemoveRemnants = response
End Function

Private Function RemoveBSEscape(ByVal response As String) As String
    ' Remove the ```json or ``` from the response
    'response = Replace(response, "```json", "")
    'response = Replace(response, "[", "")
    'response = Replace(response, "]", "")
    'response = Replace(response, "\\", "\")
    response = Replace(response, "{" & vbCrLf & "  " & """projects"": ", "")
Debug.Print ("{" & vbCrLf & "  " & """projects"": ")
    response = Replace(response, "]" & vbCrLf & "}", "]")
    response = Replace(response, "```", "")
    ' Remove extra newlines
    'response = Replace(response, vbCrLf, "")
    'response = Replace(response, vbLf, "")
    
    ' Return the cleaned response
    RemoveBSEscape = response
End Function

Private Sub ProcessDictionary(ByVal dict As Variant, ByRef parsedData As Object, Optional parentKey As String = "")
    Dim keysArr As Variant, i As Long
    Dim k As Variant, v As Variant
    Dim fullKey As String
    Dim today As String: today = Format(Date, "mmmm dd, yyyy")
    Dim keyStr As String
    
    If TypeName(dict) <> "Dictionary" Then
        Debug.Print "ProcessDictionary skipped non-dictionary type: "; TypeName(dict)
        Exit Sub
    End If
    
    keysArr = dict.keys  ' safe snapshot of keys
    
    For i = LBound(keysArr) To UBound(keysArr)
        k = keysArr(i)
        keyStr = CStr(k)
        
        ' ---- get value safely ----
        On Error Resume Next
        If IsObject(dict(keyStr)) Then
            Set v = dict(keyStr)
        Else
            v = dict(keyStr)
        End If
        
        If Err.Number <> 0 Then
            Debug.Print "KeyErr at key=[" & keyStr & "] type=" & TypeName(k) & _
                        " err=" & Err.Number & " " & Err.Description
            Err.Clear
            GoTo ContinueLoop
        End If
        On Error GoTo 0
        
        ' *** FIX: Build proper hierarchical key ***
        If Len(parentKey) > 0 Then
            fullKey = parentKey & "." & keyStr  ' Build path like "project.name"
        Else
            fullKey = keyStr
        End If
        
        Select Case True
            Case IsObject(v)
                If TypeName(v) = "Dictionary" Then
                    ' *** FIX: Pass fullKey so nested items get proper paths ***
                    ProcessDictionary v, parsedData, fullKey
                ElseIf TypeName(v) = "Collection" Then
                    Dim itm As Variant, autoK As String, idx As Long
                    idx = 1
                    For Each itm In v
                        If IsObject(itm) Then
                            ' For collections of objects, use indexed keys
                            ProcessDictionary itm, parsedData, fullKey & "." & idx
                            idx = idx + 1
                        Else
                            ' For collections of scalars, concatenate with line breaks
                            If parsedData.Exists(fullKey) Then
                                parsedData(fullKey) = parsedData(fullKey) & vbCrLf & CStr(itm)
                            Else
                                parsedData.Add fullKey, CStr(itm)
                            End If
                        End If
                    Next itm
                End If
            Case VarType(v) = vbString
                v = UnescapeJsonString(CStr(v))
                If parsedData.Exists(fullKey) Then
                    parsedData(fullKey) = parsedData(fullKey) & vbCrLf & v
                Else
                    parsedData.Add fullKey, v
                End If
            Case Else
                ' Handle numbers, booleans, etc.
                If parsedData.Exists(fullKey) Then
                    parsedData(fullKey) = parsedData(fullKey) & vbCrLf & CStr(v)
                Else
                    parsedData.Add fullKey, CStr(v)
                End If
        End Select
        
        ' Clean up object reference if needed
        If IsObject(v) Then Set v = Nothing
        
ContinueLoop:
    Next i
    
    ' Only add Date at the root level
    If parentKey = "" And Not parsedData.Exists("Date") Then
        parsedData.Add "Date", today
    End If
End Sub



Private Sub ProcessDictionary_(ByVal dict As Variant, ByRef parsedData As Object, Optional parentKey As String = "")
    Dim key As Variant
    Dim value As Variant
    Dim today As Date
    Dim formattedDate As String
    
    ' Get today's date in "yyyy-mm-dd" format
    today = Date
    formattedDate = Format(today, "mmmm dd, yyyy")
Debug.Print "formattedDate" & formattedDate
 Debug.Print "VarType(dict): " & VarType(dict)
 
 
    For Each key In dict.keys
        value = dict(key)
 Debug.Print "value: " & value
       Dim fullKey As String
        If parentKey = "" Then
            fullKey = key
        Else
            fullKey = parentKey
        End If
Debug.Print "VarType(value): " & VarType(value)

        Select Case VarType(value)
            Case vbString
                ' Unescape the string
                value = UnescapeJsonString(value)
                If parsedData.Exists(fullKey) Then
                    ' Append to existing value
                    parsedData(fullKey) = parsedData(fullKey) & vbCrLf & value
                Else
                    parsedData.Add key:=fullKey, item:=value
                End If
            Case vbObject
 Debug.Print "TypeName(value): " & TypeName(value)
                If TypeName(value) = "Dictionary" Then
                    ' Build structured string from nested dictionary
                    Dim nestedText As String
                    nestedText = BuildNestedString_(value)
                    parsedData.Add key:=fullKey, item:=nestedText
                ElseIf TypeName(value) = "Collection" Then
                    ' Handle arrays
                    Dim item As Variant
                    Dim concatenatedText As String
                    concatenatedText = ""
                    For Each item In value
                        item = UnescapeJsonString(item)
                        concatenatedText = concatenatedText & item & vbCrLf
                    Next item
                    parsedData.Add key:=fullKey, item:=concatenatedText
                End If
            Case Else
                ' Handle other types if necessary
                parsedData.Add key:=fullKey, item:=CStr(value)
        End Select
    Next key
    
    ' Add today's date to the parsed data
    parsedData.Add key:="Date", item:=formattedDate
    
    
End Sub


Private Function UnescapeJsonString(ByVal s As String) As String
    s = Replace(s, "\""", """")  ' Replace escaped double quotes
    s = Replace(s, "\\", "\")    ' Replace escaped backslashes
    s = Replace(s, "\/", "/")    ' Replace escaped forward slash
    s = Replace(s, "\b", vbBack)
    s = Replace(s, "\f", vbFormFeed)
    s = Replace(s, "\n", vbCrLf) ' Replace escaped newlines
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\t", vbTab)
    's = Replace(s, "[", "")
    's = Replace(s, "]", "")

    ' Handle Unicode escape sequences \uXXXX if necessary
    UnescapeJsonString = s
End Function

' Function to join elements of a collection or array into a single string separated by newlines
Private Function JoinCollection(col As Variant) As String
    Dim item As Variant
    Dim result As String
    result = ""
    
    ' Check if col is an array
    If IsArray(col) Then
        For Each item In col
            result = result & item & vbCrLf
        Next item
    ' Check if col is a Collection
    ElseIf TypeName(col) = "Collection" Then
        For Each item In col
            result = result & item & vbCrLf
        Next item
    Else
        ' Handle other types if necessary
        result = CStr(col)
    End If
    
    ' Remove the trailing newline
    If Len(result) > 2 Then
        result = Left(result, Len(result) - 2)
    End If
    JoinCollection = result
End Function

' Function to clean text by normalizing line breaks and removing extra spaces
Private Function CleanText(text As String) As String
    ' Normalize line breaks to spaces
    text = Replace(text, vbCrLf, " ")
    text = Replace(text, vbLf, " ")
    
    ' Remove extra spaces
    text = Trim(text)
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    
    CleanText = text
End Function

Public Function ExtractAndParsePDFs(pdfFiles As Collection, cacheFilePath As String) As Boolean
    Dim pdfFile As Variant
    Dim fileName As String
    Dim tempFile As String
    Dim promptFile As String
    Dim extractedText As String
    Dim chatGPTResponse As String
    Dim fileCounter As Integer
    Dim totalFiles As Integer
    Dim cacheJsonText As String
    Dim cacheData As Variant
    Dim isFileInCache As Boolean
    Dim projectItem As Variant
    Dim allFilesInCache As Boolean

    Debug.Print ("ExtractAndParsePDFs Start")
    
    ' Initialize the return value and the flag
    ExtractAndParsePDFs = False
    allFilesInCache = True  ' Assume all files are in cache initially

    ' Check if cache file exists
    If Dir(cacheFilePath) <> "" Then
        ' Load cache data from file
        cacheJsonText = ReadTextFile(cacheFilePath)
        
        If cacheJsonText <> "" Then
            ' Parse JSON cache data with error handling
            On Error Resume Next
            Set cacheData = ParseJSON(cacheJsonText)
            Debug.Print ("ParseJson cacheJsonText " & cacheJsonText)
            If Err.Number <> 0 Then
                MsgBox "Cache file contains invalid JSON format.", vbExclamation, "Error"
                Set cacheData = Nothing ' Clear invalid cache data
            End If
            On Error GoTo 0
        End If
    End If

    ' Create a temporary file to store the extracted text
    
    tempFile = GetRelativePath("\test output\PDFTextOutput.txt")
    promptFile = GetRelativePath("\prompts\Prompt_sheets.txt")
    
    
    ' Remove previous version of the text file if it exists
    On Error Resume Next
    If Len(Dir(tempFile)) > 0 Then
        Kill tempFile
    End If
    On Error GoTo 0
    
    ' Step 1: Ensure there are PDF files to process
    If pdfFiles.Count = 0 Then
        MsgBox "No PDF files found.", vbExclamation, "No Files"
        Exit Function
    End If

    ' Get the total number of files in pdfFiles
    totalFiles = pdfFiles.Count
Debug.Print ("WE totalFiles " & totalFiles)
    fileCounter = 1 ' Initialize the file counter
    
    For Each pdfFile In pdfFiles
        frmSearchForm.lblStatus.Caption = ""
        fileName = GetFileNameFromPath(pdfFile)
Debug.Print ("UF fileName " & fileName)

        ' Check if the file is already in the cache (only if cache exists and is valid)
        isFileInCache = False
        Debug.Print ("TypeName(cacheData): " & TypeName(cacheData))

        ' Check if cacheData is a Collection and iterate if so
        If Not IsEmpty(cacheData) And TypeName(cacheData) = "Collection" Then
            For Each projectItem In cacheData
                ' Ensure projectItem is a Dictionary to access fields
                If TypeName(projectItem) = "Dictionary" Then
                    Dim cachedFileName As String
                    cachedFileName = GetFileNameFromPath(projectItem("File Path"))
Debug.Print "DE cachedFileName: " & cachedFileName
Debug.Print "QA fileName: " & fileName
                    ' Match the file name by checking if fileName is the same as cachedFileName
                    If InStr(1, cachedFileName, fileName, vbTextCompare) > 0 Then
Debug.Print "WA > 0"
                        isFileInCache = True
                        Exit For
                    End If
                Else
                    Debug.Print "Error: projectItem is not a Dictionary."
                End If
            Next projectItem
        Else
            Debug.Print "Error: cacheData is not a Collection or is empty."
        End If

        ' Skip processing if the file is already in the cache
        If isFileInCache Then
            Debug.Print "!!! Skipping " & fileName & " as it exists in cache."
            fileCounter = fileCounter + 1
            GoTo NextFile ' Skip to the end of the loop
        End If
        
        ' If we find a file thatfs not in the cache, set allFilesInCache to False
        allFilesInCache = False
        
        ' Update the label caption with the current file number and file name
        frmSearchForm.lblStatus.Visible = True
        frmSearchForm.lblStatus.TextAlign = 1
        frmSearchForm.lblStatus.Caption = "Reading " & fileCounter & "/" & totalFiles & ": " & fileName
        DoEvents

        ' Process the PDF file
        extractedText = CStr(pdfFile) & extractedText & ExtractTextFromPDFSheet(CStr(pdfFile), tempFile) & vbCrLf & vbCrLf

        ' Increment the file counter for the next file
        fileCounter = fileCounter + 1

NextFile:
    Next pdfFile

    ' If all selected files are in cache, display a message and exit the function
    If allFilesInCache Then
        MsgBox "All selected files are already in cache.", vbInformation, "No New Files"
        Exit Function
    End If

    frmSearchForm.lblStatus.TextAlign = 2
    Debug.Print "frmSearchForm.lblStatus.TextAlign = 2"

    ' Step 3: Indicate that the process is waiting for a response
    frmSearchForm.lblStatus.Caption = "Awaiting response..."
    DoEvents
        
    chatGPTResponse = CallAPIsyn(extractedText, promptFile)
    'Debug.Print "ChatGPT Response: " & chatGPTResponse
    
    frmSearchForm.lblStatus.Caption = "Extracting data... "
    DoEvents
    
    ' Step 4: Parse the response and populate parsed data
    Call SheetsDynamicExtractItems(chatGPTResponse, parsedSheetsData)
    Debug.Print "2 Call SheetsDynamicExtractItems "
    frmSearchForm.lblStatus.Caption = " "
    frmSearchForm.lblStatus.Visible = False

    ' If successful, return True
    If Not parsedSheetsData Is Nothing Then
        ExtractAndParsePDFs = True
    End If
    
    Debug.Print ("ExtractAndParsePDFs End")

End Function



' Helper function to read the contents of a text file
Private Function ReadTextFile_(filePath As String) As String
    Dim fileNumber As Integer
    Dim fileContent As String
    fileNumber = FreeFile

    On Error GoTo ErrorHandler
    Open filePath For Input As fileNumber
    fileContent = Input(LOF(fileNumber), fileNumber)
    Close fileNumber
    ReadTextFile_ = fileContent
    Exit Function

ErrorHandler:
    ReadTextFile_ = ""
    Close fileNumber
End Function


' --- Extract Text from a Single PDF File using mutool.exe or tesseract.exe ---
Private Function ExtractTextFromPDFSheet_(pdfFilePath As Variant, tempFile As String) As String
    Dim shellCommand As String
    Dim engine As String
    Dim lineText As String
    Dim pdfText As String
    Dim fso As Object
    Dim fileExists As Boolean
    Dim startTime As Double
    Dim lastSize As Long
    Dim currentSize As Long
    Dim buffer As String ' Use buffer to avoid large string concatenation issues

    
    ' Path to MuPDF's mutool.exe
    engine = GetRelativePath("\PDFreader\mupdf-1.24.0-windows\mutool.exe")
    ' Construct the shell command to extract text using mutool.exe
    'shellCommand = """" & engine & """ extract """ & pdfFilePath & """ > """ & tempFile & """"
    shellCommand = engine & " draw -F txt -o """ & tempFile & """ """ & pdfFilePath & """"
    
    
    
   'engine = GetRelativePath("\Tesseract-OCR\tesseract.exe")
    

    ' Use tesseract to convert PDF to text
    'shellCommand = engine & " " & pdfFilePath & " " & tempFile & " -l jpn pdf"
    'command = engine & Chr(34) & pdfPath & Chr(34) & " " & Chr(34) & Left(txtPath, InStrRev(txtPath, ".") - 1) & Chr(34) & " -l jpn pdf"
    'shellCommand = Chr(34) & engine & Chr(34) & " " & _
    '      Chr(34) & pdfFilePath & Chr(34) & " " & _
    '      Chr(34) & tempFile & Chr(34) & " -l jpn"
    
    
    ' Debugging: Print the command and file paths
    Debug.Print "Shell Command: " & shellCommand
    Debug.Print "PDF File Path: " & pdfFilePath
    Debug.Print "Temp File Path: " & tempFile

    ' Execute the shell command using Shell
    Shell shellCommand, vbHide

    ' Wait for the text file to be generated, up to 30 seconds, checking file size stability
    Set fso = CreateObject("Scripting.FileSystemObject")
    startTime = Timer
    lastSize = 0
    
    Do
        ' Check if the file exists
        If fso.fileExists(tempFile) Then
            ' Get the current file size
            currentSize = FileLen(tempFile)
            
            ' Check if the file size has stabilized (i.e., it's no longer increasing)
            If currentSize = lastSize And currentSize > 0 Then
                Exit Do  ' File size hasn't changed, assume it's done writing
            End If
            
            ' Update lastSize for the next comparison
            lastSize = currentSize
        End If
        
        Debug.Print "Waiting for file size to stabilize... Current size: " & currentSize
        Application.Wait Now + TimeValue("00:00:01")  ' Wait 1 second
        
        ' Timeout after 30 seconds
        If Timer - startTime > 20 Then
            MsgBox "Timed out waiting for the PDF to be converted.", vbExclamation
            Exit Function
        End If
    Loop

    ' Wait for the command to complete and for the temp file to be generated
    'DoEvents
    'Application.Wait (Now + TimeValue("0:00:02")) ' Adjust this wait time if needed

    ' Check if the temp file exists before reading
    If Dir(tempFile) = "" Then
        MsgBox "Temporary file not found: " & tempFile, vbExclamation, "File Error"
        Exit Function
    End If

    ' Read the extracted text from the temp file
    pdfText = ""
    On Error Resume Next
    Open tempFile For Input As #1
    pdfText = pdfFilePath & vbCrLf
Debug.Print "pdfFilePath: " & pdfFilePath
    Do While Not EOF(1)
        Line Input #1, lineText
        pdfText = pdfText & lineText & vbCrLf
    Loop
    Close #1
    On Error GoTo 0
Debug.Print "pdfText: " & pdfText

    ' Return the extracted text
    ExtractTextFromPDFSheet_ = pdfText
    
End Function


Public Function ExtractTextFromPDFSheet(pdfPath As String, txtPath As String) As String
    Dim command As String
    Dim engine As String
    Dim utf8Text As String
    
    ' Path to MuPDF
    engine = GetRelativePath("\PDFreader\mupdf-1.24.0-windows\mutool.exe")
    
    ' Extract Text using MuPDF
    command = engine & " draw -F txt -o " & Chr(34) & txtPath & Chr(34) & " " & Chr(34) & pdfPath & Chr(34)
    Shell command, vbNormalFocus
    
    ' Wait for Extraction to Complete
    WaitForFile txtPath
    
    ' Convert to UTF-8
    utf8Text = ConvertShiftJISToUTF8(txtPath)
Debug.Print "AFTER convert to UTF8" & utf8Text
    
    ExtractTextFromPDFSheet = utf8Text
End Function

Public Function ConvertShiftJISToUTF8(filePath As String) As String
    Dim objStream As Object
    Dim content As String
    
    ' Use ADODB Stream to handle encoding correctly
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' Text
    objStream.Charset = "UTF-8" ' Interpret as UTF-8
    'objStream.Charset = "Shift-JIS" ' Interpret as Shift-JIS
    objStream.Open
    objStream.LoadFromFile filePath
    content = objStream.ReadText
    objStream.Close
    
    ' Re-write the text as UTF-8
    'objStream.Charset = "UTF-8"
    'objStream.Open
    'objStream.WriteText content
    'objStream.SaveToFile filePath, 2 ' Overwrite with UTF-8
    'objStream.Close
    
    ConvertShiftJISToUTF8 = content
End Function


' Helper Function to Wait for the File to be Created
Public Sub WaitForFile(filePath As String)
    Dim fso As Object
    Dim fileExists As Boolean
    Dim fileSize As Long
    Dim newSize As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = False
    fileSize = 0
    
    ' Wait Until the File Exists and Stabilizes
    Do While Not fileExists
        fileExists = fso.fileExists(filePath)
        DoEvents
    Loop
    
    Do
        newSize = fso.GetFile(filePath).Size
        If newSize > 0 And newSize = fileSize Then Exit Do
        fileSize = newSize
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop
End Sub

' Convert Extracted Text to UTF-8 (Fix Garbled Japanese)
Public Function ConvertToUTF8(filePath As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim content As String
    
    ' Read the File as ANSI/Shift-JIS
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2) ' Read with Default System Encoding
    content = ts.ReadAll
    ts.Close
    
    ' Re-write the Text as UTF-8
    Set ts = fso.OpenTextFile(filePath, 2, True, -1) ' Write as UTF-8
    ts.Write content
    ts.Close
    
    ' Return the Re-encoded Content
    ConvertToUTF8 = content
End Function


' --- Get a Sorted Collection of PDF Files ---
Private Function GetSortedPDFFiles(folderPath As String) As Collection
    Dim fileName As String
    Dim pdfFiles As New Collection
    Dim i As Integer, j As Integer
    Dim temp As String

    ' Read all PDF files in the specified folder
    fileName = Dir(folderPath & "\*.pdf")
    Do While fileName <> ""
        pdfFiles.Add folderPath & "\" & fileName
        fileName = Dir
    Loop

    ' Sort the PDF files alphabetically by their file names
    For i = 1 To pdfFiles.Count - 1
        For j = i + 1 To pdfFiles.Count
            If UCase(Dir(pdfFiles(i))) > UCase(Dir(pdfFiles(j))) Then
                ' Swap the files
                temp = pdfFiles(i)
                pdfFiles(i) = pdfFiles(j)
                pdfFiles(j) = temp
            End If
        Next j
    Next i

    Set GetSortedPDFFiles = pdfFiles
End Function



Sub PrintDictionaryKeys(dict As Object)
    Dim key As Variant
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Define the path to the cache text file
    filePath = GetRelativePath("\test output\dictionary_cache.txt")

    ' Check if the file already exists
    If Dir(filePath) = "" Then
        ' If the file does not exist, create a new one for output
        fileNum = FreeFile
        Open filePath For Output As fileNum
        Print #fileNum, "PrintDictionaryKeys: "
    Else
        ' If the file exists, open it for appending
        fileNum = FreeFile
        Open filePath For Append As fileNum
        Print #fileNum, "Appending PrintDictionaryKeys: "
    End If
    
    ' Append the dictionary keys to the file
    If dict.Count = 0 Then
        Print #fileNum, "No keys in the dictionary."
    Else
        For Each key In dict.keys
            Print #fileNum, "Key: " & key
        Next key
    End If
    
    ' Close the file
    Close fileNum

    ' Optional: Provide feedback in Immediate window (Debug.Print)
    Debug.Print "Dictionary keys have been appended to the file: " & filePath
End Sub


Function GetFileNameFromPath(filePath As Variant) As String
    Dim fileName As String
    Dim lastBackslash As Integer

    ' Find the position of the last backslash
    lastBackslash = InStrRev(filePath, "\")

    ' Extract the file name from the full file path
    If lastBackslash > 0 Then
        fileName = Mid(filePath, lastBackslash + 1)
    Else
        ' If no backslash is found, return the original string
        fileName = filePath
    End If

    ' Return the file name
    GetFileNameFromPath = EscapeJsonString(fileName)
End Function





