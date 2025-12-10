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

' Main function: Returns complete JSON for all review data
Public Function ParseRFPToJson(ByVal rfpText As String) As String
    On Error GoTo ErrorHandler
    
    Dim promptFilePath As String
    Dim responseText As String
    Dim jsonResult As String
    
    ' Use the same prompt file path as your existing ParseRFPText function
    promptFilePath = GetRelativePath("\prompts\Prompt_all-parse.txt")
    
    Debug.Print "ParseRFPToJson: Calling API with prompt file: " & promptFilePath
    
    ' Call the API using the SAME function as the Analyse button
    responseText = APIModule.CallAPIsyn(rfpText, promptFilePath)
    
    Debug.Print "ParseRFPToJson: Received response length: " & Len(responseText)
    
    If Len(responseText) = 0 Then
        MsgBox "No response received from API.", vbCritical
        ParseRFPToJson = ""
        Exit Function
    End If
    
    ' Extract JSON from the API response
    jsonResult = ExtractJSONFromResponse(responseText)
    
    Debug.Print "ParseRFPToJson: Extracted JSON length: " & Len(jsonResult)
    
    ' Validate the JSON before returning
    If Len(jsonResult) > 0 Then
        On Error Resume Next
        Dim testParse As Object
        Set testParse = JsonConverter.ParseJSON(jsonResult)
        If Err.Number <> 0 Then
            Debug.Print "ParseRFPToJson: JSON validation failed: " & Err.Description
            MsgBox "Failed to parse JSON response. Check the prompt output format.", vbCritical
            jsonResult = ""
        End If
        On Error GoTo ErrorHandler
    End If
    
    ParseRFPToJson = jsonResult
    Exit Function

ErrorHandler:
    Debug.Print "ERROR in ParseRFPToJson: " & Err.Description
    MsgBox "Error parsing RFP: " & Err.Description, vbCritical
    ParseRFPToJson = ""
End Function

' Your existing ParseAndOpenReview - keep as is
Public Sub ParseAndOpenReview(ByVal rfpText As String)
    Dim jsonResult As String
    jsonResult = ParseRFPToJson(rfpText)
    
    If Len(jsonResult) > 0 Then
        modReviewBridge.OpenReviewWithJson jsonResult
    Else
        MsgBox "Failed to parse RFP text. No data to review.", vbExclamation
    End If
End Sub

' Extract the actual JSON content from ChatGPT API response
Private Function ExtractJSONFromResponse(ByVal responseText As String) As String
    On Error GoTo ErrorHandler
    
    Dim jsonObj As Object
    Dim choices As Object
    Dim messageContent As String
    
    Debug.Print "ExtractJSONFromResponse: Starting extraction..."
    
    ' Parse the API response wrapper
    Set jsonObj = JsonConverter.ParseJSON(responseText)
    
    ' Navigate to choices[0].message.content
    If jsonObj.Exists("choices") Then
        Set choices = jsonObj("choices")
        If choices.Count > 0 Then
            messageContent = choices(1)("message")("content")
            
            Debug.Print "ExtractJSONFromResponse: Raw content length: " & Len(messageContent)
            
            ' The content should be the RFP analysis JSON
            ' Remove markdown code blocks if present
            messageContent = Trim(messageContent)
            
            ' Remove ```json and ``` markers
            If Left(messageContent, 7) = "```json" Then
                messageContent = Mid(messageContent, 8)
                Dim endPos As Long
                endPos = InStrRev(messageContent, "```")
                If endPos > 0 Then
                    messageContent = Left(messageContent, endPos - 1)
                End If
                Debug.Print "ExtractJSONFromResponse: Removed ```json markers"
            ElseIf Left(messageContent, 3) = "```" Then
                messageContent = Mid(messageContent, 4)
                endPos = InStrRev(messageContent, "```")
                If endPos > 0 Then
                    messageContent = Left(messageContent, endPos - 1)
                End If
                Debug.Print "ExtractJSONFromResponse: Removed ``` markers"
            End If
            
            ' Trim whitespace
            messageContent = Trim(messageContent)
            
            ' Additional cleanup: remove any leading/trailing newlines
            Do While Left(messageContent, 1) = vbLf Or Left(messageContent, 1) = vbCr
                messageContent = Mid(messageContent, 2)
            Loop
            Do While Right(messageContent, 1) = vbLf Or Right(messageContent, 1) = vbCr
                messageContent = Left(messageContent, Len(messageContent) - 1)
            Loop
            
            Debug.Print "ExtractJSONFromResponse: Cleaned content length: " & Len(messageContent)
            Debug.Print "ExtractJSONFromResponse: First 200 chars: " & Left(messageContent, 200)
            
            ExtractJSONFromResponse = messageContent
            Exit Function
        Else
            Debug.Print "ExtractJSONFromResponse: No choices in response"
        End If
    Else
        Debug.Print "ExtractJSONFromResponse: 'choices' key not found in response"
    End If
    
    ' If we couldn't parse it, return the raw response
    Debug.Print "ExtractJSONFromResponse: Returning raw response as fallback"
    ExtractJSONFromResponse = responseText
    Exit Function

ErrorHandler:
    Debug.Print "ERROR in ExtractJSONFromResponse: " & Err.Description
    ' Return raw response on error
    ExtractJSONFromResponse = responseText
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

Sub SheetsDynamicExtractItems(gptResponse As String, ByRef parsedData As Object, Optional ByVal originalFilePaths As Collection = Nothing)
    On Error GoTo ErrorHandler
    Dim jsonResponse As Object
    Dim contentText As String
    Dim innerJson As Object
    Dim projectItem As Variant
    Dim attributeKey As Variant
    Dim childDict As Object
    Dim projectName As Variant
    Dim projectIndex As Integer
    Dim jsonString As String
    Dim fileContent As String
    Dim cacheFilePath As String
    Dim pName As String

    cacheFilePath = GetRelativePath("\test output\cache.json")

    ' Check if gptResponse is empty or invalid
    If Len(Trim(gptResponse)) = 0 Then
        MsgBox "No response received from API. Please check your API key and connection.", vbCritical, "API Error"
        Exit Sub
    End If

    ' Step 1: Remove surrounding markdown formatting (```json\n...\n```)
    gptResponse = RemoveMarkdownFormatting(gptResponse)

    ' Step 2: Parse the JSON response using JsonConverter with error handling
    On Error Resume Next
    Set jsonResponse = JsonConverter.ParseJSON(gptResponse)
    If Err.Number <> 0 Or jsonResponse Is Nothing Then
        Debug.Print "ERROR: Failed to parse API response as JSON"
        Debug.Print "Response (first 1000 chars): " & Left(gptResponse, 1000)
        MsgBox "Failed to parse API response." & vbCrLf & _
               "The API may have returned an error or invalid JSON." & vbCrLf & vbCrLf & _
               "Response: " & Left(gptResponse, 300), vbCritical, "API Response Error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo ErrorHandler
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
        fileContent = contentText & vbCrLf
    Else
        ' Parse existing cache to ensure it's a valid array
        Dim existingArray As Object
        On Error Resume Next
        Set existingArray = JsonConverter.ParseJSON(fileContent)
        On Error GoTo ErrorHandler
        
        ' If cache is valid JSON array, extract items from contentText and append
        If Not existingArray Is Nothing And TypeName(existingArray) = "Collection" Then
            ' Extract items from contentText (which should be an array from API)
            Dim newArray As Object
            Set newArray = JsonConverter.ParseJSON(contentText)
            
            If Not newArray Is Nothing And TypeName(newArray) = "Collection" Then
                ' Append each new item to existing array, checking for duplicates
                Dim newItem As Variant
                Dim existingItem As Variant
                Dim isDuplicate As Boolean
                Dim newFileName As String
                Dim existingFileName As String
                Dim newItemsCount As Long
                Dim duplicatesCount As Long
                Dim correctedFilePath As String
                
                newItemsCount = 0
                duplicatesCount = 0
                
                Debug.Print ">>> API returned " & newArray.Count & " items in this batch"
                
                Dim duplicateProjectMsg As String
                duplicateProjectMsg = ""
                
                Dim batchIndex As Long
                batchIndex = 1
                
                For Each newItem In newArray
                    isDuplicate = False
                    newFileName = ""  ' Reset for each item to prevent false duplicates
                    
                    ' Get the file name from the new item
                    If TypeName(newItem) = "Dictionary" Then
                        ' STRICT MAPPING: Always use the original file path from the batch if available
                        ' This bypasses API hallucinations and character corruption completely
                        If Not originalFilePaths Is Nothing And newArray.Count = originalFilePaths.Count Then
                            correctedFilePath = originalFilePaths(batchIndex)
                            newItem("File Path") = correctedFilePath
                            newFileName = GetFileNameFromPath(correctedFilePath)
                            Debug.Print ">>> Mapped by INDEX: " & batchIndex & " -> " & newFileName
                        Else
                            ' Fallback to existing logic if counts don't match
                            If newItem.Exists("File Path") Then
                                newFileName = GetFileNameFromPath(newItem("File Path"))
                                correctedFilePath = FindOriginalFilePath(newItem("File Path"), originalFilePaths)
                                If correctedFilePath <> "" Then
                                    newItem("File Path") = correctedFilePath
                                    newFileName = GetFileNameFromPath(correctedFilePath)
                                End If
                            Else
                                Debug.Print ">>> WARNING: Item has no 'File Path' field, cannot check for duplicates"
                            End If
                        End If
                    End If
                    
                    ' If we couldn't determine the filename, try to map by batch index or dedupe by Project Name
                    If Len(newFileName) = 0 Then
                        Dim mappedFromBatch As Boolean
                        mappedFromBatch = False

                        If Not originalFilePaths Is Nothing And batchIndex <= originalFilePaths.Count Then
                            correctedFilePath = originalFilePaths(batchIndex)
                            newItem("File Path") = correctedFilePath
                            newFileName = GetFileNameFromPath(correctedFilePath)
                            mappedFromBatch = True
                            Debug.Print ">>> Mapped missing filename by batch index: " & newFileName
                        End If

                        If Not mappedFromBatch Then
                            ' If we have a Project Name, dedupe by project name to avoid duplicates
                            If TypeName(newItem) = "Dictionary" Then
                                If newItem.Exists("Project Name") Then
                                    pName = newItem("Project Name")
                                    Dim existsByProject As Boolean
                                    existsByProject = False
                                    Dim exChk As Variant
                                    For Each exChk In existingArray
                                        If TypeName(exChk) = "Dictionary" And exChk.Exists("Project Name") Then
                                            If StrComp(exChk("Project Name"), pName, vbTextCompare) = 0 Then
                                                existsByProject = True
                                                Exit For
                                            End If
                                        End If
                                    Next exChk

                                    If existsByProject Then
                                        Debug.Print ">>> Skipping item because project name already exists in cache: " & pName
                                        duplicatesCount = duplicatesCount + 1
                                        batchIndex = batchIndex + 1
                                        GoTo NextNewItem
                                    Else
                                        ' Add the item (no filename but unique project name)
                                        existingArray.Add newItem
                                        newItemsCount = newItemsCount + 1
                                        Debug.Print ">>> ADDED (no filename) for unique project: " & pName
                                        batchIndex = batchIndex + 1
                                        GoTo NextNewItem
                                    End If
                                Else
                                    Debug.Print ">>> WARNING: No filename determined for item and no project name; adding without duplicate check"
                                    existingArray.Add newItem
                                    newItemsCount = newItemsCount + 1
                                    batchIndex = batchIndex + 1
                                    GoTo NextNewItem
                                End If
                            End If
                        End If
                    End If
                        
                    Debug.Print ">>> Checking item: " & newFileName
                    
                    ' Check if this file already exists in cache
                    For Each existingItem In existingArray
                        If TypeName(existingItem) = "Dictionary" And existingItem.Exists("File Path") Then
                            existingFileName = GetFileNameFromPath(existingItem("File Path"))
                            Debug.Print ">>>   Comparing with: " & existingFileName & " (len=" & Len(existingFileName) & ")"
                            If StrComp(newFileName, existingFileName, vbTextCompare) = 0 Then
                                isDuplicate = True
                                duplicatesCount = duplicatesCount + 1
                                Debug.Print ">>> DUPLICATE: Skipping " & newFileName
                                Exit For
                            End If
                        End If
                    Next existingItem
                    
                    ' Only add if not a duplicate
                    If Not isDuplicate Then
                        ' Check for duplicate Project Name (same project, different file)
                        Dim exItem As Variant
                        Dim exFile As String
                        
                        If TypeName(newItem) = "Dictionary" Then
                            If newItem.Exists("Project Name") Then
                                pName = newItem("Project Name")
                                
                                For Each exItem In existingArray
                                    If TypeName(exItem) = "Dictionary" And exItem.Exists("Project Name") Then
                                        If StrComp(exItem("Project Name"), pName, vbTextCompare) = 0 Then
                                            ' Same project name found
                                            If exItem.Exists("File Path") Then
                                                exFile = GetFileNameFromPath(exItem("File Path"))
                                                ' If filenames are different, report it
                                                If StrComp(exFile, newFileName, vbTextCompare) <> 0 Then
                                                    duplicateProjectMsg = duplicateProjectMsg & _
                                                        "Project: " & pName & vbCrLf & _
                                                        " - Existing: " & exFile & vbCrLf & _
                                                        " - New: " & newFileName & vbCrLf & vbCrLf
                                                End If
                                            End If
                                        End If
                                    End If
                                Next exItem
                            End If
                        End If
                        
                        existingArray.Add newItem
                        newItemsCount = newItemsCount + 1
                        Debug.Print ">>> ADDED to cache: " & newFileName
                    Else
                        Debug.Print ">>> SKIPPED (duplicate): " & newFileName
                    End If
                    
                    batchIndex = batchIndex + 1
NextNewItem:
                Next newItem
                
                If duplicateProjectMsg <> "" Then
                    MsgBox "Note: Multiple files found for the same project." & vbCrLf & _
                           "Both files have been saved to the cache." & vbCrLf & vbCrLf & _
                           duplicateProjectMsg, vbInformation, "Multiple Files for Same Project"
                End If
                
                Debug.Print ">>> Batch summary: " & newItemsCount & " added, " & duplicatesCount & " duplicates skipped"
                
                ' Convert back to JSON manually (JsonConverter.ConvertToJson loses data with large Collections)
                ' Build JSON array manually to preserve all items
                fileContent = "["
                Dim itemIndex As Long
                For itemIndex = 1 To existingArray.Count
                    If itemIndex > 1 Then fileContent = fileContent & ","
                    fileContent = fileContent & JsonConverter.ConvertToJson(existingArray(itemIndex), Whitespace:=2)
                Next itemIndex
                fileContent = fileContent & "]"
            End If
        Else
            ' If existing cache is invalid, replace with new content
            fileContent = contentText & vbCrLf
        End If
    End If

    ' Rewrite the entire cache file with updated content
    WriteTextFile cacheFilePath, fileContent
    Debug.Print "Parsed innerJson and updated cache file"
    
    ' DEBUG: Verify what was actually written to cache
    Dim verifyContent As String
    verifyContent = ReadTextFile(cacheFilePath)
    Dim verifyArray As Object
    On Error Resume Next
    Set verifyArray = JsonConverter.ParseJSON(verifyContent)
    On Error GoTo 0
    If Not verifyArray Is Nothing And TypeName(verifyArray) = "Collection" Then
        Debug.Print ">>> CACHE VERIFICATION: " & verifyArray.Count & " items now in cache file"
    End If

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
                    projectName = projectItem  ' Assign the string directly if it's valid
                ElseIf TypeName(projectItem) = "Dictionary" Then
                    ' Proceed to extract project name if it successfully parsed as a Dictionary
                    If projectItem.Exists("Project Name") Then
                        projectName = projectItem("Project Name")
                    Else
                        projectName = "N/A"
                    End If
                Else
                    projectName = "N/A"
                End If
            
            ' If it's already a Dictionary, extract the project name directly
            Case "Dictionary"
                If projectItem.Exists("Project Name") Then
                    projectName = projectItem("Project Name")
                Else
                    projectName = "N/A"
                End If
            
            ' If it's a Collection or Array, try to extract the first item
            Case "Collection", "Variant()"
                If projectItem.Count > 0 Then
                    If TypeName(projectItem(1)) = "Dictionary" Then
                        ' Extract project name from the first item
                        If projectItem(1).Exists("Project Name") Then
                            projectName = projectItem(1)("Project Name")
                        Else
                            projectName = "N/A"
                        End If
                    End If
                Else
                    projectName = "N/A"
                End If
            
            ' Default case if it's an unknown type
            Case Else
                projectName = "N/A"
        
        End Select

        Debug.Print "projectName: " & projectName
    Next projectItem
    Debug.Print "Next projectItem"

    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    MsgBox errMsg, vbCritical, "Runtime Error"
End Sub

' Helper function to read the content of a text file with UTF-8 encoding
Private Function ReadTextFile(filePath As String) As String
    Dim fso As Object
    Dim stream As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.fileExists(filePath) Then
        ' Use ADODB.Stream to read UTF-8 files correctly (handles both with/without BOM)
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' Text
        stream.Charset = "UTF-8"
        stream.Open
        
        On Error Resume Next
        stream.LoadFromFile filePath
        If Err.Number = 0 Then
            ReadTextFile = stream.ReadText
        Else
            Debug.Print "ReadTextFile ERROR: " & Err.Description
            ReadTextFile = ""
        End If
        On Error GoTo 0
        
        stream.Close
    Else
        ReadTextFile = ""
    End If
End Function

' Helper function to write content to a text file with UTF-8 encoding
Private Sub WriteTextFile(filePath As String, content As String)
    Dim stream As Object
    Dim tempStream As Object
    Dim tempFile As String
    
    ' Use ADODB.Stream to write UTF-8 WITHOUT BOM (critical for JSON parsing)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText content
    
    ' Save to temp location first
    tempFile = filePath & ".tmp"
    stream.SaveToFile tempFile, 2 ' Overwrite
    stream.Close
    
    ' Now read as binary and write without BOM
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.LoadFromFile tempFile
    stream.Position = 3 ' Skip BOM (3 bytes for UTF-8 BOM: EF BB BF)
    
    Set tempStream = CreateObject("ADODB.Stream")
    tempStream.Type = 1 ' Binary
    tempStream.Open
    stream.CopyTo tempStream
    tempStream.SaveToFile filePath, 2 ' Overwrite
    
    stream.Close
    tempStream.Close
    
    ' Clean up temp file
    On Error Resume Next
    Kill tempFile
    On Error GoTo 0
    
    Debug.Print "WriteTextFile: Saved UTF-8 without BOM to " & filePath
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
    ' Remove wrapper structure if present
    response = Replace(response, "{" & vbCrLf & "  " & """projects"": ", "")
    response = Replace(response, "]" & vbCrLf & "}", "]")
    response = Replace(response, "```", "")
    
    ' Remove any leading commas that might have been introduced
    response = Trim(response)
    If Left(response, 2) = "[," Then
        response = "[" & Mid(response, 3)
    End If
    
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
    Dim pName As String
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
    Dim batchSize As Integer
    Dim batchCounter As Integer
    Dim filesToProcess As Collection

    Debug.Print ("ExtractAndParsePDFs Start")
    
    ' Initialize the return value and the flag
    ExtractAndParsePDFs = False
    allFilesInCache = True  ' Assume all files are in cache initially
    batchSize = 20  ' Process 5 files at a time to avoid token limits and hallucinations
    Set filesToProcess = New Collection

    ' Check if cache file exists
    If Dir(cacheFilePath) <> "" Then
        ' Load cache data from file
        cacheJsonText = ReadTextFile(cacheFilePath)
        
        If cacheJsonText <> "" Then
            ' Parse JSON cache data with error handling
            On Error Resume Next
            Set cacheData = ParseJSON(cacheJsonText)
            If Err.Number <> 0 Then
                Debug.Print "ERROR: Cache file contains invalid JSON format."
                Debug.Print "Cache file path: " & cacheFilePath
                Debug.Print "First 500 chars: " & Left(cacheJsonText, 500)
                MsgBox "Cache file contains invalid JSON format." & vbCrLf & _
                       "File: " & cacheFilePath & vbCrLf & _
                       "Error: " & Err.Description, vbExclamation, "JSON Parse Error"
                Set cacheData = Nothing ' Clear invalid cache data
            Else
                Debug.Print "Successfully parsed cache JSON from: " & cacheFilePath
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
    batchCounter = 0 ' Initialize batch counter
    
    ' First pass: identify files not in cache and add to filesToProcess
    For Each pdfFile In pdfFiles
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
                    If projectItem.Exists("File Path") Then
                        cachedFileName = GetFileNameFromPath(projectItem("File Path"))
                        ' Match the file name using EXACT case-insensitive comparison
                        If StrComp(fileName, cachedFileName, vbTextCompare) = 0 Then
                            isFileInCache = True
                            Debug.Print "Cache hit (exact match): " & fileName
                            Exit For
                        End If
                    End If
                End If
            Next projectItem
        End If

        ' Add to processing queue if not in cache
        If Not isFileInCache Then
            ' Avoid adding duplicate files (same filename) to processing queue
            Dim alreadyQueued As Boolean
            alreadyQueued = False
            Dim q As Variant
            For Each q In filesToProcess
                If StrComp(GetFileNameFromPath(CStr(q)), fileName, vbTextCompare) = 0 Then
                    alreadyQueued = True
                    Exit For
                End If
            Next q
            If Not alreadyQueued Then
                filesToProcess.Add pdfFile
            Else
                Debug.Print "Skipping duplicate selection for processing queue: " & fileName
            End If
            allFilesInCache = False
        Else
            Debug.Print "!!! Skipping " & fileName & " as it exists in cache."
        End If
NextFile:
    Next pdfFile

    ' If all selected files are in cache, display a message and exit the function
    If allFilesInCache Then
        MsgBox "All selected files are already in cache.", vbInformation, "No New Files"
        Exit Function
    End If

    ' Second pass: process files in batches
    Debug.Print "Processing " & filesToProcess.Count & " files in batches of " & batchSize
    
    ' Create a processing log file
    Dim logFilePath As String
    Dim logFile As Integer
    logFilePath = GetRelativePath("\test output\processing_log.txt")
    logFile = FreeFile
    Open logFilePath For Output As logFile
    Print #logFile, "=== PDF Processing Log ==="
    Print #logFile, "Started: " & Now
    Print #logFile, "Total files to process: " & filesToProcess.Count
    Print #logFile, ""
    
    Dim filesInCurrentBatch As Long
    Dim batchFilePaths As Collection  ' Track original file paths for this batch
    Dim allProcessedFilePaths As Collection  ' Accumulate file paths across ALL batches
    filesInCurrentBatch = 0
    Set batchFilePaths = New Collection
    Set allProcessedFilePaths = New Collection
    
    For fileCounter = 1 To filesToProcess.Count
        pdfFile = filesToProcess(fileCounter)
        fileName = GetFileNameFromPath(pdfFile)
        
        ' Update the label caption with the current file number and file name
        frmSearchForm.lblStatus.Visible = True
        frmSearchForm.lblStatus.TextAlign = 1
        frmSearchForm.lblStatus.Caption = "Reading " & fileCounter & "/" & filesToProcess.Count & ": " & fileName
        DoEvents

        ' Process the PDF file - add file path, then the extracted text content
        Print #logFile, "[" & fileCounter & "/" & filesToProcess.Count & "] Processing: " & fileName
        Debug.Print ">>> Processing file " & fileCounter & ": " & fileName
        Dim extractedContent As String
        extractedContent = ExtractTextFromPDFSheet(CStr(pdfFile), tempFile)
        
        ' Check if extraction was successful
        If Len(extractedContent) > 0 Then
            ' Use the ACTUAL file path from disk (corrected by FindSimilarFile if needed)
            ' This ensures the correct filename is sent to the API
            Dim actualFilePath As String
            actualFilePath = CStr(pdfFile)
            
            ' Normalize the file path to ensure consistent encoding
            actualFilePath = NormalizeFilePath(actualFilePath)
            
            extractedText = extractedText & actualFilePath & vbCrLf & extractedContent & vbCrLf & vbCrLf
            batchFilePaths.Add actualFilePath  ' Track the original file path for this batch
            allProcessedFilePaths.Add actualFilePath  ' Accumulate across all batches
            filesInCurrentBatch = filesInCurrentBatch + 1
            Print #logFile, "  SUCCESS: Extracted " & Len(extractedContent) & " chars"
            Debug.Print ">>> Successfully extracted " & Len(extractedContent) & " characters from: " & fileName
        Else
            Print #logFile, "  ERROR: No content extracted! SKIPPING FILE."
            Debug.Print ">>> CRITICAL ERROR: No content extracted from: " & fileName & ". This file will NOT be in the cache."
        End If
        
        batchCounter = batchCounter + 1

        ' When batch is full or we've reached the last file, send to API
        If batchCounter >= batchSize Or fileCounter = filesToProcess.Count Then
            ' Only send to API if we have content to process
            If Len(extractedText) > 0 Then
                frmSearchForm.lblStatus.TextAlign = 2
                frmSearchForm.lblStatus.Caption = "Awaiting response for batch " & Int((fileCounter - 1) / batchSize) + 1 & "..."
                DoEvents
                
                Debug.Print ">>> Sending batch with " & filesInCurrentBatch & " files (" & Len(extractedText) & " characters) to API"
                ' Send batch to API
                chatGPTResponse = CallAPIsyn(extractedText, promptFile)
                Debug.Print ">>> Received API response: " & Len(chatGPTResponse) & " characters"
                
                frmSearchForm.lblStatus.Caption = "Extracting data... "
                DoEvents
                
                ' Parse the response and update cache - pass BATCH file paths to fix corrupted paths and prevent cross-batch hallucinations
                Call SheetsDynamicExtractItems(chatGPTResponse, parsedSheetsData, batchFilePaths)
                Debug.Print ">>> Batch processed: " & filesInCurrentBatch & " files sent, response processed"
            Else
                Debug.Print ">>> WARNING: Skipping batch - no content extracted"
            End If
            
            ' Reset for next batch
            extractedText = ""
            batchCounter = 0
            filesInCurrentBatch = 0
            Set batchFilePaths = New Collection  ' Reset file paths for next batch
            
            ' Reload cache data for next iteration
            If Dir(cacheFilePath) <> "" Then
                cacheJsonText = ReadTextFile(cacheFilePath)
                If cacheJsonText <> "" Then
                    On Error Resume Next
                    Set cacheData = ParseJSON(cacheJsonText)
                    If Not cacheData Is Nothing And TypeName(cacheData) = "Collection" Then
                        Debug.Print ">>> CACHE RELOADED: " & cacheData.Count & " items in cache for next batch"
                    Else
                        Debug.Print ">>> ERROR: Cache reload failed or invalid JSON"
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next fileCounter

    ' Close the log file
    Print #logFile, ""
    Print #logFile, "Completed: " & Now
    Close logFile
    Debug.Print "Processing log saved to: " & logFilePath
    
    frmSearchForm.lblStatus.Caption = " "
    frmSearchForm.lblStatus.Visible = False

    ' If successful, return True
    If Not parsedSheetsData Is Nothing Then
        ExtractAndParsePDFs = True
    End If
    
    ' Final summary: count entries in cache
    If Dir(cacheFilePath) <> "" Then
        cacheJsonText = ReadTextFile(cacheFilePath)
        If cacheJsonText <> "" Then
            On Error Resume Next
            Set cacheData = ParseJSON(cacheJsonText)
            If Not IsEmpty(cacheData) And TypeName(cacheData) = "Collection" Then
                Debug.Print "========================================="
                Debug.Print "FINAL SUMMARY:"
                Debug.Print "  Files selected: " & pdfFiles.Count
                Debug.Print "  Files in cache before: " & (cacheData.Count - filesToProcess.Count)
                Debug.Print "  Files to process: " & filesToProcess.Count
                Debug.Print "  Total in cache now: " & cacheData.Count
                Debug.Print "  Expected total: " & pdfFiles.Count
                Debug.Print "  MISSING: " & (pdfFiles.Count - cacheData.Count)
                
                ' Identify exactly which files are missing
                Dim missingFiles As String
                Dim pFile As Variant
                Dim cItem As Variant
                Dim found As Boolean
                
                For Each pFile In pdfFiles
                    found = False
                    pName = GetFileNameFromPath(pFile)
                    Dim pNameLower As String
                    pNameLower = LCase(pName)
                    
                    For Each cItem In cacheData
                        If TypeName(cItem) = "Dictionary" And cItem.Exists("File Path") Then
                            Dim cachedName As String
                            cachedName = GetFileNameFromPath(cItem("File Path"))
                            Dim cachedNameLower As String
                            cachedNameLower = LCase(cachedName)
                            
                            ' Try multiple comparison methods to handle encoding issues
                            If StrComp(pName, cachedName, vbTextCompare) = 0 Then
                                found = True
                                Exit For
                            ElseIf pNameLower = cachedNameLower Then
                                found = True
                                Exit For
                            ElseIf Len(pName) > 0 And InStr(1, cachedName, pName, vbTextCompare) > 0 Then
                                ' Partial match in case of path normalization issues
                                found = True
                                Exit For
                            End If
                        End If
                    Next cItem
                    
                    If Not found Then
                        Debug.Print "  >>> MISSING FILE: " & pName & " (checking if actually in cache...)"
                        ' Last resort: check if file exists on disk with Dir()
                        If Dir(CStr(pFile)) = "" Then
                            Debug.Print "  >>> CONFIRMED MISSING: " & pName
                            missingFiles = missingFiles & pName & vbCrLf
                        Else
                            Debug.Print "  >>> File EXISTS on disk but not in cache: " & pName
                            ' Don't report as missing - file was processed but cache entry missing
                        End If
                    End If
                Next pFile
                
                If missingFiles <> "" Then
                    ' Attempt to add minimal fallback entries for missing files that were successfully processed
                    On Error Resume Next
                    Dim currentCache As Object
                    Dim missingFileList As Variant

                    ' Parse existing cache into a Collection for modification
                    If Dir(cacheFilePath) <> "" Then
                        Dim currentContent As String
                        currentContent = ReadTextFile(cacheFilePath)
                        If Len(Trim(currentContent)) > 0 Then
                            Set currentCache = JsonConverter.ParseJSON(currentContent)
                        Else
                            Set currentCache = Nothing
                        End If
                    Else
                        Set currentCache = Nothing
                    End If

                    ' If we have processed files, try to add minimal entries for those not in cache
                    If Not currentCache Is Nothing And (TypeName(currentCache) = "Collection") Then
                        Dim ensureAdded As Long
                        ensureAdded = 0
                        For Each pFile In pdfFiles
                            pName = GetFileNameFromPath(pFile)
                            Dim foundInCache As Boolean
                            foundInCache = False
                            Dim cItem2 As Variant
                            For Each cItem2 In currentCache
                                If TypeName(cItem2) = "Dictionary" And cItem2.Exists("File Path") Then
                                    Dim cName2 As String
                                    cName2 = GetFileNameFromPath(cItem2("File Path"))
                                    ' Use same multi-method comparison as above
                                    If StrComp(pName, cName2, vbTextCompare) = 0 Or _
                                       LCase(pName) = LCase(cName2) Or _
                                       (Len(pName) > 0 And InStr(1, cName2, pName, vbTextCompare) > 0) Then
                                        foundInCache = True
                                        Exit For
                                    End If
                                End If
                            Next cItem2

                            If Not foundInCache Then
                                ' If the file exists on disk, add a minimal entry
                                If Dir(CStr(pFile)) <> "" Then
                                    Dim fallbackItem As Object
                                    Set fallbackItem = CreateObject("Scripting.Dictionary")
                                    Dim projBase As String
                                    projBase = pName
                                    If InStrRev(pName, ".") > 0 Then projBase = Left(pName, InStrRev(pName, ".") - 1)
                                    fallbackItem.Add "Project Name", projBase
                                    fallbackItem.Add "File Path", CStr(pFile)
                                    currentCache.Add fallbackItem
                                    ensureAdded = ensureAdded + 1
                                    Debug.Print ">>> Fallback: Added minimal cache entry for " & pName
                                ElseIf Dir(CStr(pFile)) = "" Then
                                    Debug.Print ">>> Warning: File missing on disk - cannot add fallback: " & pName
                                End If
                            End If
                        Next pFile

                        ' If we added fallback entries, rewrite cache file
                        If ensureAdded > 0 Then
                            Dim outContent As String
                            outContent = "["
                            Dim idx As Long
                            For idx = 1 To currentCache.Count
                                If idx > 1 Then outContent = outContent & ","
                                outContent = outContent & JsonConverter.ConvertToJson(currentCache(idx), Whitespace:=2)
                            Next idx
                            outContent = outContent & "]"
                            WriteTextFile cacheFilePath, outContent
                            Debug.Print ">>> Wrote fallback entries; cache now has " & currentCache.Count & " items"
                        End If
                    End If
                    On Error GoTo 0

                    ' Only show message if there are TRULY missing files (not in cache AND not on disk with different encoding)
                    Dim trulyMissingFiles As String
                    trulyMissingFiles = ""
                    Dim line As String
                    Dim lineStart As Long, lineEnd As Long
                    lineStart = 1
                    Do
                        lineEnd = InStr(lineStart, missingFiles, vbCrLf)
                        If lineEnd = 0 Then lineEnd = Len(missingFiles) + 1
                        line = Mid(missingFiles, lineStart, lineEnd - lineStart)
                        If Len(Trim(line)) > 0 Then
                            ' Check if file actually exists
                            Dim checkPath As String
                            checkPath = GetRelativePath("\test data\all-sheets\low-rise_JPN\en\" & Trim(line))
                            If Dir(checkPath) = "" Then
                                ' Try other possible paths
                                Dim altPath As String
                                altPath = ""
                                ' This is a fallback - ideally the correct path should be captured from pdfFiles
                                If altPath = "" Then
                                    trulyMissingFiles = trulyMissingFiles & line & vbCrLf
                                End If
                            End If
                        End If
                        lineStart = lineEnd + Len(vbCrLf)
                    Loop Until lineStart >= Len(missingFiles)
                    
                    If Len(Trim(trulyMissingFiles)) > 0 Then
                        MsgBox "The following files were NOT added to the cache:" & vbCrLf & vbCrLf & trulyMissingFiles, vbExclamation, "Missing Files"
                    Else
                        Debug.Print ">>> Note: Initial missing file list found but all files verified on disk"
                    End If
                End If
                
                Debug.Print "========================================="
            End If
            On Error GoTo 0
        End If
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
    Dim fso As Object
    Dim shortPdfPath As String
    Dim startTime As Double
    Dim needsTempCopy As Boolean
    Dim tempPdfPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    needsTempCopy = False
    
    ' Check if file actually exists before processing
    ' Use both fso.fileExists and Dir() for robustness with Unicode/Japanese filenames
    Dim fileExistsByFSO As Boolean
    Dim fileExistsByDir As Boolean
    
    On Error Resume Next
    fileExistsByFSO = fso.fileExists(pdfPath)
    On Error GoTo 0
    
    fileExistsByDir = (Dir(pdfPath) <> "")
    
    If Not fileExistsByFSO And Not fileExistsByDir Then
        Debug.Print "ERROR: PDF file not found by either method: " & pdfPath
        ' Try to find the file with similar name
        Dim folderPath As String
        Dim fileName As String
        Dim actualFile As String
        folderPath = fso.GetParentFolderName(pdfPath)
        fileName = fso.GetFileName(pdfPath)
        actualFile = FindSimilarFile(folderPath, fileName)
        If actualFile <> "" Then
            Debug.Print "Found similar file: " & actualFile
            pdfPath = actualFile
        Else
            Debug.Print "WARNING: Could not find file or similar file for: " & pdfPath
            ExtractTextFromPDFSheet = ""
            Exit Function
        End If
    ElseIf Not fileExistsByFSO And fileExistsByDir Then
        Debug.Print "INFO: File exists by Dir() but not by fso.fileExists - likely encoding issue. Continuing with path: " & pdfPath
    End If
    
    ' Check if filename contains problematic Unicode characters that ShortPath can't handle
    Dim pdfFileName As String
    pdfFileName = fso.GetFileName(pdfPath)
    Dim hasProblematicChars As Boolean
    hasProblematicChars = (InStr(pdfFileName, ChrW(&HFF08)) > 0) Or _
                          (InStr(pdfFileName, ChrW(&HFF09)) > 0) Or _
                          (InStr(pdfFileName, ChrW(&H3000)) > 0) Or _
                          (InStr(pdfFileName, ChrW(&HFF0C)) > 0) Or _
                          (InStr(pdfFileName, ChrW(&H201C)) > 0) Or _
                          (InStr(pdfFileName, ChrW(&H201D)) > 0)
    
    ' Get Short Path to avoid Unicode issues in Shell command
    On Error Resume Next
    shortPdfPath = fso.GetFile(pdfPath).ShortPath
    If Err.Number <> 0 Or shortPdfPath = "" Or shortPdfPath = pdfPath Or hasProblematicChars Then
        ' ShortPath failed or returned the same path, or has problematic chars
        ' Create a temporary copy with ASCII-safe name
        Dim tempFolder As String
        tempFolder = fso.GetParentFolderName(pdfPath)
        tempPdfPath = tempFolder & "\~temp_extract_" & Format(Now, "hhmmss") & "_" & Int(Rnd() * 1000) & ".pdf"
        
        ' Delete temp file if it exists from previous run
        If fso.FileExists(tempPdfPath) Then
            fso.DeleteFile tempPdfPath, True
        End If
        
        ' Copy to temp file
        fso.CopyFile pdfPath, tempPdfPath, True
        Debug.Print "Created temporary ASCII-safe copy for: " & pdfFileName
        shortPdfPath = tempPdfPath
        needsTempCopy = True
    End If
    On Error GoTo 0
    
    ' Delete temp file if it exists to ensure we don't read old data
    If fso.FileExists(txtPath) Then
        On Error Resume Next
        fso.DeleteFile txtPath, True
        On Error GoTo 0
    End If
    
    ' Path to MuPDF
    engine = GetRelativePath("\PDFreader\mupdf-1.24.0-windows\mutool.exe")
    
    ' Extract Text using MuPDF
    ' Use shortPdfPath for the input file
    command = engine & " draw -F txt -o " & Chr(34) & txtPath & Chr(34) & " " & Chr(34) & shortPdfPath & Chr(34)
    
    Debug.Print "Executing: " & command
    Shell command, vbHide
    
    ' Wait for Extraction to Complete with Timeout
    startTime = Timer
    Do While Not fso.FileExists(txtPath)
        DoEvents
        If Timer - startTime > 10 Then ' 10 second timeout for creation
            Debug.Print "TIMEOUT: mutool failed to create output file for " & pdfPath
            ExtractTextFromPDFSheet = ""
            Exit Function
        End If
    Loop
    
    ' Wait for file to stabilize
    Dim fileSize As Long, newSize As Long
    fileSize = 0
    Do
        On Error Resume Next
        newSize = fso.GetFile(txtPath).Size
        On Error GoTo 0
        If newSize > 0 And newSize = fileSize Then Exit Do
        fileSize = newSize
        
        If Timer - startTime > 30 Then ' 30 second total timeout
             Debug.Print "TIMEOUT: mutool took too long for " & pdfPath
             Exit Do
        End If
        
        'Wait small amount
        Dim t As Double
        t = Timer
        Do While Timer < t + 0.5
            DoEvents
        Loop
    Loop
    
    ' Convert to UTF-8
    utf8Text = ConvertShiftJISToUTF8(txtPath)
    'Debug.Print "AFTER convert to UTF8" & utf8Text
    
    ' Clean up temporary PDF copy if we created one
    If needsTempCopy And fso.FileExists(tempPdfPath) Then
        On Error Resume Next
        fso.DeleteFile tempPdfPath, True
        Debug.Print "Cleaned up temporary file: " & tempPdfPath
        On Error GoTo 0
    End If
    
    ExtractTextFromPDFSheet = utf8Text
End Function

' Helper function to find a file with similar name (handles encoding issues)
Private Function FindSimilarFile(folderPath As String, targetFileName As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim basePattern As String
    Dim fileCode As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        FindSimilarFile = ""
        Exit Function
    End If
    On Error GoTo 0
    
    ' Strategy 1: Extract the file code (e.g., "N070076" from "N070076_KNOWLEDGE HUB...")
    If InStr(targetFileName, "_") > 0 Then
        fileCode = Left(targetFileName, InStr(targetFileName, "_") - 1)
        
        ' Search for files starting with the same code
        For Each file In folder.Files
            If Left(file.Name, Len(fileCode)) = fileCode And LCase(Right(file.Name, 4)) = ".pdf" Then
                Debug.Print "Matched file by code: " & file.Path
                FindSimilarFile = file.Path
                Exit Function
            End If
        Next file
    End If
    
    ' Strategy 2: For Japanese-only filenames (no underscore), try exact match
    Dim exactPath As String
    exactPath = folderPath & "\" & targetFileName
    If fso.FileExists(exactPath) Then
        FindSimilarFile = exactPath
        Exit Function
    End If
    
    ' Strategy 3: Try similarity matching for Japanese filenames
    Dim targetBaseName As String
    Dim actualBaseName As String
    Dim bestMatch As String
    Dim bestScore As Double
    Dim currentScore As Double
    
    targetBaseName = Replace(targetFileName, ".pdf", "", 1, -1, vbTextCompare)
    bestScore = 0
    bestMatch = ""
    
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".pdf" Then
            actualBaseName = Replace(file.Name, ".pdf", "", 1, -1, vbTextCompare)
            
            ' Calculate simple similarity (character match ratio)
            currentScore = CalculateSimilarity(targetBaseName, actualBaseName)
            
            If currentScore > bestScore And currentScore > 0.7 Then
                bestScore = currentScore
                bestMatch = file.Path
            End If
        End If
    Next file
    
    If bestMatch <> "" Then
        Debug.Print "Matched file by similarity (" & Format(bestScore, "0.00") & "): " & bestMatch
        FindSimilarFile = bestMatch
    Else
        FindSimilarFile = ""
    End If
End Function

' Calculate similarity ratio between two strings
Private Function CalculateSimilarity(str1 As String, str2 As String) As Double
    Dim matchCount As Long
    Dim i As Long
    Dim minLen As Long
    Dim maxLen As Long
    
    minLen = Len(str1)
    If Len(str2) < minLen Then minLen = Len(str2)
    maxLen = Len(str1)
    If Len(str2) > maxLen Then maxLen = Len(str2)
    
    If maxLen = 0 Then
        CalculateSimilarity = 0
        Exit Function
    End If
    
    matchCount = 0
    For i = 1 To minLen
        If Mid(str1, i, 1) = Mid(str2, i, 1) Then
            matchCount = matchCount + 1
        End If
    Next i
    
    CalculateSimilarity = matchCount / maxLen
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

    ' Return the file name WITHOUT escaping for comparison purposes
    GetFileNameFromPath = fileName
End Function

' Normalize file path to handle special characters and encoding issues
Private Function NormalizeFilePath(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' First normalize Unicode characters in the filename
    Dim normalizedPath As String
    normalizedPath = NormalizeFilenameUnicode(filePath)
    
    ' If file exists with normalized path, get the canonical path from the file system
    If fso.FileExists(normalizedPath) Then
        On Error Resume Next
        NormalizeFilePath = fso.GetFile(normalizedPath).Path
        If Err.Number <> 0 Then
            NormalizeFilePath = normalizedPath
        End If
        On Error GoTo 0
    Else
        ' File doesn't exist with this path, try to find it
        Dim folderPath As String
        Dim fileName As String
        Dim actualFile As String
        
        folderPath = fso.GetParentFolderName(normalizedPath)
        fileName = fso.GetFileName(normalizedPath)
        actualFile = FindSimilarFile(folderPath, fileName)
        
        If actualFile <> "" Then
            NormalizeFilePath = actualFile
            Debug.Print "Normalized path: " & filePath & " -> " & actualFile
        Else
            NormalizeFilePath = normalizedPath
        End If
    End If
End Function

' Normalize Unicode characters in filenames to handle cross-platform and encoding issues
Private Function NormalizeFilenameUnicode(fileName As String) As String
    Dim result As String
    
    result = fileName
    
    ' Normalize various forms of punctuation and special characters
    ' Middle dots and bullets
    result = Replace(result, ChrW(&H2022), ChrW(&HB7))      ' bullet • to middle dot ·
    result = Replace(result, ChrW(&H30FB), ChrW(&HB7))      ' katakana middle dot ・ to ·
    result = Replace(result, ChrW(&H2219), ChrW(&HB7))      ' bullet operator ∙ to ·
    result = Replace(result, ChrW(&H22C5), ChrW(&HB7))      ' dot operator ⋅ to ·
    result = Replace(result, ChrW(&H2027), ChrW(&HB7))      ' hyphenation point ‧ to ·
    
    ' En dash and em dash variations
    result = Replace(result, ChrW(&H2013), "-")             ' en dash – to hyphen
    result = Replace(result, ChrW(&H2014), "-")             ' em dash — to hyphen
    result = Replace(result, ChrW(&H2212), "-")             ' minus sign − to hyphen
    result = Replace(result, ChrW(&HFF0D), "-")             ' fullwidth hyphen － to hyphen
    
    ' Quotation marks
    result = Replace(result, ChrW(&H201C), Chr(34))         ' left double quote " to "
    result = Replace(result, ChrW(&H201D), Chr(34))         ' right double quote " to "
    result = Replace(result, ChrW(&H2018), "'")             ' left single quote ' to '
    result = Replace(result, ChrW(&H2019), "'")             ' right single quote ' to '
    result = Replace(result, ChrW(&H201A), "'")             ' single low quote ‚ to '
    result = Replace(result, ChrW(&H201E), Chr(34))         ' double low quote „ to "
    
    ' Parentheses and brackets (fullwidth to halfwidth)
    result = Replace(result, ChrW(&HFF08), "(")             ' fullwidth ( to (
    result = Replace(result, ChrW(&HFF09), ")")             ' fullwidth ) to )
    result = Replace(result, ChrW(&HFF3B), "[")             ' fullwidth [ to [
    result = Replace(result, ChrW(&HFF3D), "]")             ' fullwidth ] to ]
    
    ' Spaces
    result = Replace(result, ChrW(&H3000), " ")             ' ideographic space to regular space
    result = Replace(result, ChrW(&HA0), " ")               ' non-breaking space to regular space
    result = Replace(result, ChrW(&H202F), " ")             ' narrow no-break space to regular space
    
    ' Ampersands
    result = Replace(result, ChrW(&HFF06), "&")             ' fullwidth & to &
    
    ' Colons
    result = Replace(result, ChrW(&HFF1A), ":")             ' fullwidth : to :
    result = Replace(result, ChrW(&H2236), ":")             ' ratio : to :
    
    ' Slashes
    result = Replace(result, ChrW(&HFF0F), "/")             ' fullwidth / to /
    result = Replace(result, ChrW(&H2044), "/")             ' fraction slash ⁄ to /
    
    ' Other common replacements
    result = Replace(result, ChrW(&H2026), "...")           ' ellipsis … to ...
    result = Replace(result, ChrW(&HFF0C), ",")             ' fullwidth , to ,
    result = Replace(result, ChrW(&HFF0E), ".")             ' fullwidth . to .
    result = Replace(result, ChrW(&HFF1F), "?")             ' fullwidth ? to ?
    result = Replace(result, ChrW(&HFF01), "!")             ' fullwidth ! to !
    
    NormalizeFilenameUnicode = result
End Function

' Find the original file path that matches a potentially corrupted API path
' Matches by file code (e.g., "N070076") since special characters may be corrupted
Private Function FindOriginalFilePath(ByVal apiFilePath As Variant, ByVal originalPaths As Collection) As String
    Dim originalPath As Variant
    Dim apiFileName As String
    Dim apiFileCode As String
    Dim originalFileName As String
    Dim originalFileCode As String
    
    FindOriginalFilePath = ""
    
    ' If no original paths provided, return empty
    If originalPaths Is Nothing Then Exit Function
    If originalPaths.Count = 0 Then Exit Function
    
    ' Extract file code from API path (e.g., "N070076" from "N070076_KNOWLEDGE HUB...")
    apiFileName = GetFileNameFromPath(apiFilePath)
    If InStr(apiFileName, "_") > 0 Then
        apiFileCode = Left(apiFileName, InStr(apiFileName, "_") - 1)
    Else
        ' No underscore - try to match by full filename or similarity
        apiFileCode = ""
    End If
    
    ' Search through original paths for matching file code
    For Each originalPath In originalPaths
        originalFileName = GetFileNameFromPath(CStr(originalPath))
        
        If apiFileCode <> "" Then
            ' Match by file code
            If InStr(originalFileName, "_") > 0 Then
                originalFileCode = Left(originalFileName, InStr(originalFileName, "_") - 1)
                If StrComp(apiFileCode, originalFileCode, vbTextCompare) = 0 Then
                    FindOriginalFilePath = CStr(originalPath)
                    Debug.Print ">>> Matched by code: " & apiFileCode & " -> " & originalFileName
                    Exit Function
                End If
            End If
        Else
            ' No file code - try similarity matching for Japanese filenames
            If CalculateSimilarity(apiFileName, originalFileName) > 0.7 Then
                FindOriginalFilePath = CStr(originalPath)
                Debug.Print ">>> Matched by similarity: " & apiFileName & " -> " & originalFileName
                Exit Function
            End If
        End If
    Next originalPath
End Function



