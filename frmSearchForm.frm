VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   OleObjectBlob   =   "frmSearchForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Declare a module-level variable to track the clicked index
Private clickedIndex As Long



Public Sub InitializeSearchForm()
    Dim mDash As String
    mDash = Chr(8212)  ' Em-dash character
    
    Debug.Print "Initializing Search Form"
    
    ' Initialize all form controls with mDash
    Me.txtAreaMin.Value = mDash
    Me.txtAreaMax.Value = mDash
    Me.txtHeightMin.Value = mDash
    Me.txtHeightMax.Value = mDash
    Me.txtYearMin.Value = mDash
    Me.txtYearMax.Value = mDash
    
    ' Ensure the cache file exists and handle it
    If Not EnsureCacheFileExists Then
        PopulateListFromJsonCache
        Exit Sub ' Exit if the file was created (we don't populate dropdowns yet)
    End If
    
    ' Populate the dropdowns only after ensuring the cache file is handled
    PopulateDropdown frmSearchForm.cmbType, "Type"
    PopulateDropdown frmSearchForm.cmbCity, "City"
    PopulateDropdown frmSearchForm.cmbCountry, "Country"
    
    ' Populate the list (this doesn't require the cache check again)
    PopulateListFromJsonCache
    
End Sub


Private Function EnsureCacheFileExists() As Boolean
    Dim jsonFilePath As String
    Dim jsonText As String
    Dim fso As Object, jsonFile As Object
    Dim isFileCreated As Boolean

    jsonFilePath = GetRelativePath("\test output\cache.json")
    Set fso = CreateObject("Scripting.FileSystemObject")
    isFileCreated = False

    ' Check if the cache file exists, create if not
    If Not fso.fileExists(jsonFilePath) Then
        Set jsonFile = fso.CreateTextFile(jsonFilePath, True) ' Create new empty file
        jsonFile.Write "[  ]"  ' Write an empty JSON object
        jsonFile.Close
        MsgBox "Cache file not found. A new file has been created: " & jsonFilePath, vbExclamation, "File Created"
        isFileCreated = True
    End If

    ' Read the JSON file content
    Set jsonFile = fso.OpenTextFile(jsonFilePath, ForReading)
    jsonText = jsonFile.ReadAll
    jsonFile.Close

    ' Check if the file is empty or contains only an empty JSON object "{}"
    If Len(Trim(jsonText)) = 0 Or jsonText = "[]" Then
        ' If the file was empty, just return without showing a message
        If Not isFileCreated Then
            MsgBox "Cache file is empty. No data to populate.", vbInformation, "Empty Cache"
        End If
        EnsureCacheFileExists = False ' Return False to prevent dropdown population
        Exit Function
    End If

    ' File is valid
    EnsureCacheFileExists = True
End Function

Private Function GetKeysFromDictionary(dict As Variant) As String()
    Dim keys() As String
    Dim key As Variant
    Dim i As Integer
    
    ' Validate Dictionary type
    If TypeName(dict) <> "Dictionary" Then
        MsgBox "Expected Dictionary.", vbExclamation, "Type Error"
        Exit Function
    End If
    
    ' Initialize keys array
    ReDim keys(dict.Count - 1)
    i = 0
    For Each key In dict.keys
        keys(i) = CStr(key)
        i = i + 1
    Next key
    
    GetKeysFromDictionary = keys
End Function

Private Function CleanAndConvertValue(inputValue As String) As Double
    Dim cleanedValue As String
    cleanedValue = Replace(inputValue, ",", "")
    cleanedValue = Replace(cleanedValue, "m", "")
    cleanedValue = Replace(cleanedValue, Chr(8212), "")
    If IsNumeric(cleanedValue) Then
        CleanAndConvertValue = CDbl(cleanedValue)
    Else
        CleanAndConvertValue = 0
    End If
End Function

Private Function GetNumericValue(inputValue As String, defaultValue As Double) As Double
    If inputValue = Chr(8212) Or Not IsNumeric(inputValue) Then
        GetNumericValue = defaultValue
    Else
        GetNumericValue = CDbl(inputValue)
    End If
End Function



Private Sub cmdDel_Click()
    Dim i As Integer
    Dim itemCount As Integer
    itemCount = Me.lstResults.ListCount
    
    ' Loop through the ListBox from the last item to the first
    For i = itemCount - 1 To 0 Step -1
        ' Check if the item is selected for deletion
        If Me.lstResults.Selected(i) Then
            ' Remove the selected item
            Me.lstResults.RemoveItem i
        End If
    Next i
    
    ' Renumber the remaining items
    Dim updatedItem As String
    For i = 0 To Me.lstResults.ListCount - 1
        ' Extract the text without numbering, add new numbering, and update item
        updatedItem = Mid(Me.lstResults.List(i), InStr(Me.lstResults.List(i), ". ") + 2)
        Me.lstResults.List(i) = CStr(i + 1) & ". " & updatedItem
    Next i
End Sub



Private Sub cmdReset_Click()
    ' Call the ResetSearch routine
    ResetSearch
End Sub

Private Sub ResetSearch()
    ' Define the M-dash character
    Dim mDash As String
    mDash = Chr(8212)  ' Em-dash character
    
    ' Reset ComboBoxes to M-dash
    Me.cmbType.value = mDash
    Me.cmbCity.value = mDash
    Me.cmbCountry.value = mDash
    
    ' Set TextBoxes to M-dash for area, height, and year values
    Me.txtAreaMin.value = mDash
    Me.txtAreaMax.value = mDash
    Me.txtHeightMin.value = mDash
    Me.txtHeightMax.value = mDash
    Me.txtYearMin.value = mDash
    Me.txtYearMax.value = mDash
    
    Me.lstResults.Clear
    
    ' Print a message to the Immediate Window to confirm the reset
    Debug.Print "All controls have been reset to M-dash (�\)."
End Sub

Sub cmdSavePdf_Click_()
    Dim pdfFilePath As String
    Dim i As Integer
    Dim j As Integer
    Dim command As String
    Dim pdftkPath As String
    Dim outputPath As String
    Dim fileDialog As fileDialog
    Dim projectName As String
    Dim jsonText As String
    Dim fso As New FileSystemObject
    Dim jsonFile As textStream
    Dim cacheData As Object
    Dim item As Object
    Dim filePaths As New Collection
    Dim cacheFilePath As String
    Dim MyTime As String
    Dim ext As String
    Dim fileName As String
    
    MyTime = Format(Now, "ddmmyyyy_hhmmss")
    
    ' Path to pdftk.exe
    pdftkPath = GetRelativePath("\PdftkBuilder-4.1.6-portable\pdftk.exe")
    
    ' Path to cache file - adjust this to your actual cache file path
    cacheFilePath = GetRelativePath("\test output\cache.json")
    
    ' Read and parse the cache file
    If Not fso.fileExists(cacheFilePath) Then
        'MsgBox "2 Cache file not found: " & cacheFilePath, vbExclamation, "File Missing"
        Exit Sub
    End If
    
    Set jsonFile = fso.OpenTextFile(cacheFilePath, ForReading)
    jsonText = jsonFile.ReadAll
    jsonFile.Close
    
    ' Parse JSON data from cache
    Set cacheData = ParseJSON(jsonText)
    
    ' Initialize the command with the path to pdftk
    command = pdftkPath & " "
    
    ' Loop through each item in the ListBox and add all file paths to the collection
    For i = 0 To Me.lstResults.ListCount - 1
        ' Extract the project name from the ListBox item (remove the numbering)
        projectName = Trim(Mid(Me.lstResults.List(i), InStr(Me.lstResults.List(i), ". ") + 2))
        
        ' Find matching project in cache data
        Dim found As Boolean
        found = False
        
        For Each item In cacheData
            If item("Project Name") = projectName Then
                pdfFilePath = item("File Path")
                found = True
                ' Check if file exists before adding to the collection
                If Dir(pdfFilePath) <> "" Then
                    filePaths.Add pdfFilePath
                Else
                    MsgBox "File not found: " & pdfFilePath, vbExclamation, "File Missing"
                    Exit Sub
                End If
                Exit For
            End If
        Next item
        
        If Not found Then
            MsgBox "Project not found in cache: " & projectName, vbExclamation, "Missing Data"
            Exit Sub
        End If
    Next i
    
    ' Check if any files were found
    If filePaths.Count = 0 Then
        MsgBox "No files available to merge.", vbExclamation, "No Files"
        Exit Sub
    End If
    
    ' Build the command string with all file paths
    For i = 1 To filePaths.Count
        command = command & """" & filePaths(i) & """ "
    Next i
    
    ' Prompt user to choose the save location and file name
    Set fileDialog = Application.fileDialog(msoFileDialogSaveAs)
    ext = "*.pdf"
    fileName = "RelevantExperience" & "_" & MyTime
    For j = 1 To fileDialog.Filters.Count
'Debug.Print "fileDialog.Filters(" & j & ").Extensions: " & fileDialog.Filters(j).Extensions
        If fileDialog.Filters(j).Extensions = ext Then
            fileDialog.FilterIndex = j
            Exit For
        End If
    Next

    
    
    With fileDialog
        .title = "Save Merged PDF As"
        .AllowMultiSelect = False
        .InitialFileName = fileName  ' Default filename with .pdf extension
Debug.Print ".InitialFileName: " & .InitialFileName
Debug.Print ".FilterIndex* " & .FilterIndex

        If .Show = -1 Then
            outputPath = .selectedItems(1)
        Else
            MsgBox "Save operation cancelled.", vbExclamation, "Cancelled"
            Exit Sub
        End If
    End With
    ' Append the output command for pdftk with the chosen output path
    command = command & "cat output " & """" & outputPath & """"
    
    On Error Resume Next
    ' Execute the command
    Shell command, vbHide
    
    If Err.Number <> 0 Then
        MsgBox "Error merging PDFs: " & Err.Description, vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Notify user of success
    MsgBox "PDFs successfully merged to: " & outputPath, vbInformation, "Success"
End Sub

Sub cmdSavePdf_Click()
    Dim pdfFilePath As String
    Dim i As Integer
    Dim j As Integer
    Dim command As String
    Dim pdftkPath As String
    Dim outputPath As String
    Dim fileDialog As fileDialog
    Dim projectName As String
    Dim jsonText As String
    Dim fso As New FileSystemObject
    Dim jsonFile As textStream
    Dim cacheData As Object
    Dim item As Object
    Dim file As Object
    Dim filePaths As New Collection
    Dim cacheFilePath As String
    Dim MyTime As String
    Dim ext As String
    Dim fileName As String
    Dim matchedFile As String
    Dim confirmation As VbMsgBoxResult
    Dim similarityThreshold As Double

    MyTime = Format(Now, "ddmmyyyy_hhmmss")
    pdftkPath = GetRelativePath("\PdftkBuilder-4.1.6-portable\pdftk.exe")
    cacheFilePath = GetRelativePath("\test output\cache.json")

    If Not fso.fileExists(cacheFilePath) Then
        Exit Sub
    End If

    Set jsonFile = fso.OpenTextFile(cacheFilePath, ForReading)
    jsonText = jsonFile.ReadAll
    jsonFile.Close
    Set cacheData = ParseJSON(jsonText)

    command = pdftkPath & " "
    similarityThreshold = 0.8

    For i = 0 To Me.lstResults.ListCount - 1
        ' Only process selected (checked) items
        If Not Me.lstResults.Selected(i) Then
            GoTo NextItem
        End If
        
        projectName = Trim(Mid(Me.lstResults.List(i), InStr(Me.lstResults.List(i), ". ") + 2))
        Dim found As Boolean
        found = False

        For Each item In cacheData
'Debug.Print "item: " & CStr(item)
            If item("Project Name") = projectName Then
                pdfFilePath = item("File Path")
                found = True

                matchedFile = ""
                
                Debug.Print "pdfFilePath: " & pdfFilePath
                
                ' Check if the original file exists
                If Dir(pdfFilePath) <> "" Then
                    matchedFile = pdfFilePath
                    Debug.Print "Original file found: " & matchedFile
                Else
                    ' Try to find similar file in the same folder
                    On Error Resume Next
                    For Each file In fso.GetFolder(fso.GetParentFolderName(pdfFilePath)).Files
                        If SimilarityRatio(pdfFilePath, file.path) > similarityThreshold Then
                            matchedFile = file.path
                            Debug.Print "Similar file found: " & matchedFile
                            confirmation = MsgBox("Original file not found. A similar file '" & fso.GetFileName(matchedFile) & "' was found. Use this file?", vbYesNo + vbQuestion, "File Not Found")
                            If confirmation = vbYes Then
                                Exit For ' Exit the loop once confirmed
                            Else
                                matchedFile = ""
                            End If
                        End If
                    Next
                    On Error GoTo 0
                End If

                If matchedFile <> "" And Dir(matchedFile) <> "" Then
                    filePaths.Add matchedFile
                    Debug.Print "Added to merge list: " & matchedFile
                Else
                    MsgBox "File not found and no valid replacement: " & pdfFilePath, vbExclamation, "File Missing"
                    Exit Sub
                End If
                Exit For
            End If
        Next item

        If Not found Then
            MsgBox "Project not found in cache: " & projectName, vbExclamation, "Missing Data"
            Exit Sub
        End If
        
NextItem:
    Next i

    If filePaths.Count = 0 Then
        MsgBox "No files available to merge.", vbExclamation, "No Files"
        Exit Sub
    End If

    For i = 1 To filePaths.Count
        command = command & """" & filePaths(i) & """ "
Debug.Print "command #" & i & "; "; command
    Next i

    Set fileDialog = Application.fileDialog(msoFileDialogSaveAs)
    ext = "*.pdf"
    fileName = "RelevantExperience" & "_" & MyTime

    For j = 1 To fileDialog.Filters.Count
        If fileDialog.Filters(j).Extensions = ext Then
            fileDialog.FilterIndex = j
            Exit For
        End If
    Next

    With fileDialog
        .title = "Save Merged PDF As"
        .AllowMultiSelect = False
        .InitialFileName = fileName

        If .Show = -1 Then
            outputPath = .selectedItems(1)
        Else
            MsgBox "Save operation cancelled.", vbExclamation, "Cancelled"
            Exit Sub
        End If
    End With

    command = command & "cat output " & """" & outputPath & """"
Debug.Print "Command: " & command
    On Error Resume Next
    Shell command, vbHide

    If Err.Number <> 0 Then
        MsgBox "Error merging PDFs: " & Err.Description, vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    MsgBox "PDFs successfully merged to: " & outputPath, vbInformation, "Success"
End Sub



Function SimilarityRatio(str1 As String, str2 As String) As Double
    Dim lengthDiff As Double
    Dim matchCount As Double
    Dim i As Integer

    matchCount = 0
    For i = 1 To Application.Min(Len(str1), Len(str2))
        If Mid(str1, i, 1) = Mid(str2, i, 1) Then
            matchCount = matchCount + 1
        End If
    Next i

    SimilarityRatio = matchCount / Application.Max(Len(str1), Len(str2))
End Function



Private Sub cmdSearch_Click()
    ' Call the SearchProjects routine
    SearchProjects
End Sub

' Event handler for cmdExit: Exits the program
Private Sub cmdSExit_Click()
    ' Confirm exit with the user
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit Program")
    
    If response = vbYes Then
        ' Instead of Unload Me, use Hide to prevent the application from closing
        ResetSearch
        Me.Hide
    End If
End Sub

Private Sub SearchProjects()
    On Error GoTo ErrorHandler

    Dim typeFilter As String, cityFilter As String, countryFilter As String
    Dim areaMinFilter As Double, areaMaxFilter As Double
    Dim heightMinFilter As Double, heightMaxFilter As Double
    Dim yearMinFilter As Integer, yearMaxFilter As Integer
    Dim projectItem As Object
    Dim jsonData As Object
    Dim jsonFilePath As String, jsonText As String
    Dim fso As Object, jsonFile As Object
    Dim matchFound As Boolean
    Dim resultCount As Integer
    Dim projectName As String

    ' Define M-dash used as a default value in all inputs
    Dim mDash As String
    mDash = Chr(8212)  ' Em-dash character

    ' Define the path to the JSON cache file
    jsonFilePath = GetRelativePath("\test output\cache.json")

    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if the JSON cache file exists
    If Not fso.fileExists(jsonFilePath) Then
        'MsgBox "3 Cache file not found: " & jsonFilePath, vbExclamation, "File Missing"
        Exit Sub
    End If

    ' Read the JSON file content
    Set jsonFile = fso.OpenTextFile(jsonFilePath, ForReading)
    jsonText = jsonFile.ReadAll
    jsonFile.Close

    ' Parse the JSON data using JsonConverter
    Set jsonData = JsonConverter.ParseJSON(jsonText)

    ' Validate the parsed JSON data
    If jsonData Is Nothing Or TypeName(jsonData) <> "Collection" Then
        MsgBox "Invalid JSON structure. Expected a Collection.", vbExclamation, "Data Error"
        Exit Sub
    End If

    ' Retrieve filter values from the form
    typeFilter = Me.cmbType.value
    cityFilter = Me.cmbCity.value
    countryFilter = Me.cmbCountry.value
    areaMinFilter = GetNumericValue(Me.txtAreaMin.value, 0)
    areaMaxFilter = GetNumericValue(Me.txtAreaMax.value, 9999999)
    heightMinFilter = GetNumericValue(Me.txtHeightMin.value, 0)
    heightMaxFilter = GetNumericValue(Me.txtHeightMax.value, 9999)
    yearMinFilter = GetNumericValue(Me.txtYearMin.value, 0)
    yearMaxFilter = GetNumericValue(Me.txtYearMax.value, 9999)

    ' Clear the results list before displaying new results
    Me.lstResults.Clear
    Me.lstResults.Visible = True
    Me.cmdSavePdf.Visible = True

    ' Initialize the counter for search results
    resultCount = 0

    ' Loop through each project in the parsed JSON data
    For Each projectItem In jsonData
        matchFound = True

        ' Get the project name
        If projectItem.Exists("Project Name") Then
            projectName = projectItem("Project Name")
        Else
            GoTo ContinueLoop
        End If

        ' Check Type filter, allowing partial matches
        If typeFilter <> mDash Then
            matchFound = MultiParamMatch(projectItem("Type"), typeFilter)
        End If
        
        ' Check City filter, allowing partial matches or matching any Location
        If matchFound And cityFilter <> mDash Then
            If projectItem.Exists("Location") Then
                ' Location is an array, so we need to check each element
                Dim locationMatch As Boolean
                locationMatch = False
                
                If TypeName(projectItem("Location")) = "Collection" Then
                    Dim locItem As Variant
                    For Each locItem In projectItem("Location")
                        If InStr(1, CStr(locItem), cityFilter, vbTextCompare) > 0 Then
                            locationMatch = True
                            Exit For
                        End If
                    Next locItem
                Else
                    ' Fallback if Location is a string
                    locationMatch = InStr(1, CStr(projectItem("Location")), cityFilter, vbTextCompare) > 0
                End If
                
                matchFound = locationMatch
            Else
                matchFound = MultiParamMatch(projectItem("City"), cityFilter)
            End If
        End If

        ' Check City filter, allowing partial matches
        'If matchFound And cityFilter <> mDash Then
        '    matchFound = MultiParamMatch(projectItem("City"), cityFilter)
        'End If

        ' Check Country filter with explicit distinction between "China" and "Republic of China"
        If matchFound And countryFilter <> mDash Then
            If countryFilter = "China" And InStr(1, projectItem("Country"), "Republic of China", vbTextCompare) > 0 Then
                matchFound = False
            ElseIf countryFilter = "Republic of China" And InStr(1, projectItem("Country"), "China", vbTextCompare) > 0 And _
                   Not InStr(1, projectItem("Country"), "Republic of China", vbTextCompare) > 0 Then
                matchFound = False
            Else
                matchFound = MultiParamMatch(projectItem("Country"), countryFilter)
            End If
        End If

        ' Area filter
        If matchFound And projectItem.Exists("Total Floor Area") Then
            If CleanAndConvertValue(projectItem("Total Floor Area")) < areaMinFilter Or _
               CleanAndConvertValue(projectItem("Total Floor Area")) > areaMaxFilter Then
                matchFound = False
            End If
        End If

        ' Height filter
        If matchFound And projectItem.Exists("Building Height") Then
            If CleanAndConvertValue(projectItem("Building Height")) < heightMinFilter Or _
               CleanAndConvertValue(projectItem("Building Height")) > heightMaxFilter Then
                matchFound = False
            End If
        End If

        ' Year filter
        If matchFound And projectItem.Exists("Completion Year") Then
            If val(projectItem("Completion Year")) < yearMinFilter Or _
               val(projectItem("Completion Year")) > yearMaxFilter Then
                matchFound = False
            End If
        End If

        ' If all criteria match, add the project to the results list
        If matchFound Then
            resultCount = resultCount + 1
            Me.lstResults.AddItem CStr(resultCount) & ". " & projectName
        End If

ContinueLoop:
    Next projectItem

    ' Display message if no results found
    If resultCount = 0 Then
        Me.lstResults.AddItem "No matching projects found."
    End If

    ' Debug message for results
    Debug.Print "Search completed. " & resultCount & " matching project(s) found."
    Exit Sub

ErrorHandler:
    MsgBox "Error in SearchProjects: " & Err.Description, vbCritical, "Error"
End Sub

' Function to check if any part of a search parameter matches any part of a grouped value
Private Function MultiParamMatch(groupedValue As String, filter As String) As Boolean
    Dim part As Variant
    Dim searchPart As Variant
    Dim splitGroup() As String
    Dim intermediateFilter() As String
    Dim splitFilter() As String
    Dim tempList As Collection
    Dim i As Integer

    MultiParamMatch = False

    ' Split the grouped value by commas
    splitGroup = Split(groupedValue, ",")

    ' Step 1: Split the filter by "/" to handle multiple terms separated by "/"
    intermediateFilter = Split(filter, "/")

    ' Step 2: Further split each part of intermediateFilter by commas
    Set tempList = New Collection
    For Each searchPart In intermediateFilter
        Dim subParts() As String
        subParts = Split(searchPart, ",") ' Split each part by comma
        For i = LBound(subParts) To UBound(subParts)
            tempList.Add Trim(subParts(i)) ' Trim and add to the collection
        Next i
    Next searchPart

    ' Convert tempList collection to an array for easier iteration
    ReDim splitFilter(1 To tempList.Count)
    For i = 1 To tempList.Count
        splitFilter(i) = tempList(i)
    Next i

    ' Loop through each part of the filter and each part of the grouped value
    For Each searchPart In splitFilter
        For Each part In splitGroup
            'Debug.Print ("!!! part: " & part)
            ' Trim spaces and check if there's a partial match
            If InStr(1, Trim(part), Trim(searchPart), vbTextCompare) > 0 Then
                MultiParamMatch = True
                Exit Function
            End If
        Next part
    Next searchPart
End Function

' Function to convert values to numeric after removing non-numeric characters
Private Function CleanAndConvertValue_(inputValue As String) As Double
    Dim cleanedValue As String
    cleanedValue = inputValue
    ' Remove any non-numeric characters except for "." (period) and "-"
    cleanedValue = Replace(cleanedValue, ",", "")  ' Remove commas
    cleanedValue = Replace(cleanedValue, "m", "")  ' Remove "m" (if it appears as in "500m")
    cleanedValue = Replace(cleanedValue, Chr(8212), "")  ' Remove M-dash if present
    ' Return the cleaned value as a double
    If IsNumeric(cleanedValue) Then
        CleanAndConvertValue_ = CDbl(cleanedValue)
    Else
        CleanAndConvertValue_ = 0  ' Default to zero if not numeric
    End If
End Function

' Function to retrieve numeric value or return default value if not applicable
Private Function GetNumericValue_(inputValue As String, defaultValue As Double) As Double
    If inputValue = Chr(8212) Or Not IsNumeric(inputValue) Then
        GetNumericValue_ = defaultValue
    Else
        GetNumericValue_ = CDbl(inputValue)
    End If
End Function



Private Sub AdjustListBoxScrollbars(lst As MSForms.ListBox)
    ' Ensure the listbox has scroll bars if the content exceeds its size
    With lst
        .ColumnCount = 1
        .ColumnWidths = "350"  ' Adjust width as per content
        .IntegralHeight = False ' Allows for horizontal scroll
Debug.Print "IntegralHeight = False "
    End With
End Sub



' MouseDown: compute which row was clicked
Private Sub lstResults_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim topIndex As Long
    Dim rowHeight As Single

    If Me.lstResults.ListCount = 0 Then
        clickedIndex = -1
        Exit Sub
    End If

    ' Estimate row height from font size; if zero, fall back to 12
    rowHeight = Me.lstResults.Font.Size
    If rowHeight <= 0 Then rowHeight = 12

    topIndex = Me.lstResults.topIndex
    ' NB: "\" is integer division (your locale may show the yen glyph)
    clickedIndex = topIndex + CLng(Y \ rowHeight)

    ' Clamp to valid range
    If clickedIndex < 0 Or clickedIndex >= Me.lstResults.ListCount Then
        clickedIndex = -1
    End If
End Sub

' Click: select the computed row (if valid)
Private Sub lstResults_Click()
    If clickedIndex >= 0 And clickedIndex < Me.lstResults.ListCount Then
        Me.lstResults.ListIndex = clickedIndex
        ' TODO: handle selection here (open, load, etc.)
    Else
        ' No valid row under the click?do nothing or clear selection:
        ' Me.lstResults.ListIndex = -1
    End If
End Sub






' Helper function to find or create a group key based on similarity
Private Function FindOrCreateGroup(ByVal typeGroups As Object, ByVal value As String) As String
    Dim existingKey As Variant
    'Debug.Print "FindOrCreateGroup called for value: " & value

    For Each existingKey In typeGroups.keys
        'Debug.Print "existingKey: " & existingKey

        ' Explicitly prevent grouping "China" and "Republic of China"
        If (existingKey = "China" And InStr(value, "Republic of China")) Or (InStr(existingKey, "Republic of China") And value = "China") Then
        'If existingKey = "China" Or existingKey = "Republic of China" Then
            ' Do not group these two; move to the next iteration
            Debug.Print "Keeping 'China' and 'Republic of China' separate"
            GoTo ContinueLoop
        End If
        
        ' General grouping condition (ignoring the special case above)
        If InStr(1, existingKey, value, vbTextCompare) > 0 Or InStr(1, value, existingKey, vbTextCompare) > 0 Then
            'Debug.Print "Grouping " & value & " under existing key: " & existingKey
            FindOrCreateGroup = existingKey
            Exit Function
        End If
ContinueLoop:
    Next existingKey

    ' If no similar group is found, create a new group with this value
    'Debug.Print "Creating new group for value: " & value
    typeGroups.Add CStr(value), CreateObject("Scripting.Dictionary")
    FindOrCreateGroup = value
End Function


' Helper function to sort dictionary keys by length
Private Function SortValuesByLength(valuesDict As Object) As Collection
    Dim sortedList As New Collection
    Dim tempArray() As Variant
    Dim key As Variant, i As Integer

    ' Populate temporary array with dictionary keys
    ReDim tempArray(0 To valuesDict.Count - 1)
    i = 0
    For Each key In valuesDict.keys
        tempArray(i) = CStr(key)  ' Ensure key is treated as String
        i = i + 1
    Next key

    ' Sort the array by length
    QuickSortByLength tempArray, LBound(tempArray), UBound(tempArray)

    ' Add sorted items to collection
    For i = LBound(tempArray) To UBound(tempArray)
        sortedList.Add tempArray(i)
    Next i

    Set SortValuesByLength = sortedList
End Function

' QuickSort algorithm to sort by string length
Private Sub QuickSortByLength(arr() As Variant, low As Integer, high As Integer)
    Dim pivot As Variant, temp As Variant
    Dim i As Integer, j As Integer
    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        Do While i <= j
            Do While Len(arr(i)) < Len(pivot)
                i = i + 1
            Loop
            Do While Len(arr(j)) > Len(pivot)
                j = j - 1
            Loop
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        QuickSortByLength arr, low, j
        QuickSortByLength arr, i, high
    End If
End Sub


Private Sub cmdAdd_Click()

    Dim fd As fileDialog
    Dim selectedItems As Variant
    Dim selectedFiles As Collection
    Dim fileName As Variant
    Dim selectedFolderPath As String
    Dim pdfFiles As Collection
    Dim fileCount As Integer
    Dim cacheFilePath As String
    
    
    ' Initialize the collection to store selected PDF files
    Set selectedFiles = New Collection
    
    cacheFilePath = GetRelativePath("\test output\cache.json")

    ' Create a FileDialog object to select one or more PDF files
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    fd.title = "Select One or More PDF Files"
    fd.Filters.Add "PDF Files", "*.pdf", 1
    fd.AllowMultiSelect = True

    ' Show the dialog and capture the selected files
    If fd.Show = -1 Then
        ' Count the selected files
        fileCount = fd.selectedItems.Count
        
        ' If files are selected, add them to the collection
        If fileCount > 0 Then
            For Each fileName In fd.selectedItems
                selectedFiles.Add fileName
            Next fileName

            ' Use the selected files for processing
            Set pdfFiles = selectedFiles
        Else
            ' If no specific files are selected, select the folder path and process all PDFs in the folder
            'selectedFolderPath = fd.selectedItems(1)
            'Set pdfFiles = GetSortedPDFFiles(Left(selectedFolderPath, InStrRev(selectedFolderPath, "\")))
        End If

        ' Display the number of files selected
        MsgBox "Number of PDF files selected: " & pdfFiles.Count, vbInformation, "Files Selected"
        
        ' Step 1: Extract and parse data from the selected files
        If ExtractAndParsePDFs(pdfFiles, cacheFilePath) Then
            Debug.Print "1 TRUE ExtractAndParsePDFs(selectedFolderPath)"
            
            ' Step 2: Open frmSearchForm and populate dropdowns
            Debug.Print "2 TRUE ExtractAndParsePDFs(selectedFolderPath)"
           
            ' Call the custom initialization routine
            Call frmSearchForm.InitializeSearchForm
            Debug.Print "3 InitializeSearchForm"

        End If
    Else
        MsgBox "No files selected.", vbExclamation, "File Selection"
    End If

    ' Cleanup
    Set fd = Nothing
    Set selectedFiles = Nothing
    Set pdfFiles = Nothing

End Sub


Function GetSortedPDFFiles(folderPath As String) As Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim pdfFiles As New Collection
    Dim fileNames As New Collection
    Dim i As Long, j As Long
    Dim temp As String

    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get the folder object
    Set folder = fso.GetFolder(folderPath)

    ' Loop through each file in the folder
    For Each file In folder.Files
        ' Check if the file has a .pdf extension
        If LCase(Right(file.name, 4)) = ".pdf" Then
            ' Add the file name (with full path) to the collection
            fileNames.Add file.path
        End If
    Next file

    ' Sort the fileNames collection alphabetically
    For i = 1 To fileNames.Count - 1
        For j = i + 1 To fileNames.Count
            If fileNames(i) > fileNames(j) Then
                ' Swap fileNames(i) and fileNames(j)
                temp = fileNames(i)
                fileNames(i) = fileNames(j)
                fileNames(j) = temp
            End If
        Next j
    Next i

    ' Add the sorted file names to the pdfFiles collection
    For i = 1 To fileNames.Count
        pdfFiles.Add fileNames(i)
    Next i

    ' Return the sorted PDF files collection
    Set GetSortedPDFFiles = pdfFiles
End Function

Private Sub PopulateDropdown(cmb As MSForms.ComboBox, key As String)
    On Error GoTo ErrorHandler
    Dim uniqueValues As Collection
    Dim projectItem As Object
    Dim value As Variant
    Dim splitValue As Variant
    Dim itemKey As Variant
    Dim mDash As String
    Dim typeGroups As Object ' Dictionary to hold grouped types
    Dim groupedKey As Variant
    Dim jsonData As Object
    Dim jsonFilePath As String
    Dim jsonText As String
    Dim fso As Object, jsonFile As Object

    mDash = Chr(8212)  ' Em-dash character

    ' Initialize collections
    Set uniqueValues = New Collection
    Set typeGroups = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Define the path to the JSON cache file
    jsonFilePath = GetRelativePath("\test output\cache.json")
Debug.Print "jsonFilePath: " & jsonFilePath

    ' Check if the cache file exists
    If Not fso.fileExists(jsonFilePath) Then
        MsgBox "0 Cache file not found: " & jsonFilePath, vbExclamation, "File Missing"
        Exit Sub
    End If

    ' Read the JSON file content
    Set jsonFile = fso.OpenTextFile(jsonFilePath, ForReading)
'Debug.Print "Set jsonFile = fs "

    jsonText = jsonFile.ReadAll
'Debug.Print "jsonText: " & jsonText

    jsonFile.Close

    ' Parse the JSON data
    Set jsonData = JsonConverter.ParseJSON(jsonText)

    ' Validate jsonData as a Collection
    If jsonData Is Nothing Or TypeName(jsonData) <> "Collection" Then
        MsgBox "Invalid JSON structure. Expected a Collection.", vbExclamation, "Data Error"
        Exit Sub
    End If

Debug.Print "Populating dropdown for key: " & key

    ' Loop through each project in the JSON data
    For Each projectItem In jsonData
        ' Check if the required key exists in the project item
        If projectItem.Exists(key) Then
            value = projectItem(key)

'Debug.Print "value = : " & value
            ' Check if value is non-empty
            If Len(value) = 0 Then
                Debug.Print "Empty value for key: " & key
                GoTo ContinueLoop
            End If

            ' Split values by commas and trim each part
            For Each splitValue In Split(value, ",")
                splitValue = Trim(CStr(splitValue))
                'Debug.Print "Processing value: " & splitValue

                ' Use FindOrCreateGroup to determine the grouping key
                groupedKey = FindOrCreateGroup(typeGroups, splitValue)

                ' Ensure dictionary entry exists for groupedKey before adding
                If Not typeGroups.Exists(groupedKey) Then
                    Set typeGroups(groupedKey) = CreateObject("Scripting.Dictionary")
                End If

                ' Add splitValue to the appropriate group, avoiding duplicates
                If Not typeGroups(groupedKey).Exists(CStr(splitValue)) Then
                    typeGroups(groupedKey).Add CStr(splitValue), Len(CStr(splitValue))
                End If
            Next splitValue
        End If
ContinueLoop:
    Next projectItem

    ' Populate ComboBox with grouped and sorted values
    cmb.Clear
    cmb.AddItem mDash  ' Set M-dash as the initial value

    Dim sortedValues As Collection
    Dim finalValue As String
    Dim sortedArray() As String

    ' Iterate over each group in typeGroups
    For Each groupedKey In typeGroups.keys
        Set sortedValues = SortValuesByLength(typeGroups(groupedKey))

        ' Convert sortedValues (Collection) to an array for Join
        sortedArray = CollectionToArray(sortedValues)
        finalValue = Join(sortedArray, " / ")

        ' Add to ComboBox
        'Debug.Print "Adding to ComboBox: " & finalValue
        cmb.AddItem CStr(finalValue)
    Next groupedKey

    ' Set the M-dash as the default selected value
    cmb.value = mDash
    Debug.Print "Finished populating dropdown for key: " & key
    Exit Sub

ErrorHandler:
    MsgBox "Error in PopulateDropdown: " & Err.Description, vbCritical, "Error"
End Sub



' Helper function to convert a Collection to an array
Private Function CollectionToArray(col As Collection) As String()
    Dim arr() As String
    Dim i As Long

    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = CStr(col(i))
    Next i

    CollectionToArray = arr
End Function


Sub PopulateListFromJsonCache()
    On Error GoTo ErrorHandler
    Dim jsonData As Object
    Dim jsonFilePath As String
    Dim jsonText As String
    Dim fso As Object, jsonFile As Object
    Dim projectItem As Object
    Dim projectName As String
    Dim resultCount As Long

    ' Define the path to the JSON cache file
    jsonFilePath = GetRelativePath("\test output\cache.json")
    
    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if the JSON cache file exists
    If Not fso.fileExists(jsonFilePath) Then
        'MsgBox "1 Cache file not found: " & jsonFilePath, vbExclamation, "File Missing"
        Exit Sub
    End If

    ' Read the JSON file content
    Set jsonFile = fso.OpenTextFile(jsonFilePath, ForReading)
    jsonText = jsonFile.ReadAll
    jsonFile.Close

    ' Parse the JSON data
    Set jsonData = JsonConverter.ParseJSON(jsonText)

    ' Validate the parsed JSON data
    If jsonData Is Nothing Or TypeName(jsonData) <> "Collection" Then
        MsgBox "Invalid JSON structure. Expected a Collection.", vbExclamation, "Data Error"
        Exit Sub
    End If

    ' Clear the lstResults ListBox before populating
    Me.lstResults.Clear

    ' Initialize result counter
    resultCount = 0

    ' Loop through each project in the JSON data
    For Each projectItem In jsonData
        ' Get the project name
        If projectItem.Exists("Project Name") Then
            projectName = projectItem("Project Name")

            ' Increment the result counter
            resultCount = resultCount + 1

            ' Add the project name to the lstResults ListBox
            Me.lstResults.AddItem CStr(resultCount) & ". " & projectName
        End If
    Next projectItem

    ' Display a message if no projects were found
    If resultCount = 0 Then
        Me.lstResults.AddItem "No projects found in the cache file."
    End If

    Debug.Print "List populated from JSON cache. Total projects: " & resultCount
    Exit Sub

ErrorHandler:
    MsgBox "Error in PopulateListFromJsonCache: " & Err.Description, vbCritical, "Error"
End Sub






