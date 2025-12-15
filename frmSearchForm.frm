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

Private Const DEBUG_UI As Boolean = True
' Declare a module-level variable to track the clicked index
Private clickedIndex As Long
Private typeClickedIndex As Long
Private cityClickedIndex As Long
Private countryClickedIndex As Long
Private isInitialized As Boolean
Private isUpdatingCombo As Boolean  ' Flag to prevent recursive Click events
Private isRemovingSelection As Boolean  ' Flag to prevent recursive ListBox click events

' Variables to track checkbox state changes
' (No longer used with simplified logic)

' Module-level collections to store multi-select values for each dropdown
Private selectedTypes As Collection
Private selectedCities As Collection
Private selectedCountries As Collection

' Deferred UI work (avoid ListBox mutations during mouse events)
Private deferredScheduled As Boolean
Private pendingTypeRemovals As Collection
Private pendingCityRemovals As Collection
Private pendingCountryRemovals As Collection

Private Sub Dbg(ByVal msg As String)
    If DEBUG_UI Then Debug.Print Format$(Now, "hh:nn:ss.000") & " | " & msg
End Sub

Private Sub ScheduleDeferredUIWork()
    On Error Resume Next
    If deferredScheduled Then Exit Sub
    deferredScheduled = True

    ' NOTE: frmSearchForm is shown modally, so Application.OnTime won't run.
    ' Use a WinAPI timer instead (see MainModule).
    ScheduleDeferredSearchFormWork 50
    If Err.Number <> 0 Then
        Dbg "ScheduleDeferredUIWork ERROR: " & Err.Number & " - " & Err.Description
        Err.Clear
        deferredScheduled = False
    Else
        Dbg "ScheduleDeferredUIWork: timer scheduled"
    End If
End Sub

Private Sub CancelDeferredUIWork()
    On Error Resume Next
    CancelDeferredSearchFormWork
    deferredScheduled = False
End Sub

Public Sub ProcessDeferredUIWork()
    Dim j As Long
    Dim v As Variant
    Dim didType As Boolean
    Dim didCity As Boolean
    Dim didCountry As Boolean

    On Error GoTo Cleanup
    deferredScheduled = False
    If Not Me.Visible Then Exit Sub

    If Not pendingTypeRemovals Is Nothing Then
        For Each v In pendingTypeRemovals
            Dbg "Deferred: Type remove '" & CStr(v) & "'"
            For j = selectedTypes.Count To 1 Step -1
                If selectedTypes(j) = CStr(v) Then selectedTypes.Remove j
            Next j
            didType = True
        Next v
        Set pendingTypeRemovals = New Collection
        If didType Then
            UpdateComboDisplay Me.cmbType, selectedTypes
            RefreshSelectedTypesList
        End If
    End If

    If Not pendingCityRemovals Is Nothing Then
        For Each v In pendingCityRemovals
            Dbg "Deferred: City remove '" & CStr(v) & "'"
            For j = selectedCities.Count To 1 Step -1
                If selectedCities(j) = CStr(v) Then selectedCities.Remove j
            Next j
            didCity = True
        Next v
        Set pendingCityRemovals = New Collection
        If didCity Then
            UpdateComboDisplay Me.cmbCity, selectedCities
            RefreshSelectedCitiesList
        End If
    End If

    If Not pendingCountryRemovals Is Nothing Then
        For Each v In pendingCountryRemovals
            Dbg "Deferred: Country remove '" & CStr(v) & "'"
            For j = selectedCountries.Count To 1 Step -1
                If selectedCountries(j) = CStr(v) Then selectedCountries.Remove j
            Next j
            didCountry = True
        Next v
        Set pendingCountryRemovals = New Collection
        If didCountry Then
            UpdateComboDisplay Me.cmbCountry, selectedCountries
            RefreshSelectedCountriesList
        End If
    End If

    Exit Sub

Cleanup:
    Dbg "ProcessDeferredUIWork ERROR: " & Err.Number & " - " & Err.Description
    deferredScheduled = False
End Sub

' Ensure the form initializes itself when opened directly
Private Sub UserForm_Initialize()
    ' Always initialize multi-select collections
    Set selectedTypes = New Collection
    Set selectedCities = New Collection
    Set selectedCountries = New Collection
    
    ' Configure selected-items ListBoxes for checkbox style
    On Error Resume Next
    ' Set ListStyle to show checkboxes (fmListStyleOption = 1)
    Me.lstTypeSelected.ListStyle = fmListStyleOption
    Me.lstCitySelected.ListStyle = fmListStyleOption
    Me.lstCountrySelected.ListStyle = fmListStyleOption
    ' Set MultiSelect to allow multiple selections with checkboxes
    Me.lstTypeSelected.MultiSelect = fmMultiSelectMulti
    Me.lstCitySelected.MultiSelect = fmMultiSelectMulti
    Me.lstCountrySelected.MultiSelect = fmMultiSelectMulti
    
    ' Keep checkbox listboxes scrollable without resizing the form layout.
    ' IntegralHeight=True can round the control height down (making fewer visible rows).
    Me.lstTypeSelected.IntegralHeight = False
    Me.lstCitySelected.IntegralHeight = False
    Me.lstCountrySelected.IntegralHeight = False

    ' City listbox must behave identically to Type listbox (scroll range, row height).
    ' Under some DPI settings MSForms computes scroll bounds differently if Height/Font differ even slightly.
    SyncListBoxToTemplate Me.lstCitySelected, Me.lstTypeSelected
    SyncListBoxToTemplate Me.lstCountrySelected, Me.lstTypeSelected

    typeClickedIndex = -1
    cityClickedIndex = -1
    countryClickedIndex = -1

    If pendingTypeRemovals Is Nothing Then Set pendingTypeRemovals = New Collection
    If pendingCityRemovals Is Nothing Then Set pendingCityRemovals = New Collection
    If pendingCountryRemovals Is Nothing Then Set pendingCountryRemovals = New Collection
    
    ' Set tooltips
    Me.lstTypeSelected.ControlTipText = "Uncheck to remove"
    Me.lstCitySelected.ControlTipText = "Uncheck to remove"
    Me.lstCountrySelected.ControlTipText = "Uncheck to remove"
    On Error GoTo 0
    
    If Not isInitialized Then
        On Error Resume Next
        InitializeSearchForm
        On Error GoTo 0
    End If
End Sub

Private Sub SyncListBoxToTemplate(ByRef target As MSForms.ListBox, ByRef template As MSForms.ListBox)
    On Error Resume Next
    target.ListStyle = template.ListStyle
    target.MultiSelect = template.MultiSelect
    target.IntegralHeight = False

    target.Font.Name = template.Font.Name
    target.Font.Size = template.Font.Size
    target.Font.Bold = template.Font.Bold
    target.Font.Italic = template.Font.Italic

    target.Height = template.Height
End Sub

' If exactly one item is selected, selecting it again from the dropdown won't fire Click
' unless the ComboBox.Value changes. Force it to mDash when the user clicks the control.
Private Sub cmbType_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo Cleanup
    If isUpdatingCombo Then Exit Sub
    If selectedTypes Is Nothing Then Exit Sub
    If selectedTypes.Count = 1 Then
        isUpdatingCombo = True
        Me.cmbType.value = Chr(8212)
    End If
Cleanup:
    isUpdatingCombo = False
End Sub

Private Sub cmbCity_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo Cleanup
    If isUpdatingCombo Then Exit Sub
    If selectedCities Is Nothing Then Exit Sub
    If selectedCities.Count = 1 Then
        isUpdatingCombo = True
        Me.cmbCity.value = Chr(8212)
    End If
Cleanup:
    isUpdatingCombo = False
End Sub

Private Sub cmbCountry_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo Cleanup
    If isUpdatingCombo Then Exit Sub
    If selectedCountries Is Nothing Then Exit Sub
    If selectedCountries.Count = 1 Then
        isUpdatingCombo = True
        Me.cmbCountry.value = Chr(8212)
    End If
Cleanup:
    isUpdatingCombo = False
End Sub

Private Sub UserForm_Terminate()
    CancelDeferredUIWork
End Sub

' Helper function to read UTF-8 text files (required for Japanese characters)
Private Function ReadUTF8File(filePath As String) As String
    Dim stream As Object
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.fileExists(filePath) Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' Text
        stream.Charset = "UTF-8"
        stream.Open
        
        On Error Resume Next
        stream.LoadFromFile filePath
        If Err.Number = 0 Then
            ReadUTF8File = stream.ReadText
        Else
            Debug.Print "ReadUTF8File ERROR: " & Err.Description
            ReadUTF8File = ""
        End If
        On Error GoTo 0
        
        stream.Close
    Else
        ReadUTF8File = ""
    End If
End Function

' Normalize Unicode characters in filenames to handle cross-platform and encoding issues
Private Function NormalizeFilenameUnicode(fileName As String) As String
    Dim result As String
    Dim i As Long
    Dim charCode As Long
    Dim currentChar As String
    
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

' Normalize city names by removing " City" suffix for consistent matching
' E.g., "Suzhou City" -> "Suzhou", "Osaka City" -> "Osaka"
Private Function NormalizeCityName(cityName As String) As String
    Dim normalized As String
    normalized = Trim(cityName)
    
    ' Remove " City" suffix (case insensitive)
    If Right(LCase(normalized), 5) = " city" Then
        normalized = Left(normalized, Len(normalized) - 5)
    End If
    
    NormalizeCityName = Trim(normalized)
End Function



Public Sub InitializeSearchForm()
    Dim mDash As String
    mDash = Chr(8212)  ' Em-dash character
    
    Debug.Print "Initializing Search Form"
    
    ' Initialize multi-select collections if not already done
    If selectedTypes Is Nothing Then Set selectedTypes = New Collection
    If selectedCities Is Nothing Then Set selectedCities = New Collection
    If selectedCountries Is Nothing Then Set selectedCountries = New Collection
    
    ' Set IME mode to disable Japanese input (force English/alphanumeric) for text fields
    ' fmIMEModeAlpha = 8 forces alphanumeric (romaji/English) input
    Me.txtAreaMin.IMEMode = fmIMEModeAlpha
    Me.txtAreaMax.IMEMode = fmIMEModeAlpha
    Me.txtHeightMin.IMEMode = fmIMEModeAlpha
    Me.txtHeightMax.IMEMode = fmIMEModeAlpha
    Me.txtYearMin.IMEMode = fmIMEModeAlpha
    Me.txtYearMax.IMEMode = fmIMEModeAlpha
    
    ' Initialize all form controls with mDash
    Me.txtAreaMin.value = mDash
    Me.txtAreaMax.value = mDash
    Me.txtHeightMin.value = mDash
    Me.txtHeightMax.value = mDash
    Me.txtYearMin.value = mDash
    Me.txtYearMax.value = mDash
    
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

    ' Mark initialized so UserForm_Initialize or external callers don't re-run
    isInitialized = True
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

    ' Read the JSON file content using UTF-8 encoding
    jsonText = ReadUTF8File(jsonFilePath)

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
    cleanedValue = Trim(inputValue)
    
    ' Remove common suffixes and units
    cleanedValue = Replace(cleanedValue, ",", "")       ' Remove thousands separator commas
    cleanedValue = Replace(cleanedValue, " ", "")       ' Remove spaces
    cleanedValue = Replace(cleanedValue, "m2", "")      ' Remove m2 (square meters)
    cleanedValue = Replace(cleanedValue, "m²", "")      ' Remove m² (square meters unicode)
    cleanedValue = Replace(cleanedValue, "sqm", "")     ' Remove sqm
    cleanedValue = Replace(cleanedValue, Chr(8212), "") ' Remove M-dash
    
    ' Remove trailing 'm' only if it's a unit suffix (not part of a number)
    If Right(LCase(cleanedValue), 1) = "m" Then
        cleanedValue = Left(cleanedValue, Len(cleanedValue) - 1)
    End If
    
    ' Trim again after removals
    cleanedValue = Trim(cleanedValue)
    
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

' When chkAll is toggled, select or unselect all items in lstResults
Private Sub chkAll_Click()
    Dim i As Long
    Dim shouldSelect As Boolean

    ' If there are no items, nothing to do
    If Me.lstResults.ListCount = 0 Then Exit Sub

    ' Determine desired selection state from the checkbox
    shouldSelect = (Me.chkAll.Value = True)

    ' Loop through all items and set selection state
    For i = 0 To Me.lstResults.ListCount - 1
        Me.lstResults.Selected(i) = shouldSelect
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
    
    ' Clear multi-select collections
    Set selectedTypes = New Collection
    Set selectedCities = New Collection
    Set selectedCountries = New Collection
    
    ' Clear the selected-items ListBoxes
    On Error Resume Next
    Me.lstTypeSelected.Clear
    Me.lstCitySelected.Clear
    Me.lstCountrySelected.Clear
    On Error GoTo 0

    ' Clear any queued deferred removals
    CancelDeferredUIWork
    Set pendingTypeRemovals = New Collection
    Set pendingCityRemovals = New Collection
    Set pendingCountryRemovals = New Collection
    
    ' Set TextBoxes to M-dash for area, height, and year values
    Me.txtAreaMin.value = mDash
    Me.txtAreaMax.value = mDash
    Me.txtHeightMin.value = mDash
    Me.txtHeightMax.value = mDash
    Me.txtYearMin.value = mDash
    Me.txtYearMax.value = mDash
    
    Me.lstResults.Clear
    
    ' Print a message to the Immediate Window to confirm the reset
    Debug.Print "All controls have been reset to M-dash (?\)."
End Sub

' Event handler for cmbType - Toggle multi-select items
Private Sub cmbType_Click()
    HandleMultiSelectCombo Me.cmbType, selectedTypes
End Sub

' Safer than DropButtonClick: runs before opening the dropdown.
' Temporarily show mDash so re-picking the single selected item triggers Click and toggles off.
Private Sub cmbType_Enter()
    On Error GoTo Cleanup
    If isUpdatingCombo Then Exit Sub
    If selectedTypes Is Nothing Then Exit Sub
    If selectedTypes.Count = 1 Then
        isUpdatingCombo = True
        Me.cmbType.value = Chr(8212)
    End If
Cleanup:
    isUpdatingCombo = False
End Sub

' If the user leaves the control without picking anything, restore the display.
Private Sub cmbType_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    UpdateComboDisplay Me.cmbType, selectedTypes
End Sub

' Event handler for cmbCity - Toggle multi-select items
Private Sub cmbCity_Click()
    HandleMultiSelectCombo Me.cmbCity, selectedCities
End Sub

Private Sub cmbCity_Enter()
    On Error GoTo Cleanup
    If isUpdatingCombo Then Exit Sub
    If selectedCities Is Nothing Then Exit Sub
    If selectedCities.Count = 1 Then
        isUpdatingCombo = True
        Me.cmbCity.value = Chr(8212)
    End If
Cleanup:
    isUpdatingCombo = False
End Sub

Private Sub cmbCity_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    UpdateComboDisplay Me.cmbCity, selectedCities
End Sub

' Event handler for cmbCountry - Toggle multi-select items
Private Sub cmbCountry_Click()
    HandleMultiSelectCombo Me.cmbCountry, selectedCountries
End Sub

Private Sub cmbCountry_Enter()
    On Error GoTo Cleanup
    If isUpdatingCombo Then Exit Sub
    If selectedCountries Is Nothing Then Exit Sub
    If selectedCountries.Count = 1 Then
        isUpdatingCombo = True
        Me.cmbCountry.value = Chr(8212)
    End If
Cleanup:
    isUpdatingCombo = False
End Sub

Private Sub cmbCountry_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    UpdateComboDisplay Me.cmbCountry, selectedCountries
End Sub

' Helper function to handle multi-select behavior for ComboBoxes
Private Sub HandleMultiSelectCombo(cmb As MSForms.ComboBox, ByRef selectedItems As Collection)
    Dim mDash As String
    Dim selectedValue As String
    Dim i As Long
    Dim found As Boolean
    Dim foundIndex As Long
    Dim displayText As String
    
    ' Prevent recursive calls when we update cmb.value
    If isUpdatingCombo Then Exit Sub
    
    mDash = Chr(8212)  ' Em-dash character
    selectedValue = cmb.value
    
    ' If M-dash is selected, clear all selections
    If selectedValue = mDash Then
        Set selectedItems = New Collection
        isUpdatingCombo = True
        cmb.value = mDash
        isUpdatingCombo = False
        Exit Sub
    End If
    
    ' If the selected value contains "; " it means user clicked on the display text itself
    ' In this case, we need to find which item from the dropdown was actually clicked
    ' The ComboBox.value will be one of the actual dropdown items, not the combined display
    
    ' Check if the value is already in the collection (toggle off)
    found = False
    foundIndex = 0
    For i = selectedItems.Count To 1 Step -1
        If selectedItems(i) = selectedValue Then
            foundIndex = i
            found = True
            Exit For
        End If
    Next i
    
    ' If found, remove it (toggle off) - only remove that specific item
    If found Then
        selectedItems.Remove foundIndex
    Else
        ' If not found, add it (toggle on)
        selectedItems.Add selectedValue
    End If
    
    ' Update the ComboBox display to show all selected items
    ' Use flag to prevent recursive Click events
    isUpdatingCombo = True
    If selectedItems.Count = 0 Then
        cmb.value = mDash
    Else
        displayText = ""
        For i = 1 To selectedItems.Count
            If i > 1 Then displayText = displayText & "; "
            displayText = displayText & selectedItems(i)
        Next i
        ' Set the text display (this won't match a dropdown item, but shows the selection)
        cmb.value = displayText
    End If
    isUpdatingCombo = False
    
    ' Refresh the corresponding selected-items ListBox
    If cmb.Name = "cmbType" Then
        RefreshSelectedTypesList
    ElseIf cmb.Name = "cmbCity" Then
        RefreshSelectedCitiesList
    ElseIf cmb.Name = "cmbCountry" Then
        RefreshSelectedCountriesList
    End If
End Sub

' Helper function to get selected items as a collection for filtering
Private Function GetSelectedFilterItems(selectedItems As Collection) As Collection
    Set GetSelectedFilterItems = selectedItems
End Function

' Refresh the lstTypeSelected ListBox to show current selections
Private Sub RefreshSelectedTypesList()
    Dim i As Long
    On Error GoTo Cleanup

    isRemovingSelection = True
    Me.lstTypeSelected.Clear
    For i = 1 To selectedTypes.Count
        Me.lstTypeSelected.AddItem selectedTypes(i)
        Me.lstTypeSelected.Selected(i - 1) = True  ' Check the checkbox
    Next i
    ScrollListBoxToBottom Me.lstTypeSelected
    Me.lstTypeSelected.ListIndex = -1

Cleanup:
    isRemovingSelection = False
End Sub

' Refresh the lstCitySelected ListBox to show current selections
Private Sub RefreshSelectedCitiesList()
    Dim i As Long
    On Error GoTo Cleanup

    isRemovingSelection = True
    Me.lstCitySelected.Clear
    For i = 1 To selectedCities.Count
        Me.lstCitySelected.AddItem selectedCities(i)
        Me.lstCitySelected.Selected(i - 1) = True  ' Check the checkbox
    Next i
    ScrollListBoxToBottom Me.lstCitySelected
    Me.lstCitySelected.ListIndex = -1

Cleanup:
    isRemovingSelection = False
End Sub

' Refresh the lstCountrySelected ListBox to show current selections
Private Sub RefreshSelectedCountriesList()
    Dim i As Long
    On Error GoTo Cleanup

    isRemovingSelection = True
    Me.lstCountrySelected.Clear
    For i = 1 To selectedCountries.Count
        Me.lstCountrySelected.AddItem selectedCountries(i)
        Me.lstCountrySelected.Selected(i - 1) = True  ' Check the checkbox
    Next i
    ScrollListBoxToBottom Me.lstCountrySelected
    Me.lstCountrySelected.ListIndex = -1

Cleanup:
    isRemovingSelection = False
End Sub

Private Sub ScrollListBoxToBottom(ByRef lb As MSForms.ListBox)
    Dim lastIdx As Long
    Dim wasSelected As Boolean

    On Error Resume Next
    If lb.ListCount <= 0 Then Exit Sub

    lastIdx = lb.ListCount - 1
    ' Force the control to scroll so the last row is visible.
    ' Preserve the checkbox state of the last item.
    wasSelected = lb.Selected(lastIdx)
    lb.ListIndex = lastIdx
    DoEvents
    lb.Selected(lastIdx) = wasSelected
    lb.ListIndex = -1
End Sub

' MouseDown: capture which row was clicked (so we don't scan/remove unintended rows)
Private Sub lstTypeSelected_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim topIndex As Long
    Dim rowHeight As Single

    If Me.lstTypeSelected.ListCount = 0 Then
        typeClickedIndex = -1
        Exit Sub
    End If

    rowHeight = Me.lstTypeSelected.Font.Size + 2
    If rowHeight <= 0 Then rowHeight = 12

    topIndex = Me.lstTypeSelected.topIndex
    typeClickedIndex = topIndex + CLng(Y \ rowHeight)
    If typeClickedIndex < 0 Or typeClickedIndex >= Me.lstTypeSelected.ListCount Then typeClickedIndex = -1

    Dbg "Type MouseDown: ListCount=" & Me.lstTypeSelected.ListCount & " topIndex=" & topIndex & " rowHeight=" & rowHeight & " Y=" & Y & " computedIdx=" & typeClickedIndex & " ListIndex=" & Me.lstTypeSelected.ListIndex
End Sub

' MouseUp: process after the checkbox state has updated
Private Sub lstTypeSelected_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dbg "Type MouseUp: BEFORE DoEvents ListIndex=" & Me.lstTypeSelected.ListIndex & " computedIdx=" & typeClickedIndex
    DoEvents
    Dbg "Type MouseUp: AFTER  DoEvents ListIndex=" & Me.lstTypeSelected.ListIndex & " computedIdx=" & typeClickedIndex
    ProcessTypeListCheckboxChange
End Sub

' Click handler for lstTypeSelected - ignored (we process on MouseUp)
Private Sub lstTypeSelected_Click()
End Sub

' Helper to process checkbox changes for Type list
Private Sub ProcessTypeListCheckboxChange()
    Dim idx As Long
    Dim itemValue As String

    On Error GoTo Cleanup

    If isRemovingSelection Then Exit Sub
    If Me.lstTypeSelected.ListCount = 0 Then Exit Sub

    ' Prefer ListIndex (works reliably with scrolling). Fallback to MouseDown-estimated index.
    idx = Me.lstTypeSelected.ListIndex
    If idx < 0 Then idx = typeClickedIndex
    If idx < 0 Or idx >= Me.lstTypeSelected.ListCount Then Exit Sub

    Dbg "Type Process: START ListCount=" & Me.lstTypeSelected.ListCount & " ListIndex=" & Me.lstTypeSelected.ListIndex & " idx=" & idx

    On Error Resume Next
    Dbg "Type Process: state Selected(idx)=" & Me.lstTypeSelected.Selected(idx) & " value='" & Me.lstTypeSelected.List(idx) & "'"
    If Err.Number <> 0 Then
        Dbg "Type Process: ERROR reading Selected/List: " & Err.Number & " - " & Err.Description
        Err.Clear
        On Error GoTo Cleanup
        GoTo Cleanup
    End If
    On Error GoTo Cleanup

    ' Only act if the clicked row is now UN-checked
    If Me.lstTypeSelected.Selected(idx) = False Then
        itemValue = Me.lstTypeSelected.List(idx)
        Dbg "Type Process: QUEUE remove value='" & itemValue & "'"
        On Error Resume Next
        pendingTypeRemovals.Add itemValue, itemValue
        On Error GoTo Cleanup
        ScheduleDeferredUIWork
    End If

Cleanup:
    isRemovingSelection = False
    typeClickedIndex = -1
    If Err.Number <> 0 Then Dbg "Type Process: ERROR " & Err.Number & " - " & Err.Description
End Sub

' MouseDown: capture which row was clicked
Private Sub lstCitySelected_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim topIndex As Long
    Dim rowHeight As Single

    If Me.lstCitySelected.ListCount = 0 Then
        cityClickedIndex = -1
        Exit Sub
    End If

    rowHeight = Me.lstCitySelected.Font.Size + 2
    If rowHeight <= 0 Then rowHeight = 12

    topIndex = Me.lstCitySelected.topIndex
    cityClickedIndex = topIndex + CLng(Y \ rowHeight)
    If cityClickedIndex < 0 Or cityClickedIndex >= Me.lstCitySelected.ListCount Then cityClickedIndex = -1

    Dbg "City MouseDown: ListCount=" & Me.lstCitySelected.ListCount & " topIndex=" & topIndex & " rowHeight=" & rowHeight & " Y=" & Y & " computedIdx=" & cityClickedIndex & " ListIndex=" & Me.lstCitySelected.ListIndex
End Sub

Private Sub lstCitySelected_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dbg "City MouseUp: BEFORE DoEvents ListIndex=" & Me.lstCitySelected.ListIndex & " computedIdx=" & cityClickedIndex
    DoEvents
    Dbg "City MouseUp: AFTER  DoEvents ListIndex=" & Me.lstCitySelected.ListIndex & " computedIdx=" & cityClickedIndex
    ProcessCityListCheckboxChange
End Sub

Private Sub lstCitySelected_Click()
End Sub

' Helper to process checkbox changes for City list
Private Sub ProcessCityListCheckboxChange()
    Dim idx As Long
    Dim itemValue As String

    On Error GoTo Cleanup

    If isRemovingSelection Then Exit Sub
    If Me.lstCitySelected.ListCount = 0 Then Exit Sub

    idx = Me.lstCitySelected.ListIndex
    If idx < 0 Then idx = cityClickedIndex
    If idx < 0 Or idx >= Me.lstCitySelected.ListCount Then Exit Sub

    Dbg "City Process: START ListCount=" & Me.lstCitySelected.ListCount & " ListIndex=" & Me.lstCitySelected.ListIndex & " idx=" & idx

    If Me.lstCitySelected.Selected(idx) = False Then
        itemValue = Me.lstCitySelected.List(idx)
        Dbg "City Process: QUEUE remove value='" & itemValue & "'"
        On Error Resume Next
        pendingCityRemovals.Add itemValue, itemValue
        On Error GoTo Cleanup
        ScheduleDeferredUIWork
    End If

Cleanup:
    isRemovingSelection = False
    cityClickedIndex = -1
    If Err.Number <> 0 Then Dbg "City Process: ERROR " & Err.Number & " - " & Err.Description
End Sub

' MouseDown: capture which row was clicked
Private Sub lstCountrySelected_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim topIndex As Long
    Dim rowHeight As Single

    If Me.lstCountrySelected.ListCount = 0 Then
        countryClickedIndex = -1
        Exit Sub
    End If

    rowHeight = Me.lstCountrySelected.Font.Size + 2
    If rowHeight <= 0 Then rowHeight = 12

    topIndex = Me.lstCountrySelected.topIndex
    countryClickedIndex = topIndex + CLng(Y \ rowHeight)
    If countryClickedIndex < 0 Or countryClickedIndex >= Me.lstCountrySelected.ListCount Then countryClickedIndex = -1

    Dbg "Country MouseDown: ListCount=" & Me.lstCountrySelected.ListCount & " topIndex=" & topIndex & " rowHeight=" & rowHeight & " Y=" & Y & " computedIdx=" & countryClickedIndex & " ListIndex=" & Me.lstCountrySelected.ListIndex
End Sub

Private Sub lstCountrySelected_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dbg "Country MouseUp: BEFORE DoEvents ListIndex=" & Me.lstCountrySelected.ListIndex & " computedIdx=" & countryClickedIndex
    DoEvents
    Dbg "Country MouseUp: AFTER  DoEvents ListIndex=" & Me.lstCountrySelected.ListIndex & " computedIdx=" & countryClickedIndex
    ProcessCountryListCheckboxChange
End Sub

Private Sub lstCountrySelected_Click()
End Sub

' Helper to process checkbox changes for Country list
Private Sub ProcessCountryListCheckboxChange()
    Dim idx As Long
    Dim itemValue As String

    On Error GoTo Cleanup

    If isRemovingSelection Then Exit Sub
    If Me.lstCountrySelected.ListCount = 0 Then Exit Sub

    idx = Me.lstCountrySelected.ListIndex
    If idx < 0 Then idx = countryClickedIndex
    If idx < 0 Or idx >= Me.lstCountrySelected.ListCount Then Exit Sub

    Dbg "Country Process: START ListCount=" & Me.lstCountrySelected.ListCount & " ListIndex=" & Me.lstCountrySelected.ListIndex & " idx=" & idx

    If Me.lstCountrySelected.Selected(idx) = False Then
        itemValue = Me.lstCountrySelected.List(idx)
        Dbg "Country Process: QUEUE remove value='" & itemValue & "'"
        On Error Resume Next
        pendingCountryRemovals.Add itemValue, itemValue
        On Error GoTo Cleanup
        ScheduleDeferredUIWork
    End If

Cleanup:
    isRemovingSelection = False
    countryClickedIndex = -1
    If Err.Number <> 0 Then Dbg "Country Process: ERROR " & Err.Number & " - " & Err.Description
End Sub

' Helper to update ComboBox display text after removing from ListBox
Private Sub UpdateComboDisplay(cmb As MSForms.ComboBox, selectedItems As Collection)
    Dim mDash As String
    Dim displayText As String
    Dim i As Long

    On Error GoTo Cleanup

    mDash = Chr(8212)
    isUpdatingCombo = True

    If selectedItems Is Nothing Then
        cmb.value = mDash
        GoTo Cleanup
    End If

    If selectedItems.Count = 0 Then
        cmb.value = mDash
    Else
        displayText = ""
        For i = 1 To selectedItems.Count
            If i > 1 Then displayText = displayText & "; "
            displayText = displayText & selectedItems(i)
        Next i
        cmb.value = displayText
    End If

Cleanup:
    isUpdatingCombo = False
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
    
    ' Read with UTF-8 encoding for Japanese character support
    jsonText = ReadUTF8File(cacheFilePath)
    
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
                
                ' Normalize Unicode in file path for better matching
                pdfFilePath = NormalizeFilenameUnicode(pdfFilePath)
                
                ' Check if file exists before adding to the collection - use Dir for better Unicode support
                On Error Resume Next
                If Len(Dir(pdfFilePath)) > 0 Then
                    filePaths.Add pdfFilePath
                Else
                    ' File not found - try code-based matching
                    Dim searchFolder As String
                    Dim actualFile As String
                    actualFile = ""
                    
                    ' Get folder path
                    On Error Resume Next
                    If InStr(pdfFilePath, "\") > 0 Then
                        searchFolder = fso.GetParentFolderName(pdfFilePath)
                    ElseIf selectedFolderPath <> "" Then
                        searchFolder = selectedFolderPath
                    End If
                    
                    ' Try to find file by code
                    If searchFolder <> "" And fso.FolderExists(searchFolder) Then
                        Dim targetFileName As String
                        Dim fileCode As String
                        targetFileName = fso.GetFileName(pdfFilePath)
                        
                        If InStr(targetFileName, "_") > 0 Then
                            fileCode = Left(targetFileName, InStr(targetFileName, "_") - 1)
                            
                            ' Look for matching file
                            Dim file As Object
                            For Each file In fso.GetFolder(searchFolder).Files
                                If Left(fso.GetFileName(file.path), Len(fileCode)) = fileCode And LCase(Right(file.Name, 4)) = ".pdf" Then
                                    actualFile = file.path
                                    Exit For
                                End If
                            Next
                        Else
                            ' No underscore - try exact match first (for Japanese filenames)
                            Dim exactMatchPath As String
                            exactMatchPath = searchFolder & "\" & targetFileName
                            On Error Resume Next
                            If Len(Dir(exactMatchPath)) > 0 Or fso.FileExists(exactMatchPath) Then
                                actualFile = exactMatchPath
                            End If
                            On Error GoTo 0
                        End If
                    End If
                    On Error GoTo 0
                    
                    ' If found by code, use it; otherwise show error
                    On Error Resume Next
                    If actualFile <> "" And Len(Dir(actualFile)) > 0 Then
                        filePaths.Add actualFile
                        Debug.Print "Code-based match used: " & actualFile
                    Else
                        On Error GoTo 0
                        MsgBox "File not found: " & pdfFilePath, vbExclamation, "File Missing"
                        Exit Sub
                    End If
                    On Error GoTo 0
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

    ' Read with UTF-8 encoding for Japanese character support
    jsonText = ReadUTF8File(cacheFilePath)
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
                
                ' Normalize Unicode characters in the file path for consistent matching
                pdfFilePath = NormalizeFilenameUnicode(pdfFilePath)
                
                ' Debug: Print character codes for troubleshooting
                Dim debugMsg As String
                debugMsg = "File path from JSON: " & pdfFilePath & vbCrLf
                If InStr(pdfFilePath, "TOKYO") > 0 Then
                    Dim charPos As Integer
                    charPos = InStr(pdfFilePath, "MEGURO") + 6  ' Position after "MEGURO"
                    If charPos > 6 And charPos <= Len(pdfFilePath) Then
                        debugMsg = debugMsg & "Character code at position " & charPos & ": &H" & Hex(AscW(Mid(pdfFilePath, charPos, 1)))
                        Debug.Print debugMsg
                    End If
                End If
                
                found = True

                matchedFile = ""
                
                Debug.Print "pdfFilePath: " & pdfFilePath
                
                ' Check if the original file exists
                If Dir(pdfFilePath) <> "" Then
                    matchedFile = pdfFilePath
                    Debug.Print "Original file found: " & matchedFile
                Else
                    ' Original file not found - try code-based matching first
                    Debug.Print "Original file not found, trying code-based match..."
                    
                    On Error Resume Next
                    
                    ' Check if pdfFilePath has a parent folder, otherwise try selectedFolderPath
                    Dim searchFolder As String
                    searchFolder = ""
                    
                    If InStr(pdfFilePath, "\") > 0 Then
                        ' pdfFilePath contains a path
                        If fso.FolderExists(fso.GetParentFolderName(pdfFilePath)) Then
                            searchFolder = fso.GetParentFolderName(pdfFilePath)
                        End If
                    Else
                        ' pdfFilePath is just a filename - try to extract folder from other cache entries
                        Dim otherItem As Object
                        For Each otherItem In cacheData
                            If otherItem.Exists("File Path") Then
                                Dim otherPath As String
                                otherPath = otherItem("File Path")
                                If InStr(otherPath, "\") > 0 And fso.FolderExists(fso.GetParentFolderName(otherPath)) Then
                                    searchFolder = fso.GetParentFolderName(otherPath)
                                    Debug.Print "Using folder from other cache entry: " & searchFolder
                                    ' Try constructing the full path
                                    Dim possiblePath As String
                                    possiblePath = searchFolder & "\" & pdfFilePath
                                    If fso.FileExists(possiblePath) Then
                                        matchedFile = possiblePath
                                        Debug.Print "Reconstructed full path: " & matchedFile
                                        Exit For
                                    End If
                                End If
                            End If
                        Next otherItem
                    End If
                    
                    ' If no valid folder yet, try selectedFolderPath from MainModule
                    If searchFolder = "" And selectedFolderPath <> "" Then
                        If fso.FolderExists(selectedFolderPath) Then
                            searchFolder = selectedFolderPath
                        End If
                    End If
                    
                    ' Search for similar files if we have a valid folder
                    If searchFolder <> "" Then
                        ' First try code-based matching (more reliable for special characters)
                        Dim targetFileName As String
                        Dim fileCode As String
                        targetFileName = fso.GetFileName(pdfFilePath)
                        
                        ' Extract file code (e.g., "N070076" from "N070076_SOMETHING.pdf")
                        If InStr(targetFileName, "_") > 0 Then
                            fileCode = Left(targetFileName, InStr(targetFileName, "_") - 1)
                            Debug.Print "Searching by file code: " & fileCode
                            
                            ' Look for files with matching code
                            For Each file In fso.GetFolder(searchFolder).Files
                                Dim currentFileName As String
                                currentFileName = fso.GetFileName(file.path)
                                If Left(currentFileName, Len(fileCode)) = fileCode And LCase(Right(currentFileName, 4)) = ".pdf" Then
                                    matchedFile = file.path
                                    Debug.Print "Code-based match found: " & matchedFile
                                    Exit For
                                End If
                            Next
                        Else
                            ' No underscore - try exact match first (for Japanese filenames)
                            Debug.Print "No file code found, trying exact match for: " & targetFileName
                            Dim exactMatchPath As String
                            exactMatchPath = searchFolder & "\" & targetFileName
                            
                            ' Try direct file existence check first
                            If fso.FileExists(exactMatchPath) Then
                                matchedFile = exactMatchPath
                                Debug.Print "Exact match found: " & matchedFile
                            Else
                                ' If exact match fails, try normalizing Unicode characters and searching
                                Debug.Print "Exact match failed, trying normalized filename search..."
                                Dim normalizedTarget As String
                                normalizedTarget = NormalizeFilenameUnicode(targetFileName)
                                
                                ' Search folder for files with normalized names
                                For Each file In fso.GetFolder(searchFolder).Files
                                    Dim diskFileName As String
                                    Dim normalizedDisk As String
                                    diskFileName = fso.GetFileName(file.path)
                                    normalizedDisk = NormalizeFilenameUnicode(diskFileName)
                                    
                                    If StrComp(normalizedTarget, normalizedDisk, vbTextCompare) = 0 Then
                                        matchedFile = file.path
                                        Debug.Print "Normalized match found: " & matchedFile
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                        
                        ' If code-based matching failed, try similarity matching
                        If matchedFile = "" Then
                            Debug.Print "Code-based match failed, trying similarity..."
                            
                            ' Remove extension for better comparison
                            Dim targetBaseName As String
                            Dim actualBaseName As String
                            targetBaseName = Replace(targetFileName, ".pdf", "", 1, -1, vbTextCompare)
                            
                            For Each file In fso.GetFolder(searchFolder).Files
                                actualBaseName = Replace(fso.GetFileName(file.path), ".pdf", "", 1, -1, vbTextCompare)
                                
                                If SimilarityRatio(targetBaseName, actualBaseName) > similarityThreshold Then
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
                        End If
                    End If
                    
                    On Error GoTo 0
                End If

                If matchedFile <> "" And Dir(matchedFile) <> "" Then
                    filePaths.Add matchedFile
                    Debug.Print "Added to merge list: " & matchedFile
                Else
                    MsgBox "File not found and no valid replacement: " & pdfFilePath & vbCrLf & "Search folder: " & searchFolder, vbExclamation, "File Missing"
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

' QuickSort strings in-place (case-insensitive)
Private Sub QuickSortStrings(arr() As String, low As Long, high As Long)
    Dim pivot As String, tmp As String
    Dim i As Long, j As Long
    If low >= high Then Exit Sub
    pivot = arr((low + high) \ 2)
    i = low: j = high
    Do While i <= j
        Do While LCase(arr(i)) < LCase(pivot)
            i = i + 1
        Loop
        Do While LCase(arr(j)) > LCase(pivot)
            j = j - 1
        Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then QuickSortStrings arr, low, j
    If i < high Then QuickSortStrings arr, i, high
End Sub

' Helper to sort a dynamic string array and return it (sorted)
Private Function SortStringArray(inputArr() As String) As String()
    Dim n As Long
    n = UBound(inputArr) - LBound(inputArr) + 1
    If n <= 1 Then
        SortStringArray = inputArr
        Exit Function
    End If
    QuickSortStrings inputArr, LBound(inputArr), UBound(inputArr)
    SortStringArray = inputArr
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

    ' Read the JSON file content with UTF-8 encoding for Japanese character support
    jsonText = ReadUTF8File(jsonFilePath)

    ' Parse the JSON data using JsonConverter
    Set jsonData = JsonConverter.ParseJSON(jsonText)

    ' Validate the parsed JSON data
    If jsonData Is Nothing Or TypeName(jsonData) <> "Collection" Then
        MsgBox "Invalid JSON structure. Expected a Collection.", vbExclamation, "Data Error"
        Exit Sub
    End If

    ' Check if multi-select collections have items, otherwise use single value
    Dim hasTypeFilters As Boolean, hasCityFilters As Boolean, hasCountryFilters As Boolean
    hasTypeFilters = (selectedTypes.Count > 0)
    hasCityFilters = (selectedCities.Count > 0)
    hasCountryFilters = (selectedCountries.Count > 0)
    
    ' Retrieve numeric filter values from the form
    areaMinFilter = GetNumericValue(Me.txtAreaMin.value, 0)
    areaMaxFilter = GetNumericValue(Me.txtAreaMax.value, 9999999)
    heightMinFilter = GetNumericValue(Me.txtHeightMin.value, 0)
    heightMaxFilter = GetNumericValue(Me.txtHeightMax.value, 9999)
    yearMinFilter = GetNumericValue(Me.txtYearMin.value, 0)
    yearMaxFilter = GetNumericValue(Me.txtYearMax.value, 9999)

    ' Clear the results list before displaying new results
    ' Clear the results list before displaying new results
    Me.lstResults.Clear
    Me.lstResults.Visible = True
    Me.cmdSavePdf.Visible = True

    ' Use a temporary collection to gather matching project names, then sort
    Dim matches() As String
    Dim matchCount As Long
    matchCount = 0

    ' Loop through each project in the parsed JSON data
    For Each projectItem In jsonData
        matchFound = True

        ' Get the project name
        If projectItem.Exists("Project Name") Then
            projectName = projectItem("Project Name")
        Else
            GoTo ContinueLoop
        End If

        ' Check Type filter using multi-select collection
        If hasTypeFilters Then
            matchFound = MultiSelectMatch(projectItem("Type"), selectedTypes)
        End If
        
        ' Check City filter using multi-select collection
        If matchFound And hasCityFilters Then
            If projectItem.Exists("Location") Then
                ' Location is an array, so we need to check each element
                Dim locationMatch As Boolean
                locationMatch = False
                
                If TypeName(projectItem("Location")) = "Collection" Then
                    Dim locItem As Variant
                    Dim cityItem As Variant
                    For Each locItem In projectItem("Location")
                        For Each cityItem In selectedCities
                            If InStr(1, CStr(locItem), CStr(cityItem), vbTextCompare) > 0 Then
                                locationMatch = True
                                Exit For
                            End If
                        Next cityItem
                        If locationMatch Then Exit For
                    Next locItem
                Else
                    ' Fallback if Location is a string
                    For Each cityItem In selectedCities
                        If InStr(1, CStr(projectItem("Location")), CStr(cityItem), vbTextCompare) > 0 Then
                            locationMatch = True
                            Exit For
                        End If
                    Next cityItem
                End If
                
                matchFound = locationMatch
            Else
                ' Use normalized city matching for City field
                Dim normalizedFilter As String
                Dim normalizedCity As String
                Dim cityMatchFound As Boolean
                cityMatchFound = False
                
                If projectItem.Exists("City") Then
                    normalizedCity = NormalizeCityName(CStr(projectItem("City")))
                    For Each cityItem In selectedCities
                        normalizedFilter = NormalizeCityName(CStr(cityItem))
                        ' Match if normalized filter is contained in normalized city OR vice versa
                        If InStr(1, normalizedCity, normalizedFilter, vbTextCompare) > 0 Or _
                           InStr(1, normalizedFilter, normalizedCity, vbTextCompare) > 0 Then
                            cityMatchFound = True
                            Exit For
                        End If
                    Next cityItem
                End If
                matchFound = cityMatchFound
            End If
        End If

        ' Check Country filter using multi-select collection
        ' Also exclude projects with Country = "n/a" from all country filter searches
        If matchFound And hasCountryFilters Then
            ' First check if the project's country is "n/a" - if so, exclude it from any country filter
            If projectItem.Exists("Country") And LCase(Trim(projectItem("Country"))) = "n/a" Then
                matchFound = False
            Else
                matchFound = MultiSelectMatchCountry(projectItem("Country"), selectedCountries)
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

        ' If all criteria match, add the project name to temporary array
        If matchFound Then
            matchCount = matchCount + 1
            ReDim Preserve matches(0 To matchCount - 1)
            matches(matchCount - 1) = CStr(projectName)
        End If

ContinueLoop:
    Next projectItem

    ' Display message if no results found
    If matchCount = 0 Then
        Me.lstResults.AddItem "No matching projects found."
        Debug.Print "Search completed. 0 matching project(s) found."
        Exit Sub
    End If

    ' Sort the matches alphabetically and add with numbering
    matches = SortStringArray(matches)
    For resultCount = 0 To UBound(matches)
        Me.lstResults.AddItem CStr(resultCount + 1) & ". " & matches(resultCount)
    Next resultCount

    ' Debug message for results
    Debug.Print "Search completed. " & matchCount & " matching project(s) found."
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

' Function to check if any selected item from multi-select collection matches the grouped value
Private Function MultiSelectMatch(groupedValue As String, selectedItems As Collection) As Boolean
    Dim part As Variant
    Dim selectedItem As Variant
    Dim splitGroup() As String
    Dim splitParts() As String
    Dim searchPart As Variant
    
    MultiSelectMatch = False
    
    ' Split the grouped value by commas
    splitGroup = Split(groupedValue, ",")
    
    ' Loop through each selected item
    For Each selectedItem In selectedItems
        ' The selected item may contain " / " separators, so split those too
        splitParts = Split(CStr(selectedItem), "/")
        
        For Each searchPart In splitParts
            For Each part In splitGroup
                ' Trim spaces and check if there's a partial match
                If InStr(1, Trim(CStr(part)), Trim(CStr(searchPart)), vbTextCompare) > 0 Then
                    MultiSelectMatch = True
                    Exit Function
                End If
            Next part
        Next searchPart
    Next selectedItem
End Function

' Function to check country matches with special handling for China/Republic of China
Private Function MultiSelectMatchCountry(projectCountry As String, selectedItems As Collection) As Boolean
    Dim selectedItem As Variant
    Dim splitParts() As String
    Dim searchPart As Variant
    
    MultiSelectMatchCountry = False
    
    ' Loop through each selected country
    For Each selectedItem In selectedItems
        ' The selected item may contain " / " separators
        splitParts = Split(CStr(selectedItem), "/")
        
        For Each searchPart In splitParts
            Dim trimmedSearch As String
            trimmedSearch = Trim(CStr(searchPart))
            
            ' Special handling for China vs Republic of China
            If trimmedSearch = "China" Then
                ' Match China but not Republic of China
                If InStr(1, projectCountry, "China", vbTextCompare) > 0 And _
                   InStr(1, projectCountry, "Republic of China", vbTextCompare) = 0 Then
                    MultiSelectMatchCountry = True
                    Exit Function
                End If
            ElseIf trimmedSearch = "Republic of China" Then
                ' Match only Republic of China
                If InStr(1, projectCountry, "Republic of China", vbTextCompare) > 0 Then
                    MultiSelectMatchCountry = True
                    Exit Function
                End If
            Else
                ' Regular country match
                If InStr(1, projectCountry, trimmedSearch, vbTextCompare) > 0 Then
                    MultiSelectMatchCountry = True
                    Exit Function
                End If
            End If
        Next searchPart
    Next selectedItem
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

' Double-click: open the PDF file for the selected list item
Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim idx As Long
    Dim projectName As String
    Dim jsonFilePath As String
    Dim jsonText As String
    Dim jsonData As Object
    Dim item As Object
    Dim fso As Object
    Dim originalPath As String
    Dim resolvedPath As String

    If Me.lstResults.ListCount = 0 Then Exit Sub

    ' Use the ListIndex (selected row) which is reliable for double-clicks
    If Me.lstResults.ListIndex >= 0 Then
        idx = Me.lstResults.ListIndex
    ElseIf clickedIndex >= 0 And clickedIndex < Me.lstResults.ListCount Then
        ' fallback to clickedIndex if ListIndex not set
        idx = clickedIndex
    Else
        Exit Sub
    End If

    projectName = Trim(Mid(Me.lstResults.List(idx), InStr(Me.lstResults.List(idx), ". ") + 2))

    Set fso = CreateObject("Scripting.FileSystemObject")
    jsonFilePath = GetRelativePath("\test output\cache.json")

    If Not fso.fileExists(jsonFilePath) Then
        MsgBox "Cache file not found: " & jsonFilePath, vbExclamation, "File Missing"
        Exit Sub
    End If

    jsonText = ReadUTF8File(jsonFilePath)
    If Len(Trim(jsonText)) = 0 Then
        MsgBox "Cache file is empty.", vbExclamation, "Empty Cache"
        Exit Sub
    End If

    Set jsonData = JsonConverter.ParseJSON(jsonText)
    If jsonData Is Nothing Then
        MsgBox "Unable to parse cache file.", vbExclamation, "Parse Error"
        Exit Sub
    End If

    For Each item In jsonData
        If item.Exists("Project Name") Then
            If item("Project Name") = projectName Then
                If item.Exists("File Path") Then
                    originalPath = item("File Path")
                    resolvedPath = ResolvePdfFilePath(originalPath, jsonData)
                    If resolvedPath <> "" Then
                        On Error Resume Next
                        Application.FollowHyperlink resolvedPath
                        If Err.Number <> 0 Then
                            Err.Clear
                            Shell "explorer.exe " & Chr(34) & resolvedPath & Chr(34), vbNormalFocus
                        End If
                        On Error GoTo 0
                    Else
                        MsgBox "File not found: " & originalPath, vbExclamation, "File Missing"
                    End If
                Else
                    MsgBox "No file path recorded for project: " & projectName, vbExclamation, "Missing Data"
                End If
                Exit Sub
            End If
        End If
    Next item

    MsgBox "Project not found in cache: " & projectName, vbExclamation, "Not Found"
End Sub


' Resolve a stored PDF path to an actual existing file using the same
' matching strategies used by the merge routine (exact, code-based,
' normalized, similarity). Returns empty string if none found.
Private Function ResolvePdfFilePath(pdfFilePath As String, cacheData As Object) As String
    Dim fso As Object
    Dim targetFileName As String
    Dim searchFolder As String
    Dim file As Object
    Dim fileCode As String
    Dim matchedFile As String
    Dim otherItem As Object
    Dim exactMatchPath As String
    Dim normalizedTarget As String
    Dim normalizedDisk As String
    Dim similarityThreshold As Double

    similarityThreshold = 0.8
    matchedFile = ""
    Set fso = CreateObject("Scripting.FileSystemObject")

    pdfFilePath = NormalizeFilenameUnicode(pdfFilePath)

    ' Try original path first using both Dir() and FSO for better Unicode support
    On Error Resume Next
    If Len(Dir(pdfFilePath)) > 0 Or fso.FileExists(pdfFilePath) Then
        ResolvePdfFilePath = pdfFilePath
        Exit Function
    End If
    On Error GoTo 0

    ' Determine a search folder
    If InStr(pdfFilePath, "\") > 0 Then
        If fso.FolderExists(fso.GetParentFolderName(pdfFilePath)) Then
            searchFolder = fso.GetParentFolderName(pdfFilePath)
        End If
    Else
        ' Try to reuse a folder from other cache entries
        For Each otherItem In cacheData
            If otherItem.Exists("File Path") Then
                Dim otherPath As String
                otherPath = otherItem("File Path")
                If InStr(otherPath, "\") > 0 And fso.FolderExists(fso.GetParentFolderName(otherPath)) Then
                    searchFolder = fso.GetParentFolderName(otherPath)
                    Exit For
                End If
            End If
        Next otherItem
    End If

    ' If still empty, try global selectedFolderPath if available
    On Error Resume Next
    If searchFolder = "" Then
        If selectedFolderPath <> "" And fso.FolderExists(selectedFolderPath) Then
            searchFolder = selectedFolderPath
        End If
    End If
    On Error GoTo 0

    targetFileName = fso.GetFileName(pdfFilePath)

    If searchFolder <> "" Then
        ' Try code-based match (prefix before underscore)
        If InStr(targetFileName, "_") > 0 Then
            fileCode = Left(targetFileName, InStr(targetFileName, "_") - 1)
            For Each file In fso.GetFolder(searchFolder).Files
                If Left(fso.GetFileName(file.path), Len(fileCode)) = fileCode And LCase(Right(file.Name, 4)) = ".pdf" Then
                    matchedFile = file.path
                    Exit For
                End If
            Next file
        End If

        ' Exact match - use both Dir() and FSO for better Unicode support
        If matchedFile = "" Then
            exactMatchPath = searchFolder & "\" & targetFileName
            On Error Resume Next
            If Len(Dir(exactMatchPath)) > 0 Or fso.FileExists(exactMatchPath) Then
                matchedFile = exactMatchPath
            End If
            On Error GoTo 0
        End If

        ' Normalized filename match
        If matchedFile = "" Then
            normalizedTarget = NormalizeFilenameUnicode(targetFileName)
            For Each file In fso.GetFolder(searchFolder).Files
                normalizedDisk = NormalizeFilenameUnicode(fso.GetFileName(file.path))
                If StrComp(normalizedTarget, normalizedDisk, vbTextCompare) = 0 Then
                    matchedFile = file.path
                    Exit For
                End If
            Next file
        End If

        ' Similarity fallback
        If matchedFile = "" Then
            Dim targetBaseName As String, actualBaseName As String
            targetBaseName = Replace(targetFileName, ".pdf", "", 1, -1, vbTextCompare)
            For Each file In fso.GetFolder(searchFolder).Files
                actualBaseName = Replace(fso.GetFileName(file.path), ".pdf", "", 1, -1, vbTextCompare)
                If SimilarityRatio(targetBaseName, actualBaseName) > similarityThreshold Then
                    If MsgBox("Original file not found. A similar file '" & fso.GetFileName(file.path) & "' was found. Use this file?", vbYesNo + vbQuestion, "File Not Found") = vbYes Then
                        matchedFile = file.path
                        Exit For
                    End If
                End If
            Next file
        End If
    End If

    ResolvePdfFilePath = matchedFile
End Function






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

    ' Read the JSON file content with UTF-8 encoding for Japanese character support
    jsonText = ReadUTF8File(jsonFilePath)
'Debug.Print "jsonText: " & jsonText

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
                
                ' Skip "n/a" values entirely - don't add them to dropdowns
                If LCase(splitValue) = "n/a" Then GoTo NextSplitValue
                
                ' Normalize city names if we're populating the City dropdown
                If key = "City" Then
                    splitValue = NormalizeCityName(CStr(splitValue))
                    ' Skip empty values after normalization
                    If Len(splitValue) = 0 Then GoTo NextSplitValue
                End If
                
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
NextSplitValue:
            Next splitValue
        End If
ContinueLoop:
    Next projectItem

    ' Populate ComboBox with grouped and sorted values (alphabetically)
    cmb.Clear
    cmb.AddItem mDash  ' Set M-dash as the initial value

    Dim sortedValues As Collection
    Dim finalValue As String
    Dim sortedArray() As String
    Dim allItems() As String
    Dim itemCount As Long
    itemCount = 0

    ' First, collect all items into an array
    For Each groupedKey In typeGroups.keys
        Set sortedValues = SortValuesByLength(typeGroups(groupedKey))

        ' Convert sortedValues (Collection) to an array for Join
        sortedArray = CollectionToArray(sortedValues)
        finalValue = Join(sortedArray, " / ")

        ' Add to temporary array for sorting
        itemCount = itemCount + 1
        ReDim Preserve allItems(0 To itemCount - 1)
        allItems(itemCount - 1) = CStr(finalValue)
    Next groupedKey
    
    ' Sort all items alphabetically before adding to ComboBox
    If itemCount > 0 Then
        allItems = SortStringArray(allItems)
        
        ' Add sorted items to ComboBox
        Dim idx As Long
        For idx = 0 To UBound(allItems)
            cmb.AddItem allItems(idx)
        Next idx
    End If

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

    ' Read the JSON file content with UTF-8 encoding for Japanese character support
    jsonText = ReadUTF8File(jsonFilePath)
    
    ' DEBUG: Show what was actually read from the file
    Debug.Print "PopulateListFromJsonCache: jsonText length = " & Len(jsonText)
    Debug.Print "PopulateListFromJsonCache: First 200 chars: " & Left(jsonText, 200)

    ' Parse the JSON data
    Set jsonData = JsonConverter.ParseJSON(jsonText)

    ' Validate the parsed JSON data
    If jsonData Is Nothing Or TypeName(jsonData) <> "Collection" Then
        MsgBox "Invalid JSON structure. Expected a Collection.", vbExclamation, "Data Error"
        Exit Sub
    End If


    ' Collect project names, then sort alphabetically before populating
    Dim names() As String
    Dim countNames As Long
    countNames = 0

    For Each projectItem In jsonData
        If projectItem.Exists("Project Name") Then
            projectName = projectItem("Project Name")
            countNames = countNames + 1
            ReDim Preserve names(0 To countNames - 1)
            names(countNames - 1) = CStr(projectName)
        End If
    Next projectItem

    Me.lstResults.Clear
    If countNames = 0 Then
        Me.lstResults.AddItem "No projects found in the cache file."
        Debug.Print "List populated from JSON cache. Total projects: 0"
        Exit Sub
    End If

    names = SortStringArray(names)

    For resultCount = 0 To UBound(names)
        Me.lstResults.AddItem CStr(resultCount + 1) & ". " & names(resultCount)
    Next resultCount

    ' Display a message if no projects were found
    If resultCount = 0 Then
        Me.lstResults.AddItem "No projects found in the cache file."
    End If

    Debug.Print "List populated from JSON cache. Total projects: " & resultCount
    Exit Sub

ErrorHandler:
    MsgBox "Error in PopulateListFromJsonCache: " & Err.Description, vbCritical, "Error"
End Sub








