Attribute VB_Name = "Generator"
'Attribute VB_Name = "Generator"
' Module: Generator

Option Explicit

' Ensure that parsedDataSet is properly initialized before running GenerateFeeProposal
' Example Initialization:
' Sub InitializeParsedData()
'     Set parsedDataSet = CreateObject("Scripting.Dictionary")
'     parsedDataSet.Add "Title", "Fee Proposal for XYZ Project"
'     parsedDataSet.Add "Subtitle", "Comprehensive Services Overview"
'     parsedDataSet.Add "Content", "Line 1" & vbCrLf & "Line 2" & vbCrLf & "..." ' Continue as needed
' End Sub

Public Sub GenerateFeeProposal()
    Dim parsedData As Object
    Dim outputFilePath As String
    Dim formatFilePath As String
    Dim powerPointApp As Object
    Dim presentation As Object
    Dim relevantFolder As String ' Specify the relevant folder and output file path
    
    relevantFolder = GetRelativePath("\formats\relevant")
    
    
    ' Step 1: Get parsed data from the public variable object
    Set parsedData = parsedDataSet
    
    ' Step 2: Initialize PowerPoint Application
    Set powerPointApp = CreateObject("PowerPoint.Application")
    powerPointApp.Visible = True ' Make PowerPoint visible
    
    formatFilePath = GetRelativePath("\formats\FeeProposalFormat.pptx")
    
    ' Step 3: Open PowerPoint Template
    Set presentation = powerPointApp.Presentations.Open(formatFilePath)
    
    ' Step 4: Populate PowerPoint Template
    Call PopulatePresentation(presentation, parsedData)
    
    ' Call the function to merge slides into the passed-in presentation object
    'MergeSpecialtySlides parsedData, relevantFolder, presentation
    InsertSpecialtySlidesAndRemove parsedData, relevantFolder, presentation

    ' Step 5: Save the populated presentation
    outputFilePath = GetRelativePath("\test output\FP-test.pptx")

    presentation.SaveAs outputFilePath
    
    ' Optionally save as PDF
    ' presentation.SaveAs Replace(outputFilePath, ".pptx", ".pdf"), ppSaveAsPDF
    
    ' Cleanup
    'presentation.Close
    'powerPointApp.Quit
    Set presentation = Nothing
    Set powerPointApp = Nothing
End Sub

Sub PopulatePresentation(presentation As Object, parsedData As Object)
    On Error GoTo ErrorHandler
    
    Dim slide As Object
    Dim shape As Object
    Dim key As Variant
    Dim placeholderText As String
    Dim replacementText As String
    Dim maxLinesPerTextbox As Long
    Dim textChunks As Variant
    Dim i As Long
    Dim newSlide As Object
    Dim slideIndex As Long
    Dim shapesDict As Object ' Dictionary: key -> collection of shapes
    Dim chunksDict As Object ' Dictionary: key -> array of chunks
    
    ' Define maximum number of lines per textbox (adjust as needed)
    maxLinesPerTextbox = 14 ' Maximum 14 lines per textbox
    
    ' Initialize slide index
    slideIndex = 1
    
    ' Loop through each slide in the presentation
    Do While slideIndex <= presentation.Slides.Count
        Set slide = presentation.Slides(slideIndex)
        
        ' Initialize dictionaries
        Set shapesDict = CreateObject("Scripting.Dictionary")
        Set chunksDict = CreateObject("Scripting.Dictionary")
        
        ' First pass: Identify shapes with placeholders and map keys to shapes
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    Dim shapeText As String
                    shapeText = shape.TextFrame.TextRange.text
                    
                    ' Check for any placeholder
                    If InStr(shapeText, "{{") > 0 And InStr(shapeText, "}}") > 0 Then
                        ' Find all keys in parsedData that are present in the shape text
                        For Each key In parsedData.keys
                            placeholderText = "{{" & key & "}}"
                            If InStr(shapeText, placeholderText) > 0 Then
                                ' Add shape to the collection for this key
                                If Not shapesDict.Exists(key) Then
                                    Set shapesDict(key) = New Collection
                                End If
                                shapesDict(key).Add shape
                                
                                ' Optionally, log the mapping
                                'Debug.Print "Slide " & slideIndex & ": Shape '" & shape.Name & "' contains placeholder '" & placeholderText & "'."
                            End If
                        Next key
                    End If
                End If
            End If
        Next shape
        
        ' Debug: Print number of keys found in the slide
        'Debug.Print "Processing Slide " & slideIndex & " with " & shapesDict.Count & " key(s)."
        
        ' If no shapes to process, move to next slide
        If shapesDict.Count = 0 Then
            slideIndex = slideIndex + 1
            GoTo NextSlide
        End If
        
        ' For each key, split the replacement text into chunks only if necessary
        For Each key In shapesDict.keys
            replacementText = parsedData(key)
            textChunks = SplitTextIntoChunksIfNeeded(replacementText, maxLinesPerTextbox)
            chunksDict(key) = textChunks
            'Debug.Print "Key '" & key & "' has " & UBound(textChunks) & " chunk(s)."
        Next key
        
        ' Determine the maximum number of chunks across all keys
        Dim maxChunks As Long
        maxChunks = 1 ' At least one slide
        
        For Each key In chunksDict.keys
            Dim numChunks As Long
            numChunks = UBound(chunksDict(key)) ' assuming 1-based array
            If numChunks > maxChunks Then
                maxChunks = numChunks
            End If
        Next key
        
        ' Debug: Print maxChunks
        'Debug.Print "Maximum chunks needed: " & maxChunks
        
        ' If maxChunks > 1, duplicate the slide accordingly
        If maxChunks > 1 Then
            For i = 1 To (maxChunks - 1)
                slide.Duplicate
                'Debug.Print "Duplicated Slide " & slideIndex & " to Slide " & (presentation.Slides.Count)
            Next i
        End If
        
        ' Now, for each slide (original + duplicates), insert the corresponding chunks
        Dim s As Long
        For s = 0 To (maxChunks - 1)
            Set newSlide = presentation.Slides(slideIndex + s)
            'Debug.Print "Populating Slide " & (slideIndex + s) & " with chunk index " & (s + 1)
            
            ' For each shape in the slide, replace all placeholders
            For Each shape In newSlide.Shapes
                If shape.HasTextFrame Then
                    If shape.TextFrame.HasText Then
                        Dim originalText As String
                        originalText = shape.TextFrame.TextRange.text
                        Dim modifiedText As String
                        modifiedText = originalText
                        
                        ' Replace all placeholders in this shape
                        For Each key In parsedData.keys
                            placeholderText = "{{" & key & "}}"
                            If InStr(modifiedText, placeholderText) > 0 Then
                                Dim chunkText As String
                                chunkText = chunksDict(key)(s + 1)
                                ' Replace placeholder
                                modifiedText = Replace(modifiedText, placeholderText, chunkText)
                                'Debug.Print "Replaced '" & placeholderText & "' in Shape '" & shape.Name & "' on Slide " & (slideIndex + s)
                            End If
                        Next key
                        
                        ' Update the shape text if changes were made
                        If modifiedText <> originalText Then
                            shape.TextFrame.TextRange.text = modifiedText
                        End If
                    End If
                End If
            Next shape
        Next s
        
        ' Move to the next set of slides
        slideIndex = slideIndex + maxChunks
NextSlide:
    Loop
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description & " on Slide " & slideIndex
    Resume NextSlide
End Sub

Function SplitTextIntoChunksIfNeeded(text As String, maxLines As Long) As Variant
    Dim lines() As String
    lines = Split(text, vbCrLf)
    
    If UBound(lines) + 1 <= maxLines Then
        ' If the text fits within the maximum lines, return it as a single chunk
        Dim singleChunk(1 To 1) As String
        singleChunk(1) = text
        SplitTextIntoChunksIfNeeded = singleChunk
        Exit Function
    End If
    
    ' If the text is longer than maxLines, split it into chunks
    Dim chunks() As String
    ReDim chunks(1 To 1)
    
    Dim currentChunk As String
    Dim lineCount As Long
    Dim chunkIndex As Long
    Dim i As Long
    
    chunkIndex = 1
    
    For i = LBound(lines) To UBound(lines)
        If lineCount = maxLines Then
            ' Add the current chunk and start a new one
            chunks(chunkIndex) = currentChunk
            chunkIndex = chunkIndex + 1
            ReDim Preserve chunks(1 To chunkIndex)
            currentChunk = ""
            lineCount = 0
        End If
        
        ' Add the line to the current chunk
        If lineCount > 0 Then currentChunk = currentChunk & vbCrLf
        currentChunk = currentChunk & lines(i)
        lineCount = lineCount + 1
    Next i
    
    ' Add the last chunk
    If Len(currentChunk) > 0 Then
        chunks(chunkIndex) = currentChunk
    End If
    
    SplitTextIntoChunksIfNeeded = chunks
End Function

' Helper function to get a shape by name within a slide
Function GetShapeByName(slide As Object, shapeName As String) As Object
    Dim shp As Object
    For Each shp In slide.Shapes
        If shp.name = shapeName Then
            Set GetShapeByName = shp
            Exit Function
        End If
    Next shp
    Set GetShapeByName = Nothing
End Function


Sub InsertSpecialtySlidesAndRemove(parsedData As Object, relevantFolder As String, ByRef combinedPresentation As Object)
    Dim pptCurrent As Object
    Dim specialtyArray() As String
    Dim filePath As String
    Dim specialty As Variant
    Dim slideIndex As Integer
    Dim slideText As String
    Dim targetSlideIndex As Integer
    Dim originalPlaceholderSlide As Object
    Dim i As Integer
    
    ' Get Specialty data from parsed JSON and split it into an array
    specialtyArray = Split(parsedData("Specialty"), ",")

    ' Find the placeholder slide index containing "{{Relevant Experience}}"
    Set originalPlaceholderSlide = Nothing
    For i = 1 To combinedPresentation.Slides.Count
        slideText = GetSlideText(combinedPresentation.Slides(i))
        If InStr(1, slideText, "{{Relevant Experience}}", vbTextCompare) > 0 Then
            Set originalPlaceholderSlide = combinedPresentation.Slides(i)
            Exit For
        End If
    Next i

    ' If the placeholder is found, proceed with inserting relevant slides
    If Not originalPlaceholderSlide Is Nothing Then
        ' Get the original slide index
        targetSlideIndex = originalPlaceholderSlide.slideIndex

        ' Loop through each specialty and merge corresponding PPT files after the placeholder slide
        For Each specialty In specialtyArray
            specialty = Trim(specialty)
            
Debug.Print ("specialty found: " & specialty)

            filePath = relevantFolder & "\" & specialty & ".pptx"

            ' Check if the file exists
            If Dir(filePath) <> "" Then
                ' Open the specialty presentation
                Set pptCurrent = combinedPresentation.Application.Presentations.Open(filePath, ReadOnly:=True)

                ' Copy slides from the current presentation and insert them after the placeholder slide
                For slideIndex = 1 To pptCurrent.Slides.Count
                    pptCurrent.Slides(slideIndex).Copy
                    combinedPresentation.Slides.Paste (targetSlideIndex + 1)  ' Insert slides after the placeholder slide
                    targetSlideIndex = targetSlideIndex + 1
                Next slideIndex

                ' Close the current presentation
                pptCurrent.Close
            Else
                Debug.Print "File not found for specialty: " & specialty
            End If
        Next specialty

        ' Delete the original placeholder slide after inserting relevant slides
        originalPlaceholderSlide.Delete

    Else
        MsgBox "Placeholder '{{Relevant Experience}}' not found in the presentation.", vbExclamation, "Placeholder Missing"
    End If

    'MsgBox "Specialty slides successfully inserted and placeholder slide removed.", vbInformation, "Insertion Complete"
End Sub

' Helper function to get the concatenated text of all shapes on a slide
Private Function GetSlideText(slide As Object) As String
    Dim shape As Object
    Dim allText As String
    allText = ""
    
    ' Loop through all shapes in the slide
    For Each shape In slide.Shapes
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                allText = allText & shape.TextFrame.TextRange.text & vbCrLf
            End If
        End If
    Next shape
    
    GetSlideText = allText
End Function

