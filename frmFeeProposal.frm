VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFeeProposal 
   Caption         =   "the FP-GEN (Fee Proposal Generator)"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   OleObjectBlob   =   "frmFeeProposal.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFeeProposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Event handler for cmdExit: Exits the program
Private Sub cmdExit_Click()
    ' Confirm exit with the user
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit Program")
    
    If response = vbYes Then
        ' Unload the UserForm and quit the application if necessary
        Unload Me
        ' If you want to quit the entire Excel application, uncomment the following line:
        ' Application.Quit
    End If
End Sub







Private Sub cmdRelExp_Click()

    Call frmSearchForm.InitializeSearchForm
    Debug.Print "3 InitializeSearchForm"
    frmSearchForm.Show

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



Public Sub UserForm_Initialize()
    ' Initialize the UserForm
    ' Hide the status label initially
    'lblStatus.Visible = False
    
    ' Hide other controls initially if necessary
    txtRFPsyn.Visible = False
    cmdAnalyseRFP.Visible = False
    cmdReviewFinalize.Visible = False
    optPPT.Visible = False
    optINDD.Visible = False

    
End Sub

' Helper function to force scroll bars to appear
Private Sub ShowScrollBarsImmediately(txt As MSForms.TextBox)
    Dim txtLen As Long
    txtLen = Len(txt.text)

    ' Force the TextBox to show the scroll bars by manipulating the SelStart property
    txt.SelStart = txtLen
    txt.SelLength = 0

    ' Add DoEvents to ensure the UI updates
    DoEvents

    ' Move caret back to start (optional)
    txt.SelStart = 0
    txt.SelLength = 0
End Sub

' Event handler for cmdUploadRFP: Allows user to select an RFP file
Private Sub cmdUploadRFP_Click()
    Dim fd As fileDialog
    Dim filePath As String
    
    ' Initialize the FileDialog object
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    ' Set the dialog title
    fd.title = "Select RFP Document"
    
    ' Clear any existing filters
    fd.Filters.Clear
    
    ' Add filters for supported file types
    fd.Filters.Add "All Supported Files", "*.docx; *.pdf; *.ppt; *.pptx"
    fd.Filters.Add "Word Documents", "*.docx"
    fd.Filters.Add "PDF Files", "*.pdf"
    fd.Filters.Add "PowerPoint Files", "*.ppt; *.pptx"
    
    ' Show the dialog and capture user selection
    If fd.Show = -1 Then ' User clicked OK
        filePath = fd.selectedItems(1)
        ' Set the selected file path to the text box
        Me.txtRFPPath.Visible = True
        Me.txtRFPPath.text = filePath
        ' Make the Analyse button visible/enabled
        Me.cmdAnalyseRFP.Visible = True
    Else
        MsgBox "No file selected.", vbInformation, "Upload RFP"
    End If
    
    ' Clean up
    Set fd = Nothing
End Sub

' Event handler for cmdAnalyseRFP: Initiates RFP analysis
Private Sub cmdAnalyseRFP_Click()
    ' Disable buttons to prevent multiple clicks
    Me.cmdAnalyseRFP.Enabled = False
    Me.cmdReviewFinalize.Enabled = False
    Me.optPPT.Enabled = False
    Me.optINDD.Enabled = False
    Me.cmdUploadRFP.Enabled = False
    Me.txtRFPsyn.text = ""
    
    ' Display the status message
    Me.lblStatus.Caption = "Awaiting response..."
    Me.lblStatus.Visible = True
    
    ' Refresh the UI to show the status message
    DoEvents
    
    ' Call the RFP analysis subroutine
    Call RunRFPAnalysis
    
    ' Hide the status message after processing
    Me.lblStatus.Visible = False
    
    ' Re-enable buttons
    Me.cmdAnalyseRFP.Enabled = True
    Me.cmdReviewFinalize.Enabled = True
    Me.cmdUploadRFP.Enabled = True
    Me.cmdExit.Enabled = True
    
    Me.optPPT.Enabled = True
    Me.optINDD.Enabled = True
    
End Sub

' Event handler for cmdReviewFinalize: Initiates PPT generation
Private Sub cmdGeneratePPT_Click()
    ' Disable buttons to prevent multiple clicks
    Me.cmdReviewFinalize.Enabled = False
    Me.cmdAnalyseRFP.Enabled = False
    Me.cmdUploadRFP.Enabled = False
    Me.cmdExit.Enabled = False
    Me.optPPT.Enabled = False
    Me.optINDD.Enabled = False
    
    
    
    
    ' Display the status message
    Me.lblStatus.Caption = "Generating a Fee Proposal..."
    Me.lblStatus.Visible = True
    
    ' Refresh the UI to show the status message
    DoEvents
    
    If frmFeeProposal.optINDD.value Then
       ' Call RunGenerateINDD
    Else
        ' Call the PPT generation subroutine
        Call RunGeneratePPT
    End If
    

    
    ' Hide the status message after processing
    Me.lblStatus.Visible = False
    
    ' Re-enable buttons
    Me.cmdReviewFinalize.Enabled = True
    Me.cmdAnalyseRFP.Enabled = True
    Me.cmdUploadRFP.Enabled = True
    
    Me.optINDD.Enabled = True
    Me.cmdUploadRFP.Enabled = True
    
End Sub

Private Sub cmdReviewFinalize_Click()
    On Error GoTo ErrHandler
    
    Debug.Print "=== cmdReviewFinalize_Click START ==="
    
    Dim filePath As String
    Dim fileText As String
    Dim fileExt As String

    ' Disable buttons during processing
    Me.cmdReviewFinalize.Enabled = False
    Me.cmdAnalyseRFP.Enabled = False
    Me.cmdUploadRFP.Enabled = False
    
    ' Show status
    Me.lblStatus.Caption = "Opening Review & Finalize..."
    Me.lblStatus.Visible = True
    DoEvents

    ' Get file path
    filePath = Me.txtRFPPath.text
    Debug.Print "File path: " & filePath
    
    If Len(filePath) = 0 Then
        MsgBox "Please upload an RFP file first.", vbExclamation
        GoTo CleanExit
    End If

    ' Extract text based on file type
    fileExt = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    Debug.Print "File extension: " & fileExt
    
    Select Case fileExt
        Case "docx"
            Debug.Print "Reading Word document..."
            fileText = ReadWordDocument(filePath)
        Case "pdf"
            Debug.Print "Reading PDF document..."
            fileText = ReadPDFDocument(filePath)
        Case "ppt", "pptx"
            Debug.Print "Reading PowerPoint document..."
            fileText = ReadPowerPointDocument(filePath)
        Case Else
            MsgBox "Unsupported file format: " & fileExt, vbExclamation
            GoTo CleanExit
    End Select
    
    Debug.Print "Extracted text length: " & Len(fileText)
    
    If Len(fileText) = 0 Then
        MsgBox "Could not extract text from file.", vbExclamation
        GoTo CleanExit
    End If

    ' Update status
    Me.lblStatus.Caption = "Calling AI to parse RFP..."
    DoEvents

    ' Pass the extracted text to Parser
    Debug.Print "Calling Parser.ParseAndOpenReview..."
    Parser.ParseAndOpenReview fileText
    
    Debug.Print "=== cmdReviewFinalize_Click END ==="
    
CleanExit:
    ' Hide status and re-enable buttons
    Me.lblStatus.Visible = False
    Me.cmdReviewFinalize.Enabled = True
    Me.cmdAnalyseRFP.Enabled = True
    Me.cmdUploadRFP.Enabled = True
    Exit Sub

ErrHandler:
    Debug.Print "ERROR in cmdReviewFinalize_Click: " & Err.Description
    MsgBox "Error opening Review screen: " & Err.Description, vbCritical
    GoTo CleanExit
End Sub


' Processing Subroutine for RFP Analysis
Public Sub RunRFPAnalysis()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== RunRFPAnalysis START ==="
    
    ' Existing RFP processing code
    Dim filePath As String
    Dim readFile As String
    Dim fileExtension As String
    Dim parsedData As Object
    
    filePath = Me.txtRFPPath.text
    fileExtension = LCase(Right(filePath, Len(filePath) - InStrRev(filePath, ".")))
    Debug.Print "fileExtension = " & fileExtension
    Debug.Print "filePath = " & filePath
    
    ' Step 1: Extract Text from RFP
    Select Case fileExtension
        Case "docx"
            readFile = ReadWordDocument(filePath)
        Case "pdf"
            readFile = ReadPDFDocument(filePath)
        Case "ppt", "pptx"
            readFile = ReadPowerPointDocument(filePath)
        Case Else
            MsgBox "Unsupported file format.", vbExclamation, "Error"
            readFile = ""
    End Select
    
    Debug.Print "Extracted text length: " & Len(readFile)
    
    ' Step 2: Parse Extracted Text using ChatGPT API
    ' NOTE: This creates the OLD dictionary format for quick preview only
    Set parsedData = ParseRFPText(readFile)
    
    ' Update the public variable to use in other functions
    Set parsedDataSet = parsedData
    
    ' Show quick preview in the form
    Me.txtRFPsyn.Visible = True
    
    ' Try to show executive summary if available
    On Error Resume Next
    If parsedData.Exists("understanding.understanding") Then
        Me.txtRFPsyn.text = parsedData("understanding.understanding")
    ElseIf parsedData.Exists("Executive Summary") Then
        Me.txtRFPsyn.text = parsedData("Executive Summary")
    Else
        Me.txtRFPsyn.text = "Analysis complete. Click 'Review & Finalize' to continue."
    End If
    On Error GoTo ErrorHandler
    
    ' Show the Review & Finalize button
    Me.cmdReviewFinalize.Visible = True
    Me.optPPT.Visible = True
    Me.optINDD.Visible = True
    
    Debug.Print "=== RunRFPAnalysis END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR in RunRFPAnalysis: " & Err.Description
    MsgBox "An error occurred during RFP analysis: " & Err.Description, vbCritical, "Error"
    ' Re-enable buttons in case of error
    Me.cmdAnalyseRFP.Enabled = True
    Me.cmdReviewFinalize.Enabled = True
    Me.cmdUploadRFP.Enabled = True
    Me.optPPT.Enabled = True
    Me.optINDD.Enabled = True
    Me.lblStatus.Visible = False
End Sub

' Processing Subroutine for Generating PPT
Public Sub RunGeneratePPT()
    On Error GoTo ErrorHandler
    
    ' Existing PPT generation code
    Call GenerateFeeProposal
    
    ' After processing, the status message is hidden in the button click handler
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred during PPT generation: " & Err.Description, vbCritical, "Error"
    ' Re-enable buttons in case of error
    Me.cmdReviewFinalize.Enabled = True
    Me.cmdAnalyseRFP.Enabled = True
    Me.cmdUploadRFP.Enabled = True
    ' Hide the status message
    Me.lblStatus.Visible = False
End Sub

' Example Placeholder for GenerateFeeProposal
' Rename this subroutine to avoid naming conflicts with the "Generator" module
Private Sub GenerateFeeProposal_Simulated()
    ' Simulate processing time
    Dim i As Long
    For i = 1 To 1000000
        DoEvents ' Keep the UI responsive
    Next i
    ' Processing complete
End Sub

' Ensure proper cleanup when the UserForm is closed
Private Sub UserForm_Terminate()
    ' Hide the status label
    Me.lblStatus.Visible = False
End Sub

' Functions to Read Documents
Private Function ReadWordDocument(filePath As String) As String
    Dim objWord As Object
    Dim objDoc As Object
    Dim content As String
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open(filePath, ReadOnly:=True)
    
    content = objDoc.content.text
    
    objDoc.Close
    objWord.Quit
    
    ReadWordDocument = content
End Function

Private Function ReadPDFDocument(filePath As String) As String

    Dim txtPath As String

    txtPath = GetRelativePath("\test output\document_2.txt")
    
    ReadPDFDocument = PDFtoTxt.ProcessPDF(filePath, txtPath)
    
End Function

Private Function ReadPowerPointDocument(filePath As String) As String
    Dim objPPT As Object
    Dim objPresentation As Object
    Dim objSlide As Object
    Dim objShape As Object
    Dim content As String
    
    Set objPPT = CreateObject("PowerPoint.Application")
    Set objPresentation = objPPT.Presentations.Open(filePath, ReadOnly:=True)
    
    For Each objSlide In objPresentation.Slides
        For Each objShape In objSlide.Shapes
            If objShape.HasTextFrame Then
                If objShape.TextFrame.HasText Then
                    content = content & objShape.TextFrame.TextRange.text & vbNewLine
                End If
            End If
        Next objShape
    Next objSlide
    
    objPresentation.Close
    objPPT.Quit
    
    Debug.Print "ReadPowerPointDocument.content: " & content
    
    ReadPowerPointDocument = content
End Function





