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
'Version 5#

Option Explicit

' Store the AI response JSON at module level
Private mRfpAnalysisJson As String

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
    
    ' Clear any previous data
    mRfpAnalysisJson = ""
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
        
        ' Clear any previous analysis when new file is selected
        mRfpAnalysisJson = ""
        Me.cmdReviewFinalize.Visible = False
    Else
        MsgBox "No file selected.", vbInformation, "Upload RFP"
    End If
    
    ' Clean up
    Set fd = Nothing
End Sub

' Event handler for cmdAnalyseRFP: Initiates RFP analysis (ONE AI CALL - GETS ALL DATA)
Private Sub cmdAnalyseRFP_Click()
    ' Disable buttons to prevent multiple clicks
    Me.cmdAnalyseRFP.Enabled = False
    Me.cmdReviewFinalize.Enabled = False
    Me.optPPT.Enabled = False
    Me.optINDD.Enabled = False
    Me.cmdUploadRFP.Enabled = False
    Me.txtRFPsyn.text = ""
    
    ' Display the status message
    Me.lblStatus.Caption = "Analyzing RFP and preparing all data..."
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

' Event handler for cmdReviewFinalize: Opens review window (NO AI CALL - USES STORED DATA)
Private Sub cmdReviewFinalize_Click()
    On Error GoTo ErrHandler
    
    Debug.Print "=== cmdReviewFinalize_Click START ==="
    
    ' Check if we have parsed data
    If Len(mRfpAnalysisJson) = 0 Then
        MsgBox "Please analyze the RFP first by clicking 'Analyse RFP'.", vbExclamation
        Exit Sub
    End If

    ' Disable buttons during processing
    Me.cmdReviewFinalize.Enabled = False
    Me.cmdAnalyseRFP.Enabled = False
    Me.cmdUploadRFP.Enabled = False
    
    ' Show status
    Me.lblStatus.Caption = "Opening Review & Finalize..."
    Me.lblStatus.Visible = True
    DoEvents

    ' Use the stored JSON data - NO AI CALL HERE!
    Debug.Print "Opening review with stored JSON data (length: " & Len(mRfpAnalysisJson) & ")"
    modReviewBridge.OpenReviewWithJson mRfpAnalysisJson
    
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

' Event handler for cmdGeneratePPT: Initiates PPT generation
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

' Processing Subroutine for RFP Analysis - MAKES THE SINGLE AI CALL
Public Sub RunRFPAnalysis()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== RunRFPAnalysis START ==="
    
    Dim filePath As String
    Dim readFile As String
    Dim fileExtension As String
    
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
    
    If Len(readFile) = 0 Then
        MsgBox "Could not extract text from the document.", vbExclamation
        GoTo ErrorHandler
    End If
    
    ' Step 2: Call AI ONCE to get complete JSON with ALL data needed for review
    Debug.Print "Calling AI to parse RFP comprehensively..."
    Me.lblStatus.Caption = "Calling AI to parse entire RFP..."
    DoEvents
    
    ' Call the comprehensive parser that returns the full JSON structure
    ' This should be the SAME call that Parser.ParseAndOpenReview uses
    mRfpAnalysisJson = Parser.ParseRFPToJson(readFile)
    
    Debug.Print "AI parsing complete. JSON length: " & Len(mRfpAnalysisJson)
    
    ' Step 3: For backward compatibility, also create the old dictionary format
    ' (if your old code still needs parsedDataSet)
    Set parsedDataSet = ParseRFPText(readFile)
    
    ' Step 4: Show quick preview in the form
    Me.txtRFPsyn.Visible = True
    
    ' Try to extract a summary from the JSON for preview
    On Error Resume Next
    Dim tempDict As Object
    Set tempDict = JsonConverter.ParseJSON(mRfpAnalysisJson)
    
    If Not tempDict Is Nothing Then
        ' Build comprehensive executive summary
        Dim summary As String
        summary = "EXECUTIVE SUMMARY" & vbCrLf
        summary = summary & String(60, "=") & vbCrLf & vbCrLf
        
        ' PROJECT INFORMATION
        If tempDict.Exists("project") Or tempDict.Exists("client") Then
            summary = summary & "PROJECT INFORMATION:" & vbCrLf
            
            If tempDict.Exists("project") And TypeName(tempDict("project")) = "Dictionary" Then
                If tempDict("project").Exists("name") Then
                    summary = summary & "  " & Chr(149) & " Project: " & tempDict("project")("name") & vbCrLf
                End If
                If tempDict("project").Exists("type") Then
                    summary = summary & "  " & Chr(149) & " Type: " & tempDict("project")("type") & vbCrLf
                End If
                If tempDict("project").Exists("location") Then
                    summary = summary & "  " & Chr(149) & " Location: " & tempDict("project")("location") & vbCrLf
                End If
                If tempDict("project").Exists("siteArea") Then
                    summary = summary & "  " & Chr(149) & " Site Area: " & tempDict("project")("siteArea") & vbCrLf
                End If
            End If
            
            If tempDict.Exists("client") And TypeName(tempDict("client")) = "Dictionary" Then
                If tempDict("client").Exists("name") Then
                    summary = summary & "  " & Chr(149) & " Client: " & tempDict("client")("name") & vbCrLf
                End If
            End If
            
            summary = summary & vbCrLf
        End If
        
        ' PROJECT UNDERSTANDING
        If tempDict.Exists("understanding") Then
            Dim understanding As String
            understanding = ""
            
            If TypeName(tempDict("understanding")) = "Dictionary" Then
                If tempDict("understanding").Exists("understanding") Then
                    understanding = tempDict("understanding")("understanding")
                ElseIf tempDict("understanding").Exists("text") Then
                    understanding = tempDict("understanding")("text")
                End If
            ElseIf TypeName(tempDict("understanding")) = "String" Then
                understanding = tempDict("understanding")
            End If
            
            If Len(understanding) > 0 Then
                summary = summary & "PROJECT UNDERSTANDING:" & vbCrLf
                summary = summary & "  " & understanding & vbCrLf & vbCrLf
            End If
        End If
        
        ' SCHEDULE
        If tempDict.Exists("schedule") And TypeName(tempDict("schedule")) = "Dictionary" Then
            Dim hasScheduleInfo As Boolean
            hasScheduleInfo = False
            
            If tempDict("schedule").Exists("totalWeeks") Or tempDict("schedule").Exists("milestones") Then
                summary = summary & "SCHEDULE:" & vbCrLf
                hasScheduleInfo = True
                
                If tempDict("schedule").Exists("totalWeeks") Then
                    summary = summary & "  " & Chr(149) & " Duration: " & tempDict("schedule")("totalWeeks") & " weeks" & vbCrLf
                End If
                
                If tempDict("schedule").Exists("milestones") Then
                    Dim milestones As Object
                    Set milestones = tempDict("schedule")("milestones")
                    If TypeName(milestones) = "Collection" And milestones.Count > 0 Then
                        Dim milestoneStr As String, i As Long
                        milestoneStr = ""
                        For i = 1 To milestones.Count
                            If i > 1 Then milestoneStr = milestoneStr & ", "
                            milestoneStr = milestoneStr & milestones(i)
                        Next i
                        summary = summary & "  " & Chr(149) & " Milestones: " & milestoneStr & vbCrLf
                    End If
                End If
                
                If hasScheduleInfo Then summary = summary & vbCrLf
            End If
        End If
        
        ' KEY TEAM MEMBERS
        If tempDict.Exists("team") And TypeName(tempDict("team")) = "Dictionary" Then
            Dim hasTeamInfo As Boolean
            hasTeamInfo = False
            
            If tempDict("team").Exists("principal") Or tempDict("team").Exists("pm") Then
                summary = summary & "KEY TEAM MEMBERS:" & vbCrLf
                hasTeamInfo = True
                
                If tempDict("team").Exists("principal") And TypeName(tempDict("team")("principal")) = "Dictionary" Then
                    Dim principal As Object
                    Set principal = tempDict("team")("principal")
                    If principal.Exists("name") Then
                        Dim principalText As String
                        principalText = principal("name")
                        If principal.Exists("title") Then
                            principalText = principalText & " (" & principal("title") & ")"
                        End If
                        summary = summary & "  " & Chr(149) & " Principal: " & principalText & vbCrLf
                    End If
                End If
                
                If tempDict("team").Exists("pm") And TypeName(tempDict("team")("pm")) = "Dictionary" Then
                    Dim pm As Object
                    Set pm = tempDict("team")("pm")
                    If pm.Exists("name") Then
                        Dim pmText As String
                        pmText = pm("name")
                        If pm.Exists("title") Then
                            pmText = pmText & " (" & pm("title") & ")"
                        End If
                        summary = summary & "  " & Chr(149) & " Project Manager: " & pmText & vbCrLf
                    End If
                End If
                
                If hasTeamInfo Then summary = summary & vbCrLf
            End If
        End If
        
        ' KEY SCOPE ITEMS
        If tempDict.Exists("scope") And TypeName(tempDict("scope")) = "Dictionary" Then
            If tempDict("scope").Exists("scopeList") Then
                Dim scopeList As Object
                Set scopeList = tempDict("scope")("scopeList")
                If TypeName(scopeList) = "Collection" And scopeList.Count > 0 Then
                    summary = summary & "KEY SCOPE ITEMS:" & vbCrLf
                    Dim scopeCount As Long
                    scopeCount = 0
                    For i = 1 To scopeList.Count
                        If scopeCount >= 3 Then Exit For ' Show max 3 items
                        summary = summary & "  " & Chr(149) & " " & scopeList(i) & vbCrLf
                        scopeCount = scopeCount + 1
                    Next i
                    summary = summary & vbCrLf
                End If
            End If
            
            ' MAJOR DELIVERABLES
            If tempDict("scope").Exists("deliverablesList") Then
                Dim deliverablesList As Object
                Set deliverablesList = tempDict("scope")("deliverablesList")
                If TypeName(deliverablesList) = "Collection" And deliverablesList.Count > 0 Then
                    summary = summary & "MAJOR DELIVERABLES:" & vbCrLf
                    Dim delivCount As Long
                    delivCount = 0
                    For i = 1 To deliverablesList.Count
                        If delivCount >= 3 Then Exit For ' Show max 3 items
                        summary = summary & "  " & Chr(149) & " " & deliverablesList(i) & vbCrLf
                        delivCount = delivCount + 1
                    Next i
                End If
            End If
        End If
        
        If Len(summary) > 0 Then
            Me.txtRFPsyn.text = summary
        Else
            Me.txtRFPsyn.text = "RFP analyzed successfully. All data ready for review."
        End If
    Else
        Me.txtRFPsyn.text = "RFP analyzed successfully. Click 'Review & Finalize' to continue."
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
    ' Clear the stored JSON on error
    mRfpAnalysisJson = ""
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

' Ensure proper cleanup when the UserForm is closed
Private Sub UserForm_Terminate()
    ' Hide the status label
    Me.lblStatus.Visible = False
    ' Clear stored data
    mRfpAnalysisJson = ""
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
