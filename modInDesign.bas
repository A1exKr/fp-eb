Attribute VB_Name = "modInDesign"
'=== modInDesign.bas ===
Option Explicit

'==========================
' ENTRY POINT
'==========================
Public Sub RunGenerateINDD()
    Dim app As Object, doc As Object
    Dim inddTemplate As String, outFolder As String, outINDD As String, outPdf As String

    Set app = GetInDesignApp()
    If app Is Nothing Then
        MsgBox "Adobe InDesign is not installed or cannot be automated.", vbCritical
        Exit Sub
    End If

    ' 1) Resolve paths
    inddTemplate = ResolveINDDTemplatePath()   ' base .indd/.indt
    If Len(inddTemplate) = 0 Then Exit Sub

    outFolder = ResolveOutputFolder()
    If Len(outFolder) = 0 Then Exit Sub

    outINDD = outFolder & "\FeeProposal_" & SafeSlug(ProjectName()) & ".indd"
    outPdf = outFolder & "\FeeProposal_" & SafeSlug(ProjectName()) & ".pdf"

    On Error GoTo CleanFail

    ' 2) Open template
    Set doc = app.Open(inddTemplate)

    ' 3) Fill text placeholders (stories, text frames, variables)
    INDD_FillText doc

    ' 4) Place project images/diagrams into named frames
    INDD_PlaceImages doc

    ' 5) Insert tables (e.g., fees, schedule) from Excel ranges
    INDD_InsertTables doc

    ' 6) Do any master-page toggles / optional pages (scope vs options)
    INDD_ApplyOptionalSections doc

    ' 7) Save as INDD
    doc.Save outINDD

    ' 8) Export PDF (preset fallback-safe)
    INDD_ExportPDF doc, outPdf

    ' 9) Optionally ÅgPackageÅh the file for submission
    'INDD_Package doc, outFolder & "\Package"

    ' 10) Cleanup
    doc.Close False
    app.Quit

    MsgBox "InDesign FP created:" & vbCrLf & outINDD & vbCrLf & outPdf, vbInformation
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not app Is Nothing Then app.Quit
    MsgBox "Failed to generate InDesign FP: " & Err.Description, vbCritical
End Sub

'==========================
' INDESIGN APP
'==========================
'==========================
' INDESIGN APP (robust for 2022/17.x)
'==========================
Private Function GetInDesignApp() As Object
    Dim app As Object
    Dim progIDs As Variant
    Dim i As Long

    ' Try specific versions first, then generic
    progIDs = Array( _
        "InDesign.Application.19", _
        "InDesign.Application.CC.2024", _
        "InDesign.Application.18", _
        "InDesign.Application.CC.2023", _
        "InDesign.Application.17", _
        "InDesign.Application.CC.2022", _
        "InDesign.Application" _
    )

    On Error Resume Next
    For i = LBound(progIDs) To UBound(progIDs)
        Set app = CreateObject(progIDs(i))
        If Not app Is Nothing Then Exit For
    Next i
    On Error GoTo 0

    Set GetInDesignApp = app
End Function



'==========================
' PATH RESOLVERS
'==========================
' Return the full path to the InDesign FP base template.
' Looks for a project-specific file first, then a default.
Public Function ResolveINDDTemplatePath() As String
    Dim fso As Object
    Dim projectFolder As String
    Dim candidate As String
    Dim defaultTemplate As String

    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    projectFolder = ThisWorkbook.path

    ' 1) Exact file in project root
    candidate = projectFolder & "\230531_NIKKEN_Proposal.indd"
    If fso.fileExists(candidate) Then ResolveINDDTemplatePath = candidate: Exit Function

    ' 2) Support .indt template in project Templates
    candidate = projectFolder & "\Templates\FP-Base.indt"
    If fso.fileExists(candidate) Then ResolveINDDTemplatePath = candidate: Exit Function

    ' 3) Or .indd in Templates
    candidate = projectFolder & "\Templates\FP-Base.indd"
    If fso.fileExists(candidate) Then ResolveINDDTemplatePath = candidate: Exit Function

    ' 4) Same folder as workbook
    candidate = fso.GetParentFolderName(ThisWorkbook.FullName) & "\FP-Base.indd"
    If fso.fileExists(candidate) Then ResolveINDDTemplatePath = candidate: Exit Function

    ' 5) Corporate default (adjust as needed)
    defaultTemplate = "C:\Nikken\Templates\FP-Base.indt"
    If fso.fileExists(defaultTemplate) Then ResolveINDDTemplatePath = defaultTemplate: Exit Function

    ' Not found
    MsgBox "InDesign template not found." & vbCrLf & _
           "Expected one of: 230531_NIKKEN_Proposal.indd, FP-Base.indt, FP-Base.indd", vbCritical
    ResolveINDDTemplatePath = vbNullString
    Exit Function

ErrHandler:
    MsgBox "Error resolving InDesign template path: " & Err.Description, vbCritical
    ResolveINDDTemplatePath = vbNullString
End Function

' Return (and create if missing) a project-specific output folder.
' [ProjectFolder]\Output\[yyyy-mm-dd]_[ProjectName]\
Public Function ResolveOutputFolder() As String
    Dim fso As Object
    Dim projectFolder As String
    Dim outputRoot As String
    Dim outputFolder As String
    Dim safeProjName As String
    Dim todayStr As String

    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    projectFolder = ThisWorkbook.path

    outputRoot = projectFolder & "\Output"
    If Not fso.FolderExists(outputRoot) Then fso.CreateFolder outputRoot

    todayStr = Format(Date, "yyyy-mm-dd")
    safeProjName = SafeSlug(ProjectName())
    If Len(safeProjName) = 0 Then safeProjName = "Untitled"

    outputFolder = outputRoot & "\" & todayStr & "_" & safeProjName
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    ResolveOutputFolder = outputFolder
    Exit Function

ErrHandler:
    MsgBox "Error creating output folder: " & Err.Description, vbCritical
    ResolveOutputFolder = vbNullString
End Function

'==========================
' SAFE FILENAME SLUG
'==========================
Public Function SafeSlug(ByVal s As String) As String
    Dim cleaned As String
    cleaned = Trim$(s)
    cleaned = Replace(cleaned, "/", "-")
    cleaned = Replace(cleaned, "\", "-")
    cleaned = Replace(cleaned, ":", "-")
    cleaned = Replace(cleaned, "?", "")
    cleaned = Replace(cleaned, "*", "")
    cleaned = Replace(cleaned, """", "")
    cleaned = Replace(cleaned, "<", "")
    cleaned = Replace(cleaned, ">", "")
    cleaned = Replace(cleaned, "|", "")
    cleaned = Replace(cleaned, " ", "_")
    SafeSlug = cleaned
End Function

'==========================
' PROJECT METADATA (STUBS)
' Replace with your real getters, or map to named ranges.
'==========================
Public Function ProjectName() As String
    On Error Resume Next
    ' Example: try a Named Range first
    ProjectName = CStr(ThisWorkbook.Names("ProjectName").RefersToRange.value)
    If Len(ProjectName) = 0 Then
        ' Fallback to workbook name (no extension)
        Dim nm As String
        nm = ThisWorkbook.name
        If InStr(1, nm, ".", vbTextCompare) > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
        ProjectName = nm
    End If
End Function

Public Function ClientName() As String
    On Error Resume Next
    ClientName = CStr(ThisWorkbook.Names("ClientName").RefersToRange.value)
    If Len(ClientName) = 0 Then ClientName = ""
End Function

'==========================
' HIGH-LEVEL CONTENT PIPELINE
'==========================
Public Sub INDD_FillText(ByVal doc As Object)
    ' Variables
    INDD_SetVariable doc, "varProjectName", ProjectName()
    INDD_SetVariable doc, "varClientName", ClientName()
    INDD_SetVariable doc, "varProposalDate", Format$(Date, "yyyy-mm-dd")

    ' Labeled text frames (replace these calls with your builders)
    INDD_SetLabeledFrameText doc, "txtApproach", BuildApproachText()
    INDD_SetLabeledFrameText doc, "txtScope", BuildScopeText()
    INDD_SetLabeledFrameText doc, "txtDeliverables", BuildDeliverablesText()
    INDD_SetLabeledFrameText doc, "txtAssumptions", LoadAssumptions()
    INDD_SetLabeledFrameText doc, "txtExclusions", LoadExclusions()
    INDD_SetLabeledFrameText doc, "txtTerms", LoadTerms()
    INDD_SetLabeledFrameText doc, "txtScheduleNote", BuildScheduleNote()
    INDD_SetLabeledFrameText doc, "txtTeamList", BuildTeamList()
End Sub

Public Sub INDD_PlaceImages(ByVal doc As Object)
    ' NOTE: implement your own resolvers to actual image paths as needed.
    ' Example (guarded):
    '   INDD_PlaceIntoFrame doc, "imgCoverHero", PathToCoverHero()
    '   INDD_PlaceIntoFrame doc, "imgRef1", PathToMixedUseHero()
    ' For now, this is a safe no-op.
End Sub



Public Sub INDD_ApplyOptionalSections(ByVal doc As Object)
    ' Toggle optional pages/masters as needed.
    ' For now, safe no-op.
End Sub

'==========================
' EXPORT / PACKAGE
'==========================
Public Sub INDD_ExportPDF(ByVal doc As Object, ByVal outPdf As String)
    On Error Resume Next

    ' Try using a named preset first (recommended to create "Nikken-FP-PDF")
    Dim pdfPresets As Object, preset As Object
    Set pdfPresets = doc.Application.PDFExportPresets
    If Not pdfPresets Is Nothing Then
        For Each preset In pdfPresets
            If LCase$(CStr(preset.name)) = LCase$("Nikken-FP-PDF") Then
                doc.Export 1952403524, outPdf, False, preset ' idExportFormatPDFType
                Exit Sub
            End If
        Next preset
    End If

    ' Fallback: export with current document PDFExportPreferences
    doc.Export 1952403524, outPdf  ' may use last prefs; set up defaults in template
End Sub

Public Sub INDD_Package(ByVal doc As Object, ByVal targetFolder As String)
    On Error Resume Next
    ' Params: (to match typical needs)
    ' name, copyingFonts, copyingLinkedGraphics, copyingProfiles, updatingGraphics, includingHiddenLayers, ignorePreflightErrors, creatingReport, versionComments, forceSaving
    ' Late-binding friendly call via Package() signature may vary by version;
    ' many teams call the UI-free method on the Document object like this:
    doc.Package targetFolder, True, True, True, True, True, True, True, "Auto-packaged by FP-GEN", True
End Sub

'==========================
' LOW-LEVEL HELPERS (LATE BINDING SAFE)
'==========================
Public Sub INDD_SetVariable(ByVal doc As Object, ByVal variableName As String, ByVal valueText As String)
    On Error Resume Next
    Dim tv As Object
    For Each tv In doc.TextVariables
        If LCase$(CStr(tv.name)) = LCase$(variableName) Then
            tv.VariableOptions.Contents = valueText
            Exit For
        End If
    Next tv
End Sub

Public Sub INDD_SetLabeledFrameText(ByVal doc As Object, ByVal labelName As String, ByVal valueText As String)
    On Error Resume Next
    If Len(labelName) = 0 Then Exit Sub
    Dim it As Object
    For Each it In doc.AllPageItems    ' AllPageItems is robust for late binding
        If LCase$(CStr(it.Label)) = LCase$(labelName) Then
            ' If it's a text frame, set contents
            ' TextFrames expose .Contents or .ParentStory.Contents
            it.Contents = valueText
            Exit For
        End If
    Next it
End Sub

Public Sub INDD_PlaceIntoFrame(ByVal doc As Object, ByVal labelName As String, ByVal filePath As String)
    On Error Resume Next
    If Len(labelName) = 0 Or Len(filePath) = 0 Then Exit Sub
    If Dir$(filePath, vbNormal) = "" Then Exit Sub

    Dim it As Object, g As Object
    For Each it In doc.AllPageItems
        If LCase$(CStr(it.Label)) = LCase$(labelName) Then
            it.Place filePath   ' places into the frame
            ' Try best-effort fit without relying on enum constants
            ' Many builds have Fit method; if not, ignore
            ' Common options: idContentToFrame = 1667591796, idProportionally = 1952400175
            On Error Resume Next
            it.Fit 1667591796      ' Content to Frame (best effort)
            it.Fit 1952400175      ' Proportionally (best effort)
            Exit For
        End If
    Next it
End Sub

Public Sub INDD_InsertScheduleFromExcel(ByVal doc As Object, ByVal tableLabel As String, ByVal excelPath As String)
    ' Example scaffold if you decide to paste from clipboard later.
    ' For now, safe no-op.
End Sub


'==========================
' PATH / ITEM UTILITIES
'==========================
Private Function PathJoin(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Or Right$(a, 1) = "/" Then
        PathJoin = a & b
    Else
        PathJoin = a & "\" & b
    End If
End Function

Private Function FirstExistingPath(ParamArray candidates() As Variant) As String
    Dim i As Long, p As String
    For i = LBound(candidates) To UBound(candidates)
        p = CStr(candidates(i))
        If Len(p) > 0 Then If Dir$(p, vbNormal) <> "" Then FirstExistingPath = p: Exit Function
    Next i
    FirstExistingPath = vbNullString
End Function

Private Function INDD_FindLabeledPageItem(ByVal doc As Object, ByVal labelName As String) As Object
    On Error Resume Next
    Dim it As Object
    For Each it In doc.AllPageItems
        If LCase$(CStr(it.Label)) = LCase$(labelName) Then
            Set INDD_FindLabeledPageItem = it
            Exit Function
        End If
    Next it
End Function

Private Function INDD_GetTableStyle(ByVal doc As Object, ByVal styleName As String) As Object
    On Error Resume Next
    Dim ts As Object, styles As Object
    Set styles = doc.TableStyles
    If Not styles Is Nothing Then
        For Each ts In styles
            If LCase$(CStr(ts.name)) = LCase$(styleName) Then
                Set INDD_GetTableStyle = ts
                Exit Function
            End If
        Next ts
    End If
End Function

'==========================
' IMAGE PATH RESOLVERS
'==========================
Public Function PathToCoverHero() As String
    Dim proj As String, links As String
    proj = ThisWorkbook.path
    links = PathJoin(proj, "Links")

    PathToCoverHero = FirstExistingPath( _
        PathJoin(proj, "Assets\cover_hero.jpg"), _
        PathJoin(proj, "Assets\cover_hero.png"), _
        PathJoin(links, "cover_hero.jpg"), _
        PathJoin(links, "cover_hero.png"), _
        PathJoin(links, "placeholder.jpg"), _
        PathJoin(links, "placeholder.png") _
    )
End Function

Public Function PathToMixedUseHero() As String
    Dim proj As String, links As String
    proj = ThisWorkbook.path
    links = PathJoin(proj, "Links")

    PathToMixedUseHero = FirstExistingPath( _
        PathJoin(proj, "Assets\mixeduse_hero.jpg"), _
        PathJoin(proj, "Assets\mixeduse_hero.png"), _
        PathJoin(links, "mixeduse_hero.jpg"), _
        PathJoin(links, "mixeduse_hero.png"), _
        PathJoin(links, "placeholder.jpg"), _
        PathJoin(links, "placeholder.png") _
    )
End Function

' Add more as needed:
Public Function PathToHospHero() As String
    Dim proj As String, links As String
    proj = ThisWorkbook.path
    links = PathJoin(proj, "Links")
    PathToHospHero = FirstExistingPath( _
        PathJoin(proj, "Assets\hospitality_hero.jpg"), _
        PathJoin(proj, "Assets\hospitality_hero.png"), _
        PathJoin(links, "hospitality_hero.jpg"), _
        PathJoin(links, "hospitality_hero.png"), _
        PathJoin(links, "placeholder.jpg"), _
        PathJoin(links, "placeholder.png") _
    )
End Function


'==========================
' EXCEL RANGE Å® INDD TABLE
'==========================
Public Sub INDD_InsertTableFromRange( _
    ByVal doc As Object, _
    ByVal tableLabel As String, _
    ByVal rng As Excel.Range, _
    Optional ByVal tableStyleName As String = "")

    On Error GoTo Fail

    If rng Is Nothing Then Exit Sub

    ' 1) Find the labeled text frame
    Dim frame As Object
    Set frame = INDD_FindLabeledPageItem(doc, tableLabel)
    If frame Is Nothing Then Exit Sub

    ' 2) Build TSV (tab-separated values) from Excel range
    Dim r As Long, c As Long, rows As Long, cols As Long
    Dim arr As Variant, tsv As String
    arr = rng.Value2
    If IsArray(arr) Then
        rows = UBound(arr, 1): cols = UBound(arr, 2)
        For r = 1 To rows
            For c = 1 To cols
                tsv = tsv & SanitizeTSVCell(arr(r, c))
                If c < cols Then tsv = tsv & vbTab
            Next c
            If r < rows Then tsv = tsv & vbCr
        Next r
    Else
        ' Single cell
        tsv = SanitizeTSVCell(arr)
    End If

    ' 3) Write TSV into the frame and convert to a table
    frame.Contents = tsv
    ' Convert the whole text to a table:
    frame.Texts(1).ConvertToTable vbTab, vbCr

    ' 4) Apply a table style if provided & exists
    If Len(tableStyleName) > 0 Then
        Dim ts As Object
        Set ts = INDD_GetTableStyle(doc, tableStyleName)
        If Not ts Is Nothing Then
            ' The first (and only) table in the story:
            frame.Tables(1).AppliedTableStyle = ts
        End If
    End If

    Exit Sub
Fail:
    MsgBox "Failed to insert table '" & tableLabel & "': " & Err.Description, vbCritical
End Sub

Private Function SanitizeTSVCell(v As Variant) As String
    Dim s As String
    On Error Resume Next
    If IsError(v) Then
        s = ""
    Else
        s = CStr(v)
    End If
    ' remove embedded tabs/newlines that would break the TSV structure
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    SanitizeTSVCell = s
End Function


Public Sub INDD_InsertScheduleFromWorkbook(ByVal doc As Object, ByVal tableLabel As String, ByVal xlsxPath As String, Optional ByVal namedRange As String = "ScheduleTable", Optional ByVal tableStyleName As String = "tblSchedule")
    On Error GoTo Fail
    If Dir$(xlsxPath, vbNormal) = "" Then Exit Sub

    Dim awb As Excel.Workbook, wasVis As Boolean
    wasVis = Excel.Application.Visible
    Excel.Application.Visible = False

    Set awb = Excel.Application.Workbooks.Open(fileName:=xlsxPath, ReadOnly:=True)
    Dim rng As Excel.Range

    On Error Resume Next
    Set rng = awb.Names(namedRange).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then
        ' Fallback: take used range of the first sheet
        Set rng = awb.Worksheets(1).UsedRange
    End If

    If Not rng Is Nothing Then
        INDD_InsertTableFromRange doc, tableLabel, rng, tableStyleName
    End If

    awb.Close SaveChanges:=False
    Excel.Application.Visible = wasVis
    Exit Sub
Fail:
    MsgBox "Failed to insert schedule from workbook: " & Err.Description, vbCritical
    On Error Resume Next
    If Not awb Is Nothing Then awb.Close SaveChanges:=False
    Excel.Application.Visible = True
End Sub

Public Sub INDD_InsertTables(ByVal doc As Object)
    Dim schedPath As String
    schedPath = PathJoin(ThisWorkbook.path, "SCHEDULE.xlsx")
    If Dir$(schedPath, vbNormal) <> "" Then
        INDD_InsertScheduleFromWorkbook doc, "tblSchedule", schedPath, "ScheduleTable", "tblSchedule"
    End If

    ' Example for fees from a live Excel range:
    ' INDD_InsertTableFromRange doc, "tblFees", Sheet1.Range("FeesTable"), "tblFees"
End Sub

'==========================
' EXCEL RANGE Å® PNG Å® PLACE
'==========================
Public Function ExportRangeAsPng(ByVal rng As Excel.Range, ByVal outPath As String) As Boolean
    On Error GoTo Fail
    Dim shp As Excel.shape, ch As Excel.ChartObject
    rng.CopyPicture xlScreen, xlPicture

    ' Create a temporary chart to paste the picture into, then export
    Set ch = rng.Worksheet.ChartObjects.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width, Height:=rng.Height)
    ch.Activate
    ch.Chart.Paste
    ch.Chart.Export fileName:=outPath, FilterName:="PNG"
    ch.Delete

    ExportRangeAsPng = (Dir$(outPath, vbNormal) <> "")
    Exit Function
Fail:
    ExportRangeAsPng = False
End Function

Public Sub INDD_InsertScheduleAsImage(ByVal doc As Object, ByVal labelName As String, ByVal rng As Excel.Range)
    Dim tmpPng As String
    tmpPng = PathJoin(ResolveOutputFolder(), "schedule.png")
    If ExportRangeAsPng(rng, tmpPng) Then
        INDD_PlaceIntoFrame doc, labelName, tmpPng
    End If
End Sub












'==========================
' TEXT BUILDERS (STUBS)
' Replace these with your real generators that already exist in your PPT flow.
'==========================
'==========================
' TEXT BUILDERS (STUBS)
' Replace these with your real generators that already exist in your PPT flow.
'==========================
Public Function BuildApproachText() As String
    BuildApproachText = ""
End Function

Public Function BuildScopeText() As String
    BuildScopeText = ""
End Function

Public Function BuildDeliverablesText() As String
    BuildDeliverablesText = ""
End Function

Public Function LoadAssumptions() As String
    LoadAssumptions = ""
End Function

Public Function LoadExclusions() As String
    LoadExclusions = ""
End Function

Public Function LoadTerms() As String
    LoadTerms = ""
End Function

Public Function BuildScheduleNote() As String
    BuildScheduleNote = ""
End Function

Public Function BuildTeamList() As String
    BuildTeamList = ""
End Function



