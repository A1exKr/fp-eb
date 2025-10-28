Attribute VB_Name = "modReviewFinalize"
Option Explicit

' State for the Review screen
Public gReview As Object        ' Scripting.Dictionary
Private mCurrentSectionKey As String

Public Enum ESectionStatus
    ssSuggested = 0
    ssEdited = 1
    ssApproved = 2
    ssLocked = 3
End Enum

Public Sub Form_Initialize(ByVal f As Object)
    Static isInitializing As Boolean
    If isInitializing Then Exit Sub
    isInitializing = True
    On Error GoTo CleanExit

    Set gReview = modDataBus.ConsumeRfpData()
    If gReview Is Nothing Then
        MsgBox "No review data found. Run Test_OpenReviewWithInlineMock first.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Bind left list (hidden key in col 0, visible title in col 1)
    With f.ListSections
        .Clear
        Dim order As Variant: order = SectionOrder()
        Dim i As Long
        For i = LBound(order) To UBound(order)
            .AddItem ""
            .List(.ListCount - 1, 0) = order(i)(0)
            .List(.ListCount - 1, 1) = order(i)(1)
        Next
    End With

    ' Initialize first section safely (no event loop)
    If f.ListSections.ListCount > 0 Then
        Application.EnableEvents = False
        f.ListSections.ListIndex = 0
        mCurrentSectionKey = f.ListSections.List(0, 0)
        LoadSectionToForm f, mCurrentSectionKey
        Application.EnableEvents = True
    End If

    BindFee f
    BindSchedule f
    BindTeam f
    BindValidation f

CleanExit:
    isInitializing = False
End Sub


Public Sub Form_SectionSelect(ByVal f As Object)
    On Error GoTo SafeExit
    If f.ListSections.ListIndex < 0 Then Exit Sub

    ' Prevent spurious recursion if still initializing
    Static inSelect As Boolean
    If inSelect Then Exit Sub
    inSelect = True

    ' Save current edits before switching
    SaveCurrentEdits f

    mCurrentSectionKey = f.ListSections.List(f.ListSections.ListIndex, 0)
    LoadSectionToForm f, mCurrentSectionKey

SafeExit:
    inSelect = False
End Sub


Public Sub Form_SaveDraft(ByVal f As Object)
    SaveCurrentEdits f
    Dim json As String: json = ConvertToJson(gReview)
    Dim p As String: p = PickSavePath("Save Review Draft", "review.json")
    If Len(p) > 0 Then
        WriteAllText p, json
        MsgBox "Draft saved:" & vbCrLf & p, vbInformation
    End If
End Sub

Public Sub Form_Validate(ByVal f As Object)
    SaveCurrentEdits f
    Dim issues As Collection
    Set issues = ValidateReview(gReview)
    RenderValidation f, issues
End Sub

Public Sub Form_GenerateFP(ByVal f As Object)
    SaveCurrentEdits f
    Dim issues As Collection
    Set issues = ValidateReview(gReview)

    Dim hasBlockers As Boolean: hasBlockers = False
    Dim it As Variant
    For Each it In issues
        If it("level") = "BLOCKER" Then hasBlockers = True: Exit For
    Next it
    If hasBlockers Then
        RenderValidation f, issues
        MsgBox "Please resolve blockers before generating the final proposal.", vbExclamation
        Exit Sub
    End If

    ' Hand off to existing generator if available
    On Error Resume Next
    Application.Run "Generator_GenerateFP", ConvertToJson(gReview)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' fallback: save JSON so the team can pipe it to the generator they prefer
        Dim p As String
        p = PickSavePath("Export Data for FP Generator", "fpgen_payload.json")
        If Len(p) > 0 Then
            WriteAllText p, ConvertToJson(gReview)
            MsgBox "Data exported to:" & vbCrLf & p & vbCrLf & _
                   "Call your existing generator with this JSON.", vbInformation
        End If
    End If
End Sub

Public Sub Form_Regenerate(ByVal f As Object)
    If mCurrentSectionKey = "" Then Exit Sub
    Dim txt As String: txt = f.mpMain.Pages(0).Controls("txtText").text
    Dim protectNums As Boolean: protectNums = f.mpMain.Pages(0).Controls("chkProtectNumbers").value

    Dim newTxt As String
    On Error Resume Next
    newTxt = Application.Run("APIModule_RegenerateSection", txt, mCurrentSectionKey, protectNums)
    On Error GoTo 0

    If Len(newTxt) = 0 Then
        ' Fallback: trivial rewrite (replace double spaces, trim)
        newTxt = SimpleRewrite(txt)
    End If

    ' Store suggested text alongside edited to allow diff
    gReview("sections")(mCurrentSectionKey)("contentSuggested") = newTxt
    f.mpMain.Pages(0).Controls("txtText").text = newTxt
    SetSectionStatus mCurrentSectionKey, ssEdited
End Sub

Public Sub Form_ShowDiff(ByVal f As Object)
    If mCurrentSectionKey = "" Then Exit Sub
    Dim oldText As String: oldText = NzStr(gReview("sections")(mCurrentSectionKey)("contentEdited"))
    Dim newText As String: newText = NzStr(gReview("sections")(mCurrentSectionKey)("contentSuggested"))
    If Len(newText) = 0 And Len(oldText) = 0 Then
        MsgBox "No text available to diff.", vbInformation
        Exit Sub
    End If
    Dim diff As String: diff = diffText(oldText, newText)
    Dim fv As Object: Set fv = VBA.UserForms.Add("frmDiffViewer")
    fv.ShowDiff diff
End Sub

Public Sub Form_Approve(ByVal f As Object)
    If mCurrentSectionKey = "" Then Exit Sub
    SetSectionStatus mCurrentSectionKey, ssApproved
    UpdateListCaptions f
End Sub

Public Sub Form_Lock(ByVal f As Object)
    If mCurrentSectionKey = "" Then Exit Sub
    SetSectionStatus mCurrentSectionKey, ssLocked
    UpdateListCaptions f
End Sub

Public Sub Form_RecomputeFee(ByVal f As Object)
    On Error Resume Next
    Application.Run "Generator_RecomputeFee", gReview
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' Show current fee model (if any) in the list
        BindFee f
        MsgBox "Recompute hook not found. Populate fee data in gReview(""financial"") and re-open.", vbInformation
    Else
        BindFee f
    End If
End Sub

Public Sub Form_AttachCV(ByVal f As Object)
    On Error Resume Next
    Application.Run "Main_AttachCVToSelectedTeamMember", gReview
    On Error GoTo 0
    BindTeam f
End Sub

' ---------- helpers ----------

Public Sub LoadSectionToForm(ByVal f As Object, ByVal key As String)
    On Error Resume Next
    If gReview Is Nothing Then Exit Sub
    If Not gReview.Exists("sections") Then Exit Sub
    If Not gReview("sections").Exists(key) Then Exit Sub

    Dim sec As Object
    Set sec = gReview("sections")(key)
    Dim content As String
    content = sec("contentEdited")
    If Len(content) = 0 Then content = sec("contentSuggested")

    f.txtText.text = content
End Sub


Private Sub SaveCurrentEdits(ByVal f As Object)
    If mCurrentSectionKey = "" Then Exit Sub
    Dim txt As String: txt = f.mpMain.Pages(0).Controls("txtText").text
    gReview("sections")(mCurrentSectionKey)("contentEdited") = txt
    If NzStr(gReview("sections")(mCurrentSectionKey)("status")) = "" Then
        SetSectionStatus mCurrentSectionKey, ssEdited
    End If
End Sub

Private Sub UpdateListCaptions(ByVal f As Object)
    Dim i As Long
    For i = 0 To f.ListSections.ListCount - 1
        Dim k As String: k = f.ListSections.List(i, 0)
        Dim cap As String: cap = SectionTitle(k) & "  [" & StatusLabel(NzStr(gReview("sections")(k)("status"))) & "]"
        f.ListSections.List(i, 1) = cap ' visible column
    Next
End Sub

Private Function StatusLabel(ByVal status As String) As String
    Select Case LCase$(status)
        Case "0", "suggested": StatusLabel = "Suggested"
        Case "1", "edited": StatusLabel = "Edited"
        Case "2", "approved": StatusLabel = "Approved"
        Case "3", "locked": StatusLabel = "Locked"
        Case Else: StatusLabel = "—""
    End Select
End Function

Private Sub BindFee(ByVal f As Object)
    On Error Resume Next
    Dim lb As Object: Set lb = f.mpMain.Pages(1).Controls("lstFee")
    lb.Clear
    Dim fin As Object
    If gReview.Exists("financial") Then
        Set fin = gReview("financial")
        AddFeeRow lb, "Labor Subtotal", CStr(Nz(fin, "laborSubtotal"))
        AddFeeRow lb, "Overheads & Profit", CStr(Nz(fin, "overheads"))
        AddFeeRow lb, "Subconsultants", CStr(Nz(fin, "subconsultants"))
        AddFeeRow lb, "Reimbursables (Travel, etc.)", CStr(Nz(fin, "reimbursables"))
        AddFeeRow lb, "Total", CStr(Nz(fin, "total"))
    End If
    On Error GoTo 0
End Sub

Private Sub AddFeeRow(lb As Object, name As String, val As String)
    If Len(val) = 0 Then Exit Sub
    lb.AddItem name
    lb.List(lb.ListCount - 1, 1) = val
End Sub

Private Sub BindSchedule(ByVal f As Object)
    On Error Resume Next
    Dim lb As Object: Set lb = f.mpMain.Pages(2).Controls("lstSchedule")
    lb.Clear
    lb.AddItem "Phase": lb.List(lb.ListCount - 1, 1) = "Duration (wks)": lb.List(lb.ListCount - 1, 2) = "Notes"
    If gReview.Exists("schedule") Then
        Dim sch As Object: Set sch = gReview("schedule")
        Dim phases As Object: Set phases = NzDict(sch, "phases")
        Dim k As Variant
        For Each k In phases.keys
            lb.AddItem CStr(k)
            lb.List(lb.ListCount - 1, 1) = CStr(Nz(phases(k), "duration"))
            lb.List(lb.ListCount - 1, 2) = CStr(Nz(phases(k), "notes"))
        Next
    End If
    On Error GoTo 0
End Sub

Private Sub BindTeam(ByVal f As Object)
    On Error Resume Next
    Dim lb As Object: Set lb = f.mpMain.Pages(3).Controls("lstTeam")
    lb.Clear
    If gReview.Exists("team") Then
        Dim t As Object: Set t = gReview("team")
        Dim k As Variant
        For Each k In t.keys
            lb.AddItem CStr(k)
            lb.List(lb.ListCount - 1, 1) = CStr(Nz(t(k), "name"))
        Next
    End If
    On Error GoTo 0
End Sub

Private Sub BindValidation(ByVal f As Object)
    Dim issues As Collection: Set issues = ValidateReview(gReview)
    RenderValidation f, issues
End Sub

Public Sub RenderValidation(ByVal f As Object, ByVal issues As Collection)
    Dim lb As Object: Set lb = f.mpMain.Pages(4).Controls("lstValidation")
    lb.Clear
    Dim it As Variant
    For Each it In issues
        lb.AddItem it("level")
        lb.List(lb.ListCount - 1, 1) = it("message")
    Next
End Sub

Private Sub LoadDefaultSkeleton(ByRef st As Object)
    Set st("sections") = CreateObject("Scripting.Dictionary")
    Dim order As Variant: order = SectionOrder()
    Dim i As Long
    For i = LBound(order) To UBound(order)
        Dim sec As Object: Set sec = CreateObject("Scripting.Dictionary")
        sec("title") = order(i)(1)
        sec("contentSuggested") = ""
        sec("contentEdited") = ""
        sec("status") = ssSuggested
        st("sections")(order(i)(0)) = sec
    Next

    ' Minimal schedule / financial defaults
    Dim sch As Object: Set sch = CreateObject("Scripting.Dictionary")
    Dim phases As Object: Set phases = CreateObject("Scripting.Dictionary")
    phases("Pre-Concept") = DictOf(Array("duration", 8), Array("notes", "Baseline"))
    phases("Concept Design") = DictOf(Array("duration", 12), Array("notes", ""))
    phases("Schematic Design") = DictOf(Array("duration", 24), Array("notes", ""))
    sch("phases") = phases
    st("schedule") = sch

    Dim fin As Object: Set fin = CreateObject("Scripting.Dictionary")
    fin("laborSubtotal") = ""
    fin("overheads") = ""
    fin("subconsultants") = ""
    fin("reimbursables") = ""
    fin("total") = ""
    st("financial") = fin

    ' simple team
    Dim team As Object: Set team = CreateObject("Scripting.Dictionary")
    team("Project Principal") = DictOf(Array("name", ""))
    team("Project Manager") = DictOf(Array("name", ""))
    st("team") = team

    st("metadata") = DictOf(Array("projectName", ""), Array("clientName", ""), Array("location", ""), Array("currency", "USD"))
End Sub

Public Sub SetSectionStatus(ByVal key As String, ByVal s As ESectionStatus)
    gReview("sections")(key)("status") = s
End Sub

Public Function SectionOrder() As Variant
    ' key, title
    SectionOrder = Array( _
        Array("cover_letter", "Cover Letter"), _
        Array("project_understanding", "Project Understanding"), _
        Array("methodology", "Methodology"), _
        Array("scope_deliverables", "Scope & Deliverables"), _
        Array("schedule", "Schedule"), _
        Array("team", "Team Structure"), _
        Array("financial", "Financial Proposal"), _
        Array("relevant_experience", "Relevant Experience"), _
        Array("assumptions_exclusions", "Assumptions & Exclusions") _
    )
End Function

Public Function SectionTitle(ByVal key As String) As String
    Dim arr As Variant: arr = SectionOrder()
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i)(0) = key Then SectionTitle = arr(i)(1): Exit For
    Next
End Function

' ---------- utilities ----------

Public Function DictOf(ParamArray pairs() As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim k As Variant, v As Variant
        k = pairs(i)(0): v = pairs(i)(1)
        d(k) = v
    Next
    Set DictOf = d
End Function

Private Function Nz(ByVal d As Object, ByVal key As String) As Variant
    If Not d Is Nothing Then If d.Exists(key) Then Nz = d(key)
End Function

Private Function NzDict(ByVal d As Object, ByVal key As String) As Object
    If Not d Is Nothing Then If d.Exists(key) Then Set NzDict = d(key)
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Public Function SimpleRewrite(ByVal s As String) As String
    Dim t As String: t = Replace(s, "  ", " ")
    SimpleRewrite = Trim$(t)
End Function

Public Function ConvertToJson(ByVal d As Object) As String
    ' Requires VBA-JSON (JsonConverter) referenced in the project
    ConvertToJson = JsonConverter.ConvertToJson(d, Whitespace:=2)
End Function

Public Function ReadAllText(ByVal path As String) As String
    Dim f As Integer: f = FreeFile
    Open path For Input As #f
    ReadAllText = Input$(LOF(f), f)
    Close #f
End Function

Public Sub WriteAllText(ByVal path As String, ByVal content As String)
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub

Public Function PickSavePath(ByVal title As String, ByVal defaultName As String) As String
    On Error Resume Next
    With Application.fileDialog(msoFileDialogSaveAs)
        .title = title
        .InitialFileName = defaultName
        If .Show = -1 Then PickSavePath = .selectedItems(1)
    End With
End Function
