Attribute VB_Name = "modReviewFormFactory"
'modReviewFormFactory.bas
Option Explicit

' ===========
'  IMPORTANT
' ===========
' This module programmatically creates the UserForms (frmReviewFinalize and frmDiffViewer).
' No .frx files are required. Requirements:
'   1) Trust access to the VBA project object model (Excel/Word Options > Trust Center).
'   2) Reference: Microsoft Visual Basic for Applications Extensibility 5.3
'
' After importing:
'   - Run  InstallReviewFinalizeForm
'   - Then run ShowReviewFinalize (or call from your ribbon/menus)
'
' Fonts: All UI elements are set to "Artifact Element", 10pt as requested.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const UI_FONT_NAME As String = "Artifact Element"
Private Const UI_FONT_SIZE As Integer = 10

Public Sub InstallReviewFinalizeForm()
    Dim vbProj As VBIDE.VBProject
    Dim frm As VBIDE.vbComponent
    Dim code As VBIDE.CodeModule
    Dim d As Object

    '--- Always target this workbook ---
    Set vbProj = ThisWorkbook.VBProject

    '----------------------------------------------------------
    ' Check if form already exists Å® reuse instead of rebuilding
    '----------------------------------------------------------
    On Error Resume Next
    Set frm = vbProj.VBComponents("frmReviewFinalize")
    On Error GoTo 0

    If Not frm Is Nothing Then
        Debug.Print "frmReviewFinalize already installed; skipping rebuild."
        Exit Sub
    End If

    '----------------------------------------------------------
    ' Create new userform (first-time install only)
    '----------------------------------------------------------
    Set frm = vbProj.VBComponents.Add(vbext_ct_MSForm)
    DoEvents

    On Error Resume Next
    frm.name = "frmReviewFinalize"
    If Err.Number <> 0 Then
        Debug.Print "Rename failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '--- Get designer reference ---
    Set d = frm.Designer
    ApplyFont d


    '==========================================================
    '  MAIN FORM LAYOUT (unchanged)
    '==========================================================

    'Title label
    Dim lblTitle As Object
    Set lblTitle = d.Controls.Add("Forms.Label.1", "lblTitle", True)
    With lblTitle
        .Caption = "FP-GEN: Review & Finalize"
        .Left = 12: .Top = 8: .Width = 300: .Height = 18
        .Font.Bold = True
        ApplyFont lblTitle
    End With

    'Top buttons ------------------------------------------------
    Dim cmdSave As Object, cmdValidate As Object, cmdGen As Object
    Set cmdSave = d.Controls.Add("Forms.CommandButton.1", "cmdSaveDraft", True)
    With cmdSave: .Caption = "Save Draft": .Left = 480: .Top = 4: .Width = 84: .Height = 22: ApplyFont cmdSave: End With

    Set cmdValidate = d.Controls.Add("Forms.CommandButton.1", "cmdValidate", True)
    With cmdValidate: .Caption = "Validate": .Left = 570: .Top = 4: .Width = 78: .Height = 22: ApplyFont cmdValidate: End With

    Set cmdGen = d.Controls.Add("Forms.CommandButton.1", "cmdGenerateFP", True)
    With cmdGen: .Caption = "Generate FP": .Left = 656: .Top = 4: .Width = 96: .Height = 22: ApplyFont cmdGen: End With

    'Left section list ------------------------------------------
    Dim lst As Object
    Set lst = d.Controls.Add("Forms.ListBox.1", "ListSections", True)
    With lst
        .Left = 8: .Top = 32: .Width = 180: .Height = 360
        .IntegralHeight = False
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0 pt;172 pt"
        ApplyFont lst
    End With

    'Main multipage --------------------------------------------
    Dim mp As Object
    Set mp = d.Controls.Add("Forms.MultiPage.1", "mpMain", True)
    With mp
        .Left = 196: .Top = 28: .Width = 556: .Height = 364
        .Style = 1  ' fmTabStyleButtons
        ApplyFont mp
    End With
    If mp.Pages.Count > 1 Then mp.Pages.Remove 1

    '=== Editor page ===
    Dim txt As Object, chkNums As Object, cmdRegen As Object, cmdDiff As Object, cmdApprove As Object, cmdLock As Object
    With mp.Pages(0)
        .Caption = "Editor": ApplyFont mp.Pages(0)

        Set txt = .Controls.Add("Forms.TextBox.1", "txtText", True)
        With txt: .Left = 8: .Top = 8: .Width = 536: .Height = 272: .MultiLine = True: .ScrollBars = 3: ApplyFont txt: End With

        Set chkNums = .Controls.Add("Forms.CheckBox.1", "chkProtectNumbers", True)
        With chkNums: .Caption = "Protect numbers & dates when regenerating": .Left = 8: .Top = 284: .Width = 300: .Height = 16: .value = True: ApplyFont chkNums: End With

        Set cmdRegen = .Controls.Add("Forms.CommandButton.1", "cmdRegenerate", True)
        With cmdRegen: .Caption = "AI Re-generate": .Left = 272: .Top = 280: .Width = 92: .Height = 22: ApplyFont cmdRegen: End With

        Set cmdDiff = .Controls.Add("Forms.CommandButton.1", "cmdDiff", True)
        With cmdDiff: .Caption = "Show Diff": .Left = 368: .Top = 280: .Width = 76: .Height = 22: ApplyFont cmdDiff: End With

        Set cmdApprove = .Controls.Add("Forms.CommandButton.1", "cmdApprove", True)
        With cmdApprove: .Caption = "Approve": .Left = 448: .Top = 280: .Width = 48: .Height = 22: ApplyFont cmdApprove: End With

        Set cmdLock = .Controls.Add("Forms.CommandButton.1", "cmdLock", True)
        With cmdLock: .Caption = "Lock": .Left = 500: .Top = 280: .Width = 44: .Height = 22: ApplyFont cmdLock: End With
    End With

    '=== Financial page ===
    Dim pgFee As Object
    Set pgFee = mp.Pages.Add: pgFee.Caption = "Financial": ApplyFont pgFee
    Dim lstFee As Object, cmdRecompute As Object
    Set lstFee = pgFee.Controls.Add("Forms.ListBox.1", "lstFee", True)
    With lstFee: .Left = 8: .Top = 8: .Width = 536: .Height = 272: .IntegralHeight = False: .ColumnCount = 2: .ColumnWidths = "250 pt;200 pt": ApplyFont lstFee: End With
    Set cmdRecompute = pgFee.Controls.Add("Forms.CommandButton.1", "cmdRecomputeFee", True)
    With cmdRecompute: .Caption = "Recompute": .Left = 8: .Top = 284: .Width = 80: .Height = 22: ApplyFont cmdRecompute: End With

    '=== Schedule page ===
    Dim pgSch As Object
    Set pgSch = mp.Pages.Add: pgSch.Caption = "Schedule": ApplyFont pgSch
    Dim lstSch As Object
    Set lstSch = pgSch.Controls.Add("Forms.ListBox.1", "lstSchedule", True)
    With lstSch: .Left = 8: .Top = 8: .Width = 536: .Height = 272: .IntegralHeight = False: .ColumnCount = 3: .ColumnWidths = "150 pt;100 pt;250 pt": ApplyFont lstSch: End With

    '=== Team page ===
    Dim pgTeam As Object
    Set pgTeam = mp.Pages.Add: pgTeam.Caption = "Team": ApplyFont pgTeam
    Dim lstTeam As Object, cmdAttachCV As Object
    Set lstTeam = pgTeam.Controls.Add("Forms.ListBox.1", "lstTeam", True)
    With lstTeam: .Left = 8: .Top = 8: .Width = 536: .Height = 272: .IntegralHeight = False: .ColumnCount = 2: .ColumnWidths = "180 pt;300 pt": ApplyFont lstTeam: End With
    Set cmdAttachCV = pgTeam.Controls.Add("Forms.CommandButton.1", "cmdAttachCV", True)
    With cmdAttachCV: .Caption = "Attach CV": .Left = 8: .Top = 284: .Width = 80: .Height = 22: ApplyFont cmdAttachCV: End With

    '=== Validation page ===
    Dim pgVal As Object
    Set pgVal = mp.Pages.Add: pgVal.Caption = "Validation": ApplyFont pgVal
    Dim lstVal As Object
    Set lstVal = pgVal.Controls.Add("Forms.ListBox.1", "lstValidation", True)
    With lstVal: .Left = 8: .Top = 8: .Width = 536: .Height = 300: .IntegralHeight = False: .ColumnCount = 2: .ColumnWidths = "100 pt;380 pt": ApplyFont lstVal: End With

    '==========================================================
    '  CODE-BEHIND INJECTION
    '==========================================================
    Set code = frm.CodeModule
    Dim lineNr As Long: lineNr = 1

    lineNr = InsertLine(code, lineNr, "Option Explicit")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub UserForm_Initialize()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_Initialize Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub ListSections_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_SectionSelect Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdSaveDraft_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_SaveDraft Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdValidate_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_Validate Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdGenerateFP_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_GenerateFP Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdRegenerate_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_Regenerate Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdDiff_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_ShowDiff Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdApprove_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_Approve Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdLock_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_Lock Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdRecomputeFee_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_RecomputeFee Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdAttachCV_Click()")
    lineNr = InsertLine(code, lineNr, "    modReviewFinalize.Form_AttachCV Me")
    lineNr = InsertLine(code, lineNr, "End Sub")

    '--- also install diff viewer form if not present ---
    InstallDiffViewerForm

    MsgBox "Forms installed: frmReviewFinalize, frmDiffViewer" & vbCrLf & _
           "Open via: ShowReviewFinalize", vbInformation
End Sub

Private Function InsertLine(cm As VBIDE.CodeModule, ByVal atLine As Long, ByVal textLine As String) As Long
    cm.InsertLines atLine, textLine
    InsertLine = atLine + 1
End Function

Public Sub InstallDiffViewerForm()
    Dim vbProj As VBIDE.VBProject
    Dim frm As VBIDE.vbComponent
    Dim code As VBIDE.CodeModule

    Set vbProj = Application.vbe.ActiveVBProject

    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("frmDiffViewer")
    On Error GoTo 0

    Set frm = vbProj.VBComponents.Add(vbext_ct_MSForm)
    frm.name = "frmDiffViewer"

    Dim d As Object: Set d = frm.Designer
    ApplyFont d

    Dim txt As Object, btnClose As Object
    Set txt = d.Controls.Add("Forms.TextBox.1", "txtDiff", True)
    With txt
        .Left = 8: .Top = 8: .Width = 520: .Height = 292
        .MultiLine = True
        .ScrollBars = 3
        .WordWrap = False
        ApplyFont txt
    End With

    Set btnClose = d.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With btnClose
        .Caption = "Close"
        .Left = 448: .Top = 304: .Width = 80: .Height = 22
        ApplyFont btnClose
    End With

    Set code = frm.CodeModule
    Dim lineNr As Long: lineNr = 1
    lineNr = InsertLine(code, lineNr, "Option Explicit")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Public Sub ShowDiff(ByVal diffText As String)")
    lineNr = InsertLine(code, lineNr, "    Me.txtDiff.Text = diffText")
    lineNr = InsertLine(code, lineNr, "    Me.Show")
    lineNr = InsertLine(code, lineNr, "End Sub")
    lineNr = InsertLine(code, lineNr, "")
    lineNr = InsertLine(code, lineNr, "Private Sub cmdClose_Click()")
    lineNr = InsertLine(code, lineNr, "    Unload Me")
    lineNr = InsertLine(code, lineNr, "End Sub")
End Sub

Public Sub ShowReviewFinalize()
    Dim frmReview As Object
    Set frmReview = VBA.UserForms.Add("frmReviewFinalize")
    frmReview.Show vbModeless
End Sub

' Apply Artifact Element, 10pt to any control/container that exposes Font
Private Sub ApplyFont(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.name = UI_FONT_NAME
    ctl.Font.Size = UI_FONT_SIZE
    On Error GoTo 0
End Sub
