Attribute VB_Name = "modValidation"
Option Explicit

' Returns Collection of Dictionary items: keys "level" ("BLOCKER"/"WARNING") and "message"
Public Function ValidateReview(ByVal st As Object) As Collection
    Dim issues As New Collection

    ' Blockers: required metadata
    Dim md As Object: Set md = st("metadata")
    If Len(NzStr(md("clientName"))) = 0 Then issues.Add MsgItem("BLOCKER", "Missing Client Name.")
    If Len(NzStr(md("projectName"))) = 0 Then issues.Add MsgItem("BLOCKER", "Missing Project Name.")

    ' Blocker: Financial total
    Dim fin As Object: Set fin = st("financial")
    If Len(NzStr(fin("total"))) = 0 Then issues.Add MsgItem("BLOCKER", "Financial total is empty. Recompute or enter a lump-sum total.")

    ' Blocker: Schedule sanity
    Dim sch As Object: Set sch = st("schedule")
    Dim phases As Object: Set phases = sch("phases")
    Dim totalWks As Double: totalWks = 0
    Dim k As Variant
    For Each k In phases.keys
        Dim w As Double: w = val(NzStr(phases(k)("duration")))
        If w <= 0 Then issues.Add MsgItem("BLOCKER", "Phase '" & k & "' has zero/negative duration.")
        totalWks = totalWks + w
    Next

    ' Blocker: Team roles
    Dim team As Object: Set team = st("team")
    If Len(NzStr(team("Project Principal")("name"))) = 0 Then issues.Add MsgItem("BLOCKER", "Team: Project Principal not assigned.")
    If Len(NzStr(team("Project Manager")("name"))) = 0 Then issues.Add MsgItem("BLOCKER", "Team: Project Manager not assigned.")

    ' Blocker: Assumptions baseline
    If Not st("sections").Exists("assumptions_exclusions") Then
        issues.Add MsgItem("BLOCKER", "Assumptions & Exclusions section missing.")
    End If

    ' Warnings: baseline durations check
    If totalWks = 11 Or totalWks = 8 Or totalWks = 12 Or totalWks = 24 Then
        ' typical baseline durations – ok
    ElseIf totalWks > 0 Then
        issues.Add MsgItem("WARNING", "Schedule total (" & totalWks & " weeks) differs from common baseline patterns (e.g., 11 or 8/12/24).")
    End If

    Set ValidateReview = issues
End Function

Private Function MsgItem(ByVal level As String, ByVal message As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("level") = level
    d("message") = message
    Set MsgItem = d
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function
