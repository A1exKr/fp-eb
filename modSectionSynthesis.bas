Attribute VB_Name = "modSectionSynthesis"
'Attribute VB_Name = "modSectionSynthesis"
Option Explicit
' NOTE: ASCII-only. No Unicode bullets/dashes embedded in code.

'========================
' PUBLIC DRAFT FUNCTIONS
'========================
Public Function GetOrText(ByVal parent As Object, ByVal key As String) As String
    On Error Resume Next
    Dim v As Variant
    Dim result As String
    
    result = ""
    
    If parent Is Nothing Then
        GetOrText = ""
        Exit Function
    End If
    
    ' Check if parent has the key
    If Not parent.Exists(key) Then
        GetOrText = ""
        Exit Function
    End If
    
    ' Get the value
    v = parent(key)
    
    ' Handle different types
    If IsObject(v) Then
        ' It's a nested object (Dictionary)
        If v.Exists("defaultText") Then
            result = CStr(v("defaultText"))
        ElseIf v.Exists("text") Then
            result = CStr(v("text"))
        End If
    ElseIf VarType(v) = vbString Then
        ' It's already a string
        result = v
    Else
        ' Try to convert to string
        result = CStr(v)
    End If
    
    GetOrText = Trim(result)
End Function




Public Function Draft_CoverLetter(md As Object) As String
    Draft_CoverLetter = _
        "Dear Sirs," & vbCrLf & vbCrLf & _
        "We are honored to submit our proposal for " & Safe(md, "projectName") & _
        " for " & Safe(md, "clientName") & ". Based on our understanding of the site in " & _
        Safe(md, "location") & ", we have assembled a multidisciplinary team to deliver a " & _
        LCase$(Safe(md, "projectType")) & " of high quality with senior oversight." & vbCrLf & vbCrLf & _
        "We look forward to working together." & vbCrLf & _
        "Nikken Sekkei Ltd."
End Function

Public Function Draft_ProjectUnderstanding(md As Object, ctx As Object) As String
    Draft_ProjectUnderstanding = Safe(ctx, "understanding") & vbCrLf & _
        "Site Area: " & Safe(md, "siteArea") & "  |  Location: " & Safe(md, "location")
End Function

Public Function Draft_Methodology(ctx As Variant) As String
    Dim t As String
    
    On Error Resume Next
    
    ' Handle Dictionary object
    If IsObject(ctx) Then
        If ctx.Exists("text") Then
            t = CStr(ctx("text"))
        ElseIf ctx.Exists("defaultText") Then
            t = CStr(ctx("defaultText"))
        End If
    ' Handle string directly
    ElseIf VarType(ctx) = vbString Then
        t = CStr(ctx)
    End If
    
    On Error GoTo 0

    If Len(Trim(t)) > 0 Then
        Draft_Methodology = t
    Else
        ' Default fallback
        Draft_Methodology = _
            "Approach:" & vbCrLf & _
            "- Kick-off and site analysis" & vbCrLf & _
            "- Iterative design workshops" & vbCrLf & _
            "- Concept options and evaluation" & vbCrLf & _
            "- Refinement with client feedback" & vbCrLf & _
            "- Finalization and QA/QC"
    End If
End Function


Public Function Draft_ScopeAndDeliverables(ctx As Object) As String
    Draft_ScopeAndDeliverables = _
        "Scope:" & vbCrLf & JoinList(GetOrEmptyList(ctx, "scopeList")) & vbCrLf & _
        "Deliverables:" & vbCrLf & JoinList(GetOrEmptyList(ctx, "deliverablesList"))
End Function

Public Function Draft_Schedule(sch As Object) As String
    Draft_Schedule = "Total duration: " & NzS(sch("totalWeeks")) & " weeks" & vbCrLf & _
                     ListPhases(GetOrEmptyDict(sch, "phases"))
End Function

Public Function Draft_Team(team As Object) As String
    Draft_Team = "Key Roles:" & vbCrLf & ListTeam(team)
End Function

Public Function Draft_Financial(fin As Object) As String
    Dim t As Object: On Error Resume Next: Set t = fin("totals"): On Error GoTo 0
    If t Is Nothing Then
        Draft_Financial = "Fee totals are not available."
        Exit Function
    End If
    Draft_Financial = _
        "Lump Sum (" & Safe(fin, "currency") & "):" & vbCrLf & _
        "- Labor Subtotal: " & FormatNumber(val(NzS(t("laborSubtotal")))) & vbCrLf & _
        "- Overheads and Profit: " & FormatNumber(val(NzS(t("overheads")))) & vbCrLf & _
        "- Subconsultants: " & FormatNumber(val(NzS(t("subconsultants")))) & vbCrLf & _
        "- Reimbursables: " & FormatNumber(val(NzS(t("reimbursables")))) & vbCrLf & _
        "Total: " & FormatNumber(val(NzS(t("total")))) & vbCrLf & vbCrLf & _
        "Notes: Fees are net of taxes; travel reimbursed per trip; specialty consultants and physical models excluded; " & _
        "see Assumptions and Exclusions."
End Function

Public Function Draft_Experience(expArr As Object) As String
    Draft_Experience = "Selected References:" & vbCrLf & ListExperience(expArr)
End Function

Public Function Draft_AssumptionsExclusions(ctx As Variant) As String
    Dim t As String
    On Error Resume Next
    If IsObject(ctx) Then
        t = Safe(ctx, "defaultText")
    Else
        t = CStr(ctx)
    End If
    On Error GoTo 0

    If Len(Trim(t)) > 0 Then
        Draft_AssumptionsExclusions = t
    Else
        Draft_AssumptionsExclusions = _
            "Fees are net of taxes; specialty consultants, approvals, and models excluded." & vbCrLf & _
            "Travel reimbursed at fixed unit rate; see terms and conditions."
    End If
End Function


'========================
' INTERNAL HELPERS
'========================
Private Function Safe(d As Variant, key As String) As String
    On Error Resume Next
    Safe = ""
    
    ' Check if d is an object (Dictionary)
    If IsObject(d) Then
        If d.Exists(key) Then
            Safe = CStr(d(key))
        End If
    End If
End Function

Private Function NzS(v As Variant) As String
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Then
        NzS = ""
    Else
        NzS = CStr(v)
    End If
End Function

Private Function GetOrEmptyList(ctx As Object, key As String) As Object
    On Error Resume Next
    Dim v As Variant: v = ctx(key)
    If IsObject(v) Then
        Set GetOrEmptyList = v
    Else
        Dim a As Object: Set a = CreateObject("System.Collections.ArrayList")
        Set GetOrEmptyList = a
    End If
End Function

Private Function GetOrEmptyDict(ai As Object, path As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim v As Variant
    Set v = JPath(ai, path)
    On Error GoTo 0
    
    If Not v Is Nothing Then
        If TypeName(v) = "Dictionary" Then
            Set GetOrEmptyDict = v
            Exit Function
        End If
    End If
    Set GetOrEmptyDict = d
End Function

'========================
' LIST RENDERERS (ASCII)
'========================

Private Function JoinList(v As Variant) As String
    Dim i As Long, s As String
    On Error Resume Next
    If Not v Is Nothing Then
        For i = 0 To v.Count - 1
            s = s & "- " & CStr(v(i)) & vbCrLf
        Next i
    End If
    JoinList = s
End Function

Private Function ListPhases(ph As Object) As String
    Dim k As Variant, s As String
    If ph Is Nothing Then Exit Function
    For Each k In ph.keys
        Dim pd As Object: Set pd = ph(k)
        s = s & "- " & k & ": " & NzS(pd("duration")) & " weeks"
        If Len(NzS(pd("notes"))) > 0 Then s = s & " (" & CStr(pd("notes")) & ")"
        s = s & vbCrLf
    Next
    ListPhases = s
End Function

Private Function ListTeam(t As Object) As String
    Dim k As Variant, s As String
    If t Is Nothing Then Exit Function
    For Each k In t.keys
        s = s & "- " & k & ": " & Safe(t(k), "name") & " - " & Safe(t(k), "title") & vbCrLf
    Next
    ListTeam = s
End Function

Private Function ListExperience(a As Object) As String
    Dim i As Long, s As String
    If a Is Nothing Then Exit Function
    For i = 0 To a.Count - 1
        s = s & "- " & a(i)("name") & " (" & a(i)("location") & ") - " & a(i)("summary") & vbCrLf
    Next
    ListExperience = s
End Function


