Attribute VB_Name = "modRfpMapper"
'Attribute VB_Name = "modRfpMapper"
Option Explicit

' Maps the raw AI JSON (string) into the canonical Review Model (Dictionary).
' Relies on JsonConverter.ParseJson (VBA-JSON) and local helpers (JPath).

Public Function MapRfpJsonToReviewModel(ByVal rfpJson As String) As Object
    On Error GoTo MapError
    
    Dim review As Object: Set review = NewReviewModel()
    EnsureDefaultSections review

    Dim ai As Object
    Debug.Print "MapRfpJsonToReviewModel 1"
    
    Set ai = JsonConverter.ParseJSON(rfpJson)
    Debug.Print "Root keys: "; Join(ai.keys, ", ")

    ' ---- metadata ----
    Dim md As Object: Set md = review("metadata")
    md("projectName") = GetS(ai, "project.name")
    md("clientName") = GetS(ai, "client.name")
    md("location") = GetS(ai, "project.location")
    md("projectType") = GetS(ai, "project.type")
    md("siteArea") = GetS(ai, "project.siteArea")

    ' ---- schedule ----
    Dim sch As Object: Set sch = review("schedule")
    Dim totalW As Long: totalW = CLng(val(Def(GetS(ai, "schedule.totalWeeks"), "0")))
    If totalW = 0 Then totalW = 44
    sch("totalWeeks") = totalW

    Dim ph As Object: Set ph = CreateObject("Scripting.Dictionary")
    AddPhase ph, "Pre-Concept", 8, ""
    AddPhase ph, "Concept Design", 12, ""
    AddPhase ph, "Schematic Design", 24, ""
    If totalW = 11 Then
        ph.RemoveAll
        AddPhase ph, "Stage 1", 4, "WS1, WS2"
        AddPhase ph, "Stage 2", 7, "WS3, Final"
    End If
    Set sch("phases") = ph
    Set sch("milestones") = GetArrayOrDefault(ai, "schedule.milestones")

    ' ---- team ----
    Dim team As Object: Set team = review("team")
    AddTeam team, "Project Principal", GetS(ai, "team.principal.name"), GetS(ai, "team.principal.title"), ""
    AddTeam team, "Project Manager", GetS(ai, "team.pm.name"), GetS(ai, "team.pm.title"), ""

    ' ---- financial ----
    Dim fin As Object: Set fin = review("financial")
    fin("currency") = Def(GetS(ai, "fee.currency"), "USD")
    Set fin("rates") = GetDictOrDefault(ai, "fee.rates")
    Set fin("effort") = GetDictOrDefault(ai, "fee.effortByPhase")
    fin("overheadPct") = CDbl(val(Def(GetS(ai, "fee.overheadPct"), "0.1")))
    Set fin("travel") = GetDictOrDefault(ai, "fee.travel")
    Set fin("subconsultants") = GetArrayOrDefault(ai, "fee.subconsultants")
    
    On Error Resume Next
    modFeeEngine.ComputeFee fin
    On Error GoTo MapError

    ' ---- experience ----
    Dim exp As Object: Set exp = review("experience")
    CopyArrayObjects GetArrayOrDefault(ai, "experience"), exp

    ' ---- sections ----
    Debug.Print "MapRfpJsonToReviewModel 2"
    
    Dim secs As Object: Set secs = review("sections")
    
    ' Ensure all section dictionaries exist
    Dim key As Variant
    For Each key In Array("cover_letter", "project_understanding", "methodology", _
                          "scope_deliverables", "schedule", "team", "financial", _
                          "relevant_experience", "assumptions_exclusions")
        If Not secs.Exists(key) Then
            Dim newSec As Object: Set newSec = CreateObject("Scripting.Dictionary")
            Set secs(key) = newSec
        End If
        If Not secs(key).Exists("contentSuggested") Then
            secs(key)("contentSuggested") = ""
        End If
    Next key

    Debug.Print "MapRfpJsonToReviewModel 3"
    Debug.Print "Methodology Type: "; TypeName(ai("methodology"))
    
    ' Generate section content with error protection for each
    ' Generate section content with error protection for each
    On Error Resume Next
    
    secs("cover_letter")("contentSuggested") = modSectionSynthesis.Draft_CoverLetter(md)
    If Err.Number <> 0 Then Debug.Print "Error in Draft_CoverLetter: " & Err.Description: Err.Clear
    
    secs("project_understanding")("contentSuggested") = modSectionSynthesis.Draft_ProjectUnderstanding(md, SafeGetDict(ai, "understanding"))
    If Err.Number <> 0 Then Debug.Print "Error in Draft_ProjectUnderstanding: " & Err.Description: Err.Clear
    
    secs("methodology")("contentSuggested") = modSectionSynthesis.Draft_Methodology(GetS(ai, "methodology.text"))
    If Err.Number <> 0 Then Debug.Print "Error in Draft_Methodology: " & Err.Description: Err.Clear
    
    secs("scope_deliverables")("contentSuggested") = modSectionSynthesis.Draft_ScopeAndDeliverables(SafeGetDict(ai, "scope"))
    If Err.Number <> 0 Then Debug.Print "Error in Draft_ScopeAndDeliverables: " & Err.Description: Err.Clear
    
    secs("schedule")("contentSuggested") = modSectionSynthesis.Draft_Schedule(sch)
    If Err.Number <> 0 Then Debug.Print "Error in Draft_Schedule: " & Err.Description: Err.Clear
    
    secs("team")("contentSuggested") = modSectionSynthesis.Draft_Team(team)
    If Err.Number <> 0 Then Debug.Print "Error in Draft_Team: " & Err.Description: Err.Clear
    
    secs("financial")("contentSuggested") = modSectionSynthesis.Draft_Financial(fin)
    If Err.Number <> 0 Then Debug.Print "Error in Draft_Financial: " & Err.Description: Err.Clear
    
    secs("relevant_experience")("contentSuggested") = modSectionSynthesis.Draft_Experience(exp)
    If Err.Number <> 0 Then Debug.Print "Error in Draft_Experience: " & Err.Description: Err.Clear
    
    secs("assumptions_exclusions")("contentSuggested") = modSectionSynthesis.Draft_AssumptionsExclusions(GetS(ai, "assumptions.defaultText"))
    If Err.Number <> 0 Then Debug.Print "Error in Draft_AssumptionsExclusions: " & Err.Description: Err.Clear
    
    On Error GoTo 0

    Debug.Print "MapRfpJsonToReviewModel COMPLETE"
    Set MapRfpJsonToReviewModel = review
    Exit Function

MapError:
    Debug.Print "ERROR mapping JSON at line " & Erl & ": " & Err.Description
    MsgBox "Failed to map RFP data: " & Err.Description, vbCritical
    
    ' Return partial review even if there's an error
    If review Is Nothing Then Set review = NewReviewModel()
    Set MapRfpJsonToReviewModel = review
End Function

' Safe helper that never fails
Private Function SafeGetDict(parent As Object, path As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Dim v As Variant
    Set v = JPath(parent, path)
    
    If Err.Number = 0 And Not v Is Nothing Then
        If TypeName(v) = "Dictionary" Then
            Set result = v
        End If
    End If
    On Error GoTo 0
    
    Set SafeGetDict = result
End Function

Private Sub EnsureSectionExists(ByVal secs As Object, ByVal key As String, ByVal title As String)
    Dim sec As Object
    On Error Resume Next
    Set sec = secs(key)
    On Error GoTo 0
    If sec Is Nothing Then
        Set sec = CreateObject("Scripting.Dictionary")
        sec("title") = title
        sec("status") = 0
        sec("contentSuggested") = ""
        sec("contentEdited") = ""
        secs(key) = sec
    End If
End Sub


' ---------- helpers ----------

Private Sub AddPhase(ByRef ph As Object, ByVal name As String, ByVal weeks As Long, ByVal notes As String)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("duration") = weeks
    d("notes") = notes
    Set ph(name) = d
End Sub

Private Sub AddTeam(ByRef team As Object, ByVal role As String, ByVal name As String, ByVal title As String, ByVal cvPath As String)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("name") = name
    d("title") = title
    d("cvPath") = cvPath
    Set team(role) = d
End Sub

Private Function GetS(ai As Object, path As String) As String
    On Error Resume Next
    Dim v As Variant: v = JPath(ai, path)
    On Error GoTo 0
    If IsObject(v) Then
        GetS = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        GetS = ""
    Else
        GetS = CStr(v)
    End If
End Function

Private Function Def(ByVal v As String, ByVal fallback As String) As String
    If Len(Trim$(v)) = 0 Then Def = fallback Else Def = v
End Function

Private Function GetOrEmpty(ai As Object, path As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim v As Variant: v = JPath(ai, path)
    On Error GoTo 0
    If IsObject(v) Then
        Set GetOrEmpty = v
    Else
        Set GetOrEmpty = d
    End If
End Function



Private Function GetDictOrDefault(ai As Object, path As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim v As Variant: v = JPath(ai, path)
    On Error GoTo 0
    If IsObject(v) Then
        If TypeName(v) = "Dictionary" Then
            Set GetDictOrDefault = v
            Exit Function
        End If
    End If
    Set GetDictOrDefault = d
End Function

Private Function GetArrayOrDefault(ai As Object, path As String) As Object
    Dim arr As Object: Set arr = CreateObject("System.Collections.ArrayList")
    On Error Resume Next
    Dim v As Variant: v = JPath(ai, path)
    On Error GoTo 0

    If IsObject(v) Then
        If TypeName(v) = "Collection" Then
            Dim i As Long
            For i = 1 To v.Count
                arr.Add v.item(i)
            Next i
        End If
    End If
    Set GetArrayOrDefault = arr
End Function

Public Sub CopyArrayObjects(src As Object, ByRef dst As Object)
    On Error Resume Next
    If src Is Nothing Then Exit Sub
    Dim i As Long
    For i = 0 To src.Count - 1
        dst.Add src(i)
    Next i
    On Error GoTo 0
End Sub

' Basic dotted-path navigation into the Dictionary/Collection tree produced by VBA-JSON.
Public Function JPath(root As Variant, path As String) As Variant
    If Not IsObject(root) Then
        JPath = Empty
        Exit Function
    End If

    Dim parts() As String: parts = Split(path, ".")
    Dim cur As Variant: Set cur = root
    Dim i As Long, key As String

    On Error GoTo JPathFail
    For i = LBound(parts) To UBound(parts)
        key = parts(i)
        If TypeName(cur) = "Dictionary" Then
            If cur.Exists(key) Then
                If IsObject(cur(key)) Then
                    Set cur = cur(key)
                Else
                    ' scalar terminal
                    JPath = cur(key)
                    Exit Function
                End If
            Else
                JPath = Empty
                Exit Function
            End If
        Else
            ' not a dictionary; we do not traverse collections by index here
            JPath = Empty
            Exit Function
        End If
    Next i

    If IsObject(cur) Then
        Set JPath = cur
    Else
        JPath = cur
    End If
    Exit Function

JPathFail:
    JPath = Empty
End Function

' Additional safe helpers
Private Function GetOrEmptyDict(parent As Object, key As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    If Not parent Is Nothing Then
        If parent.Exists(key) Then
            Dim v As Variant
            Set v = parent(key)
            If Not v Is Nothing Then
                If TypeName(v) = "Dictionary" Then
                    Set d = v
                End If
            End If
        End If
    End If
    On Error GoTo 0
    
    Set GetOrEmptyDict = d
End Function

