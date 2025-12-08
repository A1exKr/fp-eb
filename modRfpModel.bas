Attribute VB_Name = "modRfpModel"
' === modRfpModel ===
Option Explicit

Public Function NewReviewModel() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    ' Use SET when assigning objects into dictionary items
    Set d("metadata") = CreateObject("Scripting.Dictionary")
    Set d("sections") = CreateObject("Scripting.Dictionary")
    Set d("schedule") = CreateObject("Scripting.Dictionary")
    Set d("team") = CreateObject("Scripting.Dictionary")
    Set d("financial") = CreateObject("Scripting.Dictionary")
    Set d("experience") = CreateObject("System.Collections.ArrayList")

    Set NewReviewModel = d
End Function

Public Sub EnsureDefaultSections(ByRef model As Object)
    Dim s As Object: Set s = model("sections")

    AddSection s, "cover_letter", "Cover Letter"
    AddSection s, "project_understanding", "Project Understanding"
    AddSection s, "methodology", "Methodology"
    AddSection s, "scope_deliverables", "Scope & Deliverables"
    AddSection s, "schedule", "Schedule"
    AddSection s, "team", "Team Structure"
    AddSection s, "financial", "Financial Proposal"
    AddSection s, "relevant_experience", "Relevant Experience"
    AddSection s, "assumptions_exclusions", "Assumptions & Exclusions"
End Sub

Private Sub AddSection(ByRef sections As Object, ByVal key As String, ByVal title As String)
    If Not sections.Exists(key) Then
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        d("title") = title
        d("status") = 0
        d("contentSuggested") = ""
        d("contentEdited") = ""
        Set sections(key) = d   ' SET here too
    End If
End Sub

