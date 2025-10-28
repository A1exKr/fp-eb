Attribute VB_Name = "modSmokeTest"
Option Explicit

' Smoke test helpers to run the Review & Finalize screen without touching the parser.

Public Sub Test_OpenReviewWithInlineMock()
    Dim js As String: js = InlineMockJson()
    OpenReviewWithJson js
End Sub

Public Sub Test_OpenReviewFromFile()
    Dim p As String
    ' Adjust this path to wherever you save rfp_mock.json
    p = CurDir$ & "\rfp_mock.json"
    If Dir$(p, vbNormal) = "" Then
        MsgBox "Mock file not found: " & p, vbExclamation
        Exit Sub
    End If
    OpenReviewWithJson ReadAllText(p)
End Sub

Private Function ReadAllText(ByVal filePath As String) As String
    Dim f As Integer: f = FreeFile
    Open filePath For Binary As #f
    ReadAllText = Space$(LOF(f))
    Get #f, , ReadAllText
    Close #f
End Function

' Minimal but complete mock JSON to exercise all tabs.
Public Function InlineMockJson() As String
    Dim s As String
    s = s & "{""project"":{""name"":""Master plan, Vietnam"",""location"":""Vietnam (50 ha)"",""type"":""Master Plan"",""siteArea"":""50 ha""}," & _
            """client"":{""name"":""Confidential Client""}," & _
            """understanding"":{""understanding"":""Develop a vibrant, sustainable, TOD-oriented mixed-use master plan integrating residential, office and commercial uses with smart-city principles.""}," & _
            """methodology"":{""workshops"":4,""notes"":""Fast-track, parallel workstreams""}," & _
            """scope"":{""scopeList"":[" & _
                """Functional cluster zoning and land-use options""," & _
                """Infrastructure and constraints analysis""," & _
                """Phasing strategy and technical/economic comparison""]," & _
                """deliverablesList"":[" & _
                """Master plan report (PDF)""," & _
                """Plans/sections/elevations (A3/A1)""," & _
                """Rendered perspectives and area schedules""]}," & _
            """schedule"":{""totalWeeks"":11,""milestones"":[""Kick-off"",""WS2"",""WS3"",""WS4"",""Final""]}," & _
            """team"":{""principal"":{""name"":""Masakazu Kimura"",""title"":""Senior Executive Officer""}," & _
                      """pm"":{""name"":""Jan Henckens"",""title"":""Senior Project Manager""}}," & _
            """fee"":{""currency"":""USD"",""rates"":{""Principal"":250,""PM"":160,""Urban Planner"":120,""Architect"":140}," & _
                      """effortByPhase"":{""Stage 1"":{""Principal"":8,""PM"":36,""Urban Planner"":80,""Architect"":40}," & _
                                         """Stage 2"":{""Principal"":8,""PM"":40,""Urban Planner"":100,""Architect"":60}}," & _
                      """overheadPct"":0.1,""travel"":{""trips"":2,""peoplePerTrip"":2,""unitCost"":6000,""currency"":""USD""}," & _
                      """subconsultants"":[{""name"":""Local Planner"",""fee"":18000,""currency"":""USD""}]}," & _
            """experience"":[{""name"":""HCMC Central District Urban Design"",""location"":""Vietnam"",""scale"":""930 ha"",""client"":""HCMC"",""period"":""2007"",""summary"":""International competition winner; delivered detailed master plan.""}," & _
                            "{""name"":""Grand Front Osaka"",""location"":""Japan"",""scale"":""567,927 m2"",""client"":""Multiple"",""period"":""2013"",""summary"":""Mixed-use TOD landmark; offices, retail, hotel and residences.""}]," & _
            """assumptions"":{""defaultText"":""Fees are net of taxes; specialty consultants, approvals and physical models excluded; travel reimbursed at fixed unit rate; see terms.""}" & _
            "}"
    InlineMockJson = s
End Function

