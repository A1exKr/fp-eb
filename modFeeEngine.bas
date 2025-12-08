Attribute VB_Name = "modFeeEngine"
Option Explicit

' Computes financial totals from rates/effort/overhead/travel/subconsultants.
' Populates fin("totals") with keys: laborSubtotal, overheads, subconsultants, reimbursables, total

Public Sub ComputeFee(ByRef fin As Object)
    Dim labor As Double: labor = SumLabor(fin)
    Dim ohPct As Double: ohPct = 0
    On Error Resume Next
    ohPct = CDbl(fin("overheadPct"))
    On Error GoTo 0

    Dim oh As Double: oh = Round(labor * ohPct, 2)
    Dim subs As Double: subs = SumSubconsultants(fin)
    Dim travel As Double: travel = TravelTotal(fin)

    Dim tot As Double: tot = labor + oh + subs + travel

    Dim t As Object: Set t = CreateObject("Scripting.Dictionary")
    t("laborSubtotal") = labor
    t("overheads") = oh
    t("subconsultants") = subs
    t("reimbursables") = travel
    t("total") = tot
    Set fin("totals") = t
End Sub

Private Function SumLabor(ByVal fin As Object) As Double
    Dim rates As Object: Set rates = fin("rates")
    Dim effort As Object: Set effort = fin("effort")
    Dim p As Variant, r As Variant, sumVal As Double

    If rates Is Nothing Or effort Is Nothing Then
        SumLabor = 0
        Exit Function
    End If

    For Each p In effort.keys
        Dim roleHours As Object: Set roleHours = effort(p)
        For Each r In roleHours.keys
            sumVal = sumVal + CDbl(roleHours(r)) * CDbl(rates(r))
        Next r
    Next p
    SumLabor = sumVal
End Function

Private Function SumSubconsultants(ByVal fin As Object) As Double
    On Error Resume Next
    Dim a As Object: Set a = fin("subconsultants")
    Dim i As Long, sumVal As Double
    If Not a Is Nothing Then
        For i = 0 To a.Count - 1
            sumVal = sumVal + CDbl(a(i)("fee"))
        Next i
    End If
    SumSubconsultants = sumVal
End Function

Private Function TravelTotal(ByVal fin As Object) As Double
    On Error Resume Next
    Dim tr As Object: Set tr = fin("travel")
    If tr Is Nothing Then
        TravelTotal = 0
    Else
        TravelTotal = CDbl(Zn(tr("trips"))) * CDbl(Zn(tr("peoplePerTrip"))) * CDbl(Zn(tr("unitCost")))
    End If
End Function

Private Function Zn(v As Variant) As Double
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Or v = "" Then
        Zn = 0
    Else
        Zn = CDbl(v)
    End If
End Function
