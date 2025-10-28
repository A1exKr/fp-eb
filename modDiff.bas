Attribute VB_Name = "modDiff"
Option Explicit

' Very simple line-by-line diff: prefix "-" for old-only, "+" for new-only, " " for common.
Public Function diffText(ByVal oldText As String, ByVal newText As String) As String
    Dim oldArr() As String, newArr() As String
    oldArr = Split(NormalizeEOL(oldText), vbLf)
    newArr = Split(NormalizeEOL(newText), vbLf)

    Dim i As Long, j As Long
    Dim buf As String

    i = 0: j = 0
    Do While i <= UBound(oldArr) Or j <= UBound(newArr)
        If i > UBound(oldArr) Then
            buf = buf & "+ " & newArr(j) & vbCrLf
            j = j + 1
        ElseIf j > UBound(newArr) Then
            buf = buf & "- " & oldArr(i) & vbCrLf
            i = i + 1
        ElseIf oldArr(i) = newArr(j) Then
            buf = buf & "  " & oldArr(i) & vbCrLf
            i = i + 1: j = j + 1
        Else
            ' naive: show deletion then addition
            buf = buf & "- " & oldArr(i) & vbCrLf
            buf = buf & "+ " & newArr(j) & vbCrLf
            i = i + 1: j = j + 1
        End If
    Loop

    diffText = buf
End Function

Private Function NormalizeEOL(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    NormalizeEOL = s
End Function
