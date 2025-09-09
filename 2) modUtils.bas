Attribute VB_Name = "modUtils"
Option Explicit

Public Function ReadAllText(ByVal path As String) As String
    Dim f As Integer: f = FreeFile
    On Error GoTo EH
    Open path For Binary As #f
    ReadAllText = Space$(LOF(f))
    Get #f, , ReadAllText
    Close #f
    Exit Function
EH:
    On Error Resume Next
    If f <> 0 Then Close #f
    ReadAllText = ""
End Function

Public Function SplitLines(ByVal s As String) As Variant
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    SplitLines = Split(s, vbLf)
End Function

Public Function FirstNumberIn(ByVal s As String) As Double
    Dim i As Long, c As String, buf As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If (c Like "[0-9]") Or c = "." Or c = "-" Or c = "+" Then
            buf = buf & c
        ElseIf Len(buf) > 0 Then
            Exit For
        End If
    Next i
    If buf = "" Then
        FirstNumberIn = 0#
    Else
        On Error Resume Next
        FirstNumberIn = CDbl(buf)
        On Error GoTo 0
    End If
End Function

Public Function FindValueNearKey(ByVal lines As Variant, ByVal key As String, Optional ByVal defaultVal As Double = 0#) As Double
    Dim i As Long, L As String
    For i = LBound(lines) To UBound(lines)
        L = lines(i)
        If InStr(1, L, key, vbTextCompare) > 0 Then
            FindValueNearKey = FirstNumberIn(L)
            Exit Function
        End If
    Next i
    FindValueNearKey = defaultVal
End Function

Public Function FindTextNearKey(ByVal lines As Variant, ByVal key As String, Optional ByVal defaultText As String = "") As String
    Dim i As Long, L As String
    For i = LBound(lines) To UBound(lines)
        L = lines(i)
        If InStr(1, L, key, vbTextCompare) > 0 Then
            Dim p As Long: p = InStr(1, L, ":", vbTextCompare)
            If p > 0 Then
                FindTextNearKey = Trim$(Mid$(L, p + 1))
                Exit Function
            Else
                FindTextNearKey = Trim$(L)
                Exit Function
            End If
        End If
    Next i
    FindTextNearKey = defaultText
End Function

Public Function NormalizeBasePolymer(ByVal txt As String) As String
    Dim t As String: t = UCase$(txt)
    If InStr(t, "PA46") > 0 Or InStr(t, "POLYAMIDE 46") > 0 Or InStr(t, "NYLON 46") > 0 Then
        NormalizeBasePolymer = "PA46": Exit Function
    End If
    If InStr(t, "PA66") > 0 Then NormalizeBasePolymer = "PA66": Exit Function
    If InStr(t, "PA6") > 0 Then NormalizeBasePolymer = "PA6": Exit Function
    If InStr(t, "PPS") > 0 Then NormalizeBasePolymer = "PPS": Exit Function
    If InStr(t, "PBT") > 0 Then NormalizeBasePolymer = "PBT": Exit Function
    If InStr(t, "PC+PBT") > 0 Then NormalizeBasePolymer = "PC+PBT": Exit Function
    If InStr(t, "PC+PET") > 0 Then NormalizeBasePolymer = "PC+PET": Exit Function
    If InStr(t, "PC") > 0 Then NormalizeBasePolymer = "PC": Exit Function
    If InStr(t, "ABS") > 0 Then NormalizeBasePolymer = "ABS": Exit Function
    NormalizeBasePolymer = Trim$(txt)
End Function

Public Function NormalizeFillerType(ByVal txt As String) As String
    Dim t As String: t = UCase$(txt)
    If InStr(t, "GLASS") > 0 Or InStr(t, "GF") > 0 Then NormalizeFillerType = "GF": Exit Function
    If InStr(t, "CARBON") > 0 Or InStr(t, "CF") > 0 Then NormalizeFillerType = "CF": Exit Function
    If InStr(t, "MINERAL") > 0 Then NormalizeFillerType = "Mineral": Exit Function
    If InStr(t, "POTASSIUM TITANATE") > 0 Or InStr(t, "TITANATE") > 0 Or InStr(t, "TISMO") > 0 Then NormalizeFillerType = "Mineral": Exit Function
    If InStr(t, "NONE") > 0 Or InStr(t, "NO FILLER") > 0 Then NormalizeFillerType = "None": Exit Function
    NormalizeFillerType = Trim$(txt)
End Function

Public Function SafeDbl(ByVal v As Variant) As Double
    On Error Resume Next
    SafeDbl = CDbl(v)
    On Error GoTo 0
End Function

Public Function NowStr() As String
    NowStr = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function
