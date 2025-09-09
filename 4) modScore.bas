Attribute VB_Name = "modScore"
Option Explicit

Private Const DB_SHEET As String = "DB"
Private Const RES_SHEET As String = "Results"
Private Const TARGET_SHEET As String = "Target"
Private Const WEIGHTS_SHEET As String = "Weights"
Private Const CONFIG_SHEET As String = "Config"

Public Sub RankSimilar()
    Dim wsDB As Worksheet, wsRes As Worksheet, wsT As Worksheet, wsW As Worksheet, wsC As Worksheet
    Set wsDB = ThisWorkbook.Worksheets(DB_SHEET)
    Set wsRes = ThisWorkbook.Worksheets(RES_SHEET)
    Set wsT = ThisWorkbook.Worksheets(TARGET_SHEET)
    Set wsW = ThisWorkbook.Worksheets(WEIGHTS_SHEET)
    Set wsC = ThisWorkbook.Worksheets(CONFIG_SHEET)
    
    ' ---- Preset ----
    Dim preset As String: preset = UCase$(Trim$(wsC.Range("B7").Value))
    Select Case preset
        Case "FLOW"
            SetWeight wsW, "Anisotropy_E1E2_ratio", 10
            SetWeight wsW, "Anisotropy_Alpha_ratio", 10
            SetWeight wsW, "Flow_VI_or_MFR", 6
            SetWeight wsW, "pvT_coeff_vector_norm", 0
        Case "WARP"
            SetWeight wsW, "Anisotropy_E1E2_ratio", 30
            SetWeight wsW, "Anisotropy_Alpha_ratio", 35
            SetWeight wsW, "Flow_VI_or_MFR", 2
            SetWeight wsW, "pvT_coeff_vector_norm", 0
        Case Else ' Balanced
            SetWeight wsW, "Anisotropy_E1E2_ratio", 20
            SetWeight wsW, "Anisotropy_Alpha_ratio", 25
            SetWeight wsW, "Flow_VI_or_MFR", 3
            SetWeight wsW, "pvT_coeff_vector_norm", 0
    End Select
    
    ' ---- Target ----
    Dim tPoly$, tFillType$
    Dim tFillPct As Double, tRho As Double, tVI As Double, tCTE As Double, tE As Double
    Dim tMeltMid As Double, tMoldMid As Double
    
    tPoly = NormalizeBasePolymer(wsT.Range("B1").Value)
    tFillType = NormalizeFillerType(wsT.Range("B2").Value)
    tFillPct = SafeDbl(wsT.Range("B3").Value)
    tRho = SafeDbl(wsT.Range("B4").Value)
    tVI = SafeDbl(wsT.Range("B5").Value)
    tCTE = SafeDbl(wsT.Range("B6").Value)
    tE = SafeDbl(wsT.Range("B7").Value)
    tMeltMid = SafeDbl(wsT.Range("B8").Value)
    tMoldMid = SafeDbl(wsT.Range("B9").Value)
    
    ' ---- Weights ----
    Dim wPoly As Double, wFillType As Double, wFillPct As Double, wRho As Double
    Dim wFlow As Double, wPvT As Double, wAnE As Double, wAnCTE As Double
    Dim wCp As Double, wk As Double, wMelt As Double, wMold As Double
    
    wPoly = GetWeight(wsW, "PolymerMismatch")
    wFillType = GetWeight(wsW, "FillerTypeMismatch")
    wFillPct = GetWeight(wsW, "FillerWtPct_per_%")
    wRho = GetWeight(wsW, "SolidDensity_per_gcc")
    wFlow = GetWeight(wsW, "Flow_VI_or_MFR")
    wPvT = GetWeight(wsW, "pvT_coeff_vector_norm")
    wAnE = GetWeight(wsW, "Anisotropy_E1E2_ratio")
    wAnCTE = GetWeight(wsW, "Anisotropy_Alpha_ratio")
    wCp = GetWeight(wsW, "Cp_avg")
    wk = GetWeight(wsW, "k_avg")
    wMelt = GetWeight(wsW, "Rec_Melt_mid")
    wMold = GetWeight(wsW, "Rec_Mold_mid")
    
    ' ---- Results init ----
    wsRes.Cells.Clear
    wsRes.Range("A1").Resize(1, 7).Value = Array("Score","Supplier","Grade","Base","Filler","DataSource","Notes")
    
    Dim last As Long: last = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, out As Long: out = 2
    
    For r = 2 To last
        Dim score As Double: score = 0#
        Dim base As String: base = NormalizeBasePolymer(wsDB.Cells(r, 3).Value)
        Dim ft As String: ft = NormalizeFillerType(wsDB.Cells(r, 4).Value)
        Dim fwt As Double: fwt = SafeDbl(wsDB.Cells(r, 5).Value)
        Dim ds As String: ds = wsDB.Cells(r, 6).Value
        Dim rho As Double: rho = SafeDbl(wsDB.Cells(r, 8).Value) ' SolidDensity
        Dim e1 As Double: e1 = SafeDbl(wsDB.Cells(r, 27).Value)
        Dim e2 As Double: e2 = SafeDbl(wsDB.Cells(r, 28).Value)
        Dim a1 As Double: a1 = SafeDbl(wsDB.Cells(r, 33).Value)
        Dim a2 As Double: a2 = SafeDbl(wsDB.Cells(r, 34).Value)
        Dim meltMin As Double: meltMin = SafeDbl(wsDB.Cells(r, 35).Value)
        Dim meltMax As Double: meltMax = SafeDbl(wsDB.Cells(r, 36).Value)
        Dim moldMin As Double: moldMin = SafeDbl(wsDB.Cells(r, 37).Value)
        Dim moldMax As Double: moldMax = SafeDbl(wsDB.Cells(r, 38).Value)
        Dim viText As String: viText = wsDB.Cells(r, 44).Value
        
        ' Polymer / Filler penalties
        If tPoly <> "" And tPoly <> "ANY" Then
            If UCase$(tPoly) <> UCase$(base) Then score = score + wPoly
        End If
        If tFillType <> "" And UCase$(tFillType) <> "ANY" Then
            If UCase$(tFillType) <> UCase$(ft) Then score = score + wFillType
        End If
        
        ' Numeric differences
        If tFillPct > 0 And fwt > 0 Then score = score + wFillPct * Abs(fwt - tFillPct)
        If tRho > 0 And rho > 0 Then score = score + wRho * Abs(rho - tRho)
        
        ' Flow approx by VI (v1簡易)
        If tVI > 0 Then
            Dim viNum As Double: viNum = ExtractVI(viText)
            If viNum > 0 Then score = score + wFlow * Abs(viNum - tVI)
        End If
        
        ' Anisotropy (E ratio / CTE ratio)
        If e1 > 0 And e2 > 0 And tE > 0 Then
            Dim tRatioE As Double: tRatioE = tE
            If tRatioE <= 5 Then
                score = score + wAnE * Abs((e1 / e2) - tRatioE)
            Else
                score = score + wAnE * Abs((e1 / 1000#) - tE) ' if user put absolute E (GPa)
            End If
        End If
        If a1 > 0 And a2 > 0 And tCTE > 0 Then
            Dim tRatioA As Double: tRatioA = tCTE
            If tRatioA <= 5 Then
                score = score + wAnCTE * Abs((a2 / a1) - tRatioA)
            Else
                score = score + wAnCTE * Abs(((a1 + a2) / 2 - tCTE) * 1E6) ' average CTE (um/mK 入力時)
            End If
        End If
        
        ' Recommended temperature window
        If tMeltMid > 0 And meltMin > 0 And meltMax > 0 Then
            score = score + wMelt * Abs(((meltMin + meltMax) / 2) - tMeltMid)
        End If
        If tMoldMid > 0 And moldMin > 0 And moldMax > 0 Then
            score = score + wMold * Abs(((moldMin + moldMax) / 2) - tMoldMid)
        End If
        
        ' Output
        wsRes.Cells(out, 1).Value = Round(score, 3)
        wsRes.Cells(out, 2).Value = wsDB.Cells(r, 1).Value
        wsRes.Cells(out, 3).Value = wsDB.Cells(r, 2).Value
        wsRes.Cells(out, 4).Value = base
        wsRes.Cells(out, 5).Value = ft & " " & fwt & "%"
        wsRes.Cells(out, 6).Value = ds
        wsRes.Cells(out, 7).Value = viText
        out = out + 1
    Next r
    
    wsRes.Range("A1").CurrentRegion.Sort Key1:=wsRes.Range("A2"), Order1:=xlAscending, Header:=xlYes
    MsgBox "ランキング完了（Scoreが小さいほど類似）", vbInformation
End Sub

Private Function GetWeight(ByVal ws As Worksheet, ByVal key As String) As Double
    Dim f As Range
    Set f = ws.Range("A:A").Find(What:=key, LookAt:=xlWhole, LookIn:=xlValues)
    If f Is Nothing Then
        GetWeight = 0#
    Else
        GetWeight = SafeDbl(f.Offset(0, 1).Value)
    End If
End Function

Private Sub SetWeight(ByVal ws As Worksheet, ByVal key As String, ByVal v As Double)
    Dim f As Range
    Set f = ws.Range("A:A").Find(What:=key, LookAt:=xlWhole, LookIn:=xlValues)
    If Not f Is Nothing Then f.Offset(0, 1).Value = v
End Sub

Private Function ExtractVI(ByVal s As String) As Double
    ' "VI(300)0157" 等の数字部分をざっくり抽出
    Dim i As Long, c As String, buf As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[0-9.]" Then buf = buf & c
    Next i
    If buf <> "" Then ExtractVI = Val(buf) Else ExtractVI = 0#
End Function
