Attribute VB_Name = "modBuild"
Option Explicit

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = nm
    End If
End Function

Public Sub BuildWorkbook(Optional ByVal SaveAsPath As String = "")
    Dim wsDB As Worksheet, wsT As Worksheet, wsW As Worksheet, wsC As Worksheet, wsR As Worksheet, wsL As Worksheet
    
    Set wsDB = EnsureSheet("DB")
    Set wsT = EnsureSheet("Target")
    Set wsW = EnsureSheet("Weights")
    Set wsC = EnsureSheet("Config")
    Set wsR = EnsureSheet("Results")
    Set wsL = EnsureSheet("ImportLog")
    
    ' ---- DB headers ----
    Dim h As Variant
    h = Array( _
        "Supplier","Grade","BasePolymer","FillerType","FillerWtPct","DataSource", _
        "MeltDensity_gcc","SolidDensity_gcc", _
        "b5","b6","b1m","b2m","b3m","b4m","b1s","b2s","b3s","b4s", _
        "CP_table(raw)","K_table(raw)", _
        "CP_80C","CP_140C","CP_200C", _
        "K_100C","K_150C","K_200C", _
        "E1_MPa","E2_MPa","nu12","nu23","G12_MPa", _
        "Alpha1_1perC","Alpha2_1perC", _
        "Rec_Melt_Tmin_C","Rec_Melt_Tmax_C", _
        "Rec_Mold_Tmin_C","Rec_Mold_Tmax_C", _
        "MaxMelt_C","MaxShearStress_MPa","MaxShearRate_1persec", _
        "FillerDensity_gcc", _
        "VI_text","MFR_g_10min","MFR_Temp_C","MFR_Load_kg", _
        "Ttrans_C", _
        "Notes","SourceFile" _
    )
    wsDB.Cells.Clear
    wsDB.Range("A1").Resize(1, UBound(h) + 1).Value = h
    wsDB.Rows(1).Font.Bold = True
    wsDB.Rows(1).Interior.ColorIndex = 15
    wsDB.Columns("A:AZ").ColumnWidth = 14
    wsDB.Rows(2).Select
    
    ' ---- Target ----
    wsT.Cells.Clear
    wsT.Range("A1:B1").Value = Array("Parameter","Value")
    Dim t As Variant
    t = Array( _
        Array("BasePolymer (e.g., PA46)","PA46"), _
        Array("FillerType (GF/Mineral/CF/None/Any)","Any"), _
        Array("FillerWtPct (%)",""), _
        Array("SolidDensity_gcc",""), _
        Array("VI or Flow index (optional)",""), _
        Array("CTE average (um/mK) or Alpha1/2 if known",""), _
        Array("E_dry_GPa (approx or ratio)",""), _
        Array("MeltTemp_mid_C (optional)",""), _
        Array("MoldTemp_mid_C (optional)",""), _
        Array("Notes","") _
    )
    Dim i As Long
    For i = 0 To UBound(t)
        wsT.Cells(i + 2, 1).Resize(1, 2).Value = t(i)
    Next
    wsT.Columns("A:B").AutoFit
    
    ' ---- Weights ----
    wsW.Cells.Clear
    wsW.Range("A1:B1").Value = Array("WeightName","Value")
    Dim w As Variant
    w = Array( _
        Array("PolymerMismatch", 100), _
        Array("FillerTypeMismatch", 40), _
        Array("FillerWtPct_per_%", 1#), _
        Array("SolidDensity_per_gcc", 20#), _
        Array("Flow_VI_or_MFR", 3#), _
        Array("pvT_coeff_vector_norm", 0#), _
        Array("Anisotropy_E1E2_ratio", 20#), _
        Array("Anisotropy_Alpha_ratio", 25#), _
        Array("Cp_avg", 2#), _
        Array("k_avg", 2#), _
        Array("Rec_Melt_mid", 0.5), _
        Array("Rec_Mold_mid", 0.5) _
    )
    For i = 0 To UBound(w)
        wsW.Cells(i + 2, 1).Resize(1, 2).Value = w(i)
    Next
    wsW.Columns("A:B").AutoFit
    
    ' ---- Config ----
    wsC.Cells.Clear
    wsC.Range("A1:B1").Value = Array("Key","Value")
    Dim c As Variant
    c = Array( _
        Array("FlowGrid_Temps_C", "260,280,300,320"), _
        Array("FlowGrid_Shear_1persec", "100,1000,10000,50000"), _
        Array("pvT_Temps_C", "40,80,120,200"), _
        Array("pvT_Pressures_MPa", "0,50,100"), _
        Array("Eval_Cp_Temps_C", "80,140,200"), _
        Array("Eval_k_Temps_C", "100,150,200"), _
        Array("Preset", "Balanced") _
    )
    For i = 0 To UBound(c)
        wsC.Cells(i + 2, 1).Resize(1, 2).Value = c(i)
    Next
    wsC.Columns("A:B").AutoFit
    
    ' ---- Results / Log ----
    wsR.Cells.Clear
    wsR.Range("A1:G1").Value = Array("Score","Supplier","Grade","Base","Filler","DataSource","Notes")
    wsR.Rows(1).Font.Bold = True
    
    wsL.Cells.Clear
    wsL.Range("A1:D1").Value = Array("When","File","ParsedRows","Warnings")
    wsL.Rows(1).Font.Bold = True
    
    If Len(SaveAsPath) > 0 Then
        Application.DisplayAlerts = False
        ThisWorkbook.SaveAs Filename:=SaveAsPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
    End If
    
    MsgBox "Workbook structure built.", vbInformation
End Sub

' ワンボタンで「作成→読み込み→ランキング」
Public Sub BuildAndRun()
    BuildWorkbook
    ImportReports
    RankSimilar
End Sub
