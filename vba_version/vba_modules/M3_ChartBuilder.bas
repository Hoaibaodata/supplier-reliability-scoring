Attribute VB_Name = "M3_ChartBuilder"
Option Explicit

' ============================================================
' M3_ChartBuilder - Tao chart truc quan tren Dashboard
' ============================================================
' Bar chart cua composite score by vendor + tier color coding.
' ============================================================

Public Sub BuildScorecardChart()
    Dim wsScore As Worksheet, wsDash As Worksheet
    On Error Resume Next
    Set wsScore = ThisWorkbook.Sheets(SHEET_SCORECARDS)
    On Error GoTo 0
    If wsScore Is Nothing Then
        MsgBox "Khong tim thay Scorecards. Run 'Compute Scores' truoc.", vbExclamation
        Exit Sub
    End If

    Set wsDash = GetOrCreateSheet(SHEET_DASHBOARD)

    ' Remove old chart if exists
    Dim ch As ChartObject
    For Each ch In wsDash.ChartObjects
        ch.Delete
    Next ch

    Dim lastR As Long: lastR = LastRow(wsScore, 1)
    If lastR < 2 Then Exit Sub

    ' Create chart
    Dim chartObj As ChartObject
    Set chartObj = wsDash.ChartObjects.Add(Left:=50, Top:=200, Width:=720, Height:=400)
    Dim cht As Chart
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Supplier Composite Score (0-100)"
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = True

        ' Set data: vendor_name (col B) as category, composite_score (col K) as value
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Composite Score"
        .SeriesCollection(1).Values = wsScore.Range("K2:K" & lastR)
        .SeriesCollection(1).XValues = wsScore.Range("B2:B" & lastR)

        .HasLegend = False
        .Axes(xlValue).MaximumScale = 100
        .Axes(xlValue).MinimumScale = 0

        ' Color bars by tier
        Dim r As Long
        Dim pt As Point
        For r = 2 To lastR
            Set pt = .SeriesCollection(1).Points(r - 1)
            Dim tier As String
            tier = CStr(wsScore.Cells(r, 12).Value)
            Select Case tier
                Case "A-Preferred"
                    pt.Format.Fill.ForeColor.RGB = RGB(56, 142, 60)
                Case "B-Standard"
                    pt.Format.Fill.ForeColor.RGB = RGB(25, 118, 210)
                Case "C-Watch"
                    pt.Format.Fill.ForeColor.RGB = RGB(245, 124, 0)
                Case "D-Risk"
                    pt.Format.Fill.ForeColor.RGB = RGB(211, 47, 47)
            End Select
        Next r
    End With

    LogInfo "INFO", "M3.BuildScorecardChart", "Chart built with " & (lastR - 1) & " vendors"
End Sub
