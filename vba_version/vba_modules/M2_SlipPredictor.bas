Attribute VB_Name = "M2_SlipPredictor"
Option Explicit

' ============================================================
' M2_SlipPredictor - Predict next-PO slip probability
' ============================================================
' Blends recent-90-day OTD vs all-time OTD per vendor.
' Output to Predictions sheet.
' ============================================================

Public Sub PredictSlipRisk()
    Dim wsPO As Worksheet, wsScore As Worksheet, wsPred As Worksheet
    On Error Resume Next
    Set wsPO = ThisWorkbook.Sheets(SHEET_PO_HISTORY)
    Set wsScore = ThisWorkbook.Sheets(SHEET_SCORECARDS)
    On Error GoTo 0
    If wsPO Is Nothing Or wsScore Is Nothing Then
        MsgBox "Can co PO_History VA Scorecards. Run 'Compute Scores' truoc.", vbExclamation
        Exit Sub
    End If

    Set wsPred = GetOrCreateSheet(SHEET_PREDICTIONS)
    wsPred.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "vendor_id", "vendor_name", "tier", "composite_score", _
        "historical_otd", "recent_90d_otd", "recent_n", _
        "predicted_slip_probability" _
    )
    WriteHeader wsPred, headers

    ' Find cutoff date = max(order_date) - 90 days
    Dim lastR As Long
    lastR = LastRow(wsPO, 1)
    Dim maxDate As Date
    maxDate = 0
    Dim r As Long
    For r = 2 To lastR
        Dim d As Date
        d = CDate(wsPO.Cells(r, 5).Value)
        If d > maxDate Then maxDate = d
    Next r
    Dim cutoff As Date
    cutoff = maxDate - 90

    ' Aggregate recent OTD per vendor
    Dim recent As Object
    Set recent = CreateObject("Scripting.Dictionary")
    For r = 2 To lastR
        Dim vid As String, orderDate As Date, isOTD As Boolean
        vid = CStr(wsPO.Cells(r, 2).Value)
        orderDate = CDate(wsPO.Cells(r, 5).Value)
        isOTD = (UCase(CStr(wsPO.Cells(r, 15).Value)) = "TRUE")
        If orderDate >= cutoff Then
            If Not recent.Exists(vid) Then
                Dim entry As Object
                Set entry = CreateObject("Scripting.Dictionary")
                entry.Add "count", 0
                entry.Add "otd_count", 0
                recent.Add vid, entry
            End If
            recent(vid)("count") = recent(vid)("count") + 1
            If isOTD Then recent(vid)("otd_count") = recent(vid)("otd_count") + 1
        End If
    Next r

    ' Loop through scorecards, compute prediction for each vendor
    Dim scoreLastR As Long
    scoreLastR = LastRow(wsScore, 1)
    Dim outRow As Long: outRow = 2

    For r = 2 To scoreLastR
        Dim svid As String: svid = CStr(wsScore.Cells(r, 1).Value)
        Dim sname As String: sname = CStr(wsScore.Cells(r, 2).Value)
        Dim stier As String: stier = CStr(wsScore.Cells(r, 12).Value)
        Dim sScore As Double: sScore = CDbl(wsScore.Cells(r, 11).Value)
        Dim histOTD As Double: histOTD = CDbl(wsScore.Cells(r, 4).Value)

        Dim recentN As Long: recentN = 0
        Dim recentOTD As Double: recentOTD = histOTD
        If recent.Exists(svid) Then
            recentN = recent(svid)("count")
            If recentN > 0 Then recentOTD = recent(svid)("otd_count") / recentN
        End If

        ' Blend recent vs historical, cap recent weight at 70%
        Dim blendWeight As Double
        blendWeight = Application.Min(recentN / 20#, 0.7)
        Dim predictedOTD As Double
        predictedOTD = blendWeight * recentOTD + (1 - blendWeight) * histOTD
        Dim slipProb As Double
        slipProb = 1 - predictedOTD

        With wsPred
            .Cells(outRow, 1).Value = svid
            .Cells(outRow, 2).Value = sname
            .Cells(outRow, 3).Value = stier
            .Cells(outRow, 4).Value = sScore
            .Cells(outRow, 5).Value = Round(histOTD, 4)
            .Cells(outRow, 6).Value = Round(recentOTD, 4)
            .Cells(outRow, 7).Value = recentN
            .Cells(outRow, 8).Value = Round(slipProb, 4)
        End With
        outRow = outRow + 1
    Next r

    ' Format percent columns
    FormatColumnAsPercent wsPred, 5
    FormatColumnAsPercent wsPred, 6
    FormatColumnAsPercent wsPred, 8

    ' Sort by slip prob desc (highest risk first)
    wsPred.Range("A2:H" & (outRow - 1)).Sort _
        Key1:=wsPred.Range("H2"), Order1:=xlDescending, Header:=xlNo

    AutoFitAllColumns wsPred

    ' Conditional format slip prob column
    Call ApplySlipRiskFormatting(wsPred, outRow - 1)

    LogInfo "INFO", "M2.PredictSlipRisk", "Predicted slip risk for " & (outRow - 2) & " vendors"
    MsgBox "Slip risk predictions done. Xem sheet '" & SHEET_PREDICTIONS & "'.", vbInformation
End Sub

Private Sub ApplySlipRiskFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = 2 To lastRow
        Dim prob As Double
        prob = CDbl(ws.Cells(r, 8).Value)
        If prob >= 0.3 Then
            ws.Cells(r, 8).Interior.Color = RGB(211, 47, 47)  ' red
            ws.Cells(r, 8).Font.Color = RGB(255, 255, 255)
        ElseIf prob >= 0.15 Then
            ws.Cells(r, 8).Interior.Color = RGB(245, 124, 0)  ' orange
            ws.Cells(r, 8).Font.Color = RGB(255, 255, 255)
        ElseIf prob >= 0.08 Then
            ws.Cells(r, 8).Interior.Color = RGB(251, 192, 45)  ' yellow
        Else
            ws.Cells(r, 8).Interior.Color = RGB(56, 142, 60)  ' green
            ws.Cells(r, 8).Font.Color = RGB(255, 255, 255)
        End If
        ws.Cells(r, 8).Font.Bold = True
    Next r
End Sub

' --- UDF: Predict slip prob given inputs ---
Public Function PredictSlip( _
    ByVal historicalOTD As Double, _
    ByVal recentOTD As Double, _
    ByVal recentN As Long) As Double
    Dim blendWeight As Double
    blendWeight = Application.Min(recentN / 20#, 0.7)
    PredictSlip = 1 - (blendWeight * recentOTD + (1 - blendWeight) * historicalOTD)
End Function
