Attribute VB_Name = "M1_KPICalculator"
Option Explicit

' ============================================================
' M1_KPICalculator - Aggregate PO history -> vendor KPIs
' ============================================================
' Mirrors Python src/reliability_scorer.py logic.
' Reads PO_History sheet, outputs to Scorecards sheet.
' ============================================================

' Expected PO_History columns (case-sensitive):
'   A: po_id
'   B: vendor_id
'   C: vendor_name
'   D: category
'   E: order_date
'   F: requested_delivery_date
'   G: requested_qty
'   H: actual_delivery_date
'   I: delivered_qty
'   J: qc_pass_qty
'   K: qc_fail_qty
'   L: requested_lt_days
'   M: actual_lt_days
'   N: slip_days
'   O: is_otd (TRUE/FALSE)
'   P: is_in_full (TRUE/FALSE)
'   Q: confirm_lag_days

Private Type VendorStats
    vendorId As String
    vendorName As String
    numPOs As Long
    otdCount As Long
    otifCount As Long  ' both OTD and in_full
    totalLT As Double
    sumLTSquared As Double  ' for std calc
    ltValues As Object  ' Collection for std dev
    totalDeliveredQty As Double
    totalQCFailQty As Double
    sumConfirmLag As Double
End Type

' --- Main entry: read PO sheet, compute KPIs, write Scorecards ---
Public Sub AggregateAndScore()
    Dim wsPO As Worksheet
    On Error Resume Next
    Set wsPO = ThisWorkbook.Sheets(SHEET_PO_HISTORY)
    On Error GoTo 0
    If wsPO Is Nothing Then
        MsgBox "Khong tim thay sheet '" & SHEET_PO_HISTORY & "'. Generate sample data truoc.", vbExclamation
        Exit Sub
    End If

    Dim lastR As Long
    lastR = LastRow(wsPO, 1)
    If lastR < 2 Then
        MsgBox "PO_History sheet trong. Generate sample data truoc.", vbExclamation
        Exit Sub
    End If

    LogInfo "INFO", "M1.AggregateAndScore", "Bat dau xu ly " & (lastR - 1) & " PO records"

    ' Aggregate by vendor using Dictionary
    Dim vendors As Object
    Set vendors = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim vid As String, vname As String
    Dim ltVal As Double
    Dim isOTD As Boolean, isInFull As Boolean
    Dim delivered As Double, qcFail As Double, confirmLag As Double

    For r = 2 To lastR
        vid = CStr(wsPO.Cells(r, 2).Value)
        vname = CStr(wsPO.Cells(r, 3).Value)
        If vid = "" Then GoTo NextPO

        ' Create vendor entry if new
        If Not vendors.Exists(vid) Then
            Dim stats As Object
            Set stats = CreateObject("Scripting.Dictionary")
            stats.Add "vendor_id", vid
            stats.Add "vendor_name", vname
            stats.Add "num_pos", 0
            stats.Add "otd_count", 0
            stats.Add "otif_count", 0
            stats.Add "total_lt", 0#
            stats.Add "lt_values", CreateObject("Scripting.Dictionary")  ' use Dict as ordered list
            stats.Add "total_delivered_qty", 0#
            stats.Add "total_qc_fail_qty", 0#
            stats.Add "sum_confirm_lag", 0#
            vendors.Add vid, stats
        End If

        ltVal = CDbl(wsPO.Cells(r, 13).Value)
        isOTD = (UCase(CStr(wsPO.Cells(r, 15).Value)) = "TRUE")
        isInFull = (UCase(CStr(wsPO.Cells(r, 16).Value)) = "TRUE")
        delivered = CDbl(wsPO.Cells(r, 9).Value)
        qcFail = CDbl(wsPO.Cells(r, 11).Value)
        confirmLag = CDbl(wsPO.Cells(r, 17).Value)

        vendors(vid)("num_pos") = vendors(vid)("num_pos") + 1
        If isOTD Then vendors(vid)("otd_count") = vendors(vid)("otd_count") + 1
        If isOTD And isInFull Then vendors(vid)("otif_count") = vendors(vid)("otif_count") + 1
        vendors(vid)("total_lt") = vendors(vid)("total_lt") + ltVal
        vendors(vid)("lt_values").Add CStr(vendors(vid)("lt_values").Count + 1), ltVal
        vendors(vid)("total_delivered_qty") = vendors(vid)("total_delivered_qty") + delivered
        vendors(vid)("total_qc_fail_qty") = vendors(vid)("total_qc_fail_qty") + qcFail
        vendors(vid)("sum_confirm_lag") = vendors(vid)("sum_confirm_lag") + confirmLag
NextPO:
    Next r

    ' Now compute final KPIs + composite score, write to Scorecards
    Dim wsScore As Worksheet
    Set wsScore = GetOrCreateSheet(SHEET_SCORECARDS)
    wsScore.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "vendor_id", "vendor_name", "num_pos", "otd_rate", "otif_rate", _
        "avg_lt_days", "sigma_lt_weeks", "quality_pass_rate", "communication_score", _
        "trade_score", "composite_score", "tier" _
    )
    WriteHeader wsScore, headers

    Dim maxPOs As Long
    Dim key As Variant
    maxPOs = 1
    For Each key In vendors.Keys
        If vendors(key)("num_pos") > maxPOs Then maxPOs = vendors(key)("num_pos")
    Next key

    Dim outRow As Long
    outRow = 2
    For Each key In vendors.Keys
        Dim v As Object
        Set v = vendors(key)

        Dim numPos As Long: numPos = v("num_pos")
        Dim otdRate As Double: otdRate = v("otd_count") / numPos
        Dim otifRate As Double: otifRate = v("otif_count") / numPos
        Dim avgLT As Double: avgLT = v("total_lt") / numPos
        Dim stdLT As Double: stdLT = ComputeStdDev(v("lt_values"), avgLT)
        Dim sigmaLTWeeks As Double: sigmaLTWeeks = stdLT / 7#
        Dim qualityRate As Double
        If v("total_delivered_qty") > 0 Then
            qualityRate = 1 - (v("total_qc_fail_qty") / v("total_delivered_qty"))
        Else
            qualityRate = 1
        End If
        Dim ltCV As Double: ltCV = IIf(avgLT > 0, stdLT / avgLT, 0)
        Dim ltConsistency As Double: ltConsistency = Application.Min(Application.Max(1 - ltCV, 0), 1)
        Dim commScore As Double
        commScore = Application.Min(Application.Max(1 - (v("sum_confirm_lag") / numPos / 7#), 0), 1)
        Dim tradeScore As Double: tradeScore = numPos / maxPOs

        ' Composite (mirror Python formula)
        Dim composite As Double
        composite = WEIGHT_OTIF * otifRate _
                  + WEIGHT_LT_CONSISTENCY * ltConsistency _
                  + WEIGHT_QUALITY * qualityRate _
                  + WEIGHT_COMMUNICATION * commScore _
                  + WEIGHT_TRADE * tradeScore
        composite = composite * 100  ' 0-100 scale

        Dim tier As String
        tier = AssignTier(composite)

        With wsScore
            .Cells(outRow, 1).Value = v("vendor_id")
            .Cells(outRow, 2).Value = v("vendor_name")
            .Cells(outRow, 3).Value = numPos
            .Cells(outRow, 4).Value = otdRate
            .Cells(outRow, 5).Value = otifRate
            .Cells(outRow, 6).Value = Round(avgLT, 2)
            .Cells(outRow, 7).Value = Round(sigmaLTWeeks, 4)
            .Cells(outRow, 8).Value = qualityRate
            .Cells(outRow, 9).Value = Round(commScore, 4)
            .Cells(outRow, 10).Value = Round(tradeScore, 4)
            .Cells(outRow, 11).Value = Round(composite, 2)
            .Cells(outRow, 12).Value = tier
        End With

        outRow = outRow + 1
    Next key

    ' Format
    FormatColumnAsPercent wsScore, 4
    FormatColumnAsPercent wsScore, 5
    FormatColumnAsPercent wsScore, 8

    ' Sort by composite_score desc
    wsScore.Range("A2:L" & (outRow - 1)).Sort _
        Key1:=wsScore.Range("K2"), Order1:=xlDescending, Header:=xlNo

    AutoFitAllColumns wsScore

    ' Apply tier conditional formatting
    Call ApplyTierFormatting(wsScore, outRow - 1)

    LogInfo "INFO", "M1.AggregateAndScore", "Hoan tat " & (outRow - 2) & " vendors"
    MsgBox "Da tinh xong " & (outRow - 2) & " vendor scorecards. Xem sheet '" & SHEET_SCORECARDS & "'.", vbInformation
End Sub

' --- Std dev tu collection of values ---
Private Function ComputeStdDev(ByVal values As Object, ByVal mean As Double) As Double
    Dim n As Long, sumSq As Double, val As Variant
    n = values.Count
    If n < 2 Then
        ComputeStdDev = 0
        Exit Function
    End If
    sumSq = 0
    For Each val In values.Items
        sumSq = sumSq + (val - mean) ^ 2
    Next val
    ComputeStdDev = Sqr(sumSq / (n - 1))
End Function

' --- UDF: Assign tier from composite score ---
Public Function AssignTier(ByVal score As Double) As String
    If score >= 85 Then
        AssignTier = "A-Preferred"
    ElseIf score >= 75 Then
        AssignTier = "B-Standard"
    ElseIf score >= 60 Then
        AssignTier = "C-Watch"
    Else
        AssignTier = "D-Risk"
    End If
End Function

' --- UDF: Compute composite score from individual scores ---
Public Function ComputeComposite( _
    ByVal otif As Double, ByVal ltConsistency As Double, _
    ByVal quality As Double, ByVal communication As Double, _
    ByVal trade As Double) As Double
    ComputeComposite = (WEIGHT_OTIF * otif _
                      + WEIGHT_LT_CONSISTENCY * ltConsistency _
                      + WEIGHT_QUALITY * quality _
                      + WEIGHT_COMMUNICATION * communication _
                      + WEIGHT_TRADE * trade) * 100
End Function

' --- Conditional format Scorecards by tier ---
Private Sub ApplyTierFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = 2 To lastRow
        Dim tier As String
        tier = CStr(ws.Cells(r, 12).Value)
        Select Case tier
            Case "A-Preferred"
                ws.Cells(r, 12).Interior.Color = RGB(56, 142, 60)  ' green
                ws.Cells(r, 12).Font.Color = RGB(255, 255, 255)
            Case "B-Standard"
                ws.Cells(r, 12).Interior.Color = RGB(25, 118, 210)  ' blue
                ws.Cells(r, 12).Font.Color = RGB(255, 255, 255)
            Case "C-Watch"
                ws.Cells(r, 12).Interior.Color = RGB(245, 124, 0)  ' orange
                ws.Cells(r, 12).Font.Color = RGB(255, 255, 255)
            Case "D-Risk"
                ws.Cells(r, 12).Interior.Color = RGB(211, 47, 47)  ' red
                ws.Cells(r, 12).Font.Color = RGB(255, 255, 255)
        End Select
        ws.Cells(r, 12).Font.Bold = True
    Next r
End Sub
