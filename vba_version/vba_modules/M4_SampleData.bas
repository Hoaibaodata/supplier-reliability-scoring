Attribute VB_Name = "M4_SampleData"
Option Explicit

' ============================================================
' M4_SampleData - Generate 200 sample POs cho demo
' ============================================================
' Mirrors Python src/po_data.py vendor profiles.
' ============================================================

Public Sub GenerateSamplePOHistory()
    Dim wsPO As Worksheet
    Set wsPO = GetOrCreateSheet(SHEET_PO_HISTORY)
    wsPO.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "po_id", "vendor_id", "vendor_name", "category", _
        "order_date", "requested_delivery_date", "requested_qty", _
        "actual_delivery_date", "delivered_qty", "qc_pass_qty", "qc_fail_qty", _
        "requested_lt_days", "actual_lt_days", "slip_days", _
        "is_otd", "is_in_full", "confirm_lag_days" _
    )
    WriteHeader wsPO, headers

    ' Vendor profiles (matches Python VENDOR_PROFILES)
    Dim vendors As Variant
    vendors = Array( _
        Array("V001", "Premium Plus Inc.", 0.95, 0.99, 14, 1), _
        Array("V002", "Reliable Co.", 0.92, 0.98, 18, 2), _
        Array("V003", "Standard Supply", 0.85, 0.95, 14, 3), _
        Array("V004", "Mid-Tier Logistics", 0.82, 0.93, 16, 4), _
        Array("V005", "Generic Parts Ltd.", 0.78, 0.9, 21, 5), _
        Array("V006", "Fast Ship Corp.", 0.9, 0.97, 7, 2), _
        Array("V007", "Budget Source", 0.7, 0.85, 25, 7), _
        Array("V008", "Global Exports", 0.88, 0.96, 20, 3), _
        Array("V009", "Risky Supplier", 0.6, 0.8, 28, 10), _
        Array("V010", "New Vendor", 0.83, 0.92, 15, 4) _
    )

    Dim categories As Variant
    categories = Array("Electronics", "Mechanical", "Plastics", "Chemicals")

    Randomize 42  ' deterministic seed (approximation in VBA)

    Dim startDate As Date
    startDate = DateSerial(2025, 6, 1)

    Dim i As Long
    Dim row As Long: row = 2
    Dim numPO As Long: numPO = 200

    For i = 1 To numPO
        Dim vIdx As Long: vIdx = Int(Rnd() * 10)
        Dim v As Variant: v = vendors(vIdx)

        Dim vid As String: vid = v(0)
        Dim vname As String: vname = v(1)
        Dim otdBase As Double: otdBase = v(2)
        Dim qcBase As Double: qcBase = v(3)
        Dim ltMean As Double: ltMean = v(4)
        Dim ltStd As Double: ltStd = v(5)

        Dim cat As String: cat = categories(Int(Rnd() * 4))

        ' Order date in past year
        Dim orderDate As Date
        orderDate = startDate + Int(Rnd() * 365)

        ' Requested LT (normal distribution approx)
        Dim reqLT As Long
        reqLT = Application.Max(1, Int(ltMean + (Rnd() - 0.5) * ltStd))
        Dim reqDeliv As Date: reqDeliv = orderDate + reqLT

        ' OTD probability (category modifier)
        Dim otdProb As Double: otdProb = otdBase
        If cat = "Chemicals" Then otdProb = otdProb * 0.95
        If cat = "Mechanical" Then otdProb = Application.Min(otdProb * 1.02, 0.99)

        Dim isOTD As Boolean: isOTD = (Rnd() < otdProb)
        Dim slipDays As Long: slipDays = IIf(isOTD, 0, Int(Rnd() * 13) + 1)
        Dim actualDeliv As Date: actualDeliv = reqDeliv + slipDays
        Dim actualLT As Long: actualLT = DateDiff("d", orderDate, actualDeliv)

        ' Quantities
        Dim reqQty As Long: reqQty = Int(Rnd() * 9500) + 500
        Dim isInFull As Boolean: isInFull = (Rnd() < 0.92)
        Dim deliveredQty As Long
        If isInFull Then
            deliveredQty = reqQty
        Else
            deliveredQty = Int(reqQty * (0.85 + Rnd() * 0.14))
        End If

        ' QC
        Dim qcRate As Double
        qcRate = Application.Min(1, Application.Max(0.5, qcBase + (Rnd() - 0.5) * 0.06))
        Dim qcPassQty As Long: qcPassQty = Int(deliveredQty * qcRate)
        Dim qcFailQty As Long: qcFailQty = deliveredQty - qcPassQty

        ' Communication lag
        Dim confirmLag As Long: confirmLag = Int(-Log(Rnd()) * 2#)
        If confirmLag < 0 Then confirmLag = 0

        ' Write row
        With wsPO
            .Cells(row, 1).Value = "PO" & Format(i, "00000")
            .Cells(row, 2).Value = vid
            .Cells(row, 3).Value = vname
            .Cells(row, 4).Value = cat
            .Cells(row, 5).Value = orderDate
            .Cells(row, 6).Value = reqDeliv
            .Cells(row, 7).Value = reqQty
            .Cells(row, 8).Value = actualDeliv
            .Cells(row, 9).Value = deliveredQty
            .Cells(row, 10).Value = qcPassQty
            .Cells(row, 11).Value = qcFailQty
            .Cells(row, 12).Value = reqLT
            .Cells(row, 13).Value = actualLT
            .Cells(row, 14).Value = slipDays
            .Cells(row, 15).Value = isOTD
            .Cells(row, 16).Value = isInFull
            .Cells(row, 17).Value = confirmLag
        End With

        row = row + 1
    Next i

    ' Format dates
    wsPO.Range("E2:E" & (row - 1)).NumberFormat = "yyyy-mm-dd"
    wsPO.Range("F2:F" & (row - 1)).NumberFormat = "yyyy-mm-dd"
    wsPO.Range("H2:H" & (row - 1)).NumberFormat = "yyyy-mm-dd"

    AutoFitAllColumns wsPO

    LogInfo "INFO", "M4.GenerateSamplePOHistory", "Da tao " & numPO & " sample POs"
    MsgBox "Da tao " & numPO & " sample PO records. Bay gio bam 'Compute Scores'.", vbInformation
End Sub
