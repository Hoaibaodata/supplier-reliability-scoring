Attribute VB_Name = "M5_Dashboard"
Option Explicit

' ============================================================
' M5_Dashboard - Action buttons + workflow controller
' ============================================================

' --- Bam nut: Run all-in-one ---
Public Sub Action_RunFullAnalysis()
    Dim wsPO As Worksheet
    On Error Resume Next
    Set wsPO = ThisWorkbook.Sheets(SHEET_PO_HISTORY)
    On Error GoTo 0
    If wsPO Is Nothing Or LastRow(wsPO, 1) < 2 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Chua co PO_History. Generate sample 200 PO truoc?", vbYesNo + vbQuestion, "Setup")
        If ans = vbYes Then
            Call GenerateSamplePOHistory
        Else
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Call AggregateAndScore  ' M1
    Call PredictSlipRisk    ' M2
    Call BuildScorecardChart ' M3
    Application.ScreenUpdating = True

    MsgBox "Phan tich xong! Xem Scorecards + Predictions + Dashboard.", vbInformation
End Sub

' --- Bam nut: Generate sample only ---
Public Sub Action_GenerateSample()
    Call GenerateSamplePOHistory
End Sub

' --- Bam nut: Compute scores ---
Public Sub Action_ComputeScores()
    Call AggregateAndScore
End Sub

' --- Bam nut: Predict slip ---
Public Sub Action_PredictSlip()
    Call PredictSlipRisk
End Sub

' --- Bam nut: Build chart ---
Public Sub Action_BuildChart()
    Call BuildScorecardChart
End Sub

' --- Bam nut: Reset (clear all outputs) ---
Public Sub Action_ResetOutputs()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Xoa Scorecards + Predictions + Dashboard charts?", vbYesNo + vbQuestion, "Reset")
    If ans <> vbYes Then Exit Sub

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SHEET_SCORECARDS).Cells.Clear
    ThisWorkbook.Sheets(SHEET_PREDICTIONS).Cells.Clear
    Dim ch As ChartObject
    For Each ch In ThisWorkbook.Sheets(SHEET_DASHBOARD).ChartObjects
        ch.Delete
    Next ch
    Application.DisplayAlerts = True
    On Error GoTo 0

    LogInfo "INFO", "M5.Action_ResetOutputs", "Outputs cleared"
    MsgBox "Reset xong.", vbInformation
End Sub
