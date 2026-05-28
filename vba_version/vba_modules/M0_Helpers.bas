Attribute VB_Name = "M0_Helpers"
Option Explicit

' ============================================================
' M0_Helpers - Utility functions dung chung
' ============================================================

Public Const SHEET_PO_HISTORY As String = "PO_History"
Public Const SHEET_SCORECARDS As String = "Scorecards"
Public Const SHEET_PREDICTIONS As String = "Predictions"
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_PARAMS As String = "Parameters"
Public Const SHEET_LOG As String = "Log"

' Composite score weights (sum to 1.0). Mirror Python src/reliability_scorer.py.
Public Const WEIGHT_OTIF As Double = 0.4
Public Const WEIGHT_LT_CONSISTENCY As Double = 0.2
Public Const WEIGHT_QUALITY As Double = 0.2
Public Const WEIGHT_COMMUNICATION As Double = 0.1
Public Const WEIGHT_TRADE As Double = 0.1

' --- Lay sheet theo ten, tao moi neu chua co ---
Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' --- Clear data row (giu header row 1) ---
Public Sub ClearSheetData(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = LastRow(ws, 1)
    If lastRow > 1 Then
        ws.Range("A2:Z" & lastRow).Clear
    End If
End Sub

' --- Last row of column ---
Public Function LastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' --- Last column of row ---
Public Function LastCol(ByVal ws As Worksheet, Optional ByVal row As Long = 1) As Long
    LastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
End Function

' --- Write header row with formatting ---
Public Sub WriteHeader(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlLeft
        End With
    Next i
    ws.Rows(1).RowHeight = 22
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)).AutoFilter
End Sub

' --- Log mot dong vao sheet Log ---
Public Sub LogInfo(ByVal level As String, ByVal source As String, ByVal msg As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_LOG)
    Dim nextRow As Long
    nextRow = LastRow(ws, 1) + 1
    If nextRow = 2 And ws.Cells(1, 1).Value = "" Then
        WriteHeader ws, Array("Timestamp", "Level", "Source", "Message")
        nextRow = 2
    End If
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = level
    ws.Cells(nextRow, 3).Value = source
    ws.Cells(nextRow, 4).Value = msg
End Sub

' --- Auto-fit columns ---
Public Sub AutoFitAllColumns(ByVal ws As Worksheet)
    ws.Cells.EntireColumn.AutoFit
End Sub

' --- Format as currency ---
Public Sub FormatColumnAsCurrency(ByVal ws As Worksheet, ByVal col As Long)
    Dim lastR As Long
    lastR = LastRow(ws, col)
    If lastR < 2 Then Exit Sub
    ws.Range(ws.Cells(2, col), ws.Cells(lastR, col)).NumberFormat = "$#,##0.00"
End Sub

' --- Format as percentage ---
Public Sub FormatColumnAsPercent(ByVal ws As Worksheet, ByVal col As Long)
    Dim lastR As Long
    lastR = LastRow(ws, col)
    If lastR < 2 Then Exit Sub
    ws.Range(ws.Cells(2, col), ws.Cells(lastR, col)).NumberFormat = "0.0%"
End Sub
