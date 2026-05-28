# Supplier Reliability Scoring — VBA Version Setup

> Excel-only edition. KHÔNG cần Python, KHÔNG cần ML libraries. Chạy 100% trong Excel + VBA.
> Mục đích: cho team operations dùng instantly khi IT chặn Python.

## Yêu cầu

- Microsoft Excel 2016+ (Windows)
- Macros enabled (Trust Center)

## Bước 1 — Generate starter Excel file

Mở PowerShell, chạy:

```powershell
cd "D:\HOC_TAP\supplier-reliability-scoring\vba_version"
python generate_starter_xlsx.py
```

→ Tạo file `SupplierReliability.xlsx` với 6 sheet pre-populated (Dashboard, PO_History, Scorecards, Predictions, Parameters, Log).

> **Tại sao cần Python ở bước này?** Chỉ để generate `.xlsx` starter có headers + formatting. Sau bước này, KHÔNG cần Python nữa.

## Bước 2 — Save As .xlsm

1. Mở `SupplierReliability.xlsx` trong Excel
2. File → Save As → đổi **Save as type** thành **"Excel Macro-Enabled Workbook (.xlsm)"**
3. Đặt tên: `SupplierReliability.xlsm`
4. Save

## Bước 3 — Import 5 VBA modules

1. Bấm **Alt + F11** → mở Visual Basic Editor (VBE)
2. File → Import File → chọn lần lượt 5 file trong `vba_modules/`:
   - `M0_Helpers.bas`
   - `M1_KPICalculator.bas`
   - `M2_SlipPredictor.bas`
   - `M3_ChartBuilder.bas`
   - `M4_SampleData.bas`
   - `M5_Dashboard.bas`
3. Quay lại Excel → Save (giữ format `.xlsm`)

## Bước 4 — Tạo nút bấm trên Dashboard

Trên sheet **Dashboard**:

1. Developer tab → Insert → Button (Form Control) → vẽ nút
2. Khi popup "Assign Macro" hiện, gán macro tương ứng:

| Caption nút | Macro |
|---|---|
| 🚀 **Run Full Analysis** | `Action_RunFullAnalysis` |
| 📥 Generate Sample PO | `Action_GenerateSample` |
| 📊 Compute Scores | `Action_ComputeScores` |
| 🎯 Predict Slip Risk | `Action_PredictSlip` |
| 📈 Build Chart | `Action_BuildChart` |
| 🔄 Reset Outputs | `Action_ResetOutputs` |

## Bước 5 — Chạy thử

Bấm **Run Full Analysis** → kết quả popup. Verify:

- Sheet **PO_History**: 200 dòng PO sample
- Sheet **Scorecards**: 10 vendors với composite score 0-100, tier color-coded
- Sheet **Predictions**: 10 vendors sắp xếp theo slip probability, top = highest risk
- Sheet **Dashboard**: chart bar đứng cao thấp theo composite score, màu theo tier
- Sheet **Log**: track lại các action

## Cấu trúc code (VBA patterns demonstrated)

| Module | Patterns |
|---|---|
| `M0_Helpers` | Constants, helper functions, sheet creation, header formatting, conditional formatting setup |
| `M1_KPICalculator` | Scripting.Dictionary cho aggregation, iterate rows, std dev calc, UDF (`AssignTier`, `ComputeComposite`) |
| `M2_SlipPredictor` | Date arithmetic, second-pass aggregation, conditional formatting based on values, UDF (`PredictSlip`) |
| `M3_ChartBuilder` | Chart objects creation, color customization per data point, ChartObjects collection |
| `M4_SampleData` | Random number generation, jagged arrays for profiles, normal distribution approximation |
| `M5_Dashboard` | Action routing, Application.ScreenUpdating optimization, message box workflow |

## So với Python version

| Feature | Python version | VBA version |
|---|---|---|
| PO history size | 500 | 200 (Excel performance) |
| Vendors | 10 | 10 (cùng profile) |
| KPI formulas | Same | Same (mirror) |
| Tier logic | Same | Same |
| Slip prediction | Same blend formula | Same |
| Charts | None | Built-in chart trên Dashboard |
| Color coding | None | Yes — tier + slip risk |
| UDFs (Excel formulas) | N/A | `=AssignTier(...)`, `=ComputeComposite(...)`, `=PredictSlip(...)` |
| Runtime | ~5 sec | ~3-5 sec |

## Sử dụng cho data thật

Thay vì click "Generate Sample":

1. Mở sheet `PO_History`
2. Paste data từ ERP/Excel của bạn vào row 2 trở đi
3. Đảm bảo cột match (xem header row 1)
4. Bấm "Compute Scores"

**Yêu cầu data:** 17 columns chính xác như header. Các giá trị `is_otd`, `is_in_full` phải là `TRUE`/`FALSE` (text hoặc Boolean).

## Troubleshooting

| Lỗi | Cách fix |
|---|---|
| "Macros disabled" khi mở file | Trust Center → Macro Settings → "Enable all macros" |
| "Compile error: Variable not defined" | Check imported đủ 5 .bas chưa. M5 reference M1-M4 |
| Chart không hiện trên Dashboard | Check sheet name "Dashboard" có tồn tại không (case-sensitive) |
| Composite score sai | Verify weights trong code (M0_Helpers) khớp với Parameters sheet |

## Roadmap mở rộng

- [ ] Đọc weights từ Parameters sheet (currently hardcoded in M0_Helpers — quick win cho user customize)
- [ ] Thêm sheet "Trend" hiển thị OTD over time per vendor
- [ ] Add Outlook integration để email scorecards weekly
- [ ] Survival analysis cho slip prediction (cần phần Python — quá phức tạp cho VBA)
