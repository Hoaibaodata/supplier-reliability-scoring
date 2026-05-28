# Supplier Reliability Scoring — VBA Version Setup

> Excel-only edition. KHÔNG cần Python để chạy, chỉ cần Excel.

## 🚀 Setup nhanh nhất (1 lệnh) — Auto-import VBA

Yêu cầu:
- Excel installed
- Python + xlwings (`pip install xlwings`)
- **Trust Center setting (1 lần):** Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → tick **"Trust access to the VBA project object model"**

Rồi chạy:

```powershell
cd "D:\HOC_TAP\supplier-reliability-scoring\vba_version"
python generate_starter_xlsx.py     # tạo SupplierReliability.xlsx (sheets + headers)
python build_xlsm_with_vba.py       # convert -> .xlsm + import 6 VBA modules
```

→ Ra ngay file `SupplierReliability.xlsm` **đầy đủ VBA**, không cần Alt+F11.

Sau đó chỉ cần:
1. Mở `.xlsm` trong Excel
2. Tạo 5-6 form buttons trên sheet Dashboard (xem **Bước 4** dưới)
3. Bấm "Run Full Analysis"

---

## Setup truyền thống (manual import — nếu xlwings không xài được)

### Bước 1 — Generate xlsx
```powershell
python generate_starter_xlsx.py
```

### Bước 2 — Save As .xlsm
Mở `SupplierReliability.xlsx` → File → Save As → **"Excel Macro-Enabled Workbook (.xlsm)"** → tên `SupplierReliability.xlsm`.

### Bước 3 — Import VBA modules thủ công
**Alt + F11** → File → Import File → chọn lần lượt 6 file trong `vba_modules/`:
- `M0_Helpers.bas`
- `M1_KPICalculator.bas`
- `M2_SlipPredictor.bas`
- `M3_ChartBuilder.bas`
- `M4_SampleData.bas`
- `M5_Dashboard.bas`

File → Save trong VBE → quay lại Excel → Save.

### Bước 4 — Tạo nút bấm trên Dashboard

Developer tab → Insert → Button (Form Control), gán macro:

| Caption | Macro |
|---|---|
| 🚀 **Run Full Analysis** | `Action_RunFullAnalysis` |
| 📥 Generate Sample PO | `Action_GenerateSample` |
| 📊 Compute Scores | `Action_ComputeScores` |
| 🎯 Predict Slip Risk | `Action_PredictSlip` |
| 📈 Build Chart | `Action_BuildChart` |
| 🔄 Reset Outputs | `Action_ResetOutputs` |

### Bước 5 — Chạy thử

Bấm **Run Full Analysis** → kết quả popup. Verify:

- Sheet **PO_History**: 200 dòng PO sample
- Sheet **Scorecards**: 10 vendors với composite score 0-100, tier color-coded
- Sheet **Predictions**: 10 vendors sắp xếp theo slip probability
- Sheet **Dashboard**: chart bar với màu theo tier
- Sheet **Log**: track lại các action

---

## VBA patterns demonstrated

| Module | Patterns |
|---|---|
| `M0_Helpers` | Constants, helper functions, sheet creation, header formatting, conditional formatting setup |
| `M1_KPICalculator` | Scripting.Dictionary cho aggregation, iterate rows, std dev calc, UDF (`AssignTier`, `ComputeComposite`) |
| `M2_SlipPredictor` | Date arithmetic, second-pass aggregation, conditional formatting based on values, UDF (`PredictSlip`) |
| `M3_ChartBuilder` | Chart objects creation, color customization per data point, ChartObjects collection |
| `M4_SampleData` | Random number generation, jagged arrays for profiles, normal distribution approximation |
| `M5_Dashboard` | Action routing, Application.ScreenUpdating optimization, message box workflow |

## UDFs callable as Excel formulas

```excel
=AssignTier(85)            → "A-Preferred"
=ComputeComposite(0.9, 0.85, 0.95, 0.8, 0.7)
=PredictSlip(0.85, 0.92, 15)
```

## So với Python version

| Feature | Python | VBA |
|---|---|---|
| PO history | 500 | 200 |
| Vendors | 10 | 10 (cùng profile) |
| Formulas | Same | Same |
| Charts | None | Built-in chart trên Dashboard |
| UDFs | N/A | 3 UDFs callable from cells |
| Runtime | ~5 sec | ~3-5 sec |

## Troubleshooting

| Lỗi | Cách fix |
|---|---|
| "Macros disabled" | Trust Center → Macro Settings → "Enable all macros" |
| "Programmatic access to VBA project not trusted" (when running build_xlsm_with_vba.py) | Trust Center → Macro Settings → tick **"Trust access to the VBA project object model"** |
| "Compile error" | Re-import đủ 6 .bas modules. M5 reference M1-M4 |
| Chart không hiện | Sheet "Dashboard" phải tồn tại (case-sensitive) |
| Composite score sai | Verify weights trong M0_Helpers |
