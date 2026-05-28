# Supplier Reliability Scoring

Score and rank suppliers on **OTD, OTIF, lead-time consistency, quality, communication, and trade history** to drive better procurement decisions.

> Companion to [inventory-optimization-portfolio](https://github.com/Hoaibaodata/inventory-optimization-portfolio) — `sigma_lead_time_weeks` output feeds back into MRP `classical_z_full_var` safety stock formula.

## Why this matters

Material Control specialist at a US manufacturing plant once stopped his production line 3.5 hours due to a supplier slip — leading him to permanently over-order at **5-6 WOS instead of 2 WOS**. That's hundreds of thousands of dollars in unnecessary carrying cost per year, driven by *behavioral risk aversion*, not formula.

This tool quantifies supplier reliability with hard numbers, so planners can:
- **Trust** A-Preferred vendors → reduce safety stock → cut carrying cost
- **Hedge** C-Watch vendors with dual-source / penalty clauses
- **Predict** next-PO slip probability and pre-emptively expedite alternatives

## Output (verified end-to-end)

5 KPIs per vendor + composite 0-100 score + tier:

```
vendor_id  vendor_name         num_pos  otif_rate  quality  sigma_LT(w)  score  tier
V001       Premium Plus Inc.   49       91.8%      98.5%    0.21         89.84  A-Preferred
V002       Reliable Co.        61       80.3%      98.2%    0.40         86.15  A-Preferred
V008       Global Exports      50       84.0%      96.2%    0.32         86.07  A-Preferred
V003       Standard Supply     54       75.9%      95.1%    0.40         81.53  B-Standard
V004       Mid-Tier Logistics  49       75.5%      92.8%    0.51         78.98  B-Standard
V006       Fast Ship Corp.     45       77.8%      96.8%    0.44         77.34  B-Standard
V005       Generic Parts Ltd.  44       68.2%      90.0%    0.57         76.24  B-Standard
V010       New Vendor          56       73.2%      91.7%    0.50         80.28  B-Standard
V007       Budget Source       42       61.9%      85.3%    0.70         72.72  C-Watch
V009       Risky Supplier      50       52.0%      80.3%    1.04         68.53  C-Watch
```

Slip prediction for next PO (recent 90-day trend blended with historical):

```
Risky Supplier predicted slip probability: 49.2%
Budget Source                              38.3%
Premium Plus                                5.6%
```

## What the scorer does

### KPI calculations (per vendor over 12-month PO history)

- **OTD rate** = % POs delivered on/before requested date
- **OTIF rate** = % POs OTD **and** delivered_qty ≥ requested_qty
- **σ_lead_time_weeks** = standard deviation of actual_lt_days / 7 → feeds MRP classical_z_full_var
- **Quality pass rate** = 1 − (qc_fail / delivered) aggregated
- **Communication score** = 1 − (avg_confirm_lag_days / 7), clipped to [0, 1]
- **Trade relationship score** = num_pos / max(num_pos) normalized

### Composite formula (weighted)

```
score = 0.40 × otif_rate
      + 0.20 × lt_consistency
      + 0.20 × quality_pass_rate
      + 0.10 × communication_score
      + 0.10 × trade_relationship_score
```

Weights tunable in `src/reliability_scorer.py`.

### Tiering

| Score | Tier | Action |
|---|---|---|
| 85-100 | A-Preferred | Strategic partner, expand share |
| 75-85 | B-Standard | Normal terms |
| 60-75 | C-Watch | Dual-source, penalty clauses |
| <60 | D-Risk | Avoid, exit relationship |

### Slip prediction (recent-trend blend)

```
recent_window = 90 days
blend_weight = min(recent_po_count / 20, 0.7)
predicted_otd = blend × recent_otd + (1 − blend) × historical_otd
slip_prob = 1 − predicted_otd
```

Conservative — defaults to historical when recent sample size is small.

## Setup

```powershell
cd "D:\HOC_TAP\supplier-reliability-scoring"
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python run_supplier_scoring.py
```

Runs in ~5 seconds. No external data needed — uses synthetic PO history.

## Synthetic data note

`src/po_data.py` generates 500 POs across 10 vendors with **realistic reliability profiles** baked in (some vendors much worse than others — see `VENDOR_PROFILES` dict). Categories: Electronics, Mechanical, Plastics, Chemicals. Each has slight quality/LT modifiers.

For real deployment, replace `generate_po_history()` with your ERP query (typical schema in `src/po_data.py:24`).

## File structure

```
supplier-reliability-scoring/
├── data/
│   ├── raw/po_history.csv      (gitignored, regenerated)
│   ├── processed/
│   └── ...
├── exports/
│   ├── supplier_scorecards.csv
│   └── slip_risk_predictions.csv
├── src/
│   ├── po_data.py              # PO history synthetic generator
│   └── reliability_scorer.py   # KPI + composite + slip predictor
├── docs/
│   ├── ROADMAP.md
│   └── IDEAS.md
├── requirements.txt
└── run_supplier_scoring.py
```

## Roadmap

See [docs/ROADMAP.md](docs/ROADMAP.md) — Phase 1 (current) covers basic scoring. Phase 2+: survival analysis for slip prediction, behavioral patterns by category, dashboard.

## License

MIT
