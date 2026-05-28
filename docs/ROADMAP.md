# Supplier Reliability Scoring — Roadmap

## Phase 1: Data model
- PO master: po_id, vendor_id, factory_id, requested_qty, requested_delivery_date
- PO delivery: actual_delivery_date, delivered_qty, qc_pass_qty, qc_fail_qty
- Communication log: confirm_lag_days, slip_notification_lag

## Phase 2: KPI calculations
- OTD = % POs delivered on or before requested_delivery_date
- OTIF = % POs OTD AND delivered_qty >= requested_qty
- LT statistics: avg, std, P95
- Quality slip rate: 1 - (qc_pass_qty / delivered_qty)

## Phase 3: Scoring engine
- Weighted score per vendor (configurable weights)
- Trend over time (rolling 90/180 days)
- Category-specific breakdowns

## Phase 4: Predictive
- Predict slip probability for next PO
- Features: historical OTD, season, quantity vs typical, communication patterns
- Output feeds into MRP `sigma_lead_time_weeks` parameter

## Phase 5: Dashboard
- Supplier scorecard (one page per supplier)
- Comparison matrix
- Alert when slip prob > threshold
