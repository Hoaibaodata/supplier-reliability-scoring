# Supplier Reliability — Tech Ideas

## OTD vs OTIF distinction
- OTD = on-time delivery (correct date)
- OTIF = on-time IN-FULL (correct date AND correct quantity)
- Industry typical: OTD 90% but OTIF 75% — over-promising is common

## Survival analysis for slip
Standard regression predicts mean LT. But what we care about is "P(LT > requested)" — tail risk. Survival analysis (Kaplan-Meier, Cox proportional hazards) handles this naturally.

## Behavior patterns from data (insight from MRP-INSIGHTS-FROM-INDUSTRY.md)
- Some suppliers slip ONLY on low-value SKU (because high-value is monitored more)
- Some slip MORE during peak season (capacity constrained)
- Some slip MORE on first-time orders (no relationship)

## Reliability score formula (draft)
```
score = 0.4 * OTIF
      + 0.2 * (1 - sigma_LT/mean_LT)  // consistency
      + 0.2 * (1 - quality_slip_rate)
      + 0.1 * communication_score
      + 0.1 * relationship_score (years of trade)
```

## Feedback loop to MRP
Output `vendor_reliability.csv` with `sigma_lead_time_weeks` per vendor → consumed by `inventory-optimization-portfolio/run_mrp.py` for `classical_z_full_var` model.

## Visualization patterns
- **Scorecard** (one page per supplier): radar chart + trend lines
- **Heatmap** SKU × supplier showing slip frequency
- **Alert table**: top 10 suppliers at risk for next PO cycle
