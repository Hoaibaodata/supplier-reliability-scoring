"""Entry point — generate PO history, score suppliers, predict slip risk."""
from __future__ import annotations

from pathlib import Path

from src.po_data import generate_po_history
from src.reliability_scorer import (
    compute_composite_score,
    compute_vendor_kpis,
    predict_slip_probability,
)


def main() -> None:
    project_root = Path(__file__).resolve().parent
    raw_dir = project_root / "data" / "raw"
    processed_dir = project_root / "data" / "processed"
    export_dir = project_root / "exports"
    for d in (raw_dir, processed_dir, export_dir):
        d.mkdir(parents=True, exist_ok=True)

    print("Generating synthetic PO history (500 POs across 10 vendors, 12 months)...")
    po_df = generate_po_history()
    po_path = raw_dir / "po_history.csv"
    po_df.to_csv(po_path, index=False)
    print(f"  Saved {len(po_df)} PO records: {po_path}")

    print("\nComputing vendor KPIs...")
    kpi_df = compute_vendor_kpis(po_df)
    print(f"  KPIs for {len(kpi_df)} vendors")

    print("\nComputing composite reliability scores...")
    scored = compute_composite_score(kpi_df)
    scored_path = export_dir / "supplier_scorecards.csv"
    scored.to_csv(scored_path, index=False)
    print(f"  Scorecards saved: {scored_path}")

    print("\nPredicting slip risk for next PO cycle...")
    predictions = predict_slip_probability(po_df, scored)
    pred_path = export_dir / "slip_risk_predictions.csv"
    predictions.to_csv(pred_path, index=False)
    print(f"  Predictions saved: {pred_path}")

    print("\n" + "=" * 80)
    print("SUPPLIER SCORECARD SUMMARY")
    print("=" * 80)
    print(
        scored[
            [
                "vendor_id",
                "vendor_name",
                "num_pos",
                "otif_rate",
                "quality_pass_rate",
                "sigma_lead_time_weeks",
                "composite_score",
                "tier",
            ]
        ].to_string(index=False)
    )
    print("=" * 80)

    print("\n" + "=" * 80)
    print("SLIP RISK PREDICTIONS (sorted by risk descending)")
    print("=" * 80)
    print(predictions.to_string(index=False))
    print("=" * 80)


if __name__ == "__main__":
    main()
