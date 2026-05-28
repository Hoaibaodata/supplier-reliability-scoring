"""
Reliability scoring engine.

Compute per-vendor KPIs and weighted composite score:
    - OTD (On-Time Delivery rate)
    - OTIF (On-Time In-Full rate)
    - σ_LT (lead time variability) — feeds back to MRP classical_z_full_var model
    - Quality slip rate
    - Communication score
    - Composite: weighted sum
"""
from __future__ import annotations

import numpy as np
import pandas as pd

# Composite score weights (sum to 1.0). Tunable based on business priority.
WEIGHTS = {
    "otif": 0.40,
    "lt_consistency": 0.20,
    "quality": 0.20,
    "communication": 0.10,
    "trade_relationship": 0.10,
}


def compute_vendor_kpis(po_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate per-vendor KPIs from PO history."""
    grouped = po_df.groupby(["vendor_id", "vendor_name"])

    kpi = grouped.agg(
        num_pos=("po_id", "count"),
        otd_rate=("is_otd", "mean"),
        in_full_rate=("is_in_full", "mean"),
        avg_lt_days=("actual_lt_days", "mean"),
        std_lt_days=("actual_lt_days", "std"),
        sigma_lead_time_weeks=("actual_lt_days", lambda s: s.std() / 7.0),
        avg_slip_days=("slip_days", "mean"),
        total_delivered_qty=("delivered_qty", "sum"),
        total_qc_fail_qty=("qc_fail_qty", "sum"),
        avg_confirm_lag=("confirm_lag_days", "mean"),
    ).reset_index()

    # OTIF = AND of OTD + In-Full per PO. Recompute from raw because mean of products != product of means
    otif_df = (
        po_df.assign(is_otif=lambda d: d["is_otd"] & d["is_in_full"])
        .groupby("vendor_id")["is_otif"]
        .mean()
        .reset_index()
        .rename(columns={"is_otif": "otif_rate"})
    )
    kpi = kpi.merge(otif_df, on="vendor_id", how="left")

    # Quality pass rate
    kpi["quality_pass_rate"] = 1 - (kpi["total_qc_fail_qty"] / kpi["total_delivered_qty"].replace(0, np.nan))
    kpi["quality_pass_rate"] = kpi["quality_pass_rate"].fillna(1.0)

    # LT consistency: lower std relative to mean = better. Normalize to 0-1 score (CV inverse).
    kpi["lt_cv"] = kpi["std_lt_days"] / kpi["avg_lt_days"].replace(0, np.nan)
    kpi["lt_cv"] = kpi["lt_cv"].fillna(0.0)
    kpi["lt_consistency_score"] = (1 - kpi["lt_cv"].clip(0, 1)).clip(0, 1)

    # Communication score: lower confirm_lag = better (normalize: 0 days = 1.0, 7+ days = 0)
    kpi["communication_score"] = (1 - (kpi["avg_confirm_lag"] / 7.0)).clip(0, 1)

    # Trade relationship: simple proxy = number of POs (more POs = more trust). Normalize per percentile.
    max_pos = kpi["num_pos"].max() if not kpi["num_pos"].empty else 1
    kpi["trade_relationship_score"] = (kpi["num_pos"] / max_pos).clip(0, 1)

    return kpi


def compute_composite_score(kpi_df: pd.DataFrame, weights: dict[str, float] | None = None) -> pd.DataFrame:
    """Compute weighted composite reliability score per vendor.

    Score 0-100. Higher is better.
    """
    if weights is None:
        weights = WEIGHTS

    score = (
        weights["otif"] * kpi_df["otif_rate"]
        + weights["lt_consistency"] * kpi_df["lt_consistency_score"]
        + weights["quality"] * kpi_df["quality_pass_rate"]
        + weights["communication"] * kpi_df["communication_score"]
        + weights["trade_relationship"] * kpi_df["trade_relationship_score"]
    )

    kpi_df = kpi_df.copy()
    kpi_df["composite_score"] = (score * 100).round(2)

    # Tier classification
    kpi_df["tier"] = pd.cut(
        kpi_df["composite_score"],
        bins=[0, 60, 75, 85, 100],
        labels=["D-Risk", "C-Watch", "B-Standard", "A-Preferred"],
    )

    return kpi_df.sort_values("composite_score", ascending=False).reset_index(drop=True)


def predict_slip_probability(po_df: pd.DataFrame, kpi_df: pd.DataFrame) -> pd.DataFrame:
    """Simple slip predictor — base rate per vendor + recent trend adjustment.

    For next PO from each vendor, probability of slip = historical slip rate
    adjusted by trend in last 90 days vs all-time.
    """
    rng_pos = po_df.sort_values("order_date")
    cutoff = rng_pos["order_date"].max() - pd.Timedelta(days=90)
    rng_pos["order_date_dt"] = pd.to_datetime(rng_pos["order_date"])

    recent_slip = (
        rng_pos[rng_pos["order_date_dt"] >= pd.to_datetime(cutoff)]
        .groupby("vendor_id")["is_otd"]
        .agg(["mean", "count"])
        .rename(columns={"mean": "recent_otd", "count": "recent_n"})
        .reset_index()
    )

    pred = kpi_df.merge(recent_slip, on="vendor_id", how="left")
    pred["recent_otd"] = pred["recent_otd"].fillna(pred["otd_rate"])
    pred["recent_n"] = pred["recent_n"].fillna(0)

    # Blend recent vs historical (more weight on recent if enough data)
    pred["blend_weight"] = (pred["recent_n"] / 20).clip(0, 0.7)  # cap at 70% recent
    pred["predicted_otd"] = (
        pred["blend_weight"] * pred["recent_otd"] + (1 - pred["blend_weight"]) * pred["otd_rate"]
    )
    pred["predicted_slip_probability"] = (1 - pred["predicted_otd"]).round(3)

    return pred[
        [
            "vendor_id",
            "vendor_name",
            "tier",
            "composite_score",
            "otd_rate",
            "recent_otd",
            "predicted_slip_probability",
            "sigma_lead_time_weeks",
        ]
    ].sort_values("predicted_slip_probability", ascending=False)
