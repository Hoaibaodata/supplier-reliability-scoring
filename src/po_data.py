"""
PO history data model + synthetic generator.

Synthetic data simulates 12 months of PO history across multiple vendors
with realistic patterns: some vendors slip more than others, some categories
have higher quality issues, etc.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

import numpy as np
import pandas as pd

DEFAULT_SEED = 42
NUM_PO = 500
NUM_VENDORS = 10
NUM_CATEGORIES = 4

# Per-vendor reliability profile (realistic — some vendors much worse than others)
VENDOR_PROFILES = {
    "V001": {"name": "Premium Plus Inc.", "otd_base": 0.95, "qc_base": 0.99, "lt_mean": 14, "lt_std": 1},
    "V002": {"name": "Reliable Co.", "otd_base": 0.92, "qc_base": 0.98, "lt_mean": 18, "lt_std": 2},
    "V003": {"name": "Standard Supply", "otd_base": 0.85, "qc_base": 0.95, "lt_mean": 14, "lt_std": 3},
    "V004": {"name": "Mid-Tier Logistics", "otd_base": 0.82, "qc_base": 0.93, "lt_mean": 16, "lt_std": 4},
    "V005": {"name": "Generic Parts Ltd.", "otd_base": 0.78, "qc_base": 0.90, "lt_mean": 21, "lt_std": 5},
    "V006": {"name": "Fast Ship Corp.", "otd_base": 0.90, "qc_base": 0.97, "lt_mean": 7, "lt_std": 2},
    "V007": {"name": "Budget Source", "otd_base": 0.70, "qc_base": 0.85, "lt_mean": 25, "lt_std": 7},
    "V008": {"name": "Global Exports", "otd_base": 0.88, "qc_base": 0.96, "lt_mean": 20, "lt_std": 3},
    "V009": {"name": "Risky Supplier", "otd_base": 0.60, "qc_base": 0.80, "lt_mean": 28, "lt_std": 10},
    "V010": {"name": "New Vendor", "otd_base": 0.83, "qc_base": 0.92, "lt_mean": 15, "lt_std": 4},
}

CATEGORIES = ["Electronics", "Mechanical", "Plastics", "Chemicals"]


def generate_po_history(
    num_po: int = NUM_PO,
    start_date: date = date(2025, 6, 1),
    seed: int = DEFAULT_SEED,
) -> pd.DataFrame:
    """Generate synthetic 12-month PO history across vendors and categories."""
    rng = np.random.default_rng(seed)
    vendors = list(VENDOR_PROFILES.keys())

    rows = []
    for po_idx in range(1, num_po + 1):
        vendor_id = rng.choice(vendors)
        profile = VENDOR_PROFILES[vendor_id]
        category = rng.choice(CATEGORIES)

        # Order date: uniform over past 12 months
        order_date = start_date + timedelta(days=int(rng.integers(0, 365)))
        requested_lt = int(rng.normal(profile["lt_mean"], profile["lt_std"] / 2))
        requested_lt = max(requested_lt, 1)
        requested_delivery = order_date + timedelta(days=requested_lt)

        # Actual delivery: usually close to requested but with slip
        otd_prob = profile["otd_base"]
        # Category modifier: Chemicals trickier, Mechanical easier
        if category == "Chemicals":
            otd_prob *= 0.95
        elif category == "Mechanical":
            otd_prob = min(otd_prob * 1.02, 0.99)

        is_otd = rng.random() < otd_prob
        if is_otd:
            slip_days = 0
        else:
            slip_days = int(rng.integers(1, 14))
        actual_delivery = requested_delivery + timedelta(days=slip_days)
        actual_lt = (actual_delivery - order_date).days

        # Quantity
        requested_qty = int(rng.integers(500, 10000))
        # In-full: 95% of OTD POs deliver full qty
        delivered_qty = requested_qty if rng.random() < 0.92 else int(requested_qty * rng.uniform(0.85, 0.99))

        # Quality at IQC
        qc_prob = profile["qc_base"]
        qc_pass_rate = max(0.5, min(1.0, rng.normal(qc_prob, 0.03)))
        qc_pass_qty = int(delivered_qty * qc_pass_rate)
        qc_fail_qty = delivered_qty - qc_pass_qty

        # Communication
        confirm_lag = int(rng.exponential(2.0)) if rng.random() < 0.95 else int(rng.exponential(7.0))
        slip_notification_lag = (
            int(rng.exponential(1.5))
            if slip_days > 0 and rng.random() < 0.7
            else (slip_days if slip_days > 0 else None)
        )

        rows.append(
            {
                "po_id": f"PO{po_idx:05d}",
                "vendor_id": vendor_id,
                "vendor_name": profile["name"],
                "category": category,
                "order_date": order_date,
                "requested_delivery_date": requested_delivery,
                "requested_qty": requested_qty,
                "actual_delivery_date": actual_delivery,
                "delivered_qty": delivered_qty,
                "qc_pass_qty": qc_pass_qty,
                "qc_fail_qty": qc_fail_qty,
                "requested_lt_days": requested_lt,
                "actual_lt_days": actual_lt,
                "slip_days": slip_days,
                "is_otd": is_otd,
                "is_in_full": delivered_qty >= requested_qty,
                "confirm_lag_days": confirm_lag,
                "slip_notification_lag_days": slip_notification_lag,
            }
        )

    return pd.DataFrame(rows).sort_values("order_date").reset_index(drop=True)
