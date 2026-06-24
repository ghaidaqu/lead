"""Pricing snapshot + shared constants for the sync pipeline.

The old price-sheet profit engine (compute_profit / resolve_customer /
resolve_platform / merchant overrides) has been **removed** — profit is now
computed per shipment from Lead's own actuals (invoice Base Cost where billed,
live carrier prices from shipping-companies.php for the current cycle), in
``sync_from_lead.shipment_record``.

What remains here is the small shared infrastructure that path still needs:
the carrier-name alias map, the realized-status helper, the money coercion, and
the ``PricingSnapshot`` that carries the live carrier prices + a few operational
constants (included weight, extra-kg rate) loaded from the DB.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

# Long carrier labels on shipment rows / the Lead site -> canonical `carriers`
# table keys.
CARRIER_ALIASES = {
    "ارامكس - ARAMEX": "ارامكس",
    "aramex ( استلام من الفرع )": "ارامكس استلام",
    "SMSA - سمسا": "سمسا",
    "SMSA ( استلام من الفرع )": "سمسا استلام",
    "RedBox - ريدبوكس": "ريد بوكس",
    "JT Express": "JT Express",
}

# Statuses excluded from profit (cancelled / draft). The realized basis also
# excludes returns — see shipment_record.
EXCLUDED_STATUSES = ("ملغي", "مسودة")
WEIGHT_INCLUDED_KG = 15.0


def money(value: Any) -> float:
    """Coerce a cell/value to float: None/""/"-" -> 0.0; else first signed number."""
    if value in (None, "", "-"):
        return 0.0
    match = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(match.group()) if match else 0.0


@dataclass(frozen=True)
class PricingSnapshot:
    """Live carrier prices (from shipping-companies.php) + the few operational
    constants the current-cycle cost computation needs (included weight, extra-kg
    rate). Loaded from the DB via :meth:`from_db`."""
    carriers: dict[str, dict[str, Any]] = field(default_factory=dict)
    extra_platform_gross: float = 0.0
    weight_included_kg: float = WEIGHT_INCLUDED_KG

    @classmethod
    def from_db(cls, snapshot: dict[str, Any]) -> "PricingSnapshot":
        s = snapshot.get("settings", {})
        weight = s.get("weight_included_kg")
        return cls(
            carriers=snapshot.get("carriers", {}),
            extra_platform_gross=money(s.get("extra_platform_gross")),
            weight_included_kg=float(weight) if weight is not None else WEIGHT_INCLUDED_KG,
        )
