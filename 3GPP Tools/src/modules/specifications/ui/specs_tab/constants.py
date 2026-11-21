"""Working group ordering and part-number regex constants for SpecificationsTab."""

from __future__ import annotations

import re

PART_NUM_REGEX = re.compile(r"\d{4,5}-(\d{1,2})(?:[-_.]|$)")

USEFUL_SA = ("SA", "SA1", "SA2", "SA3", "SA4", "SA5", "SA6")
USEFUL_RAN = ("RAN", "RAN1", "RAN2", "RAN3", "RAN4", "RAN5")
USEFUL_CT = ("CT", "CT1", "CT3", "CT4", "CT6")
USEFUL_WGS = set(USEFUL_SA + USEFUL_RAN + USEFUL_CT)

# Canonical ordering: Active groups always first in standard 3GPP hierarchy (SA -> RAN -> CT)
ACTIVE_WG_ORDER = list(USEFUL_SA) + list(USEFUL_RAN) + list(USEFUL_CT)

LEGACY_RAN = ("RAN6", "GERAN", "GERAN1", "GERAN2", "GERAN3")
LEGACY_CT = ("CN1", "CN3", "CN4", "CN5", "T2", "T3")
LEGACY_OTHER = ("SMG01", "SMG03", "AHG", "ITU")
LEGACY_WGS = set(LEGACY_RAN + LEGACY_CT + LEGACY_OTHER)

# Ordered full catalog: Active groups first, followed by legacy/closed groups grouped by TSG/domain
FULL_WG_ORDER = ACTIVE_WG_ORDER + list(LEGACY_RAN) + list(LEGACY_CT) + list(LEGACY_OTHER)
