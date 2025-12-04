"""
models.py

Data models and calculation logic for the Cable Tray Calculator.

Defines:
- CableType: basic mechanical properties of a cable.
- TrayType: basic mechanical and structural properties of a tray.
- compute_cable_tray_stats: core function to compute loading & fill metrics.

All units:
- Length: mm (for geometry), m (for weight and final results)
- Weight: kg/m
- Area: mm^2
"""

from __future__ import annotations

from dataclasses import dataclass
from math import pi
from typing import Dict, List, Tuple


@dataclass
class CableType:
    """
    Represents a cable type.

    Attributes:
        name: Human-readable name (e.g. "Cu 3C 2.5mm² PVC").
        diameter_mm: Approximate outer diameter in mm.
        weight_kg_per_m: Cable weight per metre (kg/m).
    """
    name: str
    diameter_mm: float
    weight_kg_per_m: float


@dataclass
class TrayType:
    """
    Represents a cable tray type and its structural capacity.

    Attributes:
        name: Human-readable name for display.
        width_mm: Internal usable tray width in mm.
        height_mm: Internal usable tray side height in mm (for fill height).
        max_load_kg_per_m: Maximum allowable uniformly distributed load (kg/m).
        self_weight_kg_per_m: Self-weight of the tray per metre (kg/m).
        max_fill_ratio: Maximum area fill ratio (0-1) recommended for cables.
    """
    name: str
    width_mm: float
    height_mm: float
    max_load_kg_per_m: float
    self_weight_kg_per_m: float
    max_fill_ratio: float = 0.6  # 60% recommended cable fill by default





def get_default_cables() -> List[CableType]:
    """
    Return an extended library of typical LV power, control, data,
    fibre and coax cables.

    NOTE:
        - Values are approximate / averaged from typical manufacturer data.
        - Always verify against actual cable datasheets for final design.
        - Units: diameter in mm, weight in kg/m.
    """
    return [
        # -------------------------
        # LV single-core PVC 450/750 V (Cu)
        # -------------------------
        CableType("Cu 1C 1.5mm² PVC", diameter_mm=5.0, weight_kg_per_m=0.036),
        CableType("Cu 1C 2.5mm² PVC", diameter_mm=5.5, weight_kg_per_m=0.055),
        CableType("Cu 1C 4mm² PVC", diameter_mm=6.0, weight_kg_per_m=0.075),
        CableType("Cu 1C 6mm² PVC", diameter_mm=6.8, weight_kg_per_m=0.110),
        CableType("Cu 1C 10mm² PVC", diameter_mm=8.0, weight_kg_per_m=0.170),
        CableType("Cu 1C 16mm² PVC", diameter_mm=9.5, weight_kg_per_m=0.260),
        CableType("Cu 1C 25mm² PVC", diameter_mm=11.5, weight_kg_per_m=0.380),
        CableType("Cu 1C 35mm² PVC", diameter_mm=13.0, weight_kg_per_m=0.520),
        CableType("Cu 1C 50mm² PVC", diameter_mm=15.0, weight_kg_per_m=0.720),
        CableType("Cu 1C 70mm² PVC", diameter_mm=17.0, weight_kg_per_m=1.000),
        CableType("Cu 1C 95mm² PVC", diameter_mm=19.5, weight_kg_per_m=1.290),
        CableType("Cu 1C 120mm² PVC", diameter_mm=21.0, weight_kg_per_m=1.540),

        # -------------------------
        # LV multi-core PVC (typical building power) – 3C
        # -------------------------
        CableType("Cu 3C 1.5mm² PVC", diameter_mm=9.5, weight_kg_per_m=0.22),
        CableType("Cu 3C 2.5mm² PVC", diameter_mm=11.0, weight_kg_per_m=0.30),
        CableType("Cu 3C 4mm² PVC", diameter_mm=13.0, weight_kg_per_m=0.45),
        CableType("Cu 3C 6mm² PVC", diameter_mm=15.0, weight_kg_per_m=0.65),
        CableType("Cu 3C 10mm² PVC", diameter_mm=18.0, weight_kg_per_m=1.05),
        CableType("Cu 3C 16mm² PVC", diameter_mm=21.0, weight_kg_per_m=1.55),
        CableType("Cu 3C 25mm² PVC", diameter_mm=25.0, weight_kg_per_m=2.40),
        CableType("Cu 3C 35mm² PVC", diameter_mm=28.0, weight_kg_per_m=3.20),
        CableType("Cu 3C 50mm² PVC", diameter_mm=32.0, weight_kg_per_m=4.40),
        CableType("Cu 3C 70mm² PVC", diameter_mm=36.0, weight_kg_per_m=5.90),
        CableType("Cu 3C 95mm² PVC", diameter_mm=42.0, weight_kg_per_m=7.70),

        # -------------------------
        # LV multi-core PVC – 4C
        # -------------------------
        CableType("Cu 4C 1.5mm² PVC", diameter_mm=10.0, weight_kg_per_m=0.25),
        CableType("Cu 4C 2.5mm² PVC", diameter_mm=12.0, weight_kg_per_m=0.34),
        CableType("Cu 4C 4mm² PVC", diameter_mm=14.5, weight_kg_per_m=0.52),
        CableType("Cu 4C 6mm² PVC", diameter_mm=16.5, weight_kg_per_m=0.75),
        CableType("Cu 4C 10mm² PVC", diameter_mm=20.0, weight_kg_per_m=1.20),
        CableType("Cu 4C 16mm² PVC", diameter_mm=23.0, weight_kg_per_m=1.80),
        CableType("Cu 4C 25mm² PVC", diameter_mm=27.0, weight_kg_per_m=2.70),
        CableType("Cu 4C 35mm² PVC", diameter_mm=30.0, weight_kg_per_m=3.60),

        # -------------------------
        # LV multi-core PVC – 5C (ADDED)
        # -------------------------
        CableType("Cu 5C 1.5mm² PVC", diameter_mm=11.0, weight_kg_per_m=0.29),
        CableType("Cu 5C 2.5mm² PVC", diameter_mm=13.0, weight_kg_per_m=0.40),
        CableType("Cu 5C 4mm² PVC", diameter_mm=15.5, weight_kg_per_m=0.60),
        CableType("Cu 5C 6mm² PVC", diameter_mm=17.5, weight_kg_per_m=0.85),
        CableType("Cu 5C 10mm² PVC", diameter_mm=21.0, weight_kg_per_m=1.35),
        CableType("Cu 5C 16mm² PVC", diameter_mm=24.0, weight_kg_per_m=2.00),

        # -------------------------
        # XLPE/SWA/PVC & XLPE/SWA/LSZH mains
        # (armoured LV power) – 4C
        # -------------------------
        CableType("Cu 4C 10mm² XLPE/SWA/PVC", diameter_mm=24.0, weight_kg_per_m=1.85),
        CableType("Cu 4C 16mm² XLPE/SWA/PVC", diameter_mm=27.0, weight_kg_per_m=2.45),
        CableType("Cu 4C 25mm² XLPE/SWA/PVC", diameter_mm=31.0, weight_kg_per_m=3.35),
        CableType("Cu 4C 35mm² XLPE/SWA/PVC", diameter_mm=35.0, weight_kg_per_m=4.50),
        CableType("Cu 4C 50mm² XLPE/SWA/PVC", diameter_mm=38.0, weight_kg_per_m=5.60),
        CableType("Cu 4C 70mm² XLPE/SWA/PVC", diameter_mm=43.0, weight_kg_per_m=7.50),
        CableType("Cu 4C 95mm² XLPE/SWA/PVC", diameter_mm=48.0, weight_kg_per_m=9.60),
        CableType("Cu 4C 120mm² XLPE/SWA/PVC", diameter_mm=52.0, weight_kg_per_m=11.7),

        # -------------------------
        # XLPE/SWA/PVC – 5C (ADDED)
        # -------------------------
        CableType("Cu 5C 0.75mm² XLPE/SWA/PVC", diameter_mm=17.0, weight_kg_per_m=1.55),
        CableType("Cu 5C 1.0mm² XLPE/SWA/PVC", diameter_mm=18.0, weight_kg_per_m=1.75),
        CableType("Cu 5C 1.5mm² XLPE/SWA/PVC", diameter_mm=19.5, weight_kg_per_m=2.05),
        CableType("Cu 5C 2.5mm² XLPE/SWA/PVC", diameter_mm=21.5, weight_kg_per_m=2.55),
        CableType("Cu 5C 4mm² XLPE/SWA/PVC", diameter_mm=26.0, weight_kg_per_m=2.90),
        CableType("Cu 5C 6mm² XLPE/SWA/PVC", diameter_mm=28.0, weight_kg_per_m=3.60),
        CableType("Cu 5C 10mm² XLPE/SWA/PVC", diameter_mm=32.0, weight_kg_per_m=4.80),
        CableType("Cu 5C 16mm² XLPE/SWA/PVC", diameter_mm=36.0, weight_kg_per_m=6.40),
        CableType("Cu 5C 25mm² XLPE/SWA/PVC", diameter_mm=41.0, weight_kg_per_m=8.60),

        # -------------------------
        # Control / I/O cables
        # -------------------------
        CableType("Control 7C 1.5mm² PVC", diameter_mm=13.5, weight_kg_per_m=0.33),
        CableType("Control 12C 1.5mm² PVC", diameter_mm=17.5, weight_kg_per_m=0.52),
        CableType("Control 24C 1.5mm² PVC", diameter_mm=23.0, weight_kg_per_m=0.95),

        CableType("Control 7C 2.5mm² PVC", diameter_mm=15.5, weight_kg_per_m=0.48),
        CableType("Control 12C 2.5mm² PVC", diameter_mm=20.0, weight_kg_per_m=0.78),

        # -------------------------
        # Instrumentation / twisted pair
        # -------------------------
        CableType("Instr 2x2x0.75mm² overall screen", diameter_mm=9.0, weight_kg_per_m=0.12),
        CableType("Instr 4x2x0.75mm² overall screen", diameter_mm=11.5, weight_kg_per_m=0.19),
        CableType("Instr 8x2x0.75mm² overall screen", diameter_mm=15.0, weight_kg_per_m=0.32),

        # -------------------------
        # Data cables – copper
        # -------------------------
        CableType("CAT5e U/UTP", diameter_mm=5.3, weight_kg_per_m=0.030),
        CableType("CAT5e F/UTP", diameter_mm=5.8, weight_kg_per_m=0.035),
        CableType("CAT6 U/UTP (indoor)", diameter_mm=6.1, weight_kg_per_m=0.040),
        CableType("CAT6 F/UTP", diameter_mm=6.5, weight_kg_per_m=0.045),
        CableType("CAT6A F/UTP", diameter_mm=7.6, weight_kg_per_m=0.055),
        CableType("CAT7 S/FTP", diameter_mm=8.2, weight_kg_per_m=0.065),

        # -------------------------
        # Fibre optics
        # -------------------------
        CableType("Fibre 4C tight-buffer indoor", diameter_mm=6.0, weight_kg_per_m=0.030),
        CableType("Fibre 12C loose-tube indoor", diameter_mm=8.0, weight_kg_per_m=0.045),
        CableType("Fibre 24C loose-tube indoor", diameter_mm=10.5, weight_kg_per_m=0.065),
        CableType("Fibre 48C loose-tube indoor", diameter_mm=13.0, weight_kg_per_m=0.090),

        # -------------------------
        # Coax
        # -------------------------
        CableType("RG59/U coax", diameter_mm=6.1, weight_kg_per_m=0.040),
        CableType("RG6/U coax", diameter_mm=6.9, weight_kg_per_m=0.055),
        CableType("RG11/U coax", diameter_mm=10.5, weight_kg_per_m=0.090),

        # -------------------------
        # Misc – flexible power leads
        # -------------------------
        CableType("H07RN-F 3G1.5mm²", diameter_mm=11.3, weight_kg_per_m=0.20),
        CableType("H07RN-F 3G2.5mm²", diameter_mm=12.2, weight_kg_per_m=0.26),
        CableType("H07RN-F 3G4mm²", diameter_mm=13.5, weight_kg_per_m=0.36),
        CableType("H07RN-F 5G2.5mm²", diameter_mm=14.5, weight_kg_per_m=0.36),
    ]







def get_default_trays() -> List[TrayType]:
    """
    Return an extended library of typical tray / ladder / basket sizes.

    Notes:
        - Width & height are internal usable dimensions (mm).
        - max_load_kg_per_m is an approximate uniformly distributed load
          at typical support spacing (e.g. 1.5–2.0 m).
        - self_weight_kg_per_m is approximate tray self-weight per metre.

    These are generic, blended values intended for quick planning /
    'sanity check' work. For real design, always replace with the actual
    manufacturer load tables (NEMA, IEC 61537, etc.). :contentReference[oaicite:1]{index=1}
    """
    trays: List[TrayType] = []

    # Helper to reduce repetition
    def add_tray(
        name: str,
        width: float,
        height: float,
        max_load: float,
        self_weight: float,
        fill_ratio: float,
    ) -> None:
        trays.append(
            TrayType(
                name=name,
                width_mm=width,
                height_mm=height,
                max_load_kg_per_m=max_load,
                self_weight_kg_per_m=self_weight,
                max_fill_ratio=fill_ratio,
            )
        )

    # --------------------------------------------------------
    # Heavy-duty cable ladders (NEMA 2 / 3-ish behaviour)
    # Typical 100–200+ kg/m @ 4 m span depending on width. :contentReference[oaicite:2]{index=2}
    # --------------------------------------------------------
    for width, max_load, self_w in [
        (150, 90.0, 4.5),
        (200, 110.0, 5.0),
        (300, 140.0, 6.0),
        (450, 170.0, 7.5),
        (600, 200.0, 9.0),
        (750, 220.0, 10.5),
        (900, 240.0, 12.0),
    ]:
        add_tray(
            name=f"Ladder HDG heavy {width} x 100",
            width=width,
            height=100.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.6,  # 60% recommended fill for big power :contentReference[oaicite:3]{index=3}
        )

    # --------------------------------------------------------
    # Medium-duty ladders
    # --------------------------------------------------------
    for width, max_load, self_w in [
        (150, 60.0, 3.8),
        (200, 75.0, 4.2),
        (300, 100.0, 5.0),
        (450, 125.0, 6.0),
        (600, 150.0, 7.2),
        (750, 170.0, 8.5),
        (900, 190.0, 9.8),
    ]:
        add_tray(
            name=f"Ladder HDG medium {width} x 75",
            width=width,
            height=75.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.55,
        )

    # --------------------------------------------------------
    # Light-duty ladders (shorter spans, LV/ELV)
    # --------------------------------------------------------
    for width, max_load, self_w in [
        (100, 40.0, 2.8),
        (150, 50.0, 3.2),
        (200, 60.0, 3.6),
        (300, 80.0, 4.2),
        (450, 100.0, 5.2),
        (600, 115.0, 6.0),
    ]:
        add_tray(
            name=f"Ladder HDG light {width} x 60",
            width=width,
            height=60.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.55,
        )

    # --------------------------------------------------------
    # Perforated trays (ventilated)
    # Typically lower structural capacity and height.
    # --------------------------------------------------------
    for width, max_load, self_w in [
        (100, 25.0, 2.2),
        (150, 30.0, 2.6),
        (200, 35.0, 3.0),
        (300, 45.0, 3.8),
        (450, 55.0, 4.6),
        (600, 65.0, 5.3),
    ]:
        add_tray(
            name=f"Perforated tray 50H {width} wide",
            width=width,
            height=50.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.5,  # 50% recommended fill
        )

    # Shallow perforated (for control / lighting)
    for width, max_load, self_w in [
        (100, 20.0, 1.9),
        (150, 25.0, 2.2),
        (200, 30.0, 2.5),
        (300, 40.0, 3.1),
    ]:
        add_tray(
            name=f"Perforated tray 35H {width} wide",
            width=width,
            height=35.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.45,
        )

    # --------------------------------------------------------
    # Wire mesh / basket trays
    # Typically used for data/IT – higher fill limits but
    # structurally lighter. :contentReference[oaicite:4]{index=4}
    # --------------------------------------------------------
    for width, max_load, self_w in [
        (100, 15.0, 1.2),
        (150, 18.0, 1.4),
        (200, 20.0, 1.6),
        (300, 25.0, 2.0),
        (400, 30.0, 2.4),
    ]:
        add_tray(
            name=f"Wire mesh tray 50H {width} wide",
            width=width,
            height=50.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.5,  # typical 40–50% for comms/data
        )

    # Low-profile basket for under-floor / ceiling (IT)
    for width, max_load, self_w in [
        (100, 10.0, 1.0),
        (150, 12.0, 1.1),
        (200, 14.0, 1.2),
        (300, 18.0, 1.6),
    ]:
        add_tray(
            name=f"Wire mesh tray 35H {width} wide",
            width=width,
            height=35.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.45,
        )

    # --------------------------------------------------------
    # Solid-bottom trays (for sensitive / EMC)
    # Heavier self-weight, moderate fill.
    # --------------------------------------------------------
    for width, max_load, self_w in [
        (100, 30.0, 3.0),
        (150, 35.0, 3.5),
        (200, 40.0, 4.0),
        (300, 50.0, 5.0),
        (450, 60.0, 6.3),
        (600, 70.0, 7.5),
    ]:
        add_tray(
            name=f"Solid-bottom tray 60H {width} wide",
            width=width,
            height=60.0,
            max_load=max_load,
            self_weight=self_w,
            fill_ratio=0.45,  # slightly lower to help with heat
        )

    return trays




def cable_area_mm2(diameter_mm: float) -> float:
    """
    Compute approximate cross-sectional area of a round cable, in mm².

    A = π * (d/2)²
    """
    radius = diameter_mm / 2.0
    return pi * radius * radius


def compute_cable_tray_stats(
    cables_with_qty: List[Tuple[CableType, int]],
    tray: TrayType,
    effective_fill_ratio_for_height: float = 0.9,
) -> Dict[str, float]:
    """
    Core calculation for the cable tray.

    Args:
        cables_with_qty:
            List of (CableType, quantity) tuples.
        tray:
            TrayType instance.
        effective_fill_ratio_for_height:
            Multiplier to account for the fact that not all tray height
            is practically usable (cable stacking, spacing, etc.).
            E.g. 0.9 means we use 90% of the side height as fill height.

    Returns:
        Dictionary with:
            total_cable_weight_kg_per_m
            tray_self_weight_kg_per_m
            total_weight_kg_per_m
            allowable_load_kg_per_m
            structural_utilisation_percent
            total_cable_area_mm2
            tray_usable_area_mm2
            area_fill_percent
            recommended_max_area_fill_percent
    """
    # Aggregate cable weight and area
    total_cable_weight = 0.0  # kg/m
    total_cable_area = 0.0    # mm^2

    for cable, qty in cables_with_qty:
        if qty <= 0:
            continue
        total_cable_weight += cable.weight_kg_per_m * qty
        total_cable_area += cable_area_mm2(cable.diameter_mm) * qty

    tray_self_weight = tray.self_weight_kg_per_m
    total_weight = tray_self_weight + total_cable_weight

    allowable_load = tray.max_load_kg_per_m

    # Structural utilisation (% of allowable load)
    structural_utilisation = 0.0
    if allowable_load > 0.0:
        structural_utilisation = (total_cable_weight / allowable_load) * 100.0

    # Area-based tray fill
    # Usable fill height is some fraction of tray side height
    usable_height_mm = tray.height_mm * effective_fill_ratio_for_height
    tray_usable_area_mm2 = tray.width_mm * usable_height_mm

    area_fill_percent = 0.0
    if tray_usable_area_mm2 > 0.0:
        area_fill_percent = (total_cable_area / tray_usable_area_mm2) * 100.0

    return {
        "total_cable_weight_kg_per_m": total_cable_weight,
        "tray_self_weight_kg_per_m": tray_self_weight,
        "total_weight_kg_per_m": total_weight,
        "allowable_load_kg_per_m": allowable_load,
        "structural_utilisation_percent": structural_utilisation,
        "total_cable_area_mm2": total_cable_area,
        "tray_usable_area_mm2": tray_usable_area_mm2,
        "area_fill_percent": area_fill_percent,
        "recommended_max_area_fill_percent": tray.max_fill_ratio * 100.0,
    }
