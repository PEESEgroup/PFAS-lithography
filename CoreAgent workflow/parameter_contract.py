"""Canonical parameter units and endpoint dependencies from the study model."""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Mapping


INPUT_SCHEMA_VERSION = "1.0.0"
CALCULATION_PROFILE = "holmquist_2020_pfaa_adapted_usetox_2_1"


@dataclass(frozen=True)
class ParameterSpec:
    """Canonical unit and numerical constraint for one model input."""

    unit: str
    group: str
    minimum: float | None = None
    exclusive_minimum: bool = False
    required: bool = True


# The 25 substance fields used by the deterministic fate/exposure model.
# USEtox landscape and process constants are supplied by the calculator.
# Workflow readiness is evaluated separately for each endpoint below.
PARAMETER_SPECS: dict[str, ParameterSpec] = {
    "mw_g_mol": ParameterSpec("g/mol", "physicochemical", 0.0, True),
    "pka_gain": ParameterSpec("dimensionless", "physicochemical"),
    "pka_loss": ParameterSpec("dimensionless", "physicochemical"),
    "koc_l_per_kg": ParameterSpec("L/kg", "partitioning", 0.0),
    "vapor_pressure_pa": ParameterSpec("Pa", "physicochemical", 0.0),
    "solubility_mg_l": ParameterSpec("mg/L", "physicochemical", 0.0, True),
    "kdoc_l_per_kg": ParameterSpec("L/kg", "partitioning", 0.0),
    "kdeg_air_s": ParameterSpec("1/s", "degradation", 0.0),
    "kdeg_water_s": ParameterSpec("1/s", "degradation", 0.0),
    "kdeg_sediment_s": ParameterSpec("1/s", "degradation", 0.0),
    "kdeg_soil_s": ParameterSpec("1/s", "degradation", 0.0),
    "baf_soil_root": ParameterSpec(
        "kg soil/kg vegetable", "human_exposure", 0.0
    ),
    "baf_soil_leaf": ParameterSpec(
        "kg soil/kg vegetable", "human_exposure", 0.0
    ),
    "btf_meat_d_per_kg": ParameterSpec("d/kg meat", "human_exposure", 0.0),
    "btf_milk_d_per_kg": ParameterSpec("d/kg milk", "human_exposure", 0.0),
    "baf_fish_l_per_kg": ParameterSpec("L/kg fish", "human_exposure", 0.0),
    "kqa_dimensionless": ParameterSpec("dimensionless", "partitioning", 0.0),
    "lcf_air_gas": ParameterSpec(
        "kg air gas/kg vegetable", "human_exposure", 0.0
    ),
    "lcf_air_particle": ParameterSpec(
        "kg air particle/kg vegetable", "human_exposure", 0.0
    ),
    "avlog_ec50_freshwater_log10_mg_per_l": ParameterSpec(
        "log10(mg/L)", "ecotoxicity_effect"
    ),
    "avlog_ec50_marine_log10_mg_per_l": ParameterSpec(
        "log10(mg/L)", "ecotoxicity_effect", required=False
    ),
    "ed50_noncancer_inhalation_kg_lifetime": ParameterSpec(
        "kg/lifetime", "human_effect", 0.0, True
    ),
    "ed50_noncancer_ingestion_kg_lifetime": ParameterSpec(
        "kg/lifetime", "human_effect", 0.0, True
    ),
    "ed50_cancer_inhalation_active_kg_lifetime": ParameterSpec(
        "kg/lifetime", "human_effect", 0.0, True, False
    ),
    "ed50_cancer_ingestion_active_kg_lifetime": ParameterSpec(
        "kg/lifetime", "human_effect", 0.0, True, False
    ),
}

PARAMETER_KEYS = tuple(PARAMETER_SPECS)
REQUIRED_PARAMETER_KEYS = tuple(
    key for key, spec in PARAMETER_SPECS.items() if spec.required
)

# Endpoint-specific dependency sets.  ``REQUIRED_PARAMETER_KEYS`` is retained
# for the strict Holmquist-2020 reproduction interface; workflow callers should
# use these sets so an unavailable human endpoint does not suppress CTUe (or
# the other human endpoint).
FATE_CORE_PARAMETER_KEYS = (
    "mw_g_mol",
    "pka_gain",
    "pka_loss",
    "koc_l_per_kg",
    "vapor_pressure_pa",
    "solubility_mg_l",
    "kdoc_l_per_kg",
    "kdeg_air_s",
    "kdeg_water_s",
    "kdeg_sediment_s",
    "kdeg_soil_s",
    "kqa_dimensionless",
)
HUMAN_EXPOSURE_PARAMETER_KEYS = (
    "baf_soil_root",
    "baf_soil_leaf",
    "btf_meat_d_per_kg",
    "btf_milk_d_per_kg",
    "baf_fish_l_per_kg",
    "lcf_air_gas",
    "lcf_air_particle",
)
FRESHWATER_FATE_XF_PARAMETER_KEYS = (
    *FATE_CORE_PARAMETER_KEYS,
    "baf_fish_l_per_kg",
)
FRESHWATER_CTUe_PARAMETER_KEYS = (
    *FRESHWATER_FATE_XF_PARAMETER_KEYS,
    "avlog_ec50_freshwater_log10_mg_per_l",
)
NONCANCER_CTUh_PARAMETER_KEYS = (
    *FATE_CORE_PARAMETER_KEYS,
    *HUMAN_EXPOSURE_PARAMETER_KEYS,
    "ed50_noncancer_inhalation_kg_lifetime",
    "ed50_noncancer_ingestion_kg_lifetime",
)
CANCER_CTUh_PARAMETER_KEYS = (
    *FATE_CORE_PARAMETER_KEYS,
    *HUMAN_EXPOSURE_PARAMETER_KEYS,
    "ed50_cancer_inhalation_active_kg_lifetime",
    "ed50_cancer_ingestion_active_kg_lifetime",
)
ENDPOINT_PARAMETER_KEYS: dict[str, tuple[str, ...]] = {
    "ctue_freshwater": FRESHWATER_CTUe_PARAMETER_KEYS,
    "ctuh_noncancer": NONCANCER_CTUh_PARAMETER_KEYS,
    "ctuh_cancer": CANCER_CTUh_PARAMETER_KEYS,
}

# Individual human pathways can still be audited when a complete aggregate
# intake fraction is unavailable.  The fate-core dependency is common and is
# intentionally not repeated in this mapping.
HUMAN_PATHWAY_PARAMETER_KEYS: dict[str, tuple[str, ...]] = {
    "inhalation": (),
    "drinking_water": ("baf_fish_l_per_kg",),
    "above_ground_produce": (
        "lcf_air_gas",
        "lcf_air_particle",
        "baf_soil_leaf",
    ),
    "below_ground_produce": ("baf_soil_root",),
    "meat": (
        "btf_meat_d_per_kg",
        "lcf_air_gas",
        "lcf_air_particle",
        "baf_soil_leaf",
    ),
    "dairy": (
        "btf_milk_d_per_kg",
        "lcf_air_gas",
        "lcf_air_particle",
        "baf_soil_leaf",
    ),
    "fish": ("baf_fish_l_per_kg",),
}

