"""Calculate mixture-level PFAS toxicity before and after waste management.

The editable Excel input owns the nine-substance composition, the raw
compartment-specific characterization factors, and the screening assumptions.
Substance-specific factors are used only to calculate one mixture factor for
each endpoint and receiving compartment.  All downstream quarterly, annual,
and cumulative calculations operate on aggregate PFAS mass.

The active comparison contains:

* untreated PFAS emission toxicity for the four manufacturing waste streams;
* residual parent-PFAS toxicity after 90% IER capture and Excel-managed incineration DRE;
* the validated IER process burden for EF v3.1 HTN and ECF_total;
* gas, solvent, and solid-waste incineration process burdens; and
* a managed total that keeps process burdens and residual parent-PFAS toxicity
  visible as separate components before summing compatible endpoints.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INPUT = PROJECT_ROOT / "Data input" / "toxicity" / "PFAS_toxicity_input.xlsx"
DEFAULT_ATTRIBUTION = (
    PROJECT_ROOT / "Excel results" / "attribution" / "GenAI_PFAS_attribution_results.json"
)
DEFAULT_IER_BURDEN = PROJECT_ROOT / "Excel results" / "ier" / "IER_burden_results.json"
DEFAULT_INCINERATION_BURDEN = (
    PROJECT_ROOT / "Excel results" / "treatment" / "Incineration_burden_results.json"
)
DEFAULT_OUTPUT = PROJECT_ROOT / "Excel results" / "toxicity" / "PFAS_toxicity_results.json"

SUBSTANCES = ("PFBA", "PFHxA", "PFHpA", "PFOA", "PFNA", "PFDA", "GenX", "PFBS", "PFOS")
COMPARTMENTS = ("air", "freshwater", "soil")
MATERIAL_NAMES = (
    "photoacid generators",
    "surfactant",
    "top antireflective coatings",
    "polymer",
    "immersion topcoat",
    "PBO/PI",
)
MEDIA_NAMES = ("gas", "solvent", "wastewater", "solid")
STREAM_KEYS = {
    "gas": "Manufacturing flow | gas (kg)",
    "solvent": "Manufacturing flow | solvent (kg)",
    "wastewater": "Manufacturing flow | wastewater (kg)",
    "solid": "Manufacturing flow | solid (kg)",
}
ENDPOINTS = {
    "noncancer_ctuh": {
        "raw_column": "noncancer_ctuh_raw",
        "status_column": "noncancer_ctuh_source_status",
        "unit": "CTUh (cases)",
        "factor_unit": "cases/kg PFAS emitted",
        "ier_impact_code": "HTN",
        "ier_unit": "CTUh",
    },
    "freshwater_ctue_hc50": {
        "raw_column": "freshwater_ctue_hc50_raw",
        "status_column": "freshwater_ctue_hc50_source_status",
        "unit": "CTUe (PAF m3 d)",
        "factor_unit": "PAF m3 d/kg PFAS emitted",
        "ier_impact_code": "ECF_total",
        "ier_unit": "CTUe",
    },
}
REQUIRED_SHEETS = {"Assumptions", "Composition", "Substance_Factors", "Sources"}
EXPECTED_PERIODS = tuple(
    f"{year} Q{quarter}"
    for year in range(2025, 2031)
    for quarter in range(1, 5)
)


class ToxicityInputError(ValueError):
    """Raised when an input or upstream interface violates the model contract."""


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        return None if not math.isfinite(float(value)) else float(value)
    if isinstance(value, (np.bool_, bool)):
        return bool(value)
    if pd.isna(value):
        return None
    return value


def _read_json(source: str | Path | dict[str, Any], label: str) -> tuple[dict[str, Any], str]:
    if isinstance(source, dict):
        return source, f"in-memory {label}"
    path = Path(source).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")
    return json.loads(path.read_text(encoding="utf-8")), str(path)


def _require_columns(frame: pd.DataFrame, sheet: str, columns: Iterable[str]) -> None:
    missing = set(columns).difference(frame.columns)
    if missing:
        raise ToxicityInputError(f"{sheet} is missing columns: {sorted(missing)}")


def _as_bool(value: Any, field: str) -> bool:
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    if isinstance(value, (int, float, np.integer, np.floating)) and not pd.isna(value):
        if float(value) in {0.0, 1.0}:
            return bool(int(value))
    normalized = str(value).strip().lower()
    if normalized in {"true", "yes", "y", "1"}:
        return True
    if normalized in {"false", "no", "n", "0"}:
        return False
    raise ToxicityInputError(f"{field} must be TRUE/FALSE; received {value!r}")


def _read_toxicity_input(input_path: str | Path) -> dict[str, Any]:
    path = Path(input_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Toxicity input workbook not found: {path}")
    book = pd.ExcelFile(path, engine="openpyxl")
    missing_sheets = REQUIRED_SHEETS.difference(book.sheet_names)
    if missing_sheets:
        raise ToxicityInputError(f"Toxicity workbook is missing sheets: {sorted(missing_sheets)}")

    assumptions_frame = pd.read_excel(book, sheet_name="Assumptions")
    composition = pd.read_excel(book, sheet_name="Composition")
    factors = pd.read_excel(book, sheet_name="Substance_Factors")
    sources = pd.read_excel(book, sheet_name="Sources")
    _require_columns(assumptions_frame, "Assumptions", {"parameter_id", "value"})
    _require_columns(composition, "Composition", {"pfas", "mass_fraction"})
    _require_columns(
        factors,
        "Substance_Factors",
        {
            "pfas",
            "emission_compartment",
            "noncancer_ctuh_raw",
            "noncancer_ctuh_source_status",
            "freshwater_ctue_hc50_raw",
            "freshwater_ctue_hc50_source_status",
            "source_reference",
            "basis_or_limitation",
        },
    )

    if assumptions_frame["parameter_id"].duplicated().any():
        raise ToxicityInputError("Assumptions contains duplicate parameter_id values")
    assumptions = {
        str(row["parameter_id"]).strip(): row["value"]
        for _, row in assumptions_frame.iterrows()
        if str(row["parameter_id"]).strip()
    }
    required_assumptions = {
        "incineration_parent_dre",
        "ambiguous_untreated_air_share",
        "ambiguous_untreated_freshwater_share",
        "ambiguous_untreated_soil_share",
        "include_partial_in_imputation_mean",
        "all_streams_use_common_nine_pfas_mixture",
    }
    missing_assumptions = required_assumptions.difference(assumptions)
    if missing_assumptions:
        raise ToxicityInputError(f"Missing toxicity assumptions: {sorted(missing_assumptions)}")

    dre = float(assumptions["incineration_parent_dre"])
    if not 0 < dre <= 1:
        raise ToxicityInputError("incineration_parent_dre must satisfy 0 < DRE <= 1")
    ambiguous_shares = {
        "air": float(assumptions["ambiguous_untreated_air_share"]),
        "freshwater": float(assumptions["ambiguous_untreated_freshwater_share"]),
        "soil": float(assumptions["ambiguous_untreated_soil_share"]),
    }
    if any(value < 0 for value in ambiguous_shares.values()) or not math.isclose(
        sum(ambiguous_shares.values()), 1.0, rel_tol=0, abs_tol=1e-10
    ):
        raise ToxicityInputError("Ambiguous untreated compartment shares must be nonnegative and sum to 1")
    if not _as_bool(
        assumptions["include_partial_in_imputation_mean"],
        "include_partial_in_imputation_mean",
    ):
        raise ToxicityInputError("Current contract requires partial values to participate in the arithmetic mean")
    if not _as_bool(
        assumptions["all_streams_use_common_nine_pfas_mixture"],
        "all_streams_use_common_nine_pfas_mixture",
    ):
        raise ToxicityInputError("Current contract requires one common nine-PFAS composition for all four streams")

    composition = composition.dropna(subset=["pfas"]).copy()
    composition["pfas"] = composition["pfas"].astype(str).str.strip()
    composition = composition[composition["pfas"].isin(SUBSTANCES)].copy()
    composition["mass_fraction"] = pd.to_numeric(composition["mass_fraction"], errors="raise")
    if composition["pfas"].duplicated().any():
        raise ToxicityInputError("Composition contains duplicate PFAS rows")
    if set(composition["pfas"]) != set(SUBSTANCES):
        raise ToxicityInputError(
            "Composition must contain exactly the active nine PFAS: " + ", ".join(SUBSTANCES)
        )
    if (composition["mass_fraction"] < 0).any() or not math.isclose(
        float(composition["mass_fraction"].sum()), 1.0, rel_tol=0, abs_tol=1e-8
    ):
        raise ToxicityInputError("Nine-PFAS mass fractions must be nonnegative and sum to 1")
    composition_map = {
        row["pfas"]: float(row["mass_fraction"])
        for _, row in composition.iterrows()
    }

    factors = factors.dropna(subset=["pfas", "emission_compartment"]).copy()
    factors["pfas"] = factors["pfas"].astype(str).str.strip()
    factors["emission_compartment"] = factors["emission_compartment"].astype(str).str.strip().str.lower()
    expected_keys = {(pfas, compartment) for pfas in SUBSTANCES for compartment in COMPARTMENTS}
    actual_keys = set(zip(factors["pfas"], factors["emission_compartment"]))
    if actual_keys != expected_keys or len(factors) != len(expected_keys):
        missing = sorted(expected_keys.difference(actual_keys))
        extra = sorted(actual_keys.difference(expected_keys))
        raise ToxicityInputError(
            f"Substance_Factors must contain exactly 27 unique PFAS-compartment rows; missing={missing}, extra={extra}"
        )
    if factors.duplicated(subset=["pfas", "emission_compartment"]).any():
        raise ToxicityInputError("Substance_Factors contains duplicate PFAS-compartment rows")
    for metadata in ENDPOINTS.values():
        factors[metadata["raw_column"]] = pd.to_numeric(
            factors[metadata["raw_column"]], errors="coerce"
        )
        negative = factors[metadata["raw_column"]].dropna() < 0
        if negative.any():
            raise ToxicityInputError(f"{metadata['raw_column']} contains negative factors")

    return {
        "input_path": str(path),
        "assumptions": assumptions,
        "dre": dre,
        "ambiguous_shares": ambiguous_shares,
        "composition": composition_map,
        "factors": factors,
        "sources": sources.to_dict(orient="records"),
    }


def _calculate_mixture_factors(model_input: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    factors = model_input["factors"]
    composition = model_input["composition"]
    audit_rows: list[dict[str, Any]] = []
    mixture_rows: list[dict[str, Any]] = []

    for compartment in COMPARTMENTS:
        group = factors[factors["emission_compartment"] == compartment].copy()
        for endpoint, metadata in ENDPOINTS.items():
            raw_column = metadata["raw_column"]
            available = group[raw_column].dropna().astype(float)
            if available.empty:
                raise ToxicityInputError(
                    f"No non-missing {endpoint} factors are available for {compartment}"
                )
            imputation_mean = float(available.mean())
            mixture_factor = 0.0
            imputed_count = 0
            partial_count = 0
            for pfas in SUBSTANCES:
                source = group[group["pfas"] == pfas].iloc[0]
                raw_value = source[raw_column]
                source_status = str(source[metadata["status_column"]] or "NA").strip()
                is_missing = pd.isna(raw_value)
                used_value = imputation_mean if is_missing else float(raw_value)
                imputation_flag = "arithmetic mean of non-missing values in same endpoint and compartment" if is_missing else "none"
                if is_missing:
                    imputed_count += 1
                if source_status.lower() == "partial":
                    partial_count += 1
                weight = float(composition[pfas])
                contribution = weight * used_value
                mixture_factor += contribution
                audit_rows.append(
                    {
                        "pfas": pfas,
                        "mass_fraction": weight,
                        "emission_compartment": compartment,
                        "endpoint": endpoint,
                        "factor_unit": metadata["factor_unit"],
                        "raw_factor": None if is_missing else float(raw_value),
                        "source_status": source_status,
                        "factor_used": used_value,
                        "imputation_flag": imputation_flag,
                        "weighted_contribution": contribution,
                        "source_reference": source["source_reference"],
                        "basis_or_limitation": source["basis_or_limitation"],
                    }
                )
            mixture_rows.append(
                {
                    "emission_compartment": compartment,
                    "endpoint": endpoint,
                    "factor_unit": metadata["factor_unit"],
                    "mixture_characterization_factor": mixture_factor,
                    "arithmetic_mean_used_for_missing": imputation_mean,
                    "available_factor_count": int(len(available)),
                    "imputed_substance_count": imputed_count,
                    "partial_source_count": partial_count,
                    "result_classification": "partial/imputed screening result",
                }
            )
    return audit_rows, mixture_rows


def _mixture_lookup(mixture_rows: list[dict[str, Any]]) -> dict[tuple[str, str], float]:
    return {
        (row["emission_compartment"], row["endpoint"]): float(
            row["mixture_characterization_factor"]
        )
        for row in mixture_rows
    }


def _toxicity_by_compartment(
    masses: dict[str, float],
    mixture_factors: dict[tuple[str, str], float],
    endpoint: str,
) -> tuple[dict[str, float], float]:
    values = {
        compartment: float(masses[compartment])
        * mixture_factors[(compartment, endpoint)]
        for compartment in COMPARTMENTS
    }
    return values, float(sum(values.values()))


def _sum_rows(rows: list[dict[str, Any]], keys: list[str]) -> dict[str, float]:
    return {key: float(sum(float(row[key]) for row in rows)) for key in keys}


def _managed_toxicity_summary(
    rows: list[dict[str, Any]], period_type: str
) -> list[dict[str, Any]]:
    """Return a compact, plotting-ready view of the two compatible endpoints.

    The residual terms quantify parent PFAS that still reaches air, freshwater,
    or soil after the modeled management route. Process terms are the HTN or
    ECF burdens of IER production/spent-mixture incineration plus treatment of
    gas, solvent, and solid waste. Only like endpoints are added.
    """

    summary: list[dict[str, Any]] = []
    for row in rows:
        if period_type == "quarterly":
            period = str(row["period"])
            year = int(row["year"])
            quarter: int | None = int(row["quarter"])
        elif period_type == "annual":
            period = str(int(row["year"]))
            year = int(row["year"])
            quarter = None
        else:
            period = str(row["period"])
            year = None
            quarter = None
        summary.append(
            {
                "period_type": period_type,
                "period": period,
                "year": year,
                "quarter": quarter,
                "scenario_id": row["scenario_id"],
                "role": row["role"],
                "residual_pfas_htn_ctuh": row["residual_noncancer_ctuh_total"],
                "treatment_process_htn_ctuh": row[
                    "total_treatment_process_burden_noncancer_ctuh"
                ],
                "managed_total_htn_ctuh": row["managed_total_noncancer_ctuh"],
                "residual_pfas_ecf_ctue": row[
                    "residual_freshwater_ctue_hc50_total"
                ],
                "treatment_process_ecf_ctue": row[
                    "total_treatment_process_burden_freshwater_ctue_hc50"
                ],
                "managed_total_ecf_ctue": row[
                    "managed_total_freshwater_ctue_hc50"
                ],
            }
        )
    return summary


SOURCE_TOXICITY_FIELDS = (
    "source_pfas_flow_kg",
    "residual_pfas_htn_ctuh",
    "treatment_process_htn_ctuh",
    "managed_total_htn_ctuh",
    "residual_pfas_ecf_ctue",
    "treatment_process_ecf_ctue",
    "managed_total_ecf_ctue",
)


def _aggregate_source_toxicity(
    quarterly_rows: list[dict[str, Any]], source_key: str
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    annual: list[dict[str, Any]] = []
    cumulative: list[dict[str, Any]] = []
    scenario_ids = list(dict.fromkeys(str(row["scenario_id"]) for row in quarterly_rows))
    sources = list(dict.fromkeys(str(row[source_key]) for row in quarterly_rows))
    for scenario_id in scenario_ids:
        scenario_rows = [row for row in quarterly_rows if row["scenario_id"] == scenario_id]
        role = str(scenario_rows[0]["role"])
        for source in sources:
            source_rows = [row for row in scenario_rows if row[source_key] == source]
            for year in range(2025, 2031):
                rows = [row for row in source_rows if row["year"] == year]
                if len(rows) != 4:
                    raise AssertionError(
                        f"{scenario_id}/{source}/{year} does not contain four source-toxicity rows"
                    )
                annual.append(
                    {
                        "period": str(year),
                        "year": year,
                        "scenario_id": scenario_id,
                        "role": role,
                        source_key: source,
                        **_sum_rows(rows, list(SOURCE_TOXICITY_FIELDS)),
                    }
                )
            cumulative.append(
                {
                    "period": "2025-2030",
                    "scenario_id": scenario_id,
                    "role": role,
                    source_key: source,
                    **_sum_rows(source_rows, list(SOURCE_TOXICITY_FIELDS)),
                }
            )
    return annual, cumulative


def calculate_toxicity(
    attribution_results: str | Path | dict[str, Any] = DEFAULT_ATTRIBUTION,
    ier_burden_results: str | Path | dict[str, Any] = DEFAULT_IER_BURDEN,
    incineration_burden_results: str | Path | dict[str, Any] = DEFAULT_INCINERATION_BURDEN,
    toxicity_input_path: str | Path = DEFAULT_INPUT,
) -> dict[str, Any]:
    """Calculate aggregate mixture toxicity and the managed comparison."""

    model_input = _read_toxicity_input(toxicity_input_path)
    attribution, attribution_source = _read_json(
        attribution_results, "GenAI PFAS attribution output"
    )
    ier_burden, ier_source = _read_json(ier_burden_results, "IER burden output")
    incineration_burden, incineration_source = _read_json(
        incineration_burden_results, "Other-stream incineration burden output"
    )
    factor_audit, mixture_rows = _calculate_mixture_factors(model_input)
    mixture_factors = _mixture_lookup(mixture_rows)

    quarterly_attribution = attribution.get("quarterly", [])
    if len(quarterly_attribution) != len(EXPECTED_PERIODS):
        raise ToxicityInputError(
            f"Attribution must contain {len(EXPECTED_PERIODS)} quarters; found {len(quarterly_attribution)}"
        )
    if tuple(str(row.get("Period")) for row in quarterly_attribution) != EXPECTED_PERIODS:
        raise ToxicityInputError("Attribution periods are not 2025 Q1-2030 Q4")

    mass_scenarios = ier_burden.get("mass_scenarios", [])
    if not mass_scenarios:
        raise ToxicityInputError("IER burden output contains no mass scenarios")
    primary_scenario_id = str(ier_burden.get("primary_scenario_id", ""))
    if primary_scenario_id not in {str(row["scenario_id"]) for row in mass_scenarios}:
        raise ToxicityInputError("IER primary scenario is absent from mass_scenarios")

    ier_burden_lookup: dict[tuple[str, str, str], float] = {}
    treatment_units: dict[str, set[str]] = {
        endpoint: set() for endpoint in ENDPOINTS
    }
    impact_to_endpoint = {
        metadata["ier_impact_code"]: endpoint
        for endpoint, metadata in ENDPOINTS.items()
    }
    for row in ier_burden.get("quarterly_burden", []):
        impact_code = str(row.get("impact_code"))
        if impact_code not in impact_to_endpoint:
            continue
        endpoint = impact_to_endpoint[impact_code]
        key = (str(row["scenario_id"]), str(row["period"]), endpoint)
        if key in ier_burden_lookup:
            raise ToxicityInputError(f"Duplicate IER treatment burden row: {key}")
        ier_burden_lookup[key] = float(row["total_ier_treatment_burden"])
        treatment_units[endpoint].add(str(row.get("unit", "")))
    expected_burden_rows = len(mass_scenarios) * len(EXPECTED_PERIODS) * len(ENDPOINTS)
    if len(ier_burden_lookup) != expected_burden_rows:
        raise ToxicityInputError(
            f"IER burden interface should provide {expected_burden_rows} active rows; found {len(ier_burden_lookup)}"
        )
    for endpoint, metadata in ENDPOINTS.items():
        if treatment_units[endpoint] != {metadata["ier_unit"]}:
            raise ToxicityInputError(
                f"IER {metadata['ier_impact_code']} unit mismatch: {sorted(treatment_units[endpoint])}"
            )

    incineration_burden_lookup: dict[tuple[str, str], float] = {}
    incineration_factor_lookup: dict[tuple[str, str, str], float] = {}
    incineration_units: dict[str, set[str]] = {endpoint: set() for endpoint in ENDPOINTS}
    incineration_media: dict[tuple[str, str], set[str]] = {}
    for row in incineration_burden.get("quarterly_burden", []):
        impact_code = str(row.get("impact_code"))
        if impact_code not in impact_to_endpoint:
            continue
        endpoint = impact_to_endpoint[impact_code]
        key = (str(row["period"]), endpoint)
        incineration_burden_lookup[key] = incineration_burden_lookup.get(key, 0.0) + float(
            row["process_burden"]
        )
        factor_key = (str(row["period"]), str(row["medium"]), endpoint)
        factor = float(row["factor_per_kg_pfas_sent"])
        if factor_key in incineration_factor_lookup and not math.isclose(
            incineration_factor_lookup[factor_key], factor, rel_tol=0, abs_tol=1e-12
        ):
            raise ToxicityInputError(f"Conflicting incineration factor row: {factor_key}")
        incineration_factor_lookup[factor_key] = factor
        incineration_units[endpoint].add(str(row.get("unit", "")))
        incineration_media.setdefault(key, set()).add(str(row.get("medium", "")))
    expected_incineration_rows = len(EXPECTED_PERIODS) * len(ENDPOINTS)
    if len(incineration_burden_lookup) != expected_incineration_rows:
        raise ToxicityInputError(
            "Incineration burden interface should provide "
            f"{expected_incineration_rows} period-endpoint totals; "
            f"found {len(incineration_burden_lookup)}"
        )
    for key, media in incineration_media.items():
        if media != {"gas", "solvent", "solid"}:
            raise ToxicityInputError(
                f"Incineration burden {key} does not contain all three media: {sorted(media)}"
            )
    for endpoint, metadata in ENDPOINTS.items():
        if incineration_units[endpoint] != {metadata["ier_unit"]}:
            raise ToxicityInputError(
                f"Incineration {metadata['ier_impact_code']} unit mismatch: "
                f"{sorted(incineration_units[endpoint])}"
            )

    dre = float(model_input["dre"])
    ambiguous = model_input["ambiguous_shares"]
    quarterly_rows: list[dict[str, Any]] = []
    quarterly_by_material: list[dict[str, Any]] = []
    quarterly_by_medium: list[dict[str, Any]] = []
    max_untreated_mass_balance_error = 0.0
    max_treated_mass_balance_error = 0.0
    max_ier_residual_basis_error = 0.0
    max_process_burden_identity_error = 0.0
    max_managed_total_identity_error = 0.0
    max_net_avoided_identity_error = 0.0
    max_material_stream_reconciliation_error = 0.0
    max_material_toxicity_reconciliation_error = 0.0
    max_medium_toxicity_reconciliation_error = 0.0

    for scenario in mass_scenarios:
        scenario_id = str(scenario["scenario_id"])
        captured_per_influent = float(scenario["captured_pfas_kg_per_kg_influent"])
        bypass_per_influent = float(scenario["bypass_pfas_kg_per_kg_influent"])
        if not math.isclose(
            captured_per_influent + bypass_per_influent,
            1.0,
            rel_tol=0,
            abs_tol=1e-10,
        ):
            raise ToxicityInputError(f"IER capture and bypass do not sum to 1 for {scenario_id}")

        for attribution_row in quarterly_attribution:
            period = str(attribution_row["Period"])
            stream_mass = {
                stream: float(attribution_row[key])
                for stream, key in STREAM_KEYS.items()
            }
            if any(value < 0 for value in stream_mass.values()):
                raise ToxicityInputError(f"Negative manufacturing PFAS flow in {period}")
            stream_material_mass: dict[str, dict[str, float]] = {}
            for stream in MEDIA_NAMES:
                stream_material_mass[stream] = {}
                for material in MATERIAL_NAMES:
                    key = f"Manufacturing flow | {stream} | {material} (kg)"
                    if key not in attribution_row:
                        raise ToxicityInputError(
                            f"Attribution output lacks material-by-medium mass: {period}/{key}"
                        )
                    value = float(attribution_row[key])
                    if value < 0:
                        raise ToxicityInputError(
                            f"Negative material-by-medium PFAS flow: {period}/{stream}/{material}"
                        )
                    stream_material_mass[stream][material] = value
                max_material_stream_reconciliation_error = max(
                    max_material_stream_reconciliation_error,
                    abs(sum(stream_material_mass[stream].values()) - stream_mass[stream]),
                )
            ambiguous_mass = stream_mass["solvent"] + stream_mass["solid"]
            untreated_destination = {
                "air": stream_mass["gas"] + ambiguous_mass * ambiguous["air"],
                "freshwater": stream_mass["wastewater"]
                + ambiguous_mass * ambiguous["freshwater"],
                "soil": ambiguous_mass * ambiguous["soil"],
            }
            total_stream_pfas = float(sum(stream_mass.values()))
            untreated_mass_error = sum(untreated_destination.values()) - total_stream_pfas
            max_untreated_mass_balance_error = max(
                max_untreated_mass_balance_error, abs(untreated_mass_error)
            )

            wastewater = stream_mass["wastewater"]
            captured_pfas = wastewater * captured_per_influent
            bypass_pfas = wastewater * bypass_per_influent
            other_incinerated_pfas = (
                stream_mass["gas"] + stream_mass["solvent"] + stream_mass["solid"]
            )
            incinerated_parent_pfas = captured_pfas + other_incinerated_pfas
            ier_parent_residual_air = captured_pfas * (1.0 - dre)
            other_parent_residual_air = other_incinerated_pfas * (1.0 - dre)
            residual_destination = {
                "air": ier_parent_residual_air + other_parent_residual_air,
                "freshwater": bypass_pfas,
                "soil": 0.0,
            }
            destroyed_parent_pfas = incinerated_parent_pfas * dre
            treated_mass_error = (
                total_stream_pfas
                - sum(residual_destination.values())
                - destroyed_parent_pfas
            )
            max_treated_mass_balance_error = max(
                max_treated_mass_balance_error, abs(treated_mass_error)
            )
            # This explicit identity prevents spent IER + PFAS mixture mass from
            # entering the residual-toxicity calculation.
            max_ier_residual_basis_error = max(
                max_ier_residual_basis_error,
                abs(ier_parent_residual_air - captured_pfas * (1.0 - dre)),
            )

            row: dict[str, Any] = {
                "period": period,
                "year": int(attribution_row["Year"]),
                "quarter": int(attribution_row["Quarter"]),
                "scenario_id": scenario_id,
                "role": str(scenario["role"]),
                "capacity_utilization_factor_f_c": float(
                    scenario["capacity_utilization_factor_f_c"]
                ),
                "incineration_parent_dre": dre,
                "total_modeled_manufacturing_flow_pfas_kg": total_stream_pfas,
                "gas_pfas_kg": stream_mass["gas"],
                "solvent_pfas_kg": stream_mass["solvent"],
                "wastewater_pfas_kg": wastewater,
                "solid_pfas_kg": stream_mass["solid"],
                "untreated_air_pfas_kg": untreated_destination["air"],
                "untreated_freshwater_pfas_kg": untreated_destination["freshwater"],
                "untreated_soil_pfas_kg": untreated_destination["soil"],
                "ier_captured_pfas_kg": captured_pfas,
                "ier_bypass_pfas_to_freshwater_kg": bypass_pfas,
                "other_stream_pfas_incinerated_kg": other_incinerated_pfas,
                "ier_parent_pfas_residual_to_air_kg": ier_parent_residual_air,
                "other_parent_pfas_residual_to_air_kg": other_parent_residual_air,
                "residual_air_pfas_kg": residual_destination["air"],
                "residual_freshwater_pfas_kg": residual_destination["freshwater"],
                "residual_soil_pfas_kg": residual_destination["soil"],
                "destroyed_parent_pfas_kg": destroyed_parent_pfas,
            }
            for endpoint in ENDPOINTS:
                untreated_by_compartment, untreated_total = _toxicity_by_compartment(
                    untreated_destination, mixture_factors, endpoint
                )
                residual_by_compartment, residual_total = _toxicity_by_compartment(
                    residual_destination, mixture_factors, endpoint
                )
                ier_process_burden = ier_burden_lookup[
                    (scenario_id, period, endpoint)
                ]
                other_incineration_process_burden = incineration_burden_lookup[
                    (period, endpoint)
                ]
                total_treatment_process_burden = (
                    ier_process_burden + other_incineration_process_burden
                )
                prefix = endpoint
                for compartment in COMPARTMENTS:
                    row[f"untreated_{prefix}_{compartment}"] = untreated_by_compartment[
                        compartment
                    ]
                    row[f"residual_{prefix}_{compartment}"] = residual_by_compartment[
                        compartment
                    ]
                row[f"untreated_{prefix}_total"] = untreated_total
                row[f"residual_{prefix}_total"] = residual_total
                row[f"avoided_parent_pfas_{prefix}"] = untreated_total - residual_total
                row[f"ier_process_burden_{prefix}"] = ier_process_burden
                row[f"other_incineration_process_burden_{prefix}"] = (
                    other_incineration_process_burden
                )
                row[f"total_treatment_process_burden_{prefix}"] = (
                    total_treatment_process_burden
                )
                row[f"managed_total_{prefix}"] = (
                    residual_total + total_treatment_process_burden
                )
                row[f"net_avoided_{prefix}"] = (
                    untreated_total - residual_total - total_treatment_process_burden
                )
                max_process_burden_identity_error = max(
                    max_process_burden_identity_error,
                    abs(
                        row[f"total_treatment_process_burden_{prefix}"]
                        - row[f"ier_process_burden_{prefix}"]
                        - row[f"other_incineration_process_burden_{prefix}"]
                    ),
                )
                max_managed_total_identity_error = max(
                    max_managed_total_identity_error,
                    abs(
                        row[f"managed_total_{prefix}"]
                        - row[f"residual_{prefix}_total"]
                        - row[f"total_treatment_process_burden_{prefix}"]
                    ),
                )
                max_net_avoided_identity_error = max(
                    max_net_avoided_identity_error,
                    abs(
                        row[f"net_avoided_{prefix}"]
                        - row[f"untreated_{prefix}_total"]
                        + row[f"managed_total_{prefix}"]
                    ),
                )
            quarterly_rows.append(row)

            def source_toxicity_row(
                source_key: str, source_name: str, source_stream_mass: dict[str, float]
            ) -> dict[str, Any]:
                source_wastewater = source_stream_mass["wastewater"]
                source_residual_destination = {
                    "air": (
                        source_wastewater * captured_per_influent
                        + source_stream_mass["gas"]
                        + source_stream_mass["solvent"]
                        + source_stream_mass["solid"]
                    ) * (1.0 - dre),
                    "freshwater": source_wastewater * bypass_per_influent,
                    "soil": 0.0,
                }
                detail: dict[str, Any] = {
                    "period": period,
                    "year": int(attribution_row["Year"]),
                    "quarter": int(attribution_row["Quarter"]),
                    "scenario_id": scenario_id,
                    "role": str(scenario["role"]),
                    source_key: source_name,
                    "source_pfas_flow_kg": float(sum(source_stream_mass.values())),
                }
                for endpoint in ENDPOINTS:
                    _, residual = _toxicity_by_compartment(
                        source_residual_destination, mixture_factors, endpoint
                    )
                    ier_factor = (
                        ier_burden_lookup[(scenario_id, period, endpoint)] / wastewater
                    )
                    process = source_wastewater * ier_factor
                    process += sum(
                        source_stream_mass[stream]
                        * incineration_factor_lookup[(period, stream, endpoint)]
                        for stream in ("gas", "solvent", "solid")
                    )
                    if endpoint == "noncancer_ctuh":
                        detail["residual_pfas_htn_ctuh"] = residual
                        detail["treatment_process_htn_ctuh"] = process
                        detail["managed_total_htn_ctuh"] = residual + process
                    else:
                        detail["residual_pfas_ecf_ctue"] = residual
                        detail["treatment_process_ecf_ctue"] = process
                        detail["managed_total_ecf_ctue"] = residual + process
                return detail

            material_details: list[dict[str, Any]] = []
            for material in MATERIAL_NAMES:
                material_details.append(
                    source_toxicity_row(
                        "material",
                        material,
                        {
                            stream: stream_material_mass[stream][material]
                            for stream in MEDIA_NAMES
                        },
                    )
                )
            quarterly_by_material.extend(material_details)

            medium_details: list[dict[str, Any]] = []
            for medium in MEDIA_NAMES:
                medium_details.append(
                    source_toxicity_row(
                        "medium",
                        medium,
                        {
                            stream: stream_mass[medium] if stream == medium else 0.0
                            for stream in MEDIA_NAMES
                        },
                    )
                )
            quarterly_by_medium.extend(medium_details)

            reconciliation_fields = {
                "residual_pfas_htn_ctuh": "residual_noncancer_ctuh_total",
                "treatment_process_htn_ctuh": "total_treatment_process_burden_noncancer_ctuh",
                "managed_total_htn_ctuh": "managed_total_noncancer_ctuh",
                "residual_pfas_ecf_ctue": "residual_freshwater_ctue_hc50_total",
                "treatment_process_ecf_ctue": "total_treatment_process_burden_freshwater_ctue_hc50",
                "managed_total_ecf_ctue": "managed_total_freshwater_ctue_hc50",
            }
            for detail_field, total_field in reconciliation_fields.items():
                material_error = abs(
                    sum(float(item[detail_field]) for item in material_details)
                    - float(row[total_field])
                )
                medium_error = abs(
                    sum(float(item[detail_field]) for item in medium_details)
                    - float(row[total_field])
                )
                max_material_toxicity_reconciliation_error = max(
                    max_material_toxicity_reconciliation_error, material_error
                )
                max_medium_toxicity_reconciliation_error = max(
                    max_medium_toxicity_reconciliation_error, medium_error
                )

    additive_keys = [
        key
        for key, value in quarterly_rows[0].items()
        if isinstance(value, (int, float))
        and key
        not in {
            "year",
            "quarter",
            "capacity_utilization_factor_f_c",
            "incineration_parent_dre",
        }
    ]
    annual_rows: list[dict[str, Any]] = []
    cumulative_rows: list[dict[str, Any]] = []
    for scenario in mass_scenarios:
        scenario_id = str(scenario["scenario_id"])
        scenario_rows = [row for row in quarterly_rows if row["scenario_id"] == scenario_id]
        if len(scenario_rows) != len(EXPECTED_PERIODS):
            raise AssertionError(f"{scenario_id} does not contain 24 quarterly toxicity rows")
        for year in range(2025, 2031):
            year_rows = [row for row in scenario_rows if row["year"] == year]
            if len(year_rows) != 4:
                raise AssertionError(f"{scenario_id}/{year} does not contain four quarters")
            annual_rows.append(
                {
                    "year": year,
                    "scenario_id": scenario_id,
                    "role": str(scenario["role"]),
                    "capacity_utilization_factor_f_c": float(
                        scenario["capacity_utilization_factor_f_c"]
                    ),
                    "incineration_parent_dre": dre,
                    **_sum_rows(year_rows, additive_keys),
                }
            )
        cumulative_rows.append(
            {
                "period": "2025-2030",
                "scenario_id": scenario_id,
                "role": str(scenario["role"]),
                "capacity_utilization_factor_f_c": float(
                    scenario["capacity_utilization_factor_f_c"]
                ),
                "incineration_parent_dre": dre,
                **_sum_rows(scenario_rows, additive_keys),
            }
        )

    annual_by_material, cumulative_by_material = _aggregate_source_toxicity(
        quarterly_by_material, "material"
    )
    annual_by_medium, cumulative_by_medium = _aggregate_source_toxicity(
        quarterly_by_medium, "medium"
    )

    tolerance = 1e-9
    burden_identity_tolerance = 1e-8
    quality_checks = {
        "quarter_count_per_scenario": len(EXPECTED_PERIODS),
        "scenario_count": len(mass_scenarios),
        "factor_audit_row_count": len(factor_audit),
        "mixture_factor_row_count": len(mixture_rows),
        "composition_sum": float(sum(model_input["composition"].values())),
        "max_untreated_mass_balance_error_kg": max_untreated_mass_balance_error,
        "max_treated_mass_balance_error_kg": max_treated_mass_balance_error,
        "max_ier_residual_pfas_basis_error_kg": max_ier_residual_basis_error,
        "max_process_burden_identity_error": max_process_burden_identity_error,
        "max_managed_total_identity_error": max_managed_total_identity_error,
        "max_net_avoided_identity_error": max_net_avoided_identity_error,
        "max_material_stream_reconciliation_error_kg": max_material_stream_reconciliation_error,
        "max_material_toxicity_reconciliation_error": max_material_toxicity_reconciliation_error,
        "max_medium_toxicity_reconciliation_error": max_medium_toxicity_reconciliation_error,
        "ier_residual_uses_captured_pfas_not_spent_mix": max_ier_residual_basis_error <= tolerance,
        "active_ier_units_match": True,
        "active_other_incineration_units_match": True,
        "all_three_other_incineration_media_present": all(
            media == {"gas", "solvent", "solid"}
            for media in incineration_media.values()
        ),
        "all_mixture_factors_positive": all(
            float(row["mixture_characterization_factor"]) > 0 for row in mixture_rows
        ),
    }
    quality_checks["all_checks_passed"] = all(
        (
            math.isclose(quality_checks["composition_sum"], 1.0, rel_tol=0, abs_tol=1e-8),
            max_untreated_mass_balance_error <= tolerance,
            max_treated_mass_balance_error <= tolerance,
            quality_checks["ier_residual_uses_captured_pfas_not_spent_mix"],
            max_process_burden_identity_error <= burden_identity_tolerance,
            max_managed_total_identity_error <= burden_identity_tolerance,
            max_net_avoided_identity_error <= burden_identity_tolerance,
            max_material_stream_reconciliation_error <= tolerance,
            max_material_toxicity_reconciliation_error <= burden_identity_tolerance,
            max_medium_toxicity_reconciliation_error <= burden_identity_tolerance,
            quality_checks["active_ier_units_match"],
            quality_checks["active_other_incineration_units_match"],
            quality_checks["all_three_other_incineration_media_present"],
            quality_checks["all_mixture_factors_positive"],
        )
    )

    return _json_safe(
        {
            "schema_version": "1.3",
            "model_name": "Nine-PFAS mixture toxicity with complete active treatment-process integration",
            "period": "2025 Q1-2030 Q4",
            "functional_unit": (
                "Aggregate PFAS mass in each modeled manufacturing waste stream; "
                "toxicity factors are per kg PFAS emitted to the receiving compartment"
            ),
            "input_workbook": model_input["input_path"],
            "upstream_interfaces": {
                "attribution_output": attribution_source,
                "ier_burden_output": ier_source,
                "other_incineration_burden_output": incineration_source,
            },
            "active_endpoints": [
                {
                    "endpoint": endpoint,
                    "emission_factor_unit": metadata["factor_unit"],
                    "result_unit": metadata["unit"],
                    "ier_impact_code": metadata["ier_impact_code"],
                }
                for endpoint, metadata in ENDPOINTS.items()
            ],
            "composition": model_input["composition"],
            "assumptions": {
                "common_mixture_for_all_four_streams": True,
                "substance_resolution_boundary": (
                    "Nine substances are resolved only while calculating mixture characterization factors; "
                    "all later calculations use aggregate PFAS mass."
                ),
                "missing_factor_rule": (
                    "For each endpoint and emission compartment separately, replace a missing raw factor "
                    "with the arithmetic mean of all non-missing factors; partial values participate."
                ),
                "incineration_parent_dre": dre,
                "incineration_scope": (
                    "Gas, solvent, solid waste, and PFAS captured on spent IER all use the same parent-PFAS DRE."
                ),
                "incineration_residual_destination": "air",
                "ier_residual_mass_basis": "captured PFAS mass only; spent IER mass is excluded",
                "ier_bypass_destination": "freshwater",
                "untreated_gas_destination": "air",
                "untreated_wastewater_destination": "freshwater",
                "untreated_solvent_and_solid_compartment_shares": model_input[
                    "ambiguous_shares"
                ],
                "unquantified_incineration_products": (
                    "PIC/PID and mineralization completeness are outside the quantified result."
                ),
            },
            "mixture_characterization_factors": mixture_rows,
            "factor_audit": factor_audit,
            "quarterly": quarterly_rows,
            "annual": annual_rows,
            "cumulative_2025_2030": cumulative_rows,
            "managed_toxicity_summary": {
                "definition": (
                    "Managed HTN (CTUh) or ECF (CTUe) equals residual parent-PFAS "
                    "emission toxicity plus the corresponding treatment-process burden. "
                    "HTN and ECF remain separate endpoints; HTC is not included."
                ),
                "quarterly": _managed_toxicity_summary(quarterly_rows, "quarterly"),
                "annual": _managed_toxicity_summary(annual_rows, "annual"),
                "cumulative_2025_2030": _managed_toxicity_summary(
                    cumulative_rows, "cumulative_2025_2030"
                ),
            },
            "managed_toxicity_by_material": {
                "definition": (
                    "Residual and treatment-process toxicity are attributed to six material "
                    "sources using each material's medium-specific PFAS mass. The current model "
                    "assumes the same nine-PFAS mixture characterization factors for all sources."
                ),
                "quarterly": quarterly_by_material,
                "annual": annual_by_material,
                "cumulative_2025_2030": cumulative_by_material,
            },
            "managed_toxicity_by_medium": {
                "definition": (
                    "Residual and treatment-process toxicity are attributed to the originating "
                    "gas, solvent, wastewater, or solid stream before treatment."
                ),
                "quarterly": quarterly_by_medium,
                "annual": annual_by_medium,
                "cumulative_2025_2030": cumulative_by_medium,
            },
            "comparison_scope": (
                "IER plus gas, solvent, and solid-waste process burdens are included. "
                "The managed comparison remains screening-level because PIC/PID, treatment-adoption shares, "
                "and exact foreground/background characterization-method harmonization are outside the result."
            ),
            "method_compatibility_note": (
                "PFAS foreground factors and EF v3.1 background treatment factors share CTUh/CTUe units but "
                "do not come from one identical factor library. Component results remain separately reported."
            ),
            "sources": model_input["sources"],
            "primary_scenario_id": primary_scenario_id,
            "quality_checks": quality_checks,
        }
    )


def save_json(payload: dict[str, Any], output_path: str | Path = DEFAULT_OUTPUT) -> Path:
    resolved = Path(output_path).expanduser().resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    resolved.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return resolved


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--attribution", type=Path, default=DEFAULT_ATTRIBUTION)
    parser.add_argument("--ier-burden", type=Path, default=DEFAULT_IER_BURDEN)
    parser.add_argument(
        "--incineration-burden", type=Path, default=DEFAULT_INCINERATION_BURDEN
    )
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args(argv)
    payload = calculate_toxicity(
        args.attribution,
        args.ier_burden,
        args.incineration_burden,
        args.input,
    )
    if not payload["quality_checks"]["all_checks_passed"]:
        raise RuntimeError("Toxicity quality checks failed")
    output = save_json(payload, args.output)
    primary = next(
        row
        for row in payload["cumulative_2025_2030"]
        if row["scenario_id"] == payload["primary_scenario_id"]
    )
    print(
        json.dumps(
            {
                "output": str(output),
                "primary_scenario_id": payload["primary_scenario_id"],
                "cumulative_untreated_noncancer_ctuh": primary[
                    "untreated_noncancer_ctuh_total"
                ],
                "cumulative_residual_noncancer_ctuh": primary[
                    "residual_noncancer_ctuh_total"
                ],
                "cumulative_untreated_freshwater_ctue_hc50": primary[
                    "untreated_freshwater_ctue_hc50_total"
                ],
                "cumulative_residual_freshwater_ctue_hc50": primary[
                    "residual_freshwater_ctue_hc50_total"
                ],
                "all_checks_passed": payload["quality_checks"]["all_checks_passed"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
