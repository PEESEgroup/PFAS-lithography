"""Capacity-utilization sensitivity model for ion-exchange resin (IER).

The model uses an Excel-managed EPA-informed loading proxy and represents
concentration, matrix and kinetic uncertainty with a dimensionless capacity-
utilization factor ``f_C``::

    q_eff = q_ref * f_C

``q_ref`` is the reference loading (0.1 g/g) and ``0 < f_C <= 1``.
The scenarios are sensitivity bounds, not an empirical concentration-to-
capacity relationship and not confidence intervals.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INPUT = PROJECT_ROOT / "Data input" / "treatment" / "IER_model_input.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "Excel results" / "ier" / "IER_model_results.json"

REQUIRED_SHEETS = {"Assumptions", "Sensitivity_Scenarios", "Sources"}


class IERInputError(ValueError):
    """Raised when the editable IER workbook violates its input contract."""


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_json_safe(item) for item in value]
    if isinstance(value, tuple):
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


def _as_bool(value: Any, field: str) -> bool:
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    if isinstance(value, (int, float)) and not pd.isna(value):
        return bool(value)
    normalized = str(value).strip().lower()
    if normalized in {"true", "yes", "y", "1"}:
        return True
    if normalized in {"false", "no", "n", "0"}:
        return False
    raise IERInputError(f"{field} must be TRUE/FALSE; received {value!r}")


def _require_columns(frame: pd.DataFrame, sheet: str, columns: set[str]) -> None:
    missing = columns.difference(frame.columns)
    if missing:
        raise IERInputError(f"{sheet} is missing columns: {sorted(missing)}")


def _read_input(input_path: Path) -> dict[str, Any]:
    if not input_path.exists():
        raise FileNotFoundError(f"IER input workbook not found: {input_path}")

    book = pd.ExcelFile(input_path, engine="openpyxl")
    missing_sheets = REQUIRED_SHEETS.difference(book.sheet_names)
    if missing_sheets:
        raise IERInputError(f"Workbook is missing sheets: {sorted(missing_sheets)}")

    assumptions_frame = pd.read_excel(book, sheet_name="Assumptions")
    scenarios = pd.read_excel(book, sheet_name="Sensitivity_Scenarios")
    sources = pd.read_excel(book, sheet_name="Sources")

    _require_columns(assumptions_frame, "Assumptions", {"parameter", "value"})
    _require_columns(
        scenarios,
        "Sensitivity_Scenarios",
        {
            "scenario_id",
            "capacity_utilization_factor_f_c",
            "role",
            "enabled",
        },
    )
    _require_columns(sources, "Sources", {"source_id", "reference"})

    assumptions: dict[str, Any] = {}
    for row in assumptions_frame.itertuples(index=False):
        parameter = str(row.parameter).strip()
        if not parameter or parameter.lower() == "nan":
            continue
        if parameter in assumptions:
            raise IERInputError(f"Duplicate assumption parameter: {parameter}")
        assumptions[parameter] = row.value

    required_parameters = {
        "reference_capacity_g_pfas_per_g_ier",
        "influent_pfas_basis_kg",
        "target_capture_fraction",
        "reference_scenario_id",
    }
    missing_parameters = required_parameters.difference(assumptions)
    if missing_parameters:
        raise IERInputError(f"Missing assumptions: {sorted(missing_parameters)}")

    for parameter in {
        "reference_capacity_g_pfas_per_g_ier",
        "influent_pfas_basis_kg",
        "target_capture_fraction",
    }:
        try:
            assumptions[parameter] = float(assumptions[parameter])
        except (TypeError, ValueError) as exc:
            raise IERInputError(f"Assumption {parameter} must be numeric") from exc

    assumptions["reference_scenario_id"] = str(
        assumptions["reference_scenario_id"]
    ).strip()
    if assumptions["reference_capacity_g_pfas_per_g_ier"] <= 0:
        raise IERInputError("reference_capacity_g_pfas_per_g_ier must be positive")
    if assumptions["influent_pfas_basis_kg"] <= 0:
        raise IERInputError("influent_pfas_basis_kg must be positive")
    if not 0 < assumptions["target_capture_fraction"] <= 1:
        raise IERInputError("target_capture_fraction must be in (0, 1]")

    scenarios = scenarios.copy()
    scenarios["enabled"] = [
        _as_bool(value, f"Sensitivity_Scenarios.enabled row {index + 2}")
        for index, value in enumerate(scenarios["enabled"])
    ]
    scenarios = scenarios.loc[scenarios["enabled"]].reset_index(drop=True)
    if scenarios.empty:
        raise IERInputError("At least one sensitivity scenario must be enabled")
    if scenarios["scenario_id"].astype(str).duplicated().any():
        raise IERInputError("Enabled scenario_id values must be unique")

    scenarios["capacity_utilization_factor_f_c"] = pd.to_numeric(
        scenarios["capacity_utilization_factor_f_c"], errors="raise"
    )
    if not scenarios["capacity_utilization_factor_f_c"].between(
        0, 1, inclusive="right"
    ).all():
        raise IERInputError("Every enabled f_C must satisfy 0 < f_C <= 1")

    primary = scenarios.loc[scenarios["role"].astype(str) == "primary"]
    if len(primary) != 1:
        raise IERInputError("Exactly one enabled scenario must have role='primary'")

    reference_id = assumptions["reference_scenario_id"]
    reference_rows = scenarios.loc[scenarios["scenario_id"].astype(str) == reference_id]
    if len(reference_rows) != 1:
        raise IERInputError(
            f"reference_scenario_id {reference_id!r} is not an enabled scenario"
        )
    reference_factor = float(
        reference_rows.iloc[0]["capacity_utilization_factor_f_c"]
    )
    if not math.isclose(reference_factor, 1.0, rel_tol=0, abs_tol=1e-12):
        raise IERInputError("The reference scenario must use f_C = 1")
    if str(reference_rows.iloc[0]["role"]) != "primary":
        raise IERInputError("The reference scenario must also be the primary scenario")

    return {
        "assumptions": assumptions,
        "assumptions_frame": assumptions_frame,
        "scenarios": scenarios,
        "sources": sources,
    }


def run_ier_model(input_path: str | Path = DEFAULT_INPUT, *, comparison_reference_capacity_g_g=None) -> dict[str, Any]:
    """Run the Excel-driven IER capacity-utilization sensitivity model."""

    resolved_input = Path(input_path).expanduser().resolve()
    model_input = _read_input(resolved_input)
    assumptions = model_input["assumptions"]

    q_reference = assumptions["reference_capacity_g_pfas_per_g_ier"]
    comparison_q = q_reference if comparison_reference_capacity_g_g is None else float(comparison_reference_capacity_g_g)
    if not math.isfinite(comparison_q) or comparison_q <= 0:
        raise IERInputError('Comparison reference capacity must be finite and positive')
    influent_kg = assumptions["influent_pfas_basis_kg"]
    capture_fraction = assumptions["target_capture_fraction"]
    captured_kg = influent_kg * capture_fraction
    bypass_kg = influent_kg - captured_kg

    results: list[dict[str, Any]] = []
    for scenario in model_input["scenarios"].itertuples(index=False):
        factor = float(scenario.capacity_utilization_factor_f_c)
        effective_capacity = q_reference * factor
        resin_kg = captured_kg / effective_capacity
        resin_per_influent = resin_kg / influent_kg
        resin_per_captured = resin_kg / captured_kg
        spent_resin_mix_kg = resin_kg + captured_kg
        mass_balance_error = influent_kg - captured_kg - bypass_kg

        results.append(
            {
                "scenario_id": str(scenario.scenario_id),
                "role": str(scenario.role),
                "capacity_utilization_factor_f_c": factor,
                "effective_capacity_g_pfas_per_g_ier": effective_capacity,
                "influent_pfas_basis_kg": influent_kg,
                "capture_fraction": capture_fraction,
                "captured_pfas_kg": captured_kg,
                "bypass_pfas_to_water_kg": bypass_kg,
                "ier_mass_kg": resin_kg,
                "ier_kg_per_kg_influent_pfas": resin_per_influent,
                "ier_kg_per_kg_captured_pfas": resin_per_captured,
                "spent_ier_plus_captured_pfas_kg": spent_resin_mix_kg,
                "relative_ier_demand_vs_reference": comparison_q / effective_capacity,
                "comparison_reference_capacity_g_g": comparison_q,
                "mass_balance_error_kg": mass_balance_error,
                "concentration_interpretation": str(
                    getattr(scenario, "concentration_interpretation", "")
                ),
                "notes": str(getattr(scenario, "notes", "")),
            }
        )

    primary = next(row for row in results if row["role"] == "primary")
    max_mass_balance_error = max(abs(row["mass_balance_error_kg"]) for row in results)

    return _json_safe(
        {
            "model_name": "IER capacity-utilization sensitivity model",
            "model_scope": (
                "Excel-defined effective loading and utilization scenarios. The common 0.1 g/g "
                "benchmark is an EPA-informed PFOS mixture proxy; strategy values are conditional "
                "effective loadings, not a facility-specific concentration-capacity fit."
            ),
            "equation": "q_eff = q_ref * f_C; m_IER = m_PFAS,captured / q_eff",
            "input_path": str(resolved_input),
            "assumptions": assumptions,
            "primary_scenario_id": primary["scenario_id"],
            "sensitivity_results": results,
            "sources": model_input["sources"].to_dict(orient="records"),
            "quality_checks": {
                "enabled_scenario_count": len(results),
                "reference_factor_equals_one": math.isclose(
                    primary["capacity_utilization_factor_f_c"],
                    1.0,
                    rel_tol=0,
                    abs_tol=1e-12,
                ),
                "max_mass_balance_error_kg": max_mass_balance_error,
                "all_effective_capacities_positive": all(
                    row["effective_capacity_g_pfas_per_g_ier"] > 0
                    for row in results
                ),
            },
        }
    )


def save_json(payload: dict[str, Any], output_path: str | Path = DEFAULT_OUTPUT) -> Path:
    """Write IER mass-sensitivity results and return the resolved output path."""

    resolved_output = Path(output_path).expanduser().resolve()
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    resolved_output.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return resolved_output


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run the Excel-driven IER capacity-utilization sensitivity model."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    payload = run_ier_model(args.input)
    output = save_json(payload, args.output)
    primary = next(
        row
        for row in payload["sensitivity_results"]
        if row["scenario_id"] == payload["primary_scenario_id"]
    )
    print(
        json.dumps(
            {
                "output": str(output),
                "primary_scenario_id": payload["primary_scenario_id"],
                "primary_effective_capacity_g_g": primary[
                    "effective_capacity_g_pfas_per_g_ier"
                ],
                "primary_ier_kg_per_kg_influent_pfas": primary[
                    "ier_kg_per_kg_influent_pfas"
                ],
                "scenario_count": len(payload["sensitivity_results"]),
                "max_mass_balance_error_kg": payload["quality_checks"][
                    "max_mass_balance_error_kg"
                ],
            },
            indent=2,
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
