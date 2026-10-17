"""Attribute photolithography PFAS use to newly manufactured GPU-equivalents.

The AI-scale calculation retains the legacy demand equations. Quarterly new
servers are converted to GPU-equivalents, multiplied by quarterly central PFAS
intensity and each run's manufacturing multiplier, then summed within runs.
GPU stock and retirement flows are
not multiplied by manufacturing intensity as separate drivers.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

import numpy as np
import pandas as pd
import xlrd


MATERIAL_NAMES = (
    "photoacid generators",
    "surfactant",
    "top antireflective coatings",
    "polymer",
    "immersion topcoat",
    "PBO/PI",
)
MEDIA_NAMES = ("gas", "solvent", "wastewater", "solid")
START_YEAR = 2025
END_YEAR = 2030


def _project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _default_paths() -> tuple[Path, Path, Path]:
    root = _project_root()
    return (
        root / "Data input" / "legacy_data.xls",
        root / "Data input" / "PFAS_wafer_intensity_input.xlsx",
        root / "Excel results" / "attribution" / "GenAI_PFAS_attribution_results.json",
    )


def _period_label(decimal_year: float) -> tuple[str, int, int]:
    year = int(np.floor(decimal_year + 1e-9))
    quarter = int(round((decimal_year - year) * 4)) + 1
    if quarter not in (1, 2, 3, 4):
        raise ValueError(f"Invalid decimal quarter {decimal_year}")
    return f"{year} Q{quarter}", year, quarter


def _read_legacy_inputs(path: Path) -> dict[str, object]:
    workbook = xlrd.open_workbook(str(path))
    config = workbook.sheet_by_name("input Configuration")
    computation = workbook.sheet_by_name("Computation Data.")
    pfas = workbook.sheet_by_name("PFAS data")
    lithography = workbook.sheet_by_name("lithography details")
    emission = workbook.sheet_by_name("Emission factor")

    gpu_per_server = float(config.cell(3, 1).value)
    decimal_periods = np.asarray(computation.col_values(11, 1), dtype=float)

    base_labels = [value for value in pfas.col_values(3) if value != ""]
    base_shares = np.asarray(
        [value for value in pfas.col_values(4) if isinstance(value, (int, float))],
        dtype=float,
    )
    if len(base_labels) != 6 or len(base_shares) != 6:
        raise ValueError("PFAS data must provide exactly six legacy material shares")

    lithography_periods = np.asarray(lithography.col_values(0, 1), dtype=float)
    lithography_multipliers = np.asarray(
        [lithography.row_values(row, 1, 7) for row in range(1, lithography.nrows)],
        dtype=float,
    )
    if lithography_multipliers.shape != (len(lithography_periods), 6):
        raise ValueError("lithography details must provide six material multipliers per quarter")

    emission_factors = np.asarray(
        [emission.row_values(row, 1, 7) for row in range(1, 5)],
        dtype=float,
    )
    if emission_factors.shape != (4, 6):
        raise ValueError("Emission factor must provide four media by six materials")

    return {
        "gpu_per_server": gpu_per_server,
        "decimal_periods": decimal_periods,
        "base_labels": base_labels,
        "base_shares": base_shares,
        "lithography_periods": lithography_periods,
        "lithography_multipliers": lithography_multipliers,
        "emission_factors": emission_factors,
    }


def calculate_attribution(
    scale_results: dict[str, object],
    legacy_input_path: str | Path,
    wafer_input_path: str | Path,
    *, pag_replacement_fraction: float = 0.0,
) -> dict[str, object]:
    if not np.isfinite(pag_replacement_fraction) or not 0 <= pag_replacement_fraction <= 1:
        raise ValueError('PAG replacement fraction must be in [0, 1]')
    legacy_path = Path(legacy_input_path).expanduser().resolve()
    wafer_path = Path(wafer_input_path).expanduser().resolve()
    legacy = _read_legacy_inputs(legacy_path)

    from wafer_intensity.pfas_wafer_intensity import calculate as calculate_wafer_intensity

    wafer_result = calculate_wafer_intensity(wafer_path)
    quarterly_intensity = {
        row["period"]: row for row in wafer_result["quarterly_node_intensity"]
    }

    mean_servers = np.asarray(scale_results["mean_list"], dtype=float)
    sd_servers = np.asarray(scale_results["std_list"], dtype=float)
    decimal_periods = np.asarray(legacy["decimal_periods"], dtype=float)
    if len(mean_servers) != len(decimal_periods) or len(sd_servers) != len(decimal_periods):
        raise ValueError("AI-scale output length must match legacy quarterly periods")

    lithography_periods = np.asarray(legacy["lithography_periods"], dtype=float)
    lithography_multipliers = np.asarray(legacy["lithography_multipliers"], dtype=float)
    multiplier_by_period = {
        round(float(period), 2): lithography_multipliers[index]
        for index, period in enumerate(lithography_periods)
    }
    base_shares = np.asarray(legacy["base_shares"], dtype=float)
    emission_factors = np.asarray(legacy["emission_factors"], dtype=float)
    gpu_per_server = float(legacy["gpu_per_server"])

    server_samples = scale_results.get("samples")
    pfas_samples = []
    gpu_samples = []

    quarterly_rows: list[dict[str, object]] = []
    material_share_rows: list[dict[str, object]] = []
    for index, decimal_period in enumerate(decimal_periods):
        period, year, quarter = _period_label(float(decimal_period))
        if not START_YEAR <= year <= END_YEAR:
            continue
        if period not in quarterly_intensity:
            raise ValueError(f"Missing quarterly wafer intensity for {period}")
        if round(float(decimal_period), 2) not in multiplier_by_period:
            raise ValueError(f"Missing legacy lithography composition for {period}")

        intensity_row = quarterly_intensity[period]
        new_servers = float(mean_servers[index])
        new_server_sd = float(sd_servers[index])
        new_gpus = new_servers * gpu_per_server
        new_gpu_sd = new_server_sd * gpu_per_server
        pfas_mg_per_gpu = float(intensity_row["central_pfas_mg_per_gpu"])
        pfas_input_kg = new_gpus * pfas_mg_per_gpu / 1.0e6
        pfas_sd = None
        if server_samples is not None:
            quarterly_gpu = np.asarray(server_samples)[:, index] * gpu_per_server
            quarterly_pfas = (quarterly_gpu * pfas_mg_per_gpu / 1.0e6
                              * np.asarray(scale_results["manufacturing_multipliers"]))
            gpu_samples.append(quarterly_gpu)
            pfas_samples.append(quarterly_pfas)
            pfas_input_kg = float(quarterly_pfas.mean())
            pfas_sd = float(quarterly_pfas.std(ddof=1))

        composition_raw = base_shares * multiplier_by_period[round(float(decimal_period), 2)]
        composition = composition_raw / composition_raw.sum()
        material_kg = pfas_input_kg * composition
        original_pfas_kg = pfas_input_kg
        removed_pag_kg = float(material_kg[0] * pag_replacement_fraction)
        remaining_fraction = 1.0 - float(composition[0]) * pag_replacement_fraction
        if remaining_fraction <= 0:
            raise ValueError('PAG substitution leaves no reference mass for downstream propagation')
        # First allocate the baseline material masses, then remove only PAG PFAS.
        # Renormalizing the original shares before allocation would redistribute it.
        material_kg[0] *= 1.0 - pag_replacement_fraction
        pfas_input_kg *= remaining_fraction
        composition = material_kg / pfas_input_kg
        if server_samples is not None:
            pfas_samples[-1] = quarterly_pfas * remaining_fraction
            pfas_sd *= remaining_fraction
        media_by_material_kg = emission_factors * material_kg[np.newaxis, :]
        media_kg = media_by_material_kg.sum(axis=1)

        row: dict[str, object] = {
            "Period": period,
            "Year": year,
            "Quarter": quarter,
            "New servers (mean)": new_servers,
            "New servers (SD)": new_server_sd,
            "GPUs per server": gpu_per_server,
            "Total new GPU-equivalents": new_gpus,
            "New GPU-equivalents (quarter SD)": new_gpu_sd,
            "Roadmap label": intensity_row["roadmap_label"],
            "From node": intensity_row["from_node"],
            "To node": intensity_row["to_node"],
            "Intensity interpolation fraction": float(
                intensity_row["interpolation_fraction"]
            ),
            "Wafer intensity (g PFAS/300-mm wafer)": float(
                intensity_row["wafer_intensity_g_per_300mm"]
            ),
            "Central wafer equivalents/GPU-equivalent": float(
                intensity_row["central_wafer_equivalents_per_gpu"]
            ),
            "Central PFAS intensity (mg/GPU-equivalent)": pfas_mg_per_gpu,
            "PFAS input (kg)": pfas_input_kg,
            "PFAS input (t)": pfas_input_kg / 1000.0,
            "PFAS before PAG replacement (kg)": original_pfas_kg,
            "PAG PFAS replaced (kg)": removed_pag_kg,
        }
        if pfas_sd is not None:
            row["PFAS input (quarter SD kg)"] = pfas_sd
        for material_index, material in enumerate(MATERIAL_NAMES):
            row[f"Material | {material} (kg)"] = float(material_kg[material_index])
        for media_index, medium in enumerate(MEDIA_NAMES):
            row[f"Manufacturing flow | {medium} (kg)"] = float(media_kg[media_index])
            for material_index, material in enumerate(MATERIAL_NAMES):
                row[f"Manufacturing flow | {medium} | {material} (kg)"] = float(
                    media_by_material_kg[media_index, material_index]
                )
        quarterly_rows.append(row)

        share_row: dict[str, object] = {"Period": period, "Year": year, "Quarter": quarter}
        for material_index, material in enumerate(MATERIAL_NAMES):
            share_row[f"{material} share"] = float(composition[material_index])
        share_row["Share sum"] = float(composition.sum())
        material_share_rows.append(share_row)

    quarterly_frame = pd.DataFrame(quarterly_rows)
    if len(quarterly_frame) != 24:
        raise AssertionError(f"Expected 24 quarters for 2025-2030, found {len(quarterly_frame)}")
    if not np.allclose(
        quarterly_frame[[f"Material | {name} (kg)" for name in MATERIAL_NAMES]].sum(axis=1),
        quarterly_frame["PFAS input (kg)"],
        rtol=1e-12,
        atol=1e-9,
    ):
        raise AssertionError("Quarterly material mass does not reconcile to PFAS input")

    additive_columns = [
        "Total new GPU-equivalents",
        "PFAS input (kg)",
        "PFAS input (t)",
        "PFAS before PAG replacement (kg)",
        "PAG PFAS replaced (kg)",
        *[f"Material | {name} (kg)" for name in MATERIAL_NAMES],
        *[f"Manufacturing flow | {medium} (kg)" for medium in MEDIA_NAMES],
        *[
            f"Manufacturing flow | {medium} | {material} (kg)"
            for medium in MEDIA_NAMES
            for material in MATERIAL_NAMES
        ],
    ]
    annual_frame = quarterly_frame.groupby("Year", as_index=False)[additive_columns].sum()
    annual_frame.insert(1, "Quarter count", quarterly_frame.groupby("Year").size().to_numpy())
    annual_frame.insert(
        3,
        "GPU-weighted central intensity (mg/GPU-equivalent)",
        quarterly_frame.assign(weighted_central=(
            quarterly_frame["Total new GPU-equivalents"]
            * quarterly_frame["Central PFAS intensity (mg/GPU-equivalent)"]
        )).groupby("Year")["weighted_central"].sum().to_numpy()
        / annual_frame["Total new GPU-equivalents"],
    )
    annual_frame["Cumulative PFAS input 2025 onward (t)"] = annual_frame[
        "PFAS input (t)"
    ].cumsum()
    if pfas_samples:
        # Retain run identities for downstream conditional linear propagation.
        scale_results["pfas_samples_kg"] = np.stack(pfas_samples, axis=1)
        scale_results["gpu_samples_reporting"] = np.stack(gpu_samples, axis=1)
        annual_samples = scale_results["pfas_samples_kg"].reshape(-1, 6, 4).sum(axis=2)
        annual_frame["PFAS input (annual SD kg)"] = annual_samples.std(axis=0, ddof=1)
        annual_frame["PFAS input (annual CV)"] = annual_samples.std(axis=0, ddof=1) / annual_samples.mean(axis=0)
    if not (annual_frame["Quarter count"] == 4).all():
        raise AssertionError("Each annual result must contain exactly four quarters")
    if not np.isclose(
        annual_frame["PFAS input (kg)"].sum(),
        quarterly_frame["PFAS input (kg)"].sum(),
        rtol=1e-12,
        atol=1e-9,
    ):
        raise AssertionError("Annual PFAS mass does not reconcile to quarterly results")

    return {
        "schema_version": "1.1",
        "legacy_ai_scale_input": str(legacy_path),
        "wafer_intensity_input": str(wafer_path),
        "period": "2025 Q1-2030 Q4",
        "functional_unit": "Annual photolithography PFAS input attributable to newly manufactured GPU-equivalents",
        "scale_core": "Legacy demand equations with normal input sampling, nonnegative purchases and retained idle compute",
        "uncertainty_method": (
            "PFAS mass is averaged after multiplying each run's quarterly new GPUs by central intensity "
            "and that run's shared manufacturing multiplier; annual SD uses within-run quarterly sums."
            if server_samples is not None else "No sample-level uncertainty supplied"
        ),
        "deployment_scenario": scale_results.get("scenario", "legacy"),
        "gpu_equivalent_interpretation": (
            "B300-derived package geometry is a GPU-equivalent manufacturing proxy, not a claim "
            "that every 2025-2030 product is the same physical GPU."
        ),
        "stock_and_retirement_treatment": (
            "GPU stock and retirements are not multiplied by manufacturing intensity. "
            "Replacement purchases already represented in new additions remain included."
        ),
        "node_method": (
            "Annual node labels are discrete. Quarterly PFAS wafer intensity is linearly "
            "interpolated between 2025 Q1, 2026 Q1, 2028 Q1 and 2030 Q4 anchors."
        ),
        "material_method": (
            "Six legacy material shares and quarterly relative multipliers are retained and "
            "normalized before allocation. PAG replacement then removes absolute PAG PFAS "
            "mass without changing the other five material masses; emission factors are unchanged."
        ),
        "material_media_method": (
            "Each material mass is multiplied by its medium-specific legacy emission factor. "
            "The resulting 4-media by 6-material matrix is retained in every quarterly row."
        ),
        "pag_replacement_fraction": float(pag_replacement_fraction),
        "central_wafer_equivalents_per_gpu": float(
            wafer_result["central_wafer_equivalents_per_gpu"]
        ),
        "quarterly": quarterly_frame.to_dict(orient="records"),
        "annual": annual_frame.to_dict(orient="records"),
        "material_shares": material_share_rows,
        "component_geometry": wafer_result["rows"],
        "checks": {
            "quarter_count": len(quarterly_frame),
            "annual_year_count": len(annual_frame),
            "material_mass_reconciles": True,
            "annual_quarterly_mass_reconciles": True,
            "media_allocation_note": (
                "Legacy emission factors are used without normalization; mass retained in "
                "equipment is outside the four modeled media flows."
            ),
        },
    }


def save_json(payload: dict[str, object], output_path: str | Path) -> Path:
    path = Path(output_path).expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def main(argv: list[str] | None = None) -> int:
    default_legacy, default_wafer, default_output = _default_paths()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--legacy-input", type=Path, default=default_legacy)
    parser.add_argument("--wafer-input", type=Path, default=default_wafer)
    parser.add_argument("--output", type=Path, default=default_output)
    parser.add_argument("--seed", type=int, default=None)
    args = parser.parse_args(argv)

    src_dir = _project_root() / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))
    from server_projection_module import run_server_projection

    scale_results = run_server_projection(str(args.legacy_input), seed=args.seed)
    payload = calculate_attribution(scale_results, args.legacy_input, args.wafer_input)
    payload["legacy_scale_random_seed"] = scale_results['seed']
    output = save_json(payload, args.output)
    print(f"Output JSON: {output}")
    for row in payload["annual"]:
        print(
            f"{row['Year']}: {row['Total new GPU-equivalents']:,.0f} new GPU-equivalents, "
            f"{row['PFAS input (t)']:.6f} t PFAS"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
