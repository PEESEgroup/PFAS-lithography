"""Excel-driven PFAS photolithography intensity for one GPU-equivalent package.

The current main-result boundary is intentionally limited to four component
groups: GPU logic dies, HBM DRAM dies, HBM base dies, and the silicon
interposer/RDL equivalent.  Rack, platform, storage, and scale-out networking
silicon are outside this module's main result.  Component geometry determines
wafer demand, while one shared quarterly wafer-intensity path represents node
evolution across the whole package.  Component-specific node assignments are
not used in the central result.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import pandas as pd


SCENARIOS = ("Low", "Central", "High")
REQUIRED_COMPONENT_IDS = ("C01", "C02", "C03", "C04")
MAIN_START_YEAR = 2025
MAIN_END_YEAR = 2030


def _default_paths() -> tuple[Path, Path]:
    project_root = Path(__file__).resolve().parents[2]
    return (
        project_root / "Data input" / "PFAS_wafer_intensity_input.xlsx",
        project_root / "Excel results" / "wafer_intensity" / "PFAS_wafer_intensity_results.json",
    )


def _read_configuration(input_path: Path) -> dict[str, object]:
    frame = pd.read_excel(input_path, sheet_name="Configuration", header=2)
    required = {"Parameter", "Value", "Unit"}
    if not required.issubset(frame.columns):
        raise ValueError(f"Configuration sheet must contain {sorted(required)}")
    return dict(zip(frame["Parameter"], frame["Value"]))


def _validate_component_table(components: pd.DataFrame) -> pd.DataFrame:
    required_columns = {"ID", "Component", "Boundary group", "Evidence class", "Source ID", "Notes"}
    for scenario in SCENARIOS:
        required_columns.update(
            {
                f"Count/GPU {scenario}",
                f"Die area {scenario} (mm2)",
                f"Yield {scenario}",
                f"Node {scenario}",
            }
        )
    missing = sorted(required_columns.difference(components.columns))
    if missing:
        raise ValueError(f"Components sheet is missing columns: {missing}")

    selected = components.loc[components["ID"].isin(REQUIRED_COMPONENT_IDS)].copy()
    ids = tuple(selected["ID"].tolist())
    if ids != REQUIRED_COMPONENT_IDS:
        raise ValueError(
            "Main-result boundary must contain C01-C04 exactly and in order; "
            f"found {ids}"
        )
    return selected


def _quarter_index(year: int, quarter: int) -> int:
    return (year - MAIN_START_YEAR) * 4 + quarter - 1


def _quarter_label(year: int, quarter: int) -> str:
    return f"{year} Q{quarter}"


def _build_quarterly_intensity_path(
    roadmap: pd.DataFrame,
    node_layers: dict[str, float],
    node_intensity_g_per_300mm: dict[str, float],
) -> list[dict[str, object]]:
    """Interpolate PFAS intensity, never the nominal node number.

    Annual roadmap labels remain discrete.  Each milestone is an intensity
    anchor at Q1, except the final 2030 A14 endpoint, which is reached in Q4 as
    recorded in the roadmap evidence note.
    """

    active = roadmap.loc[
        roadmap["Milestone year"].between(MAIN_START_YEAR, MAIN_END_YEAR)
    ].copy()
    if active.empty:
        raise ValueError("Node_Roadmap has no milestones for 2025-2030")
    active = active.sort_values("Milestone year")
    required_years = {2025, 2026, 2028, 2030}
    present_years = set(active["Milestone year"].astype(int))
    if not required_years.issubset(present_years):
        raise ValueError(
            "Node_Roadmap must contain the 2025, 2026, 2028 and 2030 intensity anchors"
        )

    final_year = int(active["Milestone year"].max())
    anchors: list[dict[str, object]] = []
    for _, row in active.iterrows():
        year = int(row["Milestone year"])
        quarter = 4 if year == final_year else 1
        node_key = str(row["Intensity node key"])
        anchors.append(
            {
                "index": _quarter_index(year, quarter),
                "period": _quarter_label(year, quarter),
                "year": year,
                "quarter": quarter,
                "node_label": str(row["Node label"]),
                "node_key": node_key,
                "pfas_layer_proxy": float(node_layers[node_key]),
                "intensity_g_per_300mm_wafer": float(
                    node_intensity_g_per_300mm[node_key]
                ),
            }
        )

    schedule: list[dict[str, object]] = []
    last_index = _quarter_index(MAIN_END_YEAR, 4)
    for index in range(last_index + 1):
        year = MAIN_START_YEAR + index // 4
        quarter = index % 4 + 1
        exact = next((anchor for anchor in anchors if anchor["index"] == index), None)
        if exact is not None:
            lower = exact
            upper = exact
            fraction = 0.0
            label = str(exact["node_label"])
        else:
            lower = max(
                (anchor for anchor in anchors if anchor["index"] < index),
                key=lambda anchor: int(anchor["index"]),
            )
            upper = min(
                (anchor for anchor in anchors if anchor["index"] > index),
                key=lambda anchor: int(anchor["index"]),
            )
            fraction = (index - int(lower["index"])) / (
                int(upper["index"]) - int(lower["index"])
            )
            label = (
                f"{lower['node_label']} to {upper['node_label']} "
                "PFAS-intensity interpolation"
            )
        lower_intensity = float(lower["intensity_g_per_300mm_wafer"])
        upper_intensity = float(upper["intensity_g_per_300mm_wafer"])
        intensity = lower_intensity + fraction * (upper_intensity - lower_intensity)
        lower_layers = float(lower["pfas_layer_proxy"])
        upper_layers = float(upper["pfas_layer_proxy"])
        layer_proxy = lower_layers + fraction * (upper_layers - lower_layers)
        schedule.append(
            {
                "period": _quarter_label(year, quarter),
                "year": year,
                "quarter": quarter,
                "roadmap_label": label,
                "from_node": str(lower["node_label"]),
                "to_node": str(upper["node_label"]),
                "interpolation_fraction": fraction,
                "interpolated_pfas_layer_proxy": layer_proxy,
                "wafer_intensity_g_per_300mm": intensity,
                "method": "Linear interpolation of PFAS intensity; node labels remain annual/discrete",
            }
        )
    return schedule


def calculate(input_path: str | Path) -> dict[str, object]:
    path = Path(input_path).expanduser().resolve()
    config = _read_configuration(path)
    nodes = pd.read_excel(path, sheet_name="Node_Intensity", header=2)
    roadmap = pd.read_excel(path, sheet_name="Node_Roadmap", header=2)
    components = _validate_component_table(pd.read_excel(path, sheet_name="Components", header=2))

    required_node_columns = {
        "Node", "PFAS-layer proxy", "Evidence status", "Source ID", "Gap / limitation"
    }
    missing_node_columns = sorted(required_node_columns.difference(nodes.columns))
    if missing_node_columns:
        raise ValueError(f"Node_Intensity sheet is missing columns: {missing_node_columns}")

    required_roadmap_columns = {
        "Milestone year", "Node label", "Intensity node key", "PFAS proxy status",
        "Roadmap evidence status", "Source ID", "Implemented use", "Gap / limitation",
    }
    missing_roadmap_columns = sorted(required_roadmap_columns.difference(roadmap.columns))
    if missing_roadmap_columns:
        raise ValueError(f"Node_Roadmap sheet is missing columns: {missing_roadmap_columns}")

    reference_layers = float(config["Reference PFAS-layer count"])
    reference_kg_per_200mm = float(config["Calibrated reference kg/200-mm-eq wafer"])
    area_ratio = float(config["300/200 wafer area ratio"])
    wafer_diameter = float(config["Wafer diameter (mm)"])
    wafer_area = math.pi * (wafer_diameter / 2.0) ** 2

    node_layers = dict(zip(nodes["Node"], nodes["PFAS-layer proxy"]))
    node_intensity_g_per_300mm = {
        node: float(layer_count) / reference_layers * reference_kg_per_200mm * area_ratio * 1000.0
        for node, layer_count in node_layers.items()
    }
    absent_roadmap_nodes = sorted(
        set(roadmap["Intensity node key"].dropna().astype(str)).difference(node_intensity_g_per_300mm)
    )
    if absent_roadmap_nodes:
        raise ValueError(
            "Node_Roadmap contains intensity keys absent from Node_Intensity: "
            f"{absent_roadmap_nodes}"
        )

    node_input_rows = []
    for _, node_row in nodes.iterrows():
        node_name = str(node_row["Node"])
        node_input_rows.append(
            {
                "node": node_name,
                "pfas_layer_proxy": float(node_row["PFAS-layer proxy"]),
                "intensity_g_per_300mm_wafer": node_intensity_g_per_300mm[node_name],
                "evidence_status": str(node_row["Evidence status"]),
                "source_id": str(node_row["Source ID"]),
                "gap_or_limitation": str(node_row["Gap / limitation"]),
            }
        )

    roadmap_output = []
    for _, route_row in roadmap.iterrows():
        intensity_key = str(route_row["Intensity node key"])
        roadmap_output.append(
            {
                "milestone_year": int(route_row["Milestone year"]),
                "node_label": str(route_row["Node label"]),
                "intensity_node_key": intensity_key,
                "pfas_layer_proxy": float(node_layers[intensity_key]),
                "intensity_g_per_300mm_wafer": node_intensity_g_per_300mm[intensity_key],
                "pfas_proxy_status": str(route_row["PFAS proxy status"]),
                "roadmap_evidence_status": str(route_row["Roadmap evidence status"]),
                "source_id": str(route_row["Source ID"]),
                "implemented_use": str(route_row["Implemented use"]),
                "gap_or_limitation": str(route_row["Gap / limitation"]),
            }
        )

    geometry_rows: list[dict[str, object]] = []
    wafer_equivalents_by_scenario: dict[str, float] = {}
    for scenario in SCENARIOS:
        total_wafer_equivalents = 0.0
        for _, component in components.iterrows():
            count = float(component[f"Count/GPU {scenario}"])
            die_area = float(component[f"Die area {scenario} (mm2)"])
            yield_fraction = float(component[f"Yield {scenario}"])
            if not (0 < yield_fraction <= 1):
                raise ValueError(f"{component['ID']} {scenario} yield must be in (0, 1]")

            gross_dies_per_wafer = wafer_area / die_area - math.pi * wafer_diameter / math.sqrt(2.0 * die_area)
            good_dies_per_wafer = gross_dies_per_wafer * yield_fraction
            if good_dies_per_wafer <= 0:
                raise ValueError(f"{component['ID']} {scenario} gives non-positive good dies per wafer")
            wafer_equivalents = count / good_dies_per_wafer
            total_wafer_equivalents += wafer_equivalents
            geometry_rows.append(
                {
                    "Scenario": scenario,
                    "ID": component["ID"],
                    "Component": component["Component"],
                    "Boundary group": component["Boundary group"],
                    "Count/GPU": count,
                    "Die area (mm2)": die_area,
                    "Yield": yield_fraction,
                    "Gross dies/wafer": gross_dies_per_wafer,
                    "Good dies/wafer": good_dies_per_wafer,
                    "Wafer equivalents/GPU": wafer_equivalents,
                    "Evidence class": component["Evidence class"],
                    "Source ID": component["Source ID"],
                    "Notes": component["Notes"],
                }
            )
        wafer_equivalents_by_scenario[scenario] = total_wafer_equivalents

    quarterly_path = _build_quarterly_intensity_path(
        roadmap=roadmap,
        node_layers={str(key): float(value) for key, value in node_layers.items()},
        node_intensity_g_per_300mm=node_intensity_g_per_300mm,
    )
    central_wafer_equivalents = wafer_equivalents_by_scenario["Central"]
    for row in quarterly_path:
        row["central_wafer_equivalents_per_gpu"] = central_wafer_equivalents
        row["central_pfas_mg_per_gpu"] = (
            central_wafer_equivalents
            * float(row["wafer_intensity_g_per_300mm"])
            * 1000.0
        )

    first_intensity = float(quarterly_path[0]["wafer_intensity_g_per_300mm"])
    sensitivity_memo = {
        scenario: wafer_equivalents * first_intensity * 1000.0
        for scenario, wafer_equivalents in wafer_equivalents_by_scenario.items()
    }

    return {
        "schema_version": "2.0",
        "input_workbook": str(path),
        "functional_unit": "mg photolithography PFAS input per newly manufactured GPU-equivalent",
        "main_boundary": "GPU logic + HBM DRAM + HBM base dies + silicon interposer/RDL equivalent",
        "excluded_from_main": "Grace CPU, local storage, rack/network/control/power silicon and scale-out network",
        "central_node_method": (
            "All included component wafer demand uses one shared quarterly wafer PFAS intensity. "
            "Component-specific node assignments are not used in the central result."
        ),
        "node_interpretation": (
            "DUV/EUV mixed-process node/layer proxy. Node labels are estimation inputs, "
            "not claims of a single lithography technology or product-specific measured consumption."
        ),
        "advanced_node_extension": (
            "2 nm, 2 nm+ and 14 Å use the 3-nm direct proxy multiplied by the legacy "
            "four-category total-use ratio relative to 3 nm. These are bridged estimates."
        ),
        "node_intensity_inputs": node_input_rows,
        "node_roadmap": roadmap_output,
        "quarterly_node_intensity": quarterly_path,
        "wafer_equivalents_per_gpu": wafer_equivalents_by_scenario,
        "central_wafer_equivalents_per_gpu": central_wafer_equivalents,
        "summary_mg_per_gpu": sensitivity_memo,
        "sensitivity_status": (
            "Central is active. Low and High are provisional geometry-based memo values and "
            "remain unfinished sensitivity parameters."
        ),
        "rows": geometry_rows,
    }


def main(argv: list[str] | None = None) -> int:
    default_input, default_output = _default_paths()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=default_input)
    parser.add_argument("--output", type=Path, default=default_output)
    args = parser.parse_args(argv)

    payload = calculate(args.input)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Input workbook: {args.input.resolve()}")
    print(f"Output JSON: {args.output.resolve()}")
    print(
        "Central wafer equivalents/GPU: "
        f"{payload['central_wafer_equivalents_per_gpu']:.9f}"
    )
    print(
        "Central 2025 Q1: "
        f"{payload['quarterly_node_intensity'][0]['central_pfas_mg_per_gpu']:.6f} mg/GPU"
    )
    print(
        "Central 2030 Q4: "
        f"{payload['quarterly_node_intensity'][-1]['central_pfas_mg_per_gpu']:.6f} mg/GPU"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
