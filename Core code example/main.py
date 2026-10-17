"""Active GenAI PFAS manufacturing, treatment, toxicity, and economic workflow.

Stage 1 keeps the legacy AI demand equations and replaces the old 2021-baseline
proportional attribution with quarterly new GPU-equivalents multiplied by the
central quarterly PFAS intensity. Stage 2 applies the 90% wastewater capture
assumption and rebuilds IER burdens by separately scaling resin production and
spent-mixture incineration across the f_C sensitivity cases. Stage 3 reads the
authoritative toxicity Excel input, calculates one nine-PFAS mixture factor per
endpoint and receiving compartment, and applies it to untreated and residual
aggregate PFAS flows. Stage 3 also restores gas, solvent, and solid-waste
incineration burdens without applying the destruction efficiency twice. Stage 4
reconnects the unchanged legacy WBS economic models to the active PFAS and IER
mass flows. Normal path-wise parameter biases, retained idle compute, and a shared
manufacturing multiplier propagate full Monte Carlo paths through all stages.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from attribution.genai_pfas_attribution import calculate_attribution, save_json
from economic.economic_model import (
    calculate_economic_results,
    save_json as save_economic_json,
)
from ier.ier_burden import calculate_ier_burden, save_json as save_ier_burden_json
from ier.ier_model import run_ier_model, save_json as save_ier_model_json
from uncertainty_projection import run_server_projection
from uncertainty_projection import DEFAULT_UNCERTAINTY_INPUT
from uncertainty_analysis import write_uncertainty_results
from toxicity.toxicity import calculate_toxicity, save_json as save_toxicity_json
from treatment.incineration_burden import (
    calculate_incineration_burden,
    save_json as save_incineration_burden_json,
)


def main(argv: list[str] | None = None) -> int:
    project_root = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--scenario', choices=('aggressive','moderate','conservative'), default='aggressive')
    parser.add_argument('--results-root', type=Path, help='Isolated result directory for a complete scenario run.')
    parser.add_argument('--current-scenario-only', action='store_true', help='Write uncertainty for the selected deployment only.')
    parser.add_argument(
        "--legacy-input",
        type=Path,
        default=project_root / "Data input" / "legacy_data.xls",
    )
    parser.add_argument(
        "--wafer-input",
        type=Path,
        default=project_root / "Data input" / "PFAS_wafer_intensity_input.xlsx",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=project_root
        / "Excel results"
        / "attribution"
        / "GenAI_PFAS_attribution_results.json",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=None,
        help="Override the random seed in Uncertainty_model_input.xlsx.",
    )
    parser.add_argument(
        "--ier-input",
        type=Path,
        default=project_root / "Data input" / "treatment" / "IER_model_input.xlsx",
    )
    parser.add_argument(
        "--ier-mass-output",
        type=Path,
        default=project_root / "Excel results" / "ier" / "IER_model_results.json",
    )
    parser.add_argument(
        "--ier-burden-output",
        type=Path,
        default=project_root / "Excel results" / "ier" / "IER_burden_results.json",
    )
    parser.add_argument(
        "--toxicity-input",
        type=Path,
        default=project_root / "Data input" / "toxicity" / "PFAS_toxicity_input.xlsx",
        help=(
            "Authoritative nine-PFAS composition, raw compartment-specific factors, "
            "and toxicity assumptions."
        ),
    )
    parser.add_argument(
        "--toxicity-output",
        type=Path,
        default=project_root / "Excel results" / "toxicity" / "PFAS_toxicity_results.json",
    )
    parser.add_argument(
        "--incineration-burden-output",
        type=Path,
        default=project_root
        / "Excel results"
        / "treatment"
        / "Incineration_burden_results.json",
    )
    parser.add_argument(
        "--economic-output",
        type=Path,
        default=project_root / "Excel results" / "economic" / "Economic_results.json",
    )
    parser.add_argument('--uncertainty-input', type=Path, default=DEFAULT_UNCERTAINTY_INPUT)
    parser.add_argument('--uncertainty-output-dir', type=Path, default=project_root/'Excel results'/'uncertainty')
    args = parser.parse_args(argv)

    if args.results_root is not None:
        for field, relative in {
            'output': 'attribution/GenAI_PFAS_attribution_results.json',
            'ier_mass_output': 'ier/IER_model_results.json',
            'ier_burden_output': 'ier/IER_burden_results.json',
            'toxicity_output': 'toxicity/PFAS_toxicity_results.json',
            'incineration_burden_output': 'treatment/Incineration_burden_results.json',
            'economic_output': 'economic/Economic_results.json',
            'uncertainty_output_dir': 'uncertainty',
        }.items():
            setattr(args, field, args.results_root / relative)

    scale_results = run_server_projection(str(args.legacy_input), uncertainty_input=args.uncertainty_input, seed=args.seed, scenario=args.scenario)
    attribution = calculate_attribution(
        scale_results=scale_results,
        legacy_input_path=args.legacy_input,
        wafer_input_path=args.wafer_input,
    )
    attribution["legacy_scale_random_seed"] = scale_results['seed']
    attribution_output = save_json(attribution, args.output)

    ier_model_results = run_ier_model(args.ier_input)
    ier_mass_output = save_ier_model_json(ier_model_results, args.ier_mass_output)
    ier_burden = calculate_ier_burden(
        attribution_results=attribution,
        ier_model_results=ier_model_results,
        ier_input_path=args.ier_input,
    )
    if not ier_burden["quality_checks"]["all_checks_passed"]:
        raise RuntimeError("IER burden regression or mass-balance checks failed")
    ier_burden_output = save_ier_burden_json(ier_burden, args.ier_burden_output)

    incineration_burden = calculate_incineration_burden(
        attribution_results=attribution,
        legacy_input_path=args.legacy_input,
    )
    if not incineration_burden["quality_checks"]["all_checks_passed"]:
        raise RuntimeError("Other-stream incineration burden checks failed")
    incineration_burden_output = save_incineration_burden_json(
        incineration_burden, args.incineration_burden_output
    )

    toxicity = calculate_toxicity(
        attribution_results=attribution,
        ier_burden_results=ier_burden,
        incineration_burden_results=incineration_burden,
        toxicity_input_path=args.toxicity_input,
    )
    if not toxicity["quality_checks"]["all_checks_passed"]:
        raise RuntimeError("Toxicity mixture or mass-balance checks failed")
    toxicity_output = save_toxicity_json(toxicity, args.toxicity_output)

    economic = calculate_economic_results(
        attribution_results=attribution,
        ier_burden_results=ier_burden,
        legacy_input_path=args.legacy_input,
    )
    if not economic["quality_checks"]["all_checks_passed"]:
        raise RuntimeError("Economic interface checks failed")
    economic_output = save_economic_json(economic, args.economic_output)

    uncertainty_output = write_uncertainty_results(
        scale=scale_results, attribution=attribution, ier=ier_burden,
        incineration=incineration_burden, toxicity=toxicity, economic=economic,
        legacy_input=args.legacy_input, wafer_input=args.wafer_input,
        output_dir=args.uncertainty_output_dir,
        scenarios=[args.scenario] if args.current_scenario_only else None,
    )
    print(f"Sample-level uncertainty output: {uncertainty_output}")

    print(f"Active manufacturing attribution output: {attribution_output}")
    print(f"Active IER mass-sensitivity output: {ier_mass_output}")
    print(f"Active IER treatment-burden output: {ier_burden_output}")
    print(f"Active other-stream incineration-burden output: {incineration_burden_output}")
    print(f"Active PFAS toxicity output: {toxicity_output}")
    print(f"Active economic output: {economic_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
