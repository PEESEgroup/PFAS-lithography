# PFAS Usage and Burden Model for Generative-AI Semiconductor Manufacturing

## 1. Project Overview

This project links the expansion of generative artificial intelligence (AI) to the use, emissions, toxicity and management burdens of per- and polyfluoroalkyl substances (PFAS) in semiconductor photolithography. The model translates training and inference demand into new GPU-equivalent manufacturing, calculates the associated PFAS requirements, and follows PFAS through manufacturing waste streams and treatment pathways.

The analysis covers six photolithography material classes: photoacid generators, surfactants, top antireflective coatings, resist polymer-bound PFAS, immersion topcoats, and polybenzoxazole/polyimide materials. Emissions are allocated to wastewater, solvent waste, gaseous waste and solid waste. Environmental and economic calculations evaluate their management through ion-exchange capture and incineration.

The time scale of the results: 2025–2030. Aggressive, moderate and conservative AI deployment scenarios use the same calculation framework with different user and model-growth trajectories.

This repository contains the main calculation modules and the core decision chain of the PFAS toxicity agent. Input Excel workbooks, numerical result datasets and figure workbooks are being prepared separately; their locations and roles are documented in the corresponding directories.

## 2. Model Workflow

The main calculation proceeds through six connected stages:

1. **Project computing demand and GPU procurement.** Training and inference demand are calculated for each quarter. Surviving computing capacity and cohort retirement determine the additional servers and GPU-equivalents needed to meet demand.
2. **Attribute PFAS use to manufacturing.** Newly manufactured GPU-equivalents are multiplied by quarterly PFAS intensity derived from wafer calibration and technology-node assumptions. The resulting mass is distributed among six photolithography material classes.
3. **Calculate manufacturing emissions.** Material-specific emission fractions route PFAS into wastewater, solvent, gas and solid waste, connecting manufacturing demand to treatment requirements and environmental releases.
4. **Evaluate treatment and residual PFAS.** Wastewater capture determines captured PFAS and ion-exchange-resin requirements. Treatment modules calculate resin-related and incineration burdens; capture and destruction efficiencies determine residual parent-PFAS emissions.
5. **Characterize toxicity and management costs.** Substance- and emission-compartment-specific factors are combined into nine-PFAS mixture factors and applied to untreated and residual emissions. Work-breakdown-structure unit costs are applied to waste and resin flows.
6. **Aggregate results and propagate uncertainty.** Consistent Monte Carlo paths carry demand and manufacturing variation through the calculation. Quarterly contributions are summed within each simulation before annual and cumulative statistics are calculated.

The agent supports toxicity characterization through a separate parameter-preparation sequence: **resolve chemical identity → acquire and assess evidence → select parameters and identify gaps → direct targeted follow-up → calculate ready endpoints**. Selected characterization factors enter the main model through a fixed toxicity input workbook, allowing the same parameter set to be used across deployment scenarios.

## 3. Repository Structure

| Directory | Contents |
|---|---|
| `Core code example/` | Main entry point and modules for demand, manufacturing, treatment, toxicity, economics and uncertainty |
| `CoreAgent workflow/` | Agent controller, canonical parameter definitions, evidence-assessment interfaces and structured decision prompts |
| `Data input/` | Input-workbook descriptions, expected locations and connections to model modules |
| `Economic model/` | Economic input-workbook map and work-breakdown-structure cost interface |
| `Excel results/` | Description of model outputs, including JSON summaries and compressed simulation arrays |
| `Data for figures/` | Connections between calculation outputs and manuscript figures |
| `Premise-IAM/` | Scenario-specific background inventories used in treatment calculations |

## 4. Core Code Layout

Inside `Core code example/`, the calculation is organized as follows:

| File or module | Function |
|---|---|
| `PFAS usage-burden prediction.py` | Main command-line entry point |
| `main.py` | Connects the calculation stages and writes module outputs |
| `demand_equations.py` | Training and inference compute-demand equations |
| `uncertainty_projection.py` | Parameter sampling, capacity retention, purchases and cohort retirement |
| `wafer_intensity/pfas_wafer_intensity.py` | Wafer calibration and quarterly technology-node PFAS intensity |
| `attribution/genai_pfas_attribution.py` | GPU-equivalent manufacturing attribution and material-specific waste flows |
| `ier/ier_model.py` | Effective resin loading and capacity-utilization cases |
| `ier/ier_burden.py` | Captured PFAS, resin requirements and treatment-process burdens |
| `treatment/incineration_burden.py` | Incineration burdens for solvent, gaseous and solid waste |
| `toxicity/toxicity.py` | Mixture factors, untreated toxicity, residual toxicity and managed totals |
| `economic/economic_model.py` | Management costs based on mass flows and unit-cost inputs |
| `uncertainty_analysis.py` | Quarterly, annual and cumulative statistics and simulation-array exports |

Start with [`main.py`](Core%20code%20example/main.py) to follow the calculation sequence. Each module contains its input-reading functions, calculations and output structure.

## 5. Input Data

The main model reads parameters and background inventories from Excel workbooks:

| Input | Role |
|---|---|
| `Data input/legacy_data.xls` | AI demand, hardware performance, retirement distributions, material shares and emission fractions |
| `Data input/PFAS_wafer_intensity_input.xlsx` | Wafer calibration, component assumptions, node intensities and technology roadmap |
| `Data input/Uncertainty_model_input.xlsx` | Monte Carlo sample size, random seed, coefficients of variation and temporal sampling rules |
| `Data input/treatment/IER_model_input.xlsx` | Resin loading, utilization cases and treatment assumptions |
| Treatment inventories under `Data input/treatment/` | Baseline and scenario-specific environmental burdens of treatment activities |
| `Data input/toxicity/PFAS_toxicity_input.xlsx` | Mixture composition, substance-specific factors, toxicity assumptions and sources |
| Workbooks under `Economic model/` | Ion-exchange and incineration cost components and summary checks |

The toxicity mixture comprises PFBA, PFHxA, PFHpA, PFOA, PFNA, PFDA, GenX, PFBS and PFOS. Its input workbook separates assumptions, composition, substance factors and source information into named worksheets.

See [`Data input/README.md`](Data%20input/README.md) and [`Economic model/README.md`](Economic%20model/README.md) for the full workbook map. Exact sheet names, columns, cell locations and unit conversions are defined by each module's reading functions.

## 6. Agent-Guided Toxicity Parameters

The toxicity agent coordinates evidence acquisition and parameter assessment for individual PFAS. Its contract includes 25 substance fields covering environmental fate, exposure and effects. Evidence routes include existing records, United States Environmental Protection Agency Computational Toxicology and Exposure services, literature and supporting information, Open Structure–Activity/Property Relationship App (OPERA) predictions, and read-across.

The core loop evaluates available evidence, selects usable parameters and identifies remaining gaps or conflicts. The next action specifies an acquisition route, target parameters and a decision rationale. Returned evidence is assessed again, so intermediate findings determine subsequent retrieval and calculation steps.

Evidence is organized into four tiers: reviewed direct evidence, deterministic derivation, prediction, and read-across or assumptions. Selection retains chemical identity, modeled form, canonical units, source locations and supporting rationale. Endpoint dependencies are checked separately for human cancer toxicity, human non-cancer toxicity and freshwater ecotoxicity.

The files in `CoreAgent workflow/` expose this decision chain:

- [`agent_core.py`](CoreAgent%20workflow/agent_core.py): evidence validation, parameter selection, targeted follow-up and endpoint-specific calculation handoff.
- [`parameter_contract.py`](CoreAgent%20workflow/parameter_contract.py): parameter definitions, units and endpoint dependencies.
- [`decision_prompt.md`](CoreAgent%20workflow/decision_prompt.md): structured evidence-assessment and next-action contracts.

The `Services` interface connects the controller to retrieval, language-model assessment, prediction and deterministic PFAS-adapted USEtox calculation. The shared code focuses on this core chain; service deployment and the complete retrieval application are maintained separately. Further details are in [`CoreAgent workflow/README.md`](CoreAgent%20workflow/README.md).

## 7. Output Files

The main entry writes results to the selected output directory. `Excel results/` retains the repository's familiar folder name; the included calculation code writes JSON and NumPy files for subsequent tabulation and plotting.

| Output | Contents |
|---|---|
| `attribution/GenAI_PFAS_attribution_results.json` | Manufacturing PFAS use, material allocation and routed waste flows |
| `ier/IER_model_results.json` | Resin-loading and capacity-utilization calculations |
| `ier/IER_burden_results.json` | Captured PFAS, resin requirements and treatment burdens |
| `treatment/Incineration_burden_results.json` | Solvent, gas and solid-waste incineration burdens |
| `toxicity/PFAS_toxicity_results.json` | Mixture factors, untreated toxicity, residual toxicity and managed totals |
| `economic/Economic_results.json` | Costs by material, waste stream and reporting period |
| `uncertainty/` | Statistical summaries, sampled trajectories and downstream coefficient arrays |

Outputs retain units, reporting periods and model checks. The study's formatted Excel tables and figure workbooks are prepared from these numerical results. See [`Excel results/README.md`](Excel%20results/README.md) for the output map.

## 8. Environment Requirements

The main model uses Python 3.10 or later with `numpy`, `pandas`, `openpyxl` and `xlrd`. Install dependencies from the repository root:

```bash
python -m pip install -r requirements.txt
```

The agent core uses Python's standard library. External evidence services and the deterministic substance-level calculator connect through its service interfaces.

## 9. How to Run

To inspect the command-line options:

```bash
python "Core code example/PFAS usage-burden prediction.py" --help
```

After placing the input workbooks at their documented locations, run the aggressive deployment scenario:

```bash
python "Core code example/PFAS usage-burden prediction.py" \
  --scenario aggressive --current-scenario-only \
  --results-root "Excel results/aggressive"
```

Use `--scenario moderate` or `--scenario conservative` for another deployment case, with a corresponding output directory. `--current-scenario-only` limits the uncertainty export to that case; without it, the uncertainty module also evaluates the other deployment trajectories. `--seed` overrides the seed specified in the uncertainty workbook.

Default input paths are resolved relative to this repository. `--results-root` places the run's module outputs in a separate directory. On completion, the program prints the generated result paths. Environments using `python3` can substitute it for `python` in these examples.

## 10. Model Logic and Uncertainty

Manufacturing PFAS is attributed to new GPU-equivalent production. Existing hardware stock and retirement influence procurement, while manufacturing intensity is applied to new purchases. Material-specific emission fractions determine the mass entering each waste stream.

Capture and destruction act at different stages. Wastewater capture determines how much PFAS enters the resin-treatment pathway; destruction efficiency determines the parent PFAS remaining after incineration. Toxicity outputs keep residual-emission contributions separate from treatment-process impacts before combining matching endpoints.

Mixture factors are calculated by mass-fraction weighting for each endpoint and receiving compartment. Missing substance factors follow the documented same-endpoint, same-compartment mean-filling rule in the mixture calculation, with the original factor status retained for traceability.

Monte Carlo sampling preserves each parameter's deviation across its quarterly trajectory and reuses shared inputs consistently across equations. Annual and cumulative totals are calculated within each simulation before computing means, sample standard deviations and quantiles. The uncertainty workbook controls sample size and sampling assumptions.

## 11. Figure Data and Background Inventories

`Data for figures/` connects model outputs to the manuscript figures: manufacturing demand and PFAS flows in Figure 1; treatment burdens, agent-assisted toxicity characterization and residual emissions in Figure 2; costs in Figure 3; and strategy comparisons in Figure 4. Separate strategy and sensitivity analyses vary the relevant model assumptions. See [`Data for figures/README.md`](Data%20for%20figures/README.md).

`Premise-IAM/` describes the upstream use of premise and integrated assessment model scenarios to prepare background life-cycle inventories. AI deployment scenarios determine manufacturing and PFAS quantities; background scenarios determine environmental burdens per unit of treatment activity. The main calculation reads the resulting inventory workbooks, so running premise is a separate preparation step. See [`Premise-IAM/README.md`](Premise-IAM/README.md).

## 12. Suggested Workflow

1. Read `main.py` to identify the calculation stages and module connections.
2. Follow the input-workbook map to locate parameters and sources.
3. Read the agent controller alongside the toxicity module to connect evidence selection, substance factors and mixture impacts.
4. With the workbooks in place, run each deployment scenario into its own output directory.
5. Review model checks, mass flows and annual/cumulative statistics before preparing result tables and figures.

If a workbook is not found, check its location against the input map. If a package import fails, install `requirements.txt` in the Python environment used to launch the model. Scientific input checks identify missing fields, inconsistent units or incompatible inputs at the relevant calculation stage.

Code is distributed under the [MIT License](LICENSE).
