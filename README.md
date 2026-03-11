# PFAS Usage-Burden Model (Lithography-Focused)

## 1. Project Overview
This project models PFAS use and environmental burden in a lithography-focused pathway linked to server/GPU scale growth.

The workflow is organized into three computational stages:
1. **Server projection**: estimates quarterly server-scale growth with Monte Carlo sampling.
2. **Emission model**: converts scale into lithography PFAS usage, then maps usage to emissions across media.
3. **Burden model**: applies treatment/destruction burden factors to emissions and calculates burden outputs.

The model is configured to run from **2020Q1 to 2030Q4** (as defined in the input data), and outputs Excel files for downstream analysis and plotting.

## 2. Repository Structure
At the top level:

- `Core code example/`
  - Main and module code for the full pipeline.
- `Data input/`
  - Required input Excel files used by the model.
- `Excel results/`
  - All model output Excel files are written here.
- `Data for figures/`
  - Curated data prepared for manuscript plotting.
- `Premise-IAM model/`
  - Materials/scripts related to IAM-driven background database updating using premise.
- `Economic model/`
  - Additional analysis resources for economic components.

## 3. Core Code Layout
Inside `Core code example/`:

- `PFAS usage-burden prediction.py`
  - **Main entry point**. Orchestrates all modules and controls all Excel exports.
- `server_projection_module.py`
  - Server projection Monte Carlo simulation and scale statistics.
- `emission_module.py`
  - Lithography PFAS usage and media-specific emission calculations.
- `burden_module.py`
  - Treatment/destruction burden calculations by medium and by process/component.

## 4. Input Data
Inside `Data input/`:

- `data.xls`
  - Main input workbook for server projection, lithography PFAS settings, emission factors, and toxicity inputs.
- `destruction summary.xls`
  - Burden factor and uncertainty inputs for treatment/destruction pathways.

## 5. Output Files
All outputs are written to `Excel results/`:

- `Result.xlsx`
  - Core server projection summary (`mean`, `std`, cumulative stats, outflow stats).
- `PFAS_Litho_2021_2030.xlsx`
  - Lithography usage and emissions by medium, plus quarterly computing power summary.
- `toxicity-PFOA-PFOS.xlsx`
  - Toxicity summary outputs.
- `burden-type.xlsx`
  - Burden results by treatment medium (gas/solid/solvent/water).
- `burden-process.xlsx`
  - Burden results by lithography component/process category.

## 6. How to Run
Run from the project root directory:

```bash
python "Core code example/PFAS usage-burden prediction.py"
```

If your environment uses `python3`, run:

```bash
python3 "Core code example/PFAS usage-burden prediction.py"
```

On successful execution, the script prints:

```text
done
```

## 7. Environment Requirements
Typical Python dependencies:

- `numpy`
- `pandas`
- `xlrd`

Install example:

```bash
pip install numpy pandas xlrd
```

## 8. Model Logic Notes
- Current scope is **lithography-focused PFAS usage and burden**.
- The integrated pipeline follows:
  - `server projection -> emission -> burden`
- Existing formulas, uncertainty handling, and computational order are preserved in modular form.
- Main script controls output writing to keep I/O management centralized.

## 9. Data for Figures
`Data for figures/` contains curated, post-processed data used for manuscript plotting.

This folder includes the datasets for **Figure 2 through Figure 5**, and together they cover the core quantitative results presented in the paper.

## 10. Premise-IAM Model
`Premise-IAM model/` documents and supports how IAM scenario information is integrated through the **premise** tool to update the background LCI database.

Core concept:
- Information from different IAM scenarios is harmonized via premise.
- These scenario signals are used to update the ecoinvent database.
- Activity-level results are adjusted over time according to socioeconomic development and energy system transitions.

For premise details, see:
- https://github.com/polca/premise

## 11. Troubleshooting
- If input files are not found:
  - Confirm `Data input/data.xls` and `Data input/destruction summary.xls` exist.
- If output files are missing:
  - Confirm script completes and prints `done`.
  - Check write permissions for `Excel results/`.
- If import errors occur:
  - Install missing packages in the active Python environment.

## 12. Suggested Workflow
1. Update assumptions and parameters in `Data input/*.xls`.
2. Run the main script.
3. Review outputs in `Excel results/`.
4. Use `Data for figures/` for final plotting and manuscript figures.
