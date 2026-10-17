# Model input map

The model reads the following study workbooks. This code edition documents their interfaces; Excel files are prepared separately.

| Repository-relative location | Contents and consuming module |
|---|---|
| `Data input/legacy_data.xls` | AI demand, performance, retirement, six material shares and emission routing; `uncertainty_projection` and `attribution` |
| `Data input/PFAS_wafer_intensity_input.xlsx` | `Configuration`, `Node_Intensity`, `Node_Roadmap` and `Components`; wafer intensity and manufacturing attribution |
| `Data input/Uncertainty_model_input.xlsx` | `Settings`: runs, seed, parameter coefficients of variation and temporal sampling rules |
| `Data input/treatment/IER_model_input.xlsx` | `Assumptions`, `Sensitivity_Scenarios` and `Sources`; resin loading and utilization |
| `Data input/treatment/treatment_destruction_summary_new.xls` | Treatment-factor reference blocks |
| `Data input/treatment/treatment_Database-IAM_scenarios_Basic_2025.xlsx` | Baseline treatment life-cycle inventory |
| `Data input/treatment/treatment_PFAS_destruction_NDI_2030.xlsx` | 2030 national-policy background |
| `Data input/treatment/treatment_PFAS_destruction_PkDuge500_2030.xlsx` | 2030 PkDuge500 background |
| `Data input/treatment/treatment_PFAS_destruction_PkDuge1150_2030.xlsx` | 2030 PkDuge1150 background |
| `Data input/treatment/treatment_PFAS_destruction_SSP5-BASE_2030.xlsx` | 2030 SSP5 baseline background |
| `Data input/toxicity/PFAS_toxicity_input.xlsx` | `Assumptions`, `Composition`, `Substance_Factors`, `Sources`; mixture toxicity and residual emissions |

Economic workbook locations are listed in [`Economic model/README.md`](../Economic%20model/README.md).

The nine-substance mixture consists of PFBA, PFHxA, PFHpA, PFOA, PFNA, PFDA, GenX, PFBS and PFOS. `Substance_Factors` stores endpoint- and emission-compartment-specific factors; `Composition` supplies the mass fractions. The agent workflow supports the evidence and characterization calculations behind these inputs.

Workbook-reading functions in each module specify the exact sheet names, column names, fixed-cell locations and unit conversions used by the calculation.
