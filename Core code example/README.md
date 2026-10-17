# Main calculation code

`PFAS usage-burden prediction.py` is the entry point. It calls `main.py`, which connects the following modules in order.

| Module | Calculation |
|---|---|
| `demand_equations.py` | Training and inference compute demand |
| `uncertainty_projection.py` | Shared-path sampling, retained compute, purchases and cohort retirement |
| `wafer_intensity/pfas_wafer_intensity.py` | Wafer calibration, component allocation and technology-node intensity |
| `attribution/genai_pfas_attribution.py` | New GPU-equivalents to quarterly PFAS use and four waste streams |
| `ier/ier_model.py` | Resin loading and capacity-utilization cases |
| `ier/ier_burden.py` | Captured PFAS, resin production and spent-mixture incineration |
| `treatment/incineration_burden.py` | Solvent, gas and solid-waste treatment burdens |
| `toxicity/toxicity.py` | Mixture factors, untreated toxicity, residual toxicity and managed totals |
| `economic/economic_model.py` | Work-breakdown-structure unit costs applied to current mass flows |
| `uncertainty_analysis.py` | Quarterly, annual and cumulative statistics from consistent simulation paths |

The scientific modules are drawn from the current study code. Repository paths and imports are adapted for this layout. The local spreadsheet-formatting and in-place publication utilities are separate from this calculation chain.
