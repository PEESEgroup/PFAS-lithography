# Model outputs

The main entry writes structured results under the selected result directory. The repository retains the historical `Excel results` folder name; this edition exposes the numerical model outputs directly as JSON and NumPy files.

| Subdirectory | Main output |
|---|---|
| `attribution/` | `GenAI_PFAS_attribution_results.json`: manufacturing PFAS and routed waste flows |
| `ier/` | `IER_model_results.json` and `IER_burden_results.json`: resin requirements and treatment burdens |
| `treatment/` | `Incineration_burden_results.json`: solvent, gas and solid-waste treatment |
| `toxicity/` | `PFAS_toxicity_results.json`: mixture factors, untreated and residual toxicity, managed totals |
| `economic/` | `Economic_results.json`: management costs |
| `uncertainty/` | Statistics and compressed simulation arrays for quarterly, annual and cumulative results |

The study's formatted Excel tables and figure workbooks are prepared from these outputs. Numerical output datasets will be added separately.
