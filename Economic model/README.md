# Economic model

Management costs are calculated by connecting work-breakdown-structure unit costs to the updated waste-stream and resin mass flows.

The implementation is [`economic_model.py`](../Core%20code%20example/economic/economic_model.py). It reads these workbooks from this directory:

- `IER-wastewater-.xlsm`: wastewater treatment and ion-exchange cost components.
- `Incineration-solvent.xlsm`: incineration cost components.
- `economic cost.xlsx`: summary values used for consistency checks.

The cost structure is retained while manufacturing attribution and treatment mass flows use the updated model. Results are organized by waste stream, lithography material, quarter, year and cumulative period. The workbooks will be added separately.
