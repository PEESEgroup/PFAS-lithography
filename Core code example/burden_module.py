import numpy as np
import pandas as pd
import xlrd


def run_burden_model(emission_results, input_path2=r'data input/destruction summary.xls'):
    Manu_EF_gas = emission_results['Manu_EF_gas']
    Manu_EF_gas_dev = emission_results['Manu_EF_gas_dev']
    Manu_EF_water = emission_results['Manu_EF_water']
    Manu_EF_water_dev = emission_results['Manu_EF_water_dev']
    Manu_EF_solvent = emission_results['Manu_EF_solvent']
    Manu_EF_solvent_dev = emission_results['Manu_EF_solvent_dev']
    Manu_EF_solid = emission_results['Manu_EF_solid']
    Manu_EF_solid_dev = emission_results['Manu_EF_solid_dev']

    excel2 = xlrd.open_workbook(input_path2)
    GAC_data = excel2.sheet_by_index(0)
    IER_data = excel2.sheet_by_index(1)
    RO_data = excel2.sheet_by_index(2)
    Incineration_solid = excel2.sheet_by_index(3)
    Incineration_solvent = excel2.sheet_by_index(4)

    GAC_burden = [GAC_data.row_values(r, 1, 20) for r in range(1, GAC_data.nrows)]
    GAC_error = [GAC_data.row_values(r, 22, 41) for r in range(1, GAC_data.nrows)]
    GAC_burden = np.array(GAC_burden, dtype=float) * 1.e3
    GAC_error = np.array(GAC_error, dtype=float) * 1.e3

    IER_burden = [IER_data.row_values(r, 1, 20) for r in range(1, IER_data.nrows)]
    IER_error = [IER_data.row_values(r, 22, 41) for r in range(1, IER_data.nrows)]
    IER_burden = np.array(IER_burden, dtype=float) * 1.e3
    IER_error = np.array(IER_error, dtype=float) * 1.e3

    RO_burden = [RO_data.row_values(r, 1, 20) for r in range(1, RO_data.nrows)]
    RO_error = [RO_data.row_values(r, 22, 41) for r in range(1, RO_data.nrows)]
    RO_burden = np.array(RO_burden, dtype=float) * 1.e3
    RO_error = np.array(RO_error, dtype=float) * 1.e3

    Incineration_burden1 = [Incineration_solid.row_values(r, 1, 20) for r in range(1, Incineration_solid.nrows)]
    Incineration_error1 = [Incineration_solid.row_values(r, 22, 41) for r in range(1, Incineration_solid.nrows)]
    Incineration_burden1 = np.array(Incineration_burden1, dtype=float) * 1.e3
    Incineration_error1 = np.array(Incineration_error1, dtype=float) * 1.e3

    Incineration_burden2 = [Incineration_solvent.row_values(r, 1, 20) for r in range(1, Incineration_solvent.nrows)]
    Incineration_error2 = [Incineration_solvent.row_values(r, 22, 41) for r in range(1, Incineration_solvent.nrows)]
    Incineration_burden2 = np.array(Incineration_burden2, dtype=float) * 1.e3
    Incineration_error2 = np.array(Incineration_error2, dtype=float) * 1.e3

    Treat_gas = Incineration_burden1[:, :, None] * Manu_EF_gas.to_numpy()[:, None, :]
    Treat_gas_error = (Incineration_error1[:, :, None] + Incineration_burden1[:, :, None]) * Manu_EF_gas_dev.to_numpy()[:, None, :]
    Treat_gas_total = Treat_gas.sum(axis=2)
    Treat_gas_error_total = Treat_gas_error.sum(axis=2)

    Treat_water = IER_burden[:, :, None] * Manu_EF_water.to_numpy()[:, None, :]
    Treat_water_error = (IER_error[:, :, None] + IER_burden[:, :, None]) * Manu_EF_water_dev.to_numpy()[:, None, :]
    Treat_water_total = Treat_water.sum(axis=2)
    Treat_water_error_total = Treat_water_error.sum(axis=2)

    Treat_solvent = Incineration_burden2[:, :, None] * Manu_EF_solvent.to_numpy()[:, None, :]
    Treat_solvent_error = (Incineration_error2[:, :, None] + Incineration_burden2[:, :, None]) * Manu_EF_solvent_dev.to_numpy()[:, None, :]
    Treat_solvent_total = Treat_solvent.sum(axis=2)
    Treat_solvent_error_total = Treat_solvent_error.sum(axis=2)

    Treat_solid = Incineration_burden1[:, :, None] * Manu_EF_solid.to_numpy()[:, None, :]
    Treat_solid_error = (Incineration_error1[:, :, None] + Incineration_burden1[:, :, None]) * Manu_EF_solid_dev.to_numpy()[:, None, :]
    Treat_solid_total = Treat_solid.sum(axis=2)
    Treat_solid_error_total = Treat_solid_error.sum(axis=2)

    Stacked_litho = np.stack([Treat_gas, Treat_solid, Treat_solvent, Treat_water], axis=0)
    Stacked_litho_error = np.stack([Treat_gas_error, Treat_solid_error, Treat_solvent_error, Treat_water_error], axis=0)

    PAG_burden = Stacked_litho[:, :, :, 0].sum(axis=0)
    PAG_burden_error = Stacked_litho_error[:, :, :, 0].sum(axis=0)
    Sur_burden = Stacked_litho[:, :, :, 1].sum(axis=0)
    Sur_burden_error = Stacked_litho_error[:, :, :, 1].sum(axis=0)
    TARC_burden = Stacked_litho[:, :, :, 2].sum(axis=0)
    TARC_burden_error = Stacked_litho_error[:, :, :, 2].sum(axis=0)
    Polymer_burden = Stacked_litho[:, :, :, 3].sum(axis=0)
    Polymer_burden_error = Stacked_litho_error[:, :, :, 3].sum(axis=0)
    ITC_burden = Stacked_litho[:, :, :, 4].sum(axis=0)
    ITC_burden_error = Stacked_litho_error[:, :, :, 4].sum(axis=0)
    PBO_burden = Stacked_litho[:, :, :, 5].sum(axis=0)
    PBO_burden_error = Stacked_litho_error[:, :, :, 5].sum(axis=0)

    Treat_gas_combined = pd.concat([pd.DataFrame(Treat_gas_total), pd.DataFrame(Treat_gas_error_total)], axis=1)
    Treat_solid_combined = pd.concat([pd.DataFrame(Treat_solid_total), pd.DataFrame(Treat_solid_error_total)], axis=1)
    Treat_water_combined = pd.concat([pd.DataFrame(Treat_water_total), pd.DataFrame(Treat_water_error_total)], axis=1)
    Treat_solvent_combined = pd.concat([pd.DataFrame(Treat_solvent_total), pd.DataFrame(Treat_solvent_error_total)], axis=1)

    PAG_combined = pd.concat([pd.DataFrame(PAG_burden), pd.DataFrame(PAG_burden_error)], axis=1)
    Sur_combined = pd.concat([pd.DataFrame(Sur_burden), pd.DataFrame(Sur_burden_error)], axis=1)
    TARC_combined = pd.concat([pd.DataFrame(TARC_burden), pd.DataFrame(TARC_burden_error)], axis=1)
    Polymer_combined = pd.concat([pd.DataFrame(Polymer_burden), pd.DataFrame(Polymer_burden_error)], axis=1)
    ITC_combined = pd.concat([pd.DataFrame(ITC_burden), pd.DataFrame(ITC_burden_error)], axis=1)
    PBO_combined = pd.concat([pd.DataFrame(PBO_burden), pd.DataFrame(PBO_burden_error)], axis=1)

    return {
        'Treat_gas_combined': Treat_gas_combined,
        'Treat_solid_combined': Treat_solid_combined,
        'Treat_solvent_combined': Treat_solvent_combined,
        'Treat_water_combined': Treat_water_combined,
        'PAG_combined': PAG_combined,
        'Sur_combined': Sur_combined,
        'TARC_combined': TARC_combined,
        'Polymer_combined': Polymer_combined,
        'ITC_combined': ITC_combined,
        'PBO_combined': PBO_combined,
    }
