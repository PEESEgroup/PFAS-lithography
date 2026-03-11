import numpy as np
import pandas as pd
import xlrd


def run_emission_model(scale_results, input_path='data input/data.xls'):
    excel = xlrd.open_workbook(input_path)
    PFAS_data = excel.sheet_by_index(3)
    Chip_Cooling = excel.sheet_by_index(4)
    EF = excel.sheet_by_index(5)
    Litho = excel.sheet_by_index(7)

    Litho_details = np.array([Litho.row_values(r, 1, 15) for r in range(1, 45)], dtype=float)

    mean_list = scale_results['mean_list']
    std_list = scale_results['std_list']
    mean_R = scale_results['mean_R']
    std_R = scale_results['std_R']

    Total_lith = PFAS_data.col_values(colx=0)
    Total_lith.pop(0)

    Target_Cooling_Share = [x for x in PFAS_data.col_values(colx=1) if x != '']
    Target_Cooling_Share.pop(0)

    Wet_chem = [x for x in PFAS_data.col_values(colx=6) if x != '']
    Wet_chem.pop(0)

    Dry_etching = [x for x in PFAS_data.col_values(colx=7) if x != '']
    Dry_etching.pop(0)

    Plasma_clean = [x for x in PFAS_data.col_values(colx=8) if x != '']
    Plasma_clean.pop(0)

    CVD = [x for x in PFAS_data.col_values(colx=9) if x != '']
    CVD.pop(0)

    SD = [x for x in PFAS_data.col_values(colx=10) if x != '']
    SD.pop(0)

    HTF = [x for x in PFAS_data.col_values(colx=11) if x != '']
    HTF.pop(0)

    Packaging = [x for x in PFAS_data.col_values(colx=12) if x != '']
    Packaging.pop(0)

    Pump_fluid = [x for x in PFAS_data.col_values(colx=13) if x != '']
    Pump_fluid.pop(0)

    Area = np.array(
        [x for x in Chip_Cooling.col_values(colx=4) if x != ''][1:],
        dtype=float,
    )

    Ratio_lith = [x for x in PFAS_data.col_values(colx=4) if x != '']
    Lith_components = np.outer(Total_lith, Ratio_lith)
    columns_lith = ['photoacid generators', 'surfactant', 'Top antireflective coatings', 'polymer', 'immersion topcoat', 'PBO/PI']
    PFAS_Manu = pd.DataFrame(Lith_components, columns=columns_lith)
    PFAS_Manu[
        ['Wet chemistry', 'Dry etching', 'Plasma Cleaning', 'Chemical Vapour Deposition', 'Spin-on dielectrics', 'Fluorinated heat transfer fluids', 'Assembly, packaging', 'Pump fluid, lubricants']
    ] = np.column_stack((
        Wet_chem, Dry_etching, Plasma_clean, CVD, SD, HTF, Packaging, Pump_fluid,
    ))

    DOI_PFAS = [x for x in Chip_Cooling.col_values(colx=2) if x != '']
    DOI_PFAS.pop(0)

    ER_data_center = np.array(
        [x for x in Chip_Cooling.col_values(colx=3) if x != ''][1:],
        dtype=float,
    )

    Target_share = [x for x in Chip_Cooling.col_values(colx=10) if x != '']

    comp_power = [x for x in Chip_Cooling.col_values(colx=1) if x != '']
    comp_power.pop(0)
    N_server = np.array(mean_list, dtype=float) * np.array(comp_power, dtype=float)
    N_server_dev = np.array(std_list, dtype=float) * np.array(comp_power, dtype=float)

    N_server = np.array(N_server, dtype=float).reshape(-1, 1)
    DOI_PFAS = np.array(DOI_PFAS, dtype=float).reshape(-1, 1)
    Target_share = np.array(Target_share, dtype=float).reshape(-1, 1)
    ER_data_center = np.array(ER_data_center, dtype=float).reshape(-1, 1)
    Area = np.array(Area, dtype=float).reshape(-1, 1)

    Chip_lifetime = 3 * 4
    PFAS_Manu_new = PFAS_Manu.mul(mean_list, axis=0) / ((mean_list[4] + mean_list[5] + mean_list[6] + mean_list[7]) / 4)
    PFAS_Manu_dev = PFAS_Manu_new.mul(np.array(std_list) / np.array(mean_list), axis=0)
    PFAS_Manu_new = PFAS_Manu_new.iloc[:, : 6]
    PFAS_Manu_dev = PFAS_Manu_dev.iloc[:, :6]

    Lith_result = np.multiply(Litho_details, PFAS_Manu_new)
    Lith_result_dev = np.multiply(Litho_details, PFAS_Manu_dev)
    Lith_result_total = Lith_result.sum(axis=1)
    Lith_result_dev_total = Lith_result_dev.sum(axis=1)
    Toxicity_data = excel.sheet_by_index(6)
    Toxicity_PFOS = float(Toxicity_data.cell(0, 1).value)
    Toxicity_PFOA = float(Toxicity_data.cell(1, 1).value)
    PFAS_toxicity_PFOS = Toxicity_PFOS * Lith_result_total
    PFAS_toxicity_PFOA = Toxicity_PFOA * Lith_result_total
    PFAS_toxicity_PFOS_dev = Toxicity_PFOS * Lith_result_dev_total
    PFAS_toxicity_PFOA_dev = Toxicity_PFOA * Lith_result_dev_total
    Toxicity_combined = pd.concat([PFAS_toxicity_PFOA, PFAS_toxicity_PFOA_dev, PFAS_toxicity_PFOS, PFAS_toxicity_PFOS_dev], axis=1)
    Manu_EF_gas = Lith_result * EF.row_values(rowx=1)[1:]
    Manu_EF_solvent = Lith_result * EF.row_values(rowx=2)[1:]
    Manu_EF_water = Lith_result * EF.row_values(rowx=3)[1:]
    Manu_EF_solid = Lith_result * EF.row_values(rowx=4)[1:]

    Manu_EF_gas_dev = Lith_result_dev * EF.row_values(rowx=1)[1:]
    Manu_EF_solvent_dev = Lith_result_dev * EF.row_values(rowx=2)[1:]
    Manu_EF_water_dev = Lith_result_dev * EF.row_values(rowx=3)[1:]
    Manu_EF_solid_dev = Lith_result_dev * EF.row_values(rowx=4)[1:]

    Lith_result_combined = pd.concat([Lith_result, Lith_result_dev], axis=1)
    Manu_EF_gas_combined = pd.concat([Manu_EF_gas, Manu_EF_gas_dev], axis=1)
    Manu_EF_solvent_combined = pd.concat([Manu_EF_solvent, Manu_EF_solvent_dev], axis=1)
    Manu_EF_water_combined = pd.concat([Manu_EF_water, Manu_EF_water_dev], axis=1)
    Manu_EF_solid_combined = pd.concat([Manu_EF_solid, Manu_EF_solid_dev], axis=1)
    R_combined = pd.concat([pd.Series(mean_R), pd.Series(std_R)], axis=1)
    return {
        'Manu_EF_gas': Manu_EF_gas,
        'Manu_EF_gas_dev': Manu_EF_gas_dev,
        'Manu_EF_water': Manu_EF_water,
        'Manu_EF_water_dev': Manu_EF_water_dev,
        'Manu_EF_solvent': Manu_EF_solvent,
        'Manu_EF_solvent_dev': Manu_EF_solvent_dev,
        'Manu_EF_solid': Manu_EF_solid,
        'Manu_EF_solid_dev': Manu_EF_solid_dev,
        'Toxicity_combined': Toxicity_combined,
        'Lith_result_combined': Lith_result_combined,
        'Manu_EF_gas_combined': Manu_EF_gas_combined,
        'Manu_EF_solvent_combined': Manu_EF_solvent_combined,
        'Manu_EF_water_combined': Manu_EF_water_combined,
        'Manu_EF_solid_combined': Manu_EF_solid_combined,
        'R_combined': R_combined,
    }
