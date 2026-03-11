import pandas as pd
from pathlib import Path

from server_projection_module import run_server_projection
from emission_module import run_emission_model
from burden_module import run_burden_model


def main():
    project_root = Path(__file__).resolve().parent.parent
    output_dir = project_root / 'excel results'
    output_dir.mkdir(parents=True, exist_ok=True)

    input_path = str(project_root / 'data input' / 'data.xls')
    output_path = output_dir / 'Result.xlsx'

    scale_results = run_server_projection(input_path=input_path)
    emission_results = run_emission_model(scale_results, input_path=input_path)
    burden_results = run_burden_model(emission_results, input_path2=str(project_root / 'data input' / 'destruction summary.xls'))

    with pd.ExcelWriter(output_dir / 'toxicity-PFOA-PFOS.xlsx') as writer:
        emission_results['Toxicity_combined'].to_excel(writer, sheet_name='PFOAPFOS.xlsx')

    with pd.ExcelWriter(output_dir / 'PFAS_Litho_2021_2030.xlsx') as writer:
        emission_results['Lith_result_combined'].to_excel(writer, sheet_name='Usage')
        emission_results['Manu_EF_gas_combined'].to_excel(writer, sheet_name='Manu_gas')
        emission_results['Manu_EF_solvent_combined'].to_excel(writer, sheet_name='Manu_solvent')
        emission_results['Manu_EF_water_combined'].to_excel(writer, sheet_name='Manu_water')
        emission_results['Manu_EF_solid_combined'].to_excel(writer, sheet_name='Manu_solid')
        emission_results['R_combined'].to_excel(writer, sheet_name='Quarterly computing power')

    with pd.ExcelWriter(output_dir / 'burden-type.xlsx') as writer:
        burden_results['Treat_gas_combined'].to_excel(writer, sheet_name='Treat-gas')
        burden_results['Treat_solid_combined'].to_excel(writer, sheet_name='Treat-solid')
        burden_results['Treat_solvent_combined'].to_excel(writer, sheet_name='Treat-solvent')
        burden_results['Treat_water_combined'].to_excel(writer, sheet_name='Treat-water')

    with pd.ExcelWriter(output_dir / 'burden-process.xlsx') as writer:
        burden_results['PAG_combined'].to_excel(writer, sheet_name='PAG')
        burden_results['Sur_combined'].to_excel(writer, sheet_name='Sur')
        burden_results['TARC_combined'].to_excel(writer, sheet_name='TARC')
        burden_results['Polymer_combined'].to_excel(writer, sheet_name='Polymer')
        burden_results['ITC_combined'].to_excel(writer, sheet_name='ITC')
        burden_results['PBO_combined'].to_excel(writer, sheet_name='PBO')

    exp = pd.DataFrame()
    exp['mean'] = scale_results['mean_list']
    exp['std'] = scale_results['std_list']
    exp['std_cum'] = scale_results['std_cum_list']
    exp['Out'] = scale_results['mean_out']
    exp['std_out'] = scale_results['std_out_list']
    exp['std_cum_out'] = scale_results['std_cum_out_list']

    with pd.ExcelWriter(output_path) as writer:
        exp.to_excel(writer, sheet_name='S')

    print('done')


if __name__ == '__main__':
    main()
