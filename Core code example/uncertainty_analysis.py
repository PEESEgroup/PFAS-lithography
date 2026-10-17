"""Sample-level uncertainty through the fixed, linear PFAS downstream modules.

Quarter-specific coefficients come from freshly evaluated active modules, not
old cached results. Every run retains its own mass in every quarter. This is
exact while the agreed composition, treatment factors and unit costs are fixed.
Annual/cumulative statistics are calculated only after summing within each run.
"""
import json
from pathlib import Path

import numpy as np

from attribution.genai_pfas_attribution import calculate_attribution
from uncertainty_projection import SCENARIO_COLUMNS, run_server_projection


def downstream_coefficients(attribution, ier, incineration, toxicity, economic):
    periods = [r['Period'] for r in attribution['quarterly']]
    index = {p: i for i, p in enumerate(periods)}
    mass = np.array([r['PFAS input (kg)'] for r in attribution['quarterly']])
    if np.any(mass <= 0):
        raise ValueError('Need positive reference-quarter mass to construct downstream coefficients')
    primary = ier['primary_scenario_id']
    coefficients, units = {}, {}
    def add(metric, period, value, unit):
        if metric not in coefficients:
            coefficients[metric] = np.zeros(24)
            units[metric] = unit
        if units[metric] != unit:
            raise ValueError(f'Inconsistent units for {metric}')
        coefficients[metric][index[period]] += float(value) / mass[index[period]]
    for row in attribution['quarterly']:
        add('manufacturing.pfas_input_kg', row['Period'], row['PFAS input (kg)'], 'kg PFAS')
        add('manufacturing.pag_pfas_replaced_kg', row['Period'], row.get('PAG PFAS replaced (kg)', 0), 'kg PFAS')
        for key, value in row.items():
            if key.startswith(('Material | ', 'Manufacturing flow | ')):
                add('manufacturing.'+key, row['Period'], value, 'kg PFAS')
    for row in toxicity['quarterly']:
        if row['scenario_id'] != primary:
            continue
        for key, value in row.items():
            if isinstance(value, (float, int)) and key not in {
                    'year', 'quarter', 'capacity_utilization_factor_f_c', 'incineration_parent_dre'}:
                unit = 'kg PFAS' if key.endswith('_kg') else ('CTUh' if 'noncancer_ctuh' in key else 'CTUe-HC50')
                add('toxicity.'+key, row['period'], value, unit)
    first_impact = ier['impact_categories'][0]['impact_code']
    for row in ier['quarterly_burden']:
        if row['scenario_id'] != primary:
            continue
        if row['impact_code'] == first_impact:
            for key in ('influent_wastewater_pfas_kg', 'captured_pfas_kg', 'bypass_pfas_to_water_kg',
                        'ier_required_kg', 'spent_mix_incinerated_kg'):
                add('ier.'+key, row['period'], row[key], 'kg')
        for key in ('separation_burden', 'spent_mix_incineration_burden',
                    'incineration_excluding_lime_burden', 'pfas_specific_lime_burden', 'total_ier_treatment_burden'):
            add('ier.'+row['impact_code']+'.'+key, row['period'], row[key], row['unit'])
        add('treatment_process.'+row['impact_code'], row['period'], row['total_ier_treatment_burden'], row['unit'])
    for row in incineration['quarterly_burden']:
        add('incineration.'+row['medium']+'.'+row['impact_code'], row['period'], row['process_burden'], row['unit'])
        add('incineration.total.'+row['impact_code'], row['period'], row['process_burden'], row['unit'])
        add('treatment_process.'+row['impact_code'], row['period'], row['process_burden'], row['unit'])
    for row in economic['quarterly_by_medium']:
        if row['scenario_id'] != primary:
            continue
        for key in ('wastewater_ier_cost_usd', 'other_incineration_cost_usd', 'total_cost_usd'):
            add('economic.'+key, row['period'], row[key], 'USD')
        add('economic.'+row['medium']+'.total_cost_usd', row['period'], row['total_cost_usd'], 'USD')
    return coefficients, units


def statistics(values, deterministic):
    """SD describes output dispersion, not the standard error of the mean."""
    mean, sd = float(np.mean(values)), float(np.std(values, ddof=1))
    p5, p50, p95 = np.quantile(values, [.05, .50, .95])
    return {'deterministic': float(deterministic), 'mean': mean, 'sd': sd,
            'cv': sd / abs(mean) if mean else None,
            'p5': float(p5), 'p50': float(p50), 'p95': float(p95),
            'mean_change_fraction': mean / deterministic - 1 if deterministic else None}


def summarize_samples(samples, deterministic, metric, unit, periods):
    if not np.all(np.isfinite(samples)):
        raise ValueError(f'Nonfinite propagated samples: {metric}')
    quarterly = [{'metric': metric, 'unit': unit, 'period': period,
                  **statistics(samples[:, t], deterministic[t])} for t, period in enumerate(periods)]
    annual_samples = samples.reshape(samples.shape[0], 6, 4).sum(axis=2)
    annual_center = np.asarray(deterministic).reshape(6, 4).sum(axis=1)
    annual = [{'metric': metric, 'unit': unit, 'period': str(2025+y),
               **statistics(annual_samples[:, y], annual_center[y])} for y in range(6)]
    cumulative = {'metric': metric, 'unit': unit, 'period': '2025-2030',
                  **statistics(samples.sum(axis=1), sum(deterministic))}
    return quarterly, annual, cumulative


def write_uncertainty_results(*, scale, attribution, ier, incineration, toxicity, economic,
                              legacy_input, wafer_input, output_dir, scenarios=None,
                              inference_efficiency_multiplier=1.0, pag_replacement_fraction=0.0):
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    coefficients, units = downstream_coefficients(attribution, ier, incineration, toxicity, economic)
    periods = [r['Period'] for r in attribution['quarterly']]
    names = list(coefficients)
    coefficient_array = np.stack([coefficients[k] for k in names])
    # The stored coefficient matrix and per-run masses reconstruct every endpoint.
    np.savez_compressed(directory/'downstream_coefficients.npz', metrics=np.array(names),
                        units=np.array([units[k] for k in names]), periods=np.array(periods),
                        coefficients=coefficient_array)
    summary = {'schema_version': '1.0', 'settings': scale['settings'], 'seed': scale['seed'],
               'primary_treatment_scenario': ier['primary_scenario_id'],
               'sampling_rules': {
                   'distribution': 'normal, SD = mean × CV; nonpositive physical inputs raise an error',
                   'ai_inputs': 'each parameter has one independent z ~ N(0,1) per run; x_t = mean_t * (1 + CV_t * z) for all quarters including initialization',
                   'daily_tokens': 'one independent z per run, shared across quarters; normal CV=daily_token_cv',
                   'manufacturing': 'one independent normal multiplier per run, mean 1; shared by all quarters/components',
                   'reuse': 'same parameter/run/quarter sample is reused in all equations and modules',
                   'cross_parameter_correlation': 'independent assumption; shared inputs still correlate outputs',
                   'deployment': 'fixed user/model paths within each scenario; paired random numbers across scenarios',
                   'purchase': 'ceil(max(demand - surviving installed compute, 0)/new-server compute)',
                   'retirement': 'fixed legacy lifetime distribution, based on original purchase cohorts and their sampled performance',
                   'summary': 'quarterly sums within each run precede annual/cumulative mean, sample SD and quantiles',
                   'interval': 'P5–P95 simulation uncertainty interval, not a confidence interval for the mean',
                   'propagation': 'exact quarter-specific linear coefficients from freshly evaluated fixed downstream modules',
                   'scope': 'conditional on fixed composition, CF, treatment efficiency, node timing, LCI and unit costs'},
               'ai_cv_schedule': [{'decimal_quarter': float(p), 'cv': float(c)}
                                  for p,c in zip(scale['periods'], scale['ai_cv_by_quarter'])],
               'upstream_quality_checks': {name: value['quality_checks'] for name,value in
                                           [('ier',ier),('incineration',incineration),('toxicity',toxicity),('economic',economic)]},
               'scenarios': []}
    np.savez_compressed(directory/'shared_input_samples.npz',
                        **{k:v for k,v in scale['input_samples'].items() if np.asarray(v).ndim==2},
                        manufacturing_multiplier=scale['manufacturing_multipliers'], periods=scale['periods'])
    selected = list(SCENARIO_COLUMNS) if scenarios is None else list(scenarios)
    if not selected or any(name not in SCENARIO_COLUMNS for name in selected):
        raise ValueError('Unknown or empty deployment scenario selection')
    for name in selected:
        if name == scale['scenario']:
            current, current_attribution = scale, attribution
        else:
            current = run_server_projection(legacy_input, settings=scale['settings'], scenario=name, seed=scale['seed'],
                                            inference_efficiency_multiplier=inference_efficiency_multiplier)
            current_attribution = calculate_attribution(current, legacy_input, wafer_input,
                                                        pag_replacement_fraction=pag_replacement_fraction)
        central = run_server_projection(legacy_input, settings=scale['settings'], scenario=name,
                                        seed=scale['seed'], deterministic=True,
                                        inference_efficiency_multiplier=inference_efficiency_multiplier)
        calculate_attribution(central, legacy_input, wafer_input, pag_replacement_fraction=pag_replacement_fraction)
        result = {'scenario': name, 'runs': current['samples'].shape[0], 'quarterly': [], 'annual': [], 'cumulative': [],
                  'quality_checks': current['quality_checks']}
        mass, center_mass = current['pfas_samples_kg'], central['pfas_samples_kg'][0]
        arrays = {'periods': current['periods'], 'pfas_samples_kg': mass,
                  'new_servers': current['samples'], 'new_gpu_reporting': current['gpu_samples_reporting'],
                  'deterministic_pfas_kg': center_mass,
                  'ni_user': current['input_samples']['ni_user'], 'n_model': current['input_samples']['n_model']}
        for component in ('training', 'inference'):
            for key, value in current[component].items():
                arrays[component+'_'+key] = value
        np.savez_compressed(directory/f'{name}_samples.npz', **arrays)
        # Iterate without materializing all run × quarter × endpoint tensors.
        def consume(metric, samples, center, unit):
            q,a,c = summarize_samples(samples, center, metric, unit, periods)
            result['quarterly'].extend(q); result['annual'].extend(a); result['cumulative'].append(c)
        for metric in names:
            consume(metric, mass*coefficients[metric], center_mass*coefficients[metric], units[metric])
        consume('manufacturing.new_gpu_equivalents', current['gpu_samples_reporting'],
                central['gpu_samples_reporting'][0], 'GPU-equivalents')
        for component in ('training','inference'):
            mask = (current['periods'] >= 2025) & (current['periods'] < 2031)
            consume('ai.'+component+'.new_gpu_equivalents',
                    current[component]['new_servers'][:,mask]*current['gpu_per_server'],
                    central[component]['new_servers'][0,mask]*central['gpu_per_server'], 'GPU-equivalents')
        result['quality_checks']['zero_purchase_quarter_fraction_2025_2030'] = float(np.mean(current['gpu_samples_reporting']==0))
        result['quality_checks']['sample_mass_means_match_attribution'] = bool(np.allclose(
            mass.mean(axis=0), [r['PFAS input (kg)'] for r in current_attribution['quarterly']], rtol=1e-12))
        result['quality_checks']['all_checks_passed'] = bool(
            result['quality_checks']['nonnegative_purchases'] and result['quality_checks']['finite_samples']
            and result['quality_checks']['sample_mass_means_match_attribution'])
        if not result['quality_checks']['all_checks_passed']:
            raise AssertionError('Uncertainty checks failed for '+name)
        summary['scenarios'].append(result)
        headline = next(r for r in result['cumulative'] if r['metric']=='manufacturing.pfas_input_kg')
        print(f"{name}: PFAS {headline['mean']/1000:.3f} t; cumulative CV {headline['cv']:.2%}", flush=True)
    path = directory/'Uncertainty_results.json'
    path.write_text(json.dumps(summary, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
    return path
