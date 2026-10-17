# Structured evidence decisions

The assessment and action interfaces use structured language-model responses. The following contracts summarize the scientific questions supplied to the model.

## Evidence assessment

Given the resolved target identity, modeled chemical form and candidate evidence:

1. Determine whether each observation represents the requested parameter for the target substance and modeled form.
2. Read the source context, experimental conditions, unit conversion and applicability information together.
3. Distinguish independent measurements from repeated reporting of the same measurement.
4. Resolve differences caused by chemical form, endpoint definition or conversion. Identify conflicts that require additional evidence.
5. Return the accepted evidence IDs, unresolved parameter conflicts and parameter-specific reasons.

The controller subsequently applies the explicit route priorities and canonical numerical checks in `agent_core.py`.

Assessment response fields:

```text
accepted_ids: list of IDs from the supplied evidence
conflicting_parameters: list of canonical parameter keys
reasons: parameter key -> scientific selection or conflict rationale
```

## Next-action decision

Given the current selected parameters, endpoint dependencies, unresolved gaps, conflicting evidence, previous actions and available route budgets:

1. Identify the evidence gap that prevents an endpoint from being calculated.
2. Select the available acquisition route most likely to resolve that gap. Consider existing evidence and CTX, targeted literature and supplementary information, applicable OPERA predictions, and supported read-across.
3. For a conflict, direct follow-up toward the specific ambiguity: substance identity, chemical form, endpoint, unit or measurement context.
4. Select only unresolved target parameters with an available route budget.
5. Explain how the expected evidence would change parameter selection or calculation readiness.
6. Return no action when the remaining routes cannot supply a supported parameter.

Action response fields:

```text
route: snapshot | ctx | literature | opera | read_across
parameters: list of canonical target keys
rationale: reason for this action based on the current evidence state
```

Provider adapters use the selected parameter targets and rationale to construct searches, choose prediction endpoints or assess donor candidates. Their returned evidence is assessed again before calculation.
