"""Compact evidence-to-parameter agent with explicit scientific service interfaces.

The controller exposes assessment, targeted follow-up, tiered selection,
endpoint readiness and calculation handoff. Provider integrations are supplied
through Services; scientific judgments and their reasons remain in the audit.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
import math
from typing import Any, Mapping, Protocol, Sequence

from parameter_contract import ENDPOINT_PARAMETER_KEYS, PARAMETER_SPECS

ROUTES = ('snapshot', 'ctx', 'literature', 'opera', 'read_across')
TIER = {'direct': 1, 'derived': 2, 'predicted': 3, 'read_across': 4, 'assumption': 4}


@dataclass(frozen=True)
class Identity:
    name: str
    casrn: str
    modeled_form: str
    terminal_class: str  # PFCA or PFSA


@dataclass(frozen=True)
class Evidence:
    id: str
    casrn: str
    parameter: str
    value: float
    unit: str
    kind: str
    source: str
    locator: str
    measurement_id: str  # Stable source-location identity across extractions.
    modeled_form: str
    applicable: bool = True
    rationale: str = ''


@dataclass(frozen=True)
class Assessment:
    accepted_ids: tuple[str, ...]
    conflicting_parameters: tuple[str, ...]
    reasons: Mapping[str, str]


@dataclass(frozen=True)
class Action:
    route: str
    parameters: tuple[str, ...]
    rationale: str


@dataclass
class State:
    identity: Identity
    endpoints: tuple[str, ...]
    evidence: list[Evidence] = field(default_factory=list)
    selected: dict[str, dict[str, Any]] = field(default_factory=dict)
    conflicts: tuple[str, ...] = ()
    attempts: dict[str, dict[str, int]] = field(default_factory=dict)
    audit: list[dict[str, Any]] = field(default_factory=list)

    @property
    def required(self) -> set[str]:
        return {p for e in self.endpoints for p in ENDPOINT_PARAMETER_KEYS[e]}

    @property
    def gaps(self) -> tuple[str, ...]:
        return tuple(sorted(self.required - self.selected.keys()))


class Services(Protocol):
    """Scientific adapters used by the controller.

    resolve: registry/structure identity and parent-acid normalization.
    retrieve: snapshot, CTX, parameter-specific literature/SI, OPERA or read-across.
    assess: LLM evidence interpretation plus scientific applicability/conflict review.
    decide: structured LLM action using gaps, conflicts, history and remaining routes.
    derive: fixed relationships from currently selected parent parameters.
    calculate: deterministic PFAS-adapted USEtox for one ready endpoint.
    """

    def resolve(self, query: str) -> Identity: ...
    def retrieve(self, identity: Identity, action: Action) -> Sequence[Evidence]: ...
    def assess(self, identity: Identity, evidence: Sequence[Evidence]) -> Assessment: ...
    def decide(self, state: State, available: Mapping[str, tuple[str, ...]]) -> Action | None: ...
    def derive(self, identity: Identity, selected: Mapping[str, dict]) -> Sequence[Evidence]: ...
    def calculate(self, identity: Identity, endpoint: str, parameters: Mapping[str, float]) -> Any: ...


def validate_record(record: Evidence, identity: Identity) -> str | None:
    """Deterministic checks applied after evidence extraction and unit conversion."""
    spec = PARAMETER_SPECS.get(record.parameter)
    if spec is None:
        return 'Unknown parameter'
    if record.casrn != identity.casrn or record.modeled_form != identity.modeled_form:
        return 'Chemical identity or modeled form does not match'
    if record.unit != spec.unit:
        return 'Canonical unit conversion required'
    if record.kind not in TIER or not record.applicable:
        return 'Evidence route or applicability check failed'
    if not all((record.id, record.source, record.locator, record.measurement_id)):
        return 'Source and measurement location required'
    if isinstance(record.value, bool) or not isinstance(record.value, (int, float)) or not math.isfinite(record.value):
        return 'Finite numerical value required'
    if spec.minimum is not None:
        if record.value < spec.minimum or (spec.exclusive_minimum and record.value == spec.minimum):
            return 'Value outside the parameter domain'
    return None


def select_parameters(state: State, assessment: Assessment) -> dict[str, dict]:
    """Select the highest-ranked usable route, deduplicating measurement groups.

    Arithmetic means implement the study's availability-first screening policy.
    The assessment adapter resolves conflicting interpretations before aggregation.
    """
    groups: dict[str, list[Evidence]] = {}
    accepted = set(assessment.accepted_ids)
    for record in state.evidence:
        if record.id in accepted and record.parameter not in assessment.conflicting_parameters:
            groups.setdefault(record.parameter, []).append(record)
    result = {}
    route_order = {'direct': 0, 'derived': 1, 'predicted': 2, 'read_across': 3, 'assumption': 4}
    for parameter, records in groups.items():
        priority = min(route_order[r.kind] for r in records)
        chosen = [r for r in records if route_order[r.kind] == priority]
        unique: dict[str, Evidence] = {}
        for record in chosen:
            previous = unique.get(record.measurement_id)
            if previous is not None and previous.value != record.value:
                raise ValueError('Conflicting extraction of one measurement must be resolved by assessment')
            unique[record.measurement_id] = record
        selected = list(unique.values())
        result[parameter] = {
            'value': sum(r.value for r in selected) / len(selected),
            'unit': PARAMETER_SPECS[parameter].unit,
            'tier': f'T{TIER[selected[0].kind]}',
            'evidence_ids': [r.id for r in selected],
            'rationale': assessment.reasons.get(parameter, 'Highest-priority usable evidence'),
        }
    return result


def run(query: str, services: Services, *,
        endpoints: tuple[str, ...] = ('ctuh_noncancer', 'ctue_freshwater'),
        budgets: Mapping[str, int] | None = None) -> dict[str, Any]:
    """Evaluate evidence, choose a targeted action, reassess and hand off ready inputs."""
    if not endpoints or any(e not in ENDPOINT_PARAMETER_KEYS for e in endpoints):
        raise ValueError('Select supported toxicity endpoints')
    # Controller-call quotas; provider-specific document/search budgets belong to adapters.
    limits = dict(budgets if budgets is not None else {
        'snapshot': 1, 'ctx': 1, 'literature': 2, 'opera': 1, 'read_across': 1,
    })
    if set(limits) - set(ROUTES) or any(type(n) is not int or n < 0 for n in limits.values()):
        raise ValueError('Route budgets must be nonnegative integers')
    identity = services.resolve(query)
    if not identity.casrn or not identity.modeled_form or identity.terminal_class not in ('PFCA', 'PFSA'):
        return {'identity': asdict(identity), 'status': 'NA', 'reason': 'Terminal-PFAA identity unresolved'}
    state = State(identity, endpoints)
    state.audit.append({'stage': 'identity', 'identity': asdict(identity)})

    def ingest(records: Sequence[Evidence]) -> None:
        for record in records:
            issue = validate_record(record, identity)
            if issue:
                state.audit.append({'stage': 'rejected', 'evidence_id': record.id, 'reason': issue})
            else:
                previous = next((r for r in state.evidence if r.id == record.id), None)
                if previous is None:
                    state.evidence.append(record)
                elif previous != record:
                    raise ValueError('An evidence ID must identify one immutable record')

    def assess() -> None:
        assessment = services.assess(identity, tuple(state.evidence))
        if set(assessment.accepted_ids) - {r.id for r in state.evidence}:
            raise ValueError('Assessment selected unknown evidence IDs')
        state.conflicts = assessment.conflicting_parameters
        state.selected = select_parameters(state, assessment)
        state.audit.append({'stage': 'assessment', 'assessment': asdict(assessment),
                            'selected': dict(state.selected), 'gaps': state.gaps})

    stop_reason = 'Parameters ready'
    derived_ids: set[str] = set()
    while True:
        # Invalidate computed candidates when their selected parent evidence changes.
        state.evidence = [r for r in state.evidence if r.id not in derived_ids]
        assess()
        # Recompute dependent candidates from the latest selected parent parameters.
        derived = tuple(services.derive(identity, state.selected))
        if derived:
            ingest(derived)
            derived_ids = {r.id for r in derived}
            assess()
        if not state.gaps:
            break
        available = {
            p: tuple(r for r in ROUTES if state.attempts.get(p, {}).get(r, 0) < limits.get(r, 0))
            for p in state.gaps
        }
        available = {p: routes for p, routes in available.items() if routes}
        if not available:
            stop_reason = 'Supported retrieval budgets exhausted'
            break
        action = services.decide(state, available)
        if action is None:
            stop_reason = 'Agent assessment found no further supported retrieval route'
            break
        if not action.parameters or not action.rationale.strip():
            raise ValueError('Action must identify target parameters and a decision rationale')
        if len(set(action.parameters)) != len(action.parameters):
            raise ValueError('Target parameters must be unique')
        if any(action.route not in available.get(p, ()) for p in action.parameters):
            raise ValueError('Action targets a resolved parameter or an exhausted route')
        for p in action.parameters:
            counts = state.attempts.setdefault(p, {})
            counts[action.route] = counts.get(action.route, 0) + 1
        state.audit.append({'stage': 'action', **asdict(action)})
        ingest(services.retrieve(identity, action))

    results = {}
    for endpoint in endpoints:
        missing = sorted(set(ENDPOINT_PARAMETER_KEYS[endpoint]) - state.selected.keys())
        if missing:
            results[endpoint] = {'status': 'NA', 'missing_parameters': missing}
        else:
            values = {p: state.selected[p]['value'] for p in ENDPOINT_PARAMETER_KEYS[endpoint]}
            results[endpoint] = {'status': 'calculated', 'result': services.calculate(identity, endpoint, values)}
    return {'identity': asdict(identity), 'selected_parameters': state.selected,
            'endpoint_results': results, 'stop_reason': stop_reason,
            'evidence': [asdict(r) for r in state.evidence], 'audit': state.audit}
