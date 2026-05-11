from app.config import settings
from app.schemas import FeeInput


def calculate_fee(fee_input: FeeInput) -> dict:
    labor_total = 0.0
    phase_totals: dict[str, float] = {}

    for role in fee_input.roles:
        for phase, hours in role.hours_by_phase.items():
            role_cost = hours * role.rate
            phase_totals[phase] = phase_totals.get(phase, 0.0) + role_cost
            labor_total += role_cost

    overhead_pct = (
        fee_input.overhead_pct
        if fee_input.overhead_pct is not None
        else settings.default_overhead_pct
    )
    overhead = labor_total * overhead_pct

    included_subs = [s for s in fee_input.subconsultants if s.included_in_lump_sum]
    excluded_subs = [s for s in fee_input.subconsultants if not s.included_in_lump_sum]
    subconsultants_total = sum(s.fee for s in included_subs)

    travel_unit_cost = (
        fee_input.travel.unit_cost
        if fee_input.travel.unit_cost is not None
        else settings.travel_unit_cost_usd
    )
    travel_total = fee_input.travel.trips * fee_input.travel.people_per_trip * travel_unit_cost
    reimbursables_total = fee_input.misc_reimbursables
    if fee_input.travel.include_in_lump_sum:
        reimbursables_total += travel_total

    lump_sum_total = labor_total + overhead + subconsultants_total + reimbursables_total

    return {
        "currency": settings.default_currency,
        "phase_totals": phase_totals,
        "labor_total": round(labor_total, 2),
        "overhead_pct": overhead_pct,
        "overhead_total": round(overhead, 2),
        "subconsultants_included_total": round(subconsultants_total, 2),
        "subconsultants_excluded": [
            {"name": s.name, "fee": s.fee} for s in excluded_subs
        ],
        "travel": {
            "trips": fee_input.travel.trips,
            "people_per_trip": fee_input.travel.people_per_trip,
            "unit_cost": travel_unit_cost,
            "total": round(travel_total, 2),
            "included_in_lump_sum": fee_input.travel.include_in_lump_sum,
        },
        "misc_reimbursables": round(fee_input.misc_reimbursables, 2),
        "reimbursables_total": round(reimbursables_total, 2),
        "lump_sum_total": round(lump_sum_total, 2),
    }
