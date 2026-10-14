"""Mine the event log and size the value of fixing what it finds.

Produces data/insights.json for the dashboard:
  kpis, process map (directly-follows graph), variants, bottlenecks,
  segment comparisons, and a value model (driver -> KPI -> $ impact).

Usage:
  python3 pipeline/mine.py
"""
import csv
import json
from collections import Counter, defaultdict
from datetime import datetime
import os
from pathlib import Path

DATA_DIR = os.environ.get("PI_DATA", "data")
from statistics import mean, median

HAPPY_PATH = ["Lead Created", "Prospecting", "Discovery Call", "Qualification", "Needs Analysis",
              "Product Demo", "Proposal/Price Quote", "Quote Sent", "Negotiation/Review",
              "Contract Signed", "Closed Won"]
# Newer Starter orgs use a shorter sales process; same analysis, different stage names.
STARTER_HAPPY_PATH = ["Qualify", "Discovery Call", "Meet & Present", "Product Demo", "Propose",
                      "Quote Sent", "Negotiate", "Contract Signed", "Closed Won"]
STAGES = {"qualify": "Qualification", "propose": "Proposal/Price Quote", "negotiate": "Negotiation/Review"}


def load(path=None):
    path = path or f"{DATA_DIR}/event_log.csv"
    cases = defaultdict(list)
    with open(path) as f:
        for row in csv.DictReader(f):
            row["start"] = datetime.fromisoformat(row["start_time"])
            row["end"] = datetime.fromisoformat(row["end_time"])
            cases[row["case_id"]].append(row)
    for events in cases.values():
        events.sort(key=lambda e: (e["start"], e["end"]))
    return cases


def days_between(a, b):
    return max((b - a).total_seconds() / 86400, 0)


def win_rate(cases):
    closed = [c for c in cases if c[0]["outcome"] in ("Won", "Lost")]
    return sum(c[0]["outcome"] == "Won" for c in closed) / len(closed) if closed else 0


def summarize(cases):
    cycle = [days_between(c[0]["start"], c[-1]["end"]) for c in cases]
    return {"cases": len(cases), "win_rate": round(win_rate(cases), 3),
            "median_cycle_days": round(median(cycle), 1) if cycle else 0}


def process_map(cases):
    nodes, edges, waits = Counter(), Counter(), defaultdict(list)
    for events in cases:
        trace = ["▶ Start"] + [e["activity"] for e in events] + ["■ End"]
        nodes.update(trace)
        for i, (a, b) in enumerate(zip(trace, trace[1:])):
            edges[(a, b)] += 1
            if 0 < i < len(events):
                waits[(a, b)].append(days_between(events[i - 1]["end"], events[i]["start"]))
    durations = defaultdict(list)
    for events in cases:
        for e in events:
            durations[e["activity"]].append(days_between(e["start"], e["end"]))
    return {
        "nodes": [{"id": n, "count": c,
                   "median_duration_days": round(median(durations[n]), 2) if durations[n] else 0}
                  for n, c in nodes.items()],
        "edges": [{"source": a, "target": b, "count": c,
                   "median_wait_days": round(median(waits[(a, b)]), 2) if waits[(a, b)] else 0}
                  for (a, b), c in edges.items()],
    }


def variants(cases, top=8):
    groups = defaultdict(list)
    for events in cases:
        groups[tuple(e["activity"] for e in events)].append(events)
    ranked = sorted(groups.items(), key=lambda kv: -len(kv[1]))
    return len(groups), [{"path": list(path), "count": len(cs), "share": round(len(cs) / len(cases), 3),
                          **summarize(cs)} for path, cs in ranked[:top]]


def has(events, activity):
    return any(e["activity"] == activity for e in events)


def approval_wait(events):
    """Request -> approval for the first approval round of a deal."""
    approved = next((e for e in events if e["activity"] == "Discount Approved"), None)
    return days_between(approved["start"], approved["end"]) if approved else None


def bottlenecks(pm, top=6):
    ranked = sorted(pm["edges"], key=lambda e: -(e["median_wait_days"] * e["count"]))
    node_dur = sorted(pm["nodes"], key=lambda n: -(n["median_duration_days"] * n["count"]))
    return {
        "handoffs": [e for e in ranked if e["median_wait_days"] > 0][:top],
        "activities": [n for n in node_dur if n["median_duration_days"] > 0][:top],
    }


def main():
    global HAPPY_PATH
    cases_by_id = load()
    cases = list(cases_by_id.values())
    if any(e["activity"] == "Negotiate" for c in cases for e in c):
        HAPPY_PATH = STARTER_HAPPY_PATH
        STAGES.update(qualify="Qualify", propose="Propose", negotiate="Negotiate")
    all_events = [e for c in cases for e in c]
    first, last = min(e["start"] for e in all_events), max(e["end"] for e in all_events)
    months = max(days_between(first, last) / 30.4, 1)
    annualize = 12 / months

    won = [c for c in cases if c[0]["outcome"] == "Won"]
    avg_deal = mean(float(c[0]["amount"]) for c in won) if won else 0

    # Compare like with like: only deals that reached Negotiation can be re-quoted.
    negotiated = [c for c in cases if has(c, STAGES["negotiate"])]
    rework = [c for c in negotiated if sum(e["activity"] == STAGES["propose"] for e in c) > 1]
    clean = [c for c in negotiated if c not in rework]
    skipped = [c for c in cases if not has(c, STAGES["qualify"])]
    qualified = [c for c in cases if has(c, STAGES["qualify"])]
    approvals = defaultdict(list)
    for c in cases:
        wait = approval_wait(c)
        if wait is not None:
            approvals[c[0]["region"]].append((wait, c))

    pm = process_map(cases)
    n_variants, top_variants = variants(cases)
    conforming = sum(1 for c in cases if [e["activity"] for e in c] == HAPPY_PATH)

    # ---------- value model: driver -> KPI baseline -> target -> $ ----------
    region_wait = {r: round(median(w for w, _ in v), 1) for r, v in approvals.items()}
    best_region_wait = min(region_wait.values()) if region_wait else 0
    emea_cases = approvals.get("EMEA", [])
    emea_slow = [c for w, c in emea_cases if w > 7]
    emea_fast = [c for w, c in emea_cases if w <= 7]
    rate_slow, rate_fast = win_rate(emea_slow), win_rate(emea_fast)
    approval_wins = len(emea_slow) * max(rate_fast - rate_slow, 0) * annualize
    days_saved = (region_wait.get("EMEA", 0) - best_region_wait)

    rework_lift = max(win_rate(clean) - win_rate(rework), 0)
    qual_lift = max(win_rate(qualified) - win_rate(skipped), 0)
    pipeline_annual = sum(float(c[0]["amount"]) for c in cases) * annualize

    drivers = [
        {
            "id": "approval",
            "theme": "Accelerate bookings",
            "driver": "Remove EMEA discount-approval bottleneck",
            "evidence": f"EMEA approvals take {region_wait.get('EMEA', 0)} days (median) vs "
                        f"{best_region_wait} days in the best region; deals waiting >7 days win "
                        f"{rate_slow:.0%} vs {rate_fast:.0%}.",
            "kpi": "Median discount-approval cycle (days)",
            "baseline": region_wait.get("EMEA", 0), "target": best_region_wait,
            "affected_deals_per_year": round(len(emea_cases) * annualize),
            "win_rate_lift": round(rate_fast - rate_slow, 3),
            "deals_at_risk_per_year": round(len(emea_slow) * annualize),
            "default_capture": 0.5,
            "annual_value": round(approval_wins * avg_deal),
            "days_saved_per_deal": round(days_saved, 1),
            "cohorts": {"problem": len(emea_slow), "comparison": len(emea_fast),
                        "problem_label": "EMEA deals waiting >7 days", "comparison_label": "EMEA deals waiting ≤7 days"},
        },
        {
            "id": "rework",
            "theme": "Improve win rate",
            "driver": "Cut re-quote rework in Negotiation",
            "evidence": f"{len(rework) / len(negotiated):.0%} of negotiated deals loop back to Proposal; they win "
                        f"{win_rate(rework):.0%} vs {win_rate(clean):.0%} for clean deals.",
            "kpi": "Re-quote rate (% of deals)",
            "baseline": round(len(rework) / len(negotiated), 3),
            "target": round(len(rework) / len(negotiated) / 2, 3),
            "affected_deals_per_year": round(len(rework) * annualize),
            "win_rate_lift": round(rework_lift, 3),
            "default_capture": 0.5,
            "annual_value": round(len(rework) * annualize * rework_lift * avg_deal),
            "negotiated_per_year": round(len(negotiated) * annualize, 1),
            "cohorts": {"problem": len(rework), "comparison": len(clean),
                        "problem_label": "re-quoted", "comparison_label": "clean path"},
        },
        {
            "id": "qualification",
            "theme": "Increase sales productivity",
            "driver": "Enforce the Qualification gate",
            "evidence": f"{len(skipped) / len(cases):.0%} of deals skip Qualification and win "
                        f"{win_rate(skipped):.0%} vs {win_rate(qualified):.0%}.",
            "kpi": "Deals skipping Qualification (%)",
            "baseline": round(len(skipped) / len(cases), 3), "target": 0.02,
            "affected_deals_per_year": round(len(skipped) * annualize),
            "win_rate_lift": round(qual_lift, 3),
            "default_capture": 0.5,
            "annual_value": round(len(skipped) * annualize * qual_lift * avg_deal),
            "cohorts": {"problem": len(skipped), "comparison": len(qualified),
                        "problem_label": "skipped Qualification", "comparison_label": "qualified"},
        },
    ]

    ids = lambda cs: {c[0]["case_id"] for c in cs}
    overlaps = {
        "rework_and_qualification": len(ids(rework) & ids(skipped)),
        "rework_and_approval": len(ids(rework) & ids(emea_slow)),
        "qualification_and_approval": len(ids(skipped) & ids(emea_slow)),
    }

    by_region = {}
    for region in sorted({c[0]["region"] for c in cases}):
        rc = [c for c in cases if c[0]["region"] == region]
        by_region[region] = {**summarize(rc), "approval_wait_days": region_wait.get(region)}

    monthly = defaultdict(lambda: {"opened": 0, "won": 0})
    for c in cases:
        monthly[c[0]["start"].strftime("%Y-%m")]["opened"] += 1
        if c[0]["outcome"] == "Won":
            monthly[c[-1]["end"].strftime("%Y-%m")]["won"] += 1

    insights = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source": "synthetic" if Path(f"{DATA_DIR}/raw.json").exists() and
                  json.loads(Path(f"{DATA_DIR}/raw.json").read_text())["source"] == "synthetic" else "salesforce",
        "period": {"from": first.date().isoformat(), "to": last.date().isoformat(),
                   "months": round(months, 1), "span_days": round(days_between(first, last), 2)},
        "kpis": {
            **summarize(cases),
            "events": len(all_events),
            "activities": len({e["activity"] for e in all_events}),
            "variants": n_variants,
            "happy_path_conformance": round(conforming / len(cases), 3),
            "winning_trace_share_of_won": round(conforming / len(won), 3) if won else 0,
            "rework_rate": round(len(rework) / len(negotiated), 3),
            "avg_won_deal": round(avg_deal),
            "annual_pipeline": round(pipeline_annual),
            "events_by_object": dict(Counter(e["source_object"] for e in all_events)),
        },
        "process_map": pm,
        "happy_path": HAPPY_PATH,
        "variants": top_variants,
        "bottlenecks": bottlenecks(pm),
        "by_region": by_region,
        "monthly": [{"month": m, **v} for m, v in sorted(monthly.items())],
        "value_drivers": drivers,
        "total_annual_value": sum(d["annual_value"] for d in drivers),
        "overlaps": overlaps,
    }
    Path(f"{DATA_DIR}/insights.json").write_text(json.dumps(insights, indent=1))
    k = insights["kpis"]
    print(f"  {k['cases']} cases, {k['variants']} variants, win {k['win_rate']:.0%}, "
          f"median cycle {k['median_cycle_days']}d, conformance {k['happy_path_conformance']:.0%}")
    for d in drivers:
        print(f"  [{d['theme']}] {d['driver']}: ${d['annual_value']:,.0f}/yr — {d['evidence']}")
    print(f"Wrote {DATA_DIR}/insights.json")


if __name__ == "__main__":
    main()
