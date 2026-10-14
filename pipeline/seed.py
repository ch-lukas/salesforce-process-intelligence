"""Seed a Salesforce trial org with demo deals that walk through real stage changes.

Replays the simulator's lifecycles against the org: creates Accounts and
Opportunities, then advances stages and logs completed Tasks round by round, so
OpportunityHistory and Task hold genuine, system-stamped events. Timestamps are
"now" (Salesforce can't backdate them), so the live log proves the plumbing while
the simulated dataset carries the months-long behaviour.

All seeded records are prefixed "[Demo]" and can be removed with --cleanup.

Usage:
  python3 pipeline/seed.py --target-org trial [--deals 40]
  python3 pipeline/seed.py --target-org trial --cleanup
"""
import argparse
import random
import time
from collections import defaultdict

from sf_client import API_VERSION, SalesforceClient
from synthesize import simulate
from datetime import datetime, timezone

PREFIX = "[Demo]"
LEAD_SOURCES = {"Trade Show": "Other", "Outbound": "Phone Inquiry", "Webinar": "Web"}
# Simulator uses the classic sales process; newer Starter orgs use a shorter one.
STARTER_STAGES = {"Prospecting": None, "Qualification": "Qualify", "Needs Analysis": "Meet & Present",
                  "Proposal/Price Quote": "Propose", "Negotiation/Review": "Negotiate",
                  "Closed Won": "Closed Won", "Closed Lost": "Closed Lost"}


def collection(sf, method, records):
    """sObject Collections: up to 200 records per call, all-or-none off so one bad row is visible."""
    results = []
    for i in range(0, len(records), 200):
        body = {"allOrNone": False, "records": records[i:i + 200]}
        results += sf._request(method, f"/services/data/{API_VERSION}/composite/sobjects", body)
    errors = [r["errors"] for r in results if not r.get("success")]
    if errors:
        raise RuntimeError(f"{len(errors)} records failed, first: {errors[0]}")
    return [r["id"] for r in results]


def stage_map(sf, needed):
    """Return simulator stage -> org stage (None = drop), or stop if the org fits neither process."""
    active = {r["ApiName"] for r in sf.query_all("SELECT ApiName FROM OpportunityStage WHERE IsActive = true")}
    if needed <= active:
        return {s: s for s in needed}
    if {v for v in STARTER_STAGES.values() if v} <= active:
        return STARTER_STAGES
    raise SystemExit(f"The org's sales process doesn't match. Active stages: {sorted(active)}")


def seed(sf, n_deals):
    random.seed(11)
    plan = simulate(n_deals, datetime(2026, 1, 5, tzinfo=timezone.utc))
    stages = stage_map(sf, {h["StageName"] for h in plan["stage_history"]})

    # Accounts: one per deal keeps industry/country attributes per case.
    accounts = [{"attributes": {"type": "Account"}, "Name": f"{PREFIX} {o['Name'].split(' - ')[0]}",
                 "Industry": o["Account.Industry"], "BillingCountry": o["Account.BillingCountry"]}
                for o in plan["opportunities"]]
    account_ids = collection(sf, "POST", accounts)
    print(f"  created {len(account_ids)} accounts")

    # Each deal's lifecycle as an ordered list of steps: ("stage", name) or ("task", subject).
    steps = defaultdict(list)
    for h in plan["stage_history"]:
        if stages[h["StageName"]]:
            steps[h["OpportunityId"]].append((h["CreatedDate"], "stage", stages[h["StageName"]]))
    for t in plan["tasks"]:
        steps[t["WhatId"]].append((t["CreatedDate"], "task", t["Subject"]))
    for v in steps.values():
        v.sort()
        # The opportunity is created at its first stage; anything logged before that follows it.
        first = next(i for i, step in enumerate(v) if step[1] == "stage")
        v.insert(0, v.pop(first))

    opps = []
    for o, account_id in zip(plan["opportunities"], account_ids):
        first_stage = steps[o["Id"]][0][2]
        opps.append({"attributes": {"type": "Opportunity"}, "Name": f"{PREFIX} {o['Name']}",
                     "AccountId": account_id, "Amount": o["Amount"], "StageName": first_stage,
                     "CloseDate": o["CloseDate"], "LeadSource": LEAD_SOURCES.get(o["LeadSource"], o["LeadSource"])})
    opp_ids = collection(sf, "POST", opps)
    real_id = dict(zip([o["Id"] for o in plan["opportunities"]], opp_ids))
    print(f"  created {len(opp_ids)} opportunities")

    # Replay round by round so history rows get distinct, correctly ordered timestamps.
    cursors = {sim_id: 1 for sim_id in real_id}  # step 0 was the create
    round_no = 0
    while any(cursors[s] < len(steps[s]) for s in cursors):
        round_no += 1
        updates, tasks = [], []
        for sim_id, i in cursors.items():
            if i >= len(steps[sim_id]):
                continue
            _, kind, value = steps[sim_id][i]
            if kind == "stage":
                updates.append({"attributes": {"type": "Opportunity"}, "Id": real_id[sim_id], "StageName": value})
            else:
                tasks.append({"attributes": {"type": "Task"}, "WhatId": real_id[sim_id], "Subject": value,
                              "Status": "Completed", "ActivityDate": time.strftime("%Y-%m-%d")})
            cursors[sim_id] = i + 1
        if updates:
            collection(sf, "PATCH", updates)
        if tasks:
            collection(sf, "POST", tasks)
        print(f"  round {round_no:>2}: {len(updates)} stage changes, {len(tasks)} tasks")
        time.sleep(2)


def cleanup(sf):
    for sobject, field in (("Opportunity", "Name"), ("Account", "Name")):
        ids = [r["Id"] for r in sf.query_all(f"SELECT Id FROM {sobject} WHERE {field} LIKE '{PREFIX}%'")]
        sf.delete_ids(sobject, ids)
        print(f"  deleted {len(ids)} {sobject} records")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--target-org")
    parser.add_argument("--deals", type=int, default=40)
    parser.add_argument("--cleanup", action="store_true")
    args = parser.parse_args()
    sf = SalesforceClient(args.target_org)
    cleanup(sf) if args.cleanup else seed(sf, args.deals)


if __name__ == "__main__":
    main()
