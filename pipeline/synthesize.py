"""Generate a realistic, Salesforce-shaped raw.json (same schema as extract.py).

Trial / Trailhead orgs can't backdate OpportunityHistory.CreatedDate, so a live
extract only ever shows minutes of history. This simulator produces nine months
of lead-to-close behaviour with known, planted process problems so the mining
step has something real to find:

  * Discount approvals stall in EMEA (single approver) and need Finance sign-off above 25%
  * ~20% of deals bounce Negotiation -> Proposal for a re-quote (rework loop)
  * ~12% of deals skip Qualification, and win far less often
  * Enterprise deals (>$100k) wait on Legal Review

Usage:
  python3 pipeline/synthesize.py [--cases 800] [--seed 7]
"""
import argparse
import json
import random
import string
from datetime import datetime, timedelta, timezone
from pathlib import Path

REPS = [
    ("Avery Chen", "AMER"), ("Jordan Blake", "AMER"), ("Sam Ortiz", "AMER"), ("Riley Park", "AMER"),
    ("Morgan Hale", "EMEA"), ("Elena Vogt", "EMEA"), ("Tomás Rey", "EMEA"), ("Priya Shah", "EMEA"),
    ("Kenji Mori", "APAC"), ("Mia Wong", "APAC"), ("Liam Kelly", "APAC"), ("Ana Lima", "APAC"),
]
COUNTRIES = {"AMER": ["United States", "Canada"], "EMEA": ["Germany", "United Kingdom", "France"],
             "APAC": ["Australia", "Singapore", "Japan"]}
INDUSTRIES = ["Banking", "Insurance", "Healthcare", "Manufacturing", "Telecommunications", "Retail"]
LEAD_SOURCES = ["Web", "Partner Referral", "Trade Show", "Outbound", "Webinar"]
APPROVERS = {"AMER": "Dana Whitfield", "EMEA": "Klaus Berger", "APAC": "Grace Tan"}
FINANCE = "Finance Desk"
LEGAL = "Legal Ops"


def sf_id(prefix):
    return prefix + "".join(random.choices(string.ascii_letters + string.digits, k=12))


def iso(ts):
    return ts.strftime("%Y-%m-%dT%H:%M:%S.000+0000")


def work_hours(ts):
    """Nudge a timestamp into a weekday working window so the log looks human."""
    while ts.weekday() >= 5:
        ts += timedelta(days=1)
    if ts.hour < 8 or ts.hour >= 18:
        ts = ts.replace(hour=random.randint(8, 17), minute=random.randint(0, 59))
    return ts


def days(median, spread=0.6):
    return timedelta(days=random.lognormvariate(0, spread) * median)


def simulate(n_cases, start):
    users = {sf_id("005"): {"Name": n, "Region": r} for n, r in REPS}
    rep_ids = list(users)
    raw = {"source": "synthetic", "opportunities": [], "stage_history": [], "tasks": [], "leads": []}

    for i in range(n_cases):
        owner = random.choice(rep_ids)
        rep, region = users[owner]["Name"], users[owner]["Region"]
        opp_id = sf_id("006")
        amount = round(random.lognormvariate(10.6, 0.8), -2)
        discount = random.choice([0, 0, 5, 10, 10, 15, 20, 25, 30, 35])
        industry = random.choice(INDUSTRIES)
        t = work_hours(start + timedelta(days=random.uniform(0, 240), hours=random.uniform(0, 9)))
        win_odds = 0.55
        stages, tasks = [], []

        def stage(name, by=rep):
            stages.append((name, t, by))

        def task(subject, duration, by=rep):
            nonlocal t
            began = t
            t = work_hours(t + duration)
            tasks.append((subject, began, t, by))

        # Lead -> conversion
        lead_created = t - days(4)
        raw["leads"].append({"Id": sf_id("00Q"), "CreatedDate": iso(lead_created),
                             "ConvertedDate": t.strftime("%Y-%m-%d"),
                             "ConvertedOpportunityId": opp_id, "Owner.Name": rep})

        stage("Prospecting")
        task("Discovery Call", days(2))
        lost = False

        if random.random() < 0.12:
            win_odds -= 0.30  # skipped qualification
        else:
            t = work_hours(t + days(3))
            stage("Qualification")
            if random.random() < 0.10:
                lost = True

        if not lost:
            t = work_hours(t + days(4))
            stage("Needs Analysis")
            task("Product Demo", days(1.5))
            if random.random() < 0.12:
                lost = True

        rounds = 0
        while not lost:
            rounds += 1
            t = work_hours(t + days(2))
            stage("Proposal/Price Quote")
            task("Quote Sent", days(1, 0.3))
            if discount >= 20:
                task("Discount Approval Requested", timedelta(hours=random.uniform(0.2, 2)))
                wait = days(6.5 if region == "EMEA" else 1.2, 0.5)
                task("Discount Approved", wait, by=APPROVERS[region])
                if wait.days > 7:
                    win_odds -= 0.18
                if discount >= 30:
                    task("Finance Approval", days(3), by=FINANCE)
            t = work_hours(t + days(3))
            stage("Negotiation/Review")
            if rounds == 1 and random.random() < 0.20:
                task("Customer Requested Re-quote", days(2))
                win_odds -= 0.20
                continue
            break

        if not lost and amount > 100_000:
            task("Legal Review", days(5, 0.5), by=LEGAL)

        if lost or random.random() > win_odds:
            t = work_hours(t + days(5))
            stage("Closed Lost")
            won = False
        else:
            task("Contract Signed", days(3))
            stage("Closed Won")
            won = True

        country = random.choice(COUNTRIES[region])
        raw["opportunities"].append({
            "Id": opp_id, "Name": f"{industry} Co {i:04d} - Platform", "Amount": amount,
            "StageName": stages[-1][0], "IsWon": won, "IsClosed": True,
            "LeadSource": random.choice(LEAD_SOURCES), "CreatedDate": iso(stages[0][1]),
            "CloseDate": stages[-1][1].strftime("%Y-%m-%d"), "OwnerId": owner, "Owner.Name": rep,
            "Account.Industry": industry, "Account.BillingCountry": country,
            "Discount_Percent": discount,
        })
        for name, ts, by in stages:
            raw["stage_history"].append({"OpportunityId": opp_id, "StageName": name,
                                         "CreatedDate": iso(ts), "CreatedBy.Name": by})
        for subject, began, ended, by in tasks:
            raw["tasks"].append({"WhatId": opp_id, "Subject": subject, "CreatedDate": iso(began),
                                 "CompletedDateTime": iso(ended), "Owner.Name": by, "Status": "Completed"})
    return raw


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cases", type=int, default=800)
    parser.add_argument("--seed", type=int, default=7)
    parser.add_argument("--out", default="data/raw.json")
    args = parser.parse_args()

    random.seed(args.seed)
    raw = simulate(args.cases, datetime(2025, 11, 3, 9, tzinfo=timezone.utc))
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out).write_text(json.dumps(raw, indent=1))
    for key in ("opportunities", "stage_history", "tasks", "leads"):
        print(f"  {key:15} {len(raw[key]):>6} records")
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
