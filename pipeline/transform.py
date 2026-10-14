"""Turn raw Salesforce records into a process-mining event log.

Case ID   = Opportunity Id (Lead events joined in via ConvertedOpportunityId)
Activity  = stage name (OpportunityHistory) or task subject (Task)
Timestamps= start/end; stage changes are instantaneous, tasks carry duration
Resource  = who did it; case attributes = region, industry, amount, outcome...

Outputs data/event_log.csv (Apromore CSV import) and data/event_log.xes (IEEE XES).

Usage:
  python3 pipeline/transform.py
"""
import csv
import json
from datetime import datetime
import os
from pathlib import Path

DATA_DIR = os.environ.get("PI_DATA", "data")
from xml.sax.saxutils import quoteattr

REGION_BY_COUNTRY = {
    "United States": "AMER", "USA": "AMER", "US": "AMER", "Canada": "AMER",
    "Germany": "EMEA", "United Kingdom": "EMEA", "UK": "EMEA", "France": "EMEA",
    "Australia": "APAC", "Singapore": "APAC", "Japan": "APAC",
}
COLUMNS = ["case_id", "activity", "start_time", "end_time", "resource", "source_object",
           "region", "industry", "lead_source", "amount", "discount_pct", "outcome", "owner"]


def parse_ts(value):
    if not value:
        return None
    if len(value) == 10:  # Date-only fields such as ConvertedDate
        return datetime.fromisoformat(value + "T09:00:00+00:00")
    return datetime.fromisoformat(value.replace("+0000", "+00:00").replace("Z", "+00:00"))


def build_events(raw):
    opps = {o["Id"]: o for o in raw["opportunities"]}
    events = []

    def add(case_id, activity, start, end, resource, source):
        if case_id in opps and start:
            events.append({"case_id": case_id, "activity": activity, "start": start,
                           "end": end or start, "resource": resource or "Unknown",
                           "source_object": source})

    for lead in raw.get("leads", []):
        add(lead["ConvertedOpportunityId"], "Lead Created", parse_ts(lead["CreatedDate"]), None,
            lead.get("Owner.Name"), "Lead")
    for row in raw["stage_history"]:
        add(row["OpportunityId"], row["StageName"], parse_ts(row["CreatedDate"]), None,
            row.get("CreatedBy.Name"), "OpportunityHistory")
    for task in raw["tasks"]:
        add(task["WhatId"], task["Subject"], parse_ts(task["CreatedDate"]),
            parse_ts(task.get("CompletedDateTime")), task.get("Owner.Name"), "Task")

    # OpportunityHistory writes a row on any field change; keep only real stage transitions.
    events.sort(key=lambda e: (e["case_id"], e["start"], e["end"]))
    deduped, last_stage = [], {}
    for e in events:
        if e["source_object"] == "OpportunityHistory":
            if last_stage.get(e["case_id"]) == e["activity"]:
                continue
            last_stage[e["case_id"]] = e["activity"]
        deduped.append(e)

    for e in deduped:
        o = opps[e["case_id"]]
        e.update({
            "region": REGION_BY_COUNTRY.get(o.get("Account.BillingCountry") or "", "Unknown"),
            "industry": o.get("Account.Industry") or "Unknown",
            "lead_source": o.get("LeadSource") or "Unknown",
            "amount": o.get("Amount") or 0,
            "discount_pct": o.get("Discount_Percent", ""),
            "outcome": "Won" if o.get("IsWon") else ("Lost" if o.get("IsClosed") else "Open"),
            "owner": o.get("Owner.Name") or "Unknown",
        })
    return deduped


def write_csv(events, path):
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        for e in events:
            row = {k: e.get(k, "") for k in COLUMNS}
            row["start_time"] = e["start"].strftime("%Y-%m-%d %H:%M:%S")
            row["end_time"] = e["end"].strftime("%Y-%m-%d %H:%M:%S")
            writer.writerow(row)


def write_xes(events, path):
    by_case = {}
    for e in events:
        by_case.setdefault(e["case_id"], []).append(e)
    case_attrs = ["region", "industry", "lead_source", "outcome", "owner"]
    with open(path, "w") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n<log xes.version="1.0" '
                'xmlns="http://www.xes-standard.org/">\n'
                '  <extension name="Concept" prefix="concept" uri="http://www.xes-standard.org/concept.xesext"/>\n'
                '  <extension name="Time" prefix="time" uri="http://www.xes-standard.org/time.xesext"/>\n'
                '  <extension name="Lifecycle" prefix="lifecycle" uri="http://www.xes-standard.org/lifecycle.xesext"/>\n'
                '  <extension name="Organizational" prefix="org" uri="http://www.xes-standard.org/org.xesext"/>\n')
        for case_id, case_events in by_case.items():
            first = case_events[0]
            f.write(f"  <trace>\n    <string key=\"concept:name\" value={quoteattr(case_id)}/>\n")
            for attr in case_attrs:
                f.write(f"    <string key={quoteattr(attr)} value={quoteattr(str(first[attr]))}/>\n")
            f.write(f"    <float key=\"amount\" value=\"{float(first['amount'])}\"/>\n")
            for e in case_events:
                transitions = [("start", e["start"]), ("complete", e["end"])] if e["end"] > e["start"] \
                    else [("complete", e["start"])]
                for lifecycle, ts in transitions:
                    f.write("    <event>\n"
                            f"      <string key=\"concept:name\" value={quoteattr(e['activity'])}/>\n"
                            f"      <string key=\"org:resource\" value={quoteattr(e['resource'])}/>\n"
                            f"      <string key=\"lifecycle:transition\" value=\"{lifecycle}\"/>\n"
                            f"      <date key=\"time:timestamp\" value=\"{ts.isoformat()}\"/>\n"
                            "    </event>\n")
            f.write("  </trace>\n")
        f.write("</log>\n")


def main():
    raw = json.loads(Path(f"{DATA_DIR}/raw.json").read_text())
    events = build_events(raw)
    write_csv(events, f"{DATA_DIR}/event_log.csv")
    write_xes(events, f"{DATA_DIR}/event_log.xes")
    print(f"  {len(events)} events across {len({e['case_id'] for e in events})} cases "
          f"({raw['source']})")
    print(f"Wrote {DATA_DIR}/event_log.csv and {DATA_DIR}/event_log.xes")


if __name__ == "__main__":
    main()
