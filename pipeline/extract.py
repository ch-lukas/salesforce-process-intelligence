"""Extract opportunity lifecycle data from a Salesforce org into data/raw.json.

Joins four objects into one lifecycle per Opportunity:
  Lead (ConvertedOpportunityId) -> Opportunity -> OpportunityHistory (stage changes) -> Task (activities)

Usage:
  python3 pipeline/extract.py [--target-org my-trial]
"""
import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from sf_client import SalesforceClient

QUERIES = {
    "opportunities": """
        SELECT Id, Name, Amount, StageName, IsWon, IsClosed,
               LeadSource, CreatedDate, CloseDate, OwnerId, Owner.Name,
               Account.Industry, Account.BillingCountry
        FROM Opportunity""",
    "stage_history": """
        SELECT OpportunityId, StageName, CreatedDate,
               CreatedById, CreatedBy.Name
        FROM OpportunityHistory
        ORDER BY OpportunityId, CreatedDate""",
    "tasks": """
        SELECT WhatId, Subject, CreatedDate, CompletedDateTime,
               OwnerId, Owner.Name, Status
        FROM Task
        WHERE What.Type = 'Opportunity'""",
    "leads": """
        SELECT Id, CreatedDate, ConvertedDate, ConvertedOpportunityId,
               OwnerId, Owner.Name
        FROM Lead
        WHERE IsConverted = true
          AND ConvertedOpportunityId != null""",
}


def flatten(record):
    """Drop SOQL 'attributes' noise and turn Owner.Name style lookups into flat keys."""
    out = {}
    for key, value in record.items():
        if key == "attributes":
            continue
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                if sub_key != "attributes":
                    out[f"{key}.{sub_key}"] = sub_value
        else:
            out[key] = value
    return out


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--target-org")
    parser.add_argument("--out", default="data/raw.json")
    args = parser.parse_args()

    sf = SalesforceClient(args.target_org)
    # Record where the data came from without writing the org's URL into shareable files.
    raw = {"source": "salesforce", "extracted_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
    for name, soql in QUERIES.items():
        raw[name] = [flatten(r) for r in sf.query_all(" ".join(soql.split()))]
        print(f"  {name:15} {len(raw[name]):>6} records")

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out).write_text(json.dumps(raw, indent=1))
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
