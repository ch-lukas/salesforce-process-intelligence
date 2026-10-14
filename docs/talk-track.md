# Talk track: 5 minutes with the VPs

**Frame (30s).** "I wanted to show how I'd approach a first conversation with a
Salesforce-heavy prospect: start with their own CRM data, find where the process
leaks, and turn that into a number the CRO cares about."

**1. The answer first (45s).** Point at the headline figure. "Three process issues,
about $1.4M a year if we close half the gap, $2.8M at full gap. Everything below is
how I got there and how you'd check it."

**2. Discover (60s).** Process map, Frequency view. "Fourteen percent of deals follow
the designed path. There are 67 variants across 800 deals." Switch to Wait time.
"Amber is idle time. Discount Approved and Finance Approval light up." Drag the
slider: "Same abstraction idea as Apromore's path slider."

**3. Diagnose (60s).** The three cards.
- EMEA approvals: 7.0 days vs 1.2 in AMER. Same policy, one approver.
- Re-quotes: one in five negotiated deals loops back to Proposal; they win 31% vs 49%.
- Skipped Qualification: 12% of deals, win rate halves.

**4. Value engineering (60s).** Value tree. "Objective → driver → process KPI → dollars.
The capture sliders are the conversation with the customer: they own the assumption,
we own the evidence." Mention the overlap caveat before they do.

**5. How it's built (30s).** "SOQL over REST using the CLI session, no stored
credentials. Opportunity Id is the case; Leads join via ConvertedOpportunityId, Tasks
via WhatId. The output is a standard CSV/XES you can drop straight into Apromore."
Be upfront: trial orgs can't backdate history, so the dataset is simulated in the
exact Salesforce shape; the extractor runs against a live org unchanged.

**Close (15s).** "In a real engagement: validate on the production org with RevOps,
pilot a delegated approval threshold, then monitor the three KPIs in Apromore
and report value realised quarterly."

## Questions to expect

- *Object-centric?* Opportunity is the case here; Quote, Approval and Case would make a
  natural object-centric extension (OCEL), since approvals relate to quotes, not deals.
- *Data volume / refresh?* Bulk API 2.0 for large orgs; nightly incremental extract on
  `SystemModstamp`.
- *Field history limits?* OpportunityHistory is always on; other objects need Field
  History Tracking enabled (20 fields per object, 18-24 month retention).
- *Causality?* Win-rate gaps are correlations between cohorts; the pilot is how you
  prove the lift.
