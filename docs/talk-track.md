# Talk track: 8 slides, about 8 minutes

**Frame (20s).** "This shows how I'd open with a Salesforce-heavy prospect: start from their own CRM data, find where the process leaks, frame what fixing it could be worth, and be explicit about what still needs validating."

**1. Title.** "It's a simulated value case with a working Salesforce extraction pipeline. I'll be clear throughout about which parts are illustrative."

**2. Illustrative upside (60s).** "$1.4M is a gross scenario at 50% capture across three overlapping opportunities. The full gap is $2.8M. The data is 800 simulated opportunities in Salesforce's object shape; the same pipeline runs live on 40 seeded opportunities in a trial org."

**3. Approach (45s).** "Identify, Frame, Commit, Realize. Today stops at Frame on purpose: the assumptions are illustrative until the process owners and Finance sign them off."

**4. Discover (60s).** "Only 14% of deals match the selected winning trace, which is 38% of won deals. That's an exact-trace measure, not a conformance score: lost deals and legitimate approval steps don't match it." Click through to the live process map.

**5. Diagnose (60s).** "Three patterns are associated with lower win rates: EMEA approvals 5.8 days slower, re-quoted deals 18 points lower, skipped qualification 21 points lower. Cohort sizes are on the slide. These are associations, not proven causes."

**6. Business case overview (60s).** "Three opportunities with proposed owners and full-gap scenarios. They overlap: 22 deals are both re-quoted and unqualified. So the total is gross, and the checklist shows what has to happen before any value is committed."

**7. Value case (75s).** "The deep-dive: 747 negotiated deals a year × 19.2% re-quoted × a 17.5-point gap × $53.1k × 50% capture ≈ $667k, which is exactly half the $1.33M full gap. Tracking uses a frozen baseline and a comparison group, and Finance validates realized revenue."

**8. Next steps (30s).** "Validate on production data, pilot only once the causes are confirmed, then measure against the baseline with Finance sign-off."

**Tech setup deck (if asked).** Show the org, the stage-history SOQL and the object-to-event mapping. Opportunity Id is the case; tasks join on WhatId and leads on ConvertedOpportunityId.

## Tough questions

1. **"Did you find $1.4M in a real org?"** No. It's an illustrative scenario from simulated data. The live trial-org extract proves the pipeline, not the numbers.
2. **"Can we add the three opportunities up?"** Only as a gross scenario. The cohorts overlap (22, 7 and 5 deals), so I'd count incremental wins once per deal before committing.
3. **"Why would fewer re-quotes raise win rate?"** Today it's an association. I'd capture re-quote reasons, control for deal mix, and test it in a pilot.
4. **"Does 14% mean 86% of deals break the process?"** No. It's an exact match to one winning trace. Real conformance checking has to allow loss exits and conditional steps.
5. **"Why a 30% approval threshold?"** I've dropped that from the recommendation. Finance sets thresholds, and only after we know how many re-quotes approval delays actually cause.
6. **"How would you prove realized value?"** Freeze the baseline and cohort, measure the operational change against a comparison group, and get Finance sign-off on the revenue.
7. **"What changes for a production org?"** Validate activity semantics first (task timestamps can include queueing; open tasks have no completion), then add incremental refresh, stable event IDs, reconciliation and Bulk API for volume.
8. **"Have you imported it into a process mining tool?"** The pipeline exports CSV and XES with the standard case, activity, timestamp and resource columns; the import itself is the next thing to verify.
