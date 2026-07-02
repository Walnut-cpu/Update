# FundiGraph Review - 2026-07-01

Please complete manual review before applying any workbook update.

- Extracted candidates: `3`
- Dropped by overlap >= 0.80: `2`
- Sent to Gemini/manual review: `1`
- Gemini precheck status: `local_fallback`
- Gemini model: `gemini-3.1-flash-lite`
- Gemini correct/error: `0` / `1`
- Gemini report: `FundiGraph_gemini_precheck_2026-07-01.md`

Review steps:
- Mark correct items by changing `Confirm` to `[x]`.
- Mark incorrect items by changing `Error` to `[x]`, then edit the `Corrected triple` line directly.
- After review, run `python fundigraph_review_workflow.py apply`.

## Item 1

- Source file: `001_Age-related_Macular_Degeneration_2026-06-20_OphthaDT.md`
- Confidence: `0.72`
- Triple: `(: Disease {name: Age-related Macular Degeneration})-[: Requires examination]->(: Examination {name: Best Corrected Visual Acuity (BCVA)})`
- Evidence: `The study serializes longitudinal histories to forecast best corrected visual acuity (BCVA).`
- Workbook overlap: `False`
- Best matched value: `Optical coherence tomography`
- Best matched location: `row 5510, col M`
- Gemini verdict: `error`
- Gemini reason: `Fallback heuristic: lower-confidence extraction without strong workbook overlap.`
- Planned action: `needs_review_for_insert`

- Confirm: [ ]
- Error: [ ]
- Corrected triple: `(: Disease {name: Age-related Macular Degeneration})-[: Requires examination]->(: Examination {name: Best Corrected Visual Acuity (BCVA)})`
