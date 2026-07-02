# Gemini Precheck - 2026-07-01

- Status: `local_fallback`
- Model: `gemini-3.1-flash-lite`
- Correct count: `0`
- Error count: `1`
- Reason: `Gemini request failed; switched to local fallback. Original error: <urlopen error [WinError 10013] 以一种访问权限不允许的方式做了一个访问套接字的尝试。>`
- Note: `This is a local heuristic fallback, not a live Gemini response.`

## Item 1

- Triple: `(: Disease {name: Age-related Macular Degeneration})-[: Requires examination]->(: Examination {name: Best Corrected Visual Acuity (BCVA)})`
- Gemini verdict: `error`
- Reason: `Fallback heuristic: lower-confidence extraction without strong workbook overlap.`
