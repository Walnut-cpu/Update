# Literature Search Workflow

## Goal

Build a repeatable loop for disease-centered literature retrieval:

1. Read disease names from `Disease_list.xlsx`
2. Auto-generate candidate search strategies for new diseases
3. Let the user review and edit the strategies
4. Cache approved strategies locally
5. Reuse approved strategies on later runs
6. Generate new candidate strategies only for unseen diseases

## Files

- `fetch_med_lit.py`
  Main workflow script
- `search_strategy_store.json`
  Local cache of approved search strategies
- `search_strategy_review.json`
  Pending review items that the user can approve or edit
- `search_strategy_review.md`
  Human-readable review summary
- `DOI_list.xlsx`
  Retrieval output workbook. Each result row stores the date when that literature item was newly recorded
- `Markdown/`
  Saved abstract markdown files. New literature found on a given day is appended into a date-named file such as `2026-07-02.md`

## Loop

### First run

Run:

```powershell
& 'C:\Users\Siyani Chen\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' 'D:\Data\Graduate_Content\AI_Agent\Codex\fetch_med_lit.py'
```

If you want a local-only preparation pass before any network retrieval, run:

```powershell
& 'D:\Data\Graduate_Content\AI_Agent\Codex\run_literature_strategy_sync.cmd'
```

If a disease has no approved local strategy yet, the script will:

- generate a candidate strategy
- save it into `search_strategy_review.json`
- render a readable summary into `search_strategy_review.md`
- stop before formal retrieval for that disease

### Review step

Open `search_strategy_review.json`, then change:

- `"status": "pending"` to `"status": "approved"` if the strategy is acceptable
- or `"status": "rejected"` if it should not be used

You can also edit:

- `queries.openalex_terms`
- `queries.europe_pmc_query`
- `review_comment`

### Second run

Rerun `fetch_med_lit.py`.

The script will:

- import approved strategies into `search_strategy_store.json`
- reuse approved strategies directly
- skip diseases that are still pending
- generate new pending strategies only for previously unseen diseases
- compare newly retrieved literature against existing `DOI_list.xlsx` entries, using DOI first and identifier second
- append only genuinely new literature to `DOI_list.xlsx`
- append the new literature package into the date-named markdown file for that run in `Markdown/`

## Low-Approval Codex Usage

When Codex runs under a managed permission policy, persistent "approve for me" rules may be blocked even when one-time approval still works. In that case:

- use `run_literature_strategy_sync.cmd` for local strategy generation and cache sync
- edit and approve `search_strategy_review.json`
- use `run_fundigraph_extract.cmd` and `run_fundigraph_apply.cmd` for the FundiGraph review loop
- reserve one-time approval for the final network retrieval run of `fetch_med_lit.py`

## Example

### Example candidate

```json
{
  "disease_name": "Age-related Macular Degeneration",
  "canonical_key": "age_related_macular_degeneration",
  "status": "pending",
  "queries": {
    "openalex_terms": [
      "Age-related Macular Degeneration",
      "Age related Macular Degeneration",
      "AMD",
      "neovascular age-related macular degeneration",
      "nAMD",
      "dry AMD",
      "wet AMD",
      "non-neovascular age-related macular degeneration"
    ],
    "europe_pmc_query": "(\"Age-related Macular Degeneration\" OR \"Age related Macular Degeneration\" OR AMD OR \"neovascular age-related macular degeneration\" OR nAMD OR \"dry AMD\" OR \"wet AMD\" OR \"non-neovascular age-related macular degeneration\") AND FIRST_PDATE:[2026-06-06 TO *] AND OPEN_ACCESS:y AND HAS_ABSTRACT:y"
  }
}
```

### Example approval

```json
{
  "disease_name": "Age-related Macular Degeneration",
  "canonical_key": "age_related_macular_degeneration",
  "status": "approved",
  "review_comment": "Approved after checking that all terms remain disease-name-centered."
}
```

## Current rule

Strategies must stay disease-name-centered:

- allowed: disease name, standard abbreviation, direct disease subtype names
- not allowed: exam terms, imaging terms, biomarker terms, generic retina terms, treatment-only terms
