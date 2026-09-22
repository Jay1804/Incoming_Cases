# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-page Streamlit app (`Incoming.py`) that generates client case-count reports and emails them out via Outlook (Windows COM automation). It has two independent tabs:

- **Tab 1 — Manual Upload & Split** (legacy/generic): user uploads an Excel file + a distribution list, splits rows by `BM Team Member` or `BM Team Lead`, and emails each split file to the matching person.
- **Tab 2 — Daily / Monthly BM Lead Report** (primary, DB-driven): pulls live data from the MySQL production DB, bifurcates it by BM Team Lead, and builds one workbook per BM Team Lead with two sheets — a dynamic month-to-date daily breakdown and a user-ranged monthly pivot.

## Commands

Run the app (must use the anaconda interpreter — the `python`/`py` aliases on this machine resolve to the Microsoft Store stub, not a real interpreter):

```powershell
& "C:\Users\jay.chaudhary\anaconda3\python.exe" -m streamlit run Incoming.py
```

Install/refresh dependencies:

```powershell
& "C:\Users\jay.chaudhary\anaconda3\python.exe" -m pip install -r requirements.txt
```

There is no test suite, linter, or build step in this repo. To sanity-check a change without going through the Streamlit UI, byte-compile and/or exercise a module directly, e.g.:

```powershell
& "C:\Users\jay.chaudhary\anaconda3\python.exe" -m py_compile Incoming.py queries.py db.py bm_mapping.py
& "C:\Users\jay.chaudhary\anaconda3\python.exe" -c "from queries import fetch_incoming_case_count; import datetime; print(fetch_incoming_case_count(datetime.date.today(), datetime.date.today()))"
```

Windows-only: sending mail depends on `pywin32`'s `win32com.client` driving a local Outlook installation. This cannot be tested headlessly — importing/calling `init_outlook()` will actually launch Outlook COM.

## Architecture

### DB layer (`db.py`, `queries.py`)

- `db.py` builds a SQLAlchemy engine (`mysql+mysqlconnector://...`) from env vars loaded via `python-dotenv`: `DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`. Real credentials live in `.env` (gitignored in spirit, not committed); `.env.example` is the placeholder template — never put real secrets in the `.example` file.
- `queries.py` holds two parametrized SQL queries against the `checkpoint_live` MySQL DB, both scoped to a `start_date`/`end_date` range:
  - `fetch_incoming_case_count` — new cases received per client per day.
  - `fetch_sent_case` — reports sent per client per day.
- **Table naming gotcha**: despite the DB being named `checkpoint_live`, its tables do *not* carry a `checkpoint_live_` prefix (e.g. it's `ec_client`, `ec_case_master`, `ec_case_reports`, not `checkpoint_live_ec_client`). Any new query copied from an external dashboard/BI tool will likely need that prefix stripped.
- **Sargability matters here**: `ec_case_reports` has 30M+ rows. The Sent Case query filters on the raw `report_sent_on` timestamp (`report_sent_on >= start AND report_sent_on < end_exclusive`), *not* `DATE(report_sent_on)`. Wrapping the filter column in `DATE()` defeats the `REPORT_SENT_ON` index and forces a full scan — confirmed via `EXPLAIN` to turn a `range` scan (~5M rows) into a full `index` scan (~30M rows), taking a multi-month range from ~1 min to 9+ min (effectively hung). Keep any future date filters on these large tables sargable (compare against the raw column, use `< end + INTERVAL 1 DAY` instead of wrapping in `DATE()`).
- Both queries were trimmed of several `LEFT JOIN`s (`ec_case_candidates`, `ec_user_details` ×3, `ec_client_process`, `ec_master_company_locations`, `ec_case_fields`) that weren't referenced in `SELECT`/`WHERE`/`GROUP BY` — verified via direct row-count/sum comparison that dropping them doesn't change results, only speeds things up. Don't add a join back without checking it's actually used.

### BM Team Lead mapping (`bm_mapping.py`)

The DB queries have no concept of "BM Team Lead" — that mapping is sourced from a **published Google Sheet** (`MAPPING_SHEET_URL` in `bm_mapping.py`, fetched as CSV via its `/pub?...&output=csv` export URL, not the human-facing `pubhtml` link) maintained outside this app. The sheet's own headers (`Client Code`, `BM`, `BM Lead`) are renamed to this app's convention (`Client Id`, `BM Team Member`, `BM Team Lead`) in `load_bm_mapping()`. The sheet's `Client Code` column also contains non-numeric values (e.g. `AUTH-10`) that never match the DB's numeric client ids — these are coerced to `NaN` via `pd.to_numeric(errors="coerce")` and dropped. `attach_bm_lead()` then left-joins this mapping onto query results by client id; clients with no match are labelled `"Unmapped"` rather than dropped, and Tab 2 excludes `"Unmapped"` from the per-lead emails (but reports the count).

Known gap (as of 2026-09-22): the Google Sheet is missing ~71 client ids that `Incoming_Data.xlsx`'s older mapping had (confirmed by diffing unmapped client ids against that file) — this is a data-completeness gap in the sheet itself, not a join bug. Decision was to use the Google Sheet as the sole source (no fallback to `Incoming_Data.xlsx`) and accept those clients showing as `"Unmapped"` until the sheet is updated.

### Tab 2 report shapes (`Incoming.py`)

- `build_combined_report(incoming_subset, sent_subset)` — outer-joins Incoming and Sent Cases on `Client Code` + date, so a client's sent-only day (no matching incoming case that day) or incoming-only day still gets a row instead of being dropped. Produces the "Daily Incoming vs sent" sheet: `Client Code, Client_name, received_date, Sent Date, Received_count, Sent Case`.
- `build_monthly_pivot(incoming_df, sent_df, bm_lead)` — aggregates by calendar month per client, then pivots wide into one row per client with a `Received- Mon'YY` / `Sent- Mon'YY` column pair per month in the selected range, in chronological order.
- Tab 2's Daily sheet is always hardcoded to "1st of current month → today" (recalculated on every run, no user control); only the Monthly sheet's range is user-selectable.

### Export/Send Mail split (Tab 2 only)

Export and Send Mail are two separate buttons, deliberately decoupled so generated reports can be reviewed before anything is emailed:

- **Export Reports**: queries the DB, builds one workbook per BM Team Lead into a **persistent** temp dir (`tempfile.mkdtemp()`, not a context-managed `TemporaryDirectory`) because the files must survive into a later Streamlit rerun. The dir path, file paths, and the distribution dataframe are stashed in `st.session_state["tab2_*"]`. A previous run's temp dir is `shutil.rmtree`'d before creating a new one.
- **Send Mail**: reads those same session-state-held file paths and sends exactly what was previewed — it does not re-query the DB or rebuild anything.
- Since Streamlit reruns the entire script top-to-bottom on every widget interaction, any state that must outlive a single button click goes through `st.session_state`, not local variables.

### Data files read by the app (not code, but load-bearing)

- `Distribution_list.xlsx` — recipient list with columns `Designation`, `Name`, `Email_ID`. Not read from a hardcoded path; it's uploaded manually through the Streamlit UI on every run (both tabs). Tab 2 only matches rows where `Designation == "BM Team Lead"` (case-insensitive) and `Name` matches the BM Team Lead value from the mapping.
