# NHO Attendance Checker

A Streamlit app that compares an NHO attendance Excel file against the Miami-Dade onboarding tracker.

## Features

- Accepts attendance files with `N`, `NAME`, and `EMAIL` columns.
- Reads tracker sheets even when their header row or column labels differ.
- Searches the active recruiting sheets and skips archive/reference sheets.
- Matches by exact email first, exact normalized name second, and optional approximate name last.
- Shows the recruiter's tracker sheet, onboarding stage, current requirement values, notes, and every missing process from **Pre-hire Training** through **Badge Photo**.
- Exports the complete comparison to a formatted Excel file.
- Exports a print-ready PDF summary with fingerprint alerts in red.
- Highlights the full row in red when fingerprints are missing, pending, or marked `Info Requested`.
- Processes uploaded files in memory; the app does not intentionally save them.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

Create a new app from this repository, select `app.py` as the entry point, and deploy. Keep the repository private if company policy requires it; the tracker contains personally identifiable information.

## Matching rules

1. Exact normalized email
2. Exact normalized name
3. Approximate name at or above the score selected in the UI (`88` by default)

Approximate matches are labeled **Possible match – review** and should be checked manually.
