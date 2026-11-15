# Pak Kret Territory Explorer

Static dashboard. **Contains no data** — it reads a private Google Spreadsheet
that the publisher syncs to, and shows nothing without a Google sign-in.

## Repo layout

```
index.html        the dashboard -- GitHub Pages serves this from repo root
manifest.json      Android/iOS "Add to Home Screen" metadata
icons/             app icons (source + generated apple-touch/192/512 sizes)
backend/           Apps Script web app that serves the Sheet
sync/              local-only scripts: update_territory_sheet.py, fetch_links.py, intake_sources.py
                   test_update.py -- tests for what the update script publishes or deletes
dev/               drive-file-test.html -- OAuth / Sheets API diagnostic,
                   not linked from the dashboard itself
```

The page is viewer-only: all Excel reading happens in `sync/update_territory_sheet.py`.

Tests: `python3 -m unittest discover -s sync -p "test_*.py"`

Design + build notes live outside this repo (`HOSTED_DASHBOARD_PLAN.md`).
