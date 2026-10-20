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
dev/               drive-file-test.html -- OAuth / Sheets API diagnostic,
                   not linked from the dashboard itself
```

Design + build notes live outside this repo (`HOSTED_DASHBOARD_PLAN.md`).
