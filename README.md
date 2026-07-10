# Merz Legal Crystal Ball — Searchable Web App

A fully static, single-page search app for the pharmaceutical compliance tracker
`The_Merz_Legal_Crystal_Ball_.xlsm`. The workbook itself is **never** committed
(it's git-ignored); only the extracted `docs/data.json` is versioned.

- **`parse.py`** — reads the workbook read-only and regenerates `docs/data.json`.
- **`docs/index.html`** — the app (HTML/CSS/JS in one file, no build step, no
  framework). It fetches `data.json` with relative paths, so it works both from a
  local server and on GitHub Pages.

## What's in the data

Records are pulled from the enforcement, litigation, and channel sheets:
DOJ Settlements, OPDP Warning Letters, OPDP Untitled Letters, Direct Quotes,
Pharma v Pharma Lawsuits, NAD Challenges, and the channel views (TV Commercials,
YouTube/Podcasts, Online Video, Social Media, Banner & Display, Sales Aids &
Print, Email, Conferences, Investor & Press), plus the Glossary. Overview,
analysis, and backend/hidden sheets are skipped.

## Update the data after editing the tracker

The parser opens the workbook **read-only** (`data_only=True`) and never modifies
or re-saves it. After you change the `.xlsm`, make sure Excel has saved it (the
parser reads cached cell values), then rerun:

```bash
python3 parse.py                                  # auto-finds the workbook
# or point it at an explicit copy:
python3 parse.py /path/to/The_Merz_Legal_Crystal_Ball_.xlsm
```

It auto-locates the workbook next to `parse.py`, in `~/Downloads`, or in the
current directory. It prints a per-sheet record count so you can check the totals
against the workbook's own live counters, and rewrites `docs/data.json`.

Requires `openpyxl`:

```bash
python3 -m pip install --user openpyxl
```

Commit the regenerated `docs/data.json` to publish the update.

## Run the app locally

Serve the `docs/` folder over HTTP (a plain `file://` open won't work because the
app fetches `data.json`):

```bash
python3 -m http.server 8000 --directory docs
```

Then open <http://localhost:8000/>.

## Features

- One global search box — case-insensitive partial matching across every field,
  updating as you type. Glossary terms are searched too and shown as
  distinctly-styled cards.
- Filter chips for source sheet, year, and violation type / doctrine, plus a
  company dropdown. Filters combine with the search box.
- Result cards show the sheet, company/parties, year, and a snippet with the
  match highlighted; click a card to expand the full record.
- A live match count (e.g. "34 results across 6 sheets").

## Hosting on GitHub Pages

The app uses relative paths only, so it can be served from the `docs/` folder on
the `main` branch (Settings → Pages → Source: `main` / `/docs`). **Note: GitHub
Pages sites are public**, even from a private repo.
