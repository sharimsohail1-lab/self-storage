# TractIQ Hub scraper

Logs into `hub.tractiq.com`, scrapes:

- **Report 1** (`1c055fbe-1836-4950-b70c-ecc712934a6c`): every MSA in the dropdown, its data table, and "as of" date.
- **Report 2** (`7fb00aeb-5f9d-47a1-86e1-48644f34158d`): within-market spread, supply pressure, and pipeline for the top 50 MSAs.

...and writes both to a single Excel workbook.

**Run this locally, not in a network-restricted sandbox.** `hub.tractiq.com` must be reachable from wherever you run it.

## Why there's an "inspect" step

This script was written without ever being able to load hub.tractiq.com (the
environment it was built in has no network access to that host). The login
form, MSA dropdown, and table markup are all best-effort guesses based on
common dashboard patterns — they need to be verified against the real page
once, which is what `inspect` mode is for.

## Setup

```bash
cd tractiq-scraper
npm install
npx playwright install chromium
```

Set your credentials as environment variables (don't put them in a file that gets committed):

```bash
export TRACTIQ_EMAIL="you@example.com"
export TRACTIQ_PASSWORD="..."
```

## Step 1: inspect

```bash
node scrape.js inspect
```

This logs in, opens both reports, and writes to `output/inspect/`:

- Screenshots and HTML dumps of the login page and both reports
- The MSA dropdown options it found (`report1-dropdown-options.json`, `report2-dropdown-options.json`)
- Every `<table>` it found on each page (`report1-tables.json`, `report2-tables.json`) with row/column counts and headers logged to the console
- Whether it could detect an "as of" date

Review the console output and, if something wasn't found (dropdown, login
fields, tables), open the corresponding screenshot/HTML in `output/inspect/`
to find the real selector, then update the relevant `*_SELECTORS` /
`*_HINT_SELECTORS` array near the top of `scrape.js`. Selectors are tried in
order, first match wins.

If report 2 turns out to already have a single table listing all MSAs with
spread/supply/pipeline columns, the script will use it directly and rank
by row order (assumed already sorted "top" first — adjust `scrapeReport2`
if it needs an explicit sort). If it's a per-MSA drilldown like report 1
instead, the script falls back to walking the dropdown in order and treats
the first 50 entries as "top 50" — adjust if the tool's own ranking lives
elsewhere (e.g. a separate sort control).

## Step 2: run the full scrape

```bash
node scrape.js run
```

Produces `output/tractiq_data_<timestamp>.xlsx` with these sheets:

- **Report1 - All MSAs**: one row per MSA per data row, with MSA name and as-of date prepended. Assumes the table shape is consistent across MSAs (columns come from the first successfully-scraped MSA).
- **Report1 - Raw**: every table, every row, exactly as found — a safety net in case the structured sheet above doesn't line up cleanly.
- **Report2 - Top50**: MSA, within-market spread, supply pressure, pipeline.
- **Report2 - Raw**: populated only if report 2 required the per-MSA drilldown fallback — every table found on each MSA's page, in case the regex-based metric extraction missed something.
- **Metadata**: scrape timestamp, source URLs, which extraction mode report 2 used.

If any individual MSA fails to scrape, it's recorded with an error message
in the sheet rather than aborting the whole run, and a screenshot is saved
to `output/inspect/report{1,2}-failed-<msa>.png` for debugging.

## Security note

Since the TractIQ password was shared in chat to get this built, rotate it
once you're done using this script.
