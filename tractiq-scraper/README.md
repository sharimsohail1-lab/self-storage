# TractIQ Hub scraper

Logs into `hub.tractiq.com` and pulls data from two reports:

- **Report 1 — Self-Storage Market Explorer** (`1c055fbe-1836-4950-b70c-ecc712934a6c`): demographics, supply, pipeline, quarterly pricing, quarterly occupancy, and annual financials for National + all 50 largest MSAs.
- **Report 2 — Self-Storage Occupancy** (`7fb00aeb-5f9d-47a1-86e1-48644f34158d`): within-market occupancy spread (P10–P90), supply pressure (pipeline sqft/facility count vs. current occupancy), and rates & pricing, for the top 50 MSAs (ranked as in report 1).

...and writes everything to a single Excel workbook.

**Run this locally, not in a network-restricted sandbox.** `hub.tractiq.com` and `app.tractiq.com` (the OAuth login domain) must both be reachable from wherever you run it.

## How it works

Both reports render inside an `<iframe class="viewer-frame">` whose document
embeds its *entire* underlying dataset as JSON in an inline `<script>` tag —
report 1 as `const MSA_DATA = [...]` (plus `DEMOGRAPHICS`, `SUPPLY`,
`PIPELINE`, `PRICING`, `OCCUPANCY`, `FINANCIALS`), report 2 as
`window.__API__ = {...}`. Rather than driving the dropdown/chart UI and
scraping rendered DOM, this script logs in, loads each report, and parses
that embedded JSON directly. That's faster, gets the full history (not just
whatever's currently rendered), and sidesteps report 2's charts entirely —
which matters because they depend on Chart.js from `cdn.jsdelivr.net`, a
domain that's blocked in some sandboxed environments and isn't needed at all
once you're reading the underlying data instead of the rendered chart.

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

## Step 1: inspect (sanity check)

```bash
node scrape.js inspect
```

Logs in, opens both reports, confirms the expected data sections were found
(prints row counts and top-50 coverage), and writes to `output/inspect/`:
raw iframe HTML for both reports, and the fully parsed JSON
(`report1-parsed.json`, `report2-parsed.json`) for manual review.

If TractIQ changes either report's page structure, this is where it'll show
up — the script throws a clear error naming which expected section
(`MSA_DATA`, `window.__API__`, etc.) it couldn't find.

## Step 2: run the full scrape

```bash
node scrape.js run
```

Produces `output/tractiq_data_<timestamp>.xlsx` with these sheets:

- **Metadata** — scrape timestamp, source URLs, as-of dates per data section.
- **R1 - MSA Reference** — id, name, state, rank (National + top 50).
- **R1 - Demographics** — population (current/2010/2020/projected), households, median HH income, renters/homeowners.
- **R1 - Supply** — facility count, total rentable sqft, REIT/climate-controlled counts.
- **R1 - Pipeline** — under-construction facility count and sqft.
- **R1 - Pricing (Quarterly)** — street/web 10x10 rates, every quarter back to 2019.
- **R1 - Occupancy (Quarterly)** — occupancy %, every quarter back to 2019.
- **R1 - Financials (Annual)** — EGI/NOI/OpEx and margins, per year.
- **R2 - Within-Market Spread** — P10/P25/P50/P75/P90 occupancy and P90−P10 spread, top 50 MSAs.
- **R2 - Supply Pressure** — pipeline sqft/facility count paired with current occupancy, top 50 MSAs.
- **R2 - Rates & Pricing** — street/web 10x10 rates, discount %, top 50 MSAs.

All report-1 sheets share one global "as of" date (the page's own "Last
updated" banner). Report-2 sheets each carry their own as-of date, since the
underlying dispersion/supply/pricing datasets are refreshed independently.

A handful of top-50 MSAs are legitimately missing from report 2's
within-market spread (44/50 covered) and supply (49/50) sections in the
source data itself — TractIQ suppresses cells below a minimum sample size
(see the `pricing.meta.notes` value in the Metadata sheet). Those rows are
left blank in the workbook rather than guessed at.

## Security note

Since the TractIQ password was shared in chat to get this built, rotate it
once you're done using this script.
