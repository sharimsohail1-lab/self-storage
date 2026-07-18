#!/usr/bin/env node
/**
 * TractIQ Hub scraper.
 *
 * Run this LOCALLY (not in a network-restricted cloud sandbox) — hub.tractiq.com
 * must be reachable from wherever you run it.
 *
 * Setup:
 *   cd tractiq-scraper
 *   npm install
 *   npx playwright install chromium
 *
 * Credentials are read from the environment, never hardcoded here:
 *   export TRACTIQ_EMAIL="you@example.com"
 *   export TRACTIQ_PASSWORD="..."
 *
 * How this works: both reports render inside an <iframe class="viewer-frame">
 * whose document embeds its full underlying dataset as a JSON blob in a
 * <script> tag (report 1: `const MSA_DATA = [...]` etc.; report 2:
 * `window.__API__ = {...}`). Rather than driving the dropdown/chart UI and
 * scraping rendered DOM (fragile, and report 2's charts require a
 * jsdelivr-hosted Chart.js that's often blocked by sandboxed egress
 * policies), this pulls the data straight from that embedded JSON — faster,
 * complete (full history, not just the currently-rendered view), and immune
 * to chart-rendering failures.
 *
 * Step 1 — sanity check:
 *   node scrape.js inspect
 * Logs in, opens both reports, and confirms the expected data sections are
 * found, dumping raw frame HTML + parsed JSON to output/inspect/ for review.
 *
 * Step 2 — full scrape:
 *   node scrape.js run
 * Produces output/tractiq_data_<timestamp>.xlsx.
 */

const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');

const BASE_URL = 'https://hub.tractiq.com';
const REPORT_1_URL = `${BASE_URL}/#report/1c055fbe-1836-4950-b70c-ecc712934a6c`;
const REPORT_2_URL = `${BASE_URL}/#report/7fb00aeb-5f9d-47a1-86e1-48644f34158d`;

const EMAIL = process.env.TRACTIQ_EMAIL;
const PASSWORD = process.env.TRACTIQ_PASSWORD;

const OUT_DIR = path.join(__dirname, 'output');
const INSPECT_DIR = path.join(OUT_DIR, 'inspect');

const LOGIN_EMAIL_SELECTORS = ['input[type="email"]', 'input[name="email"]', 'input#email'];
const LOGIN_PASSWORD_SELECTORS = ['input[type="password"]', 'input[name="password"]', 'input#password'];
const LOGIN_SUBMIT_SELECTORS = ['input[type="submit"]', 'button[type="submit"]', 'button:has-text("Log in")'];
const OAUTH_START_SELECTORS = [
  'a.login-button',
  'a[href="/api/oauth/authorize"]',
  'a:has-text("Log in with TractIQ")',
];
const AUTHORIZE_SELECTORS = ['button:has-text("Authorize")', 'input[value="Authorize"]', 'a:has-text("Authorize")'];

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function assertCredentials() {
  if (!EMAIL || !PASSWORD) {
    console.error('Missing credentials. Set TRACTIQ_EMAIL and TRACTIQ_PASSWORD environment variables before running.');
    process.exit(1);
  }
}

function isRunningBehindProxy() {
  return Boolean(process.env.HTTPS_PROXY || process.env.https_proxy);
}

function resolveChromiumExecutable() {
  // Some environments pre-install a pinned Chromium build and block network
  // access to Playwright's own browser CDN. If PLAYWRIGHT_BROWSERS_PATH points
  // at one, use it directly instead of letting Playwright try to download.
  const candidate = process.env.PLAYWRIGHT_BROWSERS_PATH
    ? path.join(process.env.PLAYWRIGHT_BROWSERS_PATH, 'chromium')
    : null;
  if (candidate && fs.existsSync(candidate)) return candidate;
  return undefined;
}

function resolveChromiumLaunchOptions() {
  const options = { headless: true, executablePath: resolveChromiumExecutable(), args: [] };
  if (isRunningBehindProxy()) {
    options.proxy = { server: process.env.HTTPS_PROXY || process.env.https_proxy };
    // Some TLS-terminating egress proxies reset the connection on Chromium's
    // TLS 1.3 ClientHello (its post-quantum hybrid key share makes it large
    // enough to span multiple TCP segments, which such proxies mishandle).
    // Forcing TLS 1.2 avoids it. Only applied when a proxy is detected — a
    // normal machine reaching the site directly doesn't need this downgrade.
    options.args.push('--ssl-version-max=tls1.2');
  }
  return options;
}

function contextOptionsForCurrentEnv() {
  return isRunningBehindProxy() ? { ignoreHTTPSErrors: true } : {};
}

async function firstVisible(page, selectors, timeoutEachMs = 2000) {
  for (const selector of selectors) {
    try {
      const locator = page.locator(selector).first();
      await locator.waitFor({ state: 'visible', timeout: timeoutEachMs });
      return { locator, selector };
    } catch (_) {
      // try next
    }
  }
  return null;
}

async function debugDump(page, label) {
  ensureDir(INSPECT_DIR);
  const safeLabel = label.replace(/[^a-z0-9_-]/gi, '_');
  try {
    await page.screenshot({ path: path.join(INSPECT_DIR, `${safeLabel}.png`), fullPage: true });
  } catch (_) {}
  try {
    fs.writeFileSync(path.join(INSPECT_DIR, `${safeLabel}.html`), await page.content());
  } catch (_) {}
}

async function login(page) {
  console.log(`Logging in as ${EMAIL} ...`);
  await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
  await sleep(500);

  const oauthStart = await firstVisible(page, OAUTH_START_SELECTORS, 3000);
  if (oauthStart) {
    await Promise.all([
      page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {}),
      oauthStart.locator.click(),
    ]);
    await sleep(500);
  }

  const emailField = await firstVisible(page, LOGIN_EMAIL_SELECTORS, 8000);
  if (!emailField) {
    await debugDump(page, 'login-page-no-email-field');
    throw new Error(
      'Could not find an email/username field on the login page. See output/inspect/login-page-no-email-field.html/.png.'
    );
  }
  await emailField.locator.fill(EMAIL);

  const passwordField = await firstVisible(page, LOGIN_PASSWORD_SELECTORS, 4000);
  if (!passwordField) {
    await debugDump(page, 'login-page-no-password-field');
    throw new Error('Could not find a password field on the login page.');
  }
  await passwordField.locator.fill(PASSWORD);

  const submitButton = await firstVisible(page, LOGIN_SUBMIT_SELECTORS, 4000);
  if (!submitButton) {
    await debugDump(page, 'login-page-no-submit-button');
    throw new Error('Could not find a submit/login button.');
  }
  await Promise.all([
    page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {}),
    submitButton.locator.click(),
  ]);
  await sleep(1500);

  // OAuth consent screen ("Reports Gallery wants to use your TractIQ
  // account...") appears after credentials are accepted. Click through it.
  const authorizeButton = await firstVisible(page, AUTHORIZE_SELECTORS, 4000);
  if (authorizeButton) {
    console.log('OAuth consent screen detected, authorizing...');
    await Promise.all([
      page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {}),
      authorizeButton.locator.click(),
    ]);
    await sleep(1000);
  }

  const stillOnPasswordField = await page
    .locator(LOGIN_PASSWORD_SELECTORS.join(','))
    .first()
    .isVisible()
    .catch(() => false);
  if (stillOnPasswordField) {
    await debugDump(page, 'login-may-have-failed');
    throw new Error(
      'A password field is still visible after submitting — login failed (wrong credentials, or an ' +
        'unhandled 2FA/CAPTCHA step). Check output/inspect/login-may-have-failed.png.'
    );
  }
  console.log('Login succeeded.');
}

async function getReportFrame(page, url, { timeoutMs = 20000 } = {}) {
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle', { timeout: timeoutMs }).catch(() => {});
  const iframeLocator = page.locator('iframe.viewer-frame');
  await iframeLocator.waitFor({ state: 'attached', timeout: timeoutMs });
  const handle = await iframeLocator.elementHandle();
  const frame = await handle.contentFrame();
  if (!frame) throw new Error(`Could not resolve content frame for ${url}`);
  await frame.waitForLoadState('networkidle', { timeout: timeoutMs }).catch(() => {});
  await sleep(1000);
  return frame;
}

// ---------------------------------------------------------------------------
// Report 1 (Self-Storage Market Explorer): embedded `const NAME = ...;`
// declarations inside the iframe's inline <script>.
// ---------------------------------------------------------------------------

function extractJsConst(scriptSrc, name) {
  const match = scriptSrc.match(new RegExp(`const\\s+${name}\\s*=\\s*(\\[[\\s\\S]*?\\]|\\{[\\s\\S]*?\\}|[0-9.]+)\\s*;`));
  if (!match) return undefined;
  return JSON.parse(match[1]);
}

async function extractReport1Data(page) {
  const frame = await getReportFrame(page, REPORT_1_URL);
  const frameHtml = await frame.content();
  ensureDir(INSPECT_DIR);
  fs.writeFileSync(path.join(INSPECT_DIR, 'report1-frame.html'), frameHtml);

  const dataStoreMatch = frameHtml.match(/<script(?:(?!src=)[^>])*>([\s\S]*?MSA_DATA[\s\S]*?)<\/script>/);
  if (!dataStoreMatch) {
    await debugDump(page, 'report1-no-datastore');
    throw new Error(
      'Could not find the embedded MSA_DATA script in report 1. The page structure may have changed — ' +
        'see output/inspect/report1-frame.html and report1-no-datastore.png.'
    );
  }
  const scriptSrc = dataStoreMatch[1];

  const msaData = extractJsConst(scriptSrc, 'MSA_DATA');
  const demographics = extractJsConst(scriptSrc, 'DEMOGRAPHICS');
  const supply = extractJsConst(scriptSrc, 'SUPPLY');
  const pipeline = extractJsConst(scriptSrc, 'PIPELINE');
  const pricing = extractJsConst(scriptSrc, 'PRICING');
  const occupancy = extractJsConst(scriptSrc, 'OCCUPANCY');
  const financials = extractJsConst(scriptSrc, 'FINANCIALS');
  const usSqftPerCapita = extractJsConst(scriptSrc, 'US_SQFT_PER_CAPITA');

  if (!msaData || !demographics || !supply || !pipeline || !pricing || !occupancy || !financials) {
    throw new Error(
      'Report 1 data store was found but one or more expected sections (MSA_DATA/DEMOGRAPHICS/SUPPLY/' +
        'PIPELINE/PRICING/OCCUPANCY/FINANCIALS) failed to parse. See output/inspect/report1-frame.html.'
    );
  }

  const lastUpdatedMatch = frameHtml.match(/Last updated:\s*<strong>([^<]+)<\/strong>/i);
  const lastUpdated = lastUpdatedMatch ? lastUpdatedMatch[1].trim() : null;

  console.log(
    `Report 1: ${msaData.length} markets (incl. National), last updated ${lastUpdated || 'unknown'}, ` +
      `US sqft/capita ${usSqftPerCapita}`
  );

  return { msaData, demographics, supply, pipeline, pricing, occupancy, financials, usSqftPerCapita, lastUpdated };
}

// ---------------------------------------------------------------------------
// Report 2 (Self-Storage Occupancy): `window.__API__ = {...}` in the iframe.
// ---------------------------------------------------------------------------

async function extractReport2Data(page) {
  const frame = await getReportFrame(page, REPORT_2_URL);
  const frameHtml = await frame.content();
  ensureDir(INSPECT_DIR);
  fs.writeFileSync(path.join(INSPECT_DIR, 'report2-frame.html'), frameHtml);

  const match = frameHtml.match(/window\.__API__\s*=\s*(\{[\s\S]*?\});/);
  if (!match) {
    await debugDump(page, 'report2-no-api-data');
    throw new Error(
      'Could not find window.__API__ embedded data in report 2. The page structure may have changed — ' +
        'see output/inspect/report2-frame.html and report2-no-api-data.png.'
    );
  }
  const api = JSON.parse(match[1]);
  const required = ['national', 'state', 'msa', 'dispersion', 'supply', 'msa-names', 'pricing'];
  const missing = required.filter((k) => !(k in api));
  if (missing.length > 0) {
    throw new Error(`Report 2 embedded data is missing expected section(s): ${missing.join(', ')}`);
  }

  console.log(
    `Report 2: dispersion=${api.dispersion.data.length} rows, supply=${api.supply.data.length} rows, ` +
      `pricing.msa=${api.pricing.msa.length} rows, msa-names=${Object.keys(api['msa-names']).length} entries`
  );

  return api;
}

// ---------------------------------------------------------------------------
// Shape report 2's data down to the top-50 MSAs (as ranked in report 1) for
// within-market spread, supply pressure + pipeline, and rates & pricing.
// ---------------------------------------------------------------------------

function buildReport2Top50Sheets(report1Data, report2Api) {
  const top50 = report1Data.msaData.filter((m) => m.id !== 'NATIONAL').sort((a, b) => a.rank - b.rank);
  const nameFor = (id) => report2Api['msa-names'][id] || top50.find((m) => m.id === id)?.name || id;

  const dispersionById = new Map(report2Api.dispersion.data.map((r) => [r.msa_id, r]));
  const supplyById = new Map(report2Api.supply.data.map((r) => [r.msa_id, r]));
  const pricingById = new Map(report2Api.pricing.msa.map((r) => [r.id, r]));

  // Latest occupancy per MSA from the top-level `msa` time series (used to
  // pair current occupancy with pipeline for supply pressure, matching the
  // report's own "Current Occupancy (<latest month>)" framing).
  const latestOccByMsa = new Map();
  for (const row of report2Api.msa.data) {
    const existing = latestOccByMsa.get(row.msa_id);
    if (!existing || row.year_and_month > existing.year_and_month) latestOccByMsa.set(row.msa_id, row);
  }

  const spreadRows = top50.map((m) => {
    const d = dispersionById.get(m.id);
    return {
      rank: m.rank,
      msa_id: m.id,
      msa: nameFor(m.id),
      state: m.state,
      n_facilities: d ? d.n : null,
      p10_occ: d ? d.p10 : null,
      p25_occ: d ? d.p25 : null,
      p50_occ_median: d ? d.p50 : null,
      p75_occ: d ? d.p75 : null,
      p90_occ: d ? d.p90 : null,
      avg_occ: d ? d.avg_occ : null,
      within_market_spread_p90_minus_p10: d ? Number((d.p90 - d.p10).toFixed(4)) : null,
      period: report2Api.dispersion.period,
      as_of: report2Api.dispersion.lastUpdated,
    };
  });

  const supplyRows = top50.map((m) => {
    const s = supplyById.get(m.id);
    const occ = latestOccByMsa.get(m.id);
    return {
      rank: m.rank,
      msa_id: m.id,
      msa: nameFor(m.id),
      state: m.state,
      pipeline_sqft: s ? s.pipeline_sqft : null,
      pipeline_facility_count: s ? s.pipeline_count : null,
      current_occ: occ ? occ.avg_occ : null,
      current_occ_period: occ ? occ.year_and_month : null,
      as_of: report2Api.supply.lastUpdated,
    };
  });

  const pricingRows = top50.map((m) => {
    const p = pricingById.get(m.id);
    const occ = latestOccByMsa.get(m.id);
    const discount = p && p.street ? Number(((p.street - p.web) / p.street).toFixed(4)) : null;
    return {
      rank: m.rank,
      msa_id: m.id,
      msa: nameFor(m.id),
      state: m.state,
      street_rate_per_sqft: p ? p.street : null,
      web_rate_per_sqft: p ? p.web : null,
      web_discount_pct: discount,
      n_facilities_priced: p ? p.n : null,
      current_occ: occ ? occ.avg_occ : null,
      unit_size: '10x10',
      period_label: report2Api.pricing.meta.period_label,
      as_of: report2Api.pricing.meta.lastUpdated,
    };
  });

  return { spreadRows, supplyRows, pricingRows };
}

// ---------------------------------------------------------------------------
// Excel export
// ---------------------------------------------------------------------------

function addSheet(workbook, name, columns, rows) {
  const sheet = workbook.addWorksheet(name);
  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };
  for (const row of rows) sheet.addRow(row);
  return sheet;
}

async function exportToExcel(report1Data, report2Api, report2Sheets) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'tractiq-scraper';
  workbook.created = new Date();

  const msaLookup = new Map(report1Data.msaData.map((m) => [m.id, m]));
  const allIds = report1Data.msaData.map((m) => m.id); // includes NATIONAL

  addSheet(
    workbook,
    'Metadata',
    [
      { header: 'Field', key: 'field', width: 30 },
      { header: 'Value', key: 'value', width: 70 },
    ],
    [
      { field: 'Scraped at', value: new Date().toISOString() },
      { field: 'Report 1 URL', value: REPORT_1_URL },
      { field: 'Report 1 last updated', value: report1Data.lastUpdated },
      { field: 'Report 1 US sqft per capita (national constant)', value: report1Data.usSqftPerCapita },
      { field: 'Report 2 URL', value: REPORT_2_URL },
      { field: 'Report 2 dispersion (within-market spread) as of', value: report2Api.dispersion.lastUpdated },
      { field: 'Report 2 dispersion period (YYYYMM)', value: report2Api.dispersion.period },
      { field: 'Report 2 supply/pipeline as of', value: report2Api.supply.lastUpdated },
      { field: 'Report 2 pricing as of', value: report2Api.pricing.meta.lastUpdated },
      { field: 'Report 2 pricing period', value: report2Api.pricing.meta.period_label },
      { field: 'Report 2 pricing notes', value: report2Api.pricing.meta.notes },
    ]
  );

  // --- Report 1 sheets: every market (National + top 50 MSAs) ---
  addSheet(
    workbook,
    'R1 - MSA Reference',
    [
      { header: 'Rank', key: 'rank', width: 8 },
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    report1Data.msaData.map((m) => ({ ...m, as_of: report1Data.lastUpdated }))
  );

  addSheet(
    workbook,
    'R1 - Demographics',
    [
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Population', key: 'population', width: 14 },
      { header: 'Population 2010', key: 'population_2010', width: 16 },
      { header: 'Population 2020', key: 'population_2020', width: 16 },
      { header: 'Population Projected', key: 'population_projected', width: 18 },
      { header: 'Households', key: 'households', width: 14 },
      { header: 'Median HH Income', key: 'median_hh_income', width: 16 },
      { header: 'Population 25-44', key: 'pop_25_44', width: 16 },
      { header: 'Renters', key: 'renters', width: 12 },
      { header: 'Homeowners', key: 'homeowners', width: 12 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    allIds
      .filter((id) => demographicsHas(report1Data, id))
      .map((id) => ({
        id,
        name: msaLookup.get(id).name,
        state: msaLookup.get(id).state,
        ...report1Data.demographics[id],
        as_of: report1Data.lastUpdated,
      }))
  );

  addSheet(
    workbook,
    'R1 - Supply',
    [
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Facility Count', key: 'facility_count', width: 14 },
      { header: 'Total Rentable Sqft', key: 'total_rent_sqft', width: 18 },
      { header: 'Avg Facility Sqft', key: 'avg_rent_sqft', width: 16 },
      { header: 'REIT Facility Count', key: 'reit_count', width: 16 },
      { header: 'Climate-Controlled Unit Count', key: 'cc_unit_count', width: 20 },
      { header: 'Sqft Per Capita', key: 'sqft_per_capita', width: 16 },
      { header: 'US Sqft Per Capita', key: 'us_sqft_per_capita', width: 18 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    allIds
      .filter((id) => report1Data.supply[id])
      .map((id) => {
        const population = report1Data.demographics[id]?.population;
        const totalRentSqft = report1Data.supply[id].total_rent_sqft;
        return {
          id,
          name: msaLookup.get(id).name,
          state: msaLookup.get(id).state,
          ...report1Data.supply[id],
          sqft_per_capita: population ? Number((totalRentSqft / population).toFixed(2)) : null,
          us_sqft_per_capita: report1Data.usSqftPerCapita,
          as_of: report1Data.lastUpdated,
        };
      })
  );

  addSheet(
    workbook,
    'R1 - Pipeline',
    [
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Total Pipeline Facility Count', key: 'total_pipeline_count', width: 20 },
      { header: 'Total Pipeline Sqft', key: 'total_pipeline_sqft', width: 18 },
      { header: 'Known-Sqft Facility Count', key: 'known_sqft_count', width: 20 },
      { header: 'Avg Known Facility Sqft', key: 'avg_known_sqft', width: 18 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    allIds
      .filter((id) => report1Data.pipeline[id])
      .map((id) => ({
        id,
        name: msaLookup.get(id).name,
        state: msaLookup.get(id).state,
        ...report1Data.pipeline[id],
        as_of: report1Data.lastUpdated,
      }))
  );

  const pricingRows = [];
  for (const id of allIds) {
    const byYear = report1Data.pricing[id];
    if (!byYear) continue;
    for (const [year, byQuarter] of Object.entries(byYear)) {
      for (const [quarter, v] of Object.entries(byQuarter)) {
        pricingRows.push({
          id,
          name: msaLookup.get(id).name,
          state: msaLookup.get(id).state,
          year: Number(year),
          quarter: Number(quarter),
          street_rate_10x10: v.street,
          web_rate_10x10: v.web,
          n_facilities: v.n,
          as_of: report1Data.lastUpdated,
        });
      }
    }
  }
  addSheet(
    workbook,
    'R1 - Pricing (Quarterly)',
    [
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Year', key: 'year', width: 8 },
      { header: 'Quarter', key: 'quarter', width: 8 },
      { header: 'Street Rate 10x10 ($)', key: 'street_rate_10x10', width: 18 },
      { header: 'Web Rate 10x10 ($)', key: 'web_rate_10x10', width: 16 },
      { header: 'N Facilities', key: 'n_facilities', width: 12 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    pricingRows
  );

  const occupancyRows = [];
  for (const id of allIds) {
    const byYear = report1Data.occupancy[id];
    if (!byYear) continue;
    for (const [year, byQuarter] of Object.entries(byYear)) {
      for (const [quarter, v] of Object.entries(byQuarter)) {
        occupancyRows.push({
          id,
          name: msaLookup.get(id).name,
          state: msaLookup.get(id).state,
          year: Number(year),
          quarter: Number(quarter),
          occupancy_pct: v.occ,
          n_facilities: v.n,
          as_of: report1Data.lastUpdated,
        });
      }
    }
  }
  addSheet(
    workbook,
    'R1 - Occupancy (Quarterly)',
    [
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Year', key: 'year', width: 8 },
      { header: 'Quarter', key: 'quarter', width: 8 },
      { header: 'Occupancy %', key: 'occupancy_pct', width: 14 },
      { header: 'N Facilities', key: 'n_facilities', width: 12 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    occupancyRows
  );

  const financialsRows = [];
  for (const id of allIds) {
    const byYear = report1Data.financials[id];
    if (!byYear) continue;
    for (const [year, v] of Object.entries(byYear)) {
      financialsRows.push({
        id,
        name: msaLookup.get(id).name,
        state: msaLookup.get(id).state,
        year: Number(year),
        ...v,
        as_of: report1Data.lastUpdated,
      });
    }
  }
  addSheet(
    workbook,
    'R1 - Financials (Annual)',
    [
      { header: 'MSA ID', key: 'id', width: 12 },
      { header: 'Name', key: 'name', width: 45 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Year', key: 'year', width: 8 },
      { header: 'N Facilities', key: 'n', width: 12 },
      { header: 'Avg EGI ($)', key: 'avg_egi', width: 14 },
      { header: 'Avg NOI ($)', key: 'avg_noi', width: 14 },
      { header: 'Avg OpEx ($)', key: 'avg_opex', width: 14 },
      { header: 'NOI Margin (%)', key: 'noi_margin', width: 14 },
      { header: 'EGI $/sqft', key: 'egi_psf', width: 12 },
      { header: 'NOI $/sqft', key: 'noi_psf', width: 12 },
      { header: 'OpEx $/sqft', key: 'opex_psf', width: 12 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    financialsRows
  );

  // --- Report 2 sheets: top 50 MSAs (as ranked in report 1) ---
  addSheet(
    workbook,
    'R2 - Within-Market Spread',
    [
      { header: 'Rank', key: 'rank', width: 8 },
      { header: 'MSA ID', key: 'msa_id', width: 12 },
      { header: 'MSA', key: 'msa', width: 40 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'N Facilities', key: 'n_facilities', width: 12 },
      { header: 'P10 Occupancy', key: 'p10_occ', width: 14 },
      { header: 'P25 Occupancy', key: 'p25_occ', width: 14 },
      { header: 'Median (P50) Occupancy', key: 'p50_occ_median', width: 20 },
      { header: 'P75 Occupancy', key: 'p75_occ', width: 14 },
      { header: 'P90 Occupancy', key: 'p90_occ', width: 14 },
      { header: 'Avg Occupancy', key: 'avg_occ', width: 14 },
      { header: 'Within-Market Spread (P90-P10)', key: 'within_market_spread_p90_minus_p10', width: 26 },
      { header: 'Period (YYYYMM)', key: 'period', width: 14 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    report2Sheets.spreadRows
  );

  addSheet(
    workbook,
    'R2 - Supply Pressure',
    [
      { header: 'Rank', key: 'rank', width: 8 },
      { header: 'MSA ID', key: 'msa_id', width: 12 },
      { header: 'MSA', key: 'msa', width: 40 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Pipeline Sqft', key: 'pipeline_sqft', width: 16 },
      { header: 'Pipeline Facility Count', key: 'pipeline_facility_count', width: 20 },
      { header: 'Current Occupancy', key: 'current_occ', width: 16 },
      { header: 'Current Occ. Period (YYYYMM)', key: 'current_occ_period', width: 22 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    report2Sheets.supplyRows
  );

  addSheet(
    workbook,
    'R2 - Rates & Pricing',
    [
      { header: 'Rank', key: 'rank', width: 8 },
      { header: 'MSA ID', key: 'msa_id', width: 12 },
      { header: 'MSA', key: 'msa', width: 40 },
      { header: 'State', key: 'state', width: 16 },
      { header: 'Street Rate $/sqft', key: 'street_rate_per_sqft', width: 16 },
      { header: 'Web Rate $/sqft', key: 'web_rate_per_sqft', width: 16 },
      { header: 'Web Discount %', key: 'web_discount_pct', width: 14 },
      { header: 'N Facilities Priced', key: 'n_facilities_priced', width: 16 },
      { header: 'Current Occupancy', key: 'current_occ', width: 16 },
      { header: 'Unit Size', key: 'unit_size', width: 10 },
      { header: 'Period', key: 'period_label', width: 12 },
      { header: 'As Of', key: 'as_of', width: 16 },
    ],
    report2Sheets.pricingRows
  );

  ensureDir(OUT_DIR);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const outPath = path.join(OUT_DIR, `tractiq_data_${timestamp}.xlsx`);
  await workbook.xlsx.writeFile(outPath);
  console.log(`\nWrote workbook: ${outPath}`);
  return outPath;
}

function demographicsHas(report1Data, id) {
  return Boolean(report1Data.demographics[id]);
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

async function runInspect() {
  assertCredentials();
  ensureDir(INSPECT_DIR);
  const browser = await chromium.launch(resolveChromiumLaunchOptions());
  const page = await browser.newPage(contextOptionsForCurrentEnv());
  try {
    await login(page);
    console.log('\n=== Report 1 ===');
    const report1Data = await extractReport1Data(page);
    console.log('\n=== Report 2 ===');
    const report2Api = await extractReport2Data(page);
    const report2Sheets = buildReport2Top50Sheets(report1Data, report2Api);
    console.log(
      `\nTop-50 coverage — spread: ${report2Sheets.spreadRows.filter((r) => r.avg_occ != null).length}/50, ` +
        `supply: ${report2Sheets.supplyRows.filter((r) => r.pipeline_sqft != null).length}/50, ` +
        `pricing: ${report2Sheets.pricingRows.filter((r) => r.street_rate_per_sqft != null).length}/50`
    );
    fs.writeFileSync(path.join(INSPECT_DIR, 'report1-parsed.json'), JSON.stringify(report1Data, null, 2));
    fs.writeFileSync(path.join(INSPECT_DIR, 'report2-parsed.json'), JSON.stringify(report2Api, null, 2));
    console.log('\nLooks good. Run `node scrape.js run` to produce the Excel workbook.');
  } finally {
    await browser.close();
  }
}

async function runFull() {
  assertCredentials();
  ensureDir(OUT_DIR);
  const browser = await chromium.launch(resolveChromiumLaunchOptions());
  const page = await browser.newPage(contextOptionsForCurrentEnv());
  try {
    await login(page);
    const report1Data = await extractReport1Data(page);
    const report2Api = await extractReport2Data(page);
    const report2Sheets = buildReport2Top50Sheets(report1Data, report2Api);
    await exportToExcel(report1Data, report2Api, report2Sheets);
  } finally {
    await browser.close();
  }
}

if (require.main === module) {
  const mode = process.argv[2];
  if (mode === 'inspect') {
    runInspect().catch((err) => {
      console.error(err);
      process.exit(1);
    });
  } else if (mode === 'run') {
    runFull().catch((err) => {
      console.error(err);
      process.exit(1);
    });
  } else {
    console.log('Usage: node scrape.js <inspect|run>');
    process.exit(1);
  }
} else {
  module.exports = {
    login,
    getReportFrame,
    extractReport1Data,
    extractReport2Data,
    buildReport2Top50Sheets,
    resolveChromiumLaunchOptions,
    contextOptionsForCurrentEnv,
    REPORT_1_URL,
    REPORT_2_URL,
    BASE_URL,
  };
}
