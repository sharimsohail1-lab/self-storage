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
 * Step 1 — ALWAYS run inspect mode first. This site's exact DOM structure is
 * unknown to this script (it was written without ever being able to load the
 * page), so `run` mode uses best-effort/fuzzy selectors. Inspect mode logs
 * what it actually finds (screenshots, HTML, dropdown/table structure) to
 * output/inspect/ so you can confirm the guesses before trusting a full run,
 * and so the SELECTOR HINTS below can be corrected if needed.
 *
 *   node scrape.js inspect
 *
 * Step 2 — review output/inspect/*.png and *.html, and the console output.
 * If the console warns it couldn't find the MSA dropdown or a data table,
 * open the HTML dump, find the real selector, and add it to the relevant
 * *_SELECTORS array below (near the top of the list = tried first).
 *
 * Step 3 — run the full scrape:
 *
 *   node scrape.js run
 *
 * Output workbook: output/tractiq_data_<timestamp>.xlsx
 */

const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');

const BASE_URL = 'https://hub.tractiq.com';
const REPORT_1_URL = `${BASE_URL}/#report/1c055fbe-1836-4950-b70c-ecc712934a6c`;
const REPORT_2_URL = `${BASE_URL}/#report/7fb00aeb-5f9d-47a1-86e1-48644f34158d`;
const TOP_N_MSAS = 50;

const EMAIL = process.env.TRACTIQ_EMAIL;
const PASSWORD = process.env.TRACTIQ_PASSWORD;

const OUT_DIR = path.join(__dirname, 'output');
const INSPECT_DIR = path.join(OUT_DIR, 'inspect');

// ---------------------------------------------------------------------------
// SELECTOR HINTS — edit these if inspect mode shows the real DOM differs.
// Order matters: first match wins.
// ---------------------------------------------------------------------------
const LOGIN_EMAIL_SELECTORS = [
  'input[type="email"]',
  'input[name="email"]',
  'input[name="username"]',
  'input#email',
  'input#username',
  'input[autocomplete="username"]',
];
const LOGIN_PASSWORD_SELECTORS = [
  'input[type="password"]',
  'input[name="password"]',
  'input#password',
  'input[autocomplete="current-password"]',
];
const LOGIN_SUBMIT_SELECTORS = [
  'button[type="submit"]',
  'button:has-text("Log in")',
  'button:has-text("Sign in")',
  'button:has-text("Login")',
  'input[type="submit"]',
];
const LOADING_SPINNER_SELECTORS = [
  '.spinner',
  '[class*="loading" i]',
  '[class*="spinner" i]',
  '[aria-busy="true"]',
];
const MSA_DROPDOWN_HINT_SELECTORS = [
  'select[name*="msa" i]',
  'select[id*="msa" i]',
  'select[name*="market" i]',
  '[role="combobox"][aria-label*="msa" i]',
  '[role="combobox"][aria-label*="market" i]',
];

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function assertCredentials() {
  if (!EMAIL || !PASSWORD) {
    console.error(
      'Missing credentials. Set TRACTIQ_EMAIL and TRACTIQ_PASSWORD environment variables before running.'
    );
    process.exit(1);
  }
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

async function waitForDataLoad(page, { extraDelayMs = 800, timeoutMs = 20000 } = {}) {
  try {
    await page.waitForLoadState('networkidle', { timeout: timeoutMs });
  } catch (_) {
    // SPA may keep a long-lived connection open; don't hard-fail on this.
  }
  for (const sel of LOADING_SPINNER_SELECTORS) {
    try {
      await page.locator(sel).first().waitFor({ state: 'hidden', timeout: 3000 });
    } catch (_) {
      // spinner selector may not exist on this page — fine
    }
  }
  await sleep(extraDelayMs);
}

async function login(page) {
  console.log(`Logging in as ${EMAIL} ...`);
  await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });
  await waitForDataLoad(page, { extraDelayMs: 500 });

  const emailField = await firstVisible(page, LOGIN_EMAIL_SELECTORS, 8000);
  if (!emailField) {
    await debugDump(page, 'login-page-no-email-field');
    throw new Error(
      'Could not find an email/username field on the login page. ' +
        'See output/inspect/login-page-no-email-field.html and .png, then update LOGIN_EMAIL_SELECTORS in scrape.js.'
    );
  }
  await emailField.locator.fill(EMAIL);

  const passwordField = await firstVisible(page, LOGIN_PASSWORD_SELECTORS, 4000);
  if (!passwordField) {
    await debugDump(page, 'login-page-no-password-field');
    throw new Error(
      'Could not find a password field on the login page. ' +
        'See output/inspect/login-page-no-password-field.html and .png, then update LOGIN_PASSWORD_SELECTORS in scrape.js.'
    );
  }
  await passwordField.locator.fill(PASSWORD);

  const submitButton = await firstVisible(page, LOGIN_SUBMIT_SELECTORS, 4000);
  if (!submitButton) {
    await debugDump(page, 'login-page-no-submit-button');
    throw new Error(
      'Could not find a submit/login button. ' +
        'See output/inspect/login-page-no-submit-button.html and .png, then update LOGIN_SUBMIT_SELECTORS in scrape.js.'
    );
  }

  await Promise.all([
    page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {}),
    submitButton.locator.click(),
  ]);
  await sleep(1500);

  const stillOnPasswordField = await page
    .locator(LOGIN_PASSWORD_SELECTORS.join(','))
    .first()
    .isVisible()
    .catch(() => false);
  if (stillOnPasswordField) {
    await debugDump(page, 'login-may-have-failed');
    console.warn(
      'WARNING: a password field is still visible after submitting — login may have failed ' +
        '(wrong credentials, or a 2FA/CAPTCHA step this script does not handle). ' +
        'Check output/inspect/login-may-have-failed.png.'
    );
  } else {
    console.log('Login appears to have succeeded.');
  }
}

async function debugDump(page, label) {
  ensureDir(INSPECT_DIR);
  const safeLabel = label.replace(/[^a-z0-9_-]/gi, '_');
  try {
    await page.screenshot({ path: path.join(INSPECT_DIR, `${safeLabel}.png`), fullPage: true });
  } catch (_) {}
  try {
    const html = await page.content();
    fs.writeFileSync(path.join(INSPECT_DIR, `${safeLabel}.html`), html);
  } catch (_) {}
}

// ---------------------------------------------------------------------------
// Dropdown discovery + enumeration
// ---------------------------------------------------------------------------

async function findMsaDropdown(page) {
  // Strategy 1: an explicit hinted selector.
  for (const selector of MSA_DROPDOWN_HINT_SELECTORS) {
    const count = await page.locator(selector).count();
    if (count > 0) {
      return { kind: 'select-or-combobox', selector };
    }
  }

  // Strategy 2: any <select> with a plausible number of options (MSAs = tens to hundreds).
  const selects = await page.locator('select').all();
  for (const select of selects) {
    const optionCount = await select.locator('option').count();
    if (optionCount >= 10) {
      const selector = await select.evaluate((el) => {
        if (el.id) return `#${el.id}`;
        if (el.name) return `select[name="${el.name}"]`;
        return null;
      });
      return { kind: 'select', selector, locator: select };
    }
  }

  // Strategy 3: a combobox / listbox trigger.
  const combo = page.locator('[role="combobox"], [role="listbox"], [class*="dropdown" i], [class*="select" i]').first();
  if ((await combo.count()) > 0) {
    return { kind: 'combobox', locator: combo };
  }

  return null;
}

async function enumerateMsaOptions(page, dropdown) {
  if (dropdown.kind === 'select' || (dropdown.kind === 'select-or-combobox' && dropdown.selector)) {
    const selectLocator = dropdown.locator || page.locator(dropdown.selector).first();
    const options = await selectLocator.locator('option').evaluateAll((opts) =>
      opts
        .map((o) => ({ label: o.textContent.trim(), value: o.value }))
        .filter((o) => o.label && o.label.toLowerCase() !== 'select...' && o.value !== '')
    );
    return { type: 'select', locator: selectLocator, options };
  }

  // Combobox: click to open, read rendered option items, then close.
  const trigger = dropdown.locator;
  await trigger.click();
  await sleep(400);
  const optionLocator = page.locator('[role="option"], li[class*="option" i]');
  const optionCount = await optionLocator.count();
  const options = [];
  for (let i = 0; i < optionCount; i++) {
    const text = (await optionLocator.nth(i).innerText()).trim();
    if (text) options.push({ label: text, value: text });
  }
  await page.keyboard.press('Escape').catch(() => {});
  return { type: 'combobox', trigger, optionLocator, options };
}

async function selectMsaOption(page, enumerated, option) {
  if (enumerated.type === 'select') {
    await enumerated.locator.selectOption({ label: option.label }).catch(async () => {
      await enumerated.locator.selectOption(option.value);
    });
    await enumerated.locator.dispatchEvent('change').catch(() => {});
  } else {
    await enumerated.trigger.click();
    await sleep(300);
    const target = page.locator('[role="option"], li[class*="option" i]', { hasText: option.label }).first();
    await target.click();
  }
  await waitForDataLoad(page);
}

// ---------------------------------------------------------------------------
// Data extraction
// ---------------------------------------------------------------------------

async function extractAsOfDate(page) {
  const text = await page.locator('body').innerText();
  const match = text.match(/as[\s-]*of[:\s]*([A-Za-z0-9,/\-\s]{4,30})/i);
  return match ? match[1].trim().split('\n')[0] : null;
}

async function extractAllTables(page) {
  const tableLocators = await page.locator('table').all();
  const tables = [];
  for (const table of tableLocators) {
    const data = await table.evaluate((t) =>
      Array.from(t.querySelectorAll('tr')).map((r) =>
        Array.from(r.querySelectorAll('th,td')).map((c) => c.textContent.trim())
      )
    );
    if (data.length > 0) tables.push(data);
  }
  return tables;
}

function pickMainTable(tables) {
  if (tables.length === 0) return null;
  return tables.reduce((best, t) => {
    const size = t.length * (t[0] ? t[0].length : 0);
    const bestSize = best.length * (best[0] ? best[0].length : 0);
    return size > bestSize ? t : best;
  });
}

async function extractMetricByLabel(page, labelPatterns) {
  const text = await page.locator('body').innerText();
  for (const pattern of labelPatterns) {
    const re = new RegExp(pattern.source + '\\s*:?\\s*(\\$?-?[\\d.,]+%?)', 'i');
    const m = text.match(re);
    if (m) return m[1].trim();
  }
  return null;
}

// ---------------------------------------------------------------------------
// Inspect mode — figure out the real DOM before trusting the scrape logic.
// ---------------------------------------------------------------------------

async function inspectReport(page, url, label) {
  console.log(`\n=== Inspecting ${label}: ${url} ===`);
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await waitForDataLoad(page, { extraDelayMs: 1500 });
  await debugDump(page, `${label}-landing`);

  const dropdown = await findMsaDropdown(page);
  if (!dropdown) {
    console.warn(`  No dropdown found on ${label}. Check output/inspect/${label}-landing.html manually.`);
  } else {
    console.log(`  Dropdown candidate: ${JSON.stringify({ kind: dropdown.kind, selector: dropdown.selector })}`);
    const enumerated = await enumerateMsaOptions(page, dropdown);
    console.log(`  Found ${enumerated.options.length} options. First 5:`, enumerated.options.slice(0, 5));
    fs.writeFileSync(
      path.join(INSPECT_DIR, `${label}-dropdown-options.json`),
      JSON.stringify(enumerated.options, null, 2)
    );
  }

  const tables = await extractAllTables(page);
  console.log(`  Found ${tables.length} <table> element(s) on the page.`);
  tables.forEach((t, i) => {
    console.log(`    table[${i}]: ${t.length} rows x ${t[0] ? t[0].length : 0} cols. Header:`, t[0]);
  });
  fs.writeFileSync(path.join(INSPECT_DIR, `${label}-tables.json`), JSON.stringify(tables, null, 2));

  const asOf = await extractAsOfDate(page);
  console.log(`  "As of" text detected: ${asOf || '(none found — check page manually)'}`);
}

async function runInspect() {
  assertCredentials();
  ensureDir(INSPECT_DIR);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  try {
    await login(page);
    await debugDump(page, 'post-login-landing');
    await inspectReport(page, REPORT_1_URL, 'report1');
    await inspectReport(page, REPORT_2_URL, 'report2');
    console.log(
      `\nDone. Review the files in ${INSPECT_DIR}. If dropdown/table detection above looks wrong, ` +
        'update the SELECTOR HINTS near the top of scrape.js before running `node scrape.js run`.'
    );
  } finally {
    await browser.close();
  }
}

// ---------------------------------------------------------------------------
// Report 1: enumerate every MSA, capture its main table + as-of date
// ---------------------------------------------------------------------------

async function scrapeReport1(page) {
  console.log(`\n=== Scraping report 1: ${REPORT_1_URL} ===`);
  await page.goto(REPORT_1_URL, { waitUntil: 'domcontentloaded' });
  await waitForDataLoad(page, { extraDelayMs: 1500 });

  const dropdown = await findMsaDropdown(page);
  if (!dropdown) {
    await debugDump(page, 'report1-no-dropdown');
    throw new Error(
      'Could not locate the MSA dropdown on report 1. Run `node scrape.js inspect` and update MSA_DROPDOWN_HINT_SELECTORS.'
    );
  }
  const enumerated = await enumerateMsaOptions(page, dropdown);
  if (enumerated.options.length === 0) {
    throw new Error('MSA dropdown found but it has no options — check output/inspect/report1-landing.html.');
  }
  console.log(`Found ${enumerated.options.length} MSAs in report 1's dropdown.`);

  const rows = [];
  for (let i = 0; i < enumerated.options.length; i++) {
    const option = enumerated.options[i];
    process.stdout.write(`  [${i + 1}/${enumerated.options.length}] ${option.label} ... `);
    try {
      await selectMsaOption(page, enumerated, option);
      const asOfDate = await extractAsOfDate(page);
      const tables = await extractAllTables(page);
      const mainTable = pickMainTable(tables);
      rows.push({ msa: option.label, asOfDate, table: mainTable, allTables: tables });
      console.log(`ok (as of: ${asOfDate || 'n/a'}, ${tables.length} table(s))`);
    } catch (err) {
      console.log(`FAILED: ${err.message}`);
      await debugDump(page, `report1-failed-${option.label}`);
      rows.push({ msa: option.label, asOfDate: null, table: null, allTables: [], error: err.message });
    }
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Report 2: within-market spread, supply pressure, pipeline for top 50 MSAs
// ---------------------------------------------------------------------------

async function scrapeReport2(page) {
  console.log(`\n=== Scraping report 2: ${REPORT_2_URL} ===`);
  await page.goto(REPORT_2_URL, { waitUntil: 'domcontentloaded' });
  await waitForDataLoad(page, { extraDelayMs: 1500 });

  // Attempt 1: a single table already listing all MSAs with the three metrics as columns.
  const tables = await extractAllTables(page);
  const candidate = tables.find(
    (t) => t[0] && t[0].some((h) => /spread/i.test(h)) && t[0].some((h) => /supply/i.test(h))
  );

  if (candidate) {
    console.log('Found a single ranked table with spread/supply/pipeline columns — using it directly.');
    const header = candidate[0];
    const idx = {
      msa: header.findIndex((h) => /msa|market/i.test(h)),
      spread: header.findIndex((h) => /spread/i.test(h)),
      supply: header.findIndex((h) => /supply/i.test(h)),
      pipeline: header.findIndex((h) => /pipeline/i.test(h)),
    };
    const rows = candidate
      .slice(1)
      .map((r) => ({
        msa: idx.msa >= 0 ? r[idx.msa] : null,
        spread: idx.spread >= 0 ? r[idx.spread] : null,
        supply: idx.supply >= 0 ? r[idx.supply] : null,
        pipeline: idx.pipeline >= 0 ? r[idx.pipeline] : null,
      }))
      .filter((r) => r.msa);
    return { rows: rows.slice(0, TOP_N_MSAS), rawTables: tables, mode: 'single-table' };
  }

  // Attempt 2: per-MSA drilldown, same dropdown pattern as report 1, using the
  // first TOP_N_MSAS entries in the dropdown as the "top 50" (as ranked by the tool).
  console.log('No single ranked table found — falling back to per-MSA drilldown (using dropdown order as rank).');
  const dropdown = await findMsaDropdown(page);
  if (!dropdown) {
    await debugDump(page, 'report2-no-dropdown-no-table');
    throw new Error(
      'Report 2 has neither a recognizable ranked table nor an MSA dropdown. ' +
        'Run `node scrape.js inspect` and inspect output/inspect/report2-landing.html manually.'
    );
  }
  const enumerated = await enumerateMsaOptions(page, dropdown);
  const topOptions = enumerated.options.slice(0, TOP_N_MSAS);
  console.log(`Using top ${topOptions.length} MSAs from dropdown order.`);

  const rows = [];
  const rawTables = [];
  for (let i = 0; i < topOptions.length; i++) {
    const option = topOptions[i];
    process.stdout.write(`  [${i + 1}/${topOptions.length}] ${option.label} ... `);
    try {
      await selectMsaOption(page, enumerated, option);
      const spread = await extractMetricByLabel(page, [/within.?market spread/i, /\bspread\b/i]);
      const supply = await extractMetricByLabel(page, [/supply pressure/i]);
      const pipeline = await extractMetricByLabel(page, [/pipeline/i]);
      const pageTables = await extractAllTables(page);
      rawTables.push({ msa: option.label, tables: pageTables });
      rows.push({ msa: option.label, spread, supply, pipeline });
      console.log(`ok (spread=${spread ?? 'n/a'}, supply=${supply ?? 'n/a'}, pipeline=${pipeline ?? 'n/a'})`);
    } catch (err) {
      console.log(`FAILED: ${err.message}`);
      await debugDump(page, `report2-failed-${option.label}`);
      rows.push({ msa: option.label, spread: null, supply: null, pipeline: null, error: err.message });
    }
  }
  return { rows, rawTables, mode: 'per-msa-drilldown' };
}

// ---------------------------------------------------------------------------
// Excel export
// ---------------------------------------------------------------------------

async function exportToExcel(report1Rows, report2Result) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'tractiq-scraper';
  workbook.created = new Date();

  // --- Metadata ---
  const meta = workbook.addWorksheet('Metadata');
  meta.columns = [{ header: 'Field', key: 'field', width: 24 }, { header: 'Value', key: 'value', width: 70 }];
  meta.addRows([
    { field: 'Scraped at', value: new Date().toISOString() },
    { field: 'Report 1 URL', value: REPORT_1_URL },
    { field: 'Report 2 URL', value: REPORT_2_URL },
    { field: 'Report 2 extraction mode', value: report2Result.mode },
    { field: 'Top N MSAs (report 2)', value: TOP_N_MSAS },
  ]);
  meta.getRow(1).font = { bold: true };

  // --- Report 1: structured sheet (assumes consistent table shape across MSAs) ---
  const r1Sheet = workbook.addWorksheet('Report1 - All MSAs');
  const firstGoodTable = report1Rows.find((r) => r.table && r.table.length > 1);
  const headerRow = firstGoodTable ? firstGoodTable.table[0] : ['Value'];
  r1Sheet.columns = [
    { header: 'MSA', key: 'msa', width: 28 },
    { header: 'As Of Date', key: 'asOfDate', width: 16 },
    ...headerRow.map((h, i) => ({ header: h || `Col${i + 1}`, key: `c${i}`, width: 18 })),
  ];
  r1Sheet.getRow(1).font = { bold: true };
  for (const r of report1Rows) {
    if (!r.table || r.table.length < 2) {
      r1Sheet.addRow({ msa: r.msa, asOfDate: r.asOfDate, c0: r.error ? `ERROR: ${r.error}` : '(no table found)' });
      continue;
    }
    for (const dataRow of r.table.slice(1)) {
      const rowObj = { msa: r.msa, asOfDate: r.asOfDate };
      dataRow.forEach((cell, i) => {
        rowObj[`c${i}`] = cell;
      });
      r1Sheet.addRow(rowObj);
    }
  }

  // --- Report 1: raw safety-net sheet (every table, every row, untouched) ---
  const r1Raw = workbook.addWorksheet('Report1 - Raw');
  r1Raw.columns = [
    { header: 'MSA', key: 'msa', width: 28 },
    { header: 'As Of Date', key: 'asOfDate', width: 16 },
    { header: 'Table #', key: 'tableIdx', width: 10 },
    { header: 'Row #', key: 'rowIdx', width: 8 },
    { header: 'Row Values (tab-joined)', key: 'rowValues', width: 100 },
  ];
  r1Raw.getRow(1).font = { bold: true };
  for (const r of report1Rows) {
    (r.allTables || []).forEach((table, tIdx) => {
      table.forEach((rowValues, rIdx) => {
        r1Raw.addRow({
          msa: r.msa,
          asOfDate: r.asOfDate,
          tableIdx: tIdx,
          rowIdx: rIdx,
          rowValues: rowValues.join(' | '),
        });
      });
    });
  }

  // --- Report 2: top 50 spread/supply/pipeline ---
  const r2Sheet = workbook.addWorksheet('Report2 - Top50');
  r2Sheet.columns = [
    { header: 'MSA', key: 'msa', width: 28 },
    { header: 'Within-Market Spread', key: 'spread', width: 22 },
    { header: 'Supply Pressure', key: 'supply', width: 20 },
    { header: 'Pipeline', key: 'pipeline', width: 18 },
  ];
  r2Sheet.getRow(1).font = { bold: true };
  for (const r of report2Result.rows) {
    r2Sheet.addRow(r);
  }

  // --- Report 2: raw safety-net sheet (only populated in per-MSA drilldown mode) ---
  if (report2Result.rawTables && report2Result.rawTables.length > 0 && report2Result.rawTables[0].msa) {
    const r2Raw = workbook.addWorksheet('Report2 - Raw');
    r2Raw.columns = [
      { header: 'MSA', key: 'msa', width: 28 },
      { header: 'Table #', key: 'tableIdx', width: 10 },
      { header: 'Row #', key: 'rowIdx', width: 8 },
      { header: 'Row Values (tab-joined)', key: 'rowValues', width: 100 },
    ];
    r2Raw.getRow(1).font = { bold: true };
    for (const entry of report2Result.rawTables) {
      entry.tables.forEach((table, tIdx) => {
        table.forEach((rowValues, rIdx) => {
          r2Raw.addRow({ msa: entry.msa, tableIdx: tIdx, rowIdx: rIdx, rowValues: rowValues.join(' | ') });
        });
      });
    }
  }

  ensureDir(OUT_DIR);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const outPath = path.join(OUT_DIR, `tractiq_data_${timestamp}.xlsx`);
  await workbook.xlsx.writeFile(outPath);
  console.log(`\nWrote workbook: ${outPath}`);
  return outPath;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

async function runFull() {
  assertCredentials();
  ensureDir(OUT_DIR);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  try {
    await login(page);
    const report1Rows = await scrapeReport1(page);
    const report2Result = await scrapeReport2(page);
    await exportToExcel(report1Rows, report2Result);
  } finally {
    await browser.close();
  }
}

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
