import { test, expect, Page } from '@playwright/test';
import * as ExcelJS from 'exceljs';
import * as path from 'path';

/**
 * Hr Time Sheet Format — All WOs
 * 1. Read all WOIDs + Worker Names from the "WO" sheet in FieldGlass.xlsx
 * 2. Login to SAP Fieldglass once
 * 3. For each WOID: search → Time & Expense tab → set 01/07/2026–31/07/2026 → Apply Filters → extract rows
 * 4. Write all rows to "01-Jul-2026_31-Jul-2026" sheet
 */

const XLSX_PATH = path.resolve(__dirname, '../FieldGlass.xlsx');
const SHEET_NAME = '01-Jul-2026_31-Jul-2026';
const FROM_DATE  = '01/07/2026';
const TO_DATE    = '31/07/2026';

// Output column headers (matches existing sheet format)
const OUTPUT_HEADERS = [
  'WOID', 'Worker Name', 'Timesheet ID', 'Status',
  'Start Date', 'End Date', 'Approved Date',
  'ST Hours', 'OT Hours', 'DT Hours', 'Others Hours', 'NB Hours',
  'Amount (INR)',
];

// ─── Excel helpers ────────────────────────────────────────────────────────────

/** Read WOID (col C) and Worker Name (col D) from the WO sheet, skipping header row */
async function readWOSheet(): Promise<Array<{ woid: string; name: string }>> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  const ws = wb.getWorksheet('WO');
  if (!ws) throw new Error('WO sheet not found in FieldGlass.xlsx');
  const items: Array<{ woid: string; name: string }> = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum === 1) return; // skip header
    const woid = String(row.getCell(3).value ?? '').trim(); // column C
    const name = String(row.getCell(4).value ?? '').trim(); // column D
    if (woid) items.push({ woid, name });
  });
  console.log(`Read ${items.length} WOIDs from WO sheet`);
  return items;
}

/** Write all collected rows to the output sheet (clears it first) */
async function writeOutputSheet(allRows: string[][]) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);

  let ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) {
    ws = wb.addWorksheet(SHEET_NAME);
    console.log(`Created sheet "${SHEET_NAME}"`);
  } else {
    ws.spliceRows(1, ws.rowCount);
    console.log(`Cleared sheet "${SHEET_NAME}"`);
  }

  // Header row — blue background, white bold text
  const headerRow = ws.addRow(OUTPUT_HEADERS);
  headerRow.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
  headerRow.font   = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  // Data rows
  for (const row of allRows) {
    ws.addRow(row);
  }

  // Auto-fit column widths
  ws.columns.forEach(col => {
    let maxLen = 10;
    col.eachCell?.({ includeEmpty: false }, cell => {
      const v = cell.value?.toString() ?? '';
      if (v.length > maxLen) maxLen = v.length;
    });
    col.width = maxLen + 2;
  });

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`\nSaved ${allRows.length} data rows to "${SHEET_NAME}" in FieldGlass.xlsx`);
}

/**
 * Write raw timesheet rows to the "fields" sheet.
 * The sheet already has the header row: Status | ID | Start | End | Approved | ST | OT | DT | Others | NB | Amount (INR)
 * We keep the header and replace all data rows below it.
 * Each input row = [Status, ID, Start, End, Approved, ST, OT, DT, Others, NB, Amount]
 */
async function writeFieldsSheet(rows: string[][]) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);

  const ws = wb.getWorksheet('fields');
  if (!ws) throw new Error('"fields" sheet not found in FieldGlass.xlsx');

  // Keep row 1 (header), delete everything after it
  const totalRows = ws.rowCount;
  if (totalRows > 1) ws.spliceRows(2, totalRows - 1);

  // Append data rows
  for (const row of rows) {
    ws.addRow(row);
  }

  // Auto-fit column widths
  ws.columns.forEach(col => {
    let maxLen = 10;
    col.eachCell?.({ includeEmpty: false }, cell => {
      const v = cell.value?.toString() ?? '';
      if (v.length > maxLen) maxLen = v.length;
    });
    col.width = maxLen + 2;
  });

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`\nSaved ${rows.length} rows to "fields" sheet in FieldGlass.xlsx`);
}

// ─── Page helpers ─────────────────────────────────────────────────────────────

async function waitForSettle(page: Page, ms = 3000) {
  await page.waitForLoadState('load').catch(() => {});
  await page.waitForTimeout(ms);
}

async function acceptCookies(page: Page) {
  const selectors = [
    'button:has-text("Accept All")',
    'button:has-text("Accept all")',
    'button:has-text("Accept Cookies")',
    '#onetrust-accept-btn-handler',
  ];
  for (const sel of selectors) {
    try {
      const btn = page.locator(sel).first();
      if (await btn.isVisible({ timeout: 3000 })) {
        await btn.click();
        console.log(`Cookies accepted via: ${sel}`);
        await page.waitForTimeout(1500);
        return;
      }
    } catch { /* skip */ }
  }
}

async function dismissPanel(page: Page) {
  const selectors = [
    'button[aria-label="Close"]', 'button[title="Close"]',
    'button:has-text("×")', 'button:has-text("✕")',
    'button.close', 'button[class*="closeBtn"]',
  ];
  for (const sel of selectors) {
    try {
      const btn = page.locator(sel).first();
      if (await btn.isVisible({ timeout: 1500 })) {
        await btn.click();
        await page.waitForTimeout(600);
        return;
      }
    } catch { /* skip */ }
  }
  await page.keyboard.press('Escape').catch(() => {});
  await page.waitForTimeout(400);
}

async function getActivePage(
  context: import('@playwright/test').BrowserContext,
  fallback: Page
): Promise<Page> {
  await fallback.waitForTimeout(3000).catch(() => {});
  const pages = context.pages();
  for (const p of pages) {
    try {
      const url = p.url();
      if (!url.includes('login.do') && url !== 'about:blank') {
        await p.waitForLoadState('domcontentloaded').catch(() => {});
        return p;
      }
    } catch { /* closed */ }
  }
  for (const p of pages) {
    try { p.url(); return p; } catch { /* closed */ }
  }
  return fallback;
}

// ─── Text parser ──────────────────────────────────────────────────────────────

/**
 * Parse the Time Sheets section from the page's body innerText.
 * Returns raw data rows (11 cells each: Status, ID, Start, End, Approved, ST, OT, DT, Others, NB, Amount).
 */
function parseTimeSheetsFromText(rawText: string): string[][] {
  const STATUS_WORDS = ['Invoiced', 'Draft', 'Pending', 'Rejected', 'Submitted', 'Recalled'];
  const COL_COUNT = 11; // Status + 10 more

  // Step 1: merge "DD/MM/YYYY" + "HH:MM AM/PM" pairs that the page splits across 2 lines
  const rawLines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const lines: string[] = [];
  for (let i = 0; i < rawLines.length; i++) {
    const cur = rawLines[i];
    const nxt = rawLines[i + 1] ?? '';
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(cur) && /^\d{2}:\d{2}\s+(AM|PM)$/.test(nxt)) {
      lines.push(`${cur} ${nxt}`);
      i++;
    } else {
      lines.push(cur);
    }
  }

  // Step 2: find "Time Sheets" section
  const tsIdx = lines.findIndex(l => l === 'Time Sheets');
  if (tsIdx === -1) return [];

  // Step 3: find the "Status" column header right after the section heading
  let headerIdx = -1;
  for (let i = tsIdx; i < Math.min(tsIdx + 25, lines.length); i++) {
    if (lines[i] === 'Status') { headerIdx = i; break; }
  }
  if (headerIdx === -1) return [];

  // Step 4: skip the 11 column header lines, then skip any "All" filter pill
  let dataStart = headerIdx + COL_COUNT;
  if (lines[dataStart] === 'All') dataStart++;

  // Step 5: collect rows — each starts with a status word
  const rows: string[][] = [];
  let i = dataStart;
  while (i < lines.length) {
    const line = lines[i];
    if (
      line.startsWith('Clear Sort') || line.startsWith('No items') ||
      line.startsWith('Press enter') || /^\d+-\d+ of \d+/.test(line) ||
      line === 'Absences' || line === 'Expense Sheets' || line === 'Credit/Debit Memo'
    ) break;

    if (STATUS_WORDS.includes(line)) {
      const row: string[] = [line];
      for (let j = 1; j < COL_COUNT; j++) row.push(lines[i + j] ?? '');
      rows.push(row);
      i += COL_COUNT;
    } else {
      i++;
    }
  }
  return rows;
}

// ─── Per-WO fetch helper ──────────────────────────────────────────────────────

/**
 * For a given page (already on the SAP dashboard), search for woid,
 * navigate to Time & Expense, apply the date filter, and return extracted rows.
 * Returns [] if no timesheets found for this WO.
 */
async function fetchTimesheetsForWO(ap: Page, woid: string): Promise<string[][]> {
  console.log(`\n--- Fetching: ${woid} ---`);

  // Search
  const searchField = ap.locator(
    'input[placeholder*="Search by ID"], input[placeholder*="Search"], input[type="search"]'
  ).first();
  await searchField.waitFor({ state: 'visible', timeout: 20000 });
  await searchField.fill(woid);
  await searchField.press('Enter');
  await waitForSettle(ap, 3000);

  // Click Time & Expense tab
  const teTab = ap.locator('a[href*="tabId=timeAndExpense"], a:has-text("Time & Expense")').first();
  try {
    await teTab.waitFor({ state: 'visible', timeout: 15000 });
  } catch {
    console.log(`  No Time & Expense tab found for ${woid} — skipping`);
    return [];
  }
  await teTab.click();
  await waitForSettle(ap, 3000);

  // Fill date range — from
  const dateInputs = ap.locator('input[aria-label="Open Calendar (DD/MM/YYYY)"]');
  try {
    await dateInputs.first().waitFor({ state: 'visible', timeout: 12000 });
  } catch {
    console.log(`  No date inputs found for ${woid} — skipping`);
    return [];
  }

  await dateInputs.nth(0).click({ clickCount: 3 });
  await dateInputs.nth(0).fill(FROM_DATE);
  await ap.keyboard.press('Tab');
  await ap.waitForTimeout(600);

  // Fill date range — to
  await dateInputs.nth(1).click({ clickCount: 3 });
  await dateInputs.nth(1).fill(TO_DATE);
  await ap.keyboard.press('Tab'); // Tab moves focus to Apply Filters button
  await ap.waitForTimeout(600);

  // Apply Filters — button has focus after Tab, press Enter to activate
  await ap.keyboard.press('Enter');
  await waitForSettle(ap, 3000);

  // Extract data
  const rawText = await ap.evaluate(() => document.body.innerText);
  const rows = parseTimeSheetsFromText(rawText);
  console.log(`  Found ${rows.length} timesheet rows`);
  return rows;
}

// ─── Main test ────────────────────────────────────────────────────────────────

test('Hr Time Sheet Format — All WOs', async ({ page, context }) => {

  // ── Read WOIDs from WO sheet ─────────────────────────────────────────────────
  // To run for ALL WOIDs: remove the .slice(0, 1) below
  const woList = (await readWOSheet()).slice(0, 1);
  // const woList = await readWOSheet(); // ← uncomment this (and comment line above) to run all
  expect(woList.length).toBeGreaterThan(0);
  console.log(`Processing ${woList.length} WOID(s):`, woList.map(w => w.woid).join(', '));

  // ── Login ────────────────────────────────────────────────────────────────────
  await page.goto('https://www.us.fieldglass.cloud.sap/login.do');
  await page.waitForLoadState('domcontentloaded');
  await expect(page).toHaveURL(/login\.do/);
  console.log('On login page');

  await acceptCookies(page);
  await dismissPanel(page);

  // Username
  const usernameField = page.locator(
    'input[name="username"], input[id="username"], input[type="text"]'
  ).first();
  await usernameField.waitFor({ state: 'visible', timeout: 15000 });
  await usernameField.fill('Chethan K');
  await expect(usernameField).toHaveValue('Chethan K');

  // Password
  const passwordField = page.locator('input[type="password"]').first();
  await passwordField.waitFor({ state: 'visible', timeout: 10000 });
  await passwordField.fill('AmmuCA@@2002');

  // Sign In
  const newTabPromise = context.waitForEvent('page', { timeout: 30000 }).catch(() => null);
  const signInBtn = page.locator(
    'button:has-text("Sign In"), button[type="submit"], input[type="submit"]'
  ).first();
  await signInBtn.waitFor({ state: 'visible', timeout: 10000 });
  await signInBtn.click();
  console.log('Sign In clicked');

  // Resolve active page (SAP may open a new tab and close original)
  const newTab = await newTabPromise;
  let ap: Page;
  if (newTab) {
    await newTab.waitForLoadState('domcontentloaded').catch(() => {});
    ap = newTab;
    console.log('Switched to new tab');
  } else {
    ap = await getActivePage(context, page);
  }

  // Wait until URL is off login.do (handles login.do# SSO phase)
  for (let i = 0; i < 30; i++) {
    await ap.waitForTimeout(2000).catch(async () => { ap = await getActivePage(context, page); });
    let url = '';
    try { url = ap.url(); } catch { ap = await getActivePage(context, page); url = ap.url(); }
    console.log(`  [${i + 1}] ${url}`);
    if (url && !url.includes('login.do')) break;
    if (i === 29) throw new Error('Login failed — still on login.do after 60s');
  }

  await waitForSettle(ap, 2000);
  console.log('Login OK:', ap.url());
  await dismissPanel(ap);

  // ── Loop through all WOs ─────────────────────────────────────────────────────
  const allOutputRows: string[][] = [];

  for (const { woid, name } of woList) {
    let rows: string[][] = [];
    try {
      rows = await fetchTimesheetsForWO(ap, woid);
    } catch (err) {
      console.log(`  ERROR fetching ${woid}: ${err}`);
    }

    if (rows.length === 0) {
      // Record a single "no data" row so we know it was checked
      allOutputRows.push([woid, name, '', 'No Data', '', '', '', '', '', '', '', '', '']);
    } else {
      // Prepend WOID and Worker Name to each timesheet row
      // Raw row = [Status, ID, Start, End, Approved, ST, OT, DT, Others, NB, Amount]
      for (const row of rows) {
        const [status, tsId, start, end, approved, st, ot, dt, others, nb, amount] = row;
        allOutputRows.push([
          woid, name, tsId, status,
          start, end, approved,
          st, ot, dt, others, nb, amount,
        ]);
      }
    }

    console.log(`  Written ${rows.length} rows for ${woid}`);
  }

  // ── Write all results to Excel ────────────────────────────────────────────────
  // Build raw rows for the fields sheet: Status, ID, Start, End, Approved, ST, OT, DT, Others, NB, Amount
  // (skip WOs that had no data)
  const fieldsRows: string[][] = [];
  for (const outputRow of allOutputRows) {
    // outputRow = [WOID, Name, TimesheetID, Status, Start, End, Approved, ST, OT, DT, Others, NB, Amount]
    if (outputRow[3] === 'No Data') continue; // skip no-data placeholder rows
    const [, , tsId, status, start, end, approved, st, ot, dt, others, nb, amount] = outputRow;
    fieldsRows.push([status, tsId, start, end, approved, st, ot, dt, others, nb, amount]);
  }

  await writeOutputSheet(allOutputRows);
  await writeFieldsSheet(fieldsRows);

  console.log('\n=== COMPLETE ===');
  console.log(`Total WOs processed : ${woList.length}`);
  console.log(`Total rows written  : ${allOutputRows.length}`);
  console.log(`Output file         : ${XLSX_PATH}`);
  console.log(`Output sheet        : ${SHEET_NAME}`);

  expect(allOutputRows.length).toBeGreaterThan(0);
});
