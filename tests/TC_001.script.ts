import { test, Page, BrowserContext } from '@playwright/test';
import * as ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import axios from 'axios';

// ─── Config ────────────────────────────────────────────────────────────────────
const TENANT_ID     = process.env.TENANT_ID     || 'faa3f9fc-9b37-406b-b37d-58a9f045c17a';
const CLIENT_ID     = process.env.CLIENT_ID     || '753ed0ae-e240-4397-92fe-8f05171a0e32';
const CLIENT_SECRET = process.env.CLIENT_SECRET || '';   // set via env var — do not hardcode
const ONEDRIVE_USER = process.env.ONEDRIVE_USER || 'adarsh.pagadi@simplify3x.com';
const ONEDRIVE_ITEM = '01KUWJBDDPQHXJHEKAU5BLW3VEAJ7GKEBI';
const SHARING_URL   = 'https://simplify3xsoftware-my.sharepoint.com/:x:/g/personal/adarsh_pagadi_simplify3x_com/IQBvge6TkUCnQrtupAJ-ZRAoAfhLxv3atyfMDbLuuvW5ZuI?e=cGyLZG';

const XLSX_PATH       = path.join(os.tmpdir(), 'FieldGlass_temp.xlsx');
const CHECKPOINT_PATH = path.join(os.tmpdir(), 'fg_checkpoint.json');
const FROM_DATE       = '01/07/2026';
const TO_DATE         = '31/07/2026';
const SHEET_NAME      = '01-Jul-2026_31-Jul-2026';
const SUMMARY_SHEET   = 'Summary';
const FIELDS_SHEET    = 'fields';
const SAP_EXPECTED_ST = 40;

// ─── Slice control ─────────────────────────────────────────────────────────────
// Set TEST_SLICE to a number to run only the first N WOIDs (for testing).
// Set to 0 or Infinity to run ALL WOIDs.
const TEST_SLICE = 0;   // ← change to 0 to run all 85, or set a number for testing

// ─── Checkpoint helpers (crash recovery) ──────────────────────────────────────
// After every WO, results are saved to CHECKPOINT_PATH.
// On restart the script reads this and skips already-done WOIDs.
interface CheckpointData {
  detailRows:  string[][];
  summaryRows: string[][];
  fieldsRows:  string[][];
  doneWoids:   string[];
}

function loadCheckpoint(): CheckpointData {
  try {
    if (fs.existsSync(CHECKPOINT_PATH)) {
      const data = JSON.parse(fs.readFileSync(CHECKPOINT_PATH, 'utf8')) as CheckpointData;
      console.log(`[Checkpoint] Resuming — ${data.doneWoids.length} WOIDs already done`);
      return data;
    }
  } catch { /* corrupted — start fresh */ }
  return { detailRows: [], summaryRows: [], fieldsRows: [], doneWoids: [] };
}

function saveCheckpoint(cp: CheckpointData): void {
  fs.writeFileSync(CHECKPOINT_PATH, JSON.stringify(cp));
}

function clearCheckpoint(): void {
  try { if (fs.existsSync(CHECKPOINT_PATH)) fs.unlinkSync(CHECKPOINT_PATH); } catch { /* */ }
}

// ─── Types ─────────────────────────────────────────────────────────────────────
interface WORecord {
  simplify3xId: string;
  domainId:     string;
  woid:         string;
  name:         string;
  doj:          string;
  location:     string;
  email:        string;
  phone:        string;
}

interface TimesheetRow {
  status: string; tsId: string; startDt: string; endDt: string;
  approved: string; st: string; ot: string; dt: string;
  others: string; nb: string; amount: string;
}

interface DayDetail {
  dayLabel: string; isSat: boolean; isSun: boolean; hours: number;
}

// ─── Graph / SharePoint ────────────────────────────────────────────────────────
async function getGraphToken(): Promise<string> {
  const res = await axios.post(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({ grant_type: 'client_credentials', client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default' }).toString(),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );
  return res.data.access_token;
}

function encodeSharingUrl(url: string): string {
  const b64 = Buffer.from(url).toString('base64');
  return 'u!' + b64.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

async function downloadFromSharePoint(): Promise<void> {
  const token = await getGraphToken();
  try {
    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${ONEDRIVE_USER}/drive/items/${ONEDRIVE_ITEM}/content`,
      { headers: { Authorization: `Bearer ${token}` }, responseType: 'arraybuffer' });
    fs.writeFileSync(XLSX_PATH, Buffer.from(res.data));
    console.log(`[Graph] Downloaded ${res.data.byteLength} bytes`); return;
  } catch (e: any) { console.log(`[Graph] Direct failed (${e?.response?.status ?? e?.message})`); }
  const res = await axios.get(
    `https://graph.microsoft.com/v1.0/shares/${encodeSharingUrl(SHARING_URL)}/driveItem/content`,
    { headers: { Authorization: `Bearer ${token}` }, responseType: 'arraybuffer' });
  fs.writeFileSync(XLSX_PATH, Buffer.from(res.data));
  console.log(`[Graph] Downloaded via shares (${res.data.byteLength} bytes)`);
}

async function uploadToSharePoint(): Promise<void> {
  const token = await getGraphToken();
  await axios.put(
    `https://graph.microsoft.com/v1.0/users/${ONEDRIVE_USER}/drive/items/${ONEDRIVE_ITEM}/content`,
    fs.readFileSync(XLSX_PATH),
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
  console.log('[Graph] SharePoint updated ✓');
}

// ─── Excel: read WO sheet ──────────────────────────────────────────────────────
// WO sheet columns: 1=Simplify3x ID, 2=Domain ID, 3=WOID, 4=Name,
//                   5=DOJ, 6=Location, 7=Email, 8=Phone
async function readWOSheet(): Promise<WORecord[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  const ws = wb.getWorksheet('WO');
  if (!ws) throw new Error('WO sheet not found');
  const items: WORecord[] = [];
  ws.eachRow({ includeEmpty: false }, (row, n) => {
    if (n === 1) return;
    const woid = String(row.getCell(3).value ?? '').trim();
    if (!woid) return;
    items.push({
      simplify3xId: String(row.getCell(1).value ?? '').trim(),
      domainId:     String(row.getCell(2).value ?? '').trim(),
      woid,
      name:         String(row.getCell(4).value ?? '').trim(),
      doj:          String(row.getCell(5).value ?? '').trim(),
      location:     String(row.getCell(6).value ?? '').trim(),
      email:        String(row.getCell(7).value ?? '').trim(),
      phone:        String(row.getCell(8).value ?? '').trim(),
    });
  });
  return items;
}

// ─── Excel: write Detail sheet ─────────────────────────────────────────────────
// Detail sheet: removed DT Hours and Others Hours columns per requirements
// OT Hours = actual overtime (Sat/Sun>0 or weekday>8)
// NB Hours = total leave hours (weekday hours that are <8, summed as missing hours)
const DETAIL_HEADERS = [
  'WOID', 'Worker Name', 'Timesheet ID', 'Status',
  'Start Date', 'End Date', 'Approved Date',
  'ST Hours', 'OT Hours', 'NB Hours', 'Amount (INR)',
];
async function writeDetailSheet(rows: string[][]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  let ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) { ws = wb.addWorksheet(SHEET_NAME); } else { ws.spliceRows(1, ws.rowCount); }
  const hdr = ws.addRow(DETAIL_HEADERS);
  hdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
  hdr.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  hdr.alignment = { horizontal: 'center', vertical: 'middle' };
  for (const row of rows) ws.addRow(row);
  ws.columns.forEach(col => { let m = 10; col.eachCell?.({ includeEmpty: false }, c => { const v = c.value?.toString() ?? ''; if (v.length > m) m = v.length; }); col.width = m + 2; });
  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`[Excel] Detail: ${rows.length} rows → "${SHEET_NAME}"`);
}

// ─── Excel: write fields sheet ─────────────────────────────────────────────────
async function writeFieldsSheet(rows: string[][]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  const ws = wb.getWorksheet(FIELDS_SHEET);
  if (!ws) { console.log('[Excel] "fields" sheet not found — skipping'); return; }
  if (ws.rowCount > 1) ws.spliceRows(2, ws.rowCount - 1);
  for (const row of rows) ws.addRow(row);
  ws.columns.forEach(col => { let m = 10; col.eachCell?.({ includeEmpty: false }, c => { const v = c.value?.toString() ?? ''; if (v.length > m) m = v.length; }); col.width = m + 2; });
  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`[Excel] Fields: ${rows.length} rows`);
}

// ─── Excel: write Summary sheet (exact Image 1 columns) ───────────────────────
// Cols: Simplify3x ID | Domain ID | WOID | Name | Total Hrs | Timesheet status | TS Hrs | Leaves | Overtime
// Leaves   → written when total TS Hrs < 200 (under-hours) — weekday dates where hours = 0
// Overtime → written when total TS Hrs > 200 (over-hours) — dates where ST > 40 per week
const SUMMARY_HEADERS = [
  'Simplify3x ID', 'Domain ID', 'WOID', 'Name',
  'Total Hrs', 'Timesheet status', 'TS Hrs', 'NB', 'Overtime',
];
const MONTHLY_ST_THRESHOLD = 200; // monthly standard hours threshold
async function writeSummarySheet(rows: string[][]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  let ws = wb.getWorksheet(SUMMARY_SHEET);
  if (!ws) { ws = wb.addWorksheet(SUMMARY_SHEET); } else { ws.spliceRows(1, ws.rowCount); }
  const hdr = ws.addRow(SUMMARY_HEADERS);
  hdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203864' } };
  hdr.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  hdr.alignment = { horizontal: 'center', vertical: 'middle' };
  for (const row of rows) {
    const dataRow = ws.addRow(row);
    // Yellow highlight when TS Hrs ≠ 40
    const tsHrs = parseFloat(row[10] ?? '');
    if (!isNaN(tsHrs) && tsHrs !== SAP_EXPECTED_ST) {
      dataRow.eachCell({ includeEmpty: false }, cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      });
    }
  }
  ws.columns.forEach(col => { let m = 10; col.eachCell?.({ includeEmpty: false }, c => { const v = c.value?.toString() ?? ''; if (v.length > m) m = v.length; }); col.width = m + 2; });
  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`[Excel] Summary: ${rows.length} rows → "${SUMMARY_SHEET}"`);
}

// ─── Browser helpers ───────────────────────────────────────────────────────────
async function settle(page: Page, ms = 3000) {
  await page.waitForLoadState('load').catch(() => {});
  await page.waitForTimeout(ms).catch(() => {});
}

async function acceptCookies(page: Page) {
  for (const sel of ['button:has-text("Accept All")', 'button:has-text("Accept all")', '#onetrust-accept-btn-handler']) {
    try { const b = page.locator(sel).first(); if (await b.isVisible({ timeout: 3000 })) { await b.click(); await page.waitForTimeout(1500); return; } } catch { /* */ }
  }
}

async function dismissPanel(page: Page) {
  for (const sel of ['button[aria-label="Close"]', 'button[title="Close"]', 'button.close', 'button:has-text("×")']) {
    try { const b = page.locator(sel).first(); if (await b.isVisible({ timeout: 1500 })) { await b.click(); await page.waitForTimeout(600); return; } } catch { /* */ }
  }
  await page.keyboard.press('Escape').catch(() => {});
  await page.waitForTimeout(400).catch(() => {});
}

// Tab guard: disabled — do NOT close any tabs.
// SAP opens a new tab on login; we keep all tabs open and just track the active one.
function registerTabGuard(_ctx: BrowserContext): (p: Page) => void {
  const handler = (_newPage: Page) => { /* no-op — do not close any tab */ };
  return handler;
}

// ─── Login ─────────────────────────────────────────────────────────────────────
async function doLogin(page: Page, ctx: BrowserContext): Promise<Page> {
  console.log('[Login] → SAP Fieldglass...');
  await page.goto('https://www.us.fieldglass.cloud.sap/login.do');
  await page.waitForLoadState('domcontentloaded');
  await acceptCookies(page);
  await dismissPanel(page);
  await page.waitForTimeout(2000);

  const user = page.locator('input[name="username"], input[id="username"], input[type="text"]').first();
  await user.waitFor({ state: 'visible', timeout: 15000 });
  await user.fill('Chethan K');
  const pass = page.locator('input[type="password"]').first();
  await pass.waitFor({ state: 'visible', timeout: 10000 });
  await pass.fill('AmmuCA@@2002');

  const newTabP = ctx.waitForEvent('page', { timeout: 40000 }).catch(() => null);
  const btn = page.locator('button:has-text("Sign In"), button[type="submit"], input[type="submit"]').first();
  await btn.waitFor({ state: 'visible', timeout: 10000 });
  await btn.click();
  console.log('[Login] Sign In clicked');

  const newTab = await newTabP;
  let ap: Page;
  if (newTab) {
    await newTab.waitForLoadState('domcontentloaded').catch(() => {});
    // Do NOT close original tab — keep existing Chrome browser open
    ap = newTab;
  } else {
    ap = page;
  }

  // Wait until we land on a non-login page, always using the LATEST active page
  // SAP may open additional redirect tabs — pick whichever page is not login.do
  for (let i = 0; i < 30; i++) {
    // Always re-fetch the most recently active page from context
    const pages = ctx.pages();
    const active = pages.find(p => {
      try { const u = p.url(); return u && !u.includes('login.do') && u !== 'about:blank'; }
      catch { return false; }
    });
    if (active) { ap = active; break; }
    await new Promise(r => setTimeout(r, 2000));
    if (i === 29) throw new Error('[Login] Timeout — still on login.do after 60s');
  }

  await settle(ap, 5000).catch(() => {});
  console.log(`[Login] OK → ${ap.url()}`);
  await dismissPanel(ap).catch(() => {});
  await ap.waitForTimeout(3000).catch(() => {});
  return ap;
}

// ─── Timesheet list parser ─────────────────────────────────────────────────────
function parseTimesheets(rawText: string): TimesheetRow[] {
  const STATUS_PREFIXES = ['Invoiced', 'Draft', 'Pending', 'Rejected', 'Submitted', 'Recalled'];
  const isStatus = (l: string) => STATUS_PREFIXES.some(p => l === p || l.startsWith(p + ' '));
  const isDate   = (l: string) => /^\d{2}\/\d{2}\/\d{4}/.test(l);
  const isNum    = (l: string) => /^-?[\d,]+(\.\d+)?$/.test(l);
  const COL      = 11;
  const raw = rawText.split('\n').map(l => l.trim()).filter(Boolean);
  const lines: string[] = [];
  for (let i = 0; i < raw.length; i++) {
    const cur = raw[i], nxt = raw[i + 1] ?? '';
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(cur) && /^\d{2}:\d{2}\s+(AM|PM)$/.test(nxt)) { lines.push(`${cur} ${nxt}`); i++; }
    else lines.push(cur);
  }
  const tsIdx = lines.findIndex(l => l === 'Time Sheets');
  if (tsIdx === -1) return [];
  let hdrIdx = -1;
  for (let i = tsIdx; i < Math.min(tsIdx + 25, lines.length); i++) { if (lines[i] === 'Status') { hdrIdx = i; break; } }
  if (hdrIdx === -1) return [];
  let ds = hdrIdx + COL;
  if (lines[ds] === 'All') ds++;
  const isBnd = (l: string) =>
    l.startsWith('Clear Sort') || l.startsWith('Clear Filters') || l.startsWith('No items') ||
    l.startsWith('Press enter') || /^\d+-\d+ of \d+/.test(l) ||
    l === 'Absences' || l === 'Expense Sheets' || l === 'Credit/Debit Memo';
  const rows: TimesheetRow[] = [];
  let i = ds;
  while (i < lines.length) {
    const line = lines[i];
    if (isBnd(line)) break;
    if (isStatus(line)) {
      const c: string[] = [line];
      let j = i + 1;
      if (j < lines.length && !isStatus(lines[j]) && !isBnd(lines[j])) c.push(lines[j++]); else c.push('');
      for (let k = 0; k < 3; k++) { if (j < lines.length && isDate(lines[j])) c.push(lines[j++]); else c.push(''); }
      for (let k = 5; k < COL; k++) { if (j < lines.length && isNum(lines[j]) && !isStatus(lines[j]) && !isBnd(lines[j])) c.push(lines[j++]); else c.push(''); }
      rows.push({ status: c[0], tsId: c[1], startDt: c[2], endDt: c[3], approved: c[4], st: c[5], ot: c[6], dt: c[7], others: c[8], nb: c[9], amount: c[10] });
      i = j;
    } else { i++; }
  }
  return rows;
}

// ─── Time Worked parser (inside a TS detail page) ─────────────────────────────
// SAP innerText format (confirmed from debug):
//   Each row is ONE line with tab-separated values:
//   "Day\t25/7"  "Sat\t26/7"  "Sun\t27/7"  "Mon\t28/7" ...
//   "Total\t0.00\t0.00\t8.00\t8.00\t8.00\t8.00\t8.00\t40.00"
//
// Strategy:
//   1. Find consecutive lines matching "DayName\tDate" pattern (Sat\t18/7 etc.)
//      These are the column headers.
//   2. Find the "Total\t..." line after them — split by tab to get per-day hours.
//   3. The last "Total" line before the next section = grand total.
function parseTimeWorked(rawText: string): DayDetail[] {
  const lines   = rawText.split('\n').map(l => l.trim()).filter(Boolean);
  const isN     = (s: string) => /^-?\d+(\.\d+)?$/.test(s);
  const isDate  = (s: string) => /^\d{1,2}\/\d{1,2}$/.test(s);
  const isDayNm = (s: string) => /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)$/i.test(s);

  // ── Step 1: Find the column header block ─────────────────────────────────────
  // Confirmed SAP innerText format (from debug):
  //   Line N  :  "Day\t25/7"   ← row-label="Day",  col-1 date=25/7 (no day name yet)
  //   Line N+1:  "Sat\t26/7"   ← day name "Sat" belongs to col-1 (25/7), col-2 date=26/7
  //   Line N+2:  "Sun\t27/7"   ← day name "Sun" belongs to col-2 (26/7), col-3 date=27/7
  //   ...
  //   Line N+7:  "Fri\tTotal Worked" ← day name "Fri" belongs to col-7, last column
  //
  // So: dayName[k] is on line[k+1].split('\t')[0], and date[k] is on line[k].split('\t')[1]
  let dayLabels: DayDetail[] = [];
  let headerEndIdx = -1;

  for (let i = 0; i < lines.length - 4; i++) {
    const firstParts = lines[i].split('\t');
    // Start of header block: first token is "Day" or a day-name, second token is a date
    if (firstParts.length < 2 || !isDate(firstParts[1])) continue;

    // Collect dates from this block (date is secondToken of each line)
    const dates: string[] = [];
    let j = i;
    while (j < lines.length && dates.length < 7) {
      const p = lines[j].split('\t');
      if (p.length < 2 || !isDate(p[1])) break;
      dates.push(p[1]);
      j++;
    }
    if (dates.length < 5) continue;

    // Now collect day-names: dayName[k] = firstToken of line[i+k+1]
    // (offset by 1: line i+1 has "Sat" for date at line i's second token, etc.)
    const candidate: DayDetail[] = [];
    for (let k = 0; k < dates.length; k++) {
      const dnLine = lines[i + k + 1] ?? '';
      const dnParts = dnLine.split('\t').map(p => p.trim()); // trim \r and whitespace
      const dn = isDayNm(dnParts[0]) ? dnParts[0] : '?';
      candidate.push({
        dayLabel: `${dn} ${dates[k]}`,
        isSat:    dn.toLowerCase() === 'sat',
        isSun:    dn.toLowerCase() === 'sun',
        hours:    0,
      });
    }
    if (candidate.length >= 5) {
      dayLabels    = candidate;
      headerEndIdx = j;
      break;
    }
  }

  if (!dayLabels.length) return [];
  const N = dayLabels.length;

  // ── Step 2: Find "Total\tV1\tV2\t..." line — keep the LAST one ───────────────
  // This is the grand-total row at the bottom of the Time Worked table.
  let bestVals: number[] | null = null;
  for (let i = headerEndIdx; i < Math.min(headerEndIdx + 150, lines.length); i++) {
    const p = lines[i].split('\t');
    if (p[0] === 'Total' && p.length >= N + 1) {
      const vals = p.slice(1, N + 1).map(v => parseFloat(v) || 0);
      if (vals.length === N) bestVals = vals; // keep overwriting → last one wins
    }
  }

  if (bestVals) {
    for (let m = 0; m < N; m++) dayLabels[m].hours = bestVals[m];
    return dayLabels;
  }

  // ── Fallback: sum all "Time Worked\tV1\tV2\t..." lines ───────────────────────
  for (let i = headerEndIdx; i < Math.min(headerEndIdx + 150, lines.length); i++) {
    const p = lines[i].split('\t');
    if (p[0] === 'Time Worked' && p.length >= N + 1) {
      const vals = p.slice(1, N + 1).map(v => parseFloat(v) || 0);
      if (vals.length === N) {
        for (let m = 0; m < N; m++) dayLabels[m].hours += vals[m];
      }
    }
  }

  return dayLabels;
}

// ─── Drill into one TS ID — return leave days + weekend work ──────────────────
interface DrillResult {
  leaveDates:  string;  // dates of weekday absences e.g. "Mon 7/7, Tue 14/7"
  nbCount:     number;  // count of absent weekdays
  weekendWork: string;  // weekend hours e.g. "Sat 4/7(8h)"
  otHours:     number;  // total weekend hours worked
}

async function drillIntoTS(ap: Page, tsId: string): Promise<DrillResult> {
  console.log(`   → drill TS: ${tsId}`);

  // Save T&E URL BEFORE clicking — SAP's T&E page loads via POST.
  // goBack() causes ERR_CACHE_MISS; navigate directly back to saved URL.
  const teUrl = ap.url();

  // SAP renders TS IDs as links inside table cells — try multiple locator strategies
  const link = ap.locator(
    `a:has-text("${tsId}"), a[href*="${tsId}"], td:has-text("${tsId}") a, [data-id="${tsId}"] a`
  ).first();
  try {
    await link.waitFor({ state: 'visible', timeout: 10000 });
    await link.click();
    await settle(ap, 4000);
  } catch {
    // Final fallback: find any link on the page whose text exactly matches the tsId
    try {
      const fallback = ap.getByText(tsId, { exact: true }).first();
      await fallback.waitFor({ state: 'visible', timeout: 5000 });
      await fallback.click();
      await settle(ap, 4000);
    } catch {
      console.log(`   TS link not found: ${tsId}`);
      return { leaveDates: '', nbCount: 0, weekendWork: '', otHours: 0 };
    }
  }

  const rawText = await ap.evaluate(() => document.body.innerText);
  const days = parseTimeWorked(rawText);

  // Rules from image:
  //   Sat/Sun > 0            → OT (user worked on weekend)
  //   Mon-Fri > 8            → OT (extra hours on weekday)
  //   Mon-Fri < 8 (incl. 0) → Leave (user was absent or left early)
  //   Mon-Fri = 8            → Normal day
  const otList:    string[] = [];  // OT date+hours labels
  let totalOtHours  = 0;           // sum of all OT hours
  const leaveList:  string[] = [];  // Leave date+hours labels
  let totalLeaveHrs = 0;           // sum of missing hours (8 - actual)

  for (const d of days) {
    if (d.isSat || d.isSun) {
      // Weekend: any hours > 0 = OT
      if (d.hours > 0) {
        otList.push(`${d.dayLabel}(${d.hours}h)`);
        totalOtHours += d.hours;
      }
    } else {
      // Weekday: normal = 8h
      if (d.hours > 8) {
        // Worked extra — e.g. 10h means 2h OT
        const extra = d.hours - 8;
        otList.push(`${d.dayLabel}(+${extra}h)`);
        totalOtHours += extra;
      } else if (d.hours < 8) {
        // Worked less than 8 (or 0) — leave
        const missing = 8 - d.hours;
        leaveList.push(`${d.dayLabel}(${d.hours}h)`);
        totalLeaveHrs += missing;
      }
    }
  }

  console.log(`   OT:[${otList.join(',')}]  Leave:[${leaveList.join(',')}]`);

  // Navigate back to desktop (dashboard) so the next WO search finds the global search box.
  // Going back to teUrl (T&E POST page) leaves a hidden autocomplete as the first search match.
  await ap.goto('https://www.us.fieldglass.cloud.sap/desktop.do').catch(() => {});
  await settle(ap, 3000);

  return {
    leaveDates:   leaveList.join(', '),
    nbCount:      totalLeaveHrs,   // total missing hours (not count of days)
    weekendWork:  otList.join(', '),
    otHours:      totalOtHours,
  };
}

// ─── Fetch one WO ──────────────────────────────────────────────────────────────
interface FetchResult { detailRows: string[][]; summaryRow: string[]; fieldsRows: string[][]; }

function emptyResult(wo: WORecord): FetchResult {
  return {
    detailRows: [[wo.woid, wo.name, '', 'No Data', '', '', '', '', '', '', '', '', '']],
    // 9 cols matching SUMMARY_HEADERS: Simplify3x ID | Domain ID | WOID | Name | Total Hrs | TS status | TS Hrs | Leaves | Overtime
    summaryRow: [wo.simplify3xId, wo.domainId, wo.woid, wo.name, '0', 'No Data', '0', '', ''],
    fieldsRows: [],
  };
}

async function fetchWO(ap: Page, wo: WORecord): Promise<{ result: FetchResult; page: Page }> {
  console.log(`\n[WO] ${wo.woid} — ${wo.name}`);

  // Search
  const search = ap.locator('input[placeholder*="Search by ID"], input[placeholder*="Search"], input[type="search"]').first();
  await search.waitFor({ state: 'visible', timeout: 25000 });
  await ap.waitForTimeout(500);
  await search.fill(wo.woid);
  await ap.waitForTimeout(500);
  await search.press('Enter');
  await settle(ap, 4000);

  // Global search → Go to Details
  if (ap.url().includes('global_search.do')) {
    const link = ap.locator('a:has-text("Go to Details"), a[href*="work_order_detail"]').first();
    try { await link.waitFor({ state: 'visible', timeout: 12000 }); await link.click(); await settle(ap, 4000); }
    catch { console.log(`  no details link`); return { result: emptyResult(wo), page: ap }; }
  }

  // Time & Expense tab
  const teTab = ap.locator('a[href*="tabId=timeAndExpense"], a:has-text("Time & Expense")').first();
  try { await teTab.waitFor({ state: 'visible', timeout: 18000 }); }
  catch { console.log(`  no T&E tab`); return { result: emptyResult(wo), page: ap }; }
  await teTab.click();
  await settle(ap, 4000);

  // Date filter
  const dates = ap.locator('input[aria-label="Open Calendar (DD/MM/YYYY)"]');
  try { await dates.first().waitFor({ state: 'visible', timeout: 15000 }); }
  catch { console.log(`  no date inputs`); return { result: emptyResult(wo), page: ap }; }
  await dates.nth(0).click({ clickCount: 3 }); await dates.nth(0).fill(FROM_DATE); await ap.keyboard.press('Tab'); await ap.waitForTimeout(1000);
  await dates.nth(1).click({ clickCount: 3 }); await dates.nth(1).fill(TO_DATE);   await ap.keyboard.press('Tab'); await ap.waitForTimeout(1000);
  await ap.keyboard.press('Enter');
  await settle(ap, 5000);

  const timesheets = parseTimesheets(await ap.evaluate(() => document.body.innerText));
  console.log(`  ${timesheets.length} timesheets`);
  if (!timesheets.length) return { result: emptyResult(wo), page: ap };

  // Detail rows: 10 cols — WOID | Name | TS ID | Status | Start | End | Approved | ST | OT | NB | Amount
  // (DT and Others columns removed per requirements)
  const detailRows = timesheets.map(ts => [
    wo.woid, wo.name, ts.tsId, ts.status,
    ts.startDt, ts.endDt, ts.approved,
    ts.st,    // ST Hours
    ts.ot,    // OT Hours — will be overwritten by drill data below
    ts.nb,    // NB Hours — will be overwritten by drill data below
    ts.amount,
  ]);

  // Fields sheet: Status | ID | Start | End | Approved | ST | OT | NB | Amount (no DT/Others)
  // Build as same-indexed array as timesheets so drill-in can update OT/NB by index
  const fieldsRows = timesheets.map(ts => [
    ts.status, ts.tsId, ts.startDt, ts.endDt, ts.approved,
    ts.st, // ST
    ts.ot, // OT — overwritten by drill data below (index 6)
    ts.nb, // NB — overwritten by drill data below (index 7)
    ts.amount,
  ]);

  const totalST  = timesheets.reduce((s, ts) => s + (parseFloat(ts.st) || 0), 0);
  const statuses = [...new Set(timesheets.map(ts => ts.status))].join(', ');

  // ── Drill into each TS where ST ≠ 40 ──────────────────────────────────────
  // Rules (confirmed from image):
  //   Sat/Sun > 0          → OT (weekend overtime)
  //   Mon-Fri > 8          → OT (extra weekday hours)
  //   Mon-Fri < 8 (incl 0) → Leave (absent or left early)
  //   Mon-Fri = 8          → Normal day
  // OT column in detail sheet  → total OT hours from drill
  // NB column in detail sheet  → total leave hours from drill (sum of missing hours)
  // Overtime col in Summary    → OT date+hour labels (always written when found)
  // NB col in Summary          → Leave date+hour labels (always written when found)
  const allLeaveDates: string[] = [];
  const allOtDates:    string[] = [];

  for (let tsIdx = 0; tsIdx < timesheets.length; tsIdx++) {
    const ts = timesheets[tsIdx];
    const st = parseFloat(ts.st) || 0;
    // Drill only when ST ≠ 40
    if (!ts.tsId || st === SAP_EXPECTED_ST) continue;

    const d = await drillIntoTS(ap, ts.tsId);

    // Collect for Summary columns
    if (d.weekendWork) allOtDates.push(d.weekendWork);   // OT dates (Sat/Sun + weekday extra)
    if (d.leaveDates)  allLeaveDates.push(d.leaveDates); // Leave dates (weekday < 8)

    // Write OT hours to detail OT (index 8) and fields OT (index 6)
    if (d.otHours > 0) { detailRows[tsIdx][8] = String(d.otHours); fieldsRows[tsIdx][6] = String(d.otHours); }
    // Write leave hours to detail NB (index 9) and fields NB (index 7)
    if (d.nbCount > 0) { detailRows[tsIdx][9] = String(d.nbCount); fieldsRows[tsIdx][7] = String(d.nbCount); }
  }

  // ── Summary sheet: always write whatever was found ─────────────────────────
  const leavesValue   = allLeaveDates.join(', ');
  const overtimeValue = allOtDates.join(', ');

  // Summary row columns:
  // Simplify3x ID | Domain ID | WOID | Name |
  // Total Hrs | Timesheet status | TS Hrs | Leaves | Overtime
  const summaryRow = [
    wo.simplify3xId, wo.domainId, wo.woid, wo.name,
    String(totalST),   // Total Hrs
    statuses,          // Timesheet status
    String(totalST),   // TS Hrs
    leavesValue,       // Leaves — weekday dates with 0 hours (only when total < 200)
    overtimeValue,     // Overtime — OT dates (only when total > 200)
  ];

  return { result: { detailRows, summaryRow, fieldsRows }, page: ap };
}

// ─── Main test — single session, crash-safe checkpoint resume ─────────────────
test('HR Timesheet — All WOs (single session, crash-safe)', async ({ page, context }) => {
  // Exactly 1 tab at all times — guard closes any extra tab SAP opens
  const guard = registerTabGuard(context);

  try {
    await downloadFromSharePoint();
    const allWOs = await readWOSheet();
    console.log(`\n[Run] ${allWOs.length} WOIDs total`);

    // Load checkpoint — resume from where a previous crash left off
    const cp = loadCheckpoint();
    const doneSet = new Set(cp.doneWoids);
    // Apply slice for test runs — TEST_SLICE=10 runs first 10 WOIDs, TEST_SLICE=0 runs all
    const slicedWOs = (TEST_SLICE > 0 && TEST_SLICE < Infinity) ? allWOs.slice(0, TEST_SLICE) : allWOs;
    const pending = slicedWOs.filter(w => !doneSet.has(w.woid));
    console.log(`[Run] Total: ${allWOs.length} | Slice: ${slicedWOs.length} | Already done: ${doneSet.size} | Pending: ${pending.length}`);

    if (!pending.length) {
      console.log('[Run] All WOIDs already completed — writing sheets and uploading...');
    } else {
      // Login once
      let ap = await doLogin(page, context);

      for (let idx = 0; idx < pending.length; idx++) {
        const wo = pending[idx];
        console.log(`\n[Run] [${doneSet.size + idx + 1}/${allWOs.length}] ${wo.woid}`);

        // Check browser is still alive — if not, close everything and re-open exactly 1 tab
        let alive = false;
        try { ap.url(); alive = true; } catch { /* crashed */ }
        if (!alive) {
          console.log('[Run] Browser tab died — closing all tabs and re-logging in...');
          // Close all existing tabs first
          for (const p of context.pages()) {
            try { await p.close().catch(() => {}); } catch { /* */ }
          }
          // Open exactly 1 new tab and login
          try {
            const freshPage = await context.newPage();
            ap = await doLogin(freshPage, context);
          } catch (e) {
            console.log(`[Run] Re-login failed: ${e} — marking ${wo.woid} No Data and continuing`);
            cp.detailRows.push([wo.woid, wo.name, '', 'No Data', '', '', '', '', '', '', '', '', '']);
            cp.summaryRows.push([wo.simplify3xId, wo.domainId, wo.woid, wo.name, '0', 'No Data', '0', '', '']);
            cp.doneWoids.push(wo.woid);
            saveCheckpoint(cp);
            continue;
          }
        }

        try {
          const { result, page: newAp } = await fetchWO(ap, wo);
          ap = newAp;
          cp.detailRows.push(...result.detailRows);
          cp.summaryRows.push(result.summaryRow);
          cp.fieldsRows.push(...result.fieldsRows);
          cp.doneWoids.push(wo.woid);
          // Save after every WO — crash recovery point
          saveCheckpoint(cp);
          console.log(`  ✓ saved checkpoint (${cp.doneWoids.length}/${allWOs.length})`);
        } catch (err) {
          console.log(`[Run] ERROR ${wo.woid}: ${err}`);
          cp.detailRows.push([wo.woid, wo.name, '', 'Error', '', '', '', '', '', '', '', '', '']);
          cp.summaryRows.push([wo.simplify3xId, wo.domainId, wo.woid, wo.name, '0', 'Error', '0', '', '']);
          cp.doneWoids.push(wo.woid);
          saveCheckpoint(cp);
          // Try to recover browser
          try { ap.url(); } catch {
            try {
              for (const p of context.pages()) { try { await p.close().catch(() => {}); } catch { /* */ } }
              ap = await doLogin(await context.newPage(), context);
            } catch { /* will re-check alive on next iteration */ }
          }
        }
      }
    }

    // All WOIDs done — write Excel sheets and upload
    console.log('\n[Run] Writing Excel sheets...');
    await writeDetailSheet(cp.detailRows);
    await writeFieldsSheet(cp.fieldsRows);
    await writeSummarySheet(cp.summaryRows);
    await uploadToSharePoint();

    // Clear checkpoint only after successful upload
    clearCheckpoint();

    const noData  = cp.summaryRows.filter(r => r[9] === 'No Data' || r[9] === 'Error').length;
    const drilled = cp.summaryRows.filter(r => r[11] !== '' || r[12] !== '').length;

    console.log('\n═══════════════════════════════════════════');
    console.log('           EXECUTION COMPLETE              ');
    console.log('═══════════════════════════════════════════');
    console.log(`Total WOIDs          : ${allWOs.length}`);
    console.log(`Summary rows         : ${cp.summaryRows.length}`);
    console.log(`Detail rows          : ${cp.detailRows.length}`);
    console.log(`Drilled (ST≠40)      : ${drilled}`);
    console.log(`No Data / Error      : ${noData}`);
    console.log(`Sheets written       : "${SHEET_NAME}", "${SUMMARY_SHEET}", "${FIELDS_SHEET}"`);
    console.log(`SharePoint updated   : ✓`);
    console.log('═══════════════════════════════════════════');

  } finally {
    context.off('page', guard);
  }
});
