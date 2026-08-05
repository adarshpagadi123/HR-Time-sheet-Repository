import type { WalnutContext } from './walnut';
import type { Page } from '@playwright/test';
import * as ExcelJS from 'exceljs';

/** @walnut_method
 * name: Search IDs from Excel and Write Results to Output Sheet
 * description: Open workbook ${filePath} sheet ${inputSheetName} column ${idColumnName} search each ID on web UI with from date ${fromDate} and to date ${toDate} and write results into sheet ${outputSheetName}
 * actionType: custom_search_and_export_to_excel
 * context: shared
 * needsLocator: false
 * category: Data Processing
 */
export async function searchAndExportToExcel(ctx: WalnutContext) {
  const filePath        = ctx.args[0];
  const inputSheetName  = ctx.args[1];
  const idColumnName    = ctx.args[2];
  const fromDate        = ctx.args[3];
  const toDate          = ctx.args[4];
  const outputSheetName = ctx.args[5];

  // ── LOCATORS — edit here when the DOM changes, nowhere else ─────────────────
  const LOCATORS = {
    searchInputHost:  'ui5-input#searchFieldInput',
    searchInputInner: 'input#inner',
    tabTimeExpense:   '//a[@id="tab_timeAndExpense"]',
    fromDateInput:    '//input[@id="filterStartDate"]',
    toDateInput:      '//input[@id="filterEndDate"]',
    applyFilterBtn:   '//input[@id="timeAndExpenseFitlerBtn"]',
    resultRow:        'div[role="row"][id*="timeSheet_workOrder_list_byWorkerId"]',
    columns: [
      { index: 0,  header: 'Status',        child: 'span.fd-object-status__text'  },
      { index: 1,  header: 'Timesheet ID',  child: 'a.archiveLink'                },
      { index: 2,  header: 'Start Date',    child: 'div.jqx-grid-cell-left-align' },
      { index: 3,  header: 'End Date',      child: 'div.jqx-grid-cell-left-align' },
      { index: 4,  header: 'Approved Date', child: 'div.jqx-grid-cell-left-align' },
      { index: 5,  header: 'ST Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 6,  header: 'OT Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 7,  header: 'DT Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 8,  header: 'Others Hours',  child: 'div.jqx-grid-cell-left-align' },
      { index: 9,  header: 'NB Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 10, header: 'Amount (INR)',   child: 'div.jqx-grid-cell-left-align' },
    ],
    workerName:     'h1[data-help-id="TITLE_270"] span.titlePrimary',
    pageReadyProbe: '//a[@id="tab_timeAndExpense"]',
    homeBtn:        '//li[@id="homeMenuTitle"]//a[@title="Home"]',
    homeReadyProbe: 'ui5-input#searchFieldInput',
  } as const;

  // ── Resolve ExcelJS cell value to plain string ───────────────────────────────
  function cellText(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      if ('result' in value) return String((value as ExcelJS.CellFormulaValue).result ?? '');
      if ('richText' in value)
        return (value as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join('');
      if ('text' in value) return (value as ExcelJS.CellHyperlinkValue).text;
      if (value instanceof Date) return value.toISOString().slice(0, 10);
    }
    return String(value);
  }

  // ── Bold + light-blue header row ─────────────────────────────────────────────
  function styleHeaderRow(row: ExcelJS.Row, colCount: number): void {
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.font      = { bold: true };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.border    = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
    row.commit();
  }

  // ── Alternating white/grey data row ─────────────────────────────────────────
  function styleDataRow(row: ExcelJS.Row, colCount: number, dataIndex: number): void {
    const bgColor = dataIndex % 2 === 0 ? 'FFFFFFFF' : 'FFF2F2F2';
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
      cell.border    = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
    row.commit();
  }

  // ── Poll browser JS until a real Timesheet ID appears in the grid ────────────
  // Returns false if no results appear within timeoutMs (no-results ID).
  async function waitForGridReady(page: Page, timeoutMs = 90000): Promise<boolean> {
    try {
      await page.waitForFunction(
        ({ rowSel }: { rowSel: string }) => {
          const rows = document.querySelectorAll(rowSel);
          for (const row of Array.from(rows)) {
            const cell = row.querySelector('div[columnindex="1"][role="gridcell"]');
            const title = cell?.getAttribute('title') ?? '';
            if (title.trim().length > 3) return true;
          }
          return false;
        },
        { rowSel: LOCATORS.resultRow },
        { timeout: timeoutMs },
      );
      return true;
    } catch {
      return false;
    }
  }

  // ── Read all result rows in one page.evaluate() call ────────────────────────
  // Reads cell values from the `title` attribute (stamped by jqx at row-attach
  // time). Falls back to child textContent if title is empty.
  async function scrapeResultRows(
    page: Page,
    sourceId: string,
  ): Promise<Record<string, string>[]> {
    type ColDef = { index: number; header: string; child: string };
    const cols: ColDef[] = LOCATORS.columns.map((c) => ({ ...c }));

    const rawRows = await page.evaluate(
      ({ rowSel, columns }: { rowSel: string; columns: ColDef[] }) => {
        const rowEls = document.querySelectorAll(rowSel);
        const results: Record<string, string>[] = [];

        rowEls.forEach((rowEl) => {
          const record: Record<string, string> = {};
          columns.forEach(({ index, header, child }) => {
            const cell = rowEl.querySelector(`div[columnindex="${index}"][role="gridcell"]`);
            if (!cell) { record[header] = ''; return; }
            const title = (cell.getAttribute('title') ?? '').trim();
            if (title !== '') { record[header] = title; return; }
            const childEl = cell.querySelector(child);
            record[header] = (childEl?.textContent ?? '').trim();
          });
          results.push(record);
        });

        return results;
      },
      { rowSel: LOCATORS.resultRow, columns: cols },
    );

    // Skip jqx placeholder rows — real rows always have a Timesheet ID
    return rawRows
      .filter((r) => (r['Timesheet ID'] ?? '').trim().length > 3)
      .map((r) => ({ WOID: sourceId, ...r }));
  }

  // ── 1. Read IDs from input sheet ─────────────────────────────────────────────
  ctx.log(`Reading workbook: ${filePath}`);

  const inputWorkbook = new ExcelJS.Workbook();
  await inputWorkbook.xlsx.readFile(filePath);

  const inSheet = inputWorkbook.getWorksheet(inputSheetName);
  if (!inSheet) {
    const available = inputWorkbook.worksheets.map((ws: ExcelJS.Worksheet) => ws.name).join(', ');
    throw new Error(`Input sheet "${inputSheetName}" not found. Available sheets: ${available}`);
  }

  const headerRow  = inSheet.getRow(1);
  let   idColIndex = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell, colIdx: number) => {
    if (String(cell.value ?? '').trim().toLowerCase() === idColumnName.trim().toLowerCase()) {
      idColIndex = colIdx;
    }
  });

  if (idColIndex === -1) {
    const found: string[] = [];
    headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) =>
      found.push(String(cell.value ?? '')),
    );
    throw new Error(
      `Column "${idColumnName}" not found in row 1 of sheet "${inputSheetName}". ` +
      `Found headers: ${found.join(', ')}`,
    );
  }

  const ids: string[] = [];
  inSheet.eachRow({ includeEmpty: false }, (row: ExcelJS.Row, rowNumber: number) => {
    if (rowNumber === 1) return;
    const val = cellText(row.getCell(idColIndex).value).trim();
    if (val !== '') ids.push(val);
  });

  if (ids.length === 0) {
    throw new Error(`No IDs found in column "${idColumnName}" of sheet "${inputSheetName}".`);
  }
  ctx.log(`Found ${ids.length} ID(s) to process.`);

  // ── 2. Determine output column order ─────────────────────────────────────────
  // Lock to existing header if the output sheet already has one (resumed run).
  const defaultColOrder: string[] = ['WOID', 'Worker Name', ...LOCATORS.columns.map((c) => c.header)];
  let colOrder: string[] = defaultColOrder;
  {
    const existingSheet = inputWorkbook.getWorksheet(outputSheetName);
    if (existingSheet) {
      const hRow = existingSheet.getRow(1);
      const hdrs: string[] = [];
      hRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) => {
        hdrs.push(String(cell.value ?? ''));
      });
      if (hdrs.length > 0) {
        colOrder = hdrs;
        ctx.log(`Output sheet "${outputSheetName}" found — column order locked from header row.`);
      } else {
        ctx.log(`Output sheet "${outputSheetName}" found but has no header — will write header on first flush.`);
      }
    } else {
      ctx.log(`Output sheet "${outputSheetName}" not found — will create it on first flush.`);
    }
  }

  // Release input workbook — flushRows reloads the file fresh each call
  // @ts-ignore
  inputWorkbook._worksheets = [];

  // ── 3. flushRows — load fresh → append → save → discard ─────────────────────
  let dataRowCounter = 0;
  {
    const probe = new ExcelJS.Workbook();
    try {
      await probe.xlsx.readFile(filePath);
      const s = probe.getWorksheet(outputSheetName);
      if (s && s.rowCount > 1) {
        dataRowCounter = s.rowCount - 1;
      }
    } catch { /* output sheet may not exist yet */ }
  }

  async function flushRows(rows: Record<string, string>[]): Promise<void> {
    if (rows.length === 0) return;

    const cols = colOrder;
    const writeStart = Date.now();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);

    let sheet = wb.getWorksheet(outputSheetName);
    if (!sheet) {
      sheet = wb.addWorksheet(outputSheetName);
      sheet.columns = cols.map((h) => ({
        width: Math.min(40, Math.max(14, h.length + 4)),
      }));
      ctx.log(`Created output sheet "${outputSheetName}".`);
    }

    const cell1     = sheet.getRow(1).getCell(1);
    const hasHeader = cell1.value !== null && cell1.value !== undefined && cell1.value !== '';
    if (!hasHeader) {
      const hRow = sheet.getRow(1);
      cols.forEach((h, idx) => { hRow.getCell(idx + 1).value = h; });
      styleHeaderRow(hRow, cols.length);
      ctx.log(`Written header row to sheet "${outputSheetName}".`);
    }

    const nextRow = sheet.rowCount + 1;
    rows.forEach((record, idx) => {
      const row = sheet!.getRow(nextRow + idx);
      cols.forEach((h, colIdx) => {
        row.getCell(colIdx + 1).value = record[h] ?? '';
      });
      styleDataRow(row, cols.length, dataRowCounter + idx);
      row.commit();
    });

    dataRowCounter += rows.length;
    await wb.xlsx.writeFile(filePath);

    ctx.log(`[FLUSH] ${rows.length} row(s) written in ${Date.now() - writeStart}ms`);
  }

  // ── 4. Main loop ─────────────────────────────────────────────────────────────
  let totalWritten = 0;

  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    ctx.log(`[${i + 1}/${ids.length}] Processing ID: ${id}`);

    try {
      const searchHost  = ctx.page.locator(LOCATORS.searchInputHost);
      const searchInner = searchHost.locator(LOCATORS.searchInputInner);

      await searchInner.waitFor({ state: 'visible' });
      await searchInner.click();
      await searchInner.click({ clickCount: 3 });
      await searchInner.fill(id);
      await searchInner.press('Enter');

      await ctx.waitForVisible(LOCATORS.pageReadyProbe);

      const workerName = await ctx.page
        .locator(LOCATORS.workerName)
        .first()
        .textContent()
        .then((t: string | null) => t?.trim() ?? '')
        .catch(() => '');
      ctx.log(`  → Worker name: "${workerName}"`);

      await ctx.click(LOCATORS.tabTimeExpense);

      await ctx.waitForVisible(LOCATORS.fromDateInput);
      await ctx.clear(LOCATORS.fromDateInput);
      await ctx.type(LOCATORS.fromDateInput, fromDate);
      await ctx.pressKey('Tab');

      await ctx.waitForVisible(LOCATORS.toDateInput);
      await ctx.clear(LOCATORS.toDateInput);
      await ctx.type(LOCATORS.toDateInput, toDate);
      await ctx.pressKey('Tab');

      await ctx.waitForVisible(LOCATORS.applyFilterBtn);
      await ctx.click(LOCATORS.applyFilterBtn);

      const hasResults = await waitForGridReady(ctx.page, 90000);

      if (!hasResults) {
        ctx.log(`  → No results for ID: ${id} — writing sentinel row.`);
        await flushRows([{ WOID: id, 'Worker Name': workerName, Status: 'No Results' }]);
        totalWritten += 1;
      } else {
        const rows = await scrapeResultRows(ctx.page, id);
        rows.forEach((r) => { r['Worker Name'] = workerName; });
        ctx.log(`  → ${rows.length} row(s) scraped for ID: ${id}`);
        await flushRows(rows);
        totalWritten += rows.length;
        ctx.log(`  → Saved. Total rows written so far: ${totalWritten}`);
      }

      if (i < ids.length - 1) {
        ctx.log(`  → Navigating home...`);
        await ctx.click(LOCATORS.homeBtn);
        await ctx.page.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible' });
      }

    } catch (err) {
      ctx.log(`  Error processing ID "${id}": ${String(err)}`);
      await flushRows([{ WOID: id, error: String(err) }]);
      totalWritten += 1;

      try {
        await ctx.click(LOCATORS.homeBtn);
        await ctx.page.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible' });
      } catch { /* ignore — next iteration will fail early if page is broken */ }
    }

    if (i < ids.length - 1) {
      await new Promise((resolve) => setTimeout(resolve, 300));
    }
  }

  const fileName = filePath.replace(/\\/g, '/').split('/').pop() ?? filePath;
  ctx.log(`Done. ${totalWritten} total row(s) written to "${outputSheetName}" in ${fileName}.`);
}
