import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: Search IDs from Excel and Write Results to Output Sheet
 * description: Open workbook ${filePath} sheet ${inputSheetName} column ${idColumnName} search each ID on web UI with from date ${fromDate} and to date ${toDate} and write results into sheet ${outputSheetName}
 * actionType: custom_search_and_export_to_excel
 * context: web
 * needsLocator: false
 * category: Data Processing
 */
export async function searchAndExportToExcel(ctx: WalnutContext) {
  // Dynamic require via createRequire(__filename) so Node resolves 'exceljs'
  // from the live cache directory regardless of the cache hash.
  // import('module') is a Node built-in — esbuild does not bundle it.
  const { createRequire } = await import('module');
  const xl: any = createRequire(__filename)('exceljs');

  const webCtx = ctx as WalnutWebContext;

  const filePath        = ctx.args[0];
  const inputSheetName  = ctx.args[1];
  const idColumnName    = ctx.args[2];
  const fromDate        = ctx.args[3];
  const toDate          = ctx.args[4];
  const outputSheetName = ctx.args[5];

  // ── LOCATORS ──────────────────────────────────────────────────────────────────
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
    homeBtn:        '//li[@id="homeMenuTitle"]//a[@title="Home"]',
    homeReadyProbe: 'ui5-input#searchFieldInput',
  } as const;

  // ── Helpers ───────────────────────────────────────────────────────────────────

  function cellText(value: any): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      if ('result'   in value) return String(value.result ?? '');
      if ('richText' in value) return value.richText.map((r: any) => r.text).join('');
      if ('text'     in value) return value.text;
      if (value instanceof Date) return value.toISOString().slice(0, 10);
    }
    return String(value);
  }

  function styleHeaderRow(row: any, colCount: number): void {
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.font      = { bold: true };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
    row.commit();
  }

  function styleDataRow(row: any, colCount: number, dataIndex: number): void {
    const bgColor = dataIndex % 2 === 0 ? 'FFFFFFFF' : 'FFF2F2F2';
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
      cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
    row.commit();
  }

  // Polls until a real Timesheet ID (title length > 3) appears in the grid.
  // Returns false if nothing appears within the timeout (no results for this ID).
  async function waitForGridReady(timeoutMs = 90000): Promise<boolean> {
    try {
      await webCtx.page.waitForFunction(
        ({ rowSel }: { rowSel: string }) => {
          for (const row of Array.from(document.querySelectorAll(rowSel))) {
            const title = row.querySelector('div[columnindex="1"][role="gridcell"]')?.getAttribute('title') ?? '';
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

  // Reads all visible result rows in a single page.evaluate() call.
  async function scrapeResultRows(sourceId: string): Promise<Record<string, string>[]> {
    type ColDef = { index: number; header: string; child: string };
    const cols: ColDef[] = LOCATORS.columns.map((c) => ({ ...c }));

    const rawRows = await webCtx.page.evaluate(
      ({ rowSel, columns }: { rowSel: string; columns: ColDef[] }) => {
        const results: Record<string, string>[] = [];
        document.querySelectorAll(rowSel).forEach((rowEl) => {
          const record: Record<string, string> = {};
          columns.forEach(({ index, header, child }) => {
            const cell = rowEl.querySelector(`div[columnindex="${index}"][role="gridcell"]`);
            if (!cell) { record[header] = ''; return; }
            const title = (cell.getAttribute('title') ?? '').trim();
            record[header] = title !== '' ? title : (cell.querySelector(child)?.textContent ?? '').trim();
          });
          results.push(record);
        });
        return results;
      },
      { rowSel: LOCATORS.resultRow, columns: cols },
    );

    return rawRows
      .filter((r) => (r['Timesheet ID'] ?? '').trim().length > 3)
      .map((r) => ({ WOID: sourceId, ...r }));
  }

  // ── 1. Read IDs from input sheet ──────────────────────────────────────────────
  ctx.log(`Reading workbook: ${filePath}`);

  const inputWorkbook = new xl.Workbook();
  await inputWorkbook.xlsx.readFile(filePath);

  const inSheet = inputWorkbook.getWorksheet(inputSheetName);
  if (!inSheet) {
    const available = inputWorkbook.worksheets.map((ws: any) => ws.name).join(', ');
    throw new Error(`Input sheet "${inputSheetName}" not found. Available: ${available}`);
  }

  const headerRow  = inSheet.getRow(1);
  let   idColIndex = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell: any, colIdx: number) => {
    if (String(cell.value ?? '').trim().toLowerCase() === idColumnName.trim().toLowerCase()) {
      idColIndex = colIdx;
    }
  });

  if (idColIndex === -1) {
    const found: string[] = [];
    headerRow.eachCell({ includeEmpty: false }, (cell: any) => found.push(String(cell.value ?? '')));
    throw new Error(`Column "${idColumnName}" not found in "${inputSheetName}". Found: ${found.join(', ')}`);
  }

  const ids: string[] = [];
  inSheet.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    if (rowNumber === 1) return;
    const val = cellText(row.getCell(idColIndex).value).trim();
    if (val !== '') ids.push(val);
  });

  if (ids.length === 0) {
    throw new Error(`No IDs found in column "${idColumnName}" of sheet "${inputSheetName}".`);
  }
  ctx.log(`Found ${ids.length} ID(s) to process.`);

  // ── 2. Determine output column order + seed dataRowCounter ───────────────────
  const defaultColOrder: string[] = ['WOID', 'Worker Name', ...LOCATORS.columns.map((c) => c.header)];
  let colOrder: string[]          = defaultColOrder;
  let dataRowCounter              = 0;

  const existingSheet = inputWorkbook.getWorksheet(outputSheetName);
  if (existingSheet) {
    const hdrs: string[] = [];
    existingSheet.getRow(1).eachCell({ includeEmpty: false }, (cell: any) => {
      hdrs.push(String(cell.value ?? ''));
    });
    if (hdrs.length > 0) {
      colOrder       = hdrs;
      dataRowCounter = Math.max(0, existingSheet.rowCount - 1);
      ctx.log(`Output sheet "${outputSheetName}" exists — ${dataRowCounter} existing data row(s), column order locked.`);
    } else {
      ctx.log(`Output sheet "${outputSheetName}" exists but has no header — will write header on first flush.`);
    }
  } else {
    ctx.log(`Output sheet "${outputSheetName}" not found — will create it on first flush.`);
  }

  // ── 3. flushRows — reload workbook → append rows → save ──────────────────────
  async function flushRows(rows: Record<string, string>[]): Promise<void> {
    if (rows.length === 0) return;

    const start = Date.now();
    const wb    = new xl.Workbook();
    await wb.xlsx.readFile(filePath);

    let sheet = wb.getWorksheet(outputSheetName);
    if (!sheet) {
      sheet = wb.addWorksheet(outputSheetName);
      sheet.columns = colOrder.map((h: string) => ({ width: Math.min(40, Math.max(14, h.length + 4)) }));
    }

    const hasHeader = sheet.getRow(1).getCell(1).value != null && sheet.getRow(1).getCell(1).value !== '';
    if (!hasHeader) {
      const hRow = sheet.getRow(1);
      colOrder.forEach((h: string, idx: number) => { hRow.getCell(idx + 1).value = h; });
      styleHeaderRow(hRow, colOrder.length);
    }

    const nextRow = sheet.rowCount + 1;
    rows.forEach((record, idx) => {
      const row = sheet!.getRow(nextRow + idx);
      colOrder.forEach((h: string, colIdx: number) => { row.getCell(colIdx + 1).value = record[h] ?? ''; });
      styleDataRow(row, colOrder.length, dataRowCounter + idx);
      row.commit();
    });

    dataRowCounter += rows.length;
    await wb.xlsx.writeFile(filePath);
    ctx.log(`[FLUSH] ${rows.length} row(s) written in ${Date.now() - start}ms`);
  }

  // ── 4. Main loop ──────────────────────────────────────────────────────────────
  let totalWritten = 0;

  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    ctx.log(`[${i + 1}/${ids.length}] Processing ID: ${id}`);

    try {
      const searchInner = webCtx.page.locator(LOCATORS.searchInputHost).locator(LOCATORS.searchInputInner);
      await searchInner.waitFor({ state: 'visible' });
      await searchInner.click({ clickCount: 3 });
      await searchInner.fill(id);
      await searchInner.press('Enter');

      await webCtx.waitForVisible(LOCATORS.tabTimeExpense);

      const workerName = await webCtx.page.locator(LOCATORS.workerName).first()
        .textContent().then((t: string | null) => t?.trim() ?? '').catch(() => '');
      ctx.log(`  → Worker: "${workerName}"`);

      await webCtx.click(LOCATORS.tabTimeExpense);

      await webCtx.waitForVisible(LOCATORS.fromDateInput);
      await webCtx.clear(LOCATORS.fromDateInput);
      await webCtx.type(LOCATORS.fromDateInput, fromDate);
      await webCtx.pressKey('Tab');

      await webCtx.waitForVisible(LOCATORS.toDateInput);
      await webCtx.clear(LOCATORS.toDateInput);
      await webCtx.type(LOCATORS.toDateInput, toDate);
      await webCtx.pressKey('Tab');

      await webCtx.waitForVisible(LOCATORS.applyFilterBtn);
      await webCtx.click(LOCATORS.applyFilterBtn);

      const hasResults = await waitForGridReady();

      if (!hasResults) {
        ctx.log(`  → No results — writing sentinel row.`);
        await flushRows([{ WOID: id, 'Worker Name': workerName, Status: 'No Results' }]);
        totalWritten += 1;
      } else {
        const rows = await scrapeResultRows(id);
        rows.forEach((r) => { r['Worker Name'] = workerName; });
        ctx.log(`  → ${rows.length} row(s) scraped.`);
        await flushRows(rows);
        totalWritten += rows.length;
      }

      if (i < ids.length - 1) {
        await webCtx.click(LOCATORS.homeBtn);
        await webCtx.page.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible' });
        await new Promise((resolve) => setTimeout(resolve, 300));
      }

    } catch (err) {
      ctx.log(`  Error on "${id}": ${String(err)}`);
      await flushRows([{ WOID: id, error: String(err) }]);
      totalWritten += 1;
      try {
        await webCtx.click(LOCATORS.homeBtn);
        await webCtx.page.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible' });
      } catch { /* best-effort recovery */ }
    }
  }

  ctx.log(`Done. ${totalWritten} total row(s) written to "${outputSheetName}" in ${filePath.split('/').pop()}.`);
}
