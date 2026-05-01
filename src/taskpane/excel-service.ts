/* global Excel */

/**
 * Excel Service: High-performance data bridge for LLM Assistant.
 * Handles reading, writing, and analytical context gathering.
 * Designed for modular stability in Excel Online.
 */

export interface WorksheetContext {
  activeSheet: string;
  range: string;
  data: any[][];
}

/**
 * ── ANALYZE: Get full worksheet data for AI context ──
 */
export async function getFullWorksheetContext(): Promise<WorksheetContext> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange(true);
    usedRange.load(["address", "values", "rowCount"]);
    await context.sync();

    return {
      activeSheet: sheet.name,
      range: usedRange.address,
      data: usedRange.values
    };
  });
}

/**
 * ── WRITE: Safe data insertion with physical range detection ──
 */
export async function writeData(data: any[][]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange(true);
    usedRange.load("rowCount");
    await context.sync();

    let startRow = 0;
    try {
      if (usedRange && usedRange.rowCount > 0) {
        startRow = usedRange.rowCount + 2;
      }
    } catch (e) {
      startRow = 0;
    }

    const targetRange = sheet.getRangeByIndexes(startRow, 0, data.length, data[0].length);
    targetRange.values = data;
    targetRange.format.autofitColumns();
    await context.sync();

    // Professional Table Conversion
    try {
      const table = sheet.tables.add(targetRange, true);
      table.style = "TableStyleMedium2";
      await context.sync();
    } catch (e) {
      console.warn("Table styling skipped, data remains intact.");
    }
  });
}

/**
 * ── EDIT: Surgical cell and formula updates ──
 */
export async function updateCells(rangeAddr: string, values: any[][]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddr);
    range.values = values;
    await context.sync();
  });
}

/**
 * ── ANALYSIS: Create charts and visual tools ──
 */
export async function createChart(type: string): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    context.workbook.worksheets.getActiveWorksheet().charts.add(type as any, range, "Auto" as any);
    await context.sync();
  });
}

/**
 * ── STRUCTURE: Slicers, PivotTables, and Protection ──
 */
export async function applyStructure(type: "PIVOT" | "SLICER" | "PROTECT", op: any): Promise<void> {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    if (type === "PIVOT") {
      const source = workbook.worksheets.getActiveWorksheet().getRange(op.sourceRange);
      let target = workbook.worksheets.getItemOrNullObject(op.targetSheet);
      await context.sync();
      if (target.isNullObject) target = workbook.worksheets.add(op.targetSheet);
      const pivot = target.pivotTables.add(op.tableName, source, "A3");
      for (const r of op.rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(r));
      for (const v of op.values) pivot.dataHierarchies.add(pivot.hierarchies.getItem(v));
    } else if (type === "SLICER") {
      const sheet = workbook.worksheets.getItem(op.targetSheet);
      const pivot = workbook.pivotTables.getItem(op.pivotTable);
      sheet.slicers.add(pivot, op.fieldName);
    } else if (type === "PROTECT") {
      workbook.worksheets.getItem(op.sheetName).protection.protect();
    }
    await context.sync();
  });
}
