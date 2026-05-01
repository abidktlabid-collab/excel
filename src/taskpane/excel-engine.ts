/* global Excel */

export async function createChartFromSelection(type: string) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const chart = context.workbook.worksheets.getActiveWorksheet().charts.add(
      type as Excel.ChartType,
      range,
      Excel.ChartSeriesBy.auto
    );
    chart.title.text = "AI Generated Analysis";
    await context.sync();
  });
}

export async function insertFormulas(formulas: { range: string; formula: string }[]) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    for (const f of formulas) {
      const range = sheet.getRange(f.range);
      range.values = [[f.formula]];
    }
    await context.sync();
  });
}

export async function updateCells(op: { range: string; values: any[][] }) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(op.range);
    range.values = op.values;
    await context.sync();
  });
}

export async function applyValidation(v: { range: string; type: string; values: string }) {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(v.range);
    const validation = range.dataValidation;
    validation.clear();
    if (v.type === "list") {
      validation.rule = { list: { inCellDropDown: true, source: v.values } };
    }
    await context.sync();
  });
}

export async function createSheet(name: string) {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const newSheet = sheets.add(name);
    newSheet.activate();
    await context.sync();
  });
}

export async function applyConditionalFormatting(cf: { range: string; type: string; color: string }) {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cf.range);
    const format = range.conditionalFormats.add(Excel.ConditionalFormatType.topItems);
    format.topItems.format.fill.color = cf.color;
    format.topItems.rule = { rank: 10, selectionType: Excel.ConditionalTopBottomSelectionType.topItems };
    await context.sync();
  });
}

export async function copyDataBetweenSheets(op: { sourceSheet: string; sourceRange: string; targetSheet: string }) {
  await Excel.run(async (context) => {
    const sourceRange = context.workbook.worksheets.getItem(op.sourceSheet).getRange(op.sourceRange);
    const targetSheet = context.workbook.worksheets.getItem(op.targetSheet);
    const targetRange = targetSheet.getRange(op.sourceRange);
    targetRange.copyFrom(sourceRange);
    await context.sync();
  });
}

export async function createPivotTable(op: { sourceRange: string; targetSheet: string; tableName: string; rows: string[]; columns: string[]; values: string[] }) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const sourceRange = workbook.worksheets.getActiveWorksheet().getRange(op.sourceRange);
    let targetSheet: Excel.Worksheet;
    
    try {
      targetSheet = workbook.worksheets.getItem(op.targetSheet);
    } catch (e) {
      targetSheet = workbook.worksheets.add(op.targetSheet);
    }
    
    const pivotTable = targetSheet.pivotTables.add(op.tableName, sourceRange, targetSheet.getRange("A3"));
    
    // Add Row Fields
    for (const r of op.rows) pivotTable.rowHierarchy.add(pivotTable.hierarchies.getItem(r));
    // Add Column Fields
    for (const c of op.columns) pivotTable.columnHierarchy.add(pivotTable.hierarchies.getItem(c));
    // Add Data Fields
    for (const v of op.values) {
      const field = pivotTable.hierarchies.getItem(v);
      pivotTable.dataHierarchy.add(field);
    }
    
    targetSheet.activate();
    await context.sync();
  });
}

export async function addSlicer(op: { pivotTable: string; fieldName: string; targetSheet: string }) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(op.targetSheet);
    const pivot = sheet.pivotTables.getItem(op.pivotTable);
    sheet.slicers.add(pivot, op.fieldName, sheet.getRange("G3"));
    await context.sync();
  });
}

export async function protectSheet(op: { sheetName: string; password?: string }) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(op.sheetName);
    sheet.protection.protect({
      allowInsertRows: false,
      allowDeleteRows: false,
      selectionMode: Excel.ProtectionSelectionMode.unlocked
    }, op.password);
    await context.sync();
  });
}

export async function applyExecutiveTheme(op: { range: string; theme: "Modern" | "Classic" | "Dark" }) {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(op.range);
    range.format.fill.clear();
    range.format.font.name = "Segoe UI";
    
    if (op.theme === "Modern") {
      range.format.fill.color = "#F3F4F6";
      range.format.font.color = "#111827";
      range.getRow(0).format.fill.color = "#2563EB";
      range.getRow(0).format.font.color = "white";
    } else if (op.theme === "Dark") {
      range.format.fill.color = "#1F2937";
      range.format.font.color = "#F9FAFB";
      range.getRow(0).format.fill.color = "#111827";
    }
    
    range.format.autofitColumns();
    await context.sync();
  });
}

export async function writeDataToEmptyArea(data: any[][]) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // 1. Find the next safe spot
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("rowCount");
    await context.sync();

    let startRow = 0;
    if (!usedRange.isNullObject) {
      startRow = usedRange.rowCount + 2; // Add a 2-row buffer
    }

    const targetRange = sheet.getRangeByIndexes(startRow, 0, data.length, data[0].length);
    
    // 2. RAW INSERT (The most important part - ensuring data exists)
    targetRange.values = data;
    targetRange.format.autofitColumns();
    await context.sync();
    
    // 3. OPTIONAL STYLING (Try to make it a table, but don't fail if we can't)
    try {
      const table = sheet.tables.add(targetRange, true);
      table.name = `AI_Table_${Date.now()}`;
      table.style = "TableStyleMedium2";
      await context.sync();
    } catch (e) {
      console.warn("Table conversion failed, but data was inserted successfully.", e);
    }
  });
}

export async function createNamedTable(op: { range: string; name: string }) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(op.range);
    const table = sheet.tables.add(range, true);
    table.name = op.name.replace(/\s+/g, "_");
    table.style = "TableStyleMedium9";
    await context.sync();
  });
}
