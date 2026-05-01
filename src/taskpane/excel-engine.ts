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

export async function writeDataToEmptyArea(data: any[][]) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("address, rowCount");
    await context.sync();

    let startRow = 0;
    if (!usedRange.isNullObject) {
      startRow = usedRange.rowCount + 2;
    }

    const targetRange = sheet.getRangeByIndexes(startRow, 0, data.length, data[0].length);
    targetRange.values = data;
    targetRange.format.autofitColumns();
    await context.sync();
  });
}
