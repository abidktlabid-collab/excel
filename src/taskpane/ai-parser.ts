export function detectChartRequest(text: string): string | null {
  if (text.includes("PIE_CHART")) return "Pie";
  if (text.includes("BAR_CHART")) return "ColumnClustered";
  if (text.includes("LINE_CHART")) return "Line";
  return null;
}

export function extractFormulas(text: string): { range: string; formula: string }[] {
  const results: { range: string; formula: string }[] = [];
  const lines = text.split("\n");
  for (const line of lines) {
    if (line.includes("APPLY_FORMULA:") || line.includes("FORMULA:")) {
      const match = line.match(/range=([A-Z0-9:]+),?\s*formula=(.+)/i);
      if (match) results.push({ range: match[1], formula: match[2].trim() });
    }
  }
  return results;
}

export function detectUpdateCells(text: string): { range: string; values: any[][] }[] {
  const results: { range: string; values: any[][] }[] = [];
  // More flexible regex to catch various AI formatting styles
  const regex = /UPDATE_CELLS:?\s*range=([A-Z0-9:]+),?\s*values=(\[\[.+\]\])/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    try {
      const values = JSON.parse(match[2]);
      results.push({ range: match[1], values });
    } catch (e) { console.warn("Parse error in UPDATE_CELLS", e); }
  }
  return results;
}

export function detectApplyFormula(text: string): { range: string; formula: string }[] {
  const results: { range: string; formula: string }[] = [];
  const regex = /APPLY_FORMULA:\s*range=([A-Z0-9:]+),\s*formula=(=.+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ range: match[1], formula: match[2] });
  }
  return results;
}

export function detectValidation(text: string): { range: string; type: string; values: string }[] {
  const results: { range: string; type: string; values: string }[] = [];
  const regex = /VALIDATION:\s*range=([A-Z0-9:]+),\s*type=(\w+),\s*values=(.+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ range: match[1], type: match[2], values: match[3] });
  }
  return results;
}

export function detectNewSheet(text: string): string[] {
  const results: string[] = [];
  const lines = text.split("\n");
  for (const line of lines) {
    if (line.includes("NEWSHEET:")) {
      const parts = line.split("NEWSHEET:");
      if (parts[1]) results.push(parts[1].trim());
    }
  }
  return results;
}

export function detectCopyData(text: string): { sourceSheet: string; sourceRange: string; targetSheet: string }[] {
  const results: { sourceSheet: string; sourceRange: string; targetSheet: string }[] = [];
  const regex = /COPYDATA:\s*sourceSheet=(.+),\s*sourceRange=([A-Z0-9:]+),\s*targetSheet=(.+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ sourceSheet: match[1].trim(), sourceRange: match[2], targetSheet: match[3].trim() });
  }
  return results;
}

export function detectConditionalFormatting(text: string): { range: string; type: string; color: string }[] {
  const results: { range: string; type: string; color: string }[] = [];
  const regex = /HIGHLIGHT:\s*range=([A-Z0-9:]+),\s*type=(\w+),\s*color=(#[A-F0-9]{6})/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ range: match[1], type: match[2], color: match[3] });
  }
  return results;
}

export function detectPivotTable(text: string): { sourceRange: string; targetSheet: string; tableName: string; rows: string[]; columns: string[]; values: string[] }[] {
  const results: any[] = [];
  const regex = /PIVOT_TABLE:\s*sourceRange=([A-Z0-9:]+),\s*targetSheet=(.+),\s*tableName=(.+),\s*rows=\[(.+)\],\s*columns=\[(.+)\],\s*values=\[(.+)\]/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({
      sourceRange: match[1],
      targetSheet: match[2].trim(),
      tableName: match[3].trim(),
      rows: match[4].split(",").map(s => s.trim()),
      columns: match[5].split(",").map(s => s.trim()),
      values: match[6].split(",").map(s => s.trim())
    });
  }
  return results;
}

export function detectSlicer(text: string): { pivotTable: string; fieldName: string; targetSheet: string }[] {
  const results: any[] = [];
  const regex = /ADD_SLICER:\s*pivotTable=(.+),\s*fieldName=(.+),\s*targetSheet=(.+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ pivotTable: match[1].trim(), fieldName: match[2].trim(), targetSheet: match[3].trim() });
  }
  return results;
}

export function detectProtection(text: string): { sheetName: string; password?: string }[] {
  const results: any[] = [];
  const regex = /PROTECT_SHEET:\s*sheetName=(.+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ sheetName: match[1].trim() });
  }
  return results;
}

export function detectTheme(text: string): { range: string; theme: "Modern" | "Classic" | "Dark" }[] {
  const results: any[] = [];
  const regex = /APPLY_THEME:\s*range=([A-Z0-9:]+),\s*theme=(\w+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ range: match[1], theme: match[2] as any });
  }
  return results;
}

export function detectCreateTable(text: string): { range: string; name: string }[] {
  const results: any[] = [];
  const regex = /CREATE_TABLE:\s*range=([A-Z0-9:]+),\s*name=(.+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    results.push({ range: match[1], name: match[2].trim() });
  }
  return results;
}

export function parseResponseToTable(text: string): string[][] | null {
  const lines = text.split("\n").filter(l => l.trim().startsWith("|"));
  if (lines.length < 2) return null;
  return lines.map(line => {
    return line.split("|")
      .filter((_, i, arr) => i > 0 && i < arr.length - 1)
      .map(cell => cell.trim());
  }).filter(row => row.length > 0);
}
