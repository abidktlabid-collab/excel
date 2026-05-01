import { Message } from "../shared/ai-clients";
import * as AIParser from "./ai-parser";
import * as ExcelService from "./excel-service";
import { SYSTEM_PROMPT } from "./system-prompt";

/**
 * AI Orchestrator: Manages the bridge between LLM intelligence and Excel actions.
 * Handles parsing, streaming logic, and command dispatching.
 */

export async function processRequest(
  text: string, 
  client: any, 
  model: string, 
  onChunk: (chunk: string) => void
): Promise<string> {
  
  // 1. ANALYZE: Gather deep context from the sheet
  const context = await ExcelService.getFullWorksheetContext();
  
  const messages: Message[] = [
    { role: "system", content: SYSTEM_PROMPT },
    { role: "user", content: `[WORKSHEET CONTEXT]\nSheet: ${context.activeSheet}\nRange: ${context.range}\nData: ${JSON.stringify(context.data)}\n\nUser Request: ${text}` }
  ];

  let fullResponse = "";
  
  // 2. THINK: Stream the response
  await client.generateStreamingCompletion(messages, model, (chunk: string) => {
    fullResponse += chunk;
    onChunk(chunk);
  });

  return fullResponse;
}

/**
 * ── EXECUTE: Dispatch parsed commands to the Excel Service ──
 */
export async function executeAutoActions(text: string): Promise<void> {
  // Data Tables
  const tableData = AIParser.parseResponseToTable(text);
  if (tableData) {
    await ExcelService.writeData(tableData);
  }

  // Charts
  const chartType = AIParser.detectChartRequest(text);
  if (chartType) {
    await ExcelService.createChart(chartType);
  }

  // Surgical Updates (Cells/Formulas)
  const cellUpdates = AIParser.detectUpdateCells(text);
  if (cellUpdates.length > 0) {
    for (const update of cellUpdates) {
      await ExcelService.updateCells(update.range, update.values);
    }
  }

  // Structural Changes
  const pivotOps = AIParser.detectPivotTable(text);
  if (pivotOps.length > 0) {
    for (const op of pivotOps) await ExcelService.applyStructure("PIVOT", op);
  }

  const slicerOps = AIParser.detectSlicer(text);
  if (slicerOps.length > 0) {
    for (const op of slicerOps) await ExcelService.applyStructure("SLICER", op);
  }

  const protectOps = AIParser.detectProtection(text);
  if (protectOps.length > 0) {
    for (const op of protectOps) await ExcelService.applyStructure("PROTECT", op);
  }
}
