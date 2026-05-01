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
  const cellUpdates = AIParser.parseCellUpdates(text);
  if (cellUpdates.length > 0) {
    for (const update of cellUpdates) {
      await ExcelService.updateCells(update.range, update.values);
    }
  }

  // Structural Changes
  const pivotOp = AIParser.parsePivotTableCommand(text);
  if (pivotOp) await ExcelService.applyStructure("PIVOT", pivotOp);

  const slicerOp = AIParser.parseSlicerCommand(text);
  if (slicerOp) await ExcelService.applyStructure("SLICER", slicerOp);

  const protectOp = AIParser.parseProtectSheetCommand(text);
  if (protectOp) await ExcelService.applyStructure("PROTECT", protectOp);
}
