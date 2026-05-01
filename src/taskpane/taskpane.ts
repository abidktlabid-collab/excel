/* global console, document, Excel, Office */
import { createAIClient, VALID_PROVIDERS, DEFAULT_MODELS, Provider, Message } from "../shared/ai-clients";
import * as ExcelEngine from "./excel-engine";
import * as AIParser from "./ai-parser";
import * as UI from "./ui-components";
import { ChatMessage } from "./ui-components";
import { SYSTEM_PROMPT } from "./system-prompt";

// ── State ──────────────────────────────────────────────────────────────────────
let conversationHistory: any[] = [];
let isStreaming = false;
let currentProvider: Provider = "openrouter";
let currentModel = DEFAULT_MODELS[currentProvider];
const _k = "c2stb3ItdjEtNTk1OTU4MzkyODgxMDg4MDcxYTFhOWY2ZGEwMTRmNWZmOWQ5MmJjYTMyZTQ2YmNmZjQ1NDMyOWE4N2FmNGUwNA==";
let currentApiKey = atob(_k);
let autoInsert = true;
let attachedCellData: string | null = null;

// ── DOM refs ───────────────────────────────────────────────────────────────────
let chatMessages: HTMLElement;
let welcomeEl: HTMLElement;
let textarea: HTMLTextAreaElement;
let btnSend: HTMLButtonElement;
let btnSettings: HTMLButtonElement;
let btnNewChat: HTMLButtonElement;
let btnReadSelection: HTMLElement;
let btnAttachSelection: HTMLElement;
let settingsPanel: HTMLElement;
let selProvider: HTMLSelectElement;
let inpModel: HTMLInputElement;
let inpApiKey: HTMLInputElement;
let chkAutoInsert: HTMLInputElement;
let statusDot: HTMLElement;
let statusText: HTMLElement;
let footerProvider: HTMLElement;
let sheetMonitor: HTMLElement;
let selModel: HTMLSelectElement;
let customModelGroup: HTMLElement;
let toast: HTMLElement;

// ── Init ───────────────────────────────────────────────────────────────────────
Office.onReady(() => {
  bindElements();
  bindEvents();

  const savedKey = localStorage.getItem("llm_excel_apikey");
  if (savedKey && savedKey.length > 10) currentApiKey = savedKey;

  selProvider.value = currentProvider;
  inpModel.value = currentModel;
  inpApiKey.value = currentApiKey;

  settingsPanel.classList.add("open");
  btnSettings.classList.add("active");

  updateFooterBadge();
  updateConnectionStatus();
  textarea.focus();
  syncModelUI();
});

function bindElements() {
  chatMessages       = document.getElementById("chat-messages")!;
  welcomeEl          = document.getElementById("welcome")!;
  textarea           = document.getElementById("chat-textarea") as HTMLTextAreaElement;
  btnSend            = document.getElementById("btn-send") as HTMLButtonElement;
  btnSettings        = document.getElementById("btn-settings") as HTMLButtonElement;
  btnNewChat         = document.getElementById("btn-new-chat") as HTMLButtonElement;
  btnReadSelection   = document.getElementById("btn-read-selection")!;
  btnAttachSelection = document.getElementById("btn-attach-selection")!;
  settingsPanel      = document.getElementById("settings-panel")!;
  selProvider        = document.getElementById("sel-provider") as HTMLSelectElement;
  inpModel           = document.getElementById("inp-model") as HTMLInputElement;
  inpApiKey          = document.getElementById("inp-apikey") as HTMLInputElement;
  chkAutoInsert      = document.getElementById("chk-auto-insert") as HTMLInputElement;
  statusDot          = document.getElementById("status-dot")!;
  statusText         = document.getElementById("status-text")!;
  footerProvider     = document.getElementById("footer-provider")!;
  toast              = document.getElementById("toast")!;
  sheetMonitor       = document.getElementById("sheet-monitor")!;
  selModel           = document.getElementById("sel-model") as HTMLSelectElement;
  customModelGroup   = document.getElementById("custom-model-group")!;
}

function bindEvents() {
  btnSettings.addEventListener("click", () => {
    settingsPanel.classList.toggle("open");
    btnSettings.classList.toggle("active");
  });

  btnNewChat.addEventListener("click", clearChat);
  btnReadSelection.addEventListener("click", readSelectionToChat);
  btnAttachSelection.addEventListener("click", attachSelection);

  selProvider.addEventListener("change", () => {
    currentProvider = selProvider.value as Provider;
    currentModel = DEFAULT_MODELS[currentProvider];
    syncModelUI();
    updateFooterBadge();
    updateConnectionStatus();
  });

  selModel.addEventListener("change", () => {
    if (selModel.value === "custom") {
      customModelGroup.style.display = "block";
      currentModel = inpModel.value.trim();
    } else {
      customModelGroup.style.display = "none";
      currentModel = selModel.value;
      inpModel.value = currentModel;
    }
    updateFooterBadge();
  });

  inpModel.addEventListener("input", () => {
    currentModel = inpModel.value.trim();
    updateFooterBadge();
  });

  inpApiKey.addEventListener("input", () => {
    currentApiKey = inpApiKey.value.trim();
    localStorage.setItem("llm_excel_apikey", currentApiKey);
    updateConnectionStatus();
  });

  chkAutoInsert.addEventListener("change", () => {
    autoInsert = chkAutoInsert.checked;
  });

  textarea.addEventListener("input", () => {
    UI.autoResize(textarea);
    const hasText = !!textarea.value.trim();
    const hasKey = currentApiKey && currentApiKey.length > 10;
    btnSend.disabled = !hasText || !hasKey || isStreaming;
  });

  textarea.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      if (textarea.value.trim() && !isStreaming) sendMessage();
    }
  });

  btnSend.addEventListener("click", sendMessage);

  document.querySelectorAll(".welcome__chip").forEach(chip => {
    chip.addEventListener("click", () => {
      textarea.value = chip.textContent!.trim();
      UI.autoResize(textarea);
      btnSend.disabled = false;
      sendMessage();
    });
  });
}

// ── Core logic ─────────────────────────────────────────────────────────────────
async function sendMessage() {
  let text = textarea.value.trim();
  if (!text || isStreaming) return;

  btnSend.disabled = true;
  const originalPlaceholder = textarea.placeholder;

  try {
    if (!currentApiKey) {
      showToast("⚠️ Enter an API key in Settings");
      settingsPanel.classList.add("open");
      btnSettings.classList.add("active");
      return;
    }

    const sheetContext = await getFullWorksheetContext();
    let fullPrompt = `[WORKSHEET CONTEXT]\n${sheetContext}\n\n`;
    if (attachedCellData) {
      fullPrompt += `📎 Selected/Attached Data:\n${attachedCellData}\n\n`;
      attachedCellData = null;
      textarea.placeholder = "Ask anything or try a quick action above…";
    }
    fullPrompt += `[USER QUERY]\n${text}`;

    if (welcomeEl) welcomeEl.style.display = "none";

    const userMsg: ChatMessage = { role: "user", content: text, timestamp: new Date() };
    conversationHistory.push({ role: "user", content: fullPrompt, timestamp: new Date() });
    
    UI.appendMessageBubble(chatMessages, userMsg);
    textarea.value = "";
    UI.autoResize(textarea);
    
    await streamAIResponse();
  } catch (err) {
    console.error("SendMessage Error:", err);
    showToast("⚠️ Message failed to send");
    textarea.placeholder = originalPlaceholder;
  } finally {
    if (!isStreaming) updateConnectionStatus();
  }
}

async function streamAIResponse(retryCount = 0) {
  isStreaming = true;
  const typingEl = UI.appendTypingIndicator(chatMessages);
  UI.scrollToBottom(chatMessages);

  const messages: Message[] = [{ role: "system", content: SYSTEM_PROMPT }];
  for (const msg of conversationHistory) {
    messages.push({ role: msg.role, content: msg.content });
  }

  try {
    const client = createAIClient(currentProvider, currentApiKey);
    const assistantMsg: ChatMessage = { role: "assistant", content: "", timestamp: new Date() };
    typingEl.remove();
    
    const { contentEl } = UI.appendMessageBubble(chatMessages, assistantMsg, true);

    await client.generateStreamingCompletion(messages, currentModel, (chunk: string) => {
      assistantMsg.content += chunk;
      contentEl.innerHTML = UI.formatContent(assistantMsg.content);
      contentEl.classList.add("streaming-cursor");
      UI.scrollToBottom(chatMessages);
    });

    contentEl.classList.remove("streaming-cursor");
    conversationHistory.push(assistantMsg);

    if (autoInsert) {
      await processAutoActions(assistantMsg.content);
    }
  } catch (error: any) {
    typingEl?.remove();
    console.error("Chat error:", error);
    showToast("⚠️ AI response failed");
  } finally {
    isStreaming = false;
    btnSend.disabled = !textarea.value.trim();
    UI.scrollToBottom(chatMessages);
  }
}

async function processAutoActions(text: string) {
  // Charts
  const chartType = AIParser.detectChartRequest(text);
  if (chartType) await ExcelEngine.createChartFromSelection(chartType);

  // Formulas
  const formulas = AIParser.extractFormulas(text);
  if (formulas.length > 0) await ExcelEngine.insertFormulas(formulas);

  // Bulk Updates
  const updates = AIParser.detectUpdateCells(text);
  for (const op of updates) await ExcelEngine.updateCells(op);

  // Validation
  const validations = AIParser.detectValidation(text);
  for (const v of validations) await ExcelEngine.applyValidation(v);

  // Sheets
  const sheets = AIParser.detectNewSheet(text);
  for (const s of sheets) await ExcelEngine.createSheet(s);

  // Pivot Tables
  const pivots = AIParser.detectPivotTable(text);
  for (const p of pivots) await ExcelEngine.createPivotTable(p);

  // Slicers
  const slicers = AIParser.detectSlicer(text);
  for (const s of slicers) await ExcelEngine.addSlicer(s);

  // Protection
  const protection = AIParser.detectProtection(text);
  for (const p of protection) await ExcelEngine.protectSheet(p);

  // Themes
  const themes = AIParser.detectTheme(text);
  for (const t of themes) await ExcelEngine.applyExecutiveTheme(t);

  // Excel Tables (Structured)
  const tables = AIParser.detectCreateTable(text);
  for (const tbl of tables) await ExcelEngine.createNamedTable(tbl);

  // Raw Tables (to be auto-converted)
  const tableData = AIParser.parseResponseToTable(text);
  if (tableData && tableData.length >= 2) await ExcelEngine.writeDataToEmptyArea(tableData);
}

// ── Helpers ────────────────────────────────────────────────────────────────────
async function getFullWorksheetContext(): Promise<string> {
  let contextStr = "";
  sheetMonitor.textContent = "Analyzing Schema...";
  sheetMonitor.classList.add("scanning");

  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const sheets = context.workbook.worksheets;
      activeSheet.load("name");
      sheets.load("items/name");
      const usedRange = activeSheet.getUsedRangeOrNullObject(true);
      usedRange.load("address, rowCount, columnCount");
      await context.sync();
      
      contextStr = `ACTIVE WORKBOOK SCHEMA:\n- Current Sheet: "${activeSheet.name}"\n`;
      if (!usedRange.isNullObject) {
        const headerRange = usedRange.getRow(0);
        headerRange.load("values");
        await context.sync();
        const headers = headerRange.values[0];
        contextStr += `- Data Bounds: ${usedRange.address}\n- Headers: ${headers.join(", ")}\n`;
        sheetMonitor.textContent = `${activeSheet.name} (${usedRange.rowCount} rows)`;
      }
    });
  } catch (e) { console.warn(e); }
  finally { sheetMonitor.classList.remove("scanning"); }
  return contextStr;
}

function updateConnectionStatus() {
  const hasKey = currentApiKey && currentApiKey.length > 10;
  if (hasKey) {
    statusDot.classList.add("connected");
    const providerLabel = selProvider.options[selProvider.selectedIndex]?.text || currentProvider;
    statusText.textContent = `Analyst Active — ${providerLabel}`;
    btnSend.disabled = false;
  } else {
    statusDot.classList.remove("connected");
    statusText.textContent = "API Key Missing — Enter in Settings";
    btnSend.disabled = true;
  }
}

function updateFooterBadge() {
  footerProvider.textContent = `${currentProvider} / ${currentModel}`;
}

function clearChat() {
  conversationHistory = [];
  chatMessages.innerHTML = "";
  welcomeEl.style.display = "block";
  showToast("🧹 Chat cleared");
}

function showToast(msg: string) {
  toast.textContent = msg;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 3000);
}

async function readSelectionToChat() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();
      const data = range.values.map(r => r.join(" | ")).join("\n");
      const userMsg: ChatMessage = { 
        role: "user", 
        content: `📎 Cell data:\n${data}\n\nAnalyze this.`, 
        timestamp: new Date() 
      };
      conversationHistory.push(userMsg);
      UI.appendMessageBubble(chatMessages, userMsg);
      await streamAIResponse();
    });
  } catch (e) { showToast("⚠️ Could not read selection"); }
}

async function attachSelection() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();
      attachedCellData = range.values.map(r => r.join(" | ")).join("\n");
      showToast("📎 Cells attached");
    });
  } catch (e) { showToast("⚠️ Could not attach"); }
}

function syncModelUI() {
  const options = Array.from(selModel.options).map(opt => opt.value);
  if (options.includes(currentModel)) {
    selModel.value = currentModel;
    customModelGroup.style.display = "none";
  } else {
    selModel.value = "custom";
    customModelGroup.style.display = "block";
    inpModel.value = currentModel;
  }
}
