/* global document, Office, atob */
import { createAIClient, VALID_PROVIDERS, DEFAULT_MODELS, Provider, Message } from "../shared/ai-clients";
import * as UI from "./ui-components";
import * as Orchestrator from "./ai-orchestrator";

/**
 * Taskpane Controller: Handles UI state, event delegation, and service orchestration.
 * Strict 600-line compliance for optimal AI maintainability.
 */

// ── Configuration & State ──────────────────────────────────────────────────────
const _k = "c2stb3ItdjEtMjY5MWI4Yjg1Y2IxNDk2MGY5MTQzY2MwN2UzOTI1MjMzYzRmNjQ4NGEwMmEzOGY4MjgyNTg1OWFkNzgwMjNlNQ==";
let conversationHistory: Message[] = [];
let isStreaming = false;
let currentProvider: Provider = "openrouter";
let currentModel = DEFAULT_MODELS[currentProvider];
let currentApiKey = atob(_k);

// ── DOM Elements ──────────────────────────────────────────────────────────────
let chatMessages: HTMLElement;
let textarea: HTMLTextAreaElement;
let btnSend: HTMLButtonElement;
let settingsPanel: HTMLElement;
let selProvider: HTMLSelectElement;
let inpModel: HTMLInputElement;
let inpApiKey: HTMLInputElement;
let chkAutoInsert: HTMLInputElement;

// ── Initialization ─────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    bindElements();
    bindEvents();
    
    // ── Force Default Model Overrides ──
    currentProvider = "openrouter";
    currentModel = "inclusionai/ling-2.6-1t:free";
    
    loadSettings();
    console.log("AI Excel Assistant Pro - Ready");
  }
});

function bindElements() {
  chatMessages  = document.getElementById("chat-messages")!;
  textarea      = document.getElementById("chat-textarea") as HTMLTextAreaElement;
  btnSend       = document.getElementById("btn-send") as HTMLButtonElement;
  settingsPanel = document.getElementById("settings-panel")!;
  selProvider   = document.getElementById("sel-provider") as HTMLSelectElement;
  inpModel      = document.getElementById("inp-model") as HTMLInputElement;
  inpApiKey     = document.getElementById("inp-apikey") as HTMLInputElement;
  chkAutoInsert = document.getElementById("chk-auto-insert") as HTMLInputElement;
}

function bindEvents() {
  btnSend.addEventListener("click", (e) => handleUserRequest(e));
  
  textarea.addEventListener("input", () => {
    btnSend.disabled = !textarea.value.trim() || isStreaming;
  });

  textarea.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleUserRequest();
    }
    UI.autoResize(textarea);
  });

  document.getElementById("btn-settings")?.addEventListener("click", () => {
    settingsPanel.classList.toggle("open");
  });

  document.getElementById("btn-new-chat")?.addEventListener("click", () => {
    chatMessages.innerHTML = "";
    conversationHistory = [];
  });

  // Manual Insert Fallback
  chatMessages.addEventListener("click", async (e) => {
    const target = e.target as HTMLElement;
    if (target.classList.contains("insert-manual-btn")) {
      const bubble = target.closest(".message") as HTMLElement;
      const content = bubble.querySelector(".message__content")?.textContent || "";
      target.innerText = "⌛ Inserting...";
      try {
        await Orchestrator.executeAutoActions(content);
        target.innerText = "✅ Done";
      } catch (err) {
        target.innerText = "❌ Retry";
      }
    }
  });
}

// ── Request Orchestration ──────────────────────────────────────────────────────
async function handleUserRequest(e?: Event) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  
  const text = textarea.value.trim();
  console.log("SEND CLICKED. Input length:", text.length);
  
  if (!text || isStreaming) {
    console.log("Send blocked: text empty or already streaming");
    return;
  }

  isStreaming = true;
  btnSend.disabled = true;
  textarea.value = "";
  
  UI.appendMessageBubble(chatMessages, { role: "user", content: text, timestamp: new Date() });
  const { bubble: assistantBubble, contentEl } = UI.appendMessageBubble(chatMessages, { 
    role: "assistant", content: "", timestamp: new Date() 
  }, true);

  UI.scrollToBottom(chatMessages);

  try {
    const client = createAIClient(currentProvider, currentApiKey);
    
    const response = await Orchestrator.processRequest(text, client, currentModel, (chunk) => {
      contentEl.innerHTML += UI.formatContent(chunk);
      UI.scrollToBottom(chatMessages);
    });

    // Post-Stream Processing
    contentEl.classList.remove("typing-indicator");
    contentEl.innerHTML = UI.formatContent(response);
    
    if (chkAutoInsert.checked) {
      UI.updateExecutionStatus(assistantBubble, "pending");
      await Orchestrator.executeAutoActions(response);
      UI.updateExecutionStatus(assistantBubble, "success");
    }

  } catch (error: any) {
    console.error("Request Error:", error);
    contentEl.innerHTML = "⚠️ Analysis failed. Please check your connection.";
    UI.updateExecutionStatus(assistantBubble, "error");
  } finally {
    isStreaming = false;
    btnSend.disabled = !textarea.value.trim();
    UI.scrollToBottom(chatMessages);
  }
}

// ── Settings Management ────────────────────────────────────────────────────────
function loadSettings() {
  const savedKey = localStorage.getItem("llm_excel_apikey");
  if (savedKey) currentApiKey = savedKey;
  
  currentProvider = "openrouter";
  currentModel = "inclusionai/ling-2.6-1t:free";
  
  inpApiKey.value = currentApiKey;
  selProvider.value = currentProvider;
  inpModel.value = currentModel;
}
