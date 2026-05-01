export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
}

export function appendMessageBubble(
  container: HTMLElement, 
  msg: ChatMessage, 
  isTyping = false
): { bubble: HTMLElement; contentEl: HTMLElement } {
  const bubble = document.createElement("div");
  bubble.className = `message message--${msg.role}`;

  const avatar = document.createElement("div");
  avatar.className = "message__avatar";
  avatar.textContent = msg.role === "assistant" ? "AI" : "U";

  const body = document.createElement("div");
  body.className = "message__body";

  const contentEl = document.createElement("div");
  contentEl.className = "message__content";
  
  if (isTyping) {
    contentEl.classList.add("typing-indicator");
  } else {
    contentEl.innerHTML = formatContent(msg.content);
  }

  const meta = document.createElement("div");
  meta.className = "message__meta";
  
  const hasTable = msg.content.includes("|") && msg.role === "assistant";

  meta.innerHTML = `
    <span class="message__time">${formatTime(msg.timestamp)}</span>
    ${hasTable && !isTyping ? `<button class="message__action insert-manual-btn">📥 Insert into Sheet</button>` : ""}
    <span class="execution-status"></span>
  `;

  body.appendChild(contentEl);
  body.appendChild(meta);
  bubble.appendChild(avatar);
  bubble.appendChild(body);
  container.appendChild(bubble);
  
  return { bubble, contentEl };
}

export function updateExecutionStatus(bubble: HTMLElement, status: "pending" | "success" | "error") {
  const statusEl = bubble.querySelector(".execution-status");
  if (!statusEl) return;
  
  if (status === "pending") {
    statusEl.innerHTML = " &nbsp; ⚡ Applying...";
    statusEl.style.color = "var(--warning)";
  } else if (status === "success") {
    statusEl.innerHTML = " &nbsp; ✅ Applied";
    statusEl.style.color = "var(--success)";
  } else {
    statusEl.innerHTML = " &nbsp; ⚠️ Check Sheet";
    statusEl.style.color = "var(--error)";
  }
}

export function appendTypingIndicator(container: HTMLElement): HTMLElement {
  const { bubble } = appendMessageBubble(container, { 
    role: "assistant", 
    content: "...", 
    timestamp: new Date() 
  }, true);
  return bubble;
}

export function formatContent(text: string): string {
  // ── Ghost Mode: Mask technical execution commands ──
  const lines = text.split("\n");
  const cleanedLines = lines.map(line => {
    const isTech = /\[EXECUTION\]|UPDATE_CELLS:|NEW_TABLE:|CREATE_TABLE:|PIVOT_TABLE:|ADD_SLICER:|PROTECT_SHEET:|APPLY_THEME:|HIGHLIGHT:|APPLY_FORMULA:/i.test(line);
    if (isTech) {
      return `<span class="tech-mask">${escapeHtml(line)}</span>`;
    }
    return line;
  });

  let html = cleanedLines.join("\n");
  
  // ── Professional Markdown Rendering ──
  
  // Headers (e.g. ### Header)
  html = html.replace(/^### (.*$)/gim, "<h3>$1</h3>");
  
  // Bold (e.g. **text**)
  html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  
  // Lists (e.g. * Item or - Item)
  html = html.replace(/^\* (.*$)/gim, "<li>$1</li>");
  html = html.replace(/^- (.*$)/gim, "<li>$1</li>");
  
  // Wrap li in ul
  html = html.replace(/(<li>.*<\/li>)/gs, "<ul>$1</ul>");
  
  // Code Blocks (e.g. ```code```)
  html = html.replace(/```([\s\S]*?)```/g, '<pre class="code-block">$1</pre>');
  
  // Inline Code (e.g. `code`)
  html = html.replace(/`(.*?)`/g, "<code>$1</code>");
  
  // Newlines (only if not already converted to block elements)
  html = html.replace(/\n/g, "<br>");
  
  // Clean up double br around block elements
  html = html.replace(/<\/h3><br>/g, "</h3>");
  html = html.replace(/<\/ul><br>/g, "</ul>");
  html = html.replace(/<\/pre><br>/g, "</pre>");

  return html;
}

export function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

export function formatTime(date: Date): string {
  return date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

export function scrollToBottom(container: HTMLElement) {
  requestAnimationFrame(() => { container.scrollTop = container.scrollHeight; });
}

export function autoResize(el: HTMLTextAreaElement) {
  el.style.height = "auto";
  el.style.height = Math.min(el.scrollHeight, 120) + "px";
}
export function showToast(message: string, type: "success" | "error" = "success") {
  const toast = document.getElementById("toast");
  if (!toast) return;
  
  toast.innerText = message;
  toast.className = `context-toast context-toast--${type} show`;
  
  setTimeout(() => {
    toast.classList.remove("show");
  }, 4000);
}
