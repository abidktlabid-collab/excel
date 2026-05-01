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

  const meta = document.createElement("div");
  meta.className = "message__meta";
  meta.textContent = formatTime(msg.timestamp);

  const contentEl = document.createElement("div");
  contentEl.className = "message__content";
  
  if (isTyping) {
    contentEl.classList.add("typing-indicator");
  } else {
    contentEl.innerHTML = formatContent(msg.content);
  }

  body.appendChild(contentEl);
  body.appendChild(meta);
  bubble.appendChild(avatar);
  bubble.appendChild(body);
  container.appendChild(bubble);
  
  return { bubble, contentEl };
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
  let html = escapeHtml(text);
  // Simple markdown-ish bold
  html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  // Simple code blocks
  html = html.replace(/```([\s\S]*?)```/g, '<pre class="code-block">$1</pre>');
  // Newlines
  html = html.replace(/\n/g, "<br>");
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
