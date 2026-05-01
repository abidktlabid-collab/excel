/* global console */
import { OpenAI } from "openai";
import Anthropic from "@anthropic-ai/sdk";
import { MessageParam } from "@anthropic-ai/sdk/resources";
import { GoogleGenerativeAI } from "@google/generative-ai";

export type Provider = "openai" | "anthropic" | "openrouter" | "mistral" | "groq" | "gemini";
export type Message = { role: "user" | "system" | "assistant"; content: string };

export interface AIClient {
  generateCompletion(messages: Message[], model: string): Promise<string>;
  generateStreamingCompletion(messages: Message[], model: string, onChunk: (chunk: string) => void): Promise<void>;
}

// ── OpenAI ────────────────────────────────────────────────────────────────────
class OpenAIClient implements AIClient {
  private client: OpenAI;

  constructor(apiKey: string) {
    this.client = new OpenAI({ apiKey, dangerouslyAllowBrowser: true });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    const response = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
    });
    return response.choices[0].message.content || "No content in response";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    const stream = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
      stream: true,
    });
    for await (const chunk of stream) {
      const content = chunk.choices[0]?.delta?.content || "";
      onChunk(content);
    }
  }
}

// ── Anthropic ─────────────────────────────────────────────────────────────────
class AnthropicClient implements AIClient {
  private client: Anthropic;

  constructor(apiKey: string) {
    this.client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    const systemMessage = messages.find((msg) => msg.role === "system");
    const chatMessages = messages.filter((msg) => msg.role !== "system");

    const response = await this.client.messages.create({
      messages: chatMessages as MessageParam[],
      model,
      max_tokens: 4096,
      system: systemMessage?.content,
    });
    return response.content[0].type === "text" ? response.content[0].text : "";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    const systemMessage = messages.find((msg) => msg.role === "system");
    const chatMessages = messages.filter((msg) => msg.role !== "system");

    const stream = await this.client.messages.create({
      messages: chatMessages as MessageParam[],
      model,
      max_tokens: 4096,
      stream: true,
      system: systemMessage?.content,
    });
    for await (const chunk of stream) {
      if (chunk.type === "content_block_delta" && chunk.delta.type === "text_delta") {
        onChunk(chunk.delta.text);
      }
    }
  }
}

// ── OpenRouter ────────────────────────────────────────────────────────────────
class OpenRouterClient implements AIClient {
  private client: OpenAI;

  constructor(apiKey: string) {
    this.client = new OpenAI({
      apiKey,
      baseURL: "https://openrouter.ai/api/v1",
      dangerouslyAllowBrowser: true,
      defaultHeaders: {
        "HTTP-Referer": "https://excelco.netlify.app",
        "X-Title": "AI Excel Assistant Pro - excelco",
      },
    });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    const response = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
    });
    return response.choices[0].message.content || "No content in response";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    const stream = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
      stream: true,
      max_tokens: 8192,
    });
    for await (const chunk of stream) {
      const content = chunk.choices[0]?.delta?.content || "";
      onChunk(content);
    }
  }
}

// ── Mistral ───────────────────────────────────────────────────────────────────
class MistralClient implements AIClient {
  private client: OpenAI;

  constructor(apiKey: string) {
    this.client = new OpenAI({
      apiKey,
      baseURL: "https://api.mistral.ai/v1",
      dangerouslyAllowBrowser: true,
    });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    const response = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
    });
    return response.choices[0].message.content || "No content in response";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    const stream = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
      stream: true,
    });
    for await (const chunk of stream) {
      const content = chunk.choices[0]?.delta?.content || "";
      onChunk(content);
    }
  }
}

// ── Groq ──────────────────────────────────────────────────────────────────────
class GroqClient implements AIClient {
  private client: OpenAI;

  constructor(apiKey: string) {
    this.client = new OpenAI({
      apiKey,
      baseURL: "https://api.groq.com/openai/v1",
      dangerouslyAllowBrowser: true,
    });
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    const response = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
    });
    return response.choices[0].message.content || "No content in response";
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    const stream = await this.client.chat.completions.create({
      messages: messages as OpenAI.ChatCompletionMessageParam[],
      model,
      stream: true,
    });
    for await (const chunk of stream) {
      const content = chunk.choices[0]?.delta?.content || "";
      onChunk(content);
    }
  }
}

// ── Google Gemini ─────────────────────────────────────────────────────────────
class GeminiClient implements AIClient {
  private genAI: GoogleGenerativeAI;

  constructor(apiKey: string) {
    this.genAI = new GoogleGenerativeAI(apiKey);
  }

  private buildPrompt(messages: Message[]): { systemInstruction?: string; userPrompt: string } {
    const systemMessage = messages.find((m) => m.role === "system");
    const userMessages = messages.filter((m) => m.role === "user");
    return {
      systemInstruction: systemMessage?.content,
      userPrompt: userMessages.map((m) => m.content).join("\n"),
    };
  }

  async generateCompletion(messages: Message[], model: string): Promise<string> {
    const { systemInstruction, userPrompt } = this.buildPrompt(messages);
    const genModel = this.genAI.getGenerativeModel({
      model,
      ...(systemInstruction ? { systemInstruction } : {}),
    });
    const result = await genModel.generateContent(userPrompt);
    return result.response.text();
  }

  async generateStreamingCompletion(
    messages: Message[],
    model: string,
    onChunk: (chunk: string) => void
  ): Promise<void> {
    const { systemInstruction, userPrompt } = this.buildPrompt(messages);
    const genModel = this.genAI.getGenerativeModel({
      model,
      ...(systemInstruction ? { systemInstruction } : {}),
    });
    const result = await genModel.generateContentStream(userPrompt);
    for await (const chunk of result.stream) {
      onChunk(chunk.text());
    }
  }
}

// ── Factory ───────────────────────────────────────────────────────────────────
export function createAIClient(provider: Provider, apiKey: string): AIClient {
  switch (provider) {
    case "openai":
      return new OpenAIClient(apiKey);
    case "anthropic":
      return new AnthropicClient(apiKey);
    case "openrouter":
      return new OpenRouterClient(apiKey);
    case "mistral":
      return new MistralClient(apiKey);
    case "groq":
      return new GroqClient(apiKey);
    case "gemini":
      return new GeminiClient(apiKey);
    default:
      throw new Error(`Unsupported provider: ${provider}`);
  }
}

export const VALID_PROVIDERS: Provider[] = ["openai", "anthropic", "openrouter", "mistral", "groq", "gemini"];

export const DEFAULT_MODELS: Record<Provider, string> = {
  openai: "gpt-4o",
  anthropic: "anthropic/claude-haiku-4.5",
  openrouter: "inclusionai/ling-2.6-1t:free",
  mistral: "mistral-large-latest",
  groq: "llama-3.3-70b-versatile",
  gemini: "google/gemini-3.1-pro-preview",
};
