/* global  console, CustomFunctions  */
import { createAIClient, VALID_PROVIDERS, Provider, Message } from "../shared/ai-clients";

/**
 * Generates a response based on the given prompt using the specified AI model and provider.
 * @customfunction PROMPT
 * @helpUrl https://llmexcel.liminity.se/help
 * @param message The prompt message to send to the AI.
 * @param model The AI model to use for generating the response.
 * @param apiKey The API key for the AI service.
 * @param systemPrompt An optional system prompt to provide context for the AI.
 * @param provider The AI provider to use ("openai", "anthropic", "openrouter", "mistral", "groq", or "gemini").
 * @returns A promise that resolves to the generated response.
 */
export async function prompt(
  message: string,
  model: string,
  apiKey: string,
  systemPrompt: string,
  provider: string
): Promise<string> {
  try {
    if (!message || !model || !apiKey || !provider) {
      throw new Error("Missing required parameters");
    }

    const normalizedProvider = provider.toLowerCase() as Provider;
    if (!VALID_PROVIDERS.includes(normalizedProvider)) {
      throw new Error(`Invalid provider. Must be one of: ${VALID_PROVIDERS.join(", ")}`);
    }

    const client = createAIClient(normalizedProvider, apiKey);
    const messages: Message[] = systemPrompt
      ? [
          { role: "system", content: systemPrompt },
          { role: "user", content: message },
        ]
      : [{ role: "user", content: message }];

    const response = await client.generateCompletion(messages, model);
    if (!response) {
      throw new Error("Empty response from AI provider");
    }
    return response;
  } catch (error) {
    console.error("Error in prompt function:", error);
    if (error instanceof Error) {
      return `Error: ${error.message}`;
    } else {
      return "An unexpected error occurred";
    }
  }
}

/**
 * Generates a streaming response based on the given prompt using the specified AI model and provider.
 * @customfunction PROMPT_STREAM
 * @streaming
 * @helpUrl https://llmexcel.liminity.se/help
 * @param message The prompt message to send to the AI.
 * @param model The AI model to use for generating the response.
 * @param apiKey The API key for the AI service.
 * @param systemPrompt An optional system prompt to provide context for the AI.
 * @param provider The AI provider to use ("openai", "anthropic", "openrouter", "mistral", "groq", or "gemini").
 * @param invocation The streaming invocation object
 */
export function promptStream(
  message: string,
  model: string,
  apiKey: string,
  systemPrompt: string,
  provider: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  try {
    if (!message || !model || !apiKey || !provider) {
      throw new Error("Missing required parameters");
    }

    const normalizedProvider = provider.toLowerCase() as Provider;
    if (!VALID_PROVIDERS.includes(normalizedProvider)) {
      throw new Error(`Invalid provider. Must be one of: ${VALID_PROVIDERS.join(", ")}`);
    }

    const client = createAIClient(normalizedProvider, apiKey);
    let fullResponse = "";

    const messages: Message[] = systemPrompt
      ? [
          { role: "system", content: systemPrompt },
          { role: "user", content: message },
        ]
      : [{ role: "user", content: message }];

    client
      .generateStreamingCompletion(messages, model, (chunk) => {
        fullResponse += chunk;
        invocation.setResult(fullResponse);
      })
      .then(() => {
        if (!fullResponse) {
          throw new Error("Empty response from AI provider");
        }
      })
      .catch((error) => {
        console.error("Error in promptStream function:", error);
        if (error instanceof Error) {
          invocation.setResult(`Error: ${error.message}`);
        } else {
          invocation.setResult("An unexpected error occurred");
        }
      });

    invocation.onCanceled = () => {
      console.log("Stream cancelled by user");
    };
  } catch (error) {
    console.error("Error in promptStream function setup:", error);
    if (error instanceof Error) {
      invocation.setResult(`Error: ${error.message}`);
    } else {
      invocation.setResult("An unexpected error occurred");
    }
  }
}
