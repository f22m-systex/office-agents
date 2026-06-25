import type {
  AgentEvent,
  AgentMessage,
  AgentTool,
} from "@earendil-works/pi-agent-core";
import type {
  AssistantMessage,
  TextContent,
  UserMessage,
} from "@earendil-works/pi-ai";
import { generateId } from "../message-utils";
import type { ProviderConfig } from "../provider-config";

export class HermesAgent {
  state: { messages: AgentMessage[] };
  private listeners: Set<(e: AgentEvent) => void> = new Set();
  private abortController: AbortController | null = null;
  private config: ProviderConfig;
  private systemPrompt: string;
  private tools: AgentTool[];

  constructor(options: {
    config: ProviderConfig;
    systemPrompt: string;
    tools: AgentTool[];
    existingMessages: AgentMessage[];
  }) {
    this.config = options.config;
    // Apply namespace prefix to tools in the system prompt
    this.systemPrompt = options.systemPrompt.replace(
      /Use the read tool to load a skill's file/g,
      "Use the office_vfs_read tool to load a skill's file",
    );
    this.tools = options.tools;
    this.state = { messages: options.existingMessages || [] };
  }

  subscribe(listener: (e: AgentEvent) => void) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(event: AgentEvent) {
    for (const listener of this.listeners) {
      listener(event);
    }
  }

  abort() {
    if (this.abortController) {
      this.abortController.abort();
      this.abortController = null;
    }
  }

  async prompt(
    text: string,
    images?: { type: "image"; data: string; mimeType: string }[],
  ) {
    this.abort();
    this.abortController = new AbortController();

    const baseUrl = this.config.customBaseUrl || "http://localhost:8642/v1";
    const apiKey = this.config.apiKey || "change-me";

    // Prepare content for user message
    let userContent: any = text;
    if (images && images.length > 0) {
      userContent = [
        { type: "text", text },
        ...images.map((img) => ({
          type: "image_url",
          image_url: { url: `data:${img.mimeType};base64,${img.data}` },
        })),
      ];
    }

    const openaiMessages = [
      { role: "system", content: this.systemPrompt },
      ...this.state.messages.flatMap((m) => {
        if (m.role === "assistant") {
          const content = (m as AssistantMessage).content;
          const textParts =
            content
              .filter((p: any) => p.type === "text")
              .map((p: any) => p.text)
              .join("") || "";
          return [{ role: "assistant", content: textParts }];
        } else if (m.role === "user") {
          const content = (m as UserMessage).content;
          let userText = "";
          if (typeof content === "string") {
            userText = content;
          } else {
            userText = content
              .filter((p: any) => p.type === "text")
              .map((p: any) => p.text)
              .join("");
          }
          return [{ role: "user", content: userText }];
        }
        return [];
      }),
      { role: "user", content: userContent },
    ];

    try {
      const response = await fetch(`${baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model: this.config.model || "hermes-agent",
          messages: openaiMessages,
          stream: true,
        }),
        signal: this.abortController.signal,
      });

      if (!response.ok) {
        throw new Error(
          `Hermes API error: ${response.status} ${response.statusText}`,
        );
      }

      if (!response.body) {
        throw new Error("No response body");
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");

      const assistantMessage: AssistantMessage = {
        role: "assistant",
        content: [],
        timestamp: Date.now(),
        api: "openai-completions" as any,
        provider: "hermes",
        model: this.config.model || "hermes-agent",
        usage: {
          input: 0,
          output: 0,
          cacheRead: 0,
          cacheWrite: 0,
          totalTokens: 0,
          cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
        },
        stopReason: "stop",
      };

      this.emit({ type: "message_start", message: assistantMessage } as any);

      let buffer = "";
      let currentEvent: string | null = null;
      let currentToolCallId: string | null = null;
      let inReasoning = false;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() || "";

        for (const line of lines) {
          if (line.trim() === "") {
            currentEvent = null;
            continue;
          }
          if (line.startsWith("event: ")) {
            currentEvent = line.slice(7).trim();
          } else if (line.startsWith("data: ")) {
            const dataStr = line.slice(6);
            if (dataStr === "[DONE]") continue;

            try {
              const data = JSON.parse(dataStr);

              if (currentEvent === "hermes.tool.progress") {
                // Parse tool progress
                const toolCallId = data.id || data.toolCallId || generateId();
                if (data.status === "start" || data.status === "running") {
                  if (currentToolCallId !== toolCallId) {
                    currentToolCallId = toolCallId;

                    assistantMessage.content.push({
                      type: "toolCall",
                      id: toolCallId,
                      name: data.name || "hermes_tool",
                      arguments: data.args || {},
                    });
                    this.emit({
                      type: "message_update",
                      message: assistantMessage,
                    } as any);
                    this.emit({
                      type: "tool_execution_start",
                      toolCallId,
                    } as any);
                  }
                  if (data.partialResult) {
                    this.emit({
                      type: "tool_execution_update",
                      toolCallId,
                      partialResult: data.partialResult,
                    } as any);
                  }
                } else if (
                  data.status === "complete" ||
                  data.status === "error"
                ) {
                  this.emit({
                    type: "tool_execution_end",
                    toolCallId,
                    result: data.result || data.partialResult,
                    isError: data.status === "error",
                  } as any);
                  currentToolCallId = null;
                }
              } else {
                // Standard chat.completion.chunk
                if (data.choices && data.choices[0]?.delta) {
                  const delta = data.choices[0].delta;
                  let deltaContent = delta.content || "";
                  
                  if (delta.reasoning_content) {
                    if (!inReasoning) {
                      deltaContent = "<think>\n" + delta.reasoning_content;
                      inReasoning = true;
                    } else {
                      deltaContent = delta.reasoning_content;
                    }
                  } else if (inReasoning && delta.content) {
                    deltaContent = "\n</think>\n\n" + delta.content;
                    inReasoning = false;
                  }

                  if (deltaContent) {
                    let textPart = assistantMessage.content.find(
                      (p: any) => p.type === "text",
                    ) as TextContent;
                    if (!textPart) {
                      textPart = { type: "text", text: "" };
                      assistantMessage.content.push(textPart);
                    }
                    textPart.text += deltaContent;
                    this.emit({
                      type: "message_update",
                      message: assistantMessage,
                    } as any);
                  }
                }
              }
            } catch (e) {
              console.warn("Failed to parse SSE data", dataStr, e);
            }
          }
        }
      }

      if (inReasoning) {
        const textPart = assistantMessage.content.find(
          (p: any) => p.type === "text",
        ) as TextContent;
        if (textPart) {
          textPart.text += "\n</think>\n";
        }
      }

      this.emit({ type: "agent_end" } as any);

      this.state.messages.push({
        role: "user",
        content: text,
        timestamp: Date.now(),
      } as UserMessage);
      this.state.messages.push(assistantMessage);
    } catch (err: any) {
      if (err.name === "AbortError") {
        console.log("Hermes request aborted");
      } else {
        this.emit({ type: "agent_end" } as any);
        throw err;
      }
    } finally {
      this.abortController = null;
    }
  }
}
