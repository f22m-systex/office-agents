import type { ChatMessage } from "@office-agents/sdk";

/**
 * 將 ChatMessage[] 對話歷史轉換成格式良好的 Markdown 字串
 */
export function chatMessagesToMarkdown(
  messages: ChatMessage[],
  sessionName: string,
): string {
  let md = `# 對話紀錄: ${sessionName}\n\n`;

  for (const msg of messages) {
    const roleName = msg.role === "user" ? "使用者 (User)" : "助理 (Assistant)";
    md += `## ${roleName}\n\n`;

    for (const part of msg.parts) {
      if (part.type === "text") {
        md += `${part.text}\n\n`;
      } else if (part.type === "image") {
        md += `![上傳圖片](data:${part.mimeType};base64,${part.data})\n\n`;
      } else if (part.type === "thinking") {
        if (part.thinking && part.thinking.trim()) {
          md += `<details>\n<summary>思考過程 (Thinking Process)</summary>\n\n${part.thinking.trim()}\n</details>\n\n`;
        }
      } else if (part.type === "toolCall") {
        md += `### 工具呼叫: \`${part.name}\`\n\n`;
        md += `**參數:**\n\`\`\`json\n${JSON.stringify(part.args, null, 2)}\n\`\`\`\n\n`;

        if (part.status === "pending" || part.status === "running") {
          md += `*狀態: 執行中 (${part.status})...*\n\n`;
        } else if (part.status === "error") {
          md += `*狀態: 失敗 (Failed)*\n\n`;
          if (part.result) {
            md += `**錯誤訊息:**\n\`\`\`\n${part.result}\n\`\`\`\n\n`;
          }
        } else if (part.status === "complete") {
          md += `*狀態: 已完成 (Completed)*\n\n`;
          if (part.result) {
            md += `**執行結果:**\n\`\`\`\n${part.result}\n\`\`\`\n\n`;
          }
          if (part.images && part.images.length > 0) {
            md += `**結果圖片:**\n\n`;
            for (const img of part.images) {
              md += `![工具輸出圖片](data:${img.mimeType};base64,${img.data})\n\n`;
            }
          }
        }
      }
    }
    md += `---\n\n`;
  }
  return md;
}

/**
 * 將 ChatMessage[] 對話歷史轉換成格式良好的 JSON 字串
 */
export function chatMessagesToJson(
  messages: ChatMessage[],
  sessionName: string,
): string {
  const exportData = {
    sessionName,
    exportedAt: new Date().toISOString(),
    messages,
  };
  return JSON.stringify(exportData, null, 2);
}

/**
 * 觸發瀏覽器下載檔案
 */
export function downloadFile(
  content: string,
  filename: string,
  mimeType: string,
): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * 將文字複製到系統剪貼簿
 */
export async function copyToClipboard(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }

    // 針對舊瀏覽器或 Office webview 環境的 Fallback 方案
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    textarea.style.left = "-999999px";
    textarea.style.top = "-999999px";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    const success = document.execCommand("copy");
    document.body.removeChild(textarea);
    return success;
  } catch (err) {
    console.error("Failed to copy to clipboard:", err);
    return false;
  }
}

/**
 * 清理並格式化檔名，過濾不合法字元
 */
export function sanitizeFilename(name: string): string {
  return name
    .replace(/[\\/:*?"<>|]/g, "_") // 過濾 Windows 檔名不合法字元
    .replace(/\s+/g, "_") // 將空白替換為底線
    .trim();
}
