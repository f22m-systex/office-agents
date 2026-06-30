import type { AgentContext } from "@office-agents/core";
import { createBashTool, createReadTool } from "@office-agents/core";
import { createExecuteOfficeJsTool } from "./execute-office-js";
import { getDocumentStructureTool } from "./get-document-structure";
import { getDocumentTextTool } from "./get-document-text";
import { createGetOoxmlTool } from "./get-ooxml";
import { createInsertParagraphTool } from "./insert-word-paragraph";
import { createInsertTableTool } from "./insert-word-table";
import { createSafeReplaceTextTool } from "./safe-replace-paragraph-text";
import { screenshotDocumentTool } from "./screenshot-document";

export function createWordTools(ctx: AgentContext) {
  return [
    // fs tools
    createReadTool(ctx),
    createBashTool(ctx),
    // Word read tools
    screenshotDocumentTool,
    getDocumentTextTool,
    getDocumentStructureTool,
    createGetOoxmlTool(ctx),
    // Word write tools
    createInsertParagraphTool(ctx),
    createSafeReplaceTextTool(ctx),
    createInsertTableTool(ctx),
    createExecuteOfficeJsTool(ctx),
  ];
}

export {
  defineTool,
  type ToolResult,
  toolError,
  toolImage,
  toolSuccess,
  toolText,
} from "./types";
export {
  createBashTool,
  createExecuteOfficeJsTool,
  createGetOoxmlTool,
  createInsertParagraphTool,
  createInsertTableTool,
  createReadTool,
  createSafeReplaceTextTool,
  getDocumentStructureTool,
  getDocumentTextTool,
  screenshotDocumentTool,
};
