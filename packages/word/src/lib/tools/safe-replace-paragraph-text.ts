import type { AgentContext } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import { defineTool, toolError, toolSuccess } from "./types";

/* global Word */

export function createSafeReplaceTextTool(_ctx: AgentContext) {
  return defineTool({
    name: "safe_replace_paragraph_text",
    label: "Safely Replace Text Preserving Formatting",
    description:
      "Replace the text of an existing paragraph while completely preserving its original style, font family, size, color, and bold/italic states.",
    parameters: Type.Object({
      paragraphIndex: Type.Integer({
        description:
          "The 0-based index of the paragraph to modify (obtained from get_document_text).",
      }),
      newText: Type.String({ description: "The replacement text." }),
    }),
    execute: async (_toolCallId, params) => {
      try {
        await Word.run(async (context) => {
          const paragraphs = context.document.body.paragraphs;
          paragraphs.load("items");
          await context.sync();

          if (
            params.paragraphIndex < 0 ||
            params.paragraphIndex >= paragraphs.items.length
          ) {
            throw new Error(
              `Paragraph index ${params.paragraphIndex} is out of bounds (0 to ${paragraphs.items.length - 1}).`,
            );
          }

          const target = paragraphs.items[params.paragraphIndex];

          // Load original style and font properties
          target.font.load("name,size,color,bold,italic,underline");
          target.load("style");
          await context.sync();

          const origFont = target.font.name;
          const origSize = target.font.size;
          const origColor = target.font.color;
          const origBold = target.font.bold;
          const origItalic = target.font.italic;
          const origUnderline = target.font.underline;
          const origStyle = target.style;

          target.clear();
          target.insertText(params.newText, "Start");

          // Restore original style
          if (origStyle) {
            target.style = origStyle;
          }

          // Restore original font attributes if they are not null, undefined, or mixed
          if (origFont && origFont !== "Mixed") {
            target.font.name = origFont;
          }
          if (origSize && typeof origSize === "number") {
            target.font.size = origSize;
          }
          if (
            origColor &&
            origColor !== "Mixed" &&
            origColor !== "Automatic" &&
            origColor !== "#000000"
          ) {
            target.font.color = origColor;
          }
          if (
            origBold !== null &&
            origBold !== undefined &&
            typeof origBold === "boolean"
          ) {
            target.font.bold = origBold;
          }
          if (
            origItalic !== null &&
            origItalic !== undefined &&
            typeof origItalic === "boolean"
          ) {
            target.font.italic = origItalic;
          }
          if (origUnderline && origUnderline !== "Mixed") {
            target.font.underline = origUnderline;
          }

          await context.sync();
        });
        return toolSuccess({
          success: true,
          message: "Successfully replaced text while preserving formatting.",
        });
      } catch (error) {
        return toolError(
          error instanceof Error ? error.message : "Error replacing text",
        );
      }
    },
  });
}
