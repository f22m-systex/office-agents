import type { AgentContext } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import { defineTool, toolError, toolSuccess } from "./types";

/* global Word */

export function createInsertParagraphTool(_ctx: AgentContext) {
  return defineTool({
    name: "insert_word_paragraph",
    label: "Insert a Styled Paragraph",
    description:
      "Insert a new paragraph into the document with specific text and a built-in style.",
    parameters: Type.Object({
      text: Type.String({
        description: "The text content of the new paragraph.",
      }),
      insertLocation: Type.Enum(
        {
          Start: "Start",
          End: "End",
          Before: "Before",
          After: "After",
        },
        {
          description:
            "Location relative to the target paragraph (when paragraphIndex is specified) or the document body.",
        },
      ),
      paragraphIndex: Type.Optional(
        Type.Integer({
          description:
            "The 0-based index of the target paragraph to insert relative to. " +
            "If omitted, the paragraph is inserted relative to the document body, and insertLocation must be 'Start' or 'End'.",
        }),
      ),
      style: Type.Optional(
        Type.Enum(
          {
            Normal: "Normal",
            Heading1: "Heading1",
            Heading2: "Heading2",
            Heading3: "Heading3",
            Title: "Title",
            Subtitle: "Subtitle",
            ListParagraph: "ListParagraph",
            Quote: "Quote",
            IntenseQuote: "IntenseQuote",
          },
          {
            description: "The built-in Word style name. Defaults to 'Normal'.",
          },
        ),
      ),
    }),
    execute: async (_toolCallId, params) => {
      try {
        await Word.run(async (context) => {
          let paragraph: Word.Paragraph;

          if (params.paragraphIndex !== undefined) {
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

            if (
              params.insertLocation !== "Before" &&
              params.insertLocation !== "After"
            ) {
              throw new Error(
                `Insert location '${params.insertLocation}' is not supported when paragraphIndex is specified. Use 'Before' or 'After'.`,
              );
            }

            const target = paragraphs.items[params.paragraphIndex];
            paragraph = target.insertParagraph(
              params.text,
              params.insertLocation as "Before" | "After",
            );
          } else {
            const body = context.document.body;
            if (
              params.insertLocation === "Before" ||
              params.insertLocation === "After"
            ) {
              throw new Error(
                `Insert location '${params.insertLocation}' is only supported when paragraphIndex is specified. Use 'Start' or 'End'.`,
              );
            }
            paragraph = body.insertParagraph(
              params.text,
              params.insertLocation as "Start" | "End",
            );
          }

          if (params.style) {
            paragraph.styleBuiltIn = params.style as Word.BuiltInStyleName;
          } else {
            paragraph.styleBuiltIn = "Normal";
          }

          await context.sync();
        });
        return toolSuccess({
          success: true,
          message: "Successfully inserted paragraph.",
        });
      } catch (error) {
        return toolError(
          error instanceof Error ? error.message : "Error inserting paragraph",
        );
      }
    },
  });
}
