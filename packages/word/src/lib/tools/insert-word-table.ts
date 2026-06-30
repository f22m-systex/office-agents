import type { AgentContext } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import { defineTool, toolError, toolSuccess } from "./types";

/* global Word */

export function createInsertTableTool(_ctx: AgentContext) {
  return defineTool({
    name: "insert_word_table",
    label: "Insert a Table",
    description:
      "Insert a table into the Word document with the specified rows, columns, and data.",
    parameters: Type.Object({
      rows: Type.Integer({
        description: "The number of rows in the table.",
        minimum: 1,
      }),
      columns: Type.Integer({
        description: "The number of columns in the table.",
        minimum: 1,
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
            "If omitted, the table is inserted relative to the document body, and insertLocation must be 'Start' or 'End'.",
        }),
      ),
      data: Type.Optional(
        Type.Array(Type.Array(Type.String()), {
          description:
            "Optional 2D array of strings containing data for cells. Row count and column count must match the table dimensions.",
        }),
      ),
      autoFit: Type.Optional(
        Type.Boolean({
          description:
            "If true, automatically adjusts the table columns to the width of the window.",
        }),
      ),
    }),
    execute: async (_toolCallId, params) => {
      try {
        await Word.run(async (context) => {
          let table: Word.Table;
          const values = params.data || [];

          // Construct rows x columns string 2D array
          const cellValues: string[][] = [];
          for (let r = 0; r < params.rows; r++) {
            const row: string[] = [];
            for (let c = 0; c < params.columns; c++) {
              row.push(values[r]?.[c] ?? "");
            }
            cellValues.push(row);
          }

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
            table = target.insertTable(
              params.rows,
              params.columns,
              params.insertLocation as "Before" | "After",
              cellValues,
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
            table = body.insertTable(
              params.rows,
              params.columns,
              params.insertLocation as "Start" | "End",
              cellValues,
            );
          }

          if (params.autoFit) {
            table.autoFitWindow();
          }

          await context.sync();
        });
        return toolSuccess({
          success: true,
          message: "Successfully inserted table.",
        });
      } catch (error) {
        return toolError(
          error instanceof Error ? error.message : "Error inserting table",
        );
      }
    },
  });
}
