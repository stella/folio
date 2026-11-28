import { describe, expect, test } from "bun:test";

import type { Paragraph, Table } from "../../model/document";
import { resolveTableMoves } from "../tableCellMove";

const paragraph = (paraId: string, text = ""): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const insertedParagraph = (paraId: string, id: number): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [],
  pPrMark: {
    kind: "ins",
    info: { id, author: "Reviewer", date: "2026-11-28T00:00:00Z" },
  },
});

describe("table cell moves", () => {
  for (const nestedDestinationSurvives of [false, true]) {
    test(`empty move boundaries follow a ${nestedDestinationSurvives ? "surviving" : "removed"} nested destination`, () => {
      const owner = paragraph("00000100", "Owner");
      const leading = insertedParagraph("00000201", 11);
      const trailing = insertedParagraph("00000202", 12);
      const nested: Table = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [paragraph("00000301", "Nested")],
              },
            ],
          },
        ],
      };
      const table = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [
                  owner,
                  leading,
                  ...(nestedDestinationSurvives ? [nested] : []),
                  trailing,
                ],
              },
            ],
          },
        ],
        formatting: {
          gridChange: {
            id: 1,
            columnWidths: [1200],
            geometryChange: {
              type: "mergeCells",
              revisionIds: [1],
              moves: [
                {
                  kind: "empty",
                  source: "00000211",
                  destination: "00000201",
                  destinationMark: 11,
                },
                {
                  kind: "empty",
                  source: "00000212",
                  destination: "00000202",
                  destinationMark: 12,
                },
              ],
              nestedMoves: [
                {
                  rows: [20],
                  paragraphs: [["00000311", "00000301"]],
                },
              ],
            },
          },
        },
      } satisfies Table;

      const resolved = resolveTableMoves({
        table,
        reject: false,
        handoffs: { revisions: new Map(), controls: new Map() },
      });
      expect(resolved.rows[0]?.cells[0]?.content).toEqual(
        nestedDestinationSurvives
          ? [
              paragraph("00000100", "Owner"),
              paragraph("00000211"),
              nested,
              paragraph("00000212"),
            ]
          : [paragraph("00000100", "Owner")],
      );
    });
  }
});
