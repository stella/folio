import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { parseDocx } from "../../docx/parser";
import { createDocx, repackDocx } from "../../docx/rezip";
import type { Document, TableCell, Paragraph } from "../../types/document";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import {
  acceptAIEditRevision,
  acceptAllChanges,
  rejectAIEditRevision,
  rejectAllChanges,
} from "./comments";

const INFO = { id: 91, author: "Reviewer", date: "2026-09-01T00:00:00Z" };
type CellCase = { continuation: boolean; text: boolean };
type CaseOptions = { cells: readonly CellCase[]; decision: "accept" | "reject" };

const paragraph = (text: string) =>
  ({
    type: "paragraph",
    content: text ? [{ type: "run", content: [{ type: "text", text }] }] : [],
  }) satisfies Paragraph;
const documentFor = ({ cells, decision }: CaseOptions): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: cells.map((_cell, column) => ({
                type: "tableCell",
                formatting: { vMerge: "restart" },
                content: [paragraph(`Top ${column}`)],
              })),
            },
            {
              type: "tableRow",
              cells: cells.map(
                ({ continuation, text }, column): TableCell => ({
                  type: "tableCell",
                  ...(continuation
                    ? {
                        formatting: decision === "accept" ? { vMerge: "continue" } : {},
                        structuralChange: {
                          type: "tableCellMerge",
                          info: INFO,
                          verticalMerge: decision === "accept" ? "continue" : "rest",
                          verticalMergeOriginal: decision === "reject" ? "continue" : "rest",
                        },
                      }
                    : {}),
                  content: [paragraph(text ? `Below ${column}` : "")],
                }),
              ),
            },
          ],
        },
      ],
    },
  },
});

const shape = (doc: PMNode): unknown => {
  const table = doc.firstChild;
  expect(table?.type.name).toBe("table");
  return Array.from({ length: table?.childCount ?? 0 }, (_rowValue, row) => {
    const node = table?.child(row);
    return Array.from({ length: node?.childCount ?? 0 }, (_cellValue, col) => {
      const cell = node?.child(col);
      return {
        text: cell?.textContent,
        rowspan: cell?.attrs["rowspan"],
        colspan: cell?.attrs["colspan"],
      };
    });
  });
};

const exercise = async (options: CaseOptions) => {
  const document = documentFor(options);
  const expectedDocument = structuredClone(document);
  const expectedTable = expectedDocument.package.document.content.at(0);
  if (expectedTable?.type !== "table") throw new Error("Expected table");
  for (const row of expectedTable.rows) {
    for (const cell of row.cells) {
      const change = cell.structuralChange;
      if (change?.type !== "tableCellMerge") continue;
      cell.formatting = { ...cell.formatting, vMerge: "continue" };
      delete cell.structuralChange;
    }
  }
  const expected = shape(toProseDoc(expectedDocument));
  for (const path of ["targeted", "bulk"] as const) {
    let state = EditorState.create({ doc: toProseDoc(document) });
    const commands = {
      accept: { targeted: acceptAIEditRevision(INFO.id), bulk: acceptAllChanges() },
      reject: { targeted: rejectAIEditRevision(INFO.id), bulk: rejectAllChanges() },
    };
    const command = commands[options.decision][path];
    expect(
      command(state, (tr) => {
        state = state.apply(tr);
      }),
    ).toBe(true);
    expect(shape(state.doc)).toEqual(expected);
    state.doc.check();
    const saved = fromProseDoc(state.doc, document);
    const reopened = await parseDocx(await createDocx(saved));
    expect(shape(toProseDoc(reopened))).toEqual(expected);
    const reopenedAgain = await parseDocx(
      await repackDocx(fromProseDoc(toProseDoc(reopened), reopened)),
    );
    expect(shape(toProseDoc(reopenedAgain))).toEqual(expected);
  }
};

test("reject retains a continuation with text beside an ordinary cell", async () => {
  await exercise({
    cells: [
      { continuation: true, text: true },
      { continuation: false, text: true },
    ],
    decision: "reject",
  });
});

test("reject retains the sole continuation in a single-column row", async () => {
  await exercise({ cells: [{ continuation: true, text: false }], decision: "reject" });
});

test(
  "resolved vertical merges match the reader across content and row shapes",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.record({ continuation: fc.boolean(), text: fc.boolean() }), {
          minLength: 1,
          maxLength: 4,
        }),
        async (generated) => {
          const cells = generated.map((cell, index) => ({
            continuation: index === 0 || cell.continuation,
            text: cell.text,
          }));
          for (const decision of ["accept", "reject"] as const) await exercise({ cells, decision });
        },
      ),
      { numRuns: 30 },
    );
  },
  propertyTestTimeout(30_000),
);
