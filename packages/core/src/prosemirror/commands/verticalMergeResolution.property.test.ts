/**
 * Generate content, implicit origins, chains, whole-row continuations and
 * optional dates. Targeted and bulk resolution must preserve the same spans
 * and payloads; saved content must match the reader across repeated saves.
 */
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
type CaseOptions = {
  cells: readonly CellCase[];
  decision: "accept" | "reject";
  origin?: "plain" | "restart";
  continuations?: number;
  date?: "present" | "absent" | "null";
  rowTexts?: readonly (readonly boolean[])[];
};

const paragraph = (text: string) =>
  ({
    type: "paragraph",
    content: text ? [{ type: "run", content: [{ type: "text", text }] }] : [],
  }) satisfies Paragraph;
const documentFor = ({
  cells,
  decision,
  origin = "restart",
  continuations = 1,
  date = "present",
  rowTexts,
}: CaseOptions): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: cells.map(
                (_cell, column): TableCell => ({
                  type: "tableCell",
                  ...(origin === "restart" ? { formatting: { vMerge: "restart" } } : {}),
                  content: [paragraph(`Top ${column}`)],
                }),
              ),
            },
            ...Array.from({ length: continuations }, (_value, row) => ({
              type: "tableRow" as const,
              cells: cells.map(
                ({ continuation, text }, column): TableCell => ({
                  type: "tableCell",
                  ...(continuation
                    ? {
                        formatting:
                          decision === "accept" && origin === "restart"
                            ? { vMerge: "continue" }
                            : {},
                        structuralChange: {
                          type: "tableCellMerge",
                          info: date === "present" ? INFO : { id: INFO.id, author: INFO.author },
                          verticalMerge: decision === "accept" ? "continue" : "rest",
                          verticalMergeOriginal: decision === "reject" ? "continue" : "rest",
                        },
                      }
                    : {}),
                  content: [
                    paragraph(
                      (rowTexts?.at(row)?.at(column) ?? text) ? `Below ${row}-${column}` : "",
                    ),
                  ],
                }),
              ),
            })),
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
  if (options.origin === "plain") {
    const top = expectedTable.rows.at(0);
    if (!top) throw new Error("Expected origin row");
    for (const [column, cell] of top.cells.entries()) {
      if (options.cells.at(column)?.continuation) cell.formatting = { vMerge: "restart" };
    }
  }
  const expected = shape(toProseDoc(expectedDocument));
  const continuations = options.continuations ?? 1;
  const blocked = options.cells.map(() => false);
  const top = options.cells.map((_cell, column) => ({
    text: `Top ${column}`,
    rowspan: 1,
    colspan: 1,
  }));
  const expectedResolved = [top];
  for (let row = 0; row < continuations; row++) {
    expectedResolved.push(
      options.cells.flatMap(({ continuation, text }, column) => {
        const hasText = options.rowTexts?.at(row)?.at(column) ?? text;
        const folds =
          continuation && (options.decision === "reject" || (!hasText && !blocked.at(column)));
        if (folds) {
          const origin = top.at(column);
          if (!origin) throw new Error("Missing expected merge origin");
          origin.rowspan++;
          return [];
        }
        if (continuation && hasText) blocked[column] = true;
        return [{ text: hasText ? `Below ${row}-${column}` : "", rowspan: 1, colspan: 1 }];
      }),
    );
  }
  let targeted: PMNode | undefined;
  for (const path of ["targeted", "bulk"] as const) {
    let state = EditorState.create({ doc: toProseDoc(document) });
    if (options.date === "null") {
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        const marker = node.attrs["cellMarker"];
        if (marker?.kind === "merge")
          tr.setNodeAttribute(pos, "cellMarker", {
            ...marker,
            info: { ...marker.info, date: null },
          });
      });
      state = state.apply(tr);
    }
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
    expect(shape(state.doc)).toEqual(expectedResolved);
    if (path === "targeted") targeted = state.doc;
    else expect(state.doc.toJSON()).toEqual(targeted?.toJSON());
    state.doc.check();
    const saved = fromProseDoc(state.doc, document);
    // The reader's visible cells include stored continuation payload content.
    expect(shape(toProseDoc(saved))).toEqual(expected);
    const reopened = await parseDocx(await createDocx(saved));
    expect(shape(toProseDoc(reopened))).toEqual(expected);
    const reopenedAgain = await parseDocx(
      await repackDocx(fromProseDoc(toProseDoc(reopened), reopened)),
    );
    expect(shape(toProseDoc(reopenedAgain))).toEqual(expected);
  }
};

test("reject folds a continuation with text beside an ordinary cell", async () => {
  await exercise({
    cells: [
      { continuation: true, text: true },
      { continuation: false, text: true },
    ],
    decision: "reject",
  });
});

test("reject folds the sole continuation in a single-column row", async () => {
  await exercise({ cells: [{ continuation: true, text: false }], decision: "reject" });
});

test("accept keeps an empty continuation below a visible text continuation", async () => {
  await exercise({
    cells: [{ continuation: true, text: false }],
    decision: "accept",
    continuations: 2,
    rowTexts: [[true], [false]],
  });
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
        fc.constantFrom("plain", "restart"),
        fc.integer({ min: 1, max: 3 }),
        fc.constantFrom("present", "absent", "null"),
        fc.array(fc.array(fc.boolean(), { maxLength: 4 }), { maxLength: 3 }),
        async (generated, origin, continuations, date, rowTexts) => {
          const cells = generated.map((cell, index) => ({
            continuation: index === 0 || cell.continuation,
            text: cell.text,
          }));
          for (const decision of ["accept", "reject"] as const)
            await exercise({ cells, decision, origin, continuations, date, rowTexts });
        },
      ),
      { numRuns: 30 },
    );
  },
  propertyTestTimeout(30_000),
);
