import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, type Transaction } from "prosemirror-state";
import { Step } from "prosemirror-transform";

import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import { RevisionResolutionStep } from "./revisionResolutionStep";

const revision = (id: number) => ({ revisionId: id, author: "Reviewer", date: "2026-09-09" });

const paragraph = (
  index: number,
  mark: "none" | "ins" | "del",
  inline: "plain" | "ins" | "del" | "format",
) => {
  const content = [schema.text(`start${index}`)];
  if (inline === "ins" || inline === "del") {
    const type = inline === "ins" ? schema.marks.insertion : schema.marks.deletion;
    content.push(schema.text(`change${index}`, [type.create(revision(index * 10 + 1))]));
  }
  if (inline === "format") {
    content.push(
      schema.text(`format${index}`, [
        schema.marks.bold.create(),
        schema.marks.runPropertyChange.create({
          changes: [
            {
              type: "runPropertyChange",
              info: { id: index * 10 + 2, author: "Reviewer", date: "2026-09-09" },
              previousFormatting: { italic: true },
              currentFormatting: { bold: true },
            },
          ],
        }),
      ]),
    );
  }
  return schema.node(
    "paragraph",
    {
      paraId: index.toString(16).padStart(8, "0"),
      ...(mark === "none"
        ? {}
        : {
            pPrMark: {
              kind: mark,
              info: { id: index * 10 + 3, author: "Reviewer", date: "2026-09-09" },
            },
          }),
      ...(index % 2 === 0
        ? {
            alignment: "right",
            _originalFormatting: { alignment: "right" },
            _propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: { id: index * 10 + 4, author: "Reviewer", date: "2026-09-09" },
                previousFormatting: { alignment: "left" },
              },
            ],
          }
        : {}),
    },
    content,
  );
};

const table = (rowMarker: "none" | "trIns" | "trDel", cellMarker: "none" | "ins" | "del") =>
  schema.node("table", null, [
    schema.node("tableRow", null, [
      schema.node(
        "tableCell",
        cellMarker === "none" ? null : { cellMarker: { kind: cellMarker, info: revision(902) } },
        [paragraph(90, "none", "plain")],
      ),
    ]),
    schema.node("tableRow", rowMarker === "none" ? null : { [rowMarker]: revision(901) }, [
      schema.node("tableCell", null, [paragraph(91, "none", "del")]),
    ]),
  ]);

test("bulk resolution JSON replay and undo match its cached result", () => {
  const result = fc.check(
    fc.property(
      fc.array(
        fc.record({
          mark: fc.constantFrom("none", "ins", "del"),
          inline: fc.constantFrom("plain", "ins", "del", "format"),
        }),
        { minLength: 2, maxLength: 5 },
      ),
      fc.constantFrom("none", "trIns", "trDel"),
      fc.constantFrom("none", "ins", "del"),
      (items, rowMarker, cellMarker) => {
        const blocks = items.map(({ mark, inline }, index) => paragraph(index + 1, mark, inline));
        blocks.splice(1, 0, table(rowMarker, cellMarker));
        const doc = schema.node("doc", null, blocks);
        const state = EditorState.create({ schema, doc });
        for (const mode of ["accept", "reject"] as const) {
          let transaction: Transaction | null = null;
          const command = mode === "accept" ? acceptAllChanges() : rejectAllChanges();
          expect(
            command(state, (dispatched) => {
              transaction = dispatched;
            }),
          ).toBe(true);
          if (!transaction) {
            continue;
          }
          expect(transaction.steps).toHaveLength(1);
          const step = transaction.steps.at(0);
          expect(step).toBeInstanceOf(RevisionResolutionStep);
          if (!step) {
            throw new Error("Missing bulk revision step");
          }
          const resolved = transaction.doc;
          expect(step.getMap().map(0, -1)).toBe(0);
          expect(step.getMap().map(doc.content.size, 1)).toBe(resolved.content.size);
          step.getMap().forEach((oldStart, oldEnd, newStart, newEnd) => {
            expect(oldStart).toBeGreaterThanOrEqual(0);
            expect(oldEnd).toBeLessThanOrEqual(doc.content.size);
            expect(newStart).toBeGreaterThanOrEqual(0);
            expect(newEnd).toBeLessThanOrEqual(resolved.content.size);
          });
          const replayed = Step.fromJSON(schema, step.toJSON()).apply(doc);
          if (!replayed.doc?.eq(resolved)) {
            throw new Error(
              `${mode} replay failed: ${replayed.failed ?? JSON.stringify(replayed.doc?.toJSON())} expected ${JSON.stringify(resolved.toJSON())}`,
            );
          }
          const inverse = step.invert(doc);
          const undone = inverse.apply(resolved);
          if (!undone.doc?.eq(doc)) {
            throw new Error(
              `${mode} cached undo failed: ${undone.failed ?? JSON.stringify(undone.doc?.toJSON())}`,
            );
          }
          const replayedUndo = Step.fromJSON(schema, inverse.toJSON()).apply(resolved);
          if (!replayedUndo.doc?.eq(doc)) {
            throw new Error(
              `${mode} undo replay failed: ${replayedUndo.failed ?? JSON.stringify(replayedUndo.doc?.toJSON())}`,
            );
          }
        }
      },
    ),
    { seed: 260926, numRuns: 32, verbose: true },
  );
  if (result.failed) {
    throw result.errorInstance;
  }
});

test("a paragraph join after whole-table deletion has a valid replay map", () => {
  const doc = schema.node("doc", null, [
    paragraph(1, "del", "plain"),
    table("trDel", "del"),
    paragraph(2, "none", "plain"),
  ]);
  const state = EditorState.create({ schema, doc });
  let transaction: Transaction | null = null;
  expect(
    acceptAllChanges()(state, (dispatched) => {
      transaction = dispatched;
    }),
  ).toBe(true);
  if (!transaction) throw new Error("Missing bulk revision transaction");
  const step = transaction.steps.at(0);
  expect(step).toBeInstanceOf(RevisionResolutionStep);
  if (!step) throw new Error("Missing bulk revision step");
  const replayed = Step.fromJSON(schema, step.toJSON()).apply(doc);
  expect(replayed.doc?.eq(transaction.doc)).toBe(true);
  const inverse = Step.fromJSON(schema, step.invert(doc).toJSON());
  expect(inverse.apply(transaction.doc).doc?.eq(doc)).toBe(true);
});
