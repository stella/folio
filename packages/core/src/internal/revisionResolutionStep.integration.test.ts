import { expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";
import { Step } from "prosemirror-transform";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { acceptAllChanges, rejectAllChanges, rejectChange } from "../prosemirror/commands/comments";
import { createHarnessState } from "../__tests__/editorHarness";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Run } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { schema } from "../prosemirror/schema";
import { RevisionResolutionStep } from "./revisionResolutionStep";

const revision = (id: number) => ({ revisionId: id, author: "Reviewer", date: "2026-09-09" });

test(
  "restored historical runs preserve authorship through bulk and individual rejection",
  () => {
    // The old generator covered structural markers but never imported historical runs
    // beneath paragraph-mark formatting, so both paths could agree on invented overrides.
    assertProperty(
      fc.property(
        fc.record({
          bold: fc.boolean(),
          italic: fc.boolean(),
          fontSize: fc.integer({ min: 8, max: 32 }),
        }),
        fc.option(
          fc.record({
            bold: fc.boolean(),
            italic: fc.boolean(),
            fontSize: fc.integer({ min: 8, max: 32 }),
          }),
          { nil: undefined },
        ),
        fc.string({ minLength: 1, maxLength: 12 }).filter((text) => text.trim().length > 0),
        (paragraphMark, direct, text) => {
          const source = createEmptyDocument();
          const run = {
            type: "run",
            content: [{ type: "text", text }],
            ...(direct ? { formatting: direct } : {}),
          } satisfies Run;
          source.package.document.content = [
            {
              type: "paragraph",
              paraId: "1A000001",
              formatting: { runProperties: paragraphMark },
              content: [
                { type: "deletion", info: { id: 101, author: "Reviewer" }, content: [run] },
              ],
            },
          ];
          const expected = createEmptyDocument();
          expected.package.document.content = [
            {
              type: "paragraph",
              paraId: "1A000001",
              formatting: { runProperties: paragraphMark },
              content: [run],
            },
          ];
          const expectedContent = fromProseDoc(
            createHarnessState(expected, "editing").doc,
            expected,
            { stylesheetSource: { type: "package" } },
          ).package.document.content;
          for (const command of [
            rejectAllChanges(),
            rejectChange(0, toProseDoc(source).content.size),
          ]) {
            const state = createHarnessState(source, "editing");
            let resolved = state;
            expect(
              command(state, (tr) => {
                resolved = state.apply(tr);
              }),
            ).toBe(true);
            expect(
              fromProseDoc(resolved.doc, source, { stylesheetSource: { type: "package" } }).package
                .document.content,
            ).toEqual(expectedContent);
            const expectedDoc = createHarnessState(expected, "editing").doc;
            expect(resolved.doc.eq(expectedDoc)).toBe(true);
          }
        },
      ),
      {
        numRuns: 40,
        seed: 16091615,
        id: "restored historical runs preserve authorship through bulk and individual rejection",
      },
    );
  },
  propertyTestTimeout(5_000),
);

const paragraph = (
  index: number,
  mark: "none" | "ins" | "del",
  inline: "plain" | "ins" | "del" | "format" | "ins-only" | "del-only",
) => {
  // A paragraph holding only a revision empties when it is resolved away, so
  // the paragraph marks either side of it join with nothing between them.
  if (inline === "ins-only" || inline === "del-only") {
    const type = inline === "ins-only" ? schema.marks.insertion : schema.marks.deletion;
    return paragraphNode(index, mark, [
      schema.text(`change${index}`, [type.create(revision(index * 10 + 1))]),
    ]);
  }
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
  return paragraphNode(index, mark, content);
};

const paragraphNode = (index: number, mark: "none" | "ins" | "del", content: PMNode[]) =>
  schema.node(
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

test(
  "bulk resolution JSON replay and undo match its cached result",
  () => {
    assertProperty(
      fc.property(
        fc.array(
          fc.record({
            mark: fc.constantFrom("none", "ins", "del"),
            inline: fc.constantFrom("plain", "ins", "del", "format", "ins-only", "del-only"),
          }),
          { minLength: 2, maxLength: 5 },
        ),
        fc.constantFrom("none", "trIns", "trDel"),
        fc.constantFrom("none", "ins", "del"),
        fc.boolean(),
        (items, rowMarker, cellMarker, withBookmarks) => {
          const blocks = items.map(({ mark, inline }, index) => paragraph(index + 1, mark, inline));
          const between = [table(rowMarker, cellMarker)];
          if (withBookmarks) {
            between.unshift(
              schema.node("blockBookmarkBoundary", { type: "start", id: 1, name: "range" }),
            );
            between.push(schema.node("blockBookmarkBoundary", { type: "end", id: 1 }));
          }
          blocks.splice(1, 0, ...between);
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
              // A removed range reads as removed from both of its edges. Two
              // removals left adjacent in one map break this: the map reads
              // their shared boundary against the first only.
              if (oldEnd > oldStart) {
                const removedAfter = step.getMap().mapResult(oldStart, 1).deletedAfter;
                const removedBefore = step.getMap().mapResult(oldEnd, -1).deletedBefore;
                if (!removedAfter || !removedBefore) {
                  throw new Error(
                    `Invalid ${mode} map: ${JSON.stringify({ oldStart, oldEnd, removedAfter, removedBefore, mapRanges: step.toJSON().mapRanges })}`,
                  );
                }
              }
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
      {
        seed: 260926,
        numRuns: 32,
        verbose: true,
        id: "bulk resolution JSON replay and undo match its cached result",
      },
    );
  },
  propertyTestTimeout(30_000),
);

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
