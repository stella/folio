import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { createEmptyDocument } from "../utils/createDocument";
import { expectTrackedChangeMarkAttrs } from "./attrs";
import { addTrackedDeletionMark } from "./addTrackedDeletionMark";
import { resolveAllChangesInHeadlessState } from "./commands/comments";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema/index";

test(
  "whole-field deletion retains nested child insertion paths",
  () => {
    assertProperty(
      fc.property(
        fc.integer({ min: 1, max: 12 }),
        fc.integer({ min: 1, max: 4 }),
        (length, depth) => {
          const ancestors = Array.from({ length: depth }, (_, index) => ({
            type: "insertion" as const,
            revisionId: 100 + index,
            author: "Earlier Reviewer",
            outerWrapperCount: 0,
          }));
          const insertion = schema.mark("insertion", {
            revisionId: 200,
            author: "Earlier Reviewer",
            _docxRevisionAncestors: ancestors,
          });
          const existing = schema.mark("deletion", { revisionId: 300, author: "Other Reviewer" });
          const field = schema.node("structuredField", null, [
            schema.text("x".repeat(length), [insertion]),
            schema.text("hidden", [existing]),
          ]);
          const doc = schema.node("doc", null, schema.node("paragraph", null, [field]));
          for (const insertionPolicy of ["preserve-pending", "retract-own"] as const) {
            const tr = EditorState.create({ doc }).tr;
            addTrackedDeletionMark({
              tr,
              from: 1,
              to: 1 + field.nodeSize,
              mark: schema.mark("deletion", { revisionId: 400, author: "New Reviewer" }),
              insertionPolicy,
            });
            const inserted = tr.doc.nodeAt(2);
            const deleted = inserted?.marks.find(({ type }) => type.name === "deletion");
            if (!deleted) throw new TypeError("Nested inserted child requires a deletion");
            expect(
              expectTrackedChangeMarkAttrs(deleted)._docxRevisionAncestors?.map(
                ({ revisionId }) => revisionId,
              ),
            ).toEqual([...ancestors.map(({ revisionId }) => revisionId), 200]);
            expect(tr.doc.nodeAt(2 + length)?.marks).toContainEqual(existing);
            expect(tr.doc.textContent).toBe(doc.textContent);
            const first = tr.doc.toJSON();
            addTrackedDeletionMark({
              tr,
              from: 1,
              to: 1 + field.nodeSize,
              mark: schema.mark("deletion", { revisionId: 500, author: "Another Reviewer" }),
              insertionPolicy,
            });
            expect(tr.doc.toJSON()).toEqual(first);
          }
        },
      ),
      { numRuns: 30, id: "whole-field deletion retains nested child insertion paths" },
    );
  },
  propertyTestTimeout(5_000),
);

test(
  "range deletion preserves every existing deletion and is idempotent",
  async () => {
    await assertProperty(
      fc.property(
        fc.array(
          fc.record({
            kind: fc.constantFrom("text", "hardBreak"),
            pending: fc.constantFrom("none", "insertion", "deletion", "nested"),
          }),
          { minLength: 1, maxLength: 40 },
        ),
        fc.nat(),
        fc.nat(),
        fc.constantFrom("none", "sdt", "structuredField"),
        (runs, start, length, container) => {
          const content = runs.map(({ kind, pending }, index) => {
            const attrs = { revisionId: index + 1, author: "Earlier Reviewer" };
            const marks =
              pending === "none"
                ? []
                : [
                    schema.mark(
                      pending === "insertion" ? "insertion" : "deletion",
                      pending === "nested"
                        ? {
                            ...attrs,
                            _docxRevisionAncestors: [
                              {
                                type: "insertion",
                                revisionId: 100 + index,
                                author: "First Reviewer",
                                outerWrapperCount: 0,
                              },
                            ],
                          }
                        : attrs,
                    ),
                  ];
            return kind === "text"
              ? schema.text("x", marks)
              : schema.node("hardBreak", null, null, marks);
          });
          const doc = schema.node(
            "doc",
            null,
            schema.node(
              "paragraph",
              null,
              container !== "none" ? schema.node(container, null, content) : content,
            ),
          );
          const tr = EditorState.create({ doc }).tr;
          const base = container !== "none" ? 2 : 1;
          const from = base + (start % content.length);
          const to = from + (length % (base + content.length + 1 - from));
          const mark = schema.mark("deletion", { revisionId: 999, author: "New Reviewer" });
          addTrackedDeletionMark({ insertionPolicy: "preserve-pending", tr, from, to, mark });
          for (let index = 0; index < content.length; index++) {
            const original = content[index];
            const current = tr.doc.nodeAt(index + base);
            if (!original || !current) throw new Error("fixture run missing");
            const prior = original.marks.find(({ type }) => type.name === "deletion");
            const deleted = current.marks.find(({ type }) => type.name === "deletion");
            expect(deleted?.toJSON()).toEqual(
              prior?.toJSON() ??
                (index + base >= from && index + base < to ? mark.toJSON() : undefined),
            );
            expect(current.marks.filter(({ type }) => type.name !== "deletion")).toEqual(
              original.marks.filter(({ type }) => type.name !== "deletion"),
            );
          }
          const first = tr.doc.toJSON();
          addTrackedDeletionMark({
            insertionPolicy: "preserve-pending",
            tr,
            from,
            to,
            mark: schema.mark("deletion", { revisionId: 1000, author: "Another Reviewer" }),
          });
          expect(tr.doc.toJSON()).toEqual(first);
        },
      ),
      { numRuns: 100, id: "range deletion preserves every existing deletion and is idempotent" },
    );
  },
  propertyTestTimeout(5_000),
);

test(
  "deletion ownership matrix resolves alike live and after saving",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 1, max: 12 }),
        fc.constantFrom("insertion", "deletion"),
        async (length, ancestor) => {
          for (const insertionPolicy of ["preserve-pending", "retract-own"] as const) {
            const insertion = (author: string, revisionId: number) =>
              schema.mark("insertion", {
                author,
                revisionId,
                _docxRevisionAncestors: [
                  {
                    type: "insertion",
                    revisionId: 10 + revisionId,
                    author: "First Reviewer",
                    outerWrapperCount: 0,
                  },
                ],
              });
            const mine = schema.mark("deletion", { revisionId: 3, author: "Reviewer" });
            const theirs = schema.mark("deletion", {
              revisionId: 4,
              author: "Other Reviewer",
              _docxRevisionAncestors: [
                { type: ancestor, revisionId: 5, author: "First Reviewer", outerWrapperCount: 0 },
              ],
            });
            const nodes = [
              schema.text("p".repeat(length)),
              schema.text("i".repeat(length), [insertion("Reviewer", 1)]),
              schema.text("o".repeat(length), [insertion("Other Reviewer", 2)]),
              schema.text("d".repeat(length), [mine]),
              schema.text("n".repeat(length), [theirs]),
            ];
            const state = EditorState.create({
              doc: schema.node("doc", null, schema.node("paragraph", null, nodes)),
            });
            const tr = state.tr;
            addTrackedDeletionMark({
              tr,
              from: 1,
              to: state.doc.content.size - 1,
              mark: schema.mark("deletion", { revisionId: 6, author: "Reviewer" }),
              insertionPolicy,
            });
            const live = state.apply(tr);
            expect(live.doc.textContent).toBe(
              (insertionPolicy === "preserve-pending"
                ? "p".repeat(length) + "i".repeat(length)
                : "p".repeat(length)) +
                "o".repeat(length) +
                "d".repeat(length) +
                "n".repeat(length),
            );
            const source = createEmptyDocument();
            const saved = await createDocx(
              fromProseDoc(live.doc, source, { stylesheetSource: { type: "package" } }),
            );
            const reopened = EditorState.create({
              doc: toProseDoc(await parseDocx(saved, { preloadFonts: false })),
            });
            for (const current of [live, reopened]) {
              current.doc.descendants((node) => {
                if (node.text !== "o".repeat(length)) return;
                const deletion = node.marks.find(({ type }) => type.name === "deletion");
                if (!deletion) throw new Error("inserted fixture run lost its deletion");
                expect(
                  expectTrackedChangeMarkAttrs(deletion)._docxRevisionAncestors?.map(
                    ({ revisionId }) => revisionId,
                  ),
                ).toEqual([12, 2]);
              });
            }
            for (const mode of ["accept", "reject"] as const) {
              const expected =
                mode === "accept"
                  ? ""
                  : "p".repeat(length) +
                    "d".repeat(length) +
                    (ancestor === "deletion" ? "n".repeat(length) : "");
              expect(resolveAllChangesInHeadlessState(live, mode).doc.textContent).toBe(expected);
              expect(resolveAllChangesInHeadlessState(reopened, mode).doc.textContent).toBe(
                expected,
              );
            }
          }
        },
      ),
      { numRuns: 30, id: "deletion ownership matrix resolves alike live and after saving" },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "nested insertion carriers preserve protected descendant revisions under both policies",
  async () => {
    await assertProperty(
      fc.property(
        fc.array(fc.constantFrom("text", "hardBreak"), { minLength: 1, maxLength: 12 }),
        fc.constantFrom("Reviewer", "Other Reviewer"),
        fc.constantFrom("none", "sdt"),
        (kinds, author, wrapper) => {
          const protectedMark = schema.mark("deletion", {
            revisionId: 37,
            author,
            date: "2026-01-02T03:04:05Z",
            initials: "OR",
            _docxRevisionAncestors: [
              { type: "insertion", revisionId: 36, author: "First Reviewer", outerWrapperCount: 0 },
            ],
          });
          const content = kinds.map((kind) =>
            kind === "text"
              ? schema.text("hidden", [protectedMark])
              : schema.node("hardBreak", null, null, [protectedMark]),
          );
          const field = schema.node(
            "structuredField",
            null,
            [schema.text("before"), ...content, schema.text("after")],
            [schema.mark("insertion", { revisionId: 35, author: "Reviewer" })],
          );
          const doc = schema.node(
            "doc",
            null,
            schema.node(
              "paragraph",
              null,
              wrapper === "sdt" ? schema.node("sdt", null, field) : field,
            ),
          );
          for (const insertionPolicy of ["preserve-pending", "retract-own"] as const) {
            const state = EditorState.create({ doc });
            const tr = state.tr;
            addTrackedDeletionMark({
              tr,
              from: 1,
              to: doc.content.size - 1,
              mark: schema.mark("deletion", { revisionId: 38, author: "Reviewer" }),
              insertionPolicy,
            });
            const live = state.apply(tr);
            const protectedNodes = [];
            live.doc.descendants((node) => {
              const deletion = node.marks.find(({ type }) => type.name === "deletion");
              if (deletion?.attrs["revisionId"] === 37) protectedNodes.push(deletion.toJSON());
            });
            const originalNodes = [];
            doc.descendants((node) => {
              const deletion = node.marks.find(({ type }) => type.name === "deletion");
              if (deletion?.attrs["revisionId"] === 37) originalNodes.push(deletion.toJSON());
            });
            expect(protectedNodes).toEqual(originalNodes);
            expect(live.doc.textContent).toBe(doc.textContent);
            for (const mode of ["accept", "reject"] as const) {
              expect(resolveAllChangesInHeadlessState(live, mode).doc.textContent).toBe("");
            }
          }
        },
      ),
      {
        numRuns: 100,
        id: "nested insertion carriers preserve protected descendant revisions under both policies",
      },
    );
  },
  propertyTestTimeout(5_000),
);
