import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { assertProperty } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { createEmptyDocument } from "../utils/createDocument";
import { expectTrackedChangeMarkAttrs } from "./attrs";
import { addTrackedDeletionMark } from "./addTrackedDeletionMark";
import { resolveAllChangesInHeadlessState } from "./commands/comments";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema/index";

test("range deletion preserves every existing deletion and is idempotent", async () => {
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
    { numRuns: 100 },
  );
});

test("deletion ownership matrix resolves alike live and after saving", async () => {
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
          const tr = addTrackedDeletionMark({
            tr: state.tr,
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
          const saved = await createDocx(fromProseDoc(live.doc, source));
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
            expect(resolveAllChangesInHeadlessState(reopened, mode).doc.textContent).toBe(expected);
          }
        }
      },
    ),
    { numRuns: 30 },
  );
}, 30_000);
