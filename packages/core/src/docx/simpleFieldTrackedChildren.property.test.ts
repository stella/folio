import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Node as PMNode } from "prosemirror-model";
import type { ParagraphContent, SimpleField, TrackedRunChange } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { resolveAllChangesInHeadlessState } from "../prosemirror/commands/comments";
import { expectTrackedChangeMarkAttrs } from "../prosemirror/attrs";
import { createDocx } from "./rezip";
import { parseDocx } from "./parser";

const childRevisions = (doc: PMNode) => {
  const found = [];
  doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name !== "insertion" && mark.type.name !== "deletion") continue;
      const attrs = expectTrackedChangeMarkAttrs(mark);
      if (attrs.revisionId !== 37) continue;
      found.push({
        text: node.textContent,
        type: mark.type.name,
        revisionId: attrs.revisionId,
        author: attrs.author,
        date: attrs.date,
        initials: attrs.initials,
        ancestors:
          attrs._docxRevisionAncestors
            ?.filter(({ author }) => author === "Ancestor Reviewer")
            .map(({ type, revisionId, author }) => ({ type, revisionId, author })) ?? [],
      });
    }
  });
  return found;
};

test(
  "simple field child revision ownership survives editor save and reopen",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("insertion", "deletion", "moveFrom", "moveTo"),
        fc.constantFrom("none", "insertion", "deletion"),
        fc.constantFrom("run", "hyperlink", "inlineWrapper"),
        fc.integer({ min: 1, max: 8 }),
        fc.constantFrom("none", "insertion"),
        fc.constantFrom("Child Reviewer", "Outer Reviewer", "Žluťoučký", "レビュー"),
        async (kind, outer, carrier, length, ancestry, author) => {
          const run = {
            type: "run",
            content: [{ type: "text", text: "x".repeat(length) }],
          } as const;
          const revision: TrackedRunChange = {
            type: kind,
            info: {
              id: 37,
              author,
              date: "2026-01-02T03:04:05Z",
              initials: "CR",
            },
            content: [{ ...run, content: [...run.content] }],
          };
          const field: SimpleField = {
            type: "simpleField",
            instruction: "REF target",
            fieldType: "REF",
            content: [],
          };
          if (carrier === "hyperlink")
            revision.content = [
              {
                type: "hyperlink",
                anchor: "target",
                children: [{ ...run, content: [...run.content] }],
              },
            ];
          if (carrier === "inlineWrapper")
            revision.content = [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "embedding",
                direction: "rtl",
                content: [{ ...run, content: [...run.content] }],
              },
            ];
          const trackedChild: TrackedRunChange =
            ancestry === "none"
              ? revision
              : {
                  type: "insertion",
                  info: { id: 36, author: "Ancestor Reviewer" },
                  content: [revision],
                };
          field.content = [
            { type: "run", content: [{ type: "text", text: "before" }] },
            trackedChild,
            { type: "run", content: [{ type: "text", text: "after" }] },
          ];
          const source = createEmptyDocument();
          const content: ParagraphContent =
            outer === "none"
              ? field
              : { type: outer, info: { id: 35, author: "Outer Reviewer" }, content: [field] };
          source.package.document.content = [{ type: "paragraph", content: [content] }];
          const initial = toProseDoc(source);
          expect(childRevisions(initial)).toHaveLength(1);
          const live = EditorState.create({ doc: initial });
          const saved = await createDocx(fromProseDoc(initial, source));
          const reopenedSource = await parseDocx(saved, { preloadFonts: false });
          const reopened = EditorState.create({ doc: toProseDoc(reopenedSource) });
          expect(childRevisions(reopened.doc)).toEqual(childRevisions(initial));
          for (const mode of ["accept", "reject"] as const) {
            const childKept =
              mode === "accept"
                ? kind === "insertion" || kind === "moveTo"
                : ancestry === "none" && (kind === "deletion" || kind === "moveFrom");
            const outerRemoved = mode === "accept" ? outer === "deletion" : outer === "insertion";
            const expected = outerRemoved
              ? ""
              : `before${childKept ? "x".repeat(length) : ""}after`;
            expect(resolveAllChangesInHeadlessState(live, mode).doc.textContent).toBe(expected);
            expect(resolveAllChangesInHeadlessState(reopened, mode).doc.textContent).toBe(expected);
          }
          const savedAgain = await createDocx(fromProseDoc(reopened.doc, reopenedSource));
          expect(
            childRevisions(toProseDoc(await parseDocx(savedAgain, { preloadFonts: false }))),
          ).toEqual(childRevisions(initial));
        },
      ),
      { numRuns: 40 },
    );
  },
  propertyTestTimeout(30_000),
);
