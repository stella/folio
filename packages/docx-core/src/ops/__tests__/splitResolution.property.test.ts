import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { projectReview } from "../../../../../test/reviewProjection";
import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { IDENTITY_SPACES, identityKeysIn } from "../ids";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const control = (content: ParagraphContent[], id: number): ParagraphContent => ({
  type: "inlineSdt",
  properties: { sdtType: "richText", id, tag: "clause" },
  content,
});

test(
  "rejecting inserted break chains restores cut controls without removing authored empty controls",
  () => {
    assertProperty(
      fc.property(
        fc.record({
          depth: fc.integer({ min: 1, max: 3 }),
          insertions: fc.array(fc.integer({ min: 1, max: 3 }), { minLength: 1, maxLength: 4 }),
          payload: fc.constantFrom("text", "tab"),
          emptySibling: fc.constantFrom("none", "before", "after"),
          retainedBreak: fc.nat({ max: 3 }),
        }),
        ({ depth, insertions, payload, emptySibling, retainedBreak }) => {
          let content: ParagraphContent[] = [
            {
              type: "run",
              content: payload === "tab" ? [{ type: "tab" }] : [{ type: "text", text: "tail" }],
            },
          ];
          for (let id = depth; id > 0; id--) content = [control(content, id)];
          if (emptySibling === "before") content.unshift(control([], depth + 1));
          if (emptySibling === "after") content.push(control([], depth + 1));
          const source = {
            package: {
              document: { content: [{ type: "paragraph", paraId: "00000001", content }] },
            },
          } satisfies Document;
          let current: Document = source;
          let blockId = "00000001";
          let revisionId = 0;
          let prefixLength = 0;
          const revisions: number[][] = [];
          const breakRevisions: number[] = [];
          const stamp = () => ({
            id: ++revisionId,
            author: "A",
            date: `2026-01-02T03:04:${String(revisionId).padStart(2, "0")}Z`,
          });
          for (const [index, count] of insertions.entries()) {
            // Later edits leave multiple pending insertions in the first fragment.
            for (let insertion = 0; insertion < count; insertion++) {
              const inserted = applyDocumentOp(current, {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: { story: OP_STORIES.MAIN, blockId, offset: 0 },
                text: "x",
                runProps: {},
                revision: stamp(),
                newIds: { revision: [100 + revisionId], control: [200 + revisionId] },
              }).unwrap();
              current = inserted.document;
              revisions.push([...inserted.revisions]);
              prefixLength++;
            }
            const nextId = (index + 2).toString(16).padStart(8, "0");
            const split = applyDocumentOp(current, {
              type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
              at: { story: OP_STORIES.MAIN, blockId, offset: prefixLength },
              newBlockId: nextId,
              revision: stamp(),
              newIds: {
                revision: [300 + index],
                control: [400 + index * 3, 401 + index * 3, 402 + index * 3],
              },
            }).unwrap();
            current = split.document;
            revisions.push([...split.revisions]);
            breakRevisions.push(revisionId);
            // The first split retains the source id on the trailing fragment.
            if (index === 0) blockId = nextId;
          }
          const reject = (document: Document, ids: readonly number[]) =>
            applyDocumentOp(document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: ids,
              decision: REVISION_DECISIONS.REJECT,
            }).unwrap().document;
          const batch = reject(current, revisions.flat());
          let sequential = current;
          for (const ids of revisions.toReversed()) sequential = reject(sequential, ids);
          const projected = (document: Document) => projectReview({ document, projection: "π′" });
          expect(projected(batch)).toStrictEqual(projected(sequential));
          expect(projected(batch)).toStrictEqual(projected(source));
          const controlIds = (document: Document) =>
            identityKeysIn(document.package)
              .filter((key) => key.startsWith(`${IDENTITY_SPACES.CONTROL}:`))
              .sort();
          expect(controlIds(batch)).toEqual(controlIds(source));
          const kept = breakRevisions.at(retainedBreak % breakRevisions.length);
          if (kept === undefined) throw new TypeError("A generated break must exist.");
          const selected = revisions.map((ids) => ids.filter((id) => id !== kept));
          const partial = reject(current, selected.flat());
          let partialSequential = current;
          for (const ids of selected.toReversed()) {
            if (ids.length > 0) partialSequential = reject(partialSequential, ids);
          }
          expect(projected(partial)).toStrictEqual(projected(partialSequential));
          expect(
            identityKeysIn(partial.package).includes(`${IDENTITY_SPACES.REVISION}:${kept}`),
          ).toBe(true);
        },
      ),
      {
        numRuns: 80,
        examples: [
          [
            {
              depth: 1,
              insertions: [1, 1],
              payload: "tab",
              emptySibling: "none",
              retainedBreak: 0,
            },
          ],
        ],
      },
    );
  },
  propertyTestTimeout(10_000),
);
