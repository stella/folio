import { expect, test } from "bun:test";

import { projectReview } from "../../../../../test/reviewProjection";
import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { normalizeForOps } from "../contract";
import { isTrackedWrapper, stampedRevisionIds } from "../review";
import type { DocumentOp } from "../types";

const boundaries = [
  { type: "bookmarkStart", id: 9, name: "_Ref9" },
  { type: "bookmarkEnd", id: 9 },
  { type: "moveFromRangeStart", id: 9, name: "move9", author: "Earlier" },
  { type: "moveFromRangeEnd", id: 9 },
  { type: "moveToRangeStart", id: 9, name: "move9", author: "Earlier" },
  { type: "moveToRangeEnd", id: 9 },
] as const satisfies readonly ParagraphContent[];

const originalStamp = { id: 1, author: "C", date: "2026-03-04T05:00:00.000Z" };
const typingStamp = { id: 2, author: "C", date: "2026-03-04T05:00:01.000Z" };
const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});
const link = (children: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children,
});
const paragraphDocument = (content: ParagraphContent[]) =>
  normalizeForOps({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content }],
      },
    },
  } satisfies Document);

/** Every producer and resolution must remain exactly undoable on the wire. */
const apply = (document: Document, operation: DocumentOp) => {
  const result = applyDocumentOp(
    document,
    JSON.parse(JSON.stringify(operation)),
  ).unwrap();
  const undo = applyDocumentOps(
    result.document,
    JSON.parse(JSON.stringify(result.inverse)),
  ).unwrap();
  expect(undo.document).toStrictEqual(document);
  expect(
    applyDocumentOps(undo.document, undo.inverse).unwrap().document,
  ).toStrictEqual(result.document);
  return result.document;
};

const projection = (document: Document) =>
  projectReview({ document, projection: "π′" });

for (const side of ["before", "after"] as const)
  for (const depth of [0, 1, 2])
    test(`a ${side} source boundary follows its accepted container through ${depth} authored wrappers`, () => {
      for (const boundary of boundaries) {
        const source = paragraphDocument([
          {
            type: "run",
            content: [{ type: "tab" }, { type: "text", text: "a" }],
          },
        ]);
        let payload = link(
          side === "before" ? [boundary, text("a")] : [text("a"), boundary],
        );
        for (let level = 0; level < depth; level++)
          payload = {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [payload],
          };
        const inserted = apply(source, {
          type: "insertContent",
          at: { story: "main", blockId: "00000001", offset: 1 },
          slice: { openStart: 0, openEnd: 0, content: [payload] },
          revision: originalStamp,
        });
        const typed = apply(inserted, {
          type: "insertText",
          at: {
            story: "main",
            blockId: "00000001",
            offset: side === "before" ? 1 : 2,
            zeroWidthBefore: side === "before" ? 1 : 0,
          },
          text: "x",
          runProps: {},
          revision: typingStamp,
          newIds: { revision: [100, 101, 102, 103] },
        });
        const paragraphs = storyParagraphs(typed.package.document).map(
          ({ paragraph }) => paragraph,
        );
        const originalIds = stampedRevisionIds(
          paragraphs,
          originalStamp,
          new Set(),
        );
        const typingIds = stampedRevisionIds(
          paragraphs,
          typingStamp,
          new Set(),
        );
        const groups = [originalIds, typingIds];
        expect(groups.every((group) => group.length > 0)).toBe(true);
        const acceptedTyping = apply(typed, {
          type: "resolveRevision",
          story: "main",
          revisionIds: typingIds,
          decision: "accept",
        });
        const holders = storyParagraphs(acceptedTyping.package.document)
          .flatMap(({ paragraph }) => paragraph.content)
          .filter(isTrackedWrapper)
          .filter((node) => originalIds.includes(node.info.id));
        // These actual producers must exercise the explicit handoff, rather
        // than merely compare orders that never reach its guarded branch.
        expect(holders.length).toBeGreaterThan(0);
        expect(
          holders.every(
            ({ resolutionJoins: joins }) =>
              joins !== undefined &&
              joins.remove > 0 &&
              ((joins.before === 0 && joins.after > 0) ||
                (joins.after === 0 && joins.before > 0)),
          ),
        ).toBe(true);
        for (const original of ["accept", "reject"] as const)
          for (const typing of ["accept", "reject"] as const) {
            const decisions = [original, typing];
            const results = [];
            const individual = [
              { revisionIds: originalIds, decision: original },
              { revisionIds: typingIds, decision: typing },
            ];
            for (const order of [individual, individual.toReversed()]) {
              let sequential = typed;
              for (const entry of order)
                sequential = apply(sequential, {
                  type: "resolveRevision",
                  story: "main",
                  ...entry,
                });
              results.push(projection(sequential));
            }
            for (const order of [
              ["accept", "reject"],
              ["reject", "accept"],
            ] as const) {
              let batched = typed;
              for (const decision of order)
                batched = apply(batched, {
                  type: "resolveRevision",
                  story: "main",
                  revisionIds: groups.flatMap((group, index) =>
                    decisions[index] === decision ? group : [],
                  ),
                  decision,
                });
              results.push(projection(batched));
            }
            for (const result of results)
              expect(result).toStrictEqual(results.at(0));
          }
      }
    });

test("independent authored boundary and matching containers acquire no handoff", () => {
  for (const boundary of boundaries) {
    const source = paragraphDocument([
      link([text("a")]),
      boundary,
      link([text("a")]),
    ]);
    const inserted = apply(source, {
      type: "insertText",
      at: { story: "main", blockId: "00000001", offset: 1, zeroWidthBefore: 1 },
      text: "x",
      revision: typingStamp,
    });
    const paragraph = inserted.package.document.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type === "paragraph") {
      const insertion = paragraph.content.find(
        (node) => isTrackedWrapper(node) && node.info.id === typingStamp.id,
      );
      expect(
        insertion !== undefined && isTrackedWrapper(insertion)
          ? insertion.resolutionJoins?.remove
          : undefined,
      ).toBe(0);
    }
    const rejected = apply(inserted, {
      type: "resolveRevision",
      story: "main",
      revisionIds: [typingStamp.id],
      decision: "reject",
    });
    expect(rejected).toStrictEqual(source);
  }
});
