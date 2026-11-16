import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document, ParagraphContent } from "../types/document";

const shape = fc.record({
  depth: fc.integer({ min: 0, max: 2 }),
  orphan: fc.boolean(),
  reply: fc.boolean(),
});
type Shape = fc.RecordValue<typeof shape>;

const documentOf = ({ depth, orphan, reply }: Shape): Document => {
  let content: ParagraphContent[] = [{ type: "commentReference", id: 7 }];
  for (let index = 0; index < depth; index += 1)
    content =
      index === 0
        ? [
            {
              type: "hyperlink",
              href: "https://example.org",
              children: content,
            },
          ]
        : [
            {
              type: "inlineSdt",
              properties: { sdtType: "richText", id: index, tag: "clause" },
              content,
            },
          ];
  return {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content },
          { type: "paragraph", paraId: "00000002", content: [] },
        ],
        comments: [
          {
            id: 7,
            author: "A",
            content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
          },
          ...(reply
            ? [
                {
                  id: 8,
                  parentId: 7,
                  author: "B",
                  content: [
                    {
                      type: "paragraph" as const,
                      paraId: "7FFFFFF1",
                      content: [],
                    },
                  ],
                },
              ]
            : []),
          ...(orphan
            ? [
                {
                  id: 99,
                  author: "C",
                  content: [
                    {
                      type: "paragraph" as const,
                      paraId: "7FFFFFF2",
                      content: [],
                    },
                  ],
                },
              ]
            : []),
        ],
      },
    },
  };
};

const commentIds = (document: Document) =>
  (document.package.document.comments ?? []).map(({ id }) => id);
const parse = (buffer: ArrayBuffer) =>
  parseDocx(buffer, { preloadFonts: false, detectVariables: false });

test(
  "point threads retire with their last anchor while source orphans survive edits, inverse and save",
  async () => {
    await assertProperty(
      fc.asyncProperty(shape, async (example) => {
        const source = await createDocx(documentOf(example));
        const parsed = await parse(source);
        expect(
          parsed.package.document.comments?.find(({ id }) => id === 7)
            ?.anchorKind,
        ).toBe("point");
        if (example.orphan) {
          const orphan = parsed.package.document.comments?.find(
            ({ id }) => id === 99,
          );
          expect(orphan).toBeDefined();
          expect(orphan?.anchorKind).toBeUndefined();
        }
        const expected = example.orphan ? [99] : [];
        const applied = applyDocumentOp(parsed, {
          type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
          story: OP_STORIES.MAIN,
          blockIds: ["00000001"],
        }).unwrap();
        expect(commentIds(applied.document)).toEqual(expected);
        const undone = applyDocumentOps(
          applied.document,
          JSON.parse(JSON.stringify(applied.inverse)),
        ).unwrap();
        expect(undone.document).toStrictEqual(parsed);
        expect(
          applyDocumentOps(
            undone.document,
            JSON.parse(JSON.stringify(undone.inverse)),
          ).unwrap().document,
        ).toStrictEqual(applied.document);

        const reviewer = await FolioDocxReviewer.fromBuffer(source, {
          session: "canonical",
        });
        const edit = reviewer.applyOperations(
          [{ id: "remove", type: "deleteBlock", blockId: "00000001" }],
          { mode: "direct", undefinedReferences: "refuse" },
        );
        expect(edit.applied).toHaveLength(1);
        expect(reviewer.getComments().map(({ id }) => id)).toEqual(expected);
        const saved = await reviewer.toBuffer();
        expect(commentIds(await parse(saved))).toEqual(expected);
        const reopened = await FolioDocxReviewer.fromBuffer(saved, {
          session: "canonical",
        });
        expect(reopened.getComments()).toEqual(reviewer.getComments());
        expect(
          (
            await FolioDocxReviewer.fromBuffer(await reopened.toBuffer(), {
              session: "canonical",
            })
          ).getComments(),
        ).toEqual(reviewer.getComments());

        for (const type of ["insertion", "deletion"] as const) {
          const tracked = documentOf(example);
          const first = tracked.package.document.content.at(0);
          if (first?.type !== "paragraph")
            panic("Fixture needs its first paragraph");
          first.content = [
            {
              type,
              info: { id: 10, author: "A", date: "2026-01-01T00:00:00Z" },
              content: first.content,
            },
          ];
          const trackedSource = await createDocx(tracked);
          const trackedParsed = await parse(trackedSource);
          for (const decision of ["accept", "reject"] as const) {
            const keeps = (type === "insertion") === (decision === "accept");
            const resolved = applyDocumentOp(trackedParsed, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: [10],
              decision,
            }).unwrap();
            expect(commentIds(resolved.document)).toEqual(
              keeps ? commentIds(trackedParsed) : expected,
            );
            const restored = applyDocumentOps(
              resolved.document,
              JSON.parse(JSON.stringify(resolved.inverse)),
            ).unwrap();
            expect(restored.document).toStrictEqual(trackedParsed);
            expect(
              applyDocumentOps(
                restored.document,
                JSON.parse(JSON.stringify(restored.inverse)),
              ).unwrap().document,
            ).toStrictEqual(resolved.document);
            const reviewed = await FolioDocxReviewer.fromBuffer(trackedSource, {
              session: "canonical",
            });
            const result =
              decision === "accept"
                ? reviewed.acceptAll()
                : reviewed.rejectAll();
            expect(result.skipped).toEqual([]);
            expect(result.resolved).toBeGreaterThan(0);
            const savedReview = await reviewed.toBuffer();
            expect(commentIds(await parse(savedReview))).toEqual(
              keeps ? commentIds(trackedParsed) : expected,
            );
            expect(
              (
                await FolioDocxReviewer.fromBuffer(savedReview, {
                  session: "canonical",
                })
              ).getComments(),
            ).toEqual(reviewed.getComments());
          }
        }
      }),
      {
        numRuns: 24,
        id: "point comment exhaustion preserves source orphans",
        examples: [0, 1, 2].flatMap((depth) =>
          [false, true].flatMap((orphan) =>
            [false, true].map((reply): [Shape] => [{ depth, orphan, reply }]),
          ),
        ),
      },
    );
  },
  propertyTestTimeout(60_000),
);
