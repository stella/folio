import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const inner = { type: "paragraph", paraId: "10000003", content: [] } as const;
const followers = {
  preservedBlock: { type: "preservedBlock", xml: "<w:altChunk/>" },
  blockSdt: {
    type: "blockSdt",
    properties: { sdtType: "richText", id: 5 },
    content: [inner],
  },
  blockCustomXml: {
    type: "blockCustomXml",
    openingXml: '<w:customXml w:element="standing">',
    closingXml: "</w:customXml>",
    content: [inner],
  },
} satisfies Record<string, BlockContent>;

for (const [kind, follower] of Object.entries(followers))
  test(`${kind}: rejected added marks preserve final paragraphs, sections and JSON inverses`, () => {
    for (const tail of [false, true])
      for (const content of ["", "standing"])
        for (const ownSection of [false, true])
          for (const owner of [
            "unavailable",
            "available",
            "same",
            "occupied",
          ] as const) {
            const section = { pageWidth: 11906 };
            const preceding = {
              type: "paragraph",
              paraId: "10000001",
              content: [],
              ...(owner === "same" ? { sectionProperties: section } : {}),
              ...(owner === "occupied"
                ? { sectionProperties: { pageWidth: 12240 } }
                : {}),
            } satisfies Paragraph;
            const target = {
              type: "paragraph",
              paraId: "10000002",
              content: content
                ? [{ type: "run", content: [{ type: "text", text: content }] }]
                : [],
              pPrMark: { kind: "ins", info: { id: 7, author: "Reviewer" } },
              ...(ownSection ? { sectionProperties: section } : {}),
            } satisfies Paragraph;
            const ending = tail
              ? [
                  {
                    type: "paragraph",
                    paraId: "10000004",
                    content: [],
                  } as const,
                ]
              : [];
            const source = {
              package: {
                document: {
                  content: [
                    ...(owner === "unavailable" ? [] : [preceding]),
                    target,
                    follower,
                    ...ending,
                  ],
                  finalSectionProperties: { pageWidth: 15840 },
                },
              },
            } satisfies Document;
            // An added mark's section never stood in the original: it goes
            // with the mark, and no other paragraph receives it.
            // An emptied paragraph goes where another paragraph stays to
            // anchor the inverse: the block after it stays, so its container
            // ends as it did before the insertion, in a paragraph or in an
            // opaque block ending the story.
            const removable =
              content === "" && (tail || owner !== "unavailable");
            const { pPrMark: _resolved, ...unmarked } = target;
            const { sectionProperties: _added, ...rejected } = unmarked;
            const expected = {
              package: {
                document: {
                  content: [
                    ...(owner === "unavailable" ? [] : [preceding]),
                    ...(removable ? [] : [rejected]),
                    follower,
                    ...ending,
                  ],
                  finalSectionProperties:
                    source.package.document.finalSectionProperties,
                },
              },
            } satisfies Document;
            const op = {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: [7],
              decision: "reject",
            } as const;
            const edit = applyDocumentOp(source, op).unwrap();
            expect(edit.document).toEqual(expected);
            const inverse: DocumentOp[] = JSON.parse(
              JSON.stringify(edit.inverse),
            );
            const undone = applyDocumentOps(edit.document, inverse).unwrap();
            expect(undone.document).toEqual(source);
            expect(
              applyDocumentOps(undone.document, undone.inverse).unwrap()
                .document,
            ).toEqual(expected);
            expect(
              applyDocumentOp(source, { ...op, decision: "accept" }).unwrap()
                .document,
            ).toEqual({
              ...source,
              package: {
                ...source.package,
                document: {
                  ...source.package.document,
                  content: source.package.document.content.map((block) =>
                    block === target ? unmarked : block,
                  ),
                },
              },
            });
          }
  });

test("rejecting an added paragraph joining one before a block resolves as if it had never been added", () => {
  const reject = (content: BlockContent[], revisionIds: number[]) =>
    applyDocumentOp(
      {
        package: {
          document: { content, finalSectionProperties: { pageWidth: 15840 } },
        },
      },
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision: "reject",
      },
    ).unwrap().document;
  // It holds only range boundaries, under its own added mark and a property
  // change: rejected, it goes where a paragraph before it takes the
  // boundaries, and stays with its properties restored where none can.
  const target = {
    type: "paragraph",
    paraId: "10000002",
    content: [
      { type: "bookmarkStart", id: 4, name: "kept" },
      { type: "bookmarkEnd", id: 4 },
    ],
    propertyChanges: [
      {
        type: "paragraphPropertyChange",
        info: { id: 10, author: "A" },
        previousFormatting: { alignment: "end" },
      },
    ],
    pPrMark: { kind: "ins", info: { id: 7, author: "A" } },
  } satisfies Paragraph;
  const added = {
    type: "paragraph",
    paraId: "10000005",
    content: [
      {
        type: "insertion",
        info: { id: 8, author: "Reviewer" },
        content: [{ type: "run", content: [{ type: "text", text: "new" }] }],
      },
    ],
    pPrMark: { kind: "ins", info: { id: 9, author: "Reviewer" } },
  } satisfies Paragraph;
  const bookmark = {
    type: "bookmarkStart",
    id: 90,
    name: "_GoBack",
  } satisfies BlockContent;
  const paragraph = {
    type: "paragraph",
    paraId: "10000001",
    content: [],
  } satisfies Paragraph;
  const preceding: Record<string, BlockContent[]> = {
    nothing: [],
    bookmark: [bookmark],
    paragraph: [paragraph],
    "paragraph and bookmark": [paragraph, bookmark],
  };
  for (const [kind, follower] of Object.entries(followers))
    for (const [before, blocks] of Object.entries(preceding))
      expect({
        kind,
        before,
        rejected: reject([...blocks, added, target, follower], [7, 8, 9, 10]),
      }).toEqual({
        kind,
        before,
        rejected: reject([...blocks, target, follower], [7, 10]),
      });
});
