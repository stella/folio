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
            const removable =
              content === "" &&
              tail &&
              (!ownSection || owner === "available" || owner === "same");
            const { pPrMark: _resolved, ...unmarked } = target;
            const expected = {
              package: {
                document: {
                  content: [
                    ...(owner === "unavailable"
                      ? []
                      : [
                          {
                            ...preceding,
                            ...(removable && ownSection
                              ? { sectionProperties: section }
                              : {}),
                          },
                        ]),
                    ...(removable ? [] : [unmarked]),
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
