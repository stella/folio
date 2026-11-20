import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";

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
  for (const session of ["legacy", "canonical"] as const)
    test(`${session}: actual insertion before ${kind} rejects to its source`, async () => {
      const document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "10000001",
                content: [
                  { type: "run", content: [{ type: "text", text: "before" }] },
                ],
              },
              {
                type: "paragraph",
                paraId: "10000002",
                formatting: { indentLeft: 720 },
                content: [
                  { type: "run", content: [{ type: "text", text: "anchor" }] },
                ],
              },
              follower,
              { type: "paragraph", paraId: "10000004", content: [] },
            ],
          },
        },
      } satisfies Document;
      const input = (
        await prepareCanonicalDocxInput(await createDocx(document))
      )
        .unwrap()
        .slice().buffer;
      for (const text of ["", "new\nmore"])
        for (const route of ["individual", "all"] as const) {
          const reviewer = await FolioDocxReviewer.fromBuffer(input.slice(0), {
            session,
          });
          const before = reviewer.toDocument().package.document;
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [
                {
                  id: "insert",
                  type: "insertAfterBlock",
                  blockId: "10000002",
                  text,
                },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied.map(({ id }) => id)).toEqual(["insert"]);
          expect(result.skipped).toEqual([]);
          expect(reviewer.getChanges().length).toBeGreaterThan(0);
          if (route === "all") reviewer.rejectAll();
          else
            for (const change of reviewer.getChanges().toReversed())
              expect(reviewer.rejectChange(change.id)).toBe(true);
          expect(reviewer.getChanges()).toEqual([]);
          expect(reviewer.toDocument().package.document).toEqual(before);
          const saved = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
            { session },
          );
          expect(saved.getChanges()).toEqual([]);
          expect(saved.toDocument().package.document).toEqual(before);
        }
    });

for (const session of ["legacy", "canonical"] as const)
  test(`${session}: removing an empty added mark before an opaque block preserves its section owner`, async () => {
    for (const owner of [
      "unavailable",
      "available",
      "same",
      "occupied",
    ] as const)
      for (const tail of [false, true]) {
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
          content: [],
          sectionProperties: section,
          pPrMark: { kind: "ins", info: { id: 7, author: "Reviewer" } },
        } satisfies Paragraph;
        const ending = tail
          ? [{ type: "paragraph", paraId: "10000004", content: [] } as const]
          : [];
        const document = {
          package: {
            document: {
              content: [
                ...(owner === "unavailable" ? [] : [preceding]),
                target,
                followers.preservedBlock,
                ...ending,
              ],
              finalSectionProperties: { pageWidth: 15840 },
            },
          },
        } satisfies Document;
        const input = (
          await prepareCanonicalDocxInput(await createDocx(document))
        )
          .unwrap()
          .slice().buffer;
        for (const route of ["individual", "all"] as const) {
          const reviewer = await FolioDocxReviewer.fromBuffer(input.slice(0), {
            session,
          });
          const before = reviewer.toDocument().package.document;
          const removable = tail && (owner === "available" || owner === "same");
          const expectedContent = before.content.flatMap((block) => {
            if (block.type !== "paragraph") return [block];
            if (block.paraId === "10000002") {
              const { pPrMark: _resolved, ...unmarked } = block;
              return removable ? [] : [unmarked];
            }
            return [
              removable && block.paraId === "10000001"
                ? { ...block, sectionProperties: section }
                : block,
            ];
          });
          // Parse the independently stated resolved OOXML, including its
          // section grouping, rather than retaining the source's cached view.
          const expected = (
            await FolioDocxReviewer.fromBuffer(
              await createDocx({
                package: {
                  document: {
                    content: expectedContent,
                    finalSectionProperties: before.finalSectionProperties,
                  },
                },
              }),
              { session: "legacy" },
            )
          ).toDocument().package.document;
          if (route === "all") reviewer.rejectAll();
          else expect(reviewer.rejectChange(7)).toBe(true);
          // Capture/baseline symbols describe each reader's source, rather
          // than the document's section and content facts.
          expect(
            structuredClone(reviewer.toDocument().package.document),
          ).toEqual(structuredClone(expected));
          const saved = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
            { session },
          );
          expect(structuredClone(saved.toDocument().package.document)).toEqual(
            structuredClone(expected),
          );
          expect(saved.getChanges()).toEqual([]);
        }
      }
  });
