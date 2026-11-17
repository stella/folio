import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

test("empty added paragraphs inside a standing control keep exact owners through review and DOCX", async () => {
  for (const length of [2, 3])
    for (const kind of ["ins", "moveTo"] as const) {
      const source = {
        package: {
          document: {
            content: [
              {
                type: "blockSdt",
                properties: { sdtType: "group", id: 1 },
                content: Array.from({ length }, (_, index) => ({
                  type: "paragraph",
                  paraId: (index + 1).toString(16).padStart(8, "0"),
                  content: [],
                  pPrMark: {
                    kind,
                    info: { id: index + 1, author: "Reviewer" },
                  },
                })),
              },
              { type: "paragraph", paraId: "10000000", content: [] },
            ],
          },
        },
      } satisfies Document;
      const input = (await prepareCanonicalDocxInput(await createDocx(source)))
        .unwrap()
        .slice().buffer;
      const revisions = Array.from({ length }, (_, index) => index + 1);
      for (const session of ["legacy", "canonical"] as const)
        for (const decision of ["accept", "reject"] as const)
          for (const route of [
            "read",
            "whole",
            "forward",
            "reverse",
          ] as const) {
            const reviewer = await FolioDocxReviewer.fromBuffer(
              input.slice(0),
              { session },
            );
            const before = reviewer.toDocument();
            let expected = before;
            const resolve = (revisionIds: number[]) => {
              expected = applyDocumentOp(expected, {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds,
                decision,
              }).unwrap().document;
            };
            const assertFacts = (actual: Document) => {
              // Private capture symbols describe each reader's source; all
              // document fields, paragraph identities and sections are compared.
              expect(structuredClone(actual.package.document)).toEqual(
                structuredClone(expected.package.document),
              );
            };
            if (route === "read") {
              resolve(revisions);
              assertFacts(reviewer.toResolvedDocument(decision));
              expect(reviewer.toDocument()).toEqual(before);
              continue;
            }
            if (route === "whole") {
              resolve(revisions);
              expect(
                decision === "accept"
                  ? reviewer.acceptAll()
                  : reviewer.rejectAll(),
              ).toEqual({ resolved: length, skipped: [] });
              const control = reviewer
                .toDocument()
                .package.document.content.at(0);
              expect(
                control?.type === "blockSdt" &&
                  control.content.map((block) =>
                    block.type === "paragraph" ? block.paraId : block.type,
                  ),
              ).toEqual(
                decision === "accept"
                  ? revisions.map((id) => id.toString(16).padStart(8, "0"))
                  : [length.toString(16).padStart(8, "0")],
              );
            } else {
              for (const id of route === "forward"
                ? revisions
                : revisions.toReversed()) {
                resolve([id]);
                expect(
                  decision === "accept"
                    ? reviewer.acceptChange(id)
                    : reviewer.rejectChange(id),
                ).toBe(true);
                assertFacts(reviewer.toDocument());
              }
            }
            assertFacts(reviewer.toDocument());
            expect(reviewer.getChanges()).toEqual([]);
            const reopened = await FolioDocxReviewer.fromBuffer(
              await reviewer.toBuffer(),
              { session },
            );
            assertFacts(reopened.toDocument());
            expect(reopened.getChanges()).toEqual([]);
          }
    }
});
