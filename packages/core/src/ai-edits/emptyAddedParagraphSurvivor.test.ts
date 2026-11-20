import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  validateOpsDocument,
  withBodyContent,
} from "@stll/docx-core/ops";

import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

test("rejected bare added containers leave one empty carrier and preserve every outside fact through DOCX", async () => {
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
            const expected = applyDocumentOp(before, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: revisions,
              decision,
            }).unwrap().document;
            const assertFacts = (actual: Document) => {
              const owner = actual.package.document.content.at(0);
              if (owner?.type !== "blockSdt")
                return panic("The source control must remain its owner.");
              let expectedBody = expected.package.document;
              if (decision === "reject") {
                const carrier = owner.content.at(0);
                if (
                  carrier?.type !== "paragraph" ||
                  carrier.paraId === undefined
                )
                  return panic(
                    "An exhausted local container requires one identified paragraph.",
                  );
                expect(Object.keys(carrier).toSorted()).toEqual([
                  "content",
                  "paraId",
                  "type",
                ]);
                expect(
                  validateOpsDocument(structuredClone(actual)).isOk(),
                ).toBe(true);
                const sourceOwner = before.package.document.content.at(0);
                if (sourceOwner?.type !== "blockSdt")
                  return panic("The fixture source has one control.");
                // No source mark survives. This route may mint its own filler;
                // all source facts outside that bare carrier remain exact.
                expectedBody = withBodyContent(before.package.document, [
                  {
                    ...sourceOwner,
                    content: [
                      {
                        type: "paragraph",
                        paraId: carrier.paraId,
                        content: [],
                      },
                    ],
                  },
                  ...before.package.document.content.slice(1),
                ]);
              }
              expect(structuredClone(actual.package.document)).toEqual(
                structuredClone(expectedBody),
              );
              expect(owner.content).toHaveLength(
                decision === "reject" ? 1 : length,
              );
              for (const block of owner.content) {
                expect(block.type).toBe("paragraph");
                if (block.type !== "paragraph")
                  return panic("A source paragraph remains.");
                expect(block.content).toEqual([]);
                expect(block.pPrMark).toBeUndefined();
              }
            };
            if (route === "read") {
              assertFacts(reviewer.toResolvedDocument(decision));
              expect(reviewer.toDocument()).toEqual(before);
              continue;
            }
            if (route === "whole") {
              expect(
                decision === "accept"
                  ? reviewer.acceptAll()
                  : reviewer.rejectAll(),
              ).toMatchObject({ resolved: length, skipped: [] });
            } else {
              const order =
                route === "forward" ? revisions : revisions.toReversed();
              for (const [index, id] of order.entries()) {
                expect(
                  decision === "accept"
                    ? reviewer.acceptChange(id)
                    : reviewer.rejectChange(id),
                ).toBe(true);
                expect(
                  reviewer
                    .getChanges()
                    .map((change) => change.id)
                    .toSorted((a, b) => a - b),
                ).toEqual(order.slice(index + 1).toSorted((a, b) => a - b));
              }
            }
            const resolved = reviewer.toDocument();
            assertFacts(resolved);
            expect(reviewer.getChanges()).toEqual([]);
            const reopened = await FolioDocxReviewer.fromBuffer(
              await reviewer.toBuffer(),
              { session },
            );
            assertFacts(reopened.toDocument());
            expect(reopened.getChanges()).toEqual([]);
            // Each engine's chosen carrier keeps its own identity across save/reopen.
            expect(
              JSON.parse(
                JSON.stringify(reopened.toDocument().package.document),
              ),
            ).toEqual(JSON.parse(JSON.stringify(resolved.package.document)));
          }
    }
});
