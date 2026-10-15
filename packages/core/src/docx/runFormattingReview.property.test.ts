import { expect, test } from "bun:test";
import fc from "fast-check";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  PROPERTY_REVIEW_POLICIES,
  revisionIdDemand,
} from "@stll/docx-core/ops";
import type { SetRunPropsOp } from "@stll/docx-core/ops";
import type { Document } from "@stll/docx-core/model";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

const runsOf = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
  return paragraph.content.flatMap((entry) => (entry.type === "run" ? [entry] : []));
};

// Formatting generators did not combine multiple authors with a persisted overlapping range.
test(
  "overlapping formatting has one latest-author revision and the first baseline",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 3, max: 7 }),
        fc.integer({ min: 3, max: 9 }),
        fc.boolean(),
        fc.boolean(),
        async (firstEnd, secondEnd, hasBaseline, hasEmptyBaseline) => {
          const populatedBaseline = hasBaseline ? { fontSize: 22, bold: false } : undefined;
          const baseline = populatedBaseline ?? (hasEmptyBaseline ? {} : undefined);
          const original = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "12345678",
                    content: [
                      {
                        type: "run",
                        ...(baseline === undefined ? {} : { formatting: baseline }),
                        content: [{ type: "text", text: "abcdefghi" }],
                      },
                    ],
                  },
                ],
              },
            },
          } satisfies Document;
          let document: Document = original;
          const at = (offset: number) => ({ story: OP_STORIES.MAIN, blockId: "12345678", offset });
          const patches = [
            { from: 0, to: firstEnd, patch: { bold: true }, author: "A", id: 1 },
            { from: 2, to: secondEnd, patch: { italic: true }, author: "B", id: 2 },
          ];
          for (const { from, to, patch, author, id } of patches) {
            const operation = {
              type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
              from: at(from),
              to: at(to),
              patch,
              revision: { id, author, date: `2000-01-0${id}T00:00:00Z` },
              propertyReview: PROPERTY_REVIEW_POLICIES.APPEND,
            } satisfies SetRunPropsOp;
            const count = revisionIdDemand(document, operation).unwrap();
            document = applyDocumentOp(document, {
              ...operation,
              newIds: { revision: Array.from({ length: count }, (_, index) => 100 + index) },
            }).unwrap().document;
          }
          let offset = 0;
          for (const run of runsOf(document)) {
            expect(run.propertyChanges?.length ?? 0).toBe(
              offset < firstEnd || (offset >= 2 && offset < secondEnd) ? 1 : 0,
            );
            const pending = run.propertyChanges?.at(0);
            const startOffset = offset;
            const latest = patches.findLast(
              ({ from, to }) => startOffset >= from && startOffset < to,
            );
            expect(pending?.info.author).toBe(latest?.author);
            expect(pending?.info.date).toBe(
              latest === undefined ? undefined : `2000-01-0${latest.id}T00:00:00Z`,
            );
            expect(pending?.previousFormatting).toEqual(pending ? baseline : undefined);
            const text = run.content
              .flatMap((leaf) => (leaf.type === "text" ? [leaf.text] : []))
              .join("");
            for (const _character of text) {
              const expected = {
                ...(hasBaseline ? { fontSize: 22, bold: false } : {}),
                ...(offset < firstEnd ? { bold: true } : {}),
                ...(offset >= 2 && offset < secondEnd ? { italic: true } : {}),
              };
              expect(run.formatting).toEqual(
                Object.keys(expected).length > 0 ? expected : baseline,
              );
              offset++;
            }
          }
          expect(offset).toBe(9);
          const saved = await parseDocx(await createDocx(document), { preloadFonts: false });
          expect(runsOf(saved).map(({ propertyChanges }) => propertyChanges?.length)).toEqual(
            runsOf(document).map(({ propertyChanges }) => propertyChanges?.length),
          );
          for (const decision of ["accept", "reject"] as const) {
            for (const source of [document, saved]) {
              const revisionIds = [
                ...new Set(
                  runsOf(source).flatMap(
                    ({ propertyChanges }) => propertyChanges?.map(({ info }) => info.id) ?? [],
                  ),
                ),
              ];
              const resolved = applyDocumentOp(source, {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds,
                decision,
              }).unwrap().document;
              expect(
                runsOf(resolved).every(({ propertyChanges }) => !propertyChanges?.length),
              ).toBe(true);
              if (decision === "reject") {
                for (const run of runsOf(resolved)) {
                  expect(run.formatting).toEqual(baseline);
                }
              } else {
                expect(runsOf(resolved).map(({ formatting }) => formatting)).toEqual(
                  runsOf(source).map(({ formatting }) => formatting),
                );
              }
            }
          }
        },
      ),
      {
        seed: 1404002,
        numRuns: 12,
        id: "overlapping formatter actions retain one earliest baseline and latest author",
      },
    );
  },
  propertyTestTimeout(30_000),
);
