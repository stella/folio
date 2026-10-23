import { expect, test } from "bun:test";
import fc from "fast-check";

import { canonicalReviewModel } from "../../../../test/consumer-scenarios/support/review-comparison";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import type { Document, Hyperlink } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const linksOf = (document: Document) => {
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph")
    throw new Error("Expected hyperlink paragraph");
  return first.content.filter((node) => node.type === "hyperlink");
};
const link = (
  text: string,
  href = "https://example.com/notices",
): Hyperlink => ({
  type: "hyperlink",
  href,
  children: [{ type: "run", content: [{ type: "text", text }] }],
});

test(
  "generated rejected partial edits retain one hyperlink through save and reopen",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("notice", "éclair", "合同", "😀"),
        fc.constantFrom("tip", "contains & <tag>", "Unicode 合同", ""),
        async (token, tooltip) => {
          const document = createEmptyDocument();
          document.package.document.content = [
            {
              type: "paragraph",
              paraId: "A1000011",
              content: [
                {
                  ...link(`the ${token} address`),
                  tooltip,
                  history: false,
                  target: "_blank",
                  docLocation: "location",
                },
              ],
            },
          ];
          const reviewer = await FolioDocxReviewer.fromBuffer(
            await createDocx(document),
          );
          const block = reviewer.snapshot().blocks.at(0);
          if (!block) throw new Error("Expected hyperlink block");
          const result = reviewer.applyDocumentOperations(
            {
              version: 1,
              mode: "tracked-changes",
              operations: [
                {
                  id: "replace",
                  type: "replaceInBlock",
                  blockId: block.id,
                  find: token,
                  replace: "changed",
                },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.skipped).toEqual([]);
          expect(result.applied).toHaveLength(1);
          const reopened = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
          );
          const live = reviewer.toResolvedDocument("reject");
          const saved = reopened.toResolvedDocument("reject");
          expect(linksOf(live)).toHaveLength(1);
          expect(linksOf(saved)).toHaveLength(1);
          expect(canonicalReviewModel(saved)).toEqual(
            canonicalReviewModel(live),
          );
        },
      ),
      {
        id: "generated rejected partial edits retain one hyperlink through save and reopen",
        numRuns: 12,
        seed: 1683012,
      },
    );
  },
  propertyTestTimeout(10_000),
);

test("adjacent identical imported hyperlinks stay distinct through save and reopen", async () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000012",
      content: [link("first"), link("second")],
    },
  ];
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(document),
  );
  expect(linksOf(reviewer.toDocument())).toHaveLength(2);
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(linksOf(reopened.toDocument())).toHaveLength(2);
});

test("accepting a changed hyperlink target keeps different targets split", async () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000013",
      content: [
        link("left"),
        {
          type: "insertion",
          info: { id: 17, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
          content: [link("changed", "https://example.com/other")],
        },
        link("right"),
      ],
    },
  ];
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(document),
  );
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  for (const current of [reviewer, reopened]) {
    expect(
      linksOf(current.toResolvedDocument("accept")).map(({ href }) => href),
    ).toEqual([
      "https://example.com/notices",
      "https://example.com/other",
      "https://example.com/notices",
    ]);
  }
});
