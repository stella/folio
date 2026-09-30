import { expect, test } from "bun:test";
import type { Document } from "@stll/docx-core/model";

import { reviewDifferences } from "./reviewOracle";

const fixture = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            {
              type: "insertion",
              info: { id: 41, author: "Reviewer", date: "2026-05-06T07:08:09Z" },
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Tracked" }],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

test("the oracle detects lost text and revision identities", () => {
  const original = fixture();
  expect(reviewDifferences(original, structuredClone(original))).toEqual({
    messages: [],
    omitted: 0,
  });
  const changed = structuredClone(original);
  const paragraph = changed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
  const insertion = paragraph.content.at(0);
  if (insertion?.type !== "insertion") throw new Error("Expected insertion");
  insertion.info.id = 42;
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
  insertion.info.id = 41;
  insertion.content = [];
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
});

test("the oracle detects a different paragraph survivor", () => {
  const original = fixture();
  const changed = structuredClone(original);
  const paragraph = changed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
  paragraph.paraId = "00000002";
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
});
