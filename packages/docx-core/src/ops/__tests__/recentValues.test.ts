import { expect, spyOn, test } from "bun:test";

import type { DocxPackage, Paragraph } from "../../model/document";
import { packageParagraphIdIndex } from "../ids";
import { findStoryParagraph } from "../paragraphLocator";
import { recentValues } from "../recentValues";

test("recency retains at most eight values in one weak target", () => {
  const recent = recentValues<object>();
  const versions = Array.from({ length: 24 }, (_, version) => ({ version }));
  const deref = spyOn(WeakRef.prototype, "deref");
  try {
    expect(recent.values()).toEqual([]);
    for (const version of versions) recent.remember(version);
    const newest = versions.slice(-8).reverse();
    expect(recent.values()).toEqual(newest);
    const eighth = newest.at(-1);
    if (eighth === undefined) throw new Error("Expected a full recency list");
    recent.remember(eighth);
    expect(recent.values()).toEqual([eighth, ...newest.slice(0, -1)]);
    expect(new Set(deref.mock.results.map(({ value }) => value)).size).toBe(1);
  } finally {
    deref.mockRestore();
  }
});

test("package and locator indexes dereference only bounded recency lists", () => {
  const sharedId = "10000000";
  const shared: Paragraph = {
    type: "paragraph",
    paraId: sharedId,
    content: [],
  };
  const packages = Array.from(
    { length: 24 },
    (_, index) =>
      ({
        document: {
          content: [
            shared,
            {
              type: "paragraph",
              paraId: (0x20000000 + index).toString(16),
              content: [],
            },
          ],
        },
      }) satisfies DocxPackage,
  );
  const deref = spyOn(WeakRef.prototype, "deref");
  try {
    for (const pkg of packages) {
      expect(packageParagraphIdIndex(pkg).has(sharedId)).toBe(true);
      expect(findStoryParagraph(pkg.document, sharedId)?.paragraph).toBe(
        shared,
      );
    }
    // The module's lists may predate this test: a collection after the job
    // that filled them empties their references, which then read as
    // `undefined` once before a list is made again.
    const targets = new Set(
      deref.mock.results
        .map(({ value }) => value)
        .filter((value) => value !== undefined),
    );
    expect(targets.size).toBe(2);
    for (const target of targets) {
      expect(Array.isArray(target)).toBe(true);
      if (Array.isArray(target)) expect(target.length).toBe(8);
    }
  } finally {
    deref.mockRestore();
  }
});
