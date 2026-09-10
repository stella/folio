import { describe, expect, test } from "bun:test";

import { readPageBreakRunOwnerMarkAttrs } from "../../attrs";
import { schema } from "../../schema";

const parseOwnerId = (rawId: string | undefined) => {
  const getAttrs = schema.marks.pageBreakRunOwner.spec.parseDOM?.at(0)?.getAttrs;
  if (!getAttrs) {
    throw new Error("PageBreakRunOwnerExtension must define parseDOM[0].getAttrs");
  }
  const dataset = rawId === undefined ? {} : { docxPageBreakRunOwner: rawId };
  // SAFETY: getAttrs reads only HTMLElement.dataset in this test fixture.
  return getAttrs({ dataset } as unknown as HTMLElement);
};

describe("page-break run owner identity", () => {
  test.each([
    ["0", 0],
    ["1", 1],
    [String(Number.MAX_SAFE_INTEGER), Number.MAX_SAFE_INTEGER],
  ] as const)("parses canonical safe owner ID %p", (rawId, id) => {
    expect(parseOwnerId(rawId)).toEqual({ id });
  });

  test.each([
    undefined,
    "",
    "00",
    "01",
    "-1",
    "+1",
    "1.0",
    "1junk",
    " 1 ",
    String(Number.MAX_SAFE_INTEGER + 1),
  ] as const)("rejects malformed or unsafe owner ID %p", (rawId) => {
    expect(parseOwnerId(rawId)).toBe(false);
  });

  test("does not synthesize owner zero when an ID is omitted", () => {
    const mark = schema.marks.pageBreakRunOwner.create();
    const result = readPageBreakRunOwnerMarkAttrs(mark);

    expect(mark.attrs["id"]).not.toBe(0);
    expect(result.ok).toBe(false);
  });

  test.each([-1, 1.5, Number.MAX_SAFE_INTEGER + 1])("rejects non-safe model owner ID %p", (id) => {
    const result = readPageBreakRunOwnerMarkAttrs(schema.marks.pageBreakRunOwner.create({ id }));

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.issues.map((issue) => issue.path)).toContain("pageBreakRunOwner.attrs.id");
    }
  });
});
