import { describe, expect, test } from "bun:test";

import { readRunIdentityMarkAttrs } from "../../attrs";
import { schema } from "../../schema";

const parseRunIdentity = (rawId: string | undefined) => {
  const getAttrs = schema.marks.runIdentity.spec.parseDOM?.at(0)?.getAttrs;
  if (!getAttrs) {
    throw new Error("RunIdentityExtension must define parseDOM[0].getAttrs");
  }
  const dataset = rawId === undefined ? {} : { docxRunIdentity: rawId };
  // SAFETY: getAttrs reads only HTMLElement.dataset in this test fixture.
  return getAttrs({ dataset } as unknown as HTMLElement);
};

describe("run identity", () => {
  test.each([
    ["0", 0],
    ["1", 1],
    [String(Number.MAX_SAFE_INTEGER), Number.MAX_SAFE_INTEGER],
  ] as const)("parses canonical safe identity %p", (rawId, id) => {
    expect(parseRunIdentity(rawId)).toEqual({ id });
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
  ] as const)("rejects malformed or unsafe identity %p", (rawId) => {
    expect(parseRunIdentity(rawId)).toBe(false);
  });

  test("does not synthesize identity zero when an ID is omitted", () => {
    const mark = schema.marks.runIdentity.create();
    const result = readRunIdentityMarkAttrs(mark);

    expect(mark.attrs["id"]).not.toBe(0);
    expect(result.ok).toBe(false);
  });

  test.each([-1, 1.5, Number.MAX_SAFE_INTEGER + 1])("rejects non-safe model identity %p", (id) => {
    const result = readRunIdentityMarkAttrs(schema.marks.runIdentity.create({ id }));

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.issues.map((issue) => issue.path)).toContain("runIdentity.attrs.id");
    }
  });

  test("refuses a remainder entry that is not a resolved name and a string value", () => {
    const result = readRunIdentityMarkAttrs(
      schema.marks.runIdentity.create({
        id: 1,
        preservedAttributes: [{ name: "rsidR", value: 7 }],
      }),
    );

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.issues.map((issue) => issue.path)).toContain(
        "runIdentity.attrs.preservedAttributes[0].value",
      );
    }
  });

  test("refuses a sink child with no ordinal to put it back at", () => {
    const result = readRunIdentityMarkAttrs(
      schema.marks.runIdentity.create({
        id: 1,
        preserved: { children: [{ xml: "<w:u/>" }] },
      }),
    );

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.issues.map((issue) => issue.path)).toContain(
        "runIdentity.attrs.preserved.children[0].index",
      );
    }
  });

  test("the DOM carries the id and nothing else", () => {
    const mark = schema.marks.runIdentity.create({
      id: 4,
      preservedAttributes: [{ name: "rsidR", value: "00ABCDEF" }],
    });
    const spec = schema.marks.runIdentity.spec.toDOM?.(mark, false);

    expect(spec).toEqual(["span", { "data-docx-run-identity": "4" }, 0]);
  });
});
