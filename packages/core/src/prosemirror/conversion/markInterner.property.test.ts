import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { Schema } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../types/document";
import { createMarkInterner } from "./markInterner";
import { toProseDoc } from "./toProseDoc";

const testSchema = new Schema({
  nodes: { doc: { content: "text*" }, text: {} },
  marks: { probe: { attrs: { a: { default: null }, b: { default: null } } } },
});

/**
 * Values JSON collapses (`-0`/`0`, `NaN`/`null`, `undefined`/absent, key
 * order), drawn from a small domain so distinct-but-similar pairs are common.
 */
const leaf = fc.constantFrom(0, -0, Number.NaN, null, undefined, "0");
const attrValue = fc.oneof(
  leaf,
  fc.array(leaf, { maxLength: 2 }),
  fc.dictionary(fc.constantFrom("k", "l"), leaf, { maxKeys: 2 }),
);

const attrs = fc.record({ a: attrValue, b: attrValue }, { requiredKeys: [] });

/** Exact structural identity: `Object.is` leaves, own-key order, `undefined` vs absent. */
const sameValue = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) {
    return true;
  }
  if (typeof left !== "object" || typeof right !== "object" || left === null || right === null) {
    return false;
  }
  if (Array.isArray(left) !== Array.isArray(right)) {
    return false;
  }
  const leftKeys = Reflect.ownKeys(left);
  const rightKeys = Reflect.ownKeys(right);
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key, index) =>
        key === rightKeys[index] && sameValue(Reflect.get(left, key), Reflect.get(right, key)),
    )
  );
};

describe("createMarkInterner", () => {
  test(
    "shares an instance only between marks with identical attrs",
    () => {
      fc.assert(
        fc.property(fc.clone(attrs, 2), attrs, ([left, leftCopy], right) => {
          const createMark = createMarkInterner(testSchema);
          const leftMark = createMark("probe", left);
          const rightMark = createMark("probe", right);
          const built = [testSchema.mark("probe", left), testSchema.mark("probe", right)] as const;
          expect(sameValue(leftMark.attrs, built[0].attrs)).toBe(true);
          expect(sameValue(rightMark.attrs, built[1].attrs)).toBe(true);
          if (leftMark === rightMark) {
            expect(sameValue(built[0].attrs, built[1].attrs)).toBe(true);
          }
          expect(createMark("probe", leftCopy)).toBe(leftMark);
        }),
        propertyConfig({ numRuns: 2000 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test("scopes shared marks to one conversion", () => {
    const paragraph = (text: string): Paragraph => ({
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text }], formatting: { fontSize: 24 } }],
    });
    const document: Document = {
      package: { document: { content: [paragraph("one"), paragraph("two")] } },
    };
    const fontSizeMarks = (pmDoc: ReturnType<typeof toProseDoc>) => {
      const marks = pmDoc.content.content.map((block) =>
        block.firstChild?.marks.find(({ type }) => type.name === "fontSize"),
      );
      expect(marks.every((mark) => mark !== undefined)).toBe(true);
      return marks;
    };

    const first = fontSizeMarks(toProseDoc(document));
    const second = fontSizeMarks(toProseDoc(document));

    expect(first[0]).toBe(first[1]);
    expect(second[0]).toBe(second[1]);
    expect(second[0]).not.toBe(first[0]);
    expect(second[0]?.eq(first[0]!)).toBe(true);
  });
});
