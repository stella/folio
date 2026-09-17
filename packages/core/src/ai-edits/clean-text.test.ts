import { describe, expect, test } from "bun:test";
import { type Node as PMNode, Schema } from "prosemirror-model";

import { schema as folioSchema } from "../prosemirror/schema";
import { buildCleanBlockText, resolveCleanTextRange } from "./clean-text";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "inline*", group: "block" },
    pageBreakRun: {
      attrs: { clear: { default: undefined } },
      atom: true,
      group: "inline",
      inline: true,
    },
    text: { group: "inline" },
  },
});

const paragraph = (content: PMNode | readonly PMNode[]) => schema.node("paragraph", null, content);

describe("buildCleanBlockText", () => {
  test("shares one frozen empty boundary projection across ordinary blocks", () => {
    const first = buildCleanBlockText(paragraph(schema.text("first")), 0);
    const second = buildCleanBlockText(paragraph(schema.text("second")), 10);

    expect(first.structuralBoundaries).toBe(second.structuralBoundaries);
    expect(first.structuralBoundaries).toHaveLength(0);
    expect(Object.isFrozen(first.structuralBoundaries)).toBe(true);
  });

  test("allocates on the first carrier and preserves every carrier in document order", () => {
    const cleanBlock = buildCleanBlockText(
      paragraph([
        schema.text("A"),
        schema.node("pageBreakRun", { clear: "left" }),
        schema.node("pageBreakRun", { clear: "right" }),
        schema.text("B"),
      ]),
      0,
    );

    expect(cleanBlock).toEqual({
      text: "AB",
      offsets: [1, 4, 5],
      structuralBoundaries: [
        {
          type: "pageBreakRun",
          offset: 1,
          from: 2,
          to: 3,
          clear: "left",
          presentInCleanView: true,
        },
        {
          type: "pageBreakRun",
          offset: 1,
          from: 3,
          to: 4,
          clear: "right",
          presentInCleanView: true,
        },
      ],
    });
  });
});

describe("resolveCleanTextRange", () => {
  test("maps the common carrierless path directly through clean offsets", () => {
    const cleanBlock = buildCleanBlockText(paragraph(schema.text("AB")), 10);

    expect(resolveCleanTextRange({ cleanBlock, startOffset: 0, endOffset: 2 })).toEqual({
      from: 11,
      to: 13,
    });
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 1, endOffset: 1 })).toEqual({
      from: 12,
      to: 12,
    });
    expect(resolveCleanTextRange({ cleanBlock, startOffset: -1, endOffset: 1 })).toBeNull();
  });

  test("biases either side of consecutive carriers in one boundary pass", () => {
    const cleanBlock = buildCleanBlockText(
      paragraph([
        schema.text("A"),
        schema.node("pageBreakRun", { clear: "left" }),
        schema.node("pageBreakRun", { clear: "right" }),
        schema.text("B"),
      ]),
      0,
    );

    expect(resolveCleanTextRange({ cleanBlock, startOffset: 0, endOffset: 1 })).toEqual({
      from: 1,
      to: 2,
    });
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 1, endOffset: 2 })).toEqual({
      from: 4,
      to: 5,
    });
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 1, endOffset: 1 })).toEqual({
      from: 4,
      to: 4,
    });
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 0, endOffset: 2 })).toBeNull();
  });
});

describe("field results in the clean projection", () => {
  const fieldParagraph = (): PMNode =>
    folioSchema.node("paragraph", null, [
      folioSchema.text("see Clause "),
      folioSchema.node("field", {
        fieldType: "REF",
        instruction: " REF _Ref1 ",
        displayText: "3.6(a)",
        fieldKind: "complex",
      }),
      folioSchema.text(" above"),
    ]);

  test("reads the field as its result and anchors every character at the atom", () => {
    const cleanBlock = buildCleanBlockText(fieldParagraph(), 0);

    expect(cleanBlock.text).toBe("see Clause 3.6(a) above");
    expect(cleanBlock.structuralBoundaries).toEqual([
      { type: "field", offset: 11, length: 6, from: 12, to: 13 },
    ]);
    expect(cleanBlock.offsets.slice(11, 18)).toEqual([12, 12, 12, 12, 12, 12, 13]);
  });

  test("omits the result when the caller asks for an alignment coordinate", () => {
    expect(buildCleanBlockText(fieldParagraph(), 0, { fieldResults: "omitted" })).toMatchObject({
      text: "see Clause  above",
      structuralBoundaries: [],
    });
  });

  test("refuses a range that cuts into the result and allows one around it", () => {
    const cleanBlock = buildCleanBlockText(fieldParagraph(), 0);

    expect(resolveCleanTextRange({ cleanBlock, startOffset: 13, endOffset: 20 })).toBeNull();
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 0, endOffset: 13 })).toBeNull();
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 11, endOffset: 17 })).toEqual({
      from: 12,
      to: 13,
    });
    expect(resolveCleanTextRange({ cleanBlock, startOffset: 0, endOffset: 11 })).toEqual({
      from: 1,
      to: 12,
    });
  });
});
