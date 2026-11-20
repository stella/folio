import { expect, spyOn, test } from "bun:test";

import type { Paragraph } from "../../model/document";
import { findStoryParagraph } from "../paragraphLocator";

const paragraph = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [],
});

const hex = (value: number) =>
  value.toString(16).toUpperCase().padStart(8, "0");

test.each([
  ["dense", (index: number) => hex(index)],
  ["sparse", (index: number) => hex(Math.imul(index, 0x9e3779b9) >>> 0)],
] as const)(
  "settling a %s index copies a bounded number of real keys",
  (_distribution, id) => {
    for (const width of [2048, 8192]) {
      const target = paragraph(id(0));
      const nextTarget = paragraph(id(0));
      const last = paragraph(id(width - 1));
      let content = Array.from({ length: width }, (_, index) =>
        paragraph(id(index)),
      );
      content[0] = target;
      content[width - 1] = last;
      const source = { content };
      expect(findStoryParagraph(source, id(0))?.paragraph).toBe(target);
      for (let index = 1; index < 32; index += 1) {
        const replacement = index % 2 === 0 ? target : nextTarget;
        content = content.toSpliced(0, 1, replacement);
        expect(findStoryParagraph({ content }, id(0))?.paragraph).toBe(
          replacement,
        );
      }
      content = content.toSpliced(0, 1);
      const writes = spyOn(Map.prototype, "set");
      let copied = 0;
      let missing;
      try {
        missing = findStoryParagraph({ content }, id(0));
        copied = writes.mock.calls.length;
      } finally {
        writes.mockRestore();
      }
      expect(missing).toBeUndefined();
      expect(copied).toBeGreaterThan(0);
      expect(copied).toBeLessThanOrEqual(4 * 256 + 16);
      expect(findStoryParagraph({ content }, id(width - 1))?.paragraph).toBe(
        last,
      );
      expect(findStoryParagraph(source, id(0))?.paragraph).toBe(target);
    }
  },
);

test("non-hex names keep case folding, first occurrence and removals across settlement", () => {
  const first = paragraph("not-hex");
  const second = paragraph("NOT-HEX");
  const unrelated = paragraph("other");
  const hexName = paragraph("00000001");
  const suffixedName = paragraph("00000001\n");
  const source = { content: [first, second, unrelated, hexName, suffixedName] };
  expect(findStoryParagraph(source, "Not-Hex")?.paragraph).toBe(first);
  let content = source.content.toSpliced(0, 1);
  expect(findStoryParagraph({ content }, "Not-Hex")?.paragraph).toBe(second);
  for (let index = 1; index < 31; index += 1) {
    const changed = paragraph("other");
    content = content.toSpliced(1, 1, changed);
    expect(findStoryParagraph({ content }, "Other")?.paragraph).toBe(changed);
  }
  content = content.toSpliced(0, 1);
  expect(findStoryParagraph({ content }, "Not-Hex")).toBeUndefined();
  expect(findStoryParagraph(source, "Not-Hex")?.paragraph).toBe(first);
  expect(findStoryParagraph({ content }, "00000001")?.paragraph).toBe(hexName);
  expect(findStoryParagraph({ content }, "00000001\n")?.paragraph).toBe(
    suffixedName,
  );
});
