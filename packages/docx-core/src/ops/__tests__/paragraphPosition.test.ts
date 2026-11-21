import { expect, spyOn, test } from "bun:test";

import type { Paragraph } from "../../model/document";
import { spliceItems } from "../listSplice";
import { findStoryParagraph } from "../paragraphLocator";

const paragraph = (id: number): Paragraph => ({
  type: "paragraph",
  paraId: id.toString(16).padStart(8, "0"),
  content: [],
});

test("retained paragraphs shift without searching the story after each splice", () => {
  let content = Array.from({ length: 80 }, (_, i) => paragraph(i));
  const anchor = content[40];
  expect(findStoryParagraph({ content }, "00000028")?.paragraph).toBe(anchor);
  for (let i = 0; i < 64; i++) {
    content = spliceItems({
      items: content,
      start: 1,
      count: 0,
      replacement: [paragraph(100 + i)],
    });
    const search = spyOn(content, "indexOf");
    try {
      expect(findStoryParagraph({ content }, "00000028")?.index).toBe(41 + i);
      expect(search).not.toHaveBeenCalled();
    } finally {
      search.mockRestore();
    }
  }
});

test("removing a repeated reference locates the remaining occurrence", () => {
  const same = paragraph(13);
  const content = [same, paragraph(14), same];
  expect(findStoryParagraph({ content }, "0000000D")?.index).toBe(0);
  const next = spliceItems({
    items: content,
    start: 2,
    count: 1,
    replacement: [],
  });
  expect(findStoryParagraph({ content: next }, "0000000D")?.index).toBe(0);
});
