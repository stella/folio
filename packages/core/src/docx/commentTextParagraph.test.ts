import { expect, test } from "bun:test";

import { commentTextParagraph } from "./commentTextParagraph";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

test.each(["Plain", "A & B", " ", "A👍🏽B"])(
  "plain comment paragraphs retain authored properties across save: %j",
  (text) => {
    const source = commentTextParagraph(text);
    const xml = parseXmlDocument(
      `<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${serializeParagraph(source)}</w:root>`,
    );
    const element = xml?.elements?.find((child) => child.type === "element");
    if (!element || element.type !== "element")
      throw new Error("Expected paragraph");
    const saved = parseParagraph(element, null, null, null);
    expect(saved.formatting).toBeUndefined();
    expect(saved.content).toEqual(source.content);
  },
);
