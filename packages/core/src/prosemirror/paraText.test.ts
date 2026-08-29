import { describe, expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";

import { collectBlockChunks, joinChunks, offsetToDocPos } from "./plugins/pmTextScan";
import { findTextInPmParagraph, getVanillaNodeText, getVanillaTextBetween } from "./paraText";
import { schema } from "./schema";

describe("field text extraction", () => {
  test("exposes leaf and structured field results exactly once", () => {
    const hyperlink = schema.mark("hyperlink", {
      href: "https://example.test/target",
      _docxHyperlinkIndex: 1,
    });
    const structuredField = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: " REF target \\h ",
        displayText: "Target",
        fieldKind: "simple",
      },
      [
        schema.node("bookmarkBoundary", { type: "start", id: 31, name: "target" }, null, [
          hyperlink,
        ]),
        schema.text("Target", [hyperlink]),
        schema.node("bookmarkBoundary", { type: "end", id: 31 }, null, [hyperlink]),
      ],
    );
    const pageField = schema.node("field", {
      fieldType: "PAGE",
      instruction: " PAGE ",
      displayText: "",
      fieldKind: "complex",
    });
    const paragraph = schema.node("paragraph", null, [
      schema.text("Page "),
      pageField,
      schema.text(" / "),
      structuredField,
      schema.text(" after"),
    ]);
    const doc = schema.node("doc", null, [paragraph]);
    const expected = "Page {page} / Target after";

    expect(paragraph.textContent).toBe(expected);
    expect(paragraph.textBetween(0, paragraph.content.size)).toBe(expected);
    const clipboardSlice = new Slice(Fragment.from(paragraph), 0, 0);
    expect(clipboardSlice.content.textBetween(0, clipboardSlice.content.size, "\n\n")).toBe(
      expected,
    );
    expect(getVanillaNodeText(paragraph)).toBe(expected);
    expect(collectBlockChunks(doc).map(joinChunks)).toEqual([expected]);

    for (const searchText of ["{page}", "Target"]) {
      const match = findTextInPmParagraph(doc, 0, doc.content.size, searchText);
      expect(match).not.toBeNull();
      if (match) {
        expect(getVanillaTextBetween(doc, match.from, match.to)).toBe(searchText);
      }
    }
    expect(findTextInPmParagraph(doc, 0, doc.content.size, "page")).toBeNull();
    expect(findTextInPmParagraph(doc, 0, doc.content.size, "{page")).toBeNull();
    expect(findTextInPmParagraph(doc, 0, doc.content.size, "page}")).toBeNull();

    const chunks = collectBlockChunks(doc).at(0) ?? [];
    const pageOffset = joinChunks(chunks).indexOf("{page}");
    const pageFrom = offsetToDocPos(chunks, pageOffset);
    const pageTo = offsetToDocPos(chunks, pageOffset + "{page}".length, "end");
    expect(doc.nodeAt(pageFrom)?.type.name).toBe("field");
    expect(pageTo).toBe(pageFrom + pageField.nodeSize);
  });
});
