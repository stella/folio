import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";

describe("renderedPageBreakBefore round-trip", () => {
  test("attr survives PM to Document to XML", () => {
    const paragraph = schema.node("paragraph", { renderedPageBreakBefore: true }, [
      schema.text("Attachment 1"),
    ]);
    const doc = schema.node("doc", null, [paragraph]);

    const document = fromProseDoc(doc);
    const parsed = document.package.document.content.at(0);
    if (parsed?.type !== "paragraph") {
      panic("Expected the converted paragraph");
    }
    expect(parsed.renderedPageBreakBefore).toBe(true);

    const xml = serializeParagraph(parsed);
    expect(xml).toMatch(/<w:lastRenderedPageBreak\/>/u);
    expect(xml).toMatch(/<w:r[^>]*>(?:<w:rPr>.*<\/w:rPr>)?<w:lastRenderedPageBreak\/>/u);
    expect(xml.match(/<w:lastRenderedPageBreak\/>/gu)).toHaveLength(1);
  });

  test("injects the marker after run properties, including a nested previous-property snapshot", () => {
    const xml = serializeParagraph({
      type: "paragraph",
      renderedPageBreakBefore: true,
      content: [
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [
            {
              info: { id: 1, author: "Reviewer" },
              previousFormatting: { italic: true },
            },
          ],
          content: [{ type: "text", text: "Attachment 1" }],
        },
      ],
    });

    expect(xml).toMatch(
      /<w:r><w:rPr><w:b\/><w:rPrChange [^>]+><w:rPr><w:i\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:lastRenderedPageBreak\/><w:t>/u,
    );
  });

  test("preserves an inline cached boundary at its editable position", () => {
    const paragraph = schema.node("paragraph", null, [
      schema.text("Previous page"),
      schema.node("renderedPageBreak"),
      schema.text("Next page"),
    ]);
    const document = fromProseDoc(schema.node("doc", null, [paragraph]));
    const parsed = document.package.document.content.at(0);

    expect(parsed).toMatchObject({
      type: "paragraph",
      content: [
        { type: "run", content: [{ type: "text", text: "Previous page" }] },
        { type: "run", content: [{ type: "renderedPageBreak" }] },
        { type: "run", content: [{ type: "text", text: "Next page" }] },
      ],
    });
    expect(serializeParagraph(parsed).match(/<w:lastRenderedPageBreak\/>/gu)).toHaveLength(1);
  });

  test("serializer injects marker into the first run inside a hyperlink wrapper", () => {
    const paragraph = {
      type: "paragraph" as const,
      renderedPageBreakBefore: true,
      content: [
        {
          type: "hyperlink" as const,
          href: "https://example.com",
          children: [
            {
              type: "run" as const,
              content: [{ type: "text" as const, text: "link" }],
            },
          ],
        },
      ],
    };

    const xml = serializeParagraph(paragraph as never);
    expect(xml).toMatch(/<w:hyperlink[^>]*>[^<]*<w:r[^>]*><w:lastRenderedPageBreak\/>/u);
  });
});
