import { describe, expect, test } from "bun:test";

import type { Document } from "../../types/document";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { parseRunProperties, RUN_PROPERTY_OWNERS } from "../runParser";
import { serializeTextFormatting } from "../serializer/textFormattingSerializer";
import { parseXml } from "../xmlParser";
import type { XmlElement } from "../xmlParser";

const parseRPr = (xml: string): XmlElement => {
  const parsed = parseXml(
    `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</w:rPr>`,
  );
  const rPr = parsed.elements.at(0);
  if (!rPr || rPr.type !== "element") {
    throw new Error("Expected run properties element");
  }
  return rPr;
};

describe("run font hint round-trip", () => {
  for (const hint of ["default", "eastAsia", "cs"] as const) {
    test(`preserves a hint-only ${hint} font declaration`, () => {
      const formatting = parseRunProperties(
        parseRPr(`<w:rFonts w:hint="${hint}"/>`),
        null,
        RUN_PROPERTY_OWNERS.standalone,
      );

      expect(formatting?.fontFamily).toEqual({ hint });
      expect(serializeTextFormatting(formatting)).toContain(`<w:rFonts w:hint="${hint}"/>`);
    });
  }

  test("takes no font from an unknown hint, and keeps the element", () => {
    const formatting = parseRunProperties(
      parseRPr('<w:rFonts w:hint="unsupported"/>'),
      null,
      RUN_PROPERTY_OWNERS.standalone,
    );

    // The hint is outside the enum, so the reader states no font family at all
    // rather than an empty one; the element the author wrote comes back from
    // the verbatim sink instead of being dropped.
    expect(formatting?.fontFamily).toBeUndefined();
    expect(serializeTextFormatting(formatting)).toContain('<w:rFonts w:hint="unsupported"/>');
  });

  test("preserves the hint through the editor model", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  formatting: { fontFamily: { hint: "eastAsia" } },
                  content: [{ type: "text", text: "text" }],
                },
              ],
            },
          ],
        },
      },
    };

    const roundtripped = fromProseDoc(toProseDoc(document), document);
    const paragraph = roundtripped.package.document.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }
    const run = paragraph.content.at(0);
    expect(run?.type).toBe("run");
    if (run?.type !== "run") {
      throw new Error("Expected run");
    }
    expect(run.formatting?.fontFamily?.hint).toBe("eastAsia");
  });
});
