/**
 * A header or footer is written with the blocks it was read with.
 *
 * The rebuild path synthesised `<w:p><w:pPr/></w:p>` whenever the serialized
 * content came out empty, under the premise that OOXML requires a block child.
 * It does not: `CT_HdrFtr` holds a single `EG_BlockLevelElts` occurrence whose
 * choice members are all `minOccurs="0"`, so `<w:hdr/>` is valid, and Word
 * writes exactly that for a header the author left blank. Verbatim replay
 * returned such a part unchanged, so the invented line appeared only once the
 * document had been edited.
 *
 * The property generates bodies rather than one fixture: the block count is the
 * variable, and zero is only the smallest of its values.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../../test/property-testing";

import { parseFooter, parseHeader } from "../headerFooterParser";
import { clearHeaderFooterVerbatimXml } from "../headerFooterVerbatim";
import { serializeHeaderFooter } from "./headerFooterSerializer";

const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The paragraph shapes an author's header actually holds. */
const PARAGRAPH_FORMS = {
  /** `<w:p/>`: a blank line. */
  bare: "<w:p/>",
  /** Properties, no content: a blank line that carries formatting. */
  propertiesOnly: '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>',
  /** A run with no text: still a block the author wrote. */
  emptyRun: "<w:p><w:r></w:r></w:p>",
  text: "<w:p><w:r><w:t>Chapter</w:t></w:r></w:p>",
} as const;

type ParagraphForm = keyof typeof PARAGRAPH_FORMS;

const partXml = (root: "hdr" | "ftr", forms: readonly ParagraphForm[]): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:${root} xmlns:w="${W_NAMESPACE}">` +
  `${forms.map((form) => PARAGRAPH_FORMS[form]).join("")}</w:${root}>`;

const body = fc.array(
  fc.constantFrom<ParagraphForm>("bare", "propertiesOnly", "emptyRun", "text"),
  { maxLength: 3 },
);

describe("a header or footer keeps the blocks it was read with", () => {
  test("the block count survives a rebuild, zero included", async () => {
    await fc.assert(
      fc.property(fc.constantFrom<"hdr" | "ftr">("hdr", "ftr"), body, (root, forms) => {
        const parse = root === "hdr" ? parseHeader : parseFooter;
        const parsed = parse(partXml(root, forms));
        expect(parsed.content).toHaveLength(forms.length);

        // Drop the capture so the serializers run, which is what every
        // edited document does.
        clearHeaderFooterVerbatimXml(parsed);
        const rebuilt = parse(serializeHeaderFooter(parsed));

        expect(rebuilt.content).toHaveLength(forms.length);
      }),
      propertyConfig({ numRuns: 60 }),
    );
  });

  test("an empty header is written empty", () => {
    const parsed = parseHeader(partXml("hdr", []));
    clearHeaderFooterVerbatimXml(parsed);

    const rebuilt = serializeHeaderFooter(parsed);

    expect(rebuilt).not.toContain("<w:p");
  });
});
