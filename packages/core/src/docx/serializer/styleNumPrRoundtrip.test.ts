import { describe, expect, test } from "bun:test";

import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, ParagraphFormatting } from "../../types/document";
import { serializeParagraphFormatting } from "./paragraphSerializer";

describe("serializeParagraphFormatting style-sourced numPr (#765)", () => {
  test("a style-sourced numPr serializes no direct <w:numPr>", () => {
    const formatting: ParagraphFormatting = {
      styleId: "AppBody-Claim",
      numPr: { kind: "reference" as const, numId: 2 },
      numPrFromStyle: { kind: "reference", numId: 2 },
    };
    const xml = serializeParagraphFormatting(formatting);
    expect(xml).not.toContain("<w:numPr>");
    expect(xml).toContain('<w:pStyle w:val="AppBody-Claim"/>');
  });

  test("a diverged numPr (user changed numbering) still serializes <w:numPr>", () => {
    const formatting: ParagraphFormatting = {
      styleId: "AppBody-Claim",
      numPr: { kind: "reference" as const, numId: 5, ilvl: 0 },
      numPrFromStyle: { kind: "reference", numId: 2 },
    };
    const xml = serializeParagraphFormatting(formatting);
    expect(xml).toContain("<w:numPr>");
    expect(xml).toContain('<w:numId w:val="5"/>');
  });

  test("a direct numPr with no provenance serializes <w:numPr>", () => {
    const formatting: ParagraphFormatting = {
      numPr: { kind: "reference" as const, numId: 2, ilvl: 0 },
    };
    const xml = serializeParagraphFormatting(formatting);
    expect(xml).toContain("<w:numPr>");
  });
});

describe("serializeParagraphFormatting auto-spacing overrides (#823)", () => {
  test("serializes explicit false auto-spacing values", () => {
    const xml = serializeParagraphFormatting({
      spaceBefore: 240,
      beforeAutospacing: false,
      afterAutospacing: false,
    });

    expect(xml).toContain('w:before="240"');
    expect(xml).toContain('w:beforeAutospacing="0"');
    expect(xml).toContain('w:afterAutospacing="0"');
  });
});

// toProseDoc load-side: a paragraph that removes the style's numbering
// (direct numId=0 under a numbered style) drops the style's marker-positioning
// indents and keeps only the indent it states itself.
const STYLE_DEFS = {
  styles: [
    {
      styleId: "Numbered",
      type: "paragraph" as const,
      pPr: {
        numPr: { kind: "reference" as const, numId: 1 },
        indentLeft: 357,
        indentFirstLine: -357,
        hangingIndent: true,
      },
    },
  ],
};

function pmAttrsFor(formatting: ParagraphFormatting): Record<string, unknown> {
  const document: Document = {
    package: {
      document: {
        content: [{ type: "paragraph", content: [], formatting }],
      },
    },
  };
  const pmDoc = toProseDoc(document, { styles: STYLE_DEFS });
  let attrs: Record<string, unknown> = {};
  pmDoc.descendants((node) => {
    if (node.type.name === "paragraph") {
      attrs = node.attrs;
    }
    return false;
  });
  return attrs;
}

describe("style vs direct w:ind merge in toProseDoc (#765)", () => {
  test("removing style numbering (numId 0) drops the style hanging too", () => {
    const attrs = pmAttrsFor({
      styleId: "Numbered",
      numPr: { kind: "none" as const },
      indentLeft: 357,
    });
    expect(attrs["indentLeft"]).toBe(357);
    expect(attrs["indentFirstLine"] ?? null).toBeNull();
    expect(attrs["hangingIndent"] ?? false).toBe(false);
  });

  test("removing style numbering does not retain a style-only left indent", () => {
    const attrs = pmAttrsFor({
      styleId: "Numbered",
      numPr: { kind: "none" as const },
    });
    expect(attrs["indentLeft"] ?? null).toBeNull();
    expect(attrs["indentFirstLine"] ?? null).toBeNull();
    expect(attrs["hangingIndent"] ?? false).toBe(false);
  });

  test("a numbered style without a direct numId keeps the style hanging", () => {
    const attrs = pmAttrsFor({ styleId: "Numbered" });
    expect(attrs["indentLeft"]).toBe(357);
    expect(attrs["indentFirstLine"]).toBe(-357);
    expect(attrs["hangingIndent"]).toBe(true);
    // The style-sourced numPr is projected with provenance.
    expect(attrs["numPr"]).toEqual({ numId: 1 });
    expect(attrs["numPrFromStyle"]).toEqual({ numId: 1 });
  });
});
