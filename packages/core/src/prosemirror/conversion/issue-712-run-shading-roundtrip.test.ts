// eigenpal #722 (#712) — run-level shading (w:shd) used as a background.
//
// Folio models w:highlight as a strict OOXML named-palette union, so an
// arbitrary run-background fill (e.g. a Word/Google Docs `w:shd`, or a custom
// highlight another editor saved as shading) cannot ride on the highlight mark.
// It was parsed into `formatting.shading` but DROPPED at PM conversion, so the
// background silently vanished on reload. It now round-trips as a dedicated
// `runShading` mark and renders as a background.

import { describe, expect, test } from "bun:test";

import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import type { FlowBlock, TextRun } from "../../layout-engine/types";
import type { Document, Paragraph, Run, ShadingProperties } from "../../types/document";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const wrap = (formatting: Run["formatting"]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "x" }], formatting }],
        },
      ],
    },
  },
});

const firstParagraph = (document: Document): Paragraph => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("expected first block to be a paragraph");
  }
  return block;
};

const firstRunFormatting = (document: Document): Run["formatting"] => {
  for (const content of firstParagraph(document).content) {
    if (content.type === "run") {
      return content.formatting;
    }
  }
  throw new Error("expected a run");
};

// The runShading mark color, as it would round-trip out to the model on save.
const roundTripShading = (formatting: Run["formatting"]): ShadingProperties | undefined =>
  firstRunFormatting(fromProseDoc(toProseDoc(wrap(formatting)), wrap(formatting)))?.shading;

// The resolved CSS background the painter receives for the run.
const renderedBackground = (formatting: Run["formatting"]): string | undefined => {
  const blocks: FlowBlock[] = toFlowBlocks(toProseDoc(wrap(formatting)));
  const paragraph = blocks.find((b) => b.kind === "paragraph");
  if (paragraph?.kind !== "paragraph") {
    throw new Error("expected a paragraph flow block");
  }
  const textRun = paragraph.runs.find((r): r is TextRun => r.kind === "text");
  return textRun?.shading;
};

describe("Issue #712 — run shading round-trips and renders", () => {
  test("a solid fill renders as a background and survives the round-trip", () => {
    const shading = { fill: { rgb: "FFFF00" } };
    expect(renderedBackground({ shading })).toBe("#FFFF00");
    expect(roundTripShading({ shading })?.fill?.rgb).toBe("FFFF00");
  });

  test("the default `clear` pattern remains authored while rendering as a solid fill", () => {
    const shading = { pattern: "clear" as const, fill: { rgb: "00B050" } };
    expect(renderedBackground({ shading })).toBe("#00B050");
    const out = roundTripShading({ shading });
    expect(out?.fill?.rgb).toBe("00B050");
    expect(out?.pattern).toBe("clear");
  });

  test("a non-clear pattern is carried for export fidelity", () => {
    const shading = { pattern: "pct25" as const, fill: { rgb: "FFFF00" } };
    expect(roundTripShading({ shading })?.pattern).toBe("pct25");
  });

  test("a non-clear pattern's foreground color round-trips", () => {
    const shading = {
      pattern: "pct25" as const,
      color: { rgb: "FF0000" },
      fill: { rgb: "DDDDDD" },
    };
    const out = roundTripShading({ shading });
    expect(out?.pattern).toBe("pct25");
    expect(out?.color?.rgb).toBe("FF0000");
    expect(out?.fill?.rgb).toBe("DDDDDD");
  });

  test("a solid pattern with a color but no fill renders and round-trips", () => {
    // `<w:shd w:val="solid" w:color="FF0000"/>` paints the color as a solid
    // background without changing which OOXML slot authored the value.
    const shading = { pattern: "solid" as const, color: { rgb: "FF0000" } };
    expect(renderedBackground({ shading })).toBe("#FF0000");
    expect(roundTripShading({ shading })).toEqual(shading);
  });

  test("a solid pattern paints the color over the fill (color wins)", () => {
    // `<w:shd w:val="solid" w:color="FF0000" w:fill="00FF00"/>` — the solid
    // pattern covers the fill, so the visible background is the color.
    const shading = {
      pattern: "solid" as const,
      color: { rgb: "FF0000" },
      fill: { rgb: "00FF00" },
    };
    expect(renderedBackground({ shading })).toBe("#FF0000");
    expect(roundTripShading({ shading })).toEqual(shading);
  });

  test("a theme-color fill round-trips (themeColor preserved)", () => {
    const formatting: Run["formatting"] = {
      shading: { fill: { themeColor: "accent1" } },
    };
    expect(roundTripShading(formatting)?.fill?.themeColor).toBe("accent1");
  });

  test("highlight and shading both reach the flow run (painter resolves precedence)", () => {
    // Both fields flow through independently; the painter prefers `highlight`
    // (covered in renderParagraph-run-shading.test.ts). Here we only assert the
    // data is not lost.
    const formatting: Run["formatting"] = {
      // eslint-disable-next-line no-inline-style-colors/no-inline-style-colors -- a named OOXML highlight value, not UI styling
      highlight: "green",
      shading: { fill: { rgb: "FFFF00" } },
    };
    const blocks = toFlowBlocks(toProseDoc(wrap(formatting)));
    const paragraph = blocks.find((b) => b.kind === "paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("expected a paragraph flow block");
    }
    const run = paragraph.runs.find((r): r is TextRun => r.kind === "text");
    expect(run?.highlight).toBeDefined();
    expect(run?.shading).toBe("#FFFF00");
  });

  test("an `auto` fill produces no shading background", () => {
    expect(renderedBackground({ shading: { fill: { auto: true } } })).toBe(undefined);
    expect(roundTripShading({ shading: { fill: { auto: true } } })).toEqual({
      fill: { auto: true },
    });
  });

  test("a fill-less shading remains authored without producing a visible background", () => {
    const formatting: Run["formatting"] = { shading: { pattern: "pct25" } };
    expect(renderedBackground(formatting)).toBeUndefined();
    expect(roundTripShading(formatting)).toEqual({ pattern: "pct25" });
  });

  test.each([
    {
      authored: { pattern: "clear" as const, fill: { rgb: "00B050" } },
      serialized: { pattern: "clear" as const, fill: { rgb: "00B050" } },
    },
    {
      authored: { pattern: "solid" as const, color: { rgb: "FF0000" } },
      serialized: { pattern: "solid" as const, color: { rgb: "FF0000" } },
    },
    {
      authored: {
        pattern: "solid" as const,
        color: { rgb: "FF0000" },
        fill: { rgb: "00FF00" },
      },
      serialized: {
        pattern: "solid" as const,
        color: { rgb: "FF0000" },
        fill: { rgb: "00FF00" },
      },
    },
    {
      authored: { fill: { auto: true } },
      // CT_Shd requires w:val, so the serializer supplies its canonical default.
      serialized: { pattern: "clear" as const, fill: { auto: true } },
    },
    {
      authored: { pattern: "pct25" as const },
      serialized: { pattern: "pct25" as const },
    },
  ])(
    "preserves an otherwise invisible authored shading shape through DOCX reopen",
    async ({ authored, serialized }) => {
      const input = wrap({ shading: authored });
      const firstBuffer = await createDocx(fromProseDoc(toProseDoc(input), input));
      const firstReopen = await parseDocx(firstBuffer, {
        detectVariables: false,
        preloadFonts: false,
      });
      expect(firstRunFormatting(firstReopen)?.shading).toEqual(serialized);

      const secondBuffer = await createDocx(fromProseDoc(toProseDoc(firstReopen), firstReopen));
      const secondReopen = await parseDocx(secondBuffer, {
        detectVariables: false,
        preloadFonts: false,
      });
      expect(firstRunFormatting(secondReopen)?.shading).toEqual(serialized);
    },
  );

  test("a run with no shading is unaffected", () => {
    expect(renderedBackground({ bold: true })).toBeUndefined();
    expect(roundTripShading({ bold: true })).toBeUndefined();
  });
});

describe("Issue #712 — HTML paste of a custom background", () => {
  const rule = schema.marks["runShading"]?.spec.parseDOM?.[0];
  const getAttrs = (value: string): unknown =>
    (rule?.getAttrs as ((v: string) => unknown) | undefined)?.(value);

  test("a custom background-color is claimed as run shading", () => {
    expect(getAttrs("#D9D9D9")).toEqual({ rgb: "D9D9D9" });
    expect(getAttrs("rgb(217, 217, 217)")).toEqual({ rgb: "D9D9D9" });
  });

  test("a named-palette color is left to the highlight mark", () => {
    // highlight maps these; runShading must not double-claim them.
    expect(getAttrs("yellow")).toBe(false);
    expect(getAttrs("#FFFF00")).toBe(false);
  });

  test("a non-color value is ignored", () => {
    expect(getAttrs("transparent")).toBe(false);
  });
});
