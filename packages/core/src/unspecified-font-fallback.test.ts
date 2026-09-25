/**
 * Text whose cascade names no Latin font anywhere measures with
 * `DEFAULT_FONT_FAMILY`; a font named at any level, including through a theme
 * reference in `w:docDefaults`, keeps the fallback out of layout entirely.
 */

import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { createDocx } from "./docx/rezip";
import { layoutDocxHeadless } from "./headless-layout";
import type { Document, StyleDefinitions, Theme } from "./types/document";
import { DEFAULT_FONT_FAMILY, ptToPx } from "./layout-engine/measure/measureHelpers";
import { getMeasureProvider, setMeasureProvider } from "./layout-engine/measure/measureProvider";
import type { FontStyle } from "./layout-engine/measure/measureTypes";

let installed = getMeasureProvider();
let measuredFamilies = new Set<string | undefined>();

const metricsOf = (fontSize: number | undefined) => {
  const px = ptToPx(fontSize ?? 11);
  return {
    fontSize: fontSize ?? 11,
    ascent: px * 0.8,
    descent: px * 0.2,
    fontBoxAscent: px * 0.9,
    fontBoxDescent: px * 0.25,
    lineHeight: px,
    fontFamily: "fixed",
    singleLineRatio: 1.15,
  };
};
const widthOf = (text: string, style: FontStyle) => {
  measuredFamilies.add(style.fontFamily);
  return [...text].length * ptToPx(style.fontSize ?? 11) * 0.5;
};

beforeEach(() => {
  installed = getMeasureProvider();
  measuredFamilies = new Set();
  setMeasureProvider({
    getFontMetrics: (style) => {
      // The pipeline's own provider check asks with no family at all.
      if (style.fontFamily !== undefined) {
        measuredFamilies.add(style.fontFamily);
      }
      return metricsOf(style.fontSize);
    },
    measureTextWidth: (text, style) => widthOf(text, style),
    measureText: (text, style) => {
      const metrics = metricsOf(style.fontSize);
      return {
        width: widthOf(text, style),
        height: metrics.ascent + metrics.descent,
        ascent: metrics.ascent,
        descent: metrics.descent,
      };
    },
    measureRun: (text, style) => {
      const width = widthOf(text, style);
      const per = ptToPx(style.fontSize ?? 11) * 0.5;
      return {
        width,
        charWidths: [...text].flatMap((char) => (char.length === 2 ? [per, 0] : [per])),
        metrics: metricsOf(style.fontSize),
      };
    },
  });
});

afterEach(() => {
  setMeasureProvider(installed);
});

const documentWith = (styles: StyleDefinitions, theme?: Theme): Document => ({
  package: {
    styles,
    ...(theme ? { theme } : {}),
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Body text" }] }],
        },
        // An empty paragraph is sized by its mark's face.
        { type: "paragraph", content: [] },
      ],
    },
  },
});

const familiesMeasured = async (document: Document): Promise<Set<string | undefined>> => {
  const result = await layoutDocxHeadless(await createDocx(document));
  expect(result.isErr()).toBe(false);
  return measuredFamilies;
};

describe("unspecified Latin font fallback", () => {
  test("is Times New Roman", () => {
    expect(DEFAULT_FONT_FAMILY).toBe("Times New Roman");
  });

  test("measures text whose cascade names no font with the fallback", async () => {
    const families = await familiesMeasured(
      documentWith({ docDefaults: { rPr: { fontSize: 24 } }, styles: [] }),
    );
    expect([...families]).toEqual([DEFAULT_FONT_FAMILY]);
  });

  test("never reaches the fallback when w:docDefaults names a face", async () => {
    const families = await familiesMeasured(
      documentWith({
        docDefaults: { rPr: { fontFamily: { ascii: "Arial", hAnsi: "Arial" } } },
        styles: [],
      }),
    );
    expect([...families]).toEqual(["Arial"]);
  });

  test("never reaches the fallback when w:docDefaults names the theme minor font", async () => {
    const families = await familiesMeasured(
      documentWith(
        {
          docDefaults: {
            rPr: { fontFamily: { asciiTheme: "minorHAnsi", hAnsiTheme: "minorHAnsi" } },
          },
          styles: [],
        },
        { fontScheme: { minorFont: { latin: "Aptos" }, majorFont: { latin: "Aptos Display" } } },
      ),
    );
    expect([...families]).toEqual(["Aptos"]);
  });
});
