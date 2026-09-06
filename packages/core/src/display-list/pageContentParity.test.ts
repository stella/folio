/**
 * Nothing the engine put on a page may go unpainted.
 *
 * The display list is now the only thing a backend sees, so a character the
 * builder fails to emit is a character that appears in no output at all: not
 * in the editor, not in an export, and with no error anywhere. Unit tests of
 * the builder cannot catch that, because they assert what was emitted rather
 * than what was owed.
 *
 * The invariant asserted here is a subsequence, not an equality. The painted
 * text legitimately carries things the source text does not (list markers,
 * substituted field values, a discretionary hyphen) and legitimately drops
 * whitespace the layout collapsed. What it may never do is lose a
 * non-whitespace character, or reorder two of them.
 *
 * ## Why the per-page check skips pages carrying a table
 *
 * `getPageTextFromLayout` reads each fragment's `[pmStart, pmEnd)`, and a
 * `TableFragment` never narrows that range to its own `[fromRow, toRow)`
 * window the way a `ParagraphFragment` narrows to its line window. So for a
 * table split across pages, every one of its fragments reports the whole
 * table, and a page holding two fragments reports it twice. That is a defect
 * in the engine rather than in the painter (it also makes the editor's
 * page-text wrong for any split table), and fixing it means changing a field
 * that click mapping also reads, so it is not a drive-by. Until it is fixed,
 * the per-page oracle is only trustworthy on pages with no table, and the
 * whole-document assertion below carries the load for the rest.
 */

import { afterEach, describe, expect, test } from "bun:test";

import { buildDisplayList } from "./build/buildDisplayList";
import type { DisplayPage, DisplayPrimitive } from "./types";
import { layoutDocxHeadless } from "../headless-layout";
import { getMeasureProvider, setMeasureProvider } from "../layout-engine/measure/measureProvider";
import { ptToPx } from "../layout-engine/measure/measureHelpers";
import { getPageTextFromLayout } from "../paged-layout/pageText";

const FIXTURES = [
  new URL("../../../../tests/visual/fixtures/sample.docx", import.meta.url),
  new URL("../../../../tests/visual/fixtures/docx-editor-demo.docx", import.meta.url),
  new URL("../../../../tests/visual/fixtures/podily-bps.docx", import.meta.url),
] as const;

const FIXED_ADVANCE_RATIO = 0.5;

/**
 * A provider with no canvas and no font files, so the assertion is about the
 * builder rather than about any face being installed.
 */
const installFixedWidthProvider = (): void => {
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
  const widthOf = (text: string, fontSize: number | undefined) =>
    [...text].length * ptToPx(fontSize ?? 11) * FIXED_ADVANCE_RATIO;

  setMeasureProvider({
    getFontMetrics: (style) => metricsOf(style.fontSize),
    measureTextWidth: (text, style) => widthOf(text, style.fontSize),
    measureText: (text, style) => {
      const metrics = metricsOf(style.fontSize);
      return {
        width: widthOf(text, style.fontSize),
        height: metrics.ascent + metrics.descent,
        ascent: metrics.ascent,
        descent: metrics.descent,
      };
    },
    measureRun: (text, style) => {
      const per = ptToPx(style.fontSize ?? 11) * FIXED_ADVANCE_RATIO;
      const charWidths = [...text].flatMap((char) => (char.length === 2 ? [per, 0] : [per]));
      return {
        width: charWidths.reduce((total, width) => total + width, 0),
        charWidths,
        metrics: metricsOf(style.fontSize),
      };
    },
  });
};

const glyphTextOf = (primitives: readonly DisplayPrimitive[]): string => {
  let text = "";
  for (const primitive of primitives) {
    switch (primitive.kind) {
      case "glyphRun":
        text += primitive.text;
        break;
      case "clipGroup":
      case "rotateGroup":
      case "opacityGroup":
        text += glyphTextOf(primitive.children);
        break;
      case "rect":
      case "line":
      case "image":
        break;
      default:
        primitive satisfies never;
    }
  }
  return text;
};

const paintedTextOf = (page: DisplayPage): string => glyphTextOf(page.primitives);

/**
 * Case is folded because `w:caps` and `w:smallCaps` are rendering transforms:
 * the painted glyph legitimately differs in case from the character in the
 * document. Losing a character is a defect; changing its case is the feature.
 */
const comparable = (text: string): string => text.replace(/\s+/gu, "").toLocaleUpperCase();

/** Index of the first character of `needle` that `haystack` fails to cover. */
const firstUncoveredIndex = (needle: string, haystack: string): number => {
  let cursor = 0;
  for (let index = 0; index < needle.length; index += 1) {
    const found = haystack.indexOf(needle[index] ?? "", cursor);
    if (found === -1) {
      return index;
    }
    cursor = found + 1;
  }
  return -1;
};

describe("every character the engine placed on a page is painted", () => {
  // The provider is process-wide state that the bun preload installs once, so
  // it is restored after every test rather than at the end of the happy path:
  // a failed assertion would otherwise leave the fixed-width provider in place
  // and fail whichever file bun runs next, as that file's bug.
  const installed = getMeasureProvider();

  afterEach(() => {
    setMeasureProvider(installed);
  });

  for (const fixture of FIXTURES) {
    const name = fixture.pathname.split("/").at(-1) ?? "fixture";

    test(`${name} loses no content between layout and paint`, async () => {
      installFixedWidthProvider();
      const bytes = await Bun.file(fixture).arrayBuffer();

      const laidOut = await layoutDocxHeadless(bytes);
      expect(laidOut.isErr()).toBe(false);
      if (laidOut.isErr()) {
        return;
      }

      const list = buildDisplayList({
        layout: laidOut.value.layout,
        blockLookup: laidOut.value.blockLookup,
      });
      expect(list.pages).toHaveLength(laidOut.value.layout.pages.length);

      // The document as a whole: nothing may be lost anywhere. This is the
      // assertion that holds for every fixture, table or not.
      const wholeDocument = comparable(
        laidOut.value.proseDoc.textBetween(0, laidOut.value.proseDoc.content.size, "\n"),
      );
      const wholePainted = comparable(list.pages.map((page) => paintedTextOf(page)).join(""));
      const lostAt = firstUncoveredIndex(wholeDocument, wholePainted);
      expect({
        scope: "document",
        uncoveredAt: lostAt,
        missingFrom: lostAt === -1 ? "" : wholeDocument.slice(lostAt, lostAt + 40),
      }).toEqual({ scope: "document", uncoveredAt: -1, missingFrom: "" });

      // Per page, wherever the oracle can be trusted: content must land on the
      // page the engine put it on, not merely somewhere in the document.
      for (const [index, page] of list.pages.entries()) {
        if (laidOut.value.layout.pages.at(index)?.fragments.some((f) => f.kind === "table")) {
          continue;
        }
        const expected = comparable(
          getPageTextFromLayout(laidOut.value.layout, laidOut.value.proseDoc, index + 1) ?? "",
        );
        const uncovered = firstUncoveredIndex(expected, comparable(paintedTextOf(page)));
        expect({
          page: index + 1,
          uncoveredAt: uncovered,
          missingFrom: uncovered === -1 ? "" : expected.slice(uncovered, uncovered + 40),
        }).toEqual({ page: index + 1, uncoveredAt: -1, missingFrom: "" });
      }
    });
  }
});
