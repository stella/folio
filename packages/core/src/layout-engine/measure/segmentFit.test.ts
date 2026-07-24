import { afterEach, describe, expect, test } from "bun:test";

import type { ParagraphBlock, TextRun } from "../types";
import { fixedCharWidth, withFakeTextMeasure } from "./__tests__/fakeTextMeasure";
import { clearAllCaches } from "./cache";
import { setFolioMeasurementFlags } from "./featureFlags";
import { measureParagraph } from "./measureParagraph";
import {
  isSegmentFitActive,
  resetSegmentFitEngine,
  runSegmentFitWalk,
  setSegmentFitEngine,
  styleSupportsSegmentFit,
  type SegmentFitEngine,
  type SegmentFitLine,
} from "./segmentFit";

const fakeMeasure = { charWidth: fixedCharWidth(5) };
const CHAR_W = 5;

function textRun(text: string, extra?: Partial<TextRun>): TextRun {
  return { kind: "text", text, fontFamily: "Stub", fontSize: 12, ...extra };
}

function para(runs: TextRun[], attrs?: ParagraphBlock["attrs"]): ParagraphBlock {
  return { kind: "paragraph", id: "p1", runs, ...(attrs ? { attrs } : {}) };
}

/** Space-splitting engine sharing the fixed 5px/char math. Records calls. */
function makeFakeEngine() {
  const calls: { prepares: string[] } = { prepares: [] };
  const engine: SegmentFitEngine = {
    prepare(text: string) {
      calls.prepares.push(text);
      return { text };
    },
    fitLine(prepared, cursor, maxWidth): SegmentFitLine | null {
      const { text } = prepared as { text: string };
      const start = (cursor as number | null) ?? 0;
      if (start >= text.length) {
        return null;
      }
      // Greedy words; a line's trailing break-space hangs (fit width
      // excludes it), mirroring folio's legacy trim + pretext behaviour.
      let end = start;
      for (;;) {
        let next = text.indexOf(" ", end);
        next = next === -1 ? text.length : next + 1;
        const fitWidth = (next - start) * CHAR_W - (text[next - 1] === " " ? CHAR_W : 0);
        if (fitWidth > maxWidth) {
          break;
        }
        end = next;
        if (end >= text.length) {
          break;
        }
      }
      if (end === start) {
        return null; // overlong word: refuse, core decides
      }
      return { endChar: end, width: (end - start) * CHAR_W, cursor: end };
    },
  };
  return { engine, calls };
}

afterEach(() => {
  resetSegmentFitEngine();
  setFolioMeasurementFlags(undefined);
});

describe("segmentFit registry + gates", () => {
  test("inactive without the feature flag even when an engine is installed", () => {
    setSegmentFitEngine(makeFakeEngine().engine);
    expect(isSegmentFitActive()).toBe(false);
    setFolioMeasurementFlags({ segmentFitLineBreaking: true });
    expect(isSegmentFitActive()).toBe(true);
  });

  test("style allowlist declines letterSpacing, eastAsia fonts, transforms, scaling", () => {
    expect(styleSupportsSegmentFit({ fontFamily: "X", fontSize: 12 })).toBe(true);
    expect(styleSupportsSegmentFit({ fontVariant: "small-caps" })).toBe(true);
    expect(styleSupportsSegmentFit({ letterSpacing: 0.5 })).toBe(false);
    expect(styleSupportsSegmentFit({ eastAsiaFontFamily: "SimSun" })).toBe(false);
    expect(styleSupportsSegmentFit({ textTransform: "uppercase" })).toBe(false);
    expect(styleSupportsSegmentFit({ horizontalScale: 150 })).toBe(false);
    expect(styleSupportsSegmentFit({ horizontalScale: 100 })).toBe(true);
  });

  test("runSegmentFitWalk consults supportsText before prepare", () => {
    const { engine, calls } = makeFakeEngine();
    setSegmentFitEngine({ ...engine, supportsText: () => false });
    const consumed = runSegmentFitWalk("hello", "16px X", {
      spaceLeft: () => 100,
      lineHasContent: () => false,
      commit: () => {},
      wrap: () => {},
    });
    expect(consumed).toBe(0);
    expect(calls.prepares.length).toBe(0);
  });
});

describe("measureParagraph segment-fit wiring", () => {
  test("flag off: registered engine is never consulted", () => {
    withFakeTextMeasure(() => {
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      measureParagraph(para([textRun("hello world wraps here")]), 50);
      expect(calls.prepares.length).toBe(0);
    }, fakeMeasure);
  });

  test("flag on: engine path matches legacy line breaks and heights for spaced text", () => {
    withFakeTextMeasure(() => {
      const block = para([textRun("aaaa bbbb cccc dddd")]);
      const legacy = measureParagraph(block, 50);

      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      const seg = measureParagraph(block, 50);

      expect(calls.prepares).toEqual(["aaaa bbbb cccc dddd"]);
      expect(seg.lines.map((l) => [l.fromChar, l.toChar])).toEqual(
        legacy.lines.map((l) => [l.fromChar, l.toChar]),
      );
      expect(seg.lines.map((l) => l.width)).toEqual(legacy.lines.map((l) => l.width));
      expect(seg.totalHeight).toBeCloseTo(legacy.totalHeight, 4);
    }, fakeMeasure);
  });

  test("justified paragraphs bypass the engine (shrink-tolerance stays legacy)", () => {
    withFakeTextMeasure(() => {
      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      measureParagraph(para([textRun("justify me across lines")], { alignment: "justify" }), 50);
      expect(calls.prepares.length).toBe(0);
    }, fakeMeasure);
  });

  test("automatic-hyphenation paragraphs bypass the engine (mid-word breaks stay legacy)", () => {
    withFakeTextMeasure(() => {
      // The seam fits by whole segments; automatic hyphenation inserts breaks
      // INSIDE a word, which a space-fit engine cannot reproduce. Such runs
      // must stay on the legacy walk or line breaks would diverge.
      const mk = () =>
        para([textRun("hyphenationworthy longwords everywhere")], {
          automaticHyphenation: { enabled: true },
        });
      const legacy = measureParagraph(mk(), 40);

      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      const seg = measureParagraph(mk(), 40);

      expect(calls.prepares.length).toBe(0);
      expect(seg.lines.map((l) => [l.fromChar, l.toChar])).toEqual(
        legacy.lines.map((l) => [l.fromChar, l.toChar]),
      );
    }, fakeMeasure);
  });

  test("letterSpacing runs bypass the engine but measure identically to legacy", () => {
    withFakeTextMeasure(() => {
      const mk = () => para([textRun("spaced out text", { letterSpacing: 1 })]);
      const legacy = measureParagraph(mk(), 40);

      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      const seg = measureParagraph(mk(), 40);

      expect(calls.prepares.length).toBe(0);
      expect(seg.lines.map((l) => [l.fromChar, l.toChar])).toEqual(
        legacy.lines.map((l) => [l.fromChar, l.toChar]),
      );
    }, fakeMeasure);
  });

  test("run-tail glue candidates (#991) bypass the engine for the glued run", () => {
    withFakeTextMeasure(() => {
      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      // Run 1 ends without a break char and run 2 starts glued: the wrap
      // decision for run 1's tail needs cross-run lookahead the engine
      // cannot see, so run 1 must stay legacy. Run 2 has no glue tail and
      // may use the engine.
      const block = para([textRun("word glued"), textRun("Tail more words here")]);
      measureParagraph(block, 50);
      expect(calls.prepares).toEqual(["Tail more words here"]);
    }, fakeMeasure);
  });

  test("engine refusing an overlong token falls back to legacy hard-breaking identically", () => {
    withFakeTextMeasure(() => {
      const token = "x".repeat(30);
      const legacy = measureParagraph(para([textRun(token)]), 50);

      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      setSegmentFitEngine(makeFakeEngine().engine);
      const seg = measureParagraph(para([textRun(token)]), 50);

      expect(seg.lines.map((l) => [l.fromChar, l.toChar])).toEqual(
        legacy.lines.map((l) => [l.fromChar, l.toChar]),
      );
    }, fakeMeasure);
  });
});

describe("eastAsia dual-font bypass (folio PR #2 review)", () => {
  test("a run with eastAsiaFontFamily measures legacy-identically and never reaches the engine", () => {
    withFakeTextMeasure(() => {
      const mk = () => para([textRun("中文 text mixed 內容", { eastAsiaFontFamily: "SimSun" })]);
      const legacy = measureParagraph(mk(), 50);

      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      const seg = measureParagraph(mk(), 50);

      expect(calls.prepares.length).toBe(0);
      expect(seg.lines.map((l) => [l.fromChar, l.toChar])).toEqual(
        legacy.lines.map((l) => [l.fromChar, l.toChar]),
      );
    }, fakeMeasure);
  });

  test("clearAllCaches drops the installed engine prepared state", () => {
    let cleared = 0;
    const { engine } = makeFakeEngine();
    setSegmentFitEngine({
      ...engine,
      clearCaches: () => {
        cleared += 1;
      },
    });
    clearAllCaches();
    expect(cleared).toBe(1);
  });

  test("empty text never reaches prepare", () => {
    withFakeTextMeasure(() => {
      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      const { engine, calls } = makeFakeEngine();
      setSegmentFitEngine(engine);
      measureParagraph(para([textRun("")]), 50);
      expect(calls.prepares.length).toBe(0);
    }, fakeMeasure);
  });
});
