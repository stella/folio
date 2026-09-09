import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import type { ParagraphBlock, TextRun } from "../../layout-engine/types";
import { toFlowBlocks } from "./toFlowBlocks";

const schema = new Schema({
  nodes: {
    doc: { content: "paragraph+" },
    paragraph: {
      content: "inline*",
      group: "block",
      attrs: {
        styleId: { default: null },
        defaultTextFormatting: { default: null },
      },
    },
    text: { group: "inline" },
  },
  marks: {
    runFormattingOverride: {
      attrs: {
        allCaps: { default: null },
        bold: { default: null },
        boldCs: { default: null },
        color: { default: null },
        cs: { default: null },
        directFontProperties: { default: null },
        doubleStrike: { default: null },
        effect: { default: null },
        emboss: { default: null },
        emphasisMark: { default: null },
        fontSizeCs: { default: null },
        hidden: { default: null },
        highlight: { default: null },
        imprint: { default: null },
        italic: { default: null },
        italicCs: { default: null },
        kerning: { default: null },
        outline: { default: null },
        position: { default: null },
        rtl: { default: null },
        scale: { default: null },
        shading: { default: null },
        shadow: { default: null },
        smallCaps: { default: null },
        spacing: { default: null },
        strike: { default: null },
        underline: { default: null },
        vertAlign: { default: null },
      },
    },
    allCaps: {},
    smallCaps: {},
    emboss: {},
    imprint: {},
    hidden: {},
    textShadow: {},
    textOutline: {},
    emphasisMark: {
      attrs: { type: { default: "dot" } },
    },
    characterSpacing: {
      attrs: {
        spacing: { default: null },
        position: { default: null },
        scale: { default: null },
        kerning: { default: null },
      },
    },
    textColor: {
      attrs: {
        rgb: { default: null },
        themeColor: { default: null },
        themeTint: { default: null },
        themeShade: { default: null },
      },
    },
    highlight: {
      attrs: { color: {} },
    },
    runShading: {
      attrs: {
        rgb: { default: null },
        themeColor: { default: null },
        themeTint: { default: null },
        themeShade: { default: null },
        pattern: { default: null },
        patternColor: { default: null },
      },
    },
    strike: {
      attrs: { double: { default: false } },
    },
    rtl: {},
    language: {
      attrs: {
        val: { default: null },
        eastAsia: { default: null },
        bidi: { default: null },
      },
    },
    fontFamily: {
      attrs: {
        ascii: { default: null },
        hAnsi: { default: null },
        eastAsia: { default: null },
        cs: { default: null },
        asciiTheme: { default: null },
        hAnsiTheme: { default: null },
        eastAsiaTheme: { default: null },
        csTheme: { default: null },
      },
    },
    textEffect: {
      attrs: { effect: {} },
    },
    footnoteRef: {
      attrs: {
        id: { default: null },
        noteType: { default: "footnote" },
        vertAlign: { default: null },
      },
    },
    superscript: {},
    subscript: {},
  },
});

function buildRunWithMarks(
  text: string,
  specs: { markName: string; attrs?: Record<string, unknown> }[],
) {
  const marks = specs.map(({ markName, attrs }) => {
    const mark = schema.marks[markName]?.create(attrs);
    if (!mark) {
      throw new Error(`Unknown mark: ${markName}`);
    }
    return mark;
  });
  return schema.node("doc", null, [schema.node("paragraph", null, [schema.text(text, marks)])]);
}

function buildSingleRunDoc(text: string, markName: string, attrs?: Record<string, unknown>) {
  const mark = schema.marks[markName]?.create(attrs);
  if (!mark) {
    throw new Error(`Unknown mark: ${markName}`);
  }
  return schema.node("doc", null, [schema.node("paragraph", null, [schema.text(text, [mark])])]);
}

function firstRun(blocks: unknown[]): TextRun {
  const paragraph = blocks.find(
    (block) => (block as { kind?: string }).kind === "paragraph",
  ) as ParagraphBlock;
  return paragraph.runs[0] as TextRun;
}

describe("toFlowBlocks run-level OOXML marks", () => {
  test("resolves a theme font before its legacy fallback name", () => {
    const doc = buildSingleRunDoc("text", "fontFamily", {
      ascii: "Calibri",
      hAnsi: "Calibri",
      asciiTheme: "majorHAnsi",
      hAnsiTheme: "majorHAnsi",
    });

    const themedRun = firstRun(
      toFlowBlocks(doc, {
        theme: { fontScheme: { majorFont: { latin: "Aptos" } } },
      }),
    );
    const fallbackRun = firstRun(toFlowBlocks(doc, {}));

    expect(themedRun.fontFamily).toBe("Aptos");
    expect(fallbackRun.fontFamily).toBe("Calibri");
  });

  test("resolves a complex-script theme font from the canonical csTheme field", () => {
    const doc = buildSingleRunDoc("مكتب", "fontFamily", {
      cs: "Noto Naskh Arabic",
      csTheme: "majorBidi",
    });

    const themedRun = firstRun(
      toFlowBlocks(doc, {
        theme: { fontScheme: { majorFont: { cs: "Sakkal Majalla" } } },
      }),
    );
    const fallbackRun = firstRun(toFlowBlocks(doc, {}));

    expect(themedRun.complexScriptFontFamily).toBe("Sakkal Majalla");
    expect(fallbackRun.complexScriptFontFamily).toBe("Noto Naskh Arabic");
  });

  test("carries document-scoped alternate names for every authored font slot", () => {
    const doc = buildSingleRunDoc("text مكتب 日本", "fontFamily", {
      ascii: "Brand Sans",
      hAnsi: "Brand Sans",
      eastAsia: "Brand CJK",
      cs: "Brand Complex",
    });

    const run = firstRun(
      toFlowBlocks(doc, {
        fontAlternates: new Map([
          ["brand sans", "Calibri"],
          ["brand cjk", "MS Mincho"],
          ["brand complex", "Arial"],
        ]),
      }),
    );

    expect(run.fontFamily).toBe("Brand Sans");
    expect(run.alternateFontFamily).toBe("Calibri");
    expect(run.eastAsiaFontFamily).toBe("Brand CJK");
    expect(run.eastAsiaAlternateFontFamily).toBe("MS Mincho");
    expect(run.complexScriptFontFamily).toBe("Brand Complex");
    expect(run.complexScriptAlternateFontFamily).toBe("Arial");
  });

  test("carries an alternate name through paragraph-default formatting", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { defaultTextFormatting: { fontFamily: { ascii: "Brand Sans" } } }, [
        schema.text("Default text"),
      ]),
    ]);

    const run = firstRun(
      toFlowBlocks(doc, {
        fontAlternates: new Map([["brand sans", "Calibri"]]),
      }),
    );

    expect(run.fontFamily).toBe("Brand Sans");
    expect(run.alternateFontFamily).toBe("Calibri");
  });

  test("does not retain a paragraph alternate when a direct font replaces its primary", () => {
    const mark = schema.marks.fontFamily?.create({ ascii: "Direct Face" });
    if (!mark) {
      throw new Error("fontFamily mark is unavailable");
    }
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        { defaultTextFormatting: { fontFamily: { ascii: "Default Face" } } },
        [schema.text("Direct text", [mark])],
      ),
    ]);

    const run = firstRun(
      toFlowBlocks(doc, {
        fontAlternates: new Map([["default face", "Calibri"]]),
      }),
    );

    expect(run.fontFamily).toBe("Direct Face");
    expect(run.alternateFontFamily).toBeUndefined();
  });

  test("propagates caps and text effect marks to run formatting", () => {
    for (const markName of [
      "allCaps",
      "smallCaps",
      "emboss",
      "imprint",
      "textShadow",
      "textOutline",
    ] as const) {
      const blocks = toFlowBlocks(buildSingleRunDoc("text", markName), {});
      expect(firstRun(blocks)[markName]).toBe(true);
    }
  });

  // eigenpal #424 (w:vanish gap 9): the `hidden` PM mark must surface as
  // RunFormatting.hidden so the painter applies the dimmed dotted-underline
  // treatment for the editing view.
  test("propagates the hidden mark (w:vanish) to RunFormatting.hidden", () => {
    const blocks = toFlowBlocks(buildSingleRunDoc("text", "hidden"), {});
    expect(firstRun(blocks).hidden).toBe(true);
  });

  test("propagates character spacing position, scale, and kerning", () => {
    const blocks = toFlowBlocks(
      buildSingleRunDoc("text", "characterSpacing", {
        spacing: 16,
        position: 12,
        scale: 90,
        kerning: 16,
      }),
      {},
    );
    const run = firstRun(blocks);

    expect(run.letterSpacing).toBeCloseTo(1.0667, 3);
    expect(run.positionPx).toBeCloseTo(8, 3);
    expect(run.horizontalScale).toBe(90);
    expect(run.kerningMinPt).toBe(8);
  });

  test("does not emit no-op character spacing values", () => {
    const blocks = toFlowBlocks(
      buildSingleRunDoc("text", "characterSpacing", {
        spacing: 0,
        position: 0,
        scale: 100,
        kerning: 0,
      }),
      {},
    );
    const run = firstRun(blocks);

    expect(run.letterSpacing).toBeUndefined();
    expect(run.positionPx).toBeUndefined();
    expect(run.horizontalScale).toBeUndefined();
    expect(run.kerningMinPt).toBeUndefined();
  });

  test.each([-50, 601, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
    "normalizes malformed character scale %p at the layout boundary",
    (scale) => {
      const blocks = toFlowBlocks(buildSingleRunDoc("text", "characterSpacing", { scale }), {});

      expect(firstRun(blocks).horizontalScale).toBeUndefined();
    },
  );

  test("preserves zero-width character scale", () => {
    const blocks = toFlowBlocks(buildSingleRunDoc("text", "characterSpacing", { scale: 0 }), {});

    expect(firstRun(blocks).horizontalScale).toBe(0);
  });

  test("explicit 100% scale suppresses paragraph-default horizontal scale", () => {
    const mark = schema.marks.characterSpacing?.create({ scale: 100 });
    if (!mark) {
      throw new Error("characterSpacing mark is unavailable");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { defaultTextFormatting: { scale: 150 } }, [
        schema.text("Signature name", [mark]),
      ]),
    ]);

    expect(firstRun(toFlowBlocks(doc, {})).horizontalScale).toBe(100);
  });

  test("explicit zero spacing suppresses paragraph-default character spacing", () => {
    const mark = schema.marks.characterSpacing?.create({ spacing: 0 });
    if (!mark) {
      throw new Error("characterSpacing mark is unavailable");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { defaultTextFormatting: { spacing: -60 } }, [
        schema.text("Signature name", [mark]),
      ]),
    ]);

    expect(firstRun(toFlowBlocks(doc, {})).letterSpacing).toBeUndefined();
  });

  test("projects every run-format cancellation after active and inherited formatting", () => {
    const override = schema.marks.runFormattingOverride?.create({
      allCaps: false,
      bold: false,
      boldCs: false,
      color: "auto",
      cs: false,
      directFontProperties: ["color"],
      doubleStrike: false,
      effect: "none",
      emboss: false,
      emphasisMark: "none",
      fontSizeCs: 18,
      hidden: false,
      highlight: "none",
      imprint: false,
      italic: false,
      italicCs: false,
      kerning: 0,
      outline: false,
      position: 0,
      rtl: false,
      scale: 100,
      shading: { pattern: "nil" },
      shadow: false,
      smallCaps: false,
      spacing: 0,
      strike: false,
      underline: "none",
      vertAlign: "baseline",
    });
    const emphasis = schema.marks.emphasisMark?.create({ type: "dot" });
    const shading = schema.marks.runShading?.create({ rgb: "00AA00" });
    const effect = schema.marks.textEffect?.create({ effect: "shimmer" });
    if (!override || !emphasis || !shading || !effect) {
      throw new Error("Expected run formatting marks");
    }
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: {
            allCaps: true,
            bold: true,
            boldCs: true,
            color: { rgb: "AA0000" },
            cs: true,
            emboss: true,
            emphasisMark: "comma",
            fontSizeCs: 30,
            hidden: true,
            highlight: "yellow",
            imprint: true,
            italic: true,
            italicCs: true,
            kerning: 20,
            outline: true,
            position: 10,
            rtl: true,
            scale: 125,
            shadow: true,
            smallCaps: true,
            spacing: 40,
            strike: true,
            underline: { style: "single" },
            vertAlign: "superscript",
          },
        },
        [schema.text("plain", [effect, emphasis, shading, override])],
      ),
    ]);

    const run = firstRun(toFlowBlocks(doc, {}));

    expect(run).toMatchObject({
      allCaps: false,
      bold: false,
      complexScriptBold: false,
      complexScriptFontSize: 9,
      complexScriptItalic: false,
      emboss: false,
      forceComplexScript: false,
      hidden: false,
      horizontalScale: 100,
      imprint: false,
      italic: false,
      kerningMinPt: 0,
      positionPx: 0,
      rtl: false,
      smallCaps: false,
      strike: false,
      subscript: false,
      superscript: false,
      textOutline: false,
      textShadow: false,
      underline: false,
    });
    expect(run.color).toBeUndefined();
    expect(run.emphasisMark).toBeUndefined();
    expect(run.highlight).toBeUndefined();
    expect(run.letterSpacing).toBeUndefined();
    expect(run.shading).toBeUndefined();
    expect(run.textColorSource).toBeUndefined();
    expect(run.textEffect).toBeUndefined();
  });

  test("a double-strike cancellation does not clear a copied single strike", () => {
    const strike = schema.marks.strike?.create({ double: false });
    const override = schema.marks.runFormattingOverride?.create({ doubleStrike: false });
    if (!strike || !override) {
      throw new Error("Expected strike formatting marks");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("single", [strike, override])]),
    ]);

    expect(firstRun(toFlowBlocks(doc, {})).strike).toBe(true);
  });

  test("propagates rtl mark to run formatting", () => {
    // eigenpal #424 (gap 10) — the painter needs `rtl` on the flow run to
    // emit `dir="rtl"`; without this case the PM mark would survive the
    // ProseMirror round-trip but mixed RTL runs would still paint LTR.
    const blocks = toFlowBlocks(buildSingleRunDoc("שלום", "rtl"), {});
    expect(firstRun(blocks).rtl).toBe(true);
  });

  test("propagates Word language metadata to line layout", () => {
    const blocks = toFlowBlocks(
      buildSingleRunDoc("日本語", "language", {
        val: "en-GB",
        eastAsia: "ja-JP",
        bidi: "ar-SA",
      }),
      {},
    );

    expect(firstRun(blocks).language).toEqual({
      val: "en-GB",
      eastAsia: "ja-JP",
      bidi: "ar-SA",
    });
  });

  test("propagates textEffect mark to run formatting", () => {
    // eigenpal #424 (gap 11) — host CSS keys off `docx-text-effect-<name>`;
    // the painter only emits those classes when the flow run carries the
    // effect value.
    const blocks = toFlowBlocks(
      buildSingleRunDoc("animated", "textEffect", { effect: "shimmer" }),
      {},
    );
    expect(firstRun(blocks).textEffect).toBe("shimmer");
  });

  test("does not raise bare footnote/endnote reference anchors by default", () => {
    // eigenpal/docx-editor#994: Word renders a bare note anchor at the
    // baseline unless the run or resolved character style says superscript.
    const footnote = toFlowBlocks(
      buildSingleRunDoc("1", "footnoteRef", { id: 1, noteType: "footnote" }),
      {},
    );
    const footnoteRun = firstRun(footnote);
    expect(footnoteRun.footnoteRefId).toBe(1);
    expect(footnoteRun.superscript).toBeUndefined();

    const endnote = toFlowBlocks(
      buildSingleRunDoc("i", "footnoteRef", { id: 2, noteType: "endnote" }),
      {},
    );
    const endnoteRun = firstRun(endnote);
    expect(endnoteRun.endnoteRefId).toBe(2);
    expect(endnoteRun.superscript).toBeUndefined();
  });

  test("keeps style-derived superscript on footnote/endnote anchors", () => {
    const footnote = toFlowBlocks(
      buildRunWithMarks("1", [
        { markName: "footnoteRef", attrs: { id: 1, noteType: "footnote" } },
        { markName: "superscript" },
      ]),
      {},
    );
    const footnoteRun = firstRun(footnote);
    expect(footnoteRun.footnoteRefId).toBe(1);
    expect(footnoteRun.superscript).toBe(true);
  });

  test("keeps explicit note-reference superscript on footnote anchors", () => {
    const footnote = toFlowBlocks(
      buildSingleRunDoc("1", "footnoteRef", {
        id: 1,
        noteType: "footnote",
        vertAlign: "superscript",
      }),
      {},
    );
    const footnoteRun = firstRun(footnote);
    expect(footnoteRun.footnoteRefId).toBe(1);
    expect(footnoteRun.superscript).toBe(true);
  });

  test("does not raise a footnote anchor that carries an explicit subscript", () => {
    // The superscript default must never override an explicit subscript on the
    // same run, regardless of mark order (eigenpal/docx-editor#845).
    const blocks = toFlowBlocks(
      buildRunWithMarks("1", [
        { markName: "footnoteRef", attrs: { id: 1, noteType: "footnote" } },
        { markName: "subscript" },
      ]),
      {},
    );
    const run = firstRun(blocks);
    expect(run.footnoteRefId).toBe(1);
    expect(run.subscript).toBe(true);
    expect(run.superscript).toBeUndefined();
  });

  test("lets explicit subscript override note-reference superscript attrs", () => {
    for (const marks of [
      [
        {
          markName: "footnoteRef",
          attrs: { id: 1, noteType: "footnote", vertAlign: "superscript" },
        },
        { markName: "subscript" },
      ],
      [
        { markName: "subscript" },
        {
          markName: "footnoteRef",
          attrs: { id: 1, noteType: "footnote", vertAlign: "superscript" },
        },
      ],
    ] as const) {
      const blocks = toFlowBlocks(buildRunWithMarks("1", marks), {});
      const run = firstRun(blocks);
      expect(run.footnoteRefId).toBe(1);
      expect(run.subscript).toBe(true);
      expect(run.superscript).toBeUndefined();
    }
  });

  test("does not raise a footnote anchor with explicit baseline vertAlign", () => {
    const blocks = toFlowBlocks(
      buildSingleRunDoc("1", "footnoteRef", {
        id: 1,
        noteType: "footnote",
        vertAlign: "baseline",
      }),
      {},
    );
    const run = firstRun(blocks);
    expect(run.footnoteRefId).toBe(1);
    expect(run.superscript).toBeUndefined();
    expect(run.subscript).toBeUndefined();
  });

  test("propagates emphasis mark variants", () => {
    for (const variant of ["dot", "comma", "circle", "underDot"] as const) {
      const blocks = toFlowBlocks(buildSingleRunDoc("text", "emphasisMark", { type: variant }), {});
      expect(firstRun(blocks).emphasisMark).toBe(variant);
    }
  });

  test("cascades paragraph default text formatting to unmarked runs", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: {
            fontFamily: { ascii: "Arial Narrow", hAnsi: "Arial Narrow" },
            fontSize: 22,
            bold: true,
            color: { rgb: "C00000" },
            underline: { style: "single" },
            smallCaps: true,
          },
        },
        [schema.text("body text")],
      ),
    ]);
    const run = firstRun(toFlowBlocks(doc, {}));

    expect(run.fontFamily).toBe("Arial Narrow");
    expect(run.fontSize).toBe(11);
    expect(run.bold).toBe(true);
    expect(run.color).toBe("#C00000");
    expect(run.textColorSource).toBe("paragraphDefault");
    expect(run.underline).toEqual({ style: "single" });
    expect(run.smallCaps).toBe(true);
  });

  test("keeps inherited paragraph default black identifiable on highlighted runs", () => {
    const highlight = schema.marks.highlight.create({ color: "darkBlue" });
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: {
            color: { rgb: "000000" },
            highlight: "darkBlue",
          },
        },
        [schema.text("body text", [highlight])],
      ),
    ]);
    const run = firstRun(toFlowBlocks(doc, {}));

    expect(run.color).toBe("#000000");
    expect(run.textColorSource).toBe("paragraphDefault");
    expect(run.highlight).toBe("#00008B");
  });

  test("keeps direct black text colors marked as direct when paragraph default is also black", () => {
    const textColor = schema.marks.textColor.create({ rgb: "000000" });
    const highlight = schema.marks.highlight.create({ color: "darkBlue" });
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: {
            color: { rgb: "000000" },
          },
        },
        [schema.text("body text", [textColor, highlight])],
      ),
    ]);
    const run = firstRun(toFlowBlocks(doc, {}));

    expect(run.color).toBe("#000000");
    expect(run.textColorSource).toBe("direct");
    expect(run.highlight).toBe("#00008B");
  });

  test("keeps distinguishable direct black text colors marked as direct", () => {
    const textColor = schema.marks.textColor.create({ rgb: "000000" });
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: {
            color: { rgb: "C00000" },
          },
        },
        [schema.text("body text", [textColor])],
      ),
    ]);
    const run = firstRun(toFlowBlocks(doc, {}));

    expect(run.color).toBe("#000000");
    expect(run.textColorSource).toBe("direct");
  });

  test("omits automatic paragraph default text colors from runs", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: {
            color: { auto: true },
            highlight: "darkBlue",
          },
        },
        [schema.text("body text")],
      ),
    ]);
    const run = firstRun(toFlowBlocks(doc, {}));

    expect(run.color).toBeUndefined();
    expect(run.highlight).toBe("#00008B");
  });
});
