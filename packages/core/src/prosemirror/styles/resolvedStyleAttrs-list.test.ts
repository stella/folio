import { describe, expect, test } from "bun:test";

import { createNumberingMap, parseNumbering } from "../../docx/numberingParser";
import {
  listAttrsFromResolvedStyle,
  paragraphAttrsFromResolvedStyle,
  listLevelAttrPatch,
  listLevelIndentRemovalPatch,
} from "./resolvedStyleAttrs";
import { paragraphNumberingAttr } from "../numberingAttr";
import {
  directParagraphIndentation,
  paragraphIndentationFromFormatting,
} from "../paragraphIndentation";
import { expectParagraphAttrs } from "../attrs/index";
import { schema } from "../schema/index";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

// numId 1 -> abstractNum 6 (custom decimalZero4, "[%1]", level ind 360/360)
// numId 2 -> abstractNum 10 (decimal, "[Claim %1]", level ind 360/360)
const NUMBERING_CUSTOM = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W} ${MC}>
  <w:abstractNum w:abstractNumId="6">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <mc:AlternateContent>
        <mc:Choice Requires="w14">
          <w:numFmt w:val="custom" w:format="0001, 0002, ..."/>
        </mc:Choice>
        <mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback>
      </mc:AlternateContent>
      <w:lvlText w:val="[%1]"/>
      <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="10">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="[Claim %1]"/>
      <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="11">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="none"/>
      <w:lvlText w:val=""/>
      <w:pPr><w:ind w:left="700" w:hanging="700"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="6"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="10"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="11"/></w:num>
</w:numbering>`;

describe("listAttrsFromResolvedStyle (#765 applyStyle)", () => {
  const numbering = parseNumbering(NUMBERING_CUSTOM);
  const map = createNumberingMap({
    abstractNums: numbering.definitions.abstractNums,
    nums: numbering.definitions.nums,
  });

  test("projects the style numPr into numPr + marker attrs", () => {
    const attrs = listAttrsFromResolvedStyle(
      { paragraphFormatting: { numPr: { kind: "reference", numId: 2 }, indentLeft: 1134 } },
      map,
    );
    expect(attrs).not.toBeNull();
    expect(attrs?.["numPr"]).toBeNull();
    expect(attrs?.["numPrFromStyle"]).toEqual({ kind: "reference", numId: 2, ilvl: 0 });
    expect(attrs?.["listMarker"]).toBe("[Claim %1]");
    expect(attrs?.["listNumFmt"]).toBe("decimal");
    expect(attrs?.["listAbstractNumId"]).toBe(10);
    // The paired projection preserves the style's indent over its level.
    expect(attrs?.["indentLeft"]).toBe(1134);
  });

  test("projects a custom zero-padded style numbering", () => {
    const attrs = listAttrsFromResolvedStyle(
      { paragraphFormatting: { numPr: { kind: "reference", numId: 1 } } },
      map,
    );
    expect(attrs?.["listMarker"]).toBe("[%1]");
    expect(attrs?.["listNumFmt"]).toBe("decimalZero4");
    expect(attrs?.["listLevelNumFmts"]).toEqual(["decimalZero4"]);
  });

  test("falls back to the numbering level indents when the style has none", () => {
    const attrs = listAttrsFromResolvedStyle(
      { paragraphFormatting: { numPr: { kind: "reference", numId: 2 } } },
      map,
    );
    expect(attrs?.["indentLeft"]).toBe(360);
    expect(attrs?.["indentFirstLine"]).toBe(-360);
    expect(attrs?.["hangingIndent"]).toBe(true);
    expect(attrs?._resolvedFormatting).toMatchObject({
      indentLeft: 360,
      indentFirstLine: -360,
      hangingIndent: true,
    });
  });

  test("returns null for styles without numbering or with numId 0", () => {
    expect(
      listAttrsFromResolvedStyle({ paragraphFormatting: { indentLeft: 100 } }, map),
    ).toBeNull();
    expect(
      listAttrsFromResolvedStyle({ paragraphFormatting: { numPr: { kind: "none" } } }, map),
    ).toBeNull();
  });

  test("without numbering definitions returns numPr but null marker attrs", () => {
    const attrs = listAttrsFromResolvedStyle(
      { paragraphFormatting: { numPr: { kind: "reference", numId: 2 } } },
      null,
    );
    expect(attrs?.["numPr"]).toBeNull();
    expect(attrs?.["listMarker"]).toBeNull();
  });

  test("keeps markerless level indentation without an empty hanging slot", () => {
    const attrs = listAttrsFromResolvedStyle(
      { paragraphFormatting: { numPr: { kind: "reference", numId: 3 } } },
      map,
    );

    expect(attrs?.["listNumFmt"]).toBe("none");
    expect(attrs?.["indentLeft"]).toBe(700);
    expect(attrs?.["indentFirstLine"]).toBeUndefined();
    expect(attrs?.["hangingIndent"]).toBeUndefined();
  });

  test("clears the hanging slot when a list command targets a markerless level", () => {
    const before = {
      ...expectParagraphAttrs(schema.node("paragraph")),
      numPr: paragraphNumberingAttr({ kind: "reference", numId: 3, ilvl: 0 }),
    };
    const attrs = listLevelAttrPatch(before, 0, map);

    expect(attrs["indentLeft"]).toBe(700);
    expect(attrs["indentFirstLine"]).toBeNull();
    expect(attrs["hangingIndent"]).toBe(false);
  });
});

const INDENTATION_SOURCES = [
  { kind: "inherited", direct: undefined },
  { kind: "zero", direct: { indentLeft: 0, indentFirstLine: 0, hangingIndent: false } },
  { kind: "first-line", direct: { indentLeft: 901, indentFirstLine: 120, hangingIndent: false } },
  { kind: "hanging", direct: { indentLeft: 902, indentFirstLine: -240, hangingIndent: true } },
  { kind: "right-only", direct: { indentRight: 123 } },
] as const;
const LEVELS = Array.from({ length: 9 }, (_, level) => level);
const LEVEL_PROVENANCE_CASES = INDENTATION_SOURCES.flatMap((source) =>
  LEVELS.flatMap((from) => LEVELS.map((to) => ({ source, from, to }))),
);
const LEVEL_NUMBERING = parseNumbering(
  `<w:numbering ${W}><w:abstractNum w:abstractNumId="1">${LEVELS.map(
    (level) =>
      `<w:lvl w:ilvl="${String(level)}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
      `<w:lvlText w:val="%${String(level + 1)}."/>` +
      `<w:pPr><w:ind w:left="${String(720 * (level + 1))}" w:hanging="360"/></w:pPr></w:lvl>`,
  ).join("")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
);
const LEVEL_MAP = createNumberingMap(LEVEL_NUMBERING.definitions);
const PARAGRAPH_ATTRS = expectParagraphAttrs(schema.node("paragraph"));

// The initial oracle covered level-zero assignment but not changes between
// levels: stale resolved indentation made a derived level edit look authored.
test.each(LEVEL_PROVENANCE_CASES)(
  "$source.kind indentation keeps its ownership across level $from -> $to",
  ({ source, from, to }) => {
    const direct = paragraphIndentationFromFormatting(source.direct);
    const inherited = { indentLeft: 720 * (from + 1), indentFirstLine: -360, hangingIndent: true };
    const styleFormatting = {
      alignment: "center",
      indentLeft: 100,
      indentFirstLine: 20,
      hangingIndent: false,
      indentRight: 456,
    } as const;
    const styleAttrs = paragraphAttrsFromResolvedStyle(
      { paragraphFormatting: styleFormatting },
      { styleId: "IndentedList" },
    );
    const attrs = {
      ...PARAGRAPH_ATTRS,
      ...styleAttrs,
      numPr: paragraphNumberingAttr({ kind: "reference", numId: 1, ilvl: from }),
      ...inherited,
      ...direct,
      _originalFormatting: { numPr: { kind: "reference", numId: 1, ilvl: from }, ...direct },
      _resolvedFormatting: { ...styleAttrs._resolvedFormatting, ...inherited },
      indentRight: direct?.indentRight ?? styleFormatting.indentRight,
    };
    expect(directParagraphIndentation(attrs)).toEqual(direct);
    const changed = {
      ...attrs,
      ...listLevelAttrPatch(attrs, to, LEVEL_MAP),
    };
    expect(changed.numPr).toEqual({ kind: "reference", numId: 1, ilvl: to });
    expect(changed.indentLeft).toBe(direct?.indentLeft ?? 720 * (to + 1));
    expect(changed.indentFirstLine).toBe(direct?.indentFirstLine ?? -360);
    expect(changed.hangingIndent).toBe(direct?.hangingIndent ?? true);
    expect(changed.alignment).toBe("center");
    expect(changed.alignmentFromStyle).toBe("center");
    expect(changed._resolvedFormatting?.indentRight).toBe(456);
    expect(directParagraphIndentation(changed)).toEqual(direct);
    const restored = {
      ...changed,
      ...listLevelAttrPatch(changed, from, LEVEL_MAP),
    };
    expect(directParagraphIndentation(restored)).toEqual(direct);
    expect(restored.indentLeft).toBe(attrs.indentLeft);
    expect(restored.indentFirstLine).toBe(attrs.indentFirstLine);
    const removed = {
      ...changed,
      ...listLevelIndentRemovalPatch(changed, LEVEL_MAP),
      numPr: paragraphNumberingAttr({ kind: "none" }),
    };
    expect(removed.indentLeft).toBe(direct?.indentLeft ?? styleFormatting.indentLeft);
    expect(removed.indentFirstLine).toBe(
      direct?.indentFirstLine ?? styleFormatting.indentFirstLine,
    );
    expect(removed.hangingIndent).toBe(direct?.hangingIndent ?? styleFormatting.hangingIndent);
    expect(directParagraphIndentation(removed)).toEqual(direct);
  },
);

test("an authored attr edit equal to the next level stays authored", () => {
  const attrs = {
    ...PARAGRAPH_ATTRS,
    ...listLevelAttrPatch(
      {
        ...PARAGRAPH_ATTRS,
        numPr: paragraphNumberingAttr({ kind: "reference", numId: 1, ilvl: 0 }),
      },
      0,
      LEVEL_MAP,
    ),
    indentLeft: 1440,
    _originalFormatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
  };
  expect(directParagraphIndentation(attrs)).toEqual({ indentLeft: 1440 });
  const changed = {
    ...attrs,
    ...listLevelAttrPatch(attrs, 1, LEVEL_MAP),
  };
  expect(changed._resolvedFormatting?.indentLeft).toBe(1440);
  expect(directParagraphIndentation(changed)).toEqual({ indentLeft: 1440 });
});

test("style-owned numbering keeps its baseline when its effective level changes", () => {
  const attrs = {
    ...PARAGRAPH_ATTRS,
    ...paragraphAttrsFromResolvedStyle(
      {
        paragraphFormatting: { indentLeft: 100, indentFirstLine: 20, hangingIndent: false },
      },
      { styleId: "IndentedList" },
    ),
    styleId: "IndentedList",
    numPr: null,
    numPrFromStyle: paragraphNumberingAttr({ kind: "reference", numId: 1, ilvl: 0 }),
    _originalFormatting: { styleId: "IndentedList" },
  };
  const changed = {
    ...attrs,
    ...listLevelAttrPatch(attrs, 1, LEVEL_MAP),
  };
  expect(changed.numPr).toEqual({ kind: "levelOnly", ilvl: 1 });
  expect(changed.numPrFromStyle).toEqual({ kind: "reference", numId: 1, ilvl: 0 });
  expect(changed.indentLeft).toBe(100);
  expect(changed.indentFirstLine).toBe(20);
  expect(changed.hangingIndent).toBe(false);
  expect(directParagraphIndentation(changed)).toBeUndefined();
});

test.each(LEVELS)("style-owned indentation wins over list level %s", (ilvl) => {
  const numPr = paragraphNumberingAttr({ kind: "reference", numId: 1, ilvl: 0 });
  const styleFormatting = { indentLeft: 100, indentFirstLine: 0, hangingIndent: false } as const;
  const attrs = {
    ...PARAGRAPH_ATTRS,
    ...paragraphAttrsFromResolvedStyle(
      { paragraphFormatting: { ...styleFormatting, spaceBefore: 123 } },
      { styleId: "IndentedList" },
    ),
    numPr: null,
    numPrFromStyle: numPr,
    ...styleFormatting,
    _originalFormatting: { styleId: "IndentedList" },
  };
  const changed = {
    ...attrs,
    ...listLevelAttrPatch(attrs, ilvl, LEVEL_MAP),
  };
  expect(changed.indentLeft).toBe(100);
  expect(changed.indentFirstLine).toBe(0);
  expect(changed.hangingIndent).toBe(false);
  expect(changed.spaceBefore).toBe(123);
  expect(directParagraphIndentation(changed)).toBeUndefined();
  const removed = listLevelIndentRemovalPatch(changed, LEVEL_MAP);
  expect({ ...changed, ...removed }.spaceBefore).toBe(123);
});
