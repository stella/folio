/**
 * An optional `w:pPr` field has three states, and the editor round trip has to
 * keep them apart: ABSENT (inherit from the style), explicitly off, explicitly
 * on. The public-corpus gate found the third state colonising the first —
 * `bidi: absent became false`, `suppressAutoHyphens: absent became true`,
 * `widowControl: absent became false`, `borders: absent became object` — on
 * documents whose styles define those fields and whose paragraphs do not.
 *
 * `modelFixedPoint.property.test.ts` next to this file pins the whole
 * paragraph over parsed fixtures. None of those fixtures carries a style that
 * defines these fields, which is why the loss survived it. This one generates
 * the style chain as well as the paragraph, so every combination of (style
 * states the field, paragraph states it or not) is covered.
 *
 * The second property is the one that makes the first worth having: a
 * paragraph that inherits must go on inheriting. Editing the style after a
 * save has to move the paragraph's effective value, which it cannot do once
 * the save has written the inherited value into the paragraph's own `w:pPr`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { createEmptyDocument } from "../../utils/createDocument";
import type {
  BorderSpec,
  Document,
  Paragraph,
  ParagraphFormatting,
  Run,
  Style,
} from "../../types/document";
import { STYLE_RESOLVED_PARAGRAPH_FIELDS } from "../paragraphFormattingProvenance";
import { updateDocumentContent } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const STYLE_ID = "Governed";
const BASE_STYLE_ID = "GovernedBase";

const BORDER: BorderSpec = { style: "single", size: 8, color: { rgb: "FF0000" } };
const OTHER_BORDER: BorderSpec = { style: "double", size: 12, color: { rgb: "00FF00" } };

/**
 * Two distinct values per governed field, so "the style says X, the paragraph
 * says Y" is representable for every one of them. Boolean fields use the two
 * OOXML toggle states, which is exactly the explicit-off/explicit-on pair the
 * corpus signatures confused with absence.
 */
const FIELD_VALUES = {
  bidi: [false, true],
  kinsoku: [false, true],
  overflowPunctuation: [false, true],
  snapToGrid: [false, true],
  widowControl: [false, true],
  pageBreakBefore: [true, false],
  contextualSpacing: [true, false],
  suppressAutoHyphens: [true, false],
  indentLeft: [440, 880],
  indentRight: [110, 220],
  indentFirstLine: [330, 660],
  outlineLevel: [2, 5],
  borders: [{ top: BORDER, bottom: BORDER }, { left: OTHER_BORDER }],
  shading: [
    { fill: { rgb: "CCCCCC" }, pattern: "clear" },
    { fill: { rgb: "112233" }, pattern: "solid" },
  ],
  tabs: [
    [{ position: 720, alignment: "left" }],
    [{ position: 1440, alignment: "right", leader: "dot" }],
  ],
} as const satisfies Record<
  (typeof STYLE_RESOLVED_PARAGRAPH_FIELDS)[number],
  readonly [unknown, unknown]
>;

type GovernedField = keyof typeof FIELD_VALUES;

const GOVERNED_FIELDS = Object.keys(FIELD_VALUES) as GovernedField[];

/** Absent, or one of the field's two explicit values. */
type FieldState = 0 | 1 | 2;

const applyStates = (
  states: readonly FieldState[],
  pick: (field: GovernedField, state: Exclude<FieldState, 0>) => unknown,
): ParagraphFormatting => {
  const formatting: ParagraphFormatting = {};
  for (const [index, field] of GOVERNED_FIELDS.entries()) {
    const state = states[index] ?? 0;
    if (state === 0) {
      continue;
    }
    Reflect.set(formatting, field, pick(field, state));
  }
  return formatting;
};

const valueOf = (field: GovernedField, state: Exclude<FieldState, 0>): unknown =>
  FIELD_VALUES[field][state - 1];

const sameValue = (left: unknown, right: unknown): boolean =>
  JSON.stringify(left ?? null) === JSON.stringify(right ?? null);

const fieldStatesArb = fc.array(fc.constantFrom<FieldState>(0, 1, 2), {
  minLength: GOVERNED_FIELDS.length,
  maxLength: GOVERNED_FIELDS.length,
});

type Scenario = {
  /** What the base style in the chain defines. */
  base: ParagraphFormatting;
  /** What the paragraph's own style defines, over the base. */
  style: ParagraphFormatting;
  /** What the paragraph states directly. */
  direct: ParagraphFormatting;
};

const scenarioArb: fc.Arbitrary<Scenario> = fc
  .record({
    base: fieldStatesArb,
    style: fieldStatesArb,
    direct: fieldStatesArb,
  })
  .map(({ base, style, direct }) => ({
    base: applyStates(base, valueOf),
    style: applyStates(style, valueOf),
    direct: applyStates(direct, valueOf),
  }));

/** A run carrying an embedded object plus its own `w:rPr`. */
const OBJECT_RUN: Run = {
  type: "run",
  formatting: { language: { val: "en-US" } },
  content: [
    {
      type: "drawing",
      image: {
        type: "image",
        rId: "rId8",
        size: { width: 685_800, height: 485_775 },
        wrap: { type: "inline" },
      },
      rawXml: '<w:object w:dxaOrig="1080" w:dyaOrig="765"/>',
    },
  ],
};

/** Adjacent runs whose texts differ in whether `xml:space` is required. */
const ADJACENT_TEXT_RUNS: Run[] = [
  { type: "run", content: [{ type: "text", text: "This " }] },
  { type: "run", content: [{ type: "text", text: "." }] },
];

const documentFor = (scenario: Scenario): Document => {
  const document = createEmptyDocument();
  const styles: Style[] = [
    { styleId: BASE_STYLE_ID, type: "paragraph", name: BASE_STYLE_ID, pPr: scenario.base },
    {
      styleId: STYLE_ID,
      type: "paragraph",
      name: STYLE_ID,
      basedOn: BASE_STYLE_ID,
      pPr: scenario.style,
    },
  ];
  document.package.styles = {
    ...document.package.styles,
    styles: [...(document.package.styles?.styles ?? []), ...styles],
  };
  const paragraph: Paragraph = {
    type: "paragraph",
    formatting: { styleId: STYLE_ID, ...scenario.direct },
    content: [...ADJACENT_TEXT_RUNS, OBJECT_RUN],
  };
  document.package.document.content = [paragraph, { ...paragraph, formatting: undefined }];
  return document;
};

const rebuild = (document: Document): Document =>
  updateDocumentContent(
    document,
    toProseDoc(document, { styles: document.package.styles, theme: document.package.theme }),
  );

const bodyParagraphs = (document: Document): Paragraph[] =>
  document.package.document.content.filter((block) => block.type === "paragraph");

describe("a style-resolved paragraph property stays out of direct w:pPr", () => {
  test("every governed field keeps absent, explicit-off and explicit-on apart", () => {
    fc.assert(
      fc.property(scenarioArb, (scenario) => {
        const document = documentFor(scenario);
        const before = bodyParagraphs(document);
        const after = bodyParagraphs(rebuild(document));
        expect(after).toHaveLength(before.length);
        for (const [index, paragraph] of before.entries()) {
          for (const field of GOVERNED_FIELDS) {
            expect(after[index]?.formatting?.[field], `paragraph ${index} field ${field}`).toEqual(
              paragraph.formatting?.[field],
            );
          }
        }
      }),
      { numRuns: 200 },
    );
  });

  test("a later edit to the style still moves the paragraph's effective value", () => {
    fc.assert(
      fc.property(scenarioArb, fc.constantFrom(...GOVERNED_FIELDS), (scenario, field) => {
        // Only a field the paragraph inherits can be moved by the style.
        fc.pre(scenario.direct[field] === undefined);
        fc.pre(scenario.style[field] !== undefined || scenario.base[field] !== undefined);

        const saved = rebuild(documentFor(scenario));
        const restyled = structuredClone(saved);
        const style = restyled.package.styles?.styles.find(({ styleId }) => styleId === STYLE_ID);
        if (!style) {
          throw new Error("the generated document must carry the paragraph style");
        }
        const resolved = { ...style.pPr }[field];
        const moved = valueOf(field, sameValue(resolved, FIELD_VALUES[field][0]) ? 2 : 1);
        style.pPr = { ...style.pPr, [field]: moved };

        // `direction` is the attr `bidi` resolves into; every other governed
        // field keeps its own name.
        const attr = field === "bidi" ? "direction" : field;
        const before = toProseDoc(saved, { styles: saved.package.styles }).child(0).attrs[attr];
        const after = toProseDoc(restyled, { styles: restyled.package.styles }).child(0).attrs[
          attr
        ];
        expect(sameValue(after, before), `a style edit must still reach ${field}`).toBe(false);
      }),
      { numRuns: 200 },
    );
  });
});

describe("the editor round trip keeps a run's own properties", () => {
  test("an embedded-object run keeps its w:rPr, and merged text keeps its spacing", () => {
    const document = documentFor({ base: {}, style: {}, direct: {} });
    const [paragraph] = bodyParagraphs(rebuild(document));
    const objectRun = paragraph?.content.findLast((item) => item.type === "run");
    expect(objectRun?.type).toBe("run");
    expect(objectRun?.type === "run" ? objectRun.formatting : undefined).toMatchObject({
      language: { val: "en-US" },
    });
    const text = paragraph?.content
      .filter((item) => item.type === "run")
      .flatMap((run) => run.content)
      .filter((item) => item.type === "text")
      .map((item) => item.text)
      .join("");
    expect(text).toBe("This .");
  });
});
