import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { resolveAllChangesInHeadlessState } from "./comments";
import { createDocumentStylesPlugin } from "../plugins/documentStyles";
import {
  LIST_RENDERING_ATTR_DEFAULTS,
  LIST_RENDERING_ATTR_KEYS,
  PPR_CHANGE_SCOPED_ATTR_DEFAULTS,
  PPR_CHANGE_SCOPED_ATTR_KEYS,
  PPR_CHANGE_SCOPED_FORMATTING_KEYS,
  PPR_FORMATTING_FIELD_DISPOSITIONS,
} from "../schema/paragraphAttrDefaults";
import type { ParagraphAttrs } from "../schema/nodes";
import { schema } from "../schema";
import type { ParagraphFormatting } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { paragraphRejectAttrPatch, paragraphRejectOriginalFormatting } from "./propertyChangeScope";

const NON_DEFAULT_LIST_RENDERING_ATTRS = {
  listIsBullet: true,
  listIsLegal: true,
  listNumFmt: "decimal",
  listMarker: "1.",
  listMarkerTemplate: "%1.",
  listMarkerHidden: true,
  listMarkerFormatting: { bold: true },
  listMarkerAlignment: "right",
  listMarkerSuffix: "space",
  listMarkerAllCaps: true,
  listImplicitChildLevelAdvances: 1,
  listMarkerSecondSlotOffsetTwips: 720,
  listLevelNumFmts: ["decimal"],
  listLevelStarts: [1],
  listAbstractNumId: 1,
  listStartOverride: 2,
} as const satisfies {
  readonly [Key in keyof typeof LIST_RENDERING_ATTR_DEFAULTS]: Exclude<
    ParagraphAttrs[Key],
    null | undefined
  >;
};

const PARAGRAPH_FORMATTING_SAMPLES = {
  alignment: "right",
  bidi: true,
  kinsoku: false,
  overflowPunctuation: true,
  spaceBefore: 120,
  spaceAfter: 240,
  lineSpacing: 360,
  lineSpacingRule: "exact",
  snapToGrid: false,
  beforeAutospacing: true,
  afterAutospacing: false,
  spacingExplicit: { before: true, after: true },
  indentLeft: 720,
  indentRight: 360,
  indentFirstLine: -360,
  hangingIndent: true,
  borders: { top: { style: "single", size: 4 } },
  shading: { fill: { rgb: "EAEAEA" } },
  tabs: [{ position: 1440, alignment: "right", leader: "dot" }],
  keepNext: true,
  keepLines: false,
  widowControl: true,
  pageBreakBefore: false,
  contextualSpacing: true,
  numPr: { numId: 4, ilvl: 2 },
  numPrFromStyle: { numId: 4, ilvl: 2 },
  outlineLevel: 2,
  styleId: "Normal",
  frame: { width: 720 },
  suppressLineNumbers: false,
  suppressAutoHyphens: false,
  runProperties: { bold: true },
  runInWithNext: true,
} satisfies {
  readonly [Key in keyof Required<ParagraphFormatting>]: NonNullable<ParagraphFormatting[Key]>;
};

describe("canonical paragraph attr restoration", () => {
  test("the schema, property rejection, and list clearing share every canonical default", () => {
    const schemaAttrs = schema.nodes.paragraph.spec.attrs;
    for (const key of PPR_CHANGE_SCOPED_ATTR_KEYS) {
      expect(schemaAttrs?.[key]?.default).toBe(PPR_CHANGE_SCOPED_ATTR_DEFAULTS[key]);
    }
    for (const key of LIST_RENDERING_ATTR_KEYS) {
      expect(schemaAttrs?.[key]?.default).toBe(LIST_RENDERING_ATTR_DEFAULTS[key]);
    }
    for (const disposition of Object.values(PPR_FORMATTING_FIELD_DISPOSITIONS)) {
      switch (disposition.type) {
        case "attr":
          expect(Object.hasOwn(PPR_CHANGE_SCOPED_ATTR_DEFAULTS, disposition.attr)).toBe(true);
          break;
        case "mapped":
          for (const attr of disposition.attrs) {
            expect(Object.hasOwn(PPR_CHANGE_SCOPED_ATTR_DEFAULTS, attr)).toBe(true);
          }
          break;
        case "original-only":
        case "preserved-live-original":
        case "outside-change-scope":
          break;
      }
    }

    const rejected = paragraphRejectAttrPatch(undefined);
    for (const key of PPR_CHANGE_SCOPED_ATTR_KEYS) {
      expect(rejected[key]).toBe(PPR_CHANGE_SCOPED_ATTR_DEFAULTS[key]);
    }

    const populated = schema.node("paragraph", NON_DEFAULT_LIST_RENDERING_ATTRS, [
      schema.text("Clause"),
    ]);
    const cleared = populated.type.create(
      { ...populated.attrs, ...LIST_RENDERING_ATTR_DEFAULTS },
      populated.content,
    );
    const canonical = schema.node("paragraph", null, [schema.text("Clause")]);
    expect(cleared.eq(canonical)).toBe(true);
  });

  test.each([
    [{ spaceBefore: 0 }, { before: true }],
    [{ spaceAfter: 0 }, { after: true }],
    [
      { spaceBefore: 120, spaceAfter: 240 },
      { before: true, after: true },
    ],
    [{ lineSpacing: 360, lineSpacingRule: "exact" as const }, undefined],
  ] as const)("rebuilds authored spacing provenance for %#", (previous, spacingExplicit) => {
    expect(paragraphRejectOriginalFormatting(previous, null)).toEqual({
      ...previous,
      ...(spacingExplicit === undefined ? {} : { spacingExplicit }),
    });
  });

  test("list-property views reconstruct both sides across every rendering attr", () => {
    const originalParagraph = schema.node("paragraph", NON_DEFAULT_LIST_RENDERING_ATTRS, [
      schema.text("Clause"),
    ]);
    const canonicalTargetParagraph = schema.node("paragraph", null, [schema.text("Clause")]);
    const pendingParagraph = originalParagraph.type.create(
      {
        ...originalParagraph.attrs,
        ...LIST_RENDERING_ATTR_DEFAULTS,
        _propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 1, author: "reviewer", date: "2026-09-12T10:00:00.000Z" },
            previousFormatting: NON_DEFAULT_LIST_RENDERING_ATTRS,
          },
        ],
      },
      originalParagraph.content,
    );
    const pending = schema.node("doc", null, [pendingParagraph]);

    const accepted = resolveAllChangesInHeadlessState(
      EditorState.create({ doc: pending }),
      "accept",
    ).doc;
    const rejected = resolveAllChangesInHeadlessState(
      EditorState.create({ doc: pending }),
      "reject",
    ).doc;

    expect(accepted.eq(schema.node("doc", null, [canonicalTargetParagraph]))).toBe(true);
    expect(rejected.eq(schema.node("doc", null, [originalParagraph]))).toBe(true);
  });

  test("a rejected paragraph-property view is a document-model projection fixed point", () => {
    const source = createEmptyDocument({ initialText: "Clause" });
    const sourceParagraph = source.package.document.content.at(0);
    if (sourceParagraph?.type !== "paragraph") {
      throw new Error("Expected an initial paragraph");
    }
    const sourceFormatting: ParagraphFormatting = { ...PARAGRAPH_FORMATTING_SAMPLES };
    delete sourceFormatting.numPrFromStyle;
    sourceParagraph.formatting = sourceFormatting;
    const canonicalSource = toProseDoc(source);
    const canonicalParagraph = canonicalSource.firstChild;
    if (!canonicalParagraph) {
      throw new Error("Expected a projected paragraph");
    }
    const pendingParagraph = canonicalParagraph.type.create(
      {
        ...canonicalParagraph.attrs,
        kinsoku: true,
        overflowPunctuation: false,
        suppressAutoHyphens: true,
        hangingIndent: true,
        indentFirstLine: -360,
        spaceBefore: 480,
        spaceAfter: null,
        spacingExplicit: { before: true },
        _originalFormatting: {
          styleId: "Normal",
          kinsoku: true,
          overflowPunctuation: false,
          suppressAutoHyphens: true,
          hangingIndent: true,
          indentFirstLine: -360,
          spaceBefore: 480,
          spacingExplicit: { before: true },
          runProperties: PARAGRAPH_FORMATTING_SAMPLES.runProperties,
          runInWithNext: PARAGRAPH_FORMATTING_SAMPLES.runInWithNext,
        },
        _propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 1, author: "reviewer", date: "2026-09-12T10:00:00.000Z" },
            previousFormatting: Object.fromEntries(
              PPR_CHANGE_SCOPED_FORMATTING_KEYS.filter((key) => key !== "spacingExplicit").flatMap(
                (key) => {
                  const value = PARAGRAPH_FORMATTING_SAMPLES[key];
                  return value === undefined ? [] : [[key, value]];
                },
              ),
            ),
          },
        ],
      },
      canonicalParagraph.content,
    );
    const pending = canonicalSource.type.create(canonicalSource.attrs, [pendingParagraph]);
    const rejected = resolveAllChangesInHeadlessState(
      EditorState.create({
        doc: pending,
        plugins: [createDocumentStylesPlugin(source.package.styles)],
      }),
      "reject",
    ).doc;
    const paragraph = rejected.firstChild;
    expect(paragraph?.attrs["hangingIndent"]).toBe(true);
    expect(paragraph?.attrs["kinsoku"]).toBe(false);
    expect(paragraph?.attrs["overflowPunctuation"]).toBe(true);
    expect(paragraph?.attrs["suppressAutoHyphens"]).toBe(false);
    expect(paragraph?.attrs["spacingExplicit"]).toEqual({ before: true, after: true });
    expect(paragraph?.attrs["_originalFormatting"]).toEqual(sourceFormatting);
    expect(paragraph?.attrs).toEqual(canonicalParagraph.attrs);

    expect(rejected.eq(canonicalSource)).toBe(true);
    const reprojected = toProseDoc(fromProseDoc(rejected, source));
    expect(reprojected.firstChild?.attrs).toEqual(paragraph?.attrs);
    expect(reprojected.eq(rejected)).toBe(true);
  });
});
