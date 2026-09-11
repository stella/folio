import { describe, expect, test } from "bun:test";

import {
  ParagraphPropertySourceContract,
  ParagraphPropertySourceToken,
  ParagraphPropertyTransientTemplateHandle,
} from "../docx/paragraphPropertySourceIdentity";
import {
  canonicalParagraphPropertySourceFingerprintJson,
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  paragraphPropertySourceFingerprintFromFormatting,
  paragraphPropertySourceFingerprintFromParts,
  type AuthoredParagraphProperties,
} from "../docx/paragraphPropertyDescriptor";
import {
  createEditorParagraphPropertyState,
  createImportedParagraphPropertyState,
  createTransientTemplateParagraphPropertyState,
  joinParagraphPropertyStates,
  readPersistableParagraphPropertyState,
  readParagraphPropertyState,
  serializeParagraphPropertyState,
  serializePersistableParagraphPropertyState,
  splitParagraphPropertyState,
  transitionParagraphPropertyState,
} from "./paragraphPropertyState";
import {
  assertParagraphPropertyInvariant,
  createEditorParagraphProperties,
  paragraphAttrsFromExternalDomImport,
  preserveParagraphProperties,
} from "./paragraphPropertyMutation";
import type { ParagraphAttrs } from "./schema/nodes";

const ALL_AUTHORED_PROPERTIES = {
  alignment: "both",
  bidi: false,
  kinsoku: false,
  overflowPunctuation: true,
  spaceBefore: 0,
  spaceAfter: 240,
  lineSpacing: 360,
  lineSpacingRule: "exact",
  snapToGrid: false,
  beforeAutospacing: false,
  afterAutospacing: true,
  spacingExplicit: { before: false, after: true },
  indentLeft: 720,
  indentRight: 0,
  indentFirstLine: -360,
  hangingIndent: true,
  borders: {
    bottom: { style: "single", size: 8, space: 1, shadow: false },
  },
  shading: { fill: { rgb: "00FF00" }, pattern: "clear" },
  tabs: [{ position: 720, alignment: "left", leader: "dot" }],
  keepNext: false,
  keepLines: true,
  widowControl: false,
  pageBreakBefore: true,
  contextualSpacing: false,
  numPr: { ilvl: 2 },
  outlineLevel: 1,
  styleId: "BodyText",
  frame: { hAnchor: "margin", xAlign: "inside", wrap: "around" },
  suppressLineNumbers: false,
  suppressAutoHyphens: true,
} satisfies Required<AuthoredParagraphProperties>;

const token = (): ParagraphPropertySourceToken =>
  ParagraphPropertySourceToken.forOrdinal(
    ParagraphPropertySourceContract.fromDigest("a".repeat(64)),
    { type: "document" },
    0,
  );

describe("mandatory paragraph property state", () => {
  test("the exhaustive authored fixture covers every descriptor-owned pPr property", () => {
    const descriptorKeys = Object.entries(PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR)
      .filter(([, descriptor]) => descriptor.owner === "pPr-base")
      .map(([key]) => key)
      .toSorted();

    expect(Object.keys(ALL_AUTHORED_PROPERTIES).toSorted()).toEqual(descriptorKeys);
  });

  test("raw source identity covers pPr base and paragraph-mark properties only", () => {
    const parsed = paragraphPropertySourceFingerprintFromFormatting({
      ...ALL_AUTHORED_PROPERTIES,
      numberingLevelIndent: {
        type: "latent",
        numId: 9,
        ilvl: 0,
        baseline: { indentLeft: 720 },
      },
      numPrFromStyle: { numId: 9 },
      runProperties: { bold: false },
      runInWithNext: true,
    });
    const composed = paragraphPropertySourceFingerprintFromParts(ALL_AUTHORED_PROPERTIES, {
      runProperties: { bold: false },
      runInWithNext: true,
    });

    expect(parsed).toEqual(composed);
    expect(parsed).toEqual({
      pPrBase: ALL_AUTHORED_PROPERTIES,
      paragraphMark: { runProperties: { bold: false }, runInWithNext: true },
    });
    expect(canonicalParagraphPropertySourceFingerprintJson(parsed)).not.toBe(
      canonicalParagraphPropertySourceFingerprintJson(
        paragraphPropertySourceFingerprintFromParts(ALL_AUTHORED_PROPERTIES, {
          runProperties: { bold: true },
          runInWithNext: true,
        }),
      ),
    );
  });

  test("reifies wire state into nominal, deeply immutable trusted state", () => {
    const serialized = {
      type: "imported",
      token: token().serialized,
      authoredPPr: ALL_AUTHORED_PROPERTIES,
    };

    const result = readParagraphPropertyState(serialized);

    expect(result.status).toBe("valid");
    if (result.status !== "valid") {
      throw new Error("Expected valid paragraph-property state");
    }
    expect(result.value.type).toBe("imported");
    expect(Object.isFrozen(result.value)).toBe(true);
    expect(Object.isFrozen(result.value.authoredPPr)).toBe(true);
    expect(Object.isFrozen(result.value.authoredPPr.borders?.bottom)).toBe(true);
    expect(result.value.authoredPPr.kinsoku).toBe(false);
    expect(result.value.authoredPPr.numPr).toEqual({ ilvl: 2 });
    expect(serializeParagraphPropertyState(result.value)).toEqual(serialized);
  });

  test("keeps empty authored pPr distinct from missing or invalid state", () => {
    expect(readParagraphPropertyState(undefined)).toEqual({ status: "absent" });
    expect(readParagraphPropertyState({ type: "editor-created" }).status).toBe("invalid");

    const state = createEditorParagraphPropertyState({});
    expect(serializePersistableParagraphPropertyState(state)).toEqual({
      type: "editor-created",
      authoredPPr: {},
    });
  });

  test("rejects non-authored, malformed, and ambiguous wire payloads", () => {
    for (const authoredPPr of [
      { runProperties: { bold: true } },
      { numberingLevelIndent: { type: "latent" } },
      { numPrFromStyle: { numId: 1 } },
      { numPr: { numId: -1 } },
      { hangingIndent: true, indentFirstLine: 360 },
      { tabs: [{ position: 720, alignment: "invented" }] },
      { kinsoku: null },
    ]) {
      expect(readParagraphPropertyState({ type: "editor-created", authoredPPr }).status).toBe(
        "invalid",
      );
    }
    expect(
      readParagraphPropertyState({
        type: "editor-created",
        authoredPPr: {},
        token: token().serialized,
      }).status,
    ).toBe("invalid");
  });

  test("patches exact authored values without losing imported identity", () => {
    const sourceToken = token();
    const initial = createImportedParagraphPropertyState(sourceToken, {
      kinsoku: false,
      overflowPunctuation: true,
      numPr: { ilvl: 4 },
    });

    const next = transitionParagraphPropertyState(initial, {
      type: "mutate-authored",
      mutations: [
        { key: "kinsoku", mutation: { type: "set", value: true } },
        { key: "overflowPunctuation", mutation: { type: "remove" } },
        { key: "numPr", mutation: { type: "set", value: { numId: 8 } } },
      ],
    });

    expect(next.type).toBe("imported");
    if (next.type !== "imported") {
      throw new Error("Expected imported paragraph-property state");
    }
    expect(next.token).toBe(sourceToken);
    expect(next.authoredPPr).toEqual({ kinsoku: true, numPr: { numId: 8 } });
    expect(() =>
      transitionParagraphPropertyState(initial, {
        type: "mutate-authored",
        mutations: [
          { key: "kinsoku", mutation: { type: "remove" } },
          { key: "kinsoku", mutation: { type: "set", value: true } },
        ],
      }),
    ).toThrow("Paragraph-property mutation batch contains duplicate key: kinsoku");
  });

  test("models split, copy, join, and transient persistence boundaries explicitly", () => {
    const imported = createImportedParagraphPropertyState(token(), { keepNext: false });
    const split = splitParagraphPropertyState(imported, {
      type: "split-left-created-right-retains",
    });
    expect(split.type).toBe("split-left-created-right-retains");
    expect(split.left.type).toBe("editor-created");
    expect(split.left.authoredPPr).toEqual({ keepNext: false });
    expect(split.right).toBe(imported);
    expect(split.right.type).toBe("imported");
    if (split.right.type !== "imported") {
      throw new Error("Expected the right paragraph mark to retain source ownership");
    }
    expect(split.right.token).toBe(imported.token);

    const copied = transitionParagraphPropertyState(imported, { type: "editor-copy" });
    expect(copied.type).toBe("editor-created");
    expect(
      joinParagraphPropertyStates(imported, copied, {
        type: "join-right-paragraph-mark-retains",
      }),
    ).toBe(copied);
    expect(
      joinParagraphPropertyStates(imported, copied, {
        type: "join-left-paragraph-mark-retains",
      }),
    ).toBe(imported);

    const transient = createTransientTemplateParagraphPropertyState(
      ParagraphPropertyTransientTemplateHandle.forOrdinal(0),
      {},
    );
    expect(serializeParagraphPropertyState(transient)).toEqual({
      type: "transient-template",
      handle: "folio-ppr-template-v1:0",
      authoredPPr: {},
    });
    expect(() => serializePersistableParagraphPropertyState(transient)).toThrow(
      "Transient paragraph-property state cannot cross a persistence boundary",
    );
    expect(
      readPersistableParagraphPropertyState(serializeParagraphPropertyState(transient)).status,
    ).toBe("invalid");
  });
});

describe("paragraph property projection proof", () => {
  test("keeps exact authored partial numbering separate from fieldwise effective numbering", () => {
    const attrs = paragraphAttrsFromExternalDomImport({
      effectiveAttrs: { numPr: { numId: 7, ilvl: 2 }, kinsoku: true },
      authoredAttrs: { numPr: { ilvl: 2 }, kinsoku: false },
      inheritedPPr: { numPr: { numId: 7, ilvl: 0 }, kinsoku: true },
    });

    const state = assertParagraphPropertyInvariant(attrs, { type: "persistence" });
    expect(state.authoredPPr).toEqual({ numPr: { ilvl: 2 }, kinsoku: false });
    expect(attrs.numPr).toEqual({ numId: 7, ilvl: 2 });
    expect(attrs.kinsoku).toBe(false);
  });

  test("capsules expose no attrs and the preserve path rejects governed drift", () => {
    const projection = createEditorParagraphProperties({
      authoredPPr: { alignment: "center" },
      context: { inheritedPPr: {}, numberingLevelIndent: null, numPrFromStyle: null },
    });
    expect(Object.keys(projection)).toEqual([]);

    const attrs = paragraphAttrsFromExternalDomImport({
      effectiveAttrs: { alignment: "center" },
    });
    expect(() => preserveParagraphProperties(attrs, { alignment: "right" })).toThrow(
      "Non-governed paragraph mutation changed governed attr: alignment",
    );
    expect(() =>
      assertParagraphPropertyInvariant(
        { ...attrs, alignment: "right" } as ParagraphAttrs,
        { type: "internal" },
      ),
    ).toThrow("Paragraph-property projection invariant failed for attr: alignment");
  });
});
