import { describe, expect, test } from "bun:test";
import {
  canonicalParagraphPropertySourceFingerprintJson,
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  paragraphFormattingWithPropertySourceFingerprint,
  paragraphPropertySourceFingerprintFromFormatting,
  paragraphPropertySourceFingerprintFromParts,
  type AuthoredParagraphProperties,
} from "@stll/docx-core/model";

import {
  ParagraphPropertyTransientTemplateStore,
} from "../docx/paragraphPropertySourceIdentity";
import {
  createEditorParagraphPropertyState,
  createImportedParagraphPropertyState,
  createTransientTemplateParagraphPropertyState,
  joinParagraphPropertyStates,
  paragraphPropertyStateAttribute,
  readPersistableParagraphPropertyState,
  readParagraphPropertyState,
  serializePersistableParagraphPropertyState,
  splitParagraphPropertyState,
  transitionParagraphPropertyState,
} from "./paragraphPropertyState";
import {
  assertParagraphPropertyInvariant,
  createEditorParagraphProperties,
  paragraphDomGetAttrs,
  paragraphPropertiesFromExternalDomImport,
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

const SOURCE_TOKEN = "p2s:document::0";

const EMPTY_CONTEXT = {
  inheritedPPr: {},
  numberingLevelIndent: null,
  numPrFromStyle: null,
  paragraphMark: { authored: {}, effective: {} },
  spacingInheritance: {},
} as const;

type ExternalImportOptions = Parameters<typeof paragraphPropertiesFromExternalDomImport>[0];
const attrsFromExternalDom = (options: ExternalImportOptions): ParagraphAttrs => {
  const getAttrs = paragraphDomGetAttrs(() =>
    paragraphPropertiesFromExternalDomImport(options),
  );
  const attrs = getAttrs(undefined);
  if (attrs === false) {
    throw new Error("Expected projected paragraph attrs");
  }
  return attrs;
};

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
    expect(
      paragraphFormattingWithPropertySourceFingerprint(
        {
          alignment: "right",
          kinsoku: true,
          numberingLevelIndent: {
            type: "latent",
            numId: 4,
            ilvl: 1,
            baseline: { indentLeft: 720 },
          },
          runProperties: { italic: true },
        },
        paragraphPropertySourceFingerprintFromParts(
          { alignment: "center", kinsoku: false },
          { runProperties: { bold: false }, runInWithNext: false },
        ),
      ),
    ).toEqual({
      alignment: "center",
      kinsoku: false,
      numberingLevelIndent: {
        type: "latent",
        numId: 4,
        ilvl: 1,
        baseline: { indentLeft: 720 },
      },
      runProperties: { bold: false },
      runInWithNext: false,
    });
  });

  test("reifies wire state into nominal, deeply immutable trusted state", () => {
    const serialized = {
      type: "imported",
      token: SOURCE_TOKEN,
      authoredPPr: ALL_AUTHORED_PROPERTIES,
      context: EMPTY_CONTEXT,
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
    expect(serializePersistableParagraphPropertyState(result.value)).toEqual(serialized);
  });

  test("keeps empty authored pPr distinct from missing or invalid state", () => {
    expect(readParagraphPropertyState(undefined)).toEqual({ status: "absent" });
    expect(readParagraphPropertyState({ type: "editor-created" }).status).toBe("invalid");

    const state = createEditorParagraphPropertyState();
    expect(serializePersistableParagraphPropertyState(state)).toEqual({
      type: "editor-created",
      authoredPPr: {},
      context: EMPTY_CONTEXT,
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
      expect(
        readParagraphPropertyState({
          type: "editor-created",
          authoredPPr,
          context: EMPTY_CONTEXT,
        }).status,
      ).toBe("invalid");
    }
    expect(
      readParagraphPropertyState({
        type: "editor-created",
        authoredPPr: {},
        context: EMPTY_CONTEXT,
        token: SOURCE_TOKEN,
      }).status,
    ).toBe("invalid");
  });

  test("patches exact authored values without losing imported identity", () => {
    const sourceToken = SOURCE_TOKEN;
    const initial = createImportedParagraphPropertyState({
      token: sourceToken,
      authoredPPr: {
        kinsoku: false,
        overflowPunctuation: true,
        numPr: { ilvl: 4 },
      },
      context: EMPTY_CONTEXT,
    });

    const next = transitionParagraphPropertyState(initial, {
      type: "update",
      authored: {
        type: "mutate",
        mutations: [
          { key: "kinsoku", mutation: { type: "set", value: true } },
          { key: "overflowPunctuation", mutation: { type: "remove" } },
          { key: "numPr", mutation: { type: "set", value: { numId: 8 } } },
        ],
      },
      context: { type: "preserve" },
    });

    expect(next.type).toBe("imported");
    if (next.type !== "imported") {
      throw new Error("Expected imported paragraph-property state");
    }
    expect(next.token).toBe(sourceToken);
    expect(next.authoredPPr).toEqual({ kinsoku: true, numPr: { numId: 8 } });
    expect(() =>
      transitionParagraphPropertyState(initial, {
        type: "update",
        authored: {
          type: "mutate",
          mutations: [
            { key: "kinsoku", mutation: { type: "remove" } },
            { key: "kinsoku", mutation: { type: "set", value: true } },
          ],
        },
        context: { type: "preserve" },
      }),
    ).toThrow("Paragraph-property mutation batch contains duplicate key: kinsoku");
  });

  test("models split, join, and transient persistence boundaries explicitly", () => {
    const imported = createImportedParagraphPropertyState({
      token: SOURCE_TOKEN,
      authoredPPr: { keepNext: false },
      context: EMPTY_CONTEXT,
    });
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

    const copied = createEditorParagraphPropertyState({
      authoredPPr: imported.authoredPPr,
      context: imported.context,
    });
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

    const templateStore = new ParagraphPropertyTransientTemplateStore<string>(1);
    const handle = templateStore.registerAll(["capture"]).at(0);
    if (!handle) {
      throw new Error("Expected one transient template handle");
    }
    const transient = createTransientTemplateParagraphPropertyState({
      handle,
      authoredPPr: {},
      context: EMPTY_CONTEXT,
    });
    expect(paragraphPropertyStateAttribute(transient)).toBe(transient);
    expect(() => serializePersistableParagraphPropertyState(transient)).toThrow(
      "Transient paragraph-property state cannot cross a persistence boundary",
    );
    expect(() => JSON.stringify(paragraphPropertyStateAttribute(transient))).toThrow(
      "Transient paragraph-property state cannot cross a persistence boundary",
    );
    expect(readPersistableParagraphPropertyState(transient).status).toBe("invalid");
    expect(
      readParagraphPropertyState({
        type: "transient-template",
        handle: "forged-handle",
        authoredPPr: {},
        context: EMPTY_CONTEXT,
      }).status,
    ).toBe("invalid");
  });

  test("rejects malformed context and numbering provenance that can self-prove cache drift", () => {
    const validContext = {
      ...EMPTY_CONTEXT,
      inheritedPPr: { numPr: { numId: 7, ilvl: 0 }, spaceBefore: 120 },
      numberingLevelIndent: {
        type: "owned",
        numId: 7,
        ilvl: 2,
        baseline: { indentLeft: 720 },
        owned: { indentLeft: 720 },
      },
      spacingInheritance: { before: "style" },
    } as const;
    expect(
      readParagraphPropertyState({
        type: "editor-created",
        authoredPPr: { numPr: { ilvl: 2 } },
        context: validContext,
      }).status,
    ).toBe("valid");
    expect(
      readParagraphPropertyState({
        type: "editor-created",
        authoredPPr: { numPr: { ilvl: 3 } },
        context: validContext,
      }).status,
    ).toBe("invalid");
    expect(
      readParagraphPropertyState({
        type: "editor-created",
        authoredPPr: {},
        context: { ...EMPTY_CONTEXT, spacingInheritance: { before: "style" } },
      }).status,
    ).toBe("invalid");
  });

  test("mutates authored and effective paragraph-mark values explicitly", () => {
    const initial = createEditorParagraphPropertyState({
      context: {
        ...EMPTY_CONTEXT,
        paragraphMark: {
          authored: { runProperties: { bold: false }, runInWithNext: false },
          effective: { defaultTextFormatting: { bold: true }, runInWithNext: false },
        },
      },
    });
    const next = transitionParagraphPropertyState(initial, {
      type: "update",
      authored: { type: "preserve" },
      context: {
        type: "mutate-paragraph-mark",
        authored: [
          { key: "runProperties", mutation: { type: "set", value: { italic: true } } },
          { key: "runInWithNext", mutation: { type: "remove" } },
        ],
        effective: [
          {
            key: "defaultTextFormatting",
            mutation: { type: "set", value: { italic: true } },
          },
          { key: "runInWithNext", mutation: { type: "set", value: true } },
        ],
      },
    });

    expect(next.context.paragraphMark).toEqual({
      authored: { runProperties: { italic: true } },
      effective: { defaultTextFormatting: { italic: true }, runInWithNext: true },
    });
  });
});

describe("paragraph property projection proof", () => {
  test("keeps exact authored partial numbering separate from fieldwise effective numbering", () => {
    const attrs = attrsFromExternalDom({
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
      context: {
        inheritedPPr: {},
        numberingLevelIndent: null,
        numPrFromStyle: null,
        paragraphMark: {
          authored: { runProperties: { bold: false }, runInWithNext: true },
          effective: { defaultTextFormatting: { bold: true }, runInWithNext: true },
        },
        spacingInheritance: {},
      },
    });
    expect(Object.keys(projection)).toEqual([]);

    const attrs = attrsFromExternalDom({
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

  test("projects paragraph-mark source and effective run context without spread-through", () => {
    const attrs = attrsFromExternalDom({
      effectiveAttrs: {
        defaultTextFormatting: { bold: true },
        runInWithNext: true,
      },
      authoredAttrs: {
        defaultTextFormatting: { bold: false },
        runInWithNext: false,
      },
    });

    const mark = assertParagraphPropertyInvariant(attrs, { type: "internal" }).context
      .paragraphMark;
    expect(mark.authored).toEqual({
      runProperties: { bold: false },
      runInWithNext: false,
    });
    expect(mark.effective).toEqual({
      defaultTextFormatting: { bold: true },
      runInWithNext: true,
    });
    expect(attrs.defaultTextFormatting).toEqual({ bold: true });
    expect(attrs.runInWithNext).toBe(true);
    expect(() =>
      assertParagraphPropertyInvariant(
        { ...attrs, defaultTextFormatting: { bold: false } } as ParagraphAttrs,
        { type: "internal" },
      ),
    ).toThrow("Paragraph-property projection invariant failed for attr: defaultTextFormatting");
  });
});
