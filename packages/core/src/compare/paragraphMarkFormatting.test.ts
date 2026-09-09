import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { sameParagraphMarkRunFormatting } from "../ai-edits/snapshot";
import { assignParagraphMarkRunPropertyChanges } from "../docx/paragraphMarkRunPropertyChanges";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { ParagraphFormatting, TextFormatting, Theme } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";
import { planStoryCompare } from "./plan";

const OPTIONS = { author: "compare", timestamp: "2024-03-01T00:00:00.000Z" } as const;
const PARAGRAPH_MARK_STYLE_ID = "ParagraphMarkCharacter";

type ParagraphMarkFormattingValues = {
  [Property in keyof TextFormatting]-?: NonNullable<TextFormatting[Property]>;
};

/** One serialization-stable value for every modeled CT_ParaRPr property. */
const PARAGRAPH_MARK_FORMATTING_VALUES = {
  bold: false,
  boldCs: false,
  italic: false,
  italicCs: false,
  underline: { style: "none", color: { themeColor: "accent2", themeTint: "80" } },
  strike: false,
  doubleStrike: false,
  vertAlign: "baseline",
  smallCaps: false,
  allCaps: false,
  hidden: false,
  color: { themeColor: "accent1", themeShade: "40" },
  highlight: "none",
  shading: {
    pattern: "pct25",
    color: { rgb: "112233" },
    fill: { themeColor: "accent3", themeTint: "80" },
  },
  fontSize: 23,
  fontSizeCs: 27,
  fontFamily: {
    ascii: "Arial",
    hAnsi: "Calibri",
    eastAsia: "MS Mincho",
    cs: "Arial",
    hint: "eastAsia",
    asciiTheme: "majorAscii",
    hAnsiTheme: "minorHAnsi",
    eastAsiaTheme: "majorEastAsia",
    csTheme: "minorBidi",
  },
  language: { val: "en-US", eastAsia: "ja-JP", bidi: "ar-SA" },
  spacing: 17,
  position: -2,
  scale: 88,
  kerning: 24,
  effect: "shimmer",
  emphasisMark: "circle",
  emboss: false,
  imprint: false,
  outline: false,
  shadow: false,
  rtl: false,
  cs: false,
  styleId: PARAGRAPH_MARK_STYLE_ID,
} as const satisfies ParagraphMarkFormattingValues;

const formattingCases = Object.entries(PARAGRAPH_MARK_FORMATTING_VALUES).map(
  ([property, value]) => ({ property, value }),
);

const createParagraphMarkFormattingDocx = async (
  runProperties?: TextFormatting,
  theme?: Theme,
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  if (theme) {
    document.package.theme = theme;
  }
  if (runProperties?.styleId === PARAGRAPH_MARK_STYLE_ID) {
    document.package.styles = {
      styles: [
        {
          styleId: PARAGRAPH_MARK_STYLE_ID,
          type: "character",
          name: "Paragraph mark character",
        },
      ],
    };
  }
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      ...(runProperties ? { formatting: { runProperties } } : {}),
      content: [{ type: "run", content: [{ type: "text", text: "Same text" }] }],
    },
  ];
  return await createDocx(document);
};

const createParagraphMarkCharacterStyleDocx = async (
  styleFormatting: TextFormatting,
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      {
        styleId: PARAGRAPH_MARK_STYLE_ID,
        type: "character",
        name: "Paragraph mark character",
        rPr: styleFormatting,
      },
    ],
  };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      formatting: { runProperties: { styleId: PARAGRAPH_MARK_STYLE_ID } },
      content: [{ type: "run", content: [{ type: "text", text: "Same text" }] }],
    },
  ];
  return await createDocx(document);
};

const parsedParagraphMarkFormatting = async (
  buffer: ArrayBuffer,
): Promise<TextFormatting | undefined> => {
  const document = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected the fixture to contain one paragraph");
  }
  return paragraph.formatting?.runProperties;
};

const mainDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const documentXml = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!documentXml) {
    throw new Error("Expected word/document.xml");
  }
  return await documentXml.async("string");
};

const formattingWith = (property: string, value: unknown): TextFormatting => {
  const formatting: TextFormatting = {};
  Reflect.set(formatting, property, value);
  return formatting;
};

const createParagraphSequenceDocx = async (
  paragraphs: readonly {
    id: string;
    text: string;
    formatting?: ParagraphFormatting;
    runProperties?: TextFormatting;
  }[],
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = paragraphs.map(
    ({ id, text, formatting, runProperties }) => ({
      type: "paragraph",
      paraId: id,
      ...(formatting || runProperties
        ? {
            formatting: {
              ...formatting,
              ...(runProperties ? { runProperties } : {}),
            },
          }
        : {}),
      content: [{ type: "run", content: [{ type: "text", text }] }],
    }),
  );
  return await createDocx(document);
};

const parsedParagraphMarkFormattingSequence = async (
  buffer: ArrayBuffer,
): Promise<(TextFormatting | undefined)[]> => {
  const document = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  return document.package.document.content.flatMap((block) =>
    block.type === "paragraph" ? [block.formatting?.runProperties] : [],
  );
};

type ExpectParagraphMarkRoundTripOptions = {
  base: ArrayBuffer;
  target: ArrayBuffer;
  before: TextFormatting | undefined;
  after: TextFormatting | undefined;
};

const expectParagraphMarkRoundTrip = async ({
  base,
  target,
  before,
  after,
}: ExpectParagraphMarkRoundTripOptions): Promise<void> => {
  const result = await compareDocx(base, target, OPTIONS);
  if (result.isErr()) {
    throw result.error;
  }
  expect(result.value.verification).toEqual({ status: "verified" });
  expect(result.value.changes).toEqual([
    expect.objectContaining({
      kind: "paragraph-mark-format",
      properties: after ?? null,
    }),
  ]);

  const pending = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(pending.getChanges().map(({ type }) => type)).toEqual(["paragraphPropertiesChanged"]);
  expect(await parsedParagraphMarkFormatting(result.value.buffer)).toEqual(after);
  expect((await mainDocumentXml(result.value.buffer)).match(/<w:rPrChange\b/gu)).toHaveLength(1);

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(accepting.acceptAll()).toBe(1);
  const accepted = await accepting.toBuffer();
  expect(await parsedParagraphMarkFormatting(accepted)).toEqual(after);
  expect(await mainDocumentXml(accepted)).not.toContain("<w:rPrChange");
  expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(rejecting.rejectAll()).toBe(1);
  const rejected = await rejecting.toBuffer();
  expect(await parsedParagraphMarkFormatting(rejected)).toEqual(before);
  expect(await mainDocumentXml(rejected)).not.toContain("<w:rPrChange");
  expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
};

describe("paragraph-mark formatting comparison", () => {
  test("compares exact formatting independent of object-property insertion order", () => {
    const left = { bold: false, color: { rgb: "112233" } };
    const right = { color: { rgb: "112233" }, bold: false };
    const leftProjection = { formatting: left, orderedSignature: JSON.stringify(left) };
    const rightProjection = { formatting: right, orderedSignature: JSON.stringify(right) };

    expect(leftProjection.orderedSignature).not.toBe(rightProjection.orderedSignature);
    expect(sameParagraphMarkRunFormatting(leftProjection, rightProjection)).toBe(true);
  });

  test.each(formattingCases)(
    "preserves or truthfully refuses target-only $property",
    async ({ property, value }) => {
      const base = await createParagraphMarkFormattingDocx();
      const targetFormatting = formattingWith(property, value);
      const target = await createParagraphMarkFormattingDocx(targetFormatting);

      // The fixture must reach the guard: the target package itself retains
      // the direct paragraph-mark property while the base has none.
      expect(await parsedParagraphMarkFormatting(base)).toBeUndefined();
      expect(await parsedParagraphMarkFormatting(target)).toEqual(targetFormatting);

      if (property === "fontSize" || property === "styleId") {
        const refused = await compareDocx(base, target, OPTIONS);
        expect(refused.isErr()).toBe(true);
        if (refused.isErr()) {
          expect(refused.error._tag).toBe("CompareDocxRoundTripError");
          if (refused.error._tag === "CompareDocxRoundTripError") {
            expect(refused.error.cause).toBe("paragraph-mark-format");
          }
        }
      }

      const result = await compareDocx(base, target, {
        ...OPTIONS,
        onUnverified: "emit",
      });
      if (result.isErr()) {
        throw result.error;
      }
      expect(await parsedParagraphMarkFormatting(result.value.buffer)).toEqual(targetFormatting);
      const paragraphMarkChanges = result.value.changes.filter(
        (change) => change.kind === "paragraph-mark-format",
      );
      expect(paragraphMarkChanges).toEqual([
        expect.objectContaining({
          kind: "paragraph-mark-format",
          properties: targetFormatting,
        }),
      ]);
      if (property === "fontSize" || property === "styleId") {
        expect(result.value.verification.status).toBe("unverified");
        if (result.value.verification.status === "unverified") {
          expect(result.value.verification.failures).toContainEqual(
            expect.objectContaining({
              invariant: "accept-reproduces-target",
              cause: "paragraph-mark-format",
            }),
          );
        }
      } else {
        expect(result.value.verification).toEqual({ status: "verified" });
      }

      const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(accepting.acceptAll()).toBeGreaterThan(0);
      const accepted = await accepting.toBuffer();
      expect(await parsedParagraphMarkFormatting(accepted)).toEqual(targetFormatting);
      expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);

      const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(rejecting.rejectAll()).toBeGreaterThan(0);
      const rejected = await rejecting.toBuffer();
      expect(await parsedParagraphMarkFormatting(rejected)).toBeUndefined();
      expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
    },
  );

  test.each([
    { label: "all caps", formatting: { allCaps: true } },
    {
      label: "literal font family",
      formatting: { fontFamily: { ascii: "Arial", hAnsi: "Arial" } },
    },
  ] satisfies readonly { label: string; formatting: TextFormatting }[])(
    "round-trips target-only $label through accept, reject, save, and reopen",
    async ({ formatting }) => {
      const base = await createParagraphMarkFormattingDocx();
      const target = await createParagraphMarkFormattingDocx(formatting);
      await expectParagraphMarkRoundTrip({
        base,
        target,
        before: undefined,
        after: formatting,
      });
    },
  );

  test("keeps an exact paragraph-mark formatting match verified", async () => {
    const base = await createParagraphMarkFormattingDocx(PARAGRAPH_MARK_FORMATTING_VALUES);
    const target = await createParagraphMarkFormattingDocx(PARAGRAPH_MARK_FORMATTING_VALUES);
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.changes).toEqual([]);
    expect(result.value.verification).toEqual({ status: "verified" });
  });

  test.each([
    {
      label: "color slot",
      formatting: { color: { themeColor: "accent1" } },
      baseTheme: {
        name: "Base",
        colorScheme: { accent1: "0000FF" },
      },
      targetTheme: {
        name: "Target",
        colorScheme: { accent1: "FF0000" },
      },
    },
    {
      label: "font slot",
      formatting: {
        fontFamily: { asciiTheme: "majorAscii", hAnsiTheme: "majorHAnsi" },
      },
      baseTheme: {
        name: "Base",
        fontScheme: { majorFont: { latin: "Base Font" } },
      },
      targetTheme: {
        name: "Target",
        fontScheme: { majorFont: { latin: "Target Font" } },
      },
    },
  ] satisfies readonly {
    label: string;
    formatting: TextFormatting;
    baseTheme: Theme;
    targetTheme: Theme;
  }[])(
    "refuses a target paragraph-mark $label with different package semantics",
    async ({ formatting, baseTheme, targetTheme }) => {
      const base = await createParagraphMarkFormattingDocx(formatting, baseTheme);
      const target = await createParagraphMarkFormattingDocx(formatting, targetTheme);

      const refused = await compareDocx(base, target, OPTIONS);
      expect(refused.isErr()).toBe(true);
      if (refused.isErr()) {
        expect(refused.error._tag).toBe("CompareDocxRoundTripError");
        if (refused.error._tag === "CompareDocxRoundTripError") {
          expect(refused.error.cause).toBe("paragraph-mark-format");
          expect(refused.error.failures).toContainEqual(
            expect.objectContaining({
              invariant: "accept-reproduces-target",
              cause: "paragraph-mark-format",
              detail:
                "target paragraph-mark theme references have different semantics in the base package",
            }),
          );
        }
      }

      const emitted = await compareDocx(base, target, { ...OPTIONS, onUnverified: "emit" });
      if (emitted.isErr()) {
        throw emitted.error;
      }
      expect(emitted.value.verification.status).toBe("unverified");
      if (emitted.value.verification.status === "unverified") {
        expect(emitted.value.verification.failures).toContainEqual(
          expect.objectContaining({
            invariant: "accept-reproduces-target",
            cause: "paragraph-mark-format",
          }),
        );
      }
      const output = await parseDocx(emitted.value.buffer, {
        detectVariables: false,
        preloadFonts: false,
      });
      expect(output.package.theme?.name).toBe("Base");
      expect(await parsedParagraphMarkFormatting(emitted.value.buffer)).toEqual(
        await parsedParagraphMarkFormatting(target),
      );
    },
  );

  test("keeps unrelated theme differences outside the paragraph-mark guard", async () => {
    const formatting = { color: { themeColor: "accent1" as const } };
    const base = await createParagraphMarkFormattingDocx(formatting, {
      name: "Base",
      colorScheme: { accent1: "112233", accent2: "000000" },
    });
    const target = await createParagraphMarkFormattingDocx(formatting, {
      name: "Target",
      colorScheme: { accent1: "112233", accent2: "FFFFFF" },
    });

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes).toEqual([]);
  });

  test("fails closed when an unchanged character-style reference changes meaning", async () => {
    const base = await createParagraphMarkCharacterStyleDocx({ bold: true });
    const target = await createParagraphMarkCharacterStyleDocx({ italic: true });

    const refused = await compareDocx(base, target, OPTIONS);
    expect(refused.isErr()).toBe(true);
    if (refused.isErr() && refused.error._tag === "CompareDocxRoundTripError") {
      expect(refused.error.cause).toBe("paragraph-mark-format");
      expect(refused.error.failures).toContainEqual(
        expect.objectContaining({
          invariant: "accept-reproduces-target",
          cause: "paragraph-mark-format",
          detail:
            "target paragraph-mark character styles have different semantics in the base package",
        }),
      );
    }

    const emitted = await compareDocx(base, target, { ...OPTIONS, onUnverified: "emit" });
    if (emitted.isErr()) {
      throw emitted.error;
    }
    // The changed style cascade also changes the paragraph's visible runs;
    // that independently reported range does not make the mark dependency safe.
    expect(emitted.value.changes.map(({ kind }) => kind)).toEqual(["format"]);
    expect(emitted.value.verification.status).toBe("unverified");
    if (emitted.value.verification.status === "unverified") {
      expect(emitted.value.verification.failures).toContainEqual(
        expect.objectContaining({
          cause: "paragraph-mark-format",
          detail:
            "target paragraph-mark character styles have different semantics in the base package",
        }),
      );
    }
  });

  test("counts a paired paragraph-mark formatting carrier against the operation budget", async () => {
    const base = await FolioDocxReviewer.fromBuffer(await createParagraphMarkFormattingDocx());
    const target = await FolioDocxReviewer.fromBuffer(
      await createParagraphMarkFormattingDocx({ allCaps: true }),
    );
    const input = {
      story: { type: "main" as const },
      baseSnapshot: base.snapshot(),
      targetSnapshot: target.snapshot(),
    };

    expect(planStoryCompare({ ...input, maxOperations: 0 })).toBeNull();
    expect(planStoryCompare({ ...input, maxOperations: 1 })).not.toBeNull();
  });

  test("round-trips removal of a base-only paragraph-mark property", async () => {
    const before = { allCaps: true };
    const base = await createParagraphMarkFormattingDocx(before);
    const target = await createParagraphMarkFormattingDocx();
    await expectParagraphMarkRoundTrip({ base, target, before, after: undefined });
  });

  test("reports direct paragraph and paragraph-mark properties independently", async () => {
    const base = await createParagraphSequenceDocx([
      {
        id: "A1000001",
        text: "Same text",
        formatting: { alignment: "left" },
        runProperties: { allCaps: true },
      },
    ]);
    const target = await createParagraphSequenceDocx([
      {
        id: "A1000001",
        text: "Same text",
        formatting: { alignment: "right" },
        runProperties: { smallCaps: true },
      },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes).toEqual([
      expect.objectContaining({
        kind: "paragraph-format",
        properties: { alignment: "right" },
      }),
      expect.objectContaining({
        kind: "paragraph-mark-format",
        properties: { smallCaps: true },
      }),
    ]);
  });

  test("preserves an inserted paragraph's target paragraph-mark formatting", async () => {
    const base = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000003", text: "Z" },
    ]);
    const target = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000002", text: "B", runProperties: { allCaps: true } },
      { id: "A1000003", text: "Z" },
    ]);
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["insert"]);
    expect(await parsedParagraphMarkFormattingSequence(result.value.buffer)).toEqual([
      undefined,
      { allCaps: true },
      undefined,
    ]);

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await accepting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(accepted)).toEqual([
      undefined,
      { allCaps: true },
      undefined,
    ]);
    expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await rejecting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(rejected)).toEqual([undefined, undefined]);
    expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
  });

  test("keeps inserted and paired paragraph-mark formatting distinct at one boundary", async () => {
    const insertedFormatting = { allCaps: true };
    const pairedFormatting = { fontFamily: { ascii: "Arial", hAnsi: "Arial" } };
    const base = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000003", text: "Z" },
    ]);
    const target = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000002", text: "B", runProperties: insertedFormatting },
      { id: "A1000003", text: "Z", runProperties: pairedFormatting },
    ]);
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await accepting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(accepted)).toEqual([
      undefined,
      insertedFormatting,
      pairedFormatting,
    ]);
    expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await rejecting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(rejected)).toEqual([undefined, undefined]);
    expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
  });

  test.each([
    {
      label: "deleted paragraph",
      base: [
        { id: "A1000001", text: "A" },
        { id: "A1000002", text: "Delete", runProperties: { allCaps: true } },
        { id: "A1000003", text: "Z" },
      ],
      target: [
        { id: "A1000001", text: "A" },
        { id: "A1000003", text: "Z" },
      ],
      expectedKind: "delete",
    },
    {
      label: "split paragraph",
      base: [
        { id: "A1000001", text: "A" },
        { id: "A1000002", text: "Alpha beta", runProperties: { allCaps: true } },
        { id: "A1000004", text: "Z" },
      ],
      target: [
        { id: "A1000001", text: "A" },
        { id: "A1000002", text: "Alpha", runProperties: { italic: true } },
        { id: "A1000003", text: "beta", runProperties: { bold: true } },
        { id: "A1000004", text: "Z" },
      ],
      expectedKind: "split",
    },
    {
      label: "merged paragraph",
      base: [
        { id: "A1000001", text: "A" },
        { id: "A1000002", text: "Alpha", runProperties: { allCaps: true } },
        { id: "A1000003", text: "beta", runProperties: { bold: true } },
        { id: "A1000004", text: "Z" },
      ],
      target: [
        { id: "A1000001", text: "A" },
        { id: "A1000002", text: "Alpha beta", runProperties: { italic: true } },
        { id: "A1000004", text: "Z" },
      ],
      expectedKind: "merge",
    },
  ] satisfies readonly {
    label: string;
    base: readonly { id: string; text: string; runProperties?: TextFormatting }[];
    target: readonly { id: string; text: string; runProperties?: TextFormatting }[];
    expectedKind: "delete" | "split" | "merge";
  }[])("does not double-report paragraph-mark formatting for a $label", async (scenario) => {
    const base = await createParagraphSequenceDocx(scenario.base);
    const target = await createParagraphSequenceDocx(scenario.target);
    const result = await compareDocx(base, target, { ...OPTIONS, onUnverified: "emit" });
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.changes.map(({ kind }) => kind)).toEqual([scenario.expectedKind]);
  });

  test("restores a formatted final paragraph mark through an appended insertion chain", async () => {
    const before = { allCaps: true };
    const base = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000002", text: "Z", runProperties: before },
    ]);
    const target = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000002", text: "Z" },
      { id: "A1000003", text: "B" },
      { id: "A1000004", text: "C" },
    ]);
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await accepting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(accepted)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(await mainDocumentXml(accepted)).not.toContain("<w:rPrChange");
    expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await rejecting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(rejected)).toEqual([undefined, before]);
    expect(await mainDocumentXml(rejected)).not.toContain("<w:rPrChange");
    expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
  });

  test("moves target paragraph-mark formatting onto a trailing deletion carrier", async () => {
    const targetFormatting = { allCaps: true };
    const carrierFormatting = { fontFamily: { ascii: "Arial", hAnsi: "Arial" } };
    const base = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000002", text: "" },
      { id: "A1000003", text: "Delete me", runProperties: carrierFormatting },
    ]);
    const target = await createParagraphSequenceDocx([
      { id: "A1000001", text: "A" },
      { id: "A1000002", text: "", runProperties: targetFormatting },
    ]);
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await accepting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(accepted)).toEqual([
      undefined,
      targetFormatting,
    ]);
    expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await rejecting.toBuffer();
    expect(await parsedParagraphMarkFormattingSequence(rejected)).toEqual([
      undefined,
      undefined,
      carrierFormatting,
    ]);
    expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
  });

  test("fails closed before serializing a second paragraph-mark rPrChange", async () => {
    const document = createEmptyDocument();
    const paragraph = {
      type: "paragraph" as const,
      formatting: { runProperties: { bold: true } },
      content: [{ type: "run" as const, content: [{ type: "text" as const, text: "Text" }] }],
    };
    document.package.document.content = [paragraph];
    assignParagraphMarkRunPropertyChanges(paragraph, [
      {
        type: "runPropertyChange",
        info: { id: 1, author: "A" },
        previousFormatting: { italic: true },
      },
      {
        type: "runPropertyChange",
        info: { id: 2, author: "B" },
        previousFormatting: { underline: { style: "single" } },
      },
    ]);

    await expect(createDocx(document)).rejects.toThrow(
      "A paragraph mark cannot serialize more than one w:rPrChange",
    );
  });
});
