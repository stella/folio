import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIParagraphSpacing } from "../ai-edits/types";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-08T00:00:00.000Z" } as const;
const TEXT = "The agreement remains effective for the stated term.";
const INSERTED_TEXT = "The renewal period starts after written notice.";
const STYLE_ID = "SpacedBody";

type SpacingDocumentOptions = {
  directSpacing?: FolioAIParagraphSpacing;
  inheritedSpacing?: FolioAIParagraphSpacing;
};

type SpacingState = {
  label: string;
  spacing: FolioAIParagraphSpacing | undefined;
};

const spacingStates = (
  property: keyof FolioAIParagraphSpacing,
  values: readonly (number | string | boolean)[],
): SpacingState[] => [
  { label: "absent", spacing: undefined },
  ...values.map((value) => ({ label: String(value), spacing: { [property]: value } })),
];

const SPACING_STATES_BY_PROPERTY = [
  { property: "spaceBefore", states: spacingStates("spaceBefore", [0, 240]) },
  { property: "spaceAfter", states: spacingStates("spaceAfter", [0, 360]) },
  { property: "lineSpacing", states: spacingStates("lineSpacing", [-480, 0, 480]) },
  {
    property: "lineSpacingRule",
    states: spacingStates("lineSpacingRule", ["auto", "exact", "atLeast"]),
  },
  {
    property: "beforeAutospacing",
    states: spacingStates("beforeAutospacing", [false, true]),
  },
  {
    property: "afterAutospacing",
    states: spacingStates("afterAutospacing", [false, true]),
  },
] as const;

const DIRECT_SPACING_MATRIX = SPACING_STATES_BY_PROPERTY.flatMap(({ property, states }) =>
  states.flatMap((before) =>
    states
      .filter((after) => JSON.stringify(after.spacing) !== JSON.stringify(before.spacing))
      .map((after) => ({
        after: after.spacing,
        before: before.spacing,
        transition: `${property}: ${before.label} to ${after.label}`,
      })),
  ),
);

const EXPLICIT_SPACING_STATES = SPACING_STATES_BY_PROPERTY.flatMap(({ property, states }) =>
  states
    .filter(({ spacing }) => spacing !== undefined)
    .map(({ label, spacing }) => ({ label: `${property}: ${label}`, spacing })),
);

const FULL_SPACING = {
  spaceBefore: 0,
  spaceAfter: 360,
  lineSpacing: 480,
  lineSpacingRule: "exact",
  beforeAutospacing: false,
  afterAutospacing: true,
} as const satisfies FolioAIParagraphSpacing;

const paragraphModel = (
  text: string,
  paraId: string,
  directSpacing?: FolioAIParagraphSpacing,
): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  formatting: { styleId: STYLE_ID, ...directSpacing },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const spacingDocumentModel = ({ directSpacing, inheritedSpacing }: SpacingDocumentOptions) => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      {
        type: "paragraph",
        styleId: STYLE_ID,
        name: "Spaced Body",
        basedOn: "Normal",
        ...(inheritedSpacing === undefined ? {} : { pPr: inheritedSpacing }),
      },
    ],
  };
  document.package.document.content = [paragraphModel(TEXT, "12345678", directSpacing)];
  return document;
};

const documentCache = new Map<string, Promise<ArrayBuffer>>();

const spacingDocument = (options: SpacingDocumentOptions): Promise<ArrayBuffer> => {
  const key = JSON.stringify(options);
  const cached = documentCache.get(key);
  if (cached) {
    return cached;
  }
  const created = createDocx(spacingDocumentModel(options));
  documentCache.set(key, created);
  return created;
};

const mainDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    panic("expected word/document.xml");
  }
  return await part.async("string");
};

const paragraphXmls = (xml: string): string[] =>
  xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gu) ?? [];

const firstParagraphXml = (xml: string): string => {
  const paragraph = paragraphXmls(xml).at(0);
  if (!paragraph) {
    panic("expected a body paragraph");
  }
  return paragraph;
};

const spacingXml = (spacing: FolioAIParagraphSpacing | undefined): string => {
  if (spacing === undefined) {
    return "";
  }
  const attributes = [
    spacing.spaceBefore === undefined ? "" : `w:before="${String(spacing.spaceBefore)}"`,
    spacing.spaceAfter === undefined ? "" : `w:after="${String(spacing.spaceAfter)}"`,
    spacing.lineSpacing === undefined ? "" : `w:line="${String(spacing.lineSpacing)}"`,
    spacing.lineSpacingRule === undefined ? "" : `w:lineRule="${spacing.lineSpacingRule}"`,
    spacing.beforeAutospacing === undefined
      ? ""
      : `w:beforeAutospacing="${spacing.beforeAutospacing ? "1" : "0"}"`,
    spacing.afterAutospacing === undefined
      ? ""
      : `w:afterAutospacing="${spacing.afterAutospacing ? "1" : "0"}"`,
  ].filter((attribute) => attribute.length > 0);
  return `<w:spacing ${attributes.join(" ")}/>`;
};

const expectedParagraphProperties = (spacing: FolioAIParagraphSpacing | undefined): string =>
  `<w:pStyle w:val="${STYLE_ID}"/>${spacingXml(spacing)}`;

const trackedParagraphPropertyParts = (
  paragraph: string,
): { current: string; previous: string } => {
  const propertiesStart = paragraph.indexOf("<w:pPr>");
  const changeStart = paragraph.indexOf("<w:pPrChange ");
  const previous = paragraph
    .match(/<w:pPrChange\b[^>]*><w:pPr>([\s\S]*?)<\/w:pPr><\/w:pPrChange>/u)
    ?.at(1);
  if (propertiesStart < 0 || changeStart < 0 || previous === undefined) {
    panic("expected w:pPr containing one w:pPrChange");
  }
  return { current: paragraph.slice(propertiesStart + "<w:pPr>".length, changeStart), previous };
};

const untrackedParagraphProperties = (paragraph: string): string =>
  paragraph.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/u)?.at(1) ?? "";

const expectDirectSpacing = (
  reviewer: FolioDocxReviewer,
  expected: FolioAIParagraphSpacing | undefined,
  index = 0,
): void => {
  expect(reviewer.snapshot().blocks.at(index)?.directSpacing).toEqual(expected);
};

const expectCompareRoundTrip = async ({
  before,
  after,
}: {
  before: SpacingDocumentOptions;
  after: SpacingDocumentOptions;
}): Promise<void> => {
  const result = await compareDocx(
    await spacingDocument(before),
    await spacingDocument(after),
    OPTIONS,
  );
  if (result.isErr()) {
    throw result.error;
  }

  expect(result.value.verification).toEqual({ status: "verified" });
  expect(result.value.unsupported).toEqual([]);
  expect(result.value.changes).toEqual([
    expect.objectContaining({
      kind: "paragraph-format",
      properties: { spacing: after.directSpacing ?? null },
    }),
  ]);

  const pendingXml = await mainDocumentXml(result.value.buffer);
  expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
  expect(trackedParagraphPropertyParts(firstParagraphXml(pendingXml))).toEqual({
    current: expectedParagraphProperties(after.directSpacing),
    previous: expectedParagraphProperties(before.directSpacing),
  });

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expectDirectSpacing(accepting, after.directSpacing);
  expect(accepting.acceptAll()).toBe(1);
  expectDirectSpacing(accepting, after.directSpacing);
  const accepted = await accepting.toBuffer();
  const acceptedXml = firstParagraphXml(await mainDocumentXml(accepted));
  expect(acceptedXml).not.toContain("<w:pPrChange");
  expect(untrackedParagraphProperties(acceptedXml)).toBe(
    expectedParagraphProperties(after.directSpacing),
  );
  expectDirectSpacing(await FolioDocxReviewer.fromBuffer(accepted), after.directSpacing);

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expectDirectSpacing(rejecting, after.directSpacing);
  expect(rejecting.rejectAll()).toBe(1);
  expectDirectSpacing(rejecting, before.directSpacing);
  const rejected = await rejecting.toBuffer();
  const rejectedXml = firstParagraphXml(await mainDocumentXml(rejected));
  expect(rejectedXml).not.toContain("<w:pPrChange");
  expect(untrackedParagraphProperties(rejectedXml)).toBe(
    expectedParagraphProperties(before.directSpacing),
  );
  expectDirectSpacing(await FolioDocxReviewer.fromBuffer(rejected), before.directSpacing);
};

describe("paragraph spacing comparison", () => {
  test.each(DIRECT_SPACING_MATRIX)(
    "round-trips every direct spacing state transition: $transition",
    async ({ before, after }) => {
      await expectCompareRoundTrip({
        before: { directSpacing: before },
        after: { directSpacing: after },
      });
    },
  );

  test.each(EXPLICIT_SPACING_STATES)(
    "distinguishes inherited spacing from equal direct spacing: $label",
    async ({ spacing }) => {
      await expectCompareRoundTrip({
        before: { inheritedSpacing: spacing },
        after: { directSpacing: spacing, inheritedSpacing: spacing },
      });
      await expectCompareRoundTrip({
        before: { directSpacing: spacing, inheritedSpacing: spacing },
        after: { inheritedSpacing: spacing },
      });
    },
  );

  test("round-trips the complete direct w:spacing cluster atomically", async () => {
    await expectCompareRoundTrip({
      before: { directSpacing: { spaceBefore: 120, lineSpacingRule: "auto" } },
      after: { directSpacing: FULL_SPACING },
    });
  });

  test("does not project style-inherited spacing as direct formatting", async () => {
    const source = await spacingDocument({ inheritedSpacing: FULL_SPACING });
    const reviewer = await FolioDocxReviewer.fromBuffer(source);
    expectDirectSpacing(reviewer, undefined);

    const saved = await reviewer.toBuffer();
    expect(untrackedParagraphProperties(firstParagraphXml(await mainDocumentXml(saved)))).toBe(
      expectedParagraphProperties(undefined),
    );
    expectDirectSpacing(await FolioDocxReviewer.fromBuffer(saved), undefined);

    const result = await compareDocx(source, source, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes).toEqual([]);
    expect(result.value.verification).toEqual({ status: "verified" });
  });

  test("carries inserted direct spacing through accept and reject save-reopen", async () => {
    const inheritedSpacing = { spaceAfter: 360, afterAutospacing: false } as const;
    const baseModel = spacingDocumentModel({ inheritedSpacing });
    const targetModel = spacingDocumentModel({ inheritedSpacing });
    targetModel.package.document.content.push(
      paragraphModel(INSERTED_TEXT, "23456789", FULL_SPACING),
    );
    const result = await compareDocx(
      await createDocx(baseModel),
      await createDocx(targetModel),
      OPTIONS,
    );
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes).toEqual([
      expect.objectContaining({ kind: "insert", after: INSERTED_TEXT }),
    ]);
    const pendingXml = await mainDocumentXml(result.value.buffer);
    expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
    expectDirectSpacing(await FolioDocxReviewer.fromBuffer(result.value.buffer), FULL_SPACING, 1);

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await accepting.toBuffer();
    const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
    expect(reopenedAccepted.snapshot().blocks.map(({ text }) => text)).toEqual([
      TEXT,
      INSERTED_TEXT,
    ]);
    expectDirectSpacing(reopenedAccepted, FULL_SPACING, 1);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await rejecting.toBuffer();
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
    expect(reopenedRejected.snapshot().blocks.map(({ text }) => text)).toEqual([TEXT]);
    expectDirectSpacing(reopenedRejected, undefined);
  });

  test("fails closed when a second tracked paragraph-format edit targets a pending pPrChange", async () => {
    const originalSpacing = { spaceBefore: 120, afterAutospacing: false } as const;
    const firstTarget = { spaceAfter: 0, beforeAutospacing: false } as const;
    const source = await spacingDocument({ directSpacing: originalSpacing });
    const reviewer = await FolioDocxReviewer.fromBuffer(source);
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    if (!blockId) {
      panic("expected a paragraph block id");
    }

    const first = reviewer.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "first-spacing",
          type: "setBlockParagraphProperties",
          blockId,
          properties: { spacing: firstTarget },
        },
      ],
    });
    expect(first.skipped).toEqual([]);
    expectDirectSpacing(reviewer, firstTarget);

    const pending = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(pending);
    const second = reopened.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      atomic: true,
      operations: [
        {
          id: "second-alignment",
          type: "setBlockParagraphProperties",
          blockId,
          properties: { alignment: "center" },
        },
      ],
    });
    expect(second.status).toBe("rejected");
    expect(second.applied).toEqual([]);
    expect(second.skipped).toEqual([
      { id: "second-alignment", reason: "pendingParagraphPropertyChange" },
    ]);
    expect(second.issues).toEqual([
      expect.objectContaining({
        code: "pendingParagraphPropertyChange",
        recovery: "resolveTrackedChange",
        retryable: false,
      }),
    ]);
    expect(reopened.snapshot().blocks.at(0)?.directAlignment).toBeUndefined();
    expectDirectSpacing(reopened, firstTarget);

    const replacement = reopened.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "second-replacement-style",
          type: "replaceBlock",
          blockId,
          text: TEXT,
          styleId: null,
        },
      ],
    });
    expect(replacement.applied).toEqual([]);
    expect(replacement.skipped).toEqual([
      { id: "second-replacement-style", reason: "pendingParagraphPropertyChange" },
    ]);

    const stillPending = await reopened.toBuffer();
    const stillPendingXml = await mainDocumentXml(stillPending);
    expect(stillPendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
    expect(trackedParagraphPropertyParts(firstParagraphXml(stillPendingXml))).toEqual({
      current: expectedParagraphProperties(firstTarget),
      previous: expectedParagraphProperties(originalSpacing),
    });

    const accepting = await FolioDocxReviewer.fromBuffer(stillPending);
    expect(accepting.acceptAll()).toBe(1);
    expectDirectSpacing(
      await FolioDocxReviewer.fromBuffer(await accepting.toBuffer()),
      firstTarget,
    );

    const rejecting = await FolioDocxReviewer.fromBuffer(stillPending);
    expect(rejecting.rejectAll()).toBe(1);
    expectDirectSpacing(
      await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer()),
      originalSpacing,
    );
  });

  test("does not copy a pending pPrChange into an inserted terminal paragraph", async () => {
    const originalSpacing = { spaceBefore: 120 } as const;
    const firstTarget = { spaceAfter: 240 } as const;
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await spacingDocument({ directSpacing: originalSpacing }),
    );
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    if (!blockId) {
      panic("expected a paragraph block id");
    }

    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "spacing-first",
            type: "setBlockParagraphProperties",
            blockId,
            properties: { spacing: firstTarget },
          },
        ],
      }).skipped,
    ).toEqual([]);
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "insert-spaced",
            type: "insertAfterBlock",
            blockId,
            text: INSERTED_TEXT,
            spacing: FULL_SPACING,
          },
        ],
      }).skipped,
    ).toEqual([]);

    const pending = await reviewer.toBuffer();
    const pendingParagraphs = paragraphXmls(await mainDocumentXml(pending));
    expect(pendingParagraphs).toHaveLength(2);
    for (const paragraph of pendingParagraphs) {
      expect(paragraph.match(/<w:pPrChange\b/gu)?.length ?? 0).toBeLessThanOrEqual(1);
    }
    expectDirectSpacing(await FolioDocxReviewer.fromBuffer(pending), firstTarget);
    expectDirectSpacing(await FolioDocxReviewer.fromBuffer(pending), FULL_SPACING, 1);

    const accepting = await FolioDocxReviewer.fromBuffer(pending);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
    expect(accepted.snapshot().blocks.map(({ text }) => text)).toEqual([TEXT, INSERTED_TEXT]);
    expectDirectSpacing(accepted, firstTarget);
    expectDirectSpacing(accepted, FULL_SPACING, 1);

    const rejecting = await FolioDocxReviewer.fromBuffer(pending);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
    expect(rejected.snapshot().blocks.map(({ text }) => text)).toEqual([TEXT]);
    expectDirectSpacing(rejected, originalSpacing);
  });
});
