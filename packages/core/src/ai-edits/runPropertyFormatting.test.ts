import { describe, expect, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { canonicalReviewModel } from "../../../../test/consumer-scenarios/support/review-comparison";
import { parseDocx } from "../docx/parser";
import { createDocx, repackDocx } from "../docx/rezip";
import {
  acceptAIEditRevision,
  rejectAIEditRevision,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { schema } from "../prosemirror/schema";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type {
  Document,
  Paragraph,
  RunContent,
  RunPropertyChange,
  TextFormatting,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { decodeOoxmlSymbolCharacter } from "../utils/ooxmlSymbol";
import { FolioDocxReviewer } from "./headless";
import { getTrackedChangesFromDoc } from "./read";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIBlockPreviewRun, FolioAIInlineFormatting } from "./types";

type CreateFormattingBaselineOptions = {
  formatting?: TextFormatting;
};

const createFormattingBaseline = async ({
  formatting,
}: CreateFormattingBaselineOptions = {}): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      content: [
        {
          type: "run",
          ...(formatting && { formatting }),
          content: [{ type: "text", text: "Formatting target" }],
        },
      ],
    },
  ];
  return createDocx(document);
};

const createSameValuedDirectFormattingDocument =
  async (): Promise<ArrayBuffer> => {
    const document = createEmptyDocument();
    const formatting: TextFormatting = {
      fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
      fontSize: 21,
      color: { rgb: "C00000" },
    };
    document.package.styles = {
      ...document.package.styles,
      docDefaults: { ...document.package.styles?.docDefaults, rPr: formatting },
      styles: document.package.styles?.styles.map((style) =>
        style.styleId === "Normal" ? { ...style, rPr: formatting } : style,
      ),
    };
    document.package.document.content = [
      {
        type: "paragraph",
        paraId: "A1000003",
        content: [
          { type: "run", content: [{ type: "text", text: "Inherited" }] },
          {
            type: "run",
            formatting,
            content: [{ type: "text", text: "Direct" }],
          },
        ],
      },
    ];
    return createDocx(document);
  };

const createParagraphFontWithDirectBoldDocument =
  async (): Promise<ArrayBuffer> => {
    const document = createEmptyDocument();
    document.package.document.content = [
      {
        type: "paragraph",
        paraId: "A1000004",
        formatting: {
          runProperties: {
            fontFamily: { ascii: "Arial", hAnsi: "Arial" },
            fontSize: 21,
            color: { rgb: "C00000" },
          },
        },
        content: [
          {
            type: "run",
            formatting: { bold: true },
            content: [{ type: "text", text: "Direct bold" }],
          },
        ],
      },
    ];
    return createDocx(document);
  };

type SnapshotFormattingProvenanceCase = {
  label: string;
  paragraphStyleId?: string;
  formatting?: TextFormatting;
  expected: Omit<FolioAIBlockPreviewRun, "text">;
};

const SNAPSHOT_FORMATTING_PROVENANCE_CASES = [
  {
    label: "unstyled inherited",
    expected: { fontFamily: "Arial", fontSizePt: 11 },
  },
  {
    label: "unstyled equal direct",
    formatting: {
      fontFamily: { ascii: "Arial", hAnsi: "Arial" },
      fontSize: 22,
    },
    expected: {
      fontFamily: "Arial",
      fontSizePt: 11,
      directFormatting: { fontFamily: "Arial", fontSizePt: 11 },
    },
  },
  {
    label: "unstyled direct",
    formatting: { bold: true },
    expected: {
      bold: true,
      fontFamily: "Arial",
      fontSizePt: 11,
      directFormatting: { bold: true },
    },
  },
  {
    label: "paragraph style inherited",
    paragraphStyleId: "Heading1",
    expected: { bold: true, fontFamily: "Arial", fontSizePt: 20 },
  },
  {
    label: "paragraph style equal direct",
    paragraphStyleId: "Heading1",
    formatting: { bold: true, fontSize: 40 },
    expected: {
      bold: true,
      fontFamily: "Arial",
      fontSizePt: 20,
      directFormatting: { bold: true, fontSizePt: 20 },
    },
  },
  {
    label: "paragraph style direct off",
    paragraphStyleId: "Heading1",
    formatting: { bold: false },
    expected: {
      fontFamily: "Arial",
      fontSizePt: 20,
      directFormatting: { bold: false },
    },
  },
  {
    label: "character style inherited",
    formatting: { styleId: "SnapshotCharacter" },
    expected: {
      italic: true,
      underline: true,
      strike: true,
      fontFamily: "Georgia",
      fontSizePt: 13,
      color: "#C00000",
    },
  },
  {
    label: "character style plus direct",
    formatting: { styleId: "SnapshotCharacter", bold: true },
    expected: {
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      fontFamily: "Georgia",
      fontSizePt: 13,
      color: "#C00000",
      directFormatting: { bold: true },
    },
  },
  {
    label: "character style equal direct",
    formatting: {
      styleId: "SnapshotCharacter",
      italic: true,
      fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
      fontSize: 26,
      color: { rgb: "C00000" },
    },
    expected: {
      italic: true,
      underline: true,
      strike: true,
      fontFamily: "Georgia",
      fontSizePt: 13,
      color: "#C00000",
      directFormatting: {
        italic: true,
        fontFamily: "Georgia",
        fontSizePt: 13,
        color: "#C00000",
      },
    },
  },
  {
    label: "character style direct off",
    formatting: { styleId: "SnapshotCharacter", italic: false },
    expected: {
      underline: true,
      strike: true,
      fontFamily: "Georgia",
      fontSizePt: 13,
      color: "#C00000",
      directFormatting: { italic: false },
    },
  },
  {
    label: "character style direct value resets",
    formatting: {
      styleId: "SnapshotCharacter",
      underline: { style: "none" },
      strike: false,
      color: { auto: true },
    },
    expected: {
      italic: true,
      fontFamily: "Georgia",
      fontSizePt: 13,
      directFormatting: { underline: false, strike: false },
    },
  },
  {
    label: "character style double strike survives single strike reset",
    formatting: {
      styleId: "SnapshotCharacter",
      strike: false,
      doubleStrike: true,
    },
    expected: {
      italic: true,
      underline: true,
      strike: true,
      fontFamily: "Georgia",
      fontSizePt: 13,
      color: "#C00000",
      directFormatting: { strike: false },
    },
  },
] as const satisfies readonly SnapshotFormattingProvenanceCase[];

const createSnapshotFormattingProvenanceDocument =
  async (): Promise<ArrayBuffer> => {
    const document = createEmptyDocument();
    document.package.styles = {
      ...document.package.styles,
      styles: [
        ...(document.package.styles?.styles ?? []),
        {
          styleId: "SnapshotCharacter",
          type: "character",
          name: "Snapshot Character",
          rPr: {
            italic: true,
            underline: { style: "single" },
            strike: true,
            fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
            fontSize: 26,
            color: { rgb: "C00000" },
          },
        },
      ],
    };
    document.package.document.content =
      SNAPSHOT_FORMATTING_PROVENANCE_CASES.map(
        ({ label, paragraphStyleId, formatting }, index): Paragraph => ({
          type: "paragraph",
          paraId: `B200${String(index).padStart(4, "0")}`,
          ...(paragraphStyleId !== undefined && {
            formatting: { styleId: paragraphStyleId },
          }),
          content: [
            {
              type: "run",
              ...(formatting !== undefined && { formatting }),
              content: [{ type: "text", text: label }],
            },
          ],
        }),
      );
    return createDocx(document);
  };

const expectSnapshotFormattingProvenance = (
  snapshot: ReturnType<FolioDocxReviewer["snapshot"]>,
): void => {
  expect(snapshot.blocks.map(({ text }) => text)).toEqual(
    SNAPSHOT_FORMATTING_PROVENANCE_CASES.map(({ label }) => label),
  );
  for (const [
    index,
    { label, expected },
  ] of SNAPSHOT_FORMATTING_PROVENANCE_CASES.entries()) {
    expect({
      label,
      previewRuns: snapshot.blocks.at(index)?.previewRuns,
    }).toEqual({
      label,
      previewRuns: [{ text: label, ...expected }],
    });
  }
};

type ApplyTrackedFormattingOptions = {
  formatting: FolioAIInlineFormatting;
  baselineFormatting?: TextFormatting;
};

const applyTrackedFormatting = async ({
  formatting,
  baselineFormatting,
}: ApplyTrackedFormattingOptions): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createFormattingBaseline({ formatting: baselineFormatting }),
    { author: "Reviewer" },
  );
  const block = reviewer.snapshot().blocks.at(0);
  const range = block
    ? createFolioAITextRangeHandle({
        blockId: block.id,
        text: block.text,
        startOffset: 0,
        endOffset: "Formatting".length,
      })
    : null;
  if (!range) {
    throw new Error("expected a formatting range");
  }

  const result = reviewer.applyOperations(
    [{ id: "format", type: "formatRange", range, formatting }],
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied.at(0)?.revisionId).toBeNumber();
  return reviewer.toBuffer();
};

const HEADER_RELATIONSHIP_ID = "rIdFormattingHeader";

const createHeaderFormattingBaseline = async (): Promise<ArrayBuffer> => {
  const seed = await createDocx(createEmptyDocument());
  const document = await parseDocx(seed, {
    detectVariables: false,
    preloadFonts: false,
  });
  document.package.headers = new Map([
    [
      HEADER_RELATIONSHIP_ID,
      {
        type: "header" as const,
        hdrFtrType: "default" as const,
        content: [
          {
            type: "paragraph" as const,
            paraId: "A1000002",
            formatting: { styleId: "Heading1" },
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: "Header target" }],
              },
            ],
          },
        ],
      },
    ],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
  };
  return repackDocx(document, { updateModifiedDate: false });
};

const packagePart = async (
  buffer: ArrayBuffer,
  path: string,
): Promise<string | null> => {
  const zip = await ZipArchive.read(buffer);
  return zip.entry(path)?.text() ?? null;
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const xml = await packagePart(buffer, "word/document.xml");
  if (xml === null) {
    throw new Error("missing word/document.xml");
  }
  return xml;
};

const applyTrackedBold = async (): Promise<ArrayBuffer> => {
  return applyTrackedFormatting({ formatting: { bold: true } });
};

const INLINE_CARRIER_CHANGE = {
  type: "runPropertyChange",
  info: {
    id: 93,
    author: "Reviewer",
    date: "2026-09-09T00:00:00.000Z",
  },
  previousFormatting: {},
  currentFormatting: { bold: true },
} as const satisfies RunPropertyChange;

const inlineCarrierChange = (id: number): RunPropertyChange => ({
  ...INLINE_CARRIER_CHANGE,
  info: { ...INLINE_CARRIER_CHANGE.info, id },
});

const changedInlineRun = (content: RunContent, revisionId: number) => ({
  type: "run" as const,
  formatting: { bold: true },
  propertyChanges: [inlineCarrierChange(revisionId)],
  content: [content],
});

const createMixedInlineCarrierFormattingDocument = (): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000093",
      content: [
        changedInlineRun({ type: "text", text: "A" }, 93),
        changedInlineRun({ type: "tab" }, 94),
        changedInlineRun({ type: "break", breakType: "textWrapping" }, 95),
        changedInlineRun(
          { type: "symbol", font: "Wingdings", char: "F06F" },
          96,
        ),
        {
          type: "simpleField",
          instruction: " PAGE ",
          fieldType: "PAGE",
          content: [changedInlineRun({ type: "text", text: "1" }, 97)],
        },
        {
          type: "simpleField",
          instruction: " REF carrier ",
          fieldType: "REF",
          content: [
            {
              type: "hyperlink",
              anchor: "carrier",
              children: [changedInlineRun({ type: "text", text: "field" }, 98)],
            },
          ],
        },
      ],
    },
  ];
  return document;
};

const createUnformattedMixedInlineCarrierDocument = (): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000095",
      content: [
        { type: "run", content: [{ type: "text", text: "A" }] },
        { type: "run", content: [{ type: "tab" }] },
        {
          type: "run",
          content: [{ type: "break", breakType: "textWrapping" }],
        },
        {
          type: "run",
          content: [{ type: "symbol", font: "Wingdings", char: "F06F" }],
        },
        {
          type: "simpleField",
          instruction: " PAGE ",
          fieldType: "PAGE",
          content: [{ type: "run", content: [{ type: "text", text: "1" }] }],
        },
        {
          type: "simpleField",
          instruction: " REF carrier ",
          fieldType: "REF",
          content: [
            {
              type: "hyperlink",
              anchor: "carrier",
              children: [
                { type: "run", content: [{ type: "text", text: "field" }] },
              ],
            },
          ],
        },
      ],
    },
  ];
  return document;
};

const createMixedInlineCarrierDocumentWithPendingTab = (): Document => {
  const document = createUnformattedMixedInlineCarrierDocument();
  const paragraph = document.package.document.content.at(0);
  const tabRun =
    paragraph?.type === "paragraph" ? paragraph.content.at(1) : undefined;
  if (tabRun?.type !== "run") {
    throw new Error("expected a tab run fixture");
  }
  tabRun.formatting = { bold: true };
  tabRun.propertyChanges = [inlineCarrierChange(899)];
  return document;
};

const createSameIdInlineCarrierState = (): EditorState => {
  const change = schema.mark("runPropertyChange", {
    changes: [INLINE_CARRIER_CHANGE],
  });
  const marks = [schema.mark("bold"), change];
  const hyperlink = schema.mark("hyperlink", {
    href: "#carrier",
    _docxHyperlinkIndex: 0,
  });
  const fieldAttrs = {
    fieldType: "PAGE",
    instruction: " PAGE ",
    displayText: "1",
    fieldKind: "simple",
    fldLock: false,
    dirty: false,
  };
  return EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", { paraId: "A1000094" }, [
        schema.text("A", marks),
        schema.node("tab").mark(marks),
        schema.node("hardBreak").mark(marks),
        schema.node("symbol", { font: "Wingdings", char: "F06F" }).mark(marks),
        schema.node("field", fieldAttrs).mark(marks),
        schema
          .node(
            "structuredField",
            {
              ...fieldAttrs,
              fieldType: "REF",
              instruction: " REF carrier ",
              displayText: "field",
            },
            [schema.text("field", [...marks, hyperlink])],
          )
          .mark(marks),
      ]),
    ]),
  });
};

const applyRevisionDecision = (
  state: EditorState,
  mode: "accept" | "reject",
): EditorState => {
  let next = state;
  const command =
    mode === "accept"
      ? acceptAIEditRevision(INLINE_CARRIER_CHANGE.info.id)
      : rejectAIEditRevision(INLINE_CARRIER_CHANGE.info.id);
  expect(
    command(state, (transaction: Transaction) => {
      next = state.apply(transaction);
    }),
  ).toBe(true);
  return next;
};

const countXmlElements = (xml: string, localName: string): number =>
  xml.match(new RegExp(`<w:${localName}(?:[\\s/>])`, "gu"))?.length ?? 0;

/** The content of each run in `xml`. */
const runContents = (xml: string): string[] =>
  [...xml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/gu)].map(
    ([, content]) => content ?? "",
  );

/** The carriers of the mixed inline-carrier fixtures, however many runs hold them. */
const MIXED_CARRIER_ELEMENTS = { t: 3, tab: 1, br: 1, sym: 1 } as const;

const expectMixedCarriers = (xml: string): void => {
  for (const [localName, count] of Object.entries(MIXED_CARRIER_ELEMENTS))
    expect({ localName, count: countXmlElements(xml, localName) }).toEqual({
      localName,
      count,
    });
};

const assertSingularRunPropertyChanges = (
  xml: string,
  expectedRunCount: number,
): void => {
  const changedRuns = [
    ...xml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/gu),
  ].filter(([, content]) => content?.includes("<w:rPrChange "));
  expect(changedRuns).toHaveLength(expectedRunCount);
  for (const [, content] of changedRuns) {
    expect(countXmlElements(content ?? "", "rPrChange")).toBe(1);
  }
};

const HEADLESS_REVIEWER_SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

/** The run holding the word "Formatting", with its properties. */
const formattingOwnerRun = (xml: string): string | undefined =>
  runContents(xml).find((run) => run.includes("<w:t>Formatting</w:t>"));

/** The properties the first run states, without its pending change. */
const firstRunProperties = (xml: string): string =>
  xml.match(
    /<w:r>(?:<w:rPr>(.*?)(?:<w:rPrChange\b.*?<\/w:rPrChange>)?<\/w:rPr>)?<w:t\b/su,
  )?.[1] ?? "";

/**
 * Bold, then italic, on one unformatted run in one tracked batch, applied in
 * input order. Legacy refuses the italic, which would give the run a second
 * pending change; the canonical session folds it into the run's one change,
 * which still starts from the original properties, so accepting gives both
 * and rejecting neither.
 */
const TWO_EDITS_OF_ONE_RUN = {
  legacy: {
    applied: ["bold"],
    skipped: [{ id: "italic", reason: "pendingRunPropertyChange" }],
    properties: "<w:b/>",
  },
  canonical: {
    applied: ["bold", "italic"],
    skipped: [],
    properties: "<w:b/><w:i/>",
  },
} as const satisfies Record<
  HeadlessReviewerSession,
  {
    applied: readonly string[];
    skipped: readonly { id: string; reason: string }[];
    properties: string;
  }
>;

describe("tracked run formatting", () => {
  test("public snapshots distinguish direct formatting from paragraph and character styles", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSnapshotFormattingProvenanceDocument(),
    );

    expectSnapshotFormattingProvenance(reviewer.snapshot());
  });

  test("public snapshot formatting provenance survives save and reopen", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSnapshotFormattingProvenanceDocument(),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
    );

    expectSnapshotFormattingProvenance(reopened.snapshot());
  });

  test("direct formatting authors off and clearing restores inherited formatting", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSnapshotFormattingProvenanceDocument(),
    );
    const inherited = reviewer
      .snapshot()
      .blocks.find(({ text }) => text === "paragraph style inherited");
    if (!inherited) {
      throw new Error("expected an inherited formatting block");
    }
    const range = createFolioAITextRangeHandle({
      blockId: inherited.id,
      text: inherited.text,
      startOffset: 0,
      endOffset: inherited.text.length,
    });
    if (!range) {
      throw new Error("expected an inherited formatting range");
    }

    const authoredOff = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "bold-off",
            type: "formatRange",
            range,
            formatting: { bold: false },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(authoredOff.skipped).toEqual([]);
    const offBuffer = await reviewer.toBuffer();
    const reopenedOff = await FolioDocxReviewer.fromBuffer(offBuffer);
    const offBlock = reopenedOff
      .snapshot()
      .blocks.find(({ text }) => text === "paragraph style inherited");
    expect(offBlock?.previewRuns?.at(0)).toMatchObject({
      directFormatting: { bold: false },
    });
    expect(offBlock?.previewRuns?.at(0)?.bold).toBeUndefined();

    if (!offBlock) {
      throw new Error("expected the reopened direct-off block");
    }
    const clearRange = createFolioAITextRangeHandle({
      blockId: offBlock.id,
      text: offBlock.text,
      startOffset: 0,
      endOffset: offBlock.text.length,
    });
    if (!clearRange) {
      throw new Error("expected a direct-off formatting range");
    }
    const cleared = reopenedOff.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "bold-inherit",
            type: "formatRange",
            range: clearRange,
            formatting: { bold: null },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(cleared.skipped).toEqual([]);
    const reopenedCleared = await FolioDocxReviewer.fromBuffer(
      await reopenedOff.toBuffer(),
    );
    const clearedRun = reopenedCleared
      .snapshot()
      .blocks.find(({ text }) => text === "paragraph style inherited")
      ?.previewRuns?.at(0);
    expect(clearedRun?.bold).toBe(true);
    expect(clearedRun?.directFormatting).toBeUndefined();
  });

  test("snapshot preserves same-valued direct font properties", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSameValuedDirectFormattingDocument(),
    );

    expect(reviewer.snapshot().blocks.at(0)?.previewRuns).toEqual([
      {
        text: "Inherited",
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
      {
        text: "Direct",
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
        directFormatting: {
          fontFamily: "Georgia",
          fontSizePt: 10.5,
          color: "#C00000",
        },
      },
    ]);
  });

  test("snapshot does not promote paragraph font properties beside direct bold", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createParagraphFontWithDirectBoldDocument(),
    );

    expect(
      reviewer.snapshot().blocks.at(0)?.previewRuns?.at(0)?.directFormatting,
    ).toEqual({
      bold: true,
    });
  });

  test("tracks same-valued inherited font properties as new direct formatting", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSameValuedDirectFormattingDocument(),
      { author: "Reviewer" },
    );
    const block = reviewer.snapshot().blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: "Inherited".length,
        })
      : null;
    if (!range) {
      throw new Error("expected an inherited formatting range");
    }

    const result = reviewer.applyOperations(
      [
        {
          id: "format-inherited",
          type: "formatRange",
          range,
          formatting: {
            fontFamily: "Georgia",
            fontSizePt: 10.5,
            color: "C00000",
          },
        },
      ],
      { undefinedReferences: "refuse" },
    );

    expect(result.skipped).toEqual([]);
    expect(
      reviewer.readReviewedStory({ view: "current-markup" })?.changes,
    ).toEqual([
      expect.objectContaining({ type: "formatting", text: "Inherited" }),
    ]);
    expect(
      reviewer
        .readReviewedStory({ view: "final" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0)?.directFormatting,
    ).toEqual({ fontFamily: "Georgia", fontSizePt: 10.5, color: "#C00000" });
  });

  test("clears only direct font properties in a mixed inherited range", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSameValuedDirectFormattingDocument(),
      { author: "Reviewer" },
    );
    const block = reviewer.snapshot().blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: block.text.length,
        })
      : null;
    if (!range) {
      throw new Error("expected a mixed formatting range");
    }

    const result = reviewer.applyOperations(
      [
        {
          id: "clear-direct",
          type: "formatRange",
          range,
          formatting: { fontFamily: null, fontSizePt: null, color: null },
        },
      ],
      { undefinedReferences: "refuse" },
    );

    expect(result.skipped).toEqual([]);
    expect(
      reviewer.readReviewedStory({ view: "current-markup" })?.changes,
    ).toEqual([
      expect.objectContaining({ type: "formatting", text: "Direct" }),
    ]);
    expect(
      reviewer.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)
        ?.previewRuns,
    ).toEqual([
      {
        text: "InheritedDirect",
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
    ]);
  });

  test("saves, reopens, accepts, and rejects a formatting revision", async () => {
    const tracked = await applyTrackedBold();
    const trackedXml = await documentXml(tracked);
    expect(trackedXml).toContain("<w:rPrChange ");
    expect(trackedXml).toContain("<w:b/>");

    const accepting = await FolioDocxReviewer.fromBuffer(tracked);
    const acceptedChange = accepting
      .getChanges()
      .find(({ type }) => type === "formatting");
    if (!acceptedChange) {
      throw new Error("expected a formatting change after reopen");
    }
    expect(accepting.acceptChange(acceptedChange)).toBe(true);
    const acceptedXml = await documentXml(await accepting.toBuffer());
    expect(acceptedXml).not.toContain("<w:rPrChange ");
    expect(acceptedXml).toContain("<w:b/>");

    const rejecting = await FolioDocxReviewer.fromBuffer(tracked);
    const rejectedChange = rejecting
      .getChanges()
      .find(({ type }) => type === "formatting");
    if (!rejectedChange) {
      throw new Error("expected a formatting change after reopen");
    }
    expect(rejecting.rejectChange(rejectedChange)).toBe(true);
    const rejectedXml = await documentXml(await rejecting.toBuffer());
    expect(rejectedXml).not.toContain("<w:rPrChange ");
    expect(rejectedXml).not.toContain("<w:b/>");
  });

  test.each(HEADLESS_REVIEWER_SESSIONS)(
    "%s: never serializes two unresolved formatting revisions on one run",
    async (session) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createFormattingBaseline(),
        { author: "Reviewer", session },
      );
      const block = reviewer.snapshot().blocks.at(0);
      const range = block
        ? createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset: 0,
            endOffset: "Formatting".length,
          })
        : null;
      if (!range) {
        throw new Error("expected a formatting range");
      }

      const result = reviewer.applyOperations(
        [
          {
            id: "bold",
            type: "formatRange",
            range,
            formatting: { bold: true },
          },
          {
            id: "italic",
            type: "formatRange",
            range,
            formatting: { italic: true },
          },
        ],
        { undefinedReferences: "refuse" },
      );
      const outcome = TWO_EDITS_OF_ONE_RUN[session];
      expect(result.applied.map(({ id }) => id)).toEqual([...outcome.applied]);
      expect(result.skipped).toEqual([...outcome.skipped]);

      const tracked = await reviewer.toBuffer();
      const trackedXml = await documentXml(tracked);
      expect(trackedXml.match(/<w:rPrChange /gu)).toHaveLength(1);
      // The one change starts from the run's original, unformatted properties.
      expect(trackedXml).toContain("<w:rPr/></w:rPrChange>");
      expect(firstRunProperties(trackedXml)).toBe(outcome.properties);
      expect(
        (await FolioDocxReviewer.fromBuffer(tracked, { session })).getChanges(),
      ).toEqual([
        expect.objectContaining({ type: "formatting", text: "Formatting" }),
      ]);

      const accepting = await FolioDocxReviewer.fromBuffer(tracked, {
        session,
      });
      accepting.acceptAll();
      const acceptedXml = await documentXml(await accepting.toBuffer());
      expect(acceptedXml).not.toContain("<w:rPrChange ");
      expect(firstRunProperties(acceptedXml)).toBe(outcome.properties);

      const rejecting = await FolioDocxReviewer.fromBuffer(tracked, {
        session,
      });
      rejecting.rejectAll();
      const rejectedXml = await documentXml(await rejecting.toBuffer());
      expect(rejectedXml).not.toContain("<w:rPrChange ");
      expect(rejectedXml).not.toContain("<w:rPr");
      expect(rejecting.snapshot().blocks.at(0)?.text).toBe("Formatting target");
    },
  );

  test.each(HEADLESS_REVIEWER_SESSIONS)(
    "%s: refuses to clear a background over a reopened formatting change",
    async (session) => {
      const baseline = await createFormattingBaseline({
        formatting: { highlight: "yellow" },
      });
      const first = await FolioDocxReviewer.fromBuffer(baseline, {
        author: "Reviewer",
        session,
      });
      const firstBlock = first.snapshot().blocks.at(0);
      const firstRange = firstBlock
        ? createFolioAITextRangeHandle({
            blockId: firstBlock.id,
            text: firstBlock.text,
            startOffset: 0,
            endOffset: "Formatting".length,
          })
        : null;
      if (!firstRange) {
        throw new Error("expected an initial formatting range");
      }
      const firstResult = first.applyDocumentOperations(
        {
          version: 1,
          mode: "tracked-changes",
          operations: [
            {
              id: "bold",
              type: "formatRange",
              range: firstRange,
              formatting: { bold: true },
            },
          ],
        },
        {
          undefinedReferences: "refuse",
          revisionStamp: { date: "2026-01-02T03:04:05.000Z", idSeed: 700 },
        },
      );
      expect(firstResult.nextRevisionId).toBe(701);

      const tracked = await first.toBuffer();
      const reopened = await FolioDocxReviewer.fromBuffer(tracked, {
        author: "Reviewer",
        session,
      });
      const beforeReplacement = await reopened.toBuffer();
      const ownerRun = formattingOwnerRun(await documentXml(beforeReplacement));
      expect(ownerRun).toContain('<w:rPrChange w:id="700" ');
      const snapshotBeforeReplacement = reopened.snapshot();
      const reopenedBlock = reopened.snapshot().blocks.at(0);
      if (!reopenedBlock) {
        throw new Error("expected a reopened formatting block");
      }
      const replacementResult = reopened.applyDocumentOperations(
        {
          version: 1,
          mode: "tracked-changes",
          operations: [
            {
              id: "replace",
              type: "replaceInBlock",
              blockId: reopenedBlock.id,
              find: "Formatting",
              // Keeps the highlighted word, so clearing its background would
              // give the formatted run a second property change.
              replace: "Formatting, changed",
              comment: { text: "Explain this replacement." },
            },
          ],
        },
        {
          undefinedReferences: "refuse",
          revisionStamp: { date: "2026-01-02T03:04:06.000Z", idSeed: 701 },
        },
      );

      expect(replacementResult.status).toBe("committed");
      const afterReplacement = await reopened.toBuffer();
      expect(replacementResult.applied).toEqual([]);
      expect(replacementResult.skipped).toEqual([
        { id: "replace", reason: "pendingRunPropertyChange" },
      ]);
      expect(replacementResult.nextRevisionId).toBe(701);
      expect(reopened.getComments()).toEqual([]);
      expect(reopened.snapshot()).toEqual(snapshotBeforeReplacement);
      expect(await documentXml(afterReplacement)).toBe(
        await documentXml(beforeReplacement),
      );
      expect(await packagePart(afterReplacement, "word/comments.xml")).toBe(
        await packagePart(beforeReplacement, "word/comments.xml"),
      );

      const followupBlock = reopened.snapshot().blocks.at(0);
      const followupRange = followupBlock
        ? createFolioAITextRangeHandle({
            blockId: followupBlock.id,
            text: followupBlock.text,
            startOffset: followupBlock.text.length - "target".length,
            endOffset: followupBlock.text.length,
          })
        : null;
      if (!followupRange) {
        throw new Error("expected a follow-up formatting range");
      }
      const followupResult = reopened.applyDocumentOperations(
        {
          version: 1,
          mode: "tracked-changes",
          operations: [
            {
              id: "italic",
              type: "formatRange",
              range: followupRange,
              formatting: { italic: true },
            },
          ],
        },
        {
          undefinedReferences: "refuse",
          revisionStamp: {
            date: "2026-01-02T03:04:07.000Z",
            idSeed: replacementResult.nextRevisionId,
          },
        },
      );
      expect(followupResult.applied.at(0)?.revisionIds).toEqual([
        replacementResult.nextRevisionId,
      ]);
      expect(followupResult.nextRevisionId).toBe(
        replacementResult.nextRevisionId + 1,
      );
      const followedUp = await reopened.toBuffer();
      expect(formattingOwnerRun(await documentXml(followedUp))).toBe(ownerRun);

      const accepting = await FolioDocxReviewer.fromBuffer(followedUp, {
        session,
      });
      accepting.acceptAll();
      const acceptedBuffer = await accepting.toBuffer();
      const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer, {
        session,
      });
      const acceptedXml = await documentXml(acceptedBuffer);
      expect(acceptedXml).toContain("<w:b/>");
      expect(acceptedXml).toContain("<w:i/>");
      expect(acceptedXml).not.toContain("<w:rPrChange ");
      expect(accepted.getChanges()).toEqual([]);
      expect(accepted.snapshot().blocks.at(0)?.text).toBe("Formatting target");

      const rejecting = await FolioDocxReviewer.fromBuffer(followedUp, {
        session,
      });
      rejecting.rejectAll();
      const rejectedBuffer = await rejecting.toBuffer();
      const rejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer, {
        session,
      });
      const rejectedXml = await documentXml(rejectedBuffer);
      expect(rejectedXml).not.toContain("<w:b/>");
      expect(rejectedXml).not.toContain("<w:i/>");
      expect(rejectedXml).not.toContain("<w:rPrChange ");
      expect(rejected.getChanges()).toEqual([]);
      expect(rejected.snapshot().blocks.at(0)?.text).toBe("Formatting target");
      expect(firstRunProperties(rejectedXml)).toBe(
        '<w:highlight w:val="yellow"/>',
      );
    },
  );

  test.each(HEADLESS_REVIEWER_SESSIONS)(
    "%s: enumerates and resolves formatting revisions across mixed inline carriers",
    async (session) => {
      const source = await createDocx(
        createMixedInlineCarrierFormattingDocument(),
      );
      const reviewer = await FolioDocxReviewer.fromBuffer(source, { session });
      const symbol = decodeOoxmlSymbolCharacter("F06F");
      if (!symbol) {
        throw new Error("expected a decodable symbol fixture");
      }
      expect(reviewer.getChanges()).toEqual(
        ["A", "\t", "\n", symbol, "1", "field"].map((text, index) =>
          expect.objectContaining({ id: 93 + index, type: "formatting", text }),
        ),
      );

      const saved = await reviewer.toBuffer();
      assertSingularRunPropertyChanges(await documentXml(saved), 6);
      expect(
        (await FolioDocxReviewer.fromBuffer(saved, { session })).getChanges(),
      ).toEqual(reviewer.getChanges());

      for (const mode of ["accept", "reject"] as const) {
        const resolving = await FolioDocxReviewer.fromBuffer(saved, {
          session,
        });
        for (const { id } of resolving.getChanges()) {
          expect(
            mode === "accept"
              ? resolving.acceptChange(id)
              : resolving.rejectChange(id),
          ).toBe(true);
        }
        const resolved = await resolving.toBuffer();
        expect(
          (
            await FolioDocxReviewer.fromBuffer(resolved, { session })
          ).getChanges(),
        ).toEqual([]);
        const xml = await documentXml(resolved);
        expect(xml).not.toContain("<w:rPrChange ");
        expectMixedCarriers(xml);
        // Resolving may join runs the decisions leave alike, so every run
        // states the outcome rather than a count of property sets.
        expect(
          runContents(xml).map((run) => countXmlElements(run, "b")),
        ).toEqual(runContents(xml).map(() => (mode === "accept" ? 1 : 0)));
      }
    },
  );

  test("treats mixed same-id carriers and a structured result as one review change", async () => {
    const symbol = decodeOoxmlSymbolCharacter("F06F");
    if (!symbol) {
      throw new Error("expected a decodable symbol fixture");
    }
    const state = createSameIdInlineCarrierState();
    expect(getTrackedChangesFromDoc(state.doc)).toEqual([
      expect.objectContaining({
        id: INLINE_CARRIER_CHANGE.info.id,
        type: "formatting",
        text: `A\t\n${symbol}1field`,
      }),
    ]);

    for (const mode of ["accept", "reject"] as const) {
      const resolvedState = applyRevisionDecision(
        createSameIdInlineCarrierState(),
        mode,
      );
      expect(getTrackedChangesFromDoc(resolvedState.doc)).toEqual([]);
      const resolved = await createDocx(
        fromProseDoc(resolvedState.doc, undefined, {
          stylesheetSource: { type: "package" },
        }),
      );
      const reopened = await FolioDocxReviewer.fromBuffer(resolved);
      expect(reopened.getChanges()).toEqual([]);
      expect(countXmlElements(await documentXml(resolved), "b")).toBe(
        mode === "accept" ? 6 : 0,
      );
    }
  });

  test.each(["direct", "tracked-changes"] as const)(
    "%s formatting uses the serializer carrier contract for every supported inline run",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createDocx(createUnformattedMixedInlineCarrierDocument()),
        { author: "Reviewer" },
      );
      const block = reviewer.snapshot().blocks.at(0);
      const range = block
        ? createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset: 0,
            endOffset: block.text.length,
          })
        : null;
      if (!range) {
        throw new Error("expected a mixed-carrier formatting range");
      }

      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            {
              id: "format-carriers",
              type: "formatRange",
              range,
              formatting: {
                bold: true,
                fontFamily: "Georgia",
                fontSizePt: 10.5,
                color: "C00000",
              },
            },
          ],
        },
        {
          undefinedReferences: "refuse",
          revisionStamp: { date: "2026-09-09T00:00:00.000Z", idSeed: 900 },
        },
      );
      expect(result.skipped).toEqual([]);
      if (mode === "tracked-changes") {
        expect(result.applied.at(0)?.revisionIds).toEqual([
          900, 901, 902, 903, 904, 905,
        ]);
        expect(result.nextRevisionId).toBe(906);
      }
      const formatted = await reviewer.toBuffer();
      const formattedXml = await documentXml(formatted);
      expect(countXmlElements(formattedXml, "b")).toBe(6);
      expect(countXmlElements(formattedXml, "rFonts")).toBe(6);
      expect(countXmlElements(formattedXml, "sz")).toBe(6);
      expect(countXmlElements(formattedXml, "color")).toBe(6);

      if (mode === "direct") {
        expect(formattedXml).not.toContain("<w:rPrChange ");
        expect(
          (await FolioDocxReviewer.fromBuffer(formatted)).getChanges(),
        ).toEqual([]);
        return;
      }

      assertSingularRunPropertyChanges(formattedXml, 6);
      const symbol = decodeOoxmlSymbolCharacter("F06F");
      if (!symbol) {
        throw new Error("expected a decodable symbol fixture");
      }
      expect(
        (await FolioDocxReviewer.fromBuffer(formatted)).getChanges(),
      ).toEqual(
        ["A", "\t", "\n", symbol, "1", "field"].map((text, index) =>
          expect.objectContaining({
            id: 900 + index,
            type: "formatting",
            text,
          }),
        ),
      );
      for (const decision of ["accept", "reject"] as const) {
        const resolving = await FolioDocxReviewer.fromBuffer(formatted);
        if (decision === "accept") {
          resolving.acceptAll();
        } else {
          resolving.rejectAll();
        }
        const resolved = await resolving.toBuffer();
        const resolvedXml = await documentXml(resolved);
        expect(resolvedXml).not.toContain("<w:rPrChange ");
        expect(countXmlElements(resolvedXml, "b")).toBe(
          decision === "accept" ? 6 : 0,
        );
        expect(countXmlElements(resolvedXml, "rFonts")).toBe(
          decision === "accept" ? 6 : 0,
        );
        expect(countXmlElements(resolvedXml, "sz")).toBe(
          decision === "accept" ? 6 : 0,
        );
        expect(countXmlElements(resolvedXml, "color")).toBe(
          decision === "accept" ? 6 : 0,
        );
        expect(
          (await FolioDocxReviewer.fromBuffer(resolved)).getChanges(),
        ).toEqual([]);
      }
    },
  );

  const formatCarriersOverPendingTab = async (
    session: HeadlessReviewerSession,
  ) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(createMixedInlineCarrierDocumentWithPendingTab()),
      { author: "Reviewer", session },
    );
    const before = await reviewer.toBuffer();
    const snapshot = reviewer.snapshot();
    const block = snapshot.blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: block.text.length,
        })
      : null;
    if (!range) {
      throw new Error("expected a mixed-carrier formatting range");
    }

    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "format-carriers",
            type: "formatRange",
            range,
            formatting: { italic: true },
          },
        ],
      },
      {
        undefinedReferences: "refuse",
        revisionStamp: { date: "2026-09-09T00:00:00.000Z", idSeed: 900 },
      },
    );
    return { reviewer, before, snapshot, result };
  };

  test("legacy: a pending owner on one inline carrier refuses the whole formatting transaction", async () => {
    const { reviewer, before, snapshot, result } =
      await formatCarriersOverPendingTab("legacy");

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([
      { id: "format-carriers", reason: "pendingRunPropertyChange" },
    ]);
    expect(result.nextRevisionId).toBe(900);
    expect(reviewer.snapshot()).toEqual(snapshot);
    expect(await documentXml(await reviewer.toBuffer())).toBe(
      await documentXml(before),
    );
    expect(reviewer.getChanges()).toEqual([
      expect.objectContaining({ id: 899, type: "formatting", text: "\t" }),
    ]);
  });

  test("canonical: a pending owner on one inline carrier folds into the formatting transaction", async () => {
    const { reviewer, result } =
      await formatCarriersOverPendingTab("canonical");
    expect(result.skipped).toEqual([]);
    // The tab's change keeps its id: the transaction's new ids are the
    // other five carriers', and it reports the tab's change it now shares.
    expect(result.applied.at(0)?.revisionIds).toEqual([
      900, 901, 902, 903, 904, 899,
    ]);
    expect(result.nextRevisionId).toBe(905);

    const tracked = await reviewer.toBuffer();
    const trackedXml = await documentXml(tracked);
    // The tab's pending bold folds into its one change, which keeps its id
    // and the properties the tab had before either edit.
    assertSingularRunPropertyChanges(trackedXml, 6);
    expect(trackedXml).toContain(
      '<w:r><w:rPr><w:b/><w:i/><w:rPrChange w:id="899" w:author="Reviewer" w:date="2026-09-09T00:00:00.000Z"><w:rPr/></w:rPrChange></w:rPr><w:tab/></w:r>',
    );
    const symbol = decodeOoxmlSymbolCharacter("F06F");
    if (!symbol) {
      throw new Error("expected a decodable symbol fixture");
    }
    expect((await FolioDocxReviewer.fromBuffer(tracked)).getChanges()).toEqual(
      (
        [
          ["A", 900],
          ["\t", 899],
          ["\n", 901],
          [symbol, 902],
          ["1", 903],
          ["field", 904],
        ] as const
      ).map(([text, id]) =>
        expect.objectContaining({ id, type: "formatting", text }),
      ),
    );

    for (const decision of ["accept", "reject"] as const) {
      const resolving = await FolioDocxReviewer.fromBuffer(tracked, {
        session: "canonical",
      });
      if (decision === "accept") resolving.acceptAll();
      else resolving.rejectAll();
      const resolved = await resolving.toBuffer();
      expect(
        (await FolioDocxReviewer.fromBuffer(resolved)).getChanges(),
      ).toEqual([]);
      const xml = await documentXml(resolved);
      expect(xml).not.toContain("<w:rPrChange ");
      expectMixedCarriers(xml);
      const runs = runContents(xml);
      if (decision === "reject") {
        expect(countXmlElements(xml, "i")).toBe(0);
        expect(countXmlElements(xml, "b")).toBe(0);
        continue;
      }
      expect(runs.every((run) => countXmlElements(run, "i") === 1)).toBe(true);
      expect(runs.filter((run) => countXmlElements(run, "b") === 1)).toEqual([
        "<w:rPr><w:b/><w:i/></w:rPr><w:tab/>",
      ]);
    }
  });

  test.each(HEADLESS_REVIEWER_SESSIONS)(
    "formats a header in %s",
    async (session) => {
      const story = {
        type: "header" as const,
        relationshipId: HEADER_RELATIONSHIP_ID,
      };
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createHeaderFormattingBaseline(),
        {
          author: "Reviewer",
          session,
        },
      );
      const snapshot = reviewer.snapshotStory(story);
      const block = snapshot?.blocks.at(0);
      const range = block
        ? createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset: 0,
            endOffset: "Header".length,
          })
        : null;
      if (!snapshot || !range) {
        throw new Error("expected a header formatting range");
      }
      expect(snapshot.blocks.at(0)?.previewRuns).toEqual([
        {
          text: "Header target",
          bold: true,
          fontFamily: "Arial",
          fontSizePt: 20,
        },
      ]);

      const result = reviewer.applyDocumentOperationsToStory({
        undefinedReferences: "refuse",
        story,
        batch: {
          version: 1,
          mode: "tracked-changes",
          operations: [
            {
              id: "format-header",
              type: "formatRange",
              range,
              formatting: { italic: true },
            },
          ],
        },
        snapshot,
      });
      expect(result.skipped).toEqual([]);
      expect(
        reviewer
          .readReviewedStory({ story, view: "final" })
          ?.snapshot.blocks.at(0)?.previewRuns,
      ).toEqual([
        {
          text: "Header",
          bold: true,
          italic: true,
          fontFamily: "Arial",
          fontSizePt: 20,
          directFormatting: { italic: true },
        },
        {
          text: " target",
          bold: true,
          fontFamily: "Arial",
          fontSizePt: 20,
        },
      ]);

      const revisionId = result.applied.at(0)?.revisionId;
      expect(revisionId).toBeNumber();
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session },
      );
      expect(
        reopened.readReviewedStory({ story, view: "current-markup" })?.changes,
      ).toEqual([
        expect.objectContaining({
          id: revisionId,
          type: "formatting",
          text: "Header",
        }),
      ]);
    },
  );

  test("tracks target font face, size, and color through save and reopen", async () => {
    const tracked = await applyTrackedFormatting({
      formatting: { fontFamily: "Georgia", fontSizePt: 10.5, color: "c00000" },
    });
    const trackedXml = await documentXml(tracked);
    expect(trackedXml).toContain("<w:rPrChange ");
    expect(trackedXml).toContain(
      '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"',
    );
    expect(trackedXml).toContain('<w:color w:val="C00000"');
    expect(trackedXml).toContain('<w:sz w:val="21"');

    const reopened = await FolioDocxReviewer.fromBuffer(tracked);
    expect(
      reopened
        .readReviewedStory({ view: "final" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0),
    ).toMatchObject({
      fontFamily: "Georgia",
      fontSizePt: 10.5,
      color: "#C00000",
      directFormatting: {
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
    });
    expect(
      reopened
        .readReviewedStory({ view: "original" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0),
    ).toMatchObject({ fontFamily: "Arial", fontSizePt: 11 });
    expect(
      reopened
        .readReviewedStory({ view: "original" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0)?.directFormatting,
    ).toBeUndefined();
  });

  test("tracks clearing direct font properties", async () => {
    const tracked = await applyTrackedFormatting({
      formatting: { fontFamily: null, fontSizePt: null, color: null },
      baselineFormatting: {
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 21,
        color: { rgb: "C00000" },
      },
    });
    const reopened = await FolioDocxReviewer.fromBuffer(tracked);
    expect(
      reopened
        .readReviewedStory({ view: "final" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0),
    ).toMatchObject({ fontFamily: "Arial", fontSizePt: 11 });
    expect(
      reopened
        .readReviewedStory({ view: "final" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0)?.directFormatting,
    ).toBeUndefined();
    expect(
      reopened
        .readReviewedStory({ view: "original" })
        ?.snapshot.blocks.at(0)
        ?.previewRuns?.at(0),
    ).toMatchObject({
      fontFamily: "Georgia",
      fontSizePt: 10.5,
      color: "#C00000",
      directFormatting: {
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
    });
  });
});

// Earlier formatting fixtures did not pair property-set presence with partial replacements.
test.each([undefined, {}, { italic: true }] satisfies (
  TextFormatting | undefined
)[])(
  "formatting snapshots preserve property-set presence through resolution: %j",
  async (formatting) => {
    for (const patch of [
      { bold: true },
      ...(formatting?.italic === true ? [{ italic: null }] : []),
    ]) {
      for (const replace of [false, true]) {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await createFormattingBaseline({ formatting }),
        );
        const block = reviewer.snapshot().blocks.at(0);
        if (!block) throw new Error("Expected formatting block");
        const range = createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: 10,
        });
        if (!range) throw new Error("Expected formatting range");
        const result = reviewer.applyDocumentOperations(
          {
            version: 1,
            mode: "tracked-changes",
            operations: [
              { id: "format", type: "formatRange", range, formatting: patch },
              ...(replace
                ? [
                    {
                      id: "replace",
                      type: "replaceRange" as const,
                      range,
                      replace: "updated",
                    },
                  ]
                : []),
            ],
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.skipped).toEqual([]);
        expect(result.applied).toHaveLength(replace ? 2 : 1);
        for (const phase of [
          "pending",
          ...(replace ? ["restored-deletion"] : []),
        ]) {
          if (phase === "restored-deletion") {
            const deletion = reviewer
              .getChanges()
              .find(({ type }) => type === "deletion");
            if (!deletion) throw new Error("Expected replacement deletion");
            expect(reviewer.rejectChange(deletion.id)).toBe(true);
          }
          const reopened = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
          );
          for (const decision of ["accept", "reject"] as const) {
            const live = reviewer.toResolvedDocument(decision);
            const saved = reopened.toResolvedDocument(decision);
            expect(
              canonicalReviewModel(saved.package.document.content),
            ).toEqual(canonicalReviewModel(live.package.document.content));
            if (decision === "reject") {
              const paragraph = saved.package.document.content.at(0);
              if (paragraph?.type !== "paragraph")
                throw new Error("Expected resolved paragraph");
              for (const run of paragraph.content) {
                if (run.type === "run")
                  expect(run.formatting).toEqual(formatting);
              }
            }
          }
        }
      }
    }
  },
);
