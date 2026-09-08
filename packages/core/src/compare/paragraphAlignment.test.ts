import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { EditorState, type Transaction } from "prosemirror-state";

import { applyFolioAIEditOperations } from "../ai-edits/apply";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { createDocx } from "../docx/rezip";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import {
  acceptAllChanges,
  acceptSuggestion,
  getSuggestions,
  rejectAllChanges,
  rejectAllSuggestions,
  rejectSuggestion,
} from "../prosemirror/commands/comments";
import { applyStyle, setAlignment } from "../prosemirror/commands/paragraph";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createDocumentStylesPlugin } from "../prosemirror/plugins/documentStyles";
import type { Paragraph, ParagraphAlignment, Table, TableCell } from "../types/document";
import { PARAGRAPH_ALIGNMENT_VALUES } from "../types/documentEnumValues";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-08T00:00:00.000Z" } as const;
const REVISION_STAMP = { date: OPTIONS.timestamp, idSeed: 1 } as const;
const TEXT = "The agreement remains effective for the stated term.";
const STYLE_ID = "AlignedBody";
const NEXT_STYLE_ID = "NextAlignedBody";
const CELL_INSERT_TEXT = "A newly negotiated indemnity survives termination.";
const REPLACEMENT_TEXT = "The replacement paragraph keeps its intended style and alignment.";
const ORDINARY_PARAGRAPH_STATE_SIZE_BUDGET = 1_823_095;
const STYLED_PARAGRAPH_STATE_SIZE_BUDGET = 1_610_095;
const STYLED_ALIGNMENT_PROVENANCE_SIZE_BUDGET = 29_000;

const SUGGESTION_REJECTION_ORDERS = [
  { label: "oldest to newest", ids: ["suggestion-left", "suggestion-right", "suggestion-both"] },
  { label: "newest to oldest", ids: ["suggestion-both", "suggestion-right", "suggestion-left"] },
  {
    label: "middle, oldest, newest",
    ids: ["suggestion-right", "suggestion-left", "suggestion-both"],
  },
] as const;

type AlignmentDocumentOptions = {
  directAlignment?: ParagraphAlignment;
  inheritedAlignment?: ParagraphAlignment;
};

type AlignmentState = {
  alignment: ParagraphAlignment | undefined;
  label: string;
};

const ALIGNMENT_STATES: readonly AlignmentState[] = [
  { alignment: undefined, label: "absent" },
  ...PARAGRAPH_ALIGNMENT_VALUES.map((alignment) => ({ alignment, label: alignment })),
];

const DIRECT_ALIGNMENT_MATRIX = ALIGNMENT_STATES.flatMap((before) =>
  ALIGNMENT_STATES.filter((after) => after.alignment !== before.alignment).map((after) => ({
    after: after.alignment,
    before: before.alignment,
    transition: `${before.label} to ${after.label}`,
  })),
);

type StyleTransitionCase = {
  label: string;
  beforeDirect: ParagraphAlignment | undefined;
  alignmentChange: { type: "preserve" } | { type: "set"; value: ParagraphAlignment | null };
  afterDirect: ParagraphAlignment | undefined;
  afterStyleId: string | null;
  afterInherited: ParagraphAlignment | undefined;
};

const STYLE_TRANSITION_CASES = [
  {
    label: "clears the old direct value",
    beforeDirect: "center",
    alignmentChange: { type: "set", value: null },
    afterDirect: undefined,
    afterStyleId: NEXT_STYLE_ID,
    afterInherited: "both",
  },
  {
    label: "sets a new direct value",
    beforeDirect: undefined,
    alignmentChange: { type: "set", value: "lowKashida" },
    afterDirect: "lowKashida",
    afterStyleId: NEXT_STYLE_ID,
    afterInherited: "both",
  },
  {
    label: "preserves an existing direct value",
    beforeDirect: "center",
    alignmentChange: { type: "preserve" },
    afterDirect: "center",
    afterStyleId: NEXT_STYLE_ID,
    afterInherited: "both",
  },
  {
    label: "updates an inherited-only effective value",
    beforeDirect: undefined,
    alignmentChange: { type: "preserve" },
    afterDirect: undefined,
    afterStyleId: NEXT_STYLE_ID,
    afterInherited: "both",
  },
  {
    label: "preserves a direct value while clearing the style",
    beforeDirect: "center",
    alignmentChange: { type: "preserve" },
    afterDirect: "center",
    afterStyleId: null,
    afterInherited: undefined,
  },
  {
    label: "clears inherited alignment with the style",
    beforeDirect: undefined,
    alignmentChange: { type: "preserve" },
    afterDirect: undefined,
    afterStyleId: null,
    afterInherited: undefined,
  },
] as const satisfies readonly StyleTransitionCase[];

const REPLACEMENT_STYLE_TRANSITION_CASES = [
  {
    label: "inherited alignment while switching styles",
    directAlignment: undefined,
    styleId: NEXT_STYLE_ID,
    inheritedAlignment: "both",
  },
  {
    label: "direct alignment while switching styles",
    directAlignment: "center",
    styleId: NEXT_STYLE_ID,
    inheritedAlignment: "both",
  },
  {
    label: "inherited alignment while clearing the style",
    directAlignment: undefined,
    styleId: null,
    inheritedAlignment: undefined,
  },
  {
    label: "direct alignment while clearing the style",
    directAlignment: "center",
    styleId: null,
    inheritedAlignment: undefined,
  },
] as const satisfies readonly {
  label: string;
  directAlignment: ParagraphAlignment | undefined;
  styleId: string | null;
  inheritedAlignment: ParagraphAlignment | undefined;
}[];

type TrackedReplacementStyleCase = {
  label: string;
  beforeDirect: ParagraphAlignment | undefined;
  afterStyleId: string | null;
  afterInherited: ParagraphAlignment | undefined;
  beforeText: string;
  afterText: string;
  revisionIds: readonly number[];
};

const TRACKED_REPLACEMENT_STYLE_CASES = [
  {
    label: "changes text and style while preserving direct alignment",
    beforeDirect: "center",
    afterStyleId: NEXT_STYLE_ID,
    afterInherited: "both",
    beforeText: TEXT,
    afterText: REPLACEMENT_TEXT,
    revisionIds: [1, 2, 4],
  },
  {
    label: "clears a style without changing text",
    beforeDirect: undefined,
    afterStyleId: null,
    afterInherited: undefined,
    beforeText: TEXT,
    afterText: TEXT,
    revisionIds: [1],
  },
  {
    label: "changes the style of a blank paragraph",
    beforeDirect: undefined,
    afterStyleId: NEXT_STYLE_ID,
    afterInherited: "both",
    beforeText: "",
    afterText: "",
    revisionIds: [1],
  },
] as const satisfies readonly TrackedReplacementStyleCase[];

const styleTransitionProperties = ({ alignmentChange, afterStyleId }: StyleTransitionCase) =>
  alignmentChange.type === "preserve"
    ? { styleId: afterStyleId }
    : { styleId: afterStyleId, alignment: alignmentChange.value };

const alignmentDocumentModel = ({
  directAlignment,
  inheritedAlignment,
}: AlignmentDocumentOptions) => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      {
        type: "paragraph",
        styleId: "Normal",
        name: "Normal",
        default: true,
      },
      {
        type: "paragraph",
        styleId: STYLE_ID,
        name: "Aligned Body",
        basedOn: "Normal",
        ...(inheritedAlignment === undefined ? {} : { pPr: { alignment: inheritedAlignment } }),
      },
      {
        type: "paragraph",
        styleId: NEXT_STYLE_ID,
        name: "Next Aligned Body",
        basedOn: "Normal",
        pPr: { alignment: "both" },
      },
    ],
  };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      textId: "12345678",
      formatting: {
        styleId: STYLE_ID,
        ...(directAlignment === undefined ? {} : { alignment: directAlignment }),
      },
      content: [{ type: "run", content: [{ type: "text", text: TEXT }] }],
    },
  ];
  return document;
};

const documentCache = new Map<string, Promise<ArrayBuffer>>();

const alignmentDocument = (options: AlignmentDocumentOptions): Promise<ArrayBuffer> => {
  const key = JSON.stringify(options);
  const cached = documentCache.get(key);
  if (cached) {
    return cached;
  }
  const created = createDocx(alignmentDocumentModel(options));
  documentCache.set(key, created);
  return created;
};

const paragraphModel = ({
  text,
  paraId,
  directAlignment,
  styleId = STYLE_ID,
}: {
  text: string;
  paraId: string;
  directAlignment?: ParagraphAlignment;
  styleId?: string | null;
}): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  formatting: {
    ...(styleId === null ? {} : { styleId }),
    ...(directAlignment === undefined ? {} : { alignment: directAlignment }),
  },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const bodyInsertionDocuments = async ({
  inheritedAlignment,
  anchorDirectAlignment,
  insertedDirectAlignment,
  insertedStyleId = STYLE_ID,
  position = "after",
}: {
  inheritedAlignment: ParagraphAlignment;
  anchorDirectAlignment?: ParagraphAlignment;
  insertedDirectAlignment?: ParagraphAlignment;
  insertedStyleId?: string | null;
  position?: "before" | "after";
}): Promise<{ base: ArrayBuffer; target: ArrayBuffer }> => {
  const base = alignmentDocumentModel({
    directAlignment: anchorDirectAlignment,
    inheritedAlignment,
  });
  const target = alignmentDocumentModel({
    directAlignment: anchorDirectAlignment,
    inheritedAlignment,
  });
  const inserted = paragraphModel({
    text: "The added paragraph keeps its authored alignment.",
    paraId: "23456789",
    directAlignment: insertedDirectAlignment,
    styleId: insertedStyleId,
  });
  if (position === "before") {
    target.package.document.content.unshift(inserted);
  } else {
    target.package.document.content.push(inserted);
  }
  return { base: await createDocx(base), target: await createDocx(target) };
};

const emptyFinalCarrierInsertionDocuments = async ({
  baseDirectAlignment,
  targetDirectAlignment,
}: {
  baseDirectAlignment?: ParagraphAlignment;
  targetDirectAlignment?: ParagraphAlignment;
}): Promise<{ base: ArrayBuffer; target: ArrayBuffer }> => {
  const inheritedAlignment = "right";
  const base = alignmentDocumentModel({
    directAlignment: baseDirectAlignment,
    inheritedAlignment,
  });
  const target = alignmentDocumentModel({
    directAlignment: targetDirectAlignment,
    inheritedAlignment,
  });
  const baseCarrier = base.package.document.content.at(0);
  const targetCarrier = target.package.document.content.at(0);
  if (baseCarrier?.type !== "paragraph" || targetCarrier?.type !== "paragraph") {
    panic("expected body paragraph carriers");
  }
  baseCarrier.content = [];
  targetCarrier.content = [];
  target.package.document.content.push(
    paragraphModel({
      text: "The appended paragraph follows an empty carrier.",
      paraId: "23456789",
      directAlignment: targetDirectAlignment,
    }),
  );
  return { base: await createDocx(base), target: await createDocx(target) };
};

const tableCellDocumentModel = (inheritedAlignment: ParagraphAlignment) => {
  const document = alignmentDocumentModel({ inheritedAlignment });
  const cell: TableCell = {
    type: "tableCell",
    content: [
      paragraphModel({
        text: "Existing cell paragraph.",
        paraId: "3456789A",
        directAlignment: "center",
      }),
      paragraphModel({
        text: "The final cell paragraph remains unchanged.",
        paraId: "3A456789",
        directAlignment: "center",
      }),
    ],
  };
  const table: Table = {
    type: "table",
    rows: [{ type: "tableRow", cells: [cell] }],
  };
  document.package.document.content = [
    table,
    { type: "paragraph", paraId: "56789ABC", textId: "56789ABC", content: [] },
  ];
  return document;
};

const mainDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    panic("expected word/document.xml");
  }
  return await part.async("string");
};

const firstParagraphXml = (xml: string): string => {
  const paragraph = paragraphXmls(xml).at(0);
  if (!paragraph) {
    panic("expected a body paragraph");
  }
  return paragraph;
};

const paragraphXmls = (xml: string): string[] =>
  xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gu) ?? [];

const directAlignmentsIn = (xml: string): string[] =>
  [...xml.matchAll(/<w:jc\b[^>]*\bw:val="([^"]+)"[^>]*\/>/gu)].flatMap((match) =>
    match[1] === undefined ? [] : [match[1]],
  );

const trackedParagraphPropertyPartsFromParagraph = (
  paragraph: string,
): { current: string; previous: string } => {
  const propertiesStart = paragraph.indexOf("<w:pPr>");
  const changeStart = paragraph.indexOf("<w:pPrChange ");
  if (propertiesStart < 0 || changeStart < 0) {
    panic("expected w:pPr containing w:pPrChange");
  }
  const previous = paragraph
    .match(/<w:pPrChange\b[^>]*><w:pPr>([\s\S]*?)<\/w:pPr><\/w:pPrChange>/u)
    ?.at(1);
  if (previous === undefined) {
    panic("expected the previous w:pPr payload");
  }
  return { current: paragraph.slice(propertiesStart + "<w:pPr>".length, changeStart), previous };
};

const trackedParagraphPropertyParts = (xml: string): { current: string; previous: string } =>
  trackedParagraphPropertyPartsFromParagraph(firstParagraphXml(xml));

const untrackedParagraphProperties = (xml: string): string => {
  const properties = firstParagraphXml(xml)
    .match(/<w:pPr>([\s\S]*?)<\/w:pPr>/u)
    ?.at(1);
  if (properties === undefined) {
    panic("expected w:pPr");
  }
  return properties;
};

const untrackedParagraphPropertiesOrEmpty = (xml: string): string =>
  firstParagraphXml(xml)
    .match(/<w:pPr>([\s\S]*?)<\/w:pPr>/u)
    ?.at(1) ?? "";

const expectedParagraphProperties = (
  alignment: ParagraphAlignment | undefined,
  styleId: string | null = STYLE_ID,
): string =>
  `${styleId === null ? "" : `<w:pStyle w:val="${styleId}"/>`}${alignment ? `<w:jc w:val="${alignment}"/>` : ""}`;

const expectDirectAlignmentXml = (xml: string, expected: ParagraphAlignment | undefined): void => {
  expect(directAlignmentsIn(xml)).toEqual(expected === undefined ? [] : [expected]);
};

const firstParagraph = (reviewer: FolioDocxReviewer): Paragraph => {
  const block = reviewer.toDocument().package.document.content.at(0);
  if (block?.type !== "paragraph") {
    panic("expected the first document block to be a paragraph");
  }
  return block;
};

const expectDirectAlignmentModel = (
  reviewer: FolioDocxReviewer,
  expected: ParagraphAlignment | undefined,
  styleId: string | null = STYLE_ID,
): void => {
  expect(reviewer.snapshot().blocks.at(0)).toEqual({
    id: "12345678",
    kind: "paragraph",
    text: TEXT,
    ...(styleId === null ? {} : { styleId }),
    ...(expected === undefined ? {} : { directAlignment: expected }),
  });
  expect(firstParagraph(reviewer).formatting ?? {}).toEqual({
    ...(styleId === null ? {} : { styleId }),
    ...(expected === undefined ? {} : { alignment: expected }),
  });
};

type CompareRoundTripOptions = {
  before: AlignmentDocumentOptions;
  after: AlignmentDocumentOptions;
};

const expectCompareRoundTrip = async ({
  before,
  after,
}: CompareRoundTripOptions): Promise<void> => {
  const result = await compareDocx(
    await alignmentDocument(before),
    await alignmentDocument(after),
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
      properties: { alignment: after.directAlignment ?? null },
    }),
  ]);

  const pendingXml = await mainDocumentXml(result.value.buffer);
  expect(pendingXml).not.toContain("<w:ins ");
  expect(pendingXml).not.toContain("<w:del ");
  expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
  const pendingParts = trackedParagraphPropertyParts(pendingXml);
  expect(pendingParts.current).toBe(expectedParagraphProperties(after.directAlignment));
  expect(pendingParts.previous).toBe(expectedParagraphProperties(before.directAlignment));
  expectDirectAlignmentXml(pendingParts.current, after.directAlignment);
  expectDirectAlignmentXml(pendingParts.previous, before.directAlignment);

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(accepting.acceptAll()).toBe(1);
  expect(accepting.getChanges()).toEqual([]);
  expectDirectAlignmentModel(accepting, after.directAlignment);
  const accepted = await accepting.toBuffer();
  const acceptedXml = firstParagraphXml(await mainDocumentXml(accepted));
  expect(acceptedXml).not.toContain("<w:pPrChange");
  expect(untrackedParagraphProperties(acceptedXml)).toBe(
    expectedParagraphProperties(after.directAlignment),
  );
  expectDirectAlignmentXml(acceptedXml, after.directAlignment);
  const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
  expect(reopenedAccepted.getChanges()).toEqual([]);
  expectDirectAlignmentModel(reopenedAccepted, after.directAlignment);

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(rejecting.rejectAll()).toBe(1);
  expect(rejecting.getChanges()).toEqual([]);
  expectDirectAlignmentModel(rejecting, before.directAlignment);
  const rejected = await rejecting.toBuffer();
  const rejectedXml = firstParagraphXml(await mainDocumentXml(rejected));
  expect(rejectedXml).not.toContain("<w:pPrChange");
  expect(untrackedParagraphProperties(rejectedXml)).toBe(
    expectedParagraphProperties(before.directAlignment),
  );
  expectDirectAlignmentXml(rejectedXml, before.directAlignment);
  const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
  expect(reopenedRejected.getChanges()).toEqual([]);
  expectDirectAlignmentModel(reopenedRejected, before.directAlignment);
};

const paragraphFormatting = (document: ReturnType<typeof fromProseDoc>) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    panic("expected a paragraph");
  }
  return paragraph.formatting;
};

const makeView = (document: ReturnType<typeof alignmentDocumentModel>) => {
  const view = {
    state: EditorState.create({
      doc: toProseDoc(document),
      plugins: [createDocumentStylesPlugin(document.package.styles)],
    }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  return view;
};

type ApplyParagraphPropertiesOptions = {
  view: ReturnType<typeof makeView>;
  id: string;
  properties: { styleId?: string | null; alignment?: ParagraphAlignment | null };
  mode: "tracked-changes" | "suggested";
  idSeed: number;
};

const applyParagraphProperties = ({
  view,
  id,
  properties,
  mode,
  idSeed,
}: ApplyParagraphPropertiesOptions) => {
  const snapshot = createFolioAIEditSnapshot(view.state.doc);
  const block = snapshot.blocks.at(0);
  if (!block) {
    panic("expected a paragraph property target");
  }
  return applyFolioAIEditOperations({
    view,
    snapshot,
    operations: [{ id, type: "setBlockParagraphProperties", blockId: block.id, properties }],
    mode,
    author: mode === "suggested" ? "assistant" : "reviewer",
    revisionStamp: { ...REVISION_STAMP, idSeed },
  });
};

const applyThreeAlignmentSuggestions = (
  view: ReturnType<typeof makeView>,
): ReturnType<typeof makeView> => {
  for (const [index, [id, alignment]] of (
    [
      ["suggestion-left", "left"],
      ["suggestion-right", "right"],
      ["suggestion-both", "both"],
    ] as const
  ).entries()) {
    const outcome = applyParagraphProperties({
      view,
      id,
      properties: { alignment },
      mode: "suggested",
      idSeed: index + 1,
    });
    expect(outcome.skipped).toEqual([]);
    expect(outcome.applied.at(0)?.suggestionId).toBe(id);
  }
  return view;
};

describe("paragraph alignment comparison", () => {
  test.each(DIRECT_ALIGNMENT_MATRIX)(
    "round-trips every direct alignment transition: $transition",
    async ({ before, after }) => {
      await expectCompareRoundTrip({
        before: { directAlignment: before },
        after: { directAlignment: after },
      });
    },
  );

  test.each(PARAGRAPH_ALIGNMENT_VALUES)(
    "distinguishes inherited %s alignment from an equal direct value",
    async (alignment) => {
      await expectCompareRoundTrip({
        before: { inheritedAlignment: alignment },
        after: { directAlignment: alignment, inheritedAlignment: alignment },
      });
      await expectCompareRoundTrip({
        before: { directAlignment: alignment, inheritedAlignment: alignment },
        after: { inheritedAlignment: alignment },
      });
    },
  );

  test("clears a direct override back to a different inherited alignment", async () => {
    await expectCompareRoundTrip({
      before: { directAlignment: "center", inheritedAlignment: "right" },
      after: { inheritedAlignment: "right" },
    });
  });

  test("does not report style-inherited alignment as direct formatting", async () => {
    const source = await alignmentDocument({ inheritedAlignment: "right" });
    const reviewer = await FolioDocxReviewer.fromBuffer(source);

    expectDirectAlignmentModel(reviewer, undefined);
    const xml = firstParagraphXml(await mainDocumentXml(await reviewer.toBuffer()));
    expectDirectAlignmentXml(xml, undefined);

    const result = await compareDocx(source, source, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes).toEqual([]);
    expect(result.value.verification).toEqual({ status: "verified" });
  });

  test.each(PARAGRAPH_ALIGNMENT_VALUES)(
    "carries inserted direct %s alignment through final-mark rotation",
    async (alignment) => {
      const { base, target } = await bodyInsertionDocuments({
        inheritedAlignment: alignment,
        insertedDirectAlignment: alignment,
      });
      const result = await compareDocx(base, target, OPTIONS);
      if (result.isErr()) {
        throw result.error;
      }

      expect(result.value.verification).toEqual({ status: "verified" });
      expect(result.value.changes).toEqual([
        expect.objectContaining({
          kind: "insert",
          after: "The added paragraph keeps its authored alignment.",
        }),
      ]);
      const pendingXml = await mainDocumentXml(result.value.buffer);
      const pendingParagraphs = paragraphXmls(pendingXml);
      expect(pendingParagraphs).toHaveLength(2);
      expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
      expect(pendingXml).not.toContain("w:hanging");
      expect(trackedParagraphPropertyPartsFromParagraph(pendingParagraphs.at(1) ?? "")).toEqual({
        current: expectedParagraphProperties(alignment),
        previous: expectedParagraphProperties(undefined),
      });

      const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(accepting.acceptAll()).toBeGreaterThan(0);
      expect(accepting.getChanges()).toEqual([]);
      expect(accepting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
        undefined,
        alignment,
      ]);
      expect(
        accepting.snapshot().blocks.map(({ kind, text, styleId, directAlignment }) => ({
          kind,
          text,
          styleId,
          directAlignment,
        })),
      ).toEqual([
        { kind: "paragraph", text: TEXT, styleId: STYLE_ID, directAlignment: undefined },
        {
          kind: "paragraph",
          text: "The added paragraph keeps its authored alignment.",
          styleId: STYLE_ID,
          directAlignment: alignment,
        },
      ]);
      const accepted = await accepting.toBuffer();
      const acceptedXml = await mainDocumentXml(accepted);
      expect(acceptedXml).not.toContain("<w:pPrChange");
      expect(untrackedParagraphProperties(paragraphXmls(acceptedXml).at(1) ?? "")).toBe(
        expectedParagraphProperties(alignment),
      );
      const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
      expect(
        reopenedAccepted.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual([undefined, alignment]);

      const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(rejecting.rejectAll()).toBeGreaterThan(0);
      expect(rejecting.getChanges()).toEqual([]);
      expect(rejecting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
        undefined,
      ]);
      expect(
        rejecting.snapshot().blocks.map(({ kind, text, styleId, directAlignment }) => ({
          kind,
          text,
          styleId,
          directAlignment,
        })),
      ).toEqual([{ kind: "paragraph", text: TEXT, styleId: STYLE_ID, directAlignment: undefined }]);
      const rejected = await rejecting.toBuffer();
      const rejectedXml = await mainDocumentXml(rejected);
      expect(rejectedXml).not.toContain("<w:pPrChange");
      expect(rejectedXml).not.toContain("w:hanging");
      expect(untrackedParagraphProperties(firstParagraphXml(rejectedXml))).toBe(
        expectedParagraphProperties(undefined),
      );
      const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
      expect(
        reopenedRejected.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual([undefined]);
    },
  );

  test.each([
    {
      label: "direct to absent",
      baseDirectAlignment: "center",
      targetDirectAlignment: undefined,
    },
    {
      label: "absent to direct",
      baseDirectAlignment: undefined,
      targetDirectAlignment: "both",
    },
  ] as const)(
    "rejects an append after an empty final carrier whose alignment changed: $label",
    async ({ baseDirectAlignment, targetDirectAlignment }) => {
      const { base, target } = await emptyFinalCarrierInsertionDocuments({
        baseDirectAlignment,
        targetDirectAlignment,
      });
      const result = await compareDocx(base, target, OPTIONS);
      if (result.isErr()) {
        throw result.error;
      }

      expect(result.value.verification).toEqual({ status: "verified" });
      expect(result.value.changes.map(({ kind }) => kind)).toEqual(["paragraph-format", "insert"]);
      const pendingXml = await mainDocumentXml(result.value.buffer);
      const pendingParagraphs = paragraphXmls(pendingXml);
      expect(pendingParagraphs).toHaveLength(2);
      expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(2);
      expect(trackedParagraphPropertyPartsFromParagraph(pendingParagraphs.at(0) ?? "")).toEqual({
        current:
          expectedParagraphProperties(targetDirectAlignment) +
          `<w:rPr><w:ins w:id="2" w:author="${OPTIONS.author}" w:date="${OPTIONS.timestamp}"/></w:rPr>`,
        previous: expectedParagraphProperties(baseDirectAlignment),
      });
      expect(trackedParagraphPropertyPartsFromParagraph(pendingParagraphs.at(1) ?? "")).toEqual({
        current: expectedParagraphProperties(targetDirectAlignment),
        previous: expectedParagraphProperties(baseDirectAlignment),
      });
      expect(pendingXml).not.toContain("w:hanging");

      const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(accepting.acceptAll()).toBeGreaterThan(0);
      expect(accepting.getChanges()).toEqual([]);
      expect(accepting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
        targetDirectAlignment,
        targetDirectAlignment,
      ]);
      const accepted = await accepting.toBuffer();
      const acceptedXml = await mainDocumentXml(accepted);
      expect(acceptedXml).not.toContain("<w:pPrChange");
      expect(paragraphXmls(acceptedXml).map(untrackedParagraphProperties)).toEqual([
        expectedParagraphProperties(targetDirectAlignment),
        expectedParagraphProperties(targetDirectAlignment),
      ]);
      const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
      expect(
        reopenedAccepted.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual([targetDirectAlignment, targetDirectAlignment]);

      const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(rejecting.rejectAll()).toBeGreaterThan(0);
      expect(rejecting.getChanges()).toEqual([]);
      expect(rejecting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
        baseDirectAlignment,
      ]);
      const rejected = await rejecting.toBuffer();
      const rejectedXml = await mainDocumentXml(rejected);
      expect(rejectedXml).not.toContain("<w:pPrChange");
      expect(rejectedXml).not.toContain("w:hanging");
      expect(untrackedParagraphProperties(rejectedXml)).toBe(
        expectedParagraphProperties(baseDirectAlignment),
      );
      const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
      expect(
        reopenedRejected.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual([baseDirectAlignment]);
    },
  );

  test("clears copied direct alignment on an inserted style-inherited paragraph", async () => {
    const { base, target } = await bodyInsertionDocuments({
      inheritedAlignment: "right",
      anchorDirectAlignment: "center",
    });
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    const pendingParagraphs = paragraphXmls(await mainDocumentXml(result.value.buffer));
    expect(trackedParagraphPropertyPartsFromParagraph(pendingParagraphs.at(1) ?? "")).toEqual({
      current: expectedParagraphProperties(undefined),
      previous: expectedParagraphProperties("center"),
    });

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    expect(accepting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
      "center",
      undefined,
    ]);
    const accepted = await accepting.toBuffer();
    expect(
      untrackedParagraphProperties(paragraphXmls(await mainDocumentXml(accepted)).at(1) ?? ""),
    ).toBe(expectedParagraphProperties(undefined));
    const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
    expect(
      reopenedAccepted.snapshot().blocks.map(({ directAlignment }) => directAlignment),
    ).toEqual(["center", undefined]);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    expect(rejecting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
      "center",
    ]);
    const rejected = await rejecting.toBuffer();
    expect(untrackedParagraphProperties(firstParagraphXml(await mainDocumentXml(rejected)))).toBe(
      expectedParagraphProperties("center"),
    );
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
    expect(
      reopenedRejected.snapshot().blocks.map(({ directAlignment }) => directAlignment),
    ).toEqual(["center"]);
  });

  test("carries direct alignment on an inserted body paragraph before its anchor", async () => {
    const { base, target } = await bodyInsertionDocuments({
      inheritedAlignment: "right",
      anchorDirectAlignment: "center",
      insertedDirectAlignment: "left",
      position: "before",
    });
    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes).toEqual([
      expect.objectContaining({
        kind: "insert",
        after: "The added paragraph keeps its authored alignment.",
      }),
    ]);
    const pendingXml = await mainDocumentXml(result.value.buffer);
    expect(pendingXml).not.toContain("<w:pPrChange");
    expectDirectAlignmentXml(paragraphXmls(pendingXml).at(0) ?? "", "left");

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    expect(accepting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
      "left",
      "center",
    ]);
    const accepted = await accepting.toBuffer();
    const acceptedXml = await mainDocumentXml(accepted);
    expect(acceptedXml).not.toContain("<w:pPrChange");
    expect(untrackedParagraphProperties(paragraphXmls(acceptedXml).at(0) ?? "")).toBe(
      expectedParagraphProperties("left"),
    );
    const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
    expect(
      reopenedAccepted.snapshot().blocks.map(({ directAlignment }) => directAlignment),
    ).toEqual(["left", "center"]);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    expect(rejecting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
      "center",
    ]);
    const rejected = await rejecting.toBuffer();
    const rejectedXml = await mainDocumentXml(rejected);
    expect(rejectedXml).not.toContain("<w:pPrChange");
    expect(untrackedParagraphProperties(firstParagraphXml(rejectedXml))).toBe(
      expectedParagraphProperties("center"),
    );
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
    expect(
      reopenedRejected.snapshot().blocks.map(({ directAlignment }) => directAlignment),
    ).toEqual(["center"]);
  });

  test.each([
    {
      label: "switches to a differently aligned style",
      insertedStyleId: NEXT_STYLE_ID,
      insertedDirectAlignment: "lowKashida",
    },
    {
      label: "clears the copied style and direct alignment",
      insertedStyleId: null,
      insertedDirectAlignment: undefined,
    },
  ] as const)(
    "an inserted body paragraph $label",
    async ({ insertedStyleId, insertedDirectAlignment }) => {
      const { base, target } = await bodyInsertionDocuments({
        inheritedAlignment: "right",
        anchorDirectAlignment: "center",
        insertedDirectAlignment,
        insertedStyleId,
      });
      const result = await compareDocx(base, target, OPTIONS);
      if (result.isErr()) {
        throw result.error;
      }

      expect(result.value.verification).toEqual({ status: "verified" });
      expect(result.value.changes).toEqual([
        expect.objectContaining({
          kind: "insert",
          after: "The added paragraph keeps its authored alignment.",
        }),
      ]);
      const pendingParagraph =
        paragraphXmls(await mainDocumentXml(result.value.buffer)).at(1) ?? "";
      expect(trackedParagraphPropertyPartsFromParagraph(pendingParagraph)).toEqual({
        current: expectedParagraphProperties(insertedDirectAlignment, insertedStyleId),
        previous: expectedParagraphProperties("center"),
      });

      const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(accepting.acceptAll()).toBeGreaterThan(0);
      expect(
        accepting.snapshot().blocks.map(({ text, styleId, directAlignment }) => ({
          text,
          styleId,
          directAlignment,
        })),
      ).toEqual([
        { text: TEXT, styleId: STYLE_ID, directAlignment: "center" },
        {
          text: "The added paragraph keeps its authored alignment.",
          styleId: insertedStyleId ?? undefined,
          directAlignment: insertedDirectAlignment,
        },
      ]);
      const accepted = await accepting.toBuffer();
      const acceptedParagraph = paragraphXmls(await mainDocumentXml(accepted)).at(1) ?? "";
      expect(acceptedParagraph).not.toContain("<w:pPrChange");
      expectDirectAlignmentXml(acceptedParagraph, insertedDirectAlignment);
      expect(acceptedParagraph.includes("<w:pStyle ")).toBe(insertedStyleId !== null);
      if (insertedStyleId !== null) {
        expect(untrackedParagraphProperties(acceptedParagraph)).toBe(
          expectedParagraphProperties(insertedDirectAlignment, insertedStyleId),
        );
      }
      const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
      expect(reopenedAccepted.snapshot().blocks.at(1)).toEqual(
        expect.objectContaining({
          text: "The added paragraph keeps its authored alignment.",
          ...(insertedStyleId === null ? {} : { styleId: insertedStyleId }),
          ...(insertedDirectAlignment === undefined
            ? {}
            : { directAlignment: insertedDirectAlignment }),
        }),
      );

      const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      expect(rejecting.rejectAll()).toBeGreaterThan(0);
      expect(
        rejecting.snapshot().blocks.map(({ styleId, directAlignment }) => ({
          styleId,
          directAlignment,
        })),
      ).toEqual([{ styleId: STYLE_ID, directAlignment: "center" }]);
      const rejected = await rejecting.toBuffer();
      const rejectedParagraph = firstParagraphXml(await mainDocumentXml(rejected));
      expect(rejectedParagraph).not.toContain("<w:pPrChange");
      expect(untrackedParagraphProperties(rejectedParagraph)).toBe(
        expectedParagraphProperties("center"),
      );
      const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
      expect(reopenedRejected.snapshot().blocks.at(0)).toEqual(
        expect.objectContaining({ styleId: STYLE_ID, directAlignment: "center" }),
      );
    },
  );

  test.each([
    {
      label: "before with absent direct alignment",
      position: "before",
      inheritedAlignment: "right",
      insertedDirectAlignment: undefined,
    },
    {
      label: "before with direct alignment different from the style",
      position: "before",
      inheritedAlignment: "right",
      insertedDirectAlignment: "left",
    },
    {
      label: "before with direct alignment equal to the style",
      position: "before",
      inheritedAlignment: "thaiDistribute",
      insertedDirectAlignment: "thaiDistribute",
    },
    {
      label: "after with absent direct alignment",
      position: "after",
      inheritedAlignment: "right",
      insertedDirectAlignment: undefined,
    },
    {
      label: "after with direct alignment different from the style",
      position: "after",
      inheritedAlignment: "right",
      insertedDirectAlignment: "left",
    },
    {
      label: "after with direct alignment equal to the style",
      position: "after",
      inheritedAlignment: "thaiDistribute",
      insertedDirectAlignment: "thaiDistribute",
    },
  ] as const)(
    "tracks a paragraph $label when its anchor is in a table cell",
    async ({ position, inheritedAlignment, insertedDirectAlignment }) => {
      const source = tableCellDocumentModel(inheritedAlignment);
      const type = position === "before" ? "insertBeforeBlock" : "insertAfterBlock";
      const operationFor = (blockId: string) =>
        ({
          id: "insert-aligned-cell-paragraph",
          type,
          blockId,
          text: CELL_INSERT_TEXT,
          styleId: STYLE_ID,
          alignment: insertedDirectAlignment ?? null,
        }) as const;
      const applyToView = () => {
        const view = makeView(source);
        const snapshot = createFolioAIEditSnapshot(view.state.doc);
        const anchor = snapshot.blocks.at(0);
        if (!anchor?.table) {
          panic("expected a table-cell insertion anchor");
        }
        const outcome = applyFolioAIEditOperations({
          view,
          snapshot,
          operations: [operationFor(anchor.id)],
          mode: "tracked-changes",
          author: OPTIONS.author,
          revisionStamp: REVISION_STAMP,
        });
        expect(outcome.skipped).toEqual([]);
        expect(outcome.applied).toHaveLength(1);
        return view;
      };
      const insertedParagraphOf = (view: ReturnType<typeof makeView>) => {
        const insertedIndex = position === "before" ? 0 : 1;
        const paragraph = view.state.doc.child(insertedIndex);
        if (paragraph.type.name !== "paragraph") {
          panic("expected the inserted paragraph beside the table");
        }
        return paragraph;
      };
      const firstCellAfterReject = (view: ReturnType<typeof makeView>) => {
        const cell = view.state.doc.firstChild?.firstChild?.firstChild;
        if (!cell) {
          panic("expected the restored table cell");
        }
        return cell;
      };
      const insertedIndex = position === "before" ? 0 : 2;
      const expectedEffectiveAlignment = insertedDirectAlignment ?? inheritedAlignment;
      const expectedOriginalFormatting = {
        styleId: STYLE_ID,
        ...(insertedDirectAlignment === undefined ? {} : { alignment: insertedDirectAlignment }),
      };

      const pendingView = applyToView();
      expect(expectParagraphAttrs(insertedParagraphOf(pendingView))).toMatchObject({
        styleId: STYLE_ID,
        alignment: expectedEffectiveAlignment,
        alignmentFromStyle: inheritedAlignment,
        _originalFormatting: expectedOriginalFormatting,
      });
      const acceptingView = applyToView();
      acceptAllChanges()(acceptingView.state, acceptingView.dispatch);
      expect(expectParagraphAttrs(insertedParagraphOf(acceptingView))).toMatchObject({
        styleId: STYLE_ID,
        alignment: expectedEffectiveAlignment,
        alignmentFromStyle: inheritedAlignment,
        _originalFormatting: expectedOriginalFormatting,
      });
      const rejectingView = applyToView();
      rejectAllChanges()(rejectingView.state, rejectingView.dispatch);
      expect(firstCellAfterReject(rejectingView).childCount).toBe(2);
      for (let index = 0; index < 2; index += 1) {
        expect(
          expectParagraphAttrs(firstCellAfterReject(rejectingView).child(index)),
        ).toMatchObject({
          styleId: STYLE_ID,
          alignment: "center",
          alignmentFromStyle: inheritedAlignment,
          _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
        });
      }

      const sourceBuffer = await createDocx(source);
      const pendingReviewer = await FolioDocxReviewer.fromBuffer(sourceBuffer, {
        author: OPTIONS.author,
      });
      const snapshot = pendingReviewer.snapshot();
      const anchor = snapshot.blocks.at(0);
      if (!anchor?.table) {
        panic("expected a table-cell insertion anchor");
      }
      const outcome = pendingReviewer.applyOperations([operationFor(anchor.id)], {
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toHaveLength(1);
      const pendingDirectAlignments =
        position === "before"
          ? [insertedDirectAlignment, "center", "center", undefined]
          : ["center", "center", insertedDirectAlignment, undefined];
      expect(
        pendingReviewer.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual(pendingDirectAlignments);
      const pendingBuffer = await pendingReviewer.toBuffer();
      const pendingXml = await mainDocumentXml(pendingBuffer);
      expect(pendingXml).not.toContain("<w:pPrChange");
      expectDirectAlignmentXml(
        paragraphXmls(pendingXml).at(insertedIndex) ?? "",
        insertedDirectAlignment,
      );

      const accepting = await FolioDocxReviewer.fromBuffer(pendingBuffer);
      expect(accepting.acceptAll()).toBeGreaterThan(0);
      expect(accepting.getChanges()).toEqual([]);
      expect(accepting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual(
        pendingDirectAlignments,
      );
      const accepted = await accepting.toBuffer();
      const acceptedXml = await mainDocumentXml(accepted);
      expect(acceptedXml).not.toContain("<w:pPrChange");
      expect(untrackedParagraphProperties(paragraphXmls(acceptedXml).at(insertedIndex) ?? "")).toBe(
        expectedParagraphProperties(insertedDirectAlignment),
      );
      const reopenedAccepted = await FolioDocxReviewer.fromBuffer(accepted);
      expect(
        reopenedAccepted.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual(pendingDirectAlignments);

      const rejecting = await FolioDocxReviewer.fromBuffer(pendingBuffer);
      expect(rejecting.rejectAll()).toBeGreaterThan(0);
      expect(rejecting.getChanges()).toEqual([]);
      expect(rejecting.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual([
        "center",
        "center",
        undefined,
      ]);
      const rejected = await rejecting.toBuffer();
      const rejectedXml = await mainDocumentXml(rejected);
      expect(rejectedXml).not.toContain("<w:pPrChange");
      expect(untrackedParagraphProperties(paragraphXmls(rejectedXml).at(0) ?? "")).toBe(
        expectedParagraphProperties("center"),
      );
      const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejected);
      expect(
        reopenedRejected.snapshot().blocks.map(({ directAlignment }) => directAlignment),
      ).toEqual(["center", "center", undefined]);
    },
  );
});

describe("paragraph alignment provenance in editor state", () => {
  test.each(["insertBeforeBlock", "insertAfterBlock"] as const)(
    "%s applies direct alignment immediately and after reopen",
    async (type) => {
      const source = alignmentDocumentModel({
        directAlignment: "center",
        inheritedAlignment: "right",
      });
      const view = makeView(source);
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const anchor = snapshot.blocks.at(0);
      if (!anchor) {
        panic("expected an insertion anchor");
      }
      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "insert-aligned",
            type,
            blockId: anchor.id,
            text: "Direct insertion.",
            styleId: NEXT_STYLE_ID,
            alignment: "right",
          },
        ],
        mode: "direct",
        author: OPTIONS.author,
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toHaveLength(1);

      const insertedIndex = type === "insertBeforeBlock" ? 0 : 1;
      const inserted = view.state.doc.child(insertedIndex);
      expect(expectParagraphAttrs(inserted)).toMatchObject({
        styleId: NEXT_STYLE_ID,
        alignment: "right",
        alignmentFromStyle: "both",
        _originalFormatting: { styleId: NEXT_STYLE_ID, alignment: "right" },
      });
      const edited = fromProseDoc(view.state.doc, source);
      const editedParagraph = edited.package.document.content.at(insertedIndex);
      if (editedParagraph?.type !== "paragraph") {
        panic("expected the inserted paragraph in the document model");
      }
      expect(editedParagraph.formatting).toEqual({ styleId: NEXT_STYLE_ID, alignment: "right" });

      const buffer = await createDocx(edited);
      const xml = await mainDocumentXml(buffer);
      expect(xml).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
      expect(untrackedParagraphProperties(paragraphXmls(xml).at(insertedIndex) ?? "")).toBe(
        expectedParagraphProperties("right", NEXT_STYLE_ID),
      );
      const reopened = await FolioDocxReviewer.fromBuffer(buffer);
      const expectedDirect =
        type === "insertBeforeBlock" ? ["right", "center"] : ["center", "right"];
      expect(reopened.snapshot().blocks.map(({ directAlignment }) => directAlignment)).toEqual(
        expectedDirect,
      );
      expect(reopened.snapshot().blocks.at(insertedIndex)?.styleId).toBe(NEXT_STYLE_ID);
    },
  );

  test.each(
    (["direct", "tracked-changes"] as const).flatMap((mode) =>
      (["insertBeforeBlock", "insertAfterBlock"] as const).map((type) => ({ mode, type })),
    ),
  )(
    "$mode $type keeps inferred direct alignment when the inserted style inherits the same value",
    async ({ mode, type }) => {
      const source = alignmentDocumentModel({ inheritedAlignment: "right" });
      const nextStyle = source.package.styles?.styles.find(
        ({ styleId }) => styleId === NEXT_STYLE_ID,
      );
      if (!nextStyle) {
        panic("expected the insertion target style");
      }
      nextStyle.pPr = { ...nextStyle.pPr, alignment: "center" };

      const imported = toProseDoc(source);
      const importedAnchor = imported.firstChild;
      if (!importedAnchor) {
        panic("expected an imported insertion anchor");
      }
      const createdAnchor = importedAnchor.type.create(
        {
          ...importedAnchor.attrs,
          alignment: "center",
          alignmentFromStyle: "right",
          _originalFormatting: null,
        },
        importedAnchor.content,
      );
      const view = {
        state: EditorState.create({
          doc: imported.type.create(imported.attrs, [createdAnchor]),
          plugins: [createDocumentStylesPlugin(source.package.styles)],
        }),
        dispatch(transaction: Transaction) {
          view.state = view.state.apply(transaction);
        },
      };
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const anchor = snapshot.blocks.at(0);
      if (!anchor) {
        panic("expected a PM-created insertion anchor");
      }
      expect(anchor.directAlignment).toBe("center");

      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "insert-equal-direct",
            type,
            blockId: anchor.id,
            text: "Equal direct insertion.",
            styleId: NEXT_STYLE_ID,
          },
        ],
        mode,
        author: OPTIONS.author,
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);

      const insertedIndex = type === "insertBeforeBlock" ? 0 : 1;
      const pendingInserted = view.state.doc.child(insertedIndex);
      expect(expectParagraphAttrs(pendingInserted)).toMatchObject({
        styleId: NEXT_STYLE_ID,
        alignment: "center",
        alignmentFromStyle: "center",
        _originalFormatting: { styleId: NEXT_STYLE_ID, alignment: "center" },
      });
      expect(createFolioAIEditSnapshot(view.state.doc).blocks.at(insertedIndex)).toMatchObject({
        styleId: NEXT_STYLE_ID,
        directAlignment: "center",
      });

      if (mode === "tracked-changes") {
        expect(acceptAllChanges()(view.state, view.dispatch)).toBe(true);
      }
      const saved = await createDocx(fromProseDoc(view.state.doc, source));
      const xml = await mainDocumentXml(saved);
      expect(xml).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
      expect(untrackedParagraphProperties(paragraphXmls(xml).at(insertedIndex) ?? "")).toBe(
        expectedParagraphProperties("center", NEXT_STYLE_ID),
      );
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expect(reopened.snapshot().blocks.at(insertedIndex)).toMatchObject({
        styleId: NEXT_STYLE_ID,
        directAlignment: "center",
      });
    },
  );

  test.each(STYLE_TRANSITION_CASES)("resolves a style switch that $label", async (transition) => {
    const { beforeDirect, afterDirect, afterStyleId, afterInherited } = transition;
    const afterFormatting = {
      ...(afterStyleId === null ? {} : { styleId: afterStyleId }),
      ...(afterDirect === undefined ? {} : { alignment: afterDirect }),
    };
    const expectedAfterOriginalFormatting =
      Object.keys(afterFormatting).length === 0 ? null : afterFormatting;
    const source = alignmentDocumentModel({
      directAlignment: beforeDirect,
      inheritedAlignment: "right",
    });

    const applyTrackedOperation = () => {
      const view = makeView(source);
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected an editable paragraph");
      }
      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "style-and-alignment",
            type: "setBlockParagraphProperties",
            blockId: block.id,
            properties: styleTransitionProperties(transition),
          },
        ],
        mode: "tracked-changes",
        author: OPTIONS.author,
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toHaveLength(1);
      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected a paragraph");
      }
      const pending = expectParagraphAttrs(paragraph);
      expect(pending.styleId ?? null).toBe(afterStyleId);
      expect(pending.alignment ?? null).toBe(afterDirect ?? afterInherited ?? null);
      expect(pending.alignmentFromStyle).toBe(afterInherited);
      expect(pending._originalFormatting ?? null).toEqual(expectedAfterOriginalFormatting);
      expect(pending._propertyChanges).toEqual([
        {
          type: "paragraphPropertyChange",
          info: { id: 1, author: OPTIONS.author, date: OPTIONS.timestamp },
          previousFormatting: {
            styleId: STYLE_ID,
            ...(beforeDirect === undefined ? {} : { alignment: beforeDirect }),
          },
        },
      ]);
      return view;
    };

    const accepting = applyTrackedOperation();
    acceptAllChanges()(accepting.state, accepting.dispatch);
    const acceptedParagraph = accepting.state.doc.firstChild;
    if (!acceptedParagraph) {
      panic("expected an accepted paragraph");
    }
    const acceptedAttrs = expectParagraphAttrs(acceptedParagraph);
    expect(acceptedAttrs.styleId ?? null).toBe(afterStyleId);
    expect(acceptedAttrs.alignment ?? null).toBe(afterDirect ?? afterInherited ?? null);
    expect(acceptedAttrs.alignmentFromStyle).toBe(afterInherited);
    expect(acceptedAttrs._originalFormatting ?? null).toEqual(expectedAfterOriginalFormatting);
    expect(acceptedAttrs._propertyChanges).toBeUndefined();
    expect(paragraphFormatting(fromProseDoc(accepting.state.doc, source)) ?? {}).toEqual(
      afterFormatting,
    );

    const rejecting = applyTrackedOperation();
    rejectAllChanges()(rejecting.state, rejecting.dispatch);
    const rejectedParagraph = rejecting.state.doc.firstChild;
    if (!rejectedParagraph) {
      panic("expected a rejected paragraph");
    }
    const rejectedAttrs = expectParagraphAttrs(rejectedParagraph);
    expect(rejectedAttrs.styleId).toBe(STYLE_ID);
    expect(rejectedAttrs.alignment).toBe(beforeDirect ?? "right");
    expect(rejectedAttrs.alignmentFromStyle).toBe("right");
    expect(rejectedAttrs._originalFormatting).toEqual({
      styleId: STYLE_ID,
      ...(beforeDirect === undefined ? {} : { alignment: beforeDirect }),
    });
    expect(rejectedAttrs._propertyChanges).toBeUndefined();
    expect(paragraphFormatting(fromProseDoc(rejecting.state.doc, source))).toEqual({
      styleId: STYLE_ID,
      ...(beforeDirect === undefined ? {} : { alignment: beforeDirect }),
    });

    const buffer = await alignmentDocument({
      directAlignment: beforeDirect,
      inheritedAlignment: "right",
    });
    const pendingReviewer = await FolioDocxReviewer.fromBuffer(buffer, {
      author: OPTIONS.author,
    });
    const pendingBlock = pendingReviewer.snapshot().blocks.at(0);
    if (!pendingBlock) {
      panic("expected a reviewable paragraph");
    }
    const outcome = pendingReviewer.applyOperations(
      [
        {
          id: "style-and-alignment",
          type: "setBlockParagraphProperties",
          blockId: pendingBlock.id,
          properties: styleTransitionProperties(transition),
        },
      ],
      { revisionStamp: REVISION_STAMP },
    );
    expect(outcome.skipped).toEqual([]);
    expect(outcome.applied).toHaveLength(1);
    const pendingBuffer = await pendingReviewer.toBuffer();
    const pendingXml = await mainDocumentXml(pendingBuffer);
    expect(pendingXml).not.toContain("<w:ins ");
    expect(pendingXml).not.toContain("<w:del ");
    expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
    expect(trackedParagraphPropertyParts(pendingXml)).toEqual({
      current: expectedParagraphProperties(afterDirect, afterStyleId),
      previous: expectedParagraphProperties(beforeDirect, STYLE_ID),
    });

    const acceptingReviewer = await FolioDocxReviewer.fromBuffer(pendingBuffer);
    expect(acceptingReviewer.acceptAll()).toBe(1);
    expect(acceptingReviewer.getChanges()).toEqual([]);
    expectDirectAlignmentModel(acceptingReviewer, afterDirect, afterStyleId);
    const acceptedBuffer = await acceptingReviewer.toBuffer();
    const acceptedXml = firstParagraphXml(await mainDocumentXml(acceptedBuffer));
    expect(acceptedXml).not.toContain("<w:pPrChange");
    expect(untrackedParagraphPropertiesOrEmpty(acceptedXml)).toBe(
      expectedParagraphProperties(afterDirect, afterStyleId),
    );
    const reopenedAccepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
    expectDirectAlignmentModel(reopenedAccepted, afterDirect, afterStyleId);

    const rejectingReviewer = await FolioDocxReviewer.fromBuffer(pendingBuffer);
    expect(rejectingReviewer.rejectAll()).toBe(1);
    expect(rejectingReviewer.getChanges()).toEqual([]);
    expectDirectAlignmentModel(rejectingReviewer, beforeDirect, STYLE_ID);
    const rejectedBuffer = await rejectingReviewer.toBuffer();
    const rejectedXml = firstParagraphXml(await mainDocumentXml(rejectedBuffer));
    expect(rejectedXml).not.toContain("<w:pPrChange");
    expect(untrackedParagraphProperties(rejectedXml)).toBe(
      expectedParagraphProperties(beforeDirect, STYLE_ID),
    );
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    expectDirectAlignmentModel(reopenedRejected, beforeDirect, STYLE_ID);
  });

  test.each(REPLACEMENT_STYLE_TRANSITION_CASES)(
    "replaceBlock preserves $label",
    async ({ directAlignment, styleId, inheritedAlignment }) => {
      const source = alignmentDocumentModel({
        directAlignment,
        inheritedAlignment: "right",
      });
      const view = makeView(source);
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected an editable paragraph");
      }

      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "replace-and-restyle",
            type: "replaceBlock",
            blockId: block.id,
            text: REPLACEMENT_TEXT,
            styleId,
          },
        ],
        mode: "direct",
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toHaveLength(1);

      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected the replaced paragraph");
      }
      const attrs = expectParagraphAttrs(paragraph);
      expect(attrs.styleId ?? null).toBe(styleId);
      expect(attrs.alignment ?? null).toBe(directAlignment ?? inheritedAlignment ?? null);
      expect(attrs.alignmentFromStyle).toBe(inheritedAlignment);
      const expectedOriginalFormatting = {
        ...(styleId === null ? {} : { styleId }),
        ...(directAlignment === undefined ? {} : { alignment: directAlignment }),
      };
      if (Object.keys(expectedOriginalFormatting).length === 0) {
        expect(attrs._originalFormatting ?? null).toBeNull();
      } else {
        expect(attrs._originalFormatting).toEqual(expectedOriginalFormatting);
      }
      expect(createFolioAIEditSnapshot(view.state.doc).blocks.at(0)).toEqual({
        id: "12345678",
        kind: "paragraph",
        text: REPLACEMENT_TEXT,
        ...(styleId === null ? {} : { styleId }),
        ...(directAlignment === undefined ? {} : { directAlignment }),
      });

      const edited = fromProseDoc(view.state.doc, source);
      expect(paragraphFormatting(edited) ?? {}).toEqual({
        ...(styleId === null ? {} : { styleId }),
        ...(directAlignment === undefined ? {} : { alignment: directAlignment }),
      });
      const buffer = await createDocx(edited);
      expect(untrackedParagraphPropertiesOrEmpty(await mainDocumentXml(buffer))).toBe(
        expectedParagraphProperties(directAlignment, styleId),
      );

      const reopened = await FolioDocxReviewer.fromBuffer(buffer);
      expect(reopened.snapshot().blocks.at(0)).toEqual({
        id: "12345678",
        kind: "paragraph",
        text: REPLACEMENT_TEXT,
        ...(styleId === null ? {} : { styleId }),
        ...(directAlignment === undefined ? {} : { directAlignment }),
      });
      const reopenedParagraph = toProseDoc(reopened.toDocument()).firstChild;
      if (!reopenedParagraph) {
        panic("expected the reopened replacement");
      }
      const reopenedAttrs = expectParagraphAttrs(reopenedParagraph);
      expect(reopenedAttrs.styleId ?? null).toBe(styleId);
      expect(reopenedAttrs.alignment ?? null).toBe(directAlignment ?? inheritedAlignment ?? null);
      expect(reopenedAttrs.alignmentFromStyle).toBe(inheritedAlignment);
      if (Object.keys(expectedOriginalFormatting).length === 0) {
        expect(reopenedAttrs._originalFormatting ?? null).toBeNull();
      } else {
        expect(reopenedAttrs._originalFormatting).toEqual(expectedOriginalFormatting);
      }
    },
  );

  test("replaceBlock applies style provenance after dropping preserved formatting", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const view = makeView(source);
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    if (!block) {
      panic("expected a replacement block");
    }
    const outcome = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "drop-formatting-and-restyle",
          type: "replaceBlock",
          blockId: block.id,
          text: REPLACEMENT_TEXT,
          preserveFormatting: false,
          styleId: NEXT_STYLE_ID,
        },
      ],
      mode: "direct",
    });
    expect(outcome.skipped).toEqual([]);
    expect(outcome.applied).toEqual([{ id: "drop-formatting-and-restyle" }]);
    const paragraph = view.state.doc.firstChild;
    if (!paragraph) {
      panic("expected a replacement paragraph");
    }
    expect(expectParagraphAttrs(paragraph)).toMatchObject({
      styleId: NEXT_STYLE_ID,
      alignment: "both",
      alignmentFromStyle: "both",
      _originalFormatting: { styleId: NEXT_STYLE_ID },
    });
    const buffer = await createDocx(fromProseDoc(view.state.doc, source));
    const xml = firstParagraphXml(await mainDocumentXml(buffer));
    expect(untrackedParagraphProperties(xml)).toBe(
      expectedParagraphProperties(undefined, NEXT_STYLE_ID),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(buffer);
    expect(reopened.snapshot().blocks.at(0)).toEqual(
      expect.objectContaining({
        kind: "paragraph",
        text: REPLACEMENT_TEXT,
        styleId: NEXT_STYLE_ID,
      }),
    );
  });

  test("replaceBlock applies style provenance after rebuilding inline emphasis", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const view = makeView(source);
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    if (!block) {
      panic("expected a replacement block");
    }
    const outcome = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "emphasize-and-restyle",
          type: "replaceBlock",
          blockId: block.id,
          text: "The **replacement** remains effective.",
          styleId: NEXT_STYLE_ID,
        },
      ],
      mode: "direct",
    });
    expect(outcome.skipped).toEqual([]);
    expect(outcome.applied).toEqual([{ id: "emphasize-and-restyle" }]);
    const paragraph = view.state.doc.firstChild;
    if (!paragraph) {
      panic("expected an emphasized replacement paragraph");
    }
    expect(expectParagraphAttrs(paragraph)).toMatchObject({
      styleId: NEXT_STYLE_ID,
      alignment: "center",
      alignmentFromStyle: "both",
      _originalFormatting: { styleId: NEXT_STYLE_ID, alignment: "center" },
    });
    const buffer = await createDocx(fromProseDoc(view.state.doc, source));
    const xml = firstParagraphXml(await mainDocumentXml(buffer));
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("The ");
    expect(xml).toContain("replacement");
    expect(xml).not.toContain("**");
    expect(untrackedParagraphProperties(xml)).toBe(
      expectedParagraphProperties("center", NEXT_STYLE_ID),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(buffer);
    expect(reopened.snapshot().blocks.at(0)).toEqual(
      expect.objectContaining({
        id: "12345678",
        kind: "paragraph",
        text: "The replacement remains effective.",
        styleId: NEXT_STYLE_ID,
        directAlignment: "center",
      }),
    );
  });

  test.each(TRACKED_REPLACEMENT_STYLE_CASES)(
    "tracked replaceBlock $label",
    async ({ beforeDirect, afterStyleId, afterInherited, beforeText, afterText, revisionIds }) => {
      const source = alignmentDocumentModel({
        directAlignment: beforeDirect,
        inheritedAlignment: "right",
      });
      const sourceParagraph = source.package.document.content.at(0);
      if (sourceParagraph?.type !== "paragraph") {
        panic("expected a source paragraph");
      }
      sourceParagraph.content =
        beforeText.length === 0
          ? []
          : [{ type: "run", content: [{ type: "text", text: beforeText }] }];

      const applyTrackedReplacement = () => {
        const view = makeView(source);
        const block = createFolioAIEditSnapshot(view.state.doc).blocks.at(0);
        if (!block) {
          panic("expected a replacement block");
        }
        const outcome = applyFolioAIEditOperations({
          view,
          snapshot: createFolioAIEditSnapshot(view.state.doc),
          operations: [
            {
              id: "tracked-replacement-style",
              type: "replaceBlock",
              blockId: block.id,
              text: afterText,
              styleId: afterStyleId,
            },
          ],
          mode: "tracked-changes",
          author: OPTIONS.author,
          revisionStamp: REVISION_STAMP,
        });
        expect(outcome.skipped).toEqual([]);
        expect(outcome.applied).toEqual([
          {
            id: "tracked-replacement-style",
            revisionId: revisionIds.at(0),
            revisionIds,
          },
        ]);
        const paragraph = view.state.doc.firstChild;
        if (!paragraph) {
          panic("expected the pending replacement paragraph");
        }
        const expectedDirect = beforeDirect;
        const paragraphAttrs = expectParagraphAttrs(paragraph);
        expect(paragraphAttrs).toMatchObject({
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: revisionIds.at(-1), author: OPTIONS.author, date: OPTIONS.timestamp },
              previousFormatting: {
                styleId: STYLE_ID,
                ...(beforeDirect === undefined ? {} : { alignment: beforeDirect }),
              },
            },
          ],
        });
        expect(paragraphAttrs.styleId ?? null).toBe(afterStyleId);
        expect(paragraphAttrs.alignment ?? null).toBe(expectedDirect ?? afterInherited ?? null);
        expect(paragraphAttrs.alignmentFromStyle).toBe(afterInherited);
        const expectedOriginalFormatting = {
          ...(afterStyleId === null ? {} : { styleId: afterStyleId }),
          ...(expectedDirect === undefined ? {} : { alignment: expectedDirect }),
        };
        expect(paragraphAttrs._originalFormatting ?? null).toEqual(
          Object.keys(expectedOriginalFormatting).length > 0 ? expectedOriginalFormatting : null,
        );
        return view;
      };

      const pending = applyTrackedReplacement();
      const pendingBuffer = await createDocx(fromProseDoc(pending.state.doc, source));
      const pendingXml = await mainDocumentXml(pendingBuffer);
      expect(pendingXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
      expect(pendingXml.includes("<w:ins ")).toBe(beforeText !== afterText);
      expect(pendingXml.includes("<w:del ")).toBe(beforeText !== afterText);
      expect(trackedParagraphPropertyParts(pendingXml)).toEqual({
        current: expectedParagraphProperties(beforeDirect, afterStyleId),
        previous: expectedParagraphProperties(beforeDirect, STYLE_ID),
      });

      const accepting = await FolioDocxReviewer.fromBuffer(pendingBuffer);
      expect(accepting.acceptAll()).toBe(revisionIds.length);
      expect(accepting.getChanges()).toEqual([]);
      expect(accepting.snapshot().blocks.at(0)).toEqual({
        id: "12345678",
        kind: "paragraph",
        text: afterText,
        ...(afterStyleId === null ? {} : { styleId: afterStyleId }),
        ...(beforeDirect === undefined ? {} : { directAlignment: beforeDirect }),
      });
      const acceptedBuffer = await accepting.toBuffer();
      const acceptedXml = firstParagraphXml(await mainDocumentXml(acceptedBuffer));
      expect(acceptedXml).not.toContain("<w:pPrChange");
      expect(acceptedXml).not.toMatch(/<w:(?:ins|del)\b/u);
      expect(untrackedParagraphPropertiesOrEmpty(acceptedXml)).toBe(
        expectedParagraphProperties(beforeDirect, afterStyleId),
      );
      const reopenedAccepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
      expect(reopenedAccepted.snapshot().blocks.at(0)).toEqual(accepting.snapshot().blocks.at(0));

      const rejecting = await FolioDocxReviewer.fromBuffer(pendingBuffer);
      expect(rejecting.rejectAll()).toBe(revisionIds.length);
      expect(rejecting.getChanges()).toEqual([]);
      expect(rejecting.snapshot().blocks.at(0)).toEqual({
        id: "12345678",
        kind: "paragraph",
        text: beforeText,
        styleId: STYLE_ID,
        ...(beforeDirect === undefined ? {} : { directAlignment: beforeDirect }),
      });
      const rejectedBuffer = await rejecting.toBuffer();
      const rejectedXml = firstParagraphXml(await mainDocumentXml(rejectedBuffer));
      expect(rejectedXml).not.toContain("<w:pPrChange");
      expect(rejectedXml).not.toMatch(/<w:(?:ins|del)\b/u);
      expect(untrackedParagraphProperties(rejectedXml)).toBe(
        expectedParagraphProperties(beforeDirect, STYLE_ID),
      );
      const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
      expect(reopenedRejected.snapshot().blocks.at(0)).toEqual(rejecting.snapshot().blocks.at(0));
    },
  );

  test("suggested replaceBlock tracks style provenance through acceptance and rejection", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const applySuggestion = () => {
      const view = makeView(source);
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected a replacement block");
      }
      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "suggested-replacement-style",
            type: "replaceBlock",
            blockId: block.id,
            text: REPLACEMENT_TEXT,
            styleId: NEXT_STYLE_ID,
          },
        ],
        mode: "suggested",
        author: "assistant",
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toEqual([
        {
          id: "suggested-replacement-style",
          revisionId: 1,
          revisionIds: [1, 2, 4],
          suggestionId: "suggested-replacement-style",
        },
      ]);
      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected a suggested replacement paragraph");
      }
      expect(expectParagraphAttrs(paragraph)).toMatchObject({
        styleId: NEXT_STYLE_ID,
        alignment: "center",
        alignmentFromStyle: "both",
        _originalFormatting: { styleId: NEXT_STYLE_ID, alignment: "center" },
        _propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: {
              id: 4,
              author: "assistant",
              date: OPTIONS.timestamp,
              provenance: "suggested",
              suggestionId: "suggested-replacement-style",
            },
            previousFormatting: { styleId: STYLE_ID, alignment: "center" },
          },
        ],
      });
      expect(getSuggestions(view.state)).toEqual([
        expect.objectContaining({
          suggestionId: "suggested-replacement-style",
          kinds: ["insertion", "deletion", "formatting"],
          appliedAs: "tracked",
        }),
      ]);
      return view;
    };

    const pending = applySuggestion();
    const pendingBuffer = await createDocx(fromProseDoc(pending.state.doc, source));
    const pendingXml = firstParagraphXml(await mainDocumentXml(pendingBuffer));
    expect(pendingXml).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
    expect(pendingXml).toContain(TEXT);
    expect(pendingXml).not.toContain(REPLACEMENT_TEXT);
    expect(untrackedParagraphProperties(pendingXml)).toBe(
      expectedParagraphProperties("center", STYLE_ID),
    );

    expect(
      acceptSuggestion("suggested-replacement-style", {
        author: "reviewer",
        date: OPTIONS.timestamp,
      })(pending.state, pending.dispatch),
    ).toBe(true);
    expect(getSuggestions(pending.state)).toEqual([]);
    const acceptedSuggestionParagraph = pending.state.doc.firstChild;
    if (!acceptedSuggestionParagraph) {
      panic("expected an accepted suggestion paragraph");
    }
    expect(expectParagraphAttrs(acceptedSuggestionParagraph)._propertyChanges).toEqual([
      {
        type: "paragraphPropertyChange",
        info: {
          id: 4,
          author: "reviewer",
          date: OPTIONS.timestamp,
          provenance: "user",
          suggestionId: null,
        },
        previousFormatting: { styleId: STYLE_ID, alignment: "center" },
      },
    ]);
    const acceptedSuggestionBuffer = await createDocx(fromProseDoc(pending.state.doc, source));
    const acceptedSuggestionXml = await mainDocumentXml(acceptedSuggestionBuffer);
    expect(acceptedSuggestionXml).toContain('<w:ins w:id="2" w:author="reviewer"');
    expect(acceptedSuggestionXml).toContain('<w:del w:id="1" w:author="reviewer"');
    expect(acceptedSuggestionXml).toContain('<w:pPrChange w:id="4" w:author="reviewer"');
    expect(trackedParagraphPropertyParts(acceptedSuggestionXml)).toEqual({
      current: expectedParagraphProperties("center", NEXT_STYLE_ID),
      previous: expectedParagraphProperties("center", STYLE_ID),
    });

    const resolvingAccepted = await FolioDocxReviewer.fromBuffer(acceptedSuggestionBuffer);
    expect(resolvingAccepted.acceptAll()).toBe(3);
    const acceptedBuffer = await resolvingAccepted.toBuffer();
    const reopenedAccepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
    expect(reopenedAccepted.getChanges()).toEqual([]);
    expect(reopenedAccepted.snapshot().blocks.at(0)).toEqual({
      id: "12345678",
      kind: "paragraph",
      text: REPLACEMENT_TEXT,
      styleId: NEXT_STYLE_ID,
      directAlignment: "center",
    });

    const rejecting = applySuggestion();
    expect(
      rejectSuggestion("suggested-replacement-style")(rejecting.state, rejecting.dispatch),
    ).toBe(true);
    expect(getSuggestions(rejecting.state)).toEqual([]);
    const rejectedParagraph = rejecting.state.doc.firstChild;
    if (!rejectedParagraph) {
      panic("expected a rejected suggestion paragraph");
    }
    expect(expectParagraphAttrs(rejectedParagraph)).toMatchObject({
      styleId: STYLE_ID,
      alignment: "center",
      alignmentFromStyle: "right",
      _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
    });
    expect(expectParagraphAttrs(rejectedParagraph)._propertyChanges).toBeUndefined();
    const rejectedBuffer = await createDocx(fromProseDoc(rejecting.state.doc, source));
    const rejectedXml = firstParagraphXml(await mainDocumentXml(rejectedBuffer));
    expect(rejectedXml).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
    expect(rejectedXml).toContain(TEXT);
    expect(rejectedXml).not.toContain(REPLACEMENT_TEXT);
    expect(untrackedParagraphProperties(rejectedXml)).toBe(
      expectedParagraphProperties("center", STYLE_ID),
    );
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    expect(reopenedRejected.snapshot().blocks.at(0)).toEqual({
      id: "12345678",
      kind: "paragraph",
      text: TEXT,
      styleId: STYLE_ID,
      directAlignment: "center",
    });
  });

  test("suggested paragraph-only style changes remain independently resolvable", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const applySuggestion = () => {
      const view = makeView(source);
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected a paragraph property block");
      }
      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "suggested-paragraph-style",
            type: "setBlockParagraphProperties",
            blockId: block.id,
            properties: { styleId: NEXT_STYLE_ID, alignment: null },
          },
        ],
        mode: "suggested",
        author: "assistant",
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toEqual([
        {
          id: "suggested-paragraph-style",
          revisionId: 1,
          revisionIds: [1],
          suggestionId: "suggested-paragraph-style",
        },
      ]);
      expect(getSuggestions(view.state)).toEqual([
        expect.objectContaining({
          suggestionId: "suggested-paragraph-style",
          kinds: ["formatting"],
          appliedAs: "tracked",
        }),
      ]);
      return view;
    };

    const pending = applySuggestion();
    const pendingBuffer = await createDocx(fromProseDoc(pending.state.doc, source));
    const pendingXml = firstParagraphXml(await mainDocumentXml(pendingBuffer));
    expect(pendingXml).not.toContain("<w:pPrChange");
    expect(untrackedParagraphProperties(pendingXml)).toBe(
      expectedParagraphProperties("center", STYLE_ID),
    );

    expect(
      acceptSuggestion("suggested-paragraph-style", {
        author: "reviewer",
        date: OPTIONS.timestamp,
      })(pending.state, pending.dispatch),
    ).toBe(true);
    const acceptedSuggestionBuffer = await createDocx(fromProseDoc(pending.state.doc, source));
    const acceptedSuggestionXml = await mainDocumentXml(acceptedSuggestionBuffer);
    expect(acceptedSuggestionXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
    expect(trackedParagraphPropertyParts(acceptedSuggestionXml)).toEqual({
      current: expectedParagraphProperties(undefined, NEXT_STYLE_ID),
      previous: expectedParagraphProperties("center", STYLE_ID),
    });
    const resolvingAccepted = await FolioDocxReviewer.fromBuffer(acceptedSuggestionBuffer);
    expect(resolvingAccepted.acceptAll()).toBe(1);
    const acceptedBuffer = await resolvingAccepted.toBuffer();
    const reopenedAccepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
    expect(reopenedAccepted.snapshot().blocks.at(0)).toEqual({
      id: "12345678",
      kind: "paragraph",
      text: TEXT,
      styleId: NEXT_STYLE_ID,
    });

    const rejecting = applySuggestion();
    expect(rejectSuggestion("suggested-paragraph-style")(rejecting.state, rejecting.dispatch)).toBe(
      true,
    );
    const rejectedBuffer = await createDocx(fromProseDoc(rejecting.state.doc, source));
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    expect(reopenedRejected.snapshot().blocks.at(0)).toEqual({
      id: "12345678",
      kind: "paragraph",
      text: TEXT,
      styleId: STYLE_ID,
      directAlignment: "center",
    });
  });

  test("strips adjacent paragraph suggestions without overwriting a retained tracked change", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const view = makeView(source);
    const applyProperties = ({
      id,
      properties,
      mode,
      idSeed,
    }: {
      id: string;
      properties: { styleId?: string | null; alignment?: ParagraphAlignment | null };
      mode: "tracked-changes" | "suggested";
      idSeed: number;
    }) => {
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected an interleaved property-change block");
      }
      return applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [{ id, type: "setBlockParagraphProperties", blockId: block.id, properties }],
        mode,
        author: mode === "suggested" ? "assistant" : "reviewer",
        revisionStamp: { ...REVISION_STAMP, idSeed },
      });
    };

    expect(
      applyProperties({
        id: "tracked-style",
        properties: { styleId: NEXT_STYLE_ID },
        mode: "tracked-changes",
        idSeed: 1,
      }).applied,
    ).toEqual([{ id: "tracked-style", revisionId: 1, revisionIds: [1] }]);
    expect(
      applyProperties({
        id: "suggested-right",
        properties: { alignment: "right" },
        mode: "suggested",
        idSeed: 2,
      }).applied,
    ).toEqual([
      {
        id: "suggested-right",
        revisionId: 2,
        revisionIds: [2],
        suggestionId: "suggested-right",
      },
    ]);
    expect(
      applyProperties({
        id: "suggested-clear",
        properties: { alignment: null },
        mode: "suggested",
        idSeed: 3,
      }).applied,
    ).toEqual([
      {
        id: "suggested-clear",
        revisionId: 3,
        revisionIds: [3],
        suggestionId: "suggested-clear",
      },
    ]);

    const pendingParagraph = view.state.doc.firstChild;
    if (!pendingParagraph) {
      panic("expected an interleaved property-change paragraph");
    }
    const pendingAttrs = expectParagraphAttrs(pendingParagraph);
    expect(pendingAttrs.styleId).toBe(NEXT_STYLE_ID);
    expect(pendingAttrs.alignment).toBe("both");
    expect(pendingAttrs._propertyChanges?.map(({ info }) => info)).toEqual([
      { id: 1, author: "reviewer", date: OPTIONS.timestamp },
      {
        id: 2,
        author: "assistant",
        date: OPTIONS.timestamp,
        provenance: "suggested",
        suggestionId: "suggested-right",
      },
      {
        id: 3,
        author: "assistant",
        date: OPTIONS.timestamp,
        provenance: "suggested",
        suggestionId: "suggested-clear",
      },
    ]);

    const strippedBuffer = await createDocx(fromProseDoc(view.state.doc, source));
    const strippedXml = await mainDocumentXml(strippedBuffer);
    expect(strippedXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
    expect(trackedParagraphPropertyParts(strippedXml)).toEqual({
      current: expectedParagraphProperties("center", NEXT_STYLE_ID),
      previous: expectedParagraphProperties("center", STYLE_ID),
    });

    expect(rejectSuggestion("suggested-clear")(view.state, view.dispatch)).toBe(true);
    const singlyRejectedParagraph = view.state.doc.firstChild;
    if (!singlyRejectedParagraph) {
      panic("expected the singly rejected paragraph");
    }
    expect(expectParagraphAttrs(singlyRejectedParagraph).alignment).toBe("right");
    expect(getSuggestions(view.state).map(({ suggestionId }) => suggestionId)).toEqual([
      "suggested-right",
    ]);
    expect(rejectSuggestion("suggested-right")(view.state, view.dispatch)).toBe(true);
    const retainedParagraph = view.state.doc.firstChild;
    if (!retainedParagraph) {
      panic("expected the retained tracked paragraph");
    }
    const retainedAttrs = expectParagraphAttrs(retainedParagraph);
    expect(retainedAttrs.styleId).toBe(NEXT_STYLE_ID);
    expect(retainedAttrs.alignment).toBe("center");
    expect(retainedAttrs._originalFormatting).toEqual({
      styleId: NEXT_STYLE_ID,
      alignment: "center",
    });
    expect(retainedAttrs._propertyChanges).toEqual([
      {
        type: "paragraphPropertyChange",
        info: { id: 1, author: "reviewer", date: OPTIONS.timestamp },
        previousFormatting: { styleId: STYLE_ID, alignment: "center" },
      },
    ]);

    rejectAllChanges()(view.state, view.dispatch);
    const rejectedBuffer = await createDocx(fromProseDoc(view.state.doc, source));
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    expect(reopenedRejected.snapshot().blocks.at(0)).toEqual({
      id: "12345678",
      kind: "paragraph",
      text: TEXT,
      styleId: STYLE_ID,
      directAlignment: "center",
    });
  });

  test("rebases a retained tracked property change over an earlier suggestion", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const view = makeView(source);
    const applyProperties = ({
      id,
      properties,
      mode,
      idSeed,
    }: {
      id: string;
      properties: { styleId?: string; alignment?: ParagraphAlignment | null };
      mode: "tracked-changes" | "suggested";
      idSeed: number;
    }) => {
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected a preceding-suggestion block");
      }
      return applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [{ id, type: "setBlockParagraphProperties", blockId: block.id, properties }],
        mode,
        author: mode === "suggested" ? "assistant" : "reviewer",
        revisionStamp: { ...REVISION_STAMP, idSeed },
      });
    };

    expect(
      applyProperties({
        id: "suggested-right",
        properties: { alignment: "right" },
        mode: "suggested",
        idSeed: 1,
      }).applied,
    ).toEqual([
      {
        id: "suggested-right",
        revisionId: 1,
        revisionIds: [1],
        suggestionId: "suggested-right",
      },
    ]);
    expect(
      applyProperties({
        id: "tracked-style",
        properties: { styleId: NEXT_STYLE_ID },
        mode: "tracked-changes",
        idSeed: 2,
      }).applied,
    ).toEqual([{ id: "tracked-style", revisionId: 2, revisionIds: [2] }]);

    const strippedBuffer = await createDocx(fromProseDoc(view.state.doc, source));
    const strippedXml = await mainDocumentXml(strippedBuffer);
    expect(strippedXml.match(/<w:pPrChange\b/gu)).toHaveLength(1);
    expect(trackedParagraphPropertyParts(strippedXml)).toEqual({
      current: expectedParagraphProperties("right", NEXT_STYLE_ID),
      previous: expectedParagraphProperties("center", STYLE_ID),
    });

    const rejecting = await FolioDocxReviewer.fromBuffer(strippedBuffer);
    expect(rejecting.rejectAll()).toBe(1);
    const rejectedBuffer = await rejecting.toBuffer();
    const reopenedRejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    expect(reopenedRejected.snapshot().blocks.at(0)).toEqual({
      id: "12345678",
      kind: "paragraph",
      text: TEXT,
      styleId: STYLE_ID,
      directAlignment: "center",
    });
  });

  test.each([
    {
      id: "suggestion-left",
      liveAlignment: "both",
      remaining: [
        { id: 2, previousAlignment: "center" },
        { id: 3, previousAlignment: "right" },
      ],
    },
    {
      id: "suggestion-right",
      liveAlignment: "both",
      remaining: [
        { id: 1, previousAlignment: "center" },
        { id: 3, previousAlignment: "left" },
      ],
    },
    {
      id: "suggestion-both",
      liveAlignment: "right",
      remaining: [
        { id: 1, previousAlignment: "center" },
        { id: 2, previousAlignment: "left" },
      ],
    },
  ] as const)(
    "rejecting the individual $id suggestion rebases its retained property-change chain",
    async ({ id, liveAlignment, remaining }) => {
      const source = alignmentDocumentModel({
        directAlignment: "center",
        inheritedAlignment: "right",
      });
      const view = applyThreeAlignmentSuggestions(makeView(source));

      expect(rejectSuggestion(id)(view.state, view.dispatch)).toBe(true);
      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected a paragraph after individual suggestion rejection");
      }
      const attrs = expectParagraphAttrs(paragraph);
      expect(attrs.alignment).toBe(liveAlignment);
      expect(
        attrs._propertyChanges?.map((change) => ({
          id: change.info.id,
          previousAlignment: change.previousFormatting?.alignment,
        })),
      ).toEqual(remaining);

      expect(rejectAllSuggestions()(view.state, view.dispatch)).toBe(true);
      const restored = view.state.doc.firstChild;
      if (!restored) {
        panic("expected a paragraph after rejecting the remaining suggestions");
      }
      expect(expectParagraphAttrs(restored)).toMatchObject({
        alignment: "center",
        _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
      });
      expect(expectParagraphAttrs(restored)._propertyChanges).toBeUndefined();

      const saved = await createDocx(fromProseDoc(view.state.doc, source));
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expectDirectAlignmentModel(reopened, "center", STYLE_ID);
    },
  );

  test.each(SUGGESTION_REJECTION_ORDERS)(
    "rejects a paragraph suggestion chain $label without depending on resolution order",
    async ({ ids }) => {
      const source = alignmentDocumentModel({
        directAlignment: "center",
        inheritedAlignment: "right",
      });
      const view = applyThreeAlignmentSuggestions(makeView(source));

      for (const id of ids) {
        expect(rejectSuggestion(id)(view.state, view.dispatch)).toBe(true);
      }
      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected a paragraph after ordered suggestion rejection");
      }
      expect(expectParagraphAttrs(paragraph)).toMatchObject({
        alignment: "center",
        _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
      });
      expect(expectParagraphAttrs(paragraph)._propertyChanges).toBeUndefined();
      expect(getSuggestions(view.state)).toEqual([]);

      const saved = await createDocx(fromProseDoc(view.state.doc, source));
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expectDirectAlignmentModel(reopened, "center", STYLE_ID);
    },
  );

  test("rejects every paragraph suggestion in one transaction", async () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const view = applyThreeAlignmentSuggestions(makeView(source));

    expect(rejectAllSuggestions()(view.state, view.dispatch)).toBe(true);
    const paragraph = view.state.doc.firstChild;
    if (!paragraph) {
      panic("expected a paragraph after bulk suggestion rejection");
    }
    expect(expectParagraphAttrs(paragraph)).toMatchObject({
      alignment: "center",
      _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
    });
    expect(expectParagraphAttrs(paragraph)._propertyChanges).toBeUndefined();

    const saved = await createDocx(fromProseDoc(view.state.doc, source));
    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expectDirectAlignmentModel(reopened, "center", STYLE_ID);
  });

  test.each(SUGGESTION_REJECTION_ORDERS)(
    "rebases interleaved tracked paragraph changes when suggestions resolve $label",
    async ({ ids }) => {
      const source = alignmentDocumentModel({
        directAlignment: "center",
        inheritedAlignment: "right",
      });
      const view = makeView(source);
      for (const [index, edit] of (
        [
          { id: "suggestion-left", alignment: "left", mode: "suggested" },
          { id: "tracked-right", alignment: "right", mode: "tracked-changes" },
          { id: "suggestion-right", alignment: "both", mode: "suggested" },
          { id: "tracked-both", alignment: "distribute", mode: "tracked-changes" },
          { id: "suggestion-both", alignment: null, mode: "suggested" },
        ] as const
      ).entries()) {
        const outcome = applyParagraphProperties({
          view,
          id: edit.id,
          properties: { alignment: edit.alignment },
          mode: edit.mode,
          idSeed: index + 1,
        });
        expect(outcome.skipped).toEqual([]);
      }

      for (const id of ids) {
        expect(rejectSuggestion(id)(view.state, view.dispatch)).toBe(true);
      }
      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected a paragraph after interleaved suggestion rejection");
      }
      const attrs = expectParagraphAttrs(paragraph);
      expect(attrs).toMatchObject({
        alignment: "distribute",
        _originalFormatting: { styleId: STYLE_ID, alignment: "distribute" },
      });
      expect(
        attrs._propertyChanges?.map((change) => ({
          id: change.info.id,
          provenance: change.info.provenance,
          previousAlignment: change.previousFormatting?.alignment,
        })),
      ).toEqual([
        { id: 2, provenance: undefined, previousAlignment: "center" },
        { id: 4, provenance: undefined, previousAlignment: "right" },
      ]);

      expect(rejectAllChanges()(view.state, view.dispatch)).toBe(true);
      const restored = view.state.doc.firstChild;
      if (!restored) {
        panic("expected a paragraph after rejecting retained tracked changes");
      }
      expect(expectParagraphAttrs(restored)).toMatchObject({
        alignment: "center",
        _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
      });
      expect(expectParagraphAttrs(restored)._propertyChanges).toBeUndefined();

      const saved = await createDocx(fromProseDoc(view.state.doc, source));
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expectDirectAlignmentModel(reopened, "center", STYLE_ID);
    },
  );

  test("rejects a style-only change without materializing inherited alignment", async () => {
    const source = alignmentDocumentModel({ inheritedAlignment: "right" });
    const view = {
      state: EditorState.create({ doc: toProseDoc(source) }),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    if (!block) {
      panic("expected a paragraph without a styles plugin");
    }
    const outcome = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "style-only-without-resolver",
          type: "setBlockParagraphProperties",
          blockId: block.id,
          properties: { styleId: NEXT_STYLE_ID },
        },
      ],
      mode: "tracked-changes",
      author: OPTIONS.author,
      revisionStamp: REVISION_STAMP,
    });
    expect(outcome.skipped).toEqual([]);
    expect(outcome.applied).toHaveLength(1);

    rejectAllChanges()(view.state, view.dispatch);
    const rejectedParagraph = view.state.doc.firstChild;
    if (!rejectedParagraph) {
      panic("expected the rejected paragraph");
    }
    const rejectedAttrs = expectParagraphAttrs(rejectedParagraph);
    expect(rejectedAttrs.styleId).toBe(STYLE_ID);
    expect(rejectedAttrs._originalFormatting).toEqual({ styleId: STYLE_ID });
    expect(rejectedAttrs._propertyChanges).toBeUndefined();
    expect(createFolioAIEditSnapshot(view.state.doc).blocks.at(0)?.directAlignment).toBeUndefined();

    const rejectedBuffer = await createDocx(fromProseDoc(view.state.doc, source));
    const rejectedXml = firstParagraphXml(await mainDocumentXml(rejectedBuffer));
    expect(rejectedXml).not.toContain("<w:pPrChange");
    expect(untrackedParagraphProperties(rejectedXml)).toBe(
      expectedParagraphProperties(undefined, STYLE_ID),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    expectDirectAlignmentModel(reopened, undefined, STYLE_ID);
    const reopenedParagraph = toProseDoc(reopened.toDocument()).firstChild;
    if (!reopenedParagraph) {
      panic("expected the reopened paragraph");
    }
    expect(expectParagraphAttrs(reopenedParagraph)).toMatchObject({
      styleId: STYLE_ID,
      alignment: "right",
      alignmentFromStyle: "right",
      _originalFormatting: { styleId: STYLE_ID },
    });
  });

  test("accept and reject immediately restore the correct effective and direct values", () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });

    const applyTrackedClear = () => {
      const view = makeView(source);
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected an editable paragraph");
      }
      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "clear-alignment",
            type: "setBlockParagraphProperties",
            blockId: block.id,
            properties: { alignment: null },
          },
        ],
        mode: "tracked-changes",
        author: OPTIONS.author,
        revisionStamp: REVISION_STAMP,
      });
      expect(outcome.skipped).toEqual([]);
      expect(outcome.applied).toHaveLength(1);
      const paragraph = view.state.doc.firstChild;
      if (!paragraph) {
        panic("expected a paragraph");
      }
      const pending = expectParagraphAttrs(paragraph);
      expect(pending).toMatchObject({
        alignment: "right",
        alignmentFromStyle: "right",
        _originalFormatting: { styleId: STYLE_ID },
      });
      expect(pending._propertyChanges).toEqual([
        expect.objectContaining({
          previousFormatting: {
            styleId: STYLE_ID,
            alignment: "center",
          },
        }),
      ]);
      return view;
    };

    const accepting = applyTrackedClear();
    acceptAllChanges()(accepting.state, accepting.dispatch);
    const acceptedParagraph = accepting.state.doc.firstChild;
    if (!acceptedParagraph) {
      panic("expected an accepted paragraph");
    }
    expect(expectParagraphAttrs(acceptedParagraph)).toMatchObject({
      alignment: "right",
      alignmentFromStyle: "right",
      _originalFormatting: { styleId: STYLE_ID },
    });
    const acceptedAttrs = expectParagraphAttrs(acceptedParagraph);
    expect(acceptedAttrs._propertyChanges).toBeUndefined();
    expect(paragraphFormatting(fromProseDoc(accepting.state.doc, source))).toEqual({
      styleId: STYLE_ID,
    });

    const rejecting = applyTrackedClear();
    rejectAllChanges()(rejecting.state, rejecting.dispatch);
    const rejectedParagraph = rejecting.state.doc.firstChild;
    if (!rejectedParagraph) {
      panic("expected a rejected paragraph");
    }
    expect(expectParagraphAttrs(rejectedParagraph)).toMatchObject({
      alignment: "center",
      alignmentFromStyle: "right",
      _originalFormatting: { styleId: STYLE_ID, alignment: "center" },
    });
    expect(expectParagraphAttrs(rejectedParagraph)._propertyChanges).toBeUndefined();
    expect(paragraphFormatting(fromProseDoc(rejecting.state.doc, source))).toEqual({
      styleId: STYLE_ID,
      alignment: "center",
    });
  });

  test("preserves imported direct provenance when it equals the inherited value", () => {
    const source = alignmentDocumentModel({
      directAlignment: "right",
      inheritedAlignment: "right",
    });
    const pmDocument = toProseDoc(source);
    const paragraph = pmDocument.firstChild;
    if (!paragraph) {
      panic("expected a paragraph");
    }

    expect(expectParagraphAttrs(paragraph)).toMatchObject({
      alignment: "right",
      alignmentFromStyle: "right",
      _originalFormatting: { styleId: STYLE_ID, alignment: "right" },
    });
    expect(paragraphFormatting(fromProseDoc(pmDocument, source))).toEqual({
      styleId: STYLE_ID,
      alignment: "right",
    });
  });

  test("infers direct provenance for PM-created or pasted paragraphs", () => {
    const source = alignmentDocumentModel({ inheritedAlignment: "center" });
    const imported = toProseDoc(source);
    const importedParagraph = imported.firstChild;
    if (!importedParagraph) {
      panic("expected a paragraph");
    }
    const created = imported.type.create(null, [
      importedParagraph.type.create(
        {
          ...importedParagraph.attrs,
          alignment: "right",
          alignmentFromStyle: "center",
          _originalFormatting: null,
        },
        importedParagraph.content,
      ),
    ]);

    expect(createFolioAIEditSnapshot(created).blocks.at(0)?.directAlignment).toBe("right");
    expect(paragraphFormatting(fromProseDoc(created, source))).toEqual({
      styleId: STYLE_ID,
      alignment: "right",
    });
  });

  test("alignment and style commands keep direct and inherited provenance synchronized", () => {
    const source = alignmentDocumentModel({
      directAlignment: "center",
      inheritedAlignment: "right",
    });
    const view = makeView(source);

    setAlignment("lowKashida")(view.state, view.dispatch);
    const directlyAligned = view.state.doc.firstChild;
    if (!directlyAligned) {
      panic("expected a directly aligned paragraph");
    }
    expect(expectParagraphAttrs(directlyAligned)).toMatchObject({
      alignment: "lowKashida",
      alignmentFromStyle: "right",
      _originalFormatting: { styleId: STYLE_ID, alignment: "lowKashida" },
    });
    expect(paragraphFormatting(fromProseDoc(view.state.doc, source))).toEqual({
      styleId: STYLE_ID,
      alignment: "lowKashida",
    });

    applyStyle(NEXT_STYLE_ID, { paragraphFormatting: { alignment: "both" } })(
      view.state,
      view.dispatch,
    );
    const restyled = view.state.doc.firstChild;
    if (!restyled) {
      panic("expected a restyled paragraph");
    }
    expect(expectParagraphAttrs(restyled)).toMatchObject({
      styleId: NEXT_STYLE_ID,
      alignment: "both",
      alignmentFromStyle: "both",
    });
    expect(paragraphFormatting(fromProseDoc(view.state.doc, source))).toEqual({
      styleId: NEXT_STYLE_ID,
    });
  });

  test("adds no serialized state to one thousand ordinary paragraphs", () => {
    const source = createEmptyDocument();
    source.package.document.content = Array.from({ length: 1_000 }, (_, index) => ({
      type: "paragraph" as const,
      paraId: index.toString(16).padStart(8, "0"),
      content: [{ type: "run" as const, content: [{ type: "text" as const, text: TEXT }] }],
    }));
    const json = toProseDoc(source).toJSON();
    const serialized = JSON.stringify(json);
    const withoutAlignmentProvenance = JSON.stringify(json, (key, value) =>
      key === "alignmentFromStyle" ? undefined : value,
    );

    expect(json.content).toHaveLength(1_000);
    expect(serialized.length).toBeLessThanOrEqual(ORDINARY_PARAGRAPH_STATE_SIZE_BUDGET);
    expect(serialized.length - withoutAlignmentProvenance.length).toBe(0);
  });

  test("bounds inherited-alignment provenance to one field per styled paragraph", () => {
    const source = alignmentDocumentModel({ inheritedAlignment: "right" });
    const template = source.package.document.content.at(0);
    if (template?.type !== "paragraph") {
      panic("expected a paragraph template");
    }
    source.package.document.content = Array.from({ length: 1_000 }, (_, index) => ({
      ...template,
      paraId: index.toString(16).padStart(8, "0"),
    }));
    const json = toProseDoc(source).toJSON();
    const serialized = JSON.stringify(json);
    const withoutAlignmentProvenance = JSON.stringify(json, (key, value) =>
      key === "alignmentFromStyle" ? undefined : value,
    );

    expect(json.content).toHaveLength(1_000);
    expect(serialized.length).toBeLessThanOrEqual(STYLED_PARAGRAPH_STATE_SIZE_BUDGET);
    expect(serialized.length - withoutAlignmentProvenance.length).toBe(
      STYLED_ALIGNMENT_PROVENANCE_SIZE_BUDGET,
    );
  });
});
