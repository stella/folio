/**
 * Table, row, column and cell public operations in the canonical session
 * give the package the default engine gives: the same tables, rows, cells
 * and cell texts pending, once every change is accepted, and once every
 * change is rejected, in direct and tracked mode, addressed through any
 * cell's paragraph. A suggestion the default engine cannot record gives the
 * direct edit once accepted and the document as it was once rejected.
 */

import {
  afterAll,
  beforeAll,
  describe,
  expect,
  setDefaultTimeout,
  setSystemTime,
  test,
} from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { FolioDocumentOperation } from "../ai-edits";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { TABLE_INSERTION_MODES } from "../ai-edits/tableInsertionModes";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import {
  type FolioDocumentOperationMode,
  isFolioDocumentOperationModeSupported,
} from "../document-operations";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { BlockContent, TableCell } from "../types/document";

setDefaultTimeout(propertyTestTimeout(240_000));
// Saves update core-property dates; a fixed clock makes exact ZIP equality
// distinguish document mutation from the save's modification timestamp.
beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const paragraph = (id: string, text: string) =>
  `<w:p w14:paraId="${id}">${text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ""}</w:p>`;
const cell = (id: string, text: string, properties = "") =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${properties}</w:tcPr>${paragraph(id, text)}</w:tc>`;
const table = (rows: readonly string[], columns: number) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join("")}</w:tbl>`;

/** Bodies holding one table between two paragraphs; the cell paragraphs are the anchors. */
const BODIES = {
  "a plain table": {
    xml: [
      paragraph("10000001", "Before"),
      table(
        [
          cell("20000001", "A1") + cell("20000002", "B1"),
          cell("20000003", "A2") + cell("20000004", "B2"),
        ],
        2,
      ),
      paragraph("10000002", "After"),
    ],
    anchors: ["20000001", "20000002", "20000003", "20000004"],
  },
  "a vertical merge": {
    xml: [
      paragraph("10000001", "Before"),
      table(
        [
          cell("30000001", "A1", '<w:vMerge w:val="restart"/>') +
            cell("30000002", "B1"),
          cell("30000003", "", "<w:vMerge/>") + cell("30000004", "B2"),
          cell("30000005", "A3") + cell("30000006", "B3"),
        ],
        2,
      ),
      paragraph("10000002", "After"),
    ],
    anchors: ["30000001", "30000002", "30000004", "30000005", "30000006"],
  },
  "a table ending the body": {
    xml: [
      paragraph("10000001", "Before"),
      table(
        [
          cell("40000001", "A1") +
            cell("40000002", "B1") +
            cell("40000003", "C1"),
        ],
        3,
      ),
      paragraph("10000002", ""),
    ],
    anchors: ["40000001", "40000002", "40000003"],
  },
} as const;

/** Bodies the cell and column operations also run on. */
const MORE_BODIES = {
  "a horizontal merge": {
    xml: [
      paragraph("10000001", "Before"),
      table(
        [
          cell("50000001", "A1", '<w:gridSpan w:val="2"/>') +
            cell("50000002", "C1"),
          cell("50000003", "A2") +
            cell("50000004", "") +
            cell("50000005", "C2"),
        ],
        3,
      ),
      paragraph("10000002", "After"),
    ],
    anchors: ["50000001", "50000002", "50000003", "50000004", "50000005"],
  },
  "a one-column table": {
    xml: [
      paragraph("10000001", "Before"),
      table([cell("60000001", "S1"), cell("60000002", "")], 1),
      paragraph("10000002", "After"),
    ],
    anchors: ["60000001", "60000002"],
  },
} as const;

/** A paragraph whose break is pending deletion, before a table and before a paragraph. */
const deletedBreak = (id: string, text: string) =>
  `<w:p w14:paraId="${id}"><w:pPr><w:rPr><w:del w:id="90" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

type RemovedJoinParagraphOptions = {
  properties?: string;
  keptContent?: string;
};
/** A paragraph a pending join reaches, with its words and mark deleted. */
const removedJoinParagraph = ({
  properties = "",
  keptContent = "",
}: RemovedJoinParagraphOptions = {}) =>
  `<w:p w14:paraId="70000009"><w:pPr>${properties}<w:rPr><w:del w:id="91" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="92" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Gone</w:delText></w:r></w:del>${keptContent}</w:p>`;

/** Bodies tables are inserted into, beside paragraphs and tables. */
const PLACEMENT_BODIES = {
  "a deleted break before a table": {
    xml: [
      deletedBreak("70000001", "Joined"),
      table([cell("70000002", "A1") + cell("70000003", "B1")], 2),
      paragraph("10000002", "After"),
    ],
    anchors: ["70000001", "70000002", "70000003", "10000002"],
  },
  "a deleted break before a paragraph": {
    xml: [
      paragraph("10000001", "Before"),
      deletedBreak("70000004", "Joined"),
      paragraph("70000005", "Next"),
      table([cell("70000006", "A1")], 1),
      paragraph("10000002", ""),
    ],
    anchors: ["10000001", "70000004", "70000005", "70000006"],
  },
  "a table starting the body": {
    xml: [
      table([cell("60000101", "A1") + cell("60000102", "B1")], 2),
      paragraph("10000002", "After"),
    ],
    anchors: ["60000101", "60000102", "10000002"],
  },
  "a table at the body's end": {
    xml: [paragraph("10000001", "Before"), table([cell("60000103", "A1")], 1)],
    anchors: ["10000001", "60000103"],
  },
  "two tables touching": {
    xml: [
      paragraph("10000001", "Before"),
      table([cell("60000104", "A1")], 1),
      table([cell("60000105", "X1")], 1),
      paragraph("10000002", "After"),
    ],
    anchors: ["60000104", "60000105"],
  },
  "a nested table": {
    xml: [
      table(
        [
          `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${table([cell("60000106", "N1")], 1)}${paragraph("60000107", "Outer")}</w:tc>`,
        ],
        1,
      ),
      paragraph("10000002", "After"),
    ],
    anchors: ["60000106", "60000107", "10000002"],
  },
} as const;

const ALL_BODIES = { ...BODIES, ...MORE_BODIES, ...PLACEMENT_BODIES };
type Body = keyof typeof ALL_BODIES;

/** Bodies the canonical session's own review also runs on. */
const REVIEW_BODIES = {
  ...BODIES,
  "a vertical merge over three rows": {
    xml: [
      paragraph("10000001", "Before"),
      table(
        [
          cell("70000001", "A1", '<w:vMerge w:val="restart"/>') +
            cell("70000002", "B1"),
          cell("70000003", "", "<w:vMerge/>") + cell("70000004", "B2"),
          cell("70000005", "", "<w:vMerge/>") + cell("70000006", "B3"),
          cell("70000007", "A4") + cell("70000008", "B4"),
        ],
        2,
      ),
      paragraph("10000002", "After"),
    ],
    anchors: ["70000001", "70000002", "70000004", "70000006", "70000007"],
  },
} as const;

const PENDING_BOUNDARY_BODIES = {
  "a pending join into a deleted paragraph before a table": {
    xml: [
      deletedBreak("70000008", "Kept"),
      removedJoinParagraph(),
      table([cell("70000010", "A1") + cell("70000011", "B1")], 2),
      paragraph("10000002", "After"),
    ],
    anchors: ["70000008", "70000009", "70000010", "70000011", "10000002"],
  },
  "a pending join into a paragraph retaining a range before a table": {
    xml: [
      deletedBreak("70000008", "Kept"),
      removedJoinParagraph({
        keptContent:
          '<w:bookmarkStart w:id="73" w:name="Retained"/><w:bookmarkEnd w:id="73"/>',
      }),
      table([cell("70000010", "A1") + cell("70000011", "B1")], 2),
      paragraph("10000002", "After"),
    ],
    anchors: ["70000008", "70000009", "70000010", "70000011", "10000002"],
  },
  "a pending join into a paragraph retaining a section before a table": {
    xml: [
      deletedBreak("70000008", "Kept"),
      removedJoinParagraph({
        properties: '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      }),
      table([cell("70000010", "A1") + cell("70000011", "B1")], 2),
      paragraph("10000002", "After"),
    ],
    anchors: ["70000008", "70000009", "70000010", "70000011", "10000002"],
  },
} as const;

type ReviewBody = keyof typeof REVIEW_BODIES;

const PACKAGE_BODIES = {
  ...ALL_BODIES,
  ...REVIEW_BODIES,
  ...PENDING_BOUNDARY_BODIES,
};

const packageOf = async (body: keyof typeof PACKAGE_BODIES) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${PACKAGE_BODIES[body].xml.join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** Each operation, addressed through a cell's paragraph and, for a merge's far corner, the body's last cell. */
const OPERATIONS = {
  deleteTable: (blockId: string) => ({
    id: "op",
    type: "deleteTable",
    blockId,
  }),
  deleteTableRow: (blockId: string) => ({
    id: "op",
    type: "deleteTableRow",
    blockId,
  }),
  deleteTableColumn: (blockId: string) => ({
    id: "op",
    type: "deleteTableColumn",
    blockId,
  }),
  "insertTableRow after": (blockId: string) => ({
    id: "op",
    type: "insertTableRow",
    blockId,
    position: "after",
    cellTexts: ["New"],
  }),
  "insertTableRow before, two lines": (blockId: string) => ({
    id: "op",
    type: "insertTableRow",
    blockId,
    position: "before",
    cellTexts: ["First\nSecond"],
  }),
  "insertTableRow empty": (blockId: string) => ({
    id: "op",
    type: "insertTableRow",
    blockId,
  }),
} as const satisfies Record<
  string,
  (blockId: string, last: string) => FolioDocumentOperation
>;

const CELL_OPERATIONS = {
  deleteTableColumn: OPERATIONS.deleteTableColumn,
  "insertTableColumn after": (blockId: string) => ({
    id: "op",
    type: "insertTableColumn",
    blockId,
    position: "after",
    cellTexts: ["New"],
  }),
  "insertTableColumn before, two cells": (blockId: string) => ({
    id: "op",
    type: "insertTableColumn",
    blockId,
    position: "before",
    cellTexts: ["First\nSecond", "Third"],
  }),
  "insertTableColumn empty": (blockId: string) => ({
    id: "op",
    type: "insertTableColumn",
    blockId,
  }),
  "mergeTableCells down": (blockId: string) => ({
    id: "op",
    type: "mergeTableCells",
    blockId,
    rowCount: 2,
  }),
  "mergeTableCells to the last cell": (blockId: string, last: string) => ({
    id: "op",
    type: "mergeTableCells",
    blockId,
    endBlockId: last,
  }),
  splitTableCell: (blockId: string) => ({
    id: "op",
    type: "splitTableCell",
    blockId,
  }),
} as const satisfies Record<
  string,
  (blockId: string, last: string) => FolioDocumentOperation
>;

type Mode = FolioDocumentOperationMode;

/**
 * A paragraph's words, those a pending change holds included: a pending
 * table is compared by its structure and text, and the changes themselves
 * by what accepting and rejecting them give.
 */
const wordsOf = (value: unknown): string[] => {
  if (Array.isArray(value)) return value.flatMap(wordsOf);
  if (typeof value !== "object" || value === null) return [];
  const type: unknown = Reflect.get(value, "type");
  if (type === "text") {
    const text: unknown = Reflect.get(value, "text");
    return typeof text === "string" ? [text] : [];
  }
  if (type === "run") return wordsOf(Reflect.get(value, "content"));
  if (
    type === "insertion" ||
    type === "deletion" ||
    type === "moveFrom" ||
    type === "moveTo"
  )
    return wordsOf(Reflect.get(value, "content"));
  if (type === "hyperlink") return wordsOf(Reflect.get(value, "children"));
  return [];
};

const textOf = (blocks: readonly BlockContent[]): string[] =>
  blocks.flatMap((block) =>
    block.type === "paragraph" ? [wordsOf(block.content).join("")] : [],
  );

const cellShape = (tableCell: TableCell) => ({
  text: textOf(tableCell.content),
  gridSpan: tableCell.formatting?.gridSpan,
  vMerge: tableCell.formatting?.vMerge,
});

type CellShape = ReturnType<typeof cellShape>;

/**
 * A table's rows as both engines' files read alike: the default engine
 * drops a row that holds nothing but empty continuations of merges from the
 * rows above, ending a merge that no row continues any more, where the
 * canonical session keeps the row the file states.
 */
const withoutContinuationRows = (
  rows: readonly (readonly CellShape[])[],
): CellShape[][] => {
  const kept = rows.filter(
    (row) =>
      !row.every(
        (entry) =>
          entry.vMerge === "continue" &&
          entry.text.every((text) => text === ""),
      ),
  );
  return kept.map((row, rowIndex) =>
    row.map((entry, index) =>
      entry.vMerge === "restart" &&
      kept[rowIndex + 1]?.[index]?.vMerge !== "continue"
        ? { ...entry, vMerge: undefined }
        : entry,
    ),
  );
};

/** Each body block: a paragraph's text, a table's rows of cells. */
const blocksOf = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.map((block) =>
    block.type === "table"
      ? {
          rows: withoutContinuationRows(
            block.rows.map((row) => row.cells.map(cellShape)),
          ),
        }
      : { text: textOf([block]) },
  );
};

/**
 * The package pending, with every change accepted, and with every change
 * rejected by the legacy reviewer, whichever engine wrote it;
 * `reviewedInSession` reads the canonical session's own review.
 */
const outcomes = async (bytes: ArrayBuffer) => {
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
      session: "legacy",
    });
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return blocksOf(await reviewer.toBuffer());
  };
  return {
    pending: await blocksOf(bytes),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

/**
 * The canonical session's own review of a saved package: every change
 * accepted, and every change rejected, in the session that opens it.
 */
const reviewedInSession = async (bytes: ArrayBuffer) => {
  const resolved = async (resolution: "accept" | "reject") => {
    const document = await parseDocx(bytes.slice(0), {
      preloadFonts: false,
      detectVariables: false,
    });
    const session = createCanonicalSession(
      document,
      document.package.styles,
    ).unwrap();
    const state = EditorState.create({ schema, doc: session.projection.doc });
    const revisionIds = [
      ...new Set([
        ...getTrackedChangeStatsFromDoc(state.doc).ids,
        ...[
          ...JSON.stringify(session.document.package.document.content).matchAll(
            /"info":\{"id":(\d+)/gu,
          ),
        ].map(([, id]) => Number(id)),
      ]),
    ];
    if (revisionIds.length > 0) {
      const commit = session
        .prepareResolve(state, { revisionIds, resolution })
        .unwrap();
      expect(commit.publish().isOk()).toBe(true);
    }
    return blocksOf(
      (
        await serializeCanonicalSave({
          snapshot: session.captureSaveSnapshot(),
        })
      ).buffer,
    );
  };
  return {
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

type EngineRunOptions = {
  source: ArrayBuffer;
  operation: FolioDocumentOperation;
  mode: Mode;
  resolveSuggestion?: "accept" | "reject";
};
const viaLegacyEngine = async ({
  source,
  operation,
  mode,
  resolveSuggestion,
}: EngineRunOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0), {
    session: "legacy",
  });
  const before = await reviewer.toBuffer();
  const snapshot = reviewer.snapshot();
  const result = reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [operation] },
    { undefinedReferences: "refuse" },
  );
  if (resolveSuggestion !== undefined) {
    for (const applied of result.applied) {
      if (applied.suggestionId === undefined)
        throw new Error("The signature insertion has a suggestion id.");
      if (resolveSuggestion === "accept")
        expect(reviewer.acceptSuggestion(applied.suggestionId)).toBe(true);
      else expect(reviewer.rejectSuggestion(applied.suggestionId)).toBe(true);
    }
  }
  const bytes = await reviewer.toBuffer();
  if (result.applied.length === 0) {
    expect(result.undoHandle).toBeNull();
    expect(reviewer.snapshot()).toEqual(snapshot);
    expect(new Uint8Array(bytes)).toEqual(new Uint8Array(before));
  }
  return {
    applied: result.applied,
    skipped: result.skipped.map(({ reason }) => reason),
    bytes,
  };
};

const viaCanonicalEngine = async ({
  source,
  operation,
  mode,
}: EngineRunOptions) => {
  const document = await parseDocx(source.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const before = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
  });
  const documentBefore = structuredClone(session.document);
  const documentReference = session.document;
  const versionBefore = session.version;
  let publications = 0;
  const executor = new CanonicalPublicOperations({
    undoLimit: Number.POSITIVE_INFINITY,
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      publications++;
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode, operations: [operation] },
  });
  const saved = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
  });
  if (result.applied.length === 0) {
    expect(publications).toBe(0);
    expect(result.undoHandle).toBeNull();
    expect(session.canUndo).toBe(false);
    expect(session.version).toBe(versionBefore);
    expect(session.document).toBe(documentReference);
    expect(structuredClone(session.document)).toEqual(documentBefore);
    expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(before.buffer));
  }
  return {
    applied: result.applied,
    skipped: result.skipped.map(({ reason }) => reason),
    refusals: result.issues.flatMap(({ canonicalRefusal }) =>
      canonicalRefusal === undefined ? [] : [canonicalRefusal],
    ),
    bytes: saved.buffer,
  };
};

const INSERTIONS = {
  "insertTable after": (blockId: string) => ({
    id: "op",
    type: "insertTable",
    blockId,
    position: "after",
    rows: [
      ["N1", "N2"],
      ["N3", ""],
    ],
  }),
  "insertTable before": (blockId: string) => ({
    id: "op",
    type: "insertTable",
    blockId,
    position: "before",
    rows: [["N1"]],
  }),
  "insertSignatureTable after": (blockId: string) => ({
    id: "op",
    type: "insertSignatureTable",
    blockId,
    parties: [
      { name: "Buyer", signatory: "A. Person", title: "Director" },
      { name: "Seller" },
    ],
  }),
  "insertSignatureTable before": (blockId: string) => ({
    id: "op",
    type: "insertSignatureTable",
    blockId,
    position: "before",
    parties: [{ name: "Buyer" }],
  }),
} as const satisfies Record<
  string,
  (blockId: string, last: string) => FolioDocumentOperation
>;

const TABLE_INSERTION_TYPES = [
  ...new Set(Object.values(INSERTIONS).map((build) => build("70000001").type)),
];

const caseArbitrary = (
  bodies: readonly Body[],
  operations: Readonly<
    Record<string, (blockId: string, last: string) => FolioDocumentOperation>
  >,
  /** Paragraphs beside the table the operations may also name, where the body has them. */
  besides: readonly string[] = [],
) =>
  fc.constantFrom(...bodies).chain((body) =>
    fc.record({
      body: fc.constant(body),
      anchor: fc.constantFrom(
        ...new Set<string>([
          ...ALL_BODIES[body].anchors,
          ...besides.filter((id) =>
            ALL_BODIES[body].xml.some((xml) =>
              xml.startsWith(`<w:p w14:paraId="${id}"`),
            ),
          ),
        ]),
      ),
      operation: fc.constantFrom(...Object.keys(operations)),
      mode: fc.constantFrom<Mode>("direct", "tracked-changes"),
    }),
  );

/** Whether the paragraph `anchor` is the last block of the body. */
const endsTheBody = (body: Body, anchor: string): boolean =>
  ALL_BODIES[body].xml.at(-1)?.startsWith(`<w:p w14:paraId="${anchor}"`) ??
  false;

/** Whether the body ends in a table. */
const endsInTable = (body: keyof typeof PACKAGE_BODIES): boolean =>
  PACKAGE_BODIES[body].xml.at(-1)?.startsWith("<w:tbl") ?? false;

type Outcomes = Awaited<ReturnType<typeof outcomes>>;

/**
 * A table put after the body's last paragraph, or any edit of a body that
 * ends in a table: the canonical session's body ends with a paragraph, an
 * empty one after the table, where the default engine lets the table end the
 * body. Nothing else differs.
 */
const closedOutcomes = (outcome: Outcomes): Outcomes => {
  const closed = (blocks: Outcomes["pending"]): Outcomes["pending"] => {
    const last = blocks.at(-1);
    return last !== undefined && "rows" in last
      ? [...blocks, { text: [""] }]
      : blocks;
  };
  return {
    pending: closed(outcome.pending),
    accepted: closed(outcome.accepted),
    rejected: closed(outcome.rejected),
  };
};

/** The operation in both engines gives the same package, pending, accepted and rejected. */
const sameAsDefaultEngine =
  (
    operations: Readonly<
      Record<string, (blockId: string, last: string) => FolioDocumentOperation>
    >,
  ) =>
  async ({
    body,
    anchor,
    operation,
    mode,
  }: {
    body: Body;
    anchor: string;
    operation: string;
    mode: Mode;
  }) => {
    const source = await packageOf(body);
    const last = ALL_BODIES[body].anchors.at(-1) ?? anchor;
    const build = operations[operation];
    if (build === undefined) throw new Error(`No operation ${operation}.`);
    const op = build(anchor, last);
    const legacy = await viaLegacyEngine({
      source,
      operation: op,
      mode: mode,
    });
    const canonical = await viaCanonicalEngine({
      source,
      operation: op,
      mode: mode,
    });
    expect(canonical.refusals).toEqual([]);
    const closes =
      (operation.endsWith(" after") && endsTheBody(body, anchor)) ||
      endsInTable(body);
    const asCanonicalCloses = (outcome: Outcomes) =>
      closes ? closedOutcomes(outcome) : outcome;
    // A suggestion the default engine cannot record, such as a split or a
    // column through a merged cell: accepting it gives the direct edit,
    // rejecting it the document as it was.
    if (
      mode === "tracked-changes" &&
      legacy.skipped.length > 0 &&
      canonical.skipped.length === 0
    ) {
      const direct = await viaLegacyEngine({
        source,
        operation: op,
        mode: "direct",
      });
      expect(direct.skipped).toEqual([]);
      // Every other change the body already held is resolved alike.
      const resolved = await outcomes(canonical.bytes);
      expect(resolved.accepted).toEqual(
        asCanonicalCloses(await outcomes(direct.bytes)).accepted,
      );
      expect(resolved.rejected).toEqual((await outcomes(source)).rejected);
      return;
    }
    expect(canonical.skipped).toEqual(legacy.skipped);
    if (legacy.skipped.length > 0) return;
    expect(await outcomes(canonical.bytes)).toEqual(
      asCanonicalCloses(await outcomes(legacy.bytes)),
    );
  };

describe("table structure operations in the canonical session", () => {
  test("safe table placement honors published modes and the declared signature-suggestion gap", async () => {
    expect(new Set(TABLE_INSERTION_TYPES)).toEqual(
      new Set(Object.keys(TABLE_INSERTION_MODES)),
    );
    const cases = [
      {
        body: "a pending join into a deleted paragraph before a table",
        anchor: "70000010",
      },
      { body: "a table starting the body", anchor: "60000101" },
      { body: "a table at the body's end", anchor: "60000103" },
      { body: "two tables touching", anchor: "60000104" },
      { body: "a nested table", anchor: "60000106" },
      { body: "a nested table", anchor: "60000107" },
    ] as const;
    for (const value of cases) {
      const source = await packageOf(value.body);
      const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0), {
        session: "legacy",
      });
      expect(reviewer.snapshot().anchors[value.anchor]).toBeDefined();
      for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
        for (const type of TABLE_INSERTION_TYPES) {
          for (const side of ["before", "after"] as const) {
            const operation = INSERTIONS[`${type} ${side}`](value.anchor);
            const run = { source, operation, mode };
            const legacy = await viaLegacyEngine(run);
            const canonical = await viaCanonicalEngine(run);
            const supported = isFolioDocumentOperationModeSupported(type, mode);
            expect(
              legacy.skipped,
              JSON.stringify({ ...value, type, mode, side }),
            ).toEqual(supported ? [] : ["unsupportedMode"]);
            expect(legacy.applied).toHaveLength(supported ? 1 : 0);
            if (type === "insertSignatureTable" && mode === "suggested") {
              expect(canonical.skipped).toEqual(["unsupportedBlock"]);
              expect(canonical.applied).toEqual([]);
              expect(canonical.refusals).toEqual([
                { gap: "publicOps.tableStructure" },
              ]);
              expect(await outcomes(canonical.bytes)).toEqual(
                await outcomes(source),
              );
              // Legacy whole-table suggestions commit directly when accepted;
              // later rejectAll keeps that table. Canonical ordinary tracked
              // adoption cannot represent this boundary yet.
              for (const resolveSuggestion of ["accept", "reject"] as const) {
                const resolvedLegacy = await viaLegacyEngine({
                  ...run,
                  resolveSuggestion,
                });
                const expected =
                  resolveSuggestion === "accept"
                    ? (await viaLegacyEngine({ ...run, mode: "direct" })).bytes
                    : source;
                expect(await outcomes(resolvedLegacy.bytes)).toEqual(
                  await outcomes(expected),
                );
              }
              continue;
            }
            // On the canonical session accepting runs the joined words into
            // the table's first cell: a table put before that table would
            // take them, so it waits for the join to be resolved.
            if (
              supported &&
              side === "before" &&
              value.body ===
                "a pending join into a deleted paragraph before a table"
            ) {
              expect(canonical.skipped).toEqual([
                "pendingParagraphMarkDeletion",
              ]);
              expect(await outcomes(canonical.bytes)).toEqual(
                await outcomes(source),
              );
              continue;
            }
            expect(canonical.skipped).toEqual(legacy.skipped);
            expect(canonical.applied).toHaveLength(supported ? 1 : 0);
            expect(canonical.refusals).toEqual([]);
            const expected = await outcomes(legacy.bytes);
            expect(await outcomes(canonical.bytes)).toEqual(
              supported && endsInTable(value.body)
                ? closedOutcomes(expected)
                : expected,
            );
          }
        }
      }
    }
  });

  test("table placement at a pending deleted break refuses before any mutation in every mode", async () => {
    const source = await packageOf("a deleted break before a table");
    for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
      for (const operationType of TABLE_INSERTION_TYPES) {
        for (const target of [
          { anchor: "70000001", side: "after" },
          { anchor: "70000002", side: "before" },
        ] as const) {
          const build = INSERTIONS[`${operationType} ${target.side}`];
          const operation = build(target.anchor);
          const reason = isFolioDocumentOperationModeSupported(
            operationType,
            mode,
          )
            ? "pendingParagraphMarkDeletion"
            : "unsupportedMode";
          const legacy = await viaLegacyEngine({
            source,
            operation: operation,
            mode: mode,
          });
          const canonical = await viaCanonicalEngine({
            source,
            operation: operation,
            mode: mode,
          });
          expect(
            legacy.skipped,
            `${operationType} ${mode} ${target.side} ${target.anchor}`,
          ).toEqual([reason]);
          const gap =
            operationType === "insertSignatureTable" && mode === "suggested";
          expect(
            canonical.skipped,
            `${operationType} ${mode} ${target.side} ${target.anchor}`,
          ).toEqual([gap ? "unsupportedBlock" : reason]);
          expect(canonical.refusals).toEqual(
            gap ? [{ gap: "publicOps.tableStructure" }] : [],
          );
        }
      }
    }
  });

  test("a retained range or section does not bypass the pending-join placement refusal", async () => {
    for (const body of [
      "a pending join into a paragraph retaining a range before a table",
      "a pending join into a paragraph retaining a section before a table",
    ] as const) {
      const source = await packageOf(body);
      for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
        for (const operationType of TABLE_INSERTION_TYPES) {
          const operation = INSERTIONS[`${operationType} before`]("70000010");
          const reason = isFolioDocumentOperationModeSupported(
            operationType,
            mode,
          )
            ? "pendingParagraphMarkDeletion"
            : "unsupportedMode";
          const legacy = await viaLegacyEngine({ source, operation, mode });
          const canonical = await viaCanonicalEngine({
            source,
            operation,
            mode,
          });
          expect(legacy.skipped, `${body} ${operationType} ${mode}`).toEqual([
            reason,
          ]);
          const gap =
            operationType === "insertSignatureTable" && mode === "suggested";
          expect(canonical.skipped, `${body} ${operationType} ${mode}`).toEqual(
            [gap ? "unsupportedBlock" : reason],
          );
          expect(canonical.refusals).toEqual(
            gap ? [{ gap: "publicOps.tableStructure" }] : [],
          );
        }
      }
    }
  });

  test("reviews a deleted column including its collapsed vertical continuation", async () => {
    const source = await packageOf("a vertical merge");
    const operation = OPERATIONS.deleteTableColumn("30000001");
    const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0), {
      session: "canonical",
    });
    const result = reviewer.applyDocumentOperations(
      { version: 1, mode: "tracked-changes", operations: [operation] },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const ids = result.applied.at(0)?.revisionIds ?? [];
    expect(ids).toHaveLength(3);
    const pending = await reviewer.toBuffer();
    const open = () =>
      FolioDocxReviewer.fromBuffer(pending.slice(0), { session: "canonical" });
    const rejected = await open();
    expect(
      rejected
        .getChanges()
        .map(({ id }) => id)
        .toSorted(),
    ).toEqual([...ids].toSorted());
    expect(rejected.rejectAll()).toEqual({ resolved: 3, skipped: [] });
    expect(rejected.getChanges()).toEqual([]);
    expect(await blocksOf(await rejected.toBuffer())).toEqual(
      await blocksOf(source),
    );

    const accepted = await open();
    expect(accepted.acceptAll()).toEqual({ resolved: 3, skipped: [] });
    expect(accepted.getChanges()).toEqual([]);
    const direct = await viaLegacyEngine({
      source,
      operation,
      mode: "direct",
    });
    expect(direct.skipped).toEqual([]);
    expect(await blocksOf(await accepted.toBuffer())).toEqual(
      await blocksOf(direct.bytes),
    );
  });

  test("give the default engine's package, pending, accepted and rejected", async () => {
    await assertProperty(
      fc.asyncProperty(
        caseArbitrary(Object.keys(BODIES) as Body[], OPERATIONS),
        sameAsDefaultEngine(OPERATIONS),
      ),
      {
        numRuns: 60,
        id: "table structure operations in the canonical session give the default engine's package",
      },
    );
  });

  test("cell merges, splits and column changes give the default engine's package", async () => {
    await assertProperty(
      fc.asyncProperty(
        caseArbitrary(
          [...Object.keys(BODIES), ...Object.keys(MORE_BODIES)] as Body[],
          CELL_OPERATIONS,
        ),
        sameAsDefaultEngine(CELL_OPERATIONS),
      ),
      {
        numRuns: 120,
        id: "cell merges, splits and column changes in the canonical session give the default engine's package",
      },
    );
  });

  test("tables inserted beside paragraphs and tables give the default engine's package", async () => {
    await assertProperty(
      fc.asyncProperty(
        caseArbitrary(
          [
            "a plain table",
            "a vertical merge",
            "a table ending the body",
            ...(Object.keys(PLACEMENT_BODIES) as Body[]),
          ],
          INSERTIONS,
          ["10000001", "10000002"],
        ),
        sameAsDefaultEngine(INSERTIONS),
      ),
      {
        numRuns: 160,
        id: "tables inserted beside paragraphs and tables in the canonical session give the default engine's package",
      },
    );
  });

  for (const mode of ["direct", "tracked-changes"] as const)
    test(`a table after the body's last paragraph leaves the body ending with a paragraph (${mode})`, async () => {
      const source = await packageOf("two tables touching");
      const op = INSERTIONS["insertTable after"]("10000002");
      const legacy = await outcomes(
        (await viaLegacyEngine({ source, operation: op, mode: mode })).bytes,
      );
      const canonical = await viaCanonicalEngine({
        source,
        operation: op,
        mode: mode,
      });
      expect(canonical.refusals).toEqual([]);
      expect(canonical.skipped).toEqual([]);
      const resolved = await outcomes(canonical.bytes);
      expect(resolved.pending.at(-1)).toEqual({ text: [""] });
      expect(resolved.accepted.at(-1)).toEqual({ text: [""] });
      // Only the closing paragraph differs from the default engine's package.
      expect(resolved).not.toEqual(legacy);
      expect(resolved).toEqual(closedOutcomes(legacy));
    });

  test("resolve in the canonical session as the default engine resolves them", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc
          .constantFrom(...(Object.keys(REVIEW_BODIES) as ReviewBody[]))
          .chain((body) =>
            fc.record({
              body: fc.constant(body),
              anchor: fc.constantFrom(...REVIEW_BODIES[body].anchors),
              operation: fc.constantFrom(
                ...(Object.keys(OPERATIONS) as (keyof typeof OPERATIONS)[]),
              ),
            }),
          ),
        async ({ body, anchor, operation }) => {
          const source = await packageOf(body);
          const op = OPERATIONS[operation](anchor);
          const legacy = await viaLegacyEngine({
            source,
            operation: op,
            mode: "tracked-changes",
          });
          const canonical = await viaCanonicalEngine({
            source,
            operation: op,
            mode: "tracked-changes",
          });
          expect(canonical.refusals).toEqual([]);
          expect(canonical.skipped).toEqual(legacy.skipped);
          if (legacy.skipped.length > 0) return;
          const { accepted, rejected } = await outcomes(legacy.bytes);
          expect(await reviewedInSession(canonical.bytes)).toEqual({
            accepted,
            rejected,
          });
        },
      ),
      {
        numRuns: 40,
        id: "table structure suggestions resolve in the canonical session as the default engine resolves them",
      },
    );
  });
});

/**
 * Markup the default engine already wrote resolves in the canonical session
 * as the default engine resolves it: a tracked deletion of a one-column
 * table's column marks every cell deleted, and accepting takes the table
 * (property seed -685065398).
 */
test("the canonical session resolves a one-column table's tracked column deletion the default engine wrote", async () => {
  const source = await packageOf("a deleted break before a paragraph");
  const legacy = await viaLegacyEngine({
    source,
    operation: { id: "op", type: "deleteTableColumn", blockId: "70000006" },
    mode: "tracked-changes",
  });
  expect(legacy.skipped).toEqual([]);
  const resolved = async (session: "legacy" | "canonical") => {
    const results = [];
    for (const decision of ["accept", "reject"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        legacy.bytes.slice(0),
        { session },
      );
      if (decision === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      results.push(await blocksOf(await reviewer.toBuffer()));
    }
    return results;
  };
  expect(await resolved("canonical")).toEqual(await resolved("legacy"));
});
