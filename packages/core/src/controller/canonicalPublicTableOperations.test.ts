/**
 * Table, row, column and cell public operations in the canonical session
 * give the package the default engine gives: the same tables, rows, cells
 * and cell texts pending, once every change is accepted, and once every
 * change is rejected, in direct and tracked mode, addressed through any
 * cell's paragraph. A suggestion the default engine cannot record gives the
 * direct edit once accepted and the document as it was once rejected.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
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
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { BlockContent, TableCell } from "../types/document";

setDefaultTimeout(propertyTestTimeout(240_000));

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
      table([cell("80000001", "A1") + cell("80000002", "B1")], 2),
      paragraph("10000002", "After"),
    ],
    anchors: ["80000001", "80000002", "10000002"],
  },
  "a table at the body's end": {
    xml: [paragraph("10000001", "Before"), table([cell("80000003", "A1")], 1)],
    anchors: ["10000001", "80000003"],
  },
  "two tables touching": {
    xml: [
      paragraph("10000001", "Before"),
      table([cell("80000004", "A1")], 1),
      table([cell("80000005", "X1")], 1),
      paragraph("10000002", "After"),
    ],
    anchors: ["80000004", "80000005"],
  },
  "a nested table": {
    xml: [
      table(
        [
          `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${table([cell("80000006", "N1")], 1)}${paragraph("80000007", "Outer")}</w:tc>`,
        ],
        1,
      ),
      paragraph("10000002", "After"),
    ],
    anchors: ["80000006", "80000007", "10000002"],
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

type ReviewBody = keyof typeof REVIEW_BODIES;

const PACKAGE_BODIES = { ...ALL_BODIES, ...REVIEW_BODIES };

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

type Mode = "direct" | "tracked-changes";

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

/** The package pending, with every change accepted, and with every change rejected. */
const outcomes = async (bytes: ArrayBuffer) => {
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
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

const viaDefaultEngine = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0));
  const result = reviewer.applyOperations([operation], { mode });
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    bytes: await reviewer.toBuffer(),
  };
};

const viaCanonicalEngine = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const document = await parseDocx(source.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
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
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    refusals: result.issues.flatMap(({ canonicalRefusal }) =>
      canonicalRefusal === undefined ? [] : [canonicalRefusal],
    ),
    bytes: (
      await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
    ).buffer,
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
    const legacy = await viaDefaultEngine(source, op, mode);
    const canonical = await viaCanonicalEngine(source, op, mode);
    expect(canonical.refusals).toEqual([]);
    // A suggestion the default engine cannot record, such as a split or a
    // column through a merged cell: accepting it gives the direct edit,
    // rejecting it the document as it was.
    if (
      mode === "tracked-changes" &&
      legacy.skipped.length > 0 &&
      canonical.skipped.length === 0
    ) {
      const direct = await viaDefaultEngine(source, op, "direct");
      expect(direct.skipped).toEqual([]);
      const resolved = await outcomes(canonical.bytes);
      expect(resolved.accepted).toEqual(await blocksOf(direct.bytes));
      expect(resolved.rejected).toEqual(await blocksOf(source));
      return;
    }
    expect(canonical.skipped).toEqual(legacy.skipped);
    if (legacy.skipped.length > 0) return;
    expect(await outcomes(canonical.bytes)).toEqual(
      await outcomes(legacy.bytes),
    );
  };

describe("table structure operations in the canonical session", () => {
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
          const legacy = await viaDefaultEngine(source, op, "tracked-changes");
          const canonical = await viaCanonicalEngine(
            source,
            op,
            "tracked-changes",
          );
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
