/**
 * Table, row and column public operations in the canonical session give the
 * package the default engine gives: the same tables, rows, cells and cell
 * texts pending, once every change is accepted, and once every change is
 * rejected, in direct and tracked mode, addressed through any cell's
 * paragraph.
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
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
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
import { CanonicalPublicOperations } from "./canonicalPublicOperations";

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

type Body = keyof typeof BODIES;

const packageOf = async (body: Body) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${BODIES[body].xml.join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

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
  (blockId: string) => FolioDocumentOperation
>;

type Mode = "direct" | "tracked-changes";

const textOf = (blocks: readonly BlockContent[]): string[] =>
  blocks.flatMap((block) =>
    block.type === "paragraph"
      ? [
          block.content
            .flatMap((item) =>
              item.type === "run"
                ? item.content.flatMap((leaf) =>
                    leaf.type === "text" ? [leaf.text] : [],
                  )
                : [],
            )
            .join(""),
        ]
      : [],
  );

const cellShape = (tableCell: TableCell) => ({
  text: textOf(tableCell.content),
  gridSpan: tableCell.formatting?.gridSpan,
  vMerge: tableCell.formatting?.vMerge,
});

/** Each body block: a paragraph's text, a table's rows of cells. */
const blocksOf = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.map((block) =>
    block.type === "table"
      ? { rows: block.rows.map((row) => row.cells.map(cellShape)) }
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

const caseArbitrary = fc
  .constantFrom(...(Object.keys(BODIES) as Body[]))
  .chain((body) =>
    fc.record({
      body: fc.constant(body),
      anchor: fc.constantFrom(...BODIES[body].anchors),
      operation: fc.constantFrom(
        ...(Object.keys(OPERATIONS) as (keyof typeof OPERATIONS)[]),
      ),
      mode: fc.constantFrom<Mode>("direct", "tracked-changes"),
    }),
  );

describe("table structure operations in the canonical session", () => {
  test("give the default engine's package, pending, accepted and rejected", async () => {
    await assertProperty(
      fc.asyncProperty(
        caseArbitrary,
        async ({ body, anchor, operation, mode }) => {
          const source = await packageOf(body);
          const op = OPERATIONS[operation](anchor);
          const legacy = await viaDefaultEngine(source, op, mode);
          const canonical = await viaCanonicalEngine(source, op, mode);
          expect(canonical.refusals).toEqual([]);
          expect(canonical.skipped).toEqual(legacy.skipped);
          if (legacy.skipped.length > 0) return;
          expect(await outcomes(canonical.bytes)).toEqual(
            await outcomes(legacy.bytes),
          );
        },
      ),
      {
        numRuns: 60,
        id: "table structure operations in the canonical session give the default engine's package",
      },
    );
  });
});
