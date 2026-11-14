/**
 * Table operations in both engines, read back alike: a body's paragraphs by
 * their text, its tables by their rows of cells, nested tables included, the
 * package pending, with every change accepted and with every change rejected.
 */
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import type { FolioDocumentOperation } from "../../ai-edits";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { getTrackedChangeStatsFromDoc } from "../../ai-edits/read";
import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../../canonical/canonicalSession";
import { serializeCanonicalSave } from "../../docx/canonicalSave";
import { parseDocx } from "../../docx/parser";
import { createEmptyDocx } from "../../docx/rezip";
import { schema } from "../../prosemirror/schema";
import type { BlockContent, TableCell } from "../../types/document";
import { CanonicalPublicOperations } from "../canonicalPublicOperations";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REVISION = 'w:author="Earlier" w:date="2026-01-02T03:04:05Z"';

export const paragraph = (id: string, text: string, inner = "") =>
  `<w:p w14:paraId="${id}">${inner}${text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ""}</w:p>`;
/** A paragraph whose words a pending change deletes. */
export const deletedParagraph = (id: string, text: string, revision: number) =>
  `<w:p w14:paraId="${id}"><w:pPr><w:rPr><w:del w:id="${revision + 1}" ${REVISION}/></w:rPr></w:pPr><w:del w:id="${revision}" ${REVISION}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del></w:p>`;
export const cell = (content: string, properties = "") =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;
export const textCell = (id: string, text: string, properties = "") =>
  cell(paragraph(id, text), properties);
export const row = (cells: string, properties = "") =>
  `<w:tr>${properties ? `<w:trPr>${properties}</w:trPr>` : ""}${cells}</w:tr>`;
/** A row a pending change deletes, its words with it. */
export const deletedRow = (cells: string, revision: number) =>
  row(cells, `<w:del w:id="${revision}" ${REVISION}/>`);
export const table = (rows: readonly string[], columns: number) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>${rows.join("")}</w:tbl>`;

export const packageOf = async (blocks: readonly string[]) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${blocks.join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

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
  if (
    type === "run" ||
    type === "insertion" ||
    type === "deletion" ||
    type === "moveFrom" ||
    type === "moveTo" ||
    type === "inlineSdt"
  )
    return wordsOf(Reflect.get(value, "content"));
  if (type === "hyperlink") return wordsOf(Reflect.get(value, "children"));
  return [];
};

type CellShape = {
  blocks: BlockShape[];
  gridSpan: number | undefined;
  vMerge: "restart" | "continue" | undefined;
};
type BlockShape = { text: string } | { rows: CellShape[][] };

const blockShapes = (blocks: readonly BlockContent[]): BlockShape[] =>
  blocks.flatMap((block): BlockShape[] => {
    switch (block.type) {
      case "paragraph":
        return [{ text: wordsOf(block.content).join("") }];
      case "table":
        return [
          {
            rows: withoutContinuationRows(
              block.rows.map((tableRow) => tableRow.cells.map(cellShape)),
            ),
          },
        ];
      case "blockSdt":
        return blockShapes(block.content);
      default:
        return [];
    }
  });

const cellShape = (tableCell: TableCell): CellShape => ({
  blocks: blockShapes(tableCell.content),
  gridSpan: tableCell.formatting?.gridSpan,
  vMerge: tableCell.formatting?.vMerge,
});

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
    (cells) =>
      !cells.every(
        (entry) =>
          entry.vMerge === "continue" &&
          entry.blocks.every((block) => "text" in block && block.text === ""),
      ),
  );
  return kept.map((cells, rowIndex) =>
    cells.map((entry, index) =>
      entry.vMerge === "restart" &&
      kept[rowIndex + 1]?.[index]?.vMerge !== "continue"
        ? { ...entry, vMerge: undefined }
        : entry,
    ),
  );
};

/** Each body block: a paragraph's text, a table's rows of cells. */
export const blocksOf = async (bytes: ArrayBuffer): Promise<BlockShape[]> => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return blockShapes(document.package.document.content);
};

/** The package pending, with every change accepted, and with every change rejected. */
export const outcomes = async (bytes: ArrayBuffer) => {
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
export const reviewedInSession = async (bytes: ArrayBuffer) => {
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
      if (commit.publish().isErr())
        panic(`The session could not publish the ${resolution}ed changes.`);
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

export type Mode = "direct" | "tracked-changes";

export const viaDefaultEngine = async (
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

export const viaCanonicalEngine = async (
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
