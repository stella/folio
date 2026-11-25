/**
 * Table nodes an operation inserts: a plain grid of cell texts, or the
 * signature table `insertSignatureTable` places.
 */
import type { Mark, Node as PMNode, Schema } from "prosemirror-model";

import { cleanTextInlineNodes } from "./clean-text";
import { splitCellParagraphTexts } from "./table-row-column-mutations";
import type { TableStructureRevision } from "./table-template";
import type { FolioAISignatureParty } from "./types";

/**
 * Length of the underscore signature rule. Mirrors the constant
 * in `docx-core/legal-source/compile.ts` so the on-screen signature
 * line matches what `create-document` produces.
 */
const SIGNATURE_LINE = "_".repeat(28);

type BuildTableNodeOptions = {
  schema: Schema;
  rows: readonly (readonly string[])[];
  /** Present in tracked mode: every row is stamped as an insertion. */
  revision?: TableStructureRevision;
};

/**
 * A plain table from a grid of cell texts, `null` when the schema has no
 * tables. In tracked mode every row carries `trIns` and every run inside it
 * carries `w:ins`, which is how the format says "this table is new": there is
 * no whole-table insertion element, only rows that were inserted, and a
 * consumer that reads only run-level revisions keeps the text of a rejected
 * insertion when the row marker stands alone.
 */
export const buildTableNode = ({
  schema,
  rows,
  revision,
}: BuildTableNodeOptions): PMNode | null => {
  const paragraphType = schema.nodes["paragraph"];
  const cellType = schema.nodes["tableCell"];
  const rowType = schema.nodes["tableRow"];
  const tableType = schema.nodes["table"];
  if (
    !paragraphType ||
    !cellType ||
    !rowType ||
    !tableType ||
    rows.length === 0
  ) {
    return null;
  }
  const insertionType = schema.marks["insertion"];
  const marks =
    revision && insertionType
      ? [
          insertionType.create({
            revisionId: revision.revisionId,
            author: revision.author,
            date: revision.date,
            ...(revision.initials != null && { initials: revision.initials }),
            ...(revision.provenance != null && {
              provenance: revision.provenance,
            }),
            ...(revision.suggestionId != null && {
              suggestionId: revision.suggestionId,
            }),
          }),
        ]
      : undefined;
  const rowNodes = rows.map((cells) =>
    rowType.create(
      revision ? { trIns: revision } : null,
      cells.map((text) =>
        cellType.create(
          null,
          splitCellParagraphTexts(text).map((line) =>
            paragraphType.create(
              null,
              line.length > 0
                ? cleanTextInlineNodes({ schema, text: line, marks })
                : null,
            ),
          ),
        ),
      ),
    ),
  );
  return tableType.create(null, rowNodes);
};

type BuildSignatureTableNodeOptions = {
  schema: Schema;
  parties: readonly FolioAISignatureParty[];
};

/**
 * Build a borderless PM table mirroring docx-core's
 * `signatureTable` helper: one row, one cell per party, each cell
 * containing party name (bold), two spacer paragraphs, a signature
 * rule, then optional signatory and italic title lines.
 *
 * Returns `null` when the editor schema is missing one of the
 * required node types — callers should surface the op as a skip
 * rather than crashing.
 */
export const buildSignatureTableNode = ({
  schema,
  parties,
}: BuildSignatureTableNodeOptions): PMNode | null => {
  const paragraphType = schema.nodes["paragraph"];
  const cellType = schema.nodes["tableCell"];
  const rowType = schema.nodes["tableRow"];
  const tableType = schema.nodes["table"];
  if (!paragraphType || !cellType || !rowType || !tableType) {
    return null;
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];

  const buildParagraph = (
    text: string,
    options: {
      styleId: string;
      bold?: boolean;
      italic?: boolean;
    },
  ): PMNode => {
    const marks: Mark[] = [];
    if (options.bold && boldType) {
      marks.push(boldType.create());
    }
    if (options.italic && italicType) {
      marks.push(italicType.create());
    }
    const content = text.length > 0 ? schema.text(text, marks) : null;
    return paragraphType.create({ styleId: options.styleId }, content);
  };

  const buildCell = (party: FolioAISignatureParty): PMNode => {
    const cellContent: PMNode[] = [
      buildParagraph(party.name, { styleId: "SignatureParty", bold: true }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph(SIGNATURE_LINE, { styleId: "SignatureRule" }),
    ];
    if (party.signatory && party.signatory.length > 0) {
      cellContent.push(
        buildParagraph(party.signatory, { styleId: "SignatureField" }),
      );
    }
    if (party.title && party.title.length > 0) {
      cellContent.push(
        buildParagraph(party.title, {
          styleId: "SignatureField",
          italic: true,
        }),
      );
    }
    // Cell attrs left to schema defaults — column widths and
    // borders are decided by the table renderer; we just need the
    // structural shape.
    return cellType.create({}, cellContent);
  };

  const cells = parties.map(buildCell);
  const row = rowType.create({}, cells);
  return tableType.create({}, row);
};
