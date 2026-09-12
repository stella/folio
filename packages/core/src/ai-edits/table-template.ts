/**
 * Building a table from another document's table rather than from a grid of
 * strings.
 *
 * `insertTable` and `insertTableRow` describe their content as cell texts,
 * which is the right shape for a caller that writes a table from nothing. A
 * comparison is not that caller: the table it adds already exists in the
 * document it is comparing against, complete with its `w:tblPr`, its
 * `w:tblGrid` widths, per-row `w:trPr`, per-cell `w:tcPr` — spans, vertical
 * merges, shading, borders, margins, alignment — and cell paragraphs carrying
 * their own properties. Rebuilding that from text loses every one of them, and
 * a consumer accepting the redline then gets a table that is not the one it
 * was compared to.
 *
 * So the operation may be handed the node it should place. The copy is
 * structural and keeps everything by default; what it cannot bring is what
 * only the package it came from can resolve — a relationship id (a drawing, a
 * hyperlink), a note or comment id, a bookmark, a paragraph identity, or a
 * tracked change belonging to the other document's history. Everything that
 * describes the table itself travels.
 */

import { Fragment, type Mark, type Node as PMNode, type Schema } from "prosemirror-model";

import {
  cloneTableCellsWithParagraphPropertyCaptures,
  decodeTableCellParagraphSourcePayload,
  recreateProseNodeWithDetachedParagraphPropertySource,
} from "../docx/paragraphPropertySource";
import { expectTableCellAttrs } from "../prosemirror/attrs";
import type { TrackedChangeProvenance } from "../prosemirror/schema/marks";
import { stripBlockIdentityAttrs } from "./block-identity";

/** The node an operation should place, by the id of the operation placing it. */
export type FolioTableTemplates = ReadonlyMap<string, PMNode>;

/**
 * Nodes that name something the other package owns: a drawing's relationship,
 * an anchored frame's, a bookmark's id. They resolve to nothing once the table
 * has crossed, so they are dropped rather than copied into a dangling
 * reference.
 */
const PACKAGE_BOUND_NODE_NAMES: ReadonlySet<string> = new Set([
  "image",
  "textBox",
  "textBoxAnchor",
  "shape",
  "bookmarkBoundary",
]);

/**
 * Marks that name something the other package owns. Run properties are the
 * point of copying the runs at all, so the denial is narrow: a relationship, a
 * note, a comment, and the other document's revisions.
 */
const PACKAGE_BOUND_MARK_NAMES: ReadonlySet<string> = new Set([
  "hyperlink",
  "footnoteRef",
  "comment",
  "insertion",
  "deletion",
  "runPropertyChange",
]);

/**
 * Paragraph attrs cleared on a copied paragraph: its identity, the bookmarks
 * and empty hyperlinks that name package-scoped ids, the section it ended, and
 * the tracked-change history of the document it came from. Its formatting —
 * style, alignment, spacing, indentation, borders, shading, tabs — is what the
 * copy exists to carry, and stays.
 */
const CLEARED_PARAGRAPH_ATTRS = [
  "bookmarks",
  "_emptyHyperlinks",
  "_propertyChanges",
  "pPrMark",
  "_suggestedInsert",
  "_sectionProperties",
  "sectionBreakType",
  "renderedPageBreakBefore",
] as const;

/** Table, row and cell attrs that record the other document's revisions. */
const CLEARED_TABLE_ATTRS = ["tblPrChange", "_suggestedInsert"] as const;
const CLEARED_ROW_ATTRS = ["trIns", "trDel", "trPrChange"] as const;
const CLEARED_CELL_ATTRS = ["cellMarker", "tcPrChange"] as const;

const hasAuthoredValue = (value: unknown): boolean => {
  if (value === null || value === undefined || value === false || value === "") {
    return false;
  }
  return !Array.isArray(value) || value.length > 0;
};

const clearedAttrsOf = (node: PMNode): readonly string[] => {
  switch (node.type.spec["tableRole"]) {
    case "table":
      return CLEARED_TABLE_ATTRS;
    case "row":
      return CLEARED_ROW_ATTRS;
    case "cell":
    case "header_cell":
      return CLEARED_CELL_ATTRS;
    default:
      return node.isTextblock ? CLEARED_PARAGRAPH_ATTRS : [];
  }
};

const losesAuthoredAttrs = (node: PMNode): boolean =>
  clearedAttrsOf(node).some((name) => hasAuthoredValue(node.attrs[name]));

/**
 * Whether a target table can cross into the base package without semantic
 * content being stripped from its template.
 *
 * Paragraph identities are deliberately excluded: they are regenerated in
 * the receiving document. Package-owned relationships, annotations,
 * bookmarks and unresolved revision metadata are not portable and therefore
 * make a replacement unsafe. Callers can retain their granular plan instead.
 */
export const tableTemplateCanCrossPackageLosslessly = (template: PMNode): boolean => {
  let portable = true;
  const inspect = (node: PMNode): boolean => {
    if (
      PACKAGE_BOUND_NODE_NAMES.has(node.type.name) ||
      node.marks.some(({ type }) => PACKAGE_BOUND_MARK_NAMES.has(type.name)) ||
      losesAuthoredAttrs(node)
    ) {
      portable = false;
      return false;
    }
    return true;
  };
  if (!inspect(template)) {
    return false;
  }
  template.descendants(inspect);
  return portable;
};

const withoutAttrs = (
  attrs: Record<string, unknown>,
  cleared: readonly string[],
): Record<string, unknown> => {
  const next = { ...attrs };
  for (const key of cleared) {
    next[key] = null;
  }
  return next;
};

/**
 * A structural tracked change on a table row.
 *
 * A whole-table insertion has no element of its own in the format: every row
 * carries `w:trPr/w:ins`, and every run inside it carries `w:ins` as well. A
 * consumer that reads only run-level revisions keeps the text of a rejected
 * insertion when the row marker stands alone, so the two are always written
 * together.
 */
export type TableStructureRevision = {
  revisionId: number;
  author: string;
  date: string;
  /** Optional author initials (w:initials), carried for round-trip. */
  initials?: string;
  /**
   * `"suggested"` marks the produced `trIns`/`trDel`/`cellMarker` as an AI
   * proposal that is stripped from serialized DOCX until accepted.
   */
  provenance?: TrackedChangeProvenance;
  suggestionId?: string;
};

type TemplateContext = {
  /** Present when the copy is an insertion; absent when it only supplies shape. */
  revision: TableStructureRevision | null;
  insertion: Mark | null;
  /**
   * Clamp a vertical merge on the copied row's own cells. A cell that merged
   * downwards in the document it came from reached rows that are not the
   * receiving table's.
   */
  clampRowSpan: boolean;
};

const insertionMarkOf = (schema: Schema, revision: TableStructureRevision | null): Mark | null => {
  const markType = schema.marks["insertion"];
  if (!revision || !markType) {
    return null;
  }
  const { revisionId, author, date, initials, provenance, suggestionId } = revision;
  return markType.create({
    revisionId,
    author,
    date,
    ...(initials !== undefined && { initials }),
    ...(provenance !== undefined && { provenance }),
    ...(suggestionId !== undefined && { suggestionId }),
  });
};

const copiedAttrs = (node: PMNode, context: TemplateContext): Record<string, unknown> => {
  const sourceAttrs = node.isTextblock ? stripBlockIdentityAttrs(node.attrs) : node.attrs;
  const attrs = withoutAttrs(sourceAttrs, clearedAttrsOf(node));
  switch (node.type.spec["tableRole"]) {
    case "table":
      return attrs;
    case "row":
      return context.revision ? { ...attrs, trIns: context.revision } : attrs;
    case "cell":
    case "header_cell": {
      if (context.clampRowSpan) {
        return {
          ...attrs,
          rowspan: 1,
          _docxVMergeContinuationCells: null,
        };
      }
      const continuationCells = expectTableCellAttrs(node)._docxVMergeContinuationCells;
      return continuationCells !== undefined && continuationCells !== null
        ? {
            ...attrs,
            _docxVMergeContinuationCells: cloneTableCellsWithParagraphPropertyCaptures(
              decodeTableCellParagraphSourcePayload(continuationCells),
            ),
          }
        : attrs;
    }
    default:
      return attrs;
  }
};

/**
 * One node of the template, or `null` when it names something the package it
 * came from owns. Inline content carries the insertion mark; a nested table is
 * copied by the same rules as the one holding it.
 */
const copyNode = (node: PMNode, context: TemplateContext): PMNode | null => {
  if (PACKAGE_BOUND_NODE_NAMES.has(node.type.name)) {
    return null;
  }
  const marks = node.marks.filter(({ type }) => !PACKAGE_BOUND_MARK_NAMES.has(type.name));
  if (node.isLeaf) {
    // A revision belongs on the run, which is a leaf. Putting one on an inline
    // container as well would mark the same characters twice, and resolving
    // the revision would then delete the container's range and its children's.
    return node.mark(node.isInline && context.insertion ? [...marks, context.insertion] : marks);
  }
  // A cell's rowspan is only clamped on the row being placed. A nested table's
  // rows are its own, and their merges reach rows that travel with them.
  const inner =
    node.type.spec["tableRole"] === "table" ? { ...context, clampRowSpan: false } : context;
  const content: PMNode[] = [];
  node.forEach((child) => {
    const copied = copyNode(child, inner);
    if (copied) {
      content.push(copied);
    }
  });
  return recreateProseNodeWithDetachedParagraphPropertySource(node, {
    attrs: copiedAttrs(node, context),
    content: Fragment.fromArray(content),
    marks,
  });
};

type TableFromTemplateOptions = {
  schema: Schema;
  template: PMNode;
  /** Present in tracked mode: the whole table is stamped as one insertion. */
  revision?: TableStructureRevision;
};

/**
 * The template table, ready to insert: its properties, grid, rows and cells
 * carried verbatim, nested tables included, and every row plus every run
 * stamped as an insertion when the caller is tracking changes.
 *
 * `null` when the template is not a table or holds no rows, so the caller can
 * fall back to the operation's own cell texts rather than place something that
 * is not one.
 */
export const tableFromTemplate = ({
  schema,
  template,
  revision,
}: TableFromTemplateOptions): PMNode | null => {
  if (template.type.spec["tableRole"] !== "table" || template.childCount === 0) {
    return null;
  }
  return copyNode(template, {
    revision: revision ?? null,
    insertion: insertionMarkOf(schema, revision ?? null),
    clampRowSpan: false,
  });
};

type TableRowFromTemplateOptions = {
  template: PMNode;
  /** Grid columns the row has to occupy, from the table receiving it. */
  columnCount: number;
};

/**
 * The template row, ready to insert into an existing table.
 *
 * Refused — `null`, so the caller falls back to the operation's cell texts —
 * when the row's cells do not span exactly the receiving table's grid: a row
 * of a different width would leave the table's map inconsistent with its own
 * grid. A cell that merged vertically in the document it came from is placed
 * unmerged, because the rows it reached into are not this table's.
 */
export const tableRowFromTemplate = ({
  template,
  columnCount,
}: TableRowFromTemplateOptions): PMNode | null => {
  if (template.type.spec["tableRole"] !== "row" || template.childCount === 0) {
    return null;
  }
  let spanned = 0;
  template.forEach((cell) => {
    const colspan: unknown = cell.attrs["colspan"];
    spanned += typeof colspan === "number" ? colspan : 1;
  });
  if (spanned !== columnCount) {
    return null;
  }
  // The caller stamps the row's own insertion and marks its runs, so the copy
  // carries neither.
  return copyNode(template, { revision: null, insertion: null, clampRowSpan: true });
};
