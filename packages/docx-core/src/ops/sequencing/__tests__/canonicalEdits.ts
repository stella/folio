/**
 * Canonical edits for sequencing properties: editor intents compiled the way
 * a canonical session compiles them, and the package operations the session
 * builds directly (notes, styles, sections, headers, review decisions and
 * comment threads, as the canonical comment compiler emits them).
 */
import fc from "fast-check";

import {
  type Comment,
  type Document,
  type Paragraph,
  paragraphNumberingReference,
  type Table,
  type TableRow,
} from "../../../model/document";
import { applyDocumentOps } from "../../apply";
import { storyBody, storyParagraphs } from "../../blocks";
import { freshCommentId } from "../../comments";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
  type EditorIntentMode,
} from "../../editorIntent";
import { formatParaId, type IdLease } from "../../idLease";
import { leafSpans } from "../../leaves";
import { paragraphLength } from "../../offsets";
import { packageResourcesOpOf } from "../../packageResources";
import { allocateCommentAnchorIds } from "../../plan";
import {
  type CommentAnchor,
  type CreateCommentOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type TextPosition,
} from "../../types";

export const EDIT_KINDS = [
  "type",
  "erase",
  "replace",
  "tab",
  "format",
  "paragraphFormat",
  "split",
  "join",
  "insertParagraph",
  "deleteParagraph",
  "paste",
  "list",
  "newList",
  "insertRow",
  "deleteRow",
  "insertColumn",
  "cellProps",
  "tableProps",
  "insertTable",
  "deleteTable",
  "note",
  "deleteNote",
  "resources",
  "section",
  "header",
  "decide",
  "comment",
  "pointComment",
  "reply",
  "editComment",
  "resolveComment",
  "deleteComment",
] as const;
export type EditKind = (typeof EDIT_KINDS)[number];

/** Edits a suggesting author records as tracked changes. */
const SUGGESTIBLE: ReadonlySet<EditKind> = new Set([
  "type",
  "erase",
  "replace",
  "split",
  "insertParagraph",
  "deleteParagraph",
  "insertRow",
  "deleteRow",
]);

export type EditSeed = {
  kind: EditKind;
  a: number;
  b: number;
  suggest: boolean;
  /**
   * Edit the body's first paragraph at offset `a`, so concurrent edits meet
   * in one paragraph, rather than a paragraph and offset both drawn from `a`.
   */
  focus?: true;
};

export const editSeedArbitrary: fc.Arbitrary<EditSeed> = fc.record({
  kind: fc.constantFrom(...EDIT_KINDS),
  a: fc.nat({ max: 100_000 }),
  b: fc.nat({ max: 100_000 }),
  suggest: fc.boolean(),
});

export type EditAuthor = {
  actor: string;
  /** Distinguishes this author's records, such as relationship ids. */
  tag: string;
  lease: IdLease | undefined;
  allocate: ReturnType<typeof createEditorIntentIdAllocator>;
};

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});
const textParagraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [run(text)],
});
const cellRow = (ids: readonly string[], text: string): TableRow => ({
  type: "tableRow",
  cells: ids.map((paraId) => ({
    type: "tableCell",
    content: [textParagraph(paraId, text)],
  })),
});

const DATE = "2026-01-01T00:00:00Z";

/**
 * A body with plain, listed and suggested paragraphs, a table, a list, a
 * footnote and a section.
 */
export const canonicalBase = (): Document => {
  const paragraphs = Array.from({ length: 6 }, (_, index) =>
    textParagraph(formatParaId(0x1000_0000 + index), `Paragraph ${index} text`),
  );
  const listed: Paragraph = {
    ...textParagraph(formatParaId(0x1000_0010), "Listed item"),
    formatting: { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) },
  };
  const suggested: Paragraph = {
    type: "paragraph",
    paraId: formatParaId(0x1000_0011),
    content: [
      run("Kept "),
      {
        type: "insertion",
        info: { id: 10, author: "base", date: DATE },
        content: [run("suggested")],
      },
      run(" tail"),
    ],
  };
  const table: Table = {
    type: "table",
    columnWidths: [2400, 2400],
    rows: [0, 1].map((row) =>
      cellRow(
        [0, 1].map((cell) => formatParaId(0x2000_0000 + row * 16 + cell)),
        `Cell ${row}`,
      ),
    ),
  };
  const document: Document = {
    package: {
      // Fixed dates: a save stamps the current time on a package without them.
      properties: {
        created: new Date(Date.UTC(2026, 0, 1)),
        modified: new Date(Date.UTC(2026, 0, 1)),
      },
      document: {
        content: [
          ...paragraphs.slice(0, 3),
          table,
          listed,
          suggested,
          ...paragraphs.slice(3),
        ],
        finalSectionProperties: { pageWidth: 12_240, pageHeight: 15_840 },
      },
      numbering: {
        abstractNums: [
          {
            abstractNumId: 1,
            levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
          },
        ],
        nums: [{ numId: 1, abstractNumId: 1 }],
      },
      styles: { styles: [{ styleId: "Normal", type: "paragraph" }] },
    },
  };
  // The note goes in as the canonical note intent puts it there, at the end
  // of the last paragraph.
  const last = paragraphs.at(-1);
  const noted = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.ADD_NOTE,
      at: {
        story: OP_STORIES.MAIN,
        blockId: last?.paraId ?? "",
        offset: last === undefined ? 0 : paragraphLength(last),
      },
      note: {
        type: "footnote",
        id: 1,
        content: [textParagraph(formatParaId(0x1000_0020), "Base note")],
      },
    },
  ]);
  if (noted.isErr()) throw noted.error;
  return noted.value.document;
};

const pick = <Value>(
  values: readonly Value[],
  seed: number,
): Value | undefined =>
  values.length === 0 ? undefined : values[seed % values.length];

const nextParagraphId = (lease: IdLease | undefined, seed: number): string =>
  formatParaId(
    lease?.blocks.find(({ kind }) => kind === "paragraph")?.start ??
      0x5000_0000 + seed,
  );

const revisionIds = (document: Document): number[] => {
  const ids: number[] = [];
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const child of node) visit(child);
      return;
    }
    if (typeof node !== "object" || node === null) return;
    if (
      "type" in node &&
      (node.type === "insertion" || node.type === "deletion") &&
      "info" in node &&
      typeof node.info === "object" &&
      node.info !== null &&
      "id" in node.info &&
      typeof node.info.id === "number"
    )
      ids.push(node.info.id);
    for (const child of Object.values(node)) visit(child);
  };
  visit(document.package.document.content);
  return ids;
};

// The note references of the main story, where a note deletion names them.
const noteReferences = (
  document: Document,
): { at: TextPosition; id: number }[] =>
  storyParagraphs(storyBody(document, OP_STORIES.MAIN)).flatMap(
    ({ paragraph }) =>
      leafSpans(paragraph.content).flatMap(({ node, before }) =>
        node.type === "footnoteRef" && paragraph.paraId !== undefined
          ? [
              {
                at: {
                  story: OP_STORIES.MAIN,
                  blockId: paragraph.paraId,
                  offset: before.offset,
                },
                id: node.id,
              },
            ]
          : [],
      ),
  );

// The root of the thread a comment belongs to.
const threadRoot = (
  comments: readonly Comment[],
  comment: Comment,
): Comment => {
  const byId = new Map(comments.map((record) => [record.id, record]));
  let current = comment;
  const seen = new Set<number>();
  while (current.parentId !== undefined && !seen.has(current.id)) {
    seen.add(current.id);
    const parent = byId.get(current.parentId);
    if (parent === undefined) break;
    current = parent;
  }
  return current;
};

/**
 * Compile one canonical edit against `document`, or undefined when the seed
 * names an edit this document cannot take.
 */
export const compileCanonicalEdit = (
  document: Document,
  author: EditAuthor,
  seed: EditSeed,
  serial: number,
): readonly DocumentOp[] | undefined => {
  const body = storyBody(document, OP_STORIES.MAIN);
  const locations = storyParagraphs(body).filter(
    ({ paragraph }) => paragraph.paraId !== undefined,
  );
  const location =
    seed.focus === true ? locations.at(0) : pick(locations, seed.a);
  const paraId = location?.paragraph.paraId;
  if (location === undefined || paraId === undefined) return undefined;
  const length = paragraphLength(location.paragraph);
  const offset =
    seed.focus === true
      ? seed.a % (length + 1)
      : Math.floor(seed.a / 7) % (length + 1);
  const at = (point: number): TextPosition => ({
    story: OP_STORIES.MAIN,
    blockId: paraId,
    offset: point,
  });
  const end = Math.min(length, offset + 1 + (seed.b % 3));
  const topLevel = locations.filter(({ list }) => list.length === 0);
  const inTable = locations.filter(({ list }) =>
    list.some((step) => step.kind === "tableCell"),
  );
  const tableParagraph = pick(inTable, seed.b)?.paragraph.paraId;
  const topParagraph = pick(topLevel, seed.b)?.paragraph.paraId;
  const tableOf = (blockId: string) => {
    const step = inTable
      .find(({ paragraph }) => paragraph.paraId === blockId)
      ?.list.at(0);
    const block = step === undefined ? undefined : body.content[step.block];
    return block?.type === "table" ? block : undefined;
  };
  const direct = (op: DocumentOp): readonly DocumentOp[] => [op];
  // Comment commands apply as the canonical comment compiler checks them.
  const checked = (
    ops: readonly DocumentOp[],
  ): readonly DocumentOp[] | undefined =>
    applyDocumentOps(document, ops).isOk() ? ops : undefined;
  const commentBody = (text: string): Paragraph[] => [
    textParagraph(nextParagraphId(author.lease, serial), text),
  ];
  const comments = document.package.document.comments ?? [];
  const picked = pick(comments, seed.b);
  // Authored by the actor at a time of its own: no receiver stamps either.
  const stamp = new Date(Date.UTC(2026, 0, 2) + serial * 1000).toISOString();
  const createComment = (
    anchor: CommentAnchor,
  ): readonly DocumentOp[] | undefined => {
    const id = freshCommentId(document, author.lease);
    if (id.isErr()) return undefined;
    const op: CreateCommentOp = {
      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      comment: {
        id: id.value,
        author: author.actor,
        date: stamp,
        content: commentBody(`Comment ${serial}`),
      },
      anchor,
    };
    const newIds = allocateCommentAnchorIds(document, op, author.lease);
    return newIds.isOk()
      ? checked([{ ...op, newIds: newIds.value }])
      : undefined;
  };
  let intent: EditorIntent | undefined;
  switch (seed.kind) {
    case "type":
      intent = {
        type: "replaceText",
        from: at(offset),
        to: at(offset),
        text: seed.b % 2 === 0 ? "x" : "yz",
      };
      break;
    case "erase":
    case "replace":
      if (offset >= length) return undefined;
      intent = {
        type: "replaceText",
        from: at(offset),
        to: at(end),
        text: seed.kind === "erase" ? "" : "q",
      };
      break;
    case "tab":
      intent = {
        type: "insertAtom",
        from: at(offset),
        to: at(offset),
        atom: { type: "tab" },
      };
      break;
    case "format":
      if (offset >= length) return undefined;
      intent = {
        type: "formatRun",
        from: at(offset),
        to: at(end),
        patch: [
          { bold: true },
          { italic: true },
          { strike: true },
          { fontSize: 28 },
        ][seed.b % 4] ?? { bold: true },
      };
      break;
    case "paragraphFormat":
      intent = {
        type: "formatParagraph",
        at: at(0),
        patch: [
          { alignment: "center" as const },
          { keepNext: true },
          { indentLeft: 720 },
          { spaceAfter: 120 },
        ][seed.b % 4] ?? { keepNext: true },
      };
      break;
    case "split":
      intent = { type: "splitParagraph", at: at(offset), newBlockId: "" };
      break;
    case "join": {
      const index = topLevel.findIndex(
        ({ paragraph }) => paragraph.paraId === paraId,
      );
      const next = topLevel.at(index + 1);
      if (
        index < 0 ||
        next?.paragraph.paraId === undefined ||
        next.index !== (topLevel[index]?.index ?? -2) + 1
      )
        return undefined;
      intent = {
        type: "joinParagraphs",
        story: OP_STORIES.MAIN,
        blockId: paraId,
        nextBlockId: next.paragraph.paraId,
      };
      break;
    }
    case "insertParagraph":
      intent = {
        type: "insertParagraphs",
        story: OP_STORIES.MAIN,
        at: { type: seed.b % 2 === 0 ? "after" : "before", blockId: paraId },
        paragraphs: [{ type: "paragraph", content: [run(`New ${serial}`)] }],
      };
      break;
    case "deleteParagraph":
      if (topParagraph === undefined || topLevel.length < 3) return undefined;
      intent = {
        type: "deleteParagraphs",
        story: OP_STORIES.MAIN,
        blockIds: [topParagraph],
      };
      break;
    case "paste":
      intent = {
        type: "replaceFragment",
        from: at(offset),
        to: at(offset),
        paragraphs: [
          { type: "paragraph", content: [run("Pasted ")] },
          { type: "paragraph", content: [run("lines")] },
        ],
        openStart: 1,
        openEnd: 1,
      };
      break;
    case "list":
      intent = {
        type: "setList",
        items: [{ at: at(0), ilvl: 0 }],
        target: { type: "existing", numId: 1 },
      };
      break;
    case "newList": {
      const numIds = document.package.numbering?.nums.map(({ numId }) => numId);
      intent = {
        type: "setList",
        items: [{ at: at(0), ilvl: 0 }],
        target: {
          type: "new",
          num: { numId: Math.max(0, ...(numIds ?? [])) + 1, abstractNumId: 1 },
        },
      };
      break;
    }
    case "insertRow":
    case "deleteRow": {
      if (tableParagraph === undefined) return undefined;
      const table = tableOf(tableParagraph);
      const row = table?.rows.at(0);
      if (table === undefined || row === undefined) return undefined;
      if (seed.kind === "deleteRow" && table.rows.length < 2) return undefined;
      intent = {
        type: "tableRow",
        operation:
          seed.kind === "insertRow"
            ? {
                type: DOCUMENT_OP_TYPES.INSERT_ROW,
                story: OP_STORIES.MAIN,
                blockId: tableParagraph,
                at: seed.a % (table.rows.length + 1),
                row: cellRow(
                  row.cells.map(() => "00000000"),
                  `Row ${serial}`,
                ),
              }
            : {
                type: DOCUMENT_OP_TYPES.DELETE_ROW,
                story: OP_STORIES.MAIN,
                blockId: tableParagraph,
              },
      };
      break;
    }
    case "insertColumn":
      if (tableParagraph === undefined) return undefined;
      intent = {
        type: "table",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          story: OP_STORIES.MAIN,
          blockId: tableParagraph,
          column: seed.a % 3,
          width: 1200,
        },
      };
      break;
    case "cellProps":
      if (tableParagraph === undefined) return undefined;
      intent = {
        type: "table",
        operation: {
          type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
          story: OP_STORIES.MAIN,
          blockId: tableParagraph,
          patch: { verticalAlign: "center" },
        },
      };
      break;
    case "tableProps":
      if (tableParagraph === undefined) return undefined;
      intent = {
        type: "table",
        operation: {
          type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
          story: OP_STORIES.MAIN,
          blockId: tableParagraph,
          patch: { justification: "center" },
        },
      };
      break;
    case "insertTable":
      if (topParagraph === undefined) return undefined;
      intent = {
        type: "tableBlock",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story: OP_STORIES.MAIN,
          blockId: topParagraph,
          table: {
            type: "table",
            columnWidths: [2400, 2400],
            rows: [cellRow(["00000000", "00000000"], `Table ${serial}`)],
          },
        },
      };
      break;
    case "deleteTable":
      if (tableParagraph === undefined) return undefined;
      intent = {
        type: "tableBlock",
        operation: {
          type: DOCUMENT_OP_TYPES.DELETE_TABLE,
          story: OP_STORIES.MAIN,
          blockId: tableParagraph,
        },
      };
      break;
    case "note": {
      const ids = (document.package.footnotes ?? []).map(({ id }) => id);
      return direct({
        type: DOCUMENT_OP_TYPES.ADD_NOTE,
        at: at(offset),
        note: {
          type: "footnote",
          id: Math.max(0, ...ids) + 1,
          content: [
            textParagraph(nextParagraphId(author.lease, serial), "Note"),
          ],
        },
      });
    }
    case "deleteNote": {
      const reference = pick(noteReferences(document), seed.b);
      if (reference === undefined) return undefined;
      return direct({
        type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
        at: reference.at,
        story: { kind: "footnote", id: reference.id },
      });
    }
    case "comment":
      if (offset >= length) return undefined;
      return createComment({ kind: "range", from: at(offset), to: at(end) });
    case "pointComment":
      return createComment({ kind: "point", at: at(offset) });
    case "reply":
      if (picked === undefined) return undefined;
      return createComment({
        kind: "reply",
        parentId: threadRoot(comments, picked).id,
      });
    case "editComment":
      if (picked === undefined) return undefined;
      return checked([
        {
          type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
          id: picked.id,
          content: commentBody(`Edit ${serial}`),
        },
      ]);
    case "resolveComment":
      if (picked === undefined) return undefined;
      // Resolving a resolved comment reopens it.
      return checked([
        {
          type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
          id: picked.id,
          status: picked.done === true ? "open" : "resolved",
        },
      ]);
    case "deleteComment":
      if (picked === undefined) return undefined;
      return checked([
        {
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: picked.id,
          scope:
            picked.parentId !== undefined &&
            comments.some(({ id }) => id === picked.parentId)
              ? "reply"
              : "thread",
        },
      ]);
    case "resources": {
      const styles = document.package.styles ?? { styles: [] };
      return direct(
        packageResourcesOpOf({
          before: document,
          after: {
            ...document,
            package: {
              ...document.package,
              styles: {
                ...styles,
                styles: [
                  ...styles.styles,
                  {
                    styleId: `Pasted${author.tag}x${serial}`,
                    type: "paragraph",
                  },
                ],
              },
            },
          },
        }),
      );
    }
    case "section":
      return direct({
        type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
        sectionIndex: 0,
        patch: { marginTop: 1000 + serial },
      });
    case "header":
      return direct({
        type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
        sectionIndex: 0,
        story: { kind: "header", rId: `rIdHeader${author.tag}x${serial}` },
        referenceType:
          (["default", "first", "even"] as const)[seed.b % 3] ?? "default",
        // An empty header paragraph has no run, as the header intent makes it.
        content: [
          {
            type: "paragraph",
            paraId: nextParagraphId(author.lease, serial),
            content: [],
          },
        ],
      });
    case "decide": {
      const id = pick(revisionIds(document), seed.b);
      if (id === undefined) return undefined;
      return direct({
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [id],
        decision: seed.b % 2 === 0 ? "accept" : "reject",
      });
    }
    default: {
      const exhaustive: never = seed.kind;
      return exhaustive;
    }
  }
  const request =
    intent.type === "splitParagraph"
      ? { type: intent.type, at: intent.at }
      : intent;
  const allocation = author.allocate(document, request, author.lease);
  const compiled =
    intent.type === "splitParagraph"
      ? { ...intent, newBlockId: allocation.newBlockId }
      : intent;
  const leased = author.lease === undefined ? {} : { idLease: author.lease };
  const mode: EditorIntentMode =
    seed.suggest && SUGGESTIBLE.has(seed.kind)
      ? {
          type: "suggesting",
          revision: {
            id: allocation.revisionId,
            author: author.actor,
            // Distinct per edit, so no new record merges into an earlier one.
            date: new Date(Date.UTC(2026, 0, 1) + serial * 1000).toISOString(),
          },
          newIds: allocation.newIds,
          ...leased,
        }
      : { type: "editing", newIds: allocation.newIds, ...leased };
  const result = compileEditorIntent(document, { intent: compiled, mode });
  return result.isOk() ? result.value.ops : undefined;
};
