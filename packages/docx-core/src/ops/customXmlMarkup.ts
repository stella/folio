/**
 * A block custom XML element's own tags, changed without touching what they
 * wrap: blocks wrapped in an element, an element's tags taken away with its
 * content left standing, or an element's own tracked records set.
 *
 * The tags are markup, not content (ECMA-376 Part 1, §17.5.1), so their
 * insertion and deletion are recorded as ranges around the element
 * (`w:customXmlInsRangeStart` and its kin, §17.13.5.4–7) rather than by a
 * `w:ins` or `w:del` holding it. This operation is the one that writes those
 * records and resolves them.
 */

import { Result } from "better-result";

import type {
  BlockContent,
  BlockCustomXml,
  Document,
  Paragraph,
  TableCell,
  TableRow,
} from "../model/document";
import {
  blockListAt,
  storyParagraphs,
  updateBlockList,
  withBodyContent,
} from "./blocks";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { idKey } from "./ids";
import { cloneModel } from "./modelClone";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import { replaceStoryBody, storyBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type OpStory,
  REVISION_DECISIONS,
  type RevisionDecision,
  type SetCustomXmlMarkupOp,
} from "./types";

type Applied = Result<DocumentEdit, DocumentOpRefusal>;

const refuse = (
  op: SetCustomXmlMarkupOp,
  reason: DocumentOpRefusalReason,
  message: string,
): Applied =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));

const PROPERTIES_TAG = /^<(?:[^\s<>/:]+:)?customXmlPr[\s/>]/u;

/**
 * An element's `w:customXmlPr`, which the reader keeps as a captured child of
 * the element: part of its start tag's markup, not of what it wraps.
 */
export const isCustomXmlProperties = (block: BlockContent): boolean =>
  block.type === "preservedBlock" && PROPERTIES_TAG.test(block.xml);

/**
 * `blocks` read with every custom XML element's tags taken away, in content
 * controls and table cells too: what the tags wrap, in order.
 */
export const withoutCustomXmlTags = (
  blocks: readonly BlockContent[],
): BlockContent[] =>
  blocks.flatMap((block): BlockContent[] => {
    switch (block.type) {
      case "blockCustomXml":
        return withoutCustomXmlTags(
          block.content.filter((child) => !isCustomXmlProperties(child)),
        );
      case "blockSdt":
        return [{ ...block, content: withoutCustomXmlTags(block.content) }];
      case "table":
        return [
          {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({
                ...cell,
                content: withoutCustomXmlTags(cell.content),
              })),
            })),
          },
        ];
      default:
        return [block];
    }
  });

/**
 * How deep in custom XML elements each paragraph ending a section stands,
 * outside tables: a section ends in the list holding it, so no change of the
 * tags may move one.
 */
const sectionEndDepths = (
  blocks: readonly BlockContent[],
  depth = 0,
): number[] =>
  blocks.flatMap((block): number[] => {
    if (block.type === "paragraph")
      return block.sectionProperties === undefined ? [] : [depth];
    if (block.type === "blockCustomXml")
      return sectionEndDepths(block.content, depth + 1);
    if (block.type === "blockSdt")
      return sectionEndDepths(block.content, depth);
    return [];
  });

/** Every paragraph a block list holds, at any depth, in order. */
const paragraphsIn = (blocks: readonly BlockContent[]): Paragraph[] =>
  storyParagraphs({ content: [...blocks] }).map(({ paragraph }) => paragraph);

/** Apply a {@link SetCustomXmlMarkupOp}; its inverse swaps `expected` and `blocks`. */
export const setCustomXmlMarkup = (
  document: Document,
  op: SetCustomXmlMarkupOp,
): Applied => {
  const body = storyBody(document, op.story);
  const anchor = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.blockId),
  );
  if (anchor === undefined)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The paragraph addressing the custom XML markup does not exist.",
    );
  const depth = op.ancestorDepth ?? 0;
  if (!Number.isSafeInteger(depth) || depth < 0 || depth > anchor.list.length)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The addressing paragraph has no such enclosing block list.",
    );
  const list = anchor.list.slice(0, anchor.list.length - depth);
  const before = blockListAt(body.content, list);
  if (!equalForStaleness(before, op.expected))
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The blocks around the custom XML markup have changed.",
    );
  if (
    !structurallyEqual(
      withoutCustomXmlTags(before),
      withoutCustomXmlTags(op.blocks),
    )
  )
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "Custom XML markup changes only the elements' tags and their own records, never what they wrap.",
    );
  // A list a paragraph must end (a story body or a table cell) keeps one.
  if (
    before.at(-1)?.type === "paragraph" &&
    op.blocks.at(-1)?.type !== "paragraph"
  )
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A block list ending in a paragraph keeps it outside every custom XML element.",
    );
  if (!structurallyEqual(sectionEndDepths(before), sectionEndDepths(op.blocks)))
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      "A paragraph ending a section stays in the block list holding it.",
    );
  if (structurallyEqual(before, op.blocks))
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  const blocks = cloneModel([...op.blocks]);
  const content = updateBlockList(body.content, list, () => blocks);
  const edited = withBodyContent(body, content);
  // The addressing paragraph may now stand in or out of an element: the
  // inverse names the list from where it stands after.
  const after = storyParagraphs(edited).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.blockId),
  );
  const afterDepth = (after?.list.length ?? list.length) - list.length;
  const inverse: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_CUSTOM_XML_MARKUP,
    story: op.story,
    blockId: op.blockId,
    ...(afterDepth === 0 ? {} : { ancestorDepth: afterDepth }),
    expected: cloneModel(blocks),
    blocks: cloneModel([...before]),
  };
  return Result.ok({
    document: replaceStoryBody({ document, story: op.story, body: edited }),
    inverse: [inverse],
    touched: {
      modified: paragraphsIn(blocks).flatMap(({ paraId }) =>
        paraId === undefined ? [] : [paraId],
      ),
      inserted: [],
      removed: [],
    },
  });
};

/**
 * `blocks` with no custom XML element's own records, at any depth: what a
 * paste copies is the element, whose earlier records belong to the document
 * it was copied from.
 */
export const withoutCustomXmlRecords = (
  blocks: readonly BlockContent[],
): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "blockCustomXml":
        return withoutElementRecords(block);
      case "blockSdt":
        return { ...block, content: withoutCustomXmlRecords(block.content) };
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: withoutCustomXmlRecords(cell.content),
            })),
          })),
        };
      default:
        return block;
    }
  });

/** One element without its own records, nor any element it holds. */
export const withoutElementRecords = (
  element: BlockCustomXml,
): BlockCustomXml => {
  const { insertion: _insertion, deletion: _deletion, ...rest } = element;
  return { ...rest, content: withoutCustomXmlRecords(element.content) };
};

type MarkupResolution = {
  ids: ReadonlySet<number>;
  decision: RevisionDecision;
  /** The ids an element's record held, which the resolution settled. */
  settled: Set<number>;
};

/**
 * `blocks` with the element records `ids` selects resolved, at any depth.
 * Accepting an insertion of the tags drops its record and rejecting it takes
 * the tags away; accepting a deletion of them takes them away and rejecting
 * it drops its record. Taking the tags away leaves what they wrap, which
 * carries records of its own, so the order the two are resolved in does not
 * change what stays.
 */
const resolveMarkupIn = (
  blocks: readonly BlockContent[],
  resolution: MarkupResolution,
): BlockContent[] =>
  blocks.flatMap((block): BlockContent[] => {
    switch (block.type) {
      case "blockCustomXml": {
        const content = resolveMarkupIn(block.content, resolution);
        const selected = (record: BlockCustomXml["insertion"]) =>
          record !== undefined && resolution.ids.has(record.info.id);
        const inserted = selected(block.insertion);
        const deleted = selected(block.deletion);
        if (inserted && block.insertion !== undefined)
          resolution.settled.add(block.insertion.info.id);
        if (deleted && block.deletion !== undefined)
          resolution.settled.add(block.deletion.info.id);
        const accept = resolution.decision === REVISION_DECISIONS.ACCEPT;
        if ((accept && deleted) || (!accept && inserted))
          return content.filter((child) => !isCustomXmlProperties(child));
        const { insertion, deletion, ...rest } = block;
        return [
          {
            ...rest,
            content,
            ...(insertion === undefined || inserted ? {} : { insertion }),
            ...(deletion === undefined || deleted ? {} : { deletion }),
          },
        ];
      }
      case "blockSdt":
        return [
          { ...block, content: resolveMarkupIn(block.content, resolution) },
        ];
      case "table":
        return [
          {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({
                ...cell,
                content: resolveMarkupIn(cell.content, resolution),
              })),
            })),
          },
        ];
      default:
        return [block];
    }
  });

/**
 * The custom XML elements' own records a resolution selects, settled as one
 * operation on the story's blocks, with the revision ids to resolve after.
 */
export const settleCustomXmlMarkup = ({
  document,
  story,
  revisionIds,
  decision,
}: {
  document: Document;
  story: OpStory;
  revisionIds: readonly number[];
  decision: RevisionDecision;
}): { ops: DocumentOp[]; ids: number[] } => {
  const body = storyBody(document, story);
  const resolution: MarkupResolution = {
    ids: new Set(revisionIds),
    decision,
    settled: new Set(),
  };
  const blocks = resolveMarkupIn(body.content, resolution);
  const anchor = storyParagraphs(body).find(
    ({ paragraph }) => paragraph.paraId !== undefined,
  );
  if (resolution.settled.size === 0 || anchor?.paragraph.paraId === undefined)
    return { ops: [], ids: [...revisionIds] };
  // An id the tags' record shares with other records (the breaks a moved
  // element's removal takes) goes on to resolve those.
  return {
    ops: [
      {
        type: DOCUMENT_OP_TYPES.SET_CUSTOM_XML_MARKUP,
        story,
        blockId: anchor.paragraph.paraId,
        ...(anchor.list.length === 0
          ? {}
          : { ancestorDepth: anchor.list.length }),
        expected: body.content,
        blocks,
      },
    ],
    ids: [...revisionIds],
  };
};

/** Every block list of `blocks`, in content controls and table cells too. */
const mapLists = (
  blocks: readonly BlockContent[],
  map: (list: readonly BlockContent[]) => BlockContent[],
): BlockContent[] =>
  map(
    blocks.map((block): BlockContent => {
      switch (block.type) {
        case "blockSdt":
          return { ...block, content: mapLists(block.content, map) };
        case "blockCustomXml":
          return { ...block, content: mapLists(block.content, map) };
        case "table":
          return {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({
                ...cell,
                content: mapLists(cell.content, map),
              })),
            })),
          };
        default:
          return block;
      }
    }),
  );

/**
 * A custom XML element a resolution reads through: its tags, the elements
 * around it, and the blocks on either side of its content in its list read
 * without tags, which say where it goes back when none of its content does.
 */
type Shell = {
  element: BlockCustomXml;
  parent: Chain;
  before: BlockContent | undefined;
  after: BlockContent | undefined;
  /** Where the block after it stands in its list read without tags. */
  end: number;
};

/** The elements, outermost first, a block stands inside in its block list. */
type Chain = readonly Shell[];

/** Every paragraph id a block holds, at any depth. */
const paragraphKeysIn = (block: BlockContent): string[] =>
  storyParagraphs({ content: [block] }).flatMap(({ paragraph }) =>
    paragraph.paraId === undefined ? [] : [idKey(paragraph.paraId)],
  );

/** Whether an element holds anything its tags wrap, its properties apart. */
const wrapsContent = (element: BlockCustomXml): boolean =>
  element.content.some((child) => !isCustomXmlProperties(child));

/**
 * Where each block of a story stood among its list's custom XML elements,
 * read so that it is found again after a resolution rebuilt it: a paragraph
 * by its id, a table or content control by any paragraph it holds, anything
 * else as the record it is. Lists nest in table cells and content controls,
 * counted by `depth`; custom XML elements are steps of a list, not lists.
 */
type Placements = {
  byKey: Map<string, Chain>;
  byRecord: WeakMap<object, Chain>;
  shells: Shell[];
  /** The elements put back, around content or emptied. */
  placed: Set<Shell>;
};

const placementKey = (depth: number, paraKey: string): string =>
  `${String(depth)}:${paraKey}`;

/** Whether a resolved block is the one `original` was, or what it became. */
const sameOrigin = (
  resolved: BlockContent,
  original: BlockContent,
): boolean => {
  if (resolved === original) return true;
  if (resolved.type !== original.type) return false;
  if (
    resolved.type !== "paragraph" &&
    resolved.type !== "table" &&
    resolved.type !== "blockSdt"
  )
    return false;
  const keys = new Set(paragraphKeysIn(original));
  return paragraphKeysIn(resolved).some((key) => keys.has(key));
};

/**
 * `blocks` with every custom XML element wrapping content taken away, written
 * into `out` (one list's blocks, through the elements it holds).
 */
const untagInto = (
  blocks: readonly BlockContent[],
  depth: number,
  placements: Placements,
  chain: Chain,
  out: BlockContent[],
): void => {
  for (const block of blocks) {
    if (block.type === "blockCustomXml" && wrapsContent(block)) {
      const shell: Shell = {
        element: block,
        parent: chain,
        before: out.at(-1),
        after: undefined,
        end: 0,
      };
      placements.shells.push(shell);
      untagInto(
        block.content.filter((child) => !isCustomXmlProperties(child)),
        depth,
        placements,
        [...chain, shell],
        out,
      );
      shell.end = out.length;
      continue;
    }
    let placed: BlockContent = block;
    if (block.type === "blockSdt")
      placed = {
        ...block,
        content: untagged(block.content, depth + 1, placements),
      };
    else if (block.type === "table")
      placed = {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            content: untagged(cell.content, depth + 1, placements),
          })),
        })),
      };
    if (block.type === "paragraph") {
      if (block.paraId !== undefined)
        placements.byKey.set(placementKey(depth, idKey(block.paraId)), chain);
    } else if (block.type === "table" || block.type === "blockSdt")
      for (const key of paragraphKeysIn(block))
        placements.byKey.set(placementKey(depth, key), chain);
    placements.byRecord.set(placed, chain);
    out.push(placed);
  }
};

const untagged = (
  blocks: readonly BlockContent[],
  depth: number,
  placements: Placements,
): BlockContent[] => {
  const out: BlockContent[] = [];
  const first = placements.shells.length;
  untagInto(blocks, depth, placements, [], out);
  // The block after each element's content, once the list's blocks are in.
  for (const shell of placements.shells.slice(first))
    if (shell.after === undefined) shell.after = out[shell.end];
  return out;
};

/** Where a block of a resolved list stands, if it is found. */
const chainOf = (
  block: BlockContent,
  depth: number,
  placements: Placements,
): Chain | undefined => {
  const byRecord = placements.byRecord.get(block);
  if (byRecord !== undefined) return byRecord;
  if (
    block.type !== "paragraph" &&
    block.type !== "table" &&
    block.type !== "blockSdt"
  )
    return undefined;
  for (const key of paragraphKeysIn(block)) {
    const chain = placements.byKey.get(placementKey(depth, key));
    if (chain !== undefined) return chain;
  }
  return undefined;
};

const commonPrefix = (left: Chain, right: Chain): Chain => {
  let length = 0;
  while (
    length < left.length &&
    length < right.length &&
    left[length] === right[length]
  )
    length += 1;
  return left.slice(0, length);
};

/**
 * A resolved list with the tags put back around the blocks that stood inside
 * them: each element around the run of its blocks that remain, nested as it
 * was. A block the resolution made anew stands in what the blocks on both
 * sides of it share. An element none of whose content remains goes with
 * that content when a rejection took it, as an element inserted with it
 * goes; one whose content an acceptance removed stays, emptied, where it
 * stood, as removing that content directly leaves it. `undefined` when an
 * element's blocks no longer stand together, or an emptied element has no
 * block left beside it to stand by.
 */
const retagged = (
  blocks: readonly BlockContent[],
  depth: number,
  placements: Placements,
  keepEmptied: boolean,
): BlockContent[] | undefined => {
  const found = blocks.map((block) => chainOf(block, depth, placements));
  const entries: { block: BlockContent | undefined; chain: Chain }[] =
    blocks.map((block, index) => {
      const chain = found[index];
      if (chain !== undefined) return { block, chain };
      const before = found
        .slice(0, index)
        .findLast((each) => each !== undefined);
      const after = found.slice(index + 1).find((each) => each !== undefined);
      return {
        block,
        chain:
          before === undefined || after === undefined
            ? []
            : commonPrefix(before, after),
      };
    });
  if (keepEmptied) {
    // An emptied element of this list stands where its content stood: before
    // the block that followed it, else after the one before it. Only the
    // innermost of nested emptied elements needs a place: the elements
    // around it stand around it.
    const standing = new Set(entries.flatMap(({ chain }) => chain));
    const emptied = placements.shells.filter(
      (shell) =>
        !standing.has(shell) &&
        [shell.before, shell.after].some(
          (side) =>
            side !== undefined &&
            blocks.some((block) => sameOrigin(block, side)),
        ),
    );
    const innermost = emptied.filter(
      (shell) => !emptied.some((other) => other.parent.includes(shell)),
    );
    for (const shell of innermost) {
      const at = blocks.findIndex(
        (block) => shell.after !== undefined && sameOrigin(block, shell.after),
      );
      const by = blocks.findIndex(
        (block) =>
          shell.before !== undefined && sameOrigin(block, shell.before),
      );
      const entry = { block: undefined, chain: [...shell.parent, shell] };
      if (at >= 0) {
        const index = entries.findIndex(({ block }) => block === blocks[at]);
        entries.splice(index, 0, entry);
      } else if (by >= 0) {
        const index = entries.findIndex(({ block }) => block === blocks[by]);
        entries.splice(index + 1, 0, entry);
      } else return undefined;
    }
  }
  type Frame = { shell: Shell; content: BlockContent[] };
  const root: BlockContent[] = [];
  const open: Frame[] = [];
  const closed = new Set<Shell>();
  const close = () => {
    const frame = open.pop();
    if (frame === undefined) return;
    closed.add(frame.shell);
    (open.at(-1)?.content ?? root).push({
      ...frame.shell.element,
      content: [
        ...frame.shell.element.content.filter(isCustomXmlProperties),
        ...frame.content,
      ],
    });
  };
  for (const { block: source, chain } of entries) {
    let shared = 0;
    while (
      shared < open.length &&
      shared < chain.length &&
      open[shared]?.shell === chain[shared]
    )
      shared += 1;
    while (open.length > shared) close();
    for (const shell of chain.slice(shared)) {
      if (closed.has(shell)) return undefined;
      placements.placed.add(shell);
      open.push({ shell, content: [] });
    }
    if (source === undefined) continue;
    let block = source;
    if (block.type === "blockSdt") {
      const content = retagged(
        block.content,
        depth + 1,
        placements,
        keepEmptied,
      );
      if (content === undefined) return undefined;
      block = { ...block, content };
    } else if (block.type === "table") {
      const rows: TableRow[] = [];
      for (const row of block.rows) {
        const cells: TableCell[] = [];
        for (const cell of row.cells) {
          const content = retagged(
            cell.content,
            depth + 1,
            placements,
            keepEmptied,
          );
          if (content === undefined) return undefined;
          cells.push(Object.assign({}, cell, { content }));
        }
        rows.push(Object.assign({}, row, { cells }));
      }
      block = Object.assign({}, block, { rows });
    }
    (open.at(-1)?.content ?? root).push(block);
  }
  while (open.length > 0) close();
  return root;
};

/** Whether `value` holds a record with one of `ids`. */
const holdsRevision = (value: unknown, ids: ReadonlySet<number>): boolean => {
  if (typeof value !== "object" || value === null) return false;
  if (Array.isArray(value))
    return value.some((item) => holdsRevision(item, ids));
  if (
    "info" in value &&
    typeof value.info === "object" &&
    value.info !== null &&
    "id" in value.info &&
    typeof value.info.id === "number" &&
    ids.has(value.info.id)
  )
    return true;
  return Object.values(value).some((field) => holdsRevision(field, ids));
};

/**
 * How a resolution reads through the custom XML elements of a story: the
 * operation taking their tags away, and the one putting them back around
 * what their content left. A resolution reaching into an element reads its
 * content as if the tags were not there, at any depth of elements, through
 * tables and content controls: a mark ending an element joins the paragraph
 * after it across every element end it stands inside, as it would without
 * them. The tags then go back around their content that remains; an element
 * whose content a rejection took goes too (no empty element is left where
 * the suggestion added none), and one whose content an
 * acceptance removed stays emptied, as removing it directly leaves it. So the
 * order the tags' own records and their content's resolve in leaves the same
 * document. `undefined` when the resolution reaches into no
 * element.
 */
export const customXmlReadThrough = ({
  document,
  story,
  revisionIds,
  decision,
}: {
  document: Document;
  story: OpStory;
  revisionIds: readonly number[];
  decision: RevisionDecision;
}):
  | {
      unwrap: DocumentOp;
      rewrap: (resolved: Document) => DocumentOp | undefined;
    }
  | undefined => {
  const body = storyBody(document, story);
  const selected = new Set(revisionIds);
  let reached = false;
  mapLists(body.content, (list) => {
    if (
      list.some(
        (block) =>
          block.type === "blockCustomXml" &&
          wrapsContent(block) &&
          holdsRevision(block.content, selected),
      )
    )
      reached = true;
    return [...list];
  });
  const anchor = storyParagraphs(body).find(
    ({ paragraph }) => paragraph.paraId !== undefined,
  );
  if (!reached || anchor?.paragraph.paraId === undefined) return undefined;
  const address = (located: {
    list: readonly unknown[];
    paragraph: Paragraph;
  }) => ({
    blockId: located.paragraph.paraId ?? "",
    ...(located.list.length === 0
      ? {}
      : { ancestorDepth: located.list.length }),
  });
  const placements: Placements = {
    byKey: new Map(),
    byRecord: new WeakMap(),
    shells: [],
    placed: new Set(),
  };
  const unwrap: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_CUSTOM_XML_MARKUP,
    story,
    ...address(anchor),
    expected: body.content,
    blocks: untagged(body.content, 0, placements),
  };
  const rewrap = (resolved: Document): DocumentOp | undefined => {
    const resolvedBody = storyBody(resolved, story);
    const keepEmptied = decision === REVISION_DECISIONS.ACCEPT;
    placements.placed.clear();
    const content = retagged(resolvedBody.content, 0, placements, keepEmptied);
    // An emptied element an acceptance keeps that found no place would be lost.
    if (
      keepEmptied &&
      placements.shells.some((shell) => !placements.placed.has(shell))
    )
      return undefined;
    const located = storyParagraphs(resolvedBody).find(
      ({ paragraph }) => paragraph.paraId !== undefined,
    );
    if (content === undefined || located === undefined) return undefined;
    return {
      type: DOCUMENT_OP_TYPES.SET_CUSTOM_XML_MARKUP,
      story,
      ...address(located),
      expected: resolvedBody.content,
      blocks: content,
    };
  };
  return { unwrap, rewrap };
};
