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

/**
 * A custom XML element a resolution reads through: one wrapping paragraphs
 * only, each with an id, so where its content stands after the resolution
 * says where its tags go back.
 */
type ReadThrough = {
  element: BlockCustomXml;
  /** Its paragraphs' ids, in order. */
  members: readonly string[];
  /** The paragraph after it, then the one before it, read through tags. */
  after: string | undefined;
  before: string | undefined;
};

const simpleElement = (element: BlockCustomXml): boolean =>
  element.content.every(
    (child) =>
      isCustomXmlProperties(child) ||
      (child.type === "paragraph" &&
        child.paraId !== undefined &&
        child.sectionProperties === undefined),
  );

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

const paragraphKeyOf = (block: BlockContent | undefined): string | undefined =>
  block?.type === "paragraph" && block.paraId !== undefined
    ? idKey(block.paraId)
    : undefined;

/** The paragraphs' ids a block puts in its list's flow, read through tags. */
const flowKeysOf = (block: BlockContent): string[] => {
  if (block.type === "blockCustomXml")
    return block.content.flatMap((child) => flowKeysOf(child));
  const key = paragraphKeyOf(block);
  return key === undefined ? [] : [key];
};

/**
 * How a resolution reads through the custom XML elements of a story: the
 * operation taking their tags away, and the one putting them back around
 * what their content left. A paragraph mark resolved at an element's end
 * (on its last paragraph) joins across the tags as it would without them, and the tags go back around the element's
 * paragraphs that stand after, or where it stood when none does. So the
 * order the tags' own records and their content's are resolved in leaves the
 * same document. `undefined` when no resolved mark ends an element,
 * or when an element wraps more than paragraphs: the resolution then meets
 * the tags as blocks.
 */
export const customXmlReadThrough = ({
  document,
  story,
  revisionIds,
}: {
  document: Document;
  story: OpStory;
  revisionIds: readonly number[];
}):
  | {
      unwrap: DocumentOp;
      rewrap: (resolved: Document) => DocumentOp | undefined;
    }
  | undefined => {
  const body = storyBody(document, story);
  const selected = new Set(revisionIds);
  const shells: ReadThrough[] = [];
  let simple = true;
  let atEdge = false;
  const resolvedMark = (block: BlockContent | undefined): boolean =>
    block?.type === "paragraph" &&
    block.pPrMark !== undefined &&
    (selected.has(block.pPrMark.info.id) ||
      (block.pPrMark.deletion !== undefined &&
        selected.has(block.pPrMark.deletion.info.id)));
  mapLists(body.content, (list) => {
    const flow = list.map(flowKeysOf);
    for (const [index, block] of list.entries()) {
      if (block.type !== "blockCustomXml") continue;
      if (!simpleElement(block)) {
        simple = false;
        continue;
      }
      const paragraphs = block.content.filter(
        (child): child is Paragraph => child.type === "paragraph",
      );
      // Before an element, a resolved mark's paragraph stays or goes by what
      // it holds, as before any block; only the element's last mark has
      // nowhere to join inside it.
      if (resolvedMark(paragraphs.at(-1))) atEdge = true;
      shells.push({
        element: block,
        members: flow[index] ?? [],
        after: flow
          .slice(index + 1)
          .flat()
          .at(0),
        before: flow.slice(0, index).flat().at(-1),
      });
    }
    return [...list];
  });
  const anchor = storyParagraphs(body).find(
    ({ paragraph }) => paragraph.paraId !== undefined,
  );
  if (
    !simple ||
    !atEdge ||
    shells.length === 0 ||
    anchor?.paragraph.paraId === undefined
  )
    return undefined;
  const address = (located: {
    list: readonly unknown[];
    paragraph: Paragraph;
  }) => ({
    blockId: located.paragraph.paraId ?? "",
    ...(located.list.length === 0
      ? {}
      : { ancestorDepth: located.list.length }),
  });
  const unwrap: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_CUSTOM_XML_MARKUP,
    story,
    ...address(anchor),
    expected: body.content,
    blocks: mapLists(body.content, (list) =>
      list.flatMap((block) =>
        block.type === "blockCustomXml"
          ? block.content.filter((child) => !isCustomXmlProperties(child))
          : [block],
      ),
    ),
  };
  const rewrap = (resolved: Document): DocumentOp | undefined => {
    const resolvedBody = storyBody(resolved, story);
    const present = new Set(
      storyParagraphs(resolvedBody).flatMap(({ paragraph }) =>
        paragraph.paraId === undefined ? [] : [idKey(paragraph.paraId)],
      ),
    );
    const runs = new Map<string, ReadThrough>();
    const emptyBefore = new Map<string, ReadThrough[]>();
    const emptyAfter = new Map<string, ReadThrough[]>();
    for (const shell of shells) {
      const first = shell.members.find((key) => present.has(key));
      if (first !== undefined) {
        runs.set(first, shell);
        continue;
      }
      if (shell.after !== undefined && present.has(shell.after))
        emptyBefore.set(shell.after, [
          ...(emptyBefore.get(shell.after) ?? []),
          shell,
        ]);
      else if (shell.before !== undefined && present.has(shell.before))
        emptyAfter.set(shell.before, [
          ...(emptyAfter.get(shell.before) ?? []),
          shell,
        ]);
      else return undefined;
    }
    let placed = 0;
    let failed = false;
    const emptied = (shell: ReadThrough): BlockContent => {
      placed += 1;
      return {
        ...shell.element,
        content: shell.element.content.filter(isCustomXmlProperties),
      };
    };
    const content = mapLists(resolvedBody.content, (list) => {
      const out: BlockContent[] = [];
      for (let index = 0; index < list.length; index += 1) {
        const block = list[index];
        if (block === undefined) continue;
        const key = paragraphKeyOf(block);
        if (key === undefined) {
          out.push(block);
          continue;
        }
        out.push(...(emptyBefore.get(key) ?? []).map(emptied));
        const shell = runs.get(key);
        let last = key;
        if (shell === undefined) out.push(block);
        else {
          const remaining = shell.members.filter((member) =>
            present.has(member),
          );
          const run = list.slice(index, index + remaining.length);
          if (
            run.length !== remaining.length ||
            run.some((each, at) => paragraphKeyOf(each) !== remaining[at])
          )
            failed = true;
          out.push({
            ...shell.element,
            content: [
              ...shell.element.content.filter(isCustomXmlProperties),
              ...run,
            ],
          });
          placed += 1;
          index += run.length - 1;
          last = remaining.at(-1) ?? key;
        }
        out.push(...(emptyAfter.get(last) ?? []).map(emptied));
      }
      return out;
    });
    if (failed || placed !== shells.length) return undefined;
    const located = storyParagraphs(resolvedBody).find(
      ({ paragraph }) => paragraph.paraId !== undefined,
    );
    if (located === undefined) return undefined;
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
