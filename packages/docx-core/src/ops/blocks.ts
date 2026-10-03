/**
 * Finding and replacing paragraphs in a story.
 *
 * A story is a tree of block lists: the body, each table cell and each block
 * content control holds one. A paragraph is found by its `w14:paraId`
 * wherever it sits in that tree, and a replacement rebuilds only the records
 * on the path to it, so every other block is `===` to its input.
 */

import { panic } from "better-result";

import type {
  BlockContent,
  Document,
  DocumentBody,
  HeaderFooter,
  HeaderFooterType,
  HeaderReference,
  Paragraph,
  Section,
} from "../model/document";
import { structurallyEqual } from "./equality";
import type { OpStory, SectionViewEntry } from "./types";
import { storyBody, replaceStoryBody } from "./stories";

export { storyBody } from "./stories";

/** One step from a block list down to a block list nested in one of its blocks. */
type BlockListStep =
  | { kind: "tableCell"; block: number; row: number; cell: number }
  | { kind: "blockSdt"; block: number }
  | { kind: "blockCustomXml"; block: number };

/** Where a paragraph sits: the block list that holds it, and its index there. */
export type ParagraphLocation = {
  list: readonly BlockListStep[];
  index: number;
  paragraph: Paragraph;
};

const collectParagraphs = (
  blocks: readonly BlockContent[],
  list: readonly BlockListStep[],
  out: ParagraphLocation[],
): void => {
  for (const [index, block] of blocks.entries()) {
    switch (block.type) {
      case "paragraph":
        out.push({ list, index, paragraph: block });
        break;
      case "table":
        for (const [row, tableRow] of block.rows.entries()) {
          for (const [cell, tableCell] of tableRow.cells.entries()) {
            collectParagraphs(
              tableCell.content,
              [...list, { kind: "tableCell", block: index, row, cell }],
              out,
            );
          }
        }
        break;
      case "blockSdt":
        collectParagraphs(block.content, [...list, { kind: "blockSdt", block: index }], out);
        break;
      case "blockCustomXml":
        collectParagraphs(block.content, [...list, { kind: "blockCustomXml", block: index }], out);
        break;
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
};

/** Every paragraph of a story's block tree, in document order. */
export const storyParagraphs = (body: DocumentBody): ParagraphLocation[] => {
  const out: ParagraphLocation[] = [];
  collectParagraphs(body.content, [], out);
  return out;
};

const sameStep = (left: BlockListStep, right: BlockListStep): boolean => {
  switch (left.kind) {
    case "tableCell":
      return (
        right.kind === "tableCell" &&
        left.block === right.block &&
        left.row === right.row &&
        left.cell === right.cell
      );
    case "blockSdt":
      return right.kind === "blockSdt" && left.block === right.block;
    case "blockCustomXml":
      return right.kind === "blockCustomXml" && left.block === right.block;
    default: {
      const unreachable: never = left;
      return unreachable;
    }
  }
};

/** Whether two locations are in the same block list. */
export const sameBlockList = (
  left: readonly BlockListStep[],
  right: readonly BlockListStep[],
): boolean =>
  left.length === right.length &&
  left.every((step, index) => {
    const other = right[index];
    return other !== undefined && sameStep(step, other);
  });

/** The block list a location's steps lead to. */
export const blockListAt = (
  blocks: readonly BlockContent[],
  list: readonly BlockListStep[],
): readonly BlockContent[] => {
  let current = blocks;
  for (const step of list) {
    const block = current[step.block];
    switch (step.kind) {
      case "tableCell": {
        const cell = block?.type === "table" ? block.rows[step.row]?.cells[step.cell] : undefined;
        if (cell === undefined) {
          return panic(`Block list step does not name a table cell at block ${step.block}.`);
        }
        current = cell.content;
        break;
      }
      case "blockSdt":
        if (block?.type !== "blockSdt") {
          return panic(`Block list step does not name a block content control at ${step.block}.`);
        }
        current = block.content;
        break;
      case "blockCustomXml":
        if (block?.type !== "blockCustomXml") {
          return panic(`Block list step does not name a custom XML wrapper at ${step.block}.`);
        }
        current = block.content;
        break;
      default: {
        const unreachable: never = step;
        return unreachable;
      }
    }
  }
  return current;
};

/**
 * Whether a block is the last one of its story body or table cell. A block
 * content control or custom XML wrapper is a step in the block list, not a
 * container: the last block inside one ends the container only when the
 * wrapper does.
 */
export const endsItsContainer = (
  body: DocumentBody,
  { list, index }: { list: readonly BlockListStep[]; index: number },
): boolean => {
  let steps = list;
  let position = index;
  for (;;) {
    if (position !== blockListAt(body.content, steps).length - 1) {
      return false;
    }
    const last = steps.at(-1);
    if (last === undefined || last.kind === "tableCell") {
      return true;
    }
    steps = steps.slice(0, -1);
    position = last.block;
  }
};

export const updateBlockList = (
  blocks: readonly BlockContent[],
  list: readonly BlockListStep[],
  update: (blocks: readonly BlockContent[]) => BlockContent[],
): BlockContent[] => {
  const [step, ...rest] = list;
  if (step === undefined) {
    return update(blocks);
  }
  const block = blocks[step.block];
  const out = [...blocks];
  switch (step.kind) {
    case "tableCell": {
      const row = block?.type === "table" ? block.rows[step.row] : undefined;
      const cell = row?.cells[step.cell];
      if (block?.type !== "table" || row === undefined || cell === undefined) {
        return panic(`Block list step does not name a table cell at block ${step.block}.`);
      }
      const cells = [...row.cells];
      cells[step.cell] = { ...cell, content: updateBlockList(cell.content, rest, update) };
      const rows = [...block.rows];
      rows[step.row] = { ...row, cells };
      out[step.block] = { ...block, rows };
      return out;
    }
    case "blockSdt": {
      if (block?.type !== "blockSdt") {
        return panic(`Block list step does not name a block content control at ${step.block}.`);
      }
      out[step.block] = { ...block, content: updateBlockList(block.content, rest, update) };
      return out;
    }
    case "blockCustomXml": {
      if (block?.type !== "blockCustomXml") {
        return panic(`Block list step does not name a custom XML wrapper at ${step.block}.`);
      }
      out[step.block] = { ...block, content: updateBlockList(block.content, rest, update) };
      return out;
    }
    default: {
      const unreachable: never = step;
      return unreachable;
    }
  }
};

const breaksSection = (block: BlockContent): boolean =>
  block.type === "paragraph" && block.sectionProperties !== undefined;

/**
 * The body's blocks grouped as the parser groups them into sections: each
 * top-level paragraph carrying section properties ends one and the rest end
 * the last. The last group is dropped when it is empty and an earlier section
 * exists, as the parser does.
 */
const sectionGroups = (content: readonly BlockContent[]): BlockContent[][] => {
  const groups: BlockContent[][] = [];
  let current: BlockContent[] = [];
  for (const block of content) {
    current.push(block);
    if (breaksSection(block)) {
      groups.push(current);
      current = [];
    }
  }
  if (current.length > 0 || groups.length === 0) {
    groups.push(current);
  }
  return groups;
};

/**
 * `sections`, the parser's per-section view of the body's top-level blocks,
 * derived again from the body after an edit. Section `i` keeps every field of
 * the section it replaces (properties, headers, footers) and holds group `i`;
 * a section whose group is unchanged is the same object. This path owns
 * replacements that preserve the section count.
 */
const deriveSections = (content: readonly BlockContent[], previous: Section[]): Section[] => {
  const groups = sectionGroups(content);
  if (groups.length !== previous.length) {
    return panic(`An edit changed the section count from ${previous.length} to ${groups.length}.`);
  }
  let changed = false;
  const out: Section[] = [];
  for (const [index, section] of previous.entries()) {
    const group = groups[index] ?? [];
    const same =
      group.length === section.content.length &&
      group.every((block, position) => block === section.content[position]);
    changed ||= !same;
    out.push(same ? section : { ...section, content: group });
  }
  return changed ? out : previous;
};

/** Capture only section metadata, with Maps represented as JSON-safe entries. */
export const captureSectionView = (sections: readonly Section[]): SectionViewEntry[] =>
  sections.map(({ properties, headers, footers }) => ({
    properties,
    ...(headers !== undefined && { headers: [...headers] }),
    ...(footers !== undefined && { footers: [...footers] }),
  }));

const restoreSectionView = (
  content: readonly BlockContent[],
  snapshots: readonly SectionViewEntry[],
): Section[] => {
  const groups = sectionGroups(content);
  if (groups.length !== snapshots.length) {
    return panic("The section snapshot does not match the restored section boundaries.");
  }
  return snapshots.map(({ properties, headers, footers }, index) => {
    const group = groups.at(index);
    if (group === undefined) return panic("The restored section lost its block group.");
    return {
      properties,
      content: group,
      ...(headers !== undefined && { headers: new Map(headers) }),
      ...(footers !== undefined && { footers: new Map(footers) }),
    };
  });
};

type SectionPartsOptions = {
  references: readonly HeaderReference[] | undefined;
  parts: ReadonlyMap<string, HeaderFooter> | undefined;
};

const sectionParts = ({ references, parts }: SectionPartsOptions) => {
  if (references === undefined || references.length === 0) return undefined;
  const resolved = new Map<HeaderFooterType, HeaderFooter>();
  for (const { type, rId } of references) {
    const part = parts?.get(rId);
    if (part === undefined) return panic(`A section reference names unavailable part ${rId}.`);
    resolved.set(type, part);
  }
  return resolved;
};

type RebuildSectionsOptions = {
  document: Document;
  content: readonly BlockContent[];
  previous: readonly Section[];
};

/** A boundary keeps its section metadata when earlier section boundaries change. */
const rebuildSections = ({ document, content, previous }: RebuildSectionsOptions): Section[] => {
  const body = document.package.document;
  const previousGroups = sectionGroups(body.content);
  const sectionsByBoundary = new Map<string, Section>();
  let finalSection: Section | undefined;
  for (const [index, group] of previousGroups.entries()) {
    const section = previous.at(index);
    if (section === undefined) return panic("The previous section lost its block group.");
    const boundary = group.at(-1);
    if (boundary?.type === "paragraph" && boundary.sectionProperties !== undefined) {
      if (boundary.paraId === undefined) return panic("A section boundary has no paragraph id.");
      sectionsByBoundary.set(boundary.paraId, section);
    } else {
      finalSection = section;
    }
  }
  return sectionGroups(content).map((group) => {
    const boundary = group.at(-1);
    const stated = boundary?.type === "paragraph" ? boundary.sectionProperties : undefined;
    let matching = finalSection;
    if (stated !== undefined && boundary?.type === "paragraph") {
      matching = undefined;
      if (boundary.paraId !== undefined) matching = sectionsByBoundary.get(boundary.paraId);
    }
    const properties = stated ?? body.finalSectionProperties ?? finalSection?.properties;
    if (properties === undefined) return panic("A rebuilt section has no canonical properties.");
    if (matching !== undefined && structurallyEqual(matching.properties, properties)) {
      const section: Section = { properties, content: group };
      if (matching.headers !== undefined) section.headers = matching.headers;
      if (matching.footers !== undefined) section.footers = matching.footers;
      return section;
    }
    const headers = sectionParts({
      references: properties.headerReferences,
      parts: document.package.headers,
    });
    const footers = sectionParts({
      references: properties.footerReferences,
      parts: document.package.footers,
    });
    const section: Section = { properties, content: group };
    if (headers !== undefined) section.headers = headers;
    if (footers !== undefined) section.footers = footers;
    return section;
  });
};

/**
 * A body holding other blocks, its section view derived from them. A view
 * that does not line up with the blocks (a different section count) is left
 * for {@link sectionsInStep} to report.
 */
export const withBodyContent = (body: DocumentBody, content: BlockContent[]): DocumentBody => {
  const next: DocumentBody = { ...body, content };
  if (body.sections !== undefined && sectionGroups(content).length === body.sections.length) {
    next.sections = deriveSections(content, body.sections);
  }
  return next;
};

/**
 * Whether a body's section view says what its blocks say: the same groups
 * of structurally equal blocks, each ended by a paragraph whose section
 * properties the section carries. A view out of step with its body cannot be
 * edited: every edit derives the view again, and its inverse could not give
 * back the stale one.
 */
export const sectionsInStep = (body: DocumentBody): boolean => {
  const { sections } = body;
  if (sections === undefined) {
    return true;
  }
  const groups = sectionGroups(body.content);
  if (groups.length !== sections.length) {
    return false;
  }
  return sections.every((section, index) => {
    const group = groups[index] ?? [];
    const sameBlocks =
      group.length === section.content.length &&
      group.every(
        (block, position) =>
          block === section.content[position] ||
          structurallyEqual(block, section.content[position]),
      );
    if (!sameBlocks) {
      return false;
    }
    const last = group.at(-1);
    const stated =
      last?.type === "paragraph" && last.sectionProperties !== undefined
        ? last.sectionProperties
        : body.finalSectionProperties;
    return stated === undefined || structurallyEqual(stated, section.properties);
  });
};

type ReplaceParagraphsOptions = {
  document: Document;
  story: OpStory;
  /** The first paragraph replaced; the rest follow it in the same block list. */
  at: ParagraphLocation;
  count: number;
  replacement: readonly Paragraph[];
  /** Exact inverse metadata; section content is rebuilt from the restored blocks. */
  restoreSections?: readonly SectionViewEntry[];
};

/** The document with `count` paragraphs from `at` replaced. */
export const replaceParagraphs = ({
  document,
  story,
  at,
  count,
  replacement,
  restoreSections,
}: ReplaceParagraphsOptions): Document => {
  const body = storyBody(document, story);
  const content = updateBlockList(body.content, at.list, (blocks) => {
    const out = [...blocks];
    out.splice(at.index, count, ...replacement);
    return out;
  });
  const nextBody: DocumentBody = { ...body, content };
  if (restoreSections !== undefined) {
    nextBody.sections = restoreSectionView(content, restoreSections);
  } else if (body.sections !== undefined) {
    nextBody.sections =
      sectionGroups(content).length === body.sections.length
        ? deriveSections(content, body.sections)
        : rebuildSections({ document, content, previous: body.sections });
  }
  return replaceStoryBody({ document, story, body: nextBody });
};
