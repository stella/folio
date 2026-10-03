/**
 * Finding and replacing paragraphs in a story.
 *
 * A story is a tree of block lists: the body, each table cell and each block
 * content control holds one. A paragraph is found by its `w14:paraId`
 * wherever it sits in that tree, and a replacement rebuilds only the records
 * on the path to it, so every other block is `===` to its input.
 */

import { panic, Result, TaggedError } from "better-result";

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
import { DOCUMENT_OP_REFUSAL_REASONS } from "./refusal";
import {
  type OpStory,
  type SectionViewEntry,
  type SectionMapState,
  type SectionViewState,
} from "./types";
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

/** A section edit that cannot preserve the canonical section contract. */
export class SectionViewError extends TaggedError("SectionViewError")<{
  message: string;
  reason:
    | typeof DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY
    | typeof DOCUMENT_OP_REFUSAL_REASONS.STALE;
}> {}

const sectionFailure = (message: string) =>
  Result.err(
    new SectionViewError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      message,
    }),
  );

const sectionMapState = (section: Section, key: "headers" | "footers"): SectionMapState => {
  if (!Object.hasOwn(section, key)) return { type: "omitted" };
  const value = section[key];
  return value === undefined ? { type: "undefined" } : { type: "entries", value: [...value] };
};

/** Capture only section metadata, with explicit JSON-safe map presence. */
export const captureSectionView = (sections: readonly Section[]): SectionViewEntry[] =>
  sections.map((section) => ({
    properties: section.properties,
    headers: sectionMapState(section, "headers"),
    footers: sectionMapState(section, "footers"),
  }));

export const captureSectionViewState = (body: DocumentBody): SectionViewState => {
  if (!Object.hasOwn(body, "sections")) return { type: "omitted" };
  return body.sections === undefined
    ? { type: "undefined" }
    : { type: "sections", value: captureSectionView(body.sections) };
};

const restoreSectionMap = (
  section: Section,
  key: "headers" | "footers",
  state: SectionMapState,
): void => {
  switch (state.type) {
    case "omitted":
      return;
    case "undefined":
      section[key] = undefined;
      return;
    case "entries":
      section[key] = new Map(state.value);
      return;
    default: {
      const unreachable: never = state;
      return unreachable;
    }
  }
};

export const restoreSectionView = (
  content: readonly BlockContent[],
  snapshots: readonly SectionViewEntry[],
): Result<Section[], SectionViewError> => {
  const groups = sectionGroups(content);
  if (groups.length !== snapshots.length)
    return Result.err(
      new SectionViewError({
        reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
        message: "The section snapshot does not match the restored section boundaries.",
      }),
    );
  const sections: Section[] = [];
  for (const [index, snapshot] of snapshots.entries()) {
    const group = groups.at(index);
    if (group === undefined) return sectionFailure("The restored section lost its block group.");
    const section: Section = { properties: snapshot.properties, content: group };
    restoreSectionMap(section, "headers", snapshot.headers);
    restoreSectionMap(section, "footers", snapshot.footers);
    sections.push(section);
  }
  return Result.ok(sections);
};

export const restoreSectionViewState = (
  body: DocumentBody,
  state: SectionViewState,
): Result<DocumentBody, SectionViewError> => {
  const next = { ...body };
  delete next.sections;
  switch (state.type) {
    case "omitted":
      return Result.ok(next);
    case "undefined":
      return Result.ok({ ...next, sections: undefined });
    case "sections": {
      const restored = restoreSectionView(body.content, state.value);
      if (restored.isErr()) return restored;
      next.sections = restored.value;
      return Result.ok(next);
    }
    default: {
      const unreachable: never = state;
      return unreachable;
    }
  }
};

type SectionPartsOptions = {
  references: readonly HeaderReference[] | undefined;
  parts: ReadonlyMap<string, HeaderFooter> | undefined;
};

const sectionParts = ({ references, parts }: SectionPartsOptions) => {
  if (references === undefined || references.length === 0) return Result.ok(undefined);
  const resolved = new Map<HeaderFooterType, HeaderFooter>();
  for (const { type, rId } of references) {
    const part = parts?.get(rId);
    if (part === undefined)
      return sectionFailure(`A section reference names unavailable part ${rId}.`);
    resolved.set(type, part);
  }
  return Result.ok(resolved);
};

type RebuildSectionsOptions = {
  document: Document;
  content: readonly BlockContent[];
  previous: readonly Section[];
};

/** A boundary keeps its section metadata when earlier section boundaries change. */
export const rebuildSections = ({
  document,
  content,
  previous,
}: RebuildSectionsOptions): Result<Section[], SectionViewError> => {
  const body = document.package.document;
  const previousGroups = sectionGroups(body.content);
  const sectionsByBoundary = new Map<string, Section>();
  let finalSection: Section | undefined;
  for (const [index, group] of previousGroups.entries()) {
    const section = previous.at(index);
    if (section === undefined) return sectionFailure("The previous section lost its block group.");
    const boundary = group.at(-1);
    if (boundary?.type === "paragraph" && boundary.sectionProperties !== undefined) {
      if (boundary.paraId === undefined)
        return sectionFailure("A section boundary has no paragraph id.");
      sectionsByBoundary.set(boundary.paraId, section);
    } else {
      finalSection = section;
    }
  }
  const sections: Section[] = [];
  const nextGroups = sectionGroups(content);
  for (const [index, group] of nextGroups.entries()) {
    const boundary = group.at(-1);
    const stated = boundary?.type === "paragraph" ? boundary.sectionProperties : undefined;
    let matching = finalSection;
    if (stated !== undefined && boundary?.type === "paragraph") {
      matching = undefined;
      if (boundary.paraId !== undefined) matching = sectionsByBoundary.get(boundary.paraId);
    }
    const properties = stated ?? body.finalSectionProperties ?? finalSection?.properties;
    // A split can move an existing endpoint to a fresh paragraph without changing its section.
    if (matching === undefined && nextGroups.length === previousGroups.length)
      matching = previous.at(index);
    if (properties === undefined)
      return sectionFailure("A rebuilt section has no canonical properties.");
    if (matching !== undefined && structurallyEqual(matching.properties, properties)) {
      const same =
        group.length === matching.content.length &&
        group.every((block, position) => block === matching.content[position]);
      sections.push(same ? matching : { ...matching, properties, content: group });
      continue;
    }
    const section: Section = { ...matching, properties, content: group };
    if (
      matching === undefined ||
      !structurallyEqual(matching.properties.headerReferences, properties.headerReferences)
    ) {
      const headers = sectionParts({
        references: properties.headerReferences,
        parts: document.package.headers,
      });
      if (headers.isErr()) return headers;
      delete section.headers;
      if (headers.value !== undefined) section.headers = headers.value;
    }
    if (
      matching === undefined ||
      !structurallyEqual(matching.properties.footerReferences, properties.footerReferences)
    ) {
      const footers = sectionParts({
        references: properties.footerReferences,
        parts: document.package.footers,
      });
      if (footers.isErr()) return footers;
      delete section.footers;
      if (footers.value !== undefined) section.footers = footers.value;
    }
    sections.push(section);
  }
  return Result.ok(sections);
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
}: ReplaceParagraphsOptions): Result<Document, SectionViewError> => {
  const body = storyBody(document, story);
  const content = updateBlockList(body.content, at.list, (blocks) => {
    const out = [...blocks];
    out.splice(at.index, count, ...replacement);
    return out;
  });
  const nextBody: DocumentBody = { ...body, content };
  if (restoreSections !== undefined) {
    const restored = restoreSectionView(content, restoreSections);
    if (restored.isErr()) return restored;
    nextBody.sections = restored.value;
  } else if (body.sections !== undefined) {
    const rebuilt = rebuildSections({ document, content, previous: body.sections });
    if (rebuilt.isErr()) return rebuilt;
    nextBody.sections = rebuilt.value;
  }
  if (!sectionsInStep(nextBody))
    return sectionFailure("The restored section metadata conflicts with its canonical boundaries.");
  return Result.ok(replaceStoryBody({ document, story, body: nextBody }));
};
