import {
  documentStories,
  findStoryBody,
  type OpStory,
  withParagraphCopies,
} from "@stll/docx-core/ops";

import {
  carryParagraphPropertySource,
  copyDocumentParagraphPropertySourceContract,
} from "../docx/paragraphPropertySource";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { changedWindow } from "./canonicalSeedContent";

/**
 * Carry private paragraph source captures across immutable story operations.
 * A paragraph a reader kept is never written: one whose facts change is
 * replaced by a new paragraph carrying them, so the result is `target` or a
 * new document.
 */
export type PreservePropertySourcesOptions = {
  target: Document;
  source: Document;
  stories?: readonly OpStory[];
  indentation?: "derive" | "restore";
};

const collectParagraphs = (blocks: readonly BlockContent[]): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  for (const block of blocks) {
    if (block.type === "paragraph") {
      paragraphs.push(block);
      continue;
    }
    if (block.type !== "table") continue;
    for (const row of block.rows) {
      for (const cell of row.cells)
        paragraphs.push(...collectParagraphs(cell.content));
    }
  }
  return paragraphs;
};

/**
 * The copies that carry each original paragraph's sources onto the derived
 * paragraph with its id, or, for an original without one, onto the derived
 * paragraph at its place among `derived`'s.
 */
const carriedCopies = (
  originals: readonly BlockContent[],
  derived: readonly BlockContent[],
  indentation: "derive" | "restore",
): Map<Paragraph, Paragraph> => {
  const originalParagraphs = collectParagraphs(originals);
  const derivedParagraphs = collectParagraphs(derived);
  const paragraphs = new Map(
    derivedParagraphs.flatMap((paragraph) =>
      paragraph.paraId === undefined
        ? []
        : [[paragraph.paraId, paragraph] as const],
    ),
  );
  const replaced = new Map<Paragraph, Paragraph>();
  for (const [index, original] of originalParagraphs.entries()) {
    const next =
      original.paraId === undefined
        ? derivedParagraphs.at(index)
        : paragraphs.get(original.paraId);
    if (next === undefined || next === original) continue;
    const carried = carryParagraphPropertySource({
      target: replaced.get(next) ?? next,
      source: original,
      indentation,
    });
    if (carried !== next) replaced.set(next, carried);
  }
  return replaced;
};

/**
 * {@link preservePropertySources} walking every paragraph of every story:
 * the reference the changed-block version must equal.
 */
export const preservePropertySourcesWholeStories = ({
  target,
  source,
  stories = documentStories(source),
  indentation = "derive",
}: PreservePropertySourcesOptions): Document => {
  let document = target;
  for (const story of stories) {
    const originals = findStoryBody(source, story)?.content ?? [];
    const derived = findStoryBody(document, story)?.content ?? [];
    if (originals === derived) continue;
    document = withParagraphCopies({
      document,
      story,
      copies: carriedCopies(originals, derived, indentation),
    });
  }
  copyDocumentParagraphPropertySourceContract(document, source);
  return document;
};

/**
 * Carry each source paragraph's captures onto the target paragraph that
 * carries its id. A block the target kept from the source is the same
 * object, and every paragraph in it is the paragraph it was, so only the
 * paragraphs between a story's kept prefix and suffix are matched: a
 * paragraph id is unique in a document, so the one an original in that run
 * carries is found in the target's run or nowhere. An original without an id
 * there is matched by its place in the whole story, as the whole walk does;
 * one without an id outside the run is the paragraph it was and keeps its
 * captures, where the whole walk could pair it by place with another
 * paragraph once the run changed how many paragraphs there are. A session
 * gives every paragraph an id, so its edits meet only the first case.
 */
export const preservePropertySources = ({
  target,
  source,
  stories = documentStories(source),
  indentation = "derive",
}: PreservePropertySourcesOptions): Document => {
  let document = target;
  for (const story of stories) {
    const originals = findStoryBody(source, story)?.content ?? [];
    const derived = findStoryBody(document, story)?.content ?? [];
    if (originals === derived) continue;
    const { removed, added } = changedWindow(originals, derived);
    const placeMatched = collectParagraphs(removed).some(
      ({ paraId }) => paraId === undefined,
    );
    document = withParagraphCopies({
      document,
      story,
      copies: placeMatched
        ? carriedCopies(originals, derived, indentation)
        : carriedCopies(removed, added, indentation),
    });
  }
  copyDocumentParagraphPropertySourceContract(document, source);
  return document;
};

/**
 * The copies that carry, onto each paragraph among `blocks` with an id, the
 * sources of the paragraph `owners` last recorded for that id.
 */
export const ownedSourceCopies = (
  owners: ReadonlyMap<string, Paragraph>,
  blocks: readonly BlockContent[],
): Map<Paragraph, Paragraph> => {
  const owned = new Map<Paragraph, Paragraph>();
  for (const block of blocks) {
    if (block.type !== "paragraph" || block.paraId === undefined) continue;
    const source = owners.get(block.paraId);
    if (source === undefined || block === source) continue;
    const carried = carryParagraphPropertySource({ target: block, source });
    if (carried !== block) owned.set(block, carried);
  }
  return owned;
};
