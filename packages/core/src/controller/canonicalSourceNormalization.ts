/**
 * Source spellings the canonical session reads in its own form, and gives back
 * as read while their story is unedited.
 *
 * Two spellings carry a meaning the session models another way:
 *
 * - A literal tab character in `w:t`, which a reader shows as a tab: the
 *   session's model spells a tab as `w:tab`, so the text is split around it.
 * - A story with no blocks (a header or a footer holding nothing): a story
 *   ends with a paragraph, the one a caret goes in, so the session gives it an
 *   empty one.
 *
 * Both are the session's model only. Each normalized story keeps the blocks it
 * was read with, and a save of that story while it still holds what the
 * session opened with writes those blocks, so the part keeps its source bytes.
 * A story the user edits saves as edited, in the session's spelling.
 */

import { panic } from "better-result";

import type { BlockContent, Document, Paragraph, Run } from "@stll/docx-core";
import {
  documentStories,
  findStoryBody,
  OP_STORIES,
  type OpStory,
  packageParagraphIds,
  replaceStoryBody,
  withBodyContent,
} from "@stll/docx-core/ops";

import { canonicalJson } from "../utils/canonicalJson";

/** A story the session reads in its own form, with the blocks it was read with. */
type NormalizedStory = {
  story: OpStory;
  source: readonly BlockContent[];
};

/** A normalized story, with what it held when the session opened. */
export type CanonicalSourceRestore = NormalizedStory & {
  opened: string;
};

const TAB = "\t";

/** The run content with each literal tab in its text spelled as a tab element. */
const tabSpelledContent = (content: Run["content"]): Run["content"] => {
  const out: Run["content"] = [];
  for (const item of content) {
    if (item.type !== "text" || !item.text.includes(TAB)) {
      out.push(item);
      continue;
    }
    item.text.split(TAB).forEach((part, index) => {
      if (index > 0) out.push({ type: "tab" });
      if (part !== "") out.push({ ...item, text: part });
    });
  }
  return out;
};

const withTabElements = (run: Run): Run =>
  run.content.some((item) => item.type === "text" && item.text.includes(TAB))
    ? { ...run, content: tabSpelledContent(run.content) }
    : run;

const isPlainRecord = (value: object): value is Record<string, unknown> => {
  const prototype: unknown = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
};

/**
 * `value` with every run's literal tabs spelled as tab elements, rebuilding
 * only what changes. A run sits at any depth of inline content (inside a
 * revision, a link, a field result, a content control), so the walk goes
 * through every plain record and array and leaves everything else as it is.
 */
const withRunTabElements = <T>(value: T): T => {
  if (typeof value !== "object" || value === null) return value;
  if (Array.isArray(value)) {
    let changed = false;
    const mapped = value.map((item: unknown) => {
      const next = withRunTabElements(item);
      if (next !== item) changed = true;
      return next;
    });
    // The mapped array holds the same element types as `value`.
    return (changed ? mapped : value) as T;
  }
  if (!isPlainRecord(value)) return value;
  const changes = new Map<string, unknown>();
  for (const [key, child] of Object.entries(value)) {
    const mapped = withRunTabElements(child);
    if (mapped !== child) changes.set(key, mapped);
  }
  // Spread keeps the record's symbol-keyed bindings.
  let next: Record<string, unknown> =
    changes.size === 0 ? value : { ...value, ...Object.fromEntries(changes) };
  if (next["type"] === "run" && Array.isArray(next["content"]))
    // A record whose type is "run" is a run of the model.
    next = withTabElements(next as unknown as Run) as unknown as Record<
      string,
      unknown
    >;
  // The rebuilt record keeps every field `value` has.
  return (next === value ? value : next) as T;
};

/** The first id not in `used`, spelled as eight uppercase hex digits. */
const freshParaId = (used: Set<string>): string => {
  for (let candidate = 1; candidate < 0x80000000; candidate += 1) {
    const id = candidate.toString(16).toUpperCase().padStart(8, "0");
    if (!used.has(id)) {
      used.add(id);
      return id;
    }
  }
  return panic("Every paragraph id is in use.");
};

const isSeparatorNote = (document: Document, story: OpStory): boolean => {
  if (story === OP_STORIES.MAIN) return false;
  if (story.kind !== "footnote" && story.kind !== "endnote") return false;
  const notes =
    story.kind === "footnote"
      ? document.package.footnotes
      : document.package.endnotes;
  const note = notes?.find(({ id }) => id === story.id);
  return note?.noteType !== undefined && note.noteType !== "normal";
};

/** The document in the session's own spelling, with every story it changed. */
export const normalizeCanonicalSource = (
  document: Document,
): { document: Document; normalized: NormalizedStory[] } => {
  let current = document;
  const normalized: NormalizedStory[] = [];
  const used = new Set(
    packageParagraphIds(document.package).map((id) => id.toUpperCase()),
  );
  for (const story of documentStories(document)) {
    const body = findStoryBody(current, story);
    if (body === undefined) continue;
    let content = withRunTabElements(body.content);
    // A separator or continuation note is the application's own, not a story
    // a caret goes in; the session leaves it as read.
    if (content.length === 0 && !isSeparatorNote(document, story)) {
      const paragraph: Paragraph = {
        type: "paragraph",
        paraId: freshParaId(used),
        content: [],
      };
      content = [paragraph];
    }
    if (content === body.content) continue;
    normalized.push({ story, source: body.content });
    current = replaceStoryBody({
      document: current,
      story,
      body: withBodyContent(body, [...content]),
    });
  }
  return { document: current, normalized };
};

/** What each normalized story holds in the document the session opens. */
export const canonicalSourceRestores = (
  document: Document,
  normalized: readonly NormalizedStory[],
): CanonicalSourceRestore[] =>
  normalized.flatMap((entry) => {
    const body = findStoryBody(document, entry.story);
    return body === undefined
      ? []
      : [{ ...entry, opened: canonicalJson(body.content) }];
  });

/** The document with every story that still holds what it opened with given back as read. */
export const withUneditedSourceRestored = (
  document: Document,
  restores: readonly CanonicalSourceRestore[],
): Document => {
  let current = document;
  for (const { story, source, opened } of restores) {
    const body = findStoryBody(current, story);
    if (body === undefined || canonicalJson(body.content) !== opened) continue;
    current = replaceStoryBody({
      document: current,
      story,
      body: withBodyContent(body, [...source]),
    });
  }
  return current;
};
