/**
 * Link targets over unchanged paragraphs. The content snapshot exposes run
 * formatting but no hyperlink target, so this reads the public document model
 * returned by the reviewer. A paragraph is comparable only when its plain
 * run/link text agrees with the reader's block text and its paraId is stable.
 */

import type { FolioDocxReviewer } from "@stll/folio-core/server";

type Paragraph = Extract<
  ReturnType<FolioDocxReviewer["toDocument"]>["package"]["document"]["content"][number],
  { type: "paragraph" }
>;
type Inline = Paragraph["content"][number];
type Link = Extract<Inline, { type: "hyperlink" }>;
type Run = Extract<Inline, { type: "run" }>;

const isZeroWidthMarker = (type: string): boolean =>
  type === "bookmarkStart" ||
  type === "bookmarkEnd" ||
  type === "commentRangeStart" ||
  type === "commentRangeEnd" ||
  type === "commentReference";

export type LinkedBlock = { text: string; targets: readonly (string | null)[] };
export type LinkSnapshot = ReadonlyMap<string, LinkedBlock>;

const runText = (run: Run): string | null => {
  let text = "";
  for (const content of run.content) {
    if (content.type !== "text") return null;
    text += content.text;
  }
  return text;
};

const linkTarget = (link: Link): string | null => {
  if (link.href !== undefined) return link.href;
  if (link.anchor !== undefined) return `#${link.anchor}`;
  return null;
};

const linkedBlock = (paragraph: Paragraph): LinkedBlock | null => {
  let text = "";
  const targets: (string | null)[] = [];
  for (const item of paragraph.content) {
    if (item.type === "run") {
      const part = runText(item);
      if (part === null) return null;
      text += part;
      targets.push(...Array.from({ length: part.length }, () => null));
      continue;
    }
    if (item.type === "hyperlink") {
      const target = linkTarget(item);
      if (target === null) return null;
      for (const child of item.children) {
        if (isZeroWidthMarker(child.type)) continue;
        if (child.type !== "run") return null;
        const part = runText(child);
        if (part === null) return null;
        text += part;
        targets.push(...Array.from({ length: part.length }, () => target));
      }
      continue;
    }
    if (isZeroWidthMarker(item.type)) continue;
    // Bookmarks, revisions, fields, and wrappers require separate clean-text
    // projection rules. Skipping them avoids assigning links to wrong offsets.
    return null;
  }
  return { text, targets };
};

/** Capture link targets for simple, paraId-backed blocks in the main story. */
export const captureLinks = (reviewer: FolioDocxReviewer): LinkSnapshot => {
  const paragraphs = new Map<string, LinkedBlock>();
  const duplicateIds = new Set<string>();
  const add = (paragraph: Paragraph): void => {
    if (!paragraph.paraId) return;
    if (paragraphs.has(paragraph.paraId)) {
      duplicateIds.add(paragraph.paraId);
      return;
    }
    const linked = linkedBlock(paragraph);
    if (linked) paragraphs.set(paragraph.paraId, linked);
  };
  for (const block of reviewer.toDocument().package.document.content) {
    if (block.type === "paragraph") add(block);
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const child of cell.content) {
            if (child.type === "paragraph") add(child);
          }
        }
      }
    }
  }
  const captured = new Map<string, LinkedBlock>();
  for (const row of reviewer.getContent()) {
    if (duplicateIds.has(row.id)) continue;
    const linked = paragraphs.get(row.id);
    if (linked?.text === row.text) captured.set(row.id, linked);
  }
  return captured;
};

/**
 * Compare every stable block whose text survived unchanged. This covers
 * formatting, style, comment, and neighboring structural edits. Text edits
 * need an explicit link-inheritance rule and are left to a later model.
 */
type ComparePreservedLinksArgs = {
  before: LinkSnapshot;
  after: LinkSnapshot;
  afterRows: readonly { id: string; text: string }[];
};

export const comparePreservedLinks = ({
  before,
  after,
  afterRows,
}: ComparePreservedLinksArgs): { problems: string[]; checked: number } => {
  const problems: string[] = [];
  const readerText = new Map(afterRows.map(({ id, text }) => [id, text]));
  let checked = 0;
  for (const [id, expected] of before) {
    // A real text edit changes this reader text. A missing model projection
    // must not look like a text edit when the reader still shows the old text.
    if (readerText.get(id) !== expected.text) continue;
    const actual = after.get(id);
    checked++;
    if (!actual || actual.text !== expected.text) {
      problems.push(`block ${id} ("${expected.text}") has no matching link projection`);
      continue;
    }
    if (JSON.stringify(actual.targets) !== JSON.stringify(expected.targets)) {
      problems.push(
        `block ${id} ("${expected.text}") has link targets ${JSON.stringify(actual.targets)}, expected ${JSON.stringify(expected.targets)}`,
      );
    }
  }
  return { problems, checked };
};
