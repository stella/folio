/**
 * Edit variants: the target side of every benchmark pair.
 *
 * Each variant is a deterministic XML rewrite of a generated package. Building
 * the target by editing the package rather than by driving folio's applier
 * keeps the measurement honest twice over: the target owes nothing to the code
 * under test, so an applier bug cannot manufacture a pair the comparison
 * trivially agrees with, and no engine work happens inside the setup the stage
 * timings exclude.
 *
 * The variants are graded so a run says where cost comes from: `identical`
 * isolates the fixed cost of parsing and serializing, `light` the cost of a
 * realistic review pass, `churn` and `structural` the cost of alignment work,
 * and `rewrite` the cost of applying markup to nearly every block.
 *
 * Selection is by paragraph ordinal, never by position in the body's child
 * list. A body child can be a bookmark marker or a whole table, so choosing
 * "every third child" silently edits every paragraph of one document class and
 * none of another, and the suite then reports coverage it does not have.
 */

import {
  blockText,
  documentPartOf,
  replaceBodyChildren,
  splitBodyChildren,
  withBlockText,
  type DocxPackage,
  type PackagePart,
} from "./package-xml";

export const EDIT_VARIANTS = Object.freeze([
  "identical",
  "light",
  "heavy",
  "churn",
  "reorder",
  "structural",
  "tablecount",
  "numbering",
  "notes",
  "headers",
  "everywhere",
  "rewrite",
] as const);

export type EditVariant = (typeof EDIT_VARIANTS)[number];

/**
 * `w:p` never nests in a generated document, so a non-greedy scan recovers
 * every paragraph at any depth: top level, inside a table cell, inside a
 * nested table. `w:pPr` does not match, since the pattern requires a `>` or a
 * space after the element name.
 */
const PARAGRAPH_PATTERN = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gu;

/** A body child that is itself one paragraph, rather than a table or a marker. */
const isParagraph = (blockXml: string): boolean =>
  blockXml.startsWith("<w:p>") || blockXml.startsWith("<w:p ");

type ParagraphTransform = (paragraphXml: string, ordinal: number) => string;

/** Rewrite every paragraph in `children`, numbering them across the whole body. */
const mapParagraphs = (children: readonly string[], transform: ParagraphTransform): string[] => {
  let ordinal = 0;
  return children.map((child) =>
    child.replaceAll(PARAGRAPH_PATTERN, (paragraph) => transform(paragraph, ordinal++)),
  );
};

/** Replace every nth word of a paragraph's text, leaving the rest identical. */
const editWords = (paragraphXml: string, stride: number, suffix: string): string => {
  const words = blockText(paragraphXml).split(" ");
  if (words.length === 0) {
    return paragraphXml;
  }
  return withBlockText(
    paragraphXml,
    words.map((word, index) => (index % stride === 0 ? `${word}${suffix}` : word)).join(" "),
  );
};

type BodyRewrite = (children: readonly string[]) => string[];

const identical: BodyRewrite = (children) => [...children];

/** One paragraph in fifteen gains a word: the shape a review pass leaves. */
const light: BodyRewrite = (children) =>
  mapParagraphs(children, (paragraph, ordinal) =>
    ordinal % 15 === 3 ? editWords(paragraph, 7, " promptly") : paragraph,
  );

const heavy: BodyRewrite = (children) =>
  mapParagraphs(children, (paragraph, ordinal) =>
    ordinal % 3 === 1 ? editWords(paragraph, 4, " materially") : paragraph,
  );

/**
 * Insertions, deletions and edits interleaved, so no run of unchanged
 * paragraphs lets the alignment coast. Applied to top-level paragraphs only:
 * adding or removing a paragraph inside a cell is the structural variant's
 * job, and mixing the two would stop either from isolating anything.
 */
const churn: BodyRewrite = (children) => {
  const rewritten: string[] = [];
  let ordinal = 0;
  for (const child of children) {
    if (!isParagraph(child)) {
      rewritten.push(child);
      continue;
    }
    const step = ordinal++ % 7;
    if (step === 0) {
      rewritten.push(withBlockText(child, "A newly added obligation of the counterparty."), child);
      continue;
    }
    if (step === 3) {
      continue;
    }
    rewritten.push(step === 5 ? editWords(child, 3, " revised") : child);
  }
  return rewritten;
};

/** Relocate a contiguous run of top-level paragraphs later in the document. */
const reorder: BodyRewrite = (children) => {
  const indexes = children.flatMap((child, index) => (isParagraph(child) ? [index] : []));
  const runLength = Math.max(1, Math.floor(indexes.length / 12));
  const from = Math.floor(indexes.length / 6);
  const anchor = indexes[Math.floor((indexes.length * 2) / 3)];
  const moved = new Set(indexes.slice(from, from + runLength));
  const run = [...moved].map((index) => children[index] ?? "");
  const rewritten: string[] = [];
  for (const [index, child] of children.entries()) {
    if (moved.has(index)) {
      continue;
    }
    if (index === anchor) {
      rewritten.push(...run);
    }
    rewritten.push(child);
  }
  return rewritten;
};

/** Split a paragraph in two at its midpoint word. */
const splitParagraph = (paragraphXml: string): [string, string] => {
  const words = blockText(paragraphXml).split(" ");
  const middle = Math.max(1, Math.floor(words.length / 2));
  return [
    withBlockText(paragraphXml, words.slice(0, middle).join(" ")),
    withBlockText(paragraphXml, words.slice(middle).join(" ")),
  ];
};

const changeListLevel = (paragraphXml: string): string =>
  paragraphXml.replaceAll(
    /<w:ilvl w:val="(\d+)"\/>/gu,
    (_match, level: string) => `<w:ilvl w:val="${String(Math.min(2, Number(level) + 1))}"/>`,
  );

const deleteFirstTableRow = (blockXml: string): string => {
  const open = blockXml.indexOf("<w:tr>");
  if (open === -1) {
    return blockXml;
  }
  const close = blockXml.indexOf("</w:tr>", open);
  return close === -1
    ? blockXml
    : blockXml.slice(0, open) + blockXml.slice(close + "</w:tr>".length);
};

/**
 * Splits, merges, a deleted table row, a changed list level and an added item:
 * the edits that move a paragraph mark rather than the words around it, which
 * is the class a word-level diff gets wrong.
 */
const structural: BodyRewrite = (children) => {
  const rewritten: string[] = [];
  let pendingMerge: string | null = null;
  let ordinal = 0;
  for (const child of children) {
    if (!isParagraph(child)) {
      rewritten.push(deleteFirstTableRow(child));
      continue;
    }
    if (pendingMerge !== null) {
      rewritten.push(withBlockText(pendingMerge, `${blockText(pendingMerge)} ${blockText(child)}`));
      pendingMerge = null;
      ordinal += 1;
      continue;
    }
    const step = ordinal++ % 11;
    if (step === 1) {
      rewritten.push(...splitParagraph(child));
      continue;
    }
    if (step === 4) {
      pendingMerge = child;
      continue;
    }
    if (step === 7) {
      rewritten.push(changeListLevel(child));
      continue;
    }
    if (step === 9) {
      rewritten.push(child, withBlockText(child, "An added item in the same list."));
      continue;
    }
    rewritten.push(child);
  }
  if (pendingMerge !== null) {
    rewritten.push(pendingMerge);
  }
  return rewritten;
};

/**
 * A table removed and a table added: the pair a row-level vocabulary cannot
 * express at all. Applied only to a class that has tables; every other class
 * reports the variant as not applicable rather than pretending to cover it.
 */
const tablecount: BodyRewrite = (children) => {
  const tableIndexes = children.flatMap((child, index) =>
    child.startsWith("<w:tbl>") ? [index] : [],
  );
  const firstTable = tableIndexes.at(0);
  if (firstTable === undefined) {
    return [...children];
  }
  const added =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
    "<w:tr>" +
    '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Added heading</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Added value</w:t></w:r></w:p></w:tc>' +
    "</w:tr></w:tbl>";
  const rewritten = children.filter((_child, index) => index !== firstTable);
  // Before the body's closing paragraph, not after it: a table may not be a
  // body's last child, and a target that breaks the rule would be measuring
  // the engine against malformed input rather than against the edit.
  rewritten.splice(rewritten.length - 1, 0, added);
  return rewritten;
};

const rewrite: BodyRewrite = (children) =>
  mapParagraphs(children, (paragraph) => {
    const text = blockText(paragraph);
    return text.length === 0
      ? paragraph
      : withBlockText(paragraph, `Restated: ${text.split(" ").toReversed().join(" ")}`);
  });

const BODY_REWRITES = {
  identical,
  light,
  heavy,
  churn,
  reorder,
  structural,
  tablecount,
  numbering: identical,
  notes: identical,
  headers: identical,
  everywhere: light,
  rewrite,
} as const satisfies Record<EditVariant, BodyRewrite>;

const NOTE_PARTS = Object.freeze(["word/footnotes.xml", "word/endnotes.xml"] as const);

const NUMBERING_PART = "word/numbering.xml";

/**
 * Renumber the list without touching a word: every level's format becomes
 * `lowerRoman`. Labels are rendered from the numbering definitions, so no
 * block's text changes and a text-only comparison sees nothing at all.
 */
const rewriteNumberingPart = (parts: Map<string, PackagePart>): boolean => {
  const part = parts.get(NUMBERING_PART);
  if (typeof part !== "string") {
    return false;
  }
  const rewritten = part.replaceAll(
    /<w:numFmt w:val="[^"]*"\/>/gu,
    '<w:numFmt w:val="lowerRoman"/>',
  );
  if (rewritten === part) {
    return false;
  }
  parts.set(NUMBERING_PART, rewritten);
  return true;
};

const CHROME_PARTS = Object.freeze(["word/header1.xml", "word/footer1.xml"] as const);

/**
 * The stories that are not the body. `notes` rewrites the footnote and endnote
 * parts, `headers` the header and footer parts, and both leave the main story
 * byte-identical: the pairs that tell a main-story-only engine apart from one
 * that compares every story.
 */
const SECONDARY_STORY_PARTS = {
  notes: NOTE_PARTS,
  headers: CHROME_PARTS,
  /**
   * `everywhere` edits the body as well, so the comparison writes revisions
   * into more than one story in one call. That is the only shape in which two
   * stories can claim the same `w:id`, which is what
   * `revision-ids-are-unique` exists to catch.
   */
  everywhere: Object.freeze([...NOTE_PARTS, ...CHROME_PARTS]),
} as const satisfies Partial<Record<EditVariant, readonly string[]>>;

const rewriteStoryParts = (
  parts: Map<string, PackagePart>,
  names: readonly string[],
  stride: number,
): void => {
  for (const name of names) {
    const part = parts.get(name);
    if (typeof part !== "string") {
      continue;
    }
    let index = 0;
    parts.set(
      name,
      part.replaceAll(/<w:t(?:\s[^>]*?)?>([\s\S]*?)<\/w:t>/gu, (match, text: string) =>
        index++ % stride === 0 ? `<w:t xml:space="preserve">${text} as amended</w:t>` : match,
      ),
    );
  }
};

export type ApplyVariantOptions = {
  parts: DocxPackage;
  variant: EditVariant;
};

/**
 * The target package for one variant, or `null` when the variant does not
 * apply to this document class (a notes edit on a document with no notes would
 * be an identical pair wearing another name, and would quietly inflate the
 * suite's apparent coverage).
 */
export const applyVariant = ({ parts, variant }: ApplyVariantOptions): DocxPackage | null => {
  const target = new Map(parts);
  if (variant === "numbering") {
    return rewriteNumberingPart(target) ? target : null;
  }

  if (variant === "notes" || variant === "headers" || variant === "everywhere") {
    const names = SECONDARY_STORY_PARTS[variant];
    if (!names.some((name) => typeof parts.get(name) === "string")) {
      return null;
    }
    // Every string in a header or footer is short, so a stride of three would
    // leave some of them untouched; the note parts keep theirs.
    rewriteStoryParts(target, names, variant === "notes" ? 3 : 1);
    if (variant !== "everywhere") {
      return target;
    }
  }

  const documentXml = documentPartOf(parts);
  const children = splitBodyChildren(documentXml);
  const rewritten = BODY_REWRITES[variant](children);
  if (variant !== "identical" && rewritten.join("") === children.join("")) {
    return null;
  }
  target.set("word/document.xml", replaceBodyChildren(documentXml, rewritten));
  return target;
};
