/**
 * Selective XML Patch Module
 *
 * Patches only changed paragraphs in document.xml, preserving
 * unchanged content byte-for-byte. Uses string offset tracking
 * with proper tag depth counting (not regex) to handle nested elements.
 */

import {
  cloneElement,
  getAttributeByNamespaceUri,
  findAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespacePrefix,
  getNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";
import { patchBreaksCommentRangeBalance } from "./commentRangeIntegrity";
import { resolveParagraphIdentities } from "./paraIdAttribute";
import { captureVerbatimXml } from "./verbatimCapture";

/**
 * Whether `char` ends an element's tag name in XML — a whitespace separator
 * (space, tab, CR, or LF, all valid before attributes per XML 1.0 §3.1), the
 * tag close `>`, or a self-close `/`. Manual tag scanners must accept every
 * whitespace form, not just a literal space, so newline-formatted markup
 * (`<w:p\n  w14:paraId="…">`) is still recognized as the element rather than
 * mistaken for a longer-named sibling.
 */
export function isXmlNameBoundary(char: string | undefined): boolean {
  return (
    char === " " || char === "\t" || char === "\n" || char === "\r" || char === ">" || char === "/"
  );
}

/**
 * Find the exact string start and end offsets of a <w:p> element
 * identified by its w14:paraId attribute.
 *
 * Handles nested <w:p> elements (e.g. inside mc:AlternateContent)
 * via proper depth counting.
 *
 * Returns null if paraId not found or appears more than once (ambiguous).
 */
export function findParagraphOffsets(
  xml: string,
  paraId: string,
): { start: number; end: number } | null {
  // Find all <w:p elements that contain this paraId.
  // Pattern matches <w:p followed by whitespace or >, then any attrs, then the paraId.
  // This covers all attribute orderings since [^>]* matches any attributes before paraId.
  const escaped = escapeRegExp(paraId);
  const pattern = new RegExp(`<w:p[\\s][^>]*w14:paraId="${escaped}"`, "gu");

  const matches: number[] = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    matches.push(match.index);
  }

  if (matches.length === 0) {
    return null;
  }
  if (matches.length > 1) {
    // Duplicate paraId — ambiguous, cannot safely patch
    return null;
  }

  // SAFETY: matches.length === 1 verified above
  const start = matches[0]!;

  // Now find the matching </w:p> by counting depth
  // Start after the <w:p opening
  let pos = start;
  let depth = 0;

  while (pos < xml.length) {
    // Find next tag
    const tagStart = xml.indexOf("<", pos);
    if (tagStart === -1) {
      break;
    }

    // Check if it's a <w:p or </w:p tag
    if (xml.startsWith("<w:p", tagStart)) {
      const charAfterTag = xml[tagStart + 4];
      // Must be <w:p> or <w:p or <w:p/ (not <w:pPr, <w:pStyle, etc.)
      if (isXmlNameBoundary(charAfterTag)) {
        // Check for self-closing: <w:p ... />
        const tagEnd = xml.indexOf(">", tagStart);
        if (tagEnd === -1) {
          break;
        }

        if (xml[tagEnd - 1] === "/") {
          // Self-closing <w:p ... /> — doesn't change depth
          if (depth === 0) {
            // This IS our paragraph and it's self-closing
            return { start, end: tagEnd + 1 };
          }
        } else {
          depth++;
        }
        pos = tagEnd + 1;
      } else {
        // It's something like <w:pPr — skip
        pos = tagStart + 1;
      }
    } else if (xml.startsWith("</w:p>", tagStart)) {
      depth--;
      if (depth === 0) {
        return { start, end: tagStart + 6 }; // 6 = '</w:p>'.length
      }
      pos = tagStart + 6;
    } else {
      pos = tagStart + 1;
    }
  }

  // Couldn't find matching close tag
  return null;
}

/**
 * Extract the serialized XML for a specific paragraph by paraId
 * from a fully serialized document.xml string.
 */
export function extractParagraphXml(serializedXml: string, paraId: string): string | null {
  const offsets = findParagraphOffsets(serializedXml, paraId);
  if (!offsets) {
    return null;
  }
  return serializedXml.slice(offsets.start, offsets.end);
}

export type ParagraphOffsets = { start: number; end: number };

/**
 * Single linear-scan index of every `<w:p>` element's start/end offsets,
 * keyed by `w14:paraId`. Generalizes {@link findParagraphOffsets} (one
 * regex-scan-plus-depth-walk per lookup) to build every paragraph's offsets
 * in a single pass, so a caller that needs many paragraphs from the same XML
 * (e.g. `collectChangedNoteParaIds` in rezip.ts, walking every note paraId)
 * does O(1) map lookups afterward instead of re-scanning the whole XML once
 * per id — O(ids * XML length) collapses to O(XML length).
 *
 * A `<w:p>` nested inside another (e.g. inside `mc:AlternateContent`) is
 * indexed too via a depth stack, matching every element `findParagraphOffsets`
 * can resolve. A paraId that appears on more than one element is ambiguous —
 * mirroring {@link findParagraphOffsets}'s single-match requirement — and is
 * omitted from the index, so a lookup misses exactly where the per-id
 * function would return null.
 */
export function buildParagraphOffsetIndex(xml: string): Map<string, ParagraphOffsets> {
  const seenCount = new Map<string, number>();
  const ranges = new Map<string, ParagraphOffsets>();
  for (const { start, end, paraId } of scanParagraphs(xml)) {
    if (paraId === undefined) {
      continue;
    }
    seenCount.set(paraId, (seenCount.get(paraId) ?? 0) + 1);
    if (end > start) {
      ranges.set(paraId, { start, end });
    }
  }

  const index = new Map<string, ParagraphOffsets>();
  for (const [id, range] of ranges) {
    if (seenCount.get(id) === 1) {
      index.set(id, range);
    }
  }
  return index;
}

/**
 * The containers a `<w:p>` sits in, as far as a splice has to care.
 *
 * A paragraph is addressed inside one *story*: the main flow of the part (its
 * body, table cells and block-level content controls included), or the text
 * of a text box (`w:txbxContent`, DrawingML or VML). The two are separate
 * ordinal spaces, because the model can legitimately represent a text box
 * differently from the source (a VML box read as DrawingML, a box it cannot
 * read at all) without that saying anything about the main flow.
 *
 * `mc:Fallback` is a copy of its `mc:Choice` for consumers that cannot read
 * the Choice. The model reads the Choice and owns nothing in the Fallback, so
 * a Fallback paragraph is never addressed: it is neither a splice target nor
 * an ordinal, which is also the rule `ensureParaIds` stamps by.
 */
export type ParagraphContainer = {
  /** Inside `mc:Fallback`. */
  inFallback: boolean;
  /** Inside `mc:AlternateContent`, whose Fallback repeats this paragraph. */
  inAlternateContent: boolean;
  /** How many text-box stories (`w:txbxContent`) enclose the paragraph. */
  textBoxDepth: number;
  /** How many table cells (`w:tc`) enclose the paragraph. */
  tableDepth: number;
};

/** One `<w:p>` of a part: its byte range, the paraId its open tag carries, and where it sits. */
export type ScannedParagraph = ParagraphOffsets & {
  paraId: string | undefined;
  container: ParagraphContainer;
};

/** The containers {@link scanParagraphs} tracks, by the literal tag the part writes. */
const TRACKED_CONTAINERS = (
  [
    { name: "mc:Fallback", key: "fallback" },
    { name: "mc:AlternateContent", key: "alternateContent" },
    { name: "w:txbxContent", key: "textBox" },
    { name: "w:tc", key: "tableCell" },
  ] as const
).map(({ name, key }) => ({ key, open: `<${name}`, close: `</${name}>` }));

type TrackedContainerKey = (typeof TRACKED_CONTAINERS)[number]["key"];

/**
 * Every `<w:p>` element of `xml`, in the document order of its opening tags,
 * with the `w14:paraId` that tag carries and the containers around it.
 *
 * Document order is what makes the array an ordinal space: within one story
 * (see {@link ParagraphContainer}), the *n*th paragraph of the source part and
 * the *n*th paragraph of the model's serialization name the same paragraph,
 * which is the only way to address a paragraph the producer gave no id. A
 * `<w:p>` nested inside another (inside `mc:AlternateContent`, a text box) is
 * one entry of its own, exactly as {@link countParagraphElements} counts it.
 * An unterminated paragraph keeps `end === start`: it occupies its ordinal but
 * no splice can be built from it.
 */
export function scanParagraphs(xml: string): ScannedParagraph[] {
  const paragraphs: ScannedParagraph[] = [];
  const open: number[] = [];
  const depth: Record<TrackedContainerKey, number> = {
    fallback: 0,
    alternateContent: 0,
    textBox: 0,
    tableCell: 0,
  };
  let pos = 0;

  scan: while (pos < xml.length) {
    const tagStart = xml.indexOf("<", pos);
    if (tagStart === -1) {
      break;
    }

    if (xml.startsWith("</w:p>", tagStart)) {
      const slot = open.pop();
      const end = tagStart + "</w:p>".length;
      const paragraph = slot === undefined ? undefined : paragraphs[slot];
      if (paragraph) {
        paragraph.end = end;
      }
      pos = end;
      continue;
    }

    for (const { key, open: openLiteral, close } of TRACKED_CONTAINERS) {
      if (xml.startsWith(close, tagStart)) {
        depth[key] = Math.max(0, depth[key] - 1);
        pos = tagStart + close.length;
        continue scan;
      }
      if (
        xml.startsWith(openLiteral, tagStart) &&
        isXmlNameBoundary(xml[tagStart + openLiteral.length])
      ) {
        const tagEnd = xml.indexOf(">", tagStart);
        if (tagEnd === -1) {
          break scan;
        }
        if (xml[tagEnd - 1] !== "/") {
          depth[key] += 1;
        }
        pos = tagEnd + 1;
        continue scan;
      }
    }

    if (!xml.startsWith("<w:p", tagStart) || !isXmlNameBoundary(xml[tagStart + 4])) {
      pos = tagStart + 1;
      continue;
    }

    const tagEnd = xml.indexOf(">", tagStart);
    if (tagEnd === -1) {
      break;
    }
    const openTag = xml.slice(tagStart, tagEnd + 1);
    const paraId = /\bw14:paraId="(?<id>[^"]+)"/u.exec(openTag)?.groups?.["id"];
    const container: ParagraphContainer = {
      inFallback: depth.fallback > 0,
      inAlternateContent: depth.alternateContent > 0,
      textBoxDepth: depth.textBox,
      tableDepth: depth.tableCell,
    };

    if (xml[tagEnd - 1] === "/") {
      // Self-closing <w:p ... /> — resolved immediately, never pushed.
      paragraphs.push({ start: tagStart, end: tagEnd + 1, paraId, container });
    } else {
      open.push(paragraphs.length);
      paragraphs.push({ start: tagStart, end: tagStart, paraId, container });
    }
    pos = tagEnd + 1;
  }

  return paragraphs;
}

/**
 * Count <w:p> elements in an XML string (top-level paragraph count).
 * Counts opening <w:p tags that are NOT self-closing.
 */
export function countParagraphElements(xml: string): number {
  let count = 0;
  const pattern = /<w:p[\s>]/gu;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    // Verify this is actually <w:p and not <w:pPr etc. The regex already
    // required a whitespace-or-`>` boundary, so this rejects only <w:pPr-style
    // names while accepting every whitespace form (space, tab, CR, LF).
    const idx = match.index;
    if (isXmlNameBoundary(xml[idx + 4])) {
      count++;
    }
  }
  return count;
}

/**
 * Find all paraIds in an XML string and return counts (to detect duplicates).
 */
export function collectParaIds(xml: string): Map<string, number> {
  const ids = new Map<string, number>();
  const pattern = /w14:paraId="(?<id>[^"]+)"/gu;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    // SAFETY: named group always present when regex matches
    const id = match.groups!["id"]!;
    ids.set(id, (ids.get(id) ?? 0) + 1);
  }
  return ids;
}

export type PatchValidationResult = {
  safe: boolean;
  reason?: string;
};

/** Paragraph ids folio writes and this module may have to remove again. */
const MINTED_PARA_ID_ATTRIBUTE = /\sw14:(?:para|text)Id="[^"]*"/gu;

/**
 * The replacement for a paragraph whose source open tag carries no paraId.
 *
 * The model always has a key, because folio mints one for every paragraph that
 * arrives without one, and the serializer writes whatever key the model holds.
 * Writing a minted key into the package would upgrade a producer's id-less
 * document to a partly-id'd one on any edit, declare a `w14` namespace the
 * part never had, and turn an id the minter itself calls positional into an
 * authored one the next reader trusts. So the save keeps the author's
 * convention: `ensureParaIds` is the one pass that gives a package ids, and it
 * runs when a host asks for it. An id the source *does* carry (a `w14:textId`
 * beside no paraId) is the author's and stays.
 */
const withoutMintedIds = (paragraphXml: string, sourceOpenTag: string): string => {
  const tagEnd = paragraphXml.indexOf(">");
  if (tagEnd === -1) {
    return paragraphXml;
  }
  const openTag = paragraphXml
    .slice(0, tagEnd + 1)
    .replace(MINTED_PARA_ID_ATTRIBUTE, (attribute) =>
      sourceOpenTag.includes(attribute.trim()) ? attribute : "",
    );
  return openTag + paragraphXml.slice(tagEnd + 1);
};

/** Where each changed paragraph's re-serialized XML goes, or why it cannot. */
type ParagraphRouting =
  | { type: "routed"; splices: XmlSplice[] }
  | { type: "refused"; reason: string };

/** An ordinal space of one part; see {@link ParagraphContainer}. */
type Story = "main" | "text-box";

const storyOf = ({ textBoxDepth }: ParagraphContainer): Story =>
  textBoxDepth > 0 ? "text-box" : "main";

const sameContainer = (a: ParagraphContainer, b: ParagraphContainer): boolean =>
  a.inAlternateContent === b.inAlternateContent &&
  a.textBoxDepth === b.textBoxDepth &&
  a.tableDepth === b.tableDepth;

/** The paragraphs of one part a splice may address, indexed the ways routing asks for them. */
type AddressableParagraphs = {
  stories: ReadonlyMap<Story, readonly ScannedParagraph[]>;
  /** Each paragraph's ordinal within its story. */
  ordinals: ReadonlyMap<ScannedParagraph, number>;
  byParaId: ReadonlyMap<string, readonly ScannedParagraph[]>;
};

const addressableParagraphs = (xml: string): AddressableParagraphs => {
  const stories = new Map<Story, ScannedParagraph[]>();
  const ordinals = new Map<ScannedParagraph, number>();
  const byParaId = new Map<string, ScannedParagraph[]>();
  for (const paragraph of scanParagraphs(xml)) {
    if (paragraph.container.inFallback) {
      continue;
    }
    const story = storyOf(paragraph.container);
    const sequence = stories.get(story) ?? [];
    ordinals.set(paragraph, sequence.length);
    sequence.push(paragraph);
    stories.set(story, sequence);
    if (paragraph.paraId !== undefined) {
      const named = byParaId.get(paragraph.paraId) ?? [];
      named.push(paragraph);
      byParaId.set(paragraph.paraId, named);
    }
  }
  return { stories, ordinals, byParaId };
};

/** The paraIds written on `<w:p>` elements inside `mc:Fallback`. */
const fallbackParaIds = (xml: string): ReadonlySet<string> =>
  new Set(
    scanParagraphs(xml)
      .filter(({ container, paraId }) => container.inFallback && paraId !== undefined)
      .map(({ paraId }) => paraId as string),
  );

/**
 * The prefix of the part's root element. The scan reads WordprocessingML by
 * its conventional `w:` prefix, so a part that binds it to another prefix has
 * no paragraphs the scan can see, and a spliced `w:` paragraph would not
 * resolve under its root.
 */
const rootElementName = (xml: string): string | undefined =>
  /<(?<name>[A-Za-z_][\w.-]*(?::[\w.-]+)?)/u.exec(xml)?.groups?.["name"];

/**
 * The id of the nearest paragraph on one side of `paragraph` in its story that
 * both parts name exactly once: the witness that places it in its story.
 */
const nearestSharedParaId = (
  own: AddressableParagraphs,
  other: AddressableParagraphs,
  paragraph: ScannedParagraph,
  step: -1 | 1,
): string | null => {
  const sequence = own.stories.get(storyOf(paragraph.container)) ?? [];
  for (
    let ordinal = (own.ordinals.get(paragraph) ?? -1) + step;
    ordinal >= 0 && ordinal < sequence.length;
    ordinal += step
  ) {
    const paraId = sequence[ordinal]?.paraId;
    if (
      paraId !== undefined &&
      own.byParaId.get(paraId)?.length === 1 &&
      other.byParaId.get(paraId)?.length === 1
    ) {
      return paraId;
    }
  }
  return null;
};

/**
 * Route every changed paragraph from the model's serialization to its region
 * of the source part.
 *
 * One owner for the question the patch turns on: *which paragraph of the file
 * is this model paragraph?* The answer is local to the paragraph. A paraId the
 * source writes once names it, robust to anything the model does elsewhere; a
 * paragraph the source never named is located by its ordinal within its story,
 * which is sound exactly when {@link resolveParagraphIdentities} says that
 * story lines up. Nothing else in the part has to agree: the model may drop a
 * text box it cannot read or write a Fallback it does not own, and an edit in
 * the main flow is still one paragraph of the main flow.
 *
 * What the routing still refuses is a paragraph it cannot place with
 * certainty, or one whose place changed:
 * - a paraId missing from, or written twice by, either part (among the
 *   paragraphs outside `mc:Fallback`);
 * - an id the source writes only inside `mc:Fallback`, which the model does
 *   not own;
 * - a paragraph inside `mc:AlternateContent`, whose Fallback repeats it and
 *   would contradict the edited Choice;
 * - a paragraph whose container (table-cell nesting, text box, alternate
 *   content) differs between the two parts;
 * - an authored paragraph whose nearest shared neighbours differ, i.e. that
 *   moved within its story;
 * - an id-less paragraph whose story's ordinals do not line up.
 */
const routeChangedParagraphs = (
  originalXml: string,
  serializedXml: string,
  changedIds: ReadonlySet<string>,
): ParagraphRouting => {
  const rootName = rootElementName(originalXml);
  if (rootName !== undefined && !rootName.startsWith("w:")) {
    return { type: "refused", reason: `non-canonical-wordprocessingml-prefix: ${rootName}` };
  }

  const original = addressableParagraphs(originalXml);
  const serialized = addressableParagraphs(serializedXml);
  let originalFallbackIds: ReadonlySet<string> | undefined;
  const storyAlignment = new Map<Story, boolean>();
  const storyAligned = (story: Story): boolean => {
    let aligned = storyAlignment.get(story);
    if (aligned === undefined) {
      aligned = resolveParagraphIdentities({
        sourceParaIds: (original.stories.get(story) ?? []).map(({ paraId }) => paraId),
        serializedParaIds: (serialized.stories.get(story) ?? []).map(({ paraId }) => paraId),
      }).ordinalsAligned;
      storyAlignment.set(story, aligned);
    }
    return aligned;
  };

  const splices: XmlSplice[] = [];
  for (const id of changedIds) {
    const inOriginal = original.byParaId.get(id) ?? [];
    const inSerialized = serialized.byParaId.get(id) ?? [];
    if (inOriginal.length > 1) {
      return { type: "refused", reason: `duplicate-paraId-in-original: ${id}` };
    }
    if (inOriginal.length === 1 && inSerialized.length === 0) {
      return { type: "refused", reason: `paraId-not-found-in-serialized: ${id}` };
    }
    if (inSerialized.length === 0) {
      return { type: "refused", reason: `paraId-not-found-in-original: ${id}` };
    }
    if (inSerialized.length > 1) {
      return { type: "refused", reason: `duplicate-paraId-in-serialized: ${id}` };
    }
    // SAFETY: inSerialized.length === 1 verified above
    const replacement = inSerialized[0]!;
    if (replacement.end <= replacement.start) {
      return { type: "refused", reason: `unterminated-paragraph: ${id}` };
    }
    const newXml = serializedXml.slice(replacement.start, replacement.end);

    const authored = inOriginal[0];
    let source: ScannedParagraph;
    if (authored) {
      source = authored;
    } else {
      originalFallbackIds ??= fallbackParaIds(originalXml);
      if (originalFallbackIds.has(id)) {
        return { type: "refused", reason: `paraId-only-in-fallback: ${id}` };
      }
      // The model minted this id: the paragraph is its ordinal in its story.
      const story = storyOf(replacement.container);
      if (!storyAligned(story)) {
        return { type: "refused", reason: `unaligned-paragraph-ordinals: ${id}` };
      }
      const positional = original.stories.get(story)?.[serialized.ordinals.get(replacement) ?? -1];
      if (!positional) {
        return { type: "refused", reason: `paraId-not-found-in-original: ${id}` };
      }
      source = positional;
    }

    if (source.end <= source.start) {
      return { type: "refused", reason: `unterminated-paragraph: ${id}` };
    }
    if (source.container.inAlternateContent) {
      return { type: "refused", reason: `paragraph-in-alternate-content: ${id}` };
    }
    if (!sameContainer(source.container, replacement.container)) {
      return { type: "refused", reason: `container-changed: ${id}` };
    }
    if (
      authored &&
      (nearestSharedParaId(original, serialized, source, -1) !==
        nearestSharedParaId(serialized, original, replacement, -1) ||
        nearestSharedParaId(original, serialized, source, 1) !==
          nearestSharedParaId(serialized, original, replacement, 1))
    ) {
      return { type: "refused", reason: `paragraph-moved: ${id}` };
    }

    const sourceXml = originalXml.slice(source.start, source.end);
    splices.push({
      start: source.start,
      end: source.end,
      newXml: authored ? newXml : withoutMintedIds(newXml, sourceXml),
    });
  }

  return { type: "routed", splices };
};

/**
 * Validate that a selective patch can be safely applied: every changed
 * paragraph routes to exactly one region of the original part (see
 * {@link routeChangedParagraphs} for what is refused and why).
 *
 * The rest of the part does not have to match the model's serialization. The
 * splice keeps every unchanged byte of the source, so a paragraph the model
 * represents differently elsewhere (a text box it re-reads, a Fallback it
 * does not own) is simply kept as the source wrote it.
 */
export function validatePatchSafety(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): PatchValidationResult {
  if (changedIds.size === 0) {
    return { safe: true };
  }

  const routing = routeChangedParagraphs(originalXml, serializedXml, changedIds);
  return routing.type === "refused" ? { safe: false, reason: routing.reason } : { safe: true };
}

/**
 * Build a patched document.xml by splicing new paragraph XML into
 * the original at the correct offsets. Only changed paragraphs
 * are replaced; everything else is preserved byte-for-byte.
 *
 * Returns null when a changed paragraph cannot be routed (see
 * {@link validatePatchSafety}) or {@link spliceXml} refuses the result; the
 * caller then falls back to a full repack, whose parts are all re-serialized
 * from one model.
 */
export function buildPatchedDocumentXml(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  if (changedIds.size === 0) {
    return originalXml;
  }
  const routing = routeChangedParagraphs(originalXml, serializedXml, changedIds);
  return routing.type === "refused" ? null : spliceXml(originalXml, routing.splices);
}

/**
 * Build a patched note part (word/footnotes.xml / word/endnotes.xml) by
 * splicing edited note paragraphs into the original, preserving unchanged
 * content byte-for-byte.
 *
 * The routing is the document's: the model only retains the normal notes, so
 * the serialized note XML omits the separator / continuationSeparator
 * paragraphs the original part still carries, and splicing by `paraId` keeps
 * those separators and every unedited note byte-exact while replacing only the
 * edited ones. (An id-less note paragraph does not route here, because the
 * separators misalign its story's ordinals; {@link buildPatchedNotePartXml}
 * addresses it by note instead.)
 *
 * Returns null if any changed id is missing or ambiguous in either input, so
 * the caller can fall back to preserving the original part verbatim.
 */
export function buildPatchedNoteXml(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  return buildPatchedDocumentXml(originalXml, serializedXml, changedIds);
}

/** One region of a part replaced by re-serialized XML. */
export type XmlSplice = { start: number; end: number; newXml: string };

/**
 * Apply `splices` to `xml`, end-to-start so earlier offsets stay valid.
 *
 * Every selective patch is a splice, and every splice goes through here, so
 * the refusal below is a property of the operation rather than a check each
 * site has to remember. A patch rewrites the regions an edit touched and keeps
 * the rest of the part byte-for-byte, so it can write half a comment range:
 * invalid OOXML that anchors the comment to nothing. Answers null when it
 * would, leaving the caller to rewrite a wider region — ultimately the whole
 * part from the model, which is balanced with itself.
 */
export const spliceXml = (xml: string, splices: readonly XmlSplice[]): string | null => {
  let result = xml;
  for (const { start, end, newXml } of [...splices].toSorted((a, b) => b.start - a.start)) {
    result = result.slice(0, start) + newXml + result.slice(end);
  }
  return patchBreaksCommentRangeBalance(xml, result) ? null : result;
};

// ============================================================================
// NUMBERING DEFINITION PATCHING (word/numbering.xml)
// ============================================================================
//
// Numbering definitions carry no `w14:paraId`; a `w:abstractNum` is keyed by its
// `w:abstractNumId` attribute and a `w:num` by its `w:numId`. These helpers
// mirror the paragraph splice above but target whole definition elements by id,
// so an edited list definition is re-emitted from the model while every
// untouched definition (and the parts the model omits — `w:nsid`, `w:tmpl`,
// `w:numPicBullet`, unmodeled level sub-elements) stays byte-exact.

/**
 * Depth-count the end of the element that opens at `start` (an `<openLiteral…>`
 * offset), returning its full range. Nested same-name opens increment depth;
 * the matching `closeTag` (or a self-close) at depth 0 ends it. The boundary
 * check skips longer-named siblings (`<w:numFmt>` when scanning `<w:num`).
 */
function scanElementRange(
  xml: string,
  start: number,
  openLiteral: string,
  closeTag: string,
): { start: number; end: number } | null {
  const afterOpenIndex = openLiteral.length;
  let pos = start;
  let depth = 0;
  while (pos < xml.length) {
    const tagStart = xml.indexOf("<", pos);
    if (tagStart === -1) {
      break;
    }
    if (xml.startsWith(openLiteral, tagStart)) {
      if (!isXmlNameBoundary(xml[tagStart + afterOpenIndex])) {
        pos = tagStart + 1;
        continue;
      }
      const tagEnd = xml.indexOf(">", tagStart);
      if (tagEnd === -1) {
        break;
      }
      if (xml[tagEnd - 1] === "/") {
        if (depth === 0) {
          return { start, end: tagEnd + 1 };
        }
      } else {
        depth++;
      }
      pos = tagEnd + 1;
    } else if (xml.startsWith(closeTag, tagStart)) {
      depth--;
      if (depth === 0) {
        return { start, end: tagStart + closeTag.length };
      }
      pos = tagStart + closeTag.length;
    } else {
      pos = tagStart + 1;
    }
  }
  return null;
}

/**
 * Find the exact start/end offsets of the element `<openLiteral … idAttr="id">`,
 * depth-counting its matching close tag. Returns null when the id is absent or
 * ambiguous (appears more than once).
 */
function findElementByIdAttr(
  xml: string,
  openLiteral: string,
  closeTag: string,
  idAttr: string,
  id: string,
): { start: number; end: number } | null {
  const pattern = new RegExp(
    `${escapeRegExp(openLiteral)}[\\s][^>]*${escapeRegExp(idAttr)}="${escapeRegExp(id)}"`,
    "gu",
  );
  const matches: number[] = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    matches.push(match.index);
  }
  if (matches.length !== 1) {
    return null;
  }
  // SAFETY: matches.length === 1 verified above
  return scanElementRange(xml, matches[0]!, openLiteral, closeTag);
}

/**
 * Extract the serialized XML for `<openLiteral … idAttr="id">…`, or null when it
 * cannot be resolved uniquely.
 */
function extractElementByIdAttr(
  xml: string,
  openLiteral: string,
  closeTag: string,
  idAttr: string,
  id: string,
): string | null {
  const offsets = findElementByIdAttr(xml, openLiteral, closeTag, idAttr, id);
  return offsets ? xml.slice(offsets.start, offsets.end) : null;
}

type NoteElementName = "footnote" | "endnote";

type NoteElementSyntax = {
  elementName: string;
  idAttributeName: string;
  elementPrefix: string;
  attributePrefix: string;
};

const qualifiedName = (prefix: string, localName: string): string =>
  prefix.length === 0 ? localName : `${prefix}:${localName}`;

const collectNoteElementSyntax = (
  xml: string,
  elementName: NoteElementName,
): Map<string, NoteElementSyntax[]> => {
  const byId = new Map<string, NoteElementSyntax[]>();
  const root = parseXmlDocument(xml);
  for (const element of getChildElements(root)) {
    if (
      getLocalName(element.name) !== elementName ||
      !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
    ) {
      continue;
    }
    const idAttribute = findAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, "id");
    if (!element.name || !idAttribute) {
      continue;
    }
    const syntax = {
      elementName: element.name,
      idAttributeName: idAttribute.name,
      elementPrefix: getNamespacePrefix(element.name) ?? "",
      attributePrefix: getNamespacePrefix(idAttribute.name) ?? "",
    };
    const entries = byId.get(idAttribute.value);
    if (entries) {
      entries.push(syntax);
    } else {
      byId.set(idAttribute.value, [syntax]);
    }
  }
  return byId;
};

const syntaxLiteral = (syntax: NoteElementSyntax) => ({
  openLiteral: `<${syntax.elementName}`,
  closeTag: `</${syntax.elementName}>`,
  idAttr: syntax.idAttributeName,
});

const extractNoteElement = (xml: string, syntax: NoteElementSyntax, id: string): string | null => {
  const { openLiteral, closeTag, idAttr } = syntaxLiteral(syntax);
  return extractElementByIdAttr(xml, openLiteral, closeTag, idAttr, id);
};

const findNoteElement = (
  xml: string,
  syntax: NoteElementSyntax,
  id: string,
): ParagraphOffsets | null => {
  const { openLiteral, closeTag, idAttr } = syntaxLiteral(syntax);
  return findElementByIdAttr(xml, openLiteral, closeTag, idAttr, id);
};

type WordPrefixMapping = {
  source: NoteElementSyntax;
  target: NoteElementSyntax;
  sourceXmlnsDeclarations: Record<string, string>;
};

const rewriteWordprocessingPrefixes = (
  xml: string,
  { source, target, sourceXmlnsDeclarations }: WordPrefixMapping,
): string => {
  let rewritten = "";
  let quote: '"' | "'" | null = null;
  let insideTag = false;
  for (let index = 0; index < xml.length; index++) {
    const character = xml[index];
    if (!insideTag) {
      rewritten += character;
      insideTag = character === "<";
      continue;
    }
    if (quote) {
      rewritten += character;
      if (character === quote) {
        quote = null;
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      rewritten += character;
      continue;
    }
    if (character === ">") {
      insideTag = false;
      rewritten += character;
      continue;
    }

    const previous = xml[index - 1];
    const isElementName = previous === "<" || (previous === "/" && xml[index - 2] === "<");
    const sourcePrefix = isElementName ? source.elementPrefix : source.attributePrefix;
    if (sourcePrefix.length > 0 && xml.startsWith(`${sourcePrefix}:`, index)) {
      const targetPrefix = isElementName ? target.elementPrefix : target.attributePrefix;
      rewritten += targetPrefix.length === 0 ? "" : `${targetPrefix}:`;
      index += sourcePrefix.length;
      continue;
    }
    rewritten += character;
  }
  return withXmlnsDeclarations(rewritten, sourceXmlnsDeclarations);
};

const paragraphRanges = (xml: string, wordPrefix: string): ParagraphOffsets[] => {
  const ranges: ParagraphOffsets[] = [];
  const paragraphName = qualifiedName(wordPrefix, "p");
  const openLiteral = `<${paragraphName}`;
  const closeTag = `</${paragraphName}>`;
  let pos = 0;
  while (pos < xml.length) {
    const start = xml.indexOf(openLiteral, pos);
    if (start === -1) {
      break;
    }
    if (!isXmlNameBoundary(xml[start + openLiteral.length])) {
      pos = start + 1;
      continue;
    }
    const range = scanElementRange(xml, start, openLiteral, closeTag);
    if (!range) {
      break;
    }
    ranges.push(range);
    pos = range.end;
  }
  return ranges;
};

export const collectChangedNoteParaIds = (baselineXml: string, currentXml: string): Set<string> => {
  const changed = new Set<string>();
  const baselineIds = collectParaIds(baselineXml);
  const baselineOffsets = buildParagraphOffsetIndex(baselineXml);
  const currentOffsets = buildParagraphOffsetIndex(currentXml);
  for (const [id, count] of collectParaIds(currentXml)) {
    if (count !== 1 || baselineIds.get(id) !== 1) {
      continue;
    }
    const before = baselineOffsets.get(id);
    const after = currentOffsets.get(id);
    if (
      before &&
      after &&
      baselineXml.slice(before.start, before.end) !== currentXml.slice(after.start, after.end)
    ) {
      changed.add(id);
    }
  }
  return changed;
};

type BuildPatchedNotePartXmlOptions = {
  originalXml: string;
  baselineXml: string;
  serializedXml: string;
  replacementXml: string;
  elementName: NoteElementName;
  changedParaIds?: ReadonlySet<string>;
};

/** Why a note part could not be patched; see {@link NotePartPatch}. */
export type NotePartPatchRefusal =
  /** A note element, or a changed paragraph, was missing or ambiguous. */
  | "unroutable-paragraph"
  /** Even rewriting whole notes would leave a comment range with one half. */
  | "comment-range-balance";

/**
 * What patching a note part produced. `refused` is not a failure to serialize:
 * the caller writes the part from the model instead, the note-part reading of
 * the fall back to a full repack the document story takes when its own splice
 * is refused.
 */
export type NotePartPatch =
  | { type: "patched"; xml: string }
  | { type: "refused"; reason: NotePartPatchRefusal };

/**
 * Patch an existing note part from its model serialization.
 *
 * Dirty paragraph ids locate their owning note in the model serialization;
 * `(note w:id, paragraph ordinal)` then locates the corresponding source XML
 * even when the producer omitted paragraph ids. Equal-shape edits replace only
 * dirty paragraphs. A tracked paragraph-break resolution can change that
 * shape, so it replaces the one affected note. Separator notes, unrelated
 * notes, and unaffected equal-shape paragraphs remain byte-exact.
 * `replacementXml` also supplies synthesized automatic note-reference marks,
 * which the parsed model intentionally omits.
 *
 * A comment can be anchored on a note's own text, so its range spans that
 * note's paragraphs and an edit inside the span moves a range half from one
 * paragraph to another. Replacing only the dirty paragraph then drops the half
 * it held and leaves the other standing, so {@link spliceXml} refuses the
 * result and the changed notes are rewritten whole instead — model content on
 * both sides of the range, with every other note still byte-exact.
 */
export function buildPatchedNotePartXml({
  originalXml,
  baselineXml,
  serializedXml,
  replacementXml,
  elementName,
  changedParaIds,
}: BuildPatchedNotePartXmlOptions): NotePartPatch {
  const currentElements = collectNoteElementSyntax(serializedXml, elementName);
  const originalElements = collectNoteElementSyntax(originalXml, elementName);
  const replacementElements = collectNoteElementSyntax(replacementXml, elementName);
  const replacementXmlnsDeclarations = collectXmlnsFromOpeningTag(replacementXml);
  const paragraphSplices: XmlSplice[] = [];
  /** The same edits at note granularity, should the paragraph splices be refused. */
  const noteSplices: XmlSplice[] = [];
  const serializedParaIds = collectParaIds(serializedXml);
  const effectiveChangedParaIds =
    changedParaIds ?? collectChangedNoteParaIds(baselineXml, serializedXml);
  const unroutedChangedParaIds = new Set(
    [...effectiveChangedParaIds].filter((paraId) => serializedParaIds.has(paraId)),
  );

  for (const [id, currentSyntaxEntries] of currentElements) {
    const originalSyntaxEntries = originalElements.get(id);
    const replacementSyntaxEntries = replacementElements.get(id);
    if (
      currentSyntaxEntries.length !== 1 ||
      originalSyntaxEntries?.length !== 1 ||
      replacementSyntaxEntries?.length !== 1
    ) {
      return { type: "refused", reason: "unroutable-paragraph" };
    }
    const currentSyntax = currentSyntaxEntries[0];
    const originalSyntax = originalSyntaxEntries[0];
    const replacementSyntax = replacementSyntaxEntries[0];
    if (!currentSyntax || !originalSyntax || !replacementSyntax) {
      return { type: "refused", reason: "unroutable-paragraph" };
    }
    const currentNote = extractNoteElement(serializedXml, currentSyntax, id);
    const originalOffsets = findNoteElement(originalXml, originalSyntax, id);
    const replacementNote = extractNoteElement(replacementXml, replacementSyntax, id);
    if (!currentNote || !originalOffsets || !replacementNote) {
      return { type: "refused", reason: "unroutable-paragraph" };
    }
    const originalNote = originalXml.slice(originalOffsets.start, originalOffsets.end);
    const currentParagraphs = paragraphRanges(currentNote, currentSyntax.elementPrefix);
    const originalParagraphs = paragraphRanges(originalNote, originalSyntax.elementPrefix);
    const replacementParagraphs = paragraphRanges(replacementNote, replacementSyntax.elementPrefix);
    const noteChangedParaIds = [...collectParaIds(currentNote).keys()].filter((paraId) =>
      unroutedChangedParaIds.has(paraId),
    );
    if (noteChangedParaIds.length === 0) {
      continue;
    }
    const wholeNoteSplice: XmlSplice = {
      start: originalOffsets.start,
      end: originalOffsets.end,
      newXml: rewriteWordprocessingPrefixes(replacementNote, {
        source: replacementSyntax,
        target: originalSyntax,
        sourceXmlnsDeclarations: replacementXmlnsDeclarations,
      }),
    };
    noteSplices.push(wholeNoteSplice);
    if (
      currentParagraphs.length !== originalParagraphs.length ||
      currentParagraphs.length !== replacementParagraphs.length
    ) {
      for (const paraId of noteChangedParaIds) {
        unroutedChangedParaIds.delete(paraId);
      }
      paragraphSplices.push(wholeNoteSplice);
      continue;
    }

    for (let index = 0; index < currentParagraphs.length; index++) {
      const currentRange = currentParagraphs[index];
      const originalRange = originalParagraphs[index];
      const replacementRange = replacementParagraphs[index];
      if (!currentRange || !originalRange || !replacementRange) {
        return { type: "refused", reason: "unroutable-paragraph" };
      }
      const currentParagraph = currentNote.slice(currentRange.start, currentRange.end);
      const routedIds = [...collectParaIds(currentParagraph).keys()].filter((paraId) =>
        unroutedChangedParaIds.has(paraId),
      );
      if (routedIds.length === 0) {
        continue;
      }
      for (const paraId of routedIds) {
        unroutedChangedParaIds.delete(paraId);
      }
      paragraphSplices.push({
        start: originalOffsets.start + originalRange.start,
        end: originalOffsets.start + originalRange.end,
        newXml: rewriteWordprocessingPrefixes(
          replacementNote.slice(replacementRange.start, replacementRange.end),
          {
            source: replacementSyntax,
            target: originalSyntax,
            sourceXmlnsDeclarations: replacementXmlnsDeclarations,
          },
        ),
      });
    }
  }

  if (unroutedChangedParaIds.size > 0) {
    return { type: "refused", reason: "unroutable-paragraph" };
  }
  const patched = spliceXml(originalXml, paragraphSplices);
  if (patched !== null) {
    return { type: "patched", xml: patched };
  }
  const wholeNotes = spliceXml(originalXml, noteSplices);
  return wholeNotes === null
    ? { type: "refused", reason: "comment-range-balance" }
    : { type: "patched", xml: wholeNotes };
}

/**
 * The full range of the first `<openLiteral …>…</closeTag>` element, or null.
 * Used to locate an unkeyed sub-element (a level's `mc:AlternateContent`).
 */
function findFirstElement(
  xml: string,
  openLiteral: string,
  closeTag: string,
): { start: number; end: number } | null {
  let pos = 0;
  while (pos < xml.length) {
    const idx = xml.indexOf(openLiteral, pos);
    if (idx === -1) {
      return null;
    }
    if (isXmlNameBoundary(xml[idx + openLiteral.length])) {
      return scanElementRange(xml, idx, openLiteral, closeTag);
    }
    pos = idx + 1;
  }
  return null;
}

/**
 * Collect the ids carried by opening `<openLiteral … idAttr="…">` tags, with
 * counts so duplicates can be rejected. The `[\s]` after the literal restricts
 * matches to the exact element (excluding `<w:numFmt` when scanning `<w:num`).
 */
function collectElementIds(xml: string, openLiteral: string, idAttr: string): Map<string, number> {
  const ids = new Map<string, number>();
  const pattern = new RegExp(
    `${escapeRegExp(openLiteral)}[\\s][^>]*?${escapeRegExp(idAttr)}="(?<id>[^"]+)"`,
    "gu",
  );
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    // SAFETY: named group `id` always present when the regex matches
    const id = match.groups!["id"]!;
    ids.set(id, (ids.get(id) ?? 0) + 1);
  }
  return ids;
}

type NumberingElementKind = "abstractNum" | "num";

// Canonical-prefix parts use byte-preserving splices; other bindings use the namespace-aware fallback.
const NUMBERING_OPEN_LITERAL: Record<NumberingElementKind, string> = {
  abstractNum: "<w:abstractNum",
  num: "<w:num",
};

const NUMBERING_CLOSE_TAG: Record<NumberingElementKind, string> = {
  abstractNum: "</w:abstractNum>",
  num: "</w:num>",
};

const NUMBERING_ID_ATTR: Record<NumberingElementKind, string> = {
  abstractNum: "w:abstractNumId",
  num: "w:numId",
};

const LEVEL_OPEN_LITERAL = "<w:lvl";
const LEVEL_CLOSE_TAG = "</w:lvl>";
const LEVEL_ID_ATTR = "w:ilvl";

function findNumberingElementOffsets(
  xml: string,
  kind: NumberingElementKind,
  id: string,
): { start: number; end: number } | null {
  return findElementByIdAttr(
    xml,
    NUMBERING_OPEN_LITERAL[kind],
    NUMBERING_CLOSE_TAG[kind],
    NUMBERING_ID_ATTR[kind],
    id,
  );
}

function extractNumberingElementXml(
  xml: string,
  kind: NumberingElementKind,
  id: string,
): string | null {
  return extractElementByIdAttr(
    xml,
    NUMBERING_OPEN_LITERAL[kind],
    NUMBERING_CLOSE_TAG[kind],
    NUMBERING_ID_ATTR[kind],
    id,
  );
}

type ElementOffsets = { start: number; end: number };

function buildNumberingElementOffsetIndex(
  xml: string,
  kind: NumberingElementKind,
): Map<string, ElementOffsets | null> {
  const index = new Map<string, ElementOffsets | null>();
  const openLiteral = NUMBERING_OPEN_LITERAL[kind];
  const pattern = new RegExp(
    `${escapeRegExp(openLiteral)}[\\s][^>]*?${escapeRegExp(NUMBERING_ID_ATTR[kind])}="(?<id>[^"]+)"`,
    "gu",
  );
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    // SAFETY: named group `id` always exists when this pattern matches.
    const id = match.groups!["id"]!;
    if (index.has(id)) {
      index.set(id, null);
      continue;
    }
    index.set(id, scanElementRange(xml, match.index, openLiteral, NUMBERING_CLOSE_TAG[kind]));
  }
  return index;
}

// A custom number format, as the model re-emits it. Valid OOXML on its own —
// the model used to mint a synthetic `decimalZero{3,4,5}` here, which was not —
// but it is the flattening of whatever the source wrote, so `restoreLevelNumFmts`
// puts the source's own element back when the format has not changed.
const CUSTOM_NUM_FMT_PATTERN = /<w:numFmt w:val="custom"(?: w:format="(?<format>[^"]*)")?\/>/u;

export type ChangedNumberingDefs = {
  abstractNums: Set<string>;
  nums: Set<string>;
};

/**
 * The numbering definitions whose current serialization differs from the
 * baseline (re-parsed original) serialization — the ones actually edited. An id
 * is only considered when it resolves uniquely in BOTH inputs, so a definition
 * added or removed relative to the original (which cannot be spliced by id) is
 * left out; both save paths defer those to the byte-exact original part.
 */
export function collectChangedNumberingDefs(
  baselineXml: string,
  currentXml: string,
): ChangedNumberingDefs {
  const changedForKind = (kind: NumberingElementKind): Set<string> => {
    const changed = new Set<string>();
    const baselineIndex = buildNumberingElementOffsetIndex(baselineXml, kind);
    const currentIndex = buildNumberingElementOffsetIndex(currentXml, kind);
    for (const [id, afterRange] of currentIndex) {
      const beforeRange = baselineIndex.get(id);
      if (!beforeRange || !afterRange) {
        continue;
      }
      const before = baselineXml.slice(beforeRange.start, beforeRange.end);
      const after = currentXml.slice(afterRange.start, afterRange.end);
      if (before !== after) {
        changed.add(id);
      }
    }
    return changed;
  };
  return { abstractNums: changedForKind("abstractNum"), nums: changedForKind("num") };
}

/**
 * The element representing a level's number format in the source XML: the
 * `<mc:AlternateContent>` block Word wraps a custom format in, else a
 * self-closing `<w:numFmt …/>` (covers `<w:numFmt w:val="custom" w:format=…/>`).
 */
function extractLevelNumFmtElement(levelXml: string): string | null {
  const alt = findFirstElement(levelXml, "<mc:AlternateContent", "</mc:AlternateContent>");
  if (alt) {
    return levelXml.slice(alt.start, alt.end);
  }
  const match = /<w:numFmt\b[^>]*\/>/u.exec(levelXml);
  return match ? match[0] : null;
}

/**
 * The `@w:format` an original level's number-format element declares, or null
 * when it declares no custom format at all. Telling a preserved format from an
 * edited one is what keeps a genuine change from being reverted.
 */
function originalCustomNumFmtFormat(numFmtElement: string): string | null {
  for (const match of numFmtElement.matchAll(/<w:numFmt\b[^>]*\/>/gu)) {
    const tag = match[0];
    if (!/\bw:val="custom"/u.test(tag)) {
      continue;
    }
    return /\bw:format="(?<format>[^"]*)"/u.exec(tag)?.groups?.["format"] ?? null;
  }
  return null;
}

/**
 * Collect the `xmlns` / `xmlns:*` declarations from an element's opening tag
 * (the substring up to its first `>`).
 */
function collectXmlnsFromOpeningTag(elementXml: string): Record<string, string> {
  const out: Record<string, string> = {};
  const tagEnd = elementXml.indexOf(">");
  const openTag = tagEnd === -1 ? elementXml : elementXml.slice(0, tagEnd);
  const pattern = /\s(?<name>xmlns(?::[\w.-]+)?)="(?<uri>[^"]*)"/gu;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(openTag)) !== null) {
    // SAFETY: named groups `name`/`uri` present when the regex matches
    out[match.groups!["name"]!] = match.groups!["uri"]!;
  }
  return out;
}

/**
 * Carry the given `xmlns` declarations onto `fragmentXml`'s root element so it
 * resolves every prefix even after the ancestor that declared them is replaced.
 * Declarations the fragment already carries win (mirrors the VML
 * `cloneWithXmlnsDeclarations` self-contained-clone behaviour, in the string
 * domain the byte-exact splice needs).
 */
function withXmlnsDeclarations(fragmentXml: string, xmlnsDecls: Record<string, string>): string {
  const own = collectXmlnsFromOpeningTag(fragmentXml);
  const additions = Object.entries(xmlnsDecls)
    .filter(([name]) => !(name in own))
    .map(([name, uri]) => ` ${name}="${uri}"`);
  if (additions.length === 0) {
    return fragmentXml;
  }
  let nameEnd = 1; // skip the opening "<"
  while (nameEnd < fragmentXml.length && !isXmlNameBoundary(fragmentXml[nameEnd])) {
    nameEnd += 1;
  }
  return fragmentXml.slice(0, nameEnd) + additions.join("") + fragmentXml.slice(nameEnd);
}

/**
 * Swap each re-serialized custom numFmt back to the original level's
 * number-format element.
 *
 * The model holds a custom format as `custom` plus its `@w:format`, which the
 * serializer re-emits as a bare `<w:numFmt w:val="custom" w:format="…"/>`. That
 * is valid OOXML and reparses to the same model, but it is not what the source
 * wrote: Word wraps a custom format in an `mc:AlternateContent` whose Fallback
 * carries a plain format for pre-w14 readers, and the model has no field for
 * the Fallback. When an unrelated field of the definition is edited the whole
 * definition is re-serialized, so without this pass that wrapper would be
 * flattened away. Restore it by ilvl only when the model's format still matches
 * the original's; a level whose format was actually edited keeps the model's
 * own element, which is the edit.
 */
function restoreLevelNumFmts(originalDefXml: string, currentDefXml: string): string {
  const replacements: XmlSplice[] = [];
  for (const [ilvl, count] of collectElementIds(currentDefXml, LEVEL_OPEN_LITERAL, LEVEL_ID_ATTR)) {
    if (count !== 1) {
      continue;
    }
    const curOffsets = findElementByIdAttr(
      currentDefXml,
      LEVEL_OPEN_LITERAL,
      LEVEL_CLOSE_TAG,
      LEVEL_ID_ATTR,
      ilvl,
    );
    if (!curOffsets) {
      continue;
    }
    const curLevel = currentDefXml.slice(curOffsets.start, curOffsets.end);
    const custom = CUSTOM_NUM_FMT_PATTERN.exec(curLevel);
    if (!custom) {
      continue;
    }
    const origLevel = extractElementByIdAttr(
      originalDefXml,
      LEVEL_OPEN_LITERAL,
      LEVEL_CLOSE_TAG,
      LEVEL_ID_ATTR,
      ilvl,
    );
    const currentFormat = custom.groups?.["format"] ?? null;
    const original = origLevel ? extractLevelNumFmtElement(origLevel) : null;
    if (!original || !origLevel || originalCustomNumFmtFormat(original) !== currentFormat) {
      // The format was edited, or the original cannot be resolved: the model's
      // own element already says `custom` with the current format, which is
      // valid OOXML and is the edit.
      continue;
    }
    // Format unchanged — restore the original element verbatim. It (e.g.
    // mc:AlternateContent) may use a prefix bound on the replaced ancestors
    // (w:abstractNum / w:lvl) instead of the numbering root; carry those
    // declarations onto it so it stays resolvable.
    const ancestorXmlns = {
      ...collectXmlnsFromOpeningTag(originalDefXml),
      ...collectXmlnsFromOpeningTag(origLevel),
    };
    const restoredLevel = spliceXml(curLevel, [
      {
        start: custom.index,
        end: custom.index + custom[0].length,
        newXml: withXmlnsDeclarations(original, ancestorXmlns),
      },
    ]);
    if (restoredLevel === null) {
      continue;
    }
    replacements.push({ start: curOffsets.start, end: curOffsets.end, newXml: restoredLevel });
  }

  // A refused restore leaves the model's own serialization of the definition,
  // which is what this pass was correcting; that is the conservative answer.
  return spliceXml(currentDefXml, replacements) ?? currentDefXml;
}

/**
 * Build a patched `word/numbering.xml` by splicing the changed `w:abstractNum` /
 * `w:num` definitions from `currentXml` (the model's serialization) into
 * `originalXml`, preserving every other byte. Returns the original unchanged
 * when nothing changed, or null when a changed id cannot be resolved uniquely
 * in either input (so the caller preserves the original part verbatim).
 */
export function buildPatchedNumberingXml(
  originalXml: string,
  currentXml: string,
  changed: ChangedNumberingDefs,
): string | null {
  if (changed.abstractNums.size === 0 && changed.nums.size === 0) {
    return originalXml;
  }

  const replacements: XmlSplice[] = [];
  const collect = (kind: NumberingElementKind, ids: Set<string>): boolean => {
    for (const id of ids) {
      const origOffsets = findNumberingElementOffsets(originalXml, kind, id);
      if (!origOffsets) {
        return false;
      }
      const reserialized = extractNumberingElementXml(currentXml, kind, id);
      if (reserialized === null) {
        return false;
      }
      // Restore any custom/mc:AlternateContent numFmt the model flattened to a
      // synthetic value, so editing an unrelated field never corrupts the format.
      const originalDefXml = originalXml.slice(origOffsets.start, origOffsets.end);
      const newXml = restoreLevelNumFmts(originalDefXml, reserialized);
      replacements.push({ start: origOffsets.start, end: origOffsets.end, newXml });
    }
    return true;
  };

  if (!collect("abstractNum", changed.abstractNums) || !collect("num", changed.nums)) {
    return null;
  }

  // abstractNum and num elements are disjoint siblings, so the splice ordering
  // is total.
  return spliceXml(originalXml, replacements);
}

/**
 * Numbering definitions present in `currentXml` (the model's serialization)
 * but absent from `baselineXml` (the re-parsed original): definitions a
 * transform minted. `collectChangedNumberingDefs` deliberately skips these
 * because they cannot be spliced by id; they are appended instead.
 */
export function collectAddedNumberingDefs(
  baselineXml: string,
  currentXml: string,
): ChangedNumberingDefs {
  const addedForKind = (kind: NumberingElementKind): Set<string> => {
    const added = new Set<string>();
    const baselineIndex = buildNumberingElementOffsetIndex(baselineXml, kind);
    const currentIndex = buildNumberingElementOffsetIndex(currentXml, kind);
    for (const [id, range] of currentIndex) {
      if (range && !baselineIndex.has(id)) {
        added.add(id);
      }
    }
    return added;
  };
  return { abstractNums: addedForKind("abstractNum"), nums: addedForKind("num") };
}

const NUMBERING_CLOSE_ROOT = "</w:numbering>";

type PatchNumberingDefinitionsOptions = {
  originalXml: string;
  baselineXml: string;
  currentXml: string;
};

const numberingDefinitionIdentity = (
  element: XmlElement,
): { kind: NumberingElementKind; id: string } | null => {
  if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")) return null;
  const kind = getLocalName(element.name);
  if (kind !== "abstractNum" && kind !== "num") return null;
  const id = getAttributeByNamespaceUri(
    element,
    WORDPROCESSINGML_NAMESPACE_URIS,
    kind === "num" ? "numId" : "abstractNumId",
  );
  return id === null ? null : { kind, id };
};

type PatchNumberingByNamespaceOptions = {
  original: XmlElement;
  currentXml: string;
  changed: ChangedNumberingDefs;
  added: ChangedNumberingDefs;
};

const patchNumberingByNamespace = ({
  original,
  currentXml,
  changed,
  added,
}: PatchNumberingByNamespaceOptions): string | null => {
  const current = parseXmlDocument(currentXml);
  if (!current) return null;
  const replacements = new Map<string, XmlElement>();
  const namespaceDeclarations = Object.fromEntries(
    Object.entries(current.attributes ?? {}).filter(
      ([name]) => name === "xmlns" || name.startsWith("xmlns:"),
    ),
  );
  for (const child of getChildElements(current)) {
    const identity = numberingDefinitionIdentity(child);
    if (!identity) continue;
    const detached = parseXmlDocument(
      captureVerbatimXml(
        cloneElement(child, { attributes: { ...namespaceDeclarations, ...child.attributes } }),
      ),
    );
    if (!detached) return null;
    replacements.set(`${identity.kind}:${identity.id}`, detached);
  }
  const elements: XmlElement[] = [];
  for (const child of original.elements ?? []) {
    const identity = numberingDefinitionIdentity(child);
    if (!identity || !changed[identity.kind === "num" ? "nums" : "abstractNums"].has(identity.id)) {
      elements.push(child);
      continue;
    }
    const replacement = replacements.get(`${identity.kind}:${identity.id}`);
    if (!replacement) return null;
    elements.push(replacement);
  }
  for (const [kind, ids] of [
    ["abstractNum", added.abstractNums],
    ["num", added.nums],
  ] as const) {
    for (const id of ids) {
      const replacement = replacements.get(`${kind}:${id}`);
      if (!replacement) return null;
      const firstNum =
        kind === "abstractNum"
          ? elements.findIndex((child) => numberingDefinitionIdentity(child)?.kind === "num")
          : -1;
      elements.splice(firstNum < 0 ? elements.length : firstNum, 0, replacement);
    }
  }
  return captureVerbatimXml(cloneElement(original, { elements }));
};

/** Both save paths must write changed definitions and newly referenced instances together. */
export const patchNumberingDefinitions = ({
  originalXml,
  baselineXml,
  currentXml,
}: PatchNumberingDefinitionsOptions): string | null => {
  const changed = collectChangedNumberingDefs(baselineXml, currentXml);
  const added = collectAddedNumberingDefs(baselineXml, currentXml);
  const original = parseXmlDocument(originalXml);
  if (
    !original ||
    getLocalName(original.name) !== "numbering" ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(original) ?? "")
  )
    return null;
  if (getNamespacePrefix(original.name ?? "") !== "w") {
    return patchNumberingByNamespace({ original, currentXml, changed, added });
  }
  const spliced = buildPatchedNumberingXml(originalXml, currentXml, changed);
  return spliced === null ? null : appendNumberingDefs(spliced, currentXml, added);
};

/**
 * Append added `w:abstractNum` / `w:num` definitions from `currentXml` to
 * `xml`. ECMA-376 §17.9 orders every `w:abstractNum` before the first `w:num`,
 * so abstract definitions go right before the first `<w:num ` (or the root
 * close when the part has none) and instances before the root close. A
 * synthetic numFmt in an added abstract is restored from the original
 * definition it was cloned from (the one with an identical body), so a custom
 * format survives cloning. Returns null when an added id cannot be extracted
 * or the part has no `</w:numbering>`.
 */
export function appendNumberingDefs(
  xml: string,
  currentXml: string,
  added: ChangedNumberingDefs,
): string | null {
  if (added.abstractNums.size === 0 && added.nums.size === 0) {
    return xml;
  }
  const rootClose = xml.lastIndexOf(NUMBERING_CLOSE_ROOT);
  if (rootClose < 0) {
    return null;
  }
  const abstractXmls: string[] = [];
  for (const id of added.abstractNums) {
    const def = extractNumberingElementXml(currentXml, "abstractNum", id);
    if (def === null) {
      return null;
    }
    abstractXmls.push(restoreClonedLevelNumFmts(xml, currentXml, def, id));
  }
  const numXmls: string[] = [];
  for (const id of added.nums) {
    const def = extractNumberingElementXml(currentXml, "num", id);
    if (def === null) {
      return null;
    }
    numXmls.push(def);
  }

  const firstNum = findFirstElement(xml, NUMBERING_OPEN_LITERAL.num, NUMBERING_CLOSE_TAG.num);
  const abstractInsertAt = firstNum ? firstNum.start : rootClose;
  const head = xml.slice(0, abstractInsertAt);
  const middle = xml.slice(abstractInsertAt, rootClose);
  const tail = xml.slice(rootClose);
  return head + abstractXmls.join("") + middle + numXmls.join("") + tail;
}

/**
 * For an added abstract definition that carries a flattened custom numFmt, find
 * the original abstract it was cloned from (same serialized body apart from the
 * id) and restore the level formats from it.
 */
function restoreClonedLevelNumFmts(
  originalXml: string,
  currentXml: string,
  addedDefXml: string,
  addedId: string,
): string {
  if (!CUSTOM_NUM_FMT_PATTERN.test(addedDefXml)) {
    return addedDefXml;
  }
  const addedBody = stripAbstractNumId(addedDefXml, addedId);
  const originalIds = collectElementIds(
    originalXml,
    NUMBERING_OPEN_LITERAL.abstractNum,
    NUMBERING_ID_ATTR.abstractNum,
  );
  for (const [id] of originalIds) {
    const candidate = extractNumberingElementXml(currentXml, "abstractNum", id);
    if (candidate === null || stripAbstractNumId(candidate, id) !== addedBody) {
      continue;
    }
    const originalDef = extractNumberingElementXml(originalXml, "abstractNum", id);
    if (originalDef !== null) {
      return restoreLevelNumFmts(originalDef, addedDefXml);
    }
  }
  return restoreLevelNumFmts("", addedDefXml);
}

const stripAbstractNumId = (defXml: string, id: string): string =>
  defXml.replace(`${NUMBERING_ID_ATTR.abstractNum}="${id}"`, "");

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}
