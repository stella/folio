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
import { resolveParagraphIdentities, type ParagraphIdentity } from "./paraIdAttribute";
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

/** One `<w:p>` of a part: its byte range and the paraId its open tag carries. */
export type ScannedParagraph = ParagraphOffsets & { paraId: string | undefined };

/**
 * Every `<w:p>` element of `xml`, in the document order of its opening tags,
 * with the `w14:paraId` that tag carries.
 *
 * Document order is what makes the array an ordinal space: the *n*th entry of
 * the source part and the *n*th entry of the model's serialization name the
 * same paragraph, which is the only way to address a paragraph the producer
 * gave no id. A `<w:p>` nested inside another (inside `mc:AlternateContent`,
 * a text box) is one entry of its own, exactly as
 * {@link countParagraphElements} counts it, so the two never disagree about
 * what an ordinal is. An unterminated paragraph keeps `end === start`: it
 * occupies its ordinal but no splice can be built from it.
 */
export function scanParagraphs(xml: string): ScannedParagraph[] {
  const paragraphs: ScannedParagraph[] = [];
  const open: number[] = [];
  let pos = 0;

  while (pos < xml.length) {
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

    if (xml[tagEnd - 1] === "/") {
      // Self-closing <w:p ... /> — resolved immediately, never pushed.
      paragraphs.push({ start: tagStart, end: tagEnd + 1, paraId });
    } else {
      open.push(paragraphs.length);
      paragraphs.push({ start: tagStart, end: tagStart, paraId });
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

export type PatchSafetyOptions = {
  /**
   * Require the original and serialized XML to hold the same number of `<w:p>`
   * elements. Guards document.xml against structural drift. Notes disable it:
   * the model only retains the normal notes, so a serialized note part
   * legitimately has fewer paragraphs than the original (which still carries
   * the separator notes). See {@link buildPatchedNoteXml}.
   */
  checkParagraphCount?: boolean;
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

/**
 * Route every changed paragraph from the model's serialization to its region
 * of the source part.
 *
 * One owner for the question the patch turns on: *which paragraph of the file
 * is this model paragraph?* {@link resolveParagraphIdentities} answers it per
 * paragraph, and the union is consumed exhaustively here, so a paragraph the
 * source names by id keeps the id lookup — robust to any reordering — and one
 * the source never named falls back to its ordinal, which is sound exactly
 * when the identity plan says the two sequences line up. A package with ids on
 * some paragraphs and not others therefore needs no special case: its authored
 * ids are the witnesses that prove the ordinals for the rest.
 */
const routeChangedParagraphs = (
  originalXml: string,
  serializedXml: string,
  changedIds: ReadonlySet<string>,
): ParagraphRouting => {
  const original = scanParagraphs(originalXml);
  const serialized = scanParagraphs(serializedXml);
  const { identities, ordinalsAligned } = resolveParagraphIdentities({
    sourceParaIds: original.map(({ paraId }) => paraId),
    serializedParaIds: serialized.map(({ paraId }) => paraId),
  });

  const originalParaIds = collectParaIds(originalXml);
  const serializedParaIds = collectParaIds(serializedXml);
  const identityByParaId = new Map<string, ParagraphIdentity>();
  for (const identity of identities) {
    if (identity.type !== "anonymous" && serializedParaIds.get(identity.paraId) === 1) {
      identityByParaId.set(identity.paraId, identity);
    }
  }

  const splices: XmlSplice[] = [];
  for (const id of changedIds) {
    const originalCount = originalParaIds.get(id) ?? 0;
    if (originalCount > 1) {
      return { type: "refused", reason: `duplicate-paraId-in-original: ${id}` };
    }
    const serializedCount = serializedParaIds.get(id) ?? 0;
    if (originalCount === 1 && serializedCount === 0) {
      return { type: "refused", reason: `paraId-not-found-in-serialized: ${id}` };
    }
    if (serializedCount === 0) {
      return { type: "refused", reason: `paraId-not-found-in-original: ${id}` };
    }
    if (serializedCount > 1) {
      return { type: "refused", reason: `duplicate-paraId-in-serialized: ${id}` };
    }

    const identity = identityByParaId.get(id);
    if (!identity) {
      return { type: "refused", reason: `paraId-not-found-in-serialized: ${id}` };
    }
    const replacement = serialized[identity.ordinal];
    if (!replacement || replacement.end <= replacement.start) {
      return { type: "refused", reason: `unterminated-paragraph: ${id}` };
    }
    const newXml = serializedXml.slice(replacement.start, replacement.end);

    switch (identity.type) {
      case "authored": {
        const offsets = findParagraphOffsets(originalXml, identity.paraId);
        if (!offsets) {
          return { type: "refused", reason: `paraId-not-found-in-original: ${id}` };
        }
        splices.push({ start: offsets.start, end: offsets.end, newXml });
        break;
      }
      case "minted": {
        if (!ordinalsAligned) {
          return { type: "refused", reason: `unaligned-paragraph-ordinals: ${id}` };
        }
        const source = original[identity.ordinal];
        if (!source || source.end <= source.start) {
          return { type: "refused", reason: `paraId-not-found-in-original: ${id}` };
        }
        splices.push({
          start: source.start,
          end: source.end,
          newXml: withoutMintedIds(newXml, originalXml.slice(source.start, source.end)),
        });
        break;
      }
      case "anonymous": {
        // Unreachable: `identityByParaId` only holds the two keyed branches.
        return { type: "refused", reason: `paraId-not-found-in-serialized: ${id}` };
      }
      default: {
        const unreachable: never = identity;
        return unreachable;
      }
    }
  }

  return { type: "routed", splices };
};

/**
 * Validate that a selective patch can be safely applied.
 *
 * Checks:
 * - Every changed paraId routes to one region of the original XML, by its
 *   authored id or, when the producer wrote none, by its paragraph ordinal
 * - All changed paraIds exist in serialized XML (exactly once)
 * - Paragraph count matches between original and serialized (unless disabled)
 */
export function validatePatchSafety(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
  options: PatchSafetyOptions = {},
): PatchValidationResult {
  if (changedIds.size === 0) {
    return { safe: true };
  }

  const routing = routeChangedParagraphs(originalXml, serializedXml, changedIds);
  if (routing.type === "refused") {
    return { safe: false, reason: routing.reason };
  }

  if (options.checkParagraphCount === false) {
    return { safe: true };
  }

  // Check paragraph counts match
  const originalCount = countParagraphElements(originalXml);
  const serializedCount = countParagraphElements(serializedXml);
  if (originalCount !== serializedCount) {
    return {
      safe: false,
      reason: `paragraph-count-mismatch: original=${originalCount}, serialized=${serializedCount}`,
    };
  }

  return { safe: true };
}

/**
 * Build a patched document.xml by splicing new paragraph XML into
 * the original at the correct offsets. Only changed paragraphs
 * are replaced; everything else is preserved byte-for-byte.
 *
 * Returns null if any step fails.
 */
export function buildPatchedDocumentXml(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  if (changedIds.size === 0) {
    return originalXml;
  }

  // Validate safety first
  const validation = validatePatchSafety(originalXml, serializedXml, changedIds);
  if (!validation.safe) {
    return null;
  }

  return spliceChangedParagraphs(originalXml, serializedXml, changedIds);
}

/**
 * Build a patched note part (word/footnotes.xml / word/endnotes.xml) by
 * splicing edited note paragraphs into the original, preserving unchanged
 * content byte-for-byte.
 *
 * Unlike {@link buildPatchedDocumentXml} this does NOT require the paragraph
 * counts to match: the document model only retains the normal notes, so the
 * serialized note XML omits the separator / continuationSeparator paragraphs
 * the original part still carries. Splicing by `paraId` keeps those separators
 * and every unedited note byte-exact while replacing only the edited ones.
 *
 * Returns null if any changed id is missing or ambiguous in either input, so
 * the caller can fall back to preserving the original part verbatim.
 */
export function buildPatchedNoteXml(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  if (changedIds.size === 0) {
    return originalXml;
  }

  const validation = validatePatchSafety(originalXml, serializedXml, changedIds, {
    checkParagraphCount: false,
  });
  if (!validation.safe) {
    return null;
  }

  return spliceChangedParagraphs(originalXml, serializedXml, changedIds);
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

/**
 * Replace each changed paragraph in `originalXml` with its re-serialized form
 * extracted from `serializedXml`. Assumes safety has already been validated.
 * Returns null if an offset or extraction unexpectedly fails, or if
 * {@link spliceXml} refuses the result; the caller then falls back to a full
 * repack, whose parts are all re-serialized from one model.
 */
function spliceChangedParagraphs(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  const routing = routeChangedParagraphs(originalXml, serializedXml, changedIds);
  return routing.type === "refused" ? null : spliceXml(originalXml, routing.splices);
}

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
