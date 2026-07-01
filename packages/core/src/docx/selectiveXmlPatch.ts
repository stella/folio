/**
 * Selective XML Patch Module
 *
 * Patches only changed paragraphs in document.xml, preserving
 * unchanged content byte-for-byte. Uses string offset tracking
 * with proper tag depth counting (not regex) to handle nested elements.
 */

/**
 * Whether `char` ends an element's tag name in XML — a whitespace separator
 * (space, tab, CR, or LF, all valid before attributes per XML 1.0 §3.1), the
 * tag close `>`, or a self-close `/`. Manual tag scanners must accept every
 * whitespace form, not just a literal space, so newline-formatted markup
 * (`<w:p\n  w14:paraId="…">`) is still recognized as the element rather than
 * mistaken for a longer-named sibling.
 */
function isXmlNameBoundary(char: string | undefined): boolean {
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
          pos = tagEnd + 1;
        } else {
          depth++;
          pos = tagEnd + 1;
        }
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

/**
 * Validate that a selective patch can be safely applied.
 *
 * Checks:
 * - All changed paraIds exist in original XML (exactly once)
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

  const originalParaIds = collectParaIds(originalXml);
  const serializedParaIds = collectParaIds(serializedXml);

  // Check all changed IDs exist in original (exactly once)
  for (const id of changedIds) {
    const origCount = originalParaIds.get(id) || 0;
    if (origCount === 0) {
      return { safe: false, reason: `paraId-not-found-in-original: ${id}` };
    }
    if (origCount > 1) {
      return { safe: false, reason: `duplicate-paraId-in-original: ${id}` };
    }
  }

  // Check all changed IDs exist in serialized (exactly once)
  for (const id of changedIds) {
    const serCount = serializedParaIds.get(id) || 0;
    if (serCount === 0) {
      return { safe: false, reason: `paraId-not-found-in-serialized: ${id}` };
    }
    if (serCount > 1) {
      return { safe: false, reason: `duplicate-paraId-in-serialized: ${id}` };
    }
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

/**
 * Replace each changed paragraph in `originalXml` with its re-serialized form
 * extracted from `serializedXml`, splicing end-to-start so earlier offsets stay
 * valid. Assumes safety has already been validated. Returns null if an offset
 * or extraction unexpectedly fails.
 */
function spliceChangedParagraphs(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  const replacements: { start: number; end: number; newXml: string }[] = [];

  for (const paraId of changedIds) {
    const origOffsets = findParagraphOffsets(originalXml, paraId);
    if (!origOffsets) {
      return null;
    }

    const newXml = extractParagraphXml(serializedXml, paraId);
    if (!newXml) {
      return null;
    }

    replacements.push({
      start: origOffsets.start,
      end: origOffsets.end,
      newXml,
    });
  }

  // Sort by start offset descending so we can splice end-to-start
  // (this preserves earlier offsets when replacing later sections)
  replacements.sort((a, b) => b.start - a.start);

  let result = originalXml;
  for (const { start, end, newXml } of replacements) {
    result = result.slice(0, start) + newXml + result.slice(end);
  }

  return result;
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
        pos = tagEnd + 1;
      } else {
        depth++;
        pos = tagEnd + 1;
      }
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

function collectNumberingElementIds(xml: string, kind: NumberingElementKind): Map<string, number> {
  return collectElementIds(xml, NUMBERING_OPEN_LITERAL[kind], NUMBERING_ID_ATTR[kind]);
}

// Synthetic number formats folio's parser mints from custom / mc:AlternateContent
// formats (`decimalZero3/4/5`). These are NOT valid OOXML values, so a changed
// definition must never emit them verbatim — see restoreLevelNumFmts.
const SYNTHETIC_NUM_FMT_PATTERN = /<w:numFmt w:val="decimalZero(?<width>[345])"\/>/u;

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
    const baselineIds = collectNumberingElementIds(baselineXml, kind);
    for (const [id, count] of collectNumberingElementIds(currentXml, kind)) {
      if (count !== 1 || baselineIds.get(id) !== 1) {
        continue;
      }
      const before = extractNumberingElementXml(baselineXml, kind, id);
      const after = extractNumberingElementXml(currentXml, kind, id);
      if (before !== null && after !== null && before !== after) {
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
 * A valid OOXML custom number format equivalent to a `decimalZero{width}`
 * synthetic value: a zero-padded first token folio's parser reads back to the
 * same width. Used only when the original element cannot be resolved, so a
 * changed definition never emits a non-OOXML synthetic value.
 */
function reconstructCustomNumFmt(width: string): string {
  const token = `${"0".repeat(Number.parseInt(width, 10) - 1)}1`;
  return `<w:numFmt w:val="custom" w:format="${token}"/>`;
}

/**
 * The synthetic `decimalZero{width}` width the parser would derive from an
 * original level's number-format element (its custom `w:format` first token),
 * as the same digit-count string the SYNTHETIC_NUM_FMT_PATTERN captures. Mirrors
 * `parseCustomNumberFormat` (first comma-separated token, `^0+1$`, clamped to 5).
 * Returns null when the element carries no zero-padded custom format, so a
 * genuine format change is never mistaken for the original.
 */
function originalNumFmtWidth(numFmtElement: string): string | null {
  for (const match of numFmtElement.matchAll(/<w:numFmt\b[^>]*\/>/gu)) {
    const tag = match[0];
    if (!/\bw:val="custom"/u.test(tag)) {
      continue;
    }
    const format = /\bw:format="(?<format>[^"]*)"/u.exec(tag)?.groups?.["format"];
    if (format === undefined) {
      continue;
    }
    const firstToken = format.split(",")[0]?.trim() ?? "";
    return /^0+1$/u.test(firstToken) ? String(Math.min(firstToken.length, 5)) : null;
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
 * Swap each synthetic `decimalZero{3,4,5}` numFmt in a re-serialized definition
 * back to the original level's number-format element.
 *
 * folio's parser collapses a Word custom / mc:AlternateContent number format
 * into a synthetic `decimalZero{3,4,5}` value the model cannot re-emit as valid
 * OOXML. When an unrelated field of the definition is edited, the whole
 * definition is re-serialized, so without this pass the level's custom format
 * would be replaced by an invalid value and lost on reparse. For each level
 * whose re-emitted numFmt is synthetic, restore the original level's numFmt
 * element by ilvl ONLY when the model's synthetic value still matches the
 * original's (the format was not edited, just preserved past an unrelated
 * change). When the model deliberately changed the width (e.g. decimalZero4 ->
 * decimalZero5), or the original cannot be resolved, emit a valid OOXML custom
 * format for the CURRENT width instead — honoring the edit and never leaving a
 * synthetic value.
 */
function restoreLevelNumFmts(originalDefXml: string, currentDefXml: string): string {
  const replacements: { start: number; end: number; newXml: string }[] = [];
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
    const synthetic = SYNTHETIC_NUM_FMT_PATTERN.exec(curLevel);
    if (!synthetic) {
      continue;
    }
    const origLevel = extractElementByIdAttr(
      originalDefXml,
      LEVEL_OPEN_LITERAL,
      LEVEL_CLOSE_TAG,
      LEVEL_ID_ATTR,
      ilvl,
    );
    // SAFETY: named group `width` present because SYNTHETIC_NUM_FMT_PATTERN matched
    const currentWidth = synthetic.groups!["width"]!;
    const original = origLevel ? extractLevelNumFmtElement(origLevel) : null;
    let replacement: string;
    if (original && origLevel && originalNumFmtWidth(original) === currentWidth) {
      // Format unchanged — restore the original element verbatim. It (e.g.
      // mc:AlternateContent) may use a prefix bound on the replaced ancestors
      // (w:abstractNum / w:lvl) instead of the numbering root; carry those
      // declarations onto it so it stays resolvable.
      const ancestorXmlns = {
        ...collectXmlnsFromOpeningTag(originalDefXml),
        ...collectXmlnsFromOpeningTag(origLevel),
      };
      replacement = withXmlnsDeclarations(original, ancestorXmlns);
    } else {
      // The model changed the width (or the original is unresolvable): honor the
      // current value via a valid OOXML custom format, never the synthetic literal.
      replacement = reconstructCustomNumFmt(currentWidth);
    }
    const restoredLevel =
      curLevel.slice(0, synthetic.index) +
      replacement +
      curLevel.slice(synthetic.index + synthetic[0].length);
    replacements.push({ start: curOffsets.start, end: curOffsets.end, newXml: restoredLevel });
  }

  replacements.sort((a, b) => b.start - a.start);
  let result = currentDefXml;
  for (const { start, end, newXml } of replacements) {
    result = result.slice(0, start) + newXml + result.slice(end);
  }
  return result;
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

  const replacements: { start: number; end: number; newXml: string }[] = [];
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

  // Splice end-to-start so earlier offsets stay valid. abstractNum and num
  // elements are disjoint siblings, so descending-by-start ordering is total.
  replacements.sort((a, b) => b.start - a.start);
  let result = originalXml;
  for (const { start, end, newXml } of replacements) {
    result = result.slice(0, start) + newXml + result.slice(end);
  }
  return result;
}

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}
