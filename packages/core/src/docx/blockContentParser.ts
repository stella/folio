/**
 * Shared OOXML block-content parser.
 *
 * The document body, headers, footers, and SDT content all expose the same
 * block-level model: paragraphs, tables, and nested structured document tags.
 * Keeping the parser shared prevents body-only fixes, especially for drawings
 * like text boxes that can appear in headers and footers too.
 */

import type {
  BlockContent,
  BlockSdt,
  BookmarkEnd,
  BookmarkStart,
  MediaFile,
  Paragraph,
  RelationshipMap,
  Theme,
} from "../types/document";
import { parseBookmarkEnd, parseBookmarkStart } from "./bookmarkParser";
import {
  appendBookmarkMarkerToLastParagraphInBlocks,
  prependBookmarkMarkersToFirstParagraphInBlocks,
} from "./bookmarkPlacement";
import type { BookmarkMarker } from "./bookmarkPlacement";
import { convertBulletToUnicode } from "./bulletMarkers";
import { padDecimal, type NumberingMap } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import { enrichParagraphTextBoxes } from "./paragraphTextBoxEnrichment";
import { parseSdtProperties } from "./sdtProperties";
import type { StyleMap } from "./styleParser";
import { parseTable } from "./tableParser";
import {
  elementToXml,
  findChild,
  getChildElements,
  getLocalName,
  mergeXmlnsDeclarations,
  type XmlElement,
} from "./xmlParser";

type ParseBlockContentOptions = {
  inHeaderFooter?: boolean;
  // Source root `xmlns:*` declarations, threaded to the run parser so a captured
  // VML `w:pict` replay stays self-contained under non-canonical prefixes.
  rootXmlns?: Record<string, string>;
};

const toRoman = (numParam: number): string => {
  let num = numParam;
  const romanNumerals: [number, string][] = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];

  let result = "";
  for (const [value, symbol] of romanNumerals) {
    while (num >= value) {
      result += symbol;
      num -= value;
    }
  }
  return result;
};

const toRepeatedLetter = (value: number, baseCodePoint: number): string => {
  if (value <= 0) {
    return "0";
  }
  const zeroBased = value - 1;
  const letter = String.fromCodePoint(baseCodePoint + (zeroBased % 26));
  return letter.repeat(Math.floor(zeroBased / 26) + 1);
};

const formatNumber = (value: number, numFmt: string): string => {
  switch (numFmt) {
    case "decimal":
      return String(value);
    case "decimalZero":
      return padDecimal(value, 2);
    case "decimalZero3":
      return padDecimal(value, 3);
    case "decimalZero4":
      return padDecimal(value, 4);
    case "decimalZero5":
      return padDecimal(value, 5);
    case "lowerLetter":
      return toRepeatedLetter(value, 97);
    case "upperLetter":
      return toRepeatedLetter(value, 65);
    case "lowerRoman":
      return toRoman(value).toLowerCase();
    case "upperRoman":
      return toRoman(value);
    case "bullet":
      return "\u2022";
    case "none":
      return "";
    default:
      return String(value);
  }
};

const computeListMarker = (
  paragraph: Paragraph,
  numbering: NumberingMap | null,
  listCounters: Map<number, number[]>,
  abstractCounters: Map<number, number[]>,
): void => {
  const listRendering = paragraph.listRendering;
  if (!listRendering || !numbering) {
    return;
  }

  const { numId, level } = listRendering;
  if (numId === 0) {
    return;
  }

  if (!listCounters.has(numId)) {
    listCounters.set(numId, Array.from<number>({ length: 9 }).fill(Number.NaN));
  }

  const counters = listCounters.get(numId);
  if (!counters) {
    return;
  }

  const abstractNumId = numbering.getAbstractNumId(numId);
  const styleNumbering = paragraph.formatting?.numPrFromStyle;
  let resumedAbstractCounters: number[] | undefined;
  if (abstractNumId !== null && styleNumbering) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    if (latestAbstractCounters) {
      // A paragraph whose numbering comes only from its style resumes the
      // latest compatible list instance. Word does this when an attachment
      // starts a fresh w:num (with a startOverride) and later paragraphs fall
      // back to the style's original w:num: the style continues the attachment
      // sequence instead of reviving its stale counters from earlier content.
      for (let i = 0; i < counters.length; i += 1) {
        counters[i] = latestAbstractCounters[i] ?? Number.NaN;
      }
      resumedAbstractCounters = latestAbstractCounters;
    }
  }
  if (abstractNumId !== null && level > 0) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    const missingParentCounters = counters.slice(0, level).every(Number.isNaN);
    if (missingParentCounters) {
      for (let i = 0; i < level; i += 1) {
        const latestCounter = latestAbstractCounters?.[i];
        counters[i] =
          latestCounter !== undefined && !Number.isNaN(latestCounter)
            ? latestCounter
            : (numbering.getLevel(numId, i)?.start ?? 1);
      }
    }
  }

  if (Number.isNaN(counters[level])) {
    counters[level] = (numbering.getLevel(numId, level)?.start ?? 1) - 1;
  }
  counters[level] = (counters[level] ?? 0) + 1;

  for (let i = level + 1; i < counters.length; i += 1) {
    counters[i] = Number.NaN;
  }

  // Word's default LISTNUM field advances the counter at one ilvl deeper
  // than the host paragraph. Mirror the toFlowBlocks logic here so the
  // marker substituted at parse time agrees with the renderer's counters —
  // otherwise a follow-up paragraph at that depth picks up the stale,
  // pre-substituted "(a)" instead of "(b)".
  const childAdvances = listRendering.implicitChildLevelAdvances ?? 0;
  if (childAdvances > 0 && level + 1 < counters.length) {
    const childCounter = counters[level + 1];
    counters[level + 1] =
      (childCounter === undefined || Number.isNaN(childCounter) ? 0 : childCounter) + childAdvances;
  }

  if (abstractNumId !== null) {
    if (resumedAbstractCounters) {
      for (const [otherNumId, otherCounters] of listCounters) {
        if (
          otherNumId === numId ||
          numbering.getAbstractNumId(otherNumId) !== abstractNumId ||
          !otherCounters.every((value, index) => Object.is(value, resumedAbstractCounters[index]))
        ) {
          continue;
        }
        for (let i = 0; i < otherCounters.length; i += 1) {
          otherCounters[i] = counters[i] ?? Number.NaN;
        }
      }
    }
    abstractCounters.set(abstractNumId, [...counters]);
  }

  const pattern = listRendering.marker;

  if (listRendering.isBullet) {
    listRendering.marker = convertBulletToUnicode(pattern || "");
    return;
  }

  let computedMarker = pattern;
  const currentLevelInfo = numbering.getLevel(numId, level);
  const useLegalNumbering = currentLevelInfo?.isLgl === true || listRendering.isLegal === true;

  for (let lvl = 0; lvl <= level; lvl += 1) {
    const placeholder = `%${lvl + 1}`;
    if (computedMarker.includes(placeholder)) {
      const value = counters[lvl] ?? 0;
      const levelInfo = numbering.getLevel(numId, lvl);
      const formatted = formatNumber(
        value,
        useLegalNumbering ? "decimal" : levelInfo?.numFmt || "decimal",
      );
      computedMarker = computedMarker.replaceAll(placeholder, formatted);
    }
  }

  listRendering.marker = computedMarker;
};

type ParseBlockContentState = {
  listCounters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  options: ParseBlockContentOptions | undefined;
};

export const parseBlockContent = (
  parent: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options?: ParseBlockContentOptions,
): BlockContent[] =>
  parseBlockContentWithState(parent, styles, theme, numbering, rels, media, {
    listCounters: new Map(),
    abstractCounters: new Map(),
    // Accumulate the container's own xmlns onto the inherited in-scope set so a
    // captured VML `w:pict` replay resolves prefixes scoped on this level too.
    options: {
      ...options,
      rootXmlns: mergeXmlnsDeclarations(options?.rootXmlns ?? {}, parent),
    },
  });

const parseBlockContentWithState = (
  parent: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  state: ParseBlockContentState,
): BlockContent[] => {
  const content: BlockContent[] = [];
  const children = getChildElements(parent);
  const pendingBookmarkMarkers: BookmarkMarker[] = [];

  for (const child of children) {
    const name = child.name ?? "";
    const localName = getLocalName(name);

    if (localName === "p") {
      const paragraph = parseParagraph(child, styles, theme, numbering, rels, media, state.options);
      prependPendingBookmarkMarkers(paragraph, pendingBookmarkMarkers);
      enrichParagraphTextBoxes(paragraph, child, styles, theme, numbering, rels, media);
      computeListMarker(paragraph, numbering, state.listCounters, state.abstractCounters);
      content.push(paragraph);
      continue;
    }

    if (localName === "tbl") {
      const table = parseTable(child, styles, theme, numbering, rels, media, state.options);
      if (prependBookmarkMarkersToFirstParagraphInBlocks([table], pendingBookmarkMarkers)) {
        pendingBookmarkMarkers.length = 0;
      }
      content.push(table);
      continue;
    }

    if (localName === "sdt") {
      const sdtPr = findChild(child, "w", "sdtPr");
      const sdtEndPr = findChild(child, "w", "sdtEndPr");
      const sdtContent = findChild(child, "w", "sdtContent");
      const properties = parseSdtProperties(sdtPr, sdtEndPr);
      // Capture non-content direct children of <w:sdt> (bookmark / comment /
      // tracked-change / custom XML range markers — MS-OE376 §2.5.2.30) so a
      // comment thread or tracked change that crosses an SDT boundary
      // doesn't lose a delimiter on round-trip. Split by position relative
      // to sdtContent.
      const captured = captureSdtSiblingMarkers(child);
      if (captured.before.length > 0) {
        properties.rawSdtChildrenBeforeContent = captured.before;
      }
      if (captured.after.length > 0) {
        properties.rawSdtChildrenAfterContent = captured.after;
      }
      const blockSdt: BlockSdt = {
        type: "blockSdt",
        properties,
        content: sdtContent
          ? parseBlockContentWithState(sdtContent, styles, theme, numbering, rels, media, state)
          : [],
      };
      if (
        prependBookmarkMarkersToFirstParagraphInBlocks(blockSdt.content, pendingBookmarkMarkers)
      ) {
        pendingBookmarkMarkers.length = 0;
      }
      content.push(blockSdt);
      continue;
    }

    if (localName === "bookmarkStart" || localName === "bookmarkEnd") {
      const marker = parseBookmarkMarker(child, localName);
      if (!appendBookmarkMarkerToLastParagraphInBlocks(content, marker)) {
        pendingBookmarkMarkers.push(marker);
      }
    }
  }

  if (pendingBookmarkMarkers.length > 0) {
    content.push({
      type: "paragraph",
      content: [...pendingBookmarkMarkers],
    });
  }

  return content;
};

/**
 * Walk a `<w:sdt>` element's direct children and return the verbatim XML
 * for every child that is NOT `<w:sdtPr>`, `<w:sdtEndPr>`, or
 * `<w:sdtContent>` — split by position relative to sdtContent.
 *
 * Per MS-OE376 §2.5.2.30, Word emits 16 range-marker elements (bookmark,
 * comment, custom XML, tracked-change ranges) as direct sdt siblings of
 * sdtContent. Without preserving them, a comment thread or tracked
 * change that crosses an SDT boundary loses a delimiter when folio
 * serializes the parsed model back out.
 */
const captureSdtSiblingMarkers = (sdt: XmlElement): { before: string; after: string } => {
  const beforeParts: string[] = [];
  const afterParts: string[] = [];
  let sawContent = false;
  for (const ch of sdt.elements ?? []) {
    if (ch.type !== "element" || !ch.name) {
      continue;
    }
    const local = getLocalName(ch.name);
    if (local === "sdtPr" || local === "sdtEndPr") {
      continue;
    }
    if (local === "sdtContent") {
      sawContent = true;
      continue;
    }
    const xml = elementToXml(ch);
    if (sawContent) {
      afterParts.push(xml);
    } else {
      beforeParts.push(xml);
    }
  }
  return { before: beforeParts.join(""), after: afterParts.join("") };
};

const parseBookmarkMarker = (
  child: XmlElement,
  localName: "bookmarkStart" | "bookmarkEnd",
): BookmarkStart | BookmarkEnd => {
  if (localName === "bookmarkStart") {
    return parseBookmarkStart(child);
  }
  return parseBookmarkEnd(child);
};

const prependPendingBookmarkMarkers = (
  paragraph: Paragraph,
  pendingBookmarkMarkers: BookmarkMarker[],
): void => {
  if (pendingBookmarkMarkers.length === 0) {
    return;
  }

  paragraph.content.unshift(...pendingBookmarkMarkers);
  pendingBookmarkMarkers.length = 0;
};
