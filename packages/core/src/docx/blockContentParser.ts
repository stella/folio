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
  MediaFile,
  PreservedBlock,
  Paragraph,
  RelationshipMap,
  Theme,
} from "../types/document";
import { parseBookmarkEnd, parseBookmarkStart } from "./bookmarkParser";
import {
  CAPTURE,
  dispatchChildren,
  OWNED_ELSEWHERE,
  withPreservedChildren,
} from "./containerChildren";
import {
  appendBookmarkMarkerToLastParagraphInBlocks,
  prependBookmarkMarkersToFirstParagraphInBlocks,
} from "./bookmarkPlacement";
import type { BookmarkMarker } from "./bookmarkPlacement";
import { convertBulletToUnicode } from "./bulletMarkers";
import type { NumberingMap } from "./numberingParser";
import { isNumberingReference } from "./numberingReference";
import { formatOoxmlCounter } from "./ooxmlCounterFormatter";
import { parseParagraph } from "./paragraphParser";
import type { ParseContext } from "./parseContext";
import { enrichParagraphTextBoxes } from "./paragraphTextBoxEnrichment";
import { parseSdtProperties } from "./sdtProperties";
import type { StyleMap } from "./styleParser";
import { parseTable } from "./tableParser";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findChild,
  getLocalName,
  mergeXmlnsDeclarations,
  selectAlternateContentBranch,
  type XmlElement,
} from "./xmlParser";

type ParseBlockContentOptions = {
  inHeaderFooter?: boolean;
  // Source root `xmlns:*` declarations, threaded to the run parser so a captured
  // VML `w:pict` replay stays self-contained under non-canonical prefixes.
  rootXmlns?: Record<string, string>;
  // The normalisation channel, threaded to the text-box enrichment pass so an
  // outline dash outside `ST_PresetLineDashVal` is reported rather than kept in
  // silence. Absent for the tiers that have no collector yet.
  context?: ParseContext | undefined;
};

type PreviousListState = {
  abstractNumId: number | null;
  fromStyle: boolean;
  numId: number | null;
};

type ComputeListMarkerOptions = {
  numbering: NumberingMap | null;
  listCounters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  restartedNumIds: Set<number>;
  previousList: PreviousListState;
};

const computeListMarker = (
  paragraph: Paragraph,
  {
    numbering,
    listCounters,
    abstractCounters,
    restartedNumIds,
    previousList,
  }: ComputeListMarkerOptions,
): void => {
  const listRendering = paragraph.listRendering;
  if (!listRendering || !numbering) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  const { numId, level } = listRendering;
  if (!isNumberingReference(numId)) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  const firstEncounter = !listCounters.has(numId);
  if (firstEncounter) {
    listCounters.set(numId, Array.from<number>({ length: 9 }).fill(Number.NaN));
  }

  let counters = listCounters.get(numId);
  if (!counters) {
    return;
  }

  const abstractNumId = numbering.getAbstractNumId(numId);
  const styleNumbering = paragraph.formatting?.numPrFromStyle;
  const resumesRestartedInstance =
    firstEncounter &&
    listRendering.startOverride === undefined &&
    abstractNumId !== null &&
    previousList.abstractNumId === abstractNumId &&
    previousList.numId !== null &&
    previousList.numId !== numId &&
    restartedNumIds.has(previousList.numId);
  const resumesStyleInstance =
    firstEncounter &&
    !styleNumbering &&
    listRendering.startOverride === undefined &&
    abstractNumId !== null &&
    previousList.abstractNumId === abstractNumId &&
    previousList.fromStyle === true;
  if (
    abstractNumId !== null &&
    (styleNumbering || resumesRestartedInstance || resumesStyleInstance)
  ) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    if (latestAbstractCounters) {
      // A paragraph whose numbering comes only from its style resumes the
      // latest compatible list instance. Word does this when an attachment
      // starts a fresh w:num (with a startOverride) and later paragraphs fall
      // back to the style's original w:num: the style continues the attachment
      // sequence instead of reviving its stale counters from earlier content.
      counters = latestAbstractCounters;
      listCounters.set(numId, counters);
    }
  }
  if (
    listRendering.startOverride !== undefined ||
    resumesRestartedInstance ||
    (previousList.numId === numId && restartedNumIds.has(numId))
  ) {
    restartedNumIds.add(numId);
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
    abstractCounters.set(abstractNumId, counters);
  }
  previousList.abstractNumId = abstractNumId;
  previousList.fromStyle = Boolean(styleNumbering);
  previousList.numId = numId;

  const pattern = listRendering.marker;

  if (listRendering.isBullet) {
    listRendering.marker = convertBulletToUnicode(pattern || "");
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
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
      const formatted = formatOoxmlCounter(
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
  restartedNumIds: Set<number>;
  previousList: PreviousListState;
  options: ParseBlockContentOptions | undefined;
};

const withContainerXmlns = (
  state: ParseBlockContentState,
  element: XmlElement,
): ParseBlockContentState => ({
  ...state,
  options: {
    ...state.options,
    rootXmlns: mergeXmlnsDeclarations(state.options?.rootXmlns ?? {}, element),
  },
});

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
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
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
  const modelled: BlockContent[] = [];
  const pendingBookmarkMarkers: BookmarkMarker[] = [];

  const preserved = dispatchChildren({
    element: parent,
    container: "block-content",
    modelledCount: () => modelled.length,
    undeclared: {
      // `mc:AlternateContent`, whose selected branch folio reads. The others
      // are undeclared in the `w:` sense only because they belong to another
      // namespace, and the sink's default is right for them.
      AlternateContent: (child) => {
        const selectedBranch = selectAlternateContentBranch(child);
        if (!selectedBranch) {
          return;
        }
        modelled.push(
          ...parseBlockContentWithState(
            selectedBranch,
            styles,
            theme,
            numbering,
            rels,
            media,
            withContainerXmlns(withContainerXmlns(state, child), selectedBranch),
          ),
        );
      },
    },
    handlers: {
      p: (child) => {
        const paragraph = parseParagraph(
          child,
          styles,
          theme,
          numbering,
          rels,
          media,
          state.options,
        );
        prependPendingBookmarkMarkers(paragraph, pendingBookmarkMarkers);
        enrichParagraphTextBoxes(
          paragraph,
          child,
          styles,
          theme,
          numbering,
          rels,
          media,
          parseTable,
          state.options?.context,
        );
        computeListMarker(paragraph, {
          numbering,
          listCounters: state.listCounters,
          abstractCounters: state.abstractCounters,
          restartedNumIds: state.restartedNumIds,
          previousList: state.previousList,
        });
        modelled.push(paragraph);
      },
      tbl: (child) => {
        const table = parseTable(child, styles, theme, numbering, rels, media, state.options);
        if (!table) {
          return;
        }
        if (prependBookmarkMarkersToFirstParagraphInBlocks([table], pendingBookmarkMarkers)) {
          pendingBookmarkMarkers.length = 0;
        }
        modelled.push(table);
      },
      sdt: (child) => {
        const blockSdt = parseBlockSdt(child, styles, theme, numbering, rels, media, state);
        if (
          prependBookmarkMarkersToFirstParagraphInBlocks(blockSdt.content, pendingBookmarkMarkers)
        ) {
          pendingBookmarkMarkers.length = 0;
        }
        modelled.push(blockSdt);
      },
      bookmarkStart: (child) => {
        collectBookmarkMarker(child, "bookmarkStart", modelled, pendingBookmarkMarkers);
      },
      bookmarkEnd: (child) => {
        collectBookmarkMarker(child, "bookmarkEnd", modelled, pendingBookmarkMarkers);
      },
      // The body's own `w:sectPr` is read by the document parser and a cell's
      // `w:tcPr` by the table parser, each from the container element; the two
      // are in this map because the union covers every block container, not
      // because this walker reads them.
      sectPr: OWNED_ELSEWHERE,
      tcPr: OWNED_ELSEWHERE,
      altChunk: CAPTURE,
      commentRangeEnd: CAPTURE,
      commentRangeStart: CAPTURE,
      customXml: CAPTURE,
      customXmlDelRangeEnd: CAPTURE,
      customXmlDelRangeStart: CAPTURE,
      customXmlInsRangeEnd: CAPTURE,
      customXmlInsRangeStart: CAPTURE,
      customXmlMoveFromRangeEnd: CAPTURE,
      customXmlMoveFromRangeStart: CAPTURE,
      customXmlMoveToRangeEnd: CAPTURE,
      customXmlMoveToRangeStart: CAPTURE,
      del: CAPTURE,
      ins: CAPTURE,
      moveFrom: CAPTURE,
      moveFromRangeEnd: CAPTURE,
      moveFromRangeStart: CAPTURE,
      moveTo: CAPTURE,
      moveToRangeEnd: CAPTURE,
      moveToRangeStart: CAPTURE,
      permEnd: CAPTURE,
      permStart: CAPTURE,
      proofErr: CAPTURE,
    },
  });

  const content = withPreservedChildren(
    modelled,
    preserved,
    (xml): PreservedBlock => ({ type: "preservedBlock", xml }),
  );
  if (pendingBookmarkMarkers.length > 0) {
    content.push({ type: "paragraph", content: [...pendingBookmarkMarkers] });
  }
  return content;
};

const parseBlockSdt = (
  child: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  state: ParseBlockContentState,
): BlockSdt => {
  const sdtContent = findChild(child, "w", "sdtContent");
  const properties = parseSdtProperties(
    findChild(child, "w", "sdtPr"),
    findChild(child, "w", "sdtEndPr"),
  );
  // Capture non-content direct children of <w:sdt> (bookmark / comment /
  // tracked-change / custom XML range markers — MS-OE376 §2.5.2.30) so a
  // comment thread or tracked change that crosses an SDT boundary doesn't
  // lose a delimiter on round-trip. Split by position relative to sdtContent.
  const captured = captureSdtSiblingMarkers(child);
  if (captured.before.length > 0) {
    properties.rawSdtChildrenBeforeContent = captured.before;
  }
  if (captured.after.length > 0) {
    properties.rawSdtChildrenAfterContent = captured.after;
  }
  return {
    type: "blockSdt",
    properties,
    content: sdtContent
      ? parseBlockContentWithState(
          sdtContent,
          styles,
          theme,
          numbering,
          rels,
          media,
          withContainerXmlns(withContainerXmlns(state, child), sdtContent),
        )
      : [],
  };
};

const collectBookmarkMarker = (
  child: XmlElement,
  localName: "bookmarkStart" | "bookmarkEnd",
  content: readonly BlockContent[],
  pending: BookmarkMarker[],
): void => {
  const marker =
    localName === "bookmarkStart" ? parseBookmarkStart(child) : parseBookmarkEnd(child);
  if (!appendBookmarkMarkerToLastParagraphInBlocks(content, marker)) {
    pending.push(marker);
  }
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
    const xml = captureVerbatimXml(ch);
    if (sawContent) {
      afterParts.push(xml);
    } else {
      beforeParts.push(xml);
    }
  }
  return { before: beforeParts.join(""), after: afterParts.join("") };
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
