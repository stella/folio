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
  BlockCustomXml,
  BlockSdt,
  MediaFile,
  PreservedBlock,
  RelationshipMap,
  Theme,
} from "../types/document";
import { parseBookmarkEnd, parseBookmarkStart } from "./bookmarkParser";
import { blockCustomXmlShell } from "./blockCustomXmlShell";
import {
  CAPTURE,
  type ChildHandlers,
  type ChildReader,
  dispatchChildrenWithContext,
  ownedElsewhere,
  withPreservedChildren,
} from "./containerChildren";
import type { NumberingMap } from "./numberingParser";
import { computeListMarker, type PreviousListState } from "./listMarkerComputation";
import { parseParagraph } from "./paragraphParser";
import type { ParseContext } from "./parseContext";
import { type PreviewLedger, standalonePreviewLedger } from "./previewBudget";
import { enrichParagraphTextBoxes } from "./paragraphTextBoxEnrichment";
import { captureSdtSiblingMarkers, parseSdtProperties } from "./sdtProperties";
import type { StyleMap } from "./styleParser";
import { parseTable } from "./tableParser";
import {
  findWordprocessingChild,
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
  // The ledger of the package these blocks belong to. Blocks read on their own
  // charge their previews to no package.
  previews?: PreviewLedger;
};

/** The options every block below the entry point reads, with its ledger settled. */
type BlockContentScope = ParseBlockContentOptions & { previews: PreviewLedger };

type ParseBlockContentState = {
  listCounters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  restartedNumIds: Set<number>;
  previousList: PreviousListState;
  options: BlockContentScope;
};

const withContainerXmlns = (
  state: ParseBlockContentState,
  element: XmlElement,
): ParseBlockContentState => ({
  ...state,
  options: {
    ...state.options,
    rootXmlns: mergeXmlnsDeclarations(state.options.rootXmlns ?? {}, element),
  },
});

/**
 * The two block children this walk skips, and who reads them instead.
 *
 * Both are read off the container element rather than met here: the body's own
 * `w:sectPr` by the document parser, a cell's `w:tcPr` by the table parser.
 * They are in this map because it is total over the union every block
 * container shares, not because this walk reads them.
 *
 * Declared at module scope so the claim is registered when the module loads,
 * rather than the first time a document is parsed.
 */
const BLOCK_CHILD_OWNERS = {
  sectPr: ownedElsewhere({
    container: "block-content",
    child: "sectPr",
    reader: "documentParser#parseDocumentBody",
  }),
  tcPr: ownedElsewhere({
    container: "block-content",
    child: "tcPr",
    reader: "tableParser#parseTableCell",
  }),
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
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
    // Accumulate the container's own xmlns onto the inherited in-scope set so a
    // captured VML `w:pict` replay resolves prefixes scoped on this level too.
    options: {
      ...options,
      rootXmlns: mergeXmlnsDeclarations(options?.rootXmlns ?? {}, parent),
      previews: options?.previews ?? standalonePreviewLedger(),
    },
  });

/** The parts and maps a block sequence is read against. */
type BlockContentResources = {
  styles: StyleMap | null;
  theme: Theme | null;
  numbering: NumberingMap | null;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
};

/** One block container's children as they are read. */
type BlockContentWalk = {
  resources: BlockContentResources;
  state: ParseBlockContentState;
  modelled: BlockContent[];
};

const BLOCK_CONTENT_UNDECLARED = {
  // `mc:AlternateContent`, whose selected branch folio reads. The others
  // are undeclared in the `w:` sense only because they belong to another
  // namespace, and the sink's default is right for them.
  AlternateContent: (
    child,
    { resources: { styles, theme, numbering, rels, media }, state, modelled },
  ) => {
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
} as const satisfies Record<string, ChildReader<BlockContentWalk>>;

const parseBlockCustomXml = (
  child: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  state: ParseBlockContentState,
): BlockCustomXml => {
  return {
    type: "blockCustomXml",
    ...blockCustomXmlShell(child),
    content: parseBlockContentWithState(
      child,
      styles,
      theme,
      numbering,
      rels,
      media,
      withContainerXmlns(state, child),
    ),
  };
};

export const BLOCK_CONTENT_HANDLERS = {
  p: (child, { resources: { styles, theme, numbering, rels, media }, state, modelled }) => {
    const paragraph = parseParagraph(child, styles, theme, numbering, rels, media, {
      ...state.options,
      runConsolidation: "deferred",
    });
    enrichParagraphTextBoxes(
      paragraph,
      child,
      styles,
      theme,
      numbering,
      rels,
      media,
      parseTable,
      state.options.previews,
      state.options.context,
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
  tbl: (child, { resources: { styles, theme, numbering, rels, media }, state, modelled }) => {
    const table = parseTable(child, styles, theme, numbering, rels, media, state.options);
    if (!table) {
      const hasUnmodeledTableContent =
        child.elements?.some((tableChild) => {
          const localName = getLocalName(tableChild.name ?? "");
          return localName !== "tblPr" && localName !== "tblGrid";
        }) ?? false;
      return hasUnmodeledTableContent ? CAPTURE : undefined;
    }
    modelled.push(table);
    return undefined;
  },
  sdt: (child, { resources: { styles, theme, numbering, rels, media }, state, modelled }) => {
    modelled.push(parseBlockSdt(child, styles, theme, numbering, rels, media, state));
  },
  customXml: (child, { resources: { styles, theme, numbering, rels, media }, state, modelled }) => {
    modelled.push(parseBlockCustomXml(child, styles, theme, numbering, rels, media, state));
  },
  // A block container declares the marker beside its blocks, so the model
  // keeps it there: it is a block in its own right, between the same two
  // siblings the source wrote it between. Re-anchoring it into a
  // neighbouring paragraph saved the element and changed the range.
  bookmarkStart: (child, { modelled }) => {
    modelled.push(parseBookmarkStart(child));
  },
  bookmarkEnd: (child, { modelled }) => {
    modelled.push(parseBookmarkEnd(child));
  },
  ...BLOCK_CHILD_OWNERS,
  altChunk: CAPTURE,
  commentRangeEnd: CAPTURE,
  commentRangeStart: CAPTURE,
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
} as const satisfies ChildHandlers<"block-content", BlockContentWalk>;

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

  const preserved = dispatchChildrenWithContext({
    element: parent,
    container: "block-content",
    capturePosition: () => modelled.length,
    undeclared: BLOCK_CONTENT_UNDECLARED,
    handlers: BLOCK_CONTENT_HANDLERS,
    context: { resources: { styles, theme, numbering, rels, media }, state, modelled },
  });

  return withPreservedChildren(
    modelled,
    preserved,
    (xml): PreservedBlock => ({ type: "preservedBlock", xml }),
  );
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
  const sdtContent = findWordprocessingChild(child, "sdtContent");
  const properties = parseSdtProperties(
    findWordprocessingChild(child, "sdtPr"),
    findWordprocessingChild(child, "sdtEndPr"),
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
