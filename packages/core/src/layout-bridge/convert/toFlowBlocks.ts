/**
 * ProseMirror to FlowBlock Converter
 *
 * Converts a ProseMirror document into FlowBlock[] for the layout engine.
 * Tracks pmStart/pmEnd positions for click-to-position mapping.
 */

import type { Node as PMNode } from "prosemirror-model";
import { resolveDocumentGridLinePitch } from "../../docx/documentGrid";
import { buildPageBreakRunDescendantIndex } from "../../internal/pageBreakRunDescendantIndex";
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
  PageBreakBlock,
  ColumnBreakBlock,
  SectionBreakBlock,
  ParagraphAttrs,
  SdtGroup,
} from "../../layout-engine/types";
import { createStyleEngine } from "../../style-engine";
import { getColumns } from "../sectionColumns";
import {
  expectBlockSdtAttrs,
  expectHardBreakAttrs,
  expectParagraphAttrs,
  expectTextBoxAttrs,
} from "../../prosemirror/attrs";
import { expectTextBoxAnchorAttrs } from "../../prosemirror/textBoxAnchorAttrs";
import { getPageNumbering } from "../../paged-layout/sectionGeometry";
import { assertValidProseMirrorDocument } from "../../prosemirror/validation";
import { cloneListCounterState, type ListCounterState } from "../../prosemirror/listMarker";
import { resolveNumberedRefFields } from "../../prosemirror/numberedRefFields";
import { groupParagraphFrames } from "./paragraphFrames";
import { twipsToPixels, nextBlockId, resetBlockIdCounter } from "./flowConversionShared";
import type { ToFlowBlocksOptions, FlowConversionOptions } from "./flowConversionShared";
import { convertImage } from "./imageConversion";
import { collectNoteRefs } from "./noteReferences";
import type { PageBreakRunProjection } from "./paragraphRuns";
import { drawsParagraphBorder, convertParagraph } from "./paragraphConversion";
import { splitParagraphAtPageBreaks } from "./pageBreakSplitting";
import {
  suppressFinalEmptyParagraphAfterTable,
  suppressFinalParagraphInRepeatedEmptySuffix,
  reserveLeadingEmptyOutlineHeight,
  mergeRunInParagraphs,
} from "./flowBlockPostprocessing";
import {
  getLastMapKey,
  applySectionStartsToBoundaries,
  coalesceTrailingPageBreakBeforeContinuousSection,
  readFinalSectionStart,
  applySectionDocumentGrid,
} from "./sectionBoundaries";
import { convertTable, convertTextBoxNode } from "./tableConversion";

export { formatCounter, resolveListTemplate } from "../../prosemirror/listMarker";
export { resetBlockIdCounter } from "./flowConversionShared";
export type { ToFlowBlocksOptions } from "./flowConversionShared";
export { collectNoteRefs, isNoteReferenceMarkXml } from "./noteReferences";
export { convertBorderSpecToLayout } from "./flowBorders";

/**
 * Word's layout fallback when an imported OOXML style hierarchy never supplies
 * `w:sz`. ECMA-376 deliberately leaves that terminal fallback to the consumer;
 * this value matches Word and is therefore applied by the DOCX composition
 * roots, while standalone ProseMirror documents keep the layout engine's 11pt
 * default.
 */
export const WORD_UNSPECIFIED_FONT_SIZE = 10;

/**
 * Convert a ProseMirror document to FlowBlock array.
 *
 * Walks the document tree, converting each node to the appropriate block type.
 * Tracks pmStart/pmEnd positions for each block for click-to-position mapping.
 */
export function toFlowBlocks(doc: PMNode, options: ToFlowBlocksOptions = {}): FlowBlock[] {
  assertValidProseMirrorDocument(doc, "Cannot layout invalid ProseMirror document");

  resetBlockIdCounter();

  const listCounters = options.listCounters ?? new Map<number, number[]>();
  const originalListCounters = options.originalListCounters ?? new Map<number, number[]>();
  const lastAdvancedNumId = getLastMapKey(listCounters);
  const lastAdvancedOriginalNumId = getLastMapKey(originalListCounters);
  const listCounterState: ListCounterState = {
    counters: listCounters,
    abstractCounters: options.listAbstractCounters ?? new Map<number, number[]>(),
    seenLevels: options.listSeenNumIds ?? new Set<string>(),
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
    ...(lastAdvancedNumId !== undefined ? { lastAdvancedNumId } : {}),
  };
  const originalListCounterState: ListCounterState = {
    counters: originalListCounters,
    abstractCounters: options.originalListAbstractCounters ?? new Map<number, number[]>(),
    seenLevels: options.originalListSeenNumIds ?? new Set<string>(),
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
    ...(lastAdvancedOriginalNumId !== undefined
      ? { lastAdvancedNumId: lastAdvancedOriginalNumId }
      : {}),
  };

  const { firstPageBreakRunPosition } = buildPageBreakRunDescendantIndex(doc);

  const opts: FlowConversionOptions = {
    ...options,
    listCounters: listCounterState.counters,
    listAbstractCounters: listCounterState.abstractCounters,
    listSeenNumIds: listCounterState.seenLevels,
    originalListCounters: originalListCounterState.counters,
    originalListAbstractCounters: originalListCounterState.abstractCounters,
    originalListSeenNumIds: originalListCounterState.seenLevels,
    listCounterStreams: {
      final: listCounterState,
      original: originalListCounterState,
    },
    firstPageBreakRunPosition,
    textBoxAnchorBlockIds: new Map(),
    numberingTabIgnoresIndent: doc.attrs["_doNotUseIndentAsNumberingTabStop"] === true,
    styleResolver: createStyleEngine(options.styles),
    numberedRefResults: resolveNumberedRefFields(doc, {
      listCounterState: cloneListCounterState(listCounterState),
      originalListCounterState: cloneListCounterState(originalListCounterState),
    }),
  };

  const blocks: FlowBlock[] = [];
  const offset = 0; // Start at document beginning; kept for clarity.
  void offset;
  let lastSectionMarginsTwips = {
    top: 1440,
    bottom: 1440,
    left: 1440,
    right: 1440,
  };

  let sdtSeq = 0;
  const sdtStack: SdtGroup[] = [];

  /**
   * Stamp the active SDT stack (outer→inner) onto every block produced by the
   * current call. ParagraphBlock/TableBlock accept `sdtGroups`; section breaks
   * and page breaks deliberately do not — they bracket layout, not content.
   */
  const tagBlockWithSdtStack = (block: FlowBlock): void => {
    if (sdtStack.length === 0) {
      return;
    }
    if (block.kind === "paragraph" || block.kind === "table") {
      block.sdtGroups = [...sdtStack];
    }
  };

  const trackedPush = (block: FlowBlock): void => {
    tagBlockWithSdtStack(block);
    blocks.push(block);
  };

  const pushParagraphProjection = (
    node: PMNode,
    pos: number,
    stripLeadingLineBreak = false,
  ): void => {
    const pageBreaks: PageBreakRunProjection[] = [];
    const paragraph = convertParagraph(node, pos, opts, pageBreaks);

    if (stripLeadingLineBreak && paragraph.runs.at(0)?.kind === "lineBreak") {
      paragraph.runs.shift();
    }
    const projected = splitParagraphAtPageBreaks({
      pageBreaks,
      paragraph,
      splitPageBreakAndParagraphMark: options.splitPageBreakAndParagraphMark === true,
    });
    if (pageBreaks.length > 0) {
      placeTextBoxAnchorsBeforePageBreaks(node, pos, projected);
    }
    for (const block of projected) {
      trackedPush(block);
    }
  };

  /** Blocks emitted before the current section's first block. */
  let sectionStartBlockCount = 0;
  /** Section breaks whose empty w:sectPr paragraph projected no block. */
  const sectionBreaksWithoutMarker = new Set<SectionBreakBlock["id"]>();
  /**
   * The page break a text box anchored in an earlier part of its paragraph
   * must be laid out before, keyed by `_docxAnchorId`.
   */
  const textBoxPageBreakByAnchorId = new Map<string, PageBreakBlock>();
  /** Text boxes held back until every block is projected. */
  const textBoxesBeforePageBreak = new Map<PageBreakBlock, TextBoxBlock[]>();

  /**
   * An anchored object sits on the page holding its anchor run. When that run
   * precedes a `w:br w:type="page"` in the same paragraph, the object belongs
   * to the page the break ends. Either way its paragraph-relative position is
   * measured from the part of the paragraph holding the anchor.
   */
  const placeTextBoxAnchorsBeforePageBreaks = (
    node: PMNode,
    pos: number,
    projected: readonly (ParagraphBlock | PageBreakBlock)[],
  ): void => {
    // Anchors come in document order, so one forward sweep finds each one's
    // next page break.
    let breakIndex = 0;
    node.descendants((child, childOffset) => {
      if (child.type.name !== "textBoxAnchor") {
        return true;
      }
      const anchorPos = pos + 1 + childOffset;
      while (breakIndex < projected.length) {
        const block = projected[breakIndex];
        if (block?.kind === "pageBreak" && (block.pmStart ?? 0) > anchorPos) {
          break;
        }
        breakIndex += 1;
      }
      const { anchorId } = expectTextBoxAnchorAttrs(child);
      const pageBreak = projected[breakIndex];
      if (pageBreak?.kind === "pageBreak") {
        textBoxPageBreakByAnchorId.set(anchorId, pageBreak);
      }
      const anchorFragment = pageBreak ? projected[breakIndex - 1] : projected.at(-1);
      if (anchorFragment?.kind === "paragraph") {
        opts.textBoxAnchorBlockIds.set(anchorId, anchorFragment.id);
      } else {
        // The part of the paragraph holding the anchor paints no line, so it
        // starts where the flow stands when the box is placed.
        opts.textBoxAnchorBlockIds.delete(anchorId);
      }
      return false;
    });
  };

  const pushTextBox = (node: PMNode, pos: number): void => {
    const textBox = convertTextBoxNode(node, pos, opts);
    const anchorId = expectTextBoxAttrs(node)._docxAnchorId;
    const pageBreak = anchorId === undefined ? undefined : textBoxPageBreakByAnchorId.get(anchorId);
    if (!pageBreak) {
      trackedPush(textBox);
      return;
    }
    tagBlockWithSdtStack(textBox);
    const heldBack = textBoxesBeforePageBreak.get(pageBreak);
    if (heldBack) {
      heldBack.push(textBox);
    } else {
      textBoxesBeforePageBreak.set(pageBreak, [textBox]);
    }
  };

  /** Put every held-back text box just before its page break, in one pass. */
  const releaseTextBoxesBeforePageBreaks = (): void => {
    if (textBoxesBeforePageBreak.size === 0) {
      return;
    }
    const inFlowOrder = blocks.splice(0);
    for (const block of inFlowOrder) {
      const heldBack = block.kind === "pageBreak" ? textBoxesBeforePageBreak.get(block) : undefined;
      for (const textBox of heldBack ?? []) {
        blocks.push(textBox);
      }
      blocks.push(block);
    }
  };

  const trailingPageBreakSectionPositions = new Set<number>();
  const consumedPageBreakPositions = new Set<number>();
  const collectTrailingPageBreakSections = (parent: PMNode, contentStart: number): void => {
    let childStart = contentStart;
    for (let index = 0; index < parent.childCount; index += 1) {
      const child = parent.child(index);
      const next = index + 1 < parent.childCount ? parent.child(index + 1) : undefined;
      const paragraphAttrs =
        child.type.name === "paragraph" ? expectParagraphAttrs(child) : undefined;
      if (
        paragraphAttrs &&
        next?.type.name === "pageBreak" &&
        paragraphAttrs._trailingPageBreak === true &&
        paragraphAttrs._sectionProperties !== undefined
      ) {
        trailingPageBreakSectionPositions.add(childStart);
        consumedPageBreakPositions.add(childStart + child.nodeSize);
      }
      if (child.type.name === "blockSdt") {
        collectTrailingPageBreakSections(child, childStart + 1);
      }
      childStart += child.nodeSize;
    }
  };
  collectTrailingPageBreakSections(doc, offset);

  // Refactored visit-style traversal so blockSdt can recurse into its
  // children without duplicating the per-block conversion code.
  const visit = (node: PMNode, pos: number): void => {
    switch (node.type.name) {
      case "blockSdt": {
        const attrs = expectBlockSdtAttrs(node);
        sdtSeq += 1;
        const group: SdtGroup = {
          id: `sdt-${sdtSeq}`,
          pmPos: pos,
          sdtType: attrs.sdtType,
        };
        if (attrs.alias) {
          group.alias = attrs.alias;
        }
        if (attrs.tag) {
          group.tag = attrs.tag;
        }
        if (typeof attrs.id === "number") {
          group.sdtId = attrs.id;
        }
        if (attrs.lock) {
          group.lock = attrs.lock;
        }
        if (attrs.showingPlaceholder) {
          group.showingPlaceholder = true;
        }
        if (typeof attrs.checked === "boolean") {
          group.checked = attrs.checked;
        }
        if (attrs.dateFormat) {
          group.dateFormat = attrs.dateFormat;
        }
        if (attrs.listItems) {
          group.listItemsJson = attrs.listItems;
        }

        sdtStack.push(group);
        const startIndex = blocks.length;
        let childOffset = pos + 1; // skip the blockSdt opening token
        for (let i = 0; i < node.childCount; i += 1) {
          const child = node.child(i);
          visit(child, childOffset);
          childOffset += child.nodeSize;
        }
        sdtStack.pop();
        // Stamp first/middle/last/only on the innermost group of each block
        // that was emitted inside this SDT so the painter chrome continues
        // visually across the block sequence. Only paragraph/table blocks
        // carry sdtGroups; section breaks etc. were skipped at tag time.
        const groupBlocks: (ParagraphBlock | TableBlock)[] = [];
        for (let i = startIndex; i < blocks.length; i += 1) {
          const b = blocks[i];
          if (b && (b.kind === "paragraph" || b.kind === "table")) {
            groupBlocks.push(b);
          }
        }
        // We're finalizing the SDT that `group` represents — locate that
        // entry by `pmPos` instead of always taking `at(-1)`. For blocks
        // that sit inside an inner SDT, `at(-1)` is the inner group, and
        // overwriting it would clobber the inner SDT's first/middle/last
        // markers when the outer iterates the same range later. Matching
        // by pmPos keeps the inner positions intact.
        const ourPmPos = group.pmPos;
        for (let i = 0; i < groupBlocks.length; i += 1) {
          const b = groupBlocks[i];
          if (!b || !b.sdtGroups) {
            continue;
          }
          const idx = b.sdtGroups.findIndex((g) => g.pmPos === ourPmPos);
          if (idx === -1) {
            continue;
          }
          let position: NonNullable<SdtGroup["position"]>;
          if (groupBlocks.length === 1) {
            position = "only";
          } else if (i === 0) {
            position = "first";
          } else if (i === groupBlocks.length - 1) {
            position = "last";
          } else {
            position = "middle";
          }
          // Replace just the entry for this SDT, leaving inner/outer
          // sibling entries untouched. (SdtGroup objects are shared
          // across blocks of the same group; copy-on-write here keeps
          // the other blocks' references stable.)
          const next = [...b.sdtGroups];
          const existing = next[idx];
          if (!existing) {
            continue;
          }
          next[idx] = { ...existing, position };
          b.sdtGroups = next;
        }
        return;
      }
      default:
        break;
    }

    switch (node.type.name) {
      case "paragraph": {
        const pmAttrs = expectParagraphAttrs(node);
        const secProps = pmAttrs._sectionProperties;
        const hasSectionBreak = secProps !== undefined;
        const hasListFormatting =
          (pmAttrs.numPr !== null && pmAttrs.numPr !== undefined) ||
          (pmAttrs.listMarker !== null && pmAttrs.listMarker !== undefined);
        const firstChild = node.firstChild;
        const startsWithColumnBreak =
          firstChild?.type.name === "hardBreak" &&
          expectHardBreakAttrs(firstChild).breakType === "column";
        const isStandaloneColumnBreak = node.childCount === 1 && startsWithColumnBreak;
        const isEmptySectionMark = hasSectionBreak && node.content.size === 0;
        const opensSection = blocks.length === sectionStartBlockCount;
        let markerDropped = false;

        if (isStandaloneColumnBreak) {
          const columnBreak: ColumnBreakBlock = {
            kind: "columnBreak",
            id: nextBlockId(),
            pmStart: pos,
            pmEnd: pos + node.nodeSize,
          };
          trackedPush(columnBreak);
        } else if (startsWithColumnBreak && firstChild) {
          const columnBreak: ColumnBreakBlock = {
            kind: "columnBreak",
            id: nextBlockId(),
            pmStart: pos + 1,
            pmEnd: pos + 1 + firstChild.nodeSize,
          };
          trackedPush(columnBreak);

          pushParagraphProjection(node, pos, true);
        } else if (!isEmptySectionMark) {
          pushParagraphProjection(node, pos);
        } else if (hasListFormatting || drawsParagraphBorder(pmAttrs.borders) || opensSection) {
          // An empty w:sectPr paragraph is only a section marker, except when
          // it paints a number or a border rule, or when it is its section's
          // only block and so is that section's content.
          pushParagraphProjection(node, pos);
          const mark = blocks.at(-1);
          if (!opensSection && mark?.kind === "paragraph" && mark.attrs?.pageBreakBefore) {
            // After its section's content the marker ends the section where
            // that content ended; w:pageBreakBefore does not move it on.
            delete mark.attrs.pageBreakBefore;
          }
        } else {
          markerDropped = true;
        }

        // Emit section break block if this paragraph ends a section
        if (hasSectionBreak) {
          if (trailingPageBreakSectionPositions.has(pos)) {
            const sourceParagraph = blocks.at(-1);
            if (sourceParagraph?.kind === "paragraph" && sourceParagraph.pmStart === pos) {
              const pageBreak: PageBreakBlock = {
                kind: "pageBreak",
                id: nextBlockId(),
                pmStart: pos + node.nodeSize,
                pmEnd: pos + node.nodeSize + 1,
              };
              trackedPush(pageBreak);
              const sourceAttrs = sourceParagraph.attrs;
              const carrierSpacing = sourceAttrs?.spacing ? { ...sourceAttrs.spacing } : undefined;
              if (carrierSpacing) {
                delete carrierSpacing.before;
              }
              const carrierAttrs: ParagraphAttrs = {
                ...(carrierSpacing ? { spacing: carrierSpacing } : {}),
                ...(sourceAttrs?.automaticSpacing?.after === true
                  ? { automaticSpacing: { after: true } }
                  : {}),
                ...(sourceAttrs?.spacingExplicit?.after === true
                  ? { spacingExplicit: { after: true } }
                  : {}),
                ...(sourceAttrs?.hasDirectParagraphFormatting === true
                  ? { hasDirectParagraphFormatting: true }
                  : {}),
                ...(sourceAttrs?.hasDirectParagraphMarkFormatting === true
                  ? { hasDirectParagraphMarkFormatting: true }
                  : {}),
                ...(sourceAttrs?.snapToGrid !== undefined
                  ? { snapToGrid: sourceAttrs.snapToGrid }
                  : {}),
                ...(sourceAttrs?.documentGridLinePitch !== undefined
                  ? { documentGridLinePitch: sourceAttrs.documentGridLinePitch }
                  : {}),
                ...(sourceAttrs?.defaultFontSize !== undefined
                  ? { defaultFontSize: sourceAttrs.defaultFontSize }
                  : {}),
                ...(sourceAttrs?.defaultFontFamily !== undefined
                  ? { defaultFontFamily: sourceAttrs.defaultFontFamily }
                  : {}),
                ...(sourceAttrs?.defaultAlternateFontFamily !== undefined
                  ? { defaultAlternateFontFamily: sourceAttrs.defaultAlternateFontFamily }
                  : {}),
              };
              const carrier: ParagraphBlock = {
                kind: "paragraph",
                id: nextBlockId(),
                runs: [],
                attrs: carrierAttrs,
                pmStart: pos + node.nodeSize,
                pmEnd: pos + node.nodeSize,
              };
              trackedPush(carrier);
            }
          }
          const sectionBreak: SectionBreakBlock = {
            kind: "sectionBreak",
            id: nextBlockId(),
          };
          const breakType = secProps?.sectionStart;
          if (breakType) {
            sectionBreak.type = breakType;
          }

          if (secProps) {
            sectionBreak.pageNumbering = getPageNumbering(secProps);
            const documentGridLinePitchTwips = resolveDocumentGridLinePitch(secProps.docGrid);
            if (documentGridLinePitchTwips !== undefined) {
              sectionBreak.documentGridLinePitchTwips = documentGridLinePitchTwips;
            }
            // Populate page size
            if (secProps.pageWidth || secProps.pageHeight) {
              sectionBreak.pageSize = {
                w: twipsToPixels(secProps.pageWidth ?? 12_240),
                h: twipsToPixels(secProps.pageHeight ?? 15_840),
              };
            }
            // Populate margins
            if (
              secProps.marginTop !== undefined ||
              secProps.marginBottom !== undefined ||
              secProps.marginLeft !== undefined ||
              secProps.marginRight !== undefined
            ) {
              lastSectionMarginsTwips = {
                top: secProps.marginTop ?? lastSectionMarginsTwips.top,
                bottom: secProps.marginBottom ?? lastSectionMarginsTwips.bottom,
                left: secProps.marginLeft ?? lastSectionMarginsTwips.left,
                right: secProps.marginRight ?? lastSectionMarginsTwips.right,
              };
              sectionBreak.margins = {
                top: twipsToPixels(lastSectionMarginsTwips.top),
                bottom: twipsToPixels(lastSectionMarginsTwips.bottom),
                left: twipsToPixels(lastSectionMarginsTwips.left),
                right: twipsToPixels(lastSectionMarginsTwips.right),
              };
              if (secProps.headerDistance !== undefined) {
                sectionBreak.margins.header = twipsToPixels(secProps.headerDistance);
              }
              if (secProps.footerDistance !== undefined) {
                sectionBreak.margins.footer = twipsToPixels(secProps.footerDistance);
              }
            }
            // Populate columns
            const columns = getColumns(secProps);
            if (columns) {
              sectionBreak.columns = columns;
            }
          }

          trackedPush(sectionBreak);
          sectionStartBlockCount = blocks.length;
          if (markerDropped) {
            sectionBreaksWithoutMarker.add(sectionBreak.id);
          }
        }
        break;
      }

      case "table":
        trackedPush(convertTable(node, pos, opts));
        break;

      case "image": {
        // Standalone image block (if not inline)
        const image = convertImage(node, pos, opts.pageContentHeight);
        if (image !== undefined) {
          trackedPush(image);
        }
        break;
      }

      case "textBox":
        pushTextBox(node, pos);
        break;

      case "horizontalRule":
      case "pageBreak": {
        if (consumedPageBreakPositions.has(pos)) {
          break;
        }
        const pb: PageBreakBlock = {
          kind: "pageBreak",
          id: nextBlockId(),
          pmStart: pos,
          pmEnd: pos + node.nodeSize,
        };
        trackedPush(pb);
        break;
      }
      default:
        break;
    }
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  doc.forEach((node, nodeOffset) => {
    visit(node, offset + nodeOffset);
  });
  releaseTextBoxesBeforePageBreaks();

  reserveLeadingEmptyOutlineHeight(blocks);
  // The terminal-anchor collapse is for the document's last paragraph; one
  // followed by an endnote area is not it.
  const trailingEndnoteIds = options.trailingEndnoteIds;
  const endsInEndnoteArea =
    trailingEndnoteIds !== undefined &&
    trailingEndnoteIds.size > 0 &&
    collectNoteRefs(blocks, "endnoteRefId").some(({ noteId }) => trailingEndnoteIds.has(noteId));
  if (!endsInEndnoteArea) {
    suppressFinalEmptyParagraphAfterTable(blocks);
    suppressFinalParagraphInRepeatedEmptySuffix(blocks);
  }
  const mergedBlocks = mergeRunInParagraphs(blocks);
  const tableCellLinePitch =
    doc.attrs["_adjustLineHeightInTable"] === true ? "sectionGrid" : undefined;
  const griddedBlocks = applySectionDocumentGrid(mergedBlocks, {
    finalLinePitchTwips: opts.finalSectionDocumentGridLinePitchTwips,
    tableCellLinePitch,
  });
  const boundaryBlocks = applySectionStartsToBoundaries(griddedBlocks, readFinalSectionStart(doc));
  const reconciledBlocks = coalesceTrailingPageBreakBeforeContinuousSection(
    boundaryBlocks,
    options.splitPageBreakAndParagraphMark === true,
    sectionBreaksWithoutMarker,
  );
  return groupParagraphFrames(reconciledBlocks, nextBlockId);
}
