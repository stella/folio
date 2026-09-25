/**
 * Converts a paragraph's inline content (text, tabs, breaks, fields, symbols,
 * images, note references) to layout runs, projecting page breaks and
 * inline marks.
 */

import type { Node as PMNode, Mark } from "prosemirror-model";
import { getFontAlternate } from "../../fonts/fontAlternates";
import type { Run, TextRun, TabRun, FieldRun, RunFormatting } from "../../layout-engine/types";
import {
  expectFieldAttrs,
  expectImageAttrs,
  expectMathAttrs,
  expectParagraphAttrs,
  expectPreservedXmlAttrs,
  expectSymbolAttrs,
} from "../../prosemirror/attrs";
import { decodeOoxmlSymbolCharacter } from "../../utils/ooxmlSymbol";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import type { FlowConversionOptions } from "./flowConversionShared";
import {
  markDefaultBlackTextColorSource,
  mergeRunFormatting,
  applyCharacterStyleToggleFormatting,
} from "./runFormattingMerge";
import { paragraphRunDefaults } from "./textFormattingConversion";
import { extractRunFormatting } from "./runMarkFormatting";
import {
  constrainImageToPage,
  buildImageRun,
  hasRelationshipBackedImageBox,
  hasPaintableImageSource,
} from "./imageConversion";
import { isNoteReferenceMarkXml } from "./noteReferences";

/**
 * In TOC paragraphs, strip the resolved Hyperlink character-style colour and
 * underline so the painter's link fallback doesn't fire. The PM doc keeps the
 * original marks so copy/paste out of a TOC still carries the Hyperlink
 * styling like Word does. Applies to both text and field runs — a TOC entry's
 * page number is a PAGEREF field inside the entry's hyperlink.
 *
 * Mutates `formatting` in place; cheaper than re-cloning per run.
 */
function stripTocHyperlinkStyle(formatting: RunFormatting): void {
  if (!formatting.hyperlink) {
    return;
  }
  formatting.hyperlink.noDefaultStyle = true;
  delete formatting.color;
  delete formatting.underline;
}

/**
 * Convert a paragraph node to runs.
 */
export type PageBreakRunProjection = {
  pmStart: number;
  pmEnd: number;
  trackedChange?: Pick<
    RunFormatting,
    | "isInsertion"
    | "changeAuthor"
    | "changeDate"
    | "changeRevisionId"
    | "isSuggestion"
    | "suggestionId"
  >;
};

const isTrackedRunMark = ({ type }: Mark): boolean =>
  type.name === "insertion" || type.name === "deletion";

/**
 * Inline wrappers own marks that apply to every projected descendant. A
 * descendant's mark of the same type is more specific, and any descendant
 * revision replaces (rather than combines with) an inherited revision.
 */
const mergeProjectedInlineMarks = (
  inheritedMarks: readonly Mark[],
  ownMarks: readonly Mark[],
): readonly Mark[] => {
  if (inheritedMarks.length === 0) {
    return ownMarks;
  }

  const ownMarkTypes = new Set(ownMarks.map(({ type }) => type));
  const hasOwnTrackedMark = ownMarks.some(isTrackedRunMark);
  return [
    ...inheritedMarks.filter(
      (mark) => !ownMarkTypes.has(mark.type) && !(hasOwnTrackedMark && isTrackedRunMark(mark)),
    ),
    ...ownMarks,
  ];
};

export function paragraphToRuns(
  node: PMNode,
  startPos: number,
  _options: FlowConversionOptions,
  pageBreaks?: PageBreakRunProjection[],
): Run[] {
  const runs: Run[] = [];
  const offset = startPos + 1; // +1 for opening tag
  const theme = _options.theme;
  const fontAlternates = _options.fontAlternates;
  const pmAttrs = expectParagraphAttrs(node);
  const paraDefaults = paragraphRunDefaults(pmAttrs, theme, fontAlternates);
  const paragraphStyleId = pmAttrs.styleId;
  const inTocParagraph =
    pmAttrs._tableOfContentsLevel !== undefined ||
    tableOfContentsStyleLevel({ styleId: paragraphStyleId }) !== undefined;
  let leadingRenderedPageBreakPending = pmAttrs.renderedPageBreakBefore === true;

  // Single dispatcher for one inline PM child. Recurses on `sdt` so nested
  // content controls keep contributing runs at the right pmStart/pmEnd.
  // Used for both the top-level paragraph iteration and the descent into
  // SDT children — the previous SDT branch only handled text/hardBreak/
  // tab/image and silently dropped fields, math, and nested SDTs even
  // when the parser preserved them (see eigenpal #482).
  function pushRunsForChild(
    child: PMNode,
    childPos: number,
    inheritedMarks: readonly Mark[] = [],
  ): void {
    if (child.type.name === "bookmarkBoundary") {
      return;
    }
    if (child.type.name === "renderedPageBreak") {
      if (leadingRenderedPageBreakPending) {
        leadingRenderedPageBreakPending = false;
        return;
      }
      runs.push({
        kind: "renderedPageBreak",
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    const effectiveMarks = mergeProjectedInlineMarks(inheritedMarks, child.marks);
    if (child.type.name === "pageBreakRun") {
      const trackedChange = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      if (trackedChange.isDeletion) {
        runs.push({
          kind: "text",
          text: "",
          isDeletion: true,
          ...(trackedChange.changeAuthor !== undefined
            ? { changeAuthor: trackedChange.changeAuthor }
            : {}),
          ...(trackedChange.changeDate !== undefined
            ? { changeDate: trackedChange.changeDate }
            : {}),
          ...(trackedChange.changeRevisionId !== undefined
            ? { changeRevisionId: trackedChange.changeRevisionId }
            : {}),
          ...(trackedChange.isSuggestion === true ? { isSuggestion: true } : {}),
          ...(trackedChange.suggestionId !== undefined
            ? { suggestionId: trackedChange.suggestionId }
            : {}),
          pmStart: childPos,
          pmEnd: childPos + child.nodeSize,
        });
        return;
      }
      pageBreaks?.push({
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
        ...(trackedChange.isInsertion
          ? {
              trackedChange: {
                isInsertion: true,
                ...(trackedChange.changeAuthor !== undefined
                  ? { changeAuthor: trackedChange.changeAuthor }
                  : {}),
                ...(trackedChange.changeDate !== undefined
                  ? { changeDate: trackedChange.changeDate }
                  : {}),
                ...(trackedChange.changeRevisionId !== undefined
                  ? { changeRevisionId: trackedChange.changeRevisionId }
                  : {}),
                ...(trackedChange.isSuggestion === true ? { isSuggestion: true } : {}),
                ...(trackedChange.suggestionId !== undefined
                  ? { suggestionId: trackedChange.suggestionId }
                  : {}),
              },
            }
          : {}),
      });
      return;
    }
    if (child.type.name !== "sdt") {
      leadingRenderedPageBreakPending = false;
    }
    if (child.isText && child.text) {
      const formatting = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({
        formatting,
        marks: effectiveMarks,
        paragraphFormatting: pmAttrs.defaultTextFormatting,
        styleResolver: _options.styleResolver,
      });
      if (inTocParagraph) {
        stripTocHyperlinkStyle(formatting);
      }
      const run: TextRun = {
        kind: "text",
        text: child.text,
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
      return;
    }
    if (child.type.name === "symbol") {
      const attrs = expectSymbolAttrs(child);
      const text = decodeOoxmlSymbolCharacter(attrs.char);
      if (text === null) {
        return;
      }
      const formatting = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({
        formatting,
        marks: effectiveMarks,
        paragraphFormatting: pmAttrs.defaultTextFormatting,
        styleResolver: _options.styleResolver,
      });
      if (inTocParagraph) {
        stripTocHyperlinkStyle(formatting);
      }
      const alternateFontFamily = getFontAlternate(attrs.font, fontAlternates);
      formatting.fontFamily = attrs.font;
      if (alternateFontFamily) {
        formatting.alternateFontFamily = alternateFontFamily;
      }
      runs.push({
        kind: "text",
        text,
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "preservedXml") {
      // An opaque atom lays out as the text it puts on the line and nothing
      // else: a `w:ruby` base is a word the reader measures and clicks into,
      // while markup that paints nothing takes no space.
      const { xml, text: capturedText } = expectPreservedXmlAttrs(child);
      const text =
        capturedText === "" &&
        _options.noteReferenceMarkText !== undefined &&
        isNoteReferenceMarkXml(xml)
          ? _options.noteReferenceMarkText
          : capturedText;
      if (text === "") {
        return;
      }
      const formatting = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({
        formatting,
        marks: effectiveMarks,
        paragraphFormatting: pmAttrs.defaultTextFormatting,
        styleResolver: _options.styleResolver,
      });
      runs.push({
        kind: "text",
        text,
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "hardBreak") {
      runs.push({
        kind: "lineBreak",
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "tab") {
      const formatting = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({
        formatting,
        marks: effectiveMarks,
        paragraphFormatting: pmAttrs.defaultTextFormatting,
        styleResolver: _options.styleResolver,
      });
      const run: TabRun = {
        kind: "tab",
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
      return;
    }
    if (child.type.name === "image") {
      const attrs = expectImageAttrs(child);
      if (!hasPaintableImageSource(attrs) && !hasRelationshipBackedImageBox(attrs)) {
        // Unsupported DrawingML shapes can survive the parser as image nodes
        // without a relationship target or authored extent. They have no
        // paintable payload; a 100x100 fallback box would incorrectly consume
        // paragraph flow. Relationship-backed package images keep their real
        // line box even when the browser cannot paint that image format.
        return;
      }
      const constrained = constrainImageToPage(
        attrs.width ?? 100,
        attrs.height ?? 100,
        _options.pageContentHeight,
      );
      // Lift tracked-change marks off the image node so an inserted/deleted
      // picture paints in the revision colour and resolves with the rest of
      // the change. eigenpal #641.
      const trackedFmt = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      const run = buildImageRun(
        attrs,
        constrained,
        childPos,
        childPos + child.nodeSize,
        trackedFmt,
        _options.forceAnchorLayoutInCell,
      );
      runs.push(run);
      return;
    }
    if (child.type.name === "structuredField") {
      let containsPageBreak = false;
      child.descendants((descendant) => {
        if (descendant.type.name === "pageBreakRun") {
          containsPageBreak = true;
          return false;
        }
        return true;
      });
      if (containsPageBreak) {
        const fieldContentStart = childPos + 1;
        // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
        child.forEach((fieldChild, fieldChildOffset) => {
          pushRunsForChild(fieldChild, fieldContentStart + fieldChildOffset, effectiveMarks);
        });
        return;
      }
    }
    if (child.type.name === "field" || child.type.name === "structuredField") {
      // Marks on the field node (bold/italic/underline applied to the
      // field result inside `<w:fldChar separate>...</w:fldChar end>`)
      // must propagate to the run formatting, otherwise complex REF
      // fields whose visible text was authored as underlined (e.g.
      // cross-references like "Exhibit A" / "Section 1.3" in NVCA-style
      // templates) render with no underline. Reuse the same extractor
      // text runs use.
      const attrs = expectFieldAttrs(child);
      const ft = attrs.fieldType;
      let mappedType: FieldRun["fieldType"] = "OTHER";
      if (ft === "PAGE") {
        mappedType = "PAGE";
      } else if (ft === "NUMPAGES") {
        mappedType = "NUMPAGES";
      } else if (ft === "DATE") {
        mappedType = "DATE";
      } else if (ft === "TIME") {
        mappedType = "TIME";
      }
      const extractedFieldFormatting = extractRunFormatting(effectiveMarks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({
        formatting: extractedFieldFormatting,
        marks: effectiveMarks,
        paragraphFormatting: pmAttrs.defaultTextFormatting,
        styleResolver: _options.styleResolver,
      });
      if (inTocParagraph) {
        stripTocHyperlinkStyle(extractedFieldFormatting);
      }
      const fieldFormatting = markDefaultBlackTextColorSource(
        extractedFieldFormatting,
        paraDefaults,
      );
      const run: FieldRun = {
        kind: "field",
        fieldType: mappedType,
        instruction: attrs.instruction,
        fallback: _options.numberedRefResults?.get(child) ?? attrs.displayText ?? "",
        ...(attrs.fldLock ? { fldLock: true } : {}),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
        ...fieldFormatting,
      };
      runs.push(run);
      return;
    }
    if (child.type.name === "math") {
      const attrs = expectMathAttrs(child);
      const plainText = attrs.plainText || "[equation]";
      runs.push({
        kind: "math",
        display: attrs.display ?? "inline",
        ommlXml: attrs.ommlXml,
        plainText,
        italic: true,
        fontFamily: "Cambria Math",
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "sdt") {
      const sdtInnerOffset = childPos + 1; // +1 for opening tag
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
      child.forEach((sdtChild, sdtChildOffset) => {
        pushRunsForChild(sdtChild, sdtInnerOffset + sdtChildOffset, effectiveMarks);
      });
    }
  }

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, childOffset) => {
    pushRunsForChild(child, offset + childOffset);
  });

  const { defaultFont, defaultSize } = _options;
  if (defaultFont !== undefined || defaultSize !== undefined) {
    for (const run of runs) {
      switch (run.kind) {
        case "text":
        case "tab":
        case "renderedPageBreak":
        case "field":
        case "math":
          if (defaultFont !== undefined) {
            run.fontFamily ??= defaultFont;
          }
          if (defaultSize !== undefined) {
            run.fontSize ??= defaultSize;
          }
          break;
        case "image":
        case "lineBreak":
          break;
        default:
          run satisfies never;
      }
    }
  }

  return runs;
}
