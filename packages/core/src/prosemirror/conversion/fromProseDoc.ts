/**
 * ProseMirror to Document Conversion
 *
 * Converts a ProseMirror document back to our Document type.
 * This enables round-trip editing: DOCX -> Document -> PM -> Document -> DOCX
 *
 * Key responsibilities:
 * - Coalesce consecutive text with same marks into single Runs
 * - Preserve paragraph attributes (paraId, textId, formatting)
 * - Handle marks -> TextFormatting conversion
 */

import { panic } from "better-result";
import { DRAWING_RAW_XML_MODES, relationshipIdOf } from "@stll/docx-core/model";
import type { Node as PMNode, Mark } from "prosemirror-model";
import { Fragment } from "prosemirror-model";

import {
  isStyleSourcedParagraphNumbering,
  modelParagraphFormattingEmission,
} from "../../internal/paragraphFormattingSerialization";
import { joinCommentRangesAcrossParagraphs } from "../../docx/commentRangeJoin";
import { completeCommentReferences } from "../../docx/commentReferenceCompletion";
import { isInlineSdtContent } from "../../docx/inlineWrapperContent";
import { visitDocxParagraphs } from "../../docx/paragraphTraversal";
import {
  paragraphNumberingReferenceId,
  sameStatedParagraphNumbering,
} from "../../docx/numberingReference";
import { DATE_UTC_ATTRIBUTE } from "../../docx/trackedChangeInfo";
import { createStyleEngine, type StyleEngine } from "../../style-engine";
import {
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  ParagraphPropertySourceValidationError,
  type ParagraphPropertySourceValidationCode,
  type TableCellParagraphPropertySourceBinding,
  copyDocumentParagraphPropertySourceContract,
  copyDocumentParagraphPropertySources,
  copyParagraphPropertyCapture,
  copyParagraphPropertySource,
  decodeTableCellParagraphSourcePayload,
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceCandidate,
  getParagraphPropertySourceToken,
  getParagraphPropertySourceTransferId,
  getProseDocumentParagraphPropertySourceContract,
  getProseParagraphPropertySourceToken,
  isParagraphPropertySourceToken,
  linkParagraphPropertySourceCandidate,
  paragraphPropertySourceTokenMatchesContract,
  paragraphPropertySourceBelongsToDocument,
  recreateProseNodeWithParagraphPropertySource,
  restoreTableCellsWithParagraphPropertySources,
  visitDocumentStoryParagraphs,
  visitTableCellParagraphPropertySourceBindings,
} from "../../docx/paragraphPropertySource";
import { canonicalJson } from "../../utils/canonicalJson";
import { EDITED_PREVIEW_FINGERPRINT, imageRawXmlFingerprint } from "../../docx/imageRawXml";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import { readAuthoredTransform } from "../authoredTransformAttrs";
import { parseShapeGeometryAdjustments } from "../shapeGeometryAdjustments";
import { copiedWrapPolygon } from "../../docx/wrapPolygon";
import type { PresetLineDashVal } from "@stll/docx-core/model";

import type {
  ImageWrap,
  ImageWrapPolygon,
  EffectExtentSlots,
  WrapDistanceSlots,
  ImagePosition,
  ShapeFill,
  ShapeOutline,
  PositionedBookmarkMarker,
  PreservedAttribute,
  SectionProperties,
  SectionStart,
} from "../../types/content";
import type {
  BlockContent,
  TableCellBlock,
  BlockSdt,
  BookmarkEnd,
  BookmarkStart,
  Document,
  DocumentBody,
  Paragraph,
  ParagraphPropertyChange,
  PreservedInline,
  Run,
  TextFormatting,
  ParagraphFormatting,
  TextContent,
  BreakContent,
  TabContent,
  DrawingAnchor,
  DrawingContent,
  Image,
  Hyperlink,
  InlineWrapper,
  ParagraphContent,
  Table,
  TableRow,
  TableCell,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
  TableBorders,
  ShapeContent,
  Shape,
  SymbolContent,
  NoteReferenceContent,
  SimpleField,
  ComplexField,
  InlineSdt,
  SdtProperties,
  TrackedChangeInfo,
  MathEquation,
  ColorValue,
  CellMargins,
} from "../../types/document";
import { normalizeShapeTextAnchor, presetDashForOutlineAttr } from "../../types/documentEnumValues";
import { presetDashForCssBorderStyle } from "../../utils/borderCss";
import { emuToPixels, emuToStrokePixels, pixelsToEmu } from "../../utils/units";
import { bookmarkMarkerFromAttrs, expectBookmarkBoundaryAttrs } from "../bookmarkBoundaryAttrs";
import { expectCommentReferenceAttrs } from "../commentReferenceAttrs";
import {
  expectCharacterSpacingMarkAttrs,
  expectCharacterStyleMarkAttrs,
  expectCommentMarkAttrs,
  expectInlineWrapperMarkAttrs,
  expectEmphasisMarkAttrs,
  expectTextEffectMarkAttrs,
  expectFieldAttrs,
  expectFontFamilyMarkAttrs,
  expectLanguageMarkAttrs,
  expectFontSizeMarkAttrs,
  expectFootnoteRefMarkAttrs,
  expectHardBreakAttrs,
  expectHighlightMarkAttrs,
  expectRunShadingMarkAttrs,
  expectHyperlinkMarkAttrs,
  expectImageAttrs,
  expectMathAttrs,
  expectPageBreakRunAttrs,
  expectPageBreakRunOwnerMarkAttrs,
  expectParagraphAttrs,
  expectRunFormattingOverrideMarkAttrs,
  expectRunPropertyChangeMarkAttrs,
  expectBlockSdtAttrs,
  expectSdtAttrs,
  expectShapeAttrs,
  expectPreservedBlockAttrs,
  expectPreservedXmlAttrs,
  expectSymbolAttrs,
  expectStrikeMarkAttrs,
  expectTabAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
  expectTextBoxAttrs,
  expectTextColorMarkAttrs,
  expectTrackedChangeMarkAttrs,
  expectUnderlineMarkAttrs,
} from "../attrs";
import { autospacingMatchesBase, hasAutospacingBaseSide } from "../autospacingBase";
import { directionToBidi } from "../paragraphDirection";
import { directParagraphAlignment } from "../paragraphAlignment";
import type { StyleResolvedParagraphField } from "../paragraphFormattingProvenance";
import { directParagraphSpacing } from "../paragraphSpacing";
import {
  paragraphRejectAttrPatch,
  paragraphRejectOriginalFormatting,
  removeParagraphPropertyChanges,
} from "../commands/propertyChangeScope";
import { RUN_FORMATTING_MARK_NAMES } from "../runFormattingMarkNames";
import {
  authoredRunFormattingFromAttrs,
  hasAuthoredRunFormattingProvenance,
} from "../runFormattingProvenance";
import {
  paragraphFormattingForRun,
  paragraphRunStyleContext,
  resolveEffectiveRunStyleFormatting,
  suppressParagraphMarkFormatting,
  type RunStyleResolver,
} from "../runStyleFormatting";
import {
  applyRunFormattingOverrideAttrs,
  buildRunFormattingOverrideAttrs,
} from "../extensions/marks/RunFormattingOverrideExtension";
import { inlineWrapperMember, inlineWrapperStackKey } from "../inlineWrapperStack";
import { INLINE_WRAPPER_MARK_NAME } from "../extensions/marks/InlineWrapperExtension";
import { withMoveRanges } from "../moveRangeCarrier";
import { schema } from "../schema";
import type { InlineWrapperLayer, RunFormattingOverrideAttrs } from "../schema/marks";
import { PRESERVED_XML_LEVELS } from "../schema/nodes";
import type {
  ParagraphAttrs,
  ParagraphPropertyChangeAttrs,
  TableAttrs,
  TableRowAttrs,
  TableCellAttrs,
  AuthoredEmuAttrs,
  ImageAttrs,
  ImagePositionAttrs,
  TextBoxAttrs,
} from "../schema/nodes";
import { assertValidProseMirrorDocument } from "../validation";
import { listRenderingFromAttrs as listRenderingFieldsFromAttrs } from "../listRenderingAttrs";
import { resolveNumberedRefFields } from "../numberedRefFields";
import { expectTextBoxAnchorAttrs } from "../textBoxAnchorAttrs";
import { runShadingAttrsToShading, shadingToRunShadingAttrs } from "./runShadingMark";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import { decodeSdtListItems, sdtPropertiesFromAttrs, sdtPropertiesMatchAttrs } from "./sdtAttrs";
// `fromProseDoc` and `toProseDoc` are the two halves of one round-trip and
// already reference each other (`toProseDoc` imports `marksToTextFormatting`
// from here). Reusing the inverse converter to revert a stripped suggested
// run-property change closes that pair; both edges are call-time function
// references, so there is no initialization-order hazard.
// oxlint-disable-next-line import/no-cycle
import { textFormattingToMarks } from "../extensions/marks/markUtils";

/**
 * The dash an `outlineStyle` attribute names.
 *
 * `"none"` is the explicit no-outline sentinel rather than a dash, so it
 * returns undefined here: the serializer's no-outline guard drops the `<a:ln>`
 * before this is called.
 */
function normalizeShapeOutlineDash(style: string | undefined): PresetLineDashVal | undefined {
  const dash = presetDashForOutlineAttr(style);
  return dash === undefined || dash === "none" ? undefined : dash;
}

function imagePositionFromAttrs(attrs: ImagePositionAttrs | undefined): ImagePosition | undefined {
  const horizontalPosition = attrs?.horizontal;
  const verticalPosition = attrs?.vertical;
  if (!horizontalPosition || !verticalPosition) {
    return undefined;
  }

  const horizontal: ImagePosition["horizontal"] = {
    relativeTo: horizontalPosition.relativeTo || "column",
  };
  if (horizontalPosition.align) {
    horizontal.alignment = horizontalPosition.align;
  }
  if (horizontalPosition.posOffset !== undefined) {
    horizontal.posOffset = horizontalPosition.posOffset;
  }

  const vertical: ImagePosition["vertical"] = {
    relativeTo: verticalPosition.relativeTo || "paragraph",
  };
  if (verticalPosition.align) {
    vertical.alignment = verticalPosition.align;
  }
  if (verticalPosition.posOffset !== undefined) {
    vertical.posOffset = verticalPosition.posOffset;
  }

  return { horizontal, vertical };
}

/**
 * The EMUs behind one pixel attribute of a drawing.
 *
 * `toProseDoc` carries the authored EMU beside every pixel value it derived,
 * because EMU → px → EMU does not land back on the same number. While the
 * pixel attribute still projects from the authored EMU nothing has moved it,
 * so the document gets its own value back; once an editor command has, the
 * pixel attribute is the truth and converts.
 */
const emuFromPixels = <Key extends string>(
  px: number,
  key: Key,
  authored: AuthoredEmuAttrs<Key> | undefined,
  project: (emu: number) => number,
): number => {
  const emu = authored?.[key];
  return emu !== undefined && project(emu) === px ? emu : pixelsToEmu(px);
};

/** The attributes every drawing's wrap insets and polygon live in. */
type WrapDistanceSource = {
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
  wrapDistanceSlots?: WrapDistanceSlots;
  wrapEffectExtentSlots?: EffectExtentSlots;
  wrapPolygon?: ImageWrapPolygon;
  _docxAuthoredEmu?: AuthoredEmuAttrs<"distTop" | "distBottom" | "distLeft" | "distRight">;
};

/**
 * Copy the wrap insets, the slot each was authored on, and the wrap polygon
 * onto a wrap. Shared by the image, shape and text-box paths so one rule
 * decides when an inset keeps its authored EMU.
 */
const assignWrapDistances = (wrap: ImageWrap, attrs: WrapDistanceSource): void => {
  const authored = attrs._docxAuthoredEmu;
  if (attrs.distTop !== undefined) {
    wrap.distT = emuFromPixels(attrs.distTop, "distTop", authored, emuToPixels);
  }
  if (attrs.distBottom !== undefined) {
    wrap.distB = emuFromPixels(attrs.distBottom, "distBottom", authored, emuToPixels);
  }
  if (attrs.distLeft !== undefined) {
    wrap.distL = emuFromPixels(attrs.distLeft, "distLeft", authored, emuToPixels);
  }
  if (attrs.distRight !== undefined) {
    wrap.distR = emuFromPixels(attrs.distRight, "distRight", authored, emuToPixels);
  }
  const slots = copiedWrapDistanceSlots(attrs.wrapDistanceSlots);
  if (slots !== undefined) {
    wrap.distanceSlots = slots;
  }
  const extents = copiedEffectExtentSlots(attrs.wrapEffectExtentSlots);
  if (extents !== undefined) {
    wrap.effectExtentSlots = extents;
  }
  const polygon = copiedWrapPolygon(attrs.wrapPolygon);
  if (polygon !== undefined) {
    wrap.polygon = polygon;
  }
};

/** A copy, so the model cannot share a slot record with a ProseMirror attr. */
const copiedWrapDistanceSlots = (
  slots: WrapDistanceSlots | undefined,
): WrapDistanceSlots | undefined =>
  slots === undefined
    ? undefined
    : {
        ...(slots.drawing === undefined ? {} : { drawing: { ...slots.drawing } }),
        ...(slots.wrapChild === undefined ? {} : { wrapChild: { ...slots.wrapChild } }),
      };

/** The same copy, one carrier over. */
const copiedEffectExtentSlots = (
  slots: EffectExtentSlots | undefined,
): EffectExtentSlots | undefined =>
  slots === undefined
    ? undefined
    : {
        ...(slots.drawing === undefined ? {} : { drawing: { ...slots.drawing } }),
        ...(slots.wrapChild === undefined ? {} : { wrapChild: { ...slots.wrapChild } }),
      };

function textBoxWrapFromAttrs(attrs: TextBoxAttrs): ImageWrap | undefined {
  const hasWrapData =
    (attrs.wrapType !== undefined && attrs.wrapType !== "inline") ||
    attrs.wrapText !== undefined ||
    attrs.distTop !== undefined ||
    attrs.distBottom !== undefined ||
    attrs.distLeft !== undefined ||
    attrs.distRight !== undefined;
  if (!hasWrapData) {
    return undefined;
  }

  const wrap: ImageWrap = { type: attrs.wrapType ?? "inline" };
  if (attrs.wrapText !== undefined) {
    wrap.wrapText = attrs.wrapText;
  }
  assignWrapDistances(wrap, attrs);
  return wrap;
}

const assignUniqueParagraph = (
  paragraphs: Map<string, Paragraph | null>,
  paraId: string,
  paragraph: Paragraph,
): void => {
  const existing = paragraphs.get(paraId);
  if (!paragraphs.has(paraId) || existing === paragraph) {
    paragraphs.set(paraId, paragraph);
    return;
  }
  paragraphs.set(paraId, null);
};

const uniqueParagraphsById = (
  content: BlockContent[],
  includeTransferIds = false,
): Map<string, Paragraph | null> => {
  const paragraphs = new Map<string, Paragraph | null>();
  visitDocxParagraphs({ documentBody: { content } }, (paragraph) => {
    const { paraId } = paragraph;
    if (paraId) {
      assignUniqueParagraph(paragraphs, paraId, paragraph);
    }
    const transferId = includeTransferIds
      ? getParagraphPropertySourceTransferId(paragraph)
      : undefined;
    if (transferId) {
      assignUniqueParagraph(paragraphs, transferId, paragraph);
    }
  });
  return paragraphs;
};

const restoreParagraphPropertySource = (paragraph: Paragraph, baseParagraph: Paragraph): void => {
  copyParagraphPropertySource(paragraph, baseParagraph);

  const baseFormatting = baseParagraph.formatting;
  if (!baseFormatting) {
    return;
  }
  const { numPr, numPrFromStyle } = baseFormatting;
  if (!numPr || !numPrFromStyle || !isStyleSourcedParagraphNumbering(numPr, numPrFromStyle)) {
    return;
  }
  if (
    canonicalJson(modelParagraphFormattingEmission(paragraph.formatting)) !==
    canonicalJson(modelParagraphFormattingEmission(baseFormatting))
  ) {
    return;
  }
  paragraph.formatting = { ...paragraph.formatting, numPr, numPrFromStyle };
};

const restoreParagraphPropertySources = (
  content: BlockContent[],
  baseContent: BlockContent[],
  linkedTargets: ReadonlySet<Paragraph>,
  linkedSources: ReadonlySet<Paragraph>,
): void => {
  const baseParagraphs = uniqueParagraphsById(baseContent, true);
  for (const [paraId, paragraph] of uniqueParagraphsById(content)) {
    const baseParagraph = baseParagraphs.get(paraId);
    if (
      !paragraph ||
      !baseParagraph ||
      linkedTargets.has(paragraph) ||
      linkedSources.has(baseParagraph) ||
      !getParagraphPropertySource(baseParagraph)
    ) {
      continue;
    }
    restoreParagraphPropertySource(paragraph, baseParagraph);
  }
};

const sourceValidationError = (
  code: ParagraphPropertySourceValidationCode,
  message: string,
): ParagraphPropertySourceValidationError =>
  new ParagraphPropertySourceValidationError({
    code,
    message,
  });

const validateParagraphPropertySourceTokens = (
  pmDoc: PMNode,
  baseDocument: Document,
  contract: string,
): Map<string, Paragraph> => {
  const sourceParagraphs = copyDocumentParagraphPropertySources(baseDocument);
  if (!sourceParagraphs) {
    panic("A paragraph-property source contract lost its source registry");
  }
  const currentTokens = new Set<string>();
  visitDocumentStoryParagraphs(baseDocument.package.document.content, (paragraph) => {
    const token = getParagraphPropertySourceToken(paragraph);
    if (!token) {
      if (paragraphPropertySourceBelongsToDocument(paragraph, baseDocument)) {
        throw sourceValidationError(
          "invalid_token",
          "A source-bound paragraph is missing its paragraph-property token.",
        );
      }
      return;
    }
    if (!paragraphPropertySourceTokenMatchesContract(token, contract)) {
      throw sourceValidationError(
        "invalid_token",
        "The source document contains an invalid paragraph-property token.",
      );
    }
    if (!sourceParagraphs.has(token)) {
      throw sourceValidationError(
        "unknown_token",
        "The source document contains an unknown paragraph-property token.",
      );
    }
    if (currentTokens.has(token)) {
      throw sourceValidationError(
        "duplicate_token",
        "The source document contains a duplicate paragraph-property token.",
      );
    }
    currentTokens.add(token);
  });

  const seen = new Set<string>();
  const validateToken = (token: unknown): void => {
    if (token === null || token === undefined) {
      return;
    }
    if (!isParagraphPropertySourceToken(token)) {
      throw sourceValidationError(
        "invalid_token",
        "A paragraph contains a malformed paragraph-property token.",
      );
    }
    if (!paragraphPropertySourceTokenMatchesContract(token, contract)) {
      throw sourceValidationError(
        "unknown_token",
        "A paragraph-property token belongs to a different source document.",
      );
    }
    if (seen.has(token)) {
      throw sourceValidationError(
        "duplicate_token",
        "A paragraph-property token is attached to more than one paragraph.",
      );
    }
    if (!sourceParagraphs.has(token)) {
      throw sourceValidationError(
        "unknown_token",
        "A paragraph-property token is not present in the source document.",
      );
    }
    seen.add(token);
  };
  const validateTableCellBinding = (binding: TableCellParagraphPropertySourceBinding): void => {
    switch (binding.type) {
      case "authored":
        return;
      case "source":
        validateToken(binding.token);
        return;
      default: {
        const exhaustiveBinding: never = binding;
        return exhaustiveBinding;
      }
    }
  };
  pmDoc.descendants((node) => {
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const continuationCells = expectTableCellAttrs(node)._docxVMergeContinuationCells;
      if (continuationCells !== undefined && continuationCells !== null) {
        visitTableCellParagraphPropertySourceBindings(
          decodeTableCellParagraphSourcePayload(continuationCells),
          (binding) => validateTableCellBinding(binding),
        );
      }
      return true;
    }
    if (node.type.name !== "paragraph") {
      return true;
    }
    validateToken(getProseParagraphPropertySourceToken(node));
    return false;
  });
  return sourceParagraphs;
};

const restoreParagraphPropertySourcesByToken = (
  content: BlockContent[],
  baseParagraphs: ReadonlyMap<string, Paragraph>,
): void => {
  visitDocumentStoryParagraphs(content, (paragraph) => {
    const token = getParagraphPropertySourceToken(paragraph);
    if (!token) {
      return;
    }
    const baseParagraph = baseParagraphs.get(token);
    if (!baseParagraph) {
      panic("Validated paragraph-property token lost its source owner");
    }
    restoreParagraphPropertySource(paragraph, baseParagraph);
  });
};

type LinkedParagraphPropertySources = {
  targets: ReadonlySet<Paragraph>;
  sources: ReadonlySet<Paragraph>;
};

const restoreLinkedParagraphPropertySources = (
  content: BlockContent[],
): LinkedParagraphPropertySources => {
  const targetsBySource = new Map<Paragraph, Paragraph[]>();
  const linkedTargets = new Set<Paragraph>();
  visitDocxParagraphs({ documentBody: { content } }, (paragraph) => {
    const source = getParagraphPropertySourceCandidate(paragraph);
    if (!source) {
      return;
    }
    linkedTargets.add(paragraph);
    const targets = targetsBySource.get(source);
    if (targets) {
      targets.push(paragraph);
    } else {
      targetsBySource.set(source, [paragraph]);
    }
  });
  for (const [source, targets] of targetsBySource) {
    if (targets.length === 1) {
      const target = targets.at(0);
      if (target) {
        copyParagraphPropertyCapture(target, source);
      }
    }
  }
  return { targets: linkedTargets, sources: new Set(targetsBySource.keys()) };
};

/** Convert a ProseMirror document to the document model. */
export function fromProseDoc(pmDoc: PMNode, baseDocument?: Document): Document {
  assertValidProseMirrorDocument(
    pmDoc,
    "Cannot convert invalid ProseMirror document to DOCX model",
  );

  const baseContract = baseDocument
    ? getDocumentParagraphPropertySourceContract(baseDocument)
    : undefined;
  const proseContractAttribute = pmDoc.attrs[PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR];
  const proseContract = getProseDocumentParagraphPropertySourceContract(pmDoc);
  const proseCarriesInvalidContract =
    proseContractAttribute !== null &&
    proseContractAttribute !== undefined &&
    proseContract === null;
  if (proseCarriesInvalidContract || (baseContract ?? null) !== proseContract) {
    throw sourceValidationError(
      "contract_mismatch",
      "The ProseMirror document does not match its paragraph-property source document.",
    );
  }
  const tokenSources =
    baseContract && proseContract && baseDocument
      ? validateParagraphPropertySourceTokens(pmDoc, baseDocument, baseContract)
      : null;

  const blocks = extractBlocks(
    pmDoc,
    "resolve",
    baseDocument?.package.styles ? createStyleEngine(baseDocument.package.styles) : null,
  );
  joinCommentRangesAcrossParagraphs(blocks);
  completeCommentReferences(blocks);
  const linkedSources = restoreLinkedParagraphPropertySources(blocks);
  if (tokenSources) {
    restoreParagraphPropertySourcesByToken(blocks, tokenSources);
  } else if (baseDocument) {
    restoreParagraphPropertySources(
      blocks,
      baseDocument.package.document.content,
      linkedSources.targets,
      linkedSources.sources,
    );
  }

  // Preserve section properties (margins, headers, footers) from base document
  const documentBody: DocumentBody = { content: blocks };
  if (baseDocument?.package.document.finalSectionProperties) {
    documentBody.finalSectionProperties = baseDocument.package.document.finalSectionProperties;
  }
  if (baseDocument?.package.document.sections) {
    documentBody.sections = baseDocument.package.document.sections;
  }
  if (baseDocument?.package.document.comments) {
    documentBody.comments = baseDocument.package.document.comments;
  }

  // If we have a base document, preserve its package structure
  if (baseDocument) {
    const updatedDocument: Document = {
      ...baseDocument,
      package: {
        ...baseDocument.package,
        document: documentBody,
      },
    };
    copyDocumentParagraphPropertySourceContract(updatedDocument, baseDocument);
    return updatedDocument;
  }

  // Create a minimal document structure
  return {
    package: {
      document: documentBody,
    },
  };
}

const isSuggestedMark = (mark: Mark): boolean => mark.attrs["provenance"] === "suggested";

const hasSuggestedInsertion = (marks: readonly Mark[]): boolean =>
  marks.some((mark) => mark.type.name === "insertion" && isSuggestedMark(mark));

/**
 * Compute the mark set an inline node should keep once its suggested tracked
 * changes are stripped. Suggested deletions are dropped so their text survives
 * as plain content; a suggested `runPropertyChange` is dropped and the run's
 * live formatting is reverted to the recorded `previousFormatting`.
 *
 * Returns the input array unchanged when there is nothing to strip so callers
 * can skip rebuilding the node.
 */
function stripSuggestedInlineMarks(
  marks: readonly Mark[],
  {
    baseParagraphFormatting,
    inheritedFormatting,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle,
    styleResolver,
  }: RunFormattingContext,
): readonly Mark[] {
  const hasSuggestedDeletion = marks.some(
    (mark) => mark.type.name === "deletion" && isSuggestedMark(mark),
  );
  const suggestedRunPropertyChange = marks.find(
    (mark) => mark.type.name === "runPropertyChange" && isSuggestedMark(mark),
  );
  if (!hasSuggestedDeletion && !suggestedRunPropertyChange) {
    return marks;
  }

  let next: readonly Mark[] = marks.filter(
    (mark) =>
      !(
        (mark.type.name === "deletion" || mark.type.name === "runPropertyChange") &&
        isSuggestedMark(mark)
      ),
  );

  if (suggestedRunPropertyChange) {
    const previousFormatting = expectRunPropertyChangeMarkAttrs(
      suggestedRunPropertyChange,
    ).changes.at(0)?.previousFormatting;
    const characterStyleMark = marks.find(({ type }) => type.name === "characterStyle");
    const characterStyleAttrs = characterStyleMark
      ? expectCharacterStyleMarkAttrs(characterStyleMark)
      : undefined;
    const preservedCharacterStyleAttrs =
      previousFormatting?.styleId !== undefined &&
      characterStyleAttrs?.styleId === previousFormatting.styleId
        ? characterStyleAttrs
        : undefined;
    const paragraphFormatting = paragraphFormattingForRun({
      marks,
      context: {
        baseParagraphFormatting,
        paragraphFormatting: inheritedFormatting,
        paragraphMarkFormatting,
        paragraphMarkPrecedesStyle,
      },
      ...(previousFormatting !== undefined ? { directFormatting: previousFormatting } : {}),
    });
    const styleFormatting = preservedCharacterStyleAttrs
      ? resolveEffectiveRunStyleFormatting({
          marks,
          paragraphFormatting,
          styleResolver,
        })
      : paragraphFormatting;
    const effectivePreviousFormatting = mergeTextFormatting(styleFormatting, previousFormatting);
    next = next.filter((mark) => !RUN_FORMATTING_MARK_NAMES.has(mark.type.name));
    for (const restored of textFormattingToMarks(effectivePreviousFormatting, schema, {
      overrideFormatting: previousFormatting,
      directFormatting: previousFormatting,
    })) {
      next = restored.addToSet(next);
    }
    if (previousFormatting?.styleId) {
      const characterStyle = suggestedRunPropertyChange.type.schema.marks["characterStyle"];
      if (characterStyle) {
        next = characterStyle
          .create(preservedCharacterStyleAttrs ?? { styleId: previousFormatting.styleId })
          .addToSet(next);
      }
    }
  }

  return next;
}

/**
 * Compute the node attrs a block/structural node keeps once its suggested
 * revision markers are neutralized. Returns `null` when nothing changes.
 *
 * - a suggested `trDel` / `cellMarker` (insertion or deletion) is cleared so
 *   the row/cell serializes as though the proposed change never happened;
 *   merge markers never carry suggestion provenance (structurally excluded);
 * - suggested INSERT markers are handled by the caller, which drops the whole
 *   node instead of clearing an attr.
 *
 * Reads go through the typed attr readers (adapter-boundary convention).
 */
function stripSuggestedNodeAttrs(node: PMNode): Record<string, unknown> | null {
  const name = node.type.name;
  if (name === "paragraph") {
    const attrs = expectParagraphAttrs(node);
    const propertyChanges = attrs._propertyChanges;
    if (!Array.isArray(propertyChanges)) {
      return null;
    }
    if (!propertyChanges.some(({ info }) => info.provenance === "suggested")) {
      return null;
    }
    const removal = removeParagraphPropertyChanges(
      propertyChanges,
      ({ info }) => info.provenance === "suggested",
    );
    if (removal.type === "unchanged") {
      return null;
    }
    const nextAttrs: Record<string, unknown> = {
      ...node.attrs,
      _propertyChanges: removal.remaining.length > 0 ? removal.remaining : null,
    };
    // Only a trailing run of suggestions determines the live pPr. Suggested
    // changes before a retained tracked entry instead rewrite that entry's
    // previous snapshot above, so removing the proposal cannot overwrite the
    // later authored state.
    if (removal.type === "restore-previous") {
      Object.assign(nextAttrs, paragraphRejectAttrPatch(removal.previousFormatting));
      nextAttrs["_originalFormatting"] = paragraphRejectOriginalFormatting(
        removal.previousFormatting,
        nextAttrs["_originalFormatting"],
      );
    }
    return nextAttrs;
  }
  if (name === "tableRow") {
    const rowAttrs = expectTableRowAttrs(node);
    if (rowAttrs.trDel?.provenance === "suggested") {
      return { ...rowAttrs, trDel: null };
    }
    return null;
  }
  if (name === "tableCell" || name === "tableHeader") {
    const cellAttrs = expectTableCellAttrs(node);
    const marker = cellAttrs.cellMarker;
    if (marker && marker.kind !== "merge" && marker.info.provenance === "suggested") {
      return { ...cellAttrs, cellMarker: null };
    }
  }
  return null;
}

/**
 * Whether a block/structural node is a suggested whole-node INSERT and must be
 * dropped from serialization entirely (paragraph, table, row, or column cell).
 */
function isSuggestedInsertedNode(node: PMNode): boolean {
  const name = node.type.name;
  if (name === "paragraph") {
    return expectParagraphAttrs(node)._suggestedInsert != null;
  }
  if (name === "table") {
    return expectTableAttrs(node)._suggestedInsert != null;
  }
  if (name === "tableRow") {
    return expectTableRowAttrs(node).trIns?.provenance === "suggested";
  }
  if (name === "tableCell" || name === "tableHeader") {
    const marker = expectTableCellAttrs(node).cellMarker;
    return marker?.kind === "ins" && marker.info.provenance === "suggested";
  }
  return false;
}

/**
 * Recursively rewrite a ProseMirror node tree, removing every suggested
 * tracked change. Returns `null` when the node itself must be dropped: a
 * suggested-inserted inline node, or a suggested-inserted block/row/cell.
 * Surviving block nodes are rebuilt with their suggested delete/merge markers
 * cleared and their children stripped.
 */
function mapSuggestionStrippedNode(
  node: PMNode,
  formattingContext: RunFormattingContext = {
    baseParagraphFormatting: undefined,
    inheritedFormatting: undefined,
    paragraphMarkFormatting: undefined,
    paragraphMarkPrecedesStyle: false,
    styleResolver: null,
  },
): PMNode | null {
  if (node.isInline) {
    if (hasSuggestedInsertion(node.marks)) {
      return null;
    }
    const marks = stripSuggestedInlineMarks(node.marks, formattingContext);
    return marks === node.marks ? node : node.mark(marks);
  }

  if (isSuggestedInsertedNode(node)) {
    return null;
  }

  const nextAttrs = stripSuggestedNodeAttrs(node);
  const children: PMNode[] = [];
  let changed = nextAttrs !== null;
  const paragraphStyleContext =
    node.type.name === "paragraph"
      ? paragraphRunStyleContext(node, formattingContext.styleResolver)
      : undefined;
  const childFormattingContext = paragraphStyleContext
    ? {
        baseParagraphFormatting: paragraphStyleContext.baseParagraphFormatting,
        inheritedFormatting: paragraphStyleContext.paragraphFormatting,
        paragraphMarkFormatting: paragraphStyleContext.paragraphMarkFormatting,
        paragraphMarkPrecedesStyle: paragraphStyleContext.paragraphMarkPrecedesStyle,
        styleResolver: formattingContext.styleResolver,
      }
    : formattingContext;
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    const mapped = mapSuggestionStrippedNode(child, childFormattingContext);
    if (mapped === null) {
      changed = true;
      return;
    }
    if (mapped !== child) {
      changed = true;
    }
    children.push(mapped);
  });
  if (!changed) {
    return node;
  }
  const content = Fragment.fromArray(children);
  return recreateProseNodeWithParagraphPropertySource(node, {
    ...(nextAttrs === null ? {} : { attrs: nextAttrs }),
    content,
  });
}

/**
 * Strip suggested (AI-proposed) tracked changes from a document so it
 * serializes as though the suggestion never happened. See the call site in
 * {@link extractBlocks} for why this is the sole serialization boundary.
 */
function stripSuggestedProvenance(doc: PMNode, styleResolver: StyleEngine | null): PMNode {
  return (
    mapSuggestionStrippedNode(doc, {
      baseParagraphFormatting: undefined,
      inheritedFormatting: undefined,
      paragraphMarkFormatting: undefined,
      paragraphMarkPrecedesStyle: false,
      styleResolver,
    }) ?? doc
  );
}

/**
 * Extract block content (paragraphs, tables, block SDTs) from a ProseMirror
 * document.
 */
type RefResolutionMode = "resolve" | "inherit";

function materializeNumberedRefValues(doc: PMNode): PMNode {
  const results = resolveNumberedRefFields(doc);
  if (results.size === 0) {
    return doc;
  }
  const visit = (node: PMNode): PMNode => {
    if (node.type.name === "field" || node.type.name === "structuredField") {
      const displayText = results.get(node);
      if (displayText !== undefined) {
        return recreateProseNodeWithParagraphPropertySource(node, {
          attrs: { ...node.attrs, displayText },
        });
      }
      return node;
    }
    if (node.childCount === 0) {
      return node;
    }
    const children: PMNode[] = [];
    let changed = false;
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child) => {
      const mappedChild = visit(child);
      children.push(mappedChild);
      changed ||= mappedChild !== child;
    });
    return changed
      ? recreateProseNodeWithParagraphPropertySource(node, {
          content: Fragment.fromArray(children),
        })
      : node;
  };
  return visit(doc);
}

function extractBlocks(
  inputDoc: PMNode,
  refResolution: RefResolutionMode = "resolve",
  styleResolver: StyleEngine | null = null,
): BlockContent[] {
  // CLASS GUARD: every serialization path (export, copy, header/footer
  // conversion, previews) funnels through `extractBlocks`. Stripping suggested
  // provenance here — with no opt-out — makes it structurally impossible for an
  // AI-proposed edit to reach OOXML output before a human accepts it.
  const strippedDoc = stripSuggestedProvenance(inputDoc, styleResolver);
  const pmDoc =
    refResolution === "resolve" ? materializeNumberedRefValues(strippedDoc) : strippedDoc;
  const blocks: BlockContent[] = [];
  const textBoxAnchorMarkers = new Map<string, Run>();
  const documentCounts = buildDocumentTrackedChangeCounts(pmDoc);
  let pendingPageBreaks = 0;
  let previousStandaloneTextBox: PreviousStandaloneTextBox | null = null;

  const flushPendingPageBreaks = (): void => {
    for (let index = 0; index < pendingPageBreaks; index += 1) {
      blocks.push(createPageBreakParagraph());
    }
    pendingPageBreaks = 0;
  };
  const appendPendingPageBreaksToPreviousParagraph = (): boolean => {
    const previousBlock = blocks.at(-1);
    if (previousBlock?.type !== "paragraph") {
      return false;
    }
    appendPageBreaks(previousBlock, pendingPageBreaks);
    pendingPageBreaks = 0;
    return true;
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  pmDoc.forEach((node) => {
    if (node.type.name === "pageBreak") {
      pendingPageBreaks += 1;
      previousStandaloneTextBox = null;
      return;
    }

    if (node.type.name === "paragraph") {
      const paragraph = convertPMParagraph(
        node,
        documentCounts,
        textBoxAnchorMarkers,
        styleResolver,
      );
      prependPageBreaks(paragraph, pendingPageBreaks);
      pendingPageBreaks = 0;
      blocks.push(paragraph);
      previousStandaloneTextBox = null;
    } else if (node.type.name === "table") {
      if (pendingPageBreaks > 0 && !appendPendingPageBreaksToPreviousParagraph()) {
        flushPendingPageBreaks();
      }
      blocks.push(convertPMTable(node, documentCounts, styleResolver));
      previousStandaloneTextBox = null;
    } else if (node.type.name === "textBox") {
      previousStandaloneTextBox = appendTextBoxBlock(blocks, node, {
        pendingPageBreaks,
        previousStandaloneTextBox,
        textBoxAnchorMarkers,
        styleResolver,
      });
      pendingPageBreaks = 0;
    } else if (node.type.name === "blockSdt") {
      if (pendingPageBreaks > 0 && !appendPendingPageBreaksToPreviousParagraph()) {
        flushPendingPageBreaks();
      }
      blocks.push(convertPMBlockSdt(node, styleResolver));
      previousStandaloneTextBox = null;
    } else if (node.type.name === "preservedBlock") {
      blocks.push({ type: "preservedBlock", xml: expectPreservedBlockAttrs(node).xml });
      previousStandaloneTextBox = null;
    } else if (node.type.name === "blockBookmarkBoundary") {
      blocks.push(blockBookmarkMarker(node));
      previousStandaloneTextBox = null;
    }
  });

  if (pendingPageBreaks > 0 && !appendPendingPageBreaksToPreviousParagraph()) {
    flushPendingPageBreaks();
  }

  removeUnresolvedTextBoxAnchors(blocks, textBoxAnchorMarkers);
  keepOneAttributeRemainderPerRecord(blocks);

  return blocks;
}

/**
 * The attribute remainder follows the record it was authored on, and only it.
 *
 * ProseMirror copies a node's attrs when a command splits it, so pressing
 * Enter in the middle of a paragraph produces two nodes holding the *same*
 * remainder array. Writing it back on both would give the new half a
 * revision-session id nobody assigned to it: `w:rsidR` says which editing
 * session wrote this paragraph, and a copy of it is a claim about history the
 * author did not make. The rule is that the half that comes first in document
 * order keeps the authored identity, and the other half is a new record with
 * no remainder — the same answer a paragraph the editor created from scratch
 * gets, which is none.
 *
 * Reference identity is what tells the two cases apart, and it is exact: two
 * records that each parsed their own attributes hold different arrays however
 * equal their contents, and only a copy made by the editor shares one.
 */
const keepOneAttributeRemainderPerRecord = (blocks: readonly BlockContent[]): void => {
  const seen = new WeakSet<object>();
  /** Whether this array is the authored one rather than a copy's reference. */
  const isAuthored = (value: object | undefined): boolean => {
    if (value === undefined) {
      return true;
    }
    if (seen.has(value)) {
      return false;
    }
    seen.add(value);
    return true;
  };

  const keepFirst = (record: {
    preservedAttributes?: PreservedAttribute[];
    bookmarks?: PositionedBookmarkMarker[];
  }): void => {
    if (!isAuthored(record.preservedAttributes)) {
      delete record.preservedAttributes;
    }
    // A row's or a table's bookmark markers follow the same rule: a copy the
    // editor made shares the array, and writing it back on both would
    // duplicate a `w:bookmarkStart` the author wrote once.
    if (!isAuthored(record.bookmarks)) {
      delete record.bookmarks;
    }
  };

  const walk = (content: readonly BlockContent[]): void => {
    for (const block of content) {
      switch (block.type) {
        case "paragraph":
          keepFirst(block);
          break;
        case "table":
          keepFirst(block);
          for (const row of block.rows) {
            keepFirst(row);
            for (const cell of row.cells) {
              walk(cell.content);
            }
          }
          break;
        case "blockSdt":
          walk(block.content);
          break;
        default:
          break;
      }
    }
  };

  walk(blocks);
};

type AppendTextBoxBlockOptions = {
  pendingPageBreaks: number;
  previousStandaloneTextBox: PreviousStandaloneTextBox | null;
  textBoxAnchorMarkers: Map<string, Run>;
  styleResolver: StyleEngine | null;
};

type PreviousStandaloneTextBox = {
  paragraph: Paragraph;
  groupId: string;
};

/**
 * A bookmark marker standing between two blocks, back in the model.
 *
 * The node's place in the document is the whole of its position, so nothing is
 * read off it but the marker's own attributes.
 */
function blockBookmarkMarker(node: PMNode): BookmarkStart | BookmarkEnd {
  return bookmarkMarkerFromAttrs(expectBookmarkBoundaryAttrs(node));
}

function convertPMBlockSdt(node: PMNode, styleResolver: StyleEngine | null): BlockSdt {
  const attrs = expectBlockSdtAttrs(node);
  const properties: SdtProperties = { sdtType: attrs.sdtType };
  if (attrs.alias) {
    properties.alias = attrs.alias;
  }
  if (attrs.tag) {
    properties.tag = attrs.tag;
  }
  if (typeof attrs.id === "number") {
    properties.id = attrs.id;
  }
  if (attrs.lock) {
    properties.lock = attrs.lock;
  }
  if (attrs.placeholder) {
    properties.placeholder = attrs.placeholder;
  }
  // Preserve the explicit boolean — including `false`. When the widget
  // / editor-ref path fills a placeholder-bearing control,
  // `replaceBlockSdtChildren` sets `showingPlaceholder: false` so the
  // serializer's `reconcileRawSdtPr` knows to remove the source DOCX's
  // `<w:showingPlcHdr/>`. A truthy-only check (the prior shape) would
  // drop that `false` and the saved file would keep marking the
  // newly filled body as placeholder text.
  if (attrs.showingPlaceholder !== undefined) {
    properties.showingPlaceholder = attrs.showingPlaceholder;
  }
  if (attrs.dateFormat) {
    properties.dateFormat = attrs.dateFormat;
  }
  if (attrs.dateValueISO) {
    properties.dateValueISO = attrs.dateValueISO;
  }
  if (attrs.listItems) {
    properties.listItems = decodeSdtListItems(attrs.listItems);
  }
  if (typeof attrs.dropdownLastValue === "string") {
    properties.dropdownLastValue = attrs.dropdownLastValue;
  }
  if (typeof attrs.checked === "boolean") {
    properties.checked = attrs.checked;
  }
  if (attrs.rawPropertiesXml) {
    properties.rawPropertiesXml = attrs.rawPropertiesXml;
  }
  if (attrs.rawEndPropertiesXml) {
    properties.rawEndPropertiesXml = attrs.rawEndPropertiesXml;
  }
  if (attrs.rawSdtChildrenBeforeContent) {
    properties.rawSdtChildrenBeforeContent = attrs.rawSdtChildrenBeforeContent;
  }
  if (attrs.rawSdtChildrenAfterContent) {
    properties.rawSdtChildrenAfterContent = attrs.rawSdtChildrenAfterContent;
  }

  // Recursively materialize children. PM `blockSdt` content is `block+`, so a
  // mini-doc node is a convenient way to reuse extractBlocks.
  const innerDoc = node.type.schema.node("doc", null, node.content);
  const extracted = extractBlocks(innerDoc, "inherit", styleResolver);

  // `toProseDoc` inserts a synthetic filler paragraph into any blockSdt
  // whose source had an empty `<w:sdtContent/>` and stamps the
  // `_originallyEmpty` marker on the PM node. Use that explicit marker
  // (not a shape heuristic) to drop the filler on save — an authored
  // `<w:sdtContent><w:p/></w:sdtContent>` does NOT carry the marker
  // and its empty paragraph survives the round trip intact. If the
  // user typed into the filler we still preserve the paragraph,
  // matching their intent.
  const wasOriginallyEmpty = attrs._originallyEmpty === true;
  const content = wasOriginallyEmpty && isStillSyntheticFiller(extracted) ? [] : extracted;

  return { type: "blockSdt", properties, content };
}

function isStillSyntheticFiller(blocks: BlockContent[]): boolean {
  if (blocks.length !== 1) {
    return false;
  }
  const block = blocks[0];
  if (!block || block.type !== "paragraph") {
    return false;
  }
  if (block.content.length !== 0) {
    return false;
  }
  // Reject anything that signals real editing (formatting, mark changes,
  // section properties) — if the user authored content, preserve it.
  if (block.formatting !== undefined) {
    return false;
  }
  if (block.sectionProperties !== undefined) {
    return false;
  }
  if (block.propertyChanges !== undefined) {
    return false;
  }
  if (block.pPrMark !== undefined) {
    return false;
  }
  return true;
}

function appendTextBoxBlock(
  blocks: BlockContent[],
  node: PMNode,
  options: AppendTextBoxBlockOptions,
): PreviousStandaloneTextBox | null {
  const attrs = expectTextBoxAttrs(node);
  const paragraph = convertPMTextBox(node, options.styleResolver);
  const previousBlock = blocks.at(-1);
  if (attrs._docxPlacement === "inlineWithPrevious" && previousBlock?.type === "paragraph") {
    appendPageBreaks(previousBlock, options.pendingPageBreaks);
    const anchorId = attrs._docxAnchorId;
    const anchorMarker = anchorId ? options.textBoxAnchorMarkers.get(anchorId) : undefined;
    const textBoxRun = findTextBoxShapeRun(paragraph.content);
    if (
      anchorId &&
      anchorMarker &&
      textBoxRun &&
      replaceTextBoxAnchorInBlocks(blocks, anchorMarker, textBoxRun)
    ) {
      options.textBoxAnchorMarkers.delete(anchorId);
      return null;
    }
    if (
      !mergeTextBoxIntoTrailingInlineSdts(previousBlock, paragraph, attrs._docxInlineSdts ?? [])
    ) {
      previousBlock.content.push(...paragraph.content);
    }
    return null;
  }

  if (
    attrs._docxPlacement === "standalone" &&
    attrs._docxGroupId &&
    options.previousStandaloneTextBox?.groupId === attrs._docxGroupId
  ) {
    if (
      !mergeTextBoxIntoTrailingInlineSdts(
        options.previousStandaloneTextBox.paragraph,
        paragraph,
        attrs._docxInlineSdts ?? [],
      )
    ) {
      options.previousStandaloneTextBox.paragraph.content.push(...paragraph.content);
    }
    return options.previousStandaloneTextBox;
  }

  // This node stands in for the `w:p` it was lifted out of, so the host's
  // attribute remainder goes back on the paragraph rebuilt for it.
  if (attrs._docxPlacement === "standalone" && attrs._preservedAttributes?.length) {
    paragraph.preservedAttributes = attrs._preservedAttributes;
  }
  prependPageBreaks(paragraph, options.pendingPageBreaks);
  blocks.push(paragraph);
  return attrs._docxPlacement === "standalone" && attrs._docxGroupId
    ? { paragraph, groupId: attrs._docxGroupId }
    : null;
}

function mergeTextBoxIntoTrailingInlineSdts(
  target: Paragraph,
  incoming: Paragraph,
  inlineSdts: NonNullable<TextBoxAttrs["_docxInlineSdts"]>,
): boolean {
  if (inlineSdts.length === 0 || incoming.content.length !== 1) {
    return false;
  }
  const incomingSdt = incoming.content.at(0);
  const outerAttrs = inlineSdts.at(0);
  if (!outerAttrs || incomingSdt?.type !== "inlineSdt") {
    return false;
  }
  let targetSdt: InlineSdt | undefined;
  for (let index = target.content.length - 1; index >= 0; index -= 1) {
    const content = target.content[index];
    if (content?.type === "inlineSdt" && sdtPropertiesMatchAttrs(content.properties, outerAttrs)) {
      targetSdt = content;
      break;
    }
  }
  if (!targetSdt) {
    return false;
  }
  return mergeInlineSdtNodes(targetSdt, incomingSdt, inlineSdts, 0);
}

function mergeInlineSdtNodes(
  target: InlineSdt,
  incoming: InlineSdt,
  inlineSdts: NonNullable<TextBoxAttrs["_docxInlineSdts"]>,
  index: number,
): boolean {
  const attrs = inlineSdts[index];
  if (
    !attrs ||
    !sdtPropertiesMatchAttrs(target.properties, attrs) ||
    !sdtPropertiesMatchAttrs(incoming.properties, attrs)
  ) {
    return false;
  }
  if (index === inlineSdts.length - 1) {
    target.content.push(...incoming.content);
    return true;
  }
  const incomingNested = incoming.content.at(0);
  const nestedAttrs = inlineSdts[index + 1];
  if (!nestedAttrs || incomingNested?.type !== "inlineSdt") {
    return false;
  }
  let targetNested: InlineSdt | undefined;
  for (let nestedIndex = target.content.length - 1; nestedIndex >= 0; nestedIndex -= 1) {
    const content = target.content[nestedIndex];
    if (content?.type === "inlineSdt" && sdtPropertiesMatchAttrs(content.properties, nestedAttrs)) {
      targetNested = content;
      break;
    }
  }
  if (!targetNested) {
    target.content.push(incomingNested);
    return true;
  }
  return mergeInlineSdtNodes(targetNested, incomingNested, inlineSdts, index + 1);
}

function findTextBoxShapeRun(content: readonly ParagraphContent[]): Run | undefined {
  for (const item of content) {
    if (item.type === "run") {
      if (
        item.content.some(
          (runContent) => runContent.type === "shape" && runContent.shape.shapeType === "textBox",
        )
      ) {
        return item;
      }
      continue;
    }
    if (
      item.type === "inlineSdt" ||
      item.type === "hyperlink" ||
      item.type === "inlineWrapper" ||
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      const nestedContent = item.type === "hyperlink" ? item.children : item.content;
      const run = findTextBoxShapeRun(nestedContent);
      if (run) {
        return run;
      }
    }
  }
  return undefined;
}

function replaceTextBoxAnchorInBlocks(
  blocks: BlockContent[],
  marker: Run,
  textBoxRun: Run,
): boolean {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      if (editTextBoxAnchorInContent(block.content, marker, textBoxRun)) {
        return true;
      }
      continue;
    }
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (replaceTextBoxAnchorInBlocks(cell.content, marker, textBoxRun)) {
            return true;
          }
        }
      }
      continue;
    }
    if (block.type !== "blockSdt") {
      continue;
    }
    if (replaceTextBoxAnchorInBlocks(block.content, marker, textBoxRun)) {
      return true;
    }
  }
  return false;
}

/**
 * The inline list a text-box anchor run can be spliced out of, or nothing when
 * the item holds no such list.
 *
 * A complex field keeps its runs in `fieldCode`/`fieldResult`, two arrays
 * rather than one, and an anchor has never been placed in either; splicing one
 * would have to pick a side. The marker stays where it is and
 * `removeUnresolvedTextBoxAnchors` leaves it, as before.
 */
const textBoxAnchorHost = (item: ParagraphContent): ParagraphContent[] | undefined => {
  switch (item.type) {
    case "hyperlink":
      return item.children;
    case "simpleField":
      return item.content;
    case "inlineSdt":
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
    case "inlineWrapper":
      return item.content;
    case "run":
    case "complexField":
    case "bookmarkStart":
    case "bookmarkEnd":
    case "commentRangeStart":
    case "commentRangeEnd":
    case "commentReference":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
    case "mathEquation":
    case "preservedInline":
      return undefined;
    default: {
      const unsupported: never = item;
      panic(`Unsupported paragraph content: ${JSON.stringify(unsupported)}`);
    }
  }
};

function editTextBoxAnchorInContent(
  content: ParagraphContent[],
  marker: Run,
  replacement?: Run,
): boolean {
  for (const [index, item] of content.entries()) {
    if (item === marker) {
      content.splice(index, 1, ...(replacement ? [replacement] : []));
      return true;
    }
    const nestedContent = textBoxAnchorHost(item);
    if (!nestedContent) {
      continue;
    }
    const hadContent = nestedContent.length > 0;
    if (editTextBoxAnchorInContent(nestedContent, marker, replacement)) {
      if (!replacement && hadContent && nestedContent.length === 0) {
        content.splice(index, 1);
      }
      return true;
    }
  }
  return false;
}

function removeUnresolvedTextBoxAnchors(
  blocks: BlockContent[],
  markers: ReadonlyMap<string, Run>,
): void {
  for (const marker of markers.values()) {
    removeTextBoxAnchorFromBlocks(blocks, marker);
  }
}

function removeTextBoxAnchorFromBlocks(blocks: BlockContent[], marker: Run): boolean {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      if (editTextBoxAnchorInContent(block.content, marker)) {
        return true;
      }
      continue;
    }
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (removeTextBoxAnchorFromBlocks(cell.content, marker)) {
            return true;
          }
        }
      }
      continue;
    }
    if (block.type !== "blockSdt") {
      continue;
    }
    if (removeTextBoxAnchorFromBlocks(block.content, marker)) {
      return true;
    }
  }
  return false;
}

/**
 * Inverse of toProseDoc's listRendering → list* attrs flattening. Markdown
 * export (`toMarkdown`) and re-layout of the rebuilt Document key off
 * `paragraph.listRendering`; without this, every edited document loses its
 * list markers on the way out of the editor.
 */
function listRenderingFromAttrs(attrs: ParagraphAttrs): Paragraph["listRendering"] {
  const numId = paragraphNumberingReferenceId(attrs.numPr);
  if (numId === undefined) {
    return undefined;
  }
  const hasRenderingInfo =
    attrs.listMarker != null || attrs.listIsBullet || attrs.listNumFmt != null;
  if (!hasRenderingInfo) {
    return undefined;
  }
  return listRenderingFieldsFromAttrs({ attrs, numId });
}

/**
 * Create a paragraph containing only a page break run (for DOCX serialization)
 */
function createPageBreakParagraph(): Paragraph {
  const breakContent: BreakContent = { type: "break", breakType: "page" };
  const run: Run = { type: "run", content: [breakContent] };
  return {
    type: "paragraph",
    content: [run],
  };
}

function createPageBreakRun(): Run {
  return {
    type: "run",
    content: [{ type: "break", breakType: "page" }],
  };
}

function prependPageBreaks(paragraph: Paragraph, count: number): void {
  for (let index = 0; index < count; index += 1) {
    paragraph.content.unshift(createPageBreakRun());
  }
}

function appendPageBreaks(paragraph: Paragraph, count: number): void {
  for (let index = 0; index < count; index += 1) {
    paragraph.content.push(createPageBreakRun());
  }
}

/**
 * Convert a ProseMirror paragraph node to our Paragraph type
 */
function convertPMParagraph(
  node: PMNode,
  documentCounts?: TrackedChangeCounts,
  textBoxAnchorMarkers?: Map<string, Run>,
  styleResolver: StyleEngine | null = null,
): Paragraph {
  const attrs = expectParagraphAttrs(node);
  const paragraphStyleContext = paragraphRunStyleContext(node, styleResolver);
  let content = extractParagraphContent(
    node,
    documentCounts,
    attrs._emptyHyperlinks ?? undefined,
    textBoxAnchorMarkers,
    attrs.renderedPageBreakBefore === true,
    {
      baseParagraphFormatting: paragraphStyleContext.baseParagraphFormatting,
      inheritedFormatting: paragraphStyleContext.paragraphFormatting,
      paragraphMarkFormatting: paragraphStyleContext.paragraphMarkFormatting,
      paragraphMarkPrecedesStyle: paragraphStyleContext.paragraphMarkPrecedesStyle,
      styleResolver,
    },
  );

  // Emit BookmarkStart/End from bookmarks attr (for TOC anchors, cross-references)
  const bookmarks = attrs.bookmarks as { id: number; name: string }[] | undefined;
  if (bookmarks && bookmarks.length > 0) {
    const starts: ParagraphContent[] = bookmarks.map((b) => ({
      type: "bookmarkStart" as const,
      id: b.id,
      name: b.name,
    }));
    const ends: ParagraphContent[] = bookmarks.map((b) => ({
      type: "bookmarkEnd" as const,
      id: b.id,
    }));
    content = [...starts, ...content, ...ends];
  }

  if (attrs._moveRanges && attrs._moveRanges.length > 0) {
    content = withMoveRanges(content, attrs._moveRanges);
  }

  const paragraph: Paragraph = {
    type: "paragraph",
    content,
  };
  if (attrs.paraId) {
    paragraph.paraId = attrs.paraId;
  }
  if (attrs.textId) {
    paragraph.textId = attrs.textId;
  }
  if (attrs.reviewCarrier) {
    paragraph.reviewCarrier = attrs.reviewCarrier;
  }
  const pFormatting = paragraphAttrsToFormatting(attrs);
  if (pFormatting) {
    paragraph.formatting = pFormatting;
  }
  const listRendering = listRenderingFromAttrs(attrs);
  if (listRendering) {
    paragraph.listRendering = listRendering;
  }
  if (attrs.renderedPageBreakBefore) {
    paragraph.renderedPageBreakBefore = true;
  }

  // Restore full section properties (round-trip) or fallback to break type only
  if (attrs._sectionProperties) {
    paragraph.sectionProperties = attrs._sectionProperties as SectionProperties;
  } else if (attrs.sectionBreakType) {
    paragraph.sectionProperties = {
      sectionStart: attrs.sectionBreakType as SectionStart,
    };
  }

  // Restore `w:pPrChange` entries that PM carried opaquely. The editor
  // doesn't surface them in UI, but they must survive an edit so the
  // saved DOCX still contains the property-change history Word relies
  // on. Shallow-clone the array so the rebuilt Folio document doesn't
  // share a mutable reference with PM's attrs.
  if (attrs._propertyChanges && attrs._propertyChanges.length > 0) {
    paragraph.propertyChanges = attrs._propertyChanges.map(propertyChangeFromAttrs);
  }

  if (attrs.pPrMark) {
    paragraph.pPrMark = attrs.pPrMark;
  }

  // The attribute remainder, by reference: `keepOneAttributeRemainderPerRecord`
  // below is what decides whether this paragraph is the one that authored it.
  if (attrs._preservedAttributes && attrs._preservedAttributes.length > 0) {
    paragraph.preservedAttributes = attrs._preservedAttributes;
  }

  linkParagraphPropertySourceCandidate(paragraph, node);
  return paragraph;
}

const propertyChangeFromAttrs = (change: ParagraphPropertyChangeAttrs): ParagraphPropertyChange => {
  const { previousFormatting, currentFormatting, info, ...changeInfo } = change;
  const serializedInfo = { ...info };
  Reflect.deleteProperty(serializedInfo, "provenance");
  Reflect.deleteProperty(serializedInfo, "suggestionId");
  const normalizedChangeInfo = { ...changeInfo, info: serializedInfo };
  if (previousFormatting === undefined) {
    return currentFormatting === undefined
      ? normalizedChangeInfo
      : { ...normalizedChangeInfo, currentFormatting };
  }
  // `null` is the attr's tombstone — the paragraph carried no numbering before
  // the change — and the model spells that as the field being absent.
  const { numPr, ...previousWithoutNumPr } = previousFormatting;
  const normalizedPrevious =
    numPr == null ? previousWithoutNumPr : { ...previousWithoutNumPr, numPr };
  return {
    ...normalizedChangeInfo,
    previousFormatting: normalizedPrevious,
    ...(currentFormatting !== undefined && { currentFormatting }),
  };
};

/**
 * What the paragraph states itself for one `w:pPr` field, as opposed to what
 * the style cascade resolved into the same attr.
 *
 * `undefined` means the paragraph states nothing and the field must stay
 * absent: a materialised value is a direct override that outranks the style it
 * was read from, and a later edit to that style no longer reaches the
 * paragraph. A value the source stated directly is always authored, even when
 * it happens to equal the style's.
 */
const authoredParagraphValue = <Field extends keyof ParagraphFormatting>(
  field: Field,
  value: ParagraphFormatting[Field] | undefined,
  orig: ParagraphFormatting | undefined,
  resolved: ParagraphFormatting | undefined,
): ParagraphFormatting[Field] | undefined => {
  if (value === undefined) {
    return undefined;
  }
  if (orig?.[field] !== undefined) {
    return value;
  }
  return sameResolvedValue(value, resolved?.[field]) ? undefined : value;
};

/**
 * OOXML boolean paragraph toggles are tri-state: `true` (on), `false` (explicit
 * off, serialized as `w:val="0"`), and `null`/`undefined` (inherit). A
 * truthiness check (`if (attrs.key)`) silently collapses explicit `false` into
 * "inherit", dropping the user's decision on save. Branch on `== null` so every
 * toggle routed through here preserves `false`. (Direction is handled
 * separately via the `direction` discriminated union, not this helper.)
 */
type BooleanToggleKey = Extract<
  StyleResolvedParagraphField,
  | "pageBreakBefore"
  | "widowControl"
  | "keepNext"
  | "keepLines"
  | "runInWithNext"
  | "snapToGrid"
  | "kinsoku"
  | "overflowPunctuation"
  | "suppressAutoHyphens"
>;

function assignBooleanToggle(
  result: ParagraphFormatting,
  attrs: ParagraphAttrs,
  orig: ParagraphFormatting,
  key: BooleanToggleKey,
): void {
  // Tri-state: `true`/`false` are explicit decisions to keep; `null`/`undefined`
  // is "undecided" and clears the toggle. `exactOptionalPropertyTypes` forbids
  // assigning `undefined`, and `no-dynamic-delete` forbids `delete result[key]`,
  // so the undecided branch clears via `Reflect.deleteProperty`. This preserves
  // an explicit `false` that a truthiness check would have dropped.
  const value = authoredParagraphValue(
    key,
    attrs[key] ?? undefined,
    orig,
    attrs._resolvedFormatting,
  );
  if (value === (orig[key] ?? undefined)) {
    return;
  }
  if (value === undefined) {
    Reflect.deleteProperty(result, key);
  } else {
    result[key] = value;
  }
}

function paragraphAttrsToFormatting(attrs: ParagraphAttrs): ParagraphFormatting | undefined {
  const directAlignment = directParagraphAlignment(attrs);
  const directSpacing = directParagraphSpacing(attrs);
  // If we have the original inline formatting from the DOCX, use it as a base
  // for lossless round-trip. This preserves properties like contextualSpacing,
  // widowControl, beforeAutospacing, runProperties, etc. that aren't tracked
  // as individual PM attrs. It also avoids "inlining" style-inherited values
  // (spacing, indentation, numPr) which would override style definitions
  // and break rendering in Word/Pages/Google Docs.
  //
  // We then apply overrides for any properties the user may have changed
  // via editor commands (alignment, list toggle, etc.).
  const spaceBefore = Reflect.get(attrs, "spaceBefore");
  const spaceAfter = Reflect.get(attrs, "spaceAfter");
  const beforeHasAutospacingBase = hasAutospacingBaseSide(attrs._autospacingBase, "before");
  const afterHasAutospacingBase = hasAutospacingBaseSide(attrs._autospacingBase, "after");
  const beforeOriginalAutospacing = attrs._originalFormatting?.beforeAutospacing === true;
  const afterOriginalAutospacing = attrs._originalFormatting?.afterAutospacing === true;
  const beforeAutospacingEdited = beforeHasAutospacingBase
    ? !autospacingMatchesBase(attrs._autospacingBase, "before", spaceBefore)
    : beforeOriginalAutospacing && attrs._autospacingBase == null;
  const afterAutospacingEdited = afterHasAutospacingBase
    ? !autospacingMatchesBase(attrs._autospacingBase, "after", spaceAfter)
    : afterOriginalAutospacing && attrs._autospacingBase == null;
  const beforeIsInherited =
    attrs.spacingFromDocDefaults?.before === true ||
    attrs.spacingFromImplicitDefaultStyle?.before === true;
  const afterIsInherited =
    attrs.spacingFromDocDefaults?.after === true ||
    attrs.spacingFromImplicitDefaultStyle?.after === true;
  const shouldSerializeSpaceBefore =
    typeof spaceBefore === "number" &&
    (attrs.spacingExplicit?.before === true ||
      beforeAutospacingEdited ||
      (!beforeIsInherited && !beforeHasAutospacingBase));
  const shouldSerializeSpaceAfter =
    typeof spaceAfter === "number" &&
    (attrs.spacingExplicit?.after === true ||
      afterAutospacingEdited ||
      (!afterIsInherited && !afterHasAutospacingBase));
  const hasDirectLineSpacing = directSpacing?.lineSpacing !== undefined;
  const hasDirectLineSpacingRule = directSpacing?.lineSpacingRule !== undefined;

  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };

    if (beforeAutospacingEdited) {
      result.beforeAutospacing = false;
      if (typeof spaceBefore === "number") {
        result.spaceBefore = spaceBefore;
      } else {
        delete result.spaceBefore;
      }
    }
    if (afterAutospacingEdited) {
      result.afterAutospacing = false;
      if (typeof spaceAfter === "number") {
        result.spaceAfter = spaceAfter;
      } else {
        delete result.spaceAfter;
      }
    }

    // A spacing command is a direct override even when the imported value was
    // inherited from a style. Keep the explicit zero instead of dropping the
    // side and letting the style value reappear on the next load.
    if (orig.spaceBefore !== undefined || attrs.spacingExplicit?.before) {
      if (typeof spaceBefore === "number") {
        result.spaceBefore = spaceBefore;
      } else {
        Reflect.deleteProperty(result, "spaceBefore");
      }
    }
    if (orig.spaceAfter !== undefined || attrs.spacingExplicit?.after) {
      if (typeof spaceAfter === "number") {
        result.spaceAfter = spaceAfter;
      } else {
        Reflect.deleteProperty(result, "spaceAfter");
      }
    }

    const originalHasDirectLineSpacing = orig.lineSpacing !== undefined;
    if (hasDirectLineSpacing || originalHasDirectLineSpacing) {
      if (typeof attrs.lineSpacing === "number") {
        result.lineSpacing = attrs.lineSpacing;
      } else {
        Reflect.deleteProperty(result, "lineSpacing");
      }
    }
    const originalHasDirectLineSpacingRule = orig.lineSpacingRule !== undefined;
    if (hasDirectLineSpacingRule || originalHasDirectLineSpacingRule) {
      if (attrs.lineSpacingRule) {
        result.lineSpacingRule = attrs.lineSpacingRule;
      } else {
        Reflect.deleteProperty(result, "lineSpacingRule");
      }
    }

    // The effective value stays available for layout, but only the separately
    // tracked direct value belongs in `w:pPr/w:jc`.
    if (directAlignment === undefined) {
      Reflect.deleteProperty(result, "alignment");
    } else {
      result.alignment = directAlignment;
    }
    const statedNumbering = attrs.numPr;
    if (isStyleSourcedParagraphNumbering(statedNumbering, attrs.numPrFromStyle)) {
      // The numbering still comes verbatim from the paragraph style — don't
      // materialize it as direct formatting (see ParagraphAttrs.numPrFromStyle).
      delete result.numPr;
      delete result.numPrFromStyle;
    } else if (!sameStatedParagraphNumbering(statedNumbering, orig.numPr)) {
      if (statedNumbering !== undefined) {
        result.numPr = statedNumbering;
      } else {
        delete result.numPr;
      }
      delete result.numPrFromStyle;
    }
    if (attrs.styleId !== (orig.styleId ?? undefined)) {
      if (attrs.styleId) {
        result.styleId = attrs.styleId;
      } else {
        delete result.styleId;
      }
    }
    assignBooleanToggle(result, attrs, orig, "pageBreakBefore");
    assignBooleanToggle(result, attrs, orig, "widowControl");
    assignBooleanToggle(result, attrs, orig, "keepNext");
    assignBooleanToggle(result, attrs, orig, "keepLines");
    assignBooleanToggle(result, attrs, orig, "runInWithNext");
    assignBooleanToggle(result, attrs, orig, "snapToGrid");
    assignBooleanToggle(result, attrs, orig, "kinsoku");
    assignBooleanToggle(result, attrs, orig, "overflowPunctuation");
    assignBooleanToggle(result, attrs, orig, "suppressAutoHyphens");
    if (attrs.spacingExplicit !== orig.spacingExplicit) {
      if (attrs.spacingExplicit) {
        result.spacingExplicit = attrs.spacingExplicit;
      } else {
        delete result.spacingExplicit;
      }
    }
    // Resolve the paragraph direction to the OOXML `w:bidi` tri-state. An
    // explicit `false` (forced LTR) is preserved so it serializes as
    // `<w:bidi w:val="0"/>` and survives save/reload; `undefined` (undecided)
    // clears it. A direction that only echoes the style's `w:bidi` is not the
    // paragraph's own decision and must not become direct formatting.
    const bidi = authoredParagraphValue(
      "bidi",
      directionToBidi(attrs.direction),
      orig,
      attrs._resolvedFormatting,
    );
    if (bidi !== (orig.bidi ?? undefined)) {
      if (bidi === undefined) {
        delete result.bidi;
      } else {
        result.bidi = bidi;
      }
    }

    return result;
  }

  // Fallback: reconstruct formatting from individual attrs (e.g. for
  // newly created paragraphs that don't have _originalFormatting). Every
  // `style-resolved-attr` field is read through `authored`, so a value the
  // paragraph only inherits stays out of the direct `w:pPr`.
  const resolved = attrs._resolvedFormatting;
  const authored = <Field extends StyleResolvedParagraphField>(
    field: Field,
    value: ParagraphFormatting[Field] | null | undefined,
  ): ParagraphFormatting[Field] | undefined =>
    authoredParagraphValue(field, value ?? undefined, undefined, resolved);

  const outlineLevel = authored("outlineLevel", Reflect.get(attrs, "outlineLevel"));
  const bidi = authored("bidi", directionToBidi(attrs.direction));
  const snapToGrid = authored("snapToGrid", attrs.snapToGrid);
  const indentLeft = authored("indentLeft", attrs.indentLeft);
  const indentRight = authored("indentRight", attrs.indentRight);
  const indentFirstLine = authored("indentFirstLine", attrs.indentFirstLine);
  const borders = authored("borders", attrs.borders);
  const shading = authored("shading", attrs.shading);
  const tabs = authored("tabs", attrs.tabs);
  const contextualSpacing = authored("contextualSpacing", attrs.contextualSpacing);
  const pageBreakBefore = authored("pageBreakBefore", attrs.pageBreakBefore);
  const widowControl = authored("widowControl", attrs.widowControl);
  const keepNext = authored("keepNext", attrs.keepNext);
  const keepLines = authored("keepLines", attrs.keepLines);
  const runInWithNext = authored("runInWithNext", attrs.runInWithNext);
  const kinsoku = authored("kinsoku", attrs.kinsoku);
  const overflowPunctuation = authored("overflowPunctuation", attrs.overflowPunctuation);
  const suppressAutoHyphens = authored("suppressAutoHyphens", attrs.suppressAutoHyphens);

  const hasDirectAlignment = directAlignment !== undefined;
  const hasFormatting =
    hasDirectAlignment ||
    shouldSerializeSpaceBefore ||
    shouldSerializeSpaceAfter ||
    beforeAutospacingEdited ||
    afterAutospacingEdited ||
    hasDirectLineSpacing ||
    hasDirectLineSpacingRule ||
    snapToGrid != null ||
    indentLeft ||
    indentRight ||
    indentFirstLine ||
    attrs.numPr ||
    attrs.styleId ||
    borders ||
    shading ||
    tabs ||
    outlineLevel !== undefined ||
    contextualSpacing ||
    attrs.spacingExplicit ||
    // Tri-state toggles: an explicit `false` is meaningful formatting and must
    // keep the paragraph from short-circuiting to "no formatting".
    bidi != null ||
    pageBreakBefore != null ||
    widowControl != null ||
    keepNext != null ||
    keepLines != null ||
    runInWithNext != null ||
    kinsoku != null ||
    overflowPunctuation != null ||
    suppressAutoHyphens != null;

  if (!hasFormatting) {
    return undefined;
  }

  const f: ParagraphFormatting = {};
  if (directAlignment !== undefined) {
    f.alignment = directAlignment;
  }
  if (shouldSerializeSpaceBefore) {
    f.spaceBefore = spaceBefore;
  }
  if (beforeAutospacingEdited) {
    f.beforeAutospacing = false;
  }
  if (shouldSerializeSpaceAfter) {
    f.spaceAfter = spaceAfter;
  }
  if (afterAutospacingEdited) {
    f.afterAutospacing = false;
  }
  if (hasDirectLineSpacing && typeof attrs.lineSpacing === "number") {
    f.lineSpacing = attrs.lineSpacing;
  }
  if (hasDirectLineSpacingRule && attrs.lineSpacingRule) {
    f.lineSpacingRule = attrs.lineSpacingRule;
  }
  if (snapToGrid != null) {
    f.snapToGrid = snapToGrid;
  }
  if (attrs.spacingExplicit) {
    f.spacingExplicit = attrs.spacingExplicit;
  }
  if (indentLeft) {
    f.indentLeft = indentLeft;
  }
  if (indentRight) {
    f.indentRight = indentRight;
  }
  if (indentFirstLine) {
    f.indentFirstLine = indentFirstLine;
  }
  if (attrs.hangingIndent && indentFirstLine) {
    f.hangingIndent = attrs.hangingIndent;
  }
  if (
    attrs.numPr !== undefined &&
    !isStyleSourcedParagraphNumbering(attrs.numPr, attrs.numPrFromStyle)
  ) {
    f.numPr = attrs.numPr;
  }
  if (attrs.styleId) {
    f.styleId = attrs.styleId;
  }
  if (borders) {
    f.borders = borders;
  }
  if (shading) {
    f.shading = shading;
  }
  if (tabs) {
    f.tabs = tabs;
  }
  if (outlineLevel !== undefined) {
    f.outlineLevel = outlineLevel;
  }
  if (contextualSpacing) {
    f.contextualSpacing = contextualSpacing;
  }
  // Preserve explicit tri-state decisions, including `false` (which serializes
  // as `w:val="0"`); only undecided `null` is omitted.
  if (bidi != null) {
    f.bidi = bidi;
  }
  if (pageBreakBefore != null) {
    f.pageBreakBefore = pageBreakBefore;
  }
  if (widowControl != null) {
    f.widowControl = widowControl;
  }
  if (keepNext != null) {
    f.keepNext = keepNext;
  }
  if (keepLines != null) {
    f.keepLines = keepLines;
  }
  if (runInWithNext != null) {
    f.runInWithNext = runInWithNext;
  }
  if (kinsoku != null) {
    f.kinsoku = kinsoku;
  }
  if (overflowPunctuation != null) {
    f.overflowPunctuation = overflowPunctuation;
  }
  if (suppressAutoHyphens != null) {
    f.suppressAutoHyphens = suppressAutoHyphens;
  }
  return f;
}

/**
 * Extract paragraph content (runs, hyperlinks) from ProseMirror paragraph
 *
 * Coalesces consecutive text with the same marks into single Runs
 * for efficient DOCX representation.
 */
type TrackedRunWrapper = Extract<
  ParagraphContent,
  { type: "insertion" | "deletion" | "moveFrom" | "moveTo" }
>;

function createTrackedRunWrapper(
  type: TrackedRunWrapper["type"],
  info: TrackedChangeInfo,
  child?: TrackedRunWrapper["content"][number],
): TrackedRunWrapper {
  const content = child ? [child] : [];
  if (type === "insertion") {
    return { type, info, content };
  }
  if (type === "deletion") {
    return { type, info, content };
  }
  if (type === "moveFrom") {
    return { type, info, content };
  }
  return { type, info, content };
}

type RunFormattingContext = {
  baseParagraphFormatting: TextFormatting | undefined;
  inheritedFormatting: TextFormatting | undefined;
  paragraphMarkFormatting: TextFormatting | undefined;
  paragraphMarkPrecedesStyle: boolean;
  styleResolver: RunStyleResolver | null;
};

/**
 * Comment ids in the order their markers are written at a shared boundary.
 *
 * Two ranges that begin or end at the same position have no order in the
 * editor, where each is a mark, so every emission site orders them by id: a
 * boundary written in the order the marks happened to open is a different
 * document each time the same one is saved.
 */
const byCommentId = (ids: Iterable<number>): number[] => [...ids].toSorted((a, b) => a - b);

/** The wrappers `node` sits inside, outermost first; empty when it sits in none. */
const inlineWrapperStackOf = (node: PMNode): readonly InlineWrapperLayer[] => {
  const mark = node.marks.find((candidate) => candidate.type.name === INLINE_WRAPPER_MARK_NAME);
  return mark ? expectInlineWrapperMarkAttrs(mark).stack : [];
};

/** Where one run of inline items carrying the same wrapper stack begins. */
type InlineWrapperGroup = {
  /** An index into the emitted content, not into the editor's inline sequence. */
  start: number;
  stack: readonly InlineWrapperLayer[];
};

const isRevisionWrapper = (item: ParagraphContent): item is TrackedRunWrapper =>
  item.type === "insertion" ||
  item.type === "deletion" ||
  item.type === "moveFrom" ||
  item.type === "moveTo";

/** `stack` closed around `content`, outermost layer first. */
const nestInlineWrappers = (
  stack: readonly InlineWrapperLayer[],
  content: ParagraphContent[],
): InlineWrapper => {
  const innermost = stack.at(-1);
  if (innermost === undefined) {
    panic("An inline wrapper group carries no layer");
  }
  let wrapper = inlineWrapperMember(innermost, content);
  for (const layer of stack.slice(0, -1).toReversed()) {
    wrapper = inlineWrapperMember(layer, [wrapper]);
  }
  return wrapper;
};

/**
 * The emitted sequence with each wrapper group nested back into the wrappers
 * `toProseDoc` lifted off it.
 *
 * The revision stays outermost. folio already writes a revision outside the
 * hyperlink it spans, the parse leg is revision-owned, and accepting or
 * rejecting one is a range operation over the revision's own content; a
 * wrapper placed outside two revisions would also have to mint one revision id
 * per wrapper. So a revision in the group keeps its place and takes the nest
 * inside it, and the items around it share a nest of their own.
 *
 * A group with nothing in it writes no wrapper: an authored wrapper whose
 * content was deleted or rejected is gone, and an empty one says nothing.
 */
const nestInlineWrapperGroups = (
  items: readonly ParagraphContent[],
  groups: readonly InlineWrapperGroup[],
): ParagraphContent[] => {
  const nested: ParagraphContent[] = [];
  for (const [index, group] of groups.entries()) {
    const slice = items.slice(group.start, groups[index + 1]?.start ?? items.length);
    if (group.stack.length === 0) {
      nested.push(...slice);
      continue;
    }
    let pending: ParagraphContent[] = [];
    const closeNest = (): void => {
      if (pending.length === 0) {
        return;
      }
      nested.push(nestInlineWrappers(group.stack, pending));
      pending = [];
    };
    for (const item of slice) {
      if (!isRevisionWrapper(item)) {
        pending.push(item);
        continue;
      }
      closeNest();
      if (item.content.length > 0) {
        item.content = [nestInlineWrappers(group.stack, item.content)];
      }
      nested.push(item);
    }
    closeNest();
  }
  return nested;
};

function extractParagraphContent(
  paragraph: PMNode,
  // Parameter retained for signature compatibility with the call sites
  // threaded through tables/cells. The body no longer needs the counts
  // — `moveFrom`/`moveTo` round-trip is now driven by the explicit
  // `moveKind` mark attribute set by `toProseDoc`.
  _documentCounts?: TrackedChangeCounts,
  emptyHyperlinks?: NonNullable<ParagraphAttrs["_emptyHyperlinks"]>,
  textBoxAnchorMarkers?: Map<string, Run>,
  skipLeadingRenderedPageBreak = false,
  inheritedFormattingOverride?: RunFormattingContext,
): ParagraphContent[] {
  const content: ParagraphContent[] = [];
  const paragraphStyleContext =
    paragraph.type.name === "paragraph" ? paragraphRunStyleContext(paragraph) : undefined;
  const formattingContext =
    inheritedFormattingOverride ??
    ({
      baseParagraphFormatting: paragraphStyleContext?.baseParagraphFormatting,
      inheritedFormatting: paragraphStyleContext?.paragraphFormatting,
      paragraphMarkFormatting: paragraphStyleContext?.paragraphMarkFormatting,
      paragraphMarkPrecedesStyle: paragraphStyleContext?.paragraphMarkPrecedesStyle ?? false,
      styleResolver: null,
    } satisfies RunFormattingContext);
  const sortedEmptyHyperlinks = (emptyHyperlinks ?? [])
    .map((attrs, order) => ({ attrs, order }))
    .toSorted((left, right) => left.attrs.offset - right.attrs.offset || left.order - right.order);
  let nextEmptyHyperlink = 0;
  let leadingRenderedPageBreakPending = skipLeadingRenderedPageBreak;

  // The wrapper groups the emitted sequence is cut into, in document order.
  // The walk records boundaries and `nestInlineWrapperGroups` builds the
  // wrappers at the end: a revision or a hyperlink is only finished once the
  // walk has left it, and the nesting has to know which of them the group
  // holds whole.
  const wrapperGroups: InlineWrapperGroup[] = [{ start: 0, stack: [] }];
  let wrapperGroupKey = "";

  // Track current run being built
  let currentRun: Run | null = null;
  let currentMarksKey: string | null = null;
  let currentHyperlink: Hyperlink | null = null;
  let currentHyperlinkKey: string | null = null;
  let currentTrackedChange:
    | {
        type: "direct";
        key: string;
        wrapper: TrackedRunWrapper;
      }
    | {
        type: "hyperlink";
        key: string;
        wrapper: TrackedRunWrapper;
        hyperlink: Hyperlink;
        hyperlinkKey: string;
      }
    | undefined;
  const sourceRunOwners = new WeakMap<Run, number>();
  const openedComments = new Set<number>();

  // A single comment id must round-trip to a single contiguous comment range.
  // Pre-compute the last offset at which each comment appears so the range
  // spans any interrupting node that drops the mark (a tracked-change run, or
  // an atom whose node spec can't carry the comment mark) instead of being
  // split into several ranges for one id — invalid OOXML that Word rejects as
  // unreadable content (eigenpal/docx-editor#927).
  const commentLastOffset = new Map<number, number>();
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  paragraph.forEach((node, offset) => {
    for (const commentId of getCommentMarkIds(node.marks)) {
      commentLastOffset.set(commentId, offset);
    }
  });

  const flushCurrentInline = () => {
    if (currentRun) {
      content.push(currentRun);
      currentRun = null;
      currentMarksKey = null;
    }
    if (currentHyperlink) {
      content.push(currentHyperlink);
      currentHyperlink = null;
      currentHyperlinkKey = null;
    }
  };

  const appendDirectRun = (node: PMNode, run: Run, coalescePlainText = false): void => {
    const marksKey = getMarksKey(node.marks);
    const ownerId = pageBreakRunOwnerId(node);
    const currentOwnerId = currentRun ? sourceRunOwners.get(currentRun) : undefined;
    const currentRunIsPlainText = currentRun?.content.every(
      (runContent) => runContent.type === "text",
    );
    const joinsOwnedSourceRun =
      ownerId !== undefined &&
      ownerId === currentOwnerId &&
      currentRun !== null &&
      runsShareProperties(currentRun, run);
    const joinsOrdinaryText =
      coalescePlainText &&
      ownerId === undefined &&
      currentOwnerId === undefined &&
      currentRunIsPlainText === true &&
      currentMarksKey === marksKey;

    if (currentRun && (joinsOwnedSourceRun || joinsOrdinaryText)) {
      if (joinsOrdinaryText) {
        for (const runContent of run.content) {
          if (runContent.type === "text") {
            appendTextToRun(currentRun, runContent.text);
          }
        }
      } else {
        currentRun.content.push(...run.content);
      }
      return;
    }
    if (currentRun) {
      content.push(currentRun);
    }
    currentRun = run;
    currentMarksKey = marksKey;
    rememberSourceRunOwner(run, node, sourceRunOwners);
  };

  const syncCommentRanges = (node: PMNode, offset: number) => {
    const nodeCommentIds = getCommentMarkIds(node.marks);

    // Close an open range only once this node is past the comment's last
    // occurrence; a comment that reappears later stays open so the range spans
    // the gap (eigenpal/docx-editor#927).
    const toClose: number[] = [];
    for (const commentId of openedComments) {
      const lastOffset = commentLastOffset.get(commentId) ?? offset;
      if (!nodeCommentIds.has(commentId) && lastOffset <= offset) {
        toClose.push(commentId);
      }
    }

    const toOpen: number[] = [];
    for (const commentId of nodeCommentIds) {
      if (!openedComments.has(commentId)) {
        toOpen.push(commentId);
      }
    }

    if (toClose.length === 0 && toOpen.length === 0) {
      return;
    }

    flushCurrentInline();
    currentTrackedChange = undefined;

    for (const commentId of byCommentId(toClose)) {
      content.push({ type: "commentRangeEnd", id: commentId });
      openedComments.delete(commentId);
    }
    for (const commentId of byCommentId(toOpen)) {
      content.push({ type: "commentRangeStart", id: commentId });
      openedComments.add(commentId);
    }
  };

  const flushEmptyHyperlinksThroughOffset = (offset: number): void => {
    while (nextEmptyHyperlink < sortedEmptyHyperlinks.length) {
      const item = sortedEmptyHyperlinks[nextEmptyHyperlink];
      if (!item || item.attrs.offset > offset) {
        break;
      }
      nextEmptyHyperlink += 1;
      flushCurrentInline();
      currentTrackedChange = undefined;
      content.push(createEmptyHyperlink(item.attrs));
    }
  };

  /**
   * Cut the emitted sequence where the wrapper the nodes sit inside changes.
   *
   * Nothing that spans nodes may span the cut: a run, a hyperlink and a
   * revision each sit inside one wrapper or outside it, never half in. The
   * comment ranges `syncCommentRanges` writes are cut with the rest, which
   * puts a range that opens or closes inside a wrapper inside it, as Word
   * writes it; a range marker says the same thing at either level, because the
   * wrapper is transparent and the position is unchanged.
   */
  const enterWrapperGroup = (node: PMNode): void => {
    const stack = inlineWrapperStackOf(node);
    const key = inlineWrapperStackKey(stack);
    if (key === wrapperGroupKey) {
      return;
    }
    flushCurrentInline();
    currentTrackedChange = undefined;
    wrapperGroupKey = key;
    wrapperGroups.push({ start: content.length, stack });
  };

  const processInlineNode = (node: PMNode, offset: number): void => {
    if (node.type.name === "renderedPageBreak" && leadingRenderedPageBreakPending) {
      leadingRenderedPageBreakPending = false;
      return;
    }
    leadingRenderedPageBreakPending = false;
    enterWrapperGroup(node);
    syncCommentRanges(node, offset);

    // A comment reference is paragraph content in the model, so it never joins
    // a run, a hyperlink or a tracked wrapper: it is emitted where the editor
    // holds it, right after the range ends `syncCommentRanges` just wrote.
    if (node.type.name === "commentReference") {
      flushCurrentInline();
      currentTrackedChange = undefined;
      content.push({ type: "commentReference", id: expectCommentReferenceAttrs(node).commentId });
      return;
    }

    const linkMark = node.marks.find((m) => m.type.name === "hyperlink");

    const noteRefMark = node.marks.find((m) => m.type.name === "footnoteRef");
    const insertionMark = node.marks.find((m) => m.type.name === "insertion");
    const deletionMark = node.marks.find((m) => m.type.name === "deletion");
    if (!insertionMark && !deletionMark) {
      currentTrackedChange = undefined;
    }
    if (node.type.name === "textBoxAnchor") {
      flushCurrentInline();
      currentTrackedChange = undefined;
      if (!textBoxAnchorMarkers) {
        return;
      }
      const { anchorId } = expectTextBoxAnchorAttrs(node);
      const marker: Run = { type: "run", content: [] };
      if (textBoxAnchorMarkers.has(anchorId)) {
        return;
      }
      textBoxAnchorMarkers.set(anchorId, marker);
      const anchoredContent: Run | Hyperlink = linkMark
        ? { ...createHyperlink(linkMark), children: [marker] }
        : marker;
      const changeMark = insertionMark ?? deletionMark;
      if (!changeMark) {
        content.push(anchoredContent);
        return;
      }
      const changeAttrs = expectTrackedChangeMarkAttrs(changeMark);
      const info: TrackedChangeInfo = {
        id: changeAttrs.revisionId,
        author: changeAttrs.author || "Unknown",
        ...(changeAttrs.date ? { date: changeAttrs.date } : {}),
        ...(changeAttrs.utcDate
          ? { utcDate: { attribute: DATE_UTC_ATTRIBUTE, value: changeAttrs.utcDate } }
          : {}),
        ...(changeAttrs.initials ? { initials: changeAttrs.initials } : {}),
      };
      if (insertionMark) {
        content.push({
          type: changeAttrs.moveKind === "moveTo" ? "moveTo" : "insertion",
          info,
          content: [anchoredContent],
        });
      } else {
        content.push({
          type: changeAttrs.moveKind === "moveFrom" ? "moveFrom" : "deletion",
          info,
          content: [anchoredContent],
        });
      }
      return;
    }
    if (insertionMark || deletionMark) {
      // Finish any current content
      flushCurrentInline();

      const changeMark = insertionMark ?? deletionMark;
      if (!changeMark) {
        return;
      }
      const changeAttrs = expectTrackedChangeMarkAttrs(changeMark);
      // Filter out the tracked change mark for text formatting extraction
      const otherMarks = node.marks.filter(
        (m) => m.type.name !== "insertion" && m.type.name !== "deletion",
      );
      const info: TrackedChangeInfo = {
        id: changeAttrs.revisionId,
        author: changeAttrs.author || "Unknown",
      };
      if (changeAttrs.date) {
        info.date = changeAttrs.date;
      }
      if (changeAttrs.utcDate) {
        info.utcDate = { attribute: DATE_UTC_ATTRIBUTE, value: changeAttrs.utcDate };
      }
      if (changeAttrs.initials) {
        info.initials = changeAttrs.initials;
      }
      // The mark itself records whether it originated as a
      // `w:moveTo` / `w:moveFrom`. The previous "is there both an
      // insertion AND a deletion with the same revisionId somewhere
      // in the document?" heuristic was unsound: OOXML doesn't
      // require `w:moveFrom`/`w:moveTo` to share `w:id` (they
      // typically don't), and unrelated `w:ins w:id="5"` /
      // `w:del w:id="5"` from different reviewers would coincidentally
      // fuse into a phantom move pair.
      let type: TrackedRunWrapper["type"];
      if (insertionMark) {
        type = changeAttrs.moveKind === "moveTo" ? "moveTo" : "insertion";
      } else {
        type = changeAttrs.moveKind === "moveFrom" ? "moveFrom" : "deletion";
      }
      const trackedChangeKey = `${type}:${JSON.stringify(info)}`;
      if (linkMark) {
        const linkKey = getLinkKey(linkMark);
        if (
          !currentTrackedChange ||
          currentTrackedChange.type !== "hyperlink" ||
          currentTrackedChange.key !== trackedChangeKey ||
          currentTrackedChange.hyperlinkKey !== linkKey
        ) {
          const hyperlink = createHyperlink(linkMark);
          const wrapper = createTrackedRunWrapper(type, info, hyperlink);
          content.push(wrapper);
          currentTrackedChange = {
            type: "hyperlink",
            key: trackedChangeKey,
            wrapper,
            hyperlink,
            hyperlinkKey: linkKey,
          };
        }

        if (currentTrackedChange.type !== "hyperlink") {
          panic("A tracked hyperlink lost its serialization parent");
        }
        if (node.type.name === "bookmarkBoundary") {
          addNodeToHyperlink({
            ...formattingContext,
            hyperlink: currentTrackedChange.hyperlink,
            node,
            sourceRunOwners,
          });
          return;
        }

        // Inside the link as well as inside the wrapper: a capture authored
        // in a tracked, linked range is accepted and rejected with both, and
        // `CT_R` admits none of the elements this level holds.
        if (isInlineLevelPreservedXml(node)) {
          currentTrackedChange.hyperlink.children.push(createPreservedInline(node));
          return;
        }
        const run = createTrackedChangeRun({
          ...formattingContext,
          marks: otherMarks,
          node,
        });
        if (run) {
          appendRunToHyperlink(currentTrackedChange.hyperlink, run, node, sourceRunOwners);
        }
        return;
      }

      if (
        !currentTrackedChange ||
        currentTrackedChange.type !== "direct" ||
        currentTrackedChange.key !== trackedChangeKey
      ) {
        const wrapper = createTrackedRunWrapper(type, info);
        content.push(wrapper);
        currentTrackedChange = { type: "direct", key: trackedChangeKey, wrapper };
      }
      if (node.type.name === "bookmarkBoundary") {
        currentTrackedChange.wrapper.content.push(
          bookmarkMarkerFromAttrs(expectBookmarkBoundaryAttrs(node)),
        );
        return;
      }
      if (node.type.name === "field" || node.type.name === "structuredField") {
        currentTrackedChange.wrapper.content.push(
          createFieldFromNode(node, {
            ...formattingContext,
            marks: otherMarks,
            textBoxAnchorMarkers,
          }),
        );
        return;
      }
      if (node.type.name === "math") {
        currentTrackedChange.wrapper.content.push(createMathFromNode(node));
        return;
      }
      // Inside the wrapper, not beside it: markup lifted out of a `w:ins` is
      // markup the reviewer no longer accepts or rejects with the change.
      if (isInlineLevelPreservedXml(node)) {
        currentTrackedChange.wrapper.content.push(createPreservedInline(node));
        return;
      }
      const run = createTrackedChangeRun({
        ...formattingContext,
        marks: otherMarks,
        node,
      });
      if (run) {
        appendRunToTrackedWrapper(currentTrackedChange.wrapper, run, node, sourceRunOwners);
      }
      return;
    }

    const ownedRun =
      linkMark || pageBreakRunOwnerId(node) === undefined
        ? null
        : createTrackedChangeRun({ ...formattingContext, marks: node.marks, node });
    if (ownedRun) {
      appendDirectRun(node, ownedRun, node.isText);
      return;
    }

    // A tracked note reference must reach the branch above so it stays inside
    // the w:ins/w:del wrapper. Plain references serialize directly.
    if (noteRefMark && !linkMark) {
      flushCurrentInline();
      content.push(createNoteReferenceRun(noteRefMark, node.marks, formattingContext));
      return;
    }

    if (linkMark) {
      // Start or continue hyperlink
      const linkKey = getLinkKey(linkMark);

      if (currentHyperlink && currentHyperlinkKey === linkKey) {
        // Continue current hyperlink
      } else {
        // Finish previous content
        flushCurrentInline();

        // Start new hyperlink
        currentHyperlink = createHyperlink(linkMark);
        currentHyperlinkKey = linkKey;
      }
      addNodeToHyperlink({
        ...formattingContext,
        hyperlink: currentHyperlink,
        node,
        sourceRunOwners,
      });
      return;
    }

    // Not in hyperlink - finish any current hyperlink
    if (currentHyperlink) {
      flushCurrentInline();
    }

    // Handle node types
    if (node.type.name === "bookmarkBoundary") {
      flushCurrentInline();
      content.push(bookmarkMarkerFromAttrs(expectBookmarkBoundaryAttrs(node)));
    } else if (node.isText) {
      appendDirectRun(
        node,
        createRunFromText({
          ...formattingContext,
          marks: node.marks,
          text: node.text || "",
        }),
        true,
      );
    } else if (node.type.name === "pageBreakRun") {
      const { clear } = expectPageBreakRunAttrs(node);
      appendDirectRun(
        node,
        createPageBreakCarrierRun({
          ...formattingContext,
          clear,
          marks: node.marks,
        }),
      );
    } else if (node.type.name === "symbol") {
      flushCurrentInline();
      content.push(createSymbolRun(node, node.marks, formattingContext));
    } else if (node.type.name === "preservedXml") {
      flushCurrentInline();
      content.push(
        isInlineLevelPreservedXml(node)
          ? createPreservedInline(node)
          : createPreservedXmlRun(node, node.marks, formattingContext),
      );
    } else if (node.type.name === "hardBreak") {
      // Hard break ends current run
      flushCurrentInline();
      content.push(createBreakRun(expectHardBreakAttrs(node), node.marks, formattingContext));
    } else if (node.type.name === "image") {
      // Image ends current run
      flushCurrentInline();
      content.push(createImageRun(node));
    } else if (node.type.name === "shape") {
      // Shape ends current run
      flushCurrentInline();
      content.push(createShapeRun(node));
    } else if (node.type.name === "tab") {
      // Tab ends current run
      flushCurrentInline();
      content.push(createTabRun(node, node.marks, formattingContext));
    } else if (node.type.name === "renderedPageBreak") {
      flushCurrentInline();
      content.push(createRenderedPageBreakRun());
    } else if (node.type.name === "field" || node.type.name === "structuredField") {
      // Field ends current run and emits a field content item
      flushCurrentInline();
      content.push(
        createFieldFromNode(node, {
          ...formattingContext,
          marks: node.marks,
          textBoxAnchorMarkers,
        }),
      );
    } else if (node.type.name === "sdt") {
      // SDT ends current run and emits an InlineSdt content item
      flushCurrentInline();
      content.push(createInlineSdtFromNode(node, textBoxAnchorMarkers, formattingContext));
    } else if (node.type.name === "math") {
      // Math ends current run and emits a MathEquation content item
      flushCurrentInline();
      content.push(createMathFromNode(node));
    }
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  paragraph.forEach((node, offset) => {
    flushEmptyHyperlinksThroughOffset(offset);

    if (node.isText && node.text) {
      let consumed = 0;
      const textEndOffset = offset + node.nodeSize;
      while (nextEmptyHyperlink < sortedEmptyHyperlinks.length) {
        const item = sortedEmptyHyperlinks[nextEmptyHyperlink];
        if (!item || item.attrs.offset >= textEndOffset) {
          break;
        }

        const splitOffset = Math.max(item.attrs.offset - offset, consumed);
        const segment = node.text.slice(consumed, splitOffset);
        if (segment) {
          processInlineNode(node.type.schema.text(segment, node.marks), offset);
        }
        flushCurrentInline();
        content.push(createEmptyHyperlink(item.attrs));
        nextEmptyHyperlink += 1;
        consumed = splitOffset;
      }

      const remainder = node.text.slice(consumed);
      if (remainder) {
        processInlineNode(node.type.schema.text(remainder, node.marks), offset);
      }
      return;
    }

    processInlineNode(node, offset);
  });

  flushEmptyHyperlinksThroughOffset(Number.POSITIVE_INFINITY);

  // Don't forget the last run/hyperlink
  flushCurrentInline();
  for (const commentId of byCommentId(openedComments)) {
    content.push({ type: "commentRangeEnd", id: commentId });
  }

  return nestInlineWrapperGroups(content, wrapperGroups);
}

type CreateTrackedChangeRunOptions = RunFormattingContext & {
  marks: readonly Mark[];
  node: PMNode;
};

function createTrackedChangeRun({
  baseParagraphFormatting,
  inheritedFormatting,
  marks,
  node,
  paragraphMarkFormatting,
  paragraphMarkPrecedesStyle,
  styleResolver,
}: CreateTrackedChangeRunOptions): Run | null {
  const formattingContext = {
    baseParagraphFormatting,
    inheritedFormatting,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle,
    styleResolver,
  };
  const noteRefMark = marks.find((mark) => mark.type.name === "footnoteRef");
  let run: Run | null = null;
  if (noteRefMark) {
    run = createNoteReferenceRun(noteRefMark, marks, formattingContext);
  } else if (node.isText) {
    const formatting = marksToTextFormatting(marks, formattingContext);
    run = {
      type: "run",
      content: node.text ? [{ type: "text", text: node.text }] : [],
      ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
    };
    restoreRunPropertyChanges(run, marks);
  } else if (node.type.name === "symbol") {
    run = createSymbolRun(node, marks, formattingContext);
  } else if (node.type.name === "preservedXml") {
    run = createPreservedXmlRun(node, marks, formattingContext);
  } else if (node.type.name === "hardBreak") {
    run = createBreakRun(expectHardBreakAttrs(node), marks, formattingContext);
  } else if (node.type.name === "pageBreakRun") {
    const { clear } = expectPageBreakRunAttrs(node);
    run = createPageBreakCarrierRun({ ...formattingContext, clear, marks });
  } else if (node.type.name === "image") {
    run = createImageRun(node);
  } else if (node.type.name === "shape") {
    run = createShapeRun(node);
  } else if (node.type.name === "tab") {
    run = createTabRun(node, marks, formattingContext);
  } else if (node.type.name === "renderedPageBreak") {
    run = createRenderedPageBreakRun();
  }

  if (!run || pageBreakRunOwnerId(node) === undefined) {
    return run;
  }
  if (
    node.type.name === "image" ||
    node.type.name === "shape" ||
    node.type.name === "renderedPageBreak"
  ) {
    const formatting = getRunFormattingFromMarks(marks, formattingContext);
    if (formatting) {
      run.formatting = formatting;
    }
    restoreRunPropertyChanges(run, marks);
  }
  return run;
}

function createEmptyHyperlink(
  attrs: NonNullable<ParagraphAttrs["_emptyHyperlinks"]>[number],
): Hyperlink {
  const hyperlink: Hyperlink = { type: "hyperlink", children: [] };
  if (attrs.href !== undefined) {
    hyperlink.href = attrs.href;
  }
  if (attrs.anchor !== undefined) {
    hyperlink.anchor = attrs.anchor;
  }
  if (attrs.tooltip !== undefined) {
    hyperlink.tooltip = attrs.tooltip;
  }
  if (attrs.rId !== undefined) {
    hyperlink.rId = attrs.rId;
  }
  return hyperlink;
}

function getCommentMarkIds(marks: readonly Mark[]): Set<number> {
  const commentIds = new Set<number>();
  for (const mark of marks) {
    if (mark.type.name === "comment") {
      commentIds.add(expectCommentMarkAttrs(mark).commentId);
    }
  }
  return commentIds;
}

type TrackedChangeCounts = {
  insertionById: Map<number, number>;
  deletionById: Map<number, number>;
};

/**
 * Build document-wide tracked change counts by scanning all nodes.
 * Used for cross-paragraph move pair detection (moveFrom in one paragraph,
 * moveTo in another).
 */
function buildDocumentTrackedChangeCounts(pmDoc: PMNode): TrackedChangeCounts {
  const insertionById = new Map<number, number>();
  const deletionById = new Map<number, number>();

  pmDoc.descendants((node) => {
    const insertionMark = node.marks.find((m) => m.type.name === "insertion");
    const deletionMark = node.marks.find((m) => m.type.name === "deletion");

    if (insertionMark) {
      const { revisionId } = expectTrackedChangeMarkAttrs(insertionMark);
      insertionById.set(revisionId, (insertionById.get(revisionId) ?? 0) + 1);
    }
    if (deletionMark) {
      const { revisionId } = expectTrackedChangeMarkAttrs(deletionMark);
      deletionById.set(revisionId, (deletionById.get(revisionId) ?? 0) + 1);
    }
  });

  return { insertionById, deletionById };
}

/**
 * Create a unique key for a link mark
 */
function getLinkKey(mark: Mark): string {
  const attrs = expectHyperlinkMarkAttrs(mark);
  return [attrs.href, attrs.rId ?? "", attrs.tooltip ?? "", attrs._docxHyperlinkIndex ?? ""].join(
    "\u0000",
  );
}

/**
 * Create a unique key for a set of marks (excluding hyperlink)
 */
function getMarksKey(marks: readonly Mark[]): string {
  const nonLinkMarks = marks.filter((m) => m.type.name !== "hyperlink");
  if (nonLinkMarks.length === 0) {
    return "";
  }

  return nonLinkMarks
    .map((m) => `${m.type.name}:${JSON.stringify(m.attrs)}`)
    .toSorted()
    .join("|");
}

/**
 * Create a Hyperlink from a link mark
 */
function createHyperlink(linkMark: Mark): Hyperlink {
  const attrs = expectHyperlinkMarkAttrs(linkMark);
  const href = attrs.href;
  // Internal bookmark links use the anchor property in OOXML
  if (href.startsWith("#")) {
    const hyperlink: Hyperlink = {
      type: "hyperlink",
      anchor: href.slice(1),
      children: [],
    };
    if (attrs.tooltip) {
      hyperlink.tooltip = attrs.tooltip;
    }
    return hyperlink;
  }
  const hyperlink: Hyperlink = {
    type: "hyperlink",
    href,
    children: [],
  };
  if (attrs.tooltip) {
    hyperlink.tooltip = attrs.tooltip;
  }
  if (attrs.rId) {
    hyperlink.rId = attrs.rId;
  }
  return hyperlink;
}

/**
 * Add a node to a hyperlink
 */
type AddNodeToHyperlinkOptions = RunFormattingContext & {
  hyperlink: Hyperlink;
  node: PMNode;
  sourceRunOwners: WeakMap<Run, number>;
};

function addNodeToHyperlink({
  baseParagraphFormatting,
  hyperlink,
  inheritedFormatting,
  node,
  paragraphMarkFormatting,
  paragraphMarkPrecedesStyle,
  sourceRunOwners,
  styleResolver,
}: AddNodeToHyperlinkOptions): void {
  if (node.type.name === "bookmarkBoundary") {
    hyperlink.children.push(bookmarkMarkerFromAttrs(expectBookmarkBoundaryAttrs(node)));
    return;
  }
  const formattingContext = {
    baseParagraphFormatting,
    inheritedFormatting,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle,
    styleResolver,
  };
  const nonLinkMarks = node.marks.filter((mark) => mark.type.name !== "hyperlink");
  const ownedRun =
    pageBreakRunOwnerId(node) === undefined
      ? null
      : createTrackedChangeRun({ ...formattingContext, marks: nonLinkMarks, node });
  if (ownedRun) {
    appendRunToHyperlink(hyperlink, ownedRun, node, sourceRunOwners);
    return;
  }

  const noteRefMark = node.marks.find((m) => m.type.name === "footnoteRef");
  if (noteRefMark) {
    hyperlink.children.push(
      createNoteReferenceRun(noteRefMark, node.marks, {
        baseParagraphFormatting,
        inheritedFormatting,
        paragraphMarkFormatting,
        paragraphMarkPrecedesStyle,
        styleResolver,
      }),
    );
    return;
  }

  if (node.isText && node.text) {
    const run = createRunFromText({
      baseParagraphFormatting,
      inheritedFormatting,
      marks: nonLinkMarks,
      paragraphMarkFormatting,
      paragraphMarkPrecedesStyle,
      styleResolver,
      text: node.text,
    });
    hyperlink.children.push(run);
    return;
  }

  if (node.type.name === "symbol") {
    hyperlink.children.push(
      createSymbolRun(node, nonLinkMarks, {
        baseParagraphFormatting,
        inheritedFormatting,
        paragraphMarkFormatting,
        paragraphMarkPrecedesStyle,
        styleResolver,
      }),
    );
    return;
  }

  if (node.type.name === "preservedXml") {
    // The atom records the level it was read at, and the two go back to
    // different places: `w:ruby` is a run child and `w:permStart` is a link
    // child, because `CT_Hyperlink` admits the second and `CT_R` does not.
    hyperlink.children.push(
      isInlineLevelPreservedXml(node)
        ? createPreservedInline(node)
        : createPreservedXmlRun(node, nonLinkMarks, {
            baseParagraphFormatting,
            inheritedFormatting,
            paragraphMarkFormatting,
            paragraphMarkPrecedesStyle,
            styleResolver,
          }),
    );
    return;
  }

  if (node.type.name === "hardBreak") {
    hyperlink.children.push(
      createBreakRun(expectHardBreakAttrs(node), nonLinkMarks, {
        baseParagraphFormatting,
        inheritedFormatting,
        paragraphMarkFormatting,
        paragraphMarkPrecedesStyle,
        styleResolver,
      }),
    );
    return;
  }

  if (node.type.name === "pageBreakRun") {
    const { clear } = expectPageBreakRunAttrs(node);
    hyperlink.children.push(
      createPageBreakCarrierRun({ ...formattingContext, clear, marks: nonLinkMarks }),
    );
    return;
  }

  if (node.type.name === "tab") {
    hyperlink.children.push(
      createTabRun(node, nonLinkMarks, {
        baseParagraphFormatting,
        inheritedFormatting,
        paragraphMarkFormatting,
        paragraphMarkPrecedesStyle,
        styleResolver,
      }),
    );
    return;
  }

  if (node.type.name === "renderedPageBreak") {
    hyperlink.children.push(createRenderedPageBreakRun());
    return;
  }

  if (node.type.name === "image") {
    hyperlink.children.push(createImageRun(node));
    return;
  }

  if (node.type.name === "shape") {
    hyperlink.children.push(createShapeRun(node));
  }
}

function getNoteReferenceVertAlign(
  noteAttrs: ReturnType<typeof expectFootnoteRefMarkAttrs>,
  marks: readonly Mark[],
): "baseline" | "superscript" | "subscript" | undefined {
  if (marks.some((mark) => mark.type.name === "subscript")) {
    return "subscript";
  }
  if (marks.some((mark) => mark.type.name === "superscript")) {
    return "superscript";
  }
  if (noteAttrs.vertAlign === "baseline" || noteAttrs.vertAlign === "superscript") {
    return noteAttrs.vertAlign;
  }
  return undefined;
}

function createNoteReferenceRun(
  noteRefMark: Mark,
  marks: readonly Mark[],
  formattingContext?: MarksToTextFormattingOptions,
): Run {
  const noteAttrs = expectFootnoteRefMarkAttrs(noteRefMark);
  const noteType = noteAttrs.noteType === "endnote" ? "endnoteRef" : "footnoteRef";
  const noteId =
    typeof noteAttrs.id === "string" ? Number.parseInt(noteAttrs.id, 10) || 0 : noteAttrs.id;
  const noteRef: NoteReferenceContent = {
    type: noteType,
    id: noteId,
  };
  const run: Run = {
    type: "run",
    content: [noteRef],
  };
  const formatting = getAtomRunFormattingFromMarks(marks, formattingContext);
  const vertAlign = getNoteReferenceVertAlign(noteAttrs, marks);
  if (formatting) {
    run.formatting = formatting;
  }
  if (vertAlign && formatting?.vertAlign === undefined && formatting?.styleId === undefined) {
    run.formatting = { ...formatting, vertAlign };
  }
  restoreRunPropertyChanges(run, marks);
  return run;
}

/**
 * Create a Run from text and marks
 */
type CreateRunFromTextOptions = RunFormattingContext & {
  marks: readonly Mark[];
  text: string;
};

function createRunFromText({
  baseParagraphFormatting,
  inheritedFormatting,
  marks,
  paragraphMarkFormatting,
  paragraphMarkPrecedesStyle,
  styleResolver,
  text,
}: CreateRunFromTextOptions): Run {
  const formatting = getRunFormattingFromMarks(marks, {
    baseParagraphFormatting,
    inheritedFormatting,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle,
    styleResolver,
  });
  const textContent: TextContent = {
    type: "text",
    text,
  };

  const run: Run = { type: "run", content: [textContent] };
  if (formatting) {
    run.formatting = formatting;
  }
  restoreRunPropertyChanges(run, marks);
  return run;
}

function createSymbolRun(
  node: PMNode,
  marks: readonly Mark[],
  formattingContext?: MarksToTextFormattingOptions,
): Run {
  const { font, char } = expectSymbolAttrs(node);
  const symbolContent: SymbolContent = { type: "symbol", font, char };
  const run: Run = { type: "run", content: [symbolContent] };
  const formatting = getAtomRunFormattingFromMarks(marks, formattingContext);
  if (formatting) {
    run.formatting = formatting;
  }
  restoreRunPropertyChanges(run, marks);
  return run;
}

/**
 * Whether the atom's markup is a paragraph child rather than a run child.
 *
 * The two are the same node in the editor and different elements in the file:
 * `w:ruby` has to go back inside a `w:r` and `w:permStart` may not, because
 * the schema admits no such child of a run and Word reports the package as
 * unreadable content.
 */
const isInlineLevelPreservedXml = (node: PMNode): boolean =>
  node.type.name === "preservedXml" &&
  expectPreservedXmlAttrs(node).level === PRESERVED_XML_LEVELS.inline;

/** The paragraph-level capture an inline-level atom writes back. */
const createPreservedInline = (node: PMNode): PreservedInline => {
  const { xml, text } = expectPreservedXmlAttrs(node);
  return { type: "preservedInline", xml, text };
};

/**
 * Rebuild the run around a preserved child. The markup is opaque and comes
 * back byte for byte; only the run properties around it are rebuilt from the
 * atom's marks, exactly as for a symbol.
 */
function createPreservedXmlRun(
  node: PMNode,
  marks: readonly Mark[],
  formattingContext?: MarksToTextFormattingOptions,
): Run {
  const { xml, text } = expectPreservedXmlAttrs(node);
  const run: Run = { type: "run", content: [{ type: "preservedXml", xml, text }] };
  const formatting = getAtomRunFormattingFromMarks(marks, formattingContext);
  if (formatting) {
    run.formatting = formatting;
  }
  restoreRunPropertyChanges(run, marks);
  return run;
}

function restoreRunPropertyChanges(run: Run, marks: readonly Mark[]): void {
  const changeMark = marks.find((mark) => mark.type.name === "runPropertyChange");
  if (!changeMark) {
    return;
  }
  const { changes } = expectRunPropertyChangeMarkAttrs(changeMark);
  if (changes.length === 0) {
    return;
  }
  run.propertyChanges = [...changes];
}

function getRunFormattingFromMarks(
  marks: readonly Mark[] | undefined,
  options?: MarksToTextFormattingOptions,
): TextFormatting | undefined {
  if (!marks || (marks.length === 0 && !options)) {
    return undefined;
  }

  const formatting = marksToTextFormatting(marks, options);
  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

function getAtomRunFormattingFromMarks(
  marks: readonly Mark[] | undefined,
  options?: MarksToTextFormattingOptions,
): TextFormatting | undefined {
  if (!marks?.some(({ type }) => RUN_FORMATTING_MARK_NAMES.has(type.name))) {
    return undefined;
  }
  return getRunFormattingFromMarks(marks, options);
}

/**
 * Append text to an existing run
 */
function appendTextToRun(run: Run, text: string): void {
  const lastContent = run.content.at(-1);
  if (lastContent && lastContent.type === "text") {
    lastContent.text += text;
  } else {
    run.content.push({ type: "text", text });
  }
}

function runsShareProperties(left: Run, right: Run): boolean {
  return (
    canonicalJson({
      formatting: left.formatting,
      propertyChanges: left.propertyChanges,
    }) ===
    canonicalJson({
      formatting: right.formatting,
      propertyChanges: right.propertyChanges,
    })
  );
}

function pageBreakRunOwnerId(node: PMNode): number | undefined {
  const mark = node.marks.find(({ type }) => type.name === "pageBreakRunOwner");
  return mark ? expectPageBreakRunOwnerMarkAttrs(mark).id : undefined;
}

function canJoinOwnedRuns(
  previous: Run | undefined,
  run: Run,
  node: PMNode,
  sourceRunOwners: WeakMap<Run, number>,
): previous is Run {
  if (!previous || !runsShareProperties(previous, run)) {
    return false;
  }
  const ownerId = pageBreakRunOwnerId(node);
  return ownerId !== undefined && sourceRunOwners.get(previous) === ownerId;
}

function rememberSourceRunOwner(
  run: Run,
  node: PMNode,
  sourceRunOwners: WeakMap<Run, number>,
): void {
  const ownerId = pageBreakRunOwnerId(node);
  if (ownerId !== undefined) {
    sourceRunOwners.set(run, ownerId);
  }
}

function appendRunToTrackedWrapper(
  wrapper: TrackedRunWrapper,
  run: Run,
  node: PMNode,
  sourceRunOwners: WeakMap<Run, number>,
): void {
  const previous = wrapper.content.at(-1);
  if (previous?.type === "run" && canJoinOwnedRuns(previous, run, node, sourceRunOwners)) {
    previous.content.push(...run.content);
    return;
  }
  wrapper.content.push(run);
  rememberSourceRunOwner(run, node, sourceRunOwners);
}

function appendRunToHyperlink(
  hyperlink: Hyperlink,
  run: Run,
  node: PMNode,
  sourceRunOwners: WeakMap<Run, number>,
): void {
  const previous = hyperlink.children.at(-1);
  if (previous?.type === "run" && canJoinOwnedRuns(previous, run, node, sourceRunOwners)) {
    previous.content.push(...run.content);
    return;
  }
  hyperlink.children.push(run);
  rememberSourceRunOwner(run, node, sourceRunOwners);
}

type CreatePageBreakCarrierRunOptions = RunFormattingContext & {
  clear?: BreakContent["clear"];
  marks?: readonly Mark[];
};

function createPageBreakCarrierRun({
  baseParagraphFormatting,
  clear,
  inheritedFormatting,
  marks,
  paragraphMarkFormatting,
  paragraphMarkPrecedesStyle,
  styleResolver,
}: CreatePageBreakCarrierRunOptions): Run {
  const run: Run = {
    type: "run",
    content: [
      {
        type: "break",
        breakType: "page",
        ...(clear !== undefined ? { clear } : {}),
      },
    ],
  };
  const formatting = getAtomRunFormattingFromMarks(marks, {
    baseParagraphFormatting,
    inheritedFormatting,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle,
    styleResolver,
  });
  if (formatting) {
    run.formatting = formatting;
  }
  if (marks) {
    restoreRunPropertyChanges(run, marks);
  }
  return run;
}

/**
 * Create a Run containing a line break
 */
function createBreakRun(
  attrs: Pick<BreakContent, "breakType" | "clear">,
  marks?: readonly Mark[],
  formattingContext?: MarksToTextFormattingOptions,
): Run {
  const breakContent: BreakContent = {
    type: "break",
    ...(attrs.breakType !== undefined ? { breakType: attrs.breakType } : {}),
    ...(attrs.clear !== undefined ? { clear: attrs.clear } : {}),
  };

  const run: Run = {
    type: "run",
    content: [breakContent],
  };
  const formatting = getAtomRunFormattingFromMarks(marks, formattingContext);
  if (formatting) {
    run.formatting = formatting;
  }
  if (marks) {
    restoreRunPropertyChanges(run, marks);
  }
  return run;
}

function createRenderedPageBreakRun(): Run {
  return {
    type: "run",
    content: [{ type: "renderedPageBreak" }],
  };
}

/**
 * Create a Run containing a tab
 */
function createTabRun(
  node: PMNode,
  marks?: readonly Mark[],
  formattingContext?: MarksToTextFormattingOptions,
): Run {
  const { positional } = expectTabAttrs(node);
  const tabContent: TabContent = {
    type: "tab",
    ...(positional ? { positional } : {}),
  };

  const run: Run = {
    type: "run",
    content: [tabContent],
  };
  const formatting = getAtomRunFormattingFromMarks(marks, formattingContext);
  if (formatting) {
    run.formatting = formatting;
  }
  if (marks) {
    restoreRunPropertyChanges(run, marks);
  }
  return run;
}

/**
 * Create a SimpleField or ComplexField from a PM field node
 */
type CreateFieldFromNodeOptions = Partial<RunFormattingContext> & {
  marks?: readonly Mark[];
  textBoxAnchorMarkers?: Map<string, Run> | undefined;
};

function createFieldFromNode(
  node: PMNode,
  {
    baseParagraphFormatting,
    inheritedFormatting,
    marks,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle,
    styleResolver,
    textBoxAnchorMarkers,
  }: CreateFieldFromNodeOptions,
): SimpleField | ComplexField {
  const attrs = expectFieldAttrs(node);
  const formatting =
    marks && marks.length > 0
      ? marksToTextFormatting(marks, {
          baseParagraphFormatting,
          inheritedFormatting,
          paragraphMarkFormatting,
          paragraphMarkPrecedesStyle,
          styleResolver,
        })
      : undefined;
  const fieldFormattingContext = {
    baseParagraphFormatting,
    inheritedFormatting,
    paragraphMarkFormatting,
    paragraphMarkPrecedesStyle: paragraphMarkPrecedesStyle ?? false,
    styleResolver: styleResolver ?? null,
  };
  const extractedContent = extractParagraphContent(
    node,
    undefined,
    undefined,
    textBoxAnchorMarkers,
    false,
    fieldFormattingContext,
  ).filter(
    (content): content is SimpleField["content"][number] =>
      content.type === "run" || content.type === "hyperlink" || content.type === "preservedInline",
  );
  // A result-less PAGE/NUMPAGES field gets its visible fallback from
  // `materializeSerializerFieldFallbacks`, before the walk reaches here, so a
  // read can opt out of it while a save keeps it.
  const displayText = attrs.displayText ?? "";

  const displayRun: Run = {
    type: "run",
    content: [{ type: "text" as const, text: displayText }],
    ...(formatting && Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
  if (marks && node.type.name === "field") {
    restoreRunPropertyChanges(displayRun, marks);
  }
  const fieldContent =
    extractedContent.length > 0
      ? synchronizeFieldDisplayText(extractedContent, displayText, displayRun)
      : [];
  const fallbackFieldContent = displayText ? [displayRun] : [];

  if (attrs.fieldKind === "complex") {
    const complex: ComplexField = {
      type: "complexField",
      instruction: attrs.instruction,
      fieldType: attrs.fieldType,
      fieldCode: [],
      fieldResult:
        fieldContent.length > 0
          ? fieldContent.filter((content): content is Run => content.type === "run")
          : fallbackFieldContent,
    };
    if (attrs.fldLock !== undefined) {
      complex.fldLock = attrs.fldLock;
    }
    if (attrs.dirty !== undefined) {
      complex.dirty = attrs.dirty;
    }
    return complex;
  }

  const simple: SimpleField = {
    type: "simpleField",
    instruction: attrs.instruction,
    fieldType: attrs.fieldType,
    content: fieldContent.length > 0 ? fieldContent : fallbackFieldContent,
  };
  if (attrs.fldLock !== undefined) {
    simple.fldLock = attrs.fldLock;
  }
  if (attrs.dirty !== undefined) {
    simple.dirty = attrs.dirty;
  }
  return simple;
}

const synchronizeFieldDisplayText = (
  content: SimpleField["content"],
  displayText: string,
  fallbackRun: Run,
): SimpleField["content"] => {
  let currentText = "";
  const visitRuns = (visit: (run: Run) => void): void => {
    for (const child of content) {
      if (child.type === "run") {
        visit(child);
        continue;
      }
      if (child.type !== "hyperlink") {
        continue;
      }
      for (const hyperlinkChild of child.children) {
        if (hyperlinkChild.type === "run") {
          visit(hyperlinkChild);
        }
      }
    }
  };
  visitRuns((run) => {
    for (const runContent of run.content) {
      if (runContent.type === "text") {
        currentText += runContent.text;
      }
    }
  });
  if (currentText === displayText) {
    return content;
  }

  let replacedText = false;
  visitRuns((run) => {
    for (const runContent of run.content) {
      if (runContent.type !== "text") {
        continue;
      }
      runContent.text = replacedText ? "" : displayText;
      replacedText = true;
    }
  });
  if (replacedText) {
    return content;
  }

  const hyperlink = content.find((child) => child.type === "hyperlink");
  if (!hyperlink) {
    content.push(fallbackRun);
    return content;
  }
  const firstEnd = hyperlink.children.findIndex((child) => child.type === "bookmarkEnd");
  hyperlink.children.splice(firstEnd < 0 ? hyperlink.children.length : firstEnd, 0, fallbackRun);
  return content;
};

/**
 * Create a MathEquation from a PM math node
 */
function createMathFromNode(node: PMNode): MathEquation {
  const attrs = expectMathAttrs(node);

  const math: MathEquation = {
    type: "mathEquation",
    display: attrs.display ?? "inline",
    ommlXml: attrs.ommlXml,
  };
  if (attrs.plainText) {
    math.plainText = attrs.plainText;
  }
  return math;
}

/**
 * A revision covering every child of a control, written around the control.
 *
 * `w:ins > w:sdt` and `w:sdt > w:ins` reach the editor as the same revision
 * mark on the same leaves, so the save leg has to pick one and writes the
 * canonical order: the revision outermost, as it already is around a hyperlink
 * and around a transparent wrapper. Outermost is also the only form in which
 * accepting or rejecting the change is an operation over the whole control —
 * the reader who inserted a bound field inserted the field, not its text.
 *
 * A revision that covers only part of the content has no such form and stays
 * where the editor holds it, per child.
 */
const hoistUniformRevision = (sdt: InlineSdt): InlineSdt | TrackedRunWrapper => {
  const only = sdt.content.length === 1 ? sdt.content.at(0) : undefined;
  if (only === undefined || !isRevisionWrapper(only)) {
    return sdt;
  }
  // `CT_RunTrackChange` admits the bookmark boundaries `CT_SdtContentRun` does
  // not, and moving one inside the control would write markup the schema
  // rejects. Nothing is hoisted in that case.
  const admitted = only.content.filter(isInlineSdtContent);
  if (admitted.length !== only.content.length) {
    return sdt;
  }
  return { ...only, content: [{ ...sdt, content: admitted }] };
};

/**
 * Create an InlineSdt from a PM sdt node
 */
function createInlineSdtFromNode(
  node: PMNode,
  textBoxAnchorMarkers?: Map<string, Run>,
  formattingContext?: RunFormattingContext,
): InlineSdt | TrackedRunWrapper {
  const attrs = expectSdtAttrs(node);
  const properties = sdtPropertiesFromAttrs(attrs);

  // The control keeps everything `CT_SdtContentRun` admits, so docProps-bound
  // fields, reviewed template content and markup folio does not model all
  // survive the round trip. What it admits is `INLINE_SDT_CONTENT`, bound to
  // the content type itself, so this cannot drift from the model, from the
  // parser, or from `serializeInlineSdt`.
  const content = extractParagraphContent(
    node,
    undefined,
    undefined,
    textBoxAnchorMarkers,
    false,
    formattingContext,
  ).filter(isInlineSdtContent);

  return hoistUniformRevision({
    type: "inlineSdt",
    properties,
    content,
  });
}

/**
 * Create a Run containing an image
 */
/**
 * Restore the `w:rPr` an inline atom carried from its source run.
 *
 * Image and shape nodes do not hold the run's formatting marks
 * (`withRunBoundaryMarks` in toProseDoc keeps them off), so the run rebuilt
 * around one has no other record of the properties it was authored with.
 */
const carriedRunFormatting = (
  formatting: TextFormatting | undefined,
): Pick<Run, "formatting"> | Record<string, never> =>
  formatting && Object.keys(formatting).length > 0 ? { formatting } : {};

/**
 * The anchor record off a node's attrs, copied so the model never aliases a
 * ProseMirror attribute object.
 */
const restoredDrawingAnchor = (anchor: DrawingAnchor | undefined): DrawingAnchor | undefined =>
  anchor === undefined
    ? undefined
    : {
        ...anchor,
        ...(anchor.simplePosition === undefined
          ? {}
          : { simplePosition: { ...anchor.simplePosition } }),
      };

function createImageRun(node: PMNode): Run {
  const attrs = expectImageAttrs(node);

  // Determine wrap type from attrs (default: inline)
  const wrapType = attrs.wrapType || "inline";

  const wrap: ImageWrap = { type: wrapType };
  assignWrapDistances(wrap, attrs);

  // Restore wrapText from PM attr
  if (attrs.wrapText) {
    wrap.wrapText = attrs.wrapText;
  }

  const authoredEmu = attrs._docxAuthoredEmu;
  // The attr is an unvalidated string, and a node projected from a drawing
  // that carries no relationship holds nothing or holds `""`; both are the
  // same fact, and the model has one spelling for it.
  const rId = relationshipIdOf(attrs.rId);
  const image: Image = {
    type: "image",
    ...(rId === undefined ? {} : { rId }),
    src: attrs.src,
    size: {
      width: emuFromPixels(attrs.width || 0, "width", authoredEmu, emuToPixels),
      height: emuFromPixels(attrs.height || 0, "height", authoredEmu, emuToPixels),
    },
    wrap,
  };
  if (attrs.docPrName != null) {
    image.docPrName = attrs.docPrName;
  }
  if (attrs.alt != null) {
    image.alt = attrs.alt;
  }
  if (attrs.title != null) {
    image.title = attrs.title;
  }

  const imageTransform = readAuthoredTransform(attrs);
  if (imageTransform) {
    image.transform = imageTransform;
  }

  // eigenpal #424 (opacity render pipeline). PM schema default is `null`;
  // use `!= null` so the model only carries an explicit opacity value.
  if (attrs.opacity != null) {
    image.opacity = attrs.opacity;
  }

  if (attrs.brightness != null || attrs.contrast != null) {
    image.effects = {
      ...(attrs.brightness == null ? {} : { brightness: attrs.brightness }),
      ...(attrs.contrast == null ? {} : { contrast: attrs.contrast }),
    };
  }

  const imagePosition = imagePositionFromAttrs(attrs.position);
  if (imagePosition) {
    image.position = imagePosition;
  }
  const imageAnchor = restoredDrawingAnchor(attrs.anchor);
  if (imageAnchor) {
    image.anchor = imageAnchor;
  }
  if (attrs.decorative !== undefined) {
    image.decorative = attrs.decorative;
  }
  if (attrs.hidden !== undefined) {
    image.hidden = attrs.hidden;
  }
  if (attrs.docPrExtensions !== undefined) {
    // Copy, so the model never aliases the PM attr array.
    image.docPrExtensions = [...attrs.docPrExtensions];
  }
  if (attrs.frameLocks !== undefined) {
    // Copy, so the model never aliases the PM attr object.
    image.frameLocks = { ...attrs.frameLocks };
  }

  // Round-trip border/outline
  if (attrs.borderWidth && attrs.borderWidth > 0) {
    const outline: ShapeOutline = {
      width: emuFromPixels(attrs.borderWidth, "borderWidth", authoredEmu, emuToStrokePixels),
      dash: presetDashForCssBorderStyle(attrs.borderStyle),
    };
    if (attrs.borderColor) {
      outline.color = { rgb: attrs.borderColor.replace("#", "") };
    }
    image.outline = outline;
  }

  // Round-trip image hyperlink
  if (attrs.hlinkHref) {
    image.hlinkHref = attrs.hlinkHref;
  }
  if (attrs.docPrId !== undefined) {
    image.id = attrs.docPrId;
  }
  if (attrs.hlinkClickSource !== undefined) {
    image.hlinkClickSource = { ...attrs.hlinkClickSource };
  }
  if (attrs.hlinkHoverXml !== undefined) {
    image.hlinkHoverXml = attrs.hlinkHoverXml;
  }
  if (attrs.hlinkRId) {
    image.hlinkRId = attrs.hlinkRId;
  }

  // eigenpal #424: fold crop fractions back into Image.crop. PM defaults are
  // `null`, so `!= null` catches both null and undefined; zero sides are
  // omitted to keep the serialized <a:srcRect/> terse.
  const cropTop = attrs.cropTop != null && attrs.cropTop > 0 ? attrs.cropTop : undefined;
  const cropRight = attrs.cropRight != null && attrs.cropRight > 0 ? attrs.cropRight : undefined;
  const cropBottom =
    attrs.cropBottom != null && attrs.cropBottom > 0 ? attrs.cropBottom : undefined;
  const cropLeft = attrs.cropLeft != null && attrs.cropLeft > 0 ? attrs.cropLeft : undefined;
  if (
    cropTop !== undefined ||
    cropRight !== undefined ||
    cropBottom !== undefined ||
    cropLeft !== undefined
  ) {
    const crop: NonNullable<Image["crop"]> = {};
    if (cropTop !== undefined) {
      crop.top = cropTop;
    }
    if (cropRight !== undefined) {
      crop.right = cropRight;
    }
    if (cropBottom !== undefined) {
      crop.bottom = cropBottom;
    }
    if (cropLeft !== undefined) {
      crop.left = cropLeft;
    }
    image.crop = crop;
  }

  // Fold wp:effectExtent back into Image.padding. Unlike crop, a zero side is
  // kept: the parser only produces padding when some side is non-zero, and it
  // then records all four, so dropping zeros would reshape the parsed model.
  const { paddingTop, paddingRight, paddingBottom, paddingLeft } = attrs;
  if (
    paddingTop !== undefined ||
    paddingRight !== undefined ||
    paddingBottom !== undefined ||
    paddingLeft !== undefined
  ) {
    const padding: NonNullable<Image["padding"]> = {};
    if (paddingTop !== undefined) {
      padding.top = paddingTop;
    }
    if (paddingRight !== undefined) {
      padding.right = paddingRight;
    }
    if (paddingBottom !== undefined) {
      padding.bottom = paddingBottom;
    }
    if (paddingLeft !== undefined) {
      padding.left = paddingLeft;
    }
    image.padding = padding;
  }

  return {
    type: "run",
    ...carriedRunFormatting(attrs._docxRunFormatting),
    content: [drawingFromImageAttrs(image, attrs)],
  };
}

/**
 * Rebuild the `DrawingContent` union member the image node came from.
 *
 * Each branch is listed in full rather than spread over a base object, so a
 * field from one mode cannot leak into another.
 */
const drawingFromImageAttrs = (image: Image, attrs: ImageAttrs): DrawingContent => {
  const mode = attrs._docxRawXmlMode;
  switch (mode) {
    case undefined:
      return {
        type: "drawing",
        image,
        ...(attrs._docxRawXml ? { rawXml: attrs._docxRawXml } : {}),
        ...(attrs._docxRawXml ? { rawImageFingerprint: imageRawXmlFingerprint(image) } : {}),
      };
    case DRAWING_RAW_XML_MODES.PRESERVE_ONLY:
      return {
        type: "drawing",
        image,
        rawXml:
          attrs._docxRawXml ??
          panic("Preservation-only ProseMirror image attrs must include raw XML."),
        rawXmlMode: DRAWING_RAW_XML_MODES.PRESERVE_ONLY,
      };
    case DRAWING_RAW_XML_MODES.PREVIEW_ONLY:
      // `mergeImageAttrs` drops the capture marker when an edit reaches a
      // preview, so its absence — not its value — reports the edit. The
      // captured value cannot be compared against the image rebuilt here: the
      // projection is lossy (EMU→px→EMU rounding alone shifts the size), so a
      // round-trip that changed nothing would look edited.
      return {
        type: "drawing",
        image,
        rawXml:
          attrs._docxRawXml ?? panic("Preview-only ProseMirror image attrs must include raw XML."),
        rawImageFingerprint:
          attrs._docxRawImageFingerprint === undefined
            ? EDITED_PREVIEW_FINGERPRINT
            : imageRawXmlFingerprint(image),
        rawXmlMode: DRAWING_RAW_XML_MODES.PREVIEW_ONLY,
      };
    default:
      return mode satisfies never;
  }
};

/**
 * Create a Run from a ProseMirror shape node
 */
function createShapeRun(node: PMNode): Run {
  const attrs = expectShapeAttrs(node);

  const authoredEmu = attrs._docxAuthoredEmu;
  const shape: Shape = {
    type: "shape",
    shapeType: (attrs.shapeType || "rect") as Shape["shapeType"],
    size: {
      width: attrs.width ? emuFromPixels(attrs.width, "width", authoredEmu, emuToPixels) : 0,
      height: attrs.height ? emuFromPixels(attrs.height, "height", authoredEmu, emuToPixels) : 0,
    },
  };
  if (attrs.shapeId) {
    shape.id = attrs.shapeId;
  }
  // `""` is a name someone wrote, so presence is the test, not truthiness.
  if (attrs.shapeName !== undefined) {
    shape.name = attrs.shapeName;
  }
  if (attrs.alt !== undefined) {
    shape.alt = attrs.alt;
  }
  if (attrs.title !== undefined) {
    shape.title = attrs.title;
  }
  const geometryAdjustments = parseShapeGeometryAdjustments(attrs.geometryAdjustments);
  if (geometryAdjustments !== undefined) {
    shape.geometryAdjustments = geometryAdjustments;
  }
  const shapeTransform = readAuthoredTransform(attrs);
  if (shapeTransform) {
    shape.transform = shapeTransform;
  }

  const wrap: ImageWrap = { type: attrs.wrapType || "inline" };
  assignWrapDistances(wrap, attrs);
  if (attrs.wrapText) {
    wrap.wrapText = attrs.wrapText;
  }
  shape.wrap = wrap;

  const shapePosition = imagePositionFromAttrs(attrs.position);
  if (shapePosition) {
    shape.position = shapePosition;
  }
  const shapeAnchor = restoredDrawingAnchor(attrs.anchor);
  if (shapeAnchor) {
    shape.anchor = shapeAnchor;
  }

  // Fill
  if (attrs.fillType === "gradient" && attrs.gradientStops) {
    // Round-trip gradient fill
    try {
      const parsed = JSON.parse(attrs.gradientStops) as {
        position: number;
        color: string;
      }[];
      const gradient: NonNullable<ShapeFill["gradient"]> = {
        type: (attrs.gradientType || "linear") as "linear" | "radial" | "rectangular" | "path",
        stops: parsed.map((s) => ({
          position: s.position,
          color: { rgb: s.color.replace("#", "") },
        })),
      };
      if (attrs.gradientAngle) {
        gradient.angle = attrs.gradientAngle;
      }
      shape.fill = { type: "gradient", gradient };
    } catch {
      shape.fill = {
        type: "solid",
        color: { rgb: (attrs.fillColor || "000000").replace("#", "") },
      };
    }
  } else if (attrs.fillColor) {
    shape.fill = {
      type: attrs.fillType ?? "solid",
      color: attrs.fillColorValue ?? { rgb: attrs.fillColor.replace("#", "") },
    };
  } else if (attrs.fillType === "none") {
    shape.fill = { type: "none" };
  }

  // Outline. `outlineStyle === "none"` is the explicit "no outline" sentinel
  // (see OUTLINE_STYLE_ATTR_VALUES): suppress the `<a:ln>` element entirely so a
  // border-free shape round-trips border-free, even if other outline attrs (a
  // leftover colour/width) linger on the node.
  if (
    attrs.outlineStyle !== "none" &&
    ((attrs.outlineWidth !== undefined && attrs.outlineWidth > 0) ||
      attrs.outlineColor ||
      attrs.outlineStyle ||
      attrs.outlineCap ||
      attrs.outlineJoin ||
      attrs.outlineHeadEnd ||
      attrs.outlineTailEnd)
  ) {
    const shapeOutline: ShapeOutline = {};
    if (attrs.outlineWidth !== undefined && attrs.outlineWidth > 0) {
      shapeOutline.width = emuFromPixels(
        attrs.outlineWidth,
        "outlineWidth",
        authoredEmu,
        emuToStrokePixels,
      );
    }
    if (attrs.outlineStyle) {
      const dash = normalizeShapeOutlineDash(attrs.outlineStyle);
      if (dash !== undefined) {
        shapeOutline.dash = dash;
      }
    }
    if (attrs.outlineCap) {
      shapeOutline.cap = attrs.outlineCap;
    }
    if (attrs.outlineJoin) {
      shapeOutline.join = attrs.outlineJoin;
    }
    if (attrs.outlineHeadEnd) {
      shapeOutline.headEnd = attrs.outlineHeadEnd;
    }
    if (attrs.outlineTailEnd) {
      shapeOutline.tailEnd = attrs.outlineTailEnd;
    }
    if (attrs.outlineColor) {
      shapeOutline.color = attrs.outlineColorValue ?? {
        rgb: attrs.outlineColor.replace("#", ""),
      };
    }
    shape.outline = shapeOutline;
  }

  const shapeContent: ShapeContent = { type: "shape", shape };

  return {
    type: "run",
    ...carriedRunFormatting(attrs._docxRunFormatting),
    content: [shapeContent],
  };
}

/**
 * Convert ProseMirror marks to TextFormatting
 */
type MarksToTextFormattingOptions = {
  baseParagraphFormatting?: TextFormatting | undefined;
  inheritedFormatting?: TextFormatting | undefined;
  paragraphMarkFormatting?: TextFormatting | undefined;
  paragraphMarkPrecedesStyle?: boolean | undefined;
  styleResolver?: RunStyleResolver | null | undefined;
};

const RUN_FORMATTING_VISUAL_GROUPS = {
  bold: "bold",
  boldCs: null,
  italic: "italic",
  italicCs: null,
  underline: "underline",
  strike: "strike",
  doubleStrike: "strike",
  vertAlign: "vertAlign",
  smallCaps: "smallCaps",
  allCaps: "allCaps",
  hidden: "hidden",
  noProof: null,
  color: "color",
  highlight: "highlight",
  shading: "shading",
  fontSize: "fontSize",
  fontSizeCs: null,
  fontFamily: "fontFamily",
  language: "language",
  spacing: "characterSpacing",
  position: "characterSpacing",
  scale: "characterSpacing",
  kerning: "characterSpacing",
  effect: "effect",
  emphasisMark: "emphasisMark",
  emboss: "emboss",
  imprint: "imprint",
  outline: "outline",
  shadow: "shadow",
  rtl: "rtl",
  cs: null,
  styleId: null,
  preserved: null,
} as const satisfies Record<keyof TextFormatting, string | null>;

type VisualFormattingGroup = Exclude<
  (typeof RUN_FORMATTING_VISUAL_GROUPS)[keyof typeof RUN_FORMATTING_VISUAL_GROUPS],
  null
>;

const RUN_FORMATTING_FAST_PATH_DISPOSITION = {
  bold: "visual",
  boldCs: "structural",
  italic: "visual",
  italicCs: "structural",
  underline: "visual",
  strike: "visual",
  doubleStrike: "visual",
  vertAlign: "visual",
  smallCaps: "visual",
  allCaps: "visual",
  hidden: "visual",
  noProof: "structural",
  color: "visual",
  highlight: "visual",
  shading: "visual",
  fontSize: "visual",
  fontSizeCs: "structural",
  fontFamily: "visual",
  language: "visual",
  spacing: "visual",
  position: "visual",
  scale: "visual",
  kerning: "visual",
  effect: "visual",
  emphasisMark: "visual",
  emboss: "visual",
  imprint: "visual",
  outline: "visual",
  shadow: "visual",
  rtl: "visual",
  cs: "structural",
  styleId: "character-style",
  // Bytes, not a visual: a run holding one has formatting no mark carries, so
  // it must not take the fast path that rebuilds the run from marks alone.
  preserved: "structural",
} as const satisfies Record<keyof TextFormatting, "character-style" | "structural" | "visual">;

const hasOnlyCarrierlessVisualFormatting = (formatting: TextFormatting): boolean => {
  for (const property of Object.keys(formatting) as (keyof TextFormatting)[]) {
    if (RUN_FORMATTING_FAST_PATH_DISPOSITION[property] !== "visual") {
      return false;
    }
  }
  return true;
};

const visibleUnderline = (
  formatting: TextFormatting | undefined,
): TextFormatting["underline"] | undefined => {
  const underline = formatting?.underline;
  return underline?.style === "none" ? undefined : underline;
};

const visibleStrike = (formatting: TextFormatting | undefined): "double" | "single" | undefined => {
  if (formatting?.doubleStrike === true) {
    return "double";
  }
  return formatting?.strike === true ? "single" : undefined;
};

const visibleColor = (
  formatting: TextFormatting | undefined,
): TextFormatting["color"] | undefined =>
  formatting?.color?.auto === true ? undefined : formatting?.color;

const visibleValue = <Value>(value: Value | "none" | undefined): Value | undefined =>
  value === "none" ? undefined : value;

/**
 * Compare the exact presentation encoded by PM run marks without allocating the
 * intermediate maps used by reconciliation. Non-visual authored state is
 * deliberately absent here: the carrierless fast path below excludes every
 * structural carrier before calling this predicate.
 */
const sameVisualFormatting = (
  left: TextFormatting | undefined,
  right: TextFormatting | undefined,
): boolean => {
  if (
    (left?.bold === true) !== (right?.bold === true) ||
    (left?.italic === true) !== (right?.italic === true) ||
    !sameFormattingValue(visibleUnderline(left), visibleUnderline(right)) ||
    visibleStrike(left) !== visibleStrike(right) ||
    !sameFormattingValue(visibleColor(left), visibleColor(right)) ||
    visibleValue(left?.highlight) !== visibleValue(right?.highlight) ||
    (left?.fontSize || undefined) !== (right?.fontSize || undefined) ||
    !sameFormattingValue(left?.fontFamily, right?.fontFamily) ||
    !sameFormattingValue(left?.language, right?.language) ||
    (left?.vertAlign === "superscript" || left?.vertAlign === "subscript"
      ? left.vertAlign
      : undefined) !==
      (right?.vertAlign === "superscript" || right?.vertAlign === "subscript"
        ? right.vertAlign
        : undefined) ||
    (left?.allCaps === true) !== (right?.allCaps === true) ||
    (left?.smallCaps === true) !== (right?.smallCaps === true) ||
    (left?.hidden === true) !== (right?.hidden === true) ||
    (left?.emboss === true) !== (right?.emboss === true) ||
    (left?.imprint === true) !== (right?.imprint === true) ||
    (left?.shadow === true) !== (right?.shadow === true) ||
    (left?.outline === true) !== (right?.outline === true) ||
    (left?.rtl === true) !== (right?.rtl === true) ||
    (typeof left?.spacing === "number" ? left.spacing : null) !==
      (typeof right?.spacing === "number" ? right.spacing : null) ||
    (typeof left?.position === "number" ? left.position : null) !==
      (typeof right?.position === "number" ? right.position : null) ||
    (normalizeHorizontalScalePercent(left?.scale) ?? null) !==
      (normalizeHorizontalScalePercent(right?.scale) ?? null) ||
    (typeof left?.kerning === "number" ? left.kerning : null) !==
      (typeof right?.kerning === "number" ? right.kerning : null) ||
    visibleValue(left?.effect) !== visibleValue(right?.effect) ||
    visibleValue(left?.emphasisMark) !== visibleValue(right?.emphasisMark)
  ) {
    return false;
  }

  return sameFormattingValue(
    shadingToRunShadingAttrs(left?.shading),
    shadingToRunShadingAttrs(right?.shading),
  );
};

const visualFormattingGroups = (
  formatting: TextFormatting | undefined,
): ReadonlyMap<VisualFormattingGroup, unknown> => {
  const groups = new Map<VisualFormattingGroup, unknown>();
  if (!formatting) {
    return groups;
  }
  if (formatting.bold) {
    groups.set("bold", true);
  }
  if (formatting.italic) {
    groups.set("italic", true);
  }
  if (formatting.underline && formatting.underline.style !== "none") {
    groups.set("underline", {
      style: formatting.underline.style || "single",
      ...(formatting.underline.color ? { color: formatting.underline.color } : {}),
    });
  }
  if (formatting.strike || formatting.doubleStrike) {
    groups.set("strike", formatting.doubleStrike ? "double" : "single");
  }
  if (formatting.color && !formatting.color.auto) {
    const { rgb, themeColor, themeTint, themeShade } = formatting.color;
    groups.set("color", {
      ...(rgb ? { rgb } : {}),
      ...(themeColor ? { themeColor } : {}),
      ...(themeTint ? { themeTint } : {}),
      ...(themeShade ? { themeShade } : {}),
    });
  }
  if (formatting.highlight && formatting.highlight !== "none") {
    groups.set("highlight", formatting.highlight);
  }
  const shadingAttrs = shadingToRunShadingAttrs(formatting.shading);
  if (shadingAttrs) {
    groups.set("shading", runShadingAttrsToShading(shadingAttrs));
  }
  if (formatting.fontSize) {
    groups.set("fontSize", formatting.fontSize);
  }
  if (formatting.fontFamily) {
    const { ascii, hAnsi, eastAsia, cs, hint, asciiTheme, hAnsiTheme, eastAsiaTheme, csTheme } =
      formatting.fontFamily;
    groups.set("fontFamily", {
      ...(ascii ? { ascii } : {}),
      ...(hAnsi ? { hAnsi } : {}),
      ...(eastAsia ? { eastAsia } : {}),
      ...(cs ? { cs } : {}),
      ...(hint ? { hint } : {}),
      ...(asciiTheme ? { asciiTheme } : {}),
      ...(hAnsiTheme ? { hAnsiTheme } : {}),
      ...(eastAsiaTheme ? { eastAsiaTheme } : {}),
      ...(csTheme ? { csTheme } : {}),
    });
  }
  if (formatting.language) {
    const { val, eastAsia, bidi } = formatting.language;
    groups.set("language", {
      ...(val ? { val } : {}),
      ...(eastAsia ? { eastAsia } : {}),
      ...(bidi ? { bidi } : {}),
    });
  }
  if (formatting.vertAlign === "superscript" || formatting.vertAlign === "subscript") {
    groups.set("vertAlign", formatting.vertAlign);
  }
  for (const [property, group] of [
    ["allCaps", "allCaps"],
    ["smallCaps", "smallCaps"],
    ["hidden", "hidden"],
    ["emboss", "emboss"],
    ["imprint", "imprint"],
    ["shadow", "shadow"],
    ["outline", "outline"],
    ["rtl", "rtl"],
  ] as const) {
    if (formatting[property]) {
      groups.set(group, true);
    }
  }
  const spacing = typeof formatting.spacing === "number" ? formatting.spacing : null;
  const position = typeof formatting.position === "number" ? formatting.position : null;
  const scale = normalizeHorizontalScalePercent(formatting.scale) ?? null;
  const kerning = typeof formatting.kerning === "number" ? formatting.kerning : null;
  if (spacing !== null || position !== null || scale !== null || kerning !== null) {
    groups.set("characterSpacing", { spacing, position, scale, kerning });
  }
  if (formatting.effect && formatting.effect !== "none") {
    groups.set("effect", formatting.effect);
  }
  if (formatting.emphasisMark && formatting.emphasisMark !== "none") {
    groups.set("emphasisMark", formatting.emphasisMark);
  }
  return groups;
};

const expectedRunFormattingOverrideAttrs = (
  authoredFormatting: TextFormatting,
): RunFormattingOverrideAttrs | undefined => {
  const attrs = buildRunFormattingOverrideAttrs(authoredFormatting, {
    type: "authored-baseline",
    formatting: authoredFormatting,
  });
  const hasDirectFormatting = Object.keys(authoredFormatting).some(
    (property) => property !== "styleId",
  );
  if (!attrs && !hasDirectFormatting) {
    return undefined;
  }

  const expected: RunFormattingOverrideAttrs = { ...attrs };
  const directFontProperties = (["color", "fontFamily", "fontSize"] as const).filter(
    (property) => authoredFormatting[property] !== undefined,
  );
  if (directFontProperties.length > 0) {
    expected.directFontProperties = directFontProperties;
  }
  const complexScriptPropertyAbsences = (
    [
      ["bold", "boldCs"],
      ["italic", "italicCs"],
      ["fontSize", "fontSizeCs"],
    ] as const
  )
    .filter(
      ([ordinary, complex]) =>
        authoredFormatting[ordinary] !== undefined && authoredFormatting[complex] === undefined,
    )
    .map(([, complex]) => complex);
  if (complexScriptPropertyAbsences.length > 0) {
    expected.complexScriptPropertyAbsences = complexScriptPropertyAbsences;
  }
  return Object.keys(expected).length > 0 ? expected : undefined;
};

const changedVisualFormattingGroups = (
  observedFormatting: TextFormatting,
  expectedFormatting: TextFormatting | undefined,
): ReadonlySet<VisualFormattingGroup> => {
  const actual = visualFormattingGroups(observedFormatting);
  const expected = visualFormattingGroups(expectedFormatting);
  const changed = new Set<VisualFormattingGroup>();
  for (const group of new Set([...actual.keys(), ...expected.keys()])) {
    const actualValue = actual.get(group);
    const expectedValue = expected.get(group);
    if (!sameFormattingValue(actualValue, expectedValue)) {
      changed.add(group);
    }
  }
  return changed;
};

const VISUAL_BOOLEAN_FORMATTING_PROPERTIES = new Set<keyof TextFormatting>([
  "allCaps",
  "bold",
  "doubleStrike",
  "emboss",
  "hidden",
  "imprint",
  "italic",
  "outline",
  "rtl",
  "shadow",
  "smallCaps",
  "strike",
]);

const fontFamilyDifference = (
  actual: NonNullable<TextFormatting["fontFamily"]>,
  inherited: TextFormatting["fontFamily"],
): NonNullable<TextFormatting["fontFamily"]> | undefined => {
  const difference: NonNullable<TextFormatting["fontFamily"]> = {};
  for (const property of [
    "ascii",
    "hAnsi",
    "eastAsia",
    "cs",
    "hint",
    "asciiTheme",
    "hAnsiTheme",
    "eastAsiaTheme",
    "csTheme",
  ] as const) {
    const value = actual[property];
    if (value !== undefined && value !== inherited?.[property]) {
      Reflect.set(difference, property, value);
    }
  }
  return Object.keys(difference).length > 0 ? difference : undefined;
};

type ReconcileAuthoredFormattingOptions = {
  authoredFormatting: TextFormatting;
  carrierlessContext: "ordinary" | "paragraph-mark";
  currentOverrideAttrs: RunFormattingOverrideAttrs | undefined;
  inheritedFormatting: TextFormatting | undefined;
  observedFormatting: TextFormatting;
};

const DIRECT_OVERRIDE_FORMATTING_PROPERTIES = [
  "allCaps",
  "bold",
  "boldCs",
  "cs",
  "doubleStrike",
  "emboss",
  "fontSizeCs",
  "hidden",
  "noProof",
  "imprint",
  "italic",
  "italicCs",
  "outline",
  "rtl",
  "shadow",
  "smallCaps",
  "strike",
  "underline",
] as const satisfies readonly (keyof TextFormatting & keyof RunFormattingOverrideAttrs)[];

const POSITIVE_OVERRIDE_PROPERTIES_REPRESENTED_BY_VISUAL_MARKS = new Set<
  (typeof DIRECT_OVERRIDE_FORMATTING_PROPERTIES)[number]
>(["allCaps", "emboss", "hidden", "imprint", "outline", "rtl", "shadow", "smallCaps", "strike"]);

const setFormattingFromOverrideSignal = (
  formatting: TextFormatting,
  property: (typeof DIRECT_OVERRIDE_FORMATTING_PROPERTIES)[number],
  value: RunFormattingOverrideAttrs[typeof property],
): void => {
  Reflect.deleteProperty(formatting, property);
  if (value === undefined) {
    return;
  }
  if (property === "underline") {
    formatting.underline = { style: "none" };
    return;
  }
  Reflect.set(formatting, property, value);
};

const sameFormattingValue = (left: unknown, right: unknown): boolean => {
  if (left === right) {
    return true;
  }
  if (typeof left !== "object" || left === null || typeof right !== "object" || right === null) {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    return left.every((value, index) => sameFormattingValue(value, right[index]));
  }
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  if (leftKeys.length !== rightKeys.length) {
    return false;
  }
  return leftKeys.every(
    (key) =>
      Object.prototype.hasOwnProperty.call(right, key) &&
      sameFormattingValue(Reflect.get(left, key), Reflect.get(right, key)),
  );
};

const reconcileOverrideSignals = (
  formatting: TextFormatting,
  current: RunFormattingOverrideAttrs | undefined,
  expected: RunFormattingOverrideAttrs | undefined,
  inheritedFormatting: TextFormatting | undefined,
  observedFormatting: TextFormatting,
): void => {
  for (const property of DIRECT_OVERRIDE_FORMATTING_PROPERTIES) {
    if (sameFormattingValue(current?.[property], expected?.[property])) {
      continue;
    }
    if (
      current?.[property] === undefined &&
      expected?.[property] === true &&
      POSITIVE_OVERRIDE_PROPERTIES_REPRESENTED_BY_VISUAL_MARKS.has(property)
    ) {
      continue;
    }
    if (
      current &&
      hasAuthoredRunFormattingProvenance(current) &&
      expected?.[property] === undefined &&
      sameFormattingValue(current[property], inheritedFormatting?.[property])
    ) {
      continue;
    }
    setFormattingFromOverrideSignal(formatting, property, current?.[property]);
  }

  for (const property of ["color", "fontFamily", "fontSize"] as const) {
    const currentIsDirect = current?.directFontProperties?.includes(property) === true;
    const expectedIsDirect = expected?.directFontProperties?.includes(property) === true;
    if (currentIsDirect === expectedIsDirect) {
      continue;
    }
    if (!currentIsDirect) {
      Reflect.deleteProperty(formatting, property);
      continue;
    }
    const observed = observedFormatting[property];
    if (observed !== undefined) {
      Reflect.set(formatting, property, observed);
    }
  }

  for (const property of ["boldCs", "italicCs", "fontSizeCs"] as const) {
    const currentIsAbsent = current?.complexScriptPropertyAbsences?.includes(property) === true;
    const expectedIsAbsent = expected?.complexScriptPropertyAbsences?.includes(property) === true;
    if (currentIsAbsent !== expectedIsAbsent && currentIsAbsent) {
      Reflect.deleteProperty(formatting, property);
    }
  }
};

/**
 * Reconcile source-authored provenance with the current visual mark set at the
 * one save boundary shared by editor commands, raw transactions and compare.
 * Unchanged groups keep exact source state (including equal direct values and
 * otherwise invisible sentinels); changed groups derive the smallest direct
 * override needed relative to the current style cascade.
 */
const reconcileAuthoredFormatting = ({
  authoredFormatting,
  carrierlessContext,
  currentOverrideAttrs,
  inheritedFormatting,
  observedFormatting,
}: ReconcileAuthoredFormattingOptions): TextFormatting => {
  const hasAuthoredFormatting = Object.keys(authoredFormatting).length > 0;
  if (
    carrierlessContext === "ordinary" &&
    currentOverrideAttrs === undefined &&
    !hasAuthoredFormatting &&
    hasOnlyCarrierlessVisualFormatting(observedFormatting)
  ) {
    if (sameVisualFormatting(observedFormatting, inheritedFormatting)) {
      return authoredFormatting;
    }
    if (
      observedFormatting.fontFamily === undefined &&
      sameVisualFormatting(inheritedFormatting, undefined)
    ) {
      return observedFormatting;
    }
  }
  const expectedFormatting = hasAuthoredFormatting
    ? mergeTextFormatting(inheritedFormatting, authoredFormatting)
    : inheritedFormatting;
  const changedGroups = changedVisualFormattingGroups(observedFormatting, expectedFormatting);
  const reconciled: TextFormatting = { ...authoredFormatting };
  for (const property of Object.keys(RUN_FORMATTING_VISUAL_GROUPS) as Array<keyof TextFormatting>) {
    const group = RUN_FORMATTING_VISUAL_GROUPS[property];
    if (group === null || !changedGroups.has(group)) {
      continue;
    }
    Reflect.deleteProperty(reconciled, property);
    if (VISUAL_BOOLEAN_FORMATTING_PROPERTIES.has(property)) {
      const observed = observedFormatting[property] === true;
      const inherited = inheritedFormatting?.[property] === true;
      if (observed !== inherited) {
        Reflect.set(reconciled, property, observed);
      }
      continue;
    }

    const observed = observedFormatting[property];
    if (observed === undefined) {
      continue;
    }
    if (property === "fontFamily") {
      const difference = fontFamilyDifference(
        observed as NonNullable<TextFormatting["fontFamily"]>,
        inheritedFormatting?.fontFamily,
      );
      if (difference) {
        reconciled.fontFamily = difference;
      }
      continue;
    }
    if (!sameFormattingValue(observed, inheritedFormatting?.[property])) {
      Reflect.set(reconciled, property, observed);
    }
  }
  if (currentOverrideAttrs !== undefined) {
    reconcileOverrideSignals(
      reconciled,
      currentOverrideAttrs,
      expectedRunFormattingOverrideAttrs(authoredFormatting),
      inheritedFormatting,
      observedFormatting,
    );
  }
  return reconciled;
};

const formattingAuthorshipCost = (formatting: TextFormatting): number => {
  let cost = 0;
  const visit = (value: unknown): void => {
    if (value === undefined) {
      return;
    }
    if (value === null || typeof value !== "object") {
      cost++;
      return;
    }
    const entries = Object.entries(value);
    if (entries.length === 0) {
      cost++;
      return;
    }
    for (const [, nested] of entries) {
      visit(nested);
    }
  };
  for (const [property, value] of Object.entries(formatting)) {
    if (property !== "styleId") {
      visit(value);
    }
  }
  return cost;
};

export function marksToTextFormatting(
  marks: readonly Mark[],
  options?: MarksToTextFormattingOptions,
): TextFormatting {
  const formatting: TextFormatting = {};
  let directOverrideFormatting: TextFormatting | undefined;
  let directFontProperties: RunFormattingOverrideAttrs["directFontProperties"];
  let characterStyleId: string | undefined;
  let runFormattingOverrideMark: Mark | undefined;
  let currentOverrideAttrs: RunFormattingOverrideAttrs | undefined;

  for (const mark of marks) {
    switch (mark.type.name) {
      case "bold":
        formatting.bold = true;
        break;

      case "italic":
        formatting.italic = true;
        break;

      case "underline": {
        const attrs = expectUnderlineMarkAttrs(mark);
        const uline: NonNullable<TextFormatting["underline"]> = {
          style: attrs.style || "single",
        };
        if (attrs.color) {
          uline.color = attrs.color;
        }
        formatting.underline = uline;
        break;
      }

      case "strike":
        if (expectStrikeMarkAttrs(mark).double) {
          formatting.doubleStrike = true;
        } else {
          formatting.strike = true;
        }
        break;

      case "textColor": {
        const attrs = expectTextColorMarkAttrs(mark);
        const colorVal: ColorValue = {};
        if (attrs.rgb) {
          colorVal.rgb = attrs.rgb;
        }
        if (attrs.themeColor) {
          colorVal.themeColor = attrs.themeColor;
        }
        if (attrs.themeTint) {
          colorVal.themeTint = attrs.themeTint;
        }
        if (attrs.themeShade) {
          colorVal.themeShade = attrs.themeShade;
        }
        formatting.color = colorVal;
        break;
      }

      case "highlight":
        formatting.highlight = expectHighlightMarkAttrs(mark).color;
        break;

      case "runShading":
        // Rebuild the model `w:shd` fill so the run serializer re-emits it.
        formatting.shading = runShadingAttrsToShading(expectRunShadingMarkAttrs(mark));
        break;

      case "fontSize": {
        const attrs = expectFontSizeMarkAttrs(mark);
        formatting.fontSize = attrs.size;
        break;
      }

      case "fontFamily": {
        const attrs = expectFontFamilyMarkAttrs(mark);
        const ff: NonNullable<TextFormatting["fontFamily"]> = {};
        if (attrs.ascii) {
          ff.ascii = attrs.ascii;
        }
        if (attrs.hAnsi) {
          ff.hAnsi = attrs.hAnsi;
        }
        if (attrs.eastAsia) {
          ff.eastAsia = attrs.eastAsia;
        }
        if (attrs.hint) {
          ff.hint = attrs.hint;
        }
        if (attrs.cs) {
          ff.cs = attrs.cs;
        }
        // asciiTheme needs to be cast to the proper type
        if (attrs.asciiTheme) {
          ff.asciiTheme = attrs.asciiTheme as NonNullable<
            NonNullable<TextFormatting["fontFamily"]>["asciiTheme"]
          >;
        }
        if (attrs.hAnsiTheme) {
          ff.hAnsiTheme = attrs.hAnsiTheme;
        }
        if (attrs.eastAsiaTheme) {
          ff.eastAsiaTheme = attrs.eastAsiaTheme;
        }
        if (attrs.csTheme) {
          ff.csTheme = attrs.csTheme;
        }
        formatting.fontFamily = ff;
        break;
      }

      case "language": {
        const attrs = expectLanguageMarkAttrs(mark);
        formatting.language = {
          ...(attrs.val ? { val: attrs.val } : {}),
          ...(attrs.eastAsia ? { eastAsia: attrs.eastAsia } : {}),
          ...(attrs.bidi ? { bidi: attrs.bidi } : {}),
        };
        break;
      }

      case "superscript":
        formatting.vertAlign = "superscript";
        break;

      case "subscript":
        formatting.vertAlign = "subscript";
        break;

      case "allCaps":
        formatting.allCaps = true;
        break;

      case "smallCaps":
        formatting.smallCaps = true;
        break;

      case "characterSpacing": {
        const attrs = expectCharacterSpacingMarkAttrs(mark);
        if (attrs.spacing !== undefined) {
          formatting.spacing = attrs.spacing;
        }
        if (attrs.position !== undefined) {
          formatting.position = attrs.position;
        }
        const horizontalScale = normalizeHorizontalScalePercent(attrs.scale);
        if (horizontalScale !== undefined) {
          formatting.scale = horizontalScale;
        }
        if (attrs.kerning !== undefined) {
          formatting.kerning = attrs.kerning;
        }
        break;
      }

      case "emboss":
        formatting.emboss = true;
        break;

      case "imprint":
        formatting.imprint = true;
        break;

      case "hidden":
        // eigenpal #424 (w:vanish gap 9): mark closes the round-trip so
        // `<w:vanish/>` survives parse → PM → serialize.
        formatting.hidden = true;
        break;

      case "textShadow":
        formatting.shadow = true;
        break;

      case "emphasisMark":
        formatting.emphasisMark = expectEmphasisMarkAttrs(mark).type || "dot";
        break;

      case "textOutline":
        formatting.outline = true;
        break;

      case "rtl":
        formatting.rtl = true;
        break;

      case "textEffect":
        formatting.effect = expectTextEffectMarkAttrs(mark).effect;
        break;

      case "runFormattingOverride":
        runFormattingOverrideMark = mark;
        break;

      case "characterStyle": {
        const attrs = expectCharacterStyleMarkAttrs(mark);
        formatting.styleId = attrs.styleId;
        characterStyleId = attrs.styleId;
        break;
      }

      // hyperlink is handled separately
      default:
        break;
    }
  }

  let authoredFormatting: TextFormatting = {};
  if (runFormattingOverrideMark) {
    const overrideAttrs = expectRunFormattingOverrideMarkAttrs(runFormattingOverrideMark);
    currentOverrideAttrs = overrideAttrs;
    directFontProperties = overrideAttrs.directFontProperties;
    directOverrideFormatting = {};
    applyRunFormattingOverrideAttrs(directOverrideFormatting, overrideAttrs);
    for (const property of directFontProperties ?? []) {
      const value = formatting[property];
      if (value !== undefined) {
        Reflect.set(directOverrideFormatting, property, value);
      }
    }
    authoredFormatting = authoredRunFormattingFromAttrs(overrideAttrs) ?? directOverrideFormatting;
  }

  if (characterStyleId !== undefined && !options?.styleResolver) {
    const conservativeFormatting = mergeTextFormatting(formatting, directOverrideFormatting) ?? {};
    for (const property of currentOverrideAttrs?.complexScriptPropertyAbsences ?? []) {
      Reflect.deleteProperty(conservativeFormatting, property);
    }
    return { ...conservativeFormatting, styleId: characterStyleId };
  }

  const runContext = {
    baseParagraphFormatting: options?.baseParagraphFormatting,
    paragraphFormatting: options?.inheritedFormatting,
    paragraphMarkFormatting: options?.paragraphMarkFormatting,
    paragraphMarkPrecedesStyle: options?.paragraphMarkPrecedesStyle ?? false,
  };
  const paragraphFormatting = paragraphFormattingForRun({
    context: runContext,
    directFormatting: authoredFormatting,
    marks,
  });
  const inheritedFormatting = resolveEffectiveRunStyleFormatting({
    marks,
    paragraphFormatting,
    ...(options?.styleResolver !== undefined ? { styleResolver: options.styleResolver } : {}),
  });
  let reconciled = reconcileAuthoredFormatting({
    authoredFormatting,
    carrierlessContext:
      options?.paragraphMarkFormatting === undefined ? "ordinary" : "paragraph-mark",
    currentOverrideAttrs,
    inheritedFormatting,
    observedFormatting: formatting,
  });
  if (
    runFormattingOverrideMark === undefined &&
    characterStyleId === undefined &&
    options?.paragraphMarkFormatting !== undefined
  ) {
    const suppressedParagraphFormatting = suppressParagraphMarkFormatting({
      baseFormatting: options.baseParagraphFormatting,
      directFormatting: authoredFormatting,
      paragraphMarkFormatting: options.paragraphMarkFormatting,
      paragraphMarkPrecedesStyle: options.paragraphMarkPrecedesStyle ?? false,
    });
    const suppressedInheritedFormatting = resolveEffectiveRunStyleFormatting({
      marks,
      paragraphFormatting: suppressedParagraphFormatting,
      ...(options.styleResolver !== undefined ? { styleResolver: options.styleResolver } : {}),
    });
    const suppressed = reconcileAuthoredFormatting({
      authoredFormatting,
      carrierlessContext: "paragraph-mark",
      currentOverrideAttrs,
      inheritedFormatting: suppressedInheritedFormatting,
      observedFormatting: formatting,
    });
    if (formattingAuthorshipCost(suppressed) < formattingAuthorshipCost(reconciled)) {
      reconciled = suppressed;
    }
  }
  return {
    ...reconciled,
    ...(characterStyleId !== undefined ? { styleId: characterStyleId } : {}),
  };
}

// ============================================================================
// TABLE CONVERSION
// ============================================================================

/**
 * Convert a ProseMirror table node to our Table type
 */
/**
 * A table the editor created carries no `w:tblPr`, so the borders its cells
 * were built with are the only record of what the table looks like. Lift the
 * first bordered cell's set to the table.
 *
 * Only for such a table: a table read from a package states its own
 * `w:tblBorders` or deliberately states none, and its cells' borders may be
 * the table style's rather than their own. Inferring there invents a
 * `w:tblBorders` the source never had, and it outranks the table style on
 * reload (`formatting.borders: absent became object` in the corpus census).
 */
function inferTableBorders(rows: TableRow[]): TableBorders | undefined {
  for (const row of rows) {
    for (const cell of row.cells) {
      const borders = cell.formatting?.borders;
      if (borders) {
        const base =
          borders.top ||
          borders.left ||
          borders.right ||
          borders.bottom ||
          borders.insideH ||
          borders.insideV;
        if (!base) {
          return undefined;
        }
        return {
          top: borders.top ?? base,
          bottom: borders.bottom ?? base,
          left: borders.left ?? base,
          right: borders.right ?? base,
          insideH: borders.insideH ?? borders.bottom ?? base,
          insideV: borders.insideV ?? borders.right ?? base,
        };
      }
    }
  }
  return undefined;
}

function convertPMTable(
  node: PMNode,
  documentCounts?: TrackedChangeCounts,
  styleResolver: StyleEngine | null = null,
): Table {
  const attrs = expectTableAttrs(node);
  const rows = convertPMTableRows(node, documentCounts, styleResolver);

  const formatting = tableAttrsToFormatting(attrs) || undefined;
  if (!formatting?.borders && !attrs._originalFormatting) {
    const inferredBorders = inferTableBorders(rows);
    if (inferredBorders) {
      if (formatting) {
        formatting.borders = inferredBorders;
      } else {
        // No other formatting — create a minimal formatting object with borders
        // so borders persist on round-trip.
        const minTable: Table = {
          type: "table",
          formatting: { borders: inferredBorders },
          rows,
        };
        if (attrs.columnWidths) {
          minTable.columnWidths = attrs.columnWidths;
        }
        restoreTablePropertyChanges(minTable, attrs);
        return minTable;
      }
    }
  }

  const table: Table = { type: "table", rows };
  if (attrs.columnWidths) {
    table.columnWidths = attrs.columnWidths;
  }
  // An empty `w:tblPr` says nothing, and a table imported without one carries
  // `_originalFormatting` as an empty record to mark it as imported, so the
  // model must not gain a formatting object the document never had.
  if (formatting && Object.keys(formatting).length > 0) {
    table.formatting = formatting;
  }
  restoreTablePropertyChanges(table, attrs);
  return table;
}

/**
 * Restore `w:tblPrChange` entries that PM carried opaquely on the table attrs
 * (same rationale as the paragraph `_propertyChanges` attr): they must survive
 * an edit so the saved DOCX keeps the tracked property-change history.
 */
function restoreTablePropertyChanges(table: Table, attrs: TableAttrs): void {
  if (Array.isArray(attrs.tblPrChange) && attrs.tblPrChange.length > 0) {
    table.propertyChanges = [...attrs.tblPrChange];
  }
  // The markers the `w:tbl` held beside its rows, by reference so a table the
  // editor copied does not claim them too.
  if (attrs._bookmarks && attrs._bookmarks.length > 0) {
    table.bookmarks = attrs._bookmarks;
  }
}

type ActiveVerticalMerge = {
  remainingRows: number;
  colspan: number;
  continuationCells?: TableCell[];
};

function convertPMTableRows(
  node: PMNode,
  documentCounts?: TrackedChangeCounts,
  styleResolver: StyleEngine | null = null,
): TableRow[] {
  const rows: TableRow[] = [];
  const activeVerticalMerges = new Map<number, ActiveVerticalMerge>();

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((rowNode) => {
    if (rowNode.type.name === "tableRow") {
      rows.push(convertPMTableRow(rowNode, documentCounts, activeVerticalMerges, styleResolver));
    }
  });

  return rows;
}

/**
 * Build CellMargins from PM margin attrs (top/bottom/left/right as number|null|undefined)
 */
function buildCellMarginsFromAttrs(m: {
  top?: number | null;
  bottom?: number | null;
  left?: number | null;
  right?: number | null;
}): CellMargins {
  const margins: CellMargins = {};
  if (m.top !== null && m.top !== undefined) {
    margins.top = { value: m.top, type: "dxa" };
  }
  if (m.bottom !== null && m.bottom !== undefined) {
    margins.bottom = { value: m.bottom, type: "dxa" };
  }
  if (m.left !== null && m.left !== undefined) {
    margins.left = { value: m.left, type: "dxa" };
  }
  if (m.right !== null && m.right !== undefined) {
    margins.right = { value: m.right, type: "dxa" };
  }
  return margins;
}

/**
 * Convert ProseMirror table attrs to TableFormatting
 */
/**
 * Whether a live attr still holds what the style cascade resolved to.
 *
 * `toProseDoc` seeds an attr with the EFFECTIVE value — a border a table
 * style supplied, a margin a table declared — because that is what the editor
 * renders with. A save must write only what the node itself states, or the
 * inherited value becomes a direct override that outranks the style it came
 * from on reload. The resolved companion is what tells the two apart.
 */
const sameResolvedValue = (live: unknown, resolved: unknown): boolean =>
  resolved !== undefined && resolved !== null && sameFormattingValue(live, resolved);

export function tableAttrsToFormatting(attrs: TableAttrs): TableFormatting | undefined {
  // If we have the original formatting from the DOCX, use it as a base
  // for lossless round-trip. This preserves properties like cellSpacing,
  // indent, layout, bidi, overlap, shading that aren't tracked as PM attrs.
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };

    // Override properties that user may have changed via editor commands
    if (attrs.styleId !== (orig.styleId ?? undefined)) {
      if (attrs.styleId) {
        result.styleId = attrs.styleId;
      } else {
        delete result.styleId;
      }
    }
    if (attrs.justification !== (orig.justification ?? undefined)) {
      if (attrs.justification) {
        result.justification = attrs.justification;
      } else {
        delete result.justification;
      }
    }
    if (attrs.floating !== (orig.floating ?? undefined)) {
      if (attrs.floating) {
        result.floating = attrs.floating;
      } else {
        delete result.floating;
      }
    }
    if (attrs.look !== (orig.look ?? undefined)) {
      if (attrs.look) {
        result.look = attrs.look;
      } else {
        delete result.look;
      }
    }
    // Borders: toProseDoc seeds attrs.borders with the same reference as
    // orig.borders, so a difference means a border command replaced them.
    if (attrs.borders !== (orig.borders ?? undefined)) {
      if (attrs.borders) {
        result.borders = attrs.borders;
      } else {
        delete result.borders;
      }
    }
    // Width: check if changed
    const tableWidth = attrs.width;
    const tableWidthType = attrs.widthType;
    const origWidthVal = orig.width?.value;
    const origWidthType = orig.width?.type;
    if (tableWidth !== origWidthVal || tableWidthType !== origWidthType) {
      if (tableWidth !== undefined || tableWidthType !== undefined) {
        result.width = {
          value: tableWidth ?? 0,
          type: tableWidthType ?? "dxa",
        };
      } else {
        delete result.width;
      }
    }
    // Only what the table states: `cellMargins` also carries what the table
    // style resolved to, and writing that into `w:tblCellMar` would turn a
    // style's default into the table's own override.
    if (attrs.cellMargins && !sameResolvedValue(attrs.cellMargins, attrs._resolvedCellMargins)) {
      result.cellMargins = buildCellMarginsFromAttrs(attrs.cellMargins);
    }

    return result;
  }

  // Fallback: reconstruct formatting from individual attrs (e.g. for
  // newly created tables that don't have _originalFormatting)
  const tableWidth = attrs.width;
  const tableWidthType = attrs.widthType;
  const hasFormatting =
    attrs.styleId ||
    tableWidth !== undefined ||
    tableWidthType !== undefined ||
    attrs.justification ||
    attrs.floating ||
    attrs.cellMargins ||
    attrs.look ||
    attrs.borders;

  if (!hasFormatting) {
    return undefined;
  }

  // Convert cellMargins back to CellMargins format (twips → TableMeasurement)
  const cellMargins = attrs.cellMargins ? buildCellMarginsFromAttrs(attrs.cellMargins) : undefined;

  // Restore width — handle width=0 with type="auto" (common OOXML pattern)
  let width: TableFormatting["width"];
  if (tableWidth !== undefined || tableWidthType !== undefined) {
    width = {
      value: tableWidth ?? 0,
      type: tableWidthType ?? "dxa",
    };
  }

  const f: TableFormatting = {};
  if (attrs.styleId) {
    f.styleId = attrs.styleId;
  }
  if (width) {
    f.width = width;
  }
  if (attrs.justification) {
    f.justification = attrs.justification;
  }
  if (attrs.floating) {
    f.floating = attrs.floating;
  }
  if (cellMargins) {
    f.cellMargins = cellMargins;
  }
  if (attrs.look) {
    f.look = attrs.look;
  }
  if (attrs.borders) {
    f.borders = attrs.borders;
  }
  return f;
}

/**
 * Convert a ProseMirror table row node to our TableRow type
 */
function convertPMTableRow(
  node: PMNode,
  documentCounts?: TrackedChangeCounts,
  activeVerticalMerges?: Map<number, ActiveVerticalMerge>,
  styleResolver: StyleEngine | null = null,
): TableRow {
  const attrs = expectTableRowAttrs(node);
  const cells: TableCell[] = [];
  let gridColumn = 0;

  const appendActiveVerticalMerges = (): void => {
    if (!activeVerticalMerges) {
      return;
    }

    let activeMerge = activeVerticalMerges.get(gridColumn);
    while (activeMerge) {
      const preservedCell = activeMerge.continuationCells?.shift();
      cells.push(preservedCell ?? createVerticalMergeContinuationCell(activeMerge.colspan));
      activeMerge.remainingRows -= 1;
      if (activeMerge.remainingRows <= 0) {
        activeVerticalMerges.delete(gridColumn);
      }
      gridColumn += activeMerge.colspan;
      activeMerge = activeVerticalMerges.get(gridColumn);
    }
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((cellNode) => {
    if (cellNode.type.name === "tableCell" || cellNode.type.name === "tableHeader") {
      const cellAttrs = expectTableCellAttrs(cellNode);
      const colspan = Math.max(cellAttrs.colspan, 1);
      if (cellAttrs._omittedGridSlot) {
        // `w:gridBefore` and `w:gridAfter` describe absent grid positions,
        // not authored cells. The PM-only placeholder keeps TableMap
        // rectangular, while the row formatting carries the OOXML semantics.
        const placeholderIsEmpty =
          cellNode.childCount === 1 &&
          cellNode.child(0).type.name === "paragraph" &&
          cellNode.child(0).content.size === 0;
        if (!placeholderIsEmpty) {
          panic("An omitted table grid slot acquired authored content");
        }
        gridColumn += colspan;
        return;
      }
      appendActiveVerticalMerges();
      cells.push(convertPMTableCell(cellNode, documentCounts, styleResolver));
      if (cellAttrs.rowspan > 1) {
        const continuationCells =
          cellAttrs._docxVMergeContinuationCells !== undefined &&
          cellAttrs._docxVMergeContinuationCells !== null
            ? restoreTableCellsWithParagraphPropertySources(
                decodeTableCellParagraphSourcePayload(cellAttrs._docxVMergeContinuationCells),
              )
            : undefined;
        activeVerticalMerges?.set(gridColumn, {
          remainingRows: cellAttrs.rowspan - 1,
          colspan,
          ...(continuationCells ? { continuationCells: [...continuationCells] } : {}),
        });
      }
      gridColumn += colspan;
    }
  });
  appendActiveVerticalMerges();

  const row: TableRow = { type: "tableRow", cells };
  const rowFormatting = tableRowAttrsToFormatting(attrs);
  if (rowFormatting) {
    row.formatting = rowFormatting;
  }
  // The table properties the row overrides, carried whole through PM.
  if (attrs._tablePropertyExceptions) {
    row.tablePropertyExceptions = attrs._tablePropertyExceptions;
  }
  if (Array.isArray(attrs.tblPrExChange) && attrs.tblPrExChange.length > 0) {
    row.tablePropertyExceptionChanges = [...attrs.tblPrExChange];
  }
  // Restore `w:trPrChange` entries PM carried opaquely (see the paragraph
  // `_propertyChanges` attr for the rationale).
  if (Array.isArray(attrs.trPrChange) && attrs.trPrChange.length > 0) {
    row.propertyChanges = [...attrs.trPrChange];
  }
  if (attrs._preservedAttributes && attrs._preservedAttributes.length > 0) {
    row.preservedAttributes = attrs._preservedAttributes;
  }
  if (attrs._bookmarks && attrs._bookmarks.length > 0) {
    row.bookmarks = attrs._bookmarks;
  }
  if (attrs.trIns) {
    row.structuralChange = {
      type: "tableRowInsertion",
      info: {
        id: attrs.trIns.revisionId,
        author: attrs.trIns.author,
        ...(attrs.trIns.date != null && { date: attrs.trIns.date }),
        ...(attrs.trIns.utcDate != null && {
          utcDate: { attribute: DATE_UTC_ATTRIBUTE, value: attrs.trIns.utcDate },
        }),
        ...(attrs.trIns.initials != null && { initials: attrs.trIns.initials }),
      },
    };
  } else if (attrs.trDel) {
    row.structuralChange = {
      type: "tableRowDeletion",
      info: {
        id: attrs.trDel.revisionId,
        author: attrs.trDel.author,
        ...(attrs.trDel.date != null && { date: attrs.trDel.date }),
        ...(attrs.trDel.utcDate != null && {
          utcDate: { attribute: DATE_UTC_ATTRIBUTE, value: attrs.trDel.utcDate },
        }),
        ...(attrs.trDel.initials != null && { initials: attrs.trDel.initials }),
      },
    };
  }
  return row;
}

function createVerticalMergeContinuationCell(colspan: number): TableCell {
  const formatting: TableCellFormatting = { vMerge: "continue" };
  if (colspan > 1) {
    formatting.gridSpan = colspan;
  }
  return {
    type: "tableCell",
    content: [{ type: "paragraph", content: [] }],
    formatting,
  };
}

/**
 * Convert ProseMirror table row attrs to TableRowFormatting
 */
export function tableRowAttrsToFormatting(attrs: TableRowAttrs): TableRowFormatting | undefined {
  // If we have the original formatting from the DOCX, use it as a base
  // for lossless round-trip. This preserves properties like cantSplit,
  // justification, hidden, conditionalFormat that aren't tracked as PM attrs.
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };

    // Override properties that user may have changed via editor commands
    if (attrs.height !== (orig.height?.value ?? undefined)) {
      if (attrs.height) {
        result.height = { value: attrs.height, type: "dxa" as const };
      } else {
        delete result.height;
      }
    }
    if (attrs.heightRule !== (orig.heightRule ?? undefined)) {
      if (attrs.heightRule) {
        result.heightRule = attrs.heightRule;
      } else {
        delete result.heightRule;
      }
    }
    if (attrs.isHeader !== (orig.header ?? undefined)) {
      if (attrs.isHeader) {
        result.header = attrs.isHeader;
      } else {
        delete result.header;
      }
    }
    if (attrs.hidden !== orig.hidden) {
      if (attrs.hidden === undefined) {
        delete result.hidden;
      } else {
        result.hidden = attrs.hidden;
      }
    }

    return result;
  }

  // Fallback: reconstruct formatting from individual attrs
  const hasFormatting = attrs.height || attrs.isHeader || attrs.hidden !== undefined;

  if (!hasFormatting) {
    return undefined;
  }

  const f: TableRowFormatting = {};
  if (attrs.height) {
    f.height = { value: attrs.height, type: "dxa" };
  }
  if (attrs.heightRule) {
    f.heightRule = attrs.heightRule;
  }
  if (attrs.isHeader) {
    f.header = attrs.isHeader;
  }
  if (attrs.hidden !== undefined) {
    f.hidden = attrs.hidden;
  }
  return f;
}

/**
 * Convert a ProseMirror table cell node to our TableCell type
 */
function convertPMTableCell(
  node: PMNode,
  documentCounts?: TrackedChangeCounts,
  styleResolver: StyleEngine | null = null,
): TableCell {
  const attrs = expectTableCellAttrs(node);
  const content: TableCellBlock[] = [];
  const textBoxAnchorMarkers = new Map<string, Run>();
  let previousStandaloneTextBox: PreviousStandaloneTextBox | null = null;

  // Extract cell content (paragraphs and nested tables)
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((contentNode) => {
    if (contentNode.type.name === "paragraph") {
      content.push(
        convertPMParagraph(contentNode, documentCounts, textBoxAnchorMarkers, styleResolver),
      );
      previousStandaloneTextBox = null;
    } else if (contentNode.type.name === "table") {
      content.push(convertPMTable(contentNode, documentCounts, styleResolver));
      previousStandaloneTextBox = null;
    } else if (contentNode.type.name === "textBox") {
      previousStandaloneTextBox = appendTextBoxBlock(content, contentNode, {
        pendingPageBreaks: 0,
        previousStandaloneTextBox,
        textBoxAnchorMarkers,
        styleResolver,
      });
    } else if (contentNode.type.name === "preservedBlock") {
      content.push({ type: "preservedBlock", xml: expectPreservedBlockAttrs(contentNode).xml });
      previousStandaloneTextBox = null;
    } else if (contentNode.type.name === "blockBookmarkBoundary") {
      content.push(blockBookmarkMarker(contentNode));
      previousStandaloneTextBox = null;
    }
  });

  removeUnresolvedTextBoxAnchors(content, textBoxAnchorMarkers);

  const cell: TableCell = { type: "tableCell", content };
  const cellFormatting = tableCellAttrsToFormatting(attrs);
  if (cellFormatting) {
    cell.formatting = cellFormatting;
  }
  // Restore `w:tcPrChange` entries PM carried opaquely (see the paragraph
  // `_propertyChanges` attr for the rationale).
  if (Array.isArray(attrs.tcPrChange) && attrs.tcPrChange.length > 0) {
    cell.propertyChanges = [...attrs.tcPrChange];
  }
  if (attrs.cellMarker) {
    const info = {
      id: attrs.cellMarker.info.revisionId,
      author: attrs.cellMarker.info.author,
      ...(attrs.cellMarker.info.date != null && { date: attrs.cellMarker.info.date }),
      ...(attrs.cellMarker.info.utcDate != null && {
        utcDate: { attribute: DATE_UTC_ATTRIBUTE, value: attrs.cellMarker.info.utcDate },
      }),
      ...(attrs.cellMarker.info.initials != null && { initials: attrs.cellMarker.info.initials }),
    };
    if (attrs.cellMarker.kind === "merge") {
      cell.structuralChange = {
        type: "tableCellMerge",
        info,
        ...(attrs.cellMarker.verticalMerge !== undefined
          ? { verticalMerge: attrs.cellMarker.verticalMerge }
          : {}),
        ...(attrs.cellMarker.verticalMergeOriginal !== undefined
          ? { verticalMergeOriginal: attrs.cellMarker.verticalMergeOriginal }
          : {}),
      };
    } else {
      cell.structuralChange = {
        type: attrs.cellMarker.kind === "ins" ? "tableCellInsertion" : "tableCellDeletion",
        info,
      };
    }
  }
  return cell;
}

export const standaloneTableCellFromProseMirror = (node: PMNode): TableCell =>
  convertPMTableCell(node);

/**
 * Convert ProseMirror table cell attrs to TableCellFormatting
 * Borders are stored as full BorderSpec objects — no conversion needed.
 */
type CellShading = NonNullable<TableCellFormatting["shading"]>;

const cellShadingFromAttrs = (attrs: TableCellAttrs): CellShading =>
  attrs.backgroundColor ? { fill: { rgb: attrs.backgroundColor } } : { pattern: "nil" };

export function tableCellAttrsToFormatting(attrs: TableCellAttrs): TableCellFormatting | undefined {
  const backgroundChanged = attrs.backgroundColor !== attrs._resolvedBackgroundColor;

  // If we have the original formatting from the DOCX, use it as a base
  // for lossless round-trip. This preserves properties like vMerge, fitText,
  // hideMark, conditionalFormat that aren't tracked as PM attrs.
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };

    // Override properties that user may have changed via editor commands
    if (attrs.colspan > 1) {
      result.gridSpan = attrs.colspan;
    } else {
      delete result.gridSpan;
    }
    if (attrs.rowspan > 1) {
      result.vMerge = "restart";
    } else if (result.vMerge === "restart" && !attrs._preserveVMergeRestart) {
      delete result.vMerge;
    }
    // Only what the cell states: `attrs.width` also carries the width the
    // table resolved from its grid, and a merge clears an unsafe preferred
    // width, so neither `_originalFormatting.width` nor the rendered width may
    // decide whether a `w:tcW` goes back.
    const authoredWidth = attrs._authoredWidth;
    if (authoredWidth) {
      result.width = authoredWidth;
    } else {
      delete result.width;
    }
    if (attrs.verticalAlign !== (orig.verticalAlign ?? undefined)) {
      if (attrs.verticalAlign) {
        result.verticalAlign = attrs.verticalAlign;
      } else {
        delete result.verticalAlign;
      }
    }
    if (backgroundChanged) {
      result.shading = cellShadingFromAttrs(attrs);
    }
    // Only what the cell states: both attrs also carry what the table and the
    // table style resolved to, and writing those into `w:tcPr` would turn an
    // inherited value into the cell's own override.
    if (attrs.borders && !sameResolvedValue(attrs.borders, attrs._resolvedBorders)) {
      result.borders = attrs.borders;
    }
    if (attrs.margins && !sameResolvedValue(attrs.margins, attrs._resolvedMargins)) {
      result.margins = buildCellMarginsFromAttrs(attrs.margins);
    }
    if (attrs.textDirection !== (orig.textDirection ?? undefined)) {
      if (attrs.textDirection) {
        result.textDirection = attrs.textDirection;
      } else {
        delete result.textDirection;
      }
    }

    return result;
  }

  // Fallback: reconstruct formatting from individual attrs
  const authoredWidth = attrs._authoredWidth;
  const hasFormatting =
    attrs.colspan > 1 ||
    attrs.rowspan > 1 ||
    authoredWidth !== undefined ||
    attrs.verticalAlign ||
    backgroundChanged ||
    attrs.borders ||
    attrs.margins ||
    attrs.textDirection;

  if (!hasFormatting) {
    return undefined;
  }

  const f: TableCellFormatting = {};
  if (attrs.colspan > 1) {
    f.gridSpan = attrs.colspan;
  }
  if (attrs.rowspan > 1) {
    f.vMerge = "restart";
  }
  if (authoredWidth) {
    f.width = authoredWidth;
  }
  if (attrs.verticalAlign) {
    f.verticalAlign = attrs.verticalAlign;
  }
  if (attrs.textDirection) {
    f.textDirection = attrs.textDirection;
  }
  if (backgroundChanged) {
    f.shading = cellShadingFromAttrs(attrs);
  }
  if (attrs.borders) {
    f.borders = attrs.borders;
  }
  if (attrs.margins) {
    f.margins = buildCellMarginsFromAttrs(attrs.margins);
  }
  return f;
}

// ============================================================================
// TEXT BOX CONVERSION
// ============================================================================

type TextBoxBodyContentProjection =
  | { readonly type: "source-empty" }
  | { readonly type: "authored"; readonly content: (Paragraph | Table)[] };

const isUnchangedSourceEmptyPlaceholder = (blocks: readonly (Paragraph | Table)[]): boolean => {
  const paragraph = blocks.at(0);
  if (blocks.length !== 1 || paragraph?.type !== "paragraph" || paragraph.content.length !== 0) {
    return false;
  }

  return Object.keys(paragraph).every(
    (key) => key === "type" || key === "content" || key === "paraId" || key === "textId",
  );
};

const projectTextBoxBodyContent = (
  state: TextBoxAttrs["_docxTextBodyContentState"],
  blocks: (Paragraph | Table)[],
): TextBoxBodyContentProjection => {
  switch (state.type) {
    case "source-empty":
      return isUnchangedSourceEmptyPlaceholder(blocks)
        ? { type: "source-empty" }
        : { type: "authored", content: blocks };
    case "authored":
      return { type: "authored", content: blocks };
    default: {
      const exhaustive: never = state;
      return exhaustive;
    }
  }
};

/**
 * Convert a ProseMirror textBox node back to a Paragraph wrapping a ShapeContent run.
 * The text box content becomes a Shape with textBody.
 */
function convertPMTextBox(node: PMNode, styleResolver: StyleEngine | null = null): Paragraph {
  const attrs = expectTextBoxAttrs(node);
  const authoredEmu = attrs._docxAuthoredEmu;
  const verticalAlign = normalizeShapeTextAnchor(attrs.verticalAlign);

  // Extract child paragraphs from the text box content
  const childBlocks: (Paragraph | Table)[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    if (child.type.name === "paragraph") {
      childBlocks.push(convertPMParagraph(child, undefined, undefined, styleResolver));
    } else if (child.type.name === "table") {
      childBlocks.push(convertPMTable(child, undefined, styleResolver));
    }
  });
  const textBodyContent = projectTextBoxBodyContent(attrs._docxTextBodyContentState, childBlocks);

  // Build shape with text body
  const shape: Shape = {
    type: "shape",
    shapeType: "textBox",
    size: {
      width: attrs.width ? emuFromPixels(attrs.width, "width", authoredEmu, emuToPixels) : 0,
      height: attrs.height ? emuFromPixels(attrs.height, "height", authoredEmu, emuToPixels) : 0,
    },
    textBody: {
      content: textBodyContent.type === "source-empty" ? [] : textBodyContent.content,
      ...(attrs.wordArt !== undefined ? { wordArt: attrs.wordArt } : {}),
      ...(attrs.autoFit !== undefined ? { autoFit: attrs.autoFit } : {}),
      ...(attrs.textWrap !== undefined ? { textWrap: attrs.textWrap } : {}),
      ...(verticalAlign !== undefined ? { anchor: verticalAlign } : {}),
      margins: (() => {
        const m: {
          top?: number;
          bottom?: number;
          left?: number;
          right?: number;
        } = {};
        if (typeof attrs.marginTop === "number") {
          m.top = emuFromPixels(attrs.marginTop, "marginTop", authoredEmu, emuToPixels);
        }
        if (typeof attrs.marginBottom === "number") {
          m.bottom = emuFromPixels(attrs.marginBottom, "marginBottom", authoredEmu, emuToPixels);
        }
        if (typeof attrs.marginLeft === "number") {
          m.left = emuFromPixels(attrs.marginLeft, "marginLeft", authoredEmu, emuToPixels);
        }
        if (typeof attrs.marginRight === "number") {
          m.right = emuFromPixels(attrs.marginRight, "marginRight", authoredEmu, emuToPixels);
        }
        return m;
      })(),
    },
  };

  if (attrs.textBoxId) {
    shape.id = attrs.textBoxId;
  }
  // `""` is a name someone wrote, so presence is the test, not truthiness.
  if (attrs.textBoxName !== undefined) {
    shape.name = attrs.textBoxName;
  }
  if (attrs.alt !== undefined) {
    shape.alt = attrs.alt;
  }
  if (attrs.title !== undefined) {
    shape.title = attrs.title;
  }

  const transform = readAuthoredTransform(attrs);
  if (transform) {
    shape.transform = transform;
  }

  // Convert fill color back
  if (attrs.fillColor) {
    shape.fill = {
      type: "solid",
      color: { rgb: attrs.fillColor.replace("#", "") },
    };
  }

  // Convert outline back. `outlineStyle === "none"` is the explicit no-outline
  // sentinel: drop the `<a:ln>` even if a width lingers, matching the shape path.
  if (attrs.outlineStyle !== "none" && attrs.outlineWidth && attrs.outlineWidth > 0) {
    const tbOutline: ShapeOutline = {
      width: emuFromPixels(attrs.outlineWidth, "outlineWidth", authoredEmu, emuToStrokePixels),
      dash: normalizeShapeOutlineDash(attrs.outlineStyle) ?? "solid",
    };
    if (attrs.outlineColor) {
      tbOutline.color = { rgb: attrs.outlineColor.replace("#", "") };
    }
    shape.outline = tbOutline;
  }

  const wrap = textBoxWrapFromAttrs(attrs);
  if (wrap) {
    shape.wrap = wrap;
  }
  const position = imagePositionFromAttrs(attrs.position);
  if (position) {
    shape.position = position;
  }
  const anchor = restoredDrawingAnchor(attrs.anchor);
  if (anchor) {
    shape.anchor = anchor;
  }

  // Wrap the shape in a paragraph with a run containing ShapeContent
  const shapeContent: ShapeContent = { type: "shape", shape };
  const run: Run = { type: "run", content: [shapeContent] };
  const trackedChange = attrs._docxTrackedChange;
  const inlineSdts = attrs._docxInlineSdts ?? [];
  if (inlineSdts.length > 0) {
    let wrapped: InlineSdt["content"][number] = trackedChange
      ? { type: trackedChange.type, info: trackedChange.info, content: [run] }
      : run;
    for (let index = inlineSdts.length - 1; index >= 0; index -= 1) {
      const sdtAttrs = inlineSdts[index];
      if (!sdtAttrs) {
        continue;
      }
      wrapped = {
        type: "inlineSdt",
        properties: sdtPropertiesFromAttrs(sdtAttrs),
        content: [wrapped],
      };
    }
    return {
      type: "paragraph",
      content: [wrapped],
    };
  }

  return {
    type: "paragraph",
    content: trackedChange
      ? [{ type: trackedChange.type, info: trackedChange.info, content: [run] }]
      : [run],
  };
}

/**
 * Update a Document with content from a ProseMirror document
 * Preserves all non-content parts of the original document
 */
export function updateDocumentContent(originalDocument: Document, pmDoc: PMNode): Document {
  return fromProseDoc(pmDoc, originalDocument);
}

/**
 * Convert a ProseMirror document back to an array of `BlockContent` blocks
 * (paragraphs, tables, and block-level content controls).
 *
 * Used for converting edited header/footer PM content back to the document
 * model.
 */
export function proseDocToBlocks(
  pmDoc: PMNode,
  baseContent?: BlockContent[],
  styles?: NonNullable<Document["package"]>["styles"],
): BlockContent[] {
  const blocks = extractBlocks(pmDoc, "resolve", styles ? createStyleEngine(styles) : null);
  joinCommentRangesAcrossParagraphs(blocks);
  completeCommentReferences(blocks);
  const linkedSources = restoreLinkedParagraphPropertySources(blocks);
  if (baseContent) {
    restoreParagraphPropertySources(
      blocks,
      baseContent,
      linkedSources.targets,
      linkedSources.sources,
    );
  }
  return blocks;
}
