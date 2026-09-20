/**
 * The functions that own a reserved value, as `"<repo-relative module>#<name>"`.
 *
 * A bare function name is not a unique key in this repository: `toProseDoc.ts`
 * and `markUtils.ts` both declare `textFormattingToMarks`. The module path
 * disambiguates, and it is also what the reserved-value lint needs: a bare
 * comparison against a sentinel is allowed inside the owning module and
 * nowhere else.
 *
 * A path here is data, not a dependency. `@stll/docx-core` declares the model
 * and its decisions; the readers live in `@stll/folio-core`, which depends on
 * this package, so the registry can only name them.
 */

const CORE = "packages/core/src";

export const RESERVED_VALUE_READERS = {
  /** `w:numId` 0 names no numbering definition. */
  numberingReference: `${CORE}/docx/numberingReference.ts#isNumberingReference`,
  /** Reads `w:numPr`, `w:outlineLvl`, `w:ind`, `w:framePr` and the pPr toggles. */
  paragraphProperties: `${CORE}/docx/paragraphParser.ts#parseParagraphProperties`,
  /** Reads `w:rPr`: `w:highlight`, `w:vertAlign`, `w:effect`, `w:em`, `w:rFonts`. */
  runProperties: `${CORE}/docx/runParser.ts#parseRunProperties`,
  /** `ST_OnOff` lexical space, including the explicit `off` that is not absence. */
  onOffValue: `${CORE}/docx/xmlParser.ts#parseOnOffValue`,
  /** Combines a toggle property across two levels of the style hierarchy by XOR. */
  toggleCascade: `${CORE}/prosemirror/styles/styleToggleCascade.ts#cascadeStyleTextFormatting`,
  /** `auto` colour; resolves it against the context default. */
  color: `${CORE}/utils/colorResolver.ts#resolveColor`,
  /** `w:highlight` `none`. */
  highlight: `${CORE}/utils/colorResolver.ts#resolveHighlightColor`,
  /** `CT_Border`: `nil` and `none` are two distinct "no border" tokens. */
  borderSpec: `${CORE}/docx/borderParser.ts#parseBorderSpec`,
  /** `w:shd` `nil` (no shading) vs `clear` (pattern-less fill), and `auto` colours. */
  shading: `${CORE}/docx/shadingParser.ts#parseShading`,
  /** `w:tblW`/`w:tcW`: resolves `dxa` and `pct` only, so `auto` and `nil` carry no number. */
  tableWidth: `${CORE}/layout-engine/types.ts#resolveTableWidthPx`,
  /** `w:gridSpan` 1/0, `w:vMerge` absent. */
  tableCellProperties: `${CORE}/docx/tableParser.ts#parseTableCellProperties`,
  /** `w:trHeight@hRule` `auto`. */
  tableRowProperties: `${CORE}/docx/tableParser.ts#parseTableRowProperties`,
  /** `w:tblLayout@type`, `w:tblOverlap@val`. */
  tableProperties: `${CORE}/docx/tableParser.ts#parseTableProperties`,
  /** `w:tblLook`: which flag a stated attribute and a `w:val` bit resolve to. */
  tableLook: `${CORE}/docx/tableLook.ts#resolveTableLook`,
  /** `w:tab@val` `clear`/`bar` and `w:tab@leader` `none`. */
  tabStops: `${CORE}/layout-engine/measure/tabCalculator.ts#computeTabStops`,
  /** `w:u@val` `none` cancels an inherited underline. */
  underline: `${CORE}/prosemirror/extensions/marks/markUtils.ts#textFormattingToMarks`,
  /** `w:outlineLvl` 9 is body text, not a tenth heading level. */
  outlineLevel: `${CORE}/utils/headingCollector.ts#collectHeadings`,
  /** `w:sectPr/w:type` absent means `nextPage`. */
  sectionBreak: `${CORE}/layout-engine/section-breaks.ts#normalizeSectionBreakType`,
  /** `w:numFmt` `none` means the level renders no marker. */
  numberingLevelMarker: `${CORE}/docx/numberingParser.ts#numberingLevelHasMarkerSlot`,
  /** `w:lvl` `w:start`, `w:lvlRestart`, `w:suff`, `w:numFmt` `custom`. */
  numbering: `${CORE}/docx/numberingParser.ts#parseNumbering`,
  /** `w:spacing@line` counts 240ths of a line under `@lineRule="auto"` and twips otherwise. */
  lineSpacing: `${CORE}/layout-bridge/convert/toFlowBlocks.ts#convertParagraphAttrs`,
  /** An unknown `w:pStyle`/`w:rStyle`/`w:tblStyle` falls back to the type's default style. */
  styleChain: `${CORE}/prosemirror/styles/styleResolver.ts#createStyleResolver`,
  /** Separator and continuation-separator notes, by `@w:type`. */
  noteType: `${CORE}/docx/footnoteParser.ts#parseFootnotes`,
  /** Run children: `w:br@clear`, `w:ptab@leader`, field-character flags. */
  runContent: `${CORE}/docx/runParser.ts#parseRun`,
  /** `wp:anchor` flags and the picture's alpha and crop. */
  drawing: `${CORE}/docx/imageParser.ts#parseDrawing`,
  /** `wp:anchor@behindDoc`, which selects between the `behind` and `inFront` wraps. */
  behindDoc: `${CORE}/docx/drawingUtils.ts#parseAnchorBehindDoc`,
  /** `a:noFill`, `a:ln w="0"`, `a:bodyPr@wrap` and the autofit kind. */
  shape: `${CORE}/docx/shapeParser.ts#parseShape`,
  /** A text box's autofit, wrap and margins. */
  textBox: `${CORE}/docx/textBoxParser.ts#parseTextBox`,
  /** Content-control flags and the dropdown's last value. */
  sdtProperties: `${CORE}/docx/sdtProperties.ts#parseSdtProperties`,
  /** `settings.xml` toggles, which fall back to Word's own defaults when absent. */
  settings: `${CORE}/docx/settingsParser.ts#parseSettings`,
  /** `w:style` flags, and the `w:default="1"` style each type falls back to. */
  styleDefinitions: `${CORE}/docx/styleParser.ts#parseStyleDefinitions`,
  /** `w:pitch`, `w:family` and `w:charset` on a font-table entry. */
  fontTable: `${CORE}/docx/fontTableParser.ts#parseFontTable`,
  /** `w15:commentEx@done` and the comment's parent. */
  comments: `${CORE}/docx/commentParser.ts#parseComments`,
  /** A header or footer reference's `@w:type`, which names a slot rather than a default. */
  headerFooterType: `${CORE}/docx/sectionParser.ts#getDefaultSectionProperties`,
} as const;
