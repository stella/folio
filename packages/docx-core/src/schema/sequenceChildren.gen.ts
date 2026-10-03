/**
 * GENERATED FILE — do not edit.
 *
 * The containers whose content model declares one order for its children, with
 * the children written in it, derived from the committed schema graph by
 * `scripts/generate-container-children.ts`.
 *
 * It lives in the lower package because two packages write a property set and
 * the dependency runs one way: `@stll/docx-core` compiles a `w:rPr` from a
 * legal source and cannot import `@stll/folio-core`. folio-core's
 * declared-child table spreads this one in, so both write the same order and
 * neither restates it.
 *
 * Regenerate with:
 *
 *   bun run generate:container-children
 */

export const SEQUENCE_CHILDREN = {
  "w:numbering": ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"],
  "w:abstractNum": ["nsid", "multiLevelType", "tmpl", "name", "styleLink", "numStyleLink", "lvl"],
  "w:lvl": ["start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText", "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr"],
  "w:num": ["abstractNumId", "lvlOverride"],
  "w:lvlOverride": ["startOverride", "lvl"],
  "paragraph-properties": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
  "table-properties": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"],
  "table-property-exceptions": ["tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblPrExChange"],
  "row-properties": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"],
  "cell-properties": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"],
  "run-properties": ["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
  "content-control-properties": ["rPr", "alias", "tag", "id", "lock", "placeholder", "temporary", "showingPlcHdr", "dataBinding", "label", "tabIndex", "equation", "comboBox", "date", "docPartObj", "docPartList", "dropDownList", "picture", "richText", "text", "citation", "group", "bibliography"],
  "section-properties": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
} as const;

export const SETTINGS_CHILDREN = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge", "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"] as const;

/** A container whose declared children have an order. */
export type SequenceContainer = keyof typeof SEQUENCE_CHILDREN;

/** Every child the schema declares for `Container`, in declaration order. */
export type SequenceChild<Container extends SequenceContainer> =
  (typeof SEQUENCE_CHILDREN)[Container][number];
