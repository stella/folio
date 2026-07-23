/**
 * Paragraph Serializer - Serialize paragraphs to OOXML XML
 *
 * Converts Paragraph objects back to <w:p> XML format for DOCX files.
 * Handles all paragraph properties and child content (runs, hyperlinks, fields, bookmarks).
 *
 * OOXML Reference:
 * - Paragraph: w:p
 * - Paragraph properties: w:pPr
 * - Runs, hyperlinks, bookmarks, fields as child elements
 */

import type {
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphMarkChange,
  Run,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  SimpleField,
  ComplexField,
  InlineSdt,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  MoveFromRangeStart,
  MoveToRangeStart,
  ParagraphPropertyChange,
  SdtProperties,
  TabStop,
  ShadingProperties,
  TextFormatting,
  TrackedChangeInfo,
} from "../../types/document";
import { normalizeRevisionId } from "@stll/docx-core/model";
import { isValidHexColor } from "../../utils/colorResolver";
import { numPrEqual } from "../numberingParser";
import { reconcileRawSdtPr } from "../sdtPropertiesPatch";
import { serializeBorder } from "./borderSerializer";
// oxlint-disable-next-line import/no-cycle -- OOXML model is mutually recursive: paragraphs hold runs, shape-textbox runs hold paragraphs
import { serializeRun, serializeTextFormatting } from "./runSerializer";
import { serializeSectionProperties } from "./sectionPropertiesSerializer";
import { escapeXml, intAttr, isSingleWellFormedElement } from "./xmlUtils";

// ============================================================================
// BORDER SERIALIZATION
// ============================================================================

/**
 * Serialize paragraph borders (w:pBdr)
 */
function serializeParagraphBorders(borders: ParagraphFormatting["borders"]): string {
  if (!borders) {
    return "";
  }

  const parts: string[] = [];

  if (borders.top) {
    const topXml = serializeBorder(borders.top, "top");
    if (topXml) {
      parts.push(topXml);
    }
  }

  if (borders.left) {
    const leftXml = serializeBorder(borders.left, "left");
    if (leftXml) {
      parts.push(leftXml);
    }
  }

  if (borders.bottom) {
    const bottomXml = serializeBorder(borders.bottom, "bottom");
    if (bottomXml) {
      parts.push(bottomXml);
    }
  }

  if (borders.right) {
    const rightXml = serializeBorder(borders.right, "right");
    if (rightXml) {
      parts.push(rightXml);
    }
  }

  if (borders.between) {
    const betweenXml = serializeBorder(borders.between, "between");
    if (betweenXml) {
      parts.push(betweenXml);
    }
  }

  if (borders.bar) {
    const barXml = serializeBorder(borders.bar, "bar");
    if (barXml) {
      parts.push(barXml);
    }
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:pBdr>${parts.join("")}</w:pBdr>`;
}

// ============================================================================
// SHADING SERIALIZATION
// ============================================================================

/**
 * Serialize shading properties (w:shd)
 */
function serializeShading(shading: ShadingProperties | undefined): string {
  if (!shading) {
    return "";
  }

  const attrs: string[] = [];

  // Pattern/val
  if (shading.pattern) {
    attrs.push(`w:val="${escapeXml(shading.pattern)}"`);
  } else {
    attrs.push('w:val="clear"');
  }

  // Color (pattern color)
  if (shading.color?.rgb && isValidHexColor(shading.color.rgb)) {
    attrs.push(`w:color="${escapeXml(shading.color.rgb)}"`);
  } else if (shading.color?.auto) {
    attrs.push('w:color="auto"');
  }

  // Fill (background color)
  if (shading.fill?.rgb && isValidHexColor(shading.fill.rgb)) {
    attrs.push(`w:fill="${escapeXml(shading.fill.rgb)}"`);
  } else if (shading.fill?.auto) {
    attrs.push('w:fill="auto"');
  }

  // Theme fill
  if (shading.fill?.themeColor) {
    attrs.push(`w:themeFill="${escapeXml(shading.fill.themeColor)}"`);
  }

  if (shading.fill?.themeTint) {
    attrs.push(`w:themeFillTint="${escapeXml(shading.fill.themeTint)}"`);
  }

  if (shading.fill?.themeShade) {
    attrs.push(`w:themeFillShade="${escapeXml(shading.fill.themeShade)}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:shd ${attrs.join(" ")}/>`;
}

// ============================================================================
// TAB STOPS SERIALIZATION
// ============================================================================

/**
 * Serialize tab stops (w:tabs)
 */
function serializeTabStops(tabs: TabStop[] | undefined): string {
  if (!tabs || tabs.length === 0) {
    return "";
  }

  const tabElements = tabs.map((tab) => {
    const attrs: string[] = [`w:val="${tab.alignment}"`, `w:pos="${intAttr(tab.position)}"`];

    if (tab.leader && tab.leader !== "none") {
      attrs.push(`w:leader="${tab.leader}"`);
    }

    return `<w:tab ${attrs.join(" ")}/>`;
  });

  return `<w:tabs>${tabElements.join("")}</w:tabs>`;
}

// ============================================================================
// SPACING SERIALIZATION
// ============================================================================

/**
 * Serialize spacing properties (w:spacing)
 */
function serializeSpacing(formatting: ParagraphFormatting): string {
  const attrs: string[] = [];

  if (formatting.spaceBefore !== undefined) {
    attrs.push(`w:before="${intAttr(formatting.spaceBefore)}"`);
  }

  if (formatting.spaceAfter !== undefined) {
    attrs.push(`w:after="${intAttr(formatting.spaceAfter)}"`);
  }

  if (formatting.lineSpacing !== undefined) {
    attrs.push(`w:line="${intAttr(formatting.lineSpacing)}"`);
  }

  if (formatting.lineSpacingRule) {
    attrs.push(`w:lineRule="${formatting.lineSpacingRule}"`);
  }

  if (formatting.beforeAutospacing !== undefined) {
    attrs.push(`w:beforeAutospacing="${formatting.beforeAutospacing ? "1" : "0"}"`);
  }

  if (formatting.afterAutospacing !== undefined) {
    attrs.push(`w:afterAutospacing="${formatting.afterAutospacing ? "1" : "0"}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:spacing ${attrs.join(" ")}/>`;
}

// ============================================================================
// INDENTATION SERIALIZATION
// ============================================================================

/**
 * Serialize indentation properties (w:ind)
 */
function serializeIndentation(formatting: ParagraphFormatting): string {
  const attrs: string[] = [];

  if (formatting.indentLeft !== undefined) {
    attrs.push(`w:left="${intAttr(formatting.indentLeft)}"`);
  }

  if (formatting.indentRight !== undefined) {
    attrs.push(`w:right="${intAttr(formatting.indentRight)}"`);
  }

  if (formatting.indentFirstLine !== undefined) {
    if (formatting.hangingIndent) {
      // Hanging indent is stored as positive value but uses w:hanging attribute
      attrs.push(`w:hanging="${intAttr(Math.abs(formatting.indentFirstLine))}"`);
    } else {
      attrs.push(`w:firstLine="${intAttr(formatting.indentFirstLine)}"`);
    }
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:ind ${attrs.join(" ")}/>`;
}

// ============================================================================
// NUMBERING SERIALIZATION
// ============================================================================

/**
 * Serialize numbering properties (w:numPr)
 */
function serializeNumbering(numPr: ParagraphFormatting["numPr"]): string {
  if (!numPr) {
    return "";
  }

  const parts: string[] = [];

  if (numPr.ilvl !== undefined) {
    parts.push(`<w:ilvl w:val="${intAttr(numPr.ilvl)}"/>`);
  }

  if (numPr.numId !== undefined) {
    parts.push(`<w:numId w:val="${intAttr(numPr.numId)}"/>`);
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:numPr>${parts.join("")}</w:numPr>`;
}

// ============================================================================
// FRAME PROPERTIES SERIALIZATION
// ============================================================================

/**
 * Serialize frame properties (w:framePr)
 */
function serializeFrameProperties(frame: ParagraphFormatting["frame"]): string {
  if (!frame) {
    return "";
  }

  const attrs: string[] = [];

  if (frame.dropCap) {
    attrs.push(`w:dropCap="${frame.dropCap}"`);
  }

  if (frame.lines !== undefined) {
    attrs.push(`w:lines="${intAttr(frame.lines)}"`);
  }

  if (frame.width !== undefined) {
    attrs.push(`w:w="${intAttr(frame.width)}"`);
  }

  if (frame.height !== undefined) {
    attrs.push(`w:h="${intAttr(frame.height)}"`);
  }

  if (frame.hSpace !== undefined) {
    attrs.push(`w:hSpace="${intAttr(frame.hSpace)}"`);
  }

  if (frame.vSpace !== undefined) {
    attrs.push(`w:vSpace="${intAttr(frame.vSpace)}"`);
  }

  if (frame.hAnchor) {
    attrs.push(`w:hAnchor="${frame.hAnchor}"`);
  }

  if (frame.vAnchor) {
    attrs.push(`w:vAnchor="${frame.vAnchor}"`);
  }

  if (frame.x !== undefined) {
    attrs.push(`w:x="${frame.x}"`);
  }

  if (frame.y !== undefined) {
    attrs.push(`w:y="${frame.y}"`);
  }

  if (frame.xAlign) {
    attrs.push(`w:xAlign="${frame.xAlign}"`);
  }

  if (frame.yAlign) {
    attrs.push(`w:yAlign="${frame.yAlign}"`);
  }

  if (frame.wrap) {
    attrs.push(`w:wrap="${frame.wrap}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:framePr ${attrs.join(" ")}/>`;
}

// ============================================================================
// PARAGRAPH PROPERTIES SERIALIZATION
// ============================================================================

/**
 * Serialize paragraph formatting properties to w:pPr XML
 */
function serializeTrackedChangeAttrs(info: TrackedChangeInfo): string {
  // NOTE: `w:initials` is intentionally NOT emitted — ECMA-376 CT_TrackChange
  // defines only w:id/w:author/w:date. Initials are carried in-model for UI
  // attribution only (w:comment is the sole standards-clean initials target).
  // Bound `w:id` here so an overflowing id cannot reach the XML (eigenpal #1093).
  const parts = [`w:id="${normalizeRevisionId(info.id)}"`, `w:author="${escapeXml(info.author)}"`];
  if (info.date !== undefined) {
    parts.push(`w:date="${escapeXml(info.date)}"`);
  }
  return parts.join(" ");
}

function serializeParagraphMarkChange(mark: ParagraphMarkChange): string {
  const attrs = serializeTrackedChangeAttrs(mark.info);
  return `<w:${mark.kind} ${attrs}/>`;
}

export function serializeParagraphFormatting(
  formatting: ParagraphFormatting | undefined,
  propertyChanges?: ParagraphPropertyChange[],
  pPrMark?: ParagraphMarkChange,
): string {
  const parts: string[] = [];

  // Emit a boolean toggle: a bare element for true, `w:val="0"` for an explicit
  // false (which disables a value inherited from a style, so the override
  // survives round-trip), and nothing when absent ("inherit").
  const pushToggle = (name: string, value: boolean | undefined): void => {
    if (value === true) {
      parts.push(`<w:${name}/>`);
    } else if (value === false) {
      parts.push(`<w:${name} w:val="0"/>`);
    }
  };

  if (formatting) {
    // Style reference (must be first)
    if (formatting.styleId) {
      parts.push(`<w:pStyle w:val="${escapeXml(formatting.styleId)}"/>`);
    }

    // Keep next/lines, contextual spacing, page break before.
    pushToggle("keepNext", formatting.keepNext);
    pushToggle("keepLines", formatting.keepLines);
    pushToggle("contextualSpacing", formatting.contextualSpacing);
    pushToggle("pageBreakBefore", formatting.pageBreakBefore);

    // Frame properties
    const frameXml = serializeFrameProperties(formatting.frame);
    if (frameXml) {
      parts.push(frameXml);
    }

    // Widow control
    pushToggle("widowControl", formatting.widowControl);
    pushToggle("snapToGrid", formatting.snapToGrid);

    // Numbering. Skip numPr that still equals its style-sourced value (see
    // ParagraphFormatting.numPrFromStyle) — the parser materialized it from
    // the style and writing it back as direct formatting would flip Word's
    // level-indent precedence on the saved file. Guards the direct
    // serialize-a-parsed-Document path; the PM save path already drops it
    // in fromProseDoc.
    const styleSourcedNumPr =
      formatting.numPrFromStyle != null && numPrEqual(formatting.numPr, formatting.numPrFromStyle);
    const numPrXml = styleSourcedNumPr ? "" : serializeNumbering(formatting.numPr);
    if (numPrXml) {
      parts.push(numPrXml);
    }

    // Paragraph borders
    const bordersXml = serializeParagraphBorders(formatting.borders);
    if (bordersXml) {
      parts.push(bordersXml);
    }

    // Shading
    const shadingXml = serializeShading(formatting.shading);
    if (shadingXml) {
      parts.push(shadingXml);
    }

    // Tabs
    const tabsXml = serializeTabStops(formatting.tabs);
    if (tabsXml) {
      parts.push(tabsXml);
    }

    // Suppress line numbers / auto hyphens
    pushToggle("suppressLineNumbers", formatting.suppressLineNumbers);
    pushToggle("suppressAutoHyphens", formatting.suppressAutoHyphens);
    pushToggle("kinsoku", formatting.kinsoku);
    pushToggle("overflowPunct", formatting.overflowPunctuation);

    // Spacing
    const spacingXml = serializeSpacing(formatting);
    if (spacingXml) {
      parts.push(spacingXml);
    }

    // Indentation
    const indXml = serializeIndentation(formatting);
    if (indXml) {
      parts.push(indXml);
    }

    // Text direction (bidi)
    pushToggle("bidi", formatting.bidi);

    // Justification
    if (formatting.alignment) {
      parts.push(`<w:jc w:val="${formatting.alignment}"/>`);
    }

    // Outline level
    if (formatting.outlineLevel !== undefined) {
      parts.push(`<w:outlineLvl w:val="${formatting.outlineLevel}"/>`);
    }

    // Run properties (default run formatting for paragraph)
    // Round-trip `<w:specVanish/>` (run-in heading marker, ECMA-376
    // §17.3.1.32) by injecting it into the paragraph mark's rPr.
    // The parser populates `formatting.runInWithNext` from this
    // element; the layout engine consumes it via toFlowBlocks'
    // run-in merge. Without serializing it back, saving a doc
    // through Folio loses the soft paragraph break and the heading
    // becomes a normal separate paragraph in Word.
    //
    // EG_ParaRPrTrackChanges (ECMA-376 §17.13.5 / wml.xsd:1837) puts
    // <w:ins>/<w:del> FIRST inside the paragraph mark's rPr; strict
    // readers reject other orderings.
    if (pPrMark || formatting.runProperties || formatting.runInWithNext) {
      const pPrMarkXml = pPrMark ? serializeParagraphMarkChange(pPrMark) : "";
      const innerRPr = formatting.runProperties
        ? extractRPrInner(serializeTextFormatting(formatting.runProperties))
        : "";
      const specVanishXml = formatting.runInWithNext ? "<w:specVanish/>" : "";
      const fullInner = `${pPrMarkXml}${innerRPr}${specVanishXml}`;
      if (fullInner.length > 0) {
        parts.push(`<w:rPr>${fullInner}</w:rPr>`);
      }
    }
  } else if (pPrMark) {
    parts.push(`<w:rPr>${serializeParagraphMarkChange(pPrMark)}</w:rPr>`);
  }

  if (propertyChanges && propertyChanges.length > 0) {
    parts.push(...propertyChanges.map((change) => serializeParagraphPropertyChange(change)));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:pPr>${parts.join("")}</w:pPr>`;
}

function extractPPrInner(pPrXml: string): string {
  if (!pPrXml.startsWith("<w:pPr>") || !pPrXml.endsWith("</w:pPr>")) {
    return "";
  }
  return pPrXml.slice("<w:pPr>".length, -"</w:pPr>".length);
}

/**
 * Strip the outer `<w:rPr>...</w:rPr>` wrapper so callers can splice
 * additional rPr children (e.g. `<w:specVanish/>`) and re-emit a
 * single rPr element.
 */
function extractRPrInner(rPrXml: string): string {
  if (!rPrXml.startsWith("<w:rPr>") || !rPrXml.endsWith("</w:rPr>")) {
    return "";
  }
  return rPrXml.slice("<w:rPr>".length, -"</w:rPr>".length);
}

function serializeParagraphPropertyChange(change: ParagraphPropertyChange): string {
  const normalizedId = normalizeRevisionId(change.info.id);
  const authorCandidate = typeof change.info.author === "string" ? change.info.author.trim() : "";
  const normalizedAuthor = authorCandidate.length > 0 ? authorCandidate : "Unknown";
  const normalizedDate = typeof change.info.date === "string" ? change.info.date.trim() : undefined;
  const normalizedRsid = typeof change.info.rsid === "string" ? change.info.rsid.trim() : undefined;
  const attrs = [`w:id="${normalizedId}"`, `w:author="${escapeXml(normalizedAuthor)}"`];
  if (normalizedDate) {
    attrs.push(`w:date="${escapeXml(normalizedDate)}"`);
  }
  if (normalizedRsid) {
    attrs.push(`w:rsid="${escapeXml(normalizedRsid)}"`);
  }

  const previousPPrXml = serializeParagraphFormatting(change.previousFormatting) || "<w:pPr/>";
  const previousPPrInner = extractPPrInner(previousPPrXml);
  const normalizedPreviousPPr =
    previousPPrInner.length > 0 ? `<w:pPr>${previousPPrInner}</w:pPr>` : "<w:pPr/>";
  return `<w:pPrChange ${attrs.join(" ")}>${normalizedPreviousPPr}</w:pPrChange>`;
}

// ============================================================================
// CONTENT SERIALIZATION
// ============================================================================

/**
 * Serialize a hyperlink (w:hyperlink)
 */
function serializeHyperlink(hyperlink: Hyperlink): string {
  const attrs: string[] = [];

  if (hyperlink.rId) {
    attrs.push(`r:id="${hyperlink.rId}"`);
  }

  if (hyperlink.anchor) {
    attrs.push(`w:anchor="${escapeXml(hyperlink.anchor)}"`);
  }

  if (hyperlink.tooltip) {
    attrs.push(`w:tooltip="${escapeXml(hyperlink.tooltip)}"`);
  }

  if (hyperlink.target) {
    attrs.push(`w:tgtFrame="${escapeXml(hyperlink.target)}"`);
  }

  // Round-trip an explicit `w:history` either way. The parser only sets
  // `history` from a present `w:history="1"`/`"0"`, so emitting nothing for
  // `true` used to drop the attribute on save.
  if (hyperlink.history === true) {
    attrs.push('w:history="1"');
  } else if (hyperlink.history === false) {
    attrs.push('w:history="0"');
  }

  if (hyperlink.docLocation) {
    attrs.push(`w:docLocation="${escapeXml(hyperlink.docLocation)}"`);
  }

  // Serialize children
  const childrenXml = hyperlink.children
    .map((child) => {
      if (child.type === "run") {
        return serializeRun(child);
      }
      if (child.type === "bookmarkStart") {
        return serializeBookmarkStart(child);
      }
      return serializeBookmarkEnd(child);
    })
    .join("");

  const attrsStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return `<w:hyperlink${attrsStr}>${childrenXml}</w:hyperlink>`;
}

/**
 * Serialize bookmark start (w:bookmarkStart)
 */
function serializeBookmarkStart(bookmark: BookmarkStart): string {
  const attrs: string[] = [`w:id="${bookmark.id}"`, `w:name="${escapeXml(bookmark.name)}"`];

  if (bookmark.colFirst !== undefined) {
    attrs.push(`w:colFirst="${bookmark.colFirst}"`);
  }

  if (bookmark.colLast !== undefined) {
    attrs.push(`w:colLast="${bookmark.colLast}"`);
  }

  return `<w:bookmarkStart ${attrs.join(" ")}/>`;
}

/**
 * Serialize bookmark end (w:bookmarkEnd)
 */
function serializeBookmarkEnd(bookmark: BookmarkEnd): string {
  return `<w:bookmarkEnd w:id="${bookmark.id}"/>`;
}

/** Serialize a simple field without changing its authored OOXML field form. */
function serializeSimpleField(field: SimpleField): string {
  const attrs = [`w:instr="${escapeXml(field.instruction)}"`];
  if (field.fldLock) {
    attrs.push('w:fldLock="true"');
  }
  if (field.dirty) {
    attrs.push('w:dirty="true"');
  }

  const contentXml = field.content
    .map((item) => (item.type === "run" ? serializeRun(item) : serializeHyperlink(item)))
    .join("");

  return `<w:fldSimple ${attrs.join(" ")}>${contentXml}</w:fldSimple>`;
}

/**
 * Serialize a complex field
 * Complex fields are represented by multiple runs with fldChar elements,
 * so we convert them back to that structure
 */
function serializeComplexField(field: ComplexField): string {
  const parts: string[] = [];

  // Formatting for the structural runs (begin/separate/end). Prefer the field's
  // captured run formatting: the parser re-captures ComplexField.formatting from
  // the first non-empty field-run `w:rPr`, so re-presenting it on the begin run
  // keeps that value stable across a save→parse round-trip. Fall back to the
  // first result run's formatting when the field has no captured formatting.
  // The collapsed PAGE field (no result run) still recovers its `w:rPr`
  // (size/color) from field.formatting (eigenpal/docx-editor#909).
  const structuralFormatting = field.formatting ?? field.fieldResult[0]?.formatting;
  const rPrXml = structuralFormatting ? serializeTextFormatting(structuralFormatting) : "";

  // Begin field character (never set dirty — dirty causes apps to recalculate
  // and potentially discard run formatting)
  const beginAttrs: string[] = ['w:fldCharType="begin"'];
  if (field.fldLock) {
    beginAttrs.push('w:fldLock="true"');
  }
  parts.push(`<w:r>${rPrXml}<w:fldChar ${beginAttrs.join(" ")}/></w:r>`);

  // Field code (instrText)
  if (field.fieldCode.length > 0) {
    parts.push(...field.fieldCode.map((run) => serializeRun(run)));
  } else if (field.instruction.length > 0) {
    // Fallback: create instrText from instruction
    const needsPreserve =
      field.instruction.startsWith(" ") ||
      field.instruction.endsWith(" ") ||
      field.instruction.includes("  ");
    const spaceAttr = needsPreserve ? ' xml:space="preserve"' : "";
    parts.push(
      `<w:r>${rPrXml}<w:instrText${spaceAttr}>${escapeXml(field.instruction)}</w:instrText></w:r>`,
    );
  }

  // Separate field character
  parts.push(`<w:r>${rPrXml}<w:fldChar w:fldCharType="separate"/></w:r>`);

  // Field result
  parts.push(...field.fieldResult.map((run) => serializeRun(run)));

  // End field character
  parts.push(`<w:r>${rPrXml}<w:fldChar w:fldCharType="end"/></w:r>`);

  return parts.join("");
}

/**
 * Synthesize a `<w:sdtPr>` from the modeled {@link SdtProperties}.
 *
 * Only reached for an inline SDT that carries no captured `rawPropertiesXml`
 * (constructed programmatically rather than parsed from a DOCX). Mirrors the
 * block-SDT fallback: emit `w:id` first so the parsed numeric id survives,
 * then the shared identity fields, then the type-defining marker.
 */
function synthesizeInlineSdtPr(props: SdtProperties): string {
  const prParts: string[] = [];

  if (typeof props.id === "number") {
    prParts.push(`<w:id w:val="${props.id}"/>`);
  }
  if (props.alias) {
    prParts.push(`<w:alias w:val="${escapeXml(props.alias)}"/>`);
  }
  if (props.tag) {
    prParts.push(`<w:tag w:val="${escapeXml(props.tag)}"/>`);
  }
  if (props.lock && props.lock !== "unlocked") {
    prParts.push(`<w:lock w:val="${props.lock}"/>`);
  }
  if (props.placeholder) {
    // OOXML shape: `<w:placeholder><w:docPart w:val="..."/></w:placeholder>`.
    // The placeholder identifier lives in `w:val` on the nested `w:docPart`,
    // mirroring the parse in `paragraphParser.ts`.
    prParts.push(
      `<w:placeholder><w:docPart w:val="${escapeXml(props.placeholder)}"/></w:placeholder>`,
    );
  }
  if (props.showingPlaceholder) {
    prParts.push("<w:showingPlcHdr/>");
  }

  // Type-specific properties
  switch (props.sdtType) {
    case "plainText":
      prParts.push("<w:text/>");
      break;
    case "date": {
      // `w:date@w:fullDate` is the ISO-8601 bound value; `w:dateFormat` is
      // the display format. Older code (before the shared parser
      // split these) wrote the format into `w:fullDate`, which corrupted
      // round-trip — keep them on separate model fields and emit each
      // into its right element.
      const fullDateAttr = props.dateValueISO
        ? ` w:fullDate="${escapeXml(props.dateValueISO)}"`
        : "";
      const formatChild = props.dateFormat
        ? `<w:dateFormat w:val="${escapeXml(props.dateFormat)}"/>`
        : "";
      if (fullDateAttr || formatChild) {
        prParts.push(`<w:date${fullDateAttr}>${formatChild}</w:date>`);
      } else {
        prParts.push("<w:date/>");
      }
      break;
    }
    case "dropdown": {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXml(i.displayText)}" w:value="${escapeXml(i.value)}"/>`,
        )
        .join("");
      prParts.push(`<w:dropDownList>${items}</w:dropDownList>`);
      break;
    }
    case "comboBox": {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXml(i.displayText)}" w:value="${escapeXml(i.value)}"/>`,
        )
        .join("");
      prParts.push(`<w:comboBox>${items}</w:comboBox>`);
      break;
    }
    case "checkbox":
      prParts.push(
        `<w14:checkbox><w14:checked w14:val="${props.checked ? "1" : "0"}"/></w14:checkbox>`,
      );
      break;
    case "picture":
      prParts.push("<w:picture/>");
      break;
    case "richText":
    case "buildingBlockGallery":
    case "group":
    case "unknown":
      // These SDT variants carry no type-specific properties in OOXML;
      // the surrounding sdtPr fields (alias/tag/lock/...) carry all
      // round-trippable state for them.
      break;
  }

  return `<w:sdtPr>${prParts.join("")}</w:sdtPr>`;
}

/**
 * Serialize an inline SDT (w:sdt).
 *
 * Replays the captured `<w:sdtPr>` / `<w:sdtEndPr>` verbatim so unmodeled
 * OOXML features (`w:id`, `w:dataBinding`, `w15:*`, custom XML mappings)
 * survive the round-trip, mirroring the block-SDT serializer. Without this
 * the properties block was re-synthesized from the modeled projection alone,
 * silently dropping every unmodeled feature (and `w:sdtEndPr`) on save.
 */
function serializeInlineSdt(sdt: InlineSdt): string {
  const props = sdt.properties;

  const contentXml = sdt.content
    .map((item): string => {
      switch (item.type) {
        case "run":
          return serializeRun(item);
        case "hyperlink":
          return serializeHyperlink(item);
        case "simpleField":
          return serializeSimpleField(item);
        case "complexField":
          return serializeComplexField(item);
        case "inlineSdt":
          return serializeInlineSdt(item);
        case "insertion":
          return serializeTrackedChange("ins", item);
        case "deletion":
          return serializeTrackedChange("del", item);
        case "moveFrom":
          return serializeTrackedChange("moveFrom", item);
        case "moveTo":
          return serializeTrackedChange("moveTo", item);
        case "mathEquation":
          // Round-trip the raw OMML XML directly
          return item.ommlXml || "";
        default: {
          // Exhaustiveness check: if a new type is added to
          // InlineSdt['content'] (see docx-core/src/model/content.ts)
          // without a matching case here, TypeScript errors out instead
          // of silently dropping content on save. Keep this in sync with
          // the filter in createInlineSdtFromNode (fromProseDoc.ts).
          const _exhaustive: never = item;
          return _exhaustive;
        }
      }
    })
    .join("");

  // Reconcile any modeled interactive edit (checkbox toggle, date pick,
  // dropdown selection) into the raw properties before replay so it is not
  // discarded, exactly as the block-SDT serializer does. Unmodeled markers
  // inside the raw string are left untouched.
  // Replay the captured snapshot only when it is structurally a single
  // `<w:sdtPr>`/`<w:sdtEndPr>` element — a malformed or attacker-supplied
  // string (e.g. one that closes `<w:sdt>` early or injects sibling markup)
  // falls back to a synthesized properties block instead of being spliced
  // into the document verbatim.
  const baseSdtPr =
    props.rawPropertiesXml && isSingleWellFormedElement(props.rawPropertiesXml, "sdtPr")
      ? props.rawPropertiesXml
      : synthesizeInlineSdtPr(props);
  const dateFullDate =
    props.sdtType === "date" && props.dateValueISO ? props.dateValueISO : undefined;
  const dropdownLastValue =
    (props.sdtType === "dropdown" || props.sdtType === "comboBox") &&
    typeof props.dropdownLastValue === "string"
      ? props.dropdownLastValue
      : undefined;
  const sdtPrXml = reconcileRawSdtPr(baseSdtPr, props, {
    ...(dateFullDate !== undefined ? { dateFullDate } : {}),
    ...(dropdownLastValue !== undefined ? { dropdownLastValue } : {}),
  });
  const sdtEndPrXml =
    props.rawEndPropertiesXml && isSingleWellFormedElement(props.rawEndPropertiesXml, "sdtEndPr")
      ? props.rawEndPropertiesXml
      : "";

  return `<w:sdt>${sdtPrXml}${sdtEndPrXml}<w:sdtContent>${contentXml}</w:sdtContent></w:sdt>`;
}

function serializeMoveRangeStart(
  tag: "moveFromRangeStart" | "moveToRangeStart",
  marker: MoveFromRangeStart | MoveToRangeStart,
): string {
  const attrs = [`w:id="${marker.id}"`, `w:name="${escapeXml(marker.name)}"`];
  return `<w:${tag} ${attrs.join(" ")}/>`;
}

/**
 * Serialize a tracked change wrapper (ins/del/moveFrom/moveTo)
 */
function rewriteRunTextAsDeleted(xml: string): string {
  return xml
    .replace(/<w:t\b/gu, "<w:delText")
    .replace(/<\/w:t>/gu, "</w:delText>")
    .replace(/<w:instrText\b/gu, "<w:delInstrText")
    .replace(/<\/w:instrText>/gu, "</w:delInstrText>");
}

function serializeTrackedChange(
  tag: "ins" | "del" | "moveFrom" | "moveTo",
  change: Insertion | Deletion | MoveFrom | MoveTo,
): string {
  const info = change.info;
  const normalizedId = normalizeRevisionId(info.id);
  const authorCandidate = typeof info.author === "string" ? info.author.trim() : "";
  const normalizedAuthor = authorCandidate.length > 0 ? authorCandidate : "Unknown";
  const normalizedDate = typeof info.date === "string" ? info.date.trim() : undefined;
  // `w:initials` is intentionally NOT emitted (non-standard on CT_TrackChange).
  const attrs = [`w:id="${normalizedId}"`, `w:author="${escapeXml(normalizedAuthor)}"`];
  if (normalizedDate) {
    attrs.push(`w:date="${escapeXml(normalizedDate)}"`);
  }

  const serializeDeletedRun = (run: Run): string => {
    const xml = serializeRun(run);
    const hasDrawingContent = run.content.some((c) => c.type === "drawing" || c.type === "shape");
    if (!hasDrawingContent) {
      return rewriteRunTextAsDeleted(xml);
    }

    const hasTextualContent = run.content.some((c) => c.type !== "drawing" && c.type !== "shape");
    if (!hasTextualContent) {
      return xml;
    }

    return run.content
      .map((content) => {
        const contentXml = serializeRun({ ...run, content: [content] });
        if (content.type === "drawing" || content.type === "shape") {
          return contentXml;
        }
        return rewriteRunTextAsDeleted(contentXml);
      })
      .join("");
  };

  const contentXml = change.content
    .map((item) => {
      if (item.type === "run") {
        // A deleted drawing/shape run keeps its content verbatim: a picture
        // has no `<w:t>`, and a shape's nested textbox text
        // (`<w:txbxContent><w:t>`) must NOT be rewritten to `<w:delText>` —
        // that markup belongs only to a run's own deleted text, not to a
        // nested textbox document. eigenpal #641.
        if (tag === "del" || tag === "moveFrom") {
          return serializeDeletedRun(item);
        }
        return serializeRun(item);
      }
      return serializeHyperlink(item);
    })
    .join("");

  return `<w:${tag} ${attrs.join(" ")}>${contentXml}</w:${tag}>`;
}

/** Emit the `<w:commentReference>` run Word places after a comment range end. */
function serializeCommentReferenceRun(id: number): string {
  return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
}

/**
 * Serialize a single paragraph content item.
 *
 * `explicitCommentReferenceIds` holds the comment ids that already have their
 * own `commentReference` node in this paragraph (the parsed-from-Word shape).
 * For those, the `commentRangeEnd` marker must NOT also synthesize a reference
 * run, or a save→parse round-trip doubles the `<w:commentReference>`. The
 * editor (fromProseDoc) path emits range markers with no reference node, so the
 * synthetic run is still written when the id is absent from the set.
 */
function serializeParagraphContent(
  content: ParagraphContent,
  explicitCommentReferenceIds: ReadonlySet<number>,
): string {
  switch (content.type) {
    case "run":
      return serializeRun(content);
    case "hyperlink":
      return serializeHyperlink(content);
    case "bookmarkStart":
      return serializeBookmarkStart(content);
    case "bookmarkEnd":
      return serializeBookmarkEnd(content);
    case "simpleField":
      return serializeSimpleField(content);
    case "complexField":
      return serializeComplexField(content);
    case "inlineSdt":
      return serializeInlineSdt(content);
    case "commentRangeStart":
      return `<w:commentRangeStart w:id="${content.id}"/>`;
    case "commentRangeEnd":
      return explicitCommentReferenceIds.has(content.id)
        ? `<w:commentRangeEnd w:id="${content.id}"/>`
        : `<w:commentRangeEnd w:id="${content.id}"/>${serializeCommentReferenceRun(content.id)}`;
    case "commentReference":
      return serializeCommentReferenceRun(content.id);
    case "insertion":
      return serializeTrackedChange("ins", content);
    case "deletion":
      return serializeTrackedChange("del", content);
    case "moveFrom":
      return serializeTrackedChange("moveFrom", content);
    case "moveTo":
      return serializeTrackedChange("moveTo", content);
    case "moveFromRangeStart":
      return serializeMoveRangeStart("moveFromRangeStart", content as MoveFromRangeStart);
    case "moveFromRangeEnd":
      return `<w:moveFromRangeEnd w:id="${content.id}"/>`;
    case "moveToRangeStart":
      return serializeMoveRangeStart("moveToRangeStart", content as MoveToRangeStart);
    case "moveToRangeEnd":
      return `<w:moveToRangeEnd w:id="${content.id}"/>`;
    case "mathEquation":
      // Round-trip the raw OMML XML directly
      return content.ommlXml || "";
    default:
      return "";
  }
}

// ============================================================================
// MAIN SERIALIZATION
// ============================================================================

/**
 * Serialize a paragraph to OOXML XML (w:p)
 *
 * @param paragraph - The paragraph to serialize
 * @returns XML string for the paragraph
 */
export function serializeParagraph(paragraph: Paragraph): string {
  const parts: string[] = [];

  // Paragraph ID attributes
  const attrs: string[] = [];
  if (paragraph.paraId) {
    attrs.push(`w14:paraId="${escapeXml(paragraph.paraId)}"`);
  }
  if (paragraph.textId) {
    attrs.push(`w14:textId="${escapeXml(paragraph.textId)}"`);
  }
  const attrsStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";

  // Add paragraph properties if present
  const pPrXml = serializeParagraphFormatting(
    paragraph.formatting,
    paragraph.propertyChanges,
    paragraph.pPrMark,
  );
  const sectionPropertiesXml = serializeSectionProperties(paragraph.sectionProperties);
  if (pPrXml || sectionPropertiesXml) {
    parts.push(`<w:pPr>${extractPPrInner(pPrXml)}${sectionPropertiesXml}</w:pPr>`);
  }

  // Comment ids whose reference run is modeled explicitly (parsed-from-Word),
  // so the matching commentRangeEnd does not double-emit it (see
  // serializeParagraphContent).
  const explicitCommentReferenceIds = new Set<number>();
  for (const content of paragraph.content) {
    if (content.type === "commentReference") {
      explicitCommentReferenceIds.add(content.id);
    }
  }

  // Add paragraph content
  let pendingRenderedPageBreak = paragraph.renderedPageBreakBefore === true;
  for (const content of paragraph.content) {
    let contentXml = serializeParagraphContent(content, explicitCommentReferenceIds);
    if (contentXml) {
      if (pendingRenderedPageBreak) {
        const next = injectRenderedPageBreakIntoFirstRun(contentXml);
        if (next) {
          contentXml = next;
          pendingRenderedPageBreak = false;
        }
      }
      parts.push(contentXml);
    }
  }

  return `<w:p${attrsStr}>${parts.join("")}</w:p>`;
}

function injectRenderedPageBreakIntoFirstRun(xml: string): string | null {
  const runOpeningTag = /<w:r(?=[\s>/])[^>]*>/u;
  const openingTag = runOpeningTag.exec(xml);
  if (!openingTag) {
    return null;
  }
  const runEnd = xml.indexOf("</w:r>", openingTag.index + openingTag[0].length);
  if (
    runEnd !== -1 &&
    xml
      .slice(openingTag.index + openingTag[0].length, runEnd)
      .includes("<w:lastRenderedPageBreak/>")
  ) {
    return xml;
  }
  return xml.replace(runOpeningTag, (match) => `${match}<w:lastRenderedPageBreak/>`);
}

/**
 * Serialize multiple paragraphs to OOXML XML
 *
 * @param paragraphs - The paragraphs to serialize
 * @returns XML string for all paragraphs
 */
export function serializeParagraphs(paragraphs: Paragraph[]): string {
  return paragraphs.map(serializeParagraph).join("");
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if a paragraph has any content
 */
export function hasParagraphContent(paragraph: Paragraph): boolean {
  return paragraph.content.length > 0;
}

/**
 * Check if a paragraph has formatting
 */
export function hasParagraphFormatting(paragraph: Paragraph): boolean {
  return paragraph.formatting !== undefined && Object.keys(paragraph.formatting).length > 0;
}

/**
 * Get plain text from a paragraph (for comparison/debugging)
 */
export function getParagraphPlainText(paragraph: Paragraph): string {
  const texts: string[] = [];

  for (const content of paragraph.content) {
    if (content.type === "run") {
      for (const item of content.content) {
        if (item.type === "text") {
          texts.push(item.text);
        } else if (item.type === "tab") {
          texts.push("\t");
        } else if (item.type === "break") {
          texts.push("\n");
        }
      }
    } else if (content.type === "hyperlink") {
      for (const child of content.children) {
        if (child.type === "run") {
          for (const item of child.content) {
            if (item.type === "text") {
              texts.push(item.text);
            }
          }
        }
      }
    } else if (
      content.type === "simpleField" ||
      content.type === "inlineSdt" ||
      content.type === "insertion" ||
      content.type === "deletion" ||
      content.type === "moveFrom" ||
      content.type === "moveTo"
    ) {
      for (const item of content.content) {
        if (item.type === "run") {
          for (const subItem of item.content) {
            if (subItem.type === "text") {
              texts.push(subItem.text);
            }
          }
        }
      }
    } else if (content.type === "complexField") {
      for (const run of content.fieldResult) {
        for (const item of run.content) {
          if (item.type === "text") {
            texts.push(item.text);
          }
        }
      }
    }
  }

  return texts.join("");
}

/**
 * Create an empty paragraph
 */
export function createEmptyParagraph(formatting?: ParagraphFormatting): Paragraph {
  return {
    type: "paragraph",
    ...(formatting !== undefined ? { formatting } : {}),
    content: [],
  };
}

/**
 * Create a paragraph with a single text run
 */
export function createTextParagraph(
  text: string,
  paragraphFormatting?: ParagraphFormatting,
  textFormatting?: TextFormatting,
): Paragraph {
  return {
    type: "paragraph",
    ...(paragraphFormatting !== undefined ? { formatting: paragraphFormatting } : {}),
    content: [
      {
        type: "run",
        ...(textFormatting !== undefined ? { formatting: textFormatting } : {}),
        content: [{ type: "text", text }],
      },
    ],
  };
}

/**
 * Check if paragraph is a list item
 */
export function isListParagraph(paragraph: Paragraph): boolean {
  return paragraph.formatting?.numPr !== undefined;
}

/**
 * Get list level of a paragraph (0-8, or -1 if not a list)
 */
export function getListLevel(paragraph: Paragraph): number {
  return paragraph.formatting?.numPr?.ilvl ?? -1;
}
