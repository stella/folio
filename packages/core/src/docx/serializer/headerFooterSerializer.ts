/**
 * Header/Footer Serializer - Serialize headers/footers to OOXML XML
 *
 * Converts HeaderFooter objects back to valid header*.xml / footer*.xml format.
 * Reuses paragraph and table serializers for content.
 *
 * OOXML Reference:
 * - Header root: w:hdr
 * - Footer root: w:ftr
 * - Content: w:p, w:tbl (same as document body)
 */

import type { BlockContent, HeaderFooter, Watermark } from "../../types/document";
import { getHeaderFooterVerbatimXml, canReplayHeaderFooterVerbatim } from "../headerFooterVerbatim";
import { serializeBlockSdt } from "./blockSdtSerializer";
import { serializePartElement, type OoxmlNamespacePrefix, type SourcePart } from "./partNamespaces";
import { serializeParagraph } from "./paragraphSerializer";
import { serializeTable } from "./tableSerializer";
import { escapeXml } from "./xmlUtils";

// Prefixes a header/footer declares whether or not the body uses them. Mirrors
// the document serializer's baseline so any raw replay path (`rawPropertiesXml`,
// unmodeled OOXML extensions inside a captured SDT) lands on a root that
// declares every standard prefix it might use. `a` and `pic` cover DrawingML
// watermarks: `rawWatermarkXml` carries `<a:graphic>` / `<a:txBody>` /
// `<pic:pic>` descendants, but the hosting paragraph does not preserve the
// original header's ancestor declarations.
const HEADER_FOOTER_BASELINE_PREFIXES = [
  "wpc",
  "mc",
  "o",
  "r",
  "m",
  "v",
  "a",
  "pic",
  "wp14",
  "wp",
  "w10",
  "w",
  "w14",
  "w15",
  "w16",
  "w16cex",
  "w16cid",
  "w16sdtdh",
  "w16se",
  "wpg",
  "wps",
] as const satisfies readonly OoxmlNamespacePrefix[];

/**
 * Serialize a block content item (paragraph, table, or block-level SDT) for
 * header/footer.
 */
function serializeBlock(block: BlockContent): string {
  if (block.type === "paragraph") {
    return serializeParagraph(block);
  }
  if (block.type === "table") {
    return serializeTable(block, serializeParagraph);
  }
  return serializeBlockSdt(block, serializeBlock);
}

/**
 * Serialize a HeaderFooter object to valid OOXML XML
 *
 * @param hf - HeaderFooter object to serialize
 * @param source - The part being replaced, so a prefix only the source
 *   document bound keeps its URI
 * @returns Complete XML string for header*.xml or footer*.xml
 */
export function serializeHeaderFooter(hf: HeaderFooter, source?: SourcePart): string {
  const verbatim = getHeaderFooterVerbatimXml(hf);
  if (verbatim && canReplayHeaderFooterVerbatim(hf)) {
    return verbatim;
  }

  const rootTag = hf.type === "header" ? "w:hdr" : "w:ftr";

  // Watermark replay. The parser captured the hosting paragraph's
  // verbatim XML (`rawWatermarkXml`) and detached it from `content`,
  // so emit it back at its original block position (tracked in
  // `watermarkBlockIndex`). If a caller mutated `hf.watermark` without
  // updating the raw XML (e.g. the setDocumentWatermark path), the
  // model-driven synthesizer takes over and emits a freshly-built VML
  // watermark paragraph at the top of the header.
  const watermarkXml = serializeWatermarkParagraph(hf);
  const watermarkInsertIndex =
    hf.watermarkBlockIndex !== undefined
      ? Math.max(0, Math.min(hf.watermarkBlockIndex, hf.content.length))
      : 0;

  const blocksXml = hf.content.map((block) => serializeBlock(block));
  let contentXml: string;
  if (watermarkXml) {
    contentXml =
      blocksXml.slice(0, watermarkInsertIndex).join("") +
      watermarkXml +
      blocksXml.slice(watermarkInsertIndex).join("");
  } else {
    contentXml = blocksXml.join("");
  }

  // Ensure at least one empty paragraph (required by OOXML spec)
  if (!contentXml) {
    contentXml = "<w:p><w:pPr/></w:p>";
  }

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    serializePartElement({
      partPath: source?.path ?? `word/${hf.type}.xml`,
      rootName: rootTag,
      baselinePrefixes: HEADER_FOOTER_BASELINE_PREFIXES,
      sourceBindings: source?.bindings,
      body: contentXml,
    })
  );
}

function serializeWatermarkParagraph(hf: HeaderFooter): string {
  if (hf.rawWatermarkXml) {
    return hf.rawWatermarkXml;
  }
  if (hf.watermark) {
    return synthesizeWatermarkParagraph(hf.watermark);
  }
  return "";
}

function synthesizeWatermarkParagraph(watermark: Watermark): string {
  if (watermark.kind === "text") {
    return synthesizeTextWatermark(watermark);
  }
  return synthesizePictureWatermark(watermark);
}

function synthesizeTextWatermark(watermark: Extract<Watermark, { kind: "text" }>): string {
  // VML shape values mirror what Word's "Insert → Watermark" UI emits.
  // The shapetype id 136 is the WordArt template the model is anchored
  // on (gating in the parser); the size and offset numbers come from
  // Word's default text-watermark layout. Sufficient for round-trip
  // when the caller programmatically built the watermark; a parsed-
  // then-saved DOCX takes the raw replay path above.
  const rotation = watermark.diagonal === false ? 0 : 315;
  // `color: "auto"` is a documented model value meaning "use the
  // producer default"; we map it to Word's silver fallback rather than
  // emitting an invalid VML `fillcolor="#auto"`.
  const fillcolor =
    watermark.color && watermark.color !== "auto" ? `#${watermark.color}` : "#C0C0C0";
  const fontFamily = watermark.font ?? "Calibri";
  const text = escapeXml(watermark.text);
  // VML opacity rides on a `<v:fill>` child rather than the shape's
  // own `fillcolor` attribute. Word reads the decimal form (`opacity=
  // ".5"`) and the fixed-point form (`opacity="32768f"`); we use the
  // decimal form for clarity. Skip emission when the model carries no
  // explicit opacity so the saved DOCX matches what Word emits at the
  // default transparency.
  const fillChild =
    watermark.opacity !== undefined ? `<v:fill opacity="${watermark.opacity}"/>` : "";
  return `<w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:415pt;height:207pt;rotation:${rotation};z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="${fillcolor}" stroked="f">${fillChild}<v:textpath style="font-family:&quot;${escapeXml(fontFamily)}&quot;;font-size:1pt" string="${text}"/></v:shape></w:pict></w:r></w:p>`;
}

// Default picture-watermark dimensions Word's "Insert → Watermark"
// UI emits when no scale is specified. `scale` in the model is a
// multiplicative factor (1.0 = native), so 0.5 → half-size, 1.5 →
// one-and-a-half size; we apply it to both axes to preserve the
// caller's intended ratio.
const PICTURE_WATERMARK_DEFAULT_WIDTH_PT = 415;
const PICTURE_WATERMARK_DEFAULT_HEIGHT_PT = 207;
const PICTURE_WATERMARK_WASHOUT_GAIN = "19661f";
const PICTURE_WATERMARK_WASHOUT_BLACKLEVEL = "22938f";

function synthesizePictureWatermark(watermark: Extract<Watermark, { kind: "picture" }>): string {
  // Same VML shapetype convention as Word's UI: shape id begins with
  // `WordPictureWatermark` so a future round-trip parses cleanly via
  // the id-prefix guard.
  const rId = escapeXml(watermark.imageRId);
  // `scale` (Word's default-box multiplier) wins when set: it is the documented
  // resize knob, and the parser only records it for uniform (2:1) watermarks
  // where it agrees with the captured dimensions. When absent — e.g. a non-2:1
  // source — the captured per-image dimensions preserve the aspect ratio.
  const widthPt =
    watermark.scale === undefined
      ? (watermark.widthPt ?? PICTURE_WATERMARK_DEFAULT_WIDTH_PT)
      : PICTURE_WATERMARK_DEFAULT_WIDTH_PT * watermark.scale;
  const heightPt =
    watermark.scale === undefined
      ? (watermark.heightPt ?? PICTURE_WATERMARK_DEFAULT_HEIGHT_PT)
      : PICTURE_WATERMARK_DEFAULT_HEIGHT_PT * watermark.scale;
  const washoutAttrs =
    watermark.washout === false
      ? ""
      : ` gain="${PICTURE_WATERMARK_WASHOUT_GAIN}" blacklevel="${PICTURE_WATERMARK_WASHOUT_BLACKLEVEL}"`;
  return (
    `<w:p><w:r><w:pict>` +
    `<v:shape id="WordPictureWatermark1" type="#_x0000_t75" ` +
    `style="position:absolute;margin-left:0;margin-top:0;width:${widthPt}pt;height:${heightPt}pt;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin">` +
    `<v:imagedata r:id="${rId}" o:title=""${washoutAttrs}/>` +
    `</v:shape></w:pict></w:r></w:p>`
  );
}
