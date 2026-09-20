/**
 * Run Parser - Parse text runs (w:r) with complete formatting
 *
 * A run is a contiguous region of text with the same character formatting.
 * Runs can contain:
 * - Text (w:t)
 * - Tabs (w:tab)
 * - Line breaks (w:br)
 * - Symbols (w:sym)
 * - Footnote/endnote references
 * - Field characters
 * - Drawings/images (w:drawing)
 * - And more...
 *
 * OOXML Reference:
 * - Run: w:r
 * - Run properties: w:rPr
 * - Text content: w:t
 */

import type {
  Run,
  RunContent,
  TextContent,
  TabContent,
  BreakContent,
  SymbolContent,
  NoteReferenceContent,
  FieldCharContent,
  InstrTextContent,
  SoftHyphenContent,
  NoBreakHyphenContent,
  DrawingContent,
  RunPropertyChange,
  TextFormatting,
  ColorValue,
  Theme,
  Image,
  RelationshipMap,
  MediaFile,
  ShapeContent,
} from "../types/document";
import { DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";
import { attributeRemainder, NO_MODELLED_ATTRIBUTES } from "./attributeRemainder";
import {
  CAPTURE,
  type ChildHandlers,
  dispatchChildren,
  keptUnless,
  ownedElsewhere,
  sequencePositions,
} from "./containerChildren";
import { isGroupDrawing, parseGroupDrawing } from "./groupDrawingParser";
import { parseDiagramPreview } from "./diagramPreview";
import { parseImage } from "./imageParser";
import { imageRawXmlFingerprint } from "./imageRawXml";
import {
  EmphasisMarkSchema,
  FontHintSchema,
  FontThemeSchema,
  HighlightColorSchema,
  PositionalTabAlignmentSchema,
  PositionalTabLeaderSchema,
  PositionalTabRelativeToSchema,
  TextEffectSchema,
  UnderlineStyleSchema,
  narrowEnum,
} from "./parserEnums";
import { parseThemeColorAttribute } from "./themeColorAttribute";
import { parseShapeFromDrawing, shouldPreserveRawShapeDrawing } from "./shapeParser";
import type { StyleMap } from "./styleParser";
import { isTextBoxDrawing } from "./textBoxParser";
import { parseVmlImageContent, shouldPreserveRawVmlPict } from "./vmlImageParser";
import { resolveThemeFontRef } from "./themeParser";
import { parseHorizontalScalePercent } from "../utils/horizontalScale";
import { preserveRunChild } from "./preservedRunContent";
import { captureVerbatimXml } from "./verbatimCapture";
import { parseShading } from "./shadingParser";
import {
  cloneWithXmlnsDeclarations,
  findAllDeep,
  findChild,
  findChildren,
  getAttribute,
  getChildElements,
  getLocalName,
  getTextContent,
  mergeXmlnsDeclarations,
  parseBooleanElement,
  parseNumericAttribute,
  selectAlternateContentBranch,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";
import { parseFieldState } from "./fieldState";
import { parsePropertyChangeInfo } from "./trackedChangeInfo";

/**
 * Sanity cap on `w:lang` `@w:val`/`@w:eastAsia`/`@w:bidi` tag length. BCP-47
 * tags top out well under this; a hostile/corrupt tag here would drive the
 * hyphenation dictionary lookup and segmenter cache keying with an
 * attacker-sized string per run.
 */
const MAX_LANGUAGE_TAG_LENGTH = 35;

const truncateLanguageTag = (value: string | undefined): string | undefined =>
  value === undefined ? undefined : value.slice(0, MAX_LANGUAGE_TAG_LENGTH);

/**
 * Parse color value from attributes
 */
type ColorAttributes = {
  rgb: string | null;
  themeColor: string | null;
  themeTint: string | null;
  themeShade: string | null;
  element: string;
};

function parseColorValue({
  rgb,
  themeColor,
  themeTint,
  themeShade,
  element,
}: ColorAttributes): ColorValue {
  const color: ColorValue = {};

  if (rgb && rgb !== "auto") {
    color.rgb = rgb;
  } else if (rgb === "auto") {
    color.auto = true;
  }

  const parsedThemeColor = parseThemeColorAttribute({ raw: themeColor, element });
  if (parsedThemeColor !== undefined) {
    color.themeColor = parsedThemeColor;
  }

  if (themeTint) {
    color.themeTint = themeTint;
  }

  if (themeShade) {
    color.themeShade = themeShade;
  }

  return color;
}

/**
 * Which record beside this `w:rPr` reads children out of it.
 *
 * The four owners of a run property set share one handler map, because they
 * share one content model: `CT_RPr` is `EG_RPrBase` plus `w:rPrChange`, and
 * `CT_ParaRPr` opens that with `EG_ParaRPrTrackChanges`. They differ only in
 * which children a *sibling* record has already claimed, and a child claimed
 * twice is written twice. Naming the owner states that once, at the call site
 * that knows it, instead of leaving it to whichever reader ran first.
 */
export const RUN_PROPERTY_OWNERS = {
  /** A `w:r`'s own properties: `Run.propertyChanges` reads the `w:rPrChange`. */
  run: "run",
  /**
   * `w:pPr/w:rPr`. The paragraph's record reads the mark's revision
   * (`ParagraphMarkChange`) and its `w:specVanish`
   * (`ParagraphFormatting.runInWithNext`).
   */
  paragraphMark: "paragraphMark",
  /**
   * A style, a numbering level, a comment's reference mark, the snapshot
   * inside a `w:rPrChange`: nothing beside it reads any of its children.
   */
  standalone: "standalone",
} as const;

export type RunPropertyOwner = (typeof RUN_PROPERTY_OWNERS)[keyof typeof RUN_PROPERTY_OWNERS];

/**
 * The children each owner's sibling record claims.
 *
 * Not the decision map: that one is total over the schema and lives in
 * {@link parseRunProperties}. This names only the children a given owner reads
 * somewhere else, so both cannot write the same element.
 */
const OWNED_BY_A_SIBLING_RECORD = {
  run: {
    rPrChange: ownedElsewhere({
      container: "run-properties",
      child: "rPrChange",
      reader: "runParser#parseRun",
    }),
  },
  paragraphMark: {
    ins: ownedElsewhere({
      container: "run-properties",
      child: "ins",
      reader: "paragraphParser#parseParagraphProperties",
    }),
    del: ownedElsewhere({
      container: "run-properties",
      child: "del",
      reader: "paragraphParser#parseParagraphProperties",
    }),
    moveFrom: ownedElsewhere({
      container: "run-properties",
      child: "moveFrom",
      reader: "paragraphParser#parseParagraphProperties",
    }),
    moveTo: ownedElsewhere({
      container: "run-properties",
      child: "moveTo",
      reader: "paragraphParser#parseParagraphProperties",
    }),
    specVanish: ownedElsewhere({
      container: "run-properties",
      child: "specVanish",
      reader: "paragraphParser#parseParagraphProperties",
    }),
  },
  standalone: {},
} as const satisfies Record<RunPropertyOwner, Readonly<Partial<ChildHandlers<"run-properties">>>>;

/**
 * Read each declared child once; a repeat keeps its bytes.
 *
 * `EG_RPrBase` declares every property `maxOccurs="1"`, so a second
 * `<w:b w:val="0"/>` is not schema-valid and folio has no defined answer for
 * which one wins. The first occurrence is read, exactly as the hand-written
 * walk this replaced did, and the rest answer `CAPTURE` — they took nothing,
 * so they keep their bytes rather than overwriting a value the source stated
 * first or falling off the end of the walk.
 *
 * The name comes off the child rather than from the map key, so the guard
 * cannot name a property the handler does not read.
 */
const readEachChildOnce = (): ((
  read: (child: XmlElement) => typeof CAPTURE | void,
) => (child: XmlElement) => typeof CAPTURE | void) => {
  const taken = new Set<string>();
  return (read) => (child) => {
    const name = getLocalName(child.name);
    if (taken.has(name)) {
      return CAPTURE;
    }
    taken.add(name);
    return read(child);
  };
};

/**
 * `w:vertAlign` `baseline` is the reserved value that means "no vertical
 * offset", not a third offset beside `superscript` and `subscript`. It is
 * answered here because this module reads the slot, so a consumer deciding
 * whether a run sits on the baseline cannot drift from how it was parsed.
 */
export const isBaselineVertAlign = (vertAlign: TextFormatting["vertAlign"]): boolean =>
  vertAlign === "baseline";

/**
 * Read `w:rFonts` into `fontFamily`, resolving each theme slot to the font the
 * theme names when the element did not spell one out.
 *
 * @returns whether the element stated any font at all.
 */
const readFontFamily = (
  rFonts: XmlElement,
  theme: Theme | null,
  formatting: TextFormatting,
): boolean => {
  const fontFamily: NonNullable<TextFormatting["fontFamily"]> = {};
  const ascii = getAttribute(rFonts, "w", "ascii");
  if (ascii) {
    fontFamily.ascii = ascii;
  }
  const hAnsi = getAttribute(rFonts, "w", "hAnsi");
  if (hAnsi) {
    fontFamily.hAnsi = hAnsi;
  }
  const eastAsia = getAttribute(rFonts, "w", "eastAsia");
  if (eastAsia) {
    fontFamily.eastAsia = eastAsia;
  }
  const csFont = getAttribute(rFonts, "w", "cs");
  if (csFont) {
    fontFamily.cs = csFont;
  }
  const hint = narrowEnum(getAttribute(rFonts, "w", "hint"), FontHintSchema);
  if (hint) {
    fontFamily.hint = hint;
  }

  const asciiTheme = narrowEnum(getAttribute(rFonts, "w", "asciiTheme"), FontThemeSchema);
  if (asciiTheme) {
    fontFamily.asciiTheme = asciiTheme;
    // Also resolve the actual font name for convenience.
    if (theme && !fontFamily.ascii) {
      const resolved = resolveThemeFontRef(theme, asciiTheme);
      if (resolved) {
        fontFamily.ascii = resolved;
      }
    }
  }

  const hAnsiTheme = narrowEnum(getAttribute(rFonts, "w", "hAnsiTheme"), FontThemeSchema);
  if (hAnsiTheme) {
    fontFamily.hAnsiTheme = hAnsiTheme;
    if (theme && !fontFamily.hAnsi) {
      const resolved = resolveThemeFontRef(theme, hAnsiTheme);
      if (resolved) {
        fontFamily.hAnsi = resolved;
      }
    }
  }

  const eastAsiaTheme = narrowEnum(getAttribute(rFonts, "w", "eastAsiaTheme"), FontThemeSchema);
  if (eastAsiaTheme) {
    fontFamily.eastAsiaTheme = eastAsiaTheme;
    if (theme && !fontFamily.eastAsia) {
      const resolved = resolveThemeFontRef(theme, eastAsiaTheme);
      if (resolved) {
        fontFamily.eastAsia = resolved;
      }
    }
  }

  // OOXML spells this attribute all-lowercase, unlike its camelCase siblings.
  const csTheme = narrowEnum(getAttribute(rFonts, "w", "cstheme"), FontThemeSchema);
  if (csTheme) {
    fontFamily.csTheme = csTheme;
    if (theme && !fontFamily.cs) {
      const resolved = resolveThemeFontRef(theme, csTheme);
      if (resolved) {
        fontFamily.cs = resolved;
      }
    }
  }

  if (Object.keys(fontFamily).length === 0) {
    return false;
  }
  formatting.fontFamily = fontFamily;
  return true;
};

/** @returns whether `w:u` named an underline style the model admits. */
const readUnderline = (u: XmlElement, formatting: TextFormatting): boolean => {
  const style = narrowEnum(getAttribute(u, "w", "val"), UnderlineStyleSchema);
  if (!style) {
    return false;
  }
  formatting.underline = { style };
  const colorVal = getAttribute(u, "w", "color");
  const themeColor = getAttribute(u, "w", "themeColor");
  if (colorVal || themeColor) {
    formatting.underline.color = parseColorValue({
      rgb: colorVal,
      themeColor,
      themeTint: getAttribute(u, "w", "themeTint"),
      themeShade: getAttribute(u, "w", "themeShade"),
      element: u.name ?? "w:u",
    });
  }
  return true;
};

/** @returns whether `w:lang` named a tag for any of the three scripts. */
const readLanguage = (lang: XmlElement, formatting: TextFormatting): boolean => {
  const val = truncateLanguageTag(getAttribute(lang, "w", "val") || undefined);
  const eastAsia = truncateLanguageTag(getAttribute(lang, "w", "eastAsia") || undefined);
  const bidi = truncateLanguageTag(getAttribute(lang, "w", "bidi") || undefined);
  if (!val && !eastAsia && !bidi) {
    return false;
  }
  formatting.language = {
    ...(val ? { val } : {}),
    ...(eastAsia ? { eastAsia } : {}),
    ...(bidi ? { bidi } : {}),
  };
  return true;
};

/**
 * Parse run formatting properties (`w:rPr`).
 *
 * Every child the content model declares carries a decision: a handler that
 * reads it into {@link TextFormatting}, `ownedElsewhere` when the owner's
 * sibling record reads it, or `CAPTURE` for markup folio models nothing for.
 * A handler that looked and took nothing answers `CAPTURE` as well — `<w:sz/>`
 * states no size and `<w:highlight w:val="chartreuse"/>` states a value the
 * reader's enumeration does not admit, and neither can be decided by the
 * child's name. What no reader took goes to `TextFormatting.preserved` at its
 * schema ordinal, so a rebuild puts it back between the same two siblings.
 */
export function parseRunProperties(
  rPr: XmlElement | null,
  theme: Theme | null,
  owner: RunPropertyOwner,
): TextFormatting | undefined {
  if (!rPr) {
    return undefined;
  }

  const formatting: TextFormatting = {};
  const once = readEachChildOnce();

  const handlers: ChildHandlers<"run-properties"> = {
    // `EG_ParaRPrTrackChanges`. Only a paragraph mark declares these, and
    // there the paragraph's own record reads them; under a run the schema
    // declares no such child, so anything wearing the name is markup folio
    // keeps rather than a revision it understands.
    ins: CAPTURE,
    del: CAPTURE,
    moveFrom: CAPTURE,
    moveTo: CAPTURE,

    rStyle: once((child) => {
      const val = getAttribute(child, "w", "val");
      if (val) {
        formatting.styleId = val;
      }
      return keptUnless(Boolean(val));
    }),
    rFonts: once((child) => keptUnless(readFontFamily(child, theme, formatting))),

    // `CT_OnOff`: an empty element is the value `on`, so the tri-state reader
    // always takes something and none of these can refuse.
    b: once((child) => {
      formatting.bold = parseBooleanElement(child);
    }),
    bCs: once((child) => {
      formatting.boldCs = parseBooleanElement(child);
    }),
    i: once((child) => {
      formatting.italic = parseBooleanElement(child);
    }),
    iCs: once((child) => {
      formatting.italicCs = parseBooleanElement(child);
    }),
    caps: once((child) => {
      formatting.allCaps = parseBooleanElement(child);
    }),
    smallCaps: once((child) => {
      formatting.smallCaps = parseBooleanElement(child);
    }),
    strike: once((child) => {
      formatting.strike = parseBooleanElement(child);
    }),
    dstrike: once((child) => {
      formatting.doubleStrike = parseBooleanElement(child);
    }),
    outline: once((child) => {
      formatting.outline = parseBooleanElement(child);
    }),
    shadow: once((child) => {
      formatting.shadow = parseBooleanElement(child);
    }),
    emboss: once((child) => {
      formatting.emboss = parseBooleanElement(child);
    }),
    imprint: once((child) => {
      formatting.imprint = parseBooleanElement(child);
    }),
    noProof: once((child) => {
      formatting.noProof = parseBooleanElement(child);
    }),
    vanish: once((child) => {
      formatting.hidden = parseBooleanElement(child);
    }),
    rtl: once((child) => {
      formatting.rtl = parseBooleanElement(child);
    }),
    cs: once((child) => {
      formatting.cs = parseBooleanElement(child);
    }),

    /** Whether the run follows the section's document grid; no layout slot. */
    snapToGrid: CAPTURE,
    /** Web-view-only hiding, distinct from `w:vanish`; nothing reads it. */
    webHidden: CAPTURE,

    color: once((child) => {
      const color = parseColorValue({
        rgb: getAttribute(child, "w", "val"),
        themeColor: getAttribute(child, "w", "themeColor"),
        themeTint: getAttribute(child, "w", "themeTint"),
        themeShade: getAttribute(child, "w", "themeShade"),
        element: child.name ?? "w:color",
      });
      if (Object.keys(color).length === 0) {
        return CAPTURE;
      }
      formatting.color = color;
      return undefined;
    }),
    spacing: once((child) => {
      const val = parseNumericAttribute(child, "w", "val");
      if (val !== undefined) {
        formatting.spacing = val;
      }
      return keptUnless(val !== undefined);
    }),
    w: once((child) => {
      const val = parseHorizontalScalePercent(getAttribute(child, "w", "val"));
      if (val !== undefined) {
        formatting.scale = val;
      }
      return keptUnless(val !== undefined);
    }),
    kern: once((child) => {
      const val = parseNumericAttribute(child, "w", "val");
      if (val !== undefined) {
        formatting.kerning = val;
      }
      return keptUnless(val !== undefined);
    }),
    position: once((child) => {
      const val = parseNumericAttribute(child, "w", "val");
      if (val !== undefined) {
        formatting.position = val;
      }
      return keptUnless(val !== undefined);
    }),
    sz: once((child) => {
      const val = parseNumericAttribute(child, "w", "val");
      if (val !== undefined) {
        formatting.fontSize = val;
      }
      return keptUnless(val !== undefined);
    }),
    szCs: once((child) => {
      const val = parseNumericAttribute(child, "w", "val");
      if (val !== undefined) {
        formatting.fontSizeCs = val;
      }
      return keptUnless(val !== undefined);
    }),
    highlight: once((child) => {
      const val = narrowEnum(getAttribute(child, "w", "val"), HighlightColorSchema);
      if (val) {
        formatting.highlight = val;
      }
      return keptUnless(Boolean(val));
    }),
    u: once((child) => keptUnless(readUnderline(child, formatting))),
    effect: once((child) => {
      const val = narrowEnum(getAttribute(child, "w", "val"), TextEffectSchema);
      if (val) {
        formatting.effect = val;
      }
      return keptUnless(Boolean(val));
    }),
    /** A text border around the run; `BorderSpec` is a paragraph/table slot. */
    bdr: CAPTURE,
    shd: once((child) => {
      const shading = parseShading(child);
      if (shading) {
        formatting.shading = shading;
      }
      return keptUnless(shading !== undefined);
    }),
    /** Compress the run's text into a fixed width; no layout slot holds it. */
    fitText: CAPTURE,
    vertAlign: once((child) => {
      const val = getAttribute(child, "w", "val");
      if (val === "superscript" || val === "subscript" || val === "baseline") {
        formatting.vertAlign = val;
        return undefined;
      }
      return CAPTURE;
    }),
    em: once((child) => {
      const val = narrowEnum(getAttribute(child, "w", "val"), EmphasisMarkSchema);
      if (val) {
        formatting.emphasisMark = val;
      }
      return keptUnless(Boolean(val));
    }),
    lang: once((child) => keptUnless(readLanguage(child, formatting))),
    /** East Asian two-lines-in-one and horizontal-in-vertical typesetting. */
    eastAsianLayout: CAPTURE,
    /**
     * The run-in heading marker. On a paragraph mark the paragraph's record
     * reads it (`runInWithNext`); under a run the schema declares it and folio
     * models nothing for it.
     */
    specVanish: CAPTURE,
    /** The run is part of an equation; folio has no run-level maths slot. */
    oMath: CAPTURE,
    /**
     * The tracked property change. A run's record reads it into
     * `Run.propertyChanges`; a paragraph mark has no such record, so its
     * snapshot is kept whole rather than dropped with the revision.
     */
    rPrChange: CAPTURE,

    ...OWNED_BY_A_SIBLING_RECORD[owner],
  };

  const preserved = dispatchChildren({
    element: rPr,
    container: "run-properties",
    handlers,
    capturePosition: sequencePositions("run-properties", rPr),
  });
  if (preserved) {
    formatting.preserved = preserved;
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

function parseRunPropertyChanges(
  rPr: XmlElement | null,
  theme: Theme | null,
  currentFormatting: TextFormatting | undefined,
): RunPropertyChange[] | undefined {
  if (!rPr) {
    return undefined;
  }

  const changes = findChildren(rPr, "w", "rPrChange")
    .map((changeElement): RunPropertyChange => {
      const previousRPr = findChild(changeElement, "w", "rPr");
      const change: RunPropertyChange = {
        type: "runPropertyChange",
        info: parsePropertyChangeInfo(changeElement),
      };
      const previousFormatting = parseRunProperties(
        previousRPr,
        theme,
        RUN_PROPERTY_OWNERS.standalone,
      );
      if (previousFormatting) {
        change.previousFormatting = previousFormatting;
      }
      if (currentFormatting) {
        change.currentFormatting = currentFormatting;
      }
      return change;
    })
    .filter((change) => change.previousFormatting || change.currentFormatting);

  return changes.length > 0 ? changes : undefined;
}

/**
 * Parse text content (w:t)
 */
function parseTextContent(element: XmlElement): TextContent {
  // `xml:space` is not read: it is a property of the serialized text, and the
  // serializer re-derives it with `requiresXmlSpacePreserve`.
  return { type: "text", text: getTextContent(element) };
}

/**
 * Parse tab element (w:tab)
 */
function parseTabContent(): TabContent {
  return { type: "tab" };
}

function parsePositionalTabContent(element: XmlElement): TabContent {
  const positional: NonNullable<TabContent["positional"]> = {};
  const relativeTo = narrowEnum(
    getAttribute(element, "w", "relativeTo"),
    PositionalTabRelativeToSchema,
  );
  const alignment = narrowEnum(
    getAttribute(element, "w", "alignment"),
    PositionalTabAlignmentSchema,
  );
  const leader = narrowEnum(getAttribute(element, "w", "leader"), PositionalTabLeaderSchema);
  if (relativeTo !== undefined) {
    positional.relativeTo = relativeTo;
  }
  if (alignment !== undefined) {
    positional.alignment = alignment;
  }
  if (leader !== undefined) {
    positional.leader = leader;
  }
  return { type: "tab", positional };
}

/**
 * Parse break element (w:br)
 */
function parseBreakContent(element: XmlElement): BreakContent {
  const breakType = getAttribute(element, "w", "type");
  const clear = getAttribute(element, "w", "clear");

  const content: BreakContent = { type: "break" };

  if (breakType === "page" || breakType === "column" || breakType === "textWrapping") {
    content.breakType = breakType;
  }

  if (clear === "none" || clear === "left" || clear === "right" || clear === "all") {
    content.clear = clear;
  }

  return content;
}

/**
 * Parse symbol element (w:sym)
 */
function parseSymbolContent(element: XmlElement): SymbolContent {
  const font = getAttribute(element, "w", "font") ?? "";
  const char = getAttribute(element, "w", "char") ?? "";

  return {
    type: "symbol",
    font,
    char,
  };
}

/**
 * Parse footnote reference (w:footnoteReference)
 */
function parseFootnoteReference(element: XmlElement): NoteReferenceContent {
  const id = parseNumericAttribute(element, "w", "id") ?? 0;

  return {
    type: "footnoteRef",
    id,
  };
}

/**
 * Parse endnote reference (w:endnoteReference)
 */
function parseEndnoteReference(element: XmlElement): NoteReferenceContent {
  const id = parseNumericAttribute(element, "w", "id") ?? 0;

  return {
    type: "endnoteRef",
    id,
  };
}

/**
 * Parse field character (w:fldChar)
 */
function parseFieldChar(element: XmlElement): FieldCharContent {
  const fldCharType = getAttribute(element, "w", "fldCharType");

  let charType: FieldCharContent["charType"] = "begin";
  if (fldCharType === "separate") {
    charType = "separate";
  } else if (fldCharType === "end") {
    charType = "end";
  }

  const content: FieldCharContent = {
    type: "fieldChar",
    charType,
    ...parseFieldState(element),
  };
  // Self-numbering fields (LISTNUM, AUTONUM, …) often skip the `separate`
  // run and stash their last-rendered display value on a `<w:numberingChange
  // w:original="…"/>` child of the end fldChar instead. Capture it so the
  // paragraph parser can fall back to it when the field carries no result.
  const numberingChange = findChild(element, "w", "numberingChange");
  if (numberingChange) {
    const original = getAttribute(numberingChange, "w", "original");
    if (original !== null) {
      content.originalValue = original;
    }
  }
  return content;
}

/**
 * Parse instruction text (w:instrText)
 */
function parseInstrText(element: XmlElement): InstrTextContent {
  const text = getTextContent(element);

  return {
    type: "instrText",
    text,
  };
}

/**
 * Wrap raw XML the model cannot project at all.
 *
 * `DrawingContent` always carries an `Image`, so preservation-only content
 * gets a placeholder one. It names no relationship, which keeps
 * `classifyDrawingSafety` and the serializer on the replay path instead of
 * regenerating DrawingML from the placeholder.
 */
const preserveOnlyDrawing = (rawXml: string): DrawingContent => ({
  type: "drawing",
  image: {
    type: "image",
    size: { width: 0, height: 0 },
    wrap: { type: "inline" },
  },
  rawXml,
  rawXmlMode: DRAWING_RAW_XML_MODES.PRESERVE_ONLY,
});

/**
 * Parse drawing content (w:drawing).
 *
 * Dispatches by graphicData payload:
 * - `pic:pic` → image (handled by imageParser).
 * - `wps:wsp` with `<wps:txbx>` → text-box; returns null so
 *   `blockContentParser.enrichParagraphTextBoxes` can rebuild the shape
 *   with its inner paragraph content (it needs the style/numbering/theme
 *   context that is only available at the block parser level).
 * - `wps:wsp` without text body → generic shape; parsed via
 *   `shapeParser.parseShapeFromDrawing` into a `ShapeContent`.
 * - anything the model cannot project (a group the rasterizer declines,
 *   a diagram, a shape with unmodeled properties) → preservation-only raw XML.
 */
function parseDrawingContent(
  element: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
): DrawingContent | ShapeContent | null {
  const groupImage = parseGroupDrawing(element, rels ?? undefined, media ?? undefined);
  if (groupImage) {
    // The rasterized group is a preview, not a projection: it replays while
    // untouched, and an edit must block the save rather than regenerate one
    // child picture in place of the group.
    return {
      type: "drawing",
      image: groupImage,
      rawXml: captureVerbatimXml(element),
      rawImageFingerprint: imageRawXmlFingerprint(groupImage),
      rawXmlMode: DRAWING_RAW_XML_MODES.PREVIEW_ONLY,
    };
  }
  // A group without a preview still holds every child shape. parseImage would
  // model it as whichever picture it finds first inside the group, so an edit
  // to that projection would serialize one picture in place of the group.
  if (isGroupDrawing(element)) {
    return preserveOnlyDrawing(captureVerbatimXml(element));
  }
  const diagramImage = parseDiagramPreview(element, rels ?? undefined, media ?? undefined);
  if (diagramImage) {
    return {
      type: "drawing",
      image: diagramImage,
      rawXml: captureVerbatimXml(element),
      rawXmlMode: DRAWING_RAW_XML_MODES.PRESERVE_ONLY,
    };
  }
  if (shouldPreserveRawShapeDrawing(element)) {
    return preserveOnlyDrawing(captureVerbatimXml(element));
  }

  // Generic shapes (rect/ellipse/line/arrow/...) come in here as wps:wsp
  // with no text body. Text-box shapes are left for the block-content
  // post-pass; image drawings fall through to parseImage.
  const shape = parseShapeFromDrawing(element);
  if (shape) {
    return { type: "shape", shape };
  }

  const image = parseImage(element, rels ?? undefined, media ?? undefined);
  if (!image) {
    return null;
  }
  const rawXml = captureVerbatimXml(element);
  if (image.rId === undefined) {
    // No `a:blip`, so no picture: a chart, an OLE frame, or an anchor with no
    // `a:graphic` at all. `image` records the anchor's geometry for layout, but
    // the authored XML is the content, so the drawing replays unconditionally
    // rather than regenerating into a `pic:pic` it never held.
    return { type: "drawing", image, rawXml, rawXmlMode: DRAWING_RAW_XML_MODES.PRESERVE_ONLY };
  }
  return { type: "drawing", image, rawXml, rawImageFingerprint: imageRawXmlFingerprint(image) };
}

/**
 * Parse all content within a run element
 */
function parseRunContents(
  runElement: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  rootXmlns: Record<string, string> = {},
): RunContent[] {
  const contents: RunContent[] = [];
  const children = getChildElements(runElement);

  for (const child of children) {
    const localName = getLocalName(child.name);

    switch (localName) {
      case "t":
        // Text content
        contents.push(parseTextContent(child));
        break;

      case "tab":
        // Tab character
        contents.push(parseTabContent());
        break;

      case "ptab":
        contents.push(parsePositionalTabContent(child));
        break;

      case "br":
        // Line/page/column break
        contents.push(parseBreakContent(child));
        break;

      case "sym":
        // Symbol character
        contents.push(parseSymbolContent(child));
        break;

      case "footnoteReference":
        // Footnote reference
        contents.push(parseFootnoteReference(child));
        break;

      case "endnoteReference":
        // Endnote reference
        contents.push(parseEndnoteReference(child));
        break;

      case "fldChar":
        // Field character (begin/separate/end)
        contents.push(parseFieldChar(child));
        break;

      case "instrText":
        // Field instruction text
        contents.push(parseInstrText(child));
        break;

      case "softHyphen": {
        const softHyphen: SoftHyphenContent = { type: "softHyphen" };
        contents.push(softHyphen);
        break;
      }

      case "noBreakHyphen": {
        const noBreakHyphen: NoBreakHyphenContent = { type: "noBreakHyphen" };
        contents.push(noBreakHyphen);
        break;
      }

      case "drawing": {
        // Drawing/image
        const drawing = parseDrawingContent(child, rels, media);
        if (drawing) {
          contents.push(drawing);
        }
        break;
      }

      case "pict": {
        // Legacy VML inline picture (e.g. an old-format header logo). Resolve
        // it to the same drawing/image node a DrawingML image produces so it
        // renders through the existing image path; the original VML round-trips
        // verbatim via the drawing's rawXml.
        const vmlDrawing = parseVmlImageContent(child, rels, media, rootXmlns);
        if (vmlDrawing) {
          contents.push(vmlDrawing);
          break;
        }
        // A VML shape with no image relationship still paints; the serializer
        // emits DrawingML only, so raw replay is the sole way to keep it.
        if (shouldPreserveRawVmlPict(child)) {
          contents.push(
            preserveOnlyDrawing(captureVerbatimXml(cloneWithXmlnsDeclarations(child, rootXmlns))),
          );
        }
        break;
      }

      case "object": {
        // Embedded objects can carry a relationship-backed VML preview. Route
        // that preview through the image path while retaining the source XML.
        const objectPreview = parseVmlImageContent(child, rels, media, rootXmlns);
        // No preview resolved: nothing else claims a `w:object`, so the sink is
        // the only thing between the embedding and the floor.
        contents.push(objectPreview ?? preserveRunChild(child));
        break;
      }

      case "rPr":
        // Run properties - already handled separately
        break;

      case "commentReference":
        // Owned one level up: `parseParagraphContents` lifts it out of the run
        // into a sibling `commentReference` item and the comment serializer
        // re-emits its own run for it, so sinking it here would write the
        // reference twice.
        break;

      case "lastRenderedPageBreak":
        contents.push({ type: "renderedPageBreak" });
        break;

      case "cr": {
        // Carriage return - treat as line break
        const cr: BreakContent = { type: "break", breakType: "textWrapping" };
        contents.push(cr);
        break;
      }

      case "AlternateContent": {
        // mc:AlternateContent — folio cannot evaluate `mc:Requires`, so it
        // renders the Choice by default. A supported grouped Choice retains
        // authored group geometry, so it takes priority over a flattened
        // fallback. Otherwise, prefer a VML `w:pict` in the Fallback only when
        // it resolves to a real media part (not a textbox / empty pict / broken
        // relationship), then fall through to the Choice's DrawingML image.
        // The whole AlternateContent is kept on save so the Choice is not lost.
        const alternateChildren = getChildElements(child);
        const choiceEl = alternateChildren.find((el) => getLocalName(el.name) === "Choice");
        const fallbackEl = alternateChildren.find((el) => getLocalName(el.name) === "Fallback");
        const contentsBeforeAlternate = contents.length;
        const choiceTextBoxDrawing = choiceEl
          ? getChildElements(choiceEl).find(
              (element) => getLocalName(element.name) === "drawing" && isTextBoxDrawing(element),
            )
          : undefined;

        const groupedChoiceDrawing = choiceEl
          ? getChildElements(choiceEl).find(
              (element) =>
                getLocalName(element.name) === "drawing" &&
                findAllDeep(element, "wpg", "wgp").length > 0,
            )
          : undefined;
        if (groupedChoiceDrawing) {
          const groupedDrawing = parseDrawingContent(groupedChoiceDrawing, rels, media);
          // Widen the captured XML from the Choice to the whole
          // mc:AlternateContent so the Fallback replays too. Narrowing on the
          // mode keeps the result a preview-only member rather than a widened
          // object: a group that did not rasterize falls through instead.
          if (
            groupedDrawing?.type === "drawing" &&
            groupedDrawing.rawXmlMode === DRAWING_RAW_XML_MODES.PREVIEW_ONLY &&
            groupedDrawing.image.src
          ) {
            contents.push({ ...groupedDrawing, rawXml: captureVerbatimXml(child) });
            break;
          }
        }

        const fallbackPict = fallbackEl
          ? getChildElements(fallbackEl).find((el) => getLocalName(el.name) === "pict")
          : undefined;
        const fallbackVml =
          fallbackPict && !choiceTextBoxDrawing
            ? parseVmlImageContent(fallbackPict, rels, media, rootXmlns)
            : null;
        if (fallbackVml?.image.src) {
          fallbackVml.rawXml = captureVerbatimXml(cloneWithXmlnsDeclarations(child, rootXmlns));
          contents.push(fallbackVml);
          break;
        }

        const targetEl = selectAlternateContentBranch(child);
        if (targetEl) {
          for (const innerChild of getChildElements(targetEl)) {
            const innerName = getLocalName(innerChild.name);
            if (innerName === "drawing") {
              const innerDrawing = parseDrawingContent(innerChild, rels, media);
              // Keep package-referenced drawings even when the browser cannot render
              // the media. The serializer must preserve the relationship reference.
              if (innerDrawing) {
                if (
                  innerDrawing.type === "drawing" &&
                  (innerDrawing.rawXml !== undefined || !innerDrawing.image.src)
                ) {
                  innerDrawing.rawXml = captureVerbatimXml(child);
                }
                contents.push(innerDrawing);
              }
            } else if (innerName === "pict") {
              // A VML picture in the chosen Choice (no Fallback image present).
              const innerVml = parseVmlImageContent(innerChild, rels, media, rootXmlns);
              if (innerVml) {
                innerVml.rawXml = captureVerbatimXml(cloneWithXmlnsDeclarations(child, rootXmlns));
                contents.push(innerVml);
              }
            } else {
              // Parse one selected child at a time so text and preserved visual
              // carriers retain their original interleaving.
              contents.push(
                ...parseRunContents(
                  { ...targetEl, elements: [innerChild] },
                  rels,
                  media,
                  rootXmlns,
                ),
              );
            }
          }
        }

        // Every branch declined: without this the whole block, both
        // alternatives included, would leave no trace in the model and vanish
        // on save. A text-box drawing in either branch is excluded because
        // `enrichParagraphTextBoxes` rebuilds it as an editable shape, and
        // preserving it here too would emit the shape twice.
        const hasTextBoxBranch = [choiceEl, fallbackEl].some((branch) =>
          getChildElements(branch).some(
            (element) => getLocalName(element.name) === "drawing" && isTextBoxDrawing(element),
          ),
        );
        if (contents.length === contentsBeforeAlternate && !hasTextBoxBranch) {
          contents.push(
            preserveOnlyDrawing(captureVerbatimXml(cloneWithXmlnsDeclarations(child, rootXmlns))),
          );
        }
        break;
      }

      default:
        // Every remaining child goes to the verbatim sink, at its source
        // position: `w:ruby`, `w:contentPart`, `w:pgNum`, `w:annotationRef`,
        // the note markers `w:footnoteRef`/`w:endnoteRef`, the note separators,
        // the date placeholders, a foreign namespace, an element a later OOXML
        // revision adds. `w:rPr` is excluded above because
        // `parseRunProperties` reads the same element.
        contents.push(preserveRunChild(child));
        break;
    }
  }

  return contents;
}

/**
 * Parse a run element (w:r)
 *
 * @param node - The w:r XML element
 * @param _styles - Unread: a run's properties are parsed as the source wrote
 *   them, and style resolution happens above this parser
 * @param theme - Theme for resolving theme colors/fonts
 * @param rels - Relationship map for resolving image references
 * @param media - Media files map for image data
 * @returns Parsed Run object
 */
export function parseRun(
  node: XmlElement,
  _styles: StyleMap | null,
  theme: Theme | null,
  rels: RelationshipMap | null = null,
  media: Map<string, MediaFile> | null = null,
  rootXmlns: Record<string, string> = {},
): Run {
  const run: Run = {
    type: "run",
    content: [],
  };

  // Parse run properties (w:rPr)
  const rPr = findChild(node, "w", "rPr");
  if (rPr) {
    const formattingResult = parseRunProperties(rPr, theme, RUN_PROPERTY_OWNERS.run);
    if (formattingResult) {
      run.formatting = formattingResult;
    }
    const propertyChangesResult = parseRunPropertyChanges(rPr, theme, run.formatting);
    if (propertyChangesResult) {
      run.propertyChanges = propertyChangesResult;
    }
  }

  // Parse run contents (text, tabs, breaks, images, etc.). Accumulate the run's
  // own xmlns onto the inherited set so a captured VML `w:pict` replay resolves
  // any prefix scoped on the run itself.
  run.content = parseRunContents(node, rels, media, mergeXmlnsDeclarations(rootXmlns, node));

  // `CT_R` declares `w:rsidR`, `w:rsidDel` and `w:rsidRPr`, and `serializeRun`
  // writes no attribute of its own, so every attribute the source wrote is
  // the remainder.
  const remainder = attributeRemainder({ element: node, modelled: NO_MODELLED_ATTRIBUTES });
  if (remainder) {
    run.preservedAttributes = remainder;
  }

  return run;
}

/**
 * Get plain text from a run
 *
 * @param run - Parsed Run object
 * @returns Concatenated text content
 */
export function getRunText(run: Run): string {
  let text = "";

  for (const content of run.content) {
    if (content.type === "text") {
      text += content.text;
    } else if (content.type === "tab") {
      text += "\t";
    } else if (content.type === "break") {
      if (content.breakType === "page") {
        text += "\f"; // Form feed for page break
      } else {
        text += "\n";
      }
    } else if (content.type === "softHyphen") {
      text += "\u00AD"; // Soft hyphen Unicode
    } else if (content.type === "noBreakHyphen") {
      text += "\u2011"; // Non-breaking hyphen Unicode
    }
  }

  return text;
}

/**
 * Check if a run contains any actual content
 *
 * @param run - Parsed Run object
 * @returns true if run has visible content
 */
export function hasContent(run: Run): boolean {
  return run.content.length > 0;
}

/**
 * Check if a run contains a drawing/image
 *
 * @param run - Parsed Run object
 * @returns true if run contains an image
 */
export function hasImage(run: Run): boolean {
  return run.content.some((c) => c.type === "drawing");
}

/**
 * Get all images from a run
 *
 * @param run - Parsed Run object
 * @returns Array of Image objects
 */
export function getImages(run: Run): Image[] {
  return run.content.filter((c): c is DrawingContent => c.type === "drawing").map((c) => c.image);
}

/**
 * Check if a run is part of a complex field
 *
 * @param run - Parsed Run object
 * @returns true if run contains field characters
 */
export function hasFieldChar(run: Run): boolean {
  return run.content.some((c) => c.type === "fieldChar");
}

/**
 * Get field character type if present
 *
 * @param run - Parsed Run object
 * @returns Field character type or null
 */
export function getFieldCharType(run: Run): "begin" | "separate" | "end" | null {
  const fieldChar = run.content.find((c): c is FieldCharContent => c.type === "fieldChar");
  return fieldChar?.charType ?? null;
}
