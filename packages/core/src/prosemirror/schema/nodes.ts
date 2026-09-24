/**
 * ProseMirror Node Type Interfaces
 *
 * Type definitions for node attributes used by conversion modules,
 * extensions, and other consumers. NodeSpec definitions have moved
 * to the extension system (extensions/core/ and extensions/nodes/).
 */

import type { FloatingTableProperties, TableLook } from "../../types";
import type {
  OutlineLevel,
  ParagraphAlignment,
  ParagraphFormatting,
  ParagraphMarkChange,
  ParagraphPropertyChange,
  PositionalTab,
  PositionedBookmarkMarker,
  PreservedAttribute,
  PreservedMarkup,
  DisplacedByCustomXml,
  DrawingAnchor,
  DrawingRawXmlMode,
  ImageDocPrLink,
  FieldType,
  Hyperlink,
  LineSpacingRule,
  ImageFrameLocks,
  ImagePosition,
  PreviewDescriptor,
  ImageWrap,
  ImageWrapPolygon,
  EffectExtentSlots,
  WrapDistanceSlots,
  BorderSpec,
  ShadingProperties,
  TabStop,
  TextFormatting,
  CounterFormat,
  ListMarkerFormatting,
  TableBorders,
  TableCell,
  TableCellBorders,
  TableFormatting,
  TablePropertyExceptionFormatting,
  TablePropertyChange,
  TablePropertyExceptionChange,
  TablePreservedMarkup,
  TableRowFormatting,
  TableRowPropertyChange,
  TableCellFormatting,
  TableCellPropertyChange,
  TableWidthType,
  SectionProperties,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  SdtEndProperties,
  SdtProperties,
  SdtType,
  TrackedChangeInfo,
} from "../../types/document";
import type { OutlineStyleAttr } from "../../types/documentEnumValues";
import type { SpacingExplicit } from "../../types/formatting";
import type { ParagraphNumberingAttr } from "../numberingAttr";
import type { ParagraphDirection } from "../paragraphDirection";
import type { InlineWrapperLayer, TrackedChangeProvenance } from "./marks";

export type HardBreakAttrs =
  | {
      sourceElement: "cr";
      breakType?: never;
      clear?: never;
    }
  | {
      /** Absent and `br` both serialize as the standard break element. */
      sourceElement?: "br";
      /** Absent means the source `w:br` omitted `w:type`. */
      breakType?: "column" | "textWrapping";
      clear?: "none" | "left" | "right" | "all";
    };

/** Attributes preserved from an authored `<w:br w:type="page"/>`. */
export type PageBreakRunAttrs = {
  clear?: "none" | "left" | "right" | "all";
};

export type TabAttrs = {
  positional?: PositionalTab;
};

export type SymbolAttrs = {
  font: string;
  char: string;
};

/** A run child folio does not model, carried through the editor untouched. */
/**
 * Which container the captured markup came out of, and goes back into.
 *
 * `w:ruby` is a run child and has to be written back inside a `w:r`;
 * `w:permStart` is a paragraph child and the schema admits none inside a run,
 * so writing one there would produce a package Word repairs. One atom serves
 * both because the editor treats them identically — opaque, zero-width unless
 * the markup shows text, carrying whatever marks surround it — and only the
 * save path has to tell them apart.
 */
export const PRESERVED_XML_LEVELS = { run: "run", inline: "inline" } as const;

export type PreservedXmlLevel = (typeof PRESERVED_XML_LEVELS)[keyof typeof PRESERVED_XML_LEVELS];

/** A run or inline child folio does not model, carried through the editor untouched. */
export type PreservedXmlAttrs = {
  /** Replayable markup, as `captureVerbatimXml` wrote it. */
  xml: string;
  /** The visible text the markup puts on the line, empty when it shows none. */
  text: string;
  level: PreservedXmlLevel;
};

/** A block child folio does not model, carried through the editor untouched. */
export type PreservedBlockAttrs = {
  /** Replayable markup, as `captureVerbatimXml` wrote it. */
  xml: string;
};

export type BookmarkBoundaryAttrs =
  | {
      type: "start";
      id: number;
      name: string;
      colFirst?: number;
      colLast?: number;
      displacedByCustomXml?: DisplacedByCustomXml;
    }
  | {
      type: "end";
      id: number;
      displacedByCustomXml?: DisplacedByCustomXml;
    };

/**
 * The position of a comment's `w:commentReference`: the run that paints the
 * visible comment mark. It is authored data, not a consequence of where the
 * range ends, so the editor carries it as a node of its own.
 */
export type CommentReferenceAttrs = {
  commentId: number;
};

/**
 * Paragraph node attributes - maps to ParagraphFormatting
 */
export type ParagraphAttrs = {
  // Identity
  paraId?: string;
  textId?: string;
  reviewCarrier?: "terminal-table";

  // Alignment
  alignment?: ParagraphAlignment;
  /** Effective inherited alignment beneath any direct `w:jc` override. */
  alignmentFromStyle?: ParagraphAlignment;
  /** Effective East Asian line-edge policy (`w:kinsoku`). */
  kinsoku?: boolean;
  /** Effective hanging-punctuation policy (`w:overflowPunct`). */
  overflowPunctuation?: boolean;
  /** Effective paragraph opt-out from document automatic hyphenation. */
  suppressAutoHyphens?: boolean;

  // Spacing (in twips)
  spaceBefore?: number;
  spaceAfter?: number;
  lineSpacing?: number;
  lineSpacingRule?: LineSpacingRule;
  /**
   * Which line-spacing attributes came from this paragraph's own `w:spacing`.
   * `true` is accepted for editor states written before provenance became
   * field-specific; newly created state uses the exact discriminator.
   */
  lineSpacingExplicit?: boolean | "value" | "rule" | "both";
  snapToGrid?: boolean;
  spacingExplicit?: SpacingExplicit;
  /** Layout provenance: document defaults survive on empty paragraphs. */
  spacingFromDocDefaults?: SpacingExplicit;
  /** Layout provenance: resolved paragraph/table-style spacing survives on empty paragraphs. */
  spacingFromImplicitDefaultStyle?: SpacingExplicit;

  // Indentation (in twips)
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  hangingIndent?: boolean;

  // List properties
  /**
   * The stated `w:numPr` (17.3.1.19), as the model carries it. Minted by
   * `paragraphNumberingAttr`, so a model value cannot land here unconverted.
   */
  numPr?: ParagraphNumberingAttr;
  /**
   * The style-sourced numPr value when `numPr` came from the paragraph
   * style rather than direct formatting. While `numPr` still equals this,
   * fromProseDoc omits it from serialized formatting (writing it as direct
   * `<w:numPr>` would flip Word's level-indent precedence on reload). List
   * commands that change `numPr` make the values diverge, which re-enables
   * direct serialization — no explicit clearing needed.
   */
  numPrFromStyle?: ParagraphNumberingAttr;
  /** List number format (decimal, lowerRoman, upperRoman, etc.) for CSS counter styling */
  listNumFmt?: CounterFormat;
  /** Whether this is a bullet list */
  listIsBullet?: boolean;
  /** Whether this level uses legal numbering (parent placeholders render decimal). */
  listIsLegal?: boolean;
  /** Computed list marker text (e.g., "1.", "1.1.", "•") */
  listMarker?: string;
  /** Source numbering pattern used to compute markers for inserted list siblings. */
  listMarkerTemplate?: string;
  /** Whether the list marker is hidden (w:vanish on numbering level rPr) */
  listMarkerHidden?: boolean;
  /** Canonical numbering-level marker typography, in OOXML units. */
  listMarkerFormatting?: ListMarkerFormatting;
  /** Horizontal alignment of the marker around the paragraph's list anchor. */
  listMarkerAlignment?: "left" | "center" | "right";
  /**
   * `w:suff` (§17.9.25) — what follows the marker before body text.
   * `tab` (default) grows the marker to the next tab stop; `space` adds one
   * space glyph; `nothing` lets body text butt against the marker.
   */
  listMarkerSuffix?: "tab" | "space" | "nothing";
  /** `w:caps` on the numbering level rPr — render marker in upper case. */
  listMarkerAllCaps?: boolean;
  /**
   * Inline LISTNUM count carried by this paragraph. Each advances the
   * counter at `ilvl + 1` so a later sibling at that depth renders the
   * next marker letter.
   */
  listImplicitChildLevelAdvances?: number;
  /**
   * When the marker text contains a TAB separator, this column offset (in
   * twips) is where the slot after the tab should land — used to align an
   * inline LISTNUM "(a)" with the deeper level's marker column.
   */
  listMarkerSecondSlotOffsetTwips?: number;
  /** Number format for each level used by multi-level marker templates. */
  listLevelNumFmts?: CounterFormat[];
  /** Initial counter for each level used by multi-level marker templates. */
  listLevelStarts?: number[];
  /** Abstract numbering ID shared by numbering instances. */
  listAbstractNumId?: number;
  /** Numbering start override for this numId/level. */
  listStartOverride?: number;

  // Style reference
  styleId?: string;
  /** Imported built-in TOC entry level, resolved from the canonical style name. */
  _tableOfContentsLevel?: number;

  // Borders
  borders?: {
    top?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
    right?: BorderSpec;
    between?: BorderSpec;
    bar?: BorderSpec;
  };

  // Background/Shading
  shading?: ShadingProperties;

  // Tab stops
  tabs?: TabStop[];

  // Page break control
  pageBreakBefore?: boolean;
  /** Word's cached rendered-page-break marker; preserved for round-trip only. */
  renderedPageBreakBefore?: boolean;
  /** Internal import marker for a paragraph whose only run content is a hard page break. */
  _pageBreakCarrier?: boolean;
  /** Internal import marker for a hard page break after this paragraph's text. */
  _trailingPageBreak?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  /** Contextual spacing — suppress space between same-style paragraphs */
  contextualSpacing?: boolean;

  // Default text formatting for empty paragraphs (persists when navigating away)
  // Maps to OOXML pPr/rPr (paragraph's default run properties)
  defaultTextFormatting?: TextFormatting;
  /** Internal table-style run overlay, used to resolve body-run provenance. */
  _tableRunFormatting?: TextFormatting;

  /**
   * Base text direction as a discriminated union (undecided when absent). Maps
   * to the serialized OOXML `w:bidi` tri-state via `directionToBidi`; the
   * `source` discriminates an authoritative manual/import decision from a
   * re-evaluable auto-detected one. See `paragraphDirection.ts`.
   */
  direction?: ParagraphDirection | null;

  /** The stated `w:outlineLvl`, style-resolved: a heading level or body text. */
  outlineLevel?: OutlineLevel;

  // Bookmarks on this paragraph (for TOC anchors, cross-references)
  bookmarks?: { id: number; name: string }[];

  /** A `w:hyperlink` with no projectable leaf cannot be represented as a text mark.
   *  Preserve its metadata and any empty transparent-wrapper children at the
   *  paragraph boundary so an editor round trip does not drop either element. */
  _emptyHyperlinks?: {
    /** ProseMirror inline offset where the zero-width hyperlink appeared. */
    offset: number;
    href?: Hyperlink["href"];
    anchor?: Hyperlink["anchor"];
    tooltip?: Hyperlink["tooltip"];
    rId?: Hyperlink["rId"];
    target?: Hyperlink["target"];
    history?: Hyperlink["history"];
    docLocation?: Hyperlink["docLocation"];
    /** Empty transparent-wrapper nests authored as children of this link. */
    _docxEmptyWrapperStacks?: readonly (readonly InlineWrapperLayer[])[];
  }[];

  /**
   * Run-in heading flag: this paragraph's mark carries
   * `<w:specVanish/>` and the next paragraph should flow inline on
   * the same line (NVCA-style "6.11 Severability" → "The
   * invalidity..." merges). Layout consumes this in toFlowBlocks.
   */
  runInWithNext?: boolean;

  /** Original inline paragraph formatting from DOCX (pre-style-resolution).
   *  Used by fromProseDoc for lossless round-trip serialization. */
  _originalFormatting?: ParagraphFormatting;

  /**
   * The `w:pPr` the style cascade resolves to for this paragraph: document
   * defaults, the enclosing table style, then the `w:pStyle` chain. PM-only;
   * never serialized.
   *
   * The formatting attrs above hold the EFFECTIVE value, because that is what
   * the editor renders with. A save must write only what the paragraph states
   * itself, so it needs this companion to tell an inherited value from an
   * authored one — writing an inherited value back as direct `w:pPr` outranks
   * the style it came from, and a later edit to that style stops reaching the
   * paragraph. `TableCellAttrs._resolvedBorders` is the same device.
   */
  _resolvedFormatting?: ParagraphFormatting;

  /** Import-effective spacing baseline for HTML auto-spacing detection.
   *  PM-only; never serialized back into DOCX formatting. */
  _autospacingBase?: {
    before?: number | null;
    after?: number | null;
  };

  /**
   * The section that ends at this paragraph's mark, and the whole of that
   * state: the break type is a field of the record, derived through
   * `sectionBreakTypeOf`, never a second attr beside it.
   *
   * The record is shared by reference, never cloned. Among the paragraphs
   * holding one object the from-leg writes only the last, which is how a split
   * paragraph's two halves stay one section.
   */
  _sectionProperties?: SectionProperties;

  /**
   * Attributes the authored `w:p` carried and the model has no field for
   * (`w:rsidR` and its family), carried opaquely so an edit does not rewrite
   * the document's revision history.
   *
   * The remainder follows the record. ProseMirror copies a node's attrs to
   * both halves of a split, so `fromProseDoc` gives it to the first paragraph
   * that carries it and to no other: a paragraph the editor created never had
   * those attributes and must not inherit them from a neighbour.
   */
  _preservedAttributes?: PreservedAttribute[];

  /** Paragraph-property-change tracking entries (`w:pPrChange`).
   *  Preserved opaquely through ProseMirror — the editor does not surface
   *  them in UI today, but stripping them on every edit would corrupt the
   *  `w:pPrChange` history Word relies on for "show previous formatting"
   *  and for reverting an accepted property change. */
  _propertyChanges?: ParagraphPropertyChangeAttrs[];

  /** Paragraph-mark insertion / deletion (`<w:pPr><w:rPr><w:ins/>` /
   *  `<w:del/>`). Word emits this when the paragraph break itself was
   *  authored in track-changes mode — pressing Enter mid-paragraph
   *  produces an `ins`, Backspace-at-start / Delete-at-end produce a
   *  `del`. Stored as a discriminated union to mirror folio's
   *  `TrackedChangeWrapperType` model rather than two parallel attrs. */
  pPrMark?: ParagraphMarkChange;
  /**
   * Marks this whole paragraph as a *suggested* insertion (AI proposal). The
   * paragraph is dropped from serialized DOCX until accepted; accepting turns
   * it into a real inserted-paragraph tracked change (paragraph-mark `w:ins`).
   */
  _suggestedInsert?: SuggestedStructuralMarker | null;
};

/**
 * A partial write over a paragraph's attrs: every key carries its own attr
 * type, or the attr's absent state. Which spelling that is belongs to the
 * node spec's per-attr `default` (`null` for most, `undefined` for a few,
 * `alignmentFromStyle` among them), so both are admitted here.
 *
 * Derived from {@link ParagraphAttrs} rather than hand-listed, so an attr
 * added to the node spec is writable without a second edit, and a producer
 * that assembles a patch is held to each attr's type. `numPr` is the reason
 * it exists: a patch typed `Record<string, unknown>` can store a raw
 * `ParagraphNumberingOverride` there, which is exactly what the
 * {@link ParagraphNumberingAttr} brand makes impossible.
 */
export type ParagraphAttrsPatch = {
  [K in keyof ParagraphAttrs]?: ParagraphAttrs[K] | null | undefined;
};

/**
 * ProseMirror property-change attrs may also carry the editor's list-marker
 * snapshot fields alongside the canonical paragraph formatting fields.
 * Keeping that shape typed here lets layout consume validated attrs directly.
 */
export type ParagraphPropertyChangeAttrs = Omit<
  ParagraphPropertyChange,
  "info" | "previousFormatting" | "currentFormatting"
> & {
  info: ParagraphPropertyChange["info"] & {
    /** Editor-only proposal provenance; stripped or re-authored before OOXML serialization. */
    provenance?: TrackedChangeProvenance;
    suggestionId?: string | null;
  };
  previousFormatting?: Omit<ParagraphFormatting, "numPr"> & {
    // The attr, not the model field: what a list command records here is the
    // attr it replaced. `null` is the third state and means the paragraph
    // carried no numbering before the change.
    numPr?: ParagraphAttrs["numPr"] | null;
  } & Partial<
      Pick<
        ParagraphAttrs,
        | "listIsBullet"
        | "listIsLegal"
        | "listNumFmt"
        | "listMarker"
        | "listMarkerTemplate"
        | "listMarkerHidden"
        | "listMarkerFormatting"
        | "listMarkerAlignment"
        | "listMarkerSuffix"
        | "listMarkerAllCaps"
        | "listImplicitChildLevelAdvances"
        | "listMarkerSecondSlotOffsetTwips"
        | "listLevelNumFmts"
        | "listLevelStarts"
        | "listAbstractNumId"
        | "listStartOverride"
        | "lineSpacingExplicit"
        | "direction"
        | "_autospacingBase"
      >
    >;
  currentFormatting?: ParagraphFormatting;
};

/**
 * Image position for floating images (horizontal and vertical positioning)
 */
export type ImagePositionAttrs = {
  horizontal?: {
    relativeTo?: NonNullable<ImagePosition["horizontal"]["relativeTo"]>;
    posOffset?: number; // In EMU
    align?: NonNullable<ImagePosition["horizontal"]["alignment"]>;
  };
  vertical?: {
    relativeTo?: NonNullable<ImagePosition["vertical"]["relativeTo"]>;
    posOffset?: number; // In EMU
    align?: NonNullable<ImagePosition["vertical"]["alignment"]>;
  };
};

/**
 * The EMUs a drawing's pixel attributes were projected from, keyed by the
 * pixel attribute each one became.
 *
 * EMU → px → EMU does not land back on the same number: a size rounds to whole
 * pixels, a stroke or an inset to two. Without the authored value beside the
 * projected one, opening a document and saving it again moved every drawing
 * off the number its author wrote. `fromProseDoc` writes the authored EMU back
 * while the pixel attribute still projects from it, and converts the pixel
 * attribute once an editor command has moved it.
 *
 * A key carries `undefined` when the source authored no such value, so an
 * absent value stays absent rather than acquiring one the document never had.
 */
export type AuthoredEmuAttrs<Key extends string> = {
  readonly [K in Key]?: number | undefined;
};

/** The pixel attributes a drawing's `wp:wrap*` insets are projected into. */
type WrapDistanceAttr = "distTop" | "distBottom" | "distLeft" | "distRight";

/** The pixel attributes a text box's internal margins are projected into. */
type TextBoxMarginAttr = "marginTop" | "marginBottom" | "marginLeft" | "marginRight";

/**
 * Image node attributes
 */
export type ImageAttrs = {
  src: string;
  /**
   * How to draw a drawing that carries no image data, set instead of `src`.
   *
   * Carried on the attr rather than in a side table so an editor state is
   * still self-contained: a state that loses the table would lose the picture.
   * It is shape geometry in a serialized attr, which is a real cost in a
   * collaboration update, but the descriptor is bounded by the shape cap and
   * is three orders of magnitude smaller than the base64 raster this attr used
   * to hold for the same drawing.
   *
   * A state saved before this existed carries the raster in `src` and still
   * paints from it, so nothing stored needs migrating.
   */
  preview?: PreviewDescriptor;
  docPrName?: string;
  alt?: string;
  title?: string;
  /** Width in pixels (already converted from EMU) */
  width?: number;
  /** Height in pixels (already converted from EMU) */
  height?: number;
  rId?: string;
  /** Wrap type from DOCX: inline, square, tight, through, topAndBottom, behind, inFront */
  wrapType?: ImageWrap["type"];
  /** Display mode for CSS: inline (flows with text), float (left/right float), block (centered) */
  displayMode?: "inline" | "float" | "block";
  /** CSS float direction for floating images */
  cssFloat?: "left" | "right" | "none";
  /** CSS transform string, derived from the authored transform below. */
  transform?: string;
  /** Authored `a:xfrm@rot` in degrees; `null` when the drawing states none. */
  docxRotation?: number | null;
  /** Authored `a:xfrm@flipH`; `null` when the drawing states none. */
  docxFlipH?: boolean | null;
  /** Authored `a:xfrm@flipV`; `null` when the drawing states none. */
  docxFlipV?: boolean | null;
  /**
   * Opacity in [0, 1] from `<a:alphaModFix amt>`. Undefined / 1 means fully
   * opaque (no CSS `opacity` emitted). eigenpal #424.
   */
  opacity?: number;
  /** DrawingML `a:lum@bright` as a signed percentage in [-100, 100]. */
  brightness?: number;
  /** DrawingML `a:lum@contrast` as a signed percentage in [-100, 100]. */
  contrast?: number;
  /** Distance from text above (pixels) */
  distTop?: number;
  /** Distance from text below (pixels) */
  distBottom?: number;
  /** Distance from text left (pixels) */
  distLeft?: number;
  /** Distance from text right (pixels) */
  distRight?: number;
  /**
   * wp:srcRect crop fractions in [0, 1]. Carried through the editor so a
   * cropped image survives parse → edit → serialize. eigenpal #424
   * (image-crop subset).
   */
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  /**
   * `wp:effectExtent` reservation, in EMU — not pixels like `dist*` above.
   * Nothing renders it, and `emuToPixels` rounds (12700 EMU would land back
   * as 9525), so the editor carries the authored units untouched.
   */
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  /** Position for floating images (horizontal and vertical alignment) */
  position?: ImagePositionAttrs;
  /** `wp:anchor`'s own attributes, carried whole so a rebuild can restate them. */
  anchor?: DrawingAnchor;
  /**
   * Which element stated each wrap inset. The insets above are the value in
   * force; this is the slot each was authored on, without which a rebuild
   * collects them all onto the drawing.
   */
  wrapDistanceSlots?: WrapDistanceSlots;
  /**
   * Which element stated a `wp:effectExtent`. The drawing's is the reservation
   * in force; the wrap child's is the one the text flow is computed against,
   * and without the slots a rebuild writes neither back where it was authored.
   */
  wrapEffectExtentSlots?: EffectExtentSlots;
  /** `wp:wrapPolygon` as authored, in the path's own units. */
  wrapPolygon?: ImageWrapPolygon;
  /**
   * The image carries no information, so assistive technology skips it
   * (`wp:docPr`'s decorative extension). Not {@link hidden}: a decorative
   * image is still displayed.
   */
  decorative?: boolean;
  /** `wp:docPr @hidden`: the drawing is not displayed. */
  hidden?: boolean;
  /** `wp:docPr` extensions folio does not model, carried through verbatim. */
  docPrExtensions?: string[];
  /**
   * Authored `a:graphicFrameLocks`. Carried through the editor so a resize,
   * which forces the serializer to regenerate DrawingML, cannot silently
   * relax a lock the author set.
   */
  frameLocks?: ImageFrameLocks;
  /** Border width in pixels */
  borderWidth?: number;
  /** The EMUs behind `width`, `height`, `borderWidth` and the wrap insets. */
  _docxAuthoredEmu?: AuthoredEmuAttrs<"width" | "height" | "borderWidth" | WrapDistanceAttr>;
  /** Border color as CSS color string */
  borderColor?: string;
  /** Border style (CSS border-style value) */
  borderStyle?: string;
  /** Wrap text setting from DOCX (left, right, bothSides, largest) for round-trip */
  wrapText?: NonNullable<ImageWrap["wrapText"]>;
  /** Hyperlink URL for clickable image */
  hlinkHref?: string;
  /** DOCX relationship ID for the clickable image hyperlink */
  hlinkRId?: string;
  /**
   * `a:hlinkClick` as the source wrote it, and `a:hlinkHover` beside it. The
   * model has a field for the click target and none for the tooltip, the target
   * frame or the history flag, so an edit that did not touch the link has to
   * hand the element's own bytes back for the rebuild to replay.
   */
  hlinkClickSource?: ImageDocPrLink;
  hlinkHoverXml?: string;
  /**
   * `wp:docPr@id`, the drawing's own identity. Without it a rebuilt drawing is
   * renumbered from the serializer's counter on every edit.
   */
  docPrId?: string;
  /** Original OOXML for opaque/unsupported DOCX drawings. */
  _docxRawXml?: string;
  /** Raw XML preserved without an editable image projection. */
  _docxRawXmlMode?: DrawingRawXmlMode;
  /**
   * The fingerprint captured with `_docxRawXml`. A preview-only drawing must
   * compare against this rather than re-baseline on its own edited projection.
   */
  _docxRawImageFingerprint?: string;
  /** Embedded-object previews use their authored box as the exact line height. */
  _docxObjectPreview?: boolean;
  /**
   * The `w:rPr` of the run this atom came from. Inline atoms do not carry the
   * run's formatting marks (see `withRunBoundaryMarks`), so without this the
   * run properties of an embedded object, picture or shape are lost on save
   * (`content[].content[].formatting: object became absent` in the corpus
   * census).
   */
  _docxRunFormatting?: TextFormatting;
};

/**
 * Field node attributes
 */
export type FieldAttrs = {
  /** Field type: PAGE, NUMPAGES, DATE, MERGEFIELD, etc. */
  fieldType: FieldType;
  /** Full field instruction (e.g. "PAGE \\* MERGEFORMAT") */
  instruction: string;
  /** Current/cached display text */
  displayText: string;
  /** Imported cache that proved numbered REF resolution for this field. */
  _numberedRefBaseline?: string;
  /** Whether the field came from w:fldSimple or a complex fldChar range */
  fieldKind: "simple" | "complex";
  /** Field is locked */
  fldLock?: boolean;
  /** Field is dirty and should be recalculated by the host application */
  dirty?: boolean;
};

/**
 * Math equation node attributes
 */
export type MathAttrs = {
  /** Whether this is inline OMML or a block equation paragraph */
  display?: "inline" | "block";
  /** Raw OMML XML for round-trip preservation */
  ommlXml: string;
  /** Plain text fallback used by the editor and layout engine */
  plainText?: string;
};

/**
 * Structured document tag node attributes
 */
export type SdtAttrs = {
  /** SDT type */
  sdtType: SdtType;
  /** Alias (friendly name) */
  alias?: string;
  /** Tag (developer identifier) */
  tag?: string;
  /** Numeric `w:id/@w:val`. */
  id?: number;
  /** Lock setting */
  lock?: NonNullable<SdtProperties["lock"]>;
  /** Placeholder text */
  placeholder?: string;
  /** Whether showing placeholder */
  showingPlaceholder?: boolean;
  /** Date format for date controls */
  dateFormat?: string;
  /** ISO 8601 bound date value (`w:date@w:fullDate`). */
  dateValueISO?: string;
  /** Dropdown/combobox list items as JSON string */
  listItems?: string;
  /** Selected dropdown / comboBox value (`w:dropDownList@w:lastValue`). */
  dropdownLastValue?: string;
  /** Checkbox checked state */
  checked?: boolean;
  /**
   * The `w:sdtPr` children folio does not model, at their schema ordinal.
   *
   * The control keeps its own bytes through the editor the way a paragraph
   * keeps its attribute remainder: the serializer merges them back with the
   * modelled children in `CT_SdtPr` order.
   */
  _preserved?: PreservedMarkup;
  /** Captured `<w:sdtEndPr>…</w:sdtEndPr>` for round-trip replay. */
  rawEndPropertiesXml?: string;
  /**
   * `w:sdtEndPr` as a record, for a control the editor rebuilds and which has
   * no captured bytes left to replay. See `SdtProperties.endProperties`.
   */
  endProperties?: SdtEndProperties;
};

/**
 * Block-level structured document tag node attributes. Mirrors `SdtAttrs`
 * plus an optional numeric `w:id` and the verbatim `w:sdtPr`/`w:sdtEndPr`
 * strings captured by the parser for lossless round-trip.
 */
export type BlockSdtAttrs = {
  sdtType: SdtType;
  alias?: string;
  tag?: string;
  /** Numeric `w:id/@w:val`. */
  id?: number;
  lock?: NonNullable<SdtProperties["lock"]>;
  placeholder?: string;
  showingPlaceholder?: boolean;
  dateFormat?: string;
  /** ISO 8601 bound date value (`w:date@w:fullDate`). */
  dateValueISO?: string;
  /** Dropdown/combobox list items as JSON string. */
  listItems?: string;
  /** Selected dropdown / comboBox value (`w:dropDownList@w:lastValue`). */
  dropdownLastValue?: string;
  checked?: boolean;
  /**
   * Marker: source had empty `<w:sdtContent/>`. `toProseDoc` inserts a
   * filler paragraph to satisfy PM's `block+`; on save the converter
   * uses this flag to drop the filler instead of guessing from shape.
   */
  _originallyEmpty?: boolean;
  /**
   * The `w:sdtPr` children folio does not model, at their schema ordinal.
   *
   * The control keeps its own bytes through the editor the way a paragraph
   * keeps its attribute remainder: the serializer merges them back with the
   * modelled children in `CT_SdtPr` order.
   */
  _preserved?: PreservedMarkup;
  /** Captured `<w:sdtEndPr>…</w:sdtEndPr>` for round-trip replay. */
  rawEndPropertiesXml?: string;
  /** `w:sdtEndPr` as a record; see `SdtAttrs.endProperties`. */
  endProperties?: SdtEndProperties;
  /** Verbatim XML for sdt siblings before sdtContent (range markers). */
  rawSdtChildrenBeforeContent?: string;
  /** Verbatim XML for sdt siblings after sdtContent (range markers). */
  rawSdtChildrenAfterContent?: string;
};

/**
 * Shape node attributes
 */
export type ShapeAttrs = {
  /** Shape type preset */
  shapeType?: string;
  /** Preset geometry adjustments serialized as JSON. */
  geometryAdjustments?: string;
  /** Unique identifier */
  shapeId?: string;
  /** Authored non-visual drawing name (`wp:docPr@name` / `wps:cNvPr@name`) */
  shapeName?: string;
  /** Alt text for accessibility (`wp:docPr@descr`) */
  alt?: string;
  /** Authored non-visual drawing title (`wp:docPr@title`) */
  title?: string;
  /** Width in pixels */
  width?: number;
  /** Height in pixels */
  height?: number;
  /** Fill color as CSS color */
  fillColor?: string;
  /** Authored OOXML fill color, retained when it references a theme slot. */
  fillColorValue?: ShapeFill["color"];
  /** Fill type: none, solid, gradient, pattern, picture */
  fillType?: ShapeFill["type"];
  /** Gradient type: linear, radial, rectangular, path */
  gradientType?: NonNullable<ShapeFill["gradient"]>["type"];
  /** Gradient angle in degrees (for linear) */
  gradientAngle?: number;
  /** Gradient stops as JSON string: [{position, color}] */
  gradientStops?: string;
  /** Outline width in pixels */
  outlineWidth?: number;
  /** Outline color as CSS color */
  outlineColor?: string;
  /** Authored OOXML outline color, retained when it references a theme slot. */
  outlineColorValue?: ShapeOutline["color"];
  /** Outline dash style, or `"none"` for an explicit no-outline. */
  outlineStyle?: OutlineStyleAttr;
  /** Line cap */
  outlineCap?: NonNullable<ShapeOutline["cap"]>;
  /** Line join */
  outlineJoin?: NonNullable<ShapeOutline["join"]>;
  /** Head arrow/end marker */
  outlineHeadEnd?: NonNullable<ShapeOutline["headEnd"]>;
  /** Tail arrow/end marker */
  outlineTailEnd?: NonNullable<ShapeOutline["tailEnd"]>;
  /** CSS transform, derived from the authored transform below. */
  transform?: string;
  /** Authored `a:xfrm@rot` in degrees; `null` when the drawing states none. */
  docxRotation?: number | null;
  /** Authored `a:xfrm@flipH`; `null` when the drawing states none. */
  docxFlipH?: boolean | null;
  /** Authored `a:xfrm@flipV`; `null` when the drawing states none. */
  docxFlipV?: boolean | null;
  /** Display mode */
  displayMode?: "inline" | "float" | "block";
  /** CSS float */
  cssFloat?: "left" | "right" | "none";
  /** Wrap type */
  wrapType?: ImageWrap["type"];
  /** Wrap text setting from DOCX (left, right, bothSides, largest) for round-trip */
  wrapText?: NonNullable<ImageWrap["wrapText"]>;
  /** Distance from text above (pixels) */
  distTop?: number;
  /** Distance from text below (pixels) */
  distBottom?: number;
  /** Distance from text left (pixels) */
  distLeft?: number;
  /** Distance from text right (pixels) */
  distRight?: number;
  /** Position for floating shapes (horizontal and vertical alignment) */
  position?: ImagePositionAttrs;
  /** `wp:anchor`'s own attributes, carried whole so a rebuild can restate them. */
  anchor?: DrawingAnchor;
  /**
   * Which element stated each wrap inset. The insets above are the value in
   * force; this is the slot each was authored on, without which a rebuild
   * collects them all onto the drawing.
   */
  wrapDistanceSlots?: WrapDistanceSlots;
  /**
   * Which element stated a `wp:effectExtent`. The drawing's is the reservation
   * in force; the wrap child's is the one the text flow is computed against,
   * and without the slots a rebuild writes neither back where it was authored.
   */
  wrapEffectExtentSlots?: EffectExtentSlots;
  /** `wp:wrapPolygon` as authored, in the path's own units. */
  wrapPolygon?: ImageWrapPolygon;
  /** Shadow color as CSS color */
  shadowColor?: string;
  /** Shadow blur radius in pixels */
  shadowBlur?: number;
  /** Shadow X offset in pixels */
  shadowOffsetX?: number;
  /** Shadow Y offset in pixels */
  shadowOffsetY?: number;
  /** Glow color as CSS color */
  glowColor?: string;
  /** Glow radius in pixels */
  glowRadius?: number;
  /**
   * The `w:rPr` of the run this atom came from. Inline atoms do not carry the
   * run's formatting marks (see `withRunBoundaryMarks`), so without this the
   * run properties of an embedded object, picture or shape are lost on save
   * (`content[].content[].formatting: object became absent` in the corpus
   * census).
   */
  _docxRunFormatting?: TextFormatting;
  /** The EMUs behind `width`, `height`, `outlineWidth` and the wrap insets. */
  _docxAuthoredEmu?: AuthoredEmuAttrs<"width" | "height" | "outlineWidth" | WrapDistanceAttr>;
};

/**
 * Text box node attributes
 */
export const TEXT_BOX_TEXT_BODY_CONTENT_STATE_TYPES = Object.freeze([
  "source-empty",
  "authored",
] as const);

export type TextBoxTextBodyContentState = {
  [Type in (typeof TEXT_BOX_TEXT_BODY_CONTENT_STATE_TYPES)[number]]: {
    readonly type: Type;
  };
}[(typeof TEXT_BOX_TEXT_BODY_CONTENT_STATE_TYPES)[number]];

export type TextBoxAttrs = {
  /** Width in pixels */
  width?: number;
  /** Height in pixels */
  height?: number;
  /** The EMUs behind the size, `outlineWidth`, the wrap insets and the margins. */
  _docxAuthoredEmu?: AuthoredEmuAttrs<
    "width" | "height" | "outlineWidth" | WrapDistanceAttr | TextBoxMarginAttr
  >;
  /** Text fitting behavior */
  autoFit?: ShapeTextBody["autoFit"];
  /** Authored DrawingML WordArt metadata. */
  wordArt?: ShapeTextBody["wordArt"];
  /** Horizontal text wrapping inside the box */
  textWrap?: ShapeTextBody["textWrap"];
  /** Unique identifier */
  textBoxId?: string;
  /** Authored non-visual drawing name (`wp:docPr@name` / `wps:cNvPr@name`) */
  textBoxName?: string;
  /** Alt text for accessibility (`wp:docPr@descr`) */
  alt?: string;
  /** Authored non-visual drawing title (`wp:docPr@title`) */
  title?: string;
  /** Fill color as CSS color */
  fillColor?: string;
  /** Outline width in pixels */
  outlineWidth?: number;
  /** Outline color as CSS color */
  outlineColor?: string;
  /** Outline dash style, or `"none"` for an explicit no-outline. */
  outlineStyle?: OutlineStyleAttr;
  /** DrawingML rotation and/or flips, serialized as CSS transform functions. */
  transform?: string;
  /** Authored `a:xfrm@rot` in degrees; `null` when the drawing states none. */
  docxRotation?: number | null;
  /** Authored `a:xfrm@flipH`; `null` when the drawing states none. */
  docxFlipH?: boolean | null;
  /** Authored `a:xfrm@flipV`; `null` when the drawing states none. */
  docxFlipV?: boolean | null;
  /** Internal margin top in pixels */
  marginTop?: number;
  /** Internal margin bottom in pixels */
  marginBottom?: number;
  /** Internal margin left in pixels */
  marginLeft?: number;
  /** Internal margin right in pixels */
  marginRight?: number;
  /** Vertical text alignment */
  verticalAlign?: string;
  /** Display mode */
  displayMode?: "inline" | "float" | "block";
  /** CSS float direction */
  cssFloat?: "left" | "right" | "none";
  /** Wrap type */
  wrapType?: ImageWrap["type"];
  /** OOXML wrapText direction for anchored text boxes (eigenpal #474). */
  wrapText?: "bothSides" | "left" | "right" | "largest";
  /** Wrap distance from top edge, in pixels (OOXML distT, EMU-converted). */
  distTop?: number;
  /** Wrap distance from bottom edge, in pixels. */
  distBottom?: number;
  /** Wrap distance from left edge, in pixels. */
  distLeft?: number;
  /** Wrap distance from right edge, in pixels. */
  distRight?: number;
  /** Position for floating/anchored text boxes */
  position?: ImagePositionAttrs;
  /** `wp:anchor`'s own attributes, carried whole so a rebuild can restate them. */
  anchor?: DrawingAnchor;
  /**
   * Which element stated each wrap inset. The insets above are the value in
   * force; this is the slot each was authored on, without which a rebuild
   * collects them all onto the drawing.
   */
  wrapDistanceSlots?: WrapDistanceSlots;
  /**
   * Which element stated a `wp:effectExtent`. The drawing's is the reservation
   * in force; the wrap child's is the one the text flow is computed against,
   * and without the slots a rebuild writes neither back where it was authored.
   */
  wrapEffectExtentSlots?: EffectExtentSlots;
  /** `wp:wrapPolygon` as authored, in the path's own units. */
  wrapPolygon?: ImageWrapPolygon;
  /** Original DOCX placement hint for save-path reconstruction. */
  _docxPlacement?: "standalone" | "inlineWithPrevious";
  /** Original DOCX paragraph group for standalone text-box reconstruction. */
  _docxGroupId?: string;
  /** Inline anchor linking this block node to its source run position. */
  _docxAnchorId?: string;
  /**
   * Ownership of the schema-required placeholder paragraph. A source text
   * body with no children needs one paragraph while it is editable, but that
   * paragraph is not authored document content.
   */
  _docxTextBodyContentState: TextBoxTextBodyContentState;
  /** Original run-level revision wrapper for save-path reconstruction. */
  _docxTrackedChange?:
    | { type: "insertion"; info: TrackedChangeInfo }
    | { type: "deletion"; info: TrackedChangeInfo }
    | { type: "moveFrom"; info: TrackedChangeInfo }
    | { type: "moveTo"; info: TrackedChangeInfo };
  /** Original inline content-control ancestry for save-path reconstruction. */
  _docxInlineSdts?: SdtAttrs[];
  /**
   * The attribute remainder of the `w:p` this node was lifted out of.
   *
   * A paragraph whose only content was an anchored drawing has no paragraph
   * node in the editor: this node stands in for it, so it carries the host's
   * remainder the way `ParagraphAttrs._preservedAttributes` carries a
   * paragraph's own. Only a `"standalone"` placement has a host to speak for;
   * an `"inlineWithPrevious"` text box sits in a paragraph that is projected
   * itself and keeps its own.
   */
  _preservedAttributes?: PreservedAttribute[];
  /**
   * The paragraph properties of the `w:p` this node was lifted out of.
   *
   * An inline drawing is run content of its paragraph, so that paragraph's
   * spacing, indentation and alignment still place the box and must be
   * written back around it. Only a `"standalone"` placement has a host to
   * speak for. The host's identity (`paraId`, source token, attribute
   * remainder) is not carried here: a copied node must not duplicate it.
   */
  _docxHostParagraph?: Omit<ParagraphAttrs, "paraId" | "textId" | "_preservedAttributes">;
};

/** Internal inline position marker for an extracted text box block. */
export type TextBoxAnchorAttrs = {
  anchorId: string;
};

/**
 * Table node attributes
 */
export type TableAttrs = {
  /** Table style ID */
  styleId?: string;
  /** Table width (in twips) */
  width?: number;
  /** Table width type ('auto', 'pct', 'dxa') */
  widthType?: TableWidthType;
  /** Table placement (`w:tblPr/w:jc`) */
  justification?: NonNullable<TableFormatting["justification"]>;
  /** Column widths (in twips) from w:tblGrid */
  columnWidths?: number[];
  /** Floating table properties (w:tblpPr) */
  floating?: FloatingTableProperties;
  /** Default cell margins for the table (w:tblCellMar), in twips */
  cellMargins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  /** Table look flags for conditional formatting (w:tblLook) */
  look?: TableLook;
  /** Table-level borders (w:tblBorders) — full BorderSpec per side */
  borders?: TableBorders;
  /**
   * Effective default cell margins after the table-style cascade. PM-only;
   * never serialized. What `cellMargins` resolves to when the table declares
   * none of its own, so a save can tell a value a style supplied from one the
   * table states — and stop writing the style's into the table's `w:tblPr`.
   */
  _resolvedCellMargins?: TableAttrs["cellMargins"];
  /** Effective table indent after style resolution. PM-only; never serialized. */
  _resolvedIndent?: NonNullable<TableFormatting["indent"]>;
  /** Style-derived table justification fallback. PM-only; never serialized. */
  _resolvedJustification?: NonNullable<TableFormatting["justification"]>;
  /** Effective table direction after style resolution. PM-only; never serialized. */
  _resolvedBidi?: boolean;
  /** Original table formatting from DOCX for lossless round-trip serialization */
  _originalFormatting?: TableFormatting;
  /** Tracked table property changes (w:tblPrChange) for round-trip + accept/reject */
  tblPrChange?: TablePropertyChange[];
  /**
   * Bookmark markers the authored element held beside its own children, with
   * the position each was read at.
   *
   * There is no node for a boundary between two rows or two cells — a row's
   * children are cells — so the markers ride the record by reference, the way
   * `_preservedAttributes` does and for the same reason: a copy the editor
   * made shares the array, an authored record holds its own, and only the
   * first in document order keeps it.
   */
  _bookmarks?: PositionedBookmarkMarker[];
  /**
   * Markup the authored `w:tbl` carried beside its rows — a bookmark or
   * permission boundary, a proofing error, a custom-XML revision range —
   * with the row count that places it back between the same two rows.
   *
   * Carried by reference for the reason `_preservedAttributes` is: the sink
   * follows the record it was authored on, so a table the editor created has
   * none and a copy does not inherit one.
   */
  _preserved?: TablePreservedMarkup;
  /**
   * Marks this whole table as a *suggested* insertion (AI proposal). The table
   * is dropped from serialized DOCX until accepted; because OOXML has no tracked
   * whole-table-insert primitive, accepting applies it directly.
   */
  _suggestedInsert?: SuggestedStructuralMarker | null;
};

/**
 * Table row attributes
 */
export type TableRowAttrs = {
  /** Row height (in twips) */
  height?: number;
  /** Height rule ('auto', 'exact', 'atLeast') */
  heightRule?: NonNullable<TableRowFormatting["heightRule"]>;
  /** Is header row */
  isHeader?: boolean;
  /** Whether the row is hidden (`w:hidden`) */
  hidden?: boolean;
  /** Style-derived row justification fallback. PM-only; never serialized. */
  _resolvedJustification?: NonNullable<TableRowFormatting["justification"]>;
  /** Original row formatting from DOCX for lossless round-trip serialization */
  _originalFormatting?: TableRowFormatting;
  /**
   * The table properties the authored row overrides (`w:tblPrEx`), carried
   * whole for the reason `_originalFormatting` is: the editor surfaces none of
   * them, and rebuilding the element from the handful of attrs it does surface
   * would drop the rest.
   */
  _tablePropertyExceptions?: TablePropertyExceptionFormatting;
  /** Tracked row property changes (w:trPrChange) for round-trip + accept/reject */
  trPrChange?: TableRowPropertyChange[];
  /** Tracked changes to the property exceptions (w:tblPrExChange), carried opaquely */
  tblPrExChange?: TablePropertyExceptionChange[];
  /**
   * Attributes the authored `w:tr` carried and the model has no field for
   * (`w:rsidR`, `w:rsidDel`, `w:rsidTr`, `w:rsidRPr`), carried opaquely for
   * the reason `ParagraphAttrs._preservedAttributes` gives.
   */
  _preservedAttributes?: PreservedAttribute[];
  /**
   * Bookmark markers the authored element held beside its own children, with
   * the position each was read at.
   *
   * There is no node for a boundary between two rows or two cells — a row's
   * children are cells — so the markers ride the record by reference, the way
   * `_preservedAttributes` does and for the same reason: a copy the editor
   * made shares the array, an authored record holds its own, and only the
   * first in document order keeps it.
   */
  _bookmarks?: PositionedBookmarkMarker[];
  /**
   * The row-level content controls (`CT_SdtRow`) this row sits inside,
   * outermost first.
   *
   * The record travels on the row because a table's children are rows and
   * ProseMirror has no node to spare for a wrapper that is not one. Splitting
   * or moving a row takes its controls with it, which an index between rows
   * would not; the save re-opens one wrapper per run of consecutive rows that
   * name the same control. See `TableRow.contentControls`.
   */
  contentControls?: SdtProperties[];
  /**
   * Markup the authored `w:tr` carried beside its cells, with the cell count
   * that places it back between the same two cells. Follows the record the
   * same way `TableAttrs._preserved` does.
   */
  _preserved?: TablePreservedMarkup;
} & (
  | {
      /**
       * Tracked structural row insertion (w:trPr/w:ins). A `"suggested"`
       * provenance marks a whole-row AI proposal: stripped from serialized DOCX
       * until accepted, at which point it becomes a real (user) `trIns`. Rows use
       * this, NOT `_suggestedInsert` (that marker is paragraph/table only).
       */
      trIns: {
        revisionId: number;
        author: string;
        date?: string | null;
        utcDate?: string | null;
        initials?: string | null;
        provenance?: TrackedChangeProvenance;
        suggestionId?: string | null;
      };
      trDel?: never;
    }
  | {
      trIns?: never;
      /** Tracked structural row deletion (w:trPr/w:del). */
      trDel: {
        revisionId: number;
        author: string;
        date?: string | null;
        utcDate?: string | null;
        initials?: string | null;
        provenance?: TrackedChangeProvenance;
        suggestionId?: string | null;
      };
    }
  | {
      trIns?: never;
      trDel?: never;
    }
);

/**
 * A whole-node *suggested* insertion marker (AI proposal) carried on a
 * paragraph or table node. The node is dropped entirely from serialized DOCX
 * until accepted; accepting converts it to a real tracked change (paragraph
 * mark `w:ins`) or, where OOXML has no tracked representation (whole table),
 * applies it directly.
 */
export type SuggestedStructuralMarker = {
  suggestionId: string;
  revisionId: number;
  author: string;
  date?: string | null;
  initials?: string | null;
};

/**
 * Table cell attributes
 */
export type TableCellAttrs = {
  /** Authored OOXML cell identifier. */
  _docxCellId?: TableCell["id"];
  /** Column span */
  colspan: number;
  /** Row span */
  rowspan: number;
  /**
   * A non-authored cell which occupies a `w:gridBefore` or `w:gridAfter`
   * slot so ProseMirror can retain a rectangular table map. It is invisible
   * in the editor and omitted when projecting back to OOXML.
   */
  _omittedGridSlot?: "before" | "after";
  /** Column widths for prosemirror-tables resizing (array of pixel widths) */
  colwidth?: number[] | null;
  /** Cell width (in twips) */
  width?: number;
  /** Cell width type */
  widthType?: TableWidthType;
  /**
   * The preferred width the cell itself states, absent when it states none.
   *
   * `width` is the width the cell *renders* at, which the table resolves from
   * its grid when the cell declares no `w:tcW`, so a save that read it would
   * give every cell in the document a preferred width its author never wrote.
   * The companion to `width` that `_resolvedBorders` is to `borders`, and the
   * one the save leg writes `w:tcW` from. A command that moves a cell's width
   * states one: `mergeTableCellAttrs` records it for every command that goes
   * through it.
   */
  _authoredWidth?: { value: number; type: TableWidthType };
  /** Vertical alignment */
  verticalAlign?: "top" | "center" | "bottom";
  /** Background color (RGB hex) */
  backgroundColor?: string;
  /** Resolved source color. PM-only; distinguishes theme rendering from a user override. */
  _resolvedBackgroundColor?: string;
  /** OOXML text direction (e.g. 'tbRl', 'btLr') */
  textDirection?: NonNullable<TableCellFormatting["textDirection"]>;
  /** No text wrapping in cell */
  noWrap?: boolean;
  /** Effective end-of-cell marker suppression (`w:hideMark`). */
  hideMark?: boolean;
  /** Cell borders — full BorderSpec per side (style, color, size) */
  borders?: TableCellBorders;
  /**
   * Effective borders after the table and table-style cascade. PM-only; never
   * serialized. The companion to `borders` that `_resolvedBackgroundColor` is
   * to `backgroundColor`: what the cell renders with when it declares none of
   * its own, so a save can tell a border a style supplied from one the cell
   * states.
   */
  _resolvedBorders?: TableCellBorders;
  /** Cell margins/padding in twips per side */
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  /** Effective margins after the cascade. PM-only; never serialized. */
  _resolvedMargins?: { top?: number; bottom?: number; left?: number; right?: number };
  /** Original cell formatting from DOCX for lossless round-trip serialization */
  _originalFormatting?: TableCellFormatting;
  /** Tracked cell property changes (w:tcPrChange) for round-trip + accept/reject */
  tcPrChange?: TableCellPropertyChange[];
  /** Tracked cell structural revision for round-trip + accept/reject. */
  cellMarker?:
    | {
        kind: "ins" | "del";
        info: {
          revisionId: number;
          author: string;
          date?: string | null;
          utcDate?: string | null;
          initials?: string | null;
          /** `"suggested"` marks this as an AI proposal (stripped until accepted). */
          provenance?: TrackedChangeProvenance;
          suggestionId?: string | null;
        };
      }
    | {
        kind: "merge";
        /**
         * Merge markers resolve via cell-merge resolution and are never
         * produced in suggested mode (cell merge/split reports
         * `unsupportedMode`), so they carry no suggestion provenance.
         */
        info: {
          revisionId: number;
          author: string;
          date?: string | null;
          utcDate?: string | null;
          initials?: string | null;
        };
        verticalMerge?: "continue" | "rest";
        verticalMergeOriginal?: "continue" | "rest";
      };
  /**
   * The cell-level content controls (`CT_SdtCell`) this cell sits inside,
   * outermost first. The row's twin, one level down.
   */
  contentControls?: SdtProperties[];
  /** Preserve a DOCX vMerge restart even when PM cannot model it as a rowspan. */
  _preserveVMergeRestart?: boolean;
  /** Original DOCX vMerge continuation cells skipped into this PM rowspan. */
  _docxVMergeContinuationCells?: unknown;
};
