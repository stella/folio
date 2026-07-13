/**
 * Multi-engine DOCX rendering comparison: shared data contract.
 *
 * All geometry is expressed in PDF points (1pt = 1/72in), page-relative with
 * the origin at the page's top-left corner. Reference-renderer geometry comes
 * from `mutool draw -F stext` over an exported PDF; folio-side geometry comes
 * from the painted playground DOM (CSS px at 96dpi, converted with the px-to-pt
 * factor 72/96 after normalising against the page element width so playground
 * zoom cannot skew coordinates).
 */

export type ReferenceRendererId = "libreoffice" | "word";

export type ReferenceRendererInfo = {
  id: ReferenceRendererId;
  displayName: string;
  version?: string;
};

export type Region = "body" | "header" | "footer" | "footnote" | "unknown";

export type LineBox = {
  /** Raw extracted text of the visual line. */
  text: string;
  /** NFC-normalised, whitespace-collapsed, trimmed text used for alignment. */
  normText: string;
  /** Left edge of the line's ink/text start, in pt from the page left. */
  xPt: number;
  /** Top edge of the line box, in pt from the page top. PDF-extracted boxes
   * are ink bounds and folio-side boxes are line-height boxes, so absolute
   * yPt is only compared after per-page median-offset correction. */
  yPt: number;
  widthPt: number;
  heightPt: number;
  /** Primary font of the line (first run), when known. */
  fontName?: string;
  fontSizePt?: number;
  region: Region;
  /** DOM visual container, used to keep adjacent table cells distinct. */
  visualGroup?: string;
  /** Page-local identity shared by segments extracted from one logical line. */
  logicalLineGroup?: string;
};

export type PageGeom = {
  /** 1-based page number. */
  number: number;
  widthPt: number;
  heightPt: number;
  /** Lines sorted by (yPt, xPt). */
  lines: LineBox[];
};

export type DocGeom = {
  source: ReferenceRendererId | "folio";
  /** Absolute path of the source .docx. */
  file: string;
  pages: PageGeom[];
  /** Extractor provenance: renderer version, mutool version, playground URL, px-to-pt scale, ... */
  meta: Record<string, string>;
};

/** A divergence between an external reference renderer and folio. */
export type Divergence =
  | { kind: "page-count"; reference: number; folio: number }
  /** Same line of text landed on different pages. */
  | { kind: "pagination"; text: string; referencePage: number; folioPage: number }
  /** One reference line corresponds to 2+ folio lines, or vice versa. */
  | { kind: "line-break"; page: number; referenceTexts: string[]; folioTexts: string[] }
  /** Line present in reference output but never matched in folio. */
  | { kind: "missing-line"; page: number; text: string }
  /** Line present in folio output but never matched in the reference. */
  | { kind: "extra-line"; page: number; text: string }
  /** Matched line starts at a different horizontal position. */
  | { kind: "x-drift"; page: number; text: string; deltaPt: number }
  /** Matched line's vertical residual after per-page median-offset correction. */
  | { kind: "y-drift"; page: number; text: string; residualPt: number }
  /** Matched line's ink width differs (font metric / measurement divergence). */
  | { kind: "width-drift"; page: number; text: string; deltaPt: number }
  /** Lines aligned positionally but their text content differs. */
  | { kind: "text-mismatch"; page: number; referenceText: string; folioText: string };

export type DivergenceKind = Divergence["kind"];

/** A divergence enriched with the OOXML features active at its location. */
export type AttributedDivergence = {
  divergence: Divergence;
  /** Feature tags of the containing paragraph (e.g. "table", "tab", "numbering",
   * "float-anchor", "justify", "spacing-atLeast", "field", "cjk"), plus
   * doc-level tags prefixed "doc:" when no paragraph could be matched. */
  features: string[];
};

export type ComparisonTolerances = {
  /** Max |x delta| in pt before an x-drift divergence is reported. */
  xPt: number;
  /** Max |y residual| in pt (after median-offset correction) before y-drift. */
  yResidualPt: number;
  /** Max |width delta| in pt before width-drift. */
  widthPt: number;
  /** Widths may also pass on relative terms: |delta| <= widthRelative * referenceWidth. */
  widthRelative: number;
};

export type ParityResult = {
  file: string;
  /** 0..1: fraction of reference lines matched to a folio line on the same page
   * with all geometric deltas within tolerance. */
  score: number;
  referencePages: number;
  folioPages: number;
  totalReferenceLines: number;
  matchedLines: number;
  /** Systematic per-doc vertical offset (median of per-line y deltas);
   * informational, already subtracted from y-drift residuals. */
  medianYOffsetPt: number;
  divergences: Divergence[];
};

export type FeatureAttributedResult = ParityResult & {
  attributed: AttributedDivergence[];
  /** Doc-level feature tags (e.g. "multi-column", "landscape", "footnotes"). */
  docFeatures: string[];
};

/** A cross-corpus cluster: one feature co-occurring with one divergence kind. */
export type Cluster = {
  kind: DivergenceKind;
  feature: string;
  /** Divergences of this kind carrying this feature, across the corpus. */
  count: number;
  /** Docs contributing to the cluster. */
  docs: string[];
  /** Over-representation: P(feature | divergence of this kind) / P(feature | any line). >1 means the feature attracts this failure. */
  lift: number;
  /** Mean |deltaPt| / |residualPt| where the kind has a magnitude. */
  meanMagnitudePt?: number;
  /** Up to 5 concrete examples for triage. */
  examples: AttributedDivergence[];
};

export type CorpusReport = {
  generatedAt: string;
  reference: ReferenceRendererInfo;
  results: FeatureAttributedResult[];
  clusters: Cluster[];
};
