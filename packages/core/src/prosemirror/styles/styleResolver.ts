/**
 * Style Resolver for ProseMirror Editor
 *
 * Resolves OOXML style definitions to final paragraph and run properties.
 * Handles the cascade:
 * 1. Document defaults (docDefaults)
 * 2. Normal style (if no explicit styleId)
 * 3. Style chain (basedOn inheritance - already resolved by styleParser)
 * 4. Inline properties
 *
 * For paragraphs inside a table cell, `resolveParagraphStyleInTable` inserts
 * an additional layer between docDefaults and the Normal/style chain: the
 * enclosing table style's own paragraph-spacing (`w:pPr`). See its doc
 * comment for the full cascade order.
 *
 * Based on ECMA-376 style cascade rules.
 */

import type {
  StyleDefinitions,
  Style,
  DocDefaults,
  ParagraphFormatting,
  TextFormatting,
} from "../../types/document";
import { mergeParagraphFormatting } from "../../utils/paragraphFormattingMerge";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";

/**
 * Resolved style properties ready for rendering
 */
export type ResolvedParagraphStyle = {
  /** Paragraph formatting (alignment, spacing, indentation, etc.) */
  paragraphFormatting?: ParagraphFormatting;
  /** Default run formatting from the style */
  runFormatting?: TextFormatting;
};

/**
 * Paragraph-spacing fields sourced from the table style enclosing a cell
 * paragraph: the table style's own `w:pPr`, then the applicable
 * `w:tblStylePr` conditional region's `w:pPr` layered on top (already merged
 * by the caller — see `resolveTableBaseStyle`/`resolveTableStyleConditional`
 * in toProseDoc.ts).
 *
 * Deliberately narrower than the full {@link ParagraphFormatting}: this is
 * the layer-2 overlay described on {@link StyleResolver.resolveParagraphStyleInTable}
 * and is scoped to the spacing fields responsible for folio's table-row
 * height bug (table styles like Word's `TableGrid` zero these out relative
 * to `docDefaults`). Alignment, indentation, borders, etc. are out of scope
 * here — applying the whole `pPr` would risk regressing unrelated cell
 * layout that isn't part of this fix.
 */
export type TableCellParagraphSpacingOverlay = Pick<
  ParagraphFormatting,
  "spaceBefore" | "spaceAfter" | "lineSpacing" | "lineSpacingRule" | "contextualSpacing"
>;

/**
 * Word's default-template Normal style, used as a last-resort fallback for a
 * bare document that ships neither a `Normal` style nor `w:docDefaults`: 8pt
 * (160 twips) after spacing, 1.08x line spacing — the values Word writes into
 * a freshly created document's Normal.
 *
 * It is NOT applied when the document provides `w:docDefaults`. Those are the
 * authoritative paragraph defaults per ECMA-376 §17.7.5; an empty
 * `w:pPrDefault` means zero spacing / single line, so synthesizing this style
 * over it wrongly inflates style-less paragraphs (e.g. generated-document
 * table cells whose author relied on the compact default). See
 * eigenpal/docx-editor#909.
 */
const BUILTIN_NORMAL_STYLE: Style = {
  styleId: "Normal",
  type: "paragraph",
  name: "Normal",
  default: true,
  pPr: {
    spaceAfter: 160,
    lineSpacing: 259,
    lineSpacingRule: "auto",
  },
};

/**
 * StyleResolver provides efficient access to resolved style properties
 */
export class StyleResolver {
  private readonly stylesById: Map<string, Style>;
  private readonly docDefaults: DocDefaults | undefined;
  private readonly defaultParagraphStyle: Style | undefined;
  private readonly defaultCharacterStyle: Style | undefined;
  private readonly defaultTableStyle: Style | undefined;

  constructor(styleDefinitions: StyleDefinitions | undefined) {
    this.stylesById = new Map();
    this.docDefaults = styleDefinitions?.docDefaults;

    // Build lookup map
    if (styleDefinitions?.styles) {
      for (const style of styleDefinitions.styles) {
        if (style.styleId) {
          this.stylesById.set(style.styleId, style);
        }
      }
    }

    // Find default paragraph style
    this.defaultParagraphStyle = this.findDefaultStyle("paragraph");
    this.defaultCharacterStyle = this.findDefaultStyle("character");
    this.defaultTableStyle = this.findDefaultStyle("table");
  }

  /**
   * Get a style by ID
   */
  getStyle(styleId: string): Style | undefined {
    return this.stylesById.get(styleId);
  }

  /**
   * Get all available paragraph styles (for toolbar dropdown)
   */
  getParagraphStyles(): Style[] {
    const styles: Style[] = [];
    for (const style of this.stylesById.values()) {
      if (style.type === "paragraph" && !style.hidden && !style.semiHidden) {
        styles.push(style);
      }
    }
    // Sort by uiPriority, then by name
    return styles.toSorted((a, b) => {
      const priorityA = a.uiPriority ?? 99;
      const priorityB = b.uiPriority ?? 99;
      if (priorityA !== priorityB) {
        return priorityA - priorityB;
      }
      return (a.name ?? a.styleId).localeCompare(b.name ?? b.styleId);
    });
  }

  /**
   * Resolve paragraph style properties, including docDefaults cascade
   *
   * @param styleId - The style ID to resolve (e.g., 'Heading1', 'Normal')
   * @returns Resolved paragraph and run formatting
   */
  resolveParagraphStyle(styleId: string | undefined | null): ResolvedParagraphStyle {
    return this.resolveParagraphStyleCascade(styleId, undefined);
  }

  /**
   * Resolve paragraph style properties for a paragraph inside a table cell,
   * layering the enclosing table style's paragraph-spacing fields into the
   * cascade between docDefaults and the paragraph's own style chain.
   *
   * Per ECMA-376 §17.7.2, a cell paragraph's properties resolve in this
   * order (later wins):
   *   1. docDefaults (`w:pPrDefault`)
   *   2. the enclosing table style's `w:pPr` (`tableParagraphOverlay`, here)
   *   3. the paragraph's own style chain (`w:pStyle` + `basedOn` ancestors)
   *   4. direct formatting on the paragraph (`w:pPr`)
   *
   * Layer 4 is applied by the caller (`paragraphFormattingToAttrs`), which
   * already prefers the paragraph's direct formatting over this method's
   * result. Layer 3 is applied below, after the overlay, so an explicit
   * paragraph style still wins over the table for any field it sets.
   *
   * @param styleId - The paragraph's style ID (e.g., 'Heading1', 'Normal')
   * @param tableParagraphOverlay - Paragraph-spacing fields from the
   *   enclosing table style (base `pPr` merged with the applicable
   *   conditional region), or `undefined` when the paragraph isn't in a
   *   table cell, the table has no style, or the style sets no spacing.
   * @returns Resolved paragraph and run formatting
   */
  resolveParagraphStyleInTable(
    styleId: string | undefined | null,
    tableParagraphOverlay: TableCellParagraphSpacingOverlay | undefined,
  ): ResolvedParagraphStyle {
    return this.resolveParagraphStyleCascade(styleId, tableParagraphOverlay);
  }

  private resolveParagraphStyleCascade(
    styleId: string | undefined | null,
    tableParagraphOverlay: TableCellParagraphSpacingOverlay | undefined,
  ): ResolvedParagraphStyle {
    const result: ResolvedParagraphStyle = {};

    // Layer 1: document defaults
    if (this.docDefaults?.pPr) {
      result.paragraphFormatting = { ...this.docDefaults.pPr };
    }
    if (this.docDefaults?.rPr) {
      result.runFormatting = { ...this.docDefaults.rPr };
    }

    // Layer 2: enclosing table style's paragraph spacing (cell paragraphs
    // only — undefined for everything else, a no-op here).
    if (tableParagraphOverlay) {
      const merged = mergeParagraphFormatting(result.paragraphFormatting, tableParagraphOverlay);
      if (merged !== undefined) {
        result.paragraphFormatting = merged;
      }
    }

    // Layer 3: the paragraph's own style chain (Normal when styleId is absent)
    if (!styleId) {
      if (this.defaultParagraphStyle) {
        this.mergeStyleIntoResult(result, this.defaultParagraphStyle);
      }
      return result;
    }

    // Get the requested style (already has basedOn chain resolved by styleParser)
    const style = this.stylesById.get(styleId);
    if (!style) {
      // Style not found, fall back to Normal
      if (this.defaultParagraphStyle) {
        this.mergeStyleIntoResult(result, this.defaultParagraphStyle);
      }
      return result;
    }

    // Merge style properties into result
    this.mergeStyleIntoResult(result, style);

    return result;
  }

  /**
   * Resolve the style applied to the paragraph that follows one styled with
   * `styleId` when the user presses Enter (OOXML `w:next`, §17.7.4.10).
   *
   * Returns null when the style has no `w:next`, when `next` points back at
   * the same style (the common heading-stays-heading case is handled by the
   * caller), when the style is unknown, when `next` is a dangling reference
   * to a style that does not exist, or when the target is not a paragraph
   * style (malformed DOCX may point `w:next` at a character or table style;
   * writing such an ID into a paragraph's `styleId` would create an invalid
   * reference).
   */
  getNextStyleId(styleId: string | undefined | null): string | null {
    if (!styleId) {
      return null;
    }
    const next = this.stylesById.get(styleId)?.next;
    if (!next || next === styleId) {
      return null;
    }
    const target = this.stylesById.get(next);
    if (!target || target.type !== "paragraph") {
      return null;
    }
    return next;
  }

  /**
   * Get all available table styles (for style gallery)
   */
  getTableStyles(): Style[] {
    const styles: Style[] = [];
    for (const style of this.stylesById.values()) {
      if (style.type === "table" && !style.hidden && !style.semiHidden) {
        styles.push(style);
      }
    }
    return styles.toSorted((a, b) => {
      const priorityA = a.uiPriority ?? 99;
      const priorityB = b.uiPriority ?? 99;
      if (priorityA !== priorityB) {
        return priorityA - priorityB;
      }
      return (a.name ?? a.styleId).localeCompare(b.name ?? b.styleId);
    });
  }

  /**
   * Resolve run (character) style properties
   *
   * @param styleId - The character style ID to resolve
   * @returns Resolved text formatting
   */
  resolveRunStyle(styleId: string | undefined | null): TextFormatting | undefined {
    // Start with document defaults
    let result: TextFormatting = {};
    if (this.docDefaults?.rPr) {
      result = { ...this.docDefaults.rPr };
    }

    const defaultCharacterRpr = this.defaultCharacterStyle?.rPr;
    if (defaultCharacterRpr) {
      result = mergeTextFormatting(result, defaultCharacterRpr) ?? {};
    }

    // Get the requested style
    const style = styleId ? this.stylesById.get(styleId) : undefined;
    if (!style?.rPr) {
      return Object.keys(result).length > 0 ? result : undefined;
    }

    // Merge style's run properties
    const merged = mergeTextFormatting(result, style.rPr);

    return merged && Object.keys(merged).length > 0 ? merged : undefined;
  }

  /**
   * Get a character style's own properties WITHOUT docDefaults.
   * Used when the caller already has docDefaults applied (e.g., from paragraph style resolution).
   * This prevents docDefault fonts from incorrectly overriding paragraph style fonts.
   */
  getRunStyleOwnProperties(styleId: string | undefined | null): TextFormatting | undefined {
    if (!styleId) {
      return undefined;
    }

    const style = this.stylesById.get(styleId);
    if (!style?.rPr) {
      return undefined;
    }

    return Object.keys(style.rPr).length > 0 ? { ...style.rPr } : undefined;
  }

  /**
   * Get document defaults
   */
  getDocDefaults(): DocDefaults | undefined {
    return this.docDefaults;
  }

  /**
   * Get default paragraph style (usually "Normal")
   */
  getDefaultParagraphStyle(): Style | undefined {
    return this.defaultParagraphStyle;
  }

  /**
   * Get default character style.
   */
  getDefaultCharacterStyle(): Style | undefined {
    return this.defaultCharacterStyle;
  }

  /**
   * Get default table style.
   */
  getDefaultTableStyle(): Style | undefined {
    return this.defaultTableStyle;
  }

  /**
   * Check if a style exists
   */
  hasStyle(styleId: string): boolean {
    return this.stylesById.has(styleId);
  }

  // ============================================================================
  // Private helpers
  // ============================================================================

  private findDefaultStyle(type: "paragraph" | "character" | "table"): Style | undefined {
    // First try to find explicitly marked default
    for (const style of this.stylesById.values()) {
      if (style.type === type && style.default) {
        return style;
      }
    }
    // Fall back to "Normal" for paragraph styles. Only synthesize the
    // default-template Normal for a bare document with no docDefaults; when
    // docDefaults exist they are the authoritative paragraph defaults and an
    // absent Normal must not re-introduce the template's 8pt/1.08 spacing.
    if (type === "paragraph") {
      return this.stylesById.get("Normal") ?? (this.docDefaults ? undefined : BUILTIN_NORMAL_STYLE);
    }
    return undefined;
  }

  private mergeStyleIntoResult(result: ResolvedParagraphStyle, style: Style): void {
    if (style.pPr) {
      const merged = mergeParagraphFormatting(result.paragraphFormatting, style.pPr);
      if (merged !== undefined) {
        result.paragraphFormatting = merged;
      }
    }
    if (style.rPr) {
      const merged = mergeTextFormatting(result.runFormatting, style.rPr);
      if (merged !== undefined) {
        result.runFormatting = merged;
      }
    }
  }
}

/**
 * Create a style resolver from document's style definitions
 */
export function createStyleResolver(styleDefinitions: StyleDefinitions | undefined): StyleResolver {
  return new StyleResolver(styleDefinitions);
}
